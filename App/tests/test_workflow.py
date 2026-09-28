import json
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path
import shutil
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from pypdf import PdfReader

from certificate_core import read_batches, status_value, supplier_lines, vendor_font
from run_certificates import run
from workbook_store import WorkbookSession, exclusive_file, update_issued_status
from workflow_support import load_settings, notify_once, run_lock, save_json

ROOT = Path(__file__).resolve().parents[1]
TODAY = date(2026, 9, 23)
HEADERS = ["SN", "Inward Date", "Supplier Name", "Location", "Item", "Quantity in Pcs", "Per Pcs Weight in (KG)", "Weight (KG)", "Form 6 No", "Client Email", "Reference No.", "COD No", "COD Date", "Certificate Issued", "Notes"]


def make_book(path, rows):
    book = Workbook()
    sheet = book.active
    sheet.title = "Inward Register"
    sheet.append(HEADERS)
    for index, overrides in enumerate(rows, 2):
        values = {"SN": str(index - 1), "Inward Date": TODAY - timedelta(days=31), "Supplier Name": f"Dummy Client {index - 1} Ltd", "Location": "Mumbai", "Item": "Dummy computers", "Quantity in Pcs": 2, "Per Pcs Weight in (KG)": 3.5, "Weight (KG)": f"=F{index}*G{index}", "Form 6 No": f"FORM6-{index}", "Client Email": f"client{index}@example.test", "Reference No.": f"REF-{index}", "Certificate Issued": "Ready"}
        values.update(overrides)
        sheet.append([values.get(h) for h in HEADERS])
        for cell in sheet[index]:
            cell.border = Border(bottom=Side(style="thin", color="CCCCCC"))
            cell.alignment = Alignment(vertical="center")
        sheet.cell(index, 14).fill = PatternFill("solid", fgColor="FFFFFF")
    sheet.freeze_panes = "D2"
    sheet.column_dimensions["C"].width = 44
    sheet.auto_filter.ref = sheet.dimensions
    validation = DataValidation(type="list", formula1='"Ready,N/A,Blank,Completed,Sent"', allow_blank=True)
    sheet.add_data_validation(validation)
    validation.add("N2:N1000")
    for value, fill, color in (("READY", "C6EFCE", "006100"), ("N/A", "FFFFFF", "202020"),
                               ("BLANK", "FFC7CE", "9C0006"), ("COMPLETED", "FFEB9C", "7F6000"),
                               ("SENT", "DDEBF7", "1F4E78")):
        sheet.conditional_formatting.add(
            "N2:N1000",
            FormulaRule(formula=[f'UPPER(TRIM($N2))="{value}"'],
                        fill=PatternFill("solid", fgColor=fill), font=Font(color=color)),
        )
    sheet["A1"].font = Font(bold=True)
    sheet.merge_cells("O2:P2")
    book.save(path)
    book.close()
    with ZipFile(path, "a") as archive:
        archive.writestr("customXml/preserve.xml", b"<original>Preserve this native workbook content</original>")


class FakeOutlook:
    saved = []

    def __init__(self, mailbox):
        self.mailbox = mailbox

    def save_draft(self, batch, subject, body, pdf, banner):
        self.saved.append((batch["source_id"], batch["case_id"], subject, str(pdf)))
        return f"DUMMY-DRAFT-{len(self.saved)}"


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.paths = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsx"]
        save_json(self.root / "settings.example.json", {"eligibility_days": 30, "certificate_number_start": 1, "initial_cod_confirmed": True, "certificate_number_width": 4, "certificate_number_prefix": "", "workbook_paths": ["", ""]})
        (self.root / "templates").mkdir()
        shutil.copy2(ROOT / "templates" / "Certificate of Recycling.png", self.root / "templates" / "Certificate of Recycling.png")
        FakeOutlook.saved = []
        self.mock = patch("run_certificates.OutlookWriter", FakeOutlook)
        self.mock.start()

    def tearDown(self):
        self.mock.stop()
        self.temp.cleanup()

    def execute(self, **kwargs):
        result = run(self.paths, "dummy@example.test", self.root, as_of=TODAY, progress=lambda _: None, **kwargs)
        self.report = json.loads((result[2] / "Results.json").read_text())
        return result

    def one_ready(self):
        make_book(self.paths[0], [{}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0061"}])

    def test_shared_sequence_status_gate_and_no_duplicates_on_repeat(self):
        make_book(self.paths[0], [{"COD No": "0061"}, {"SN": "2"}, {"SN": "3", "Certificate Issued": ""}, {"SN": "4", "Certificate Issued": "N/a", "Client Email": None}])
        make_book(self.paths[1], [{"COD No": "0064"}, {"SN": "2"}])
        created, review, _ = self.execute()
        self.assertEqual((created, review), (4, 1), self.report)
        self.assertEqual(read_batches(self.paths[0])[1]["cod_no"], "0065")
        self.assertEqual(read_batches(self.paths[1])[1]["cod_no"], "0066")
        self.assertEqual(len(FakeOutlook.saved), 4)
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 4)
        self.assertEqual(self.report["created"], 0)

    def test_preview_does_not_consume_number_or_change_workbooks(self):
        self.one_ready()
        before = [p.read_bytes() for p in self.paths]
        created, _, output = self.execute(create_drafts=False, limit=1)
        self.assertEqual(created, 1, self.report)
        self.assertIn("PREVIEW", PdfReader(next(output.glob("*.pdf"))).pages[0].extract_text())
        self.assertEqual(before, [p.read_bytes() for p in self.paths])
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        self.assertIsNone(db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone())
        self.assertEqual(db.execute("SELECT count(*) FROM batches").fetchone()[0], 0)
        db.close()
        self.execute()
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0062")

    def test_missing_fields_and_white_rows_stay_visible_before_the_reminder_age(self):
        make_book(self.paths[0], [{"Certificate Issued": "", "Inward Date": TODAY}, {"SN": "2", "Reference No.": None}, {"SN": "3", "Inward Date": None}])
        make_book(self.paths[1], [{"Certificate Issued": "Not applicable", "Reference No.": None}])
        self.assertEqual(self.execute()[:2], (0, 3), self.report)
        self.assertIn("blank (white)", self.report["rows"][0]["detail"])
        self.assertFalse(FakeOutlook.saved)

    def test_unknown_status_and_missing_form6_are_reviewed_before_numbering(self):
        make_book(self.paths[0], [{"Certificate Issued": "Completed"}, {"SN": "2", "Form 6 No": None}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A"}])
        self.execute()
        detail = str(self.report)
        self.assertIn("Drafted needs a matching saved Outlook draft", detail)
        self.assertIn("missing Form 6 number / manifest number", detail)
        self.assertFalse(FakeOutlook.saved)
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        self.assertEqual(db.execute("SELECT count(*) FROM batches").fetchone()[0], 0)
        db.close()

    def test_explicit_blank_status_is_pending_and_never_allocates_or_drafts(self):
        make_book(self.paths[0], [{"Certificate Issued": "Blank"}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A"}])
        self.assertEqual(self.execute()[:2], (0, 1), self.report)
        self.assertEqual(status_value("Blank"), "blank")
        self.assertEqual(status_value(" blank "), "blank")
        self.assertIn("marked Blank", self.report["rows"][0]["detail"])
        self.assertFalse(FakeOutlook.saved)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "")
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        self.assertEqual(db.execute("SELECT count(*) FROM batches").fetchone()[0], 0)
        db.close()

    def test_form6_mapping_and_issue_date(self):
        self.one_ready()
        self.execute()
        text = PdfReader(FakeOutlook.saved[0][3]).pages[0].extract_text()
        self.assertIn("FORM6-2", text)
        self.assertIn(TODAY.strftime("%d.%m.%Y"), text)
        batch = read_batches(self.paths[0])[0]
        self.assertEqual(batch["cod_date"].date(), TODAY)

    def test_completed_to_ready_redraft_reuses_same_cod_and_date(self):
        self.one_ready()
        self.execute()
        first = read_batches(self.paths[0])[0]
        self.assertEqual((first["cod_no"], first["issued"]), ("0062", "Drafted"))
        original_date = first["cod_date"].date()

        from draft_service import History
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            history.acknowledge_sent("1", "0062")
        finally:
            history.close()

        book = load_workbook(self.paths[0])
        book["Inward Register"]["N2"] = "Ready"
        book.save(self.paths[0])
        book.close()
        self.execute()

        redrafted = read_batches(self.paths[0])[0]
        self.assertEqual((redrafted["cod_no"], redrafted["issued"]), ("0062", "Drafted"))
        self.assertEqual(redrafted["cod_date"].date(), original_date)
        self.assertEqual(len(FakeOutlook.saved), 2)
        self.assertIn("0062", FakeOutlook.saved[1][2])
        reissue = next(row for row in self.report["rows"] if row["outcome"] == "draft_created")
        self.assertIn("existing COD 0062 and its original date", reissue["detail"])
        self.assertIn("earlier draft was left untouched", reissue["detail"])
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        try:
            self.assertEqual(db.execute("SELECT generation,certificate_no,status FROM batches WHERE case_id='1'").fetchone(), (2, "0062", "saved"))
            archived = db.execute("SELECT generation,certificate_no,status,sent_at FROM archived_issuances WHERE case_id='1'").fetchone()
            self.assertEqual(archived[:3], (1, "0062", "saved"))
            self.assertTrue(archived[3])
            self.assertIsNone(db.execute("SELECT 1 FROM sent_acknowledgements WHERE case_id='1'").fetchone())
        finally:
            db.close()

        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 2)

        book = load_workbook(self.paths[0])
        sheet = book["Inward Register"]
        next_row = sheet.max_row + 1
        next_values = {
            "SN": "NEXT-1", "Inward Date": TODAY - timedelta(days=31),
            "Supplier Name": "Next Test Client", "Location": "Pune",
            "Item": "Dummy laptops", "Quantity in Pcs": 1,
            "Per Pcs Weight in (KG)": 2.5, "Weight (KG)": f"=F{next_row}*G{next_row}",
            "Form 6 No": "FORM6-NEXT-1", "Client Email": "next@example.test",
            "Reference No.": "REF-NEXT-1", "Certificate Issued": "Ready",
        }
        for col, header in enumerate(HEADERS, 1):
            sheet.cell(next_row, col).value = next_values.get(header)
        book.save(self.paths[0])
        book.close()
        self.execute()
        next_batch = next(batch for batch in read_batches(self.paths[0]) if batch["case_id"] == "NEXT-1")
        self.assertEqual(next_batch["cod_no"], "0063")
        self.assertEqual(len(FakeOutlook.saved), 3)

    def test_original_parts_formulas_layout_validation_and_backup_preserved(self):
        self.one_ready()
        original = self.paths[0].read_bytes()
        self.execute()
        with ZipFile(BytesIO(original)) as before, ZipFile(self.paths[0]) as after:
            for name in before.namelist():
                if name not in {"xl/worksheets/sheet1.xml", "xl/styles.xml"}:
                    self.assertEqual(before.read(name), after.read(name), name)
        book = load_workbook(self.paths[0])
        sheet = book.active
        self.assertEqual(sheet["H2"].value, "=F2*G2")
        self.assertEqual(sheet.freeze_panes, "D2")
        self.assertEqual(sheet.column_dimensions["C"].width, 44)
        self.assertEqual(str(sheet.merged_cells), "O2:P2")
        self.assertEqual(len(sheet.data_validations.dataValidation), 1)
        self.assertEqual(sheet.data_validations.dataValidation[0].formula1, '"Ready,N/A,Blank"')
        rules = sheet.conditional_formatting["N2:N1000"]
        by_formula = {rule.formula[0]: rule.dxf.fill.fgColor.rgb[-6:] for rule in rules}
        self.assertEqual(by_formula['UPPER(TRIM($N2))="DRAFTED"'], "FFEB9C")
        self.assertEqual(by_formula['UPPER(TRIM($N2))="COMPLETED"'], "DDEBF7")
        book.close()
        backups = list((self.root / "Data" / "Workbook Backups").glob("Plant 1_*.xlsx"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), original)

    def test_excel_write_failure_recovers_same_cod_without_second_draft(self):
        self.one_ready()
        with patch.object(WorkbookSession, "commit", side_effect=OSError("simulated save failure")):
            self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "")
        self.assertIn("retry only the Excel update", self.report["rows"][-1]["detail"])
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0062")
        self.assertEqual(self.report["updated"], 1)

    def test_uncertain_outlook_save_never_retries_automatically(self):
        self.one_ready()
        with patch.object(FakeOutlook, "save_draft", side_effect=RuntimeError("save uncertain")) as save:
            self.execute()
            self.execute()
            self.assertEqual(save.call_count, 1)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "")
        self.assertIn("unconfirmed", str(self.report))

    def test_open_workbook_stops_before_number_allocation(self):
        self.one_ready()
        with exclusive_file(self.paths[1]):
            self.execute()
        self.assertFalse(FakeOutlook.saved)
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        self.assertEqual(db.execute("SELECT count(*) FROM batches").fetchone()[0], 0)
        db.close()

    def test_na_after_a_saved_certificate_is_reported_without_a_new_alert_or_draft(self):
        self.one_ready()
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)

        book = load_workbook(self.paths[0])
        book["Inward Register"]["N2"] = "N/A"
        book.save(self.paths[0])
        book.close()

        self.execute(check_only=True)
        row = next(row for row in self.report["rows"] if row["sn"] == "1")
        self.assertEqual(row["outcome"], "already_prepared")
        self.assertFalse(row.get("system_alert", False))
        self.assertIn("N/A skips new drafts but does not delete that draft", row["detail"])
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_missing_plant_and_duplicate_cod_block_numbering(self):
        self.one_ready()
        missing = self.root / "missing.xlsx"
        original = self.paths[1]
        self.paths[1] = missing
        self.execute()
        self.assertFalse(FakeOutlook.saved)
        self.paths[1] = original
        make_book(self.paths[0], [{"COD No": "0061"}, {"SN": "2"}])
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertIn("also entered", str(self.report))
        self.assertIn("0062", FakeOutlook.saved[0][2])

    def test_unknown_number_format_is_not_guessed(self):
        self.one_ready()
        make_book(self.paths[1], [{"COD No": "0061/2026"}])
        self.execute()
        self.assertFalse(FakeOutlook.saved)
        self.assertIn("number format", str(self.report))

    def test_ready_today_ignores_legacy_thirty_day_setting(self):
        make_book(self.paths[0], [{"Inward Date": TODAY}])
        make_book(self.paths[1], [{"Inward Date": TODAY - timedelta(days=30)}])
        self.assertEqual(self.execute()[0], 2)
        self.assertEqual([r[0] for r in FakeOutlook.saved], ["plant1", "plant2"])

    def test_future_receipt_requires_review_without_allocating(self):
        make_book(self.paths[0], [{"Inward Date": TODAY + timedelta(days=1)}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A"}])
        self.assertEqual(self.execute()[:2], (0, 1))
        self.assertIn('future', str(self.report))

    def test_manual_completion_confirmation_is_idempotent_reversible_and_does_not_redraft(self):
        from draft_service import History
        self.one_ready()
        self.execute()
        history = History(self.root / 'Data' / 'certificate-history.sqlite')
        try:
            with self.assertRaises(ValueError):
                history.acknowledge_sent('1', '9999')
            with self.assertRaises(ValueError):
                history.acknowledge_sent('plant2|1', '0061')
            history.acknowledge_completed('1', '0062')
            original = history.completed_acknowledgement('1')
            created_row = next(row for row in self.report['rows'] if row['outcome'] == 'draft_created')
            self.assertTrue(created_row['awaiting_send'])
            self.assertEqual(created_row['certificate_status'], 'drafted')
            self.assertTrue(created_row['reminder_due'])
            history.acknowledge_completed('1', '0062')
            self.assertEqual(original, history.completed_acknowledgement('1'))
            update_issued_status(self.paths[0], self.root / 'Data' / 'Workbook Backups',
                                 '1', 'Inward Register', 2, '0062', 'Drafted', 'Completed')
            self.execute()
            self.assertEqual(len(FakeOutlook.saved), 1)
            prepared_row = next(row for row in self.report['rows'] if row['outcome'] == 'completed')
            self.assertFalse(prepared_row.get('awaiting_send', False))
            self.assertEqual(prepared_row['certificate_status'], 'completed')
            self.assertFalse(prepared_row['reminder_due'])
            history.undo_completed_acknowledgement('1')
            self.assertIsNone(history.completed_acknowledgement('1'))
            update_issued_status(self.paths[0], self.root / 'Data' / 'Workbook Backups',
                                 '1', 'Inward Register', 2, '0062', 'Completed', 'Drafted')
            self.execute()
            self.assertEqual(len(FakeOutlook.saved), 1)
            prepared_row = next(row for row in self.report['rows'] if row['outcome'] == 'already_prepared')
            self.assertTrue(prepared_row['awaiting_send'])
            self.assertEqual(prepared_row['certificate_status'], 'drafted')
            self.assertTrue(prepared_row['reminder_due'])
        finally:
            history.close()
        book = load_workbook(self.paths[0])
        try:
            rules = book["Inward Register"].conditional_formatting["N2:N1000"]
            by_formula = {rule.formula[0]: rule.dxf.fill.fgColor.rgb[-6:] for rule in rules}
            self.assertEqual(by_formula['UPPER(TRIM($N2))="DRAFTED"'], "FFEB9C")
            self.assertEqual(by_formula['UPPER(TRIM($N2))="COMPLETED"'], "DDEBF7")
            self.assertEqual(book["Inward Register"].data_validations.dataValidation[0].formula1,
                             '"Ready,N/A,Blank"')
        finally:
            book.close()

    def test_legacy_completed_draft_remains_overdue_until_operator_confirms(self):
        self.one_ready()
        self.execute()
        book = load_workbook(self.paths[0])
        book["Inward Register"]["N2"] = "Completed"
        book.save(self.paths[0])
        book.close()

        self.execute(check_only=True)
        row = next(item for item in self.report["rows"] if item["sn"] == "1")
        self.assertEqual(row["certificate_status"], "drafted")
        self.assertTrue(row["reminder_due"])
        self.assertTrue(row["awaiting_send"])
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(read_batches(self.paths[0])[0]["issued"], "Completed",
                         "read-only check must not rewrite old workbooks")

    def test_no_arbitrary_start_when_both_files_have_no_earlier_cod(self):
        make_book(self.paths[0], [{}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A"}])
        save_json(self.root / "settings.local.json", {"initial_cod_confirmed": False})
        self.execute()
        self.assertFalse(FakeOutlook.saved)
        self.assertIn("approved first COD", str(self.report))

    def test_check_only_needs_no_start_number_and_can_preserve_latest_report(self):
        make_book(self.paths[0], [{}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A"}])
        save_json(self.root / "settings.local.json", {"initial_cod_confirmed": False})
        latest_json = self.root / "Output" / "Latest Results.json"
        latest_text = self.root / "Output" / "Latest Results.txt"
        latest_json.parent.mkdir(parents=True, exist_ok=True)
        latest_json.write_text('{"marker":"plant report"}', encoding="utf-8")
        latest_text.write_text("plant report", encoding="utf-8")

        created, review, output = run(
            self.paths, "", self.root, as_of=TODAY, progress=lambda _: None,
            create_drafts=False, check_only=True, publish_latest=False,
        )

        result = json.loads((output / "Results.json").read_text(encoding="utf-8"))
        self.assertEqual((created, review), (0, 0))
        self.assertEqual(sum(row["outcome"] == "ready" for row in result["rows"]), 1)
        self.assertEqual(sum(row["outcome"] == "not_applicable" for row in result["rows"]), 1)
        self.assertEqual(latest_json.read_text(encoding="utf-8"), '{"marker":"plant report"}')
        self.assertEqual(latest_text.read_text(encoding="utf-8"), "plant report")
        self.assertFalse(FakeOutlook.saved)

    def test_read_only_check_keeps_an_existing_unsent_draft_in_the_alert_queue(self):
        from draft_service import History
        self.one_ready()
        self.execute()
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            history.db.execute("UPDATE batches SET status='draft_saved' WHERE case_id='1'")
            history.db.commit()
        finally:
            history.close()

        self.execute(create_drafts=False, check_only=True)
        rows = [row for row in self.report["rows"] if row["outcome"] == "ready"]
        self.assertEqual(len(rows), 1)
        self.assertTrue(rows[0]["awaiting_send"])
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_one_draft_action_uses_the_shared_workflow_and_updates_one_row(self):
        paths = [self.root / "Dummy_Inward_Plant_1.xlsx", self.root / "Dummy_Inward_Plant_2.xlsx"]
        make_book(paths[0], [{}])
        make_book(paths[1], [{"Certificate Issued": "N/A"}])
        self.paths = paths
        save_json(self.root / "settings.local.json", {"certificate_number_start": 302, "initial_cod_confirmed": True})
        second_before = paths[1].read_bytes()

        created, review, output = run(paths, "dummy@example.test", self.root, limit=1, as_of=TODAY, progress=lambda _: None)
        report = json.loads((output / "Results.json").read_text(encoding="utf-8"))

        self.assertEqual((created, review), (1, 0), report)
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertIn("0302", FakeOutlook.saved[0][2])
        self.assertIn("FORM6-2", PdfReader(FakeOutlook.saved[0][3]).pages[0].extract_text())
        self.assertEqual(read_batches(paths[0])[0]["cod_no"], "0302")
        self.assertEqual(read_batches(paths[0])[0]["cod_date"].date(), TODAY)
        self.assertEqual(paths[1].read_bytes(), second_before)
        self.assertTrue((self.root / "Data" / "certificate-history.sqlite").is_file())
        self.assertEqual(sum(row["outcome"] == "draft_created" for row in report["rows"]), 1)

        run(paths, "dummy@example.test", self.root, limit=1, as_of=TODAY, progress=lambda _: None)
        self.assertEqual(len(FakeOutlook.saved), 1, "rechecking the same row must not create a duplicate draft")

    def test_batch_action_processes_every_eligible_row_regardless_of_filename(self):
        paths = [self.root / "Dummy_Inward_Plant_1.xlsx", self.root / "Dummy_Inward_Plant_2.xlsx"]
        make_book(paths[0], [{}, {"SN": "2"}])
        make_book(paths[1], [{"Certificate Issued": "N/A"}])
        self.paths = paths

        created, review, output = run(paths, "dummy@example.test", self.root, as_of=TODAY, progress=lambda _: None)
        report = json.loads((output / "Results.json").read_text(encoding="utf-8"))

        self.assertEqual((created, review), (2, 0), report)
        self.assertEqual(sum(row["outcome"] == "draft_created" for row in report["rows"]), 2)
        self.assertEqual(sum(row["outcome"] == "not_applicable" for row in report["rows"]), 1)
        self.assertEqual([read_batches(paths[0])[index]["cod_no"] for index in range(2)], ["0001", "0002"])
        self.assertEqual(len(FakeOutlook.saved), 2)

    def test_selected_workbook_filename_does_not_restrict_recipient_domain(self):
        paths = [self.root / "Dummy_Inward_Plant_1.xlsx", self.root / "Dummy_Inward_Plant_2.xlsx"]
        make_book(paths[0], [{"Client Email": "client@company.com"}])
        make_book(paths[1], [{"Certificate Issued": "N/A"}])
        self.paths = paths

        created, review, output = run(paths, "dummy@example.test", self.root, limit=1, as_of=TODAY, progress=lambda _: None)
        report = json.loads((output / "Results.json").read_text(encoding="utf-8"))

        self.assertEqual((created, review), (1, 0), report)
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(read_batches(paths[0])[0]["cod_no"], "0001")

    def test_existing_excel_cod_without_history_is_used_as_custom_and_advances_sequence(self):
        paths = [self.root / "Dummy_Inward_Plant_1.xlsx", self.root / "Dummy_Inward_Plant_2.xlsx"]
        make_book(paths[0], [{"COD No": "0302", "COD Date": TODAY}, {"SN": "2"}])
        make_book(paths[1], [{"Certificate Issued": "N/A"}])
        self.paths = paths
        save_json(self.root / "settings.local.json", {"certificate_number_start": 302, "initial_cod_confirmed": True})

        created, review, output = run(paths, "dummy@example.test", self.root, as_of=TODAY, progress=lambda _: None)
        report = json.loads((output / "Results.json").read_text(encoding="utf-8"))
        plant1 = read_batches(paths[0])

        self.assertEqual((created, review), (2, 0), report)
        self.assertEqual(plant1[0]["cod_no"], "0302")
        self.assertEqual(plant1[0]["issued"], "Drafted")
        self.assertEqual(plant1[1]["cod_no"], "0303")
        self.assertEqual(len(FakeOutlook.saved), 2)
        self.assertIn("0302", FakeOutlook.saved[0][2])
        self.assertIn("0303", FakeOutlook.saved[1][2])

    def test_duplicate_custom_cod_only_blocks_the_rows_that_claim_it(self):
        make_book(self.paths[0], [{"COD No": "0061"}, {"SN": "2", "COD No": "0061"}, {"SN": "3"}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0060"}])
        created, review, _ = self.execute()
        self.assertEqual((created, review), (1, 2), self.report)
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(FakeOutlook.saved[0][1], "3")
        self.assertIn("0062", FakeOutlook.saved[0][2])
        self.assertEqual(read_batches(self.paths[0])[2]["cod_no"], "0062")

    def test_reissue_restores_removed_or_changed_cod_and_date_from_history(self):
        self.one_ready()
        self.execute()
        original = read_batches(self.paths[0])[0]
        original_date = original["cod_date"].date()

        book = load_workbook(self.paths[0])
        sheet = book["Inward Register"]
        sheet["N2"] = "Ready"
        sheet["L2"] = "0999"
        sheet["M2"] = TODAY - timedelta(days=1)
        book.save(self.paths[0])
        book.close()

        created, review, _ = self.execute()
        redrafted = read_batches(self.paths[0])[0]
        self.assertEqual((created, review), (1, 0), self.report)
        self.assertEqual(len(FakeOutlook.saved), 2)
        self.assertEqual((redrafted["cod_no"], redrafted["issued"]), (original["cod_no"], "Drafted"))
        self.assertEqual(redrafted["cod_date"].date(), original_date)
        self.assertIn(original["cod_no"], FakeOutlook.saved[-1][2])
        self.assertNotIn("0999", FakeOutlook.saved[-1][2])

    def test_removed_cod_from_legacy_draft_is_restored_without_duplicate(self):
        self.one_ready()
        self.execute()
        original = read_batches(self.paths[0])[0]
        original_date = original["cod_date"].date()
        db = sqlite3.connect(self.root / "Data" / "certificate-history.sqlite")
        db.execute("UPDATE batches SET status_written=0 WHERE case_id='1'")
        db.commit()
        db.close()

        book = load_workbook(self.paths[0])
        sheet = book["Inward Register"]
        sheet["N2"] = "Ready"
        sheet["L2"] = None
        sheet["M2"] = None
        book.save(self.paths[0])
        book.close()

        created, review, _ = self.execute()
        restored = read_batches(self.paths[0])[0]
        self.assertEqual((created, review), (0, 0), self.report)
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual((restored["cod_no"], restored["issued"]), (original["cod_no"], "Drafted"))
        self.assertEqual(restored["cod_date"].date(), original_date)

    def test_sent_row_restores_removed_cod_and_date_without_creating_a_message(self):
        self.one_ready()
        self.execute()
        original = read_batches(self.paths[0])[0]
        original_date = original["cod_date"].date()

        from draft_service import History
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            history.acknowledge_sent("1", original["cod_no"])
        finally:
            history.close()

        book = load_workbook(self.paths[0])
        sheet = book["Inward Register"]
        sheet["N2"] = "Sent"
        sheet["L2"] = None
        sheet["M2"] = None
        book.save(self.paths[0])
        book.close()

        created, review, _ = self.execute()
        restored = read_batches(self.paths[0])[0]
        self.assertEqual((created, review), (0, 0), self.report)
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual((restored["cod_no"], restored["issued"]), (original["cod_no"], "Completed"))
        self.assertEqual(restored["cod_date"].date(), original_date)

    def test_first_cod_setting_is_used_as_the_sequence_floor(self):
        self.one_ready()
        save_json(self.root / "settings.local.json", {"certificate_number_start": 302, "initial_cod_confirmed": True})
        self.execute()
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0302")

    def test_settings_migrate_one_path_and_remember_both(self):
        save_json(self.root / "settings.local.json", {"workbook_path": str(self.paths[0])})
        settings = load_settings(self.root)
        self.assertEqual(settings["workbook_paths"], [str(self.paths[0]), ""])
        settings["workbook_paths"][1] = str(self.paths[1])
        save_json(self.root / "settings.local.json", settings)
        self.assertEqual(load_settings(self.root)["workbook_paths"], [str(p) for p in self.paths])

    def test_desktop_remembers_selected_files_after_restart(self):
        import tkinter as tk
        from desktop_app import App
        window = tk.Tk()
        window.withdraw()
        try:
            app = App(window, self.root)
            with patch("desktop_app.filedialog.askopenfilename", side_effect=[str(p) for p in self.paths]):
                app.choose_workbook(0)
                app.choose_workbook(1)
            window.update_idletasks()
        finally:
            window.destroy()
        window = tk.Tk()
        window.withdraw()
        try:
            reopened = App(window, self.root)
            self.assertEqual([v.get() for v in reopened.workbooks], [str(p) for p in self.paths])
        finally:
            window.destroy()

    def test_several_certificates_in_one_workbook_keep_all_codes(self):
        make_book(self.paths[0], [{}, {"SN": "2"}, {"SN": "3"}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0061"}])
        self.assertEqual(self.execute()[:2], (3, 0), self.report)
        self.assertEqual([b["cod_no"] for b in read_batches(self.paths[0])], ["0062", "0063", "0064"])

    def test_protected_cod_cell_is_not_changed_or_given_a_number(self):
        self.one_ready()
        book = load_workbook(self.paths[0])
        book.active.protection.sheet = True
        book.save(self.paths[0])
        book.close()
        before = self.paths[0].read_bytes()
        self.execute()
        self.assertFalse(FakeOutlook.saved)
        self.assertEqual(before, self.paths[0].read_bytes())
        self.assertIn("protected", str(self.report))

    def test_no_overlapping_runs(self):
        self.one_ready()
        with run_lock(self.root):
            with self.assertRaisesRegex(RuntimeError, "already running"):
                self.execute()

    def test_notification_summary_repeats_on_new_runs_without_same_run_duplicates(self):
        notices = []
        notifier = lambda title, body: notices.append((title, body)) or True
        first = {"created": 0, "rows": [{"outcome": "needs_review", "plant": "plant1", "sn": "1", "detail": "blank", "reminder_due": True}]}
        notify_once(self.root, first, notifier, TODAY, repeat_pending=True)
        notify_once(self.root, first, notifier, TODAY, repeat_pending=True)
        self.assertEqual(len(notices), 1)
        notify_once(self.root, first, notifier, TODAY + timedelta(days=1), repeat_pending=True)
        self.assertEqual(len(notices), 2)
        second = {"created": 0, "rows": first["rows"] + [{"outcome": "needs_review", "plant": "plant2", "sn": "2", "detail": "blank", "reminder_due": True}]}
        notify_once(self.root, second, notifier, TODAY, repeat_pending=True)
        self.assertEqual(len(notices), 3)
        notify_once(self.root, second, notifier, TODAY + timedelta(days=1), repeat_pending=True)
        self.assertEqual(len(notices), 4)

    def test_long_supplier_has_two_centred_lines_within_bounds(self):
        batch = {"supplier": "Dummy Advanced Electronics and Resource Recovery Solutions Private Limited", "location": "Mumbai"}
        lines, size = supplier_lines(batch)
        self.assertEqual(len(lines), 2)
        from reportlab.pdfbase.pdfmetrics import stringWidth
        self.assertTrue(all(stringWidth(line, vendor_font(), size) <= 620 for line in lines))
        self.assertGreaterEqual(size, 12)


if __name__ == "__main__":
    unittest.main()
