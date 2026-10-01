from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO
import json
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from pypdf import PdfReader, PdfWriter

from certificate_core import (
    batch_content, certificate_layout, certificate_reminder_due, number, parse_date,
    read_batches, status_value,
)
from desktop_app import App, reminder_picker_parts, reminder_time_from_picker
from run_certificates import _safe_cod_name, certificate_filename, run
import scheduled_run
from workbook_store import WorkbookSession, _sheet_member, upsert_certificate_details
from workflow_support import Report, load_settings, notify_once, open_alert_workbook, save_json


APP = Path(__file__).resolve().parents[1]
AS_OF = date(2026, 10, 1)
HEADERS = [
    "SN", "Inward Date", "Supplier Name", "Location", "Item",
    "Quantity in Pcs", "Per Pcs Weight in (KG)", "Weight (KG)",
    "Form 6 No", "Reference No.", "COD No", "COD Date", "Certificate Issued",
]


def make_workbook(path, clients):
    book = Workbook()
    sheet = book.active
    sheet.title = "Inward Register"
    sheet.append(HEADERS)
    status_validation = DataValidation(type="list", formula1='"Ready,N/A,Blank"', allow_blank=True)
    sheet.add_data_validation(status_validation)
    for index, item in enumerate(clients, 2):
        row = [
            item.get("sn", str(index + 1000)),
            item.get("date", AS_OF - timedelta(days=item.get("age", 2))),
            item.get("supplier", f"Dummy Client {index - 1} Ltd"),
            item.get("location", "Mumbai"),
            item.get("item", "Recycled computer equipment"),
            item.get("quantity", 2),
            item.get("per_piece_weight", 3.5),
            item.get("weight", 7.0),
            item.get("form6", f"FORM6-{index}"),
            item.get("reference", f"REF-{index}"),
            item.get("cod", ""),
            item.get("cod_date"),
            item.get("status", "Ready"),
        ]
        sheet.append(row)
        status_validation.add(f"M{index}")
    book.save(path)
    book.close()


class AutomationTestBase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name) / "certificate-automation"
        (self.base / "App" / "templates").mkdir(parents=True)
        shutil.copy2(APP / "templates" / "Certificate of Recycling.png", self.base / "App" / "templates" / "Certificate of Recycling.png")
        shutil.copy2(APP / "settings.example.json", self.base / "App" / "settings.example.json")
        self.plant1 = self.base / "Plant 1.xlsx"
        self.plant2 = self.base / "Plant 2.xlsx"
        self.paths = [self.plant1, self.plant2]
        save_json(self.base / "settings.local.json", {
            "workbook_paths": [str(self.plant1), str(self.plant2)],
            "certificate_number_prefix": "",
            "certificate_number_start": 1,
            "certificate_number_width": 4,
            "reminder_time": "11:00",
            "reminders_enabled": True,
        })

    def tearDown(self):
        self.temp.cleanup()

    def execute(self, *, check_only=False, selected=None, limit=None, custom_cod=None,
                allow_warnings=False, plants=None):
        return run(self.paths, self.base, as_of=AS_OF, check_only=check_only,
                   selected=selected, limit=limit, custom_cod=custom_cod,
                   allow_warnings=allow_warnings, plants=plants, progress=lambda _: None)

    def results(self):
        return json.loads((self.base / "Data" / "Reports" / "Latest Results.json").read_text(encoding="utf-8"))

class CertificateWorkflowTests(AutomationTestBase):
    def test_bulk_generation_scoped_to_plant1_leaves_plant2_ready(self):
        make_workbook(self.plant1, [{"sn": "P1-ONLY", "supplier": "Plant One Ltd"}])
        make_workbook(self.plant2, [{"sn": "P2-ONLY", "supplier": "Plant Two Ltd"}])

        created, review, _ = self.execute(plants=("plant1",))

        self.assertEqual((created, review), (1, 0))
        plant1 = read_batches(self.plant1)[0]
        plant2 = read_batches(self.plant2)[0]
        self.assertEqual(plant1["issued"], "Generated")
        self.assertTrue(plant1["cod_no"])
        self.assertEqual((plant2["issued"], plant2["cod_no"]), ("Ready", ""))
        pdfs = list((self.base / "Certificates").glob("*.pdf"))
        self.assertEqual([path.name for path in pdfs], [certificate_filename(plant1, 1)])
        self.assertTrue(pdfs[0].name.endswith("_1.pdf"))
        self.assertEqual({row["plant"] for row in self.results()["rows"]}, {"plant1"})

    def test_bulk_generation_scoped_to_plant2_leaves_plant1_ready(self):
        make_workbook(self.plant1, [{"sn": "P1-ONLY", "supplier": "Plant One Ltd"}])
        make_workbook(self.plant2, [{"sn": "P2-ONLY", "supplier": "Plant Two Ltd"}])

        created, review, _ = self.execute(plants=("plant2",))

        self.assertEqual((created, review), (1, 0))
        plant1 = read_batches(self.plant1)[0]
        plant2 = read_batches(self.plant2)[0]
        self.assertEqual((plant1["issued"], plant1["cod_no"]), ("Ready", ""))
        self.assertEqual(plant2["issued"], "Generated")
        self.assertTrue(plant2["cod_no"])
        pdfs = list((self.base / "Certificates").glob("*.pdf"))
        self.assertEqual([path.name for path in pdfs], [certificate_filename(plant2, 2)])
        self.assertTrue(pdfs[0].name.endswith("_2.pdf"))
        self.assertEqual({row["plant"] for row in self.results()["rows"]}, {"plant2"})

    def test_bulk_generation_for_all_plants_generates_both_independently(self):
        make_workbook(self.plant1, [{"sn": "P1-ONLY", "supplier": "Plant One Ltd"}])
        make_workbook(self.plant2, [{"sn": "P2-ONLY", "supplier": "Plant Two Ltd"}])

        created, review, _ = self.execute(plants=("plant1", "plant2"))

        self.assertEqual((created, review), (2, 0))
        plant1, plant2 = read_batches(self.plant1)[0], read_batches(self.plant2)[0]
        self.assertEqual((plant1["issued"], plant2["issued"]), ("Generated", "Generated"))
        self.assertNotEqual(plant1["cod_no"], plant2["cod_no"])
        pdfs = sorted(path.name for path in (self.base / "Certificates").glob("*.pdf"))
        self.assertEqual(pdfs, sorted([certificate_filename(plant1, 1), certificate_filename(plant2, 2)]))
        self.assertEqual({row["plant"] for row in self.results()["rows"]}, {"plant1", "plant2"})

    def test_certificate_overlay_values_are_centered_and_company_line_uses_city_only(self):
        batch = {
            "supplier": "Plant 2 Godavari Biorefineries Ltd",
            "location": "Mumbai, Maharashtra, India",
            "inward_date": date(2026, 4, 25),
            "display_weight": 217.2,
            "total_weight": 220,
            "reference": "2026040047",
            "form6": "10164",
        }

        fitted, company_lines, supplier_size = certificate_layout(
            batch, "0061", date(2026, 6, 1),
        )

        self.assertEqual(company_lines, ["GODAVARI BIOREFINERIES LTD, MUMBAI"])
        self.assertEqual(supplier_size, 28)
        self.assertAlmostEqual(fitted[0][1], 523.9, places=1)
        self.assertAlmostEqual(fitted[1][1], 584.0, places=1)
        for field in fitted[2:]:
            self.assertAlmostEqual(field[1], 558.8, places=1)

    def test_newly_appended_client_row_is_found_and_can_generate(self):
        make_workbook(self.plant1, [{"sn": "EXISTING", "status": "N/A"}])
        make_workbook(self.plant2, [{"sn": "OTHER", "status": "N/A"}])
        book = load_workbook(self.plant1)
        book["Inward Register"].append([
            "NEW-CLIENT-1", AS_OF - timedelta(days=1), "New Client Ltd", "Pune",
            "Recycled laptops", 4, 2.5, 10.0, "FORM6-NEW", "REF-NEW", "", None, "Ready",
        ])
        book.save(self.plant1)
        book.close()

        self.assertEqual(self.execute(selected=("plant1", "NEW-CLIENT-1"))[0], 1)
        batches = read_batches(self.plant1)
        added = next(batch for batch in batches if batch["case_id"] == "NEW-CLIENT-1")
        self.assertEqual((added["cod_no"], added["issued"]), ("0001", "Generated"))
        self.assertTrue((self.base / "Certificates" / certificate_filename(added, 1)).is_file())
        self.assertEqual(self.results()["certificate_folder"], str(self.base / "Certificates"))

    def test_full_100_client_two_plant_generation_smoke(self):
        make_workbook(self.plant1, [{"sn": f"P1-{index:03d}", "supplier": f"Plant One Client {index} Ltd", "reference": f"P1-REF-{index}"}
                                    for index in range(1, 51)])
        make_workbook(self.plant2, [{"sn": f"P2-{index:03d}", "supplier": f"Plant Two Client {index} Ltd", "reference": f"P2-REF-{index}"}
                                    for index in range(1, 51)])

        created, review, _ = self.execute()

        output = self.base / "Certificates"
        results = self.results()
        failures = [(row["plant"], row["sn"], row["detail"])
                    for row in results["rows"] if row["outcome"] == "needs_review"]
        self.assertEqual((created, review), (100, 0), failures)
        self.assertTrue(results["scan_complete"])
        self.assertEqual(len(results["rows"]), 100)
        self.assertEqual(len(list(output.glob("*.pdf"))), 100)
        codes = [batch["cod_no"] for path in self.paths for batch in read_batches(path)]
        self.assertEqual(sorted(codes), [f"{index:04d}" for index in range(1, 101)])
        self.assertTrue(all(batch["issued"] == "Generated"
                            for path in self.paths for batch in read_batches(path)))

    def test_unreadable_second_plant_prevents_partial_generation(self):
        make_workbook(self.plant1, [{"sn": "MUST-WAIT", "status": "Ready"}])
        original_plant1 = self.plant1.read_bytes()
        self.plant2.write_bytes(b"not a valid Excel workbook")

        created, review, _ = self.execute()

        self.assertEqual(created, 0)
        self.assertGreaterEqual(review, 1)
        self.assertFalse(self.results()["scan_complete"])
        self.assertEqual(self.plant1.read_bytes(), original_plant1)
        self.assertFalse((self.base / "Certificates").exists())

    def test_continuation_item_rows_contribute_to_the_batch_weight(self):
        make_workbook(self.plant1, [{"sn": "MULTI-ITEM", "status": "Ready", "quantity": 2,
                                    "per_piece_weight": 5, "weight": 10, "item": "Laptops"}])
        make_workbook(self.plant2, [{"sn": "OTHER", "status": "N/A"}])
        book = load_workbook(self.plant1)
        book["Inward Register"].append([None, None, None, None, "Servers", 3, 4, "=F3*G3",
                                        None, None, None, None, None])
        book.save(self.plant1)
        book.close()

        batches = read_batches(self.plant1)
        self.assertEqual(len(batches), 1)
        self.assertEqual([item["item"] for item in batches[0]["items"]], ["Laptops", "Servers"])
        self.assertEqual(batches[0]["total_weight"], 22)
        self.assertEqual(self.execute()[0], 1)
        pdf = self.base / "Certificates" / certificate_filename(read_batches(self.plant1)[0], 1)
        text_content = "\n".join(page.extract_text() or "" for page in PdfReader(pdf).pages)
        self.assertIn("22.0 KG", text_content)

    def test_unsupported_excel_weight_formula_is_not_guessed(self):
        make_workbook(self.plant1, [{"sn": "FORMULA-1", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OTHER", "status": "N/A"}])
        book = load_workbook(self.plant1)
        book["Inward Register"]["H2"] = "=SUM(F2:G2)"
        book.save(self.plant1)
        book.close()

        created, review, _ = self.execute()
        row = next(item for item in self.results()["rows"] if item["sn"] == "FORMULA-1")
        self.assertEqual((created, review), (0, 1))
        self.assertIn("unsupported weight formula", row["detail"])
        self.assertEqual((read_batches(self.plant1)[0]["cod_no"], read_batches(self.plant1)[0]["issued"]),
                         ("", "Ready"))

    def test_weight_parser_handles_indian_grouping_and_rejects_malformed_values(self):
        self.assertEqual(number("1,23,456.75"), 123456.75)
        self.assertEqual(number("12,34,567"), 1234567)
        self.assertEqual(number("1,234.5"), 1234.5)
        for bad in ("1,23,45", "1,,000", "NaN", "Infinity", True):
            with self.subTest(value=bad):
                self.assertIsNone(number(bad))

    def test_windows_reserved_or_oversized_cod_filename_is_rejected(self):
        for cod in ("CON", "con.pdf", "LPT1", "COM³"):
            with self.subTest(cod=cod), self.assertRaisesRegex(ValueError, "reserved Windows filename"):
                _safe_cod_name(cod)
        with self.assertRaisesRegex(ValueError, "too long"):
            _safe_cod_name("7" * 241)
        self.assertEqual(_safe_cod_name("COD-0042"), "COD-0042")

    def test_failed_excel_write_removes_new_pdf_and_retry_reuses_cod(self):
        make_workbook(self.plant1, [{"sn": "ROLLBACK-1", "age": 2, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OTHER-1", "status": "N/A"}])

        with patch("run_certificates.WorkbookSession.commit", side_effect=OSError("simulated Excel write failure")):
            created, review, _ = self.execute()

        folder = self.base / "Certificates"
        self.assertEqual((created, review), (0, 1))
        self.assertEqual(list(folder.glob("*.pdf")), [])
        row = read_batches(self.plant1)[0]
        self.assertEqual((row["cod_no"], row["issued"]), ("", "Ready"))

        self.assertEqual(self.execute()[0], 1)
        batch = read_batches(self.plant1)[0]
        self.assertTrue((folder / certificate_filename(batch, 1)).is_file())
        self.assertFalse(any(path.name.endswith("_2.pdf") for path in folder.glob("*.pdf")))

    def test_failed_excel_write_restores_existing_same_cod_pdf(self):
        make_workbook(self.plant1, [{"sn": "ROLLBACK-2", "age": 2, "cod": "0042", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OTHER-2", "status": "N/A"}])
        folder = self.base / "Certificates"
        folder.mkdir(parents=True)
        original_pdf = b"previous certificate must survive failed reissue"
        (folder / "0042.pdf").write_bytes(original_pdf)

        with patch("run_certificates.WorkbookSession.commit", side_effect=OSError("simulated Excel write failure")):
            created, review, _ = self.execute()

        self.assertEqual((created, review), (0, 1))
        self.assertEqual((folder / "0042.pdf").read_bytes(), original_pdf)
        self.assertEqual((read_batches(self.plant1)[0]["cod_no"], read_batches(self.plant1)[0]["issued"]),
                         ("0042", "Ready"))

    def test_daily_reminder_check_never_generates_a_pdf(self):
        make_workbook(self.plant1, [{"sn": "ALERT-ONLY", "age": 30, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OTHER-ALERT", "age": 30, "status": "N/A"}])
        settings = load_settings(self.base)
        settings["reminder_time"] = "00:00"
        save_json(self.base / "settings.local.json", settings)

        with patch("scheduled_run.notify_once", return_value=False) as notify:
            self.assertEqual(scheduled_run.main(root=self.base), 0)

        self.assertTrue(notify.called)
        self.assertEqual(notify.call_args.args[1]["mode"], "Check plant files")
        self.assertEqual(notify.call_args.args[1]["created"], 0)
        self.assertEqual(list((self.base / "Certificates").glob("*.pdf")), [])
        self.assertEqual(read_batches(self.plant1)[0]["issued"], "Ready")

    def test_shared_sequence_custom_cod_blank_cod_and_single_flat_output(self):
        make_workbook(self.plant1, [
            {"sn": "P1-1", "age": 25, "supplier": "Plant One Ready Ltd"},
            {"sn": "P1-2", "age": 4, "cod": "SPECIAL/REF-0077", "cod_date": date(2026, 9, 26)},
            {"sn": "P1-3", "age": 31, "cod": "0045", "status": "Completed"},
            {"sn": "P1-4", "age": 31, "status": "N/A"},
        ])
        make_workbook(self.plant2, [
            {"sn": "P2-1", "age": 30, "supplier": "Plant Two Ready Ltd", "reference": "P2-REF-1"},
            {"sn": "P2-2", "age": 31, "cod": "0046", "status": "Generated"},
            {"sn": "P2-3", "age": 2, "status": "Blank"},
        ])
        legacy_folder = self.base / "Output" / "Certificates"
        legacy_folder.mkdir(parents=True)
        (legacy_folder / "COD_0050.pdf").write_bytes(b"existing COD filename is only a sequence marker")
        folder = self.base / "Certificates"
        folder.mkdir()
        existing_filename = certificate_filename(read_batches(self.plant2)[1], 2)
        existing_pdf = PdfWriter()
        existing_pdf.add_blank_page(width=960, height=540)
        existing_pdf.write(folder / existing_filename)
        completed_filename = certificate_filename(read_batches(self.plant1)[2], 1)
        existing_pdf.write(folder / completed_filename)

        created, review, _ = self.execute()

        self.assertEqual(created, 3)
        self.assertEqual(review, 1)  # The intentionally Blank row needs a decision.
        output = sorted(path.name for path in folder.glob("*.pdf"))
        expected = sorted([
            existing_filename,
            completed_filename,
            certificate_filename(read_batches(self.plant1)[0], 1),
            certificate_filename(read_batches(self.plant1)[1], 1),
            certificate_filename(read_batches(self.plant2)[0], 2),
        ])
        self.assertEqual(output, expected)
        self.assertTrue((legacy_folder / "COD_0050.pdf").is_file())
        rows = self.results()["rows"]
        values = {(row["plant"], row["sn"]): row for row in rows}
        self.assertEqual(values[("plant1", "P1-1")]["certificate_no"], "0078")
        self.assertEqual(values[("plant1", "P1-2")]["certificate_no"], "SPECIAL/REF-0077")
        self.assertEqual(values[("plant2", "P2-1")]["certificate_no"], "0079")
        self.assertEqual(values[("plant1", "P1-2")]["certificate_date"], "2026-09-26")
        self.assertEqual(values[("plant2", "P2-1")]["certificate_status"], "generated")
        self.assertEqual(values[("plant2", "P2-3")]["certificate_status"], "blank")
        first_pdf = folder / certificate_filename(read_batches(self.plant1)[0], 1)
        self.assertTrue(first_pdf.is_file())
        pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(first_pdf).pages)
        self.assertIn("PLANT ONE READY LTD", pdf_text)
        self.assertIn("FORM6-2", pdf_text)
        self.assertIn("REF-2", pdf_text)
        self.assertNotIn("email", " ".join(read_batches(self.plant1)[0].get("issues", [])).casefold())

        book = load_workbook(self.plant1, data_only=False)
        try:
            sheet = book["Inward Register"]
            self.assertEqual((sheet["K2"].value, sheet["M2"].value), ("0078", "Generated"))
            self.assertEqual((sheet["K3"].value, sheet["L3"].value.date(), sheet["M3"].value),
                             ("SPECIAL/REF-0077", date(2026, 9, 26), "Generated"))
        finally:
            book.close()

    def test_ready_is_the_only_generation_trigger_and_reissue_uses_current_excel_values(self):
        make_workbook(self.plant1, [{"sn": "EDIT-1", "age": 8, "cod": "USER-COD-1", "cod_date": date(2026, 9, 20)}])
        make_workbook(self.plant2, [{"sn": "OTHER-1", "age": 8, "status": "N/A"}])
        self.assertEqual(self.execute(check_only=True)[0], 0)
        self.assertFalse((self.base / "Certificates" / "USER-COD-1.pdf").exists())

        self.assertEqual(self.execute()[0], 1)
        first = self.base / "Certificates" / certificate_filename(read_batches(self.plant1)[0], 1)
        self.assertTrue(first.is_file())

        book = load_workbook(self.plant1)
        sheet = book["Inward Register"]
        sheet["C2"] = "Updated Client Name Ltd"
        sheet["K2"] = "NEW-CUSTOM-88"
        sheet["L2"] = date(2026, 9, 28)
        sheet["M2"] = "Ready"
        book.save(self.plant1)
        book.close()

        self.assertEqual(self.execute()[0], 1)
        rows = self.results()["rows"]
        row = next(item for item in rows if item["sn"] == "EDIT-1")
        self.assertEqual((row["certificate_no"], row["certificate_date"], row["certificate_status"]),
                         ("NEW-CUSTOM-88", "2026-09-28", "generated"))
        self.assertTrue(first.is_file(), "The earlier COD-named PDF is retained in the same folder.")
        updated_pdf = self.base / "Certificates" / certificate_filename(read_batches(self.plant1)[0], 1)
        self.assertTrue(updated_pdf.is_file())
        text = "\n".join(page.extract_text() or "" for page in PdfReader(updated_pdf).pages)
        self.assertIn("UPDATED CLIENT NAME LTD", text)

    def test_cleared_cod_is_reassigned_after_existing_certificate_numbers(self):
        make_workbook(self.plant1, [{"sn": "CLEAR-1", "age": 3, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "HIST-1", "age": 3, "status": "N/A", "cod": "0094"}])
        folder = self.base / "Certificates"
        folder.mkdir(parents=True)
        (folder / "0098.pdf").write_bytes(b"prior certificate")
        self.execute()
        self.assertEqual(read_batches(self.plant1)[0]["cod_no"], "0099")

    def test_unprefixed_custom_pdf_keeps_shared_sequence_floor(self):
        make_workbook(self.plant1, [{"sn": "FLAT-CUSTOM", "age": 3, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        folder = self.base / "Certificates"
        folder.mkdir(parents=True)
        (folder / "CUSTOM-REF-0198.pdf").write_bytes(b"prior certificate")

        self.assertEqual(self.execute()[0], 1)
        self.assertTrue((folder / certificate_filename(read_batches(self.plant1)[0], 1)).is_file())
        self.assertEqual(read_batches(self.plant1)[0]["cod_no"], "0199")

    def test_numeric_excel_cod_preserves_zero_padded_display_value(self):
        make_workbook(self.plant1, [{"sn": "NUM-COD", "age": 3}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        book = load_workbook(self.plant1)
        book["Inward Register"]["K2"] = 47
        book["Inward Register"]["K2"].number_format = "0000"
        book.save(self.plant1)
        book.close()
        self.assertEqual(read_batches(self.plant1)[0]["cod_no"], "0047")
        self.assertEqual(self.execute()[0], 1)
        self.assertTrue((self.base / "Certificates" / certificate_filename(read_batches(self.plant1)[0], 1)).is_file())

    def test_real_sample_workbook_copies_keep_content_and_gain_alert_rules(self):
        shutil.copy2(APP / "sample" / "Dummy_Inward_10_Clients.xlsx", self.plant1)
        shutil.copy2(APP / "sample" / "Dummy_Inward_5_Clients.xlsx", self.plant2)
        original = read_batches(self.plant1)
        _, _, _ = self.execute(check_only=True)
        self.assertTrue(self.results()["scan_complete"])
        updated = read_batches(self.plant1)
        self.assertEqual(len(updated), len(original))
        self.assertEqual([(item["case_id"], item["supplier"], item["reference"]) for item in updated],
                         [(item["case_id"], item["supplier"], item["reference"]) for item in original])
        book = load_workbook(self.plant1, data_only=False)
        try:
            sheet = book["Inward Register"]
            self.assertTrue(any("X2:X10000" in str(item.sqref) for item in sheet.data_validations.dataValidation))
            formulas = [" ".join(rule.formula or []) for region in sheet.conditional_formatting for rule in sheet.conditional_formatting[region]]
            self.assertTrue(any("TODAY()-INT($B2)>=25" in formula for formula in formulas))
            self.assertEqual(sheet.max_row, 25)
        finally:
            book.close()

    def test_selected_generation_processes_only_selected_ready_row(self):
        make_workbook(self.plant1, [{"sn": "PICK-1"}, {"sn": "PICK-2"}])
        make_workbook(self.plant2, [{"sn": "PICK-3"}])
        created, _, _ = self.execute(selected=("plant2", "PICK-3"))
        self.assertEqual(created, 1)
        self.assertEqual([path.name for path in (self.base / "Certificates").glob("*.pdf")],
                         [certificate_filename(read_batches(self.plant2)[0], 2)])
        self.assertEqual(read_batches(self.plant2)[0]["issued"], "Generated")
        self.assertEqual(read_batches(self.plant1)[0]["issued"], "Ready")

    def test_selected_generation_accepts_manual_cod_and_writes_it_back(self):
        make_workbook(self.plant1, [{"sn": "MANUAL-COD", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OTHER", "status": "N/A"}])

        created, review, _ = self.execute(
            selected=("plant1", "MANUAL-COD"), custom_cod="CUSTOM-0320",
        )

        self.assertEqual((created, review), (1, 0))
        batch = read_batches(self.plant1)[0]
        self.assertEqual((batch["cod_no"], batch["issued"]), ("CUSTOM-0320", "Generated"))
        self.assertTrue((self.base / "Certificates" / certificate_filename(batch, 1)).is_file())

    def test_manual_cod_cannot_replace_a_different_clients_number(self):
        make_workbook(self.plant1, [{"sn": "MANUAL-COLLISION", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "OWNS-COD", "cod": "CUSTOM-0320", "status": "N/A"}])

        created, review, _ = self.execute(
            selected=("plant1", "MANUAL-COLLISION"), custom_cod="CUSTOM-0320",
        )

        self.assertEqual((created, review), (0, 1))
        self.assertIn("already assigned to another plant row", self.results()["rows"][0]["detail"])
        self.assertEqual(read_batches(self.plant1)[0]["cod_no"], "")
        self.assertFalse(list((self.base / "Certificates").glob("*.pdf")))

    def test_duplicate_custom_cod_is_reported_and_never_overwrites_another_client(self):
        make_workbook(self.plant1, [{"sn": "DUP-1", "cod": "COD-X"}])
        make_workbook(self.plant2, [{"sn": "DUP-2", "cod": "COD-X"}])
        created, review, _ = self.execute()
        self.assertEqual((created, review), (0, 2))
        self.assertFalse(list((self.base / "Certificates").glob("*.pdf")))
        self.assertTrue(all("unique COD" in row["detail"] for row in self.results()["rows"]))

    def test_reference_is_optional_but_form6_warning_can_be_accepted(self):
        make_workbook(self.plant1, [{"sn": "MISSING-REF", "reference": "", "age": 27}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        created, review, _ = self.execute()
        self.assertEqual((created, review), (1, 0))
        row = next(item for item in self.results()["rows"] if item["sn"] == "MISSING-REF")
        self.assertNotIn("reference", row["detail"].casefold())
        self.assertEqual(read_batches(self.plant1)[0]["issued"], "Generated")

    def test_all_source_issues_are_reported_then_force_generation_is_possible(self):
        make_workbook(self.plant1, [{"sn": "WARN-1", "form6": "", "age": 27}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        self.assertEqual(self.execute(), (0, 1, self.base / "Data" / "Reports"))
        self.assertEqual(self.execute(allow_warnings=True)[0], 1)
        row = next(item for item in self.results()["rows"] if item["sn"] == "WARN-1")
        self.assertTrue(row["system_alert"])
        self.assertTrue(row["warnings"])
        self.assertIn("missing Form 6", row["detail"])
        self.assertEqual(read_batches(self.plant1)[0]["issued"], "Generated")

    def test_certificate_weight_column_replaces_inward_total(self):
        make_workbook(self.plant1, [{"sn": "NET-1", "weight": 10.0, "age": 2}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        book = load_workbook(self.plant1)
        sheet = book["Inward Register"]
        sheet["N1"] = "Certificate Weight (KG)"
        sheet["N2"] = 8.5
        book.save(self.plant1)
        book.close()
        batch = read_batches(self.plant1)[0]
        self.assertEqual((batch["display_weight"], batch["weight_source"]), (8.5, "certificate"))
        self.assertEqual(self.execute()[0], 1)
        pdf = self.base / "Certificates" / certificate_filename(read_batches(self.plant1)[0], 1)
        pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(pdf).pages)
        self.assertIn("8.5 KG", pdf_text)

    def test_file_detail_upsert_appends_row_and_can_create_certificate(self):
        make_workbook(self.plant1, [{"sn": "EXISTING", "status": "N/A"}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        batch = upsert_certificate_details(
            self.plant1, self.base / "Data" / "Workbook Backups", {
                "case_id": "FROM-FORM", "inward_date": "01.09.2026",
                "supplier": "Entered From App Ltd", "location": "Pune",
                "item": "Servers", "quantity": "2", "per_piece_weight": "5",
                "weight": "10", "certificate_weight": "8", "form6": "FORM6-FORM",
                "reference": "", "cod_no": "", "cod_date": "",
            },
        )
        self.assertEqual((batch["case_id"], batch["issued"], batch["form6"]),
                         ("FROM-FORM", "Ready", "FORM6-FORM"))
        self.assertEqual(self.execute(selected=("plant1", "FROM-FORM"))[0], 1)
        generated = read_batches(self.plant1)[-1]
        self.assertTrue((self.base / "Certificates" / certificate_filename(generated, 1)).is_file())

    def test_selected_certificate_overrides_are_written_to_excel_and_pdf(self):
        make_workbook(self.plant1, [{"sn": "OVERRIDE-1", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "N-A", "status": "N/A"}])
        self.assertEqual(run(
            self.paths, self.base, as_of=AS_OF, selected=("plant1", "OVERRIDE-1"),
            overrides={
                "supplier": "Replacement Company", "location": "Surat",
                "reference": "REF-REPLACED", "certificate_weight": "5.25",
            }, allow_warnings=True, progress=lambda _: None,
        )[0], 1)
        batch = read_batches(self.plant1)[0]
        self.assertEqual((batch["supplier"], batch["location"], batch["reference"], batch["certificate_weight"]),
                         ("Replacement Company", "Surat", "REF-REPLACED", 5.25))
        pdf = self.base / "Certificates" / certificate_filename(batch, 1)
        pdf_text = "\n".join(page.extract_text() or "" for page in PdfReader(pdf).pages)
        self.assertIn("REPLACEMENT COMPANY, SURAT", pdf_text)
        self.assertIn("5.2 KG", pdf_text)


class ReminderAndExcelTests(AutomationTestBase):
    def test_reminder_exact_day_25_through_day_30_then_overdue_until_complete(self):
        batch = {"case_id": "A", "inward_date": AS_OF - timedelta(days=24), "issued": "Blank"}
        self.assertFalse(certificate_reminder_due(batch, AS_OF))
        batch["inward_date"] = parse_date((AS_OF - timedelta(days=24)).strftime("%d.%m.%Y"))
        self.assertFalse(certificate_reminder_due(batch, AS_OF))
        batch["inward_date"] = AS_OF - timedelta(days=25)
        self.assertTrue(certificate_reminder_due(batch, AS_OF))
        batch["inward_date"] = parse_date((AS_OF - timedelta(days=25)).strftime("%d.%m.%Y"))
        self.assertTrue(certificate_reminder_due(batch, AS_OF))
        report = Report(self.base, AS_OF, "test")
        report.add(batch, "needs_review", "pending")
        self.assertTrue(report.rows[-1]["reminder_due"])
        self.assertFalse(report.rows[-1]["overdue"])
        batch["inward_date"] = AS_OF - timedelta(days=30)
        report.add(batch, "needs_review", "pending")
        self.assertTrue(report.rows[-1]["reminder_due"])
        self.assertFalse(report.rows[-1]["overdue"])
        batch["inward_date"] = AS_OF - timedelta(days=31)
        report.add(batch, "needs_review", "pending")
        self.assertTrue(report.rows[-1]["reminder_due"])
        self.assertTrue(report.rows[-1]["overdue"])
        for status in ("Completed", "N/A", "Sent"):
            self.assertFalse(certificate_reminder_due(dict(batch, issued=status), AS_OF))

    def test_legacy_labels_normalize_to_canonical_values_and_alert_never_becomes_ready(self):
        for source, expected in (("Created", "generated"), ("Sent", "completed"),
                                 ("Check Details", "alert"), ("Working Draft", "alert"),
                                 ("Generated Manual", "generated_manual")):
            with self.subTest(source=source):
                self.assertEqual(status_value(source), expected)

    def test_excel_refresh_migrates_legacy_statuses_but_keeps_blank_and_na_distinct(self):
        make_workbook(self.plant1, [
            {"sn": "CREATED-1", "status": "Created"},
            {"sn": "SENT-1", "status": "Sent"},
            {"sn": "DETAILS-1", "status": "Check Details"},
            {"sn": "DRAFT-1", "status": "Working Draft"},
            {"sn": "MANUAL-1", "status": "Generated Manual"},
            {"sn": "BLANK-1", "status": ""},
            {"sn": "NA-1", "status": "N/A"},
        ])
        with WorkbookSession(self.plant1, self.base / "Data" / "Workbook Backups") as session:
            session.ensure_alert_formatting("Inward Register")
        book = load_workbook(self.plant1, data_only=False)
        try:
            sheet = book["Inward Register"]
            self.assertEqual([sheet.cell(row, 13).value for row in range(2, 9)],
                             ["Generated", "Completed", "Alert", "Alert", "Generated Manual", None, "N/A"])
            validation = next(item for item in sheet.data_validations.dataValidation if "M2:M10000" in str(item.sqref))
            self.assertIn("Generated Manual", validation.formula1)
            self.assertNotIn("Working Draft", validation.formula1)
            self.assertNotIn("Created", validation.formula1)
            self.assertNotIn("Sent", validation.formula1)
        finally:
            book.close()

    def test_alerts_are_reported_for_ready_generated_and_blank_but_not_completed_or_na(self):
        due = AS_OF - timedelta(days=25)
        for status, expected in (("Ready", True), ("Generated", True), ("Blank", True),
                                 ("Completed", False), ("N/A", False)):
            with self.subTest(status=status):
                report = Report(self.base, AS_OF, "test")
                report.add({"case_id": "ID-1", "supplier": "Client", "inward_date": due, "issued": status}, "state", "")
                self.assertEqual(report.rows[0]["reminder_due"], expected)

    def test_excel_alert_conditional_formatting_dropdown_and_cod_writeback(self):
        make_workbook(self.plant1, [{"sn": "RED-1", "age": 25, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "RED-2", "age": 31, "status": "N/A"}])
        # Reproduce the old workbook's duplicate status-color rules. The app
        # must replace these instead of layering another set over them.
        book = load_workbook(self.plant1)
        sheet = book["Inward Register"]
        legacy_fill = PatternFill(fill_type="solid", fgColor="FF000000")
        for value in ("READY", "ALERT", "N/A", "BLANK", "GENERATED", "GENERATED MANUAL", "COMPLETED"):
            sheet.conditional_formatting.add(
                "M2:M1000",
                FormulaRule(formula=[f'UPPER(TRIM($M2))="{value}"'], fill=legacy_fill),
            )
        legacy_alert = (
            'AND($A2<>"",OR(NOT(ISNUMBER($B2)),$B2>TODAY(),TODAY()-$B2>=25),'
            'UPPER(TRIM($M2))<>"COMPLETED",'
            'SUBSTITUTE(UPPER(TRIM($M2))," ","")<>"N/A")'
        )
        sheet.conditional_formatting.add(
            "C2:C1000 A2:A1000 J2:J1000",
            FormulaRule(formula=[legacy_alert], fill=legacy_fill),
        )
        sheet.conditional_formatting.add(
            "A2:A1000",
            FormulaRule(formula=['$A2="KEEP THIS RULE"'], fill=PatternFill(fill_type="solid", fgColor="FF4472C4")),
        )
        book.save(self.plant1)
        book.close()

        self.execute()
        self.execute(check_only=True)  # A second pass must not layer duplicate rules.
        book = load_workbook(self.plant1, data_only=False)
        try:
            sheet = book["Inward Register"]
            validations = [item for item in sheet.data_validations.dataValidation if "M2:M10000" in str(item.sqref)]
            self.assertEqual(len(validations), 1)
            self.assertIn("Ready,Alert,N/A,Blank,Generated,Generated Manual,Completed", validations[0].formula1)
            matching = []
            status_colors = {}
            status_formulas = {}
            status_rule_count = 0
            priorities = []
            unrelated_rule_kept = False
            for region in sheet.conditional_formatting:
                for rule in sheet.conditional_formatting[region]:
                    formula = " ".join(rule.formula or [])
                    priorities.append(rule.priority)
                    unrelated_rule_kept |= formula == '$A2="KEEP THIS RULE"'
                    if "TODAY()-INT($B2)>=25" in formula:
                        matching.append((str(region.sqref), rule))
                    if ("M2:M10000" in str(region.sqref) and "$M2" in formula
                            and any(f'"{value}"' in formula for value in
                                    ("READY", "ALERT", "N/A", "BLANK", "GENERATED", "GENERATED MANUAL", "COMPLETED"))):
                        status_rule_count += 1
                        for label in ("READY", "ALERT", "N/A", "BLANK", "GENERATED", "GENERATED MANUAL", "COMPLETED"):
                            if f'"{label}"' in formula or (label == "N/A" and '"N/A"' in formula):
                                dxf = book._differential_styles.styles[rule.dxfId]
                                status_colors[label] = dxf.fill.fgColor.rgb
                                status_formulas[label] = formula
                                self.assertEqual(dxf.fill.fgColor.rgb, dxf.fill.bgColor.rgb)
                                self.assertEqual(dxf.fill.bgColor.type, "rgb")
            self.assertEqual(len(matching), 1)
            self.assertEqual(status_rule_count, 7, "Only one authoritative color rule per status should remain.")
            self.assertEqual(len(priorities), len(set(priorities)), "Excel rule priorities must be unique.")
            self.assertTrue(unrelated_rule_kept, "Formatting in unrelated user-owned cells must be preserved.")
            self.assertIn('AND($A2<>"",OR($M2="",UPPER(TRIM($M2))="BLANK"))', status_formulas["BLANK"])
            self.assertEqual(matching[0][0], "C2:C10000")
            self.assertNotIn("A2:A10000", matching[0][0])
            self.assertNotIn("J2:J10000", matching[0][0])
            rule = matching[0][1]
            self.assertIn("IF(ISNUMBER($B2)", " ".join(rule.formula or []))
            self.assertIn("DATE(VALUE(RIGHT(TRIM($B2),4))", " ".join(rule.formula or []))
            self.assertIn('MID(TRIM($B2),3,1)="."', " ".join(rule.formula or []))
            self.assertIn('MID(TRIM($B2),3,1)="/"', " ".join(rule.formula or []))
            self.assertIn('MID(TRIM($B2),3,1)="-"', " ".join(rule.formula or []))
            self.assertIn("OR(INT($B2)>TODAY()", " ".join(rule.formula or []))
            self.assertNotIn('UPPER(TRIM($M2))<>"SENT"', " ".join(rule.formula or []))
            self.assertNotIn("NOT(ISNUMBER($B2))", " ".join(rule.formula or []))
            dxf = book._differential_styles.styles[rule.dxfId]
            self.assertEqual(dxf.fill.fgColor.rgb, "FFC00000")
            self.assertEqual(dxf.fill.fgColor.rgb, dxf.fill.bgColor.rgb)
            self.assertEqual(dxf.font.color.rgb, "FFFFFFFF")
            self.assertEqual(status_colors["READY"], "FFB6D7A8")
            self.assertEqual(status_colors["ALERT"], "FFFCE4D6")
            self.assertEqual(status_colors["N/A"], "FFFFFFFF")
            self.assertEqual(status_colors["BLANK"], "FFC00000")
            self.assertEqual(status_colors["GENERATED"], "FFFFE699")
            self.assertEqual(status_colors["GENERATED MANUAL"], "FFDDEBF7")
            self.assertEqual(status_colors["COMPLETED"], "FFFFE699")
            self.assertEqual(sheet["M2"].value, "Generated")
        finally:
            book.close()
        backups = list((self.base / "Data" / "Workbook Backups").glob("Plant 1_*.xlsx"))
        self.assertTrue(backups, "Workbook formatting and writeback must preserve a recoverable backup.")

    def test_alert_rule_parses_text_day_first_dates_for_day_25_boundary(self):
        due_date = (AS_OF - timedelta(days=25)).strftime("%d.%m.%Y")
        recent_date = (AS_OF - timedelta(days=24)).strftime("%d.%m.%Y")
        make_workbook(self.plant1, [{"sn": "TEXT-DUE", "date": due_date, "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "TEXT-RECENT", "date": recent_date, "status": "Ready"}])

        self.execute(check_only=True)

        for path in self.paths:
            book = load_workbook(path, data_only=False)
            try:
                sheet = book["Inward Register"]
                matching = [
                    " ".join(rule.formula or [])
                    for region in sheet.conditional_formatting
                    for rule in sheet.conditional_formatting[region]
                    if "TODAY()-DATE(VALUE(RIGHT(TRIM($B2),4))" in " ".join(rule.formula or [])
                ]
                self.assertEqual(len(matching), 1)
                formula = matching[0]
                self.assertIn("TODAY()-DATE(VALUE(RIGHT(TRIM($B2),4))", formula)
                self.assertIn("DATE(VALUE(RIGHT(TRIM($B2),4))", formula)
                self.assertIn(">TODAY()", formula)
                self.assertIn("TODAY()-INT($B2)>=25", formula)
            finally:
                book.close()

        parsed_due = read_batches(self.plant1)[0]["inward_date"]
        parsed_recent = read_batches(self.plant2)[0]["inward_date"]
        self.assertEqual((AS_OF - parsed_due).days, 25)
        self.assertEqual((AS_OF - parsed_recent).days, 24)

    def test_blank_rule_refresh_saves_even_when_alert_rule_is_already_current(self):
        make_workbook(self.plant1, [{"sn": "BLANK-ROW", "status": "Ready"}])
        make_workbook(self.plant2, [{"sn": "NA-ROW", "status": "N/A"}])
        self.execute(check_only=True)

        previous_formula = '<formula>AND($A2&lt;&gt;"",OR($M2="",UPPER(TRIM($M2))="BLANK"))</formula>'
        legacy_formula = '<formula>OR($M2="",UPPER(TRIM($M2))="BLANK")</formula>'
        with ZipFile(self.plant1) as source:
            member = _sheet_member(source, "Inward Register")
            xml = source.read(member).decode("utf-8")
            self.assertIn(previous_formula, xml)
            buffer = BytesIO()
            with ZipFile(buffer, "w") as target:
                target.comment = source.comment
                for info in source.infolist():
                    content = source.read(info.filename)
                    if info.filename == member:
                        content = content.decode("utf-8").replace(previous_formula, legacy_formula, 1).encode("utf-8")
                    target.writestr(info, content)
        self.plant1.write_bytes(buffer.getvalue())

        before = read_batches(self.plant1)
        with WorkbookSession(self.plant1, self.base / "Data" / "Workbook Backups") as session:
            session.ensure_alert_formatting("Inward Register")
        after = read_batches(self.plant1)
        self.assertEqual([batch_content(row) for row in before], [batch_content(row) for row in after])
        with ZipFile(self.plant1) as repaired:
            xml = repaired.read(_sheet_member(repaired, "Inward Register")).decode("utf-8")
        self.assertIn(previous_formula, xml)
        self.assertNotIn(legacy_formula, xml)

    def test_reminder_default_and_time_picker_values(self):
        self.assertEqual(load_settings(self.base)["reminder_time"], "11:00")
        self.assertEqual(reminder_picker_parts("11:00"), ("11", "00", "AM"))
        self.assertEqual(reminder_time_from_picker("3", "25", "PM"), "15:25")
        self.assertEqual(reminder_time_from_picker("12", "00", "AM"), "00:00")
        self.assertEqual(reminder_time_from_picker("12", "00", "PM"), "12:00")

    def test_save_and_update_reminder_passes_time_as_one_option_pair(self):
        app = App.__new__(App)
        app.save_setup = Mock(return_value=True)
        app.reminders_enabled = Mock()
        app.reminders_enabled.get.return_value = True
        app.settings = {"reminder_time": "14:16"}
        app.selected_paths = Mock()
        app.status = Mock()
        app._time_label = Mock(return_value="2:16 PM")
        app._powershell = Mock(return_value=Mock(returncode=0, stdout="", stderr=""))

        with patch("desktop_app.messagebox.showerror") as show_error:
            app.save_setup_and_schedule()

        app._powershell.assert_called_once_with(
            "Install", (("-ReminderTime", "14:16"),)
        )
        show_error.assert_not_called()
        self.assertIn("2:16 PM", app.status.set.call_args.args[0])

    def test_powershell_command_builder_adds_reminder_time_without_running_it(self):
        app = App.__new__(App)
        app.base = self.base
        completed = Mock(returncode=0, stdout="", stderr="")

        with patch("desktop_app.subprocess.run", return_value=completed) as run_process:
            result = app._powershell("Install", (("-ReminderTime", "14:16"),))

        self.assertIs(result, completed)
        command = run_process.call_args.args[0]
        self.assertEqual(command[-4:], ["-Mode", "Install", "-ReminderTime", "14:16"])
        run_process.assert_called_once_with(
            command, capture_output=True, text=True, timeout=45, check=False
        )

    def test_daily_summary_is_centered_actionable_and_snooze_can_repeat_same_day(self):
        report = Report(self.base, AS_OF, "test")
        report.add({"case_id": "25", "supplier": "Day 25 Ltd", "inward_date": AS_OF - timedelta(days=25), "issued": "Generated", "source_id": "plant1"}, "generated", "")
        result = {"run_id": "a", "rows": report.rows}
        notices = []
        self.assertTrue(notify_once(self.base, result, lambda *args: notices.append(args), today=AS_OF))
        self.assertIn("day 25–30", notices[0][1])
        self.assertFalse(notify_once(self.base, result, lambda *args: notices.append(args), today=AS_OF))
        self.assertTrue(notify_once(self.base, result, lambda *args: notices.append(args), today=AS_OF, repeat_pending=True))
        self.assertEqual(len(notices), 2)

    def test_reminder_opens_the_plant_workbook_for_its_first_actionable_client(self):
        first_workbook = self.base / "Plant 1.xlsx"
        second_workbook = self.base / "Plant 2.xlsx"
        first_workbook.write_bytes(b"workbook")
        second_workbook.write_bytes(b"workbook")
        result = {"rows": [
            {"file": "", "reminder_due": False, "system_alert": True},
            {"file": str(first_workbook), "reminder_due": True, "sn": "1001"},
            {"file": str(second_workbook), "reminder_due": True, "sn": "2001"},
        ]}

        with patch("workflow_support.os.startfile", create=True) as startfile:
            opened = open_alert_workbook(result)

        self.assertEqual(opened, first_workbook)
        startfile.assert_called_once_with(str(first_workbook))

    def test_reminder_reports_when_no_linked_plant_workbook_is_available(self):
        with self.assertRaisesRegex(FileNotFoundError, "could not be found"):
            open_alert_workbook({"rows": [{
                "file": str(self.base / "Missing Plant.xlsx"), "reminder_due": True,
            }]})

    def test_failed_notification_is_not_recorded_and_can_retry(self):
        report = Report(self.base, AS_OF, "test")
        report.add({"case_id": "26", "supplier": "Retry Ltd", "inward_date": AS_OF - timedelta(days=26), "issued": "Ready", "source_id": "plant1"}, "ready", "")
        result = {"run_id": "retry", "rows": report.rows}

        def fail_notice(*_args):
            raise OSError("simulated desktop notification failure")

        with self.assertRaisesRegex(OSError, "notification failure"):
            notify_once(self.base, result, fail_notice, today=AS_OF)
        self.assertFalse((self.base / "Data" / "reminder-state.json").exists())

        notices = []
        self.assertTrue(notify_once(self.base, result, lambda *args: notices.append(args), today=AS_OF))
        self.assertEqual(len(notices), 1)
        self.assertTrue((self.base / "Data" / "reminder-state.json").exists())

    def test_pdf_only_runtime_has_no_outlook_or_com_dependencies(self):
        requirements = (APP / "requirements.txt").read_text(encoding="utf-8").casefold()
        scheduled = (APP / "scheduled_run.py").read_text(encoding="utf-8").casefold()
        self.assertNotIn("pywin32", requirements)
        self.assertNotIn("win32com", scheduled)
        self.assertNotIn("pythoncom", scheduled)
        self.assertFalse((APP / "draft_service.py").exists())
        self.assertFalse((APP / "certificate_correction.py").exists())
        self.assertFalse((APP.parent / "Run Daily Drafts.bat").exists())
        self.assertFalse((APP.parent / "scheduled_run.py").exists())


if __name__ == "__main__":
    unittest.main()
