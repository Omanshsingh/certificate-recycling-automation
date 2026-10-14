"""Transactional certificate-form edit scopes, using disposable workspaces only."""
from __future__ import annotations

from copy import deepcopy
from datetime import date
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from pypdf import PdfReader

from certificate_core import parse_date, read_batches
from certificate_drafts import clean_details
import run_certificates as runner
import web_app
from workbook_store import WorkbookSession, stage_certificate_details
from workflow_support import save_json
from tests.test_certificate_templates import write_register


APP = Path(__file__).resolve().parents[1]
AS_OF = date(2026, 10, 1)
ISSUED_DATE = date(2026, 9, 18)


class CertificateEditScopeTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        write_register(self.paths[0], continuation=True)
        write_register(self.paths[1], case_id="OTHER", cod="0040", status="N/A")
        book = load_workbook(self.paths[0])
        try:
            sheet = book.active
            # This suite exercises corrections to an already issued inward.
            # First-generation locks are covered separately in test_simple_certificate_flow.
            sheet.cell(2, 13, 'Generated')
            for column, header, value in (
                (14, "Certificate Weight (KG)", 20),
                (15, "Certificate Type", "standard"),
                (16, "Certificate Fields", json.dumps({"company_display": "Original Client"})),
            ):
                sheet.cell(1, column, header)
                sheet.cell(2, column, value)
            book.save(self.paths[0])
        finally:
            book.close()
        native = self.root / "App" / "templates" / "certificate-types"
        native.mkdir(parents=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf", native / "standard.pdf")
        self.folder = self.root / "Certificates"
        self.backups = self.root / "Data" / "Workbook Backups"
        self.results = self.root / "Data" / "Reports" / "Latest Results.json"
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "certificate_number_prefix": "", "certificate_number_start": 1,
            "certificate_number_width": 4, "reminder_time": "11:00", "reminders_enabled": False,
        })
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "BACKUP_DIR", self.backups))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", self.results))

    def details(self, **changes):
        values = {
            "case_id": "CLIENT", "supplier": "Edited Client Ltd", "location": "Surat",
            "inward_date": "05.09.2026", "item": "Laptops", "quantity": "3",
            "per_piece_weight": "7", "weight": "21", "certificate_weight": "9.5",
            "form6": "F-EDIT", "reference": "R-EDIT", "cod_no": "", "cod_date": "18.09.2026",
            "certificate_type": "standard", "certificate_fields": {"company_display": "New Client"},
        }
        values.update(changes)
        return clean_details(values)

    def batch(self, case_id="CLIENT"):
        return next(batch for batch in read_batches(self.paths[0]) if batch["case_id"] == case_id)

    def workbook_bytes(self):
        return [path.read_bytes() for path in self.paths]

    def cells(self):
        book = load_workbook(self.paths[0])
        try:
            return [tuple(cell.value for cell in row) for row in book.active.iter_rows()]
        finally:
            book.close()

    def assert_untouched(self, before):
        self.assertEqual(self.workbook_bytes(), before)
        self.assertFalse(list(self.backups.glob("*.xlsx")))
        self.assertFalse(list(self.folder.rglob("*.pdf")))

    def run_form(self, details=None, *, update_excel=True, **options):
        details = self.details() if details is None else details
        return runner.run(self.paths, self.root, selected=("plant1", details["case_id"]),
                          details=details, update_excel=update_excel, as_of=AS_OF,
                          progress=None, **options)

    def body(self, scope="update_excel", details=None, **extra):
        details = self.details() if details is None else details
        body = dict(plant="plant1", details=details, edit_scope=scope,
                    use_certificate_weight=True)
        if details['case_id'] == 'CLIENT':
            body.update(source_sn='CLIENT', source_plant='plant1')
        body.update(extra)
        return body

    def report(self):
        return json.loads(self.results.read_text(encoding="utf-8"))

    def assert_issued(self, cod="0041"):
        batch = self.batch()
        self.assertEqual((batch["cod_no"], parse_date(batch["cod_date"]), batch["issued"]),
                         (cod, ISSUED_DATE, "Generated"))
        self.assertTrue(list(self.backups.glob("*.xlsx")))
        return batch

    def merge_shared_details(self, *extra):
        book = load_workbook(self.paths[0])
        try:
            for column in ('A', 'B', 'C', 'D', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', *extra):
                book.active.merge_cells(f'{column}2:{column}3')
            book.save(self.paths[0])
            return {str(area) for area in book.active.merged_cells.ranges}
        finally:
            book.close()

    def test_preview_accepts_vertical_shared_details_without_changing_excel(self):
        self.merge_shared_details()
        before = self.workbook_bytes()
        result = web_app.preview_certificate(self.body())
        self.assertTrue(result['preview_url'])
        self.assert_untouched(before)

    def test_excel_correction_preserves_shared_merges_and_both_material_rows(self):
        merges = self.merge_shared_details()
        continuation = deepcopy(self.batch()['items'][1:])
        created, review, _ = self.run_form()
        self.assertEqual((created, review), (1, 0), self.report())
        batch = self.assert_issued()
        self.assertEqual(batch['supplier'], 'Edited Client Ltd')
        self.assertEqual(batch['reference'], 'R-EDIT')
        self.assertEqual(batch['items'][1:], continuation)
        book = load_workbook(self.paths[0])
        try:
            self.assertEqual({str(area) for area in book.active.merged_cells.ranges}, merges)
            self.assertIsNone(book.active['A3'].value)
            self.assertIsNone(book.active['C3'].value)
        finally:
            book.close()

    def test_pdf_only_correction_keeps_merged_source_details(self):
        merges = self.merge_shared_details()
        source = self.batch()
        created, review, _ = self.run_form(update_excel=False)
        self.assertEqual((created, review), (1, 0), self.report())
        batch = self.assert_issued()
        self.assertEqual(batch['supplier'], source['supplier'])
        self.assertEqual(batch['items'], source['items'])
        book = load_workbook(self.paths[0])
        try:
            self.assertEqual({str(area) for area in book.active.merged_cells.ranges}, merges)
        finally:
            book.close()
        pdf = next(self.folder.rglob('*.pdf'))
        self.assertIn('New Client', '\n'.join(page.extract_text() or '' for page in PdfReader(pdf).pages))

    def test_first_issuance_prepare_accepts_vertical_output_merges(self):
        self.merge_shared_details()
        book = load_workbook(self.paths[0])
        book.active['M2'] = 'Ready'
        book.save(self.paths[0])
        book.close()
        before = self.workbook_bytes()
        with WorkbookSession(self.paths[0], self.backups) as session:
            prepared = session.prepare(self.batch_from_session(session))
        self.assertEqual(prepared[1:4], ('K2', 'L2', 'M2'))
        self.assert_untouched(before)

    def batch_from_session(self, session):
        from io import BytesIO
        return next(batch for batch in read_batches(BytesIO(session.data)) if batch['case_id'] == 'CLIENT')

    def test_completion_status_updates_vertical_merge_anchor_only(self):
        merges = self.merge_shared_details()
        with WorkbookSession(self.paths[0], self.backups) as session:
            session.set_issued_status('CLIENT', 'Inward Register', 2, '', 'Generated', 'Completed')
        self.assertEqual(self.batch()['issued'], 'Completed')
        book = load_workbook(self.paths[0])
        try:
            self.assertEqual({str(area) for area in book.active.merged_cells.ranges}, merges)
            self.assertIsNone(book.active['M3'].value)
        finally:
            book.close()

    def test_cross_inward_shared_cell_cannot_be_saved_to_excel(self):
        book = load_workbook(self.paths[0])
        try:
            book.active.append(['NEXT', date(2026, 9, 8), 'Next Client', 'Pune', 'Printers', 1, 2, 2,
                                'FORM-NEXT', 'REF-NEXT', None, None, 'Ready'])
            book.active.merge_cells('D2:D4')
            book.save(self.paths[0])
        finally:
            book.close()
        before = self.workbook_bytes()
        with WorkbookSession(self.paths[0], self.backups) as session:
            with self.assertRaisesRegex(ValueError, 'shared across columns or inward'):
                stage_certificate_details(session, self.details())
        self.assert_untouched(before)

    def test_merged_material_lines_allow_pdf_preview_but_block_excel_save(self):
        self.merge_shared_details('E')
        before = self.workbook_bytes()
        details = self.details(items=[{'item': 'Laptops', 'weight': '21'},
                                     {'item': 'Monitors', 'weight': '12'}])
        result = web_app.preview_certificate(self.body('certificate_only', details))
        self.assertTrue(result['preview_url'])
        with WorkbookSession(self.paths[0], self.backups) as session:
            with self.assertRaisesRegex(ValueError, 'material lines share merged'):
                stage_certificate_details(session, details)
        self.assert_untouched(before)

    def test_merged_formula_and_protected_anchors_still_cannot_be_overwritten(self):
        self.merge_shared_details()
        for kind in ('formula', 'protected'):
            with self.subTest(kind=kind):
                book = load_workbook(self.paths[0])
                try:
                    book.active['C2'] = '=A2' if kind == 'formula' else 'Original Client'
                    book.active.protection.sheet = kind == 'protected'
                    book.save(self.paths[0])
                finally:
                    book.close()
                before = self.workbook_bytes()
                with WorkbookSession(self.paths[0], self.backups) as session:
                    with self.assertRaisesRegex(ValueError, 'formula or protected'):
                        stage_certificate_details(session, self.details())
                self.assert_untouched(before)

    def test_repeated_preview_and_regeneration_preserve_merged_inward(self):
        for scope in ('certificate_only', 'update_excel'):
            with self.subTest(scope=scope):
                merges = self.merge_shared_details()
                original = deepcopy(self.batch())
                other = self.paths[1].read_bytes()
                for attempt in range(3):
                    details = web_app.editor_details(self.batch())
                    details['supplier'] = f'Correction {scope} {attempt}'
                    details['reference'] = f'CORRECT-{scope}-{attempt}'
                    details['certificate_fields'] = {'company_display': details['supplier']}
                    body = self.body(scope, details, confirmed=True,
                                     override_reason='Repeated edit regression test')
                    before_preview = self.workbook_bytes()
                    preview = web_app.preview_certificate(body)
                    self.assertTrue(preview['preview_url'])
                    self.assertEqual(self.workbook_bytes(), before_preview)
                    with patch.object(web_app, 'refresh', side_effect=lambda *_: self.report()):
                        result = web_app.create_certificate(body)
                    self.assertEqual(result['created'], 1, result)
                    batch = self.batch()
                    self.assertEqual(batch['items'], original['items'])
                    self.assertEqual(batch['supplier'], details['supplier'] if scope == 'update_excel'
                                     else original['supplier'])
                    self.assertEqual(self.paths[1].read_bytes(), other)
                    book = load_workbook(self.paths[0])
                    try:
                        self.assertEqual({str(area) for area in book.active.merged_cells.ranges}, merges)
                        self.assertIsNone(book.active['A3'].value)
                    finally:
                        book.close()

    def test_first_generation_then_edit_and_regenerate_with_shared_merges(self):
        merges = self.merge_shared_details()
        book = load_workbook(self.paths[0])
        try:
            book.active['M2'] = 'Ready'
            book.save(self.paths[0])
        finally:
            book.close()
        for attempt in range(3):
            source = self.batch()
            details = web_app.editor_details(source)
            if attempt:
                details['reference'] = f'AFTER-FIRST-{attempt}'
            body = self.body('update_excel', details, confirmed=True,
                             override_reason='First generation and repeated edit regression')
            before_preview = self.workbook_bytes()
            self.assertTrue(web_app.preview_certificate(body)['preview_url'])
            self.assertEqual(self.workbook_bytes(), before_preview)
            with patch.object(web_app, 'refresh', side_effect=lambda *_: self.report()):
                result = web_app.create_certificate(body)
            self.assertEqual(result['created'], 1, result)
            self.assertEqual(self.batch()['items'], source['items'])
            if attempt:
                self.assertEqual(self.batch()['reference'], f'AFTER-FIRST-{attempt}')
            book = load_workbook(self.paths[0])
            try:
                self.assertEqual({str(area) for area in book.active.merged_cells.ranges}, merges)
            finally:
                book.close()

    def test_check_only_stages_every_edit_without_any_workbook_byte_change(self):
        for update_excel in (False, True):
            with self.subTest(update_excel=update_excel):
                before = self.workbook_bytes()
                with patch.object(runner, "make_certificate") as render:
                    created, review, _ = self.run_form(update_excel=update_excel, check_only=True)
                self.assertEqual((created, review), (0, 0), self.report())
                render.assert_not_called()
                self.assert_untouched(before)
                row = web_app.target_row(self.report(), ("plant1", "CLIENT"))
                self.assertEqual(row["outcome"], "ready")
                self.assertEqual(row["certificate_fields"], {"company_display": "New Client"})

    def test_warning_confirmation_and_cancel_never_refresh_or_save_excel(self):
        for scope in ("certificate_only", "update_excel"):
            with self.subTest(scope=scope):
                before = self.workbook_bytes()
                body = self.body(scope, self.details(inward_date=""))
                with patch.object(web_app, "refresh", side_effect=AssertionError("Warning return must not refresh")), \
                        patch.object(runner, "make_certificate") as render:
                    for _ in range(2):  # Dismissing/reopening confirmation is still read-only.
                        result = web_app.create_certificate(body)
                        self.assertTrue(result["requires_confirmation"])
                        self.assertEqual(result["created"], 0)
                        self.assertNotIn("state", result)
                        self.assertTrue(result["warnings"])
                render.assert_not_called()
                self.assert_untouched(before)

    def test_preview_uses_final_staged_weight_and_preserves_excel(self):
        for net_weight, expected in (("", 20), ("9.5", 9.5)):
            with self.subTest(net_weight=net_weight):
                before = self.workbook_bytes()
                with patch.object(web_app, "make_template_certificate", wraps=web_app.make_template_certificate) as render:
                    result = web_app.preview_certificate(self.body(
                        "certificate_only", self.details(certificate_weight=net_weight),
                        use_certificate_weight=bool(net_weight)))
                batch = render.call_args.args[0]
                self.assertEqual(batch["total_weight"], 33)
                self.assertEqual(batch["display_weight"], expected)
                self.assertEqual(batch["supplier"], "Edited Client Ltd")
                self.assertEqual(result["certificate_no"], "0041")
                self.assert_untouched(before)

    def test_unchecked_certificate_weight_uses_and_preserves_excel_weight(self):
        before = self.workbook_bytes()
        details = self.details(certificate_weight="8.5")
        body = self.body(details=details, use_certificate_weight=False)
        with patch.object(web_app, "make_template_certificate", wraps=web_app.make_template_certificate) as render:
            web_app.preview_certificate(body)
        self.assertEqual(render.call_args.args[0]["certificate_weight"], 20)
        self.assertEqual(render.call_args.args[0]["display_weight"], 20)
        self.assert_untouched(before)

        with patch.object(web_app, "refresh", side_effect=lambda *_: self.report()):
            result = web_app.create_certificate(body)
        self.assertEqual(result["created"], 1)
        batch = self.assert_issued()
        self.assertEqual(batch["certificate_weight"], 20)
        self.assertEqual(batch["display_weight"], 20)
        self.assertEqual(self.cells()[1][13], 20)

    def test_checked_certificate_weight_requires_a_positive_value(self):
        with self.assertRaisesRegex(ValueError, "positive Certificate Weight"):
            web_app.preview_certificate(self.body(
                details=self.details(certificate_weight=""), use_certificate_weight=True))
        self.assertFalse(list(self.folder.glob("*.pdf")))

    def test_preview_does_not_keep_source_warnings_after_form_corrections(self):
        book = load_workbook(self.paths[0])
        try:
            book.active["I2"] = None
            book.save(self.paths[0])
        finally:
            book.close()
        before = self.workbook_bytes()
        with patch.object(web_app, "make_template_certificate", wraps=web_app.make_template_certificate):
            result = web_app.preview_certificate(self.body(details=self.details(form6="F-CORRECTED")))
        self.assertFalse(any("Form 6" in warning or "manifest" in warning for warning in result["warnings"]), result)
        self.assert_untouched(before)

    def test_preview_uses_all_material_rows_and_sums_their_weights(self):
        before = self.workbook_bytes()
        details = self.details(certificate_weight="", items=[
            {"item": "Laptops", "quantity": "2", "per_piece_weight": "5", "weight": "10"},
            {"item": "Monitors", "quantity": "3", "per_piece_weight": "4", "weight": ""},
        ])
        with patch.object(web_app, "make_template_certificate", wraps=web_app.make_template_certificate) as render:
            web_app.preview_certificate(self.body("certificate_only", details, use_certificate_weight=False))
        batch = render.call_args.args[0]
        self.assertEqual([item["item"] for item in batch["items"]], ["Laptops", "Monitors"])
        self.assertEqual(batch["items"][1]["weight"], 12)
        self.assertEqual(batch["total_weight"], 22)
        self.assertEqual(batch["display_weight"], 20)
        self.assert_untouched(before)

    def test_update_excel_saves_each_material_in_its_own_row(self):
        details = self.details(certificate_weight="", items=[
            {"item": "Laptops", "quantity": "2", "per_piece_weight": "5", "weight": "10"},
            {"item": "Monitors", "quantity": "3", "per_piece_weight": "4", "weight": "12"},
        ])
        created, review, _ = self.run_form(details)
        self.assertEqual((created, review), (1, 0), self.report())
        batch = self.assert_issued()
        self.assertEqual([(item["item"], item["weight"]) for item in batch["items"]],
                         [("Laptops", 10), ("Monitors", 12)])
        self.assertEqual(batch["total_weight"], 22)

    def test_update_excel_preserves_formula_backed_material_weights(self):
        book = load_workbook(self.paths[0])
        try:
            sheet = book["Inward Register"]
            sheet["H2"] = "=F2*G2"
            sheet["H3"] = "=F3*G3"
            book.save(self.paths[0])
        finally:
            book.close()

        details = self.details(certificate_weight="", items=[
            {"item": "Laptops", "quantity": "3", "per_piece_weight": "7", "weight": "21"},
            {"item": "Monitors", "quantity": "3", "per_piece_weight": "4", "weight": "12"},
        ])
        created, review, _ = self.run_form(details)
        self.assertEqual((created, review), (1, 0), self.report())

        book = load_workbook(self.paths[0], data_only=False)
        try:
            sheet = book["Inward Register"]
            self.assertEqual(sheet["H2"].value, "=F2*G2")
            self.assertEqual(sheet["H3"].value, "=F3*G3")
            self.assertEqual((sheet["F2"].value, sheet["G2"].value), (3, 7))
            self.assertEqual((sheet["F3"].value, sheet["G3"].value), (3, 4))
        finally:
            book.close()

        batch = self.assert_issued()
        self.assertEqual([(item["item"], item["weight"]) for item in batch["items"]],
                         [("Laptops", 21), ("Monitors", 12)])
        self.assertEqual(batch["total_weight"], 33)

    def test_update_excel_appends_material_rows_for_last_client(self):
        details = self.details(certificate_weight="", items=[
            {"item": "Laptops", "quantity": "2", "per_piece_weight": "5", "weight": "10"},
            {"item": "Monitors", "quantity": "3", "per_piece_weight": "4", "weight": "12"},
            {"item": "Keyboards", "quantity": "4", "per_piece_weight": "1", "weight": "4"},
        ])
        created, review, _ = self.run_form(details)
        self.assertEqual((created, review), (1, 0), self.report())
        batch = self.assert_issued()
        self.assertEqual([item["item"] for item in batch["items"]], ["Laptops", "Monitors", "Keyboards"])
        self.assertEqual(batch["total_weight"], 26)

    def test_material_editor_keeps_incomplete_excel_rows_visible_for_correction(self):
        book = load_workbook(self.paths[0])
        try:
            book.active["E3"] = None
            book.save(self.paths[0])
        finally:
            book.close()
        batch = self.batch()
        self.assertEqual(len(batch["items"]), 2)
        self.assertEqual(batch["items"][1]["item"], "")
        self.assertEqual(batch["items"][1]["weight"], 12)
        self.assertIn("missing item description on row 3", batch["issues"])

    def test_adding_material_before_another_client_never_overwrites_that_client(self):
        book = load_workbook(self.paths[0])
        try:
            book.active.append(["NEXT", date(2026, 9, 8), "Next Client", "Pune", "Printers", 1, 2, 2,
                                "FORM-NEXT", "REF-NEXT", None, None, "Ready"])
            book.save(self.paths[0])
        finally:
            book.close()
        before = self.paths[0].read_bytes()
        details = self.details(items=[
            {"item": "Laptops", "weight": "10"},
            {"item": "Monitors", "weight": "12"},
            {"item": "Keyboards", "weight": "4"},
        ])
        with WorkbookSession(self.paths[0], self.backups) as session:
            with self.assertRaisesRegex(ValueError, "room for 2 material row"):
                stage_certificate_details(session, details, validate_cells=True)
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_certificate_only_commits_outputs_but_keeps_all_source_cells_and_metadata(self):
        before = self.cells()
        other = self.paths[1].read_bytes()
        original = deepcopy(self.batch())
        with patch.object(runner, "make_certificate", wraps=runner.make_certificate) as render:
            created, review, _ = self.run_form(update_excel=False)
        self.assertEqual((created, review), (1, 0), self.report())
        rendered = render.call_args.args[0]
        self.assertEqual(rendered["total_weight"], 33)
        self.assertEqual(rendered["items"][0],
                         {"item": "Laptops", "quantity": 3, "per_piece_weight": 7, "weight": 21})
        self.assertEqual(rendered["items"][1:], original["items"][1:])
        batch = self.assert_issued()
        after = self.cells()
        for row_index, row in enumerate(before):
            for column, value in enumerate(row):
                if row_index == 1 and column in {10, 11, 12}:
                    continue
                self.assertEqual(after[row_index][column], value, (row_index, column))
        for key in ("supplier", "location", "inward_date", "form6", "reference", "items",
                    "total_weight", "certificate_weight", "certificate_type", "certificate_fields"):
            self.assertEqual(batch[key], original[key], key)
        self.assertEqual(self.paths[1].read_bytes(), other)
        pdf = next(self.folder.rglob("*.pdf"))
        text = " ".join(page.extract_text() or "" for page in PdfReader(pdf).pages)
        for value in ("New Client", "R-EDIT", "F-EDIT", "9.5"):
            self.assertIn(value, text)
        self.assertNotIn("Original Client", text)

    def test_update_excel_commits_full_details_and_first_item_preserves_continuation(self):
        continuation = deepcopy(self.batch()["items"][1:])
        other = self.paths[1].read_bytes()
        created, review, _ = self.run_form()
        self.assertEqual((created, review), (1, 0), self.report())
        batch = self.assert_issued()
        for key, expected in {"supplier": "Edited Client Ltd", "location": "Surat",
                              "inward_date": date(2026, 9, 5), "form6": "F-EDIT", "reference": "R-EDIT",
                              "certificate_weight": 9.5, "certificate_type": "standard",
                              "certificate_fields": {"company_display": "New Client"}}.items():
            self.assertEqual(batch[key], expected, key)
        first = batch["items"][0]
        for key, expected in {"item": "Laptops", "quantity": 3, "per_piece_weight": 7, "weight": 21}.items():
            self.assertEqual(first[key], expected, key)
        self.assertEqual(batch["items"][1:], continuation)
        self.assertEqual(batch["total_weight"], 33)
        self.assertEqual(self.paths[1].read_bytes(), other)

    def test_web_success_uses_requested_scope_and_refreshes_only_after_commit(self):
        original = self.batch()["supplier"]
        def after_commit(*_args, **_kwargs):
            self.assert_issued()
            self.assertEqual(self.batch()["supplier"], original)
            return self.report()
        with patch.object(web_app, "refresh", side_effect=after_commit) as refresh:
            result = web_app.create_certificate(self.body("certificate_only"))
        self.assertFalse(result["requires_confirmation"])
        self.assertEqual(result["created"], 1)
        refresh.assert_called_once()
        self.assertIn("original inward details were kept", web_app.target_row(result["state"], ("plant1", "CLIENT"))["detail"])

    def test_changed_template_type_is_rendered_but_not_persisted_in_certificate_only_scope(self):
        original_metadata = (self.batch()["certificate_type"], self.batch()["certificate_fields"])
        details = self.details(certificate_type="sita", certificate_fields={"company_display": "Sita Client"})
        with patch.object(runner, "make_certificate", wraps=runner.make_certificate) as render:
            created, review, _ = self.run_form(details, update_excel=False)
        self.assertEqual((created, review), (1, 0), self.report())
        self.assertEqual(render.call_args.args[0]["certificate_type"], "sita")
        batch = self.assert_issued()
        self.assertEqual((batch["certificate_type"], batch["certificate_fields"]), original_metadata)
        text = " ".join(page.extract_text() or "" for page in PdfReader(next(self.folder.rglob("*.pdf"))).pages)
        self.assertIn("Sita Client", text)

    def test_default_web_scope_preserves_excel_details_for_corrections(self):
        body = self.body(details=self.details(certificate_type="sita", certificate_fields={"company_display": "Sita Client"}))
        del body["edit_scope"]
        with patch.object(web_app, "refresh", side_effect=lambda *_: self.report()):
            result = web_app.create_certificate(body)
        self.assertEqual(result["created"], 1)
        batch = self.assert_issued()
        self.assertEqual(batch["supplier"], "Independent Client Ltd")
        self.assertEqual(batch["certificate_type"], "standard")
        self.assertEqual(batch["certificate_fields"], {"company_display": "Original Client"})

    def assert_explicit_status_can_issue(self, status):
        book = load_workbook(self.paths[0])
        try:
            book.active["M2"] = status
            book.save(self.paths[0])
        finally:
            book.close()
        before = self.workbook_bytes()
        created, review, _ = self.run_form(check_only=True)
        self.assertEqual((created, review), (0, 0), self.report())
        self.assertEqual(self.workbook_bytes(), before)
        created, review, _ = self.run_form(update_excel=False)
        self.assertEqual((created, review), (1, 0), self.report())
        self.assert_issued()

    def test_explicit_form_can_issue_blank_without_persisting_preflight(self):
        self.assert_explicit_status_can_issue("")

    def test_explicit_form_can_issue_na_without_persisting_preflight(self):
        self.assert_explicit_status_can_issue("N/A")

    def test_new_row_preflight_cancel_then_success_appends_only_after_confirmation(self):
        details = self.details(case_id="NEW", reference="R-NEW", inward_date="")
        before = self.workbook_bytes()
        self.run_form(details, check_only=True)
        self.assert_untouched(before)
        with patch.object(web_app, "refresh", side_effect=AssertionError("No refresh before confirmation")):
            result = web_app.create_certificate(self.body(details=details))
        self.assertTrue(result["requires_confirmation"])
        self.assert_untouched(before)
        with patch.object(web_app, "refresh", side_effect=lambda *_: self.report()):
            confirmed = web_app.create_certificate(self.body(details=details, confirmed=True))
        self.assertFalse(confirmed["requires_confirmation"])
        self.assertEqual(confirmed["created"], 1, self.report())
        batches = read_batches(self.paths[0])
        self.assertEqual([batch["case_id"] for batch in batches], ["CLIENT", "NEW"])
        self.assertEqual(batches[1]["issued"], "Generated")
        self.assertEqual(batches[1]["cod_no"], "0041")
        self.assertEqual(batches[0]["items"][1]["item"], "Monitors")
        self.assertEqual(self.paths[1].read_bytes(), before[1])

    def test_certificate_only_rejects_new_row_without_any_excel_write(self):
        before = self.workbook_bytes()
        with self.assertRaises(ValueError):
            web_app.create_certificate(self.body("certificate_only", self.details(case_id="NEW")))
        self.assert_untouched(before)

    def test_immutable_source_identity_and_invalid_scope_are_rejected_before_writes(self):
        before = self.workbook_bytes()
        for changes in ({"source_sn": "OTHER"}, {"source_plant": "plant2"}, {"edit_scope": "everything"}):
            body = self.body()
            body.update(changes)
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                web_app.create_certificate(body)
        with self.assertRaises(ValueError):
            runner.run(self.paths, self.root, details=self.details(case_id="OTHER"),
                       selected=("plant1", "CLIENT"), check_only=True, progress=None)
        self.assert_untouched(before)

    def test_render_failure_never_persists_edits_or_leaves_partial_pdf(self):
        before = self.workbook_bytes()
        with patch.object(runner, "make_certificate", side_effect=PermissionError("render failed")):
            created, review, _ = self.run_form()
        self.assertEqual((created, review), (0, 1))
        self.assertIn("render failed", str(self.report()))
        self.assert_untouched(before)

    def test_commit_failure_rolls_back_published_pdf_and_staged_workbook_edits(self):
        before = self.workbook_bytes()
        with patch.object(WorkbookSession, "_persist", side_effect=PermissionError("commit failed")):
            created, review, _ = self.run_form()
        self.assertEqual((created, review), (0, 1))
        self.assertIn("commit failed", str(self.report()))
        self.assert_untouched(before)

    def test_failed_new_row_commit_does_not_append_even_after_pdf_rendering(self):
        before = self.workbook_bytes()
        with patch.object(WorkbookSession, "_persist", side_effect=PermissionError("new row commit failed")):
            created, review, _ = self.run_form(self.details(case_id="NEW", reference="R-NEW"))
        self.assertEqual((created, review), (0, 1), self.report())
        self.assertEqual([batch["case_id"] for batch in read_batches(self.paths[0])], ["CLIENT"])
        self.assert_untouched(before)

    def test_web_failure_raises_without_refresh_or_saving_form_edits(self):
        before = self.workbook_bytes()
        with patch.object(runner, "make_certificate", side_effect=PermissionError("render failed")), \
                patch.object(web_app, "refresh") as refresh:
            with self.assertRaisesRegex(ValueError, "render failed"):
                web_app.create_certificate(self.body())
        refresh.assert_not_called()
        self.assert_untouched(before)

    def test_duplicate_cod_requires_confirmation_then_uses_shared_next_number(self):
        before = self.workbook_bytes()
        details = self.details(cod_no="0040")
        with patch.object(web_app, "refresh", side_effect=AssertionError("Warning return must not refresh")):
            result = web_app.create_certificate(self.body(details=details))
        self.assertTrue(result["requires_confirmation"])
        self.assertIn("another", str(result["warnings"]).lower())
        self.assert_untouched(before)
        created, _, _ = self.run_form(details, custom_cod="0040", allow_warnings=True)
        self.assertEqual(created, 1, self.report())
        self.assert_issued("0041")
        self.assertEqual(self.paths[1].read_bytes(), before[1])

    def test_existing_generated_explicit_reissue_preflight_and_failure_keep_original_pdf(self):
        self.assertEqual(self.run_form()[0], 1, self.report())
        before = self.workbook_bytes()
        pdf = next(self.folder.rglob("*.pdf"))
        original_pdf = pdf.read_bytes()
        edits = self.details(cod_no="0041", certificate_fields={"company_display": "Reissued Client"})
        created, review, _ = self.run_form(edits, update_excel=False, check_only=True)
        self.assertEqual((created, review), (0, 0), self.report())
        self.assertEqual(self.workbook_bytes(), before)
        with patch.object(WorkbookSession, "_persist", side_effect=PermissionError("reissue commit failed")):
            created, review, _ = self.run_form(edits, update_excel=False)
        self.assertEqual((created, review), (0, 1), self.report())
        self.assertEqual(self.workbook_bytes(), before)
        self.assertEqual(pdf.read_bytes(), original_pdf)
        self.assertEqual(list(self.folder.rglob('*.pdf')), [pdf])
        created, review, _ = self.run_form(edits, update_excel=False)
        self.assertEqual((created, review), (1, 0), self.report())
        self.assert_issued("0041")
        text = " ".join(page.extract_text() or "" for page in PdfReader(pdf).pages)
        self.assertIn("Reissued Client", text)
        self.assertEqual(self.batch()["certificate_fields"], {"company_display": "New Client"})


if __name__ == "__main__":
    unittest.main()
