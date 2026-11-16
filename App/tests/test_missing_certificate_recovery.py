"""Missing-PDF recovery integration tests, isolated from real plant files."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from pypdf import PdfWriter

from certificate_core import parse_date, read_batches
import certificate_drafts as drafts
import run_certificates as runner
import web_app
from tests.test_certificate_templates import write_register
from workflow_support import save_json


APP = Path(__file__).resolve().parents[1]
AS_OF = date(2026, 10, 1)
RECORDED_DATE = date(2026, 9, 14)
COD = "CUSTOM-077"


class MissingCertificateRecoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        self.folder = self.root / "Certificates"
        self.folder.mkdir()
        self.backups = self.root / "Data" / "Workbook Backups"
        self.results = self.root / "Data" / "Reports" / "Latest Results.json"
        self.source(status="N/A", cod="0016")
        native = self.root / "App" / "templates" / "certificate-types"
        native.mkdir(parents=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf", native / "standard.pdf")
        # The second source supplies a numbering floor without becoming an
        # additional issued-but-missing row in these focused recovery tests.
        write_register(self.paths[1], case_id="OTHER", cod="0040", status="N/A")
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "certificate_number_prefix": "", "certificate_number_start": 1,
            "certificate_number_width": 4, "reminder_time": "11:00", "reminders_enabled": False,
        })

    def source(self, *, status, cod=COD, continuation=False):
        write_register(self.paths[0], case_id="RECOVER", cod=cod, status=status, continuation=continuation)
        book = load_workbook(self.paths[0])
        try:
            book.active["L2"] = RECORDED_DATE
            book.save(self.paths[0])
        finally:
            book.close()

    def source_batch(self):
        return read_batches(self.paths[0])[0]

    def certificate_pdfs(self):
        return list(self.folder.rglob("*.pdf"))

    def report_row(self, report=None):
        if report is None:
            report = json.loads(self.results.read_text(encoding="utf-8"))
        return next(row for row in report["rows"]
                    if row.get("source_kind") == "workbook" and row.get("plant") == "plant1" and row.get("sn") == "RECOVER")

    @contextmanager
    def forbid_generation(self):
        with patch.object(runner, "make_certificate", side_effect=AssertionError("Refresh must not render a workbook PDF")) as workbook, \
                patch.object(drafts, "make_template_certificate", side_effect=AssertionError("Refresh must not render a manual PDF")) as manual:
            yield
            workbook.assert_not_called()
            manual.assert_not_called()

    def check(self):
        with self.forbid_generation():
            created, review, _ = runner.run(self.paths, self.root, as_of=AS_OF, check_only=True, progress=None)
        self.assertEqual(created, 0)
        return review

    def bind_web_workspace(self):
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "BACKUP_DIR", self.backups))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", self.results))

    def cache_issued_row(self, status, cod=COD):
        save_json(self.results, {
            "rows": [{"source_kind": "workbook", "plant": "plant1", "sn": "RECOVER",
                      "certificate_status": status.casefold(), "outcome": status.casefold(),
                      "certificate_no": cod, "certificate_date": RECORDED_DATE.isoformat(),
                      "detail": f"COD {cod} was marked {status} in the previous check.", "warnings": []}],
            "scan_complete": True, "created": 0, "review": 0,
        })

    def test_unmatched_external_pdf_stays_generated_and_can_be_filtered_as_unavailable(self):
        self.source(status="Generated Externally")
        self.check()
        row = self.report_row()
        self.assertEqual(row["issue_code"], "pdf_unavailable")
        self.assertEqual(row["certificate_status"], "generated_manual")
        self.assertTrue(row["system_alert"])
        self.assertEqual(self.source_batch()["cod_no"], COD)

    def test_cached_legacy_statuses_are_canonicalized_and_persisted_on_workspace_read(self):
        self.bind_web_workspace()
        expectations = (("Created", "generated"), ("Sent", "completed"),
                        ("Check Details", "alert"), ("Working Draft", "alert"),
                        ("Generated Manual", "generated_manual"), ("Blank", "blank"),
                        ("N/A", "na"))
        for old_status, expected in expectations:
            with self.subTest(old_status=old_status):
                self.cache_issued_row(old_status)
                state = web_app.workspace_state(recheck_obsolete=False)
                row = self.report_row(state)
                self.assertEqual(row["certificate_status"], expected)
                stored = json.loads(self.results.read_text(encoding="utf-8"))
                self.assertEqual(self.report_row(stored)["certificate_status"], expected)

    def copy_pdf(self, destination):
        # Copy a supplied asset into a disposable folder solely as an existing
        # file. Its content is not adopted as an issued certificate or edited.
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf", destination)
        return destination.read_bytes()

    def assert_ready_preserves_identity(self, cod=COD, items=None):
        batch = self.source_batch()
        self.assertEqual((batch["issued"], batch["cod_no"], parse_date(batch["cod_date"])),
                         ("Ready", cod, RECORDED_DATE))
        if items is not None:
            self.assertEqual(batch["items"], items)
        return batch

    def test_generated_and_completed_missing_pdfs_persist_ready_keep_cod_date_items_and_backup(self):
        for status in ("Generated", "Completed"):
            with self.subTest(status=status):
                self.source(status=status, continuation=True)
                original = self.paths[0].read_bytes()
                items = self.source_batch()["items"]
                self.check()
                self.assert_ready_preserves_identity(items=items)
                row = self.report_row()
                self.assertEqual(row["certificate_status"], "ready")
                self.assertIn("missing", row["detail"].casefold())
                self.assertTrue(any(path.read_bytes() == original for path in self.backups.glob("plant1_*.xlsx")))
                self.assertEqual(self.certificate_pdfs(), [])

    def test_missing_cod_generated_recovers_without_allocating_until_explicit_generation(self):
        self.source(status="Generated", cod="")
        self.check()
        self.assert_ready_preserves_identity(cod="")
        self.assertEqual(self.certificate_pdfs(), [])
        created, review, _ = runner.run(self.paths, self.root, as_of=AS_OF, selected=("plant1", "RECOVER"), progress=None)
        self.assertEqual((created, review), (1, 0))
        batch = self.source_batch()
        self.assertEqual((batch["issued"], batch["cod_no"], parse_date(batch["cod_date"])),
                         ("Generated", "0041", RECORDED_DATE))
        self.assertTrue((self.folder / runner.certificate_filename(batch, 1, "0041")).is_file())

    def test_explicit_regeneration_keeps_existing_cod_and_recorded_date_after_completed_recovery(self):
        self.source(status="Completed")
        self.check()
        self.assert_ready_preserves_identity()
        created, review, _ = runner.run(self.paths, self.root, as_of=AS_OF, selected=("plant1", "RECOVER"), progress=None)
        self.assertEqual((created, review), (1, 0))
        batch = self.source_batch()
        self.assertEqual((batch["issued"], batch["cod_no"], parse_date(batch["cod_date"])),
                         ("Generated", COD, RECORDED_DATE))

    def test_manually_placed_pdf_with_expected_name_is_picked_up_without_app_generation(self):
        self.source(status="Generated")
        batch = self.source_batch()
        filename = runner.certificate_filename(batch, 1, COD)
        manual_pdf = self.folder / filename
        writer = PdfWriter()
        writer.add_blank_page(width=612, height=792)
        with manual_pdf.open("wb") as stream:
            writer.write(stream)
        self.assertEqual(filename, "CUSTOM-077_Independent Client Ltd_Pune.pdf")

        review = self.check()

        self.assertEqual(review, 0)
        self.assertEqual(self.source_batch()["issued"], "Generated")
        row = self.report_row()
        self.assertEqual(row["certificate_status"], "generated")
        self.assertIn("generated PDF", row["detail"])
        self.assertTrue(manual_pdf.is_file())
        self.assertEqual(list(self.folder.glob("*.pdf")), [manual_pdf])

    def test_pdf_in_manually_added_folder_is_picked_up_without_app_generation(self):
        self.source(status="Generated")
        batch = self.source_batch()
        manual_pdf = self.folder / "Manually Added" / runner.certificate_filename(batch, 1, COD)
        writer = PdfWriter()
        writer.add_blank_page(width=612, height=792)
        manual_pdf.parent.mkdir(parents=True)
        with manual_pdf.open("wb") as stream:
            writer.write(stream)

        review = self.check()

        self.assertEqual(review, 0)
        self.assertEqual(self.source_batch()["issued"], "Generated")
        self.assertEqual(self.report_row()["certificate_status"], "generated")
        self.assertEqual(self.certificate_pdfs(), [manual_pdf])

    def test_legacy_underscore_company_filename_is_still_picked_up(self):
        self.source(status="Generated")
        batch = self.source_batch()
        current_name = runner.certificate_filename(batch, 1, COD)
        legacy_name = current_name.replace("Independent Client Ltd", "Independent_Client_Ltd").replace('.pdf', '_1.pdf')
        manual_pdf = self.folder / legacy_name
        writer = PdfWriter()
        writer.add_blank_page(width=612, height=792)
        with manual_pdf.open("wb") as stream:
            writer.write(stream)

        review = self.check()

        self.assertEqual(review, 0)
        self.assertEqual(self.source_batch()["issued"], "Generated")
        self.assertTrue(manual_pdf.is_file())
        self.assertEqual(list(self.folder.glob("*.pdf")), [manual_pdf])

    def test_repeated_workbook_recovery_is_idempotent_without_reapplying_status(self):
        self.source(status="Completed", continuation=True)
        self.check()
        recovered = self.source_batch()
        with patch.object(runner.WorkbookSession, "set_issued_status",
                          side_effect=AssertionError("A Ready row must not repeat missing-PDF recovery")) as transition:
            # The first normal Ready scan may normalize its alert formatting.
            # It must not repeat the recovery or alter any parsed source data.
            self.check()
        transition.assert_not_called()
        self.assertEqual(self.source_batch(), recovered)
        after = [path.read_bytes() for path in self.paths]
        backups = {path.name: path.read_bytes() for path in self.backups.glob("*.xlsx")}
        with patch.object(runner.WorkbookSession, "set_issued_status",
                          side_effect=AssertionError("A Ready row must not repeat missing-PDF recovery")) as transition:
            self.check()
            self.check()
        transition.assert_not_called()
        self.assert_ready_preserves_identity()
        self.assertEqual([path.read_bytes() for path in self.paths], after)
        self.assertEqual({path.name: path.read_bytes() for path in self.backups.glob("*.xlsx")}, backups)
        self.assertEqual(self.certificate_pdfs(), [])

    def test_archived_matching_pdf_does_not_count_as_active_presence(self):
        archived = self.root / "Data" / "Archive" / "Superseded Certificates" / f"{'a' * 32}_{COD}_Old_Client_1.pdf"
        original = self.copy_pdf(archived)
        for status in ("Generated", "Completed"):
            with self.subTest(status=status):
                self.source(status=status)
                self.check()
                self.assert_ready_preserves_identity()
                self.assertEqual(archived.read_bytes(), original)
                self.assertEqual(self.certificate_pdfs(), [])

    def test_cached_workspace_rechecks_missing_generated_completed_and_wrong_plant_files(self):
        self.bind_web_workspace()
        unrelated = self.folder / f"{COD}_Unrelated_Client_2.pdf"
        for status in ("Generated", "Completed"):
            for wrong_plant_present in (False, True):
                with self.subTest(status=status, wrong_plant_present=wrong_plant_present):
                    self.source(status=status)
                    unrelated.unlink(missing_ok=True)
                    original = self.copy_pdf(unrelated) if wrong_plant_present else None
                    self.cache_issued_row(status)
                    with self.forbid_generation(), patch.object(web_app, "run", wraps=runner.run) as check:
                        state = web_app.workspace_state()
                        check.assert_called_once()
                        self.assertTrue(check.call_args.kwargs.get("check_only"))
                        self.assert_ready_preserves_identity()
                        row = self.report_row(state)
                        self.assertEqual(row["certificate_status"], "ready")
                        self.assertFalse(row.get("has_issued_certificate", False))
                        self.assertNotIn(row["outcome"], ("generated", "completed", "generated_now"))
                        self.assertIn("missing", row["detail"].casefold())
                        web_app.workspace_state()
                        check.assert_called_once()  # Ready cache does not loop back into refresh.
                    if wrong_plant_present:
                        self.assertEqual(unrelated.read_bytes(), original)
                        self.assertEqual(self.certificate_pdfs(), [unrelated])
                    else:
                        self.assertEqual(self.certificate_pdfs(), [])

    def test_cached_generated_without_cod_persists_ready_without_rendering_or_allocating(self):
        self.bind_web_workspace()
        self.source(status="Generated", cod="")
        self.cache_issued_row("Generated", cod="")
        with self.forbid_generation(), patch.object(web_app, "run", wraps=runner.run) as check:
            state = web_app.workspace_state()
        check.assert_called_once()
        self.assertTrue(check.call_args.kwargs.get("check_only"))
        self.assert_ready_preserves_identity(cod="")
        self.assertEqual(self.report_row(state)["certificate_status"], "ready")
        self.assertEqual(self.certificate_pdfs(), [])

    def test_cached_issued_row_with_active_matching_pdf_does_not_trigger_recovery(self):
        self.bind_web_workspace()
        active = self.folder / f"{COD}_Old_Client_1.pdf"
        original = self.copy_pdf(active)
        for status in ("Generated", "Completed"):
            with self.subTest(status=status):
                self.source(status=status)
                self.cache_issued_row(status)
                before = self.paths[0].read_bytes()
                with self.forbid_generation(), patch.object(web_app, "run", side_effect=AssertionError("Active PDF must not trigger recovery")):
                    state = web_app.workspace_state()
                self.assertEqual(self.report_row(state)["certificate_status"], status.casefold())
                self.assertEqual(self.paths[0].read_bytes(), before)
                self.assertEqual(active.read_bytes(), original)

    def manual_details(self):
        return {"case_id": "MANUAL-RECOVER", "supplier": "Independent Client Ltd", "location": "Pune",
                "certificate_type": "sita", "certificate_fields": {}, "inward_date": "01.09.2026",
                "certificate_weight": "8.5", "weight": "10", "reference": "M-REC"}

    def test_generated_and_completed_manual_records_recover_idempotently_with_identity_preserved(self):
        record_id = drafts.save_draft(self.root, "plant1", self.manual_details())
        for status in ("generated", "completed"):
            with self.subTest(status=status):
                data = drafts.read_register(self.root)
                record = data["records"][0]
                record.update(status=status, pdf_filename=f"{COD}_Independent_Client_Ltd_Pune_1.pdf")
                record["details"].update(cod_no=COD, cod_date=RECORDED_DATE.isoformat())
                drafts.write_register(self.root, data)
                with self.forbid_generation():
                    drafts.reconcile_manual(self.root)
                restored = drafts.read_register(self.root)["records"][0]
                self.assertEqual((restored["id"], restored["status"], restored["details"]["cod_no"], restored["details"]["cod_date"]),
                                 (record_id, "ready", COD, RECORDED_DATE.isoformat()))
                path = self.root / "Data" / "Custom Certificates.json"
                before = path.read_bytes()
                backups = set((self.root / "Data" / "Custom Register Backups").glob("*.json"))
                with self.forbid_generation():
                    drafts.reconcile_manual(self.root)
                    drafts.reconcile_manual(self.root)
                self.assertEqual(path.read_bytes(), before)
                self.assertEqual(set((self.root / "Data" / "Custom Register Backups").glob("*.json")), backups)
                self.assertEqual(drafts.manual_rows(self.root)[0]["certificate_status"], "ready")
                self.assertEqual(self.certificate_pdfs(), [])

    def test_manually_completed_deleted_pdf_recovers_on_cached_workspace_read_without_generation(self):
        self.bind_web_workspace()
        response = drafts.generate_draft(self.root, plant="plant1", details=self.manual_details(), as_of=AS_OF)
        self.assertEqual(response["created"], 1)
        record_id = response["record_id"]
        drafts.complete_draft(self.root, record_id)
        before = drafts.read_register(self.root)["records"][0]
        self.assertEqual(before["status"], "completed")
        output = self.folder / response["relative_path"]
        output.unlink()
        save_json(self.results, {"rows": [], "scan_complete": True, "created": 0, "review": 0})
        with self.forbid_generation(), patch.object(web_app, "run", side_effect=AssertionError("Manual recovery must not generate or require a workbook scan")):
            state = web_app.workspace_state()
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual((record["id"], record["status"], record["details"]["cod_no"], record["details"]["cod_date"]),
                         (record_id, "ready", before["details"]["cod_no"], before["details"]["cod_date"]))
        row = next(row for row in state["rows"] if row.get("record_id") == record_id)
        self.assertEqual(row["certificate_status"], "ready")
        self.assertEqual(row["certificate_no"], response["certificate_no"])
        self.assertEqual(self.certificate_pdfs(), [])


if __name__ == "__main__":
    unittest.main()
