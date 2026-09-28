"""Correction workflow tests use real PDFs/XLSX and a non-sending Outlook double."""
from datetime import date
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from certificate_core import read_batches
from certificate_correction import correct_certificate
from draft_service import History
from test_workflow import FakeOutlook, ROOT, TODAY, make_book
from workflow_support import save_json
from workbook_store import WorkbookSession
from pypdf import PdfReader
from openpyxl import load_workbook


class FakeCorrectionOutlook:
    updates = {}

    def __init__(self, mailbox):
        self.mailbox = mailbox

    def validate_certificate_draft(self, entry_id, batch, expected_subject, pdf_path):
        if not Path(pdf_path).is_file():
            raise FileNotFoundError(pdf_path)

    def update_certificate_draft(self, entry_id, batch, old_subject, new_subject,
                                 old_pdf_path, new_pdf_path, banner_path=None):
        self.updates[entry_id] = (new_subject, str(new_pdf_path))


class CertificateCorrectionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.paths = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsx"]
        save_json(self.root / "settings.example.json", {
            "eligibility_days": 30, "certificate_number_start": 1,
            "initial_cod_confirmed": True, "certificate_number_width": 4,
            "certificate_number_prefix": "", "workbook_paths": ["", ""],
        })
        (self.root / "templates").mkdir()
        shutil.copy2(ROOT / "templates" / "Certificate of Recycling.png",
                     self.root / "templates" / "Certificate of Recycling.png")
        FakeOutlook.saved = []
        self.mock = patch("run_certificates.OutlookWriter", FakeOutlook)
        self.mock.start()
        FakeCorrectionOutlook.updates = {}

    def tearDown(self):
        self.mock.stop()
        self.temp.cleanup()

    def one_ready(self):
        make_book(self.paths[0], [{}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0061"}])

    def execute(self, **kwargs):
        from run_certificates import run
        created, review, output = run(
            self.paths, "dummy@example.test", self.root, as_of=TODAY,
            progress=lambda _: None, **kwargs,
        )
        import json
        self.report = json.loads((output / "Results.json").read_text())
        return created, review, output

    def correct(self, cod="0999", issued="24.09.2026", reason="Corrected after checking source record"):
        return correct_certificate(
            self.paths, "dummy@example.test", self.root, "plant1", "1",
            cod, issued, reason, progress=lambda _: None,
        )

    def create_one_confirmed_draft(self):
        self.one_ready()
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_correction_updates_existing_draft_pdf_excel_and_audit_without_sending(self):
        self.create_one_confirmed_draft()
        book = load_workbook(self.paths[0])
        book["Inward Register"]["L2"] = "0998"
        book["Inward Register"]["M2"] = TODAY.replace(day=22)
        book.save(self.paths[0])
        book.close()
        old = read_batches(self.paths[0])[0]
        old_pdf = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            old_record = old_pdf.record(old)
        finally:
            old_pdf.close()
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            result = self.correct()

        corrected = read_batches(self.paths[0])[0]
        self.assertEqual((corrected["cod_no"], corrected["cod_date"].date(), corrected["issued"]),
                         ("0999", date(2026, 9, 24), "Drafted"))
        self.assertTrue(old_record["pdf_path"])
        self.assertTrue(Path(old_record["pdf_path"]).is_file())
        self.assertIn("0999", FakeCorrectionOutlook.updates["DUMMY-DRAFT-1"][0])
        self.assertEqual(len(FakeOutlook.saved), 1, "correction must edit the existing draft, not create another")
        text = PdfReader(result["pdf_path"]).pages[0].extract_text()
        self.assertIn("0999", text)
        self.assertIn("24.09.2026", text)

        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            active = history.record({"case_id": "1"})
            self.assertEqual((active["certificate_no"], active["certificate_date"], active["generation"]),
                             ("0999", "2026-09-24", 2))
            archived = history.db.execute(
                "SELECT certificate_no,certificate_date,status FROM archived_issuances WHERE case_id='1'"
            ).fetchone()
            self.assertEqual(tuple(archived), ("0062", "2026-09-23", "saved"))
            audit = history.db.execute(
                "SELECT old_certificate_no,new_certificate_no,reason,stage,observed_excel_certificate_no,"
                "observed_excel_certificate_date FROM certificate_corrections"
            ).fetchone()
            self.assertEqual(tuple(audit), ("0062", "0999", "Corrected after checking source record",
                                            "completed", "0998", "2026-09-22"))
        finally:
            history.close()

        self.execute(create_drafts=True)
        self.assertEqual(len(FakeOutlook.saved), 1, "rechecking a corrected Drafted row must not duplicate drafts")
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0999")

    def test_duplicate_cod_in_other_plant_is_rejected_before_any_update(self):
        self.create_one_confirmed_draft()
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0999"}])
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            with self.assertRaisesRegex(ValueError, "also entered|already assigned"):
                self.correct()
        self.assertFalse(FakeCorrectionOutlook.updates)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0062")
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            self.assertEqual(history.record({"case_id": "1"})["status"], "saved")
        finally:
            history.close()

    def test_date_only_correction_keeps_the_same_cod(self):
        self.create_one_confirmed_draft()
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            result = self.correct(cod="0062", issued="24.09.2026")
        corrected = read_batches(self.paths[0])[0]
        self.assertEqual(corrected["cod_no"], "0062")
        self.assertEqual(corrected["cod_date"].date(), date(2026, 9, 24))
        self.assertIn("0062", PdfReader(result["pdf_path"]).pages[0].extract_text())
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_correction_recovers_if_user_cleared_excel_cod_and_date(self):
        self.create_one_confirmed_draft()
        book = load_workbook(self.paths[0])
        book["Inward Register"]["L2"] = None
        book["Inward Register"]["M2"] = None
        book.save(self.paths[0])
        book.close()

        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            result = self.correct()

        corrected = read_batches(self.paths[0])[0]
        self.assertEqual((corrected["cod_no"], corrected["cod_date"].date(), corrected["issued"]),
                         ("0999", date(2026, 9, 24), "Drafted"))
        self.assertIn("0999", PdfReader(result["pdf_path"]).pages[0].extract_text())
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_correction_cannot_reuse_a_cod_archived_from_an_earlier_issue(self):
        self.create_one_confirmed_draft()
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            self.correct()
            calls_after_first = len(FakeCorrectionOutlook.updates)
            with self.assertRaisesRegex(ValueError, "already been used"):
                self.correct(cod="0062", issued="25.09.2026", reason="Attempted number reuse")
        self.assertEqual(len(FakeCorrectionOutlook.updates), calls_after_first)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0999")

    def test_sent_certificate_is_never_corrected_by_this_action(self):
        self.create_one_confirmed_draft()
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            history.acknowledge_sent("1", "0062")
        finally:
            history.close()
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            with self.assertRaisesRegex(ValueError, "confirmed Completed"):
                self.correct()
        self.assertFalse(FakeCorrectionOutlook.updates)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0062")

    def test_interrupted_excel_write_can_resume_without_duplicate_draft(self):
        self.create_one_confirmed_draft()
        original_commit = WorkbookSession.commit
        failed = []

        def fail_once(session, *args, **kwargs):
            if not failed:
                failed.append(True)
                raise OSError("simulated workbook write interruption")
            return original_commit(session, *args, **kwargs)

        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook), \
             patch("certificate_correction.WorkbookSession.commit", new=fail_once):
            with self.assertRaisesRegex(OSError, "simulated workbook"):
                self.correct()
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0062")
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            self.assertEqual(history.record({"case_id": "1"})["status"], "correcting")
            self.assertEqual(history.correction_for_case("1")["stage"], "outlook_updated")
        finally:
            history.close()

        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook):
            self.correct()
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0999")
        self.assertEqual(len(FakeOutlook.saved), 1)
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            self.assertEqual(history.record({"case_id": "1"})["status"], "saved")
            self.assertIsNone(history.correction_for_case("1"))
        finally:
            history.close()

    def test_normal_draft_processing_stops_while_correction_is_pending(self):
        self.create_one_confirmed_draft()
        with patch("certificate_correction.OutlookWriter", FakeCorrectionOutlook), \
             patch("certificate_correction.WorkbookSession.commit", side_effect=OSError("stop")):
            with self.assertRaisesRegex(OSError, "stop"):
                self.correct()
        _, _, _ = self.execute(create_drafts=True)
        self.assertTrue(any(row.get("issue_code") == "certificate_correction_pending"
                            for row in self.report["rows"]))
        self.assertEqual(len(FakeOutlook.saved), 1)


if __name__ == "__main__":
    unittest.main()
