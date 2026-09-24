"""Recovery contracts: temporary registers and mock COM only; never real Outlook."""
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from datetime import timedelta
import json
from pathlib import Path
import sqlite3
import sys
import tempfile
import types
import unittest
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from draft_service import History, OutlookWriter
from run_certificates import _read_sources, run
from test_workflow import FakeOutlook, TODAY, make_book
from workbook_store import WorkbookSession, exclusive_file
from workflow_support import save_json


class RecoveryEdges(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.paths = [self.root / "plant-a.xlsx", self.root / "plant-b.xlsx"]
        self.settings = {"eligibility_days": 30, "initial_cod_confirmed": True,
                         "certificate_number_start": 1, "certificate_number_width": 4}
        save_json(self.root / "settings.example.json", self.settings)
        make_book(self.paths[0], [{}])
        make_book(self.paths[1], [{"Certificate Issued": "N/A", "COD No": "0061"}])
        FakeOutlook.saved = []
        self.addCleanup(patch.stopall)
        patch("run_certificates.OutlookWriter", FakeOutlook).start()
        patch("run_certificates.make_certificate", return_value=None).start()

    @contextmanager
    def history(self):
        history = History(self.root / "Data" / "certificate-history.sqlite")
        try:
            yield history
        finally:
            history.close()

    def execute(self, **kwargs):
        result = run(self.paths, "dummy@example.test", self.root, as_of=TODAY,
                     progress=lambda _: None, **kwargs)
        self.report = json.loads((result[2] / "Results.json").read_text())
        return result

    def edit(self, index, cell, value):
        book = load_workbook(self.paths[index])
        book.active[cell] = value
        book.save(self.paths[index])
        book.close()

    def append_ready_case(self, path, case_id):
        book = load_workbook(path)
        sheet = book["Inward Register"]
        headers = {cell.value: cell.column for cell in sheet[1] if cell.value}
        row = sheet.max_row + 1
        values = {
            "SN": case_id,
            "Inward Date": (TODAY - timedelta(days=31)).strftime("%d.%m.%Y"),
            "Supplier Name": "New Test Client",
            "Location": "Mumbai",
            "Item": "E-waste equipment",
            "Quantity in Pcs": 2,
            "Per Pcs Weight in (KG)": 1.5,
            "Weight (KG)": f"=F{row}*G{row}",
            "Form 6 No": f"FORM6-{case_id}",
            "Client Email": f"{case_id.lower()}@example.test",
            "Reference No.": f"REF-{case_id}",
            "Certificate Issued": "Ready",
        }
        for header, value in values.items():
            sheet.cell(row, headers[header]).value = value
        book.save(path)
        book.close()

    def uncertain(self):
        with patch.object(FakeOutlook, "save_draft", side_effect=RuntimeError("Save result unavailable")):
            self.execute()

    def workbook_failure(self):
        with patch.object(WorkbookSession, "commit", side_effect=OSError("temporary write failure")):
            self.execute()

    def test_partial_read_still_reports_other_plant_blank_and_waiting(self):
        make_book(self.paths[0], [{"Certificate Issued": ""}, {"SN": "2"}])
        self.paths[1] = self.root / "missing.xlsx"
        self.execute()
        self.assertEqual(self.report["observed_sources"], ["plant1"])
        self.assertFalse(self.report["scan_complete"])
        self.assertIn("blank (white)", str(self.report))
        self.assertTrue(any(r["sn"] == "2" and r["outcome"] == "waiting" for r in self.report["rows"]))
        self.assertFalse(FakeOutlook.saved)

    def test_first_source_unreadable_still_reads_second_source(self):
        self.paths[0] = self.root / "missing.xlsx"
        make_book(self.paths[1], [{"Certificate Issued": ""}])
        self.execute()
        self.assertEqual(self.report["observed_sources"], ["plant2"])
        self.assertIn("blank (white)", str(self.report))

    def test_duplicate_custom_number_marks_its_owner_but_keeps_other_ready_row_eligible_in_check(self):
        make_book(self.paths[0], [{"COD No": "0061"}, {"SN": "2"}])
        self.execute(check_only=True)
        self.assertTrue(self.report["scan_complete"])
        self.assertTrue(any(r["sn"] == "1" and r["outcome"] == "needs_review" for r in self.report["rows"]))
        self.assertTrue(any(r["sn"] == "2" and r["outcome"] == "ready" for r in self.report["rows"]))
        self.assertFalse(FakeOutlook.saved)

    def test_lock_failure_reports_all_unprocessed_candidates(self):
        with exclusive_file(self.paths[1]):
            self.execute()
        self.assertTrue(any(r["sn"] == "1" and r["outcome"] == "waiting" for r in self.report["rows"]))
        self.assertFalse(FakeOutlook.saved)

    def test_reserved_number_cannot_be_claimed_by_other_plant(self):
        self.workbook_failure()
        self.edit(1, "L2", "0062")
        self.execute()
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(_read_sources(self.paths)[0]["cod_no"], "0062")
        self.assertEqual(_read_sources(self.paths)[0]["issued"], "Completed")

    def test_slot_swap_blocks_even_if_case_details_are_identical(self):
        make_book(self.paths[1], [{}])
        self.workbook_failure()
        count = len(FakeOutlook.saved)
        self.paths.reverse()
        self.execute()
        self.assertIn("slots were swapped", str(self.report))
        self.assertEqual(len(FakeOutlook.saved), count)

    def test_new_path_with_existing_history_requires_migration(self):
        self.workbook_failure()
        replacement = self.root / "replacement.xlsx"
        self.paths[0].rename(replacement)
        self.paths[0] = replacement
        self.execute()
        self.assertIn("history migration", str(self.report))
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_uncertain_save_survives_na_and_blocks_other_new_drafts(self):
        self.uncertain()
        self.edit(0, "N2", "N/A")
        make_book(self.paths[1], [{"COD No": "0061"}, {"SN": "2"}])
        self.execute()
        self.assertIn("unconfirmed for COD 0062", str(self.report))
        self.assertIn("Deleted Items", str(self.report))
        self.assertTrue(any(r["sn"] == "2" and r["outcome"] == "waiting" for r in self.report["rows"]))
        self.assertFalse(FakeOutlook.saved)

    def test_draft_saved_status_change_is_reported_until_restored(self):
        self.workbook_failure()
        for state in ("N/A", "", "Hold"):
            with self.subTest(state=state):
                self.edit(0, "N2", state)
                self.execute()
                self.assertIn("restore Ready", str(self.report))
                self.assertEqual(self.report["updated"], 0)
        self.edit(0, "N2", "Ready")
        self.execute()
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_removed_pending_row_still_reports_reserved_cod(self):
        self.uncertain()
        make_book(self.paths[0], [{"SN": "NEW"}])
        self.execute()
        self.assertIn("unconfirmed for COD 0062", str(self.report))
        self.assertTrue(any(r["sn"] == "NEW" and r["outcome"] == "waiting" for r in self.report["rows"]))

    def test_saved_excel_then_history_failure_recovers_without_new_draft(self):
        with patch.object(History, "complete", side_effect=sqlite3.OperationalError("simulated interrupted commit")):
            self.execute()
        self.assertEqual(_read_sources(self.paths)[0]["cod_no"], "0062")
        self.execute()
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_saved_excel_date_drift_is_restored_from_confirmed_history(self):
        self.execute()
        self.edit(0, "M2", TODAY - timedelta(days=1))
        self.execute()
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(_read_sources(self.paths)[0]["cod_date"].date(), TODAY)
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_blank_saved_cod_is_restored_without_issuing_again(self):
        self.execute()
        self.edit(0, "L2", None)
        self.edit(0, "M2", None)
        self.execute()
        self.assertEqual(self.report["updated"], 1)
        self.assertEqual(_read_sources(self.paths)[0]["cod_no"], "0062")
        self.assertEqual(_read_sources(self.paths)[0]["issued"], "Completed")
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_rekeyed_saved_batch_without_its_cod_blocks_new_draft(self):
        self.execute()
        self.edit(0, "A2", "REKEYED")
        self.edit(0, "L2", None)
        self.edit(0, "M2", None)
        self.execute()
        self.assertIn("Previously prepared COD 0062 has no matching source row", str(self.report))
        self.assertEqual(len(FakeOutlook.saved), 1)
        self.assertFalse(any(row["outcome"] == "draft_created" for row in self.report["rows"]))

    def test_duplicate_sn_in_one_plant_blocks_every_new_draft(self):
        make_book(self.paths[0], [{}, {"SN": "1"}])
        self.execute()
        self.assertIn("share a certificate history key", str(self.report))
        self.assertFalse(FakeOutlook.saved)

    def test_new_batch_added_after_a_successful_run_uses_next_shared_cod(self):
        self.execute()
        self.append_ready_case(self.paths[0], "NEW-2")
        self.execute()
        batches = _read_sources(self.paths)
        new_batch = next(batch for batch in batches if batch["case_id"] == "NEW-2")
        self.assertEqual(new_batch["cod_no"], "0063")
        self.assertEqual(len(FakeOutlook.saved), 2)
        self.assertIn("0063", FakeOutlook.saved[1][2])

    def test_changed_pending_source_is_not_reconciled_or_redrafted(self):
        self.workbook_failure()
        self.edit(0, "C2", "Different dummy supplier")
        self.execute()
        self.assertIn("Source details changed", str(self.report))
        self.assertEqual(self.report["updated"], 0)
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_orphaned_history_still_blocks_new_numbers(self):
        self.workbook_failure()
        with self.history() as history:
            history.db.execute(
                "INSERT INTO batches(case_id,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,generation,status_written,prior_certificate_no,prior_certificate_date) "
                "SELECT 'other',fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,generation,status_written,prior_certificate_no,prior_certificate_date FROM batches LIMIT 1"
            )
            history.db.commit()
        self.execute()
        self.assertIn("no matching source row", str(self.report))
        self.assertTrue(any(row["outcome"] == "waiting" for row in self.report["rows"]))
        self.assertEqual(len(FakeOutlook.saved), 1)

    def test_history_namespace_collision_blocks_before_allocation(self):
        make_book(self.paths[0], [{"SN": "plant2|1"}])
        make_book(self.paths[1], [{}])
        self.execute()
        self.assertIn("share a certificate history key", str(self.report))
        self.assertFalse(FakeOutlook.saved)

    def test_blank_outlook_entry_keeps_reservation_unconfirmed(self):
        with patch.object(FakeOutlook, "save_draft", return_value=None) as save:
            self.execute()
            self.execute()
            self.assertEqual(save.call_count, 1)
        with self.history() as history:
            self.assertEqual(history.records()[0]["status"], "saving_draft")

    def test_pending_reset_cannot_delete_duplicate_prevention(self):
        self.uncertain()
        with self.history() as history:
            with self.assertRaisesRegex(ValueError, "cannot be reset"):
                history.reset_pending()
            self.assertEqual(len(history.records()), 1)

    def test_legacy_pending_with_excel_cod_is_never_marked_already_prepared(self):
        batch = _read_sources(self.paths)[0]
        with self.history() as history:
            history.reserve(batch, "0062")
        self.edit(0, "L2", "0062")
        self.execute()
        row = next(r for r in self.report["rows"] if r["plant"] == "plant1")
        self.assertEqual(row["outcome"], "needs_review")
        self.assertIn("unconfirmed for COD 0062", row["detail"])
        self.assertFalse(FakeOutlook.saved)

    def test_snapshot_change_stops_before_number_or_draft_allocation(self):
        original_enter = WorkbookSession.__enter__
        def enter(session):
            if session.path == self.paths[0]:
                self.edit(0, "C2", "Changed while run was checking")
            return original_enter(session)
        with patch.object(WorkbookSession, "__enter__", enter):
            self.execute()
        self.assertIn("changed during the check", str(self.report))
        self.assertTrue(any(r["outcome"] == "waiting" for r in self.report["rows"]))
        with self.history() as history:
            self.assertFalse(history.records())
        self.assertFalse(FakeOutlook.saved)

    def test_unprintable_reference_is_flagged_before_number_allocation(self):
        self.edit(0, "K2", "W" * 1000)
        self.execute()
        self.assertIn("too long to fit", str(self.report))
        with self.history() as history:
            self.assertFalse(history.records())
        self.assertFalse(FakeOutlook.saved)

    def test_first_uncertain_save_reports_every_remaining_candidate(self):
        make_book(self.paths[0], [{}, {"SN": "2"}, {"SN": "3"}])
        self.uncertain()
        rows = [r for r in self.report["rows"] if r["plant"] == "plant1"]
        self.assertEqual(len(rows), 3)
        self.assertEqual([r["outcome"] for r in rows], ["needs_review", "waiting", "waiting"])

    def test_save_claim_is_atomic_across_workers(self):
        batch = _read_sources(self.paths)[0]
        with self.history() as history:
            history.allocate(batch, self.settings, 62, TODAY)
        def claim(_):
            with self.history() as history:
                try:
                    history.saving_draft(batch, "dummy.pdf")
                    return True
                except ValueError:
                    return False
        with ThreadPoolExecutor(max_workers=4) as pool:
            self.assertEqual(sum(pool.map(claim, range(8))), 1)

    def test_concurrent_allocations_are_unique_and_monotonic(self):
        batch = _read_sources(self.paths)[0]
        with self.history():
            pass  # finish schema setup before independent workers contend
        def allocate(index):
            local = dict(batch, case_id=str(index), history_id=str(index))
            with self.history() as history:
                return history.allocate(local, self.settings, 62, TODAY)["certificate_no"]
        with ThreadPoolExecutor(max_workers=6) as pool:
            numbers = list(pool.map(allocate, range(12)))
        self.assertEqual(sorted(numbers), [f"{n:04d}" for n in range(62, 74)])

    def test_concurrent_same_batch_has_one_reservation(self):
        batch = _read_sources(self.paths)[0]
        with self.history():
            pass
        def allocate(_):
            with self.history() as history:
                return history.allocate(batch, self.settings, 62, TODAY)["certificate_no"]
        with ThreadPoolExecutor(max_workers=4) as pool:
            self.assertEqual(set(pool.map(allocate, range(8))), {"0062"})
        with self.history() as history:
            self.assertEqual(len(history.records()), 1)


class OutlookComContracts(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pdf = Path(self.temp.name) / "dummy.pdf"
        self.banner = Path(self.temp.name) / "banner.png"
        # Empty files suffice: the COM boundary, not rendering, is under test.
        self.pdf.touch()
        self.banner.touch()
        self.message = Mock()
        self.message.EntryID = "DUMMY-ENTRY"
        self.pdf_attachment = Mock()
        self.banner_attachment = Mock()
        self.pdf_attachment.PropertyAccessor.GetProperty.return_value = False

        def add_attachment(source, *_args):
            return self.pdf_attachment if str(source).endswith(self.pdf.name) else self.banner_attachment

        self.message.Attachments.Add.side_effect = add_attachment
        self.folder = Mock()
        self.folder.StoreID = "DUMMY-STORE"
        self.folder.EntryID = "DUMMY-DRAFT-FOLDER"
        self.folder.Items.Add.return_value = self.message
        self.account = Mock()
        self.account.SmtpAddress = "dummy@example.test"
        self.account.DeliveryStore.GetDefaultFolder.return_value = self.folder
        self.app = Mock()
        self.app.Session.Accounts = [self.account]
        self.saved = Mock()
        self.saved.Subject = "Dummy COD"
        self.saved.Sent = False
        self.saved.Parent.EntryID = self.folder.EntryID
        self.saved.UserProperties.Find.return_value.Value = "plant2|1"
        self.saved_pdf = Mock()
        self.saved_pdf.FileName = self.pdf.name
        self.saved_pdf.PropertyAccessor.GetProperty.return_value = False
        self.saved_banner = types.SimpleNamespace(FileName=self.banner.name)
        self.saved.Attachments = [self.saved_pdf, self.saved_banner]
        self.app.Session.GetItemFromID.return_value = self.saved
        client = types.ModuleType("win32com.client")
        client.Dispatch = Mock(return_value=self.app)
        module = types.ModuleType("win32com")
        module.client = client
        self.modules = patch.dict(sys.modules, {"win32com": module, "win32com.client": client})
        self.modules.start()
        self.addCleanup(self.modules.stop)
        self.batch = {"case_id": "1", "history_id": "plant2|1", "email": "client@example.test"}

    def test_save_contract_account_folder_pdf_banner_and_never_display_send(self):
        writer = OutlookWriter(" DUMMY@example.test ")
        entry = writer.save_draft(self.batch, "Dummy COD", "<html>Dummy</html>", self.pdf, self.banner)
        self.assertEqual(entry, "DUMMY-ENTRY")
        self.account.DeliveryStore.GetDefaultFolder.assert_called_once_with(16)
        self.assertFalse(self.message.Send.called)
        self.assertEqual(self.message.To, self.batch["email"])
        self.assertEqual(self.message.BodyFormat, 2)
        self.assertEqual(self.message.HTMLBody, "<html>Dummy</html>")
        self.message.UserProperties.Add.assert_called_once_with("EcorecoCertificateKey", 1)
        self.assertEqual(self.message.UserProperties.Add.return_value.Value, "plant2|1")
        self.assertEqual(self.message.Attachments.Add.call_args_list[0].args, (str(self.pdf.resolve()), 1))
        self.assertEqual(self.message.Attachments.Add.call_args_list[1].args, (str(self.banner.resolve()), 1, 0))
        self.pdf_attachment.PropertyAccessor.SetProperty.assert_called_once_with(
            "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", False)
        self.banner_attachment.PropertyAccessor.SetProperty.assert_called_once_with(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ecoreco-banner")
        self.message.Save.assert_called_once_with()
        self.message.Display.assert_not_called()
        self.message.Send.assert_not_called()
        self.app.Session.GetItemFromID.assert_called_once_with("DUMMY-ENTRY", "DUMMY-STORE")

    def test_readback_transient_failure_retries_reads_not_save(self):
        self.app.Session.GetItemFromID.side_effect = [RuntimeError("busy"), self.saved]
        with patch("draft_service.time.sleep"):
            OutlookWriter("dummy@example.test").save_draft(self.batch, "Dummy COD", "body", self.pdf, self.banner)
        self.assertEqual(self.app.Session.GetItemFromID.call_count, 2)
        self.message.Save.assert_called_once()
        self.message.Display.assert_not_called()
        self.message.Send.assert_not_called()

    def test_readback_missing_attachment_does_not_claim_confirmed_save(self):
        self.saved.Attachments = []
        with patch("draft_service.time.sleep"), self.assertRaisesRegex(RuntimeError, "could not be verified"):
            OutlookWriter("dummy@example.test").save_draft(self.batch, "Dummy COD", "body", self.pdf, self.banner)
        self.assertEqual(self.app.Session.GetItemFromID.call_count, 3)
        self.message.Save.assert_called_once()
        self.message.Send.assert_not_called()

    def test_readback_hidden_pdf_does_not_claim_confirmed_save(self):
        self.saved_pdf.PropertyAccessor.GetProperty.return_value = True
        with patch("draft_service.time.sleep"), self.assertRaisesRegex(RuntimeError, "could not be verified"):
            OutlookWriter("dummy@example.test").save_draft(self.batch, "Dummy COD", "body", self.pdf, self.banner)
        self.assertEqual(self.app.Session.GetItemFromID.call_count, 3)
        self.message.Save.assert_called_once()
        self.message.Send.assert_not_called()

    def test_readback_wrong_batch_or_wrong_folder_is_not_accepted(self):
        for wrong in ("key", "folder", "sent"):
            with self.subTest(wrong=wrong):
                self.saved.UserProperties.Find.return_value.Value = "wrong" if wrong == "key" else "plant2|1"
                self.saved.Parent.EntryID = "wrong" if wrong == "folder" else self.folder.EntryID
                self.saved.Sent = wrong == "sent"
                with patch("draft_service.time.sleep"), self.assertRaisesRegex(RuntimeError, "could not be verified"):
                    OutlookWriter("dummy@example.test").save_draft(self.batch, "Dummy COD", "body", self.pdf, self.banner)
        self.message.Display.assert_not_called()
        self.message.Send.assert_not_called()

    def test_missing_banner_stops_before_item_creation(self):
        writer = OutlookWriter("dummy@example.test")
        with self.assertRaises(FileNotFoundError):
            writer.save_draft(self.batch, "dummy", "dummy", self.pdf, self.banner.with_name("missing.png"))
        self.folder.Items.Add.assert_not_called()

    def test_unavailable_or_ambiguous_mailbox_stops_before_item_creation(self):
        for accounts in ([], [self.account, self.account]):
            with self.subTest(accounts=len(accounts)):
                self.app.Session.Accounts = accounts
                with self.assertRaisesRegex(RuntimeError, "selected Outlook mailbox"):
                    OutlookWriter("dummy@example.test")
        self.folder.Items.Add.assert_not_called()

    def test_draft_does_not_depend_on_send_account_override(self):
        class RejectAccount:
            def __init__(self, delegate):
                object.__setattr__(self, "delegate", delegate)

            def __getattr__(self, name):
                return getattr(self.delegate, name)

            @property
            def SendUsingAccount(self):
                return None
            @SendUsingAccount.setter
            def SendUsingAccount(self, value):
                raise RuntimeError("account cannot be set")
        message = RejectAccount(self.message)
        self.folder.Items.Add.return_value = message
        writer = OutlookWriter("dummy@example.test")
        self.assertEqual(writer.save_draft(self.batch, "Dummy COD", "dummy", self.pdf), "DUMMY-ENTRY")
        message.Save.assert_called_once_with()

    def test_save_error_and_empty_entry_never_display_or_send(self):
        writer = OutlookWriter("dummy@example.test")
        self.message.Save.side_effect = RuntimeError("Save result uncertain")
        with self.assertRaisesRegex(RuntimeError, "uncertain"):
            writer.save_draft(self.batch, "dummy", "dummy", self.pdf)
        self.message.Save.side_effect = None
        for value in (None, "", " "):
            self.message.EntryID = value
            with self.assertRaisesRegex(RuntimeError, "no draft ID"):
                writer.save_draft(self.batch, "dummy", "dummy", self.pdf)
        self.message.Display.assert_not_called()
        self.message.Send.assert_not_called()


if __name__ == "__main__":
    unittest.main()
