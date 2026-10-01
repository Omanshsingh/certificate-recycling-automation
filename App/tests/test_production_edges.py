"""Production boundary regressions; disposable workspaces and no HTTP servers."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path
import shutil
import tempfile
import threading
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from reportlab.pdfgen.canvas import Canvas

import certificate_drafts as drafts
import issuance_safety as safety
import run_certificates as runner
import web_app
import workbook_store
from tests.test_certificate_templates import write_register
from workflow_support import save_json
from workspace_review import workspace_revision


AS_OF = date(2026, 10, 1)
APP = Path(__file__).resolve().parents[1]


class SimulatedPowerLoss(BaseException):
    """Bypass ordinary rollback just as an abruptly terminated process would."""


def small_pdf(path, text="Certificate No. 0001"):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    canvas = Canvas(str(path))
    canvas.drawString(40, 100, text)
    canvas.save()


class ProductionEdgeTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        write_register(self.paths[0])
        write_register(self.paths[1], case_id="OTHER", cod="0040", status="N/A")
        native = self.root / "App" / "templates" / "certificate-types"
        native.mkdir(parents=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf", native / "standard.pdf")
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "reminders_enabled": False,
        })
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "BACKUP_DIR", self.root / "Data" / "Workbook Backups"))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", self.root / "Data" / "Reports" / "Latest Results.json"))
        self.enterContext(patch.object(web_app, "STOPPING", False))
        self.enterContext(patch.object(runner, "make_certificate", side_effect=self.render_workbook))
        self.enterContext(patch.object(drafts, "make_template_certificate", side_effect=self.render_manual))

    @staticmethod
    def render_workbook(batch, cod, issued_date, template, destination, **kwargs):
        small_pdf(destination, f"Certificate No. {cod} - {batch['supplier']}")

    @staticmethod
    def render_manual(batch, cod, issued_date, destination, **kwargs):
        small_pdf(destination, f"Certificate No. {cod} - {batch['supplier']}")

    @staticmethod
    def details(**changes):
        values = dict(case_id="CLIENT", supplier="Reviewed Company", location="Pune",
                      inward_date="01.09.2026", item="Servers", quantity="2",
                      per_piece_weight="5", weight="10", form6="FORM-1",
                      reference="MANUAL-UNIQUE", certificate_type="standard")
        values.update(changes)
        return values

    def run_selected(self, **kwargs):
        return runner.run(self.paths, self.root, selected=("plant1", "CLIENT"),
                          as_of=AS_OF, progress=None, **kwargs)

    def tx(self, destination=None, old_paths=()):
        source = drafts.register_path(self.root)
        save_json(source, {"version": 1, "records": []})
        temporary = self.root / "work.pdf"
        small_pdf(temporary)
        return safety.IssuanceTransaction(
            self.root, destination or self.root / "Certificates" / "0001_New_Pune_1.pdf",
            temporary, source, source.read_bytes(), source_kind="manual", plant="plant1",
            key="record", cod="0001", issued_date=AS_OF, before={}, after={}, old_paths=old_paths)

    def test_workbook_crash_immediately_after_source_replace_recovers_as_committed(self):
        replace = workbook_store.os.replace
        def crash_after_replace(source, destination):
            result = replace(source, destination)
            if Path(destination).resolve() == self.paths[0].resolve():
                # Only crash at the issuance commit, not alert-format maintenance.
                if safety.audit_records(self.root):
                    raise SimulatedPowerLoss()
            return result
        with patch.object(workbook_store.os, "replace", side_effect=crash_after_replace):
            with self.assertRaises(SimulatedPowerLoss):
                self.run_selected()
        record = safety.audit_records(self.root)[0]
        self.assertEqual(record["source_expected"], safety.file_digest(self.paths[0]))
        self.assertTrue(safety.recover_issuances(self.root))
        self.assertEqual(safety.audit_records(self.root)[0]["state"], "complete")
        self.assertIsNotNone(runner.find_certificate(self.root / "Certificates", {
            "source_id": "plant1", "case_id": "CLIENT", "cod_no": record["cod"]}))

    def test_manual_crash_immediately_after_register_replace_recovers_as_committed(self):
        replace = workbook_store.os.replace
        def crash_after_replace(source, destination):
            result = replace(source, destination)
            if Path(destination).resolve() == drafts.register_path(self.root).resolve():
                raise SimulatedPowerLoss()
            return result
        with patch.object(workbook_store.os, "replace", side_effect=crash_after_replace):
            with self.assertRaises(SimulatedPowerLoss):
                drafts.generate_draft(self.root, "plant1", self.details(), as_of=AS_OF)
        record = safety.audit_records(self.root)[0]
        self.assertEqual(record["source_expected"], safety.file_digest(drafts.register_path(self.root)))
        safety.recover_issuances(self.root)
        self.assertEqual(safety.audit_records(self.root)[0]["state"], "complete")
        self.assertEqual(drafts.read_register(self.root)["records"][0]["status"], "generated_manual")

    def test_rollback_preserves_external_edit_to_published_pdf(self):
        tx = self.tx()
        tx.publish()
        destination = self.root / tx.data["pdf_path"]
        small_pdf(destination, "An outside editor changed this PDF")
        external = destination.read_bytes()
        with self.assertRaisesRegex(ValueError, "changed outside"):
            safety.recover_issuances(self.root)
        self.assertEqual(destination.read_bytes(), external)
        self.assertEqual(safety.audit_records(self.root)[0]["state"], "published")

    def test_committed_recovery_preserves_external_edit_to_old_pdf(self):
        old = self.root / "Certificates" / "0001_Old_Pune_1.pdf"
        small_pdf(old, "Original owned certificate")
        tx = self.tx(old_paths=[old])
        tx.publish()
        tx.expect_source(b"committed source")
        drafts.register_path(self.root).write_bytes(b"committed source")
        small_pdf(old, "Externally changed former certificate")
        external = old.read_bytes()
        with self.assertRaisesRegex(ValueError, "older PDF changed"):
            safety.recover_issuances(self.root)
        self.assertEqual(old.read_bytes(), external)
        self.assertTrue((self.root / tx.data["pdf_path"]).exists())

    def test_form_creation_still_reports_success_when_postcommit_refresh_fails(self):
        body = dict(plant="plant1", details=self.details(), edit_scope="update_excel")
        with patch.object(web_app, "refresh", side_effect=OSError("Excel reopened after commit")):
            try:
                response = web_app.create_certificate(body)
            except OSError as exc:
                self.assertEqual(safety.audit_records(self.root)[0]["state"], "complete")
                self.fail(f"Committed certificate incorrectly reported as failed: {exc}")
        self.assertEqual(response["created"], 1)
        self.assertEqual(len(response["certificates"]), 1)

    def test_generate_still_reports_success_when_postcommit_refresh_fails(self):
        with patch.object(web_app, "refresh", side_effect=OSError("Excel reopened after commit")):
            try:
                response = web_app.generate(selected=("plant1", "CLIENT"))
            except OSError as exc:
                self.assertEqual(safety.audit_records(self.root)[0]["state"], "complete")
                self.fail(f"Committed certificate incorrectly reported as failed: {exc}")
        self.assertEqual(response["created"], 1)

    def test_generate_preflight_and_issuance_both_respect_selected_plant(self):
        write_register(self.paths[1], case_id="OTHER", status="Ready")

        response = web_app.generate(plants=["plant2"])

        self.assertEqual(response["created"], 1)
        first = load_workbook(self.paths[0], data_only=True)
        second = load_workbook(self.paths[1], data_only=True)
        try:
            self.assertEqual((first.active["M2"].value, first.active["K2"].value), ("Ready", None))
            self.assertEqual(second.active["M2"].value, "Generated")
            self.assertTrue(second.active["K2"].value)
        finally:
            first.close()
            second.close()
        pdfs = list((self.root / "Certificates").glob("*.pdf"))
        self.assertEqual(len(pdfs), 1)
        self.assertTrue(pdfs[0].name.endswith("_2.pdf"))

    def test_bulk_generate_endpoint_forwards_plant_scope(self):
        body = {"plants": ["plant2"], "expected_revision": "revision-1"}
        handler = object.__new__(web_app.Handler)
        handler.path = "/api/generate"
        payload = json.dumps(body).encode("utf-8")
        handler.headers = {"Content-Length": str(len(payload))}
        handler.rfile = BytesIO(payload)
        responses = []
        handler._same_origin = lambda: True
        handler.send_json = lambda value, status=200: responses.append((status, value))
        with patch.object(web_app, "reviewed_operation",
                          side_effect=lambda root, endpoint, values, action: action(values)), \
                patch.object(web_app, "generate", return_value={"created": 0}) as generate:
            handler._do_post()
        generate.assert_called_once()
        self.assertEqual(generate.call_args.kwargs["plants"], ["plant2"])
        self.assertEqual(responses[0][0], 200)

    def test_generate_rejects_external_change_between_preflight_and_issuance(self):
        original_run = web_app.run
        def run_then_external_change(*args, **kwargs):
            result = original_run(*args, **kwargs)
            if kwargs.get("check_only"):
                book = load_workbook(self.paths[0])
                try:
                    book.active["C2"] = "Unreviewed external supplier"
                    book.save(self.paths[0])
                finally:
                    book.close()
            return result
        with patch.object(web_app, "run", side_effect=run_then_external_change):
            with self.assertRaisesRegex(ValueError, "files changed"):
                web_app.generate(selected=("plant1", "CLIENT"), expected_revision=workspace_revision(self.root))
        self.assertFalse(safety.audit_records(self.root))

    def test_manual_endpoint_retains_success_if_postcommit_state_read_fails(self):
        body = dict(plant="plant1", details=self.details(), expected_revision=workspace_revision(self.root))
        handler = object.__new__(web_app.Handler)
        handler.path = "/api/manual/generate"
        payload = json.dumps(body).encode("utf-8")
        handler.headers = {"Content-Length": str(len(payload))}
        handler.rfile = BytesIO(payload)
        responses = []
        handler._same_origin = lambda: True
        handler.send_json = lambda value, status=200: responses.append((status, value))
        with patch.object(web_app, "workspace_state", side_effect=OSError("postcommit dashboard read failed")):
            handler._do_post()
        self.assertEqual(safety.audit_records(self.root)[0]["state"], "complete")
        self.assertEqual(responses[0][0], 200, responses)
        self.assertEqual(responses[0][1]["created"], 1)

    def test_finalsuccess_metadata_uses_pending_current_receipt_when_completion_fails(self):
        body = dict(plant="plant1", details=self.details(), edit_scope="update_excel")
        with patch.object(safety.IssuanceTransaction, "complete", side_effect=OSError("history disk unavailable")), \
                patch.object(web_app, "refresh", return_value={"rows": []}):
            response = web_app.create_certificate(body)
        self.assertEqual(response["created"], 1)
        self.assertEqual(len(response["certificates"]), 1,
                         "A source-committed PDF needs its current filename/link even when final history needs recovery")

    def test_bulk_save_does_not_strand_first_recovery_when_next_row_changes_same_workbook(self):
        book = load_workbook(self.paths[0])
        try:
            book.active.append(["SECOND", date(2026, 9, 1), "Second Client", "Pune",
                                "Servers", 2, 5, 10, "FORM-2", "REF-2", "", None, "Ready"])
            book.save(self.paths[0])
        finally:
            book.close()
        finish = safety.IssuanceTransaction._finish_committed
        attempts = []
        def fail_first_completion(transaction):
            attempts.append(transaction.id)
            if len(attempts) == 1:
                raise OSError("Final inventory temporarily unavailable")
            return finish(transaction)
        with patch.object(safety.IssuanceTransaction, "_finish_committed", fail_first_completion):
            result = runner.run(self.paths, self.root, as_of=AS_OF, progress=None)
        records = safety.audit_records(self.root)
        self.assertGreaterEqual(result[0], 1)
        try:
            safety.recover_issuances(self.root)
        except ValueError as exc:
            self.fail(f"The app's own later bulk commit stranded its earlier recovery: {exc}; "
                      f"created={result[0]}, states={[row['state'] for row in records]}")
        self.assertTrue(all(row["state"] == "complete" for row in safety.audit_records(self.root)))

    def test_renamed_same_cod_reissue_matches_only_latest_current_pdf(self):
        self.assertEqual(self.run_selected()[0], 1)
        first = safety.audit_records(self.root)[0]
        original = self.root / first["pdf_path"]
        original_bytes = original.read_bytes()
        details = self.details(supplier="Renamed Client", location="Mumbai", cod_no=first["cod"])
        self.assertEqual(self.run_selected(details=details)[0], 1)
        latest = safety.latest_issuance(self.root, "workbook", "plant1", "CLIENT")
        current = self.root / latest["pdf_path"]
        self.assertNotEqual(current, original)
        self.assertFalse(original.exists())
        archived = next(row for row in latest["previous_pdfs"] if row["path"] == first["pdf_path"])
        self.assertEqual((self.root / archived["archive"]).read_bytes(), original_bytes)
        batch = {"source_id": "plant1", "case_id": "CLIENT", "cod_no": first["cod"]}
        self.assertEqual(runner.find_certificate(self.root / "Certificates", batch), current)
        # An old copy returning to the folder cannot mask a missing current revision.
        current.unlink()
        original.write_bytes(original_bytes)
        self.assertIsNone(runner.find_certificate(self.root / "Certificates", batch))

    def test_recovery_does_not_retire_old_pdf_into_partially_written_archive(self):
        old = self.root / "Certificates" / "0001_Old_Pune_1.pdf"
        small_pdf(old, "Old certificate to preserve")
        before = old.read_bytes()
        tx = self.tx(old_paths=[old])
        tx.publish()
        tx.expect_source(b"committed source")
        drafts.register_path(self.root).write_bytes(b"committed source")
        archive = self.root / tx.data["previous_pdfs"][0]["archive"]
        archive.parent.mkdir(parents=True)
        # Simulate a crash partway through the non-atomic archive write.
        archive.write_bytes(before[:100])
        try:
            safety.recover_issuances(self.root)
        except (ValueError, OSError):
            pass
        self.assertTrue((archive.is_file() and archive.read_bytes() == before)
                        or (old.is_file() and old.read_bytes() == before),
                        "Recovery removed the old PDF while its archive was only a partial write")

    def test_stop_waits_for_active_operation_and_rejects_later_writes(self):
        entered, release, stop_started, stopped = (threading.Event() for _ in range(4))
        errors = []
        responses = []
        lock = threading.RLock()
        active = object.__new__(web_app.Handler)
        def active_operation():
            entered.set()
            if not release.wait(2):
                raise AssertionError("Test did not release its temporary operation")
        active._do_post = active_operation
        stop = object.__new__(web_app.Handler)
        stop.path = "/api/stop"
        stop.headers = {"Content-Length": "2"}
        stop.rfile = BytesIO(b"{}")
        stop._same_origin = lambda: True
        stop.send_json = lambda value, status=200: responses.append((status, value))
        class FakeServer:
            def shutdown(self):
                stopped.set()
        def invoke(handler, starting=None):
            try:
                if starting:
                    starting.set()
                handler.do_POST()
            except BaseException as exc:
                errors.append(exc)
        with patch.object(web_app, "OPERATION_LOCK", lock), patch.object(web_app, "SERVER", FakeServer()):
            first = threading.Thread(target=invoke, args=(active,))
            closing = threading.Thread(target=invoke, args=(stop, stop_started))
            try:
                first.start()
                self.assertTrue(entered.wait(1))
                closing.start()
                self.assertTrue(stop_started.wait(1))
                self.assertFalse(stopped.is_set())
                self.assertFalse(web_app.STOPPING)
            finally:
                release.set()
                first.join(3)
                if closing.ident is not None:
                    closing.join(3)
            self.assertFalse(first.is_alive())
            self.assertFalse(closing.is_alive())
            self.assertFalse(errors, errors)
            self.assertTrue(stopped.wait(1))
            later = object.__new__(web_app.Handler)
            later.send_error_json = lambda message, status=400: responses.append((status, {"error": message}))
            later._do_post = lambda: self.fail("A write started after stop")
            later.do_POST()
            self.assertEqual(responses[-1][0], 503)


if __name__ == "__main__":
    unittest.main()
