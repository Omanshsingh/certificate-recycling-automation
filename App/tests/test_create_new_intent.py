"""New-versus-edit intent checks using disposable workbooks and PDFs only."""
from __future__ import annotations

from datetime import date
from io import BytesIO
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from certificate_core import read_batches
from issuance_safety import audit_records
import run_certificates as runner
import web_app
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_templates import write_register


APP = Path(__file__).resolve().parents[1]


class CreateNewIntentTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(
            prefix="certificate-new-intent-")))
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        write_register(self.paths[0])
        write_register(self.paths[1], case_id="OTHER", cod="0040", status="N/A")
        templates = self.root / "App" / "templates" / "certificate-types"
        templates.mkdir(parents=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf",
                     templates / "standard.pdf")
        self.results = self.root / "Data" / "Reports" / "Latest Results.json"
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "certificate_number_prefix": "", "certificate_number_start": 1,
            "certificate_number_width": 4, "reminder_time": "11:00",
            "reminders_enabled": False,
        })
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "BACKUP_DIR", self.root / "Data" / "Workbook Backups"))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", self.results))
        self.enterContext(patch.object(web_app, "STOPPING", False))
        self.enterContext(patch.object(web_app.subprocess, "run",
                                      side_effect=AssertionError("No Windows task or process")))
        # Issuance runs for real in the disposable root; dashboard rereads are unnecessary.
        self.enterContext(patch.object(web_app, "refresh", side_effect=lambda *_: json.loads(
            self.results.read_text(encoding="utf-8"))))

    def body(self, case_id="CLIENT", plant="plant1", **changes):
        body = {
            "plant": plant, "create_new": True, "edit_scope": "update_excel",
            "details": {
                "case_id": case_id, "supplier": "Requested Company Ltd", "location": "Pune",
                "inward_date": "2026-09-01", "item": "Servers", "weight": "10",
                "form6": "FORM-REQUEST", "reference": "REF-REQUEST",
                "cod_date": "2026-10-01", "certificate_type": "standard",
                "certificate_fields": {},
            },
            "expected_revision": workspace_revision(self.root),
        }
        body.update(changes)
        return body

    def post_create(self, body):
        handler = object.__new__(web_app.Handler)
        handler.path = "/api/create"
        payload = json.dumps(body).encode("utf-8")
        handler.headers = {"Host": "127.0.0.1", "Content-Length": str(len(payload))}
        handler.rfile = BytesIO(payload)
        responses = []
        handler.send_json = lambda value, status=200: responses.append((status, value))
        handler.do_POST()  # Exercise the real request lock and revision/confirmation boundary.
        self.assertEqual(len(responses), 1)
        return responses[0]

    def edit_cell(self, path, coordinate, value):
        book = load_workbook(path)
        try:
            book.active[coordinate] = value
            book.save(path)
        finally:
            book.close()

    def append_external_row(self, path, case_id):
        book = load_workbook(path)
        try:
            book.active.append([case_id, date(2026, 9, 1), "External company", "Pune",
                                "Servers", 2, 5, 10, "FORM-EXTERNAL", "REF-EXTERNAL",
                                "", None, "Blank"])
            book.save(path)
        finally:
            book.close()

    def snapshot(self):
        return {str(path.relative_to(self.root)): path.read_bytes()
                for path in self.root.rglob("*") if path.is_file()}

    def assert_no_issuance(self):
        self.assertFalse(list((self.root / "Certificates").glob("*.pdf")))
        self.assertFalse(list((self.root / "Data" / "Workbook Backups").glob("*.xlsx")))

    def test_new_duplicate_is_rejected_for_every_excel_status_before_any_write(self):
        for excel_status in ("Ready", "", "Blank", "N/A", "Generated", "Completed"):
            with self.subTest(status=excel_status):
                self.edit_cell(self.paths[0], "M2", excel_status)
                body = self.body()
                before = self.snapshot()
                with patch.object(web_app, "run", side_effect=AssertionError("No preflight")) as run:
                    status, response = self.post_create(body)
                self.assertEqual(status, 400, response)
                self.assertIn("This Excel ID already exists", response["error"])
                self.assertIn("Choose the company from the list", response["error"])
                run.assert_not_called()
                self.assertEqual(self.snapshot(), before)
                self.assert_no_issuance()

    def test_duplicate_uses_existing_trimmed_text_and_numeric_case_keys(self):
        for saved_id, typed_id in ((" CLIENT ", "  CLIENT  "), (4018, " 4018 ")):
            with self.subTest(saved_id=saved_id, typed_id=typed_id):
                self.edit_cell(self.paths[0], "A2", saved_id)
                body = self.body(case_id=typed_id)
                before = self.snapshot()
                with patch.object(web_app, "run", side_effect=AssertionError("No preflight")) as run:
                    status, response = self.post_create(body)
                self.assertEqual(status, 400, response)
                self.assertIn("This Excel ID already exists", response["error"])
                run.assert_not_called()
                self.assertEqual(self.snapshot(), before)

    def test_duplicate_check_reads_the_selected_plant_not_cached_queue(self):
        save_json(self.results, {"rows": []})
        body = self.body(case_id="OTHER", plant="plant2")
        before = self.snapshot()
        with patch.object(web_app, "latest_result", side_effect=AssertionError("No cached queue")), \
                patch.object(web_app, "run", side_effect=AssertionError("No preflight")) as run, \
                patch.object(web_app, "read_batches", wraps=read_batches) as read:
            status, response = self.post_create(body)
        self.assertEqual(status, 400, response)
        self.assertIn("This Excel ID already exists", response["error"])
        read.assert_called_once_with(self.paths[1].resolve())
        run.assert_not_called()
        self.assertEqual(self.snapshot(), before)

    def test_new_unique_id_can_match_an_id_in_the_other_plant(self):
        other_before = self.paths[1].read_bytes()
        status, response = self.post_create(self.body(case_id="OTHER"))
        self.assertEqual(status, 200, response)
        self.assertEqual(response["created"], 1)
        batches = read_batches(self.paths[0])
        self.assertEqual([batch["case_id"] for batch in batches], ["CLIENT", "OTHER"])
        self.assertEqual(batches[0]["supplier"], "Independent Client Ltd")
        self.assertEqual(batches[1]["supplier"], "Requested Company Ltd")
        self.assertEqual(batches[1]["issued"], "Generated")
        self.assertEqual(self.paths[1].read_bytes(), other_before)
        self.assertEqual(len(list((self.root / "Certificates").rglob("*.pdf"))), 1)

    def test_existing_source_edit_remains_allowed(self):
        self.assertEqual(runner.run(self.paths, self.root, selected=('plant1', 'CLIENT'), progress=None)[0], 1)
        status, response = self.post_create(self.body(
            create_new=False, source_sn="CLIENT", source_plant="plant1"))
        self.assertEqual(status, 200, response)
        self.assertEqual(response["created"], 1)
        batches = read_batches(self.paths[0])
        self.assertEqual(len(batches), 1)
        self.assertEqual(batches[0]["case_id"], "CLIENT")
        self.assertEqual(batches[0]["supplier"], "Requested Company Ltd")
        self.assertEqual(batches[0]["issued"], "Generated")

    def test_legacy_request_without_create_new_or_source_sn_still_edits_existing(self):
        self.assertEqual(runner.run(self.paths, self.root, selected=('plant1', 'CLIENT'), progress=None)[0], 1)
        body = self.body()
        del body["create_new"]
        status, response = self.post_create(body)
        self.assertEqual(status, 200, response)
        self.assertEqual(response["created"], 1)
        batches = read_batches(self.paths[0])
        self.assertEqual(len(batches), 1)
        self.assertEqual(batches[0]["supplier"], "Requested Company Ltd")

    def test_stale_revision_is_rejected_before_duplicate_lookup(self):
        body = self.body()
        self.edit_cell(self.paths[0], "C2", "External change")
        before = self.snapshot()
        with patch.object(web_app, "read_batches", side_effect=AssertionError("No stale lookup")) as read, \
                patch.object(web_app, "run", side_effect=AssertionError("No stale preflight")) as run:
            status, response = self.post_create(body)
        self.assertEqual(status, 400, response)
        self.assertIn("files changed", response["error"])
        read.assert_not_called()
        run.assert_not_called()
        self.assertEqual(self.snapshot(), before)

    def test_external_duplicate_after_lookup_is_rejected_by_runner_revision(self):
        body = self.body(case_id="NEW")
        external_bytes = []

        def read_then_external_change(path):
            batches = read_batches(path)
            self.append_external_row(path, "NEW")
            external_bytes.append(path.read_bytes())
            return batches

        with patch.object(web_app, "read_batches", side_effect=read_then_external_change), \
                patch.object(web_app, "run", wraps=runner.run) as run, \
                patch.object(runner, "stage_certificate_details", side_effect=AssertionError("No stale staging")):
            status, response = self.post_create(body)
        self.assertEqual(status, 400, response)
        self.assertIn("files changed", response["error"])
        run.assert_not_called()  # The form's new source check catches the stale revision before staging.
        self.assertEqual(self.paths[0].read_bytes(), external_bytes[0])
        self.assertFalse(self.results.exists())
        self.assert_no_issuance()

    def test_external_duplicate_after_preflight_is_rejected_before_issuance(self):
        body = self.body(case_id="NEW")
        external_bytes = []

        def run_then_external_change(*args, **kwargs):
            result = runner.run(*args, **kwargs)
            if kwargs.get("check_only"):
                self.append_external_row(self.paths[0], "NEW")
                external_bytes.append(self.paths[0].read_bytes())
            return result

        with patch.object(web_app, "run", side_effect=run_then_external_change) as run, \
                patch.object(runner, "make_certificate", side_effect=AssertionError("No stale PDF")):
            status, response = self.post_create(body)
        self.assertEqual(status, 400, response)
        self.assertIn("files changed", response["error"])
        self.assertEqual(run.call_count, 2)
        self.assertTrue(all(call.kwargs["expected_revision"] == body["expected_revision"]
                            for call in run.call_args_list))
        self.assertEqual(self.paths[0].read_bytes(), external_bytes[0])
        self.assert_no_issuance()

    def test_new_intent_is_preserved_through_warning_confirmation(self):
        body = self.body(case_id="NEW")
        body["details"]["inward_date"] = ""
        before = [path.read_bytes() for path in self.paths]
        status, review = self.post_create(body)
        self.assertEqual(status, 200, review)
        self.assertTrue(review["requires_confirmation"])
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        self.assert_no_issuance()
        body.update(confirmed=True, review_token=review["review_token"],
                    expected_revision=review["workspace_revision"])
        status, response = self.post_create(body)
        self.assertEqual(status, 200, response)
        self.assertEqual(response["created"], 1)
        self.assertEqual([batch["case_id"] for batch in read_batches(self.paths[0])], ["CLIENT", "NEW"])
        record = next(record for record in audit_records(self.root) if record['key'] == 'NEW')
        self.assertTrue(record['accepted_warnings'])
        self.assertEqual(record['override_reason'], 'Warnings reviewed; user chose Continue anyway.')


if __name__ == "__main__":
    unittest.main()
