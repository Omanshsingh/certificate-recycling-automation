from __future__ import annotations

from io import BytesIO
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

import web_app
from workflow_support import load_settings, save_json


class SettingsFeedbackTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(
            prefix="certificate-settings-feedback-")))
        self.paths = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsx"]
        for path in self.paths:
            path.touch()  # This endpoint validates paths, not workbook contents.
        self.settings_path = self.root / "settings.local.json"
        self.original_settings = {
            "workbook_paths": [str(path) for path in self.paths],
            "reminders_enabled": True,
            "reminder_time": "11:00",
            "certificate_number_prefix": "REC",
            "certificate_number_start": 77,
            "certificate_number_width": 6,
            "backup_destination": str(self.root / "Backups"),
        }
        save_json(self.settings_path, self.original_settings)
        self.original_bytes = self.settings_path.read_bytes()
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "STOPPING", False))
        # Every test mocks the scheduler, including validation and write failures.
        self.scheduler = self.enterContext(patch.object(
            web_app.subprocess, "run",
            return_value=subprocess.CompletedProcess([], 0, "", "")))
        self.settings_writer = self.enterContext(patch.object(
            web_app, "save_json", wraps=save_json))

    def body(self, **changes):
        body = {
            "workbook_paths": [str(path) for path in self.paths],
            "reminders_enabled": True,
            "reminder_time": "14:16",
            "update_reminder": True,
        }
        body.update(changes)
        return body

    def post_settings(self, body):
        return self.post_request("/api/settings", body)

    def post_request(self, path, body):
        # Exercise request dispatch and its error handling without a live server.
        handler = object.__new__(web_app.Handler)
        handler.path = path
        payload = json.dumps(body).encode("utf-8")
        handler.headers = {
            "Host": "127.0.0.1",
            "Content-Length": str(len(payload)),
        }
        handler.rfile = BytesIO(payload)
        responses = []
        handler.send_json = lambda value, status=200: responses.append((status, value))
        handler.do_POST()
        self.assertEqual(len(responses), 1)
        return responses[0]

    def assert_saved(self, response, body):
        saved = load_settings(self.root)
        self.assertEqual(saved["workbook_paths"], [
            str(Path(path).expanduser().resolve()) for path in body["workbook_paths"]])
        self.assertEqual(saved["reminders_enabled"], body["reminders_enabled"])
        self.assertEqual(saved["reminder_time"], body["reminder_time"])
        for key in ("certificate_number_prefix", "certificate_number_start",
                    "certificate_number_width", "backup_destination"):
            self.assertEqual(saved[key], self.original_settings[key])
        self.assertEqual(response["settings"], {
            key: saved[key] for key in (
                "workbook_paths", "reminders_enabled", "reminder_time", "backup_destination")})
        self.assertNotIn("error", response)

    def assert_scheduler_command(self, enabled=True, reminder_time="14:16"):
        command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(self.root / "App" / "Install Daily Task.ps1"),
            "-Mode", "Install" if enabled else "Disable",
        ]
        if enabled:
            command.extend(["-ReminderTime", reminder_time])
        self.scheduler.assert_called_once_with(
            command, capture_output=True, text=True, timeout=60, check=False,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))

    def test_invalid_settings_do_not_write_or_schedule(self):
        cases = [
            ({"workbook_paths": "not a list"}, "Choose a file"),
            ({"workbook_paths": [str(self.paths[0])]}, "Choose a file"),
            ({"workbook_paths": [str(self.paths[0]), None]}, "Choose a file"),
            ({"workbook_paths": [str(self.paths[0]), str(self.root / "Missing.xlsx")]},
             "could not be found"),
            ({"workbook_paths": [str(self.paths[0]), str(self.root)]},
             "could not be found"),
            ({"workbook_paths": [str(self.paths[0]), str(self.paths[0])]},
             "different workbooks"),
            *[({"reminder_time": value}, "valid reminder time")
              for value in ("9:00", "25:00", "09:60", "bad", None)],
            *[({"reminders_enabled": value}, "on or off")
              for value in ("true", 0, None)],
        ]
        for changes, expected_error in cases:
            with self.subTest(changes=changes):
                status, response = self.post_settings(self.body(**changes))
                self.assertEqual(status, 400)
                self.assertIn(expected_error, response["error"])
                self.assertEqual(self.settings_path.read_bytes(), self.original_bytes)
                self.settings_writer.assert_not_called()
                self.scheduler.assert_not_called()

    def test_save_without_scheduling_keeps_existing_response(self):
        for requested in (None, False):
            with self.subTest(update_reminder=requested):
                self.settings_writer.reset_mock()
                body = self.body(update_reminder=requested)
                if requested is None:
                    del body["update_reminder"]
                status, response = self.post_settings(body)
                self.assertEqual(status, 200)
                self.assertEqual(response["message"], "Settings saved.")
                self.assertNotIn("reminder_updated", response)
                self.assertNotIn("warning", response)
                self.assert_saved(response, body)
                self.settings_writer.assert_called_once()
                self.scheduler.assert_not_called()

    def test_check_rereads_committed_public_settings_after_save(self):
        state = {"rows": [], "scan_complete": True}
        with patch.object(web_app, "refresh", return_value=state) as refresh:
            status, before = self.post_request("/api/check", {})
            self.assertEqual(status, 200)
            public_keys = ("workbook_paths", "reminders_enabled", "reminder_time", "backup_destination")
            self.assertEqual(before, {"state": state, "settings": {
                key: self.original_settings[key] for key in public_keys}})

            body = self.body(workbook_paths=[str(path) for path in reversed(self.paths)],
                             reminders_enabled=False, reminder_time="18:45", update_reminder=False)
            status, saved = self.post_settings(body)
            self.assertEqual(status, 200)
            self.assert_saved(saved, body)
            self.settings_writer.reset_mock()

            status, after = self.post_request("/api/check", {})
            self.assertEqual(status, 200)
            self.assertEqual(after, {"state": state, "settings": saved["settings"]})
            self.assertNotEqual(after["settings"], before["settings"])
            self.assertEqual(refresh.call_count, 2)
            self.settings_writer.assert_not_called()
            self.scheduler.assert_not_called()

    def test_success_reports_updated_after_settings_are_saved(self):
        body = self.body()

        def schedule(*args, **kwargs):
            self.assertEqual(load_settings(self.root)["reminder_time"], body["reminder_time"])
            self.settings_writer.assert_called_once()
            return subprocess.CompletedProcess(args[0], 0, "Installed", "")

        self.scheduler.side_effect = schedule
        status, response = self.post_settings(body)
        self.assertEqual(status, 200)
        self.assertIs(response["reminder_updated"], True)
        self.assertEqual(response["message"], "Settings and daily reminder updated.")
        self.assertNotIn("warning", response)
        self.assert_saved(response, body)
        self.assert_scheduler_command()

    def test_successful_disable_uses_disable_mode_without_time(self):
        body = self.body(reminders_enabled=False)
        status, response = self.post_settings(body)
        self.assertEqual(status, 200)
        self.assertIs(response["reminder_updated"], True)
        self.assertNotIn("warning", response)
        self.assert_saved(response, body)
        self.assert_scheduler_command(enabled=False)

    def test_nonzero_scheduler_exit_returns_saved_settings_and_plain_warning(self):
        results = [
            subprocess.CompletedProcess([], 1, "Raw scheduler output", ""),
            subprocess.CompletedProcess([], 1, "", "Raw scheduler error"),
            subprocess.CompletedProcess([], 1, "", ""),
        ]
        for index, result in enumerate(results):
            with self.subTest(result=result):
                self.scheduler.reset_mock()
                self.settings_writer.reset_mock()
                self.scheduler.return_value = result
                body = self.body(reminder_time=f"15:0{index}")
                status, response = self.post_settings(body)
                self.assertEqual(status, 200)
                self.assertIs(response["reminder_updated"], False)
                self.assertEqual(response["message"], "Settings saved.")
                self.assertIn("Settings saved", response["warning"])
                self.assertIn("Save & repair reminder", response["warning"])
                self.assertNotIn("Raw scheduler", response["warning"])
                self.assert_saved(response, body)
                self.settings_writer.assert_called_once()
                self.assert_scheduler_command(reminder_time=body["reminder_time"])

    def test_timeout_or_oserror_returns_saved_settings_and_warning(self):
        errors = [
            subprocess.TimeoutExpired("mock scheduler", 60, output=b"Raw timeout output"),
            OSError("Raw scheduler launch error"),
            FileNotFoundError("Raw powershell.exe not found"),
        ]
        for index, error in enumerate(errors):
            with self.subTest(error=error):
                self.scheduler.reset_mock()
                self.settings_writer.reset_mock()
                self.scheduler.side_effect = error
                body = self.body(reminder_time=f"16:0{index}")
                status, response = self.post_settings(body)
                self.assertEqual(status, 200)
                self.assertIs(response["reminder_updated"], False)
                self.assertEqual(response["message"], "Settings saved.")
                self.assertIn("Settings saved", response["warning"])
                self.assertIn("Save & repair reminder", response["warning"])
                self.assertNotIn("Raw", response["warning"])
                self.assert_saved(response, body)
                self.settings_writer.assert_called_once()
                self.assert_scheduler_command(reminder_time=body["reminder_time"])

    def test_repair_after_failure_reports_success_without_stale_warning(self):
        self.scheduler.side_effect = [
            subprocess.CompletedProcess([], 1, "", ""),
            subprocess.CompletedProcess([], 0, "", ""),
        ]
        body = self.body()
        failed_status, failed_response = self.post_settings(body)
        self.assertEqual(failed_status, 200)
        self.assertIs(failed_response["reminder_updated"], False)
        self.assertIn("warning", failed_response)
        status, response = self.post_settings(body)
        self.assertEqual(status, 200)
        self.assertIs(response["reminder_updated"], True)
        self.assertNotIn("warning", response)
        self.assert_saved(response, body)
        self.assertEqual(self.scheduler.call_count, 2)

    def test_settings_write_failure_is_still_an_error_and_never_schedules(self):
        self.settings_writer.side_effect = OSError("Settings file is locked")
        status, response = self.post_settings(self.body())
        self.assertEqual(status, 400)
        self.assertEqual(response["error"], "Settings file is locked")
        self.assertNotIn("warning", response)
        self.assertNotIn("reminder_updated", response)
        self.assertEqual(self.settings_path.read_bytes(), self.original_bytes)
        self.scheduler.assert_not_called()


if __name__ == "__main__":
    unittest.main()
