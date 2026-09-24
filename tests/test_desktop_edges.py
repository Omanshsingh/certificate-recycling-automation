"""Desktop/scheduler regression tests. No Outlook, task registration or installs."""
from contextlib import ExitStack
import builtins
import json
import os
from pathlib import Path
import queue
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook

import desktop_app
import scheduled_run
import validate_daily_setup
from workflow_support import load_settings, save_json

ROOT = Path(__file__).resolve().parents[1]


class TemporarySetup(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.paths = [self.base / "Plant 1.xlsx", self.base / "Plant 2.xlsx"]
        for path in self.paths:
            book = Workbook()
            book.active.append(["Dummy fixture only"])
            book.save(path)
            book.close()
        self.settings = {"workbook_paths": list(map(str, self.paths)), "outlook_mailbox": "dummy@example.test",
                         "quiet_notifications": False, "initial_cod_confirmed": True, "certificate_number_start": 7}
        save_json(self.base / "settings.local.json", self.settings)
        # Fail closed if any test accidentally tries to use the real Outlook writer.
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        self.stack.enter_context(patch("draft_service.OutlookWriter", side_effect=AssertionError("Real Outlook is forbidden in these tests")))
        self.stack.enter_context(patch("run_certificates.OutlookWriter", side_effect=AssertionError("Real Outlook is forbidden in these tests")))
        self.notifier = self.stack.enter_context(patch("workflow_support.windows_notification", return_value=True))


class DesktopEdges(TemporarySetup):
    def test_daily_preflight_accepts_two_distinct_dummy_workbooks_and_mailbox(self):
        files = validate_daily_setup.validate_daily_setup(self.base)
        self.assertEqual(files, tuple(self.paths))

    def test_daily_preflight_rejects_missing_workbook_and_duplicate_plant_files(self):
        save_json(self.base / "settings.local.json", dict(self.settings, workbook_paths=[str(self.paths[0]), str(self.base / "missing.xlsx")]))
        with self.assertRaisesRegex(ValueError, "workbook is unavailable"):
            validate_daily_setup.validate_daily_setup(self.base)

        save_json(self.base / "settings.local.json", dict(self.settings, workbook_paths=[str(self.paths[0]), str(self.paths[0])]))
        with self.assertRaisesRegex(ValueError, "different Excel file"):
            validate_daily_setup.validate_daily_setup(self.base)

    def window(self):
        window = tk.Tk()
        window.withdraw()
        def dispose():
            try:
                for token in window.tk.splitlist(window.tk.call("after", "info")):
                    window.after_cancel(token)
                window.destroy()
            except tk.TclError:
                pass
        self.addCleanup(dispose)
        return window

    def app(self):
        return desktop_app.App(self.window(), self.base)

    def test_work_queue_filters_and_opens_selected_source_without_changing_it(self):
        latest = self.base / 'Output' / 'Latest Results.json'
        save_json(latest, {'rows': [
            dict(plant='plant1', sn='1', outcome='needs_review', detail='Choose Ready or N/A', file=str(self.paths[0]), sheet='Register', row=2),
            dict(plant='plant2', sn='2', outcome='ready', detail='Approved'),
            dict(plant='plant2', sn='3', outcome='not_applicable', detail='Skip'),
        ]})
        app = self.app()
        self.assertEqual(len(app.table.get_children()), 3)
        app.queue_filter.set('Needs attention')
        app.show_results()
        self.assertEqual(len(app.table.get_children()), 1)
        self.assertIn('Ready to draft: 1', app.queue_summary.get())
        app.table.selection_set(app.table.get_children()[0])
        with patch('desktop_app.os.startfile') as launch:
            app.open_selected_plant()
            launch.assert_called_once_with(str(self.paths[0]))
        self.assertIn('row 2', app.status.get())
        self.assertEqual(app.displayed_results_path, latest.with_suffix('.txt'))

    def test_unsaved_case_cannot_be_marked_sent(self):
        save_json(self.base / 'Output' / 'Latest Results.json', {'rows': [
            dict(plant='plant1', sn='1', outcome='ready', detail='Ready')
        ]})
        app = self.app()
        app.table.selection_set(app.table.get_children()[0])
        with patch('desktop_app.messagebox.askyesno') as confirm:
            app.mark_selected_sent()
            confirm.assert_not_called()
        self.assertIn('Only a prepared', app.status.get())

    def test_followup_history_read_failure_does_not_break_work_list(self):
        save_json(self.base / 'Output' / 'Latest Results.json', {'rows': [
            dict(plant='plant1', sn='1', outcome='already_prepared', detail='Prepared')
        ]})
        app = self.app()
        with patch('draft_service.History.records', side_effect=OSError('disk unavailable')):
            app.show_results()
        self.assertIn('history unavailable', app.status.get())
        self.assertEqual(len(app.table.get_children()), 1)

    def test_mark_sent_cancel_confirm_and_undo_keep_excel_unchanged(self):
        from draft_service import History
        history = History(self.base / 'Data' / 'certificate-history.sqlite')
        self.addCleanup(history.close)
        history.db.execute("INSERT INTO batches(case_id,status,certificate_no,entry_id) VALUES('1','saved','0042','test-only')")
        history.db.commit()
        save_json(self.base / 'Output' / 'Latest Results.json', {'rows': [
            dict(plant='plant1', sn='1', outcome='already_prepared', detail='COD 0042')
        ]})
        before = [p.read_bytes() for p in self.paths]
        app = self.app()
        for confirm, expected in ((False, False), (True, True), (True, False)):
            app.table.selection_set(app.table.get_children()[0])
            with patch('desktop_app.messagebox.askyesno', return_value=confirm):
                app.mark_selected_sent()
            self.assertEqual(bool(history.sent_acknowledgement('1')), expected)
            label = app.table.item(app.table.get_children()[0], 'values')[2]
            self.assertEqual(label, 'Marked sent' if expected else 'Drafts to review')
        self.assertEqual(before, [p.read_bytes() for p in self.paths])

    def test_startup_marks_results_stale_after_a_workbook_changes(self):
        latest = self.base / "Output" / "Latest Results.json"
        save_json(latest, {"completed_at": "2000-01-01T00:00:00", "rows": [
            {"plant": "plant1", "sn": "1001", "outcome": "needs_review", "detail": "old stale error"},
        ]})
        app = self.app()
        self.assertIn("Excel files changed after the last check", app.status.get())
        self.assertIn("Check both files", app.status.get())
        self.assertEqual(app.table.get_children(), ())

    def test_startup_explains_that_pdf_preview_did_not_create_a_draft(self):
        latest = self.base / "Output" / "Latest Results.json"
        save_json(latest, {"completed_at": "2026-09-23T13:48:06", "mode": "PDF preview", "rows": []})
        app = self.app()
        self.assertIn("PDF-only preview", app.status.get())
        self.assertIn("did not create a draft", app.status.get())
        self.assertIn("Create one draft", app.status.get())

    def test_startup_summarizes_saved_draft_and_remaining_attention(self):
        latest = self.base / "Output" / "Latest Results.json"
        save_json(latest, {"completed_at": "2026-09-23T14:00:00", "mode": "Outlook drafts", "rows": [
            {"plant": "plant1", "sn": "1001", "outcome": "draft_created", "detail": "saved"},
            {"plant": "plant2", "sn": "2005", "outcome": "needs_review", "detail": "blank status"},
        ]})
        app = self.app()
        self.assertIn("1 Outlook draft saved", app.status.get())
        self.assertIn("1 row need attention", app.status.get())
        self.assertIn("Review Outlook Drafts", app.status.get())

    def test_ui_has_one_draft_and_batch_actions_without_test_only_buttons(self):
        app = self.app()
        labels = [str(control.cget("text")) for control in app.controls if control.winfo_class() == "TButton"]
        self.assertIn("Create one draft", labels)
        self.assertIn("Create all Ready drafts", labels)
        self.assertFalse(any("test draft" in label.casefold() or "sample" in label.casefold() for label in labels))

    def test_one_draft_action_is_confirmed_without_test_only_copy(self):
        app = self.app()
        dummy_paths = [self.base / "Dummy_Inward_Plant_1.xlsx", self.base / "Dummy_Inward_Plant_2.xlsx"]
        for path in dummy_paths:
            path.touch()
        for variable, path in zip(app.workbooks, dummy_paths):
            variable.set(str(path))

        with patch("desktop_app.messagebox.askyesno", return_value=True) as confirm, patch("desktop_app.threading.Thread") as thread:
            app.start(limit=1)

        confirm.assert_called_once()
        prompt = confirm.call_args.args[1]
        self.assertIn("normal template", prompt)
        self.assertIn("one certificate", prompt)
        self.assertNotIn("dummy", prompt.casefold())
        self.assertNotIn("example.test", prompt.casefold())
        thread.assert_called_once()
        self.assertEqual(thread.call_args.kwargs["args"], (list(map(str, dummy_paths)), "dummy@example.test", True, 1, False))
        self.assertTrue(app.busy)

    def test_batch_action_is_available_for_selected_workbook_names(self):
        app = self.app()
        dummy_paths = [self.base / "Dummy_Inward_Plant_1.xlsx", self.base / "Dummy_Inward_Plant_2.xlsx"]
        for path in dummy_paths:
            path.touch()
        for variable, path in zip(app.workbooks, dummy_paths):
            variable.set(str(path))
        with patch("desktop_app.messagebox.askyesno", return_value=True), patch("desktop_app.threading.Thread") as thread:
            app.start()
        thread.assert_called_once()
        self.assertTrue(app.busy)

    def test_worker_uses_the_shared_cod_history(self):
        app = self.app()
        dummy_paths = [self.base / "Dummy_Inward_Plant_1.xlsx", self.base / "Dummy_Inward_Plant_2.xlsx"]
        output = self.base / "Output" / "dummy-run"
        with patch.dict(sys.modules, {"pythoncom": Mock()}), patch("desktop_app.run", return_value=(1, 0, output)) as process:
            app.worker(list(map(str, dummy_paths)), "dummy@example.test", True, 1, False)
        self.assertNotIn("history_path", process.call_args.kwargs)
        self.assertEqual(app.events.get_nowait()[0], "done")

    def test_corrupt_settings_survive_startup_save_and_choose_then_repair(self):
        path = self.base / "settings.local.json"
        path.write_text('{"workbook_paths": broken', encoding="utf-8")
        original = path.read_bytes()
        app = self.app()
        self.assertIn("Settings unavailable", app.status.get())
        self.assertIn("Reload settings", app.status.get())
        with patch("desktop_app.filedialog.askopenfilename") as dialog, patch("desktop_app.threading.Thread") as thread:
            app.choose_workbook(0)
            app.start(check_only=True)
            app.load_accounts()
            self.assertFalse(app.save_settings())
            dialog.assert_not_called()
            thread.assert_not_called()
        self.assertEqual(path.read_bytes(), original)
        save_json(path, self.settings)
        app.reload_settings()
        self.assertEqual(app.settings_error, "")
        self.assertEqual([v.get() for v in app.workbooks], self.settings["workbook_paths"])
        self.assertTrue(app.save_settings())

    def test_missing_processing_dependency_leaves_repair_controls_available(self):
        with patch("desktop_app.load_settings", side_effect=ModuleNotFoundError("openpyxl")):
            app = self.app()
        self.assertIn("Setup Once.bat", app.status.get())
        app.reload_settings()
        self.assertFalse(app.settings_error)
        self.assertTrue(app.save_settings())

    def test_corruption_after_startup_is_not_overwritten(self):
        app = self.app()
        path = self.base / "settings.local.json"
        path.write_text("[]", encoding="utf-8")
        self.assertFalse(app.save_settings())
        self.assertEqual(path.read_text(), "[]")
        self.assertIn("Settings unavailable", app.status.get())

    def test_failed_save_prevents_run_and_preserves_both_saved_paths(self):
        app = self.app()
        before = (self.base / "settings.local.json").read_bytes()
        previous = dict(app.settings)
        with patch("desktop_app.save_json", side_effect=PermissionError("read-only folder")), patch("desktop_app.threading.Thread") as thread:
            app.start(check_only=True)
            thread.assert_not_called()
        self.assertFalse(app.busy)
        self.assertIn("not saved", app.status.get())
        self.assertEqual(app.settings, previous)
        self.assertEqual((self.base / "settings.local.json").read_bytes(), before)

    def test_choose_failed_save_rolls_back_visible_path(self):
        app = self.app()
        previous = app.workbooks[0].get()
        with patch("desktop_app.filedialog.askopenfilename", return_value=str(self.base / "Replacement.xlsx")), patch("desktop_app.save_json", side_effect=PermissionError("denied")):
            app.choose_workbook(0)
        self.assertEqual(app.workbooks[0].get(), previous)
        self.assertIn("not saved", app.status.get())

    def test_invalid_first_cod_values_do_not_poison_settings(self):
        app = self.app()
        before = (self.base / "settings.local.json").read_bytes()
        for value in ("0", "-1", "1.5", "²", "9999999999999", "9" * 5000):
            with self.subTest(value=value[:20]):
                app.first_cod.set(value)
                self.assertFalse(app.save_settings())
                self.assertEqual((self.base / "settings.local.json").read_bytes(), before)
        app.first_cod.set("999999999999")
        self.assertTrue(app.save_settings())
        self.assertEqual(load_settings(self.base)["certificate_number_start"], 999999999999)

    def test_external_unrelated_setting_change_is_preserved(self):
        app = self.app()
        save_json(self.base / "settings.local.json", dict(self.settings, certificate_number_prefix="APPROVED-"))
        self.assertTrue(app.save_settings())
        self.assertEqual(load_settings(self.base)["certificate_number_prefix"], "APPROVED-")

    def test_duplicate_missing_and_samefile_alias_paths_stop_before_worker(self):
        app = self.app()
        alias = self.base / "Alias.xlsx"
        os.link(self.paths[0], alias)
        for second, message in ((self.paths[0], "different file"), (alias, "different file"), (self.base / "missing.xlsx", "unavailable")):
            with self.subTest(path=second), patch("desktop_app.threading.Thread") as thread:
                app.workbooks[1].set(str(second))
                app.start(check_only=True)
                self.assertIn(message, app.status.get())
                thread.assert_not_called()

    def test_blank_mailbox_blocks_drafts_but_not_check_or_preview(self):
        app = self.app()
        app.mailbox.set("   ")
        with patch("desktop_app.threading.Thread") as thread:
            app.start()
            thread.assert_not_called()
            self.assertIn("Find Outlook mailboxes", app.status.get())
            app.start(check_only=True)
            self.assertTrue(app.busy)
            self.assertTrue(thread.called)
            app.set_busy(False)
            thread.reset_mock()
            app.start(drafts=False, limit=1)
            self.assertTrue(thread.called)

    def test_thread_start_failure_restores_controls(self):
        app = self.app()
        with patch("desktop_app.threading.Thread") as thread:
            thread.return_value.start.side_effect = RuntimeError("thread unavailable")
            app.start(check_only=True)
        self.assertFalse(app.busy)
        self.assertIn("Could not start", app.status.get())
        self.assertEqual(str(app.accounts.cget("state")), "readonly")

    def test_malformed_results_do_not_crash_startup_or_stop_polling(self):
        output = self.base / "Output"
        output.mkdir()
        for content in ("broken", "[]", '{"rows": null}', '{"rows": [{}]}'):
            with self.subTest(content=content):
                (output / "Latest Results.json").write_text(content, encoding="utf-8")
                app = self.app()
                self.assertIn("Could not read", app.status.get())
                app.set_busy(True)
                app.events.put(("done", "Finished"))
                with patch.object(app.root, "after") as after:
                    app.poll()
                self.assertFalse(app.busy)
                after.assert_called_once()

    def test_account_save_failure_is_not_hidden_and_polling_continues(self):
        app = self.app()
        app.set_busy(True)
        app.events.put(("accounts", ["dummy2@example.test"]))
        with patch("desktop_app.save_json", side_effect=PermissionError("denied")), patch.object(app.root, "after") as after:
            app.poll()
        self.assertFalse(app.busy)
        self.assertIn("not saved", app.status.get())
        after.assert_called_once()

    def test_close_during_work_waits_for_completion(self):
        app = self.app()
        app.set_busy(True)
        with patch.object(app.root, "destroy") as destroy:
            app.close()
            destroy.assert_not_called()
            app.events.put(("done", "Finished"))
            app.poll()
            destroy.assert_called_once()

    def test_opening_failed_output_location_shows_error(self):
        app = self.app()
        with patch("desktop_app.os.startfile", side_effect=OSError("cannot open")):
            app.open_output()
            self.assertIn("Could not open", app.status.get())
            app.open_settings_folder()
            self.assertIn("Could not open", app.status.get())

    def test_check_and_preview_worker_do_not_require_com(self):
        app = self.app()
        for drafts, check in ((True, True), (False, False)):
            with self.subTest(drafts=drafts, check=check), patch.dict(sys.modules, {"pythoncom": None}), patch("desktop_app.run", return_value=(0, 0, self.base)) as run:
                app.worker(self.paths, "", drafts, 1, check)
                run.assert_called_once()
                self.assertEqual(app.events.get_nowait()[0], "done")
        self.notifier.assert_not_called()

    def test_draft_worker_missing_com_reports_setup_without_running(self):
        app = self.app()
        with patch.dict(sys.modules, {"pythoncom": None}), patch("desktop_app.run") as run:
            app.worker(self.paths, "dummy@example.test", True, 1, False)
        run.assert_not_called()
        kind, message = app.events.get_nowait()
        self.assertEqual(kind, "error")
        self.assertIn("Setup Once.bat", message)


class ScheduledEdges(TemporarySetup):
    def setUp(self):
        super().setUp()
        self.com = Mock()
        self.stack.enter_context(patch.dict(sys.modules, {"pythoncom": self.com}))
        self.output = self.base / "Mock output"
        self.output.mkdir()
        self.result = {"run_id": "dummy-run", "created": 1, "updated": 0, "review": 0, "rows": [], "scan_complete": True}
        save_json(self.output / "Results.json", self.result)
        self.run = self.stack.enter_context(patch("run_certificates.run", return_value=(1, 0, self.output)))

    def test_muted_run_calls_notify_to_acknowledge_results(self):
        with patch("workflow_support.notify_once") as notify:
            self.assertEqual(scheduled_run.main(self.base), 0)
        notify.assert_called_once_with(self.base, self.result, self.notifier, enabled=False, repeat_pending=True)
        self.assertTrue(self.run.call_args.kwargs["create_drafts"])
        self.com.CoInitialize.assert_called_once()
        self.com.CoUninitialize.assert_called_once()

    def test_actual_muted_state_has_no_backlog_on_unmute(self):
        self.assertEqual(scheduled_run.main(self.base), 0)
        self.notifier.assert_not_called()
        save_json(self.base / "settings.local.json", dict(self.settings, quiet_notifications=True))
        self.assertEqual(scheduled_run.main(self.base), 0)
        self.notifier.assert_not_called()

    def test_corrupt_settings_become_actionable_report_without_outlook(self):
        path = self.base / "settings.local.json"
        path.write_text("corrupt", encoding="utf-8")
        self.assertEqual(scheduled_run.main(self.base), 1)
        self.assertEqual(path.read_text(), "corrupt")
        result = json.loads((self.base / "Output" / "Latest Results.json").read_text())
        self.assertIn("settings.local.json", result["rows"][0]["detail"])
        self.run.assert_not_called()
        self.com.CoInitialize.assert_not_called()

    def test_incomplete_or_duplicate_paths_stop_before_outlook(self):
        for values in ({"workbook_paths": [str(self.paths[0]), ""]}, {"workbook_paths": [str(self.paths[0])] * 2}):
            with self.subTest(values=values):
                save_json(self.base / "settings.local.json", dict(self.settings, **values))
                self.assertEqual(scheduled_run.main(self.base), 1)
        self.run.assert_not_called()
        self.com.CoInitialize.assert_not_called()

    def test_empty_mailbox_still_checks_sheets_without_outlook(self):
        save_json(self.base / "settings.local.json", dict(self.settings, outlook_mailbox=""))
        self.assertEqual(scheduled_run.main(self.base), 1)
        self.run.assert_called_once()
        self.assertFalse(self.run.call_args.kwargs["create_drafts"])
        self.assertTrue(self.run.call_args.kwargs["check_only"])
        self.com.CoInitialize.assert_not_called()
        result = json.loads((self.base / "Output" / "Latest Results.json").read_text())
        self.assertIn("save the mailbox", result["rows"][-1]["detail"])

    def test_missing_file_delegates_partial_scan_instead_of_hiding_other_plant(self):
        paths = [str(self.paths[0]), str(self.base / "absent.xlsx")]
        save_json(self.base / "settings.local.json", dict(self.settings, workbook_paths=paths))
        partial = dict(self.result, scan_complete=False, observed_sources=["plant1"], created=0, review=2,
                       rows=[{"plant": "plant1", "sn": "1", "outcome": "needs_review", "detail": "Certificate Issued is blank"},
                             {"plant": "plant2", "sn": "", "outcome": "needs_review", "detail": "Source unavailable"}])
        save_json(self.output / "Results.json", partial)
        self.run.return_value = (0, 2, self.output)
        with patch("workflow_support.notify_once") as notify:
            self.assertEqual(scheduled_run.main(self.base), 1)
        self.assertEqual(self.run.call_args.args[0], paths)
        self.assertEqual(notify.call_args.args[1]["observed_sources"], ["plant1"])
        self.assertEqual(len(notify.call_args.args[1]["rows"]), 2)

    def test_missing_com_is_reported_but_sheets_are_still_checked(self):
        with patch.dict(sys.modules, {"pythoncom": None}):
            self.assertEqual(scheduled_run.main(self.base), 1)
        self.run.assert_called_once()
        self.assertFalse(self.run.call_args.kwargs["create_drafts"])
        self.assertTrue(self.run.call_args.kwargs["check_only"])
        result = json.loads((self.base / "Output" / "Latest Results.json").read_text())
        self.assertIn("Setup Once.bat", result["rows"][0]["detail"])

    def test_missing_import_at_start_is_caught_and_logged_without_dependencies(self):
        real_import = builtins.__import__
        def missing(name, *args, **kwargs):
            if name == "workflow_support":
                raise ModuleNotFoundError("No module named openpyxl")
            return real_import(name, *args, **kwargs)
        with patch("builtins.__import__", side_effect=missing):
            self.assertEqual(scheduled_run.main(self.base), 1)
        self.assertIn("Setup Once.bat", (self.base / "Output" / "Daily check error.txt").read_text())
        self.run.assert_not_called()

    def test_read_only_report_location_uses_per_user_diagnostic(self):
        self.run.side_effect = PermissionError("read-only project")
        # A file standing in for Output simulates an unwritable report directory.
        (self.base / "Output").write_text("blocked", encoding="utf-8")
        local = self.base / "Local app data"
        with patch.dict(os.environ, {"LOCALAPPDATA": str(local)}):
            self.assertEqual(scheduled_run.main(self.base), 1)
        logs = list(local.rglob("Daily check error.txt"))
        self.assertEqual(len(logs), 1)
        self.assertIn("read-only project", logs[0].read_text())
        self.assertIn(str(self.base), logs[0].read_text())
        self.com.CoUninitialize.assert_called_once()

    def test_notification_state_write_failure_does_not_crash_runner(self):
        with patch("workflow_support.notify_once", side_effect=PermissionError("state is read-only")):
            self.assertEqual(scheduled_run.main(self.base), 1)
        self.assertIn("notification state", (self.base / "Output" / "Daily check error.txt").read_text())

    def test_overlap_skips_without_notification_and_releases_com(self):
        self.run.side_effect = RuntimeError("Certificate processing is already running. Wait for it to finish.")
        with patch("workflow_support.notify_once") as notify:
            self.assertEqual(scheduled_run.main(self.base), 0)
        notify.assert_not_called()
        self.com.CoUninitialize.assert_called_once()
        self.assertFalse((self.base / "Output").exists())

    def test_broken_result_json_becomes_report_and_releases_com(self):
        (self.output / "Results.json").write_text("broken", encoding="utf-8")
        self.assertEqual(scheduled_run.main(self.base), 1)
        self.com.CoUninitialize.assert_called_once()
        self.assertTrue((self.base / "Output" / "Latest Results.json").is_file())


@unittest.skipUnless(os.name == "nt" and shutil.which("powershell.exe"), "Windows PowerShell contract tests")
class LauncherEdges(unittest.TestCase):
    def powershell(self, source):
        result = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", source], text=True, capture_output=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        return result.stdout.strip()

    def parsed(self):
        script = str(ROOT / "Install Daily Task.ps1").replace("'", "''")
        return f"$tokens=$null; $errors=$null; $ast=[System.Management.Automation.Language.Parser]::ParseFile('{script}',[ref]$tokens,[ref]$errors); if ($errors.Count) {{ throw ($errors | Out-String) }}; "

    def test_task_configuration_uses_hidden_interactive_catchup_no_overlap(self):
        # Extract only task construction, and replace ALL task commands before
        # evaluation. No real scheduler cmdlet is ever called by this test.
        output = self.powershell(self.parsed() + r'''
$body = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] }
$statements = $body.Body.Statements
$capture = $false
$selected = foreach ($statement in $statements) {
    $text = $statement.Extent.Text
    if ($text.StartsWith('$certificateUser =')) { $capture = $true }
    if ($capture) { $text }
    if ($text.StartsWith('Register-ScheduledTask')) { break }
}
$code = ($selected -join "`n").Replace('New-ScheduledTask', 'Test-New-ScheduledTask').Replace('Register-ScheduledTask', 'Test-Register-ScheduledTask')
function Test-New-ScheduledTaskAction { param($Execute,$Argument,$WorkingDirectory) @{Execute=$Execute;Argument=$Argument;WorkingDirectory=$WorkingDirectory} }
function Test-New-ScheduledTaskTrigger { param([switch]$Daily,$At,[switch]$AtLogOn,$User) @{Daily=[bool]$Daily;At=$At;AtLogOn=[bool]$AtLogOn;User=$User;Delay=''} }
function Test-New-ScheduledTaskSettingsSet { param([switch]$StartWhenAvailable,$MultipleInstances,$ExecutionTimeLimit,[switch]$AllowStartIfOnBatteries,[switch]$DontStopIfGoingOnBatteries) @{Catchup=[bool]$StartWhenAvailable;MultipleInstances=$MultipleInstances;Limit=$ExecutionTimeLimit.TotalSeconds;Battery=[bool]$AllowStartIfOnBatteries;KeepBattery=[bool]$DontStopIfGoingOnBatteries} }
function Test-New-ScheduledTaskPrincipal { param($UserId,$LogonType,$RunLevel) @{User=$UserId;LogonType=$LogonType;RunLevel=$RunLevel} }
function Test-Register-ScheduledTask { param($TaskName,$Action,$Trigger,$Settings,$Principal,$Description,[switch]$Force) $global:captured=@{Action=$Action;Triggers=$Trigger;Settings=$Settings;Principal=$Principal} }
$PSScriptRoot='C:\Dummy installation with spaces'
$certificatePythonw='C:\Dummy Python\pythonw.exe'
$scriptArgument='"C:\Dummy installation with spaces\scheduled_run.py"'
& ([scriptblock]::Create($code))
$global:captured | ConvertTo-Json -Depth 8 -Compress
''')
        task = json.loads(output)
        self.assertTrue(task["Action"]["Execute"].endswith("pythonw.exe"))
        self.assertEqual(task["Action"]["Argument"], '"C:\\Dummy installation with spaces\\scheduled_run.py"')
        self.assertEqual(task["Triggers"][0]["At"], "09:00")
        self.assertTrue(task["Triggers"][1]["AtLogOn"])
        self.assertEqual(task["Triggers"][1]["Delay"], "PT2M")
        self.assertEqual(task["Settings"], {"Catchup": True, "MultipleInstances": "IgnoreNew", "Limit": 0, "Battery": True, "KeepBattery": True})
        self.assertEqual(task["Principal"]["LogonType"], "Interactive")
        self.assertEqual(task["Principal"]["RunLevel"], "Limited")

    def test_runtime_probe_falls_back_from_unusable_python_and_checks_version(self):
        output = self.powershell(self.parsed() + r'''
$function = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Resolve-CertificatePython' }
. ([scriptblock]::Create($function.Extent.Text))
$PSScriptRoot='C:\Dummy installation'
$global:probes=@()
function Get-Command { param($Name,$CommandType,[switch]$All,$ErrorAction) if ($Name -eq 'python.exe') { [pscustomobject]@{Source='Probe-GoodPython'} } else { [pscustomobject]@{Source='Probe-BadPython'} } }
function Test-Path { param($LiteralPath,$PathType) return $LiteralPath -eq 'Probe-GoodPython' }
function Probe-BadPython { $global:probes += ,@($args); 'Python launcher check failed'; $global:LASTEXITCODE=1 }
function Probe-GoodPython { $global:probes += ,@($args); $global:LASTEXITCODE=0; 'Probe-GoodPython' }
$resolved = Resolve-CertificatePython
@{Path=$resolved;Probes=$global:probes} | ConvertTo-Json -Depth 8 -Compress
''')
        result = json.loads(output)
        self.assertEqual(result["Path"], "Probe-GoodPython")
        self.assertEqual(len(result["Probes"]), 4)
        self.assertEqual(result["Probes"][2][:1], ["-c"])
        self.assertIn("sys.version_info >= (3, 11)", result["Probes"][2][1])
        self.assertIn("print(sys.executable)", result["Probes"][2][1])
        self.assertNotIn('"Python 3.11 or newer is required"', result["Probes"][2][1])
        self.assertEqual(result["Probes"][3][:1], ["-c"])
        self.assertIn("pythoncom", result["Probes"][3][1])

    def test_python_install_manager_windowsapps_runtime_is_enumerated(self):
        script = (ROOT / "Install Daily Task.ps1").read_text(encoding="utf-8")
        self.assertIn("@('py.exe', 'pymanager.exe')", script)
        self.assertIn("list --format=exe", script)
        self.assertIn("$isWindowsApps -and -not $listedRuntimes", script)
        # Store aliases for python.exe are excluded, but py.exe is separately
        # queried for real installed runtime paths rather than discarded.
        self.assertIn("The WindowsApps python.exe alias can open the Store", script)

    def test_daily_preflight_script_launches_under_windows_powershell_with_spaced_paths(self):
        with tempfile.TemporaryDirectory(prefix="certificate preflight ") as directory:
            root = Path(directory)
            workbooks = [root / "Plant 1 dummy.xlsx", root / "Plant 2 dummy.xlsx"]
            for workbook in workbooks:
                workbook.write_bytes(b"dummy workbook placeholder")
            save_json(root / "settings.local.json", {
                "workbook_paths": [str(path) for path in workbooks],
                "outlook_mailbox": "dummy@example.test",
            })
            before = (root / "settings.local.json").read_bytes()
            command = f'& "{sys.executable}" "{ROOT / "validate_daily_setup.py"}"; exit $LASTEXITCODE'
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
                cwd=root,
                text=True,
                capture_output=True,
                timeout=30,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("Preflight passed", result.stdout)
            self.assertEqual((root / "settings.local.json").read_bytes(), before)

            save_json(root / "settings.local.json", {
                "workbook_paths": [str(path) for path in workbooks],
                "outlook_mailbox": "",
            })
            command = f'& "{sys.executable}" "{ROOT / "validate_daily_setup.py"}"; exit $LASTEXITCODE'
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
                cwd=root,
                text=True,
                capture_output=True,
                timeout=30,
            )
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertIn("Load Outlook accounts and select the mailbox", result.stdout)
            self.assertNotIn("Traceback", result.stdout + result.stderr)

    def test_runtime_failure_names_missing_import(self):
        output = self.powershell(self.parsed() + r'''
$function = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -eq 'Resolve-CertificatePython' }
. ([scriptblock]::Create($function.Extent.Text))
$PSScriptRoot='C:\Dummy installation'
function Get-Command { param($Name,$CommandType,[switch]$All,$ErrorAction) if ($Name -eq 'python.exe') { [pscustomobject]@{Source='Probe-MissingModule'} } }
function Test-Path { param($LiteralPath,$PathType) return $LiteralPath -eq 'Probe-MissingModule' }
function Probe-MissingModule { if ($args[1] -match 'print\(sys\.executable\)') { $global:LASTEXITCODE=0; 'Probe-MissingModule' } else { $global:LASTEXITCODE=1; "ModuleNotFoundError: No module named 'openpyxl'" } }
$resolved = Resolve-CertificatePython
@{Path=$resolved;Failures=$script:CertificatePythonProbeFailures} | ConvertTo-Json -Depth 5 -Compress
''')
        result = json.loads(output)
        self.assertIsNone(result["Path"])
        self.assertTrue(any("openpyxl" in failure for failure in result["Failures"]))

    def test_only_explicit_setup_mode_can_install_packages(self):
        script = (ROOT / "Install Daily Task.ps1").read_text()
        self.assertEqual(script.count("'pip', 'install'"), 1)
        for name, mode in (("Setup Once.bat", "Setup"), ("Prepare Certificates.bat", "Desktop"), ("Run Daily Drafts.bat", "Daily"), ("Create Daily Draft Task.bat", "Install"), ("Check Python Setup.bat", "Diagnose")):
            content = (ROOT / name).read_text()
            self.assertIn(f"-Mode {mode}", content)
            self.assertNotIn("pip install", content)
            self.assertNotIn("ExecutionPolicy Bypass", content)
        launcher = (ROOT / "Install Daily Task.ps1").read_text(encoding="utf-8")
        self.assertIn("validate_daily_setup.py", launcher)
        self.assertLess(launcher.index("if ($Mode -eq 'Diagnose')"), launcher.index("# Validate the saved setup without creating drafts or contacting Outlook."))


if __name__ == "__main__":
    unittest.main()
