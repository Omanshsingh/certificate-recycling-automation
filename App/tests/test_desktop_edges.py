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
import tkinter.ttk as ttk
import unittest
from datetime import date
from unittest.mock import Mock, patch

from openpyxl import Workbook

import desktop_app
import scheduled_run
import validate_daily_setup
import workflow_support
from workflow_support import load_settings, save_json

ROOT = Path(__file__).resolve().parents[1]


class CenteredReminderLifecycle(unittest.TestCase):
    def _run_standalone(self, trigger):
        host, window = Mock(), Mock()
        window.winfo_width.return_value = 420
        window.winfo_height.return_value = 240
        window.winfo_screenwidth.return_value = 1600
        window.winfo_screenheight.return_value = 900
        protocols, bindings, buttons = {}, {}, []
        window.protocol.side_effect = lambda key, callback: protocols.__setitem__(key, callback)
        window.bind.side_effect = lambda key, callback: bindings.__setitem__(key, callback)

        def button_factory(*_args, **kwargs):
            button = Mock()
            button.command = kwargs["command"]
            buttons.append(button)
            return button

        host.mainloop.side_effect = lambda: trigger(host, window, buttons, protocols, bindings)
        on_snooze, on_dismiss, on_review = Mock(), Mock(), Mock()
        with ExitStack() as stack:
            stack.enter_context(patch.object(workflow_support.tk, "Tk", return_value=host))
            stack.enter_context(patch.object(workflow_support.tk, "Toplevel", return_value=window))
            stack.enter_context(patch.object(workflow_support.tk, "StringVar", return_value=Mock()))
            stack.enter_context(patch.object(ttk, "Label", side_effect=lambda *_a, **_k: Mock()))
            stack.enter_context(patch.object(ttk, "Frame", side_effect=lambda *_a, **_k: Mock()))
            stack.enter_context(patch.object(ttk, "Button", side_effect=button_factory))
            action = workflow_support.show_centered_reminder(
                "Reminder", "Private-free test summary", on_snooze=on_snooze,
                on_dismiss=on_dismiss, on_review=on_review)
        return action, host, window, on_snooze, on_dismiss, on_review

    def test_each_standalone_button_quits_the_hidden_host_loop(self):
        for index, expected in enumerate(("review", "dismiss", "snooze")):
            with self.subTest(button=expected):
                action, host, window, snooze, dismiss, review = self._run_standalone(
                    lambda _host, _window, buttons, _protocols, _bindings:
                    buttons[index].command())
                self.assertEqual(action, expected)
                host.quit.assert_called_once_with()
                host.destroy.assert_called_once_with()
                window.destroy.assert_called_once_with()
                self.assertEqual(dismiss.call_count, 0 if expected == "snooze" else 1)
                self.assertEqual(review.call_count, 1 if expected == "review" else 0)
                self.assertEqual(snooze.call_count, 1 if expected == "snooze" else 0)

    def test_standalone_window_close_and_escape_quit_the_host_loop(self):
        for event in ("WM_DELETE_WINDOW", "<Escape>"):
            with self.subTest(event=event):
                def trigger(_host, _window, _buttons, protocols, bindings):
                    if event == "WM_DELETE_WINDOW":
                        protocols[event]()
                    else:
                        bindings[event](None)

                action, host, window, _snooze, dismiss, _review = self._run_standalone(trigger)
                self.assertEqual(action, "dismiss")
                host.quit.assert_called_once_with()
                host.destroy.assert_called_once_with()
                window.destroy.assert_called_once_with()
                dismiss.assert_called_once_with()

    def test_embedded_parent_path_does_not_quit_or_destroy_the_app_root(self):
        parent, window = Mock(), Mock()
        window.winfo_width.return_value = 420
        window.winfo_height.return_value = 240
        window.winfo_screenwidth.return_value = 1600
        window.winfo_screenheight.return_value = 900
        buttons, bindings = [], {}
        window.bind.side_effect = lambda key, callback: bindings.__setitem__(key, callback)

        def button_factory(*_args, **kwargs):
            button = Mock()
            button.command = kwargs["command"]
            buttons.append(button)
            return button

        parent.wait_window.side_effect = lambda _window: buttons[1].command()
        with ExitStack() as stack:
            stack.enter_context(patch.object(workflow_support.tk, "Tk",
                                             side_effect=AssertionError("embedded path made a new root")))
            stack.enter_context(patch.object(workflow_support.tk, "Toplevel", return_value=window))
            stack.enter_context(patch.object(workflow_support.tk, "StringVar", return_value=Mock()))
            stack.enter_context(patch.object(ttk, "Label", side_effect=lambda *_a, **_k: Mock()))
            stack.enter_context(patch.object(ttk, "Frame", side_effect=lambda *_a, **_k: Mock()))
            stack.enter_context(patch.object(ttk, "Button", side_effect=button_factory))
            action = workflow_support.show_centered_reminder("Reminder", "Test", parent=parent)
        self.assertEqual(action, "dismiss")
        parent.quit.assert_not_called()
        parent.destroy.assert_not_called()
        parent.wait_window.assert_called_once_with(window)
        window.destroy.assert_called_once_with()


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
        # Daily tests must never open a real reminder window.
        self.notifier = self.stack.enter_context(patch("workflow_support.show_centered_reminder", return_value="dismiss"))


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

    def test_saved_setup_is_collapsed_and_new_setup_opens_automatically(self):
        app = self.app()
        self.assertFalse(app.settings_visible)
        self.assertEqual(app.plants_frame.master, app.settings_frame)
        save_json(self.base / 'settings.local.json', {'workbook_paths': ['', '']})
        fresh = self.app()
        self.assertTrue(fresh.settings_visible)

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

    def test_overdue_filter_shows_only_30_day_items(self):
        save_json(self.base / 'Output' / 'Latest Results.json', {'rows': [
            dict(plant='plant1', sn='1', outcome='needs_review', detail='Old blank', reminder_due=True),
            dict(plant='plant2', sn='2', outcome='needs_review', detail='Recent blank', reminder_due=False),
        ]})
        app = self.app()
        app.queue_filter.set('Overdue 30+ days')
        app.show_results()
        self.assertEqual(len(app.table.get_children()), 1)
        self.assertEqual(app.table.item(app.table.get_children()[0], 'values')[1], '1')
        self.assertIn('Overdue 30+ days: 1', app.queue_summary.get())

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
        self.assertNotIn('history unavailable', app.status.get())
        self.assertEqual(len(app.table.get_children()), 1)

    def test_mark_sent_updates_excel_status_and_undo_restores_completed(self):
        from draft_service import History
        book = Workbook()
        sheet = book.active
        sheet.title = 'Register'
        headers = ["SN", "Inward Date", "Supplier Name", "Location", "Item", "Quantity",
                   "Per Piece Weight", "Weight", "Form 6 No", "Client Email", "Reference No",
                   "COD No", "COD Date", "Certificate Issued"]
        sheet.append(headers)
        sheet.append(["1", date(2026, 8, 1), "Dummy Client", "Mumbai", "Laptop", 2, 3.5,
                      7, "FORM-1", "dummy@example.test", "REF-1", "0042", date(2026, 9, 1), "Completed"])
        book.save(self.paths[0])
        book.close()
        history = History(self.base / 'Data' / 'certificate-history.sqlite')
        self.addCleanup(history.close)
        history.db.execute("INSERT INTO batches(case_id,status,certificate_no,entry_id) VALUES('1','saved','0042','test-only')")
        history.db.commit()
        save_json(self.base / 'Output' / 'Latest Results.json', {'rows': [
            dict(plant='plant1', sn='1', outcome='already_prepared', detail='COD 0042',
                 certificate_no='0042', certificate_status='completed', file=str(self.paths[0]),
                 sheet='Register', row=2, reminder_due=True, age_days=54)
        ]})
        app = self.app()
        for confirm, expected in ((False, False), (True, True), (True, False)):
            app.table.selection_set(app.table.get_children()[0])
            with patch('desktop_app.messagebox.askyesno', return_value=confirm):
                app.mark_selected_sent()
            self.assertEqual(bool(history.sent_acknowledgement('1')), expected)
            label = app.table.item(app.table.get_children()[0], 'values')[2]
            self.assertEqual(label, 'Sent' if expected else 'Drafts to review')
            if expected:
                latest = json.loads((self.base / 'Output' / 'Latest Results.json').read_text())
                self.assertEqual(latest['rows'][0]['certificate_status'], 'sent')
                self.assertIn('marked Sent', (self.base / 'Output' / 'Latest Results.txt').read_text())
        final = __import__('openpyxl').load_workbook(self.paths[0], data_only=False)
        try:
            self.assertEqual(final['Register']['N2'].value, 'Completed')
            self.assertEqual(final['Register']['L2'].value, '0042')
        finally:
            final.close()
        self.assertGreaterEqual(len(list((self.base / 'Data' / 'Workbook Backups').glob('*.xlsx'))), 2)

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

    def test_ui_has_draft_batch_and_alert_test_actions(self):
        app = self.app()
        labels = [str(control.cget("text")) for control in app.controls if control.winfo_class() == "TButton"]
        self.assertIn("Create one draft", labels)
        self.assertIn("Create all Ready drafts", labels)
        self.assertIn("Test alert", labels)
        self.assertFalse(any("test draft" in label.casefold() or "sample" in label.casefold() for label in labels))

    def test_test_alert_shows_notification_without_scanning_or_using_daily_state(self):
        app = self.app()
        settings_before = (self.base / "settings.local.json").read_bytes()
        with patch("desktop_app.run") as run, patch("workflow_support.show_centered_reminder", return_value="dismiss") as prompt:
            app.test_alert()
        prompt.assert_called_once_with(
            "Ecoreco certificate reminder — test",
            "This is what the scheduled reminder will look like. It summarizes the items that need review.",
            parent=app.root,
            test_mode=True,
        )
        run.assert_not_called()
        self.assertFalse((self.base / "Data" / "alert-state.json").exists())
        self.assertEqual((self.base / "settings.local.json").read_bytes(), settings_before)
        self.notifier.assert_not_called()
        self.assertIn("No workbook, schedule, or alert history was changed", app.status.get())

    def test_test_alert_reports_when_centered_prompt_cannot_open(self):
        app = self.app()
        with patch("workflow_support.show_centered_reminder", side_effect=RuntimeError("display unavailable")):
            app.test_alert()
        self.assertIn("Could not show the test reminder", app.status.get())
        self.assertIn("display unavailable", app.status.get())

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

    def test_reminder_picker_saves_afternoon_as_24_hour_and_updates_current_users_task(self):
        app = self.app()
        save_json(self.base / "Data" / "alert-state.json", {
            "schema_version": 3,
            "dismissed_until": "2026-09-24",
            "last_attempt_day": "2026-09-24",
            "last_attempt_pending": True,
            "active": {"pending-case": "plant1"},
            "system_active": {},
            "date_quality_active": {},
            "seen": [],
            "system_seen": [],
            "date_quality_seen": [],
            "processed_runs": [],
            "snooze_pending": False,
        })
        app.reminder_hour.set("5")
        app.reminder_minute.set("45")
        app.reminder_period.set("PM")
        with patch.object(app, "launch_worker") as launch:
            app.apply_reminder_time()
        self.assertEqual(load_settings(self.base)["reminder_time"], "17:45")
        self.assertIn("5:45 PM", app.status.get())
        launch.assert_called_once_with(app.reminder_schedule_worker, ("17:45",))
        worker, args = launch.call_args.args
        with patch("desktop_app.subprocess.run", return_value=subprocess.CompletedProcess([], 0, "updated", "")) as run:
            worker(*args)
        command = run.call_args.args[0]
        self.assertIn("-Mode", command)
        self.assertIn("Reschedule", command)
        self.assertIn("17:45", command)
        saved_alert_state = json.loads((self.base / "Data" / "alert-state.json").read_text(encoding="utf-8"))
        self.assertEqual(saved_alert_state["dismissed_until"], "")
        self.assertFalse(saved_alert_state["last_attempt_pending"])
        self.assertEqual(saved_alert_state["active"], {"pending-case": "plant1"})
        self.assertEqual(app.events.get_nowait(), ("done", "Daily check and reminder set for 5:45 PM."))

    def test_failed_schedule_update_does_not_clear_existing_dismissal(self):
        app = self.app()
        save_json(self.base / "Data" / "alert-state.json", {
            "schema_version": 3,
            "dismissed_until": "2026-09-24",
            "active": {},
            "system_active": {},
            "date_quality_active": {},
            "seen": [],
            "system_seen": [],
            "date_quality_seen": [],
            "processed_runs": [],
            "snooze_pending": False,
        })
        failure = subprocess.CompletedProcess([], 1, "", "task scheduler denied")
        with patch("desktop_app.subprocess.run", return_value=failure):
            app.reminder_schedule_worker("17:45")
        saved = json.loads((self.base / "Data" / "alert-state.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["dismissed_until"], "2026-09-24")
        self.assertEqual(app.events.get_nowait()[0], "error")

    def test_invalid_reminder_time_is_not_saved_or_scheduled(self):
        app = self.app()
        before = (self.base / "settings.local.json").read_bytes()
        app.reminder_hour.set("13")
        app.reminder_minute.set("10")
        app.reminder_period.set("PM")
        with patch.object(app, "launch_worker") as launch:
            app.apply_reminder_time()
        self.assertEqual((self.base / "settings.local.json").read_bytes(), before)
        launch.assert_not_called()

    def test_reminder_picker_handles_midnight_noon_and_the_user_selected_time(self):
        self.assertEqual(desktop_app.reminder_picker_parts("00:00"), ("12", "00", "AM"))
        self.assertEqual(desktop_app.reminder_picker_parts("12:00"), ("12", "00", "PM"))
        self.assertEqual(desktop_app.reminder_picker_parts("15:10"), ("3", "10", "PM"))
        self.assertEqual(desktop_app.reminder_time_from_picker("3", "10", "PM"), "15:10")
        self.assertEqual(desktop_app.reminder_time_from_picker("12", "10", "AM"), "00:10")
        app = self.app()
        self.assertEqual(str(app.reminder_hour_picker.cget("state")), "readonly")
        self.assertEqual(str(app.reminder_minute_picker.cget("state")), "readonly")
        self.assertEqual(str(app.reminder_period_picker.cget("state")), "readonly")
        app.set_busy(True)
        self.assertEqual(str(app.reminder_period_picker.cget("state")), "disabled")
        app.set_busy(False)
        self.assertEqual(str(app.reminder_period_picker.cget("state")), "readonly")

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

    def test_manual_check_refreshes_alert_records_without_acknowledging_the_daily_popup(self):
        app = self.app()
        result = {"run_id": "manual-check", "mode": "Check sheets", "created": 0,
                  "updated": 0, "rows": [], "scan_complete": True}
        save_json(self.base / "Results.json", result)
        with patch("desktop_app.run", return_value=(0, 0, self.base)), \
             patch("workflow_support.notify_once") as notify:
            app.worker(self.paths, "", False, None, True)
        notify.assert_called_once()
        self.assertTrue(notify.call_args.kwargs["record_only"])
        self.assertEqual(notify.call_args.kwargs["reminder_time"], "09:00")

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
        notify.assert_called_once()
        self.assertEqual(notify.call_args.args[:2], (self.base, self.result))
        self.assertTrue(callable(notify.call_args.args[2]))
        self.assertEqual(notify.call_args.kwargs, {"enabled": False, "repeat_pending": True, "reminder_time": "09:00"})
        self.assertTrue(self.run.call_args.kwargs["create_drafts"])
        self.com.CoInitialize.assert_called_once()
        self.com.CoUninitialize.assert_called_once()

    def test_snoozed_run_refreshes_read_only_without_outlook_or_draft_creation(self):
        save_json(self.base / "settings.local.json", dict(self.settings, quiet_notifications=True))
        with patch("workflow_support.notify_once") as notify, patch("workflow_support.cancel_snoozed_reminder") as cancel:
            self.assertEqual(scheduled_run.main(self.base, reminder_only=True), 0)
        self.assertEqual(notify.call_count, 2)
        self.assertEqual(notify.call_args_list[0].args[:2], (self.base, self.result))
        self.assertTrue(notify.call_args_list[0].kwargs["record_only"])
        self.assertIs(notify.call_args_list[0].kwargs["enabled"], True)
        self.assertEqual(notify.call_args_list[1].args[:2], (self.base, None))
        self.assertTrue(callable(notify.call_args_list[1].args[2]))
        self.assertEqual(notify.call_args_list[1].kwargs, {"enabled": True, "repeat_pending": True, "reminder_only": True})
        cancel.assert_called_once_with(self.base, force=True)
        self.run.assert_called_once_with(self.settings["workbook_paths"], "", self.base,
                                         create_drafts=False, check_only=True, progress=unittest.mock.ANY)
        self.assertFalse(self.run.call_args.kwargs["create_drafts"])
        self.assertTrue(self.run.call_args.kwargs["check_only"])
        self.com.CoInitialize.assert_not_called()
        self.com.CoUninitialize.assert_not_called()

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
$ReminderTime='09:00'
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

    def test_script_has_safe_reschedule_mode_and_uses_saved_time_when_installing(self):
        script = (ROOT / "Install Daily Task.ps1").read_text()
        self.assertIn("'Reschedule'", script)
        self.assertIn("Get-ScheduledTask -TaskName $taskName", script)
        self.assertIn("$existingTask.Triggers = $triggers", script)
        self.assertIn("$existingTask.Description =", script)
        self.assertIn("Set-ScheduledTask -InputObject $existingTask", script)
        self.assertNotIn("Set-ScheduledTask -TaskName $taskName -Trigger $triggers -Description", script)
        parameters = self.powershell("(Get-Command Set-ScheduledTask).Parameters.Keys -join ','")
        self.assertIn("InputObject", parameters)
        self.assertNotIn("Description", parameters)
        self.assertIn("ConvertFrom-Json", script)
        self.assertIn("New-ScheduledTaskTrigger -Daily -At $ReminderTime", script)

    def test_one_hour_task_is_interactive_one_shot_and_does_not_wake_pc(self):
        source = (ROOT / "Install Daily Task.ps1").read_text(encoding="utf-8")
        self.assertIn("'Snooze'", source)
        self.assertIn("New-ScheduledTaskTrigger -Once -At $runAt", source)
        self.assertIn("-StartWhenAvailable", source)
        self.assertIn("-LogonType Interactive", source)
        self.assertIn("--snoozed-reminder", source)
        snooze_section = source.split("if ($Mode -eq 'Snooze')", 1)[1].split("if (-not $ReminderTime)", 1)[0]
        self.assertNotIn("-WakeToRun", snooze_section)

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
            parent = ROOT.parent if ROOT.name == 'App' else ROOT
            launcher_path = parent / name
            if not launcher_path.exists():
                launcher_path = parent / 'Setup' / name
            content = launcher_path.read_text()
            self.assertIn(f"-Mode {mode}", content)
            self.assertNotIn("pip install", content)
            self.assertNotIn("ExecutionPolicy Bypass", content)
        launcher = (ROOT / "Install Daily Task.ps1").read_text(encoding="utf-8")
        self.assertIn("validate_daily_setup.py", launcher)
        self.assertLess(launcher.index("if ($Mode -eq 'Diagnose')"), launcher.index("# Validate the saved setup without creating drafts or contacting Outlook."))


if __name__ == "__main__":
    unittest.main()
