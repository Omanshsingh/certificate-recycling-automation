"""Notification failure/recovery tests. No test displays a real notification."""
import json
from datetime import date, timedelta
from pathlib import Path
import tempfile
import sys
import types
import os
import unittest
from unittest.mock import Mock, patch

from workflow_support import Report, load_settings, notify_once, run_lock, save_json, windows_notification
from workbook_store import exclusive_file


TODAY = date(2026, 9, 23)


def issue(plant="plant1", sn="1", detail="Certificate Issued is blank", **extra):
    return dict(plant=plant, sn=sn, detail=detail, outcome="needs_review", **extra)


def report(run_id, rows=(), created=0, updated=0, **extra):
    return dict(run_id=run_id, rows=list(rows), created=created, updated=updated, mode="Outlook drafts", **extra)


class AlertTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.messages = []

    def tearDown(self):
        self.temp.cleanup()

    def notify(self, title, message):
        self.messages.append(message)
        return True

    def state(self):
        return json.loads((self.root / "Data" / "alert-state.json").read_text())

    def test_no_work_no_popup(self):
        self.assertFalse(notify_once(self.root, report("1"), self.notify, TODAY))
        self.assertEqual(self.messages, [])

    def test_daily_pending_digest_stops_after_resolution_and_respects_mute(self):
        result = report('pending', [issue()])
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY, repeat_pending=True))
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY, repeat_pending=True))
        self.assertTrue(notify_once(self.root, report('next', [issue()]), self.notify, TODAY + timedelta(days=1), repeat_pending=True))
        self.assertFalse(notify_once(self.root, report('muted', [issue()]), self.notify, TODAY + timedelta(days=2), enabled=False, repeat_pending=True))
        self.assertFalse(notify_once(self.root, report('resolved'), self.notify, TODAY + timedelta(days=3), repeat_pending=True))
        self.assertEqual(len(self.messages), 2)

    def test_hundred_blank_rows_one_summary_without_client_details(self):
        rows = [issue(sn=str(i), supplier="Private Dummy Client", file="D:/confidential.xlsx") for i in range(100)]
        notify_once(self.root, report("1", rows), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        self.assertIn("100 issue(s)", self.messages[0])
        self.assertNotIn("Private Dummy", self.messages[0])
        self.assertNotIn("confidential", self.messages[0])

    def test_new_drafts_deduplicated_when_same_report_retried(self):
        result = report("1", created=4)
        self.assertFalse(notify_once(self.root, result, lambda *_: False, TODAY))
        self.assertFalse(notify_once(self.root, result, lambda *_: False, TODAY))
        self.assertEqual(self.state()["pending_created"], 4)
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY))
        self.assertIn("4 new draft(s)", self.messages[0])
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))

    def test_notification_exception_keeps_work_and_caps_unknown_attempt(self):
        result = report("1", [issue()], created=2)
        with patch("builtins.print"):
            self.assertFalse(notify_once(self.root, result, lambda *_: (_ for _ in ()).throw(RuntimeError("Windows unavailable")), TODAY))
        self.assertEqual(self.state()["pending_created"], 2)
        self.assertIn("Windows unavailable", self.state()["last_delivery_error"])
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY))
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))
        self.assertIn("2 new draft(s)", self.messages[0])
        self.assertNotIn("last_delivery_error", self.state())

    def test_many_runs_after_daily_cap_combine_once_next_day(self):
        notify_once(self.root, report("1", created=2), self.notify, TODAY)
        for i in range(2, 8):
            notify_once(self.root, report(str(i), [issue(sn=str(i))], created=3), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        notify_once(self.root, report("8", [issue(sn="7")]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 2)
        self.assertIn("18 new draft(s)", self.messages[-1])

    def test_unchanged_issue_quiet_for_many_days(self):
        for day in range(45):
            notify_once(self.root, report(str(day), [issue()]), self.notify, TODAY + timedelta(days=day))
        self.assertEqual(len(self.messages), 1)

    def test_issue_resolved_then_reappears_is_new(self):
        notify_once(self.root, report("1", [issue()]), self.notify, TODAY)
        notify_once(self.root, report("2"), self.notify, TODAY + timedelta(days=1))
        notify_once(self.root, report("3", [issue()]), self.notify, TODAY + timedelta(days=2))
        self.assertEqual(len(self.messages), 2)

    def test_pending_issue_resolved_before_next_day_does_not_raise_stale_popup(self):
        notify_once(self.root, report("1", created=1), self.notify, TODAY)
        notify_once(self.root, report("2", [issue()]), self.notify, TODAY)
        notify_once(self.root, report("3"), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 1)

    def test_replaying_old_report_does_not_resurrect_resolved_issue(self):
        old = report("1", [issue()])
        notify_once(self.root, old, self.notify, TODAY)
        notify_once(self.root, report("2"), self.notify, TODAY + timedelta(days=1))
        notify_once(self.root, old, self.notify, TODAY + timedelta(days=2))
        self.assertEqual(self.state()["active"], {})
        self.assertEqual(len(self.messages), 1)

    def test_partial_scan_preserves_unseen_plant_issues(self):
        notify_once(self.root, report("1", [issue(), issue("plant2")], scan_complete=True), self.notify, TODAY)
        partial = report("2", [issue("Setup", detail="Plant 2 unavailable")], scan_complete=False, observed_sources=["plant1"])
        notify_once(self.root, partial, self.notify, TODAY + timedelta(days=1))
        self.assertEqual(sorted(self.state()["active"].values()), ["Setup", "plant2"])
        notify_once(self.root, report("3", [issue("plant2")], scan_complete=True), self.notify, TODAY + timedelta(days=2))
        self.assertEqual(len(self.messages), 2)  # still the same plant 2 issue

    def test_new_issue_in_partial_scan_still_notifies(self):
        notify_once(self.root, report("1", [issue()], scan_complete=False, observed_sources=["plant1"]), self.notify, TODAY)
        self.assertIn("some files could not be checked", self.messages[0])

    def test_excel_recovery_message_does_not_claim_zero_new_drafts(self):
        notify_once(self.root, report("1", updated=3), self.notify, TODAY)
        self.assertIn("3 Excel update(s) recovered", self.messages[0])
        self.assertNotIn("new draft", self.messages[0])

    def test_preview_never_counts_as_new_drafts_or_clears_issues(self):
        notify_once(self.root, report("1", [issue()]), self.notify, TODAY)
        preview = report("2", created=1)
        preview["mode"] = "PDF preview"
        self.assertFalse(notify_once(self.root, preview, self.notify, TODAY + timedelta(days=1)))
        self.assertEqual(len(self.state()["active"]), 1)
        self.assertEqual(len(self.messages), 1)

    def test_muting_acknowledges_current_work_without_later_backlog(self):
        notify_once(self.root, report("1", [issue()], created=9), self.notify, TODAY, enabled=False)
        notify_once(self.root, report("2", [issue()]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(self.messages, [])
        notify_once(self.root, report("3", [issue(), issue(sn="2")]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 1)
        self.assertNotIn("9 new", self.messages[0])

    def test_corrupt_notification_file_is_backed_up_and_recovers_tomorrow(self):
        path = self.root / "Data" / "alert-state.json"
        path.parent.mkdir()
        path.write_text('{"partial"', encoding="utf-8")
        notify_once(self.root, report("1", [issue()]), self.notify, TODAY)
        self.assertEqual(self.messages, [])
        self.assertEqual(len(list(path.parent.glob("alert-state.corrupt-*.json"))), 1)
        self.assertTrue((path.parent / "Notification status.txt").exists())
        notify_once(self.root, report("2", [issue()]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 1)

    def test_wrong_state_types_recover_not_crash(self):
        for bad in ([], {"seen": "wrong"}, {"active": []}, {"pending_created": "3"}, {"last_day": []}, {"last_day": "invalid"}):
            with self.subTest(bad=bad):
                save_json(self.root / "Data" / "alert-state.json", bad)
                notify_once(self.root, report(str(bad), [issue()]), self.notify, TODAY)
                self.assertIn("state_warning", self.state())

    def test_clock_rollback_does_not_break_daily_cap(self):
        notify_once(self.root, report("1", created=1), self.notify, TODAY)
        notify_once(self.root, report("2", created=1), self.notify, TODAY - timedelta(days=1))
        self.assertEqual(len(self.messages), 1)
        notify_once(self.root, report("3"), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 2)

    def test_midnight_allows_new_summary(self):
        notify_once(self.root, report("1", created=1), self.notify, TODAY)
        notify_once(self.root, report("2", created=1), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 2)

    def test_duplicate_issue_rows_count_once_and_same_sn_at_two_plants_is_distinct(self):
        notify_once(self.root, report("1", [issue(), issue(), issue("plant2")]), self.notify, TODAY)
        self.assertIn("2 issue(s)", self.messages[0])

    def test_stable_issue_code_deduplicates_volatile_technical_details(self):
        notify_once(self.root, report("1", [issue(detail="Error in run A.pdf", issue_code="save_failed")]), self.notify, TODAY)
        notify_once(self.root, report("2", [issue(detail="Error in run B.pdf", issue_code="save_failed")]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 1)

    def test_concurrent_notification_pass_skips_without_popup(self):
        path = self.root / "Data" / "alert-state.lock"
        path.parent.mkdir()
        path.touch()
        with exclusive_file(path):
            self.assertFalse(notify_once(self.root, report("1", created=1), self.notify, TODAY))
        self.assertEqual(self.messages, [])
        # A skipped notification pass must retain its counts for the next pass.
        notify_once(self.root, report("2", created=2), self.notify, TODAY)
        self.assertIn("3 new draft(s)", self.messages[0])
        self.assertEqual(list((self.root / "Data" / "Alert inbox").glob("*.json")), [])

    def test_access_denied_is_not_misreported_as_another_running_job(self):
        with patch("workflow_support.exclusive_file", side_effect=PermissionError("Access denied")):
            with self.assertRaisesRegex(PermissionError, "Access denied"):
                with run_lock(self.root):
                    pass
            with self.assertRaisesRegex(PermissionError, "Access denied"):
                notify_once(self.root, report("1", created=1), self.notify, TODAY)
        self.assertEqual(self.messages, [])

    def test_damaged_queue_item_does_not_block_healthy_new_alerts(self):
        inbox = self.root / "Data" / "Alert inbox"
        inbox.mkdir(parents=True)
        (inbox / "damaged.json").write_text("not json")
        notify_once(self.root, report("1", created=2), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        self.assertIn("2 new draft(s)", self.messages[0])
        self.assertIn("counts are unknown", self.state()["state_warning"])
        self.assertTrue((inbox / "damaged.invalid").exists())

    def test_cannot_save_state_never_calls_notifier(self):
        with patch("workflow_support.save_json", side_effect=PermissionError("read only")):
            with self.assertRaises(PermissionError):
                notify_once(self.root, report("1", created=1), self.notify, TODAY)
        self.assertEqual(self.messages, [])

    def test_crash_during_notification_does_not_retry_same_day_or_lose_counts(self):
        result = report("1", created=7)
        with self.assertRaises(SystemExit):
            notify_once(self.root, result, lambda *_: (_ for _ in ()).throw(SystemExit()), TODAY)
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY))
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))
        self.assertIn("7 new draft(s)", self.messages[0])

    def test_invalid_report_counts_rejected_without_popup(self):
        for value in (-1, "2", True):
            with self.subTest(value=value), self.assertRaises(ValueError):
                notify_once(self.root, report(str(value), created=value), self.notify, TODAY)
        self.assertEqual(self.messages, [])


class SettingsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_corrupt_settings_not_overwritten_or_silently_reset(self):
        path = self.root / "settings.local.json"
        path.write_text("broken")
        with self.assertRaisesRegex(ValueError, "Repair or restore"):
            load_settings(self.root)
        self.assertEqual(path.read_text(), "broken")

    def test_invalid_settings_rejected(self):
        for values in ({"certificate_number_start": 0}, {"certificate_number_width": 0}, {"quiet_notifications": "false"}, {"initial_cod_confirmed": "yes"}, {"workbook_paths": "file.xlsx"}, {"workbook_paths": [None]}, {"workbook_paths": ["1", "2", "3"]}, {"outlook_mailbox": ["mail"]}, {"certificate_number_prefix": "COD\n"}):
            with self.subTest(values=values):
                save_json(self.root / "settings.local.json", values)
                with self.assertRaises(ValueError):
                    load_settings(self.root)

    def test_bom_settings_supported(self):
        (self.root / "settings.local.json").write_text('{"workbook_paths": ["Plant 1.xlsx", "Plant 2.xlsx"]}', encoding="utf-8-sig")
        self.assertEqual(load_settings(self.root)["workbook_paths"], ["Plant 1.xlsx", "Plant 2.xlsx"])

    def test_report_has_run_identity_coverage_and_timestamp(self):
        value = Report(self.root, TODAY, "Check sheets")
        value.observed_sources = ["plant1"]
        value.add({"source_id": "plant1", "case_id": "1"}, "needs_review", "blank", issue_code="blank_status")
        _, _, output = value.finish()
        result = json.loads((output / "Results.json").read_text())
        self.assertTrue(result["run_id"])
        self.assertTrue(result["completed_at"])
        self.assertFalse(result["scan_complete"])
        self.assertEqual(result["observed_sources"], ["plant1"])
        self.assertEqual(result["rows"][0]["issue_code"], "blank_status")


@unittest.skipUnless(os.name == "nt", "Windows notification API contract")
class WindowsNotificationContracts(unittest.TestCase):
    def setUp(self):
        self.gui = Mock()
        self.gui.WNDCLASS.side_effect = types.SimpleNamespace
        self.gui.NIF_ICON, self.gui.NIF_TIP, self.gui.NIF_INFO = 2, 4, 16
        self.gui.NIM_ADD, self.gui.NIM_DELETE = 0, 2
        self.gui.CreateWindow.return_value = 42
        self.gui.RegisterClass.return_value = 77
        self.modules = patch.dict(sys.modules, {
            "win32api": Mock(), "win32gui": self.gui,
            "win32con": types.SimpleNamespace(WM_DESTROY=2, IDI_INFORMATION=32516)})
        self.modules.start()
        self.addCleanup(self.modules.stop)
        self.clock = patch("workflow_support.time.monotonic", side_effect=[0, 9])
        self.clock.start()
        self.addCleanup(self.clock.stop)

    def test_silent_balloon_without_visible_window_is_cleaned_up(self):
        self.assertTrue(windows_notification("Dummy", "One summary"))
        args = self.gui.Shell_NotifyIcon.call_args_list[0].args
        self.assertEqual(args[0], self.gui.NIM_ADD)
        self.assertEqual(args[1][-1] & 0x10, 0x10)  # NIIF_NOSOUND
        self.assertEqual(self.gui.CreateWindow.call_args.args[2], 0)  # no WS_VISIBLE
        self.gui.DestroyWindow.assert_called_once_with(42)
        self.gui.UnregisterClass.assert_called_once()

    def test_window_creation_failure_releases_registered_class(self):
        self.gui.CreateWindow.side_effect = RuntimeError("window unavailable")
        with self.assertRaisesRegex(RuntimeError, "window unavailable"):
            windows_notification("Dummy", "One summary")
        self.gui.UnregisterClass.assert_called_once()
        self.gui.Shell_NotifyIcon.assert_not_called()

    def test_notification_failure_still_cleans_hidden_window(self):
        self.gui.Shell_NotifyIcon.side_effect = RuntimeError("notifier unavailable")
        with self.assertRaisesRegex(RuntimeError, "notifier unavailable"):
            windows_notification("Dummy", "One summary")
        self.gui.DestroyWindow.assert_called_once_with(42)
        self.gui.UnregisterClass.assert_called_once()


if __name__ == "__main__":
    unittest.main()
