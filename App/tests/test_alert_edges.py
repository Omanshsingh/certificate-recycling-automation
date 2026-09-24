"""Notification failure/recovery tests. No test displays a real notification."""
import json
from datetime import date, datetime, timedelta
from pathlib import Path
import tempfile
import sys
import types
import os
import unittest
from unittest.mock import Mock, patch

from workflow_support import Report, _alert_state, cancel_snoozed_reminder, load_settings, notify_once, rearm_alerts_after_schedule_change, run_lock, save_json, schedule_one_hour_reminder, windows_notification
from workbook_store import exclusive_file


TODAY = date(2026, 9, 23)


def issue(plant="plant1", sn="1", detail="Certificate Issued is blank", **extra):
    row = dict(plant=plant, sn=sn, detail=detail, outcome="needs_review", reminder_due=True,
               inward_date=(TODAY - timedelta(days=30)).isoformat(), age_days=30,
               inward_date_quality="valid", certificate_status="blank")
    row.update(extra)
    return row


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

    def test_saved_report_sets_the_30_day_flag_from_the_inward_date(self):
        source = {"source_id": "plant1", "source_path": "Plant 1.xlsx",
                  "source_sheet": "Inward Register", "source_row": 2,
                  "case_id": "1001", "inward_date": TODAY - timedelta(days=30),
                  "cod_no": "0042"}
        report_file = Report(self.root, TODAY, "Outlook drafts", publish_latest=False)
        for status, expected in (("Blank", True), ("Completed", True), ("N/A", False), ("Sent", False)):
            report_file.add(dict(source, issued=status), "needs_review", "status check")
            row = report_file.rows[-1]
            self.assertEqual(row["age_days"], 30)
            self.assertEqual(row["reminder_due"], expected)
            self.assertEqual(row["certificate_no"], "0042")

    def test_report_classifies_missing_invalid_future_and_valid_inward_dates(self):
        cases = ((None, "missing", None), ("", "missing", None),
                 ("not a date", "invalid", None),
                 (TODAY + timedelta(days=1), "future", -1),
                 (TODAY - timedelta(days=5), "valid", 5))
        report_file = Report(self.root, TODAY, "Outlook drafts", publish_latest=False)
        for index, (inward_date, expected_quality, expected_age) in enumerate(cases, 1):
            report_file.add({"source_id": "plant1", "case_id": str(index),
                             "inward_date": inward_date, "issued": "Ready"},
                            "needs_review", "field review")
            row = report_file.rows[-1]
            self.assertEqual(row["inward_date_quality"], expected_quality)
            self.assertEqual(row["age_days"], expected_age)
            self.assertFalse(row["reminder_due"])

    def test_new_or_young_blank_rows_do_not_popup_but_automation_failures_do(self):
        young = dict(issue(sn="young"), reminder_due=False, system_alert=False)
        self.assertFalse(notify_once(self.root, report("young", [young]), self.notify, TODAY))
        self.assertEqual(self.messages, [])
        failure = dict(issue(plant="Setup", sn="", reminder_due=False,
                             issue_code="source_unavailable", system_alert=True))
        self.assertTrue(notify_once(self.root, report("failure", [failure]), self.notify, TODAY))
        self.assertIn("automation problem needs attention", self.messages[0])

    def test_daily_pending_digest_stops_after_resolution_and_respects_mute(self):
        result = report('pending', [issue()])
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY, repeat_pending=True))
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY, repeat_pending=True))
        self.assertTrue(notify_once(self.root, report('next', [issue()]), self.notify, TODAY + timedelta(days=1), repeat_pending=True))
        self.assertFalse(notify_once(self.root, report('muted', [issue()]), self.notify, TODAY + timedelta(days=2), enabled=False, repeat_pending=True))
        self.assertFalse(notify_once(self.root, report('resolved'), self.notify, TODAY + timedelta(days=3), repeat_pending=True))
        self.assertEqual(len(self.messages), 2)

    def test_date_quality_digest_counts_missing_and_future_dates_but_excludes_sent_na(self):
        rows = [
            issue(sn="missing", inward_date="", age_days=None, reminder_due=False,
                  inward_date_quality="missing", supplier="Private Client A"),
            issue(sn="future", inward_date=(TODAY + timedelta(days=1)).isoformat(),
                  age_days=-1, reminder_due=False, inward_date_quality="future",
                  supplier="Private Client B"),
            issue(sn="na", inward_date="", age_days=None, reminder_due=False,
                  inward_date_quality="missing", certificate_status="N/A"),
            issue(sn="sent", inward_date=(TODAY + timedelta(days=2)).isoformat(),
                  age_days=-2, reminder_due=False, inward_date_quality="future",
                  certificate_status="Sent"),
        ]
        result = report("date-quality", rows, scan_complete=True,
                        observed_sources=["plant1"])
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY,
                                    repeat_pending=True))
        message = self.messages[0]
        self.assertIn("2 inward records have missing, invalid, or future Inward Dates", message)
        self.assertIn("cannot be age-classified", message)
        self.assertNotIn("Private Client", message)
        self.assertNotIn("SN", message)
        state = self.state()
        self.assertEqual(len(state["date_quality_active"]), 2)
        self.assertEqual(state["schema_version"], 3)

    def test_invalid_legacy_report_dates_are_classified_without_case_details(self):
        legacy_rows = [
            dict(plant="plant1", file="private.xlsx", sheet="Register", row=2,
                 sn="legacy-missing", inward_date="", age_days=None,
                 certificate_status="Ready", reminder_due=False),
            dict(plant="plant1", file="private.xlsx", sheet="Register", row=3,
                 sn="legacy-invalid", inward_date="bad date", age_days=None,
                 certificate_status="Completed", reminder_due=False),
            dict(plant="plant1", file="private.xlsx", sheet="Register", row=4,
                 sn="legacy-future", inward_date=(TODAY + timedelta(days=1)).isoformat(),
                 age_days=-1, certificate_status="Blank", reminder_due=False),
        ]
        notify_once(self.root, report("legacy-date-shapes", legacy_rows,
                                       scan_complete=True, observed_sources=["plant1"]),
                    self.notify, TODAY, repeat_pending=True)
        self.assertIn("3 inward records have missing, invalid, or future Inward Dates",
                      self.messages[0])
        self.assertNotIn("private.xlsx", self.messages[0])
        self.assertNotIn("legacy-", self.messages[0])

    def test_full_scan_resolves_date_quality_and_a_later_recurrence_is_new(self):
        missing = issue(sn="missing", inward_date="", age_days=None,
                        reminder_due=False, inward_date_quality="missing")
        notify_once(self.root, report("quality-1", [missing], scan_complete=True,
                                      observed_sources=["plant1"]), self.notify, TODAY,
                    repeat_pending=True)
        self.assertEqual(len(self.state()["date_quality_active"]), 1)
        valid = issue(sn="missing", inward_date=(TODAY - timedelta(days=10)).isoformat(),
                      age_days=10, reminder_due=False, inward_date_quality="valid")
        self.assertFalse(notify_once(
            self.root, report("quality-resolved", [valid], scan_complete=True,
                              observed_sources=["plant1"]), self.notify,
            TODAY + timedelta(days=1), repeat_pending=True))
        self.assertEqual(self.state()["date_quality_active"], {})
        self.assertEqual(len(self.messages), 1)
        notify_once(self.root, report("quality-returns", [missing], scan_complete=True,
                                      observed_sources=["plant1"]), self.notify,
                    TODAY + timedelta(days=2), repeat_pending=True)
        self.assertEqual(len(self.messages), 2)

    def test_partial_scan_retains_unobserved_date_quality_until_full_resolution(self):
        plant2_missing = issue(plant="plant2", sn="22", inward_date="", age_days=None,
                               reminder_due=False, inward_date_quality="missing")
        notify_once(self.root, report("quality-plant2", [plant2_missing],
                                      scan_complete=True,
                                      observed_sources=["plant1", "plant2"]),
                    self.notify, TODAY, repeat_pending=True)
        partial = report("quality-partial", [], scan_complete=False,
                         observed_sources=["plant1"])
        notify_once(self.root, partial, self.notify, TODAY + timedelta(days=1),
                    repeat_pending=True)
        self.assertEqual(len(self.state()["date_quality_active"]), 1)
        self.assertIn("1 inward record has a missing, invalid, or future Inward Date",
                      self.messages[-1])
        recovered = issue(plant="plant2", sn="22",
                          inward_date=(TODAY - timedelta(days=4)).isoformat(),
                          age_days=4, reminder_due=False, inward_date_quality="valid")
        self.assertFalse(notify_once(
            self.root, report("quality-full-recovery", [recovered], scan_complete=True,
                              observed_sources=["plant1", "plant2"]), self.notify,
            TODAY + timedelta(days=2), repeat_pending=True))
        self.assertEqual(self.state()["date_quality_active"], {})

    def test_muted_date_quality_does_not_backlog_after_unmuting(self):
        missing = issue(sn="muted-missing", inward_date="", age_days=None,
                        reminder_due=False, inward_date_quality="missing")
        self.assertFalse(notify_once(
            self.root, report("muted-quality", [missing], scan_complete=True,
                              observed_sources=["plant1"]), self.notify, TODAY,
            enabled=False))
        self.assertFalse(notify_once(
            self.root, report("unmuted-quality", [missing], scan_complete=True,
                              observed_sources=["plant1"]), self.notify,
            TODAY + timedelta(days=1)))
        self.assertEqual(self.messages, [])
        self.assertEqual(len(self.state()["date_quality_seen"]), 1)

    def test_date_quality_deduplication_and_run_replay_are_idempotent(self):
        missing = issue(sn="same-case", inward_date="", age_days=None,
                        reminder_due=False, inward_date_quality="missing")
        duplicate = report("quality-idempotent", [missing, missing],
                           scan_complete=True, observed_sources=["plant1"])
        notify_once(self.root, duplicate, self.notify, TODAY, repeat_pending=True)
        self.assertIn("1 inward record has a missing, invalid, or future Inward Date",
                      self.messages[0])
        self.assertEqual(len(self.state()["date_quality_active"]), 1)
        self.assertFalse(notify_once(self.root, duplicate, self.notify, TODAY,
                                     repeat_pending=True))
        self.assertTrue(notify_once(self.root, duplicate, self.notify,
                                    TODAY + timedelta(days=1), repeat_pending=True))
        self.assertEqual(len(self.messages), 2)
        self.assertEqual(len(self.state()["date_quality_active"]), 1)

    def test_schema_two_state_migrates_without_losing_existing_alerts(self):
        path = self.root / "Data" / "alert-state.json"
        path.parent.mkdir(parents=True)
        legacy = {"schema_version": 2, "last_day": TODAY.isoformat(),
                  "snooze_pending": True, "processed_runs": ["old-run"],
                  "active": {"overdue-id": "plant1"},
                  "system_active": {"failure-id": "Setup"},
                  "seen": ["overdue-id"], "system_seen": ["failure-id"]}
        path.write_text(json.dumps(legacy), encoding="utf-8")
        migrated = _alert_state(path, TODAY)
        self.assertEqual(migrated["schema_version"], 3)
        self.assertEqual(migrated["active"], legacy["active"])
        self.assertEqual(migrated["system_active"], legacy["system_active"])
        self.assertEqual(migrated["seen"], legacy["seen"])
        self.assertEqual(migrated["system_seen"], legacy["system_seen"])
        self.assertEqual(migrated["processed_runs"], ["old-run"])
        self.assertTrue(migrated["snooze_pending"])
        self.assertEqual(migrated["date_quality_active"], {})
        self.assertEqual(migrated["date_quality_seen"], [])

    def test_date_quality_and_system_failures_share_one_popup_but_keep_separate_counts(self):
        rows = [issue(sn="bad-date", inward_date="", age_days=None,
                      reminder_due=False, inward_date_quality="missing",
                      supplier="Private Client")]
        rows.append(issue(plant="Setup", sn="", detail="private technical detail",
                          reminder_due=False, system_alert=True,
                          issue_code="source_unavailable"))
        notify_once(self.root, report("combined-categories", rows), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        self.assertIn("1 inward record has a missing, invalid, or future Inward Date",
                      self.messages[0])
        self.assertIn("1 automation problem needs attention", self.messages[0])
        self.assertNotIn("Private Client", self.messages[0])
        self.assertNotIn("private technical detail", self.messages[0])

    def test_daily_reminder_waits_until_the_user_selected_local_time(self):
        pending = report("before-reminder", [issue()])
        self.assertFalse(notify_once(self.root, pending, self.notify, TODAY, repeat_pending=True,
                                     reminder_time="17:30", current_time="09:00"))
        self.assertEqual(self.messages, [])
        self.assertTrue(notify_once(self.root, report("at-reminder", [issue()]), self.notify, TODAY,
                                    repeat_pending=True, reminder_time="17:30", current_time="17:30"))
        self.assertEqual(len(self.messages), 1)

    def test_manual_check_refresh_does_not_consume_the_scheduled_reminder(self):
        pending = report("manual-check", [issue()])
        self.assertFalse(notify_once(self.root, pending, self.notify, TODAY, repeat_pending=True,
                                     reminder_time="15:10", current_time="15:07", record_only=True))
        self.assertEqual(self.state()["last_day"], "")
        self.assertEqual(self.messages, [])
        self.assertTrue(notify_once(self.root, report("scheduled-check", [issue()]), self.notify, TODAY,
                                    repeat_pending=True, reminder_time="15:10", current_time="15:10"))
        self.assertEqual(len(self.messages), 1)

    def test_no_reminds_tomorrow_without_marking_work_complete(self):
        pending = report("no-today", [issue()])
        self.assertTrue(notify_once(self.root, pending, lambda *_: "dismiss", TODAY, repeat_pending=True))
        self.assertFalse(notify_once(self.root, report("same-day", [issue()]), self.notify, TODAY, repeat_pending=True))
        self.assertEqual(len(self.state()["active"]), 1)
        self.assertTrue(notify_once(self.root, report("tomorrow", [issue()]), self.notify,
                                    TODAY + timedelta(days=1), repeat_pending=True))
        self.assertEqual(len(self.messages), 1)
        self.assertIn("1 inward record is 30+ days old", self.messages[0])

    def test_saving_reminder_schedule_rearms_a_same_day_dismissal(self):
        self.assertTrue(notify_once(
            self.root, report("dismiss-before-reschedule", [issue()]),
            lambda *_: "dismiss", TODAY, repeat_pending=True,
        ))
        self.assertFalse(notify_once(
            self.root, report("still-dismissed", [issue()]), self.notify,
            TODAY, repeat_pending=True,
        ))
        rearm_alerts_after_schedule_change(self.root)
        self.assertTrue(notify_once(
            self.root, report("after-reschedule", [issue()]), self.notify,
            TODAY, repeat_pending=True,
        ))
        self.assertEqual(len(self.messages), 1)

    def test_unresolved_reminder_can_be_shown_more_than_once_the_same_day(self):
        for run_id in ("repeat-1", "repeat-2", "repeat-3"):
            self.assertTrue(notify_once(
                self.root, report(run_id, [issue()]), lambda *_: "review", TODAY,
                repeat_pending=True,
            ))
        self.assertEqual(self.state()["last_day"], TODAY.isoformat())
        self.assertEqual(self.state()["last_action"], "review")

    def test_legacy_last_day_does_not_suppress_a_due_alert(self):
        path = self.root / "Data" / "alert-state.json"
        save_json(path, {"schema_version": 3, "last_day": TODAY.isoformat(),
                         "active": {}, "system_active": {}, "date_quality_active": {},
                         "seen": [], "system_seen": [], "date_quality_seen": [],
                         "processed_runs": [], "snooze_pending": False})
        self.assertTrue(notify_once(self.root, report("legacy-state", [issue()]),
                                    self.notify, TODAY, repeat_pending=True))
        self.assertEqual(len(self.messages), 1)

    def test_one_hour_snooze_replays_summary_and_pauses_regular_alerts(self):
        pending_draft = dict(plant="plant1", file="Plant 1.xlsx", sheet="Register", sn="8",
                             outcome="draft_created", awaiting_send=True,
                             inward_date=(TODAY - timedelta(days=30)).isoformat(),
                             age_days=30, inward_date_quality="valid",
                             certificate_status="blank")
        first = report("snooze-initial", [issue(), pending_draft], created=1)
        self.assertTrue(notify_once(self.root, first, lambda *_: "snooze", TODAY, repeat_pending=True))
        state = self.state()
        self.assertEqual(state["last_day"], TODAY.isoformat())
        self.assertEqual(state["last_action"], "snooze")
        self.assertNotIn("pending_created", state)
        self.assertFalse(notify_once(self.root, report("during-snooze", [issue()]),
                                     self.notify, TODAY, repeat_pending=True))
        self.assertEqual(self.messages, [])

        self.assertTrue(notify_once(self.root, None, self.notify, TODAY, repeat_pending=True, reminder_only=True))
        self.assertIn("1 inward record is 30+ days old", self.messages[0])
        self.assertNotIn("draft", self.messages[0].casefold())
        self.assertFalse(notify_once(self.root, None, self.notify, TODAY, repeat_pending=True, reminder_only=True))
        self.assertEqual(len(self.messages), 1)

    def test_snoozed_summary_is_suppressed_after_the_work_is_resolved(self):
        notify_once(self.root, report("unresolved", [issue()]), lambda *_: "snooze", TODAY, repeat_pending=True)
        self.assertFalse(notify_once(self.root, report("resolved", []), self.notify, TODAY, repeat_pending=True))
        self.assertFalse(notify_once(self.root, None, self.notify, TODAY, repeat_pending=True, reminder_only=True))
        self.assertEqual(self.messages, [])

    def test_one_hour_schedule_uses_local_clock_and_the_installed_user_task_script(self):
        now = datetime(2026, 9, 24, 9, 17, 42)
        with patch("workflow_support._powershell") as powershell:
            result = schedule_one_hour_reminder(self.root, now=now)
        self.assertEqual(result, datetime(2026, 9, 24, 10, 17, 0))
        powershell.assert_called_once_with(self.root, "Snooze", ("-SnoozeAt", "2026-09-24T10:17:00"))

    def test_failed_task_registration_rolls_back_saved_snooze_flag(self):
        path = self.root / "Data" / "alert-state.json"
        save_json(path, {"last_day": TODAY.isoformat(), "snooze_pending": False})
        with patch("workflow_support._powershell", side_effect=RuntimeError("Task Scheduler blocked")):
            with self.assertRaisesRegex(RuntimeError, "Task Scheduler blocked"):
                schedule_one_hour_reminder(self.root, now=datetime(2026, 9, 24, 9, 0))
        state = json.loads(path.read_text())
        self.assertFalse(state["snooze_pending"])
        self.assertNotIn("snooze_due", state)

    def test_cancel_only_calls_windows_when_a_snooze_is_pending(self):
        path = self.root / "Data" / "alert-state.json"
        save_json(path, {"snooze_pending": False})
        with patch("workflow_support._powershell") as powershell:
            cancel_snoozed_reminder(self.root)
            powershell.assert_not_called()
            save_json(path, {"snooze_pending": True})
            cancel_snoozed_reminder(self.root)
        powershell.assert_called_once_with(self.root, "CancelSnooze")
        self.assertFalse(json.loads(path.read_text())["snooze_pending"])

    def test_prepared_draft_is_only_a_case_alert_when_overdue(self):
        draft = dict(plant='plant1', file='Plant 1.xlsx', sheet='Register', sn='88',
                     outcome='already_prepared', detail='COD 0062', awaiting_send=True,
                     certificate_status='completed', reminder_due=False,
                     inward_date=(TODAY - timedelta(days=10)).isoformat(), age_days=10,
                     inward_date_quality='valid')
        first = report('draft-1', [draft], scan_complete=True, observed_sources=['plant1', 'plant2'])
        self.assertFalse(notify_once(self.root, first, self.notify, TODAY, repeat_pending=True))
        self.assertEqual(self.messages, [])
        overdue = dict(draft, reminder_due=True)
        notify_once(self.root, report('draft-overdue', [overdue], scan_complete=True,
                                      observed_sources=['plant1', 'plant2']),
                    self.notify, TODAY, repeat_pending=True)
        self.assertIn('1 inward record is 30+ days old', self.messages[0])
        sent = report('draft-3', [dict(draft, certificate_status='sent', reminder_due=False,
                                       awaiting_send=False)], scan_complete=True,
                      observed_sources=['plant1', 'plant2'])
        self.assertFalse(notify_once(self.root, sent, self.notify, TODAY + timedelta(days=2), repeat_pending=True))
        self.assertEqual(len(self.messages), 1)

    def test_partial_scan_preserves_pending_draft_from_unavailable_plant(self):
        draft = dict(plant='plant2', file='Plant 2.xlsx', sn='8', outcome='already_prepared',
                     certificate_status='completed', reminder_due=True, awaiting_send=True,
                     inward_date=(TODAY - timedelta(days=30)).isoformat(), age_days=30,
                     inward_date_quality='valid')
        notify_once(self.root, report('draft', [draft], scan_complete=True, observed_sources=['plant1', 'plant2']),
                    self.notify, TODAY, repeat_pending=True)
        partial = report('partial', [issue('Setup', detail='Plant 2 unavailable')], scan_complete=False,
                         observed_sources=['plant1'])
        notify_once(self.root, partial, self.notify, TODAY + timedelta(days=1), repeat_pending=True)
        self.assertIn('plant2', self.state()['active'].values())
        self.assertIn('automation problem', self.messages[-1])

    def test_hundred_blank_rows_one_summary_without_client_details(self):
        rows = [issue(sn=str(i), supplier="Private Dummy Client", file="D:/confidential.xlsx") for i in range(100)]
        notify_once(self.root, report("1", rows), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        self.assertIn("100 inward records are 30+ days old", self.messages[0])
        self.assertNotIn("Private Dummy", self.messages[0])
        self.assertNotIn("confidential", self.messages[0])

    def test_new_drafts_alone_do_not_trigger_a_popup(self):
        result = report("1", created=4)
        self.assertFalse(notify_once(self.root, result, lambda *_: False, TODAY))
        self.assertFalse(notify_once(self.root, result, lambda *_: False, TODAY))
        self.assertNotIn("pending_created", self.state())
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY))
        self.assertEqual(self.messages, [])
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))

    def test_notification_exception_keeps_work_and_caps_unknown_attempt(self):
        result = report("1", [issue()], created=2)
        with patch("builtins.print"):
            self.assertFalse(notify_once(self.root, result, lambda *_: (_ for _ in ()).throw(RuntimeError("Windows unavailable")), TODAY))
        self.assertNotIn("pending_created", self.state())
        self.assertIn("Windows unavailable", self.state()["last_delivery_error"])
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY))
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))
        self.assertIn("1 inward record is 30+ days old", self.messages[0])
        self.assertNotIn("last_delivery_error", self.state())

    def test_repeated_scheduled_runs_can_show_the_same_overdue_summary_same_day(self):
        for run_id in ("repeat-1", "repeat-2", "repeat-3"):
            self.assertTrue(notify_once(self.root, report(run_id, [issue()]), self.notify,
                                        TODAY, repeat_pending=True))
        self.assertEqual(len(self.messages), 3)

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
        notify_once(self.root, report("1", [issue()], created=1), self.notify, TODAY)
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
        partial = report("2", [issue("Setup", sn="", detail="Plant 2 unavailable",
                                      reminder_due=False, system_alert=True,
                                      issue_code="source_unavailable")],
                         scan_complete=False, observed_sources=["plant1"])
        notify_once(self.root, partial, self.notify, TODAY + timedelta(days=1))
        self.assertEqual(sorted(self.state()["active"].values()), ["plant2"])
        notify_once(self.root, report("3", [issue("plant2")], scan_complete=True), self.notify, TODAY + timedelta(days=2))
        self.assertEqual(len(self.messages), 2)  # still the same plant 2 issue

    def test_new_issue_in_partial_scan_still_notifies(self):
        notify_once(self.root, report("1", [issue()], scan_complete=False, observed_sources=["plant1"]), self.notify, TODAY)
        self.assertIn("some files could not be checked", self.messages[0])

    def test_excel_recovery_count_alone_does_not_trigger_a_case_popup(self):
        self.assertFalse(notify_once(self.root, report("1", updated=3), self.notify, TODAY))
        self.assertEqual(self.messages, [])

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
        self.assertEqual(len(self.messages), 1)
        self.assertIn("automation problem needs attention", self.messages[0])
        self.assertEqual(len(list(path.parent.glob("alert-state.corrupt-*.json"))), 1)
        self.assertTrue((path.parent / "Notification status.txt").exists())
        notify_once(self.root, report("2", [issue()]), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)

    def test_wrong_state_types_recover_and_surface_one_actionable_warning(self):
        for bad in ([], {"seen": "wrong"}, {"active": []}, {"pending_created": "3"}, {"last_day": []}, {"last_day": "invalid"}):
            with self.subTest(bad=bad):
                save_json(self.root / "Data" / "alert-state.json", bad)
                notify_once(self.root, report(str(bad), [issue()]), self.notify, TODAY)
                self.assertNotIn("state_warning", self.state())
                self.assertIn("automation problem", self.messages[-1])
                self.assertTrue((self.root / "Data" / "Notification status.txt").exists())

    def test_remind_tomorrow_choice_survives_clock_rollback(self):
        notify_once(self.root, report("dismiss", [issue()], created=1),
                    lambda *_: "dismiss", TODAY, repeat_pending=True)
        self.assertFalse(notify_once(self.root, report("rollback", [issue(sn="2")]),
                                     self.notify, TODAY - timedelta(days=1), repeat_pending=True))
        self.assertTrue(notify_once(self.root, report("tomorrow", [issue(sn="2")]),
                                    self.notify, TODAY + timedelta(days=1), repeat_pending=True))
        self.assertEqual(len(self.messages), 1)

    def test_midnight_allows_new_summary(self):
        notify_once(self.root, report("1", [issue()], created=1), self.notify, TODAY, repeat_pending=True)
        notify_once(self.root, report("2", [issue()], created=1), self.notify, TODAY + timedelta(days=1), repeat_pending=True)
        self.assertEqual(len(self.messages), 2)

    def test_duplicate_issue_rows_count_once_and_same_sn_at_two_plants_is_distinct(self):
        notify_once(self.root, report("1", [issue(), issue(), issue("plant2")]), self.notify, TODAY)
        self.assertIn("2 inward records are 30+ days old", self.messages[0])

    def test_stable_issue_code_deduplicates_volatile_technical_details(self):
        notify_once(self.root, report("1", [issue(detail="Error in run A.pdf", issue_code="save_failed")]), self.notify, TODAY)
        notify_once(self.root, report("2", [issue(detail="Error in run B.pdf", issue_code="save_failed")]), self.notify, TODAY + timedelta(days=1))
        self.assertEqual(len(self.messages), 1)

    def test_concurrent_notification_pass_skips_without_popup(self):
        path = self.root / "Data" / "alert-state.lock"
        path.parent.mkdir()
        path.touch()
        with exclusive_file(path):
            self.assertFalse(notify_once(self.root, report("1", [issue()], created=1), self.notify, TODAY))
        self.assertEqual(self.messages, [])
        # A skipped notification pass must retain its counts for the next pass.
        notify_once(self.root, report("2", [issue()], created=2), self.notify, TODAY)
        self.assertIn("1 inward record is 30+ days old", self.messages[0])
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
        notify_once(self.root, report("1", [issue()], created=2), self.notify, TODAY)
        self.assertEqual(len(self.messages), 1)
        self.assertIn("automation problem", self.messages[0])
        self.assertIn("counts are unknown", (inbox.parent / "Notification status.txt").read_text())
        self.assertTrue((inbox / "damaged.invalid").exists())

    def test_cannot_save_state_never_calls_notifier(self):
        with patch("workflow_support.save_json", side_effect=PermissionError("read only")):
            with self.assertRaises(PermissionError):
                notify_once(self.root, report("1", [issue()], created=1), self.notify, TODAY)
        self.assertEqual(self.messages, [])

    def test_crash_during_notification_does_not_retry_same_day_or_lose_counts(self):
        result = report("1", [issue()], created=7)
        with self.assertRaises(SystemExit):
            notify_once(self.root, result, lambda *_: (_ for _ in ()).throw(SystemExit()), TODAY)
        self.assertFalse(notify_once(self.root, result, self.notify, TODAY))
        self.assertTrue(notify_once(self.root, result, self.notify, TODAY + timedelta(days=1)))
        self.assertIn("1 inward record is 30+ days old", self.messages[0])

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
        for values in ({"certificate_number_start": 0}, {"certificate_number_width": 0}, {"quiet_notifications": "false"}, {"initial_cod_confirmed": "yes"}, {"reminder_time": "25:99"}, {"workbook_paths": "file.xlsx"}, {"workbook_paths": [None]}, {"workbook_paths": ["1", "2", "3"]}, {"outlook_mailbox": ["mail"]}, {"certificate_number_prefix": "COD\n"}):
            with self.subTest(values=values):
                save_json(self.root / "settings.local.json", values)
                with self.assertRaises(ValueError):
                    load_settings(self.root)

    def test_bom_settings_supported(self):
        (self.root / "settings.local.json").write_text('{"workbook_paths": ["Plant 1.xlsx", "Plant 2.xlsx"]}', encoding="utf-8-sig")
        self.assertEqual(load_settings(self.root)["workbook_paths"], ["Plant 1.xlsx", "Plant 2.xlsx"])

    def test_reminder_time_defaults_to_nine_and_accepts_a_valid_time(self):
        self.assertEqual(load_settings(self.root)["reminder_time"], "09:00")
        save_json(self.root / "settings.local.json", {"reminder_time": "17:30"})
        self.assertEqual(load_settings(self.root)["reminder_time"], "17:30")

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
