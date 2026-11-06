"""Scheduled alert-only pass. It never creates certificates or contacts email."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import argparse
import json
import sys
import time
import traceback

from app_paths import DATA_ROOT
from workbook_store import LockBusyError
from workflow_support import (
    alert_workbook_targets,
    cancel_snoozed_reminder,
    dismiss_reminders_today,
    launch_review_app,
    load_settings,
    notify_once,
    open_alert_workbooks,
    schedule_one_hour_reminder,
    show_centered_reminder,
)


def main(root=None, reminder_only=False):
    root = Path(root or DATA_ROOT)
    try:
        settings = load_settings(root)
        if not settings.get("reminders_enabled", True):
            return 0
        reminder_time = settings.get("reminder_time", "11:00")
        if not reminder_only and datetime.now().strftime("%H:%M") < reminder_time:
            # The sign-in catch-up is only a catch-up after the selected time;
            # it must never make a morning login show the alert early.
            return 0
        try:
            dismissal = json.loads((root / 'Data' / 'reminder-state.json').read_text(encoding='utf-8'))
        except (OSError, ValueError):
            dismissal = {}
        if (isinstance(dismissal, dict) and dismissal.get('dismissed_day') == date.today().isoformat()
                and dismissal.get('dismissed_time', reminder_time) == reminder_time):
            if reminder_only:
                cancel_snoozed_reminder(root, force=True)
            return 0
        paths = settings.get("workbook_paths", ["", ""])
        connected = any(str(value).strip() for value in paths)
        if connected and (len(paths) != 2 or not all(paths)):
            raise ValueError("Open the certificate app and save both plant workbooks in Settings.")
        from run_certificates import run
        result = None
        last_error = None
        for attempt in range(61):
            try:
                snapshot = {'rows': [], 'scan_complete': False}
                if connected:
                    run(paths, root, check_only=True, progress=lambda _: None, report_sink=snapshot)
                from certificate_drafts import reconcile_manual, manual_rows
                reconcile_manual(root)
                snapshot['rows'] = [row for row in snapshot.get('rows', []) if row.get('source_kind') != 'manual'] + manual_rows(root)
                result = snapshot
                break
            except (LockBusyError, RuntimeError) as exc:
                if (not isinstance(exc, LockBusyError) and "already running" not in str(exc).casefold()) or attempt == 60:
                    last_error = exc
                    break
                time.sleep(10)
                if not load_settings(root).get('reminders_enabled', True):
                    return 0
        if result is None:
            raise last_error or RuntimeError("The plant check could not be completed.")

        def present(title, message):
            targets = alert_workbook_targets(result, paths)
            def review():
                launch_review_app(root)
                cancel_snoozed_reminder(root)
            return show_centered_reminder(
                title, message,
                on_snooze=lambda: schedule_one_hour_reminder(root),
                on_dismiss=lambda: dismiss_reminders_today(root),
                on_open_automation=review,
                on_open_workbook=lambda: open_alert_workbooks(result, paths),
                workbook_plants=[plant for plant, _ in targets],
                error_root=root,
            )

        displayed = notify_once(
            root, result, present, enabled=True, repeat_pending=reminder_only,
            reminder_time=reminder_time,
        )
        if reminder_only and not displayed:
            cancel_snoozed_reminder(root, force=True)
        return 0
    except Exception as exc:
        try:
            folder = root / "Data" / "Reports"
            folder.mkdir(parents=True, exist_ok=True)
            (folder / "Daily reminder error.txt").write_text(
                f"{date.today().isoformat()}\n{exc}\n\n{traceback.format_exc()}", encoding="utf-8",
            )
        except OSError:
            pass
        return 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--snoozed-reminder", action="store_true")
    args, _ = parser.parse_known_args()
    raise SystemExit(main(reminder_only=args.snoozed_reminder))
