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
from workflow_support import (
    cancel_snoozed_reminder,
    launch_review_app,
    load_settings,
    notify_once,
    open_alert_workbook,
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
        paths = settings.get("workbook_paths", ["", ""])
        if len(paths) != 2 or not all(paths):
            raise ValueError("Open the certificate app and save both plant workbooks in Settings.")
        from run_certificates import run
        result = None
        last_error = None
        for attempt in range(3):
            try:
                _, _, folder = run(paths, root, check_only=True, progress=lambda _: None)
                result = json.loads((folder / "Latest Results.json").read_text(encoding="utf-8"))
                break
            except RuntimeError as exc:
                if "already running" not in str(exc).casefold() or attempt == 2:
                    last_error = exc
                    break
                time.sleep(2)
        if result is None:
            raise last_error or RuntimeError("The plant check could not be completed.")

        def present(title, message):
            return show_centered_reminder(
                title, message,
                on_snooze=lambda: schedule_one_hour_reminder(root),
                on_dismiss=lambda: cancel_snoozed_reminder(root, force=True),
                on_open_workbook=lambda: open_alert_workbook(result),
                on_open_automation=lambda: launch_review_app(root),
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
