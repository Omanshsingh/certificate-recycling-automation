"""One quiet daily pass; Windows Task Scheduler handles timing and missed runs."""
from pathlib import Path
from datetime import date
import argparse
import hashlib
import json
import os
import sys
import tempfile
import time
import uuid

from app_paths import DATA_ROOT
ROOT = DATA_ROOT


def diagnostic(root, message):
    """Best-effort diagnostics, including when the application folder is read-only."""
    root = Path(root)
    installation = hashlib.sha256(str(root.resolve()).casefold().encode()).hexdigest()[:12]
    fallback = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "Ecoreco Certificates" / installation
    for folder in (root / "Output", fallback):
        try:
            folder.mkdir(parents=True, exist_ok=True)
            path = folder / "Daily check error.txt"
            path.write_text(f"{date.today().isoformat()}\nProject: {root}\n{message}\n", encoding="utf-8")
            return path
        except OSError:
            continue
    if sys.stderr is not None:
        try:
            print(message, file=sys.stderr)
        except OSError:
            pass
    return None


def failure_result(root, exc):
    message = str(exc)
    if isinstance(exc, ImportError):
        message = "A required Python package is missing. Run Setup Once.bat before retrying the daily check. " + message
    try:
        from workflow_support import Report
        report = Report(root, date.today(), "Daily check")
        report.add({}, "needs_review", message)
        _, _, output = report.finish()
        return json.loads((output / "Results.json").read_text(encoding="utf-8"))
    except Exception as report_error:
        diagnostic(root, message + "\nCould not save the normal results: " + str(report_error))
        return {"run_id": uuid.uuid4().hex, "mode": "Daily check", "scan_complete": False, "observed_sources": [],
                "created": 0, "updated": 0, "review": 1,
                "rows": [{"plant": "Setup", "sn": "", "outcome": "needs_review", "detail": message}]}


def check_with_setup_issue(run, root, paths, mailbox, exc):
    """Keep source observations when Outlook is unavailable for this pass."""
    from workflow_support import Report
    _, _, output = run(paths, mailbox, root, create_drafts=False, check_only=True, progress=lambda _: None)
    result = json.loads((output / "Results.json").read_text(encoding="utf-8"))
    report = Report(root, date.today(), "Daily check (sheets only)")
    report.observed_sources = result.get("observed_sources", [])
    report.scan_complete = result.get("scan_complete", False)
    report.rows = list(result["rows"])
    message = str(exc)
    if isinstance(exc, ImportError):
        message = "A required Outlook package is missing. Run Setup Once.bat to enable drafts. " + message
    report.add({}, "needs_review", message)
    _, _, output = report.finish()
    return json.loads((output / "Results.json").read_text(encoding="utf-8"))


def _reminder_action(root, *, reminder_only=False):
    from workflow_support import (
        cancel_snoozed_reminder,
        launch_review_app,
        schedule_one_hour_reminder,
        show_centered_reminder,
    )
    selected = {"action": None}

    def present(title, message):
        action = show_centered_reminder(
            title,
            message,
            on_snooze=lambda: schedule_one_hour_reminder(root),
            on_dismiss=None if reminder_only else lambda: cancel_snoozed_reminder(root),
            on_review=lambda: launch_review_app(root),
        )
        selected["action"] = action
        return action

    return present, selected


def main(root=None, reminder_only=False):
    root = Path(root) if root is not None else ROOT
    if reminder_only:
        try:
            from workflow_support import cancel_snoozed_reminder, load_settings, notify_once
            settings = load_settings(root)
            if not settings.get("quiet_notifications", True):
                cancel_snoozed_reminder(root, force=True)
                return 0
            from run_certificates import run
            result = None
            refresh_error = None
            for attempt in range(3):
                try:
                    _, _, output = run(
                        settings["workbook_paths"], "", root, create_drafts=False,
                        check_only=True, progress=lambda _: None,
                    )
                    result = json.loads((output / "Results.json").read_text(encoding="utf-8"))
                    break
                except RuntimeError as exc:
                    if "already running" not in str(exc).casefold() or attempt == 2:
                        refresh_error = exc
                        break
                    time.sleep(2)
                except Exception as exc:
                    refresh_error = exc
                    break
            if result is None:
                detail = "Could not refresh the two plant files before this reminder. " + str(refresh_error or "The read-only check did not complete.")
                result = {"run_id": "snooze-refresh-" + uuid.uuid4().hex, "mode": "Reminder refresh",
                          "scan_complete": False, "observed_sources": [], "created": 0, "updated": 0,
                          "rows": [{"plant": "Setup", "sn": "", "outcome": "needs_review", "detail": detail}]}
            # Refresh the saved work list first. This check is read-only: no COD
            # is allocated, no workbook is written, and Outlook is never opened.
            notify_once(root, result, lambda *_: False, enabled=True, repeat_pending=True, record_only=True)
            present, selected = _reminder_action(root, reminder_only=True)
            notify_once(root, None, present, enabled=True,
                        repeat_pending=True, reminder_only=True)
            # The one-shot task has fired. Remove its registration unless the
            # user chose to replace it with another one-hour snooze.
            if selected["action"] != "snooze":
                from workflow_support import cancel_snoozed_reminder
                cancel_snoozed_reminder(root, force=True)
            return 0
        except Exception as exc:
            message = "Could not show the snoozed reminder: " + str(exc)
            # The one-shot trigger has already fired. Do not leave a stale
            # snooze flag/task behind if refreshing the files or opening the
            # centered prompt fails; the regular daily pass remains responsible
            # for surfacing any outstanding work.
            try:
                from workflow_support import cancel_snoozed_reminder
                cancel_snoozed_reminder(root, force=True)
            except Exception as cleanup_error:
                message += "\nCould not clear the consumed one-hour reminder: " + str(cleanup_error)
            diagnostic(root, message)
            return 1
    initialized = False
    settings = {}
    notify_once = windows_notification = None
    failed = False
    try:
        # Imports belong inside the guard: a missing dependency must leave a clue
        # even when Task Scheduler starts pythonw without a console.
        from workflow_support import load_settings, notify_once
        settings = load_settings(root)
        paths = settings["workbook_paths"]
        if len(paths) != 2 or not all(paths):
            raise ValueError("Open Certificate Automation and save the two plant Excel locations before running the daily check.")
        files = [Path(p).resolve() for p in paths]
        if files[0] == files[1] or (all(p.is_file() for p in files) and files[0].samefile(files[1])):
            raise ValueError("Choose a different Excel file for each plant in Certificate Automation.")
        mailbox = settings.get("outlook_mailbox", "").strip()
        from run_certificates import run
        setup_error = None
        if not mailbox:
            setup_error = ValueError("Open Certificate Automation, load Outlook accounts and save the mailbox before running the daily check. Sheets were checked without creating drafts.")
        else:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                initialized = True
            except Exception as exc:
                setup_error = exc
        if setup_error is not None:
            result = check_with_setup_issue(run, root, paths, mailbox, setup_error)
            failed = True
        else:
            _, review, output = run(paths, mailbox, root, create_drafts=True, progress=lambda _: None)
            failed = bool(review)
            result = json.loads((output / "Results.json").read_text(encoding="utf-8"))
    except Exception as exc:
        if isinstance(exc, RuntimeError) and "already running" in str(exc):
            return 0
        failed = True
        result = failure_result(root, exc)
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as exc:
                diagnostic(root, "Could not release the Outlook connection: " + str(exc))
                failed = True
    if notify_once is not None:
        try:
            present, _ = _reminder_action(root)
            # Muted runs still refresh issue and draft follow-up state without
            # displaying a reminder. Unsent drafts remain pending to review.
            notify_once(root, result, present,
                        enabled=settings.get("quiet_notifications", True), repeat_pending=True,
                        reminder_time=settings.get("reminder_time", "09:00"))
        except Exception as exc:
            diagnostic(root, "Could not record the daily notification state: " + str(exc))
            failed = True
    return int(failed)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--snoozed-reminder", action="store_true")
    arguments, _ = parser.parse_known_args()
    raise SystemExit(main(reminder_only=arguments.snoozed_reminder))
