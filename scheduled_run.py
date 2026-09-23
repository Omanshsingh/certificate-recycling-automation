"""One quiet daily pass; Windows Task Scheduler handles timing and missed runs."""
from pathlib import Path
from datetime import date
import hashlib
import json
import os
import sys
import tempfile
import uuid

ROOT = Path(__file__).resolve().parent


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


def main(root=None):
    root = Path(root) if root is not None else ROOT
    initialized = False
    settings = {}
    notify_once = windows_notification = None
    failed = False
    try:
        # Imports belong inside the guard: a missing dependency must leave a clue
        # even when Task Scheduler starts pythonw without a console.
        from workflow_support import load_settings, notify_once, windows_notification
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
            # Muted runs must still acknowledge observations so unmuting cannot
            # deliver an unexpected backlog of old drafts or issues.
            notify_once(root, result, windows_notification, enabled=settings.get("quiet_notifications", True), repeat_pending=True)
        except Exception as exc:
            diagnostic(root, "Could not record the daily notification state: " + str(exc))
            failed = True
    return int(failed)


if __name__ == "__main__":
    raise SystemExit(main())
