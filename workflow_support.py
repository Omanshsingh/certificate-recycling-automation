"""Saved configuration, one running batch, reports and quiet notifications."""
from __future__ import annotations

from contextlib import contextmanager, suppress
from datetime import date, datetime
import hashlib
import json
import os
from pathlib import Path
import time
import uuid

from workbook_store import LockBusyError, exclusive_file


def load_settings(root):
    root = Path(root)
    defaults = root / "settings.example.json"
    local = root / "settings.local.json"
    def read(path):
        if not path.exists():
            return {}
        try:
            value = json.loads(path.read_text(encoding="utf-8-sig"))
            if not isinstance(value, dict):
                raise ValueError("expected a settings object")
            return value
        except (ValueError, OSError) as exc:
            raise ValueError(f"Cannot read settings in {path.name}. Repair or restore this file; saved settings have not been changed. ({exc})") from exc
    settings, saved = read(defaults), read(local)
    settings.update(saved)
    paths = saved.get("workbook_paths", settings.get("workbook_paths", []))
    if "workbook_paths" not in saved and saved.get("workbook_path"):
        paths = [saved["workbook_path"], ""]
    if not isinstance(paths, list) or len(paths) > 2 or any(not isinstance(p, str) for p in paths):
        raise ValueError("workbook_paths in settings.local.json must contain two file paths. Correct the settings file before continuing.")
    settings["workbook_paths"] = (paths + ["", ""])[:2]
    settings.pop("workbook_path", None)
    # Old installations can retain this setting; it no longer gates Ready rows.
    settings.pop("eligibility_days", None)
    for key, default, lower, upper in (("certificate_number_start", 1, 1, 999999999999), ("certificate_number_width", 4, 1, 18)):
        value = settings.get(key, default)
        if type(value) is not int or not lower <= value <= upper:
            raise ValueError(f"{key} in settings.local.json must be a whole number from {lower} to {upper}.")
        settings[key] = value
    for key, default in (("quiet_notifications", True), ("initial_cod_confirmed", False)):
        if not isinstance(settings.get(key, default), bool):
            raise ValueError(f"{key} in settings.local.json must be true or false, not text.")
        settings.setdefault(key, default)
    for key in ("certificate_number_prefix", "outlook_mailbox"):
        value = settings.get(key, "")
        if not isinstance(value, str) or any(ord(c) < 32 for c in value):
            raise ValueError(f"{key} in settings.local.json must be plain text without line breaks.")
        settings[key] = value
    return settings


def save_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temp.write_text(json.dumps(value, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


@contextmanager
def run_lock(root):
    path = Path(root) / "Data" / "run.lock"
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb"):
            pass
    except FileExistsError:
        pass
    context = exclusive_file(path)
    try:
        context.__enter__()
    except LockBusyError as exc:
        raise RuntimeError("Certificate processing is already running. Wait for it to finish.") from exc
    try:
        yield
    finally:
        context.__exit__(None, None, None)


class Report:
    def __init__(self, root, as_of, mode, publish_latest=True):
        self.root = Path(root)
        stamp = datetime.now().strftime("%H%M%S") + "_" + uuid.uuid4().hex[:6]
        self.output = self.root / "Output" / as_of.isoformat() / stamp
        self.output.mkdir(parents=True, exist_ok=True)
        self.mode = mode
        self.publish_latest = publish_latest
        self.run_id = uuid.uuid4().hex
        self.observed_sources = []
        self.scan_complete = False
        self.rows = []
        self.created = 0
        self.updated = 0

    def add(self, batch, outcome, detail, issue_code=None):
        row = {"plant": batch.get("source_id", "Setup"), "file": batch.get("source_path", ""), "sheet": batch.get("source_sheet", ""), "row": batch.get("source_row", ""), "sn": batch.get("case_id", ""), "supplier": batch.get("supplier", ""), "outcome": outcome, "detail": detail}
        if issue_code:
            row["issue_code"] = issue_code
        self.rows.append(row)

    def finish(self):
        issues = [r for r in self.rows if r["outcome"] == "needs_review"]
        result = {"run_id": self.run_id, "completed_at": datetime.now().isoformat(timespec="seconds"), "mode": self.mode, "observed_sources": self.observed_sources, "scan_complete": self.scan_complete, "created": self.created, "updated": self.updated, "review": len(issues), "output": str(self.output), "rows": self.rows}
        header = f"Mode: {self.mode}\nCompleted: {result['completed_at']}\nBoth plant files checked: {'Yes' if self.scan_complete else 'No - some information could not be checked'}\nCreated: {self.created}\nExcel updates recovered: {self.updated}\nNeeds review: {len(issues)}\n\n"
        body = "\n".join(f"{r['plant']} | SN {r['sn']} | row {r['row']} | {r['supplier']} | {r['outcome']}: {r['detail']}" for r in self.rows)
        json_paths = [self.output / "Results.json"]
        text_paths = [self.output / "Results.txt"]
        if self.publish_latest:
            json_paths.append(self.root / "Output" / "Latest Results.json")
            text_paths.append(self.root / "Output" / "Latest Results.txt")
        for path in json_paths:
            save_json(path, result)
        for path in text_paths:
            path.write_text(header + body + "\n", encoding="utf-8")
        return self.created, len(issues), self.output


def _digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, default=str).encode()).hexdigest()


def _alert_state(path, today):
    if not path.exists():
        return {}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(state, dict):
            raise ValueError("invalid alert state")
        for key in ("seen", "processed_runs"):
            if not isinstance(state.get(key, []), list) or not all(isinstance(v, str) for v in state.get(key, [])):
                raise ValueError(f"invalid {key}")
        for key in ("pending_created", "pending_updated"):
            if type(state.get(key, 0)) is not int or state.get(key, 0) < 0:
                raise ValueError(f"invalid {key}")
        if not isinstance(state.get("active", {}), dict) or not all(isinstance(v, str) for v in state.get("active", {}).values()):
            raise ValueError("invalid active issues")
        if not isinstance(state.get("last_day", ""), str):
            raise ValueError("invalid last_day")
        if state.get("last_day"):
            date.fromisoformat(state["last_day"])
        return state
    except (ValueError, TypeError) as exc:
        # Keep damaged state for diagnosis; never touch the separate COD register.
        backup = path.with_name(f"alert-state.corrupt-{uuid.uuid4().hex}.json")
        backup.write_bytes(path.read_bytes())
        return {"last_day": today, "state_warning": f"Alert history was damaged and backed up to {backup.name}: {exc}. Notifications resume on a later day; inspect the results now."}


def _record_observation(state, result):
    if not isinstance(result, dict) or not isinstance(result.get("rows", []), list) or any(not isinstance(row, dict) for row in result.get("rows", [])):
        raise ValueError("Invalid queued notification report")
    for key in ("created", "updated"):
        if type(result.get(key, 0)) is not int or result.get(key, 0) < 0:
            raise ValueError(f"Invalid {key} count in queued notification report")
    processed = set(state.get("processed_runs", []))
    run_id = str(result.get("run_id") or result.get("output") or _digest(result))
    if run_id in processed:
        return
    incoming = {}
    for row in result.get("rows", []):
        if row.get("outcome") != "needs_review":
            continue
        identity = {k: row.get(k, "") for k in ("plant", "file", "sn", "sheet")}
        identity["issue"] = row.get("issue_code") or row.get("detail", "")
        incoming[_digest(identity)] = row.get("plant", "Setup")
    active = state.get("active", {})
    full_scan = result.get("scan_complete", True)  # legacy reports had no flag
    observed = set(result.get("observed_sources", []))
    if full_scan:
        active = incoming
    else:
        # A failed/partial scan cannot declare unseen plant issues resolved.
        active = {key: source for key, source in active.items() if source not in observed}
        active.update(incoming)
    processed.add(run_id)
    state.update(active=active, scan_complete=full_scan, seen=sorted(set(state.get("seen", [])) & set(active)), processed_runs=sorted(processed), pending_created=state.get("pending_created", 0) + result.get("created", 0), pending_updated=state.get("pending_updated", 0) + result.get("updated", 0))


def notify_once(root, result, notifier, today=None, enabled=True, repeat_pending=False):
    """Persist pending work, deduplicate runs, and submit at most one summary/day.

    A True return means Windows accepted the notification request, not that a user
    saw it (Focus Assist can hide it). This function never opens a modal dialog.
    """
    if result.get("mode") == "PDF preview":
        return False
    for key in ("created", "updated"):
        value = result.get(key, 0)
        if type(value) is not int or value < 0:
            raise ValueError(f"Invalid {key} count in the notification report")
    today = (today or date.today()).isoformat()
    path = Path(root) / "Data" / "alert-state.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    # A second run can finish while Windows displays the first run's balloon.
    # Queue its observations before taking the lock so those events are not lost.
    run_id = str(result.get("run_id") or result.get("output") or _digest(result))
    inbox = path.parent / "Alert inbox"
    save_json(inbox / (_digest(run_id) + ".json"), result)
    lock_path = path.with_suffix(".lock")
    try:
        with lock_path.open("xb"):
            pass
    except FileExistsError:
        pass
    context = exclusive_file(lock_path)
    try:
        context.__enter__()
    except LockBusyError:
        return False  # Another process owns the notification pass; no popup storm.
    try:
        state = _alert_state(path, today)
        queued = sorted(inbox.glob("*.json"), key=lambda p: (p.stat().st_mtime_ns, p.name))
        queue_errors = []
        for item in queued:
            try:
                _record_observation(state, json.loads(item.read_text(encoding="utf-8")))
            except (ValueError, TypeError, KeyError) as exc:
                os.replace(item, item.with_suffix(".invalid"))
                queue_errors.append(str(exc))
        if queue_errors:
            state.setdefault("active", {})[_digest("damaged_alert_queue")] = "Setup"
            state["state_warning"] = "A queued alert report was damaged and retained as .invalid in Data/Alert inbox. Its counts are unknown. Review the saved Results reports. " + "; ".join(queue_errors)
        active = state.get("active", {})
        full_scan = state.get("scan_complete", True)
        current = set(active)
        seen = set(state.get("seen", [])) & current
        pending_created = state.get("pending_created", 0)
        pending_updated = state.get("pending_updated", 0)
        if not enabled:
            # Muting does not queue a historical flood for when summaries resume.
            seen, pending_created, pending_updated = current, 0, 0
        state.update(seen=sorted(seen), pending_created=pending_created, pending_updated=pending_updated)
        save_json(path, state)
        for item in queued:
            item.unlink(missing_ok=True)  # Already durably recorded, not user data.
        if state.get("state_warning"):
            (path.parent / "Notification status.txt").write_text(state["state_warning"], encoding="utf-8")
        actionable = current if repeat_pending else current - seen
        if not enabled or not (actionable or pending_created or pending_updated) or state.get("last_day", "") >= today:
            return False
        parts = []
        if pending_created:
            parts.append(f"{pending_created} new draft(s)")
        if pending_updated:
            parts.append(f"{pending_updated} Excel update(s) recovered")
        if current:
            parts.append(f"{len(current)} issue(s) need attention")
        if not full_scan:
            parts.append("some files could not be checked")
        message = "; ".join(parts) + ". Open Certificate Automation: Needs attention. Review blank statuses and choose Ready or N/A."
        previous_day = state.get("last_day", "")
        state["last_day"] = today
        # If interrupted after OS submission, suppress another attempt today.
        save_json(path, state)
        try:
            submitted = bool(notifier("Ecoreco certificates", message))
        except Exception as exc:
            state["last_delivery_error"] = f"Notification attempt failed: {type(exc).__name__}: {exc}"
            submitted = False
        else:
            if submitted:
                state.update(seen=sorted(current), pending_created=0, pending_updated=0)
                state.pop("last_delivery_error", None)
            else:
                # Explicit rejection is safe to retry; unknown exceptions are not.
                state["last_day"] = previous_day
                state["last_delivery_error"] = "Windows did not accept the notification. Pending information is retained."
        save_json(path, state)
        warning = state.get("last_delivery_error") or state.get("state_warning")
        if warning:
            (path.parent / "Notification status.txt").write_text(warning, encoding="utf-8")
        return submitted
    finally:
        context.__exit__(None, None, None)


def windows_notification(title, message):
    """One silent Windows notification. No message boxes or console windows."""
    if os.name != "nt":
        return False
    import win32api
    import win32con
    import win32gui
    class_name = "EcorecoCertificateNotice" + uuid.uuid4().hex
    wc = win32gui.WNDCLASS()
    wc.hInstance = win32api.GetModuleHandle(None)
    wc.lpszClassName = class_name
    wc.lpfnWndProc = {win32con.WM_DESTROY: lambda *_: None}
    atom = win32gui.RegisterClass(wc)
    hwnd = None
    added = False
    try:
        hwnd = win32gui.CreateWindow(atom, title, 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None)
        icon = win32gui.LoadIcon(0, win32con.IDI_INFORMATION)
        flags = win32gui.NIF_ICON | win32gui.NIF_TIP | win32gui.NIF_INFO
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, (hwnd, 0, flags, 0, icon, title, message, 8000, title, 0x1 | 0x10))
        added = True
        deadline = time.monotonic() + 8
        while time.monotonic() < deadline:
            win32gui.PumpWaitingMessages()
            time.sleep(0.1)
        return True
    finally:
        if hwnd is not None:
            if added:
                with suppress(Exception):
                    win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (hwnd, 0))
            with suppress(Exception):
                win32gui.DestroyWindow(hwnd)
        with suppress(Exception):
            win32gui.UnregisterClass(class_name, wc.hInstance)
