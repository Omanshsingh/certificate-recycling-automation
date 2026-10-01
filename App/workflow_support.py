"""Settings, safe run locking, reports and focused user reminders."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox
import uuid

from app_paths import resource
from certificate_core import certificate_reminder_due, parse_date, status_value
from workbook_store import exclusive_file


def load_settings(root):
    root = Path(root)
    defaults_path = resource(root, "settings.example.json")
    local_path = root / "settings.local.json"

    def read(path):
        if not path.exists():
            return {}
        value = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(value, dict):
            raise ValueError(f"{path.name} must contain a settings object.")
        return value

    merged = read(defaults_path)
    merged.update(read(local_path))
    allowed = {"workbook_paths", "certificate_number_prefix", "certificate_number_start",
               "certificate_number_width", "reminder_time", "reminders_enabled", "backup_destination"}
    settings = {key: value for key, value in merged.items() if key in allowed}
    paths = settings.get("workbook_paths", ["", ""])
    if not isinstance(paths, list) or len(paths) > 2 or any(not isinstance(item, str) for item in paths):
        raise ValueError("Settings need the two plant Excel file paths.")
    settings["workbook_paths"] = (paths + ["", ""])[:2]
    for key, default, low, high in (("certificate_number_start", 1, 1, 999999999999),
                                    ("certificate_number_width", 4, 1, 18)):
        value = settings.get(key, default)
        if type(value) is not int or not low <= value <= high:
            raise ValueError(f"{key} must be a whole number from {low} to {high}.")
        settings[key] = value
    prefix = settings.get("certificate_number_prefix", "")
    if not isinstance(prefix, str) or any(ord(ch) < 32 for ch in prefix):
        raise ValueError("The COD prefix must be plain text.")
    settings["certificate_number_prefix"] = prefix
    reminder_time = settings.get("reminder_time", "11:00")
    if not isinstance(reminder_time, str) or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", reminder_time):
        raise ValueError("Choose a valid reminder time, such as 11:00 AM.")
    settings["reminder_time"] = reminder_time
    enabled = settings.get("reminders_enabled", True)
    if not isinstance(enabled, bool):
        raise ValueError("reminders_enabled must be true or false.")
    settings["reminders_enabled"] = enabled
    destination = settings.get('backup_destination', '')
    if not isinstance(destination, str) or len(destination) > 32000:
        raise ValueError('Choose a valid backup folder.')
    settings['backup_destination'] = destination
    return settings


def save_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temp.open('w', encoding='utf-8', newline='') as stream:
            stream.write(json.dumps(value, indent=2, ensure_ascii=False, default=str))
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


@contextmanager
def run_lock(root):
    lock = Path(root) / "Data" / "run.lock"
    lock.parent.mkdir(parents=True, exist_ok=True)
    try:
        with lock.open("xb"):
            pass
    except FileExistsError:
        pass
    with exclusive_file(lock):
        yield


class Report:
    def __init__(self, root, as_of, mode, publish_latest=True):
        self.root = Path(root)
        self.as_of = as_of
        self.mode = mode
        self.publish_latest = publish_latest
        self.output = self.root / "Data" / "Reports"
        self.output.mkdir(parents=True, exist_ok=True)
        self.run_id = uuid.uuid4().hex
        self.observed_sources = []
        self.scan_complete = False
        self.rows = []
        self.created = 0

    def add(self, batch, outcome, detail, issue_code=None, system_alert=False, warnings=None, **_legacy):
        if batch.get("_missing_pdf") and outcome != "generated_now":
            cod = str(batch.get("cod_no") or "").strip()
            recovery = (f"Keep COD {cod} when recreating it." if cod else
                        "A new COD will be assigned when you create it.")
            detail = f"PDF missing from Certificates: Excel was changed to Ready. {recovery} {detail}"
            system_alert = True
        raw_date = batch.get("inward_date")
        inward_date = parse_date(raw_date)
        if raw_date is None or (isinstance(raw_date, str) and not raw_date.strip()):
            quality = "missing"
        elif inward_date is None:
            quality = "invalid"
        elif inward_date > self.as_of:
            quality = "future"
        else:
            quality = "valid"
        age = (self.as_of - inward_date).days if inward_date else None
        state = batch.get("_effective_certificate_status", batch.get("issued", ""))
        reminder = dict(batch, inward_date=inward_date, _effective_certificate_status=state)
        due = certificate_reminder_due(reminder, self.as_of, days=25)
        row = {
            "source_kind": "workbook",
            "plant": batch.get("source_id", "Setup"),
            "file": batch.get("source_path", ""),
            "sheet": batch.get("source_sheet", ""),
            "row": batch.get("source_row", ""),
            "sn": str(batch.get("case_id", "") or ""),
            "supplier": str(batch.get("supplier", "") or ""),
            "location": str(batch.get("location", "") or ""),
            "form6": str(batch.get("form6", "") or ""),
            "reference": str(batch.get("reference", "") or ""),
            "certificate_weight": batch.get("certificate_weight"),
            "certificate_type": str(batch.get("certificate_type") or "").strip() or "standard",
            "certificate_fields": dict(batch.get("certificate_fields") or {}),
            "display_weight": batch.get("display_weight", batch.get("total_weight")),
            "weight_source": batch.get("weight_source", "inward"),
            "items": list(batch.get("items", [])),
            "outcome": outcome,
            "detail": detail,
            "certificate_no": str(batch.get("cod_no", "") or ""),
            "certificate_date": parse_date(batch.get("cod_date")).isoformat() if parse_date(batch.get("cod_date")) else "",
            "certificate_status": status_value(state),
            'pdf_filename': batch.get('pdf_filename', ''),
            "inward_date": inward_date.isoformat() if inward_date else "",
            "inward_date_quality": quality,
            "age_days": age,
            "reminder_due": due,
            "overdue": bool(due and age is not None and age > 30),
            "warnings": list(warnings or []),
        }
        if issue_code:
            row["issue_code"] = issue_code
        if system_alert or (outcome == "needs_review" and quality in {"missing", "invalid", "future"}):
            row["system_alert"] = True
        self.rows.append(row)

    def finish(self):
        needs_review = sum(row["outcome"] == "needs_review" for row in self.rows)
        result = {
            "run_id": self.run_id,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "as_of": self.as_of.isoformat(),
            "mode": self.mode,
            "observed_sources": self.observed_sources,
            "scan_complete": self.scan_complete,
            "created": self.created,
            "review": needs_review,
            "output": str(self.output),
            "certificate_folder": str(self.root / "Certificates"),
            "rows": self.rows,
        }
        if hasattr(self, 'certificate_inventory'):
            result['certificate_inventory'] = self.certificate_inventory
        if hasattr(self, 'result_sink') and self.result_sink is not None:
            self.result_sink.clear()
            self.result_sink.update(result)
        lines = [
            f"Mode: {self.mode}",
            f"Completed: {result['completed_at']}",
            f"Both plant files checked: {'Yes' if self.scan_complete else 'No'}",
            f"PDF certificates generated: {self.created}",
            f"Rows needing attention: {needs_review}",
            "",
        ]
        lines.extend(
            f"{row['plant']} | SN {row['sn']} | {row['supplier']} | {row['outcome']}: {row['detail']}"
            for row in self.rows
        )
        json_path = self.output / "Latest Results.json"
        text_path = self.output / "Latest Results.txt"
        save_json(json_path, result)
        text_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return self.created, needs_review, self.output


def _actionable_rows(result):
    rows = result.get("rows", []) if isinstance(result, dict) else []
    return [row for row in rows if row.get("reminder_due") or row.get("system_alert")]


def open_alert_workbook(result):
    """Open the plant workbook for the first client listed in the reminder."""
    row = next((item for item in _actionable_rows(result) if item.get("file")), None)
    if row is None:
        raise FileNotFoundError("No plant workbook is linked to the certificates in this reminder.")
    workbook = Path(row["file"])
    if not workbook.is_file():
        raise FileNotFoundError(f"The plant workbook could not be found:\n{workbook}")
    opener = getattr(os, "startfile", None)
    if opener is None:
        raise OSError("Opening the plant workbook from the reminder is available on Windows.")
    opener(str(workbook))
    return workbook


def _notice_text(rows):
    due = [row for row in rows if row.get("reminder_due")]
    urgent = [row for row in rows if row.get("system_alert") and not row.get("reminder_due")]
    parts = []
    if due:
        soon = sum(not row.get("overdue") for row in due)
        overdue = len(due) - soon
        if soon:
            parts.append(f"{soon} client certificate(s) are due for follow-up (inward day 25–30).")
        if overdue:
            parts.append(f"{overdue} client certificate(s) are overdue beyond day 30.")
    if urgent:
        parts.append(f"{len(urgent)} workbook or certificate detail issue(s) need attention.")
    examples = []
    for row in rows[:5]:
        client = row.get("supplier") or "Client not named"
        serial = row.get("sn")
        age = row.get("age_days")
        age_text = f" — day {age}" if isinstance(age, int) else ""
        examples.append(f"• {client} (ID {serial}){age_text}" if serial else f"• {client}")
    return "\n".join(parts + (["", *examples] if examples else []))


def notify_once(root, result, notifier, today=None, enabled=True, repeat_pending=False,
                reminder_time="11:00", record_only=False, reminder_only=False):
    """Show one centered daily summary for actionable reminders, not every row."""
    today = today or date.today()
    rows = _actionable_rows(result)
    if not rows:
        return False
    state_path = Path(root) / "Data" / "reminder-state.json"
    state = {}
    try:
        loaded = json.loads(state_path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            state = loaded
    except (OSError, ValueError):
        pass
    signature = ",".join(sorted(f"{row.get('plant')}|{row.get('sn')}|{row.get('certificate_status')}|{row.get('age_days')}" for row in rows))
    if not repeat_pending and state.get("last_day") == today.isoformat() and state.get("signature") == signature:
        return False
    if enabled and not record_only:
        notifier("Certificates needing attention", _notice_text(rows))
        # Record the alert only after it was successfully shown. If the desktop
        # notification fails, a later scheduled/sign-in check must retry it.
        state.update(last_day=today.isoformat(), signature=signature)
        save_json(state_path, state)
        return True
    return False


def _powershell(root, mode, extra=()):
    script = Path(root) / "App" / "Install Daily Task.ps1"
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-Mode", mode]
    for key, value in extra:
        command.extend([key, str(value)])
    return subprocess.run(command, capture_output=True, text=True, timeout=30, check=False)


def schedule_one_hour_reminder(root, now=None):
    run_at = (now or datetime.now()) + timedelta(hours=1)
    completed = _powershell(root, "Snooze", (("-SnoozeAt", run_at.strftime("%Y-%m-%dT%H:%M:%S")),))
    if completed.returncode:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Windows could not schedule the reminder.")


def cancel_snoozed_reminder(root, force=False):
    completed = _powershell(root, "CancelSnooze")
    if completed.returncode and not force:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Windows could not clear the reminder.")


def launch_review_app(root):
    script = Path(root) / "App" / "desktop_app.py"
    subprocess.Popen([sys.executable, str(script)], cwd=str(script.parent), close_fds=True)


def show_centered_reminder(title, message, parent=None, *, on_snooze=None,
                           on_dismiss=None, on_open_workbook=None,
                           on_open_automation=None):
    own_root = parent is None
    window = tk.Tk() if own_root else tk.Toplevel(parent)
    window.title(title)
    window.geometry("720x360")
    window.resizable(False, False)
    window.attributes("-topmost", True)
    window.transient(parent) if parent else None
    window.grab_set()
    outer = tk.Frame(window, padx=28, pady=24)
    outer.pack(fill="both", expand=True)
    tk.Label(outer, text=title, font=("Segoe UI", 18, "bold"), anchor="w").pack(fill="x", pady=(0, 12))
    tk.Label(outer, text=message, font=("Segoe UI", 11), justify="left", anchor="nw", wraplength=660).pack(fill="both", expand=True)
    actions = tk.Frame(outer)
    actions.pack(fill="x", pady=(18, 0))
    result = {"action": "dismiss"}

    def finish(action, callback):
        try:
            if callback:
                callback()
        except Exception as exc:
            error_titles = {
                "open_workbook": "Could not open Excel workbook",
                "open_automation": "Could not open automation",
                "snooze": "Could not schedule reminder",
                "dismiss": "Could not dismiss reminder",
            }
            messagebox.showerror(error_titles.get(action, "Action failed"), str(exc), parent=window)
            return
        result["action"] = action
        window.destroy()

    tk.Button(actions, text="Open Excel workbook", command=lambda: finish("open_workbook", on_open_workbook), padx=12).pack(side="left")
    tk.Button(actions, text="Open automation", command=lambda: finish("open_automation", on_open_automation), padx=12).pack(side="left", padx=8)
    tk.Button(actions, text="Remind me in 1 hour", command=lambda: finish("snooze", on_snooze), padx=12).pack(side="left")
    tk.Button(actions, text="Dismiss", command=lambda: finish("dismiss", on_dismiss), padx=12).pack(side="right")
    window.protocol("WM_DELETE_WINDOW", lambda: finish("dismiss", on_dismiss))
    window.bind("<Escape>", lambda _: finish("dismiss", on_dismiss))
    window.update_idletasks()
    x = max(0, (window.winfo_screenwidth() - window.winfo_width()) // 2)
    y = max(0, (window.winfo_screenheight() - window.winfo_height()) // 2)
    window.geometry(f"+{x}+{y}")
    window.focus_force()
    if own_root:
        window.mainloop()
    else:
        parent.wait_window(window)
    return result["action"]
