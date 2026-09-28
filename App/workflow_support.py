"""Saved configuration, one running batch, reports and quiet notifications."""
from __future__ import annotations

from contextlib import contextmanager, suppress
from datetime import date, datetime, timedelta
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
import tkinter as tk
import uuid

from workbook_store import LockBusyError, exclusive_file
from app_paths import resource
from certificate_core import certificate_reminder_due, parse_date, status_value


def load_settings(root):
    root = Path(root)
    defaults = resource(root, "settings.example.json")
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
    reminder_time = settings.get("reminder_time", "09:00")
    if not isinstance(reminder_time, str) or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", reminder_time):
        raise ValueError("reminder_time in settings.local.json must be a 24-hour time such as 09:00.")
    settings["reminder_time"] = reminder_time
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
        self.as_of = as_of
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

    def add(self, batch, outcome, detail, issue_code=None, awaiting_send=False,
            system_alert=False):
        raw_inward_date = batch.get("inward_date")
        inward_date = parse_date(raw_inward_date)
        if raw_inward_date is None or (isinstance(raw_inward_date, str) and not raw_inward_date.strip()):
            inward_date_quality = "missing"
        elif inward_date is None:
            inward_date_quality = "invalid"
        elif inward_date > self.as_of:
            inward_date_quality = "future"
        else:
            inward_date_quality = "valid"
        age_days = (self.as_of - inward_date).days if inward_date else None
        effective_status = batch.get("_effective_certificate_status")
        reminder_batch = dict(batch, inward_date=inward_date)
        if effective_status is not None:
            reminder_batch["_effective_certificate_status"] = effective_status
        row = {"plant": batch.get("source_id", "Setup"), "file": batch.get("source_path", ""),
               "sheet": batch.get("source_sheet", ""), "row": batch.get("source_row", ""),
               "sn": batch.get("case_id", ""), "supplier": batch.get("supplier", ""),
               "outcome": outcome, "detail": detail,
               "certificate_no": str(batch.get("cod_no", "") or ""),
               "certificate_date": (parse_date(batch.get("cod_date")).isoformat()
                                    if parse_date(batch.get("cod_date")) else ""),
               "certificate_status": (status_value(effective_status) if effective_status is not None
                                      else status_value(batch.get("issued"))),
               "inward_date": inward_date.isoformat() if inward_date else "",
               "inward_date_quality": inward_date_quality,
               "age_days": age_days,
               "reminder_due": certificate_reminder_due(reminder_batch, self.as_of)}
        if issue_code:
            row["issue_code"] = issue_code
        if awaiting_send:
            row["awaiting_send"] = True
        if (system_alert or issue_code in {"source_unavailable", "run_failure", "setup_issue",
                                          "draft_recovery", "history_reconcile", "history_source_changed",
                                          "sent_without_certificate", "sent_history_mismatch",
                                          "outlook_save_uncertain"}
                or (outcome == "needs_review" and not row["sn"])):
            row["system_alert"] = True
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


def _reported_date_issue(row, result):
    """Return a safe date-quality category for legacy and current report rows."""
    if status_value(row.get("certificate_status")) in {"na", "completed", "sent"}:
        return None
    category = row.get("inward_date_quality")
    if category in {"missing", "invalid", "future"}:
        return category

    raw = row.get("inward_date")
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return "missing"
    parsed = parse_date(raw)
    if parsed is None:
        return "invalid"
    age_days = row.get("age_days")
    if type(age_days) is int:
        return "future" if age_days < 0 else None
    reference = result.get("as_of") or result.get("completed_at")
    try:
        reference_date = parse_date(str(reference)[:10]) if reference else None
    except (TypeError, ValueError):
        reference_date = None
    return "future" if reference_date and parsed > reference_date else None


def _alert_state(path, today):
    if not path.exists():
        return {"schema_version": 3, "active": {}, "system_active": {},
                "date_quality_active": {}, "seen": [], "system_seen": [],
                "date_quality_seen": [], "processed_runs": [], "last_day": "",
                "dismissed_until": "", "last_attempt_day": "",
                "last_attempt_pending": False, "notified_runs": [],
                "snooze_pending": False}
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(state, dict):
            raise ValueError("invalid alert state")
        version = state.get("schema_version", 0)
        if type(version) is not int or version < 0:
            raise ValueError("invalid schema_version")
        for key in ("seen", "system_seen", "date_quality_seen", "processed_runs"):
            if not isinstance(state.get(key, []), list) or not all(isinstance(v, str) for v in state.get(key, [])):
                raise ValueError(f"invalid {key}")
        if not isinstance(state.get("notified_runs", []), list) or not all(isinstance(v, str) for v in state.get("notified_runs", [])):
            raise ValueError("invalid notified_runs")
        for key in ("pending_created", "pending_updated"):
            if type(state.get(key, 0)) is not int or state.get(key, 0) < 0:
                raise ValueError(f"invalid {key}")
        for key in ("active", "system_active", "date_quality_active", "pending_drafts"):
            if not isinstance(state.get(key, {}), dict) or not all(isinstance(v, str) for v in state.get(key, {}).values()):
                raise ValueError(f"invalid {key}")
        for key in ("last_day", "dismissed_until", "last_attempt_day"):
            if not isinstance(state.get(key, ""), str):
                raise ValueError(f"invalid {key}")
            if state.get(key):
                date.fromisoformat(state[key])
        for key in ("snooze_pending", "last_attempt_pending"):
            if not isinstance(state.get(key, False), bool):
                raise ValueError(f"invalid {key}")
            state.setdefault(key, False)
        if version < 2:
            # Old versions alerted on every blank row and unsent draft. Rebuild
            # those lists from a fresh scan instead of replaying stale alerts.
            state = {
                "schema_version": 3,
                "last_day": state.get("last_day", ""),
                "dismissed_until": state.get("last_day", "") if state.get("last_action") == "dismiss" else "",
                "last_attempt_day": "",
                "last_attempt_pending": False,
                "notified_runs": [],
                "snooze_pending": state.get("snooze_pending", False),
                "processed_runs": state.get("processed_runs", []),
                "active": {},
                "system_active": {},
                "date_quality_active": {},
                "seen": [],
                "system_seen": [],
                "date_quality_seen": [],
            }
        else:
            state.setdefault("active", {})
            state.setdefault("system_active", {})
            state.setdefault("date_quality_active", {})
            state.setdefault("seen", [])
            state.setdefault("system_seen", [])
            state.setdefault("date_quality_seen", [])
            # Preserve an explicit old-version "remind tomorrow" choice, but
            # do not treat a routine run's legacy last_day as an alert limit.
            state.setdefault("dismissed_until", state.get("last_day", "") if state.get("last_action") == "dismiss" else "")
            state.setdefault("last_attempt_day", "")
            state.setdefault("last_attempt_pending", False)
            state.setdefault("notified_runs", [])
            state["schema_version"] = 3
        return state
    except (ValueError, TypeError) as exc:
        # Keep damaged state for diagnosis; never touch the separate COD register.
        backup = path.with_name(f"alert-state.corrupt-{uuid.uuid4().hex}.json")
        backup.write_bytes(path.read_bytes())
        return {"schema_version": 3, "last_day": "", "dismissed_until": "",
                "last_attempt_day": "", "last_attempt_pending": False,
                "notified_runs": [], "snooze_pending": False,
                "active": {}, "system_active": {}, "date_quality_active": {},
                "seen": [], "system_seen": [], "date_quality_seen": [],
                "processed_runs": [],
                "state_warning": f"Alert history was damaged and backed up to {backup.name}: {exc}. Notifications resume on a later day; inspect the results now."}


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
    incoming_system = {}
    incoming_date_quality = {}
    for row in result.get("rows", []):
        if (row.get("system_alert") is True
                or row.get("issue_code") in {"source_unavailable", "run_failure", "setup_issue",
                                               "draft_recovery", "history_reconcile", "history_source_changed",
                                               "sent_without_certificate", "sent_history_mismatch",
                                               "outlook_save_uncertain"}
                or (row.get("outcome") == "needs_review" and not row.get("sn"))):
            identity = {k: row.get(k, "") for k in ("plant", "file", "sn", "sheet")}
            identity["issue"] = row.get("issue_code") or row.get("detail", "")
            incoming_system[_digest(identity)] = row.get("plant", "Setup")
        elif row.get("reminder_due") is True:
            identity = {k: row.get(k, "") for k in ("plant", "file", "sn", "sheet", "row")}
            identity["kind"] = "inward_over_30_days_not_sent"
            incoming[_digest(identity)] = row.get("plant", "Setup")
        if row.get("sn") and _reported_date_issue(row, result):
            identity = {k: row.get(k, "") for k in ("plant", "file", "sn", "sheet", "row")}
            identity["kind"] = "inward_date_unclassifiable"
            incoming_date_quality[_digest(identity)] = row.get("plant", "Setup")
    active = state.get("active", {})
    system_active = state.get("system_active", {})
    date_quality_active = state.get("date_quality_active", {})
    full_scan = result.get("scan_complete", True)  # legacy reports had no flag
    observed = set(result.get("observed_sources", []))
    if full_scan:
        active = incoming
        system_active = incoming_system
        date_quality_active = incoming_date_quality
    else:
        # A failed/partial scan cannot clear unseen overdue cases or failures.
        active = {key: source for key, source in active.items() if source not in observed}
        active.update(incoming)
        system_active = {key: source for key, source in system_active.items() if source not in observed}
        system_active.update(incoming_system)
        date_quality_active = {key: source for key, source in date_quality_active.items()
                               if source not in observed}
        date_quality_active.update(incoming_date_quality)
        system_active[_digest("partial_scan")] = "Setup"
    processed.add(run_id)
    state.update(schema_version=3, active=active, system_active=system_active,
                 date_quality_active=date_quality_active,
                 scan_complete=full_scan,
                 seen=sorted(set(state.get("seen", [])) & set(active)),
                 system_seen=sorted(set(state.get("system_seen", [])) & set(system_active)),
                 date_quality_seen=sorted(set(state.get("date_quality_seen", []))
                                          & set(date_quality_active)),
                 processed_runs=sorted(processed))


def rearm_alerts_after_schedule_change(root):
    """Re-arm reminders when the user saves a daily reminder time."""
    path = Path(root) / "Data" / "alert-state.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = path.with_suffix(".lock")
    try:
        with lock_path.open("xb"):
            pass
    except FileExistsError:
        pass
    context = exclusive_file(lock_path)
    try:
        context.__enter__()
    except LockBusyError as exc:
        raise RuntimeError("The reminder status is updating. Save the reminder time again in a moment.") from exc
    try:
        state = _alert_state(path, date.today().isoformat())
        state["dismissed_until"] = ""
        state["last_attempt_pending"] = False
        state.pop("last_delivery_error", None)
        save_json(path, state)
    finally:
        context.__exit__(None, None, None)


def notify_once(root, result, notifier, today=None, enabled=True, repeat_pending=False,
                reminder_time=None, current_time=None, reminder_only=False, record_only=False):
    """Persist pending work and show a summary for each eligible reminder run.

    ``result=None, reminder_only=True`` replays the last saved summary for a
    user-requested snooze. It never scans Excel, connects to Outlook, or creates
    drafts. ``record_only=True`` refreshes persisted observations without
    showing a prompt. Interactive notifiers may return ``review``, ``dismiss``
    or ``snooze``; legacy Windows-notification callbacks may still return bools.
    """
    if not reminder_only and (not isinstance(result, dict) or result.get("mode") == "PDF preview"):
        return False
    if not reminder_only:
        for key in ("created", "updated"):
            value = result.get(key, 0)
            if type(value) is not int or value < 0:
                raise ValueError(f"Invalid {key} count in the notification report")
    today = (today or date.today()).isoformat()
    path = Path(root) / "Data" / "alert-state.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    # A second run can finish while Windows displays the first run's balloon.
    # Queue its observations before taking the lock so those events are not lost.
    inbox = path.parent / "Alert inbox"
    run_id = ""
    if not reminder_only:
        run_id = str(result.get("run_id") or result.get("output") or _digest(result))
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
        system_active = state.get("system_active", {})
        date_quality_active = state.get("date_quality_active", {})
        full_scan = state.get("scan_complete", True)
        notification_key = f"{today}:{run_id}" if not reminder_only else ""
        already_notified = not reminder_only and notification_key in state.get("notified_runs", [])
        current = set(active)
        current_system = set(system_active)
        current_date_quality = set(date_quality_active)
        seen = set(state.get("seen", [])) & current
        system_seen = set(state.get("system_seen", [])) & current_system
        date_quality_seen = set(state.get("date_quality_seen", [])) & current_date_quality
        if state.get("state_warning"):
            current_system.add(_digest("alert_state_warning"))
            system_active[_digest("alert_state_warning")] = "Setup"
        if not enabled:
            # Muting does not queue a historical flood for when reminders resume.
            seen, system_seen, date_quality_seen = current, current_system, current_date_quality
        state.update(seen=sorted(seen), system_seen=sorted(system_seen),
                     date_quality_seen=sorted(date_quality_seen),
                     active=active, system_active=system_active,
                     date_quality_active=date_quality_active)
        save_json(path, state)
        for item in queued:
            item.unlink(missing_ok=True)  # Already durably recorded, not user data.
        if state.get("state_warning"):
            (path.parent / "Notification status.txt").write_text(state["state_warning"], encoding="utf-8")
        if record_only:
            return False
        actionable = current if repeat_pending else current - seen
        actionable_system = current_system if repeat_pending else current_system - system_seen
        actionable_date_quality = (current_date_quality if repeat_pending
                                   else current_date_quality - date_quality_seen)
        has_pending = bool(actionable or actionable_system or actionable_date_quality)
        if reminder_only and not has_pending:
            state["snooze_pending"] = False
            save_json(path, state)
            return False
        uncertain_same_day_attempt = (
            state.get("last_attempt_pending") and state.get("last_attempt_day") == today
        )
        dismissed_for_today = state.get("dismissed_until", "") >= today
        if (not enabled or not has_pending
                or (reminder_only and not state.get("snooze_pending"))
                or (not reminder_only and (state.get("snooze_pending") or dismissed_for_today))
                or already_notified
                or uncertain_same_day_attempt):
            return False
        if not reminder_only and reminder_time is not None:
            if not isinstance(reminder_time, str) or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", reminder_time):
                raise ValueError("reminder_time must be a 24-hour time such as 09:00")
            current_time = current_time or datetime.now().strftime("%H:%M")
            if not isinstance(current_time, str) or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", current_time):
                raise ValueError("current_time must be a 24-hour time such as 09:00")
            if current_time < reminder_time:
                return False
        parts = []
        if actionable:
            count = len(actionable)
            noun = "inward record" if count == 1 else "inward records"
            verb = "is" if count == 1 else "are"
            parts.append(f"{count} {noun} {verb} 30+ days old and not marked Completed or N/A")
        if actionable_date_quality:
            count = len(actionable_date_quality)
            noun = "inward record" if count == 1 else "inward records"
            verb = "has" if count == 1 else "have"
            date_issue = ("a missing, invalid, or future Inward Date" if count == 1
                          else "missing, invalid, or future Inward Dates")
            parts.append(f"{count} {noun} {verb} {date_issue} and cannot be age-classified")
        if actionable_system:
            count = len(actionable_system)
            noun = "automation problem" if count == 1 else "automation problems"
            verb = "needs" if count == 1 else "need"
            parts.append(f"{count} {noun} {verb} attention")
        if not full_scan:
            parts.append("some files could not be checked")
        message = "; ".join(parts) + ". Open Certificate Automation and review the listed items."
        previous_day = state.get("last_day", "")
        # Separate scheduled runs may remind more than once a day. Only an
        # interrupted, ambiguous popup attempt is held back to avoid duplicates.
        state["last_day"] = today
        state["last_attempt_day"] = today
        state["last_attempt_pending"] = True
        if not reminder_only:
            notified_runs = state.setdefault("notified_runs", [])
            notified_runs.append(notification_key)
            state["notified_runs"] = notified_runs[-512:]
        save_json(path, state)
        try:
            response = notifier("Ecoreco certificates", message)
        except Exception as exc:
            state["last_delivery_error"] = f"Notification attempt failed: {type(exc).__name__}: {exc}"
            save_json(path, state)
            (path.parent / "Notification status.txt").write_text(state["last_delivery_error"], encoding="utf-8")
            return False
        else:
            if isinstance(response, str) and response in {"review", "dismiss", "snooze"}:
                if state.get("state_warning"):
                    warning_id = _digest("alert_state_warning")
                    current_system.discard(warning_id)
                    system_active.pop(warning_id, None)
                    state.pop("state_warning", None)
                state["last_action"] = response
                state["snooze_pending"] = response == "snooze"
                state["last_attempt_pending"] = False
                if response == "dismiss":
                    state["dismissed_until"] = today
                elif response == "review":
                    state["dismissed_until"] = ""
                # A snooze must carry the same summary forward to its later prompt.
                if response != "snooze":
                    state.update(seen=sorted(current), system_seen=sorted(current_system),
                                 date_quality_seen=sorted(current_date_quality))
                state.pop("last_delivery_error", None)
                save_json(path, state)
                return True
            submitted = bool(response)
            state["last_attempt_pending"] = False
            if submitted:
                if state.get("state_warning"):
                    warning_id = _digest("alert_state_warning")
                    current_system.discard(warning_id)
                    system_active.pop(warning_id, None)
                    state.pop("state_warning", None)
                state["snooze_pending"] = False
                state.update(seen=sorted(current), system_seen=sorted(current_system),
                             date_quality_seen=sorted(current_date_quality))
                state.pop("last_delivery_error", None)
            else:
                # Explicit rejection is safe to retry; unknown exceptions are not.
                state["last_day"] = previous_day
                if not reminder_only:
                    state["notified_runs"] = [item for item in state.get("notified_runs", []) if item != notification_key]
                state["last_delivery_error"] = "Windows did not accept the notification. Pending information is retained."
        save_json(path, state)
        warning = state.get("last_delivery_error") or state.get("state_warning")
        if warning:
            (path.parent / "Notification status.txt").write_text(warning, encoding="utf-8")
        return submitted
    finally:
        context.__exit__(None, None, None)


SNOOZE_TASK_NAME = "Ecoreco Certificate Reminder Snooze"


def _powershell(root, mode, extra=()):
    if os.name != "nt":
        raise RuntimeError("One-hour reminders are available on Windows only.")
    script = resource(root, "Install Daily Task.ps1")
    if not script.is_file():
        raise RuntimeError("The reminder setup file is missing. Repair the app before snoozing.")
    executable = shutil.which("powershell.exe") or "powershell.exe"
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        result = subprocess.run(
            [executable, "-NoProfile", "-NonInteractive", "-File", str(script), "-Mode", mode, *map(str, extra)],
            cwd=str(script.parent), capture_output=True, text=True, timeout=45, creationflags=flags,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise RuntimeError(f"Windows could not update the reminder: {exc}") from exc
    if result.returncode:
        details = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part and part.strip())
        raise RuntimeError(details or "Windows Task Scheduler did not accept the reminder change.")
    return result.stdout.strip()


def schedule_one_hour_reminder(root, now=None):
    """Register a user-session one-shot reminder; never wakes the computer."""
    root = Path(root)
    now = now or datetime.now()
    run_at = (now + timedelta(hours=1)).replace(second=0, microsecond=0)
    if run_at <= now:
        run_at += timedelta(minutes=1)
    state_path = root / "Data" / "alert-state.json"
    state_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.exists() else {}
        if not isinstance(state, dict):
            raise ValueError("alert state is not an object")
        state["snooze_pending"] = True
        save_json(state_path, state)
    except (OSError, ValueError, TypeError) as exc:
        raise RuntimeError(f"Could not save the one-hour reminder. {exc}") from exc
    try:
        _powershell(root, "Snooze", ("-SnoozeAt", run_at.strftime("%Y-%m-%dT%H:%M:%S")))
    except Exception:
        state["snooze_pending"] = False
        save_json(state_path, state)
        raise
    return run_at


def cancel_snoozed_reminder(root, force=False):
    """Remove a pending one-shot snooze when the user chooses another action."""
    state_path = Path(root) / "Data" / "alert-state.json"
    state = {}
    if state_path.exists():
        try:
            state = json.loads(state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError, TypeError) as exc:
            raise RuntimeError(f"The saved reminder status could not be read. {exc}") from exc
    if not force and not (isinstance(state, dict) and state.get("snooze_pending")):
        return
    if isinstance(state, dict):
        state["snooze_pending"] = False
        save_json(state_path, state)
    # Invalidate the saved token first. Even if Windows refuses to delete the
    # one-shot task, its runner will now safely exit without showing a popup.
    _powershell(root, "CancelSnooze")


def launch_review_app(root):
    """Focus the app if it is open; otherwise start its normal desktop screen."""
    if os.name == "nt":
        try:
            import win32con
            import win32gui
            handle = win32gui.FindWindow(None, "Ecoreco Certificate Automation")
            if handle:
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(handle)
                return True
        except Exception:
            pass
    app_script = resource(root, "desktop_app.py")
    if not app_script.is_file():
        raise RuntimeError("The Certificate Automation screen could not be found.")
    executable = Path(sys.executable)
    if executable.name.casefold() == "python.exe":
        gui_executable = executable.with_name("pythonw.exe")
        if gui_executable.is_file():
            executable = gui_executable
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if executable.name.casefold() == "python.exe" else 0
    try:
        subprocess.Popen([str(executable), str(app_script), "--review"], cwd=str(app_script.parent), creationflags=flags)
    except OSError as exc:
        raise RuntimeError(f"Could not open Certificate Automation: {exc}") from exc
    return True


def show_centered_reminder(title, message, parent=None, *, test_mode=False,
                           on_snooze=None, on_dismiss=None, on_review=None):
    """Show a centered, actionable but non-blocking reminder window.

    The test version has identical layout and labels but deliberately runs none
    of the real actions, so it cannot schedule a task or alter alert history.
    """
    owned_root = parent is None
    host = parent or tk.Tk()
    if owned_root:
        host.withdraw()
    window = tk.Toplevel(host)
    window.title(title)
    window.resizable(False, False)
    if parent is not None:
        window.transient(host)
    window.configure(padx=24, pady=20)
    try:
        window.attributes("-topmost", True)
    except tk.TclError:
        pass

    ttk = __import__("tkinter.ttk", fromlist=["ttk"])
    ttk.Label(window, text="Certificate reminder", font=("Calibri", 17, "bold")).pack(anchor="w")
    ttk.Label(window, text=message, font=("Calibri", 11), wraplength=500, justify="left").pack(anchor="w", pady=(12, 10))
    note = "TEST ONLY — choices will not open the app or change any settings." if test_mode else "This does not send email or mark anything as complete."
    ttk.Label(window, text=note, font=("Calibri", 10), wraplength=500, justify="left").pack(anchor="w", pady=(0, 10))
    feedback = tk.StringVar(value="")
    ttk.Label(window, textvariable=feedback, foreground="#9C0006", wraplength=500, justify="left").pack(anchor="w", pady=(0, 6))
    choices = ttk.Frame(window)
    choices.pack(fill="x", pady=(8, 0))
    selected = {"action": None}

    def choose(action):
        try:
            if test_mode:
                selected["action"] = action
            elif action == "snooze":
                if on_snooze is None:
                    raise RuntimeError("One-hour reminder setup is unavailable.")
                on_snooze()
                selected["action"] = action
            elif action == "review":
                if on_dismiss:
                    on_dismiss()
                if on_review:
                    on_review()
                selected["action"] = action
            else:
                if on_dismiss:
                    on_dismiss()
                selected["action"] = "dismiss"
        except Exception as exc:
            feedback.set(f"Could not complete that choice. {exc}")
            return False
        window.destroy()
        if owned_root:
            host.quit()
        return True

    ttk.Button(choices, text="Yes, open review", command=lambda: choose("review")).pack(side="left", padx=(0, 8))
    dismiss_button = ttk.Button(choices, text="No, remind tomorrow", command=lambda: choose("dismiss"))
    dismiss_button.pack(side="left", padx=(0, 8))
    ttk.Button(choices, text="Remind me in 1 hour", command=lambda: choose("snooze")).pack(side="left")
    dismiss_button.focus_set()
    window.protocol("WM_DELETE_WINDOW", lambda: choose("dismiss"))
    window.bind("<Escape>", lambda _: choose("dismiss"))
    window.update_idletasks()
    width, height = window.winfo_width(), window.winfo_height()
    x = max(0, (window.winfo_screenwidth() - width) // 2)
    y = max(0, (window.winfo_screenheight() - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.lift()
    window.focus_force()
    # Briefly keep it above other windows so it is noticed, then let the user
    # switch freely without the reminder pinning itself over their work.
    window.after(1800, lambda: window.winfo_exists() and window.attributes("-topmost", False))
    if owned_root:
        host.mainloop()
        try:
            host.destroy()
        except tk.TclError:
            pass
    else:
        host.wait_window(window)
    return selected["action"]


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
