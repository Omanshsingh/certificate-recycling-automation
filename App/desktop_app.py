"""Small desktop interface for checking inward rows and generating PDFs."""
from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_paths import DATA_ROOT
from certificate_core import parse_date, status_value
from run_certificates import run
from workbook_store import update_issued_status, upsert_certificate_details
from workflow_support import load_settings, save_json


def reminder_picker_parts(value):
    try:
        parsed = datetime.strptime(value, "%H:%M")
    except (TypeError, ValueError):
        parsed = datetime.strptime("11:00", "%H:%M")
    hour = parsed.hour % 12 or 12
    return f"{hour:02d}", f"{parsed.minute:02d}", "AM" if parsed.hour < 12 else "PM"


def reminder_time_from_picker(hour, minute, period):
    hour, minute = int(hour), int(minute)
    if not 1 <= hour <= 12 or not 0 <= minute <= 59 or period not in {"AM", "PM"}:
        raise ValueError("Choose a valid reminder time.")
    hour = hour % 12 + (12 if period == "PM" else 0)
    return f"{hour:02d}:{minute:02d}"


class App:
    FILTERS = ("All rows", "Alerts due (day 25+)", "Ready", "Alert", "Generated", "Generated Manual", "Completed", "No status selected", "N/A")

    def __init__(self, root, base=None):
        self.root = root
        self.base = Path(base or DATA_ROOT)
        self.settings = load_settings(self.base)
        self.root.title("Ecoreco Certificate Automation")
        self.root.geometry("1120x720")
        self.root.minsize(880, 570)
        self.events = queue.Queue()
        self.busy = False
        self.close_when_done = False
        self.last_result = {}
        self.workbooks = [tk.StringVar(value=item) for item in self.settings["workbook_paths"]]
        self.reminders_enabled = tk.BooleanVar(value=self.settings.get("reminders_enabled", True))
        hour, minute, period = reminder_picker_parts(self.settings["reminder_time"])
        self.hour, self.minute, self.period = tk.StringVar(value=hour), tk.StringVar(value=minute), tk.StringVar(value=period)
        self.status = tk.StringVar(value="Choose both plant workbooks once, then mark approved rows Ready in Excel.")
        self.filter_value = tk.StringVar(value="All rows")
        self.custom_cod = tk.StringVar(value="")
        self.pending_selected = None
        self._build()
        self.show_results()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(150, self.poll)

    def _build(self):
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Ecoreco Certificate Automation", font=("Segoe UI", 18, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Creates certificate PDFs only. It never creates email drafts or sends mail. Mark a row Ready in Excel to generate its certificate.",
            wraplength=1050,
        ).pack(anchor="w", pady=(4, 12))

        files = ttk.LabelFrame(outer, text="Plant workbooks — choose each once", padding=9)
        files.pack(fill="x", pady=(0, 8))
        for index, variable in enumerate(self.workbooks):
            ttk.Label(files, text=f"Plant {index + 1}", width=9).grid(row=index, column=0, sticky="w", pady=3)
            ttk.Entry(files, textvariable=variable).grid(row=index, column=1, sticky="ew", padx=5, pady=3)
            ttk.Button(files, text="Choose file…", command=lambda i=index: self.choose_file(i)).grid(row=index, column=2, padx=4)
        files.columnconfigure(1, weight=1)

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(2, 8))
        self.button(actions, "Save setup", self.save_setup)
        self.button(actions, "Check both files", lambda: self.start(check_only=True))
        self.button(actions, "Review alerts", self.review_alerts)
        self.button(actions, "Generate selected PDF", lambda: self.start(selected_only=True))
        self.button(actions, "Generate all Ready PDFs", self.start)
        self.button(actions, "Create / edit from file…", self.create_from_file)
        self.button(actions, "Open certificate folder", self.open_certificates)

        custom_cod = ttk.Frame(outer)
        custom_cod.pack(fill="x", pady=(0, 8))
        ttk.Label(custom_cod, text="Custom COD for selected PDF (optional):").pack(side="left")
        ttk.Entry(custom_cod, textvariable=self.custom_cod, width=20).pack(side="left", padx=(6, 8))
        ttk.Label(custom_cod, text="Leave empty to use the next number automatically.").pack(side="left")

        reminders = ttk.LabelFrame(outer, text="Daily reminder", padding=9)
        reminders.pack(fill="x", pady=(0, 8))
        ttk.Checkbutton(reminders, text="Show one reminder when inward details reach day 25", variable=self.reminders_enabled).pack(side="left")
        ttk.Label(reminders, text="Time:").pack(side="left", padx=(20, 5))
        ttk.Spinbox(reminders, from_=1, to=12, width=3, textvariable=self.hour, format="%02.0f").pack(side="left")
        ttk.Label(reminders, text=":").pack(side="left")
        ttk.Spinbox(reminders, from_=0, to=59, width=3, textvariable=self.minute, format="%02.0f").pack(side="left")
        ttk.Combobox(reminders, state="readonly", values=("AM", "PM"), width=4, textvariable=self.period).pack(side="left", padx=4)
        self.button(reminders, "Save & update reminder", self.save_setup_and_schedule)
        ttk.Label(outer, text="Save and close Excel before checking or generating. Warnings are shown before creation; you can explicitly accept them. PDFs are named COD_Company_city_1.pdf or COD_Company_city_2.pdf; COD matches the number written to Excel.", wraplength=1060).pack(anchor="w", pady=(0, 8))

        line = ttk.Frame(outer)
        line.pack(fill="x", pady=(2, 3))
        ttk.Label(line, text="Work list:").pack(side="left")
        self.filter_box = ttk.Combobox(line, state="readonly", values=self.FILTERS, textvariable=self.filter_value, width=23)
        self.filter_box.pack(side="left", padx=(6, 14))
        self.filter_box.bind("<<ComboboxSelected>>", lambda _: self.show_results())
        self.button(line, "Open selected plant Excel", self.open_selected_workbook)
        self.button(line, "Mark selected Completed", self.mark_completed)
        self.summary = ttk.Label(line, text="")
        self.summary.pack(side="right")

        columns = ("plant", "sn", "status", "client", "reference", "cod", "age", "details")
        self.table = ttk.Treeview(outer, columns=columns, show="headings", selectmode="browse")
        headings = {"plant": "Plant", "sn": "Client ID", "status": "Status", "client": "Client", "reference": "Reference", "cod": "COD", "age": "Inward age", "details": "What to do"}
        widths = {"plant": 70, "sn": 82, "status": 110, "client": 175, "reference": 125, "cod": 75, "age": 86, "details": 390}
        for column in columns:
            self.table.heading(column, text=headings[column])
            self.table.column(column, width=widths[column], minwidth=55, stretch=column in {"client", "details"})
        yscroll = ttk.Scrollbar(outer, orient="vertical", command=self.table.yview)
        xscroll = ttk.Scrollbar(outer, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.table.pack(side="top", fill="both", expand=True)
        xscroll.pack(fill="x")
        yscroll.place(in_=self.table, relx=1.0, rely=0, relheight=1.0, anchor="ne")
        self.table.tag_configure("alert", background="#c00000", foreground="white")
        self.table.tag_configure("ready", background="#d9ead3")
        self.table.tag_configure("generated", background="#fff2cc")
        self.table.tag_configure("generated_manual", background="#dceaf4")
        self.table.tag_configure("alert", background="#fce4d6")
        self.table.tag_configure("completed", background="#d9eaf7")
        ttk.Label(outer, textvariable=self.status, wraplength=1050).pack(anchor="w", pady=(7, 0))

    @staticmethod
    def button(parent, text, command):
        button = ttk.Button(parent, text=text, command=command)
        button.pack(side="left", padx=(0, 6))
        return button

    def choose_file(self, index):
        path = filedialog.askopenfilename(
            parent=self.root, title=f"Choose Plant {index + 1} Excel workbook",
            filetypes=(("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")),
        )
        if path:
            self.workbooks[index].set(path)

    def selected_paths(self):
        paths = [variable.get().strip() for variable in self.workbooks]
        if not all(paths):
            raise ValueError("Choose the Plant 1 and Plant 2 Excel files first.")
        resolved = [Path(value).expanduser().resolve() for value in paths]
        if not all(path.is_file() for path in resolved):
            raise ValueError("A saved plant workbook cannot be found. Choose its current location in Settings.")
        if resolved[0] == resolved[1] or resolved[0].samefile(resolved[1]):
            raise ValueError("Choose two different Excel files, one for each plant.")
        return [str(path) for path in resolved]

    def save_setup(self):
        try:
            paths = [item.strip() for item in (var.get() for var in self.workbooks)]
            if any(paths) and not all(paths):
                raise ValueError("Choose both plant files, or leave both blank until they are available.")
            if all(paths):
                self.selected_paths()
            reminder_time = reminder_time_from_picker(self.hour.get(), self.minute.get(), self.period.get())
            self.settings["workbook_paths"] = paths
            self.settings["reminder_time"] = reminder_time
            self.settings["reminders_enabled"] = self.reminders_enabled.get()
            save_json(self.base / "settings.local.json", self.settings)
            self.status.set(f"Setup saved. Daily reminder: {'On' if self.reminders_enabled.get() else 'Off'} at {self._time_label(reminder_time)}.")
            return True
        except Exception as exc:
            messagebox.showerror("Setup needs attention", str(exc), parent=self.root)
            return False

    def save_setup_and_schedule(self):
        if not self.save_setup():
            return
        if not self.reminders_enabled.get():
            try:
                result = self._powershell("Disable")
                if result.returncode:
                    raise RuntimeError(result.stderr.strip() or result.stdout.strip())
                self.status.set("Daily reminders are off; the Windows reminder task was removed.")
            except Exception as exc:
                messagebox.showerror("Reminder update failed", str(exc), parent=self.root)
            return
        try:
            self.selected_paths()
            result = self._powershell("Install", (("-ReminderTime", self.settings["reminder_time"]),))
            if result.returncode:
                raise RuntimeError((result.stdout or result.stderr).strip() or "Windows could not set up the daily reminder.")
            self.status.set(f"Daily reminder is on for {self._time_label(self.settings['reminder_time'])}. It starts at inward day 25 and continues if overdue.")
        except Exception as exc:
            messagebox.showerror("Reminder setup failed", str(exc), parent=self.root)

    def _powershell(self, mode, extra=()):
        script = self.base / "App" / "Install Daily Task.ps1"
        command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-Mode", mode]
        for key, value in extra:
            command.extend([key, str(value)])
        return subprocess.run(command, capture_output=True, text=True, timeout=45, check=False)

    def start(self, check_only=False, selected_only=False):
        if self.busy:
            return
        try:
            manual_cod = self.custom_cod.get().strip()
            if manual_cod and not (selected_only or check_only):
                messagebox.showinfo(
                    "Custom COD is for one client",
                    "Select one Ready client and use Generate selected PDF, or clear the Custom COD box before generating all Ready PDFs.",
                    parent=self.root,
                )
                return
            paths = self.selected_paths()
            if not self.save_setup():
                return
            selected = None
            if selected_only:
                row = None
                if self.pending_selected:
                    selected = self.pending_selected
                    self.pending_selected = None
                else:
                    row = self.selected_row()
                if selected is None and not row:
                    messagebox.showinfo("Select a client", "Select a Ready row in the work list first.", parent=self.root)
                    return
                if selected is None and row.get("certificate_status") != "ready":
                    messagebox.showinfo("Not marked Ready", "In Excel, change this client's Certificate Issued cell to Ready, save and close the workbook, then try again.", parent=self.root)
                    return
                if selected is None:
                    selected = (row["plant"], str(row["sn"]))
                client_label = (row or {}).get("supplier") or selected[1]
                if manual_cod and not messagebox.askyesno(
                    "Use this custom COD?",
                    f"Generate {client_label} with COD {manual_cod}?\n\n"
                    "The number must be unique. After the PDF is created, this COD and its date will be saved back to the selected Excel row.",
                    parent=self.root,
                ):
                    return
            self.busy = True
            action = "Checking both plant files…" if check_only else ("Generating selected certificate PDF…" if selected else "Generating PDFs for all Ready rows…")
            self.status.set(action)
            threading.Thread(target=self._worker, args=(paths, check_only, selected, manual_cod if selected else None, False), daemon=True).start()
        except Exception as exc:
            messagebox.showerror("Cannot start", str(exc), parent=self.root)

    def _worker(self, paths, check_only, selected, manual_cod=None, allow_warnings=False):
        try:
            if not check_only and not allow_warnings:
                # A dry run prevents valid rows from being generated before
                # the operator has seen and accepted warnings on other rows.
                preflight = run(paths, self.base, check_only=True, selected=selected,
                                custom_cod=manual_cod, progress=lambda _: None)
                preflight_payload = json.loads((preflight[2] / "Latest Results.json").read_text(encoding="utf-8"))
                warnings = [row for row in preflight_payload.get("rows", [])
                            if row.get("outcome") == "needs_review"]
                if warnings:
                    self.events.put(("confirm_warnings", preflight_payload, manual_cod, paths, selected))
                    return
            result = run(paths, self.base, check_only=check_only, selected=selected,
                         custom_cod=manual_cod, allow_warnings=allow_warnings,
                         progress=lambda _: None)
            payload = json.loads((result[2] / "Latest Results.json").read_text(encoding="utf-8"))
            if selected and not check_only:
                generated_rows = {
                    (row.get("plant"), str(row.get("sn"))): row
                    for row in payload.get("rows", [])
                }
                run(paths, self.base, check_only=True, progress=lambda _: None)
                refreshed = json.loads((result[2] / "Latest Results.json").read_text(encoding="utf-8"))
                for row in refreshed.get("rows", []):
                    prior = generated_rows.get((row.get("plant"), str(row.get("sn"))))
                    if prior and prior.get("warnings"):
                        row["warnings"] = prior["warnings"]
                        row["system_alert"] = True
                        row["detail"] = prior.get("detail", row.get("detail", ""))
                payload = refreshed
                payload["created"] = result[0]
                save_json(result[2] / "Latest Results.json", payload)
            self.events.put(("done", payload, manual_cod))
        except Exception as exc:
            self.events.put(("error", str(exc), manual_cod))

    def create_from_file(self):
        """Enter certificate details from a workbook chosen in Explorer."""
        path = filedialog.askopenfilename(
            parent=self.root, title="Choose the plant workbook for this certificate",
            filetypes=(("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")),
        )
        if not path:
            return
        chosen = Path(path).expanduser().resolve()
        configured = [Path(value).expanduser().resolve() for value in (var.get().strip() for var in self.workbooks) if value.strip()]
        if chosen not in configured:
            messagebox.showerror(
                "Workbook not in setup",
                "Choose one of the configured Plant 1 or Plant 2 workbooks, or save this file as the plant workbook in setup first.",
                parent=self.root,
            )
            return
        plant = str(configured.index(chosen) + 1)
        self._show_details_dialog(chosen, plant)

    def _show_details_dialog(self, path, plant):
        fields = (
            ("case_id", "Client ID / SN"), ("inward_date", "Inward date (DD.MM.YYYY)"),
            ("supplier", "Company / supplier"), ("location", "City"),
            ("item", "Item"), ("quantity", "Quantity"),
            ("per_piece_weight", "Per-piece weight (KG)"), ("weight", "Inward weight (KG)"),
            ("certificate_weight", "Certificate / net weight (KG)"),
            ("form6", "Form 6 No. (mandatory)"), ("reference", "Reference No. (optional)"),
            ("cod_no", "COD No. (optional)"), ("cod_date", "COD date (optional)"),
        )
        existing = {}
        selected = self.selected_row()
        if selected and Path(selected.get("file", "")).resolve() == Path(path).resolve():
            existing = {
                "case_id": selected.get("sn", ""), "supplier": selected.get("supplier", ""),
                "reference": selected.get("reference", ""), "cod_no": selected.get("certificate_no", ""),
                "inward_date": selected.get("inward_date", ""),
            }
        dialog = tk.Toplevel(self.root)
        dialog.title("Certificate details — Excel will be updated")
        dialog.transient(self.root)
        dialog.grab_set()
        body = ttk.Frame(dialog, padding=12)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Enter or replace any certificate detail. The values are written to Excel, marked Ready, then checked before PDF creation.", wraplength=560).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        variables = {}
        for index, (key, label) in enumerate(fields, 1):
            ttk.Label(body, text=label).grid(row=index, column=0, sticky="w", padx=(0, 8), pady=3)
            value = existing.get(key, "")
            variables[key] = tk.StringVar(value=str(value or ""))
            ttk.Entry(body, textvariable=variables[key], width=44).grid(row=index, column=1, sticky="ew", pady=3)
        body.columnconfigure(1, weight=1)
        buttons = ttk.Frame(body)
        buttons.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="e", pady=(10, 0))

        def save_details():
            details = {key: variable.get().strip() for key, variable in variables.items()}
            if not details["case_id"]:
                messagebox.showerror("Client ID required", "Enter the Client ID / SN so the row can be found or added.", parent=dialog)
                return
            try:
                upsert_certificate_details(
                    path, self.base / "Data" / "Workbook Backups", details,
                )
                dialog.destroy()
                self.pending_selected = (plant, details["case_id"])
                self.status.set(f"Details saved to {Path(path).name}. Checking warnings before creating the PDF…")
                self.start(selected_only=True)
            except Exception as exc:
                messagebox.showerror("Could not save certificate details", str(exc), parent=dialog)

        ttk.Button(buttons, text="Save details and create PDF", command=save_details).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="left")
        dialog.wait_window()

    def review_alerts(self):
        self.filter_value.set("Alerts due (day 25+)")
        self.start(check_only=True)

    def _status_label(self, row):
        state = status_value(row.get("certificate_status"))
        if state == "blank":
            return ""
        if state in {"ready", "alert", "blank", "unknown"} and (
                row.get("outcome") == "needs_review" or row.get("system_alert")):
            return "Alert"
        return {
            "ready": "Ready",
            "alert": "Alert",
            "generated": "Generated",
            "generated_manual": "Generated Manual",
            "completed": "Completed",
            "na": "N/A",
            "blank": "",
            "unknown": "Alert",
        }.get(state, row.get("outcome", "Alert").replace("_", " ").title())

    def _filter_match(self, row, selected_filter):
        state = status_value(row.get("certificate_status"))
        if selected_filter == "All rows":
            return True
        if selected_filter == "Alerts due (day 25+)":
            return bool(row.get("reminder_due") or row.get("system_alert"))
        if selected_filter == "Alert":
            return bool(row.get("outcome") == "needs_review" or row.get("system_alert")
                        or state in {"alert", "unknown", "blank"})
        expected = {"Ready": "ready", "Generated": "generated", "Generated Manual": "generated_manual",
                    "Completed": "completed", "No status selected": "blank", "N/A": "na"}.get(selected_filter)
        return state == expected

    def show_results(self):
        for item in self.table.get_children():
            self.table.delete(item)
        path = self.base / "Data" / "Reports" / "Latest Results.json"
        try:
            result = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(result, dict) or not isinstance(result.get("rows"), list):
                raise ValueError("The saved results are not in the expected format.")
            self.last_result = result
        except (OSError, ValueError, json.JSONDecodeError):
            result = {"rows": []}
            self.last_result = result
        rows = result.get("rows", [])
        state_counts = {state: sum(status_value(row.get("certificate_status")) == state for row in rows)
                        for state in ("ready", "alert", "generated", "generated_manual", "completed", "na")}
        due_count = sum(bool(row.get("reminder_due")) for row in rows)
        self.summary.configure(text=f"Due: {due_count}  |  Ready: {state_counts['ready']}  |  Alert: {state_counts['alert']}  |  Generated: {state_counts['generated']}  |  Generated Manual: {state_counts['generated_manual']}  |  Completed: {state_counts['completed']}  |  N/A: {state_counts['na']}")
        selected_filter = self.filter_value.get()
        for row in rows:
            if not self._filter_match(row, selected_filter):
                continue
            age = row.get("age_days")
            age_label = f"Day {age}" if isinstance(age, int) else "Date missing"
            status = self._status_label(row)
            normalized_status = status_value(row.get("certificate_status"))
            tags = ("alert",) if row.get("reminder_due") or row.get("system_alert") else (normalized_status,)
            self.table.insert("", "end", values=(
                row.get("plant", ""), row.get("sn", ""), status, row.get("supplier", ""),
                row.get("reference", ""), row.get("certificate_no", ""), age_label, row.get("detail", ""),
            ), tags=tags, iid=str(len(self.table.get_children())))

    def selected_row(self):
        chosen = self.table.selection()
        if not chosen:
            return None
        item = self.table.item(chosen[0])
        values = item.get("values", [])
        plant, sn = str(values[0]), str(values[1])
        return next((row for row in self.last_result.get("rows", [])
                     if row.get("plant") == plant and str(row.get("sn")) == sn), None)

    def open_selected_workbook(self):
        row = self.selected_row()
        if not row or not row.get("file"):
            self.status.set("Select a client row first, then open its plant workbook.")
            return
        try:
            os.startfile(row["file"])
        except OSError as exc:
            self.status.set(f"Could not open the workbook: {exc}")

    def mark_completed(self):
        row = self.selected_row()
        if not row:
            self.status.set("Select a generated certificate row first.")
            return
        if status_value(row.get("certificate_status")) not in {"generated", "generated_manual"}:
            self.status.set("Only a Generated or Generated Manual certificate can be marked Completed.")
            return
        if not messagebox.askyesno("Mark certificate Completed?", f"Have you sent COD {row.get('certificate_no')} to {row.get('supplier')}?\n\nThis only updates Excel. The automation does not send email.", parent=self.root):
            return
        try:
            update_issued_status(
                row["file"], self.base / "Data" / "Workbook Backups",
                row["sn"], row["sheet"], row["row"], row["certificate_no"],
                row.get("certificate_status", "generated"), "Completed",
            )
            self.status.set("Excel updated to Completed. This client will no longer appear in the reminder list.")
            self.start(check_only=True)
        except Exception as exc:
            messagebox.showerror("Could not update Excel", str(exc), parent=self.root)

    def open_certificates(self):
        folder = self.base / "Certificates"
        folder.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(folder))
        except OSError as exc:
            self.status.set(f"Could not open the certificate folder: {exc}")

    @staticmethod
    def _time_label(value):
        return datetime.strptime(value, "%H:%M").strftime("%I:%M %p").lstrip("0")

    def poll(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind, value = event[:2]
                manual_cod = event[2] if len(event) > 2 else None
                self.busy = False
                if kind == "done":
                    self.last_result = value
                    self.show_results()
                    created = value.get("created", 0)
                    if created and manual_cod:
                        self.custom_cod.set("")
                    self.status.set(f"Finished. {created} certificate PDF(s) generated. All PDFs are in the Certificates folder.")
                elif kind == "confirm_warnings":
                    payload = value
                    warning_rows = [row for row in payload.get("rows", []) if row.get("outcome") == "needs_review"]
                    lines = []
                    for row in warning_rows[:12]:
                        lines.append(f"{row.get('plant')} / SN {row.get('sn')}: {row.get('detail')}")
                    if len(warning_rows) > 12:
                        lines.append(f"…and {len(warning_rows) - 12} more row(s).")
                    accepted = messagebox.askyesno(
                        "Warnings found — create anyway?",
                        "The following issues will be recorded as warnings. If you continue, the PDF may contain blank, fallback, or manually corrected values; Excel will be marked Generated and you must verify the result.\n\n"
                        + "\n".join(lines), parent=self.root,
                    )
                    if accepted:
                        paths, selected = event[3], event[4]
                        self.busy = True
                        self.status.set("Creating certificate PDF(s) with the accepted warnings…")
                        threading.Thread(target=self._worker, args=(paths, False, selected, manual_cod, True), daemon=True).start()
                    else:
                        self.last_result = payload
                        self.show_results()
                        self.status.set("No certificate was created. Correct the flagged rows or run again and accept the warnings.")
                else:
                    self.status.set(value)
                    messagebox.showerror("Could not complete", value, parent=self.root)
                if self.close_when_done:
                    self.root.destroy()
                    return
        except queue.Empty:
            pass
        self.root.after(150, self.poll)

    def close(self):
        if self.busy:
            self.close_when_done = True
            self.status.set("The app will close when its current file update is finished.")
        else:
            self.root.destroy()


if __name__ == "__main__":
    from certificate_inventory import ensure_certificate_folders
    ensure_certificate_folders(DATA_ROOT)
    app_root = tk.Tk()
    App(app_root)
    app_root.mainloop()
