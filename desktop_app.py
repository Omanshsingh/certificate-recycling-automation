"""Simple desktop controls for the two plant workbooks."""
from __future__ import annotations

import json
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

BASE = Path(__file__).resolve().parent


def load_settings(base):
    from workflow_support import load_settings as load
    return load(base)


def save_json(path, value):
    from workflow_support import save_json as save
    return save(path, value)


def run(*args, **kwargs):
    # Keep the controls available even when a processing dependency is missing.
    from run_certificates import run as process
    return process(*args, **kwargs)


def error_message(exc):
    if isinstance(exc, ImportError):
        return "A required Python package is missing. Run Setup Once.bat, then reopen the app. " + str(exc)
    return str(exc)


class App:
    def __init__(self, root, base=BASE):
        self.root, self.base = root, Path(base)
        self.root.title("Ecoreco Certificate Automation")
        self.root.geometry("1080x760")
        self.root.minsize(860, 600)
        self.events = queue.Queue()
        self.busy = False
        self.close_when_done = False
        self.settings_error = ""
        try:
            self.settings = load_settings(self.base)
        except Exception as exc:
            self.settings_error = error_message(exc)
            # Display placeholders only; never save over unreadable configuration.
            self.settings = {"workbook_paths": ["", ""]}
        self.workbooks = [tk.StringVar(value=p) for p in self.settings["workbook_paths"]]
        self.mailbox = tk.StringVar(value=self.settings.get("outlook_mailbox", ""))
        self.notifications = tk.BooleanVar(value=self.settings.get("quiet_notifications", True))
        self.first_cod = tk.StringVar(value=str(self.settings.get("certificate_number_start", 1)) if self.settings.get("initial_cod_confirmed", False) else "")
        self.status = tk.StringVar(value="Choose both plant files, then click Check both files.")
        self.last_result = None
        self.queue_filter = tk.StringVar(value="All")
        self.queue_summary = tk.StringVar(value="Check both files to see what needs attention.")
        self.visible_rows = {}
        self.displayed_results_path = self.base / "Output" / "Latest Results.txt"
        self.controls = []
        self.build()
        self.show_results()
        if self.settings_error:
            self.status.set("Settings unavailable: " + self.settings_error + " Repair the file, then click Reload settings.")
        elif self.last_result:
            previous_status = self.previous_results_message()
            self.status.set(previous_status)
            if "Excel files changed after the last check" in previous_status:
                self.clear_result_table()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(150, self.poll)

    def button(self, parent, label, command):
        button = ttk.Button(parent, text=label, command=command)
        button.pack(side="left", padx=(0, 8))
        self.controls.append(button)
        return button

    def build(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Ecoreco Certificate Automation", font=("Calibri", 18, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Creates certificate email drafts in Classic Outlook for review. It never sends them.", wraplength=900).pack(anchor="w", pady=(4, 8))

        self.plants_frame = ttk.LabelFrame(frame, text="1. Plant Excel files", padding=10)
        self.plants_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(
            self.plants_frame,
            text="Choose each file once. The app remembers both locations. Save and close the workbooks before creating drafts.",
            wraplength=900,
        ).pack(anchor="w", pady=(0, 4))
        for index, variable in enumerate(self.workbooks):
            row = ttk.Frame(self.plants_frame)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=f"Plant {index + 1}", width=10).pack(side="left")
            ttk.Entry(row, textvariable=variable, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 8))
            self.button(row, "Choose file", lambda i=index: self.choose_workbook(i))

        mailbox_frame = ttk.LabelFrame(frame, text="2. Outlook mailbox for drafts", padding=10)
        mailbox_frame.pack(fill="x", pady=(0, 8))
        account_row = ttk.Frame(mailbox_frame)
        account_row.pack(fill="x")
        self.accounts = ttk.Combobox(account_row, textvariable=self.mailbox, state="readonly", width=38)
        if self.mailbox.get():
            self.accounts["values"] = [self.mailbox.get()]
        self.accounts.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.accounts.bind("<<ComboboxSelected>>", lambda _: self.save_settings())
        self.button(account_row, "Find Outlook mailboxes", self.load_accounts)

        self.settings_button = ttk.Button(frame, text="Setup options", command=self.toggle_settings)
        self.settings_button.pack(anchor="w", pady=(0, 6))
        self.controls.append(self.settings_button)
        self.settings_frame = ttk.LabelFrame(frame, text="Setup options", padding=8)
        self.settings_visible = False
        initial = ttk.Frame(self.settings_frame)
        initial.pack(fill="x", pady=3)
        ttk.Label(initial, text="First certificate number", width=22).pack(side="left")
        number_entry = ttk.Entry(initial, textvariable=self.first_cod, width=12)
        number_entry.pack(side="left", padx=(0, 8))
        self.controls.append(number_entry)
        ttk.Label(initial, text="Only if neither plant file has a COD number; otherwise leave blank.", wraplength=650).pack(side="left")

        settings_actions = ttk.Frame(self.settings_frame)
        settings_actions.pack(fill="x", pady=(4, 0))
        self.button(settings_actions, "Open settings folder", self.open_settings_folder)
        self.button(settings_actions, "Reload settings", self.reload_settings)
        notification_control = ttk.Checkbutton(self.settings_frame, text="Daily reminder: one quiet summary while cases need attention (requires daily task setup)", variable=self.notifications, command=self.save_settings)
        notification_control.pack(anchor="w", pady=(4, 0))
        self.controls.append(notification_control)

        self.actions_frame = ttk.LabelFrame(frame, text="3. Check and create drafts", padding=10)
        self.actions_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(
            self.actions_frame,
            text="In Excel: Ready = draft now (no 30-day wait), N/A = skip, blank = choose Ready or N/A. Save and close Excel, then check again. COD is updated after a confirmed draft save.",
            wraplength=900,
        ).pack(anchor="w", pady=(0, 6))
        plant_actions = ttk.Frame(self.actions_frame)
        plant_actions.pack(fill="x")
        self.button(plant_actions, "Check both files", lambda: self.start(check_only=True))
        self.button(plant_actions, "Create one draft", lambda: self.start(limit=1))
        self.button(plant_actions, "Create all Ready drafts", self.start)
        extras = ttk.Frame(frame)
        extras.pack(fill="x", pady=(0, 4))
        ttk.Label(extras, text="Review saved drafts in Classic Outlook > Drafts.").pack(side="left", padx=(0, 12))
        self.button(extras, "Open results", self.open_results)
        self.button(extras, "Open output folder", self.open_output)
        queue_bar = ttk.Frame(frame)
        queue_bar.pack(fill="x", pady=4)
        queue_choices = ttk.Combobox(queue_bar, textvariable=self.queue_filter, state="readonly", width=23,
                                    values=("All", "Needs attention", "Ready to draft", "Drafts to review", "Marked sent", "Skipped"))
        queue_choices.pack(side="left", padx=(0, 8))
        queue_choices.bind("<<ComboboxSelected>>", lambda _: self.show_results())
        self.button(queue_bar, "Open selected plant", self.open_selected_plant)
        self.button(queue_bar, "Mark sent / Undo", self.mark_selected_sent)
        ttk.Label(frame, textvariable=self.queue_summary, wraplength=900).pack(anchor="w", pady=(0, 4))
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)
        self.table = ttk.Treeview(table_frame, columns=("plant", "sn", "status", "detail"), show="headings")
        for name, title, width in (("plant", "Plant", 70), ("sn", "SN", 70), ("status", "Status", 130), ("detail", "Details", 580)):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, minwidth=60)
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=scroll.set, xscrollcommand=horizontal.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.table.tag_configure("needs_review", background="#fff3cd")
        self.table.tag_configure("ready", background="#e2f0d9")
        self.table.bind("<Double-1>", lambda _: self.open_selected_plant())
        ttk.Label(frame, textvariable=self.status, wraplength=920).pack(anchor="w", pady=(12, 0))

    def toggle_settings(self, force_open=False):
        show = force_open or not self.settings_visible
        if show:
            self.settings_frame.pack(fill="x", before=self.actions_frame, pady=(0, 8))
            self.settings_button.configure(text="Hide setup options")
        else:
            self.settings_frame.pack_forget()
            self.settings_button.configure(text="Show setup options")
        self.settings_visible = show

    def choose_workbook(self, index):
        if self.busy or not self.settings_available():
            return
        previous = self.workbooks[index].get()
        path = filedialog.askopenfilename(title=f"Select Plant {index + 1} Excel workbook", initialdir=str(Path(previous).parent) if previous else str(self.base / "Input"), filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if path:
            if any(Path(v.get()).resolve() == Path(path).resolve() for i, v in enumerate(self.workbooks) if i != index and v.get()):
                self.status.set("Choose a different file for each plant.")
                return
            self.workbooks[index].set(path)
            if self.save_settings():
                self.status.set(f"Plant {index + 1} location saved. It will be reused next time.")
            else:
                self.workbooks[index].set(previous)

    def settings_available(self):
        if self.settings_error:
            self.status.set("Settings unavailable: " + self.settings_error + " Repair the file, then click Reload settings.")
            return False
        return True

    def reload_settings(self):
        if self.busy:
            return
        try:
            settings = load_settings(self.base)
        except Exception as exc:
            self.settings_error = error_message(exc)
            self.settings_available()
            return
        self.settings = settings
        self.settings_error = ""
        for variable, value in zip(self.workbooks, settings["workbook_paths"]):
            variable.set(value)
        self.mailbox.set(settings.get("outlook_mailbox", ""))
        self.accounts["values"] = [self.mailbox.get()] if self.mailbox.get() else []
        self.notifications.set(settings.get("quiet_notifications", True))
        self.first_cod.set(str(settings.get("certificate_number_start", 1)) if settings.get("initial_cod_confirmed", False) else "")
        self.status.set("Settings reloaded. Both plant locations are available for review.")

    def open_settings_folder(self):
        try:
            os.startfile(str(self.base))
        except OSError as exc:
            self.status.set(f"Could not open the settings folder: {exc}")

    def save_settings(self):
        if self.busy or not self.settings_available():
            return False
        settings = dict(self.settings)
        settings.update(workbook_paths=[v.get().strip() for v in self.workbooks], outlook_mailbox=self.mailbox.get().strip(), quiet_notifications=self.notifications.get())
        initial = self.first_cod.get().strip()
        if initial and (not initial.isascii() or not initial.isdecimal() or len(initial) > 12 or not 1 <= int(initial) <= 999999999999):
            self.status.set("First COD number must be a whole number from 1 to 999999999999, or left blank.")
            return False
        settings["initial_cod_confirmed"] = bool(initial)
        if initial:
            settings["certificate_number_start"] = int(initial)
        try:
            # Re-read to detect corruption since startup before replacing anything.
            latest = load_settings(self.base)
        except Exception as exc:
            self.settings_error = error_message(exc)
            self.settings_available()
            return False
        # Preserve unrelated settings changed outside this window since startup.
        latest.update({key: settings[key] for key in ("workbook_paths", "outlook_mailbox", "quiet_notifications", "initial_cod_confirmed")})
        if initial:
            latest["certificate_number_start"] = int(initial)
        settings = latest
        try:
            save_json(self.base / "settings.local.json", settings)
        except Exception as exc:
            self.status.set(f"Settings were not saved. Check folder write access and try again. {error_message(exc)}")
            return False
        self.settings = settings
        return True

    def set_busy(self, busy):
        self.busy = busy
        for control in self.controls:
            control.configure(state="disabled" if busy else "normal")
        self.accounts.configure(state="disabled" if busy else "readonly")

    def load_accounts(self):
        if self.busy or not self.settings_available():
            return
        self.set_busy(True)
        self.status.set("Reading Classic Outlook accounts...")
        def worker():
            initialized = False
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                initialized = True
                app = win32com.client.Dispatch("Outlook.Application")
                self.events.put(("accounts", [str(a.SmtpAddress) for a in app.Session.Accounts]))
            except Exception as exc:
                self.events.put(("error", f"Could not load Classic Outlook accounts: {error_message(exc)}"))
            finally:
                if initialized:
                    pythoncom.CoUninitialize()
        self.launch_worker(worker)

    def launch_worker(self, target, args=()):
        try:
            threading.Thread(target=target, args=args, daemon=True).start()
        except Exception as exc:
            self.set_busy(False)
            self.status.set(f"Could not start processing: {error_message(exc)}")

    def start(self, drafts=True, limit=None, check_only=False):
        if self.busy or not self.settings_available():
            return
        paths = [v.get().strip() for v in self.workbooks]
        if len(paths) != 2 or not all(paths):
            self.status.set("Choose both plant Excel files in the section above before continuing.")
            return
        try:
            resolved = [Path(p).resolve() for p in paths]
            if resolved[0] == resolved[1] or (all(p.is_file() for p in resolved) and resolved[0].samefile(resolved[1])):
                self.status.set("Choose a different file for each plant.")
                return
            for path in resolved:
                if not path.is_file():
                    self.status.set(f"Saved workbook unavailable: {path}. Reconnect its drive or choose its new location.")
                    return
        except OSError as exc:
            self.status.set(f"Could not access the plant files: {exc}")
            return
        if drafts and not check_only and not self.mailbox.get().strip():
            self.status.set("Click Find Outlook mailboxes, then choose where drafts should be saved. Checking files does not need Outlook.")
            return
        if drafts and not check_only:
            if limit == 1:
                prompt = (
                    "This creates one certificate using the normal template and saves its email as an unsent Outlook draft. "
                    "After Outlook confirms the save, its COD number and date are written to that Excel row. Continue?"
                )
                title = "Create one draft?"
            else:
                prompt = (
                    "This creates a certificate and unsent Outlook draft for each eligible Ready row in both files. "
                    "After each confirmed save, its COD number and date are written to Excel. Continue?"
                )
                title = "Create all Ready drafts?"
            if not messagebox.askyesno(title, prompt, parent=self.root):
                self.status.set("Cancelled. No drafts or Excel updates were made.")
                return
        if not self.save_settings():
            return
        self.set_busy(True)
        self.status.set("Checking both plant files..." if check_only else "Creating certificate and Outlook draft...")
        self.launch_worker(self.worker, (paths, self.mailbox.get().strip(), drafts, limit, check_only))

    def worker(self, paths, mailbox, drafts, limit, check_only):
        initialized = False
        try:
            if drafts and not check_only:
                import pythoncom
                pythoncom.CoInitialize()
                initialized = True
            created, review, output = run(
                paths, mailbox, self.base, limit=limit, create_drafts=drafts, check_only=check_only,
                progress=lambda msg: self.events.put(("status", msg)),
            )
            self.events.put(("done", f"Finished: {created} draft(s) created; {review} need attention. Open results for details."))
        except Exception as exc:
            self.events.put(("error", error_message(exc)))
        finally:
            if initialized:
                pythoncom.CoUninitialize()

    def show_results(self, result_path=None, update_last=True):
        path = Path(result_path) if result_path else self.base / "Output" / "Latest Results.json"
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            rows = data["rows"]
            if not isinstance(rows, list) or any(not isinstance(row, dict) or any(key not in row for key in ("plant", "sn", "outcome", "detail")) or not isinstance(row["outcome"], str) for row in rows):
                raise ValueError("The saved results have an invalid format.")
        except FileNotFoundError:
            return
        except (OSError, ValueError, KeyError, TypeError) as exc:
            self.status.set(f"Could not read the last results. Click Check both files to rebuild them. {exc}")
            return
        if update_last:
            self.last_result = data
        self.displayed_results_path = path.with_suffix(".txt")
        self.clear_result_table()
        self.visible_rows = {}
        counts = {}
        history = None
        try:
            from draft_service import History
            history = History(self.base / "Data" / "certificate-history.sqlite")
        except Exception as exc:
            self.status.set(f"Draft follow-up history unavailable: {exc}")
        for row in sorted(rows, key=lambda r: r["outcome"] != "needs_review"):
            key = ("plant2|" if row["plant"] == "plant2" else "") + str(row["sn"])
            label = {"needs_review": "Needs attention", "ready": "Ready to draft", "not_applicable": "Skipped",
                     "draft_created": "Drafts to review", "updated": "Drafts to review", "already_prepared": "Drafts to review"}.get(row["outcome"], row["outcome"].replace("_", " ").title())
            if history and label == "Drafts to review" and history.sent_acknowledgement(key):
                label = "Marked sent"
            counts[label] = counts.get(label, 0) + 1
            if self.queue_filter.get() not in ("All", label):
                continue
            detail = (str(row.get('supplier', '')).strip() + ' — ' if row.get('supplier') else '') + row['detail']
            item = self.table.insert("", "end", values=(row["plant"], row["sn"], label, detail), tags=(row["outcome"],))
            self.visible_rows[item] = row
        if history:
            history.close()
        self.queue_summary.set(" | ".join(f"{label}: {counts.get(label, 0)}" for label in ("Needs attention", "Ready to draft", "Drafts to review", "Marked sent")))

    def selected_row(self):
        selected = self.table.selection()
        row = self.visible_rows.get(selected[0]) if selected else None
        if not row:
            self.status.set("Select a case in the list first.")
        return row

    def open_selected_plant(self):
        if self.busy:
            return
        row = self.selected_row()
        if not row:
            return
        path = row.get("file", "")
        if not path or not Path(path).is_file():
            self.status.set("This workbook is unavailable. Check the saved plant location.")
            return
        try:
            os.startfile(path)
            self.status.set(f"Find SN {row['sn']} on sheet {row.get('sheet', '')}, row {row.get('row', '')}. Choose Ready or N/A, save and close Excel, then Check both files.")
        except OSError as exc:
            self.status.set(f"Could not open the plant workbook: {exc}")

    def mark_selected_sent(self):
        if self.busy:
            return
        row = self.selected_row()
        if not row:
            return
        if row['outcome'] not in {'draft_created', 'updated', 'already_prepared'}:
            self.status.set("Only a prepared certificate can be marked sent.")
            return
        from draft_service import History
        from workflow_support import run_lock
        history = None
        try:
            with run_lock(self.base):
                history = History(self.base / 'Data' / 'certificate-history.sqlite')
                key = ('plant2|' if row['plant'] == 'plant2' else '') + str(row['sn'])
                record = history.record({'case_id': key})
                if not record or record['status'] != 'saved':
                    raise ValueError('No confirmed local draft record. Check the existing certificate manually; it will not be marked sent automatically.')
                undo = bool(history.sent_acknowledgement(key))
                prompt = ('Return this case to Drafts to review?' if undo else
                          f"Have you already sent COD {record['certificate_no']} to the client? This only records your confirmation locally; it does not send mail or change Excel.")
                if messagebox.askyesno('Undo sent confirmation?' if undo else 'Confirm already sent', prompt, parent=self.root):
                    if undo:
                        history.undo_sent_acknowledgement(key)
                    else:
                        history.acknowledge_sent(key, record['certificate_no'])
            self.show_results()
        except Exception as exc:
            self.status.set(error_message(exc))
        finally:
            if history:
                history.close()

    def clear_result_table(self):
        for row in self.table.get_children():
            self.table.delete(row)

    def previous_results_message(self):
        try:
            completed = datetime.fromisoformat(self.last_result["completed_at"])
            cutoff = completed.timestamp() + 1
        except (KeyError, TypeError, ValueError, OSError):
            return "Previous results are shown. Click Check both files to refresh them before creating drafts."
        if str(self.last_result.get("mode", "")).casefold() == "pdf preview":
            return (
                f"Last run ({completed:%d %b, %H:%M}) was a PDF-only preview; it did not create a draft. "
                "Click Create one draft to use the normal certificate and email process."
            )
        if str(self.last_result.get("mode", "")).casefold() == "outlook drafts":
            outcomes = [row.get("outcome") for row in self.last_result.get("rows", []) if isinstance(row, dict)]
            created = outcomes.count("draft_created")
            updated = outcomes.count("updated")
            needs_review = outcomes.count("needs_review")
            if created or updated:
                parts = []
                if created:
                    parts.append(f"{created} Outlook draft{'s' if created != 1 else ''} saved")
                if updated:
                    parts.append(f"{updated} Excel update{'s' if updated != 1 else ''} recovered")
                if needs_review:
                    parts.append(f"{needs_review} row{'s' if needs_review != 1 else ''} need attention")
                return "Last run: " + "; ".join(parts) + ". Review Outlook Drafts and the results below."
            return "The last draft run created no new messages. Review the results below before trying again."
        for variable in self.workbooks:
            path = Path(variable.get())
            try:
                if path.is_file() and path.stat().st_mtime > cutoff:
                    return f"The Excel files changed after the last check ({completed:%d %b, %H:%M}). Click Check both files to refresh these results."
            except OSError:
                continue
        return f"Last check: {completed:%d %b, %H:%M}. Click Check both files to refresh before creating drafts."

    def open_results(self):
        path = self.displayed_results_path
        try:
            if path.exists():
                os.startfile(str(path))
            else:
                self.status.set("Click Check both files first to create the results list.")
        except OSError as exc:
            self.status.set(f"Could not open the results: {exc}")

    def open_output(self):
        path = self.base / "Output"
        try:
            path.mkdir(exist_ok=True)
            os.startfile(str(path))
        except OSError as exc:
            self.status.set(f"Could not open the output folder: {exc}")

    def close(self):
        if self.busy:
            self.close_when_done = True
            self.status.set("The app will close when this run finishes saving its work.")
        else:
            self.root.destroy()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind in {"done", "error", "accounts"}:
                    self.set_busy(False)
                if kind == "accounts":
                    self.accounts["values"] = value
                    if self.mailbox.get() not in value:
                        self.mailbox.set(value[0] if value else "")
                    if self.save_settings():
                        self.status.set(f"Found {len(value)} Outlook mailbox(es).")
                else:
                    self.status.set(value)
                if kind in {"done", "error", "accounts"}:
                    self.show_results()
                    if self.close_when_done:
                        self.root.destroy()
                        return
        except queue.Empty:
            pass
        self.root.after(150, self.poll)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
