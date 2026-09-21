"""Simple desktop screen for certificate generation and Outlook drafts."""
from __future__ import annotations

import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from certificate_core import read_batches
from draft_service import History
from run_certificates import run

BASE = Path(__file__).resolve().parent


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate of Recycling Automation")
        self.root.geometry("720x430")
        self.events = queue.Queue()
        self.workbook = tk.StringVar()
        self.mailbox = tk.StringVar()
        self.status = tk.StringVar(value="Choose the copied inward workbook.")
        settings_path = BASE / "settings.local.json"
        if not settings_path.exists(): settings_path = BASE / "settings.example.json"
        if settings_path.exists():
            settings = json.loads(settings_path.read_text(encoding="utf-8"))
            saved_workbook = settings.get("workbook_path", "")
            if saved_workbook and Path(saved_workbook).exists(): self.workbook.set(saved_workbook)
            self.mailbox.set(settings.get("outlook_mailbox", ""))
        self.build()
        self.root.after(100, self.poll)

    def build(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Certificate of Recycling Automation", font=("Arial", 16, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Draft-only mode: certificates are created after 30 days and saved as Outlook drafts. Nothing is sent automatically.").pack(anchor="w", pady=(4, 18))
        row = ttk.Frame(frame); row.pack(fill="x", pady=5)
        ttk.Label(row, text="Excel workbook", width=18).pack(side="left")
        ttk.Entry(row, textvariable=self.workbook).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Choose...", command=self.choose_workbook).pack(side="left", padx=(8, 0))
        row = ttk.Frame(frame); row.pack(fill="x", pady=5)
        ttk.Label(row, text="Outlook mailbox", width=18).pack(side="left")
        self.accounts = ttk.Combobox(row, textvariable=self.mailbox, state="readonly")
        self.accounts.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Load Outlook accounts", command=self.load_accounts).pack(side="left", padx=(8, 0))
        buttons = ttk.Frame(frame); buttons.pack(fill="x", pady=18)
        ttk.Button(buttons, text="Test 1 certificate PDF only", command=lambda: self.start(False, 1)).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Test 1 Outlook draft", command=lambda: self.start(True, 1)).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Prepare all eligible drafts", command=lambda: self.start(True, None)).pack(side="left")
        buttons2 = ttk.Frame(frame); buttons2.pack(fill="x", pady=4)
        ttk.Button(buttons2, text="Reset interrupted cases", command=self.reset_pending).pack(side="left")
        ttk.Button(buttons2, text="Open output folder", command=self.open_output).pack(side="left", padx=8)
        ttk.Label(frame, textvariable=self.status, wraplength=670).pack(anchor="w", pady=18)

    def choose_workbook(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx *.xlsm")])
        if path:
            self.workbook.set(path)
            self.save_settings()

    def load_accounts(self):
        try:
            import win32com.client
            app = win32com.client.Dispatch("Outlook.Application")
            values = [str(a.SmtpAddress) for a in app.Session.Accounts]
            self.accounts["values"] = values
            if values: self.mailbox.set(values[0])
            self.status.set(f"Found {len(values)} Outlook mailbox(es).")
        except Exception as exc:
            messagebox.showerror("Outlook", f"Could not read Classic Outlook accounts:\n{exc}")

    def start(self, drafts, limit):
        if not self.workbook.get():
            messagebox.showwarning("Workbook required", "Choose the copied inward Excel workbook first.")
            return
        if drafts and not self.mailbox.get():
            messagebox.showwarning("Mailbox required", "Load Outlook accounts and select the mailbox first.")
            return
        self.save_settings()
        self.status.set("Processing... please wait. Outlook drafts are saved one at a time.")
        threading.Thread(target=self.worker, args=(drafts, limit), daemon=True).start()

    def worker(self, drafts, limit):
        try:
            if drafts:
                import pythoncom
                pythoncom.CoInitialize()
            created, review, output = run(self.workbook.get(), self.mailbox.get(), BASE, limit=limit, create_drafts=drafts, progress=lambda msg: self.events.put(("status", msg)))
            self.events.put(("done", f"Finished. {created} item(s) created. {review} item(s) need review. Output: {output}"))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def save_settings(self):
        path = BASE / "settings.local.json"
        source = BASE / "settings.example.json"
        settings = json.loads(source.read_text(encoding="utf-8")) if source.exists() else {}
        if path.exists(): settings.update(json.loads(path.read_text(encoding="utf-8")))
        settings["workbook_path"] = self.workbook.get()
        settings["outlook_mailbox"] = self.mailbox.get()
        path.write_text(json.dumps(settings, indent=2), encoding="utf-8")

    def reset_pending(self):
        path = BASE / "Data" / "certificate-history.sqlite"
        if not path.exists():
            self.status.set("No interrupted cases found."); return
        if messagebox.askyesno("Confirm reset", "Only reset after checking Outlook Drafts for duplicates. Continue?"):
            history = History(path); count = history.reset_pending(); history.close()
            self.status.set(f"Reset {count} interrupted case(s).")

    def open_output(self):
        import os
        (BASE / "Output").mkdir(exist_ok=True)
        os.startfile(str(BASE / "Output"))

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                self.status.set(value)
                if kind == "done": messagebox.showinfo("Complete", value)
                if kind == "error": messagebox.showerror("Processing stopped", value)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)


if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
