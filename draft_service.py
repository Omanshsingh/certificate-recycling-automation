"""Draft-only Outlook integration and durable processing history."""
from __future__ import annotations

import hashlib
import json
import sqlite3
from pathlib import Path


class History:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path)
        self.db.execute("CREATE TABLE IF NOT EXISTS batches (case_id TEXT PRIMARY KEY, fingerprint TEXT, status TEXT, certificate_no TEXT, pdf_path TEXT, entry_id TEXT, updated_at TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
        self.db.commit()

    @staticmethod
    def fingerprint(batch):
        data = {key: batch.get(key) for key in ("case_id", "inward_date", "supplier", "location", "form6", "email", "reference")}
        data["items"] = batch["items"]
        return hashlib.sha256(json.dumps(data, sort_keys=True, default=str).encode()).hexdigest()

    def inspect(self, case_id, fingerprint):
        row = self.db.execute("SELECT fingerprint,status,certificate_no,pdf_path,entry_id FROM batches WHERE case_id=?", (case_id,)).fetchone()
        if not row:
            return None
        if row[0] != fingerprint:
            return "Needs review: source data changed since the earlier run"
        if row[1] == "saved":
            return "Already prepared"
        if row[1] == "pending":
            return "Needs review: earlier save was interrupted; check Outlook before retrying"
        return row[1]

    def reserve(self, batch, certificate_no):
        fingerprint = self.fingerprint(batch)
        result = self.inspect(batch["case_id"], fingerprint)
        if result:
            return result
        self.db.execute("INSERT INTO batches VALUES (?,?,?,?,?,?,datetime('now'))", (batch["case_id"], fingerprint, "pending", certificate_no, "", ""))
        self.db.commit()
        return None

    def complete(self, batch, pdf_path, entry_id):
        self.db.execute("UPDATE batches SET status='saved',pdf_path=?,entry_id=?,updated_at=datetime('now') WHERE case_id=?", (str(pdf_path), entry_id, batch["case_id"]))
        self.db.commit()

    def reset_pending(self):
        cur = self.db.execute("DELETE FROM batches WHERE status='pending'")
        self.db.commit()
        return cur.rowcount

    def next_sequence(self, start):
        value = self.db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone()
        if value is None:
            value = int(start)
        else:
            value = int(value[0])
        self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('next_sequence',?)", (str(value + 1),))
        self.db.commit()
        return value

    def close(self):
        self.db.close()


class OutlookWriter:
    def __init__(self, smtp):
        import win32com.client
        self.app = win32com.client.Dispatch("Outlook.Application")
        accounts = [a for a in self.app.Session.Accounts if str(a.SmtpAddress).lower() == smtp.lower()]
        if len(accounts) != 1:
            raise RuntimeError("The selected Outlook mailbox is not available in Classic Outlook.")
        self.account = accounts[0]
        self.folder = self.account.DeliveryStore.GetDefaultFolder(16)
        self.mailbox = smtp.lower() + "|" + str(self.folder.StoreID)

    def save_draft(self, batch, subject, html_body, pdf_path, banner_path=None):
        message = self.folder.Items.Add(0)
        try:
            try:
                message.SendUsingAccount = self.account
            except Exception:
                pass
            message.BodyFormat = 2
            message.To = batch["email"]
            message.Subject = subject
            message.HTMLBody = html_body
            message.Attachments.Add(str(Path(pdf_path).resolve()), 1, 0)
            if banner_path:
                banner = message.Attachments.Add(str(Path(banner_path).resolve()), 1, 0)
                # PR_ATTACH_CONTENT_ID makes the banner render inline in Outlook.
                banner.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    "ecoreco-banner",
                )
            message.Save()
            entry = str(message.EntryID)
            if not entry:
                raise RuntimeError("Outlook returned no draft ID after Save.")
            return entry
        finally:
            message = None


def draft_html(batch):
    location = f", {batch['location']}" if batch.get("location") else ""
    rows = "".join(
        f"<tr><td style='border:1px solid #777;padding:8px'>{item['item']}</td>"
        f"<td style='border:1px solid #777;padding:8px;text-align:center'>{item.get('quantity') or ''}</td>"
        f"<td style='border:1px solid #777;padding:8px;text-align:right'>{item.get('weight') or 0:.1f} KG</td></tr>"
        for item in batch["items"]
    )
    return f"""<html><body style='margin:0;background:#ffffff;color:#222;font-family:Arial,sans-serif;font-size:11pt'>
    <div style='max-width:760px;margin:0 auto;padding:12px 18px'>
      <p>Hello,</p>
      <p>Greetings from Ecoreco !!</p>
      <p><b>Eco Recycling Ltd is pleased to provide the attached Certificate of E-Waste Recycling for the material received from {batch['supplier'].upper()}{location.upper()}.</b></p>
      <p>Please find the certificate details below. The official certificate is attached as a PDF for your records.</p>
      <table cellpadding='0' cellspacing='0' style='border-collapse:collapse;width:100%;margin:18px 0'>
        <tr style='background:#f2f2f2'><th style='border:1px solid #777;padding:8px;text-align:left'>Material / Item</th><th style='border:1px solid #777;padding:8px'>Quantity</th><th style='border:1px solid #777;padding:8px;text-align:right'>Weight</th></tr>
        {rows}
        <tr><td colspan='2' style='border:1px solid #777;padding:8px;text-align:right'><b>Total Weight</b></td><td style='border:1px solid #777;padding:8px;text-align:right'><b>{batch['total_weight']:.1f} KG</b></td></tr>
      </table>
      <p>Kindly review the attached certificate and let us know if any details need correction.</p>
      <p>We look forward to your continued association.</p>
      <p><b>Thanks &amp; Regards,</b><br><br>
      <b>CRM Executive</b> | <a href='tel:+912240052951'>+91-22-4005 2951</a> / <a href='tel:+919004149714'>+91-9004149714</a> | <a href='https://www.ecoreco.com'>www.ecoreco.com</a><br>
      <b>Eco Recycling Limited</b><br>
      422, The Summit Business Park | Andheri Kurla Road | Andheri (East), Mumbai 400093</p>
      <p style='margin-top:18px'><img src='cid:ecoreco-banner' alt='EcoReco' style='display:block;width:100%;max-width:760px;height:auto;border:0'></p>
    </div>
    </body></html>"""
