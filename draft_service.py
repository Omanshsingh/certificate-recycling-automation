"""Draft-only Outlook integration and durable processing history."""
from __future__ import annotations

import hashlib
import json
import sqlite3
import time
from html import escape
from pathlib import Path

from certificate_core import certificate_number, certificate_sequence


class History:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path)
        self.db.row_factory = sqlite3.Row
        self.db.execute("CREATE TABLE IF NOT EXISTS batches (case_id TEXT PRIMARY KEY, fingerprint TEXT, status TEXT, certificate_no TEXT, pdf_path TEXT, entry_id TEXT, updated_at TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS sent_acknowledgements (case_id TEXT PRIMARY KEY, certificate_no TEXT NOT NULL, sent_at TEXT NOT NULL)")
        if "certificate_date" not in {row[1] for row in self.db.execute("PRAGMA table_info(batches)")}:
            self.db.execute("ALTER TABLE batches ADD COLUMN certificate_date TEXT")
        self.db.commit()

    @staticmethod
    def fingerprint(batch):
        data = {key: batch.get(key) for key in ("case_id", "inward_date", "supplier", "location", "form6", "email", "reference")}
        if batch.get("source_id") not in (None, "", "plant1"):
            data["source_id"] = batch["source_id"]
        data["items"] = batch["items"]
        return hashlib.sha256(json.dumps(data, sort_keys=True, default=str).encode()).hexdigest()

    @staticmethod
    def history_id(batch):
        return batch.get("history_id", batch["case_id"])

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
        history_id = self.history_id(batch)
        result = self.inspect(history_id, fingerprint)
        if result:
            return result
        self.db.execute("INSERT INTO batches(case_id,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at) VALUES (?,?,?,?,?,?,datetime('now'))", (history_id, fingerprint, "pending", certificate_no, "", ""))
        self.db.commit()
        return None

    def complete(self, batch, pdf_path, entry_id):
        self.db.execute("UPDATE batches SET status='saved',pdf_path=?,entry_id=?,updated_at=datetime('now') WHERE case_id=?", (str(pdf_path), entry_id, self.history_id(batch)))
        self.db.commit()

    def record(self, batch):
        row = self.db.execute("SELECT * FROM batches WHERE case_id=?", (self.history_id(batch),)).fetchone()
        return dict(row) if row and row["entry_id"] != "PDF_ONLY" else None

    def has_numbers(self):
        return bool(self.records())

    def records(self):
        return [dict(row) for row in self.db.execute("SELECT * FROM batches WHERE COALESCE(entry_id, '') != 'PDF_ONLY'")]

    def validate_sources(self, paths, batches, bind=False):
        """Keep legacy plant-slot keys tied to the files that own them."""
        identities = [str(Path(path).resolve()).casefold() for path in paths]
        saved = self.db.execute("SELECT value FROM meta WHERE key='source_paths'").fetchone()
        if saved and self.has_numbers() and json.loads(saved[0]) != identities:
            raise ValueError("Saved plant file identities changed or their slots were swapped. Restore the original two paths and plant order before issuing; a moved or replacement register needs an explicit history migration.")
        if not saved and self.has_numbers():
            current = {self.history_id(batch): batch for batch in batches}
            for record in self.records():
                batch = current.get(record["case_id"])
                if not batch or record["fingerprint"] != self.fingerprint(batch):
                    raise ValueError("Existing certificate history cannot be matched to these plant files. Restore the original registers and plant order before issuing.")
        if bind and (not saved or not self.has_numbers()):
            self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('source_paths',?)", (json.dumps(identities),))
            self.db.commit()

    def allocate(self, batch, settings, minimum, issued_date):
        """Reserve the next number and the batch together in one transaction."""
        self.db.execute("BEGIN IMMEDIATE")
        try:
            row = self.record(batch)
            if row:
                if row["fingerprint"] != self.fingerprint(batch):
                    raise ValueError("Source details changed after this certificate was reserved")
                if row["status"] != "allocated":
                    raise ValueError("An existing certificate cannot be allocated again; reconcile its saved history first")
                self.db.commit()
                return row
            saved = self.db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone()
            registered = list(self.db.execute("SELECT certificate_no FROM batches WHERE COALESCE(entry_id, '') != 'PDF_ONLY'"))
            if minimum <= 1 and not registered and not settings.get("initial_cod_confirmed", False):
                raise ValueError("Neither plant file contains an earlier COD number. Enter the approved first COD number once in the app before issuing.")
            sequence = max(int(settings.get("certificate_number_start", 1)), int(minimum), int(saved[0]) if saved else 1)
            for existing in registered:
                sequence = max(sequence, certificate_sequence(existing[0], settings) + 1)
            cert_no = certificate_number(sequence, settings)
            self.db.execute("INSERT OR REPLACE INTO batches(case_id,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date) VALUES (?,?,?,?,?,?,datetime('now'),?)", (self.history_id(batch), self.fingerprint(batch), "allocated", cert_no, "", "", issued_date.isoformat()))
            self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('next_sequence',?)", (str(sequence + 1),))
            self.db.commit()
            return self.record(batch)
        except Exception:
            self.db.rollback()
            raise

    def saving_draft(self, batch, pdf_path):
        changed = self.db.execute("UPDATE batches SET status='saving_draft',pdf_path=?,updated_at=datetime('now') WHERE case_id=? AND status='allocated' AND fingerprint=?", (str(pdf_path), self.history_id(batch), self.fingerprint(batch)))
        if changed.rowcount != 1:
            self.db.rollback()
            raise ValueError("This certificate is already being saved or its source changed; no additional Outlook draft was attempted")
        self.db.commit()

    def draft_saved(self, batch, pdf_path, entry_id):
        if not entry_id or not str(entry_id).strip():
            raise ValueError("Outlook returned no draft ID; save remains unconfirmed")
        self.db.execute("UPDATE batches SET status='draft_saved',pdf_path=?,entry_id=?,updated_at=datetime('now') WHERE case_id=?", (str(pdf_path), entry_id, self.history_id(batch)))
        self.db.commit()

    def reset_pending(self):
        if self.db.execute("SELECT 1 FROM batches WHERE status IN ('pending','saving_draft') LIMIT 1").fetchone():
            raise ValueError("Unconfirmed Outlook saves cannot be reset automatically. Check the saved COD number in the selected mailbox and reconcile the existing history; deleting reservations can create duplicate drafts.")
        return 0

    def next_sequence(self, start, minimum=1):
        self.db.execute("BEGIN IMMEDIATE")
        try:
            saved = self.db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone()
            value = max(int(start), int(minimum), int(saved[0]) if saved else int(start))
            self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('next_sequence',?)", (str(value + 1),))
            self.db.commit()
            return value
        except Exception:
            self.db.rollback()
            raise

    def close(self):
        self.db.close()

    def acknowledge_sent(self, case_id, certificate_no):
        """Record the operator's confirmation, never send or infer delivery."""
        with self.db:
            record = self.db.execute("SELECT status,certificate_no FROM batches WHERE case_id=?", (case_id,)).fetchone()
            if not record or record['status'] != 'saved' or record['certificate_no'] != certificate_no:
                raise ValueError("Only a confirmed draft with completed Excel updates can be marked sent. Check both files first.")
            self.db.execute("INSERT OR IGNORE INTO sent_acknowledgements VALUES(?,?,datetime('now'))", (case_id, certificate_no))

    def sent_acknowledgement(self, case_id):
        row = self.db.execute("SELECT * FROM sent_acknowledgements WHERE case_id=?", (case_id,)).fetchone()
        return dict(row) if row else None

    def undo_sent_acknowledgement(self, case_id):
        with self.db:
            self.db.execute("DELETE FROM sent_acknowledgements WHERE case_id=?", (case_id,))


class OutlookWriter:
    def __init__(self, smtp):
        import win32com.client
        self.app = win32com.client.Dispatch("Outlook.Application")
        accounts = [a for a in self.app.Session.Accounts if str(a.SmtpAddress).strip().lower() == smtp.strip().lower()]
        if len(accounts) != 1:
            raise RuntimeError("The selected Outlook mailbox is not available in Classic Outlook.")
        self.account = accounts[0]
        self.folder = self.account.DeliveryStore.GetDefaultFolder(16)
        self.mailbox = smtp.lower() + "|" + str(self.folder.StoreID)

    def save_draft(self, batch, subject, html_body, pdf_path, banner_path=None):
        for path in (pdf_path, banner_path):
            if path is not None and not Path(path).is_file():
                raise FileNotFoundError(f"Draft attachment unavailable: {path}. Restore it before preparing the draft.")
        message = self.folder.Items.Add(0)
        try:
            # The item is created inside the selected account's Drafts folder,
            # which determines where it is saved. Setting SendUsingAccount is
            # only needed to override the account when sending; Outlook 2013
            # can throw RPC_E_SERVERFAULT (-2147417851) on this setter even
            # for an unsent item. This workflow never sends the message.
            message.BodyFormat = 2
            message.To = batch["email"]
            message.Subject = subject
            message.UserProperties.Add("EcorecoCertificateKey", 1).Value = batch.get("history_id", batch["case_id"])
            message.HTMLBody = html_body
            # Do not pass Position: Outlook uses it for RTF placement and 0 can hide
            # the attachment. Explicitly mark the PDF visible for Outlook's UI.
            certificate = message.Attachments.Add(str(Path(pdf_path).resolve()), 1)
            certificate.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                False,
            )
            if banner_path:
                banner = message.Attachments.Add(str(Path(banner_path).resolve()), 1, 0)
                # PR_ATTACH_CONTENT_ID makes the banner render inline in Outlook.
                banner.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    "ecoreco-banner",
                )
            message.Save()
            entry = str(message.EntryID or "").strip()
            if not entry:
                raise RuntimeError("Outlook returned no draft ID after Save.")
            self.confirm_saved(entry, batch, subject, pdf_path, banner_path)
            return entry
        finally:
            message = None

    def confirm_saved(self, entry, batch, subject, pdf_path, banner_path=None):
        """Read back the saved item from the selected store; never save/send again."""
        expected = {Path(pdf_path).name}
        if banner_path:
            expected.add(Path(banner_path).name)
        last_error = None
        for attempt in range(3):
            saved = None
            try:
                saved = self.app.Session.GetItemFromID(entry, self.folder.StoreID)
                key = saved.UserProperties.Find("EcorecoCertificateKey")
                if not key or str(key.Value) != batch.get("history_id", batch["case_id"]):
                    raise RuntimeError("the saved batch identifier does not match")
                if str(saved.Subject) != subject or bool(saved.Sent):
                    raise RuntimeError("the item is not the expected unsent certificate draft")
                if str(saved.Parent.EntryID) != str(self.folder.EntryID):
                    raise RuntimeError("the item is not in the selected Drafts folder")
                attachments = {str(item.FileName): item for item in saved.Attachments}
                if not expected.issubset(attachments):
                    raise RuntimeError("the saved PDF or signature attachment is missing")
                pdf = attachments[Path(pdf_path).name]
                hidden = bool(pdf.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
                ))
                if hidden:
                    raise RuntimeError("the saved PDF is hidden in Outlook instead of appearing as a visible attachment")
                return
            except Exception as exc:
                last_error = exc
                if attempt < 2:
                    time.sleep(0.2 * (attempt + 1))
            finally:
                saved = None
        raise RuntimeError(f"Outlook save could not be verified by reading the saved draft. Do not create another copy; inspect Outlook first. ({last_error})") from last_error


def draft_html(batch):
    location = f", {escape(batch['location'])}" if batch.get("location") else ""
    rows = "".join(
        f"<tr><td width='250' valign='middle' style='border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;word-wrap:break-word;overflow-wrap:break-word;width:187.5pt'>"
        f"<p class='MsoNormal' style='margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly'>{escape(item['item'])}</p></td>"
        f"<td width='100' valign='middle' style='border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:center;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:75pt'>"
        f"<p class='MsoNormal' align='center' style='margin:0;text-align:center;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly'>{item.get('quantity') or ''}</p></td>"
        f"<td width='190' valign='middle' style='border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:142.5pt'>"
        f"<p class='MsoNormal' align='right' style='margin:0;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly'>{item.get('weight') or 0:.1f} KG</p></td></tr>"
        for item in batch["items"]
    )
    banner = '<img border="0" width="530" height="95" src="cid:ecoreco-banner" alt="EcoReco">'
    return f"""<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><style>
p.MsoNormal,li.MsoNormal,div.MsoNormal {{ margin:0cm; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri",sans-serif; }}
p.wordsection1,li.wordsection1,div.wordsection1 {{ margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif; }}
p.MsoNormal {{ margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif; line-height:1.15; }}
body,table,td,p,div,span,b,strong,a {{ font-size:11.0pt; font-family:"Calibri",sans-serif; }}
td {{ vertical-align:middle; }}
a:link {{ color:#0563C1; text-decoration:underline; }}
</style></head><body lang="EN-US" link="#0563C1" vlink="#954F72" style='font-size:11.0pt;font-family:"Calibri",sans-serif'><div class="WordSection1" style='font-size:11.0pt;font-family:"Calibri",sans-serif'>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt">Hello,<o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt">Greetings from Ecoreco !!<o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><b>Eco Recycling Ltd is pleased to provide the attached Certificate of E-Waste Recycling for the material received from {escape(batch['supplier'].upper())}{location.upper()}.</b><o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt">Please find the certificate details below. The official certificate is attached as a PDF for your records.<o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="540" style="width:405pt;border-collapse:collapse;table-layout:fixed;mso-table-lspace:0pt;mso-table-rspace:0pt">
<tr><td width="250" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:187.5pt"><p class="MsoNormal" style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly"><b>Material / Item</b></p></td><td width="100" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:center;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:75pt"><p class="MsoNormal" align="center" style="margin:0;text-align:center;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly"><b>Quantity</b></p></td><td width="190" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:142.5pt"><p class="MsoNormal" align="right" style="margin:0;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly"><b>Weight</b></p></td></tr>
{rows}
<tr><td colspan="2" width="350" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:262.5pt"><p class="MsoNormal" align="right" style="margin:0;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly"><b>Total Weight</b></p></td><td width="190" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;width:142.5pt"><p class="MsoNormal" align="right" style="margin:0;text-align:right;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly"><b>{batch['total_weight']:.1f} KG</b></p></td></tr>
</table>
<p class="MsoNormal" style="margin:0;font-size:11pt;line-height:10pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt">Kindly review the attached certificate and let us know if any details need correction.<o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt">We look forward to your continued association.<o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:12.0pt"><b><span style="color:#0F243E">Thanks &amp; Regards,</span></b><o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:12.0pt"><b><span style="color:#0F243E">CRM Executive</span></b><span style="color:#0F243E"> | <a href="tel:+912240052951">+91-22-4005 2951</a> / <a href="tel:+919004149714">+91-9004149714</a> | <a href="http://www.ecoreco.com/">www.ecoreco.com</a><br>Eco Recycling Limited | 422, The Summit Business Park | Andheri Kurla road | Andheri (East), Mumbai 400093<o:p></o:p></span></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:12.0pt"><b>{banner}</b><o:p></o:p></p>
<p class="wordsection1" style="margin:0cm;margin-bottom:.0001pt"><o:p>&nbsp;</o:p></p>
</div></body></html>"""
