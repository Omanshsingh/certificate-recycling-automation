"""Draft-only Outlook integration and durable processing history."""
from __future__ import annotations

import hashlib
import json
import sqlite3
import time
import uuid
from html import escape
from pathlib import Path

from certificate_core import certificate_number, certificate_sequence, parse_date


class History:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path)
        self.db.row_factory = sqlite3.Row
        self.db.execute("CREATE TABLE IF NOT EXISTS batches (case_id TEXT PRIMARY KEY, fingerprint TEXT, status TEXT, certificate_no TEXT, pdf_path TEXT, entry_id TEXT, updated_at TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)")
        self.db.execute("CREATE TABLE IF NOT EXISTS sent_acknowledgements (case_id TEXT PRIMARY KEY, certificate_no TEXT NOT NULL, sent_at TEXT NOT NULL)")
        self.db.execute("CREATE TABLE IF NOT EXISTS archived_issuances (case_id TEXT NOT NULL, generation INTEGER NOT NULL, fingerprint TEXT, status TEXT, certificate_no TEXT, pdf_path TEXT, entry_id TEXT, updated_at TEXT, certificate_date TEXT, sent_at TEXT, PRIMARY KEY(case_id,generation))")
        self.db.execute("CREATE TABLE IF NOT EXISTS certificate_corrections (correction_id TEXT PRIMARY KEY, case_id TEXT NOT NULL, generation INTEGER NOT NULL, old_certificate_no TEXT NOT NULL, old_certificate_date TEXT NOT NULL, new_certificate_no TEXT NOT NULL, new_certificate_date TEXT NOT NULL, reason TEXT NOT NULL, old_pdf_path TEXT NOT NULL, new_pdf_path TEXT NOT NULL, entry_id TEXT NOT NULL, stage TEXT NOT NULL, created_at TEXT NOT NULL, completed_at TEXT, observed_excel_certificate_no TEXT, observed_excel_certificate_date TEXT)")
        correction_columns = {row[1] for row in self.db.execute("PRAGMA table_info(certificate_corrections)")}
        for name in ("observed_excel_certificate_no", "observed_excel_certificate_date"):
            if name not in correction_columns:
                self.db.execute(f"ALTER TABLE certificate_corrections ADD COLUMN {name} TEXT")
        if "certificate_date" not in {row[1] for row in self.db.execute("PRAGMA table_info(batches)")}:
            self.db.execute("ALTER TABLE batches ADD COLUMN certificate_date TEXT")
        columns = {row[1] for row in self.db.execute("PRAGMA table_info(batches)")}
        for name, declaration in (("generation", "INTEGER NOT NULL DEFAULT 1"),
                                  ("status_written", "INTEGER NOT NULL DEFAULT 0"),
                                  ("prior_certificate_no", "TEXT"),
                                  ("prior_certificate_date", "TEXT")):
            if name not in columns:
                self.db.execute(f"ALTER TABLE batches ADD COLUMN {name} {declaration}")
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
        changed = self.db.execute(
            "UPDATE batches SET status='saved',status_written=1,prior_certificate_no=NULL,prior_certificate_date=NULL,"
            "pdf_path=?,entry_id=?,updated_at=datetime('now') WHERE case_id=?",
            (str(pdf_path), entry_id, self.history_id(batch)),
        )
        if changed.rowcount != 1:
            self.db.rollback()
            raise ValueError("The confirmed draft could not be matched to its active certificate history")
        self.db.commit()

    def mark_drafted(self, batch, certificate_no):
        changed = self.db.execute(
            "UPDATE batches SET status_written=1,updated_at=datetime('now') "
            "WHERE case_id=? AND status='saved' AND certificate_no=? AND fingerprint=?",
            (self.history_id(batch), certificate_no, self.fingerprint(batch)),
        )
        if changed.rowcount != 1:
            self.db.rollback()
            raise ValueError("The saved certificate history changed before its Drafted status could be recorded")
        self.db.commit()

    # Retain the old method name for callers from an earlier local install.
    def mark_completed(self, batch, certificate_no):
        return self.mark_drafted(batch, certificate_no)

    def record(self, batch):
        row = self.db.execute("SELECT * FROM batches WHERE case_id=?", (self.history_id(batch),)).fetchone()
        return dict(row) if row and row["entry_id"] != "PDF_ONLY" else None

    def has_numbers(self):
        return bool(self.all_records())

    def records(self):
        return [dict(row) for row in self.db.execute("SELECT * FROM batches WHERE COALESCE(entry_id, '') != 'PDF_ONLY'")]

    def all_records(self):
        """Return active and archived issues for number-history checks."""
        active = self.records()
        archived = [dict(row) for row in self.db.execute(
            "SELECT case_id,generation,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,"
            "0 AS status_written,NULL AS prior_certificate_no,NULL AS prior_certificate_date,sent_at "
            "FROM archived_issuances WHERE COALESCE(entry_id, '') != 'PDF_ONLY'"
        )]
        return active + archived

    def correction_for_case(self, case_id):
        row = self.db.execute(
            "SELECT * FROM certificate_corrections WHERE case_id=? AND stage!='completed' "
            "ORDER BY created_at DESC LIMIT 1", (case_id,),
        ).fetchone()
        return dict(row) if row else None

    def begin_correction(self, batch, certificate_no, certificate_date, reason,
                         pdf_path, settings):
        """Reserve an audited correction for an unsent, confirmed certificate."""
        certificate_no = str(certificate_no or "").strip()
        reason = str(reason or "").strip()
        certificate_date = parse_date(certificate_date)
        if not certificate_no or not certificate_date:
            raise ValueError("Enter a valid COD number and certificate date.")
        if not reason or len(reason) > 500:
            raise ValueError("Enter a correction reason (up to 500 characters).")
        number = certificate_sequence(certificate_no, settings)
        case_id = self.history_id(batch)
        self.db.execute("BEGIN IMMEDIATE")
        try:
            pending = self.db.execute(
                "SELECT * FROM certificate_corrections WHERE case_id=? AND stage!='completed' "
                "ORDER BY created_at DESC LIMIT 1", (case_id,),
            ).fetchone()
            if pending:
                raise ValueError("A certificate correction is already in progress. Use Correct certificate to resume that correction first.")
            row = self.db.execute("SELECT * FROM batches WHERE case_id=?", (case_id,)).fetchone()
            if not row or row["status"] != "saved" or not row["status_written"]:
                raise ValueError("Only a saved, unsent Drafted certificate can be corrected here. Check both plant files first.")
            if row["fingerprint"] != self.fingerprint(batch):
                raise ValueError("The client or inward details have changed since the certificate was drafted. Review the row before correcting its COD or date.")
            if self.db.execute("SELECT 1 FROM sent_acknowledgements WHERE case_id=?", (case_id,)).fetchone():
                raise ValueError("This certificate is confirmed Completed. The app will not change it; follow the company's certificate correction process so the original record remains intact.")
            if not row["entry_id"] or row["entry_id"] == "PDF_ONLY" or not row["pdf_path"]:
                raise ValueError("The saved Outlook draft or its PDF could not be identified. Check Outlook Drafts before changing this certificate.")
            if number == certificate_sequence(row["certificate_no"], settings):
                if certificate_no != row["certificate_no"]:
                    raise ValueError("That COD has the same number as the saved COD but a different format. Keep the existing COD exactly as shown, or enter a different unused number.")
            for existing in self.db.execute(
                "SELECT case_id,generation,certificate_no FROM batches WHERE COALESCE(entry_id,'')!='PDF_ONLY' "
                "UNION ALL SELECT case_id,generation,certificate_no FROM archived_issuances WHERE COALESCE(entry_id,'')!='PDF_ONLY'"
            ):
                same_active_issue = existing["case_id"] == case_id and int(existing["generation"] or 0) == int(row["generation"] or 1)
                if not same_active_issue and certificate_sequence(existing["certificate_no"], settings) == number:
                    raise ValueError(f"COD {certificate_no} has already been used. Enter a unique unused COD number.")
            if certificate_no == row["certificate_no"] and certificate_date == parse_date(row["certificate_date"]):
                raise ValueError("The COD number and date are unchanged; enter the value that needs correcting.")
            correction_id = uuid.uuid4().hex
            self.db.execute(
                "INSERT INTO certificate_corrections(correction_id,case_id,generation,old_certificate_no,old_certificate_date,"
                "new_certificate_no,new_certificate_date,reason,old_pdf_path,new_pdf_path,entry_id,stage,created_at,"
                "observed_excel_certificate_no,observed_excel_certificate_date) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,'pending',datetime('now'),?,?)",
                (correction_id, case_id, row["generation"], row["certificate_no"], row["certificate_date"],
                 certificate_no, certificate_date.isoformat(), reason, row["pdf_path"], str(pdf_path), row["entry_id"],
                 str(batch.get("cod_no", "") or ""),
                 parse_date(batch.get("cod_date")).isoformat() if parse_date(batch.get("cod_date")) else None),
            )
            changed = self.db.execute(
                "UPDATE batches SET status='correcting',updated_at=datetime('now') "
                "WHERE case_id=? AND generation=? AND status='saved' AND status_written=1",
                (case_id, row["generation"]),
            )
            if changed.rowcount != 1:
                raise ValueError("Certificate history changed while correction was being started. Check both plant files and try again.")
            self.db.commit()
            return self.correction_for_case(case_id)
        except Exception:
            self.db.rollback()
            raise

    def set_correction_stage(self, correction_id, stage):
        if stage not in {"outlook_updated", "workbook_updated"}:
            raise ValueError("Unknown certificate correction stage.")
        with self.db:
            if stage == "outlook_updated":
                changed = self.db.execute(
                    "UPDATE certificate_corrections SET stage=CASE WHEN stage='workbook_updated' THEN stage ELSE ? END "
                    "WHERE correction_id=? AND stage!='completed'", (stage, correction_id),
                )
            else:
                changed = self.db.execute(
                    "UPDATE certificate_corrections SET stage=? WHERE correction_id=? AND stage!='completed'",
                    (stage, correction_id),
                )
            if changed.rowcount != 1:
                raise ValueError("The certificate correction record could not be updated. Check both the draft and workbook before retrying.")

    def finish_correction(self, correction_id, settings):
        """Commit the audited issue history only after Outlook and Excel are verified."""
        self.db.execute("BEGIN IMMEDIATE")
        try:
            correction = self.db.execute(
                "SELECT * FROM certificate_corrections WHERE correction_id=?", (correction_id,),
            ).fetchone()
            if not correction:
                raise ValueError("The correction record is missing; no certificate history was changed.")
            if correction["stage"] == "completed":
                self.db.commit()
                return dict(correction)
            if correction["stage"] != "workbook_updated":
                raise ValueError("The Outlook draft and Excel update must both be verified before correction history is finalized.")
            row = self.db.execute("SELECT * FROM batches WHERE case_id=?", (correction["case_id"],)).fetchone()
            if (not row or row["status"] != "correcting"
                    or int(row["generation"] or 1) != int(correction["generation"])):
                raise ValueError("Certificate history changed during correction. Ask for review before continuing.")
            self.db.execute(
                "INSERT INTO archived_issuances(case_id,generation,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,sent_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,NULL)",
                (row["case_id"], row["generation"], row["fingerprint"], "saved", correction["old_certificate_no"],
                 correction["old_pdf_path"], correction["entry_id"], row["updated_at"], correction["old_certificate_date"]),
            )
            next_generation = int(row["generation"] or 1) + 1
            changed = self.db.execute(
                "UPDATE batches SET status='saved',certificate_no=?,certificate_date=?,pdf_path=?,entry_id=?,generation=?,"
                "status_written=1,prior_certificate_no=NULL,prior_certificate_date=NULL,updated_at=datetime('now') "
                "WHERE case_id=? AND status='correcting' AND generation=?",
                (correction["new_certificate_no"], correction["new_certificate_date"], correction["new_pdf_path"],
                 correction["entry_id"], next_generation, correction["case_id"], correction["generation"]),
            )
            if changed.rowcount != 1:
                raise ValueError("The corrected COD could not be saved to certificate history.")
            sequence = certificate_sequence(correction["new_certificate_no"], settings) + 1
            current = self.db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone()
            sequence = max(sequence, int(current[0]) if current else 1)
            self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('next_sequence',?)", (str(sequence),))
            self.db.execute(
                "UPDATE certificate_corrections SET stage='completed',completed_at=datetime('now') WHERE correction_id=?",
                (correction_id,),
            )
            self.db.commit()
            return dict(self.db.execute(
                "SELECT * FROM certificate_corrections WHERE correction_id=?", (correction_id,),
            ).fetchone())
        except Exception:
            self.db.rollback()
            raise

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

    def allocate(self, batch, settings, minimum, issued_date, reissue=False,
                 custom_certificate_no=None):
        """Reserve a COD, reusing the batch's existing COD for an explicit re-draft."""
        self.db.execute("BEGIN IMMEDIATE")
        try:
            row = self.record(batch)
            prior_number = prior_date = None
            if row:
                if reissue:
                    if (row["status"] != "saved" or not row["status_written"]
                            or row["fingerprint"] != self.fingerprint(batch)):
                        raise ValueError("A new issue is allowed only after a saved certificate is deliberately changed back to Ready and its source details are unchanged")
                    prior_number = row["certificate_no"]
                    prior_date = row["certificate_date"]
                    sent = self.db.execute("SELECT certificate_no,sent_at FROM sent_acknowledgements WHERE case_id=?", (self.history_id(batch),)).fetchone()
                    self.db.execute(
                        "INSERT INTO archived_issuances(case_id,generation,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,sent_at) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (row["case_id"], row["generation"], row["fingerprint"], row["status"], row["certificate_no"],
                         row["pdf_path"], row["entry_id"], row["updated_at"], row["certificate_date"],
                         sent["sent_at"] if sent and sent["certificate_no"] == row["certificate_no"] else None),
                    )
                    self.db.execute("DELETE FROM sent_acknowledgements WHERE case_id=?", (self.history_id(batch),))
                    self.db.execute("DELETE FROM batches WHERE case_id=?", (self.history_id(batch),))
                    generation = int(self.db.execute(
                        "SELECT COALESCE(MAX(generation),0)+1 FROM archived_issuances WHERE case_id=?",
                        (self.history_id(batch),),
                    ).fetchone()[0])
                    self.db.execute(
                        "INSERT INTO batches(case_id,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,generation,status_written,prior_certificate_no,prior_certificate_date) "
                        "VALUES (?,?,?,?,?,?,datetime('now'),?,?,0,NULL,NULL)",
                        (self.history_id(batch), self.fingerprint(batch), "allocated", prior_number, "", "", prior_date, generation),
                    )
                    self.db.commit()
                    return self.record(batch)
                else:
                    if row["fingerprint"] != self.fingerprint(batch):
                        raise ValueError("Source details changed after this certificate was reserved")
                    if row["status"] != "allocated":
                        raise ValueError("An existing certificate cannot be allocated again; reconcile its saved history first")
                    self.db.commit()
                    return row
            elif reissue:
                raise ValueError("The previous certificate history is missing; review it before requesting a new issue")

            saved = self.db.execute("SELECT value FROM meta WHERE key='next_sequence'").fetchone()
            registered = list(self.db.execute(
                "SELECT certificate_no FROM batches WHERE COALESCE(entry_id, '') != 'PDF_ONLY' "
                "UNION ALL SELECT certificate_no FROM archived_issuances WHERE COALESCE(entry_id, '') != 'PDF_ONLY'"
            ))
            if minimum <= 1 and not registered and not settings.get("initial_cod_confirmed", False):
                raise ValueError("Neither plant file contains an earlier COD number. Enter the approved first COD number once in the app before issuing.")
            sequence = max(int(settings.get("certificate_number_start", 1)), int(minimum), int(saved[0]) if saved else 1)
            for existing in registered:
                sequence = max(sequence, certificate_sequence(existing[0], settings) + 1)
            if custom_certificate_no:
                custom_sequence = certificate_sequence(custom_certificate_no, settings)
                owner = self.history_id(batch)
                for existing in self.db.execute(
                    "SELECT case_id,certificate_no FROM batches WHERE COALESCE(entry_id, '') != 'PDF_ONLY' "
                    "UNION ALL SELECT case_id,certificate_no FROM archived_issuances WHERE COALESCE(entry_id, '') != 'PDF_ONLY'"
                ):
                    if (existing[0] != owner
                            and certificate_sequence(existing[1], settings) == custom_sequence):
                        raise ValueError(f"COD {custom_certificate_no} is already assigned to another inward. Leave COD No blank to use the next number, or enter a unique COD.")
                cert_no = str(custom_certificate_no).strip()
                sequence = max(sequence, custom_sequence + 1)
                next_sequence = sequence
            else:
                cert_no = certificate_number(sequence, settings)
                next_sequence = sequence + 1
            generation = int(self.db.execute(
                "SELECT COALESCE(MAX(generation),0)+1 FROM archived_issuances WHERE case_id=?",
                (self.history_id(batch),),
            ).fetchone()[0])
            self.db.execute(
                "INSERT INTO batches(case_id,fingerprint,status,certificate_no,pdf_path,entry_id,updated_at,certificate_date,generation,status_written,prior_certificate_no,prior_certificate_date) "
                "VALUES (?,?,?,?,?,?,datetime('now'),?,?,0,?,?)",
                (self.history_id(batch), self.fingerprint(batch), "allocated", cert_no, "", "", issued_date.isoformat(), generation, prior_number, prior_date),
            )
            self.db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('next_sequence',?)", (str(next_sequence),))
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
                raise ValueError("Only a confirmed Drafted certificate with completed Excel updates can be marked Completed. Check both files first.")
            self.db.execute("INSERT OR IGNORE INTO sent_acknowledgements VALUES(?,?,datetime('now'))", (case_id, certificate_no))

    def acknowledge_completed(self, case_id, certificate_no):
        """Record that the operator has sent the draft and completed the case."""
        return self.acknowledge_sent(case_id, certificate_no)

    def sent_acknowledgement(self, case_id):
        row = self.db.execute("SELECT * FROM sent_acknowledgements WHERE case_id=?", (case_id,)).fetchone()
        return dict(row) if row else None

    def completed_acknowledgement(self, case_id):
        return self.sent_acknowledgement(case_id)

    def undo_sent_acknowledgement(self, case_id):
        with self.db:
            self.db.execute("DELETE FROM sent_acknowledgements WHERE case_id=?", (case_id,))

    def undo_completed_acknowledgement(self, case_id):
        return self.undo_sent_acknowledgement(case_id)


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

    def validate_certificate_draft(self, entry_id, batch, expected_subject, pdf_path):
        """Read-only preflight so a missing/edited draft fails before correction is reserved."""
        if not Path(pdf_path).is_file():
            raise FileNotFoundError(f"The original certificate PDF is missing: {pdf_path}")
        message = None
        try:
            message = self.app.Session.GetItemFromID(entry_id, self.folder.StoreID)
            key = message.UserProperties.Find("EcorecoCertificateKey")
            if not key or str(key.Value) != batch.get("history_id", batch["case_id"]):
                raise RuntimeError("the selected draft belongs to a different inward")
            if bool(message.Sent):
                raise RuntimeError("the certificate email is already sent; sent certificates cannot be changed here")
            if str(message.Parent.EntryID) != str(self.folder.EntryID):
                raise RuntimeError("the saved message is no longer in the selected Drafts folder")
            if str(message.Subject) != expected_subject:
                raise RuntimeError("the draft subject was changed outside this app; review it manually before correcting the certificate")
            names = {str(message.Attachments.Item(i).FileName)
                     for i in range(1, int(message.Attachments.Count) + 1)}
            if Path(pdf_path).name not in names:
                raise RuntimeError("the original certificate PDF is missing from the Outlook draft")
        finally:
            message = None

    def update_certificate_draft(self, entry_id, batch, old_subject, new_subject,
                                 old_pdf_path, new_pdf_path, banner_path=None):
        """Correct only the COD subject and certificate PDF of an existing unsent draft."""
        for path in (old_pdf_path, new_pdf_path, banner_path):
            if path is not None and not Path(path).is_file():
                raise FileNotFoundError(f"Certificate draft file unavailable: {path}. Restore it before continuing.")
        message = None
        try:
            message = self.app.Session.GetItemFromID(entry_id, self.folder.StoreID)
            key = message.UserProperties.Find("EcorecoCertificateKey")
            expected_key = batch.get("history_id", batch["case_id"])
            if not key or str(key.Value) != expected_key:
                raise RuntimeError("the selected draft belongs to a different inward")
            if bool(message.Sent):
                raise RuntimeError("the certificate email is already sent; sent certificates cannot be changed here")
            if str(message.Parent.EntryID) != str(self.folder.EntryID):
                raise RuntimeError("the saved message is no longer in the selected Drafts folder")
            if str(message.Subject) not in {old_subject, new_subject}:
                raise RuntimeError("the draft subject was changed outside this app; review it manually before correcting the certificate")

            old_name, new_name = Path(old_pdf_path).name, Path(new_pdf_path).name
            attachments = {}
            for index in range(1, int(message.Attachments.Count) + 1):
                attachment = message.Attachments.Item(index)
                attachments.setdefault(str(attachment.FileName), []).append(index)
            if new_name not in attachments and old_name not in attachments:
                raise RuntimeError("the existing certificate PDF is missing from the Outlook draft; check the draft before correcting it")
            if new_name not in attachments:
                attachment = message.Attachments.Add(str(Path(new_pdf_path).resolve()), 1)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", False,
                )
            if banner_path and Path(banner_path).name not in attachments:
                banner = message.Attachments.Add(str(Path(banner_path).resolve()), 1, 0)
                banner.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ecoreco-banner",
                )
            # Remove only the previous certificate PDF; preserve the banner and
            # any other attachments. Remove in descending index order so Outlook
            # does not shift a later attachment into an earlier index.
            remove_indices = list(attachments.get(old_name, []))
            if new_name in attachments and old_name != new_name:
                # Keep one corrected certificate attachment if a previous
                # interrupted attempt or manual action left duplicate copies.
                remove_indices.extend(attachments[new_name][1:])
            for index in sorted(set(remove_indices), reverse=True):
                if old_name == new_name:
                    continue
                message.Attachments.Item(index).Delete()
            message.Subject = new_subject
            message.Save()
            self.confirm_saved(entry_id, batch, new_subject, new_pdf_path, banner_path)
        except Exception as exc:
            if isinstance(exc, (FileNotFoundError, RuntimeError)):
                raise
            raise RuntimeError(f"Outlook could not safely update the existing unsent draft. Check Outlook before retrying. ({exc})") from exc
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
