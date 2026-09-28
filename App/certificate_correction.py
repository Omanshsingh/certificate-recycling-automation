"""Explicit, audited correction of an unsent certificate and its saved draft."""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import uuid

from app_paths import resource
from certificate_core import make_certificate, parse_date, status_value
from draft_service import History, OutlookWriter
from run_certificates import _custom_cod_issue, _read_sources
from workflow_support import load_settings, run_lock
from workbook_store import WorkbookSession


def _safe_name(value):
    return re.sub(r"[^A-Za-z0-9_-]", "_", str(value))


def _locked_batch(session, source_id, case_id, sheet_name):
    batches = list(_read_sources([session.path], data={str(session.path): session}, report=None))
    # _read_sources labels its one-file input plant1; restore the selected slot.
    matches = [batch for batch in batches
               if batch["case_id"] == case_id and batch["source_sheet"] == sheet_name]
    if len(matches) != 1:
        raise ValueError("The selected inward row changed or could not be found. Click Check both files and select it again.")
    batch = matches[0]
    batch["source_id"] = source_id
    batch["source_path"] = str(session.path)
    if source_id == "plant2":
        batch["history_id"] = f"plant2|{case_id}"
    return batch


def correct_certificate(workbook_paths, mailbox, root, source_id, case_id,
                        certificate_no, certificate_date, reason,
                        progress=lambda _: None):
    """Update one unsent Outlook draft, its PDF, the workbook and audit history.

    Sent certificates are deliberately excluded. If interrupted, a durable
    correction record blocks ordinary draft processing until this action resumes.
    """
    root = Path(root)
    paths = [Path(path).resolve() for path in workbook_paths]
    if len(paths) != 2 or source_id not in {"plant1", "plant2"}:
        raise ValueError("Choose both saved plant files and select a case from one of them.")
    plant_index = 0 if source_id == "plant1" else 1
    path = paths[plant_index]
    if not path.is_file():
        raise FileNotFoundError(f"The selected {source_id} workbook is unavailable: {path}")

    target_date = parse_date(certificate_date)
    target_no = str(certificate_no or "").strip()
    reason = str(reason or "").strip()
    if not target_no or not target_date:
        raise ValueError("Enter a valid COD number and date. Use DD.MM.YYYY for the date.")
    if not reason or len(reason) > 500:
        raise ValueError("Enter a correction reason (up to 500 characters).")

    with run_lock(root):
        settings = load_settings(root)
        history = History(root / "Data" / "certificate-history.sqlite")
        try:
            batches = _read_sources(paths)
            history.validate_sources(paths, batches)
            matches = [batch for batch in batches
                       if batch["source_id"] == source_id and batch["case_id"] == str(case_id)]
            if len(matches) != 1:
                raise ValueError("The selected inward row changed or could not be found. Click Check both files and select it again.")
            batch = matches[0]
            record = history.record(batch)
            pending = history.correction_for_case(history.history_id(batch))
            if history.completed_acknowledgement(history.history_id(batch)):
                raise ValueError("This certificate is confirmed Completed. The app will not change a certificate already sent to a client.")

            if pending:
                if not record or record.get("status") != "correcting":
                    raise ValueError("A correction record is incomplete or no longer matches certificate history. Ask an administrator to review it before proceeding.")
                if (pending["new_certificate_no"] != target_no
                        or parse_date(pending["new_certificate_date"]) != target_date
                        or pending["reason"] != reason):
                    raise ValueError(
                        f"A correction to COD {pending['new_certificate_no']} is already in progress. "
                        "Resume that exact correction first; do not start a different one."
                    )
                old_no = pending["old_certificate_no"]
                old_date = parse_date(pending["old_certificate_date"])
                correction = pending
            else:
                if not record or record["status"] != "saved" or not record["status_written"]:
                    raise ValueError("Only a saved, unsent Drafted certificate can be corrected here. Click Check both files first.")
                if status_value(batch["issued"]) not in {"drafted", "completed"}:
                    raise ValueError("This case is not marked Drafted in Excel. Click Check both files before correcting it.")
                if history.fingerprint(batch) != record["fingerprint"]:
                    raise ValueError("The client or inward details changed after drafting. Review the client row before correcting the certificate.")
                old_no = record["certificate_no"]
                old_date = parse_date(record["certificate_date"])
                # Validate against both plant registers and all issued history;
                # replace the selected batch in the validation set so its current
                # COD is not mistaken for a duplicate of itself.
                validation_batch = dict(batch, cod_no=target_no)
                validation_batches = [validation_batch if item is batch else item for item in batches]
                problem = _custom_cod_issue(validation_batch, validation_batches, history, settings)
                if problem:
                    raise ValueError(problem)
                correction_id = uuid.uuid4().hex
                safe_id = _safe_name(case_id)
                safe_no = _safe_name(target_no)
                new_pdf = Path(record["pdf_path"]).with_name(
                    f"Certificate_{source_id}_{safe_id}_{safe_no}_corrected_{correction_id[:8]}.pdf"
                )
                correction = None

            progress(f"{source_id} / SN {case_id}: checking the saved certificate")
            writer = OutlookWriter(mailbox)
            backup_dir = root / "Data" / "Workbook Backups"
            with WorkbookSession(path, backup_dir) as session:
                current = _locked_batch(session, source_id, str(case_id), batch["source_sheet"])
                if history.fingerprint(current) != record["fingerprint"]:
                    raise ValueError("The client or inward details changed while the workbook was being opened. No certificate was changed.")
                if status_value(current["issued"]) not in {"drafted", "completed"}:
                    raise ValueError("The Excel status is no longer Drafted. No certificate was changed.")
                current_pair = (current["cod_no"], parse_date(current["cod_date"]))
                # A changed/cleared Excel COD or date is exactly why an
                # operator may need this correction path. The workbook is
                # locked and the row's non-COD source fingerprint is checked.
                # The value observed here is saved in the correction event
                # and will be required again if the operation is resumed.
                if correction is not None:
                    observed_pair = (
                        correction.get("observed_excel_certificate_no") or "",
                        parse_date(correction.get("observed_excel_certificate_date")),
                    )
                    new_pair = (correction["new_certificate_no"], parse_date(correction["new_certificate_date"]))
                    if current_pair not in {observed_pair, new_pair}:
                        raise ValueError("The Excel COD/date changed during an earlier correction. Restore the original or corrected values before resuming; no new changes were made.")

                if correction is None:
                    if not Path(record["pdf_path"]).is_file():
                        raise FileNotFoundError("The original certificate PDF is missing. Restore it from the Output folder before correcting the draft.")
                    if not record["entry_id"] or record["entry_id"] == "PDF_ONLY":
                        raise ValueError("The saved Outlook draft could not be identified. Check Outlook Drafts before correcting this certificate.")
                    old_subject = f"Certificate of E-Waste Recycling - {current['supplier']} - {old_no}"
                    writer.validate_certificate_draft(
                        record["entry_id"], current, old_subject, record["pdf_path"],
                    )
                    new_pdf.parent.mkdir(parents=True, exist_ok=True)
                    make_certificate(
                        current, target_no, target_date,
                        resource(root, "templates") / "Certificate of Recycling.png",
                        new_pdf,
                    )
                    try:
                        correction = history.begin_correction(
                            current, target_no, target_date, reason, new_pdf, settings,
                        )
                    except Exception:
                        new_pdf.unlink(missing_ok=True)
                        raise
                    old_no = correction["old_certificate_no"]
                    old_date = parse_date(correction["old_certificate_date"])

                # The path and values are persisted before external changes so a
                # later click can safely resume after a crash or Office error.
                new_pdf = Path(correction["new_pdf_path"])
                if not new_pdf.is_file():
                    new_pdf.parent.mkdir(parents=True, exist_ok=True)
                    make_certificate(
                        current, target_no, target_date,
                        resource(root, "templates") / "Certificate of Recycling.png",
                        new_pdf,
                    )
                old_subject = f"Certificate of E-Waste Recycling - {current['supplier']} - {old_no}"
                new_subject = f"Certificate of E-Waste Recycling - {current['supplier']} - {target_no}"
                progress(f"{source_id} / SN {case_id}: updating the existing unsent Outlook draft")
                writer.update_certificate_draft(
                    correction["entry_id"], current, old_subject, new_subject,
                    correction["old_pdf_path"], new_pdf,
                    resource(root, "templates") / "signature-banner.png",
                )
                history.set_correction_stage(correction["correction_id"], "outlook_updated")

                current_pair = (current["cod_no"], parse_date(current["cod_date"]))
                old_pair = (correction.get("observed_excel_certificate_no") or "",
                            parse_date(correction.get("observed_excel_certificate_date")))
                new_pair = (target_no, target_date)
                if current_pair not in {old_pair, new_pair}:
                    raise ValueError("Excel COD/date changed during correction. The Outlook draft was updated, but the correction is safely paused; restore the recorded old or new values and resume it.")
                if current_pair == old_pair:
                    progress(f"{source_id} / SN {case_id}: saving the corrected COD and date in Excel")
                    session.commit(
                        current, target_no, target_date,
                        expected_previous_no=old_no,
                        expected_previous_date=old_date,
                        allow_history_restore=True,
                        target_status="Drafted",
                    )
                history.set_correction_stage(correction["correction_id"], "workbook_updated")
                history.finish_correction(correction["correction_id"], settings)
            return {"certificate_no": target_no, "certificate_date": target_date.isoformat(),
                    "reason": reason, "pdf_path": str(new_pdf)}
        finally:
            history.close()
