"""Read both plants, prepare ready certificates and reconcile Excel updates."""
from __future__ import annotations

from contextlib import ExitStack
from datetime import date
from io import BytesIO
from pathlib import Path
import re

from certificate_core import (batch_content, certificate_layout, certificate_number, certificate_sequence,
                              effective_certificate_status, make_certificate, parse_date, read_batches,
                              status_value, supplier_lines)
from draft_service import History, OutlookWriter, draft_html
from workbook_store import WorkbookSession
from workflow_support import Report, load_settings, run_lock
from app_paths import resource


def _read_sources(paths, data=None, report=None):
    batches = []
    for index, path in enumerate(paths, 1):
        source_id = f"plant{index}"
        try:
            source = BytesIO(data[str(path)].data) if data else path
            source_batches = list(read_batches(source))
        except Exception as exc:
            if report is None:
                raise
            report.add({"source_id": source_id, "source_path": str(path)}, "needs_review",
                       f"Saved workbook unavailable or unreadable: {path}. {exc}. Reconnect its drive, save and close Excel, or restore the register.",
                       issue_code="source_unavailable")
            continue
        if report is not None:
            report.observed_sources.append(source_id)
        for batch in source_batches:
            batch.update(source_id=source_id, source_path=str(path))
            batch["history_id"] = batch["case_id"] if index == 1 else f"{source_id}|{batch['case_id']}"
            batches.append(batch)
    if report is not None:
        report.scan_complete = len(report.observed_sources) == 2
    return batches


def _snapshot(batches):
    return [(b["source_id"], b["source_sheet"], b["source_row"], batch_content(b), b["issued"], b["cod_no"], b["cod_date"], b["missing_headers"], b["issues"]) for b in batches]


def _number_floor(batches, settings, history=None):
    highest = 0
    for batch in batches:
        if not batch["cod_no"]:
            continue
        try:
            highest = max(highest, certificate_sequence(batch["cod_no"], settings))
        except ValueError as exc:
            raise ValueError(f"COD No {batch['cod_no']!r} does not match the configured number format. Use a COD in that format or clear it if it is not an issued certificate.") from exc
    if history is not None:
        identities = [history.history_id(batch) for batch in batches]
        if len(set(identities)) != len(identities):
            raise ValueError("Two source rows share a certificate history key. Correct duplicate SN values or a plant2| prefix in the first plant's SN before issuing.")
        for record in history.all_records():
            number = certificate_sequence(record["certificate_no"], settings)
            highest = max(highest, number)
    return max(highest + 1, int(settings.get("certificate_number_start", 1)))


def _custom_cod_issue(batch, batches, history, settings):
    """Validate a user-entered COD for a first issue without stopping other rows."""
    value = batch.get("cod_no", "").strip()
    if not value:
        return None
    try:
        number = certificate_sequence(value, settings)
    except ValueError:
        return f"COD No {value!r} must match the configured number format. Leave it blank to assign the next COD automatically, or enter a unique number in that format."

    owner = history.history_id(batch)
    for other in batches:
        if other is batch or not other.get("cod_no"):
            continue
        try:
            other_number = certificate_sequence(other["cod_no"], settings)
        except ValueError:
            continue
        if other_number == number:
            return f"COD {value} is also entered for {other['source_id']} SN {other['case_id']}. Leave COD No blank to use the next number, or enter a unique COD."
    for record in history.all_records():
        if record["case_id"] == owner:
            continue
        if certificate_sequence(record["certificate_no"], settings) == number:
            return f"COD {value} is already assigned to another inward. Leave COD No blank to use the next number, or enter a unique COD."
    return None


def _recovery_detail(record):
    return (f"Earlier Outlook save is unconfirmed for COD {record['certificate_no']} "
            f"(history key {record['case_id']}; PDF {record.get('pdf_path') or 'not recorded'}). "
            "Search the selected mailbox for this COD, including Drafts, Sent Items and Deleted Items. "
            "Retain the reservation and reconcile the existing draft with certificate history; "
            "automatic re-drafting is blocked.")


def run(workbook_path, mailbox, root, limit=None, create_drafts=True, as_of=None, progress=print, check_only=False, publish_latest=True, history_path=None):
    root = Path(root)
    settings = load_settings(root)
    as_of = as_of or date.today()
    mode = "Check sheets" if check_only else ("Outlook drafts" if create_drafts else "PDF preview")
    with run_lock(root):
        report = Report(root, as_of, mode, publish_latest=publish_latest)
        history = None
        writer = None
        candidates = []
        reported_candidates = set()
        try:
            paths = [workbook_path] if isinstance(workbook_path, (str, Path)) else list(workbook_path)
            if len(paths) != 2 or not all(str(p).strip() for p in paths):
                raise ValueError("Select the two plant Excel files once in the app.")
            paths = [Path(p).resolve() for p in paths]
            if len({str(p).casefold() for p in paths}) != 2:
                raise ValueError("Choose a different Excel file for each plant.")
            if paths[0].exists() and paths[1].exists() and paths[0].samefile(paths[1]):
                raise ValueError("Choose two separate Excel files, not aliases of the same file.")
            history_path = history_path or (root / "Data" / "certificate-history.sqlite")
            history = History(history_path)
            report.observed_sources = []
            report.scan_complete = False
            batches = _read_sources(paths, report=report)
            blocked = not report.scan_complete
            try:
                if report.scan_complete:
                    history.validate_sources(paths, batches)
                minimum = _number_floor(batches, settings, history)
                if create_drafts and not any(batch["cod_no"] for batch in batches) and not history.has_numbers() and not settings.get("initial_cod_confirmed", False):
                    raise ValueError("Neither plant file contains an earlier COD number. Enter the approved first COD number once in the app before issuing.")
            except ValueError as exc:
                report.add({}, "needs_review", str(exc))
                minimum = None
                blocked = True
            present = {history.history_id(batch) for batch in batches}
            for record in history.records():
                if record["status"] in {"pending", "saving_draft"}:
                    blocked = True
                if record["case_id"] not in present:
                    key = record["case_id"]
                    plant = "plant2" if key.startswith("plant2|") else "plant1"
                    report.add({"source_id": plant, "case_id": key.removeprefix("plant2|")}, "needs_review",
                               _recovery_detail(record) if record["status"] in {"pending", "saving_draft"} else
                               (f"Previously prepared COD {record['certificate_no']} has no matching source row. Restore the original SN row and COD or ask an administrator to reconcile an archived register before issuing more certificates."
                                if record["status"] == "saved" else
                                f"Reserved COD {record['certificate_no']} has no matching source row. Restore the original row and reconcile the saved certificate before issuing."))
                    blocked = True
            for batch in batches:
                raw_state = status_value(batch["issued"])
                issues = list(batch["issues"])
                record = history.record(batch)
                completion_ack = history.sent_acknowledgement(
                    record["case_id"] if record else history.history_id(batch)
                )
                state = ("sent" if raw_state == "sent"
                         else effective_certificate_status(batch, record, completion_ack))
                batch["_effective_certificate_status"] = state
                batch["_raw_certificate_status"] = raw_state
                if record and record["status"] == "correcting":
                    pending = history.correction_for_case(record["case_id"])
                    target = (f"COD {pending['new_certificate_no']} dated {pending['new_certificate_date']}"
                              if pending else "the saved correction")
                    report.add(batch, "needs_review",
                               f"A correction to {target} was interrupted. Select this row and use Correct certificate to safely resume it; no new drafts will be created until it is finished.",
                               issue_code="certificate_correction_pending", system_alert=True)
                    blocked = True
                    continue
                if record and record["status"] in {"pending", "saving_draft"}:
                    report.add(batch, "needs_review", _recovery_detail(record),
                               issue_code="draft_recovery", system_alert=True)
                    continue
                if record and record["status"] not in {"saved", "draft_saved", "allocated"}:
                    report.add(batch, "needs_review", f"Unknown certificate history status {record['status']!r}; reconcile COD {record['certificate_no']} before issuing",
                               issue_code="history_reconcile", system_alert=True)
                    blocked = True
                    continue
                if record and record["status"] in {"allocated", "draft_saved"} and not parse_date(record["certificate_date"]):
                    report.add(batch, "needs_review", f"Reserved COD {record['certificate_no']} has no valid certificate date in history. Restore the recorded date before continuing",
                               issue_code="draft_recovery", system_alert=True)
                    continue
                if record and record["fingerprint"] != history.fingerprint(batch):
                    report.add(batch, "needs_review", "Source details changed after a certificate was saved or reserved; review the existing certificate before continuing",
                               issue_code="history_source_changed", system_alert=True)
                    continue

                if state == "sent":
                    if record and record["status"] in {"allocated", "draft_saved"}:
                        report.add(batch, "needs_review", f"This row is marked as previously sent, but COD {record['certificate_no']} still needs draft recovery. Restore Drafted and reconcile the saved draft first.",
                                   issue_code="draft_recovery", system_alert=True)
                        continue
                    if record and record["status"] == "saved":
                        if not parse_date(record["certificate_date"]):
                            report.add(batch, "needs_review", f"Completed COD {record['certificate_no']} has no saved date; check the certificate before continuing.",
                                       issue_code="sent_history_mismatch", system_alert=True)
                            continue
                        if (batch["cod_no"] != record["certificate_no"]
                                or parse_date(batch["cod_date"]) != parse_date(record["certificate_date"])):
                            batch["_history_action"] = "repair_sent"
                            candidates.append(batch)
                            continue
                        if not check_only and create_drafts:
                            history.acknowledge_sent(record["case_id"], batch["cod_no"])
                        report.add(batch, "completed", f"COD {record['certificate_no']} is marked Completed. No new draft or message will be created unless Ready is deliberately selected again.")
                        continue
                    if not batch["cod_no"] or not parse_date(batch["cod_date"]):
                        report.add(batch, "needs_review", "This row is marked as previously sent, but COD No or COD Date is missing. Check the certificate and record its number and date.",
                                   issue_code="sent_without_certificate", system_alert=True)
                        continue
                    report.add(batch, "completed", f"COD {batch['cod_no']} is marked Completed. The automation will not create another draft unless Ready is deliberately selected again.")
                    continue

                if state == "na":
                    if record and record["status"] in {"saved", "draft_saved", "allocated"}:
                        if record["status"] == "saved":
                            report.add(
                                batch, "already_prepared",
                                f"COD {record['certificate_no']} already has a confirmed Outlook draft. N/A skips new drafts but does not delete that draft.",
                                awaiting_send=True,
                            )
                        else:
                            report.add(
                                batch, "needs_review",
                                f"COD {record['certificate_no']} has an unfinished or unconfirmed draft save. N/A cannot remove it; restore Ready to finish checking the existing draft.",
                                issue_code="existing_certificate_marked_na", system_alert=True,
                            )
                        continue
                    report.add(batch, "not_applicable", "N/A: skipped")
                    continue

                if state == "drafted":
                    if record and record["status"] == "draft_saved":
                        if not parse_date(record["certificate_date"]):
                            report.add(batch, "needs_review", f"The saved draft for COD {record['certificate_no']} has no recorded date; review the existing draft before continuing",
                                       issue_code="draft_recovery", system_alert=True)
                            continue
                        batch["_history_action"] = "recover"
                        candidates.append(batch)
                        continue
                    if not record or record["status"] != "saved":
                        report.add(batch, "needs_review", "Drafted needs a matching saved Outlook draft in the local history. Check Outlook Drafts and the certificate register before continuing.",
                                   issue_code="draft_recovery", system_alert=True)
                        continue
                    if not parse_date(record["certificate_date"]):
                        report.add(batch, "needs_review", f"Drafted COD {record['certificate_no']} has no recorded date; review its history before continuing")
                        continue
                    if (batch["cod_no"] != record["certificate_no"]
                            or parse_date(batch["cod_date"]) != parse_date(record["certificate_date"])):
                        batch["_history_action"] = "repair_drafted"
                        candidates.append(batch)
                        continue
                    if not record["status_written"] and not check_only and create_drafts:
                        history.mark_drafted(batch, record["certificate_no"])
                    report.add(batch, "already_prepared", f"COD {record['certificate_no']} is Drafted in Excel. Review it in Outlook and mark it Completed here after you send it.", awaiting_send=True)
                    continue

                if state == "completed":
                    if not record or record["status"] != "saved":
                        report.add(batch, "needs_review", "Completed needs a matching saved certificate and a completion confirmation. Check Outlook Drafts and the certificate register before continuing.",
                                   issue_code="sent_history_mismatch", system_alert=True)
                        continue
                    if not parse_date(record["certificate_date"]):
                        report.add(batch, "needs_review", f"Completed COD {record['certificate_no']} has no recorded certificate date; review its history before continuing",
                                   issue_code="sent_history_mismatch", system_alert=True)
                        continue
                    if (batch["cod_no"] != record["certificate_no"]
                            or parse_date(batch["cod_date"]) != parse_date(record["certificate_date"])):
                        batch["_history_action"] = "repair_completed"
                        candidates.append(batch)
                        continue
                    report.add(batch, "completed", f"COD {record['certificate_no']} is confirmed Completed. It will not be drafted again unless Ready is deliberately selected.")
                    continue

                if batch.get('inward_date') and batch['inward_date'] > as_of:
                    issues.append('Inward Date is in the future; correct it before preparing the certificate')
                if state == "blank":
                    if "issued" in batch["missing_headers"]:
                        status_issue = "Add a Certificate Issued column. Choose Ready, N/A, or Blank in Excel; the app records Drafted and Completed"
                    elif str(batch["issued"] or "").strip().casefold() == "blank":
                        status_issue = "Certificate Issued is marked Blank: choose Ready or N/A when the certificate is approved"
                    else:
                        status_issue = "Certificate Issued is blank (white): choose Ready or N/A"
                    if record and record["status"] == "saved":
                        status_issue += f". A confirmed draft for COD {record['certificate_no']} already exists; this row will not create another draft while status is Blank"
                    issues.insert(0, status_issue)
                elif state == "unknown":
                    issues.insert(0, f"Certificate Issued value {batch['issued']!r} is not Ready, N/A, Blank, Drafted, Completed, or Sent")

                if state == "ready":
                    for key, title in (("cod_no", "COD No"), ("cod_date", "COD Date")):
                        if key in batch["missing_headers"]:
                            issues.append(f"Add the {title} column before preparing drafts")

                if record and record["status"] in {"allocated", "draft_saved"}:
                    if state != "ready":
                        report.add(batch, "needs_review", f"COD {record['certificate_no']} is reserved ({record['status']}); restore Ready or Drafted so the saved draft can be reconciled. No new draft was created")
                        continue
                    if issues:
                        report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)))
                        continue
                    batch["_history_action"] = "recover" if record["status"] == "draft_saved" else "allocated"
                    candidates.append(batch)
                    continue

                if state != "ready":
                    if issues:
                        report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)))
                    else:
                        report.add(batch, "needs_review", "Choose Ready, N/A, or Blank in Certificate Issued")
                    continue

                if record and record["status"] == "saved":
                    recorded_date = parse_date(record["certificate_date"])
                    if not recorded_date:
                        report.add(batch, "needs_review", f"Saved certificate COD {record['certificate_no']} has no valid saved date; review its history before continuing")
                        continue
                    if issues:
                        report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)))
                        continue
                    batch["_history_action"] = "reissue" if record["status_written"] else "complete_legacy"
                    if (record["status_written"] and check_only):
                        report.add(batch, "ready", f"Ready was selected again. A new unsent draft will reuse the saved COD {record['certificate_no']} and its original date; the existing draft will remain unchanged.")
                        continue
                    candidates.append(batch)
                    continue

                if batch["cod_no"]:
                    custom_issue = _custom_cod_issue(batch, batches, history, settings)
                    if custom_issue:
                        report.add(batch, "needs_review", custom_issue)
                        continue
                if issues:
                    report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)))
                    continue
                try:
                    display_number = record["certificate_no"] if record else certificate_number(minimum or settings["certificate_number_start"], settings)
                    certificate_layout(batch, display_number if create_drafts or check_only else "PREVIEW", parse_date(record["certificate_date"]) if record and record["certificate_date"] else as_of)
                except ValueError as exc:
                    report.add(batch, "needs_review", str(exc))
                    continue
                candidates.append(batch)

            if blocked:
                for batch in candidates:
                    report.add(batch, "waiting", "Not processed: both plant registers and global COD history must pass review before issuing")
                return report.finish()
            if check_only:
                for batch in candidates:
                    record = history.record(batch)
                    awaiting_send = bool(record and record["status"] in {"saved", "draft_saved"}
                                         and state not in {"completed", "sent"})
                    action = batch.get("_history_action")
                    if action == "complete_legacy":
                        report.add(batch, "already_prepared", f"The existing draft for COD {record['certificate_no']} is confirmed. Create drafts will restore its COD, date, and Drafted status in Excel without making another draft.", awaiting_send=awaiting_send)
                    elif action == "repair_drafted":
                        report.add(batch, "already_prepared", f"The existing draft for COD {record['certificate_no']} is confirmed. Create drafts will restore its COD, date, and Drafted status in Excel without making another draft.", awaiting_send=True)
                    elif action in {"repair_completed", "repair_sent"}:
                        report.add(batch, "completed", f"The completed COD {record['certificate_no']} is confirmed in history. Create drafts will restore its COD and date in Excel without creating or sending a message.")
                    elif action == "reissue":
                        report.add(batch, "ready", f"Ready was selected again. Creating drafts will reuse COD {record['certificate_no']} and its original date for a new unsent Outlook draft.", awaiting_send=awaiting_send)
                    elif action == "recover":
                        report.add(batch, "ready", f"The existing draft for COD {record['certificate_no']} will be kept; Create drafts will finish the Excel update without making another draft.", awaiting_send=awaiting_send)
                    else:
                        detail = (f"Ready and complete; custom COD {batch['cod_no']} will be used"
                                  if batch["cod_no"] else "Ready and complete; eligible for a certificate")
                        report.add(batch, "ready", detail, awaiting_send=awaiting_send)
                return report.finish()
            if not candidates:
                return report.finish()
            with ExitStack() as stack:
                sessions = {}
                if create_drafts:
                    # Lock both inputs before choosing a global number. Excel's own
                    # file locks cause a clean stop before any number is allocated.
                    for path in sorted(paths, key=lambda p: str(p).casefold()):
                        sessions[str(path)] = stack.enter_context(WorkbookSession(path, root / "Data" / "Workbook Backups"))
                    if _snapshot(_read_sources(paths, sessions)) != _snapshot(batches):
                        raise ValueError("A plant workbook changed during the check. Run again using the saved files.")
                    history.validate_sources(paths, batches, bind=True)
                for batch in candidates:
                    if limit is not None and report.created + report.updated >= limit:
                        report.add(batch, "waiting", "Not processed: the one-draft action was limited to one batch")
                        reported_candidates.add(id(batch))
                        continue
                    try:
                        progress(f"{batch['source_id']} / SN {batch['case_id']}: preparing")
                        reported_candidates.add(id(batch))
                        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", batch["case_id"])
                        action = batch.get("_history_action", "new")
                        # The eligibility scan handles both plants before any
                        # writes begin. Re-read each candidate's own status here;
                        # relying on the last scanned row can reject Drafted
                        # recovery simply because the other plant ends with N/A.
                        state = batch.get("_effective_certificate_status", status_value(batch["issued"]))
                        raw_state = batch.get("_raw_certificate_status", status_value(batch["issued"]))
                        if not create_drafts:
                            if action in {"complete_legacy", "repair_drafted"}:
                                record = history.record(batch)
                                report.add(batch, "already_prepared", f"COD {record['certificate_no']} already has a confirmed draft; no preview or additional draft was made")
                                continue
                            if action in {"repair_completed", "repair_sent"}:
                                record = history.record(batch)
                                report.add(batch, "completed", f"COD {record['certificate_no']} is already recorded as completed; preview mode does not change Excel or create a message")
                                continue
                            if action in {"allocated", "recover"}:
                                report.add(batch, "needs_review", "This certificate has a reserved or saved draft. Use Create drafts to finish its existing Excel update; no preview was made.")
                                continue
                            pdf = report.output / f"Preview_{batch['source_id']}_{safe_id}.pdf"
                            make_certificate(batch, "PREVIEW", as_of, resource(root, "templates") / "Certificate of Recycling.png", pdf)
                            report.created += 1
                            report.add(batch, "preview", f"{pdf.name}; no COD number allocated and Excel unchanged")
                            continue
                        session = sessions[batch["source_path"]]
                        record = history.record(batch)
                        expected_current_no = None
                        expected_previous_no = record.get("prior_certificate_no") if record else None
                        expected_previous_date = record.get("prior_certificate_date") if record else None
                        if record:
                            expected_current_no = (expected_previous_no if expected_previous_no and batch["cod_no"] == expected_previous_no
                                                   else record["certificate_no"])
                        elif batch["cod_no"]:
                            expected_current_no = batch["cod_no"]
                        session.prepare(
                            batch, expected_current_no,
                            allow_drafted=(state == "drafted" and action in {"recover", "repair_drafted"}),
                            allow_completed=(
                                (state == "completed" and action in {"repair_completed", "repair_sent"})
                                or (raw_state == "completed" and state == "drafted"
                                    and action in {"recover", "repair_drafted"})
                            ),
                            expected_previous_no=expected_previous_no,
                            expected_previous_date=expected_previous_date,
                            allow_history_restore=action in {"complete_legacy", "repair_drafted", "repair_completed", "repair_sent", "recover", "allocated", "reissue"},
                            allow_sent=action == "repair_sent",
                        )
                        if action in {"complete_legacy", "repair_drafted", "repair_completed", "repair_sent"}:
                            issued_date = parse_date(record["certificate_date"])
                            target_status = "Completed" if action in {"repair_completed", "repair_sent"} else "Drafted"
                            backup = session.commit(
                                batch, record["certificate_no"], issued_date,
                                allow_history_restore=True, target_status=target_status,
                                allow_legacy_sent=action == "repair_sent",
                            )
                            if action == "repair_sent":
                                history.acknowledge_sent(record["case_id"], record["certificate_no"])
                            elif action in {"complete_legacy", "repair_drafted"}:
                                history.mark_drafted(batch, record["certificate_no"])
                            report.updated += 1
                            batch["_effective_certificate_status"] = "completed" if target_status == "Completed" else "drafted"
                            awaiting_send = target_status == "Drafted"
                            if target_status == "Completed":
                                report.add(batch, "completed", f"Excel was restored to the confirmed completed COD {record['certificate_no']} and its original date. No message was created or sent. Backup: {backup}")
                            else:
                                report.add(batch, "already_prepared", f"Existing draft for COD {record['certificate_no']} confirmed; Excel COD and date were restored and status set to Drafted. No additional draft was created. Backup: {backup}", awaiting_send=awaiting_send)
                            continue
                        if action == "recover" and record and record["status"] == "draft_saved":
                            issued_date = parse_date(record["certificate_date"])
                            if not issued_date:
                                raise ValueError("The saved draft has no recorded certificate date; review before updating Excel")
                            backup = session.commit(
                                batch, record["certificate_no"], issued_date,
                                expected_previous_no=expected_previous_no,
                                expected_previous_date=expected_previous_date,
                                allow_history_restore=True,
                                target_status="Drafted",
                            )
                            history.complete(batch, record["pdf_path"], record["entry_id"])
                            report.updated += 1
                            batch["_effective_certificate_status"] = "drafted"
                            report.add(batch, "updated", f"COD {record['certificate_no']} written to Excel and marked Drafted; the existing draft was retained. Backup: {backup}",
                                       awaiting_send=True)
                            continue
                        if action == "reissue":
                            record = history.allocate(batch, settings, minimum, as_of, reissue=True)
                        elif record and record["status"] == "allocated":
                            record = history.allocate(batch, settings, minimum, as_of)
                        elif record and record["status"] != "allocated":
                            raise ValueError("This batch needs review before another draft can be prepared")
                        if writer is None:
                            writer = OutlookWriter(mailbox)
                        if not record:
                            issued_date = parse_date(batch["cod_date"]) or as_of
                            record = history.allocate(
                                batch, settings, minimum, issued_date,
                                custom_certificate_no=batch["cod_no"] or None,
                            )
                        cert_no = record["certificate_no"]
                        issued_date = parse_date(record["certificate_date"])
                        minimum = max(minimum, certificate_sequence(cert_no, settings) + 1)
                        safe_number = re.sub(r"[^A-Za-z0-9_-]", "_", cert_no)
                        pdf = report.output / f"Certificate_{batch['source_id']}_{safe_id}_{safe_number}.pdf"
                        make_certificate(batch, cert_no, issued_date, resource(root, "templates") / "Certificate of Recycling.png", pdf)
                        history.saving_draft(batch, pdf)
                        entry = writer.save_draft(batch, f"Certificate of E-Waste Recycling - {batch['supplier']} - {cert_no}", draft_html(batch), pdf, resource(root, "templates") / "signature-banner.png")
                        history.draft_saved(batch, pdf, entry)
                        report.created += 1
                        backup = session.commit(
                            batch, cert_no, issued_date,
                            expected_previous_no=record.get("prior_certificate_no"),
                            expected_previous_date=record.get("prior_certificate_date"),
                            allow_history_restore=action in {"allocated", "reissue"},
                            target_status="Drafted",
                        )
                        history.complete(batch, pdf, entry)
                        batch["_effective_certificate_status"] = "drafted"
                        if action == "reissue":
                            detail = (f"Re-draft saved using existing COD {cert_no} and its original date. "
                                      "A new unsent Outlook draft was created; the earlier draft was left untouched. "
                                      f"Excel was marked Drafted. Backup: {backup}")
                        else:
                            detail = (f"COD {cert_no}: PDF attached to an unsent Outlook draft; COD No, COD Date, "
                                      f"and Certificate Issued = Drafted were updated in Excel. Backup: {backup}")
                        report.add(batch, "draft_created", detail,
                                   awaiting_send=True)
                    except Exception as exc:
                        record = history.record(batch) if create_drafts else None
                        recovery = ""
                        if record and record["status"] == "draft_saved":
                            recovery = " The draft is saved. The next run will retry only the Excel update with the same COD number."
                        elif record and record["status"] == "saving_draft":
                            recovery = " " + _recovery_detail(record)
                        report.add(batch, "needs_review", str(exc) + recovery,
                                   issue_code="outlook_save_uncertain" if record and record["status"] == "saving_draft" else "case_processing_failure",
                                   system_alert=True)
                        if record and record["status"] == "saving_draft":
                            for remaining in candidates[candidates.index(batch) + 1:]:
                                report.add(remaining, "waiting", "Batch stopped because an Outlook save needs review")
                            break
            return report.finish()
        except Exception as exc:
            report.add({}, "needs_review", str(exc), issue_code="run_failure")
            for batch in candidates:
                if id(batch) not in reported_candidates:
                    report.add(batch, "waiting", "Not processed: the run stopped before this batch could be prepared; resolve the reported error and check again")
            return report.finish()
        finally:
            if history is not None:
                history.close()
