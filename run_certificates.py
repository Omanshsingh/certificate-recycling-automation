"""Read both plants, prepare ready certificates and reconcile Excel updates."""
from __future__ import annotations

from contextlib import ExitStack
from datetime import date
from io import BytesIO
from pathlib import Path
import re

from certificate_core import (batch_content, certificate_layout, certificate_number, certificate_sequence,
                              make_certificate, parse_date, read_batches,
                              status_value, supplier_lines)
from draft_service import History, OutlookWriter, draft_html
from workbook_store import WorkbookSession
from workflow_support import Report, load_settings, run_lock


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
    owners = {}
    highest = 0
    for batch in batches:
        if not batch["cod_no"]:
            continue
        number = certificate_sequence(batch["cod_no"], settings)
        owner = f"{batch['source_id']} SN {batch['case_id']}"
        if number in owners:
            raise ValueError(f"COD number {batch['cod_no']} is already repeated in {owners[number]} and {owner}. Resolve it before generating more.")
        owners[number] = owner
        highest = max(highest, number)
    if history is not None:
        registered = {}
        identities = [history.history_id(batch) for batch in batches]
        if len(set(identities)) != len(identities):
            raise ValueError("Two source rows share a certificate history key. Correct duplicate SN values or a plant2| prefix in the first plant's SN before issuing.")
        for record in history.records():
            number = certificate_sequence(record["certificate_no"], settings)
            if number in registered:
                raise ValueError(f"COD {record['certificate_no']} has multiple registry owners; reconcile certificate history before issuing.")
            registered[number] = record["case_id"]
            highest = max(highest, number)
        for batch in batches:
            if batch["cod_no"]:
                number = certificate_sequence(batch["cod_no"], settings)
                if number in registered and registered[number] != history.history_id(batch):
                    raise ValueError(f"COD {batch['cod_no']} in {batch['source_id']} SN {batch['case_id']} belongs to another batch in certificate history. Restore its correct owner before issuing.")
    return max(highest + 1, int(settings.get("certificate_number_start", 1)))


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
                state = status_value(batch["issued"])
                issues = list(batch["issues"])
                record = history.record(batch)
                if record and record["status"] in {"pending", "saving_draft"}:
                    report.add(batch, "needs_review", _recovery_detail(record))
                    continue
                if record and record["status"] not in {"saved", "draft_saved", "allocated"}:
                    report.add(batch, "needs_review", f"Unknown certificate history status {record['status']!r}; reconcile COD {record['certificate_no']} before issuing")
                    blocked = True
                    continue
                if record and record["status"] in {"allocated", "draft_saved"} and not parse_date(record["certificate_date"]):
                    report.add(batch, "needs_review", f"Reserved COD {record['certificate_no']} has no valid certificate date in history. Restore the recorded date before continuing")
                    continue
                if record and record["status"] != "saved" and state != "ready":
                    report.add(batch, "needs_review", f"COD {record['certificate_no']} is reserved ({record['status']}) but Certificate Issued is now {batch['issued']!r}. Review the existing draft and restore Ready to reconcile Excel; no new draft was created")
                    continue
                if state == "na":
                    report.add(batch, "not_applicable", "N/A: skipped")
                    continue
                if batch.get('inward_date') and batch['inward_date'] > as_of:
                    issues.append('Inward Date is in the future; correct it before preparing the certificate')
                if state == "blank":
                    issues.insert(0, "Certificate Issued is blank (white): choose Ready or N/A" if "issued" not in batch["missing_headers"] else "Add a Certificate Issued column and choose Ready or N/A for each batch")
                elif state == "unknown":
                    issues.insert(0, f"Certificate Issued value {batch['issued']!r} is not Ready or N/A")
                if record and record["fingerprint"] != history.fingerprint(batch):
                    issues.append("Source details changed after a certificate was reserved; review the existing certificate first")
                if batch["cod_no"] and record and (record["certificate_no"] != batch["cod_no"] or record["fingerprint"] != history.fingerprint(batch)):
                    report.add(batch, "needs_review", "The Excel COD number or source details differ from the saved certificate; review the existing certificate before changing it")
                    continue
                if batch["cod_no"] and record:
                    recorded_date = parse_date(record["certificate_date"])
                    if recorded_date and parse_date(batch["cod_date"]) != recorded_date:
                        report.add(batch, "needs_review", f"Excel COD Date differs from the saved certificate date {recorded_date.isoformat()}; reconcile the existing certificate before continuing")
                        continue
                    if record["status"] == "allocated":
                        report.add(batch, "needs_review", "Excel contains a reserved COD but no confirmed saved draft. Reconcile the existing certificate before continuing")
                        continue
                if batch["cod_no"] and state == "ready" and not (record and record["status"] == "draft_saved"):
                    report.add(batch, "already_prepared", f"COD {batch['cod_no']} already exists; no new number or draft")
                    continue
                for key, title in (("cod_no", "COD No"), ("cod_date", "COD Date")):
                    if key in batch["missing_headers"]:
                        issues.append(f"Add the {title} column before preparing drafts")
                if record and record["status"] == "saved" and not batch["cod_no"]:
                    issues.append("A certificate was already prepared but its Excel COD No is blank; review the saved certificate")
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
                    report.add(batch, "ready", "Saved draft awaiting Excel reconciliation" if history.record(batch) and history.record(batch)["status"] == "draft_saved" else "Ready and complete; eligible for a certificate")
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
                        if not create_drafts:
                            pdf = report.output / f"Preview_{batch['source_id']}_{safe_id}.pdf"
                            make_certificate(batch, "PREVIEW", as_of, root / "templates" / "Certificate of Recycling.png", pdf)
                            report.created += 1
                            report.add(batch, "preview", f"{pdf.name}; no COD number allocated and Excel unchanged")
                            continue
                        session = sessions[batch["source_path"]]
                        record = history.record(batch)
                        session.prepare(batch, record["certificate_no"] if record else None)
                        if record and record["status"] == "draft_saved":
                            issued_date = parse_date(record["certificate_date"])
                            if not issued_date:
                                raise ValueError("The saved draft has no recorded certificate date; review before updating Excel")
                            backup = session.commit(batch, record["certificate_no"], issued_date)
                            history.complete(batch, record["pdf_path"], record["entry_id"])
                            report.updated += 1
                            report.add(batch, "updated", f"COD {record['certificate_no']} written to Excel; existing draft retained. Backup: {backup}")
                            continue
                        if record and record["status"] != "allocated":
                            raise ValueError("This batch needs review before another draft can be prepared")
                        if writer is None:
                            writer = OutlookWriter(mailbox)
                        record = history.allocate(batch, settings, minimum, as_of)
                        cert_no = record["certificate_no"]
                        issued_date = parse_date(record["certificate_date"])
                        minimum = max(minimum, certificate_sequence(cert_no, settings) + 1)
                        safe_number = re.sub(r"[^A-Za-z0-9_-]", "_", cert_no)
                        pdf = report.output / f"Certificate_{batch['source_id']}_{safe_id}_{safe_number}.pdf"
                        make_certificate(batch, cert_no, issued_date, root / "templates" / "Certificate of Recycling.png", pdf)
                        history.saving_draft(batch, pdf)
                        entry = writer.save_draft(batch, f"Certificate of E-Waste Recycling - {batch['supplier']} - {cert_no}", draft_html(batch), pdf, root / "templates" / "signature-banner.png")
                        history.draft_saved(batch, pdf, entry)
                        report.created += 1
                        backup = session.commit(batch, cert_no, issued_date)
                        history.complete(batch, pdf, entry)
                        report.add(batch, "draft_created", f"COD {cert_no}: PDF attached to Outlook draft; COD No and COD Date updated in Excel. Backup: {backup}")
                    except Exception as exc:
                        record = history.record(batch) if create_drafts else None
                        recovery = ""
                        if record and record["status"] == "draft_saved":
                            recovery = " The draft is saved. The next run will retry only the Excel update with the same COD number."
                        elif record and record["status"] == "saving_draft":
                            recovery = " " + _recovery_detail(record)
                        report.add(batch, "needs_review", str(exc) + recovery)
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
