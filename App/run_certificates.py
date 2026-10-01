"""Generate recycling certificates as PDFs from the two plant workbooks.

This workflow has no email or Outlook integration. A user marks a row Ready,
generates its PDF, sends it through their normal process, and marks it Completed.
"""
from __future__ import annotations

from contextlib import ExitStack
from datetime import date
from io import BytesIO
import re
from pathlib import Path
import shutil
import sqlite3
import uuid
import json

from certificate_core import (
    certificate_layout,
    certificate_number,
    certificate_reminder_due,
    certificate_sequence,
    make_certificate,
    number,
    parse_date,
    read_batches,
    status_value,
)
from workbook_store import WorkbookSession, stage_certificate_details
from workflow_support import Report, load_settings, run_lock
from app_paths import resource
from certificate_templates import TEMPLATE_ROOT, template_id, validate_fields, layout_warnings
from certificate_inventory import scan_inventory, cod_conflicts, reference_warnings, verify_publication, issuance_paths, record_issuance
from issuance_safety import IssuanceTransaction, recover_issuances, latest_issuance, valid_pdf, file_digest


def _safe_cod_name(value: str) -> str:
    """Make a COD safe as a Windows filename without changing the workbook COD."""
    raw = str(value).strip()
    raw_device_name = raw.split(".", 1)[0].rstrip(" .").upper()
    if raw_device_name in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", raw_device_name):
        raise ValueError("This COD is a reserved Windows filename. Enter a different COD number.")
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", raw).strip(" ._")
    if not safe or safe in {".", ".."}:
        raise ValueError("Enter a COD number that can be used as a PDF filename.")
    device_name = safe.split(".", 1)[0].upper()
    if device_name in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", device_name):
        raise ValueError("This COD is a reserved Windows filename. Enter a different COD number.")
    if len(safe) > 240:
        raise ValueError("This COD is too long to use as a PDF filename. Shorten it to 240 characters or fewer.")
    return safe


def _safe_name_part(value, fallback):
    """Keep the requested human-readable PDF name safe on Windows."""
    raw = re.sub(r"\s+", "_", str(value or "").strip())
    raw = re.sub(r"[^A-Za-z0-9._-]+", "_", raw)
    raw = re.sub(r"_+", "_", raw).strip(" ._")
    return raw or fallback


def cod_identity(value):
    raw = str(value or "").strip()
    return str(int(raw)) if raw.isdecimal() and raw.isascii() else _safe_name_part(raw, "").casefold()


def certificate_filename(batch, plant_number, cod=None):
    """Return a COD-led name, e.g. 0319_Acme_Mumbai_1.pdf."""
    cod = str(cod if cod is not None else batch.get("cod_no") or "").strip()
    if not cod:
        raise ValueError("Assign a COD number before naming the certificate PDF.")
    raw_company = str(batch.get("supplier") or "").strip()
    raw_company = re.sub(r"^plant[\s_-]*[12](?=$|[\s_-])[\s_-]*", "", raw_company, flags=re.I)
    company = _safe_name_part(raw_company, "Unknown Company")
    city = _safe_name_part(batch.get("location"), "Unknown City")
    cod_part = _safe_cod_name(cod)
    plant_part = f"_{int(plant_number)}"
    company_city = f"{company}_{city}"
    available = 240 - len(cod_part) - 1 - len(plant_part)
    if available < 1:
        raise ValueError("This COD is too long to include the company, city, and plant number in the PDF filename.")
    company_city = company_city[:available].rstrip(" ._") or "Client"
    return f"{cod_part}_{company_city}{plant_part}.pdf"


def _read_sources(paths, sessions, report):
    batches = []
    for index, path in enumerate(paths, 1):
        source_id = f"plant{index}"
        try:
            source_batches = list(read_batches(BytesIO(sessions[index - 1].data)))
        except Exception as exc:
            report.add({"source_id": source_id, "source_path": str(path)}, "needs_review",
                       f"Could not read this plant workbook: {exc}", issue_code="source_unavailable")
            continue
        report.observed_sources.append(source_id)
        for batch in source_batches:
            batch.update(source_id=source_id, source_path=str(path))
            batches.append(batch)
    report.scan_complete = len(report.observed_sources) == 2
    return batches


def _highest_sequence(batches, settings, output_folder, legacy_history=None, inventory=None):
    highest = max(0, int(settings.get("certificate_number_start", 1)) - 1)
    for batch in batches:
        if batch.get("cod_no"):
            highest = max(highest, certificate_sequence(batch["cod_no"], settings))
    # Existing flat output PDFs are the durable numbering floor if an old
    # workbook row is archived or its COD cell is cleared by the user.
    output_folder = Path(output_folder).resolve()
    output_root = output_folder.parent
    inventory = inventory if inventory is not None else scan_inventory(output_root)
    for entry in inventory['entries']:
        for value in entry['cods']:
            highest = max(highest, certificate_sequence(value, settings))
    custom_register = output_root / "Data" / "Custom Certificates.json"
    if custom_register.is_file():
        custom_data = json.loads(custom_register.read_text(encoding="utf-8"))
        for record in custom_data.get("records", []):
            highest = max(highest, certificate_sequence(record.get("details", {}).get("cod_no", ""), settings))
    legacy_certificate_folder = output_root / "Output" / "Certificates"
    # Only durable certificate/history locations, never sample templates,
    # preview files or development/test directories elsewhere in the project.
    durable_folders = (output_folder, legacy_certificate_folder, output_root / 'Data' / 'Archive')
    for durable_folder in durable_folders:
        for pdf in durable_folder.rglob('*.pdf') if durable_folder.is_dir() else ():
            if pdf.name.startswith('.'):
                continue
            if pdf.is_relative_to(output_root / 'Data' / 'Archive' / 'Uncommitted Certificates'):
                # A rollback quarantine is diagnostic evidence, not an issued
                # certificate or a reason to consume its never-committed COD.
                continue
            stem = pdf.stem
            if durable_folder == output_root / 'Data' / 'Archive':
                # Recovery archives prepend a random ID, not a COD.
                stem = re.sub(r'^[a-f0-9]{32}_', '', stem, flags=re.I)
            # New COD-led PDFs end in _1 or _2 for the plant. Read the COD
            # from the first filename field, never from that trailing marker.
            new_client_name = re.fullmatch(r"[^_]+_.+_[12]", stem, re.I)
            if new_client_name:
                cod_prefix = stem.split("_", 1)[0]
                highest = max(highest, certificate_sequence(cod_prefix, settings))
                continue
            # Do not treat older client-name PDFs as a number source.
            legacy_client_name = re.fullmatch(r"cod_.+_[12]", stem, re.I)
            legacy_cod_name = (
                bool(re.fullmatch(r"(?:cod_|certificate_)\d+", stem, re.I))
                or ("_" not in stem and stem.casefold().startswith(("cod", "certificate")))
            )
            in_flat_certificate_folder = pdf.parent.resolve() in {
                output_folder,
                legacy_certificate_folder.resolve(),
            }
            if not (legacy_cod_name or (in_flat_certificate_folder and not legacy_client_name)):
                continue
            match = re.search(r"(\d+)$", stem)
            if match:
                highest = max(highest, int(match.group(1)))
    # Read only legacy COD values for sequence continuity; no mail state is used.
    if legacy_history and Path(legacy_history).is_file():
        try:
            connection = sqlite3.connect(f"file:{Path(legacy_history).resolve().as_posix()}?mode=ro", uri=True)
            try:
                tables = {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                for table in tables.intersection({"batches", "archived_issuances"}):
                    columns = {row[1] for row in connection.execute(f"PRAGMA table_info({table})")}
                    if "certificate_no" not in columns:
                        continue
                    for (value,) in connection.execute(f"SELECT certificate_no FROM {table}"):
                        highest = max(highest, certificate_sequence(value, settings))
            finally:
                connection.close()
        except sqlite3.Error:
            pass
    return highest


def _apply_overrides(batch, overrides):
    """Build the certificate view without mutating the parsed source batch."""
    if not overrides:
        return dict(batch)
    view = dict(batch)
    view["issues"] = list(batch.get("issues", []))
    for key in ("supplier", "location", "form6", "reference"):
        if key in overrides:
            view[key] = str(overrides[key] or "").strip()
    for key in ("certificate_type", "certificate_fields"):
        if key in overrides:
            view[key] = overrides[key]
    if "inward_date" in overrides:
        view["inward_date"] = parse_date(overrides["inward_date"])
    if "certificate_weight" in overrides:
        raw = overrides["certificate_weight"]
        parsed = raw if isinstance(raw, (int, float)) else number(raw)
        view["certificate_weight"] = parsed
        view["certificate_weight_invalid"] = parsed is None or parsed <= 0
    # An override is an explicit correction for the corresponding warning.
    if view.get("supplier"):
        view["issues"] = [issue for issue in view["issues"] if issue != "missing supplier name"]
    if view.get("inward_date"):
        view["issues"] = [issue for issue in view["issues"] if issue != "missing inward date"]
    if view.get("form6"):
        view["issues"] = [issue for issue in view["issues"] if issue != "missing Form 6 number / manifest number"]
    if "reference" in overrides:
        view["issues"] = [issue for issue in view["issues"] if "reference" not in issue.casefold()]
    if "certificate_weight" in overrides:
        view["issues"] = [issue for issue in view["issues"] if "certificate/net weight" not in issue]
    certificate_weight = view.get("certificate_weight")
    if certificate_weight is not None and certificate_weight > 0:
        view["display_weight"] = certificate_weight
        view["weight_source"] = "certificate"
    else:
        view["display_weight"] = view.get("total_weight") or 0
        view["weight_source"] = "inward"
        if "certificate_weight" in overrides:
            view["issues"].append("certificate/net weight override must be a positive number")
    return view


def find_certificate(output_folder, batch):
    """Match stable COD + plant, even if the company/city was renamed in Excel."""
    folder = Path(output_folder)
    cod = str(batch.get("cod_no") or "").strip()
    if not cod or not folder.is_dir():
        return None
    try:
        safe = _safe_cod_name(cod)
    except ValueError:
        return None
    plant = str(batch.get("source_id") or "")[-1:]
    receipt = latest_issuance(folder.parent, 'workbook', batch.get('source_id'), batch.get('case_id'))
    if receipt and cod_identity(receipt.get('cod')) == cod_identity(cod):
        current = folder.parent / receipt['pdf_path']
        return current if valid_pdf(current) and file_digest(current) == receipt['pdf_sha256'] else None
    for path in folder.rglob("*"):
        if path.suffix.casefold() != '.pdf' or not valid_pdf(path):
            continue
        prefix = path.stem.split("_", 1)[0]
        same_prefix = path.stem.casefold().startswith((safe + "_").casefold())
        if safe.isdecimal() and prefix.isdecimal():
            same_prefix = cod_identity(safe) == cod_identity(prefix)
        if same_prefix and path.stem.endswith(f"_{plant}"):
            return path
        if path.stem.casefold() in {safe.casefold(), f"cod_{safe}".casefold(), f"certificate_{safe}".casefold()}:
            return path
        if safe.isdecimal() and path.stem.isdecimal() and cod_identity(safe) == cod_identity(path.stem):
            return path
    return None


def run(workbook_paths, root, *, as_of=None, progress=print, check_only=False,
        publish_latest=True, limit=None, selected=None, custom_cod=None,
        plants=None, allow_warnings=False, overrides=None, details=None, update_excel=True,
        override_reason='', expected_revision=None, report_sink=None):
    """Check both workbooks or generate PDFs for rows explicitly marked Ready."""
    root = Path(root)
    settings = load_settings(root)
    as_of = as_of or date.today()
    manual_cod = str(custom_cod or "").strip()
    if plants is not None:
        if not isinstance(plants, (list, tuple)) or not plants:
            raise ValueError("Choose Plant 1, Plant 2, or All plants.")
        plants = tuple(dict.fromkeys(plants))
        if any(plant not in {"plant1", "plant2"} for plant in plants):
            raise ValueError("Choose Plant 1, Plant 2, or All plants.")
        if selected:
            raise ValueError("Choose one inward or a plant selection, not both.")
    if manual_cod and not selected:
        raise ValueError("A custom COD can only be used when generating one selected Ready row.")
    if details is not None and (not selected or str(details.get('case_id')) != str(selected[1])):
        raise ValueError("The selected inward ID cannot be changed when editing its certificate.")
    mode = "Check plant files" if check_only else "Generate certificates"
    with run_lock(root):
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        report = Report(root, as_of, mode, publish_latest=publish_latest)
        report.result_sink = report_sink
        try:
            paths = [Path(path).expanduser().resolve() for path in workbook_paths]
            if len(paths) != 2 or any(not str(path) for path in paths):
                raise ValueError("Select the Plant 1 and Plant 2 Excel files in Settings.")
            if paths[0] == paths[1] or (paths[0].is_file() and paths[1].is_file() and paths[0].samefile(paths[1])):
                raise ValueError("Plant 1 and Plant 2 must be two different Excel files.")

            backup_dir = root / "Data" / "Workbook Backups"
            output_folder = root / "Certificates"
            template = resource(root, "templates") / "Certificate of Recycling.png"
            with ExitStack() as stack:
                sessions = [stack.enter_context(WorkbookSession(path, backup_dir)) for path in paths]
                if expected_revision is not None:
                    require_revision(root, expected_revision, {str(session.path.resolve()): session.data for session in sessions})
                inventory = scan_inventory(root, persist=True)
                report.certificate_inventory = inventory
                batches = _read_sources(paths, sessions, report)
                if not report.scan_complete:
                    report.add({}, "needs_review", "Both plant files must be available so COD numbers stay in one shared sequence.",
                               issue_code="source_unavailable", system_alert=True)
                    return report.finish()

                original_form_batch = None
                staged_data = None
                numbering_batches = list(batches)
                if details is not None:
                    session = sessions[int(selected[0][-1]) - 1]
                    original_form_batch = next((batch for batch in batches
                                                if (batch['source_id'], str(batch['case_id'])) == selected), None)
                    if original_form_batch is None and not update_excel:
                        raise ValueError("Certificate-only changes require an existing inward. Choose Also update Excel or create a manual certificate.")
                    staged_data, staged_batch = stage_certificate_details(session, details, validate_cells=update_excel)
                    staged_batch.update(source_id=selected[0], source_path=str(session.path))
                    batches = [batch for batch in batches if (batch['source_id'], str(batch['case_id'])) != selected]
                    batches.append(staged_batch)
                    numbering_batches.append(staged_batch)

                for index, batch in enumerate(batches):
                    if plants is not None and batch["source_id"] not in plants:
                        continue
                    if details is None and (index == 0 or batch["source_id"] != batches[index - 1]["source_id"]):
                        sessions[int(batch["source_id"][-1]) - 1].ensure_alert_formatting(batch["source_sheet"])

                next_sequence = _highest_sequence(
                    numbering_batches, settings, output_folder,
                    root / "Data" / "certificate-history.sqlite",
                    inventory=inventory,
                ) + 1
                ready_count = 0
                for batch in batches:
                    if plants is not None and batch["source_id"] not in plants:
                        continue
                    if selected and (batch["source_id"], str(batch["case_id"])) != (selected[0], str(selected[1])):
                        continue
                    if limit is not None and ready_count >= limit and not check_only:
                        break
                    state = status_value(batch.get("issued"))
                    batch["_effective_certificate_status"] = state
                    age = (as_of - batch["inward_date"]).days if batch.get("inward_date") else None
                    due = certificate_reminder_due(batch, as_of, days=25)
                    overdue = bool(due and age is not None and age > 30)

                    if state == "na":
                        report.add(batch, "not_applicable", "N/A — no certificate is required.")
                        continue
                    if state == "alert":
                        report.add(batch, "needs_review", "This row is marked Alert. Review it and choose Ready, N/A, or Blank in Excel before creating a PDF.")
                        continue
                    if state == "unknown":
                        report.add(batch, "needs_review", f"Unknown Certificate Issued value {batch.get('issued')!r}; review it, then choose Ready, Alert, N/A, or Blank.")
                        continue
                    if state in {"generated", "generated_manual", "completed"}:
                        existing_pdf = find_certificate(output_folder, batch)
                        if existing_pdf:
                            if state == "completed":
                                report.add(batch, "completed", f"COD {batch.get('cod_no')} is marked Completed.")
                            elif state == "generated_manual":
                                report.add(batch, "generated_manual", f"COD {batch.get('cod_no')} has a manually generated PDF. Send it using your normal process, then mark Completed.")
                            else:
                                report.add(batch, "generated", f"COD {batch.get('cod_no')} has a generated PDF. Send it using your normal process, then mark Completed.")
                            continue
                        try:
                            sessions[int(batch["source_id"][-1]) - 1].set_issued_status(
                                batch["case_id"], batch["source_sheet"], batch["source_row"],
                                batch["cod_no"], batch["issued"], "Ready",
                            )
                            batch["issued"] = "Ready"
                            batch["_effective_certificate_status"] = state = "ready"
                            batch["_missing_pdf"] = True
                        except Exception as exc:
                            report.add(batch, "needs_review", f"The PDF is missing, but Excel could not be changed to Ready: {exc}", system_alert=True)
                            continue
                    if state == "blank":
                        detail = "Choose Ready when the certificate should be generated, N/A to skip, or leave Blank until decided."
                        report.add(batch, "needs_review", detail)
                        continue
                    if state != "ready":
                        report.add(batch, "needs_review", "Choose Ready, N/A, or Blank in the Certificate Issued column.")
                        continue

                    ready_count += 1
                    source_batch = batch
                    batch = _apply_overrides(batch, overrides if selected and
                                             (batch["source_id"], str(batch["case_id"])) == selected else None)
                    issues = list(batch.get("issues", []))
                    try:
                        kind = template_id(batch.get("certificate_type"))
                        batch["certificate_type"] = kind
                        batch["certificate_fields"] = validate_fields(kind, batch.get("certificate_fields"))
                    except ValueError as exc:
                        report.add(batch, "needs_review", str(exc), system_alert=True)
                        continue
                    if kind in {"sita", "siemens"}:
                        issues = [issue for issue in issues if issue != "missing Form 6 number / manifest number"]
                    if kind == "siemens":
                        for key, label in (("model_name", "model name"), ("model_serial", "model serial number"),
                                           ("tube_serial", "X-ray tube serial number"), ("site_location", "site location")):
                            if not str(batch["certificate_fields"].get(key) or "").strip():
                                issues.append(f"Missing {label}; enter N/A if it does not apply")
                    if batch.get("inward_date") and batch["inward_date"] > as_of:
                        issues.append("Inward Date is in the future; correct it before generating the certificate")
                    for key, label in (("cod_no", "COD No"), ("cod_date", "COD Date"), ("issued", "Certificate Issued")):
                        if key in batch.get("missing_headers", []):
                            issues.append(f"Add the {label} column to the plant workbook")
                    if batch.get("cod_date") not in (None, "") and parse_date(batch["cod_date"]) is None:
                        issues.append("COD Date is not a valid date; correct it or clear it to use today")
                    if str(batch.get("cod_no") or "").startswith("="):
                        issues.append("COD No contains a formula; enter a fixed COD or clear the cell")
                    owned_code = manual_cod or source_batch.get('cod_no')
                    allowed_paths = issuance_paths(inventory, owned_code, 'workbook', source_batch['source_id'], source_batch['case_id'])
                    owned = bool(allowed_paths)
                    issues.extend(reference_warnings(inventory, batch.get('reference'), allowed_paths, owned=owned))
                    for entry in inventory['entries']:
                        if entry['present'] and not entry['cods']:
                            issues.append(f"{entry['path']}: its COD could not be checked. Confirm the next COD manually before continuing.")
                    if issues and not allow_warnings:
                        report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)),
                                   warnings=list(dict.fromkeys(issues)), system_alert=True)
                        continue

                    row_cod = manual_cod or str(batch.get("cod_no") or "").strip()
                    try:
                        cod = row_cod or certificate_number(next_sequence, settings)
                        cod_owners = []
                        for owner in batches:
                            if cod_identity(owner.get("cod_no")) == cod_identity(cod):
                                cod_owners.append(owner)
                        custom_register = root / "Data" / "Custom Certificates.json"
                        if custom_register.is_file():
                            for record in json.loads(custom_register.read_text(encoding="utf-8")).get("records", []):
                                if cod_identity(record.get("details", {}).get("cod_no")) == cod_identity(cod):
                                    cod_owners.append({"source_id": "custom", "case_id": record["id"]})
                        if any(owner["source_id"] != batch["source_id"] or str(owner["case_id"]) != str(batch["case_id"])
                               for owner in cod_owners):
                            message = f"COD {cod} is already assigned to another plant row. Give each certificate a unique COD."
                            if not allow_warnings:
                                raise ValueError(message)
                            issues.append(message + " A new automatic COD was used after confirmation.")
                            cod = certificate_number(next_sequence, settings)
                        conflicts = cod_conflicts(inventory, cod, allowed_paths, owned=owned and cod_identity(cod) == cod_identity(source_batch.get('cod_no')))
                        if conflicts:
                            names = ', '.join(Path(entry['path']).name for entry in conflicts[:3])
                            replacement = certificate_number(next_sequence, settings)
                            message = f"COD {cod} is already recorded in certificate PDF history ({names}). Continuing will use new automatic COD {replacement}; existing PDFs will not be overwritten."
                            if not allow_warnings:
                                report.add(batch, 'needs_review', message, warnings=[message], system_alert=True)
                                continue
                            issues.append(message)
                            cod = replacement
                        try:
                            safe_cod = _safe_cod_name(cod)
                        except ValueError as exc:
                            if not allow_warnings:
                                raise
                            issues.append(str(exc) + " A new automatic COD was used after confirmation.")
                            cod = certificate_number(next_sequence, settings)
                            safe_cod = _safe_cod_name(cod)
                        issued_date = parse_date(batch.get("cod_date")) or as_of
                        destination = output_folder / certificate_filename(
                            batch, int(batch["source_id"][-1]), cod,
                        )
                        native_template = TEMPLATE_ROOT / f"{kind}.pdf"
                        batch["_data_root"] = str(root)
                        if native_template.is_file() and (kind != "standard" or (resource(root, "templates") / "certificate-types" / "standard.pdf").is_file()):
                            active_template = native_template
                            issues.extend(layout_warnings(batch, cod, issued_date))
                            if issues and not allow_warnings:
                                report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)), warnings=list(dict.fromkeys(issues)), system_alert=True)
                                continue
                        else:
                            active_template = template
                            certificate_layout(batch, cod, issued_date, allow_warnings=allow_warnings)
                        if check_only:
                            detail = f"Ready to generate COD {cod}."
                            if issues:
                                detail += " Warnings: " + "; ".join(dict.fromkeys(issues))
                            report.add(batch, "ready", detail, system_alert=bool(issues),
                                       warnings=list(dict.fromkeys(issues)))
                            continue

                        output_folder.mkdir(parents=True, exist_ok=True)
                        temporary = output_folder / f".{safe_cod}.{uuid.uuid4().hex}.tmp.pdf"
                        transaction = None
                        workbook_committed = False
                        history_pending = False
                        session = sessions[int(batch["source_id"][-1]) - 1]
                        try:
                            make_certificate(batch, cod, issued_date, active_template, temporary,
                                             allow_warnings=allow_warnings)
                            old_paths = issuance_paths(inventory, source_batch.get('cod_no'),
                                                       'workbook', batch['source_id'], batch['case_id'])
                            verify_publication(root, inventory, cod, batch.get('reference'), list(set(allowed_paths + old_paths)),
                                               owned=owned and cod_identity(cod) == cod_identity(source_batch.get('cod_no')),
                                               allow_warnings=allow_warnings)
                            transaction = IssuanceTransaction(
                                root, destination, temporary, session.path, session.data,
                                source_kind='workbook', plant=batch['source_id'], key=batch['case_id'],
                                cod=cod, issued_date=issued_date, before=original_form_batch or source_batch,
                                after=dict(batch, cod_no=cod, cod_date=issued_date, edit_scope='update_excel' if update_excel else 'certificate_only'),
                                warnings=list(dict.fromkeys(issues)), override_reason=override_reason,
                                old_paths=[root / path for path in old_paths], template_path=active_template)
                            transaction.publish()
                            session.before_persist = transaction.expect_source
                            if details is not None:
                                backup = session.commit(original_form_batch or source_batch, cod, issued_date,
                                                        staged_data=staged_data if update_excel else None,
                                                        allow_reissue=True, new_batch=original_form_batch is None)
                            else:
                                backup = session.commit(source_batch, cod, issued_date, target_status="Generated",
                                                        overrides=overrides if selected and
                                                        (source_batch["source_id"], str(source_batch["case_id"])) == selected else None)
                            workbook_committed = True
                            try:
                                transaction.mark_source_committed()
                                transaction.complete()
                            except (OSError, ValueError) as recovery_error:
                                history_pending = True
                                issues.append(f'Certificate and Excel saved; final history update needs recovery on refresh: {recovery_error}')
                        except Exception as exc:
                            if transaction and not workbook_committed:
                                try:
                                    transaction.rollback()
                                except (OSError, ValueError) as recovery_error:
                                    raise RuntimeError(
                                        f"{exc}. The Excel update failed and the previous PDF could not be restored; "
                                        f"saved recovery copies were preserved. Ask for help before retrying: {recovery_error}"
                                    ) from exc
                            raise
                        finally:
                            session.before_persist = None
                            try:
                                temporary.unlink(missing_ok=True)
                            except OSError:
                                pass

                        batch["cod_no"] = cod
                        batch["cod_date"] = issued_date
                        batch["issued"] = "Generated"
                        batch['pdf_filename'] = destination.name
                        batch["_effective_certificate_status"] = "generated"
                        report.created += 1
                        detail = f"Generated {destination.name}; Excel was updated to COD {cod}, {issued_date:%d.%m.%Y}, Generated. Send it manually, then mark Completed."
                        if issues:
                            detail += " Warnings accepted: " + "; ".join(dict.fromkeys(issues))
                        if backup:
                            detail += f" Workbook backup: {backup.name}."
                        report.add(batch, "generated_now", detail,
                                   system_alert=bool(issues), warnings=list(dict.fromkeys(issues)))
                        if history_pending:
                            # Do not commit another row to this source until the
                            # pending transaction's exact source hash is recovered.
                            break
                        # Reserve this successful issuance immediately for any
                        # following row in the same bulk operation.
                        inventory = scan_inventory(root, persist=True)
                        if not row_cod:
                            next_sequence += 1
                        else:
                            next_sequence = max(next_sequence, certificate_sequence(cod, settings) + 1)
                        if progress:
                            progress(f"{batch['source_id']} SN {batch['case_id']}: saved {destination.name}")
                    except Exception as exc:
                        report.add(batch, "needs_review", str(exc), issue_code="certificate_generation_failure", system_alert=True)
            report.certificate_inventory = scan_inventory(root, persist=True)
            return report.finish()
        except Exception as exc:
            report.add({}, "needs_review", str(exc), issue_code="run_failure", system_alert=True)
            return report.finish()
