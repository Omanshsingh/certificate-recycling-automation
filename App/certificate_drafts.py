"""Local drafts for certificates that have no inward workbook entry yet."""
from __future__ import annotations

from datetime import date, datetime
import json
from pathlib import Path
import shutil
import uuid
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

from certificate_core import certificate_number, material_items, number, parse_date, read_batches, status_value
from certificate_templates import TEMPLATE_NAMES, layout_warnings, make_template_certificate, template_id, validate_fields, current_warnings
from run_certificates import _highest_sequence, _safe_cod_name, certificate_filename, find_certificate, cod_identity
from certificate_inventory import (scan_inventory, cod_conflicts, reference_warnings, verify_publication,
                                   issuance_paths, record_issuance, certificate_destination, safe_certificate_path)
from workflow_support import load_settings, run_lock, save_json
from issuance_safety import IssuanceTransaction, recover_issuances, valid_pdf, file_digest, latest_issuance


def register_path(root):
    return Path(root) / "Data" / "Custom Certificates.json"


def read_register(root):
    path = register_path(root)
    if not path.is_file():
        return {"version": 1, "records": []}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("records"), list):
        raise ValueError("The custom certificate register is unreadable. Restore its backup before creating another certificate.")
    ids = set()
    for row in data["records"]:
        if (not isinstance(row, dict) or not isinstance(row.get("details"), dict)
                or not isinstance(row.get("id"), str) or row["id"] in ids):
            raise ValueError("The custom certificate register contains invalid records. Restore its backup.")
        ids.add(row["id"])
    return data


def write_register(root, data):
    path = register_path(root)
    if path.exists():
        backups = Path(root) / "Data" / "Custom Register Backups"
        backups.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backups / f"Custom Certificates_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.json")
    save_json(path, data)


def clean_details(details):
    if not isinstance(details, dict):
        raise ValueError("Enter the certificate details first.")
    kind = template_id(details.get("certificate_type"))
    result = {"certificate_type": kind, "certificate_fields": validate_fields(kind, details.get("certificate_fields"))}
    for key in ("case_id", "supplier", "location", "inward_date", "item", "quantity", "per_piece_weight",
                "weight", "certificate_weight", "form6", "reference", "cod_no", "cod_date"):
        value = str(details.get(key) if details.get(key) is not None else "").strip()
        if len(value) > 6000:
            raise ValueError(f"The {key} field is too long.")
        result[key] = value
    if "items" in details:
        items = details["items"]
        if not isinstance(items, list) or len(items) > 100:
            raise ValueError("Add no more than 100 materials at a time.")
        result["items"] = []
        for index, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise ValueError(f"Material {index} is not a valid row.")
            clean_item = {}
            for key, limit in (("item", 180), ("quantity", 40), ("per_piece_weight", 40), ("weight", 40)):
                value = str(item.get(key) if item.get(key) is not None else "").strip()
                if len(value) > limit:
                    raise ValueError(f"Material {index}: {key.replace('_', ' ')} is too long.")
                clean_item[key] = value
            result["items"].append(clean_item)
    if kind == "siemens" and not result["certificate_fields"].get("site_location", "").strip():
        result["certificate_fields"]["site_location"] = result["location"]
    return result


def _record(data, record_id):
    if record_id is None:
        return None
    result = next((record for record in data["records"] if record["id"] == record_id), None)
    if result is None:
        raise ValueError("This saved certificate was not found. Refresh the queue before editing it.")
    return result


def save_draft(root, plant, details, record_id=None, expected_revision=None):
    if plant not in {"plant1", "plant2"}:
        raise ValueError("Choose Plant 1 or Plant 2.")
    details = clean_details(details)
    with run_lock(root):
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        data = read_register(root)
        existing = _record(data, record_id)
        if existing:
            # Editing a previously generated draft keeps its stable COD unless
            # the operator explicitly enters another number.
            details["cod_no"] = details["cod_no"] or existing["details"].get("cod_no", "")
            existing.update(plant=plant, details=details, status="ready", updated_at=datetime.now().isoformat())
        else:
            existing = {"id": uuid.uuid4().hex, "plant": plant, "details": details, "status": "ready",
                        "created_at": datetime.now().isoformat(), "updated_at": datetime.now().isoformat()}
            data["records"].append(existing)
        write_register(root, data)
    return existing["id"]


def numbering_sources(root, settings):
    batches, warnings = [], []
    for index, value in enumerate(settings.get("workbook_paths", []), 1):
        if not value:
            warnings.append(f"Plant {index} workbook is not connected; confirm that the COD does not duplicate an existing certificate before sending.")
            continue
        try:
            for batch in read_batches(Path(value)):
                batch["source_id"] = f"plant{index}"
                batches.append(batch)
        except (OSError, ValueError, BadZipFile, InvalidFileException) as exc:
            warnings.append(f"Plant {index} workbook could not be checked for existing CODs: {exc}. Verify the COD before sending.")
    return batches, warnings


def prepare_certificate(root, plant, details, *, record_id=None, as_of=None, source_kind="manual"):
    """Preview/preflight without saving a draft, consuming a COD or changing Excel."""
    if plant not in {"plant1", "plant2"}:
        raise ValueError("Choose Plant 1 or Plant 2.")
    details = clean_details(details)
    data = read_register(root)
    existing = _record(data, record_id)
    as_of = as_of or date.today()
    kind = details["certificate_type"]
    settings = load_settings(root)
    batches, warnings = numbering_sources(root, settings)
    inventory = scan_inventory(root)
    next_number = _highest_sequence(batches, settings, Path(root) / "Certificates",
                                    Path(root) / "Data" / "certificate-history.sqlite", inventory=inventory) + 1
    automatic = certificate_number(next_number, settings)
    requested = details["cod_no"] or (existing["details"].get("cod_no", "") if existing else "")
    cod = requested or automatic
    allowed_paths = issuance_paths(inventory, cod, 'manual', plant, record_id)
    owned = bool(allowed_paths)
    if source_kind == 'workbook':
        own_batch = next((batch for batch in batches if batch.get('source_id') == plant and batch.get('case_id') == details['case_id']), None)
        allowed_paths = issuance_paths(inventory, cod, 'workbook', plant, details['case_id'])
        owned = bool(allowed_paths)
    owners = [batch for batch in batches if cod_identity(batch.get("cod_no")) == cod_identity(cod)
              and not (source_kind == "workbook" and batch.get("source_id") == plant and batch.get("case_id") == details["case_id"])]
    owners.extend(record for record in data["records"] if record["id"] != record_id
                  and cod_identity(record["details"].get("cod_no")) == cod_identity(cod))
    if owners:
        warnings.append(f"COD {cod} belongs to another certificate. If you continue, a new automatic COD {automatic} will be used.")
        cod = automatic
    conflicts = cod_conflicts(inventory, cod, allowed_paths, owned=owned and cod_identity(cod) == cod_identity(requested))
    if conflicts:
        names = ', '.join(Path(entry['path']).name for entry in conflicts[:3])
        warnings.append(f"COD {cod} is already recorded in certificate PDF history ({names}). If you continue, new automatic COD {automatic} will be used; existing PDFs will not be overwritten.")
        cod = automatic
    warnings.extend(reference_warnings(inventory, details.get('reference'), allowed_paths, owned=owned))
    for entry in inventory['entries']:
        if entry['present'] and not entry['cods']:
            warnings.append(f"{entry['path']}: its COD could not be checked. Confirm the next COD manually before continuing.")
    try:
        _safe_cod_name(cod)
    except ValueError as exc:
        warnings.append(f"{exc} If you continue, automatic COD {automatic} will be used.")
        cod = automatic
    receipt_date = parse_date(details["inward_date"])
    issued_date = parse_date(details["cod_date"]) or as_of
    if details["cod_date"] and parse_date(details["cod_date"]) is None:
        warnings.append("Certificate date is invalid; today's date will be used if you continue.")
    if issued_date > as_of:
        warnings.append('COD Date is in the future; check this date before creating the certificate')
    if not details["supplier"] and not details["certificate_fields"].get("company_display"):
        warnings.append("Company name is missing; the company line will be blank.")
    if kind != "siemens" and receipt_date is None:
        warnings.append("Material receipt date is missing or invalid; it will be blank on the certificate.")
    if receipt_date and receipt_date > as_of:
        warnings.append("Material receipt date is in the future; verify it before sending.")
    certificate_weight = number(details["certificate_weight"])
    if source_kind == "manual" and "items" not in details:
        items, material_issues = [], []
        inward_weight = number(details.get("weight")) or 0
    else:
        items, material_issues = material_items(details)
        inward_weight = sum(item["weight"] or 0 for item in items)
    for issue in material_issues:
        warnings.append(f"{issue.capitalize()}; check the material list and weight before sending.")
    display_weight = certificate_weight if certificate_weight is not None else inward_weight
    if display_weight is None or display_weight <= 0:
        warnings.append("Weight is missing, invalid or not positive; check the weight shown in the preview before sending.")
    if details["certificate_weight"] and certificate_weight is None:
        warnings.append("Certificate/net weight is invalid; inward weight will be used if available.")
    if kind == "standard" and not details["form6"]:
        warnings.append("Form 6 / manifest number is missing; that value will be blank.")
    if kind == "siemens":
        for key, label in (("model_name", "Model name"), ("model_serial", "Model serial number"),
                           ("tube_serial", "X-ray tube serial number"), ("site_location", "Site location")):
            if not details["certificate_fields"].get(key, "").strip():
                warnings.append(f"{label} is missing; enter N/A if it does not apply, or continue with a blank value.")
    batch = dict(details, case_id=details["case_id"] or record_id or "Custom", source_id=plant, items=items,
                 inward_date=receipt_date, certificate_weight=certificate_weight, total_weight=inward_weight,
                 display_weight=display_weight, _data_root=str(root))
    warnings.extend(layout_warnings(batch, cod, issued_date))
    return batch, cod, issued_date, list(dict.fromkeys(warnings))


def current_record_pdf(root, record):
    """Check the current revision, not a superseded same-COD certificate."""
    filename = record.get("pdf_filename")
    if filename:
        relative = record.get('pdf_relative_path') or filename
        if not isinstance(relative, str):
            return None
        try:
            path = safe_certificate_path(root, relative)
        except ValueError:
            return None
        receipt = latest_issuance(root, 'manual', record['plant'], record['id'])
        healthy = valid_pdf(path)
        if receipt and receipt.get('pdf_path') == path.relative_to(Path(root)).as_posix():
            healthy = healthy and file_digest(path) == receipt['pdf_sha256']
        elif receipt:
            healthy = False
        return path if healthy else None
    return find_certificate(Path(root) / "Certificates", dict(record["details"], source_id=record["plant"]))


def _manual_status(value):
    """Use canonical status names for custom certificates, retaining local drafts."""
    cleaned = "".join(character for character in str(value or "").lower() if character.isalnum())
    if cleaned == "draft":
        return "draft"
    state = status_value(value)
    if state == "generated":
        return "generated_manual"
    return state


def manual_rows(root):
    data = read_register(root)
    rows = []
    for record in data["records"]:
        details = record["details"]
        materials, _ = material_items(details)
        status = _manual_status(record.get("status", "ready"))
        batch = dict(details, source_id=record["plant"])
        missing = status in {"generated_manual", "completed"} and not current_record_pdf(root, record)
        if missing:
            status = "ready"
        weight = number(details.get("certificate_weight"))
        if weight is None:
            weight = (sum(item["weight"] or 0 for item in materials)
                      if materials else number(details.get("weight")))
        rows.append({
            "source_kind": "manual", "record_id": record["id"], "plant": record["plant"],
            "raw_details": dict(details),
            "case_id": details.get("case_id", ""),
            "sn": details.get("case_id") or f"Custom {record['id'][:8]}", "supplier": details.get("supplier", ""),
            "location": details.get("location", ""), "certificate_type": details.get("certificate_type", "standard"),
            "certificate_fields": details.get("certificate_fields", {}), "form6": details.get("form6", ""),
            "reference": details.get("reference", ""), "inward_date": details.get("inward_date", ""),
            "certificate_no": details.get("cod_no", ""), "certificate_date": details.get("cod_date", ""),
            "certificate_weight": number(details.get("certificate_weight")), "display_weight": weight,
            "weight_source": "certificate" if number(details.get("certificate_weight")) is not None else "inward",
            "certificate_status": status, "outcome": "ready" if missing else status,
            "detail": f"PDF missing. Ready to recreate with COD {details.get('cod_no')}." if missing else
                      f"{TEMPLATE_NAMES.get(details.get('certificate_type'), 'Custom')} · Saved in the local certificate register.",
            "warnings": current_warnings(record.get("warnings", [])), "system_alert": missing,
            "items": materials,
            "file": "", "reminder_due": False, "overdue": False,
        })
    return rows


def reconcile_manual(root):
    if not register_path(root).is_file():
        return
    with run_lock(root):
        recover_issuances(root)
        data = read_register(root)
        changed = False
        for record in data["records"]:
            canonical = _manual_status(record.get("status", "ready"))
            if canonical != record.get("status"):
                record["status"] = canonical
                changed = True
            if canonical in {"generated_manual", "completed"} and not current_record_pdf(root, record):
                record["status"] = "ready"
                changed = True
        if changed:
            write_register(root, data)


def generate_draft(root, plant=None, details=None, record_id=None, confirmed=False, as_of=None,
                   override_reason='', expected_revision=None):
    with run_lock(root):
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        inventory = scan_inventory(root, persist=True)
        data = read_register(root)
        existing = _record(data, record_id)
        if existing and details is None:
            details, plant = existing["details"], existing["plant"]
        elif existing and plant is None:
            plant = existing["plant"]
        batch, cod, issued_date, warnings = prepare_certificate(root, plant, details, record_id=record_id, as_of=as_of)
        if warnings and not confirmed:
            row = {"plant": plant, "supplier": batch["supplier"], "sn": batch["case_id"],
                   "certificate_status": "ready", "outcome": "needs_review", "warnings": warnings,
                   "detail": "; ".join(warnings)}
            return {"requires_confirmation": True, "warnings": [row], "created": 0, "record_id": record_id}
        details = clean_details(details)
        details.update(cod_no=cod, cod_date=issued_date.isoformat())
        filename = certificate_filename(batch, int(plant[-1]), cod)
        folder = Path(root) / "Certificates"
        destination = certificate_destination(root, 'manual', filename)
        relative_path = destination.relative_to(folder).as_posix()
        temporary = Path(root) / "Data" / "Previews" / f".{uuid.uuid4().hex}.tmp.pdf"
        allowed_paths = issuance_paths(inventory, cod, 'manual', plant, record_id)
        owned = bool(allowed_paths)
        previous = current_record_pdf(root, existing) if existing else None
        previous_paths = issuance_paths(inventory, existing['details'].get('cod_no'), 'manual', plant, record_id) if existing else []
        # A saved filename or typed number is not permission to retire a PDF.
        # Only an unchanged, committed issuance can be replaced or archived.
        if previous and previous.relative_to(Path(root)).as_posix() not in previous_paths:
            previous = None
        if previous:
            allowed_paths.append(previous.relative_to(Path(root)).as_posix())
        transaction = None
        committed = False
        identifier = existing['id'] if existing else uuid.uuid4().hex
        source = register_path(root)
        source_before = source.read_bytes() if source.is_file() else None
        try:
            make_template_certificate(batch, cod, issued_date, temporary, allow_warnings=confirmed)
            if expected_revision is not None:
                require_revision(root, expected_revision)
            verify_publication(root, inventory, cod, batch.get('reference'), allowed_paths, owned=owned, allow_warnings=confirmed)
            transaction = IssuanceTransaction(
                root, destination, temporary, source, source_before, source_kind='manual',
                plant=plant, key=identifier, cod=cod, issued_date=issued_date,
                before=dict(existing or {}), after=details, warnings=warnings,
                override_reason=override_reason, old_paths=[previous] if previous else [],
                template_path=__import__('certificate_templates').TEMPLATE_ROOT / f"{batch['certificate_type']}.pdf")
            transaction.publish()
            if existing is None:
                existing = {"id": identifier, "created_at": datetime.now().isoformat()}
                data["records"].append(existing)
            existing.update(plant=plant, details=details, status="generated_manual", pdf_filename=filename,
                            pdf_relative_path=relative_path,
                            warnings=warnings, updated_at=datetime.now().isoformat())
            transaction.expect_source(json.dumps(data, indent=2, ensure_ascii=False, default=str).encode('utf-8'))
            write_register(root, data)
            committed = True
            try:
                transaction.mark_source_committed()
                transaction.complete()
            except (OSError, ValueError) as exc:
                warnings.append(f'Certificate saved; final history update needs recovery on refresh: {exc}')
        except Exception:
            if transaction and not committed:
                transaction.rollback()
            raise
        finally:
            temporary.unlink(missing_ok=True)
        # Failure to cache an already-committed issuance must not report it as
        # uncreated; the durable register still reserves its COD on the retry.
        try:
            record_issuance(root, destination, cod, 'manual', plant, existing['id'])
        except (OSError, ValueError) as exc:
            warnings.append(f"PDF created, but its folder inventory could not be refreshed: {exc}. Refresh the dashboard before the next certificate.")
        return {"requires_confirmation": False, "created": 1, "record_id": existing["id"],
                "certificate_no": cod, "filename": filename, "relative_path": relative_path,
                "warnings": warnings}


def complete_draft(root, record_id, expected_revision=None):
    with run_lock(root):
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        data = read_register(root)
        existing = _record(data, record_id)
        if not existing or _manual_status(existing.get("status")) not in {"generated_manual", "generated"}:
            raise ValueError("Create this certificate PDF before marking it Completed.")
        if not current_record_pdf(root, existing):
            raise ValueError("The certificate PDF is missing. Recreate it before marking it Completed.")
        existing.update(status="completed", updated_at=datetime.now().isoformat())
        write_register(root, data)
