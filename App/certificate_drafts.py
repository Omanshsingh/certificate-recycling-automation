"""Local drafts for certificates that have no inward workbook entry yet."""
from __future__ import annotations
from contextlib import ExitStack

from datetime import date, datetime
import json
from pathlib import Path
import shutil
import uuid
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

from certificate_core import (certificate_number, material_items, number, parse_date, read_batches,
                              status_value, supplier_is_always_not_applicable)
from certificate_templates import TEMPLATE_NAMES, layout_warnings, make_template_certificate, template_id, validate_fields, current_warnings
from run_certificates import _highest_sequence, _safe_cod_name, certificate_filename, find_certificate, cod_identity, numbering_plan
from certificate_inventory import (scan_inventory, cod_conflicts, verify_publication,
                                   issuance_paths, record_issuance, certificate_destination, safe_certificate_path,
                                   unresolved_cod_entries, cod_recovery_message)
from workflow_support import load_settings, operation_error_message, run_lock, save_json
from issuance_safety import (IssuanceTransaction, recover_issuances, latest_issuance,
                             current_pdf_is_healthy, manual_receipts, LOOKUP_RECEIPT)
from manual_ownership import record_plant, normalize_records


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
    from inward_serials import project_manual_sources
    return normalize_records(project_manual_sources(root, data))


def write_register(root, data):
    normalize_records(data)
    path = register_path(root)
    if path.exists():
        backups = Path(root) / "Data" / "Custom Register Backups"
        backups.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backups / f"Custom Certificates_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.json")
    save_json(path, data)


def migrate_local_ownership(root):
    """Back up legacy local plant assignments; keep issuance journals immutable."""
    with run_lock(root):
        recover_issuances(root)
        path = register_path(root)
        if path.is_file():
            original = json.loads(path.read_text(encoding='utf-8'))
            current = read_register(root)
            if current != original:
                write_register(root, current)


def clean_details(details):
    if not isinstance(details, dict):
        raise ValueError("Enter the certificate details first.")
    kind = template_id(details.get("certificate_type"))
    result = {"certificate_type": kind, "certificate_fields": validate_fields(kind, details.get("certificate_fields"))}
    from inward_followup import FIELDS, clean_followup
    if any(key in details for key in FIELDS):
        inward = clean_followup({key:details.get(key, '') for key in FIELDS})
        result.update({key:inward[key] for key in FIELDS if key in details})
    if kind == 'destruction' and details.get('destruction_source') is not None:
        link = details['destruction_source']
        if (not isinstance(link, dict) or (link.get('plant') not in {'plant1', 'plant2'} and not link.get('record_id'))
                or not isinstance(link.get('case_id'), str) or not link['case_id'].strip()
                or len(link['case_id']) > 100):
            raise ValueError('The linked inward is invalid. Reopen its data destruction request.')
        result['destruction_source'] = {'plant': link['plant'], 'case_id': link['case_id'].strip()}
        if link.get('record_id') is not None:
            if not isinstance(link['record_id'], str) or not __import__('re').fullmatch(r'[a-f0-9]{32}', link['record_id']):
                raise ValueError('The linked local certificate is invalid.')
            result['destruction_source']['record_id'] = link['record_id']
            result['destruction_source']['plant'] = None
    for key in ("case_id", "supplier", "location", "inward_date", "item", "quantity", "per_piece_weight",
                "weight", "certificate_weight", "form6", "reference"):
        value = str(details.get(key) if details.get(key) is not None else "").strip()
        if len(value) > 6000:
            raise ValueError(f"The {key} field is too long.")
        result[key] = value
    # Keep identifiers long enough for source-specific validation and issued
    # date corrections. Draft persistence and first issuance still enforce
    # automatic values in save_draft()/prepare_certificate().
    result["cod_no"] = str(details.get("cod_no") or "").strip()
    result["cod_date"] = str(details.get("cod_date") or "").strip()
    if "items" in details:
        items = details["items"]
        if not isinstance(items, list) or len(items) > 100:
            raise ValueError("Add no more than 100 materials at a time.")
        result["items"] = []
        for index, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise ValueError(f"Material {index} is not a valid row.")
            clean_item = {}
            for key, limit in (("item", 500), ("quantity", 100), ("per_piece_weight", 100), ("weight", 100)):
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


def _manual_record_issued(root, record):
    return bool(record and (
        _manual_status(record.get('status')) in {'generated_manual', 'completed'}
        or latest_issuance(root, 'manual', record['plant'], record['id'])))


def _local_plant(record, details, supplied):
    plant = record_plant(record or {'details':details})
    if plant is not None and supplied not in (None, plant):
        raise ValueError('Keep the plant of the linked inward for this data destruction request.')
    return plant


def _validate_destruction_record(record, details, plant):
    if record:
        previous = record['details']
        if (previous.get('certificate_type') == 'destruction') != (details['certificate_type'] == 'destruction'):
            raise ValueError('Data destruction needs its own certificate record. Create a new certificate instead of replacing this one.')
        link = previous.get('destruction_source')
        if link:
            if details.get('destruction_source') not in (None, link):
                raise ValueError('The linked inward cannot be changed. Create a new request for another inward.')
            details['destruction_source'] = dict(link)
    link = details.get('destruction_source')
    if link and details['certificate_type'] != 'destruction':
        raise ValueError('The linked inward belongs to a separate Data Destruction certificate.')
    if link and not link.get('record_id') and link['plant'] != plant:
        raise ValueError('Keep the plant of the linked inward for this data destruction request.')


def save_draft(root, plant, details, record_id=None, expected_revision=None):
    details = clean_details(details)
    with run_lock(root):
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        return _save_draft_locked(root, plant, details, read_register(root), record_id)


def _save_draft_locked(root, plant, details, data, record_id=None):
    if supplier_is_always_not_applicable(details.get("supplier")):
        raise ValueError("Reverse E-Commerce Pvt. Ltd. (Book My Junk) is always N/A and does not need a certificate.")
    existing = _record(data, record_id)
    plant = _local_plant(existing, details, plant)
    _validate_destruction_record(existing, details, plant)
    if ((details.get('case_id') or details.get('form6'))
            and not details.get('destruction_source')):
        from identifier_rules import require_unique, workspace_batches
        require_unique(details, plant, 'manual', record_id, workspace_batches(root), data['records'])
    if not existing and details.get('destruction_source'):
        linked = next((record for record in data['records']
                       if record['details'].get('destruction_source') == details['destruction_source']), None)
        if linked:
            return linked['id']
    if existing:
        details["cod_no"] = existing["details"].get("cod_no", "")
        issued = _manual_record_issued(root, existing)
        submitted_date = details.get('cod_date')
        if issued and submitted_date:
            parsed_date = parse_date(submitted_date)
            if parsed_date is None:
                raise ValueError('Enter a valid certificate date.')
            details['cod_date'] = parsed_date.isoformat()
        else:
            details["cod_date"] = existing["details"].get("cod_date", "")
    else:
        details["cod_no"] = ""
        details["cod_date"] = ""
    if existing:
        # Keep issuance identifiers stable across corrections/recreation.
        status = 'na' if _manual_status(existing.get('status')) == 'na' else 'ready'
        existing.update(plant=plant, details=details, status=status, updated_at=datetime.now().isoformat())
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


def prepare_certificate(root, plant, details, *, record_id=None, as_of=None, source_kind="manual",
                        starting_cod=None):
    """Preview/preflight without saving a draft, consuming a COD or changing Excel."""
    if source_kind == 'workbook' and plant not in {"plant1", "plant2"}:
        raise ValueError("Choose Plant 1 or Plant 2.")
    details = clean_details(details)
    if supplier_is_always_not_applicable(details.get("supplier")):
        raise ValueError("Reverse E-Commerce Pvt. Ltd. (Book My Junk) is always N/A and does not need a certificate.")
    data = read_register(root)
    existing = _record(data, record_id)
    if existing and _manual_status(existing.get('status')) == 'na':
        raise ValueError('This certificate is marked N/A. Choose Undo N/A before creating a PDF.')
    if source_kind == 'manual':
        plant = _local_plant(existing, details, plant)
    _validate_destruction_record(existing, details, plant)
    if existing:
        # The COD never changes. An issued certificate may receive an
        # explicitly corrected date; unissued drafts still use the automatic date.
        details["cod_no"] = existing["details"].get("cod_no", "")
        issued = _manual_record_issued(root, existing)
        if issued:
            details["cod_date"] = details.get("cod_date") or existing["details"].get("cod_date", "")
        else:
            details["cod_date"] = ""
    elif source_kind == "manual":
        details["cod_no"] = ""
        details["cod_date"] = ""
    as_of = as_of or date.today()
    kind = details["certificate_type"]
    settings = load_settings(root)
    batches, warnings = numbering_sources(root, settings)
    from identifier_rules import require_unique
    require_unique(details, plant, source_kind,
                   details['case_id'] if source_kind == 'workbook' else record_id, batches, data['records'])
    inventory = scan_inventory(root)
    numbering = numbering_plan(batches, settings, Path(root) / 'Certificates',
                               inventory=inventory, starting_cod=starting_cod)
    automatic = numbering['next_cod']
    requested = details["cod_no"] or (existing["details"].get("cod_no", "") if existing else "")
    if not requested and numbering['needs_start'] and starting_cod is None:
        raise ValueError('Enter the starting COD number before previewing or creating this certificate.')
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
        if existing and _manual_record_issued(root, existing):
            raise ValueError(f'COD {cod} is already assigned to another certificate. Correct the duplicate before reissuing; an issued COD cannot be changed automatically.')
        warnings.append(f"COD {cod} belongs to another certificate. If you continue, a new automatic COD {automatic} will be used.")
        cod = automatic
    conflicts = cod_conflicts(inventory, cod, allowed_paths, owned=owned and cod_identity(cod) == cod_identity(requested))
    if conflicts:
        if existing and _manual_record_issued(root, existing):
            raise ValueError(f'COD {cod} is already used by another PDF. Correct the duplicate before reissuing this certificate.')
        names = ', '.join(Path(entry['path']).name for entry in conflicts[:3])
        warnings.append(f"COD {cod} is already recorded in certificate PDF history ({names}). If you continue, new automatic COD {automatic} will be used; existing PDFs will not be overwritten.")
        cod = automatic
    warnings.extend(cod_recovery_message(entry['path']) for entry in unresolved_cod_entries(inventory))
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
    if kind not in {"siemens", "destruction"} and receipt_date is None:
        warnings.append("Material receipt date is missing or invalid; it will be blank on the certificate.")
    if receipt_date and receipt_date > as_of:
        warnings.append("Material receipt date is in the future; verify it before sending.")
    certificate_weight = number(details["certificate_weight"])
    if kind == 'destruction':
        items, material_issues, inward_weight = [], [], 0
    elif source_kind == "manual" and "items" not in details:
        items, material_issues = [], []
        inward_weight = number(details.get("weight")) or 0
    else:
        items, material_issues = material_items(details)
        inward_weight = sum(item["weight"] or 0 for item in items)
    for issue in material_issues:
        warnings.append(f"{issue.capitalize()}; check the material list and weight before sending.")
    display_weight = certificate_weight if certificate_weight is not None else inward_weight
    if kind != 'destruction' and (display_weight is None or display_weight <= 0):
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
    if kind == 'destruction':
        if source_kind != 'manual':
            raise ValueError('Use Data destruction beside the inward to create an additional certificate.')
        from destruction_certificate import content_issues
        warnings.extend(content_issues(details, as_of))
    batch = dict(details, case_id=details["case_id"] or record_id or "Custom", source_id=plant, items=items,
                 inward_date=receipt_date, certificate_weight=certificate_weight, total_weight=inward_weight,
                 display_weight=display_weight, _data_root=str(root),
                 _numbering={key: value for key, value in numbering.items() if key != 'sequence'})
    warnings.extend(layout_warnings(batch, cod, issued_date))
    return batch, cod, issued_date, list(dict.fromkeys(warnings))


def current_record_pdf(root, record, *, receipt=LOOKUP_RECEIPT):
    """Check the current revision, not a superseded same-COD certificate."""
    filename = record.get("pdf_filename")
    if receipt is LOOKUP_RECEIPT:
        receipt = latest_issuance(root, 'manual', record['plant'], record['id'])
    if not filename and not receipt:
        return find_certificate(Path(root) / "Certificates", dict(record["details"], source_id=record["plant"]))
    if filename:
        relative = record.get('pdf_relative_path') or filename
    else:
        recorded_path = Path(receipt.get('pdf_path', ''))
        if recorded_path.parts[:1] != ('Certificates',) or len(recorded_path.parts) < 2:
            return None
        relative = Path(*recorded_path.parts[1:]).as_posix()
    if not isinstance(relative, str):
        return None
    try:
        path = safe_certificate_path(root, relative, expected_sha=receipt.get('pdf_sha256') if receipt else None)
        receipt_path = path
        if receipt:
            recorded_path = Path(receipt.get('pdf_path', ''))
            if recorded_path.parts[:1] != ('Certificates',) or len(recorded_path.parts) < 2:
                return None
            recorded_relative = Path(*recorded_path.parts[1:]).as_posix()
            if recorded_relative != relative:
                receipt_path = safe_certificate_path(root, recorded_relative, expected_sha=receipt.get('pdf_sha256'))
    except ValueError:
        return None
    return path if receipt_path == path and current_pdf_is_healthy(path, receipt['pdf_sha256'] if receipt else None) else None


def _manual_status(value):
    """Use canonical status names for custom certificates, retaining local drafts."""
    cleaned = "".join(character for character in str(value or "").lower() if character.isalnum())
    if cleaned == "draft":
        return "draft"
    state = status_value(value)
    if state == "generated":
        return "generated_manual"
    return state


def manual_rows(root, as_of=None, records=None, *, history=None):
    as_of = as_of or date.today()
    records = read_register(root)['records'] if records is None else records
    receipts = manual_receipts(root, records, history=history)
    rows = []
    for record in records:
        details = record["details"]
        materials, _ = material_items(details)
        status = _manual_status(record.get("status", "ready"))
        receipt = receipts.get(record['id'])
        has_issued_certificate = (status in {"generated_manual", "completed"}
                                  or bool(record.get('pdf_recovery_pending'))
                                  or bool(receipt))
        requested = record.get('destruction_requested', True)
        path = (current_record_pdf(root, record, receipt=receipt)
                if has_issued_certificate and (status != 'na' or not requested) else None)
        missing = status != 'na' and has_issued_certificate and not path
        if not requested and not path:
            continue
        if missing:
            status = "ready"
        elif status == 'generated_manual' and receipt and not receipt.get('external_upload'):
            # Local register ownership is independent of who made the PDF.
            status = 'generated'
        received = parse_date(details.get('inward_date'))
        age = (as_of - received).days if received else None
        due = status not in {'completed', 'na'} and age is not None and age >= 25
        pending_hdd = requested and details.get('certificate_type') == 'destruction' and status not in {'completed', 'na'}
        weight = number(details.get("certificate_weight"))
        if weight is None:
            weight = (sum(item["weight"] or 0 for item in materials)
                      if materials else number(details.get("weight")))
        rows.append({
            "source_kind": "manual", "record_id": record["id"], "plant": record["plant"],
            'removed_inward': record.get('removed_inward'),
            'detached_inward': record.get('detached_inward'),
            "destruction_source": details.get('destruction_source'),
            "destruction_requested": requested,
            "destruction_pending": pending_hdd,
            "raw_details": dict(details),
            "case_id": details.get("case_id", ""),
            "sn": details.get("case_id") or f"Custom {record['id'][:8]}", "supplier": details.get("supplier", ""),
            "location": details.get("location", ""), "certificate_type": details.get("certificate_type", "standard"),
            "certificate_fields": details.get("certificate_fields", {}), "form6": details.get("form6", ""),
            "reference": details.get("reference", ""), "inward_date": details.get("inward_date", ""),
            "certificate_no": details.get("cod_no", ""), "certificate_date": details.get("cod_date", ""),
            "has_issued_certificate": has_issued_certificate,
            "certificate_weight": number(details.get("certificate_weight")), "display_weight": weight,
            "weight_source": "certificate" if number(details.get("certificate_weight")) is not None else "inward",
            "certificate_status": status, "outcome": "ready" if missing else status,
            "detail": f"PDF missing. Ready to recreate with COD {details.get('cod_no')}." if missing else
                      f"{TEMPLATE_NAMES.get(details.get('certificate_type'), 'Custom')} · Saved in the local certificate register.",
            "warnings": current_warnings(record.get("warnings", [])),
            "system_alert": missing,
            "items": materials,
            "file": "", "reminder_due": due, "overdue": bool(due and age > 30),
            "age_days": age, "pdf_missing": missing, "attention_pending": pending_hdd,
        })
    return rows


def reconcile_manual(root):
    if not register_path(root).is_file():
        return
    with run_lock(root):
        recover_issuances(root)
        data = read_register(root)
        receipts = manual_receipts(root, data['records'])
        changed = False
        for record in data["records"]:
            canonical = _manual_status(record.get("status", "ready"))
            if canonical != record.get("status"):
                record["status"] = canonical
                changed = True
            if canonical in {"generated_manual", "completed"} and not current_record_pdf(
                    root, record, receipt=receipts.get(record['id'])):
                record["status"] = "ready"
                record['pdf_recovery_pending'] = True
                changed = True
        if changed:
            write_register(root, data)


def generate_draft(root, plant=None, details=None, record_id=None, confirmed=False, as_of=None,
                   override_reason='', expected_revision=None, starting_cod=None):
    with run_lock(root), ExitStack() as resources:
        recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        # Keep review inputs unchanged until the certificate commits. Updating
        # the ledger here would invalidate the revision checked below.
        inventory = scan_inventory(root)
        unresolved = unresolved_cod_entries(inventory)
        if unresolved:
            messages = [cod_recovery_message(entry['path']) for entry in unresolved[:3]]
            if len(unresolved) > 3:
                messages.append(f"There are {len(unresolved) - 3} more PDFs to check.")
            raise ValueError(' '.join(messages))
        data = read_register(root)
        existing = _record(data, record_id)
        if existing and details is None:
            details, plant = existing["details"], existing["plant"]
        elif existing and plant is None:
            plant = existing["plant"]
        candidate = clean_details(details)
        plant = _local_plant(existing, candidate, plant)
        _validate_destruction_record(existing, candidate, plant)
        if not existing and candidate.get('destruction_source'):
            if any(record['details'].get('destruction_source') == candidate['destruction_source']
                   for record in data['records']):
                raise ValueError('This inward already has a Data Destruction record. Open that record to continue.')
        if candidate['certificate_type'] != 'destruction' and not candidate.get('reference'):
            raise ValueError('Enter the reference number before creating the certificate.')
        details = candidate
        if candidate['certificate_type'] == 'destruction':
            from destruction_certificate import content_issues
            issues = content_issues(candidate, as_of)
            if issues:
                raise ValueError(' '.join(issues))
        batch, cod, issued_date, warnings = prepare_certificate(root, plant, details, record_id=record_id,
                                                               as_of=as_of, starting_cod=starting_cod)
        if warnings and not confirmed:
            row = {"plant": plant, "supplier": batch["supplier"], "sn": batch["case_id"],
                   "certificate_status": "ready", "outcome": "needs_review", "warnings": warnings,
                   "detail": "; ".join(warnings)}
            return {"requires_confirmation": True, "warnings": [row], "created": 0, "record_id": record_id}
        details = clean_details(details)
        details.update(cod_no=cod, cod_date=issued_date.isoformat())
        linked_session = None
        linked_data = None
        link = details.get('destruction_source') or {}
        if details['certificate_type'] == 'destruction' and link and not link.get('record_id'):
            from workbook_store import WorkbookSession
            from hdd_workbook import stage_hdd_tracking
            configured = load_settings(root)['workbook_paths'][int(link['plant'][-1]) - 1]
            if not configured:
                raise ValueError('Connect the linked plant workbook before creating its HDD certificate.')
            linked_session = resources.enter_context(WorkbookSession(Path(configured).expanduser().resolve(), Path(root) / 'Data' / 'Workbook Backups'))
            if linked_session.context is None:
                raise ValueError('Close the linked plant workbook in Excel before saving its HDD number and date.')
            linked_data = stage_hdd_tracking(linked_session, link['case_id'], cod, issued_date)
        filename = certificate_filename(batch, cod=cod)
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
        if previous and previous.relative_to(Path(root)).as_posix().casefold() not in {path.casefold() for path in previous_paths}:
            previous = None
        if previous:
            allowed_paths.append(previous.relative_to(Path(root)).as_posix())
        transaction = None
        committed = False
        recovery_note = ''
        identifier = existing['id'] if existing else uuid.uuid4().hex
        source = register_path(root)
        source_before = source.read_bytes() if source.is_file() else None
        try:
            make_template_certificate(batch, cod, issued_date, temporary, allow_warnings=confirmed)
            if expected_revision is not None:
                # The linked workbook is exclusively locked by this save on
                # Windows. Compare its reviewed bytes without reopening it.
                overrides = {str(linked_session.path.resolve()): linked_session.data} if linked_session else None
                require_revision(root, expected_revision, overrides)
            verify_publication(root, inventory, cod, allowed_paths=allowed_paths,
                               owned=owned, allow_warnings=confirmed)
            before_record = dict(existing or {})
            if existing is None:
                existing = {"id": identifier, "created_at": datetime.now().isoformat()}
                data["records"].append(existing)
            existing.update(plant=plant, details=details, status="generated_manual", pdf_filename=filename,
                            pdf_relative_path=relative_path, pdf_recovery_pending=False,
                            warnings=warnings, updated_at=datetime.now().isoformat())
            register_payload = json.dumps(data, indent=2, ensure_ascii=False, default=str).encode('utf-8')
            transaction = IssuanceTransaction(
                root, destination, temporary, linked_session.path if linked_session else source,
                linked_session.data if linked_session else source_before, source_kind='manual',
                plant=plant, key=identifier, cod=cod, issued_date=issued_date,
                before=before_record, after=details, warnings=warnings,
                override_reason=override_reason, old_paths=[previous] if previous else [],
                companion_source=(source, source_before, register_payload) if linked_session else None,
                template_path=__import__('certificate_templates').TEMPLATE_ROOT / f"{batch['certificate_type']}.pdf")
            transaction.publish()
            if linked_session:
                transaction.sync_companion(True)
                linked_session.before_persist = transaction.expect_source
                linked_session._persist(linked_data)
            else:
                transaction.expect_source(register_payload)
                write_register(root, data)
            committed = True
            try:
                transaction.mark_source_committed()
                transaction.complete()
            except (OSError, ValueError) as exc:
                recovery_note = operation_error_message(exc)
                warnings.append(f'Certificate saved; final history update needs recovery on refresh: {recovery_note}')
        except Exception:
            if transaction and not committed:
                transaction.rollback()
            raise
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError as exc:
                if committed:
                    warnings.append(f'Certificate saved; its temporary preview file could not be removed: {exc}.')
        return {"requires_confirmation": False, "created": 1, "record_id": existing["id"],
                "certificate_no": cod, "filename": filename, "relative_path": relative_path,
                "warnings": warnings, "recovery_note": recovery_note, "workbook_updated": linked_session is not None}


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
