"""Loopback-only browser interface for the Ecoreco certificate workflow."""
from __future__ import annotations

from datetime import datetime
from http.client import HTTPConnection
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import argparse
import os
from pathlib import Path
import subprocess
import threading
import time
from urllib.parse import urlsplit, unquote, quote
import webbrowser
import re
import uuid
import shutil

from app_paths import DATA_ROOT
from certificate_core import batch_content, inward_details_complete, number, parse_date, read_batches, source_value_issues, status_value
from run_certificates import run, find_certificate
from workbook_store import update_issued_status, WorkbookSession, stage_certificate_details
from workflow_support import load_settings, open_excel_record, save_json
from certificate_drafts import (clean_details, save_draft, generate_draft, complete_draft,
                                prepare_certificate, manual_rows, reconcile_manual,
                                read_register, current_record_pdf)
from certificate_templates import template_specs, make_template_certificate, obsolete_template_warning, display_values
from pypdf import PdfReader
from certificate_inventory import scan_inventory, safe_certificate_path, ensure_certificate_folders
from workflow_support import run_lock
from issuance_safety import recover_issuances, valid_pdf, file_digest, latest_issuance, audit_records
from workspace_review import workspace_revision, require_revision, reviewed_operation
from backup_service import create_backup, list_backups, restore_backup
from dashboard_review import DashboardReview
from client_intake import intake_details, stage_new_client


ROOT = Path(DATA_ROOT).resolve()
WEB_ROOT = Path(__file__).resolve().parent / "web"
BACKUP_DIR = ROOT / "Data" / "Workbook Backups"
RESULTS_PATH = ROOT / "Data" / "Reports" / "Latest Results.json"
SERVER: ThreadingHTTPServer | None = None
OPERATION_LOCK = threading.RLock()
STOPPING = False
WORKBOOK_CHECK_STAMPS = {}
DASHBOARD_REVIEW = DashboardReview()


def workbook_check_stamp(paths):
    """Cheap saved-file change detection; Excel edits are checked after saving."""
    stamp = []
    for path in paths:
        stat = path.stat()
        stamp.append((str(path), stat.st_mtime_ns, stat.st_size))
    return tuple(stamp)


def check_workbook_changes():
    try:
        paths = configured_paths()
        stamp = workbook_check_stamp(paths)
    except (OSError, ValueError):
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        return {'changed': True, 'state': refresh()}
    if WORKBOOK_CHECK_STAMPS.get(str(ROOT)) == stamp:
        return {"changed": False}
    return {"changed": True, "state": refresh(paths)}


def configured_paths(settings=None, allow_unavailable=False):
    settings = settings or load_settings(ROOT)
    paths = [Path(value).expanduser().resolve() for value in settings["workbook_paths"]]
    if len(paths) != 2 or (not allow_unavailable and not all(path.is_file() for path in paths)):
        raise ValueError("Choose the Plant 1 and Plant 2 workbooks in Settings first.")
    if paths[0] == paths[1] or (paths[0].is_file() and paths[1].is_file() and paths[0].samefile(paths[1])):
        raise ValueError("Plant 1 and Plant 2 must use different workbooks.")
    return paths


def add_client(body):
    plant = body.get('plant')
    if plant not in {'plant1', 'plant2'}:
        raise ValueError('Choose Plant 1 or Plant 2.')
    expected = body.get('expected_revision')
    request_id = body.get('request_id')
    if not isinstance(request_id, str) or not re.fullmatch(r'[a-f0-9]{32}', request_id):
        raise ValueError('Reopen Add client and try again.')
    details = intake_details(body.get('details'))
    signature = json.loads(json.dumps(details, default=str))
    journal_path = ROOT / 'Data' / 'Client Additions' / f'{request_id}.json'
    paths = configured_paths(allow_unavailable=True)
    path = paths[int(plant[-1]) - 1]
    with run_lock(ROOT):
        journal = json.loads(journal_path.read_text(encoding='utf-8')) if journal_path.is_file() else None
        if journal and (journal.get('plant') != plant or journal.get('path') != str(path) or journal.get('input') != signature):
            raise ValueError('This add request already belongs to different details. Check the client in Excel before starting a new entry.')
        if not journal:
            require_revision(ROOT, expected)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path.resolve()):session.data}
            if journal:
                from io import BytesIO
                batch = next((row for row in read_batches(BytesIO(session.data)) if row['case_id'] == journal['sn']), None)
                recovered = batch and (journal.get('saved') or json.loads(json.dumps(batch_content(batch), default=str)) == journal['snapshot'])
                if not recovered and (batch or journal.get('saved')):
                    raise ValueError('This client was already added and later changed or removed. Check Excel; it has not been added again.')
            else:
                recovered = False
            if not recovered:
                require_revision(ROOT, journal.get('expected_revision', expected) if journal else expected, overrides)
                updated, batch = stage_new_client(session, details)
                journal = {'plant':plant, 'path':str(path), 'input':signature, 'sn':batch['case_id'],
                           'snapshot':json.loads(json.dumps(batch_content(batch), default=str)),
                           'expected_revision':expected, 'saved':False}
                session.before_persist = lambda _: require_revision(ROOT, expected, overrides)
                save_json(journal_path, journal)
                session._persist(updated)
                journal['saved'] = True
                try:
                    save_json(journal_path, journal)
                except OSError:
                    pass  # The prepared record verifies a committed add if its response is lost.
    try:
        state = refresh()
    except Exception as exc:
        try:
            state = latest_result()
        except Exception:
            state = {'rows':[], 'scan_complete':False}
        state = dict(state, setup_warning=f'Client added to Excel, but the dashboard could not refresh: {exc}. Refresh before adding it again.')
    return {'client_added':True, 'client':{'plant':plant, 'sn':batch['case_id'],
            'supplier':batch['supplier'], 'status':status_value(batch['issued'])}, 'state':state}


def latest_result():
    if not RESULTS_PATH.is_file():
        return {"rows": [], "scan_complete": False, "created": 0, "review": 0}
    value = json.loads(RESULTS_PATH.read_text(encoding="utf-8"))
    if not isinstance(value, dict) or not isinstance(value.get("rows"), list):
        raise ValueError("The latest work list is unreadable. Refresh the workbooks to rebuild it.")
    return value


def refresh(paths=None):
    settings = load_settings(ROOT)
    reconcile_manual(ROOT)
    with run_lock(ROOT):
        recover_issuances(ROOT)
        inventory = scan_inventory(ROOT, persist=True)
    try:
        paths = paths or configured_paths(settings, allow_unavailable=True)
    except ValueError as exc:
        return {"rows": current_pdf_rows(manual_rows(ROOT)), "scan_complete": False, "created": 0,
                "review": 0, "setup_warning": str(exc), 'certificate_inventory': inventory}
    run(paths, ROOT, check_only=True, progress=None)
    state = workspace_state(recheck_obsolete=False)
    if state.get("scan_complete"):
        try:
            WORKBOOK_CHECK_STAMPS[str(ROOT)] = workbook_check_stamp(paths)
        except OSError:
            WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
            state['scan_complete'] = False
            state['setup_warning'] = 'An Excel file became unavailable while refreshing. Save and close it in Excel, check its path in Settings, then refresh again. No certificate was created.'
    return state


def workspace_state(recheck_obsolete=True):
    reconcile_manual(ROOT)
    report = latest_result()
    canonicalized = False
    for row in report.get("rows", []):
        status = status_value(row.get("certificate_status"))
        if row.get("source_kind") == "manual" and status == "generated":
            status = "generated_manual"
        if status != row.get("certificate_status"):
            row["certificate_status"] = status
            canonicalized = True
    if canonicalized:
        save_json(RESULTS_PATH, report)
    missing_pdf = any(
        row.get('certificate_status') in {'generated', 'generated_manual', 'completed'}
        and not find_certificate(ROOT / 'Certificates', {
            'cod_no': row.get('certificate_no'), 'source_id': row.get('plant'), 'case_id': row.get('sn')})
        for row in report.get('rows', []) if row.get('source_kind') != 'manual')
    if recheck_obsolete and (missing_pdf or any(
            obsolete_template_warning(row.get('detail', ''))
            or any(obsolete_template_warning(warning) for warning in (row.get('warnings') or []))
            for row in report.get('rows', []) if row.get('source_kind') != 'manual')):
        # Re-evaluate real data instead of just hiding the warning and leaving
        # stale attention flags (or incorrectly declaring the row valid).
        # Missing PDFs also invalidate cached Generated/Completed states.
        # Refresh is check-only; it cannot generate a certificate.
        return refresh()
    report['certificate_inventory'] = scan_inventory(ROOT)
    report["rows"] = current_pdf_rows(
        [row for row in report.get("rows", []) if row.get("source_kind") != "manual"] + manual_rows(ROOT))
    return report


def current_pdf_rows(rows):
    """Add validated current filenames to response rows without saving tracking."""
    issued = {'generated', 'generated_manual', 'completed'}
    records = {record['id']: record for record in read_register(ROOT)['records']} if any(
        row.get('source_kind') == 'manual' and status_value(row.get('certificate_status')) in issued
        for row in rows) else {}
    folder = (ROOT / 'Certificates').resolve()
    receipts = {(entry.get('source_kind'), entry.get('plant'), str(entry.get('key'))): entry
                for entry in audit_records(ROOT) if entry.get('state') == 'complete'}
    result = []
    for original in rows:
        row = dict(original)
        row.pop('pdf_filename', None)  # Cached filenames are not current-file evidence.
        row.pop('pdf_relative_path', None)
        row.pop('has_issued_certificate', None)
        row.pop('last_certificate_details', None)
        source_kind = row.get('source_kind', 'workbook')
        key = row.get('record_id') if source_kind == 'manual' else row.get('sn')
        receipt = receipts.get((source_kind, row.get('plant'), str(key)))
        if receipt or status_value(row.get('certificate_status')) in issued:
            row['has_issued_certificate'] = True
        if source_kind == 'workbook' and receipt and isinstance(receipt.get('after'), dict):
            row['last_certificate_details'] = editor_details(receipt['after'])
        result.append(row)
        if status_value(row.get('certificate_status')) not in issued:
            continue
        source_kind = row.get('source_kind', 'workbook')
        key = row.get('record_id') if source_kind == 'manual' else row.get('sn')
        try:
            if source_kind == 'manual':
                record = records.get(key)
                if not record or record.get('plant') != row.get('plant'):
                    continue
                path = current_record_pdf(ROOT, record)
            else:
                path = find_certificate(ROOT / 'Certificates', {
                    'cod_no': row.get('certificate_no'), 'source_id': row.get('plant'), 'case_id': key})
            if not isinstance(path, (str, os.PathLike)):
                continue
            path = Path(path)
            current = path.resolve(strict=True)
            if (path.suffix.casefold() != '.pdf' or not current.is_relative_to(folder)
                    or current.name != path.name or not valid_pdf(current)):
                continue
            relative = current.relative_to(folder).as_posix()
            if safe_certificate_path(ROOT, relative) != current:
                continue
            if receipt and (current != (ROOT / receipt['pdf_path']).resolve()
                            or file_digest(current) != receipt.get('pdf_sha256')):
                continue
            row['pdf_filename'] = path.name
            row['pdf_relative_path'] = current.relative_to(folder).as_posix()
        except (OSError, ValueError):
            # Unavailable metadata must not turn a stale filename into an Open PDF link.
            continue
    return result


def editor_details(batch):
    """Only form values from the issued snapshot, retaining decimal precision."""
    details = {key: batch.get(key, '') for key in (
        'case_id', 'supplier', 'location', 'inward_date', 'form6', 'reference',
        'certificate_weight', 'cod_no', 'cod_date', 'certificate_type', 'certificate_fields')}
    for key in ('inward_date', 'cod_date'):
        parsed = parse_date(details[key])
        if parsed:
            details[key] = parsed.isoformat()
    details['items'] = [{key: item.get(key, '') for key in ('item', 'quantity', 'per_piece_weight', 'weight')}
                        for item in batch.get('items', [])]
    if details['items']:
        details.update(details['items'][0])
    return clean_details(details)


def workbook_form(body, details, paths):
    """First issuance uses current Excel facts; only template extras are editable."""
    selected = selected_key(body.get('plant'), details.get('case_id'))
    if body.get('expected_revision') is not None:
        require_revision(ROOT, body['expected_revision'])
    if body.get('source_sn') is not None and (
            str(body['source_sn']) != selected[1] or body.get('source_plant') != selected[0]):
        raise ValueError('The selected inward ID and plant cannot be changed. Open the correct inward instead.')
    source = next((batch for batch in read_batches(paths[int(selected[0][-1]) - 1])
                   if str(batch['case_id']) == selected[1]), None)
    if body.get('expected_revision') is not None:
        require_revision(ROOT, body['expected_revision'])
    if source is None:
        if body.get('source_sn') is not None:
            raise ValueError('This inward no longer exists. Refresh the company list before continuing.')
        return details, body.get('edit_scope', 'update_excel')
    if body.get('create_new'):
        raise ValueError('This Excel ID already exists in this plant. Choose the company from the list to edit it, or enter a different ID.')
    receipt = latest_issuance(ROOT, 'workbook', *selected)
    pdf_only_edits = bool(body.get('pdf_only_edits'))
    if pdf_only_edits and body.get('edit_scope', 'certificate_only') != 'certificate_only':
        raise ValueError('First-generation corrections can change the PDF only; Excel inward details stay unchanged.')
    if receipt or status_value(source.get('issued')) in {'generated', 'generated_manual', 'completed'}:
        if pdf_only_edits:
            raise ValueError('Refresh the certificate before editing its PDF.')
        if not details.get('cod_no'):
            details['cod_no'] = source.get('cod_no') or (receipt or {}).get('cod', '')
        if not details.get('cod_date'):
            saved_date = source.get('cod_date') or (receipt or {}).get('certificate_date') or ''
            parsed = parse_date(saved_date)
            details['cod_date'] = parsed.isoformat() if parsed else str(saved_date)
        return details, body.get('edit_scope', 'certificate_only')
    if source.get('blocking_issues'):
        raise ValueError(' '.join(source['blocking_issues']))
    source_state = status_value(source.get('issued'))
    kind = details['certificate_type']
    if source_state == 'blank':
        selected_source = dict(source, certificate_type=kind)
        if not inward_details_complete(selected_source):
            issues = source_value_issues(selected_source)
            message = ' '.join(issues) or 'Complete the company, received date and material weights in Excel.'
            raise ValueError(f'Excel details need checking: {message}')
    elif source_state != 'ready':
        raise ValueError('This inward is marked ' + (source.get('issued') or 'with an unknown status') + '. Choose Ready in Excel before creating its first certificate.')
    # Check normalized staged values so harmless date/number formatting is accepted.
    with WorkbookSession(paths[int(selected[0][-1]) - 1], BACKUP_DIR) as session:
        _, proposed = stage_certificate_details(session, details, validate_cells=False)
    protected = ('supplier', 'location', 'inward_date', 'form6', 'reference', 'cod_no', 'certificate_weight')
    for key in protected:
        if key in (body.get('details') or {}) and key in details and proposed.get(key) != source.get(key):
            if not pdf_only_edits or key == 'cod_no':
                raise ValueError('Excel details are locked for the first certificate. Correct the inward in Excel, or edit the certificate after generating it.')
    if 'cod_date' in (body.get('details') or {}) and parse_date(proposed.get('cod_date')) != parse_date(source.get('cod_date')):
        raise ValueError('The certificate date comes from Excel or is chosen automatically for first generation.')
    material_keys = ('item', 'quantity', 'per_piece_weight', 'weight')
    first_material = next(iter(source.get('items', [])), {})
    for key in material_keys:
        if key not in (body.get('details') or {}):
            continue
        proposed_value = details.get(key, '') if key == 'item' else number(details.get(key))
        original_value = first_material.get(key, '') if key == 'item' else first_material.get(key)
        if proposed_value != original_value:
            if not pdf_only_edits:
                raise ValueError('Materials and weights are locked to Excel for first generation.')
    materials = lambda batch: [tuple(item.get(key) for key in material_keys) for item in batch.get('items', [])]
    if materials(proposed) != materials(source) and not pdf_only_edits:
        raise ValueError('Materials and weights are locked to Excel for first generation.')
    fields = dict(details['certificate_fields'])
    saved_fields = source.get('certificate_fields') or {}
    original_company_line = saved_fields.get('company_display', '') if source.get('certificate_type') == 'standard' else ''
    if kind == 'standard' and str(fields.get('company_display', '')) != str(original_company_line) and not pdf_only_edits:
        raise ValueError('The company name is locked to Excel for first generation.')
    if 'weight_display' in fields and fields['weight_display'] != saved_fields.get('weight_display', '') and not pdf_only_edits:
        raise ValueError('The printed weight is locked to Excel for first generation.')
    if (kind == 'siemens' and str(source.get('location') or '').strip()
            and (not pdf_only_edits or not fields.get('site_location', '').strip())):
        fields['site_location'] = source['location']
    if (kind in {'sita', 'siemens'} and kind != source.get('certificate_type')
            and not fields.get('company_display', '').strip()
            and not (pdf_only_edits and 'company_display' in fields)):
        fields['company_display'] = next(spec['default_company'] for spec in template_specs() if spec['id'] == kind)
    canonical = editor_details(dict(source, certificate_type=kind, certificate_fields=fields))
    if pdf_only_edits:
        # Keep source workbooks immutable on first generation while allowing
        # the operator to correct the printed certificate itself.
        for key in ('supplier', 'location', 'inward_date', 'reference', 'form6'):
            if key in details:
                canonical[key] = details[key]
        if 'items' in details:
            canonical['items'] = details['items']
            canonical.update(details['items'][0] if details['items'] else {})
        if body.get('use_certificate_weight') and 'certificate_weight' in details:
            canonical['certificate_weight'] = details['certificate_weight']
    else:
        canonical.pop('certificate_weight', None)  # Leave even invalid/raw Excel weight cells intact.
    return canonical, 'certificate_only'


def preview_certificate(body):
    source_kind = body.get("source_kind", "workbook")
    if source_kind not in {"workbook", "manual"}:
        raise ValueError("Choose the plant workbook or local certificate register.")
    validate_certificate_weight_override(body, source_kind)
    details = clean_details(body.get("details"))
    if source_kind == "workbook" and not body.get("use_certificate_weight"):
        details.pop("certificate_weight", None)
    if source_kind == 'workbook':
        details, _ = workbook_form(body, details, configured_paths())
    batch, cod, issued_date, warnings = prepare_certificate(
        ROOT, body.get("plant"), details, record_id=body.get("record_id") if source_kind == "manual" else None,
        source_kind=source_kind,
    )
    if source_kind == "workbook":
        paths = configured_paths()
        # Use the same staged material rows as final generation. Preview never
        # writes them, and corrections should not inherit stale source warnings.
        with WorkbookSession(paths[int(body["plant"][-1]) - 1], BACKUP_DIR) as session:
            _, staged = stage_certificate_details(session, details, validate_cells=False)
        batch.update(total_weight=staged['total_weight'], display_weight=staged['display_weight'],
                     certificate_weight=staged.get('certificate_weight'))
        if batch['display_weight'] and batch['display_weight'] > 0:
            warnings = [warning for warning in warnings if not warning.startswith('Weight is missing, invalid or not positive;')]
        warnings.extend(issue for issue in staged.get('issues', [])
                        if not (batch['certificate_type'] in {'sita', 'siemens'}
                                and issue == 'missing Form 6 number / manifest number'))
    identifier = uuid.uuid4().hex
    destination = ROOT / "Data" / "Previews" / f"{identifier}.pdf"
    make_template_certificate(batch, cod, issued_date, destination, allow_warnings=True)
    import pypdfium2 as pdfium
    with pdfium.PdfDocument(str(destination)) as document:
        page = document[0]
        bitmap = page.render(scale=1.5)
        try:
            bitmap.to_pil().save(destination.with_suffix('.png'))
        finally:
            bitmap.close()
            page.close()
    printed = display_values(batch, cod, issued_date)
    return {"preview_url": f"/previews/{identifier}.pdf", 'preview_image_url': f'/previews/{identifier}.png', "certificate_no": cod,
            "certificate_date": issued_date.isoformat(),
            "printed_details": {key: printed[key] for key in ('company_display', 'certificate_weight')},
            "warnings": list(dict.fromkeys(warnings))}


def selected_key(plant, case_id):
    if plant not in {"plant1", "plant2"} or not str(case_id or "").strip():
        raise ValueError("Choose a plant and enter a Client ID / SN.")
    return plant, str(case_id).strip()


def validate_certificate_weight_override(body, source_kind=None):
    source_kind = source_kind or body.get("source_kind", "workbook")
    if source_kind == "workbook" and body.get("use_certificate_weight"):
        details = body.get("details") or {}
        weight = number(details.get("certificate_weight")) if isinstance(details, dict) else None
        if weight is None or weight <= 0:
            raise ValueError("Enter a positive Certificate Weight, or untick the option to use Excel's weight.")


def target_row(report, selected):
    return next((row for row in report.get("rows", [])
                 if row.get("plant") == selected[0]
                 and str(row.get("sn", "")) == selected[1]), None)


def open_selected_workbook_record(plant, case_id):
    """Resolve the chosen plant + client ID against the current workbook, then navigate to that row."""
    selected = selected_key(plant, case_id)
    row = target_row(latest_result(), selected)
    if not row or not row.get("file"):
        raise ValueError("Refresh the workbooks and select the row again.")
    paths = configured_paths()
    workbook = paths[int(plant[-1]) - 1]
    if Path(row["file"]).expanduser().resolve() != workbook:
        raise ValueError("The selected workbook changed. Refresh both plant workbooks and try again.")
    matches = [batch for batch in read_batches(workbook)
               if str(batch.get("case_id")) == str(case_id)
               and (not row.get("sheet") or batch.get("source_sheet") == row.get("sheet"))]
    if len(matches) != 1:
        raise ValueError("The selected client row changed in Excel. Refresh both plant workbooks and try again.")
    batch = matches[0]
    return open_excel_record(workbook, batch["source_sheet"], batch["source_row"],
                             case_id=batch['case_id'], header_row=batch.get('source_header_row',1))


def generate(*, selected=None, plants=None, custom_cod=None, confirmed=False, expected_revision=None, override_reason=''):
    settings = load_settings(ROOT)
    paths = configured_paths(settings)
    preflight = {}
    preflight_args = {"check_only": True, "progress": None, 'report_sink': preflight}
    if expected_revision is not None:
        preflight_args['expected_revision'] = expected_revision
    if selected:
        preflight_args["selected"] = selected
    if plants is not None:
        preflight_args["plants"] = plants
    if custom_cod:
        preflight_args["custom_cod"] = custom_cod
    run(paths, ROOT, **preflight_args)
    preflight = preflight or latest_result()
    if not preflight.get('scan_complete'):
        error = next((row.get('detail') for row in preflight.get('rows', []) if row.get('source_kind') == 'system'), None)
        raise ValueError(error or 'Both plant files must be checked before creating a certificate. Refresh and review the workbook message.')
    candidates = [row for row in preflight.get("rows", [])
                  if row.get("certificate_status") == "ready"
                  and row.get("outcome") == "needs_review"]
    if selected and target_row(preflight, selected) is None:
        raise ValueError(f"Client ID {selected[1]} was not found in {selected[0].replace('plant', 'Plant ')}.")
    if candidates and not confirmed:
        return {"requires_confirmation": True, "warnings": candidates,
                "state": refresh(paths)}

    generated_report = {}
    arguments = {"progress": None, "allow_warnings": confirmed, 'override_reason': override_reason,
                 'report_sink': generated_report}
    if expected_revision is not None:
        arguments['expected_revision'] = expected_revision
    if selected:
        arguments["selected"] = selected
    if plants is not None:
        arguments["plants"] = plants
    if custom_cod:
        arguments["custom_cod"] = custom_cod
    result = run(paths, ROOT, **arguments)
    generated_report = generated_report or latest_result()
    generated_rows = {
        (row.get("plant"), str(row.get("sn"))): row
        for row in generated_report.get("rows", [])
        if row.get("outcome") == "generated_now"
    }
    state = post_commit_state(lambda: refresh(paths))
    for row in state.get("rows", []):
        generated_row = generated_rows.get((row.get("plant"), str(row.get("sn"))))
        if generated_row:
            row["detail"] = generated_row.get("detail", row.get("detail", ""))
            row["warnings"] = generated_row.get("warnings", [])
            if generated_row.get("system_alert"):
                row["system_alert"] = True
    state["created"] = result[0]
    return {"requires_confirmation": False, "state": state,
            "created": result[0], "review": result[1],
            'warnings': [state['setup_warning']] if state.get('setup_warning') else [],
            'certificates': certificate_outputs(list(generated_rows.values()))}


def create_certificate(body):
    """Preflight form edits without writing Excel; commit only with the final PDF."""
    plant = body.get('plant')
    validate_certificate_weight_override(body, "workbook")
    details = clean_details(body.get('details'))
    if not body.get('use_certificate_weight'):
        details.pop('certificate_weight', None)
    selected = selected_key(plant, details.get('case_id'))
    scope = body.get('edit_scope', 'certificate_only')
    if scope not in {'certificate_only', 'update_excel'}:
        raise ValueError('Choose This certificate only or Also update Excel.')
    if body.get('source_sn') is not None:
        if str(body['source_sn']) != selected[1] or body.get('source_plant') != plant:
            raise ValueError('The selected inward ID and plant cannot be changed. Open the correct inward instead.')
    paths = configured_paths()
    if body.get('create_new'):
        # The HTTP request lock covers this read; run() rechecks its revision.
        workbook = paths[int(plant[-1]) - 1]
        if any(selected_key(plant, batch['case_id']) == selected
               for batch in read_batches(workbook)):
            raise ValueError('This Excel ID already exists in this plant. '
                             'Choose the company from the list to edit it, or enter a different ID.')
    details, scope = workbook_form(body, details, paths)
    snapshot = {}
    arguments = {'selected': selected, 'custom_cod': details.get('cod_no') or None,
                 'details': details, 'update_excel': scope == 'update_excel', 'progress': None, 'report_sink': snapshot}
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    run(paths, ROOT, check_only=True, **arguments)
    preflight = dict(snapshot) if snapshot else latest_result()
    if not preflight.get('scan_complete'):
        error = next((row.get('detail') for row in preflight.get('rows', []) if row.get('source_kind') == 'system'), None)
        raise ValueError(error or 'Both plant files must be checked before creating a certificate.')
    row = target_row(preflight, selected)
    if row is None:
        error = next((item.get('detail') for item in preflight.get('rows', []) if item.get('system_alert')), None)
        raise ValueError(error or 'The certificate details could not be checked.')
    if row.get('issuance_blocked'):
        raise ValueError(row.get('detail') or 'Correct the Excel import issue before creating this certificate.')
    if row.get('outcome') == 'needs_review' and not body.get('confirmed'):
        return {'requires_confirmation': True, 'warnings': [row], 'created': 0}
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    result = run(paths, ROOT, allow_warnings=bool(body.get('confirmed')),
                 override_reason=body.get('override_reason', ''), **arguments)
    final_report = snapshot or latest_result()
    final_row = target_row(final_report, selected)
    if not result[0]:
        raise ValueError((final_row or {}).get('detail') or 'The certificate could not be created. Excel details were not saved.')
    state = post_commit_state(lambda: refresh(paths))
    state_row = target_row(state, selected)
    if state_row and final_row:
        state_row['detail'] = final_row['detail'] + (' Certificate-only changes: original inward details were kept.'
                                                  if scope == 'certificate_only' else ' Edited details were also saved to Excel.')
        state_row['warnings'] = final_row.get('warnings', [])
    return {'requires_confirmation': False, 'state': state, 'created': result[0], 'review': result[1],
            'warnings': [state['setup_warning']] if state.get('setup_warning') else [],
            'certificates': certificate_outputs([final_row] if final_row else [])}


def post_commit_state(read_state):
    try:
        return read_state()
    except Exception as exc:
        try:
            state = latest_result()
        except Exception:
            state = {'rows': [], 'scan_complete': False}
        return dict(state, setup_warning=f'Certificate saved successfully, but the dashboard could not refresh: {exc}. Do not create it again; refresh after correcting the issue.')


def certificate_outputs(rows):
    outputs = []
    for row in rows:
        relative_path = row.get('pdf_relative_path')
        receipt = None
        if not relative_path and row.get('pdf_filename'):
            relative_path = row['pdf_filename']
        if not relative_path:
            try:
                receipt = next((entry for entry in reversed(audit_records(ROOT)) if entry.get('state') in {'complete', 'source_committed'}
                                and entry.get('source_kind') == row.get('source_kind', 'workbook')
                                and entry.get('plant') == row.get('plant')
                                and str(entry.get('key')) == str(row.get('record_id') or row.get('sn'))), None)
            except (OSError, ValueError):
                pass
        if receipt:
            receipt_path = Path(receipt.get('pdf_path', ''))
            if receipt_path.parts[:1] != ('Certificates',) or len(receipt_path.parts) < 2:
                continue
            relative_path = Path(*receipt_path.parts[1:]).as_posix()
        if receipt:
            cod = receipt.get('cod')
        else:
            cod = row.get('certificate_no')
        try:
            destination = safe_certificate_path(ROOT, relative_path)
            if not valid_pdf(destination):
                continue
            relative_path = destination.relative_to((ROOT / 'Certificates').resolve()).as_posix()
        except (OSError, ValueError, TypeError):
            continue
        filename = destination.name
        outputs.append({'filename': filename, 'relative_path': relative_path,
                        'certificate_no': cod, 'plant': row.get('plant'), 'supplier': row.get('supplier'),
                        'source_kind':row.get('source_kind','workbook'), 'sn':row.get('sn'),
                        'pdf_url': '/api/certificates/' + quote(relative_path, safe='/')})
    return outputs


def open_path(path):
    opener = getattr(os, "startfile", None)
    if opener is None:
        raise RuntimeError("Opening local folders is available from Windows Explorer on this computer.")
    opener(str(path))


def reveal_file_in_explorer(path):
    """Select an exact file through the Shell API, without parsing Explorer arguments."""
    if os.name != 'nt':
        raise RuntimeError('Showing a selected PDF in its folder is available on Windows.')
    import ctypes

    ole = ctypes.WinDLL('ole32')
    shell = ctypes.WinDLL('shell32')
    hresult = ctypes.c_int32
    ole.CoInitializeEx.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    ole.CoInitializeEx.restype = hresult
    ole.CoUninitialize.argtypes = []
    ole.CoUninitialize.restype = None
    ole.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    ole.CoTaskMemFree.restype = None
    shell.SHParseDisplayName.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p,
        ctypes.POINTER(ctypes.c_void_p), ctypes.c_uint32, ctypes.POINTER(ctypes.c_uint32)]
    shell.SHParseDisplayName.restype = hresult
    shell.SHOpenFolderAndSelectItems.argtypes = [ctypes.c_void_p, ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_void_p), ctypes.c_uint32]
    shell.SHOpenFolderAndSelectItems.restype = hresult
    initialized = ole.CoInitializeEx(None, 2)
    # RPC_E_CHANGED_MODE means this request thread already has a COM apartment.
    if initialized < 0 and initialized != -2147417850:
        raise OSError(f'Windows could not initialize file selection (0x{initialized & 0xffffffff:08X}).')
    item = ctypes.c_void_p()
    try:
        parsed = shell.SHParseDisplayName(str(path), None, ctypes.byref(item), 0, None)
        if parsed < 0 or not item.value:
            raise OSError(f'Windows could not locate this PDF: {path} (0x{parsed & 0xffffffff:08X}).')
        # With cidl=0, the absolute file PIDL selects that file in its parent folder.
        selected = shell.SHOpenFolderAndSelectItems(item, 0, None, 0)
        if selected < 0:
            raise OSError(f'Windows could not select this PDF: {path} (0x{selected & 0xffffffff:08X}).')
    finally:
        if item.value:
            ole.CoTaskMemFree(item)
        if initialized >= 0:
            ole.CoUninitialize()


def open_certificate_locally(relative_path, *, reveal=True, open_document=True):
    destination = safe_certificate_path(ROOT, relative_path)
    if not destination.is_file() or not valid_pdf(destination):
        raise ValueError('This PDF is missing or unreadable. Refresh the certificate list and try again.')
    if reveal:
        reveal_file_in_explorer(destination)
    if open_document:
        try:
            open_path(destination)
        except OSError as exc:
            raise OSError('The PDF is selected in its folder, but Windows could not open it. '
                          'Choose a default PDF viewer in Windows and try again.') from exc
    return destination


class Handler(BaseHTTPRequestHandler):
    server_version = "EcorecoLocal/1.0"

    def log_message(self, _format, *_args):
        return

    def send_json(self, value, status=200):
        if isinstance(value, dict) and status < 400:
            value = dict(value)
            # Reads can fail while Excel is open; do not misreport an already
            # committed certificate as a failure just to obtain a UI snapshot.
            if value.get('changed') is not False:
                try:
                    value['workspace_revision'] = workspace_revision(ROOT)
                except (OSError, ValueError):
                    value.setdefault('workspace_revision', None)
        payload = json.dumps(value, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(payload)

    def send_error_json(self, exc, status=400):
        self.send_json({"error": str(exc)}, status)

    def do_GET(self):
        if urlsplit(self.path).path == '/api/dashboard-review':
            self.send_json({'application':'ecoreco-certificates', 'changed':False})
            return
        with OPERATION_LOCK:
            self._do_get()

    def _do_get(self):
        path = urlsplit(self.path).path
        if path == "/api/workbook-changes":
            try:
                self.send_json(check_workbook_changes())
            except Exception as exc:
                self.send_error_json(exc, 500)
            return
        if path == "/api/state":
            try:
                settings = load_settings(ROOT)
                self.send_json({
                    "state": workspace_state(),
                    "settings": {
                        "workbook_paths": settings["workbook_paths"],
                        "reminders_enabled": settings["reminders_enabled"],
                        "reminder_time": settings["reminder_time"],
                        'backup_destination': settings.get('backup_destination', ''),
                    },
                    "certificate_folder": str(ROOT / "Certificates"),
                    "server_time": datetime.now().isoformat(timespec="minutes"),
                })
            except Exception as exc:
                self.send_error_json(exc, 500)
            return

        if path == "/api/templates":
            try:
                self.send_json({"templates": template_specs()})
            except Exception as exc:
                self.send_error_json(exc, 500)
            return
        if path == '/api/history':
            try:
                self.send_json({'history': audit_records(ROOT)[-50:]})
            except (OSError, ValueError) as exc:
                self.send_error_json(exc)
            return
        if path.startswith('/api/certificates/'):
            relative_path = unquote(path[len('/api/certificates/'):])
            try:
                destination = safe_certificate_path(ROOT, relative_path)
            except ValueError:
                self.send_error(404)
                return
            if not valid_pdf(destination):
                self.send_error_json('This PDF is missing or unreadable. Refresh and recreate it.', 404)
                return
            payload = destination.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'application/pdf')
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(payload)
            return
        preview_match = re.fullmatch(r"/previews/([a-f0-9]{32})\.(pdf|png)", path)
        if preview_match:
            extension = preview_match[2]
            destination = ROOT / "Data" / "Previews" / f"{preview_match[1]}.{extension}"
            if not destination.is_file():
                self.send_error(404)
                return
            payload = destination.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", 'application/pdf' if extension == 'pdf' else 'image/png')
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(payload)
            return

        assets = {"/": ("index.html", "text/html; charset=utf-8"),
                  "/styles.css": ("styles.css", "text/css; charset=utf-8"),
                  "/app.js": ("app.js", "text/javascript; charset=utf-8")}
        asset = assets.get(path)
        if not asset:
            self.send_error(404)
            return
        name, mime = asset
        try:
            payload = (WEB_ROOT / name).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(payload)
        except OSError as exc:
            self.send_error_json(exc, 500)

    def do_POST(self):
        if urlsplit(self.path).path in {'/api/dashboard-presence', '/api/dashboard-review'}:
            if STOPPING:
                self.send_error_json('The dashboard is closing.', 503)
            else:
                self._do_post()
            return
        with OPERATION_LOCK:
            if STOPPING:
                self.send_error_json('The dashboard is closing. Reopen it before starting another action.', 503)
                return
            self._do_post()

    def _do_post(self):
        global SERVER
        global STOPPING
        parsed = urlsplit(self.path)
        if not self._same_origin():
            self.send_error_json("This local app only accepts requests from its own page.", 403)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 0 or length > 1_000_000:
                raise ValueError("The request is too large.")
            body = json.loads(self.rfile.read(length) or b"{}")
            if not isinstance(body, dict):
                raise ValueError("Expected an object of request values.")

            if parsed.path == '/api/dashboard-presence':
                request = DASHBOARD_REVIEW.presence(body.get('client_id'), body.get('visible'), body.get('acknowledged'))
                self.send_json({'review_request':request, 'changed':False})
            elif parsed.path == '/api/dashboard-review':
                result = DASHBOARD_REVIEW.status(body.get('request_id'), cancel=body.get('cancel') is True) if body.get('request_id') else DASHBOARD_REVIEW.request()
                self.send_json({**result, 'changed':False})
            elif parsed.path == "/api/check":
                result = refresh()
                settings = load_settings(ROOT)
                self.send_json({"state": result, "settings": {
                    "workbook_paths": settings["workbook_paths"],
                    "reminders_enabled": settings["reminders_enabled"],
                    "reminder_time": settings["reminder_time"],
                    'backup_destination': settings.get('backup_destination', ''),
                }})
            elif parsed.path == "/api/preview":
                self.send_json(preview_certificate(body))
            elif parsed.path == '/api/clients/add':
                self.send_json(add_client(body))
            elif parsed.path == "/api/manual/save":
                require_revision(ROOT, body.get('expected_revision'))
                identifier = save_draft(ROOT, body.get("plant"), body.get("details"), body.get("record_id"),
                                        expected_revision=body.get('expected_revision'))
                self.send_json({"record_id": identifier, "state": workspace_state()})
            elif parsed.path == "/api/manual/generate":
                response = reviewed_operation(ROOT, parsed.path, body, lambda values: generate_draft(
                    ROOT, values.get('plant'), values.get('details'), values.get('record_id'),
                    confirmed=bool(values.get('confirmed')), override_reason=values.get('override_reason', ''),
                    expected_revision=values.get('expected_revision')))
                response["state"] = post_commit_state(workspace_state) if response.get('created') else workspace_state()
                if response.get('created'):
                    if response['state'].get('setup_warning'):
                        response.setdefault('warnings', []).append(response['state']['setup_warning'])
                    response['certificates'] = certificate_outputs([{'source_kind': 'manual',
                        'record_id': response['record_id'], 'plant': body.get('plant'),
                        'pdf_filename': response.get('filename'), 'certificate_no': response.get('certificate_no'),
                        'pdf_relative_path': response.get('relative_path'),
                        'supplier': (body.get('details') or {}).get('supplier')}])
                self.send_json(response)
            elif parsed.path == "/api/manual/complete":
                require_revision(ROOT, body.get('expected_revision'))
                complete_draft(ROOT, body.get("record_id"), expected_revision=body.get('expected_revision'))
                self.send_json({"state": workspace_state()})
            elif parsed.path == "/api/generate":
                selected = None
                if body.get("plant") or body.get("sn"):
                    selected = selected_key(body.get("plant"), body.get("sn"))
                plants = body.get("plants")
                self.send_json(reviewed_operation(ROOT, parsed.path, body, lambda values: generate(
                    selected=selected, plants=plants, custom_cod=str(values.get('custom_cod') or '').strip() or None,
                    confirmed=bool(values.get('confirmed')), expected_revision=values.get('expected_revision'),
                    override_reason=values.get('override_reason', ''))))
            elif parsed.path == "/api/create":
                self.send_json(reviewed_operation(ROOT, parsed.path, body, create_certificate))
            elif parsed.path == "/api/pick-annexure":
                require_revision(ROOT, body.get('expected_revision'))
                import tkinter as tk
                from tkinter import filedialog
                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes("-topmost", True)
                    picked = filedialog.askopenfilename(parent=picker_root, title="Choose an annexure PDF",
                                                       filetypes=[("PDF files", "*.pdf")])
                finally:
                    picker_root.destroy()
                attachment = None
                if picked:
                    # A workbook can change while the native picker is open.
                    require_revision(ROOT, body.get('expected_revision'))
                    source = Path(picked)
                    if source.suffix.casefold() != ".pdf" or source.stat().st_size > 50_000_000:
                        raise ValueError("Choose a PDF smaller than 50 MB.")
                    reader = PdfReader(source)
                    if reader.is_encrypted or not reader.pages:
                        raise ValueError("Choose an unlocked PDF with at least one page.")
                    identifier = uuid.uuid4().hex
                    folder = ROOT / "Data" / "Attachments"
                    folder.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, folder / f"{identifier}.pdf")
                    attachment = {"id": identifier, "name": source.name, "pages": len(reader.pages)}
                    save_json(folder / f"{identifier}.json", attachment)
                self.send_json({"annexure": attachment})
            elif parsed.path == "/api/complete":
                with run_lock(ROOT):
                    recover_issuances(ROOT)
                    require_revision(ROOT, body.get('expected_revision'))
                    selected = selected_key(body.get("plant"), body.get("sn"))
                    row = target_row(latest_result(), selected)
                    if not row or not row.get("file"):
                        raise ValueError("Refresh the workbooks and select the row again.")
                    if not find_certificate(ROOT / 'Certificates', {'cod_no': row.get('certificate_no'),
                            'source_id': row['plant'], 'case_id': row['sn']}):
                        raise ValueError('The current PDF is missing, changed or unreadable. Refresh and recreate it before marking it Completed.')
                    update_issued_status(
                        row["file"], BACKUP_DIR, row["sn"], row["sheet"], row["row"],
                        row.get("certificate_no", ""), row.get("certificate_status", "Generated"), "Completed",
                    )
                self.send_json({"state": refresh()})
            elif parsed.path == "/api/open-folder":
                folder = ROOT / "Certificates"
                ensure_certificate_folders(ROOT)
                open_path(folder)
                self.send_json({"ok": True})
            elif parsed.path == "/api/open-certificate":
                open_document = bool(body.get('open_pdf', True))
                destination = open_certificate_locally(body.get('relative_path'), reveal=True,
                                                       open_document=open_document)
                self.send_json({"ok": True, 'pdf_path': str(destination),
                                'folder': str(destination.parent), 'selected': True,
                                'opened': open_document})
            elif parsed.path == "/api/open-workbook":
                path, sheet, row = open_selected_workbook_record(body.get("plant"), body.get("sn"))
                self.send_json({"ok": True, 'selected':True, 'workbook':str(path), 'sheet':sheet, 'row':row})
            elif parsed.path == "/api/pick-workbook":
                import tkinter as tk
                from tkinter import filedialog

                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes("-topmost", True)
                    picked = filedialog.askopenfilename(
                        parent=picker_root,
                        title="Choose a plant Excel workbook",
                        filetypes=(("Excel workbooks", "*.xlsx *.xlsm"),),
                    )
                finally:
                    picker_root.destroy()
                if not picked:
                    self.send_json({"cancelled": True})
                    return
                picked_path = Path(picked).expanduser().resolve()
                settings = load_settings(ROOT)
                configured = [Path(value).expanduser().resolve() for value in settings["workbook_paths"]]
                plant = next((f"plant{index + 1}" for index, path in enumerate(configured)
                              if path.is_file() and path.samefile(picked_path)), None)
                if not plant and not body.get("for_settings"):
                    raise ValueError("That workbook is not configured yet. Add its path in Settings first.")
                self.send_json({"path": str(picked_path), "plant": plant})
            elif parsed.path == "/api/settings":
                settings = load_settings(ROOT)
                paths = body.get("workbook_paths", settings["workbook_paths"])
                if not isinstance(paths, list) or len(paths) != 2 or any(not isinstance(item, str) or not item.strip() for item in paths):
                    raise ValueError("Enter both Plant 1 and Plant 2 workbook paths.")
                resolved = [Path(item).expanduser().resolve() for item in paths]
                if not all(path.is_file() for path in resolved):
                    raise ValueError("One of those workbooks could not be found.")
                if resolved[0] == resolved[1] or resolved[0].samefile(resolved[1]):
                    raise ValueError("Plant 1 and Plant 2 must use different workbooks.")
                reminder_time = str(body.get("reminder_time", settings["reminder_time"]))
                if len(reminder_time) != 5 or reminder_time[2] != ":":
                    raise ValueError("Choose a valid reminder time.")
                try:
                    datetime.strptime(reminder_time, "%H:%M")
                except ValueError as exc:
                    raise ValueError("Choose a valid reminder time.") from exc
                enabled = body.get("reminders_enabled", settings["reminders_enabled"])
                if not isinstance(enabled, bool):
                    raise ValueError("Reminder enabled must be on or off.")
                settings.update(workbook_paths=[str(path) for path in resolved],
                                reminder_time=reminder_time, reminders_enabled=enabled)
                save_json(ROOT / "settings.local.json", settings)
                message = "Settings saved."
                reminder_feedback = {}
                if body.get("update_reminder"):
                    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                               "-File", str(ROOT / "App" / "Install Daily Task.ps1"),
                               "-Mode", "Install" if enabled else "Disable"]
                    if enabled:
                        command.extend(["-ReminderTime", reminder_time])
                    # Settings are already saved; report scheduling separately.
                    try:
                        completed = subprocess.run(command, capture_output=True, text=True,
                                                   timeout=60, check=False)
                        reminder_updated = completed.returncode == 0
                    except (subprocess.TimeoutExpired, OSError):
                        reminder_updated = False
                    reminder_feedback["reminder_updated"] = reminder_updated
                    if reminder_updated:
                        message = "Settings and daily reminder updated."
                    else:
                        reminder_feedback["warning"] = (
                            "Settings saved, but Windows did not confirm the daily reminder update. "
                            "Choose Save & repair reminder to try again."
                        )
                self.send_json({"settings": {
                    "workbook_paths": settings["workbook_paths"],
                    "reminders_enabled": enabled, "reminder_time": reminder_time,
                    'backup_destination': settings.get('backup_destination', ''),
                }, "message": message, **reminder_feedback})
            elif parsed.path == '/api/backup/pick-destination':
                import tkinter as tk
                from tkinter import filedialog
                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes('-topmost', True)
                    picked = filedialog.askdirectory(parent=picker_root, title='Choose a separate backup folder')
                finally:
                    picker_root.destroy()
                if not picked:
                    self.send_json({'cancelled': True})
                    return
                chosen = Path(picked).resolve()
                if chosen == ROOT or chosen.is_relative_to(ROOT / 'Certificates') or chosen.is_relative_to(ROOT / 'Data') or chosen.is_relative_to(ROOT / 'App'):
                    raise ValueError('Choose a separate folder, outside this app and its certificate data.')
                settings = load_settings(ROOT)
                settings['backup_destination'] = str(chosen)
                with run_lock(ROOT):
                    save_json(ROOT / 'settings.local.json', settings)
                self.send_json({'backup_destination': str(chosen), 'settings': settings,
                                'message': 'Backup folder saved. A separate drive is recommended.'})
            elif parsed.path == '/api/backup/create':
                result = create_backup(ROOT)
                self.send_json({'backup': result, 'backup_destination': load_settings(ROOT).get('backup_destination', ''),
                                'message': 'Verified backup created. Keep the backup and its trust files in a safe separate location.',
                                'health': '\n'.join(result.get('warnings', []))})
            elif parsed.path == '/api/backup/list':
                self.send_json({'backups': list_backups(ROOT),
                                'backup_destination': load_settings(ROOT).get('backup_destination', '')})
            elif parsed.path == '/api/backup/restore':
                with run_lock(ROOT):
                    require_revision(ROOT, body.get('expected_revision'))
                    result = restore_backup(ROOT, body.get('backup_id'), confirmed=body.get('confirmed') is True, locked=True)
                self.send_json({'restored': result, 'state': refresh(), 'settings': load_settings(ROOT),
                                'message': 'Backup restored. The previous data was preserved in a recovery snapshot.'})
            elif parsed.path == "/api/stop":
                STOPPING = True
                self.send_json({"ok": True})
                if SERVER:
                    threading.Thread(target=SERVER.shutdown, daemon=True).start()
            else:
                self.send_error_json("Unknown action.", 404)
        except (ValueError, OSError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
            self.send_error_json(exc)
        except Exception as exc:
            self.send_error_json(exc, 500)

    def _same_origin(self):
        host = self.headers.get("Host", "").split(":", 1)[0].casefold()
        if host not in {"127.0.0.1", "localhost"}:
            return False
        origin = self.headers.get("Origin")
        if not origin:
            return True
        parsed = urlsplit(origin)
        return parsed.hostname in {"127.0.0.1", "localhost"} and parsed.port == self.server.server_port


def focus_dashboard_window():
    """Bring an already-visible dashboard browser window forward on Windows."""
    if os.name != 'nt':
        return
    import ctypes
    user = ctypes.WinDLL('user32')
    callback_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user.EnumWindows.argtypes = [callback_type, ctypes.c_void_p]
    user.GetWindowTextW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
    user.IsWindowVisible.argtypes = [ctypes.c_void_p]
    user.ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
    user.IsIconic.argtypes = [ctypes.c_void_p]
    user.SetForegroundWindow.argtypes = [ctypes.c_void_p]

    @callback_type
    def find_window(handle, _):
        title = ctypes.create_unicode_buffer(512)
        user.GetWindowTextW(handle, title, len(title))
        if user.IsWindowVisible(handle) and title.value.startswith('Certificates · Ecoreco'):
            if user.IsIconic(handle):
                user.ShowWindow(handle, 9)
            user.SetForegroundWindow(handle)
            return False
        return True
    user.EnumWindows(find_window, None)


def reuse_dashboard(port):
    """Wait for receipt in a live tab; fall back if that tab has closed."""
    def send(payload):
        connection = HTTPConnection('127.0.0.1', port, timeout=1)
        try:
            connection.request('POST', '/api/dashboard-review', json.dumps(payload), {'Content-Type':'application/json'})
            response = connection.getresponse()
            value = json.loads(response.read())
            return value if response.status == 200 and isinstance(value, dict) else {}
        finally:
            connection.close()
    try:
        result = send({})
        if not result.get('reused') or not result.get('request_id'):
            return False
        try:
            focus_dashboard_window()
        except OSError:
            pass  # A Windows focus restriction must not create a duplicate tab.
        request_id = result['request_id']
        for _ in range(20):
            if send({'request_id':request_id}).get('acknowledged'):
                return True
            time.sleep(.25)
        # Cancel atomically so a delayed old tab cannot also handle the fallback.
        return bool(send({'request_id':request_id, 'cancel':True}).get('acknowledged'))
    except (OSError, ValueError):
        return False


def main(attention=False, open_browser=True):
    global SERVER
    ensure_certificate_folders(ROOT)
    for port in range(8765, 8776):
        try:
            connection = HTTPConnection("127.0.0.1", port, timeout=0.6)
            connection.request('GET', '/api/dashboard-review')
            response = connection.getresponse()
            is_existing_dashboard = response.status == 200 and "application/json" in response.getheader("Content-Type", "")
            response.read()
            if not is_existing_dashboard:
                # Older running dashboards do not yet expose the presence API.
                connection.request('GET', '/api/state')
                response = connection.getresponse()
                is_existing_dashboard = response.status == 200 and 'application/json' in response.getheader('Content-Type', '')
                response.read()
            connection.close()
            if is_existing_dashboard:
                if attention and reuse_dashboard(port):
                    return
                if open_browser:
                    webbrowser.open(f"http://127.0.0.1:{port}/" + ('?view=attention' if attention else ''))
                return
        except OSError:
            pass
        try:
            SERVER = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if SERVER is None:
        raise RuntimeError("Could not start the local dashboard. Close any old certificate dashboard and try again.")
    # Wait for request workers before exit: a second tab cannot terminate an
    # in-progress publication halfway through a workbook/register update.
    SERVER.daemon_threads = False
    url = f"http://127.0.0.1:{SERVER.server_port}/" + ('?view=attention' if attention else '')
    if open_browser:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    SERVER.serve_forever(poll_interval=0.25)
    SERVER.server_close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--attention', action='store_true')
    parser.add_argument('--no-browser', action='store_true')
    arguments = parser.parse_args()
    main(attention=arguments.attention, open_browser=not arguments.no_browser)
