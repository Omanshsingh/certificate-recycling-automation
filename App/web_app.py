"""Loopback-only browser interface for the Ecoreco certificate workflow."""
from __future__ import annotations

from collections import OrderedDict
from datetime import date, datetime
from http.client import HTTPConnection, HTTPException
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
import json
import argparse
import copy
import os
from pathlib import Path
import subprocess
import threading
import time
import traceback
from urllib.parse import urlsplit, unquote, quote, parse_qs
import webbrowser
import re
import uuid
import shutil

from app_paths import DATA_ROOT
from certificate_core import (batch_content, inward_details_complete, number, parse_date,
                               read_batches, source_value_issues, status_value, text,
                               supplier_is_always_not_applicable)
from run_certificates import run, find_certificate, workbook_cod_ownership_issue
from workbook_store import LockBusyError, update_issued_status, WorkbookSession, WorkbookSnapshot, stage_certificate_details
from workflow_support import Report, load_settings, open_excel_record, open_excel_workbook, save_json, operation_error_message
from zipfile import BadZipFile
from certificate_drafts import (clean_details, save_draft, generate_draft, complete_draft,
                                prepare_certificate, manual_rows, material_items, reconcile_manual,
                                read_register, current_record_pdf)
from certificate_templates import template_specs, make_template_certificate, obsolete_template_warning, display_values, current_warnings
from pypdf import PdfReader
from openpyxl.utils.exceptions import InvalidFileException
from certificate_inventory import scan_inventory, safe_certificate_path, ensure_certificate_folders, identity as cod_identity
from workflow_support import run_lock
from file_fingerprints import file_signature
from issuance_safety import (recover_issuances, valid_pdf, file_digest, digest_bytes, latest_issuance,
                             audit_records, workbook_receipt_conflict, current_pdf_is_healthy,
                             manual_receipts)
from workspace_review import workspace_revision, require_revision, reviewed_operation
from backup_service import create_backup, list_backups, restore_backup
from dashboard_review import DashboardReview
from client_intake import intake_details, stage_new_client
from plant_workbook_template import create_plant_workbook
from issued_detail_changes import (REVIEW_FILE as ISSUED_CHANGE_REVIEW_FILE, clear_review,
    FIELDS as ISSUED_SOURCE_FIELDS, compare_details, is_intentionally_kept,
    intentional_reviews, keep_intentional, normalized_source_details)


ROOT = Path(DATA_ROOT).resolve()
WEB_ROOT = Path(__file__).resolve().parent / "web"
BACKUP_DIR = ROOT / "Data" / "Workbook Backups"
RESULTS_PATH = ROOT / "Data" / "Reports" / "Latest Results.json"
CERTIFICATE_TYPE_CHOICES_PATH = ROOT / 'Data' / 'Certificate Type Choices.json'
SERVER: ThreadingHTTPServer | None = None
OPERATION_LOCK = threading.RLock()
STOPPING = False
WORKBOOK_CHECK_STAMPS = {}
WORKSPACE_STATE_CACHE = None
PREVIEW_CACHE = OrderedDict()
DASHBOARD_REVIEW = DashboardReview()


def cached_workspace_state(expected_revision):
    cached = WORKSPACE_STATE_CACHE
    if (isinstance(expected_revision, str) and cached
            and cached[0] == expected_revision):
        return copy.deepcopy(cached[1])
    return None


def merge_workbook_report(expected_revision, report, *, preserve_unmentioned=False):
    """Update a checked dashboard snapshot from the report produced by a save."""
    state = cached_workspace_state(expected_revision)
    if state is None or not isinstance(report, dict) or not isinstance(report.get('rows'), list):
        return None
    previous_rows = state.get('rows', [])
    previous = {(row.get('plant'), str(row.get('sn')).strip().casefold()): row
                for row in previous_rows if row.get('source_kind') != 'manual'}
    merged = []
    for incoming in report['rows']:
        row = dict(incoming)
        if row.get('source_kind') != 'manual':
            old = previous.get((row.get('plant'), str(row.get('sn')).strip().casefold()))
            if old:
                for key in ('has_issued_certificate', 'last_certificate_details'):
                    if key in old and key not in row:
                        row[key] = copy.deepcopy(old[key])
        merged.append(row)
    if preserve_unmentioned:
        incoming_keys = {(row.get('plant'), str(row.get('sn')).strip().casefold()) for row in merged
                         if row.get('source_kind') != 'manual'}
        merged.extend(copy.deepcopy(row) for row in previous_rows if row.get('source_kind') == 'workbook'
                      and (row.get('plant'), str(row.get('sn')).strip().casefold()) not in incoming_keys)
    merged.extend(copy.deepcopy(row) for row in previous_rows if row.get('source_kind') == 'manual')
    state['rows'] = merged
    for key in ('refresh_pending', 'recovery_note'):
        state.pop(key, None)
    for key in ('scan_complete', 'certificate_inventory', 'numbering', 'created', 'review',
                'completed_at', 'as_of', 'run_id', 'maintenance_notes', 'setup_warning', 'refresh_pending',
                'recovery_note'):
        if key in report:
            state[key] = copy.deepcopy(report[key])
    return state


def merge_added_workbook_row(expected_revision, row, state=None):
    state = state if state is not None else cached_workspace_state(expected_revision)
    if state is None:
        return None
    rows = state.setdefault('rows', [])
    key = (row.get('plant'), str(row.get('sn')).strip().casefold())
    existing = next((item for item in rows if item.get('source_kind') != 'manual'
                     and (item.get('plant'), str(item.get('sn')).strip().casefold()) == key), None)
    if existing is None:
        rows.append(copy.deepcopy(row))
    else:
        existing.update(copy.deepcopy(row))
    state['scan_complete'] = True
    return state


def apply_inward_number_changes(state, plant, changes, batches):
    positions = {batch['case_id']:(batch['source_sheet'], batch['source_row']) for batch in batches}
    seen_links = set()
    for row in state.get('rows', []):
        if row.get('source_kind') != 'manual' and row.get('plant') == plant:
            old = str(row.get('sn', ''))
            row['sn'] = changes.get(old, old)
            if row['sn'] in positions:
                row['sheet'], row['row'] = positions[row['sn']]
            if isinstance(row.get('last_certificate_details'), dict):
                row['last_certificate_details']['case_id'] = row['sn']
        for details in (row, row.get('raw_details', {})):
            link = details.get('destruction_source')
            if isinstance(link, dict) and not link.get('record_id') and link.get('plant') == plant and id(link) not in seen_links:
                seen_links.add(id(link))
                link['case_id'] = changes.get(link.get('case_id'), link.get('case_id'))
    seen_owners = set()
    for entry in state.get('certificate_inventory', {}).get('entries', []):
        owner = entry.get('managed_owner', {})
        if owner.get('source_kind') == 'workbook' and owner.get('plant') == plant and id(owner) not in seen_owners:
            seen_owners.add(id(owner))
            owner['key'] = changes.get(str(owner.get('key')), owner.get('key'))
    return state


def merge_manual_record_state(expected_revision, record_id, state=None):
    """Refresh one custom certificate row after its local register is saved."""
    state = state if state is not None else cached_workspace_state(expected_revision)
    if state is None:
        return None
    record = next((item for item in read_register(ROOT).get('records', [])
                   if item.get('id') == record_id), None)
    if not record:
        return None
    if record.get('destruction_requested') is False and not current_record_pdf(ROOT, record):
        state['rows'] = [row for row in state.get('rows', []) if row.get('record_id') != record_id]
        return state
    details = record['details']
    materials, _ = material_items(details)
    status = status_value(record.get('status'))
    if status == 'generated':
        status = 'generated_manual'
    weight = number(details.get('certificate_weight'))
    if weight is None:
        weight = sum(item.get('weight') or 0 for item in materials) if materials else number(details.get('weight'))
    row = {
        'source_kind':'manual', 'record_id':record_id, 'plant':record.get('plant'),
        'destruction_source':details.get('destruction_source'),
        'destruction_requested':record.get('destruction_requested', True),
        'destruction_pending':record.get('destruction_requested', True) and details.get('certificate_type') == 'destruction' and status != 'completed',
        'raw_details':dict(details), 'case_id':details.get('case_id', ''),
        'sn':details.get('case_id') or f'Custom {record_id[:8]}',
        'supplier':details.get('supplier', ''), 'location':details.get('location', ''),
        'certificate_type':details.get('certificate_type', 'standard'),
        'certificate_fields':details.get('certificate_fields', {}), 'form6':details.get('form6', ''),
        'reference':details.get('reference', ''), 'inward_date':details.get('inward_date', ''),
        'certificate_no':details.get('cod_no', ''), 'certificate_date':details.get('cod_date', ''),
        'has_issued_certificate':status in {'generated_manual', 'completed'},
        'certificate_weight':number(details.get('certificate_weight')), 'display_weight':weight,
        'weight_source':'certificate' if number(details.get('certificate_weight')) is not None else 'inward',
        'certificate_status':status, 'outcome':status,
        'detail':'Saved in the local certificate register.',
        'warnings':current_warnings(record.get('warnings', [])), 'items':materials,
        'pdf_filename':record.get('pdf_filename', ''),
        'pdf_relative_path':record.get('pdf_relative_path', ''),
        'file':'', 'reminder_due':False, 'overdue':False,
    }
    normalized = manual_rows(ROOT, records=[record])
    if normalized:
        row.update(normalized[0])
    rows = state.setdefault('rows', [])
    existing = next((item for item in rows if item.get('source_kind') == 'manual'
                     and item.get('record_id') == record_id), None)
    if existing is None:
        rows.append(row)
    else:
        existing.update(row)
    ledger_path = ROOT / 'Data' / 'Certificate Inventory.json'
    try:
        inventory = json.loads(ledger_path.read_text(encoding='utf-8'))
        if isinstance(inventory, dict) and isinstance(inventory.get('entries'), list):
            inventory['warnings'] = [f"{item['path']}: {warning}" for item in inventory['entries']
                                     if item.get('present') for warning in item.get('warnings', [])]
            owners = {}
            for item in inventory['entries']:
                if item.get('present'):
                    content_id = item.get('sha256') or f"path:{item.get('path', '').casefold()}"
                    for cod in {cod_identity(value) for value in item.get('current_cods', item.get('cods', []))}:
                        owners.setdefault(cod, {}).setdefault(content_id, item.get('path', ''))
            for cod, paths in owners.items():
                if len(paths) > 1:
                    names = ', '.join(Path(path).name for path in list(paths.values())[:3])
                    inventory['warnings'].append(f'COD {cod} appears in different PDFs ({names}). Verify which certificate is current before sending.')
            state['certificate_inventory'] = add_certificate_type_choices(inventory)
    except (OSError, ValueError, TypeError):
        pass
    return state


def remember_workbook_stamp():
    try:
        WORKBOOK_CHECK_STAMPS[str(ROOT)] = workbook_check_stamp(configured_paths())
    except (OSError, ValueError):
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)


def certificate_type_choices(*, strict=False):
    if not CERTIFICATE_TYPE_CHOICES_PATH.is_file():
        return {}
    try:
        data = json.loads(CERTIFICATE_TYPE_CHOICES_PATH.read_text(encoding='utf-8'))
        choices = data.get('choices') if isinstance(data, dict) and data.get('version') == 1 else None
        if not isinstance(choices, dict):
            raise ValueError('The saved type choices have an invalid format.')
        return {digest: kind for digest, kind in choices.items()
                if re.fullmatch(r'[a-f0-9]{64}', str(digest))
                and kind in {'standard', 'destruction', 'siemens', 'sita'}}
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        if strict:
            raise ValueError('Saved certificate types could not be read. Fix the type choices file before saving another choice.') from exc
        return {}


def add_certificate_type_choices(inventory):
    choices = certificate_type_choices()
    for entry in inventory.get('entries', []):
        digest = entry.get('sha256')
        if entry.get('present') and isinstance(digest, str) and digest in choices:
            entry['certificate_type'] = choices[digest]
    return inventory


def save_certificate_type_choice(body):
    kind = body.get('certificate_type')
    digest = str(body.get('sha256') or '').casefold()
    relative_path = body.get('relative_path')
    if kind not in {'standard', 'destruction', 'siemens', 'sita'}:
        raise ValueError('Choose Recycling, Data Destruction, Medical, or International.')
    if not re.fullmatch(r'[a-f0-9]{64}', digest):
        raise ValueError('Refresh the saved files and choose a type again.')
    destination = safe_certificate_path(ROOT, relative_path)
    if not valid_pdf(destination) or file_digest(destination) != digest:
        raise ValueError('This PDF changed. Refresh Saved certificates before choosing a type.')
    choices = certificate_type_choices(strict=True)
    choices[digest] = kind
    save_json(CERTIFICATE_TYPE_CHOICES_PATH, {'version': 1, 'choices': choices})
    return kind


def workbook_check_stamp(paths):
    """Cheap saved-file change detection for workbooks and certificate PDFs."""
    stamp = [('as_of', date.today().isoformat())]
    for tracked in (ROOT / 'Data' / 'Custom Certificates.json', ROOT / 'Data' / 'History Reset.json',
                    ROOT / 'Data' / ISSUED_CHANGE_REVIEW_FILE, ROOT / 'settings.local.json'):
        if tracked.exists():
            stamp.append(('tracking', tracked.name, file_signature(tracked)))
    for path in paths:
        stat = path.stat()
        stamp.append(('workbook', str(path), stat.st_mtime_ns, stat.st_ctime_ns, stat.st_size))
    folder = ROOT / 'Certificates'
    if folder.is_dir():
        for path in sorted(folder.rglob('*'), key=lambda item: item.as_posix().casefold()):
            if not path.is_file() or path.suffix.casefold() != '.pdf' or path.name.startswith('.'):
                continue
            try:
                signature = file_signature(path)
            except FileNotFoundError:
                continue
            stamp.append(('pdf', path.relative_to(ROOT).as_posix().casefold(), signature))
    return tuple(stamp)


def check_workbook_changes():
    settings = load_settings(ROOT)
    if not any(value.strip() for value in settings['workbook_paths']):
        if WORKBOOK_CHECK_STAMPS.get(str(ROOT)) == ('disconnected', workbook_check_stamp([])):
            return {'changed': False}
        return {'changed': True, 'state': refresh()}
    try:
        paths = configured_paths()
        stamp = workbook_check_stamp(paths)
        if WORKBOOK_CHECK_STAMPS.get(str(ROOT)) == stamp and reusable_workbook_report(latest_result()):
            return {'changed': False}
    except (OSError, ValueError):
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        return {'changed': True, 'state': refresh()}
    return {"changed": True, "state": refresh(paths)}


def configured_paths(settings=None, allow_unavailable=False):
    settings = settings or load_settings(ROOT)
    if not all(value.strip() for value in settings['workbook_paths']):
        raise ValueError('Connect both plant Excel files in Settings to use Excel records.')
    paths = [Path(value).expanduser().resolve() for value in settings["workbook_paths"]]
    if len(paths) != 2 or (not allow_unavailable and not all(path.is_file() for path in paths)):
        raise ValueError("Choose the Plant 1 and Plant 2 workbooks in Settings first.")
    if paths[0] == paths[1] or (paths[0].is_file() and paths[1].is_file() and paths[0].samefile(paths[1])):
        raise ValueError("Plant 1 and Plant 2 must use different workbooks.")
    return paths


class ClientAddRequestEnded(ValueError):
    """The current request ID cannot be used for another add."""


def add_client(body):
    plant = body.get('plant')
    if plant not in ('plant1', 'plant2'):
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
    serial_changes = {}
    ordering_warning = ''
    with run_lock(ROOT):
        recover_issuances(ROOT)
        journal = json.loads(journal_path.read_text(encoding='utf-8')) if journal_path.is_file() else None
        if journal and (journal.get('plant') != plant or journal.get('path') != str(path) or journal.get('input') != signature):
            raise ClientAddRequestEnded(
                'This add request belongs to different details. Check Companies for the previous client. '
                'Use Review & edit to correct it, or choose Add to Excel again to add this as a new entry.')
        if not journal:
            require_revision(ROOT, expected)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path.resolve()):session.data}
            if journal:
                from inward_serials import links_from
                current_sn = links_from(session.data).get('additions', {}).get(request_id, journal['sn'])
                if current_sn is None:
                    raise ClientAddRequestEnded(
                        'This client was already added and later removed. '
                        'Review the fields and choose Add to Excel to create a new entry.')
                journal['sn'] = current_sn
                if isinstance(journal.get('snapshot'), dict):
                    journal['snapshot']['case_id'] = current_sn
                batch = next((row for row in read_batches(BytesIO(session.data)) if row['case_id'] == journal['sn']), None)
                recovered = batch and (journal.get('saved') or json.loads(json.dumps(batch_content(batch), default=str)) == journal['snapshot'])
                if not recovered and (batch or journal.get('saved')):
                    raise ClientAddRequestEnded(
                        'This client was already added and later changed or removed. '
                        'Check Companies before choosing Add to Excel to create a new entry.')
            else:
                recovered = False
            if not recovered:
                if (journal and journal.get('expected_revision', expected) != expected
                        and journal.get('before_sha256') != digest_bytes(session.data)):
                    raise ClientAddRequestEnded(
                        'The workbook changed after this interrupted add. Check Companies for the client, '
                        'then choose Add to Excel again if a new entry is needed.')
                require_revision(ROOT, expected, overrides)
                updated, batch = stage_new_client(session, details)
                if not details.get('sn'):
                    from inward_serials import resequence_added_inward
                    updated, batch, serial_changes = resequence_added_inward(
                        updated, batch, session.data, root=ROOT, plant=plant, request_id=request_id,
                        on_order_warning=lambda warning: setattr(session, 'order_warning', warning))
                from identifier_rules import require_unique, workspace_batches
                staged_overrides = {str(path.resolve()):updated}
                require_unique(batch, plant, 'workbook', batch['case_id'], workspace_batches(ROOT, staged_overrides), read_register(ROOT)['records'])
                journal = {'plant':plant, 'path':str(path), 'input':signature, 'sn':batch['case_id'],
                           'snapshot':json.loads(json.dumps(batch_content(batch), default=str)),
                           'expected_revision':expected, 'before_sha256':digest_bytes(session.data), 'saved':False}
                session.before_persist = lambda _: require_revision(ROOT, expected, overrides)
                save_json(journal_path, journal)
                session._persist(updated)
                journal['saved'] = True
                try:
                    save_json(journal_path, journal)
                except OSError:
                    pass  # The prepared record verifies a committed add if its response is lost.
            saved_batches = read_batches(BytesIO(session.data))
            ordering_warning = session.order_warning or ''
    state = cached_workspace_state(expected)
    if state is not None:
        apply_inward_number_changes(state, plant, serial_changes, saved_batches)
        row_batch = dict(batch, source_id=plant, source_path=str(path),
                         _effective_certificate_status=batch.get('issued', ''))
        report = Report(ROOT, date.today(), 'Add client', publish_latest=False)
        report.add_saved_inward(row_batch, 'ready' if status_value(batch.get('issued')) == 'ready' else 'needs_review',
                               'Client added to Excel.')
        state = merge_added_workbook_row(expected, report.rows[0], state)
        try:
            latest = latest_result()
            apply_inward_number_changes(latest, plant, serial_changes, saved_batches)
            cached_rows = latest.setdefault('rows', [])
            key = (plant, str(batch['case_id']).strip().casefold())
            existing = next((item for item in cached_rows if item.get('source_kind') != 'manual'
                             and (item.get('plant'), str(item.get('sn')).strip().casefold()) == key), None)
            if existing is None:
                cached_rows.append(report.rows[0])
            else:
                existing.update(report.rows[0])
            latest['scan_complete'] = True
            save_json(RESULTS_PATH, latest)
        except (OSError, ValueError):
            state['setup_warning'] = 'Client added to Excel. The dashboard will update when you refresh.'
        remember_workbook_stamp()
    else:
        try:
            state = refresh()
        except Exception as exc:
            try:
                state = latest_result()
            except Exception:
                state = {'rows':[], 'scan_complete':False}
            state = dict(state, setup_warning=f'Client added to Excel, but the dashboard could not refresh: {exc}. Refresh before adding it again.')
    return {'client_added':True, 'client':{'plant':plant, 'sn':batch['case_id'],
            'supplier':batch['supplier'], 'status':status_value(batch['issued'])},
            'order_warning':ordering_warning, 'state':state}


def latest_result():
    if not RESULTS_PATH.is_file():
        return {"rows": [], "scan_complete": False, "created": 0, "review": 0}
    value = json.loads(RESULTS_PATH.read_text(encoding="utf-8"))
    if not isinstance(value, dict) or not isinstance(value.get("rows"), list):
        raise ValueError("The latest work list is unreadable. Refresh the workbooks to rebuild it.")
    return value


def reusable_workbook_report(report):
    # A failed save or reminder can replace the report without changing the
    # workbooks. A file can also become writable again without a timestamp change.
    return (report.get('scan_complete') is True
            and not report.get('refresh_pending')
            and not any(row.get('workbook_read_only') for row in report['rows']))


def refresh(paths=None, *, force=False):
    settings = load_settings(ROOT)
    if not any(value.strip() for value in settings['workbook_paths']):
        with run_lock(ROOT):
            recover_issuances(ROOT)
            inventory = scan_inventory(ROOT, persist=True)
        state = disconnected_workbook_state(inventory)
        save_json(RESULTS_PATH, state)
        WORKBOOK_CHECK_STAMPS[str(ROOT)] = ('disconnected', workbook_check_stamp([]))
        return state
    try:
        paths = paths or configured_paths(settings, allow_unavailable=True)
    except ValueError as exc:
        with run_lock(ROOT):
            recover_issuances(ROOT)
            inventory = scan_inventory(ROOT, persist=True)
        return {"rows": current_pdf_rows(manual_rows(ROOT)), "scan_complete": False, "created": 0,
                "review": 0, "setup_warning": str(exc), 'certificate_inventory': add_certificate_type_choices(inventory)}
    try:
        stamp = workbook_check_stamp(paths)
        if force or WORKBOOK_CHECK_STAMPS.get(str(ROOT)) != stamp:
            if _reconcile_linked_completions():
                stamp = workbook_check_stamp(paths)
        if (not force and WORKBOOK_CHECK_STAMPS.get(str(ROOT)) == stamp and RESULTS_PATH.is_file()):
            previous = latest_result()
            if reusable_workbook_report(previous) and isinstance(previous.get('certificate_inventory'), dict):
                state = workspace_state(recheck_obsolete=False, inventory=previous['certificate_inventory'])
                state['already_up_to_date'] = True
                return state
    except (OSError, ValueError, TypeError):
        # Fall through to the full check whenever file metadata is uncertain.
        pass
    checked = {}
    run(paths, ROOT, check_only=True, progress=None, report_sink=checked)
    state = workspace_state(recheck_obsolete=False, inventory=checked.get('certificate_inventory'))
    if reusable_workbook_report(state):
        try:
            WORKBOOK_CHECK_STAMPS[str(ROOT)] = workbook_check_stamp(paths)
        except OSError:
            WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
            state['scan_complete'] = False
            state['setup_warning'] = 'An Excel file became unavailable while refreshing. Save and close it in Excel, check its path in Settings, then refresh again. No certificate was created.'
    else:
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    return state


def disconnected_workbook_state(inventory=None):
    from run_certificates import numbering_plan
    reset_path = ROOT / 'Data' / 'History Reset.json'
    inventory = inventory if inventory is not None else scan_inventory(ROOT)
    return {'rows': current_pdf_rows(manual_rows(ROOT)), 'scan_complete': False,
            'numbering': numbering_plan([], load_settings(ROOT), ROOT / 'Certificates', inventory=inventory),
            'workbooks_disconnected': True, 'created': 0, 'review': 0,
            'certificate_inventory': add_certificate_type_choices(inventory),
            'certificate_folder': str(ROOT / 'Certificates'),
            'history_reset': json.loads(reset_path.read_text(encoding='utf-8')).get('id', '') if reset_path.exists() else ''}


def workspace_state(recheck_obsolete=True, inventory=None):
    refresh_pending = False
    try:
        reconcile_manual(ROOT)
    except LockBusyError:
        refresh_pending = True
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    if not any(value.strip() for value in load_settings(ROOT)['workbook_paths']):
        return disconnected_workbook_state()
    report = latest_result()
    reset_path = ROOT / 'Data' / 'History Reset.json'
    report['history_reset'] = json.loads(reset_path.read_text(encoding='utf-8')).get('id', '') if reset_path.exists() else ''
    canonicalized = False
    for row in report.get("rows", []):
        status = status_value(row.get("certificate_status"))
        if row.get("source_kind") == "manual" and status == "generated":
            status = "generated_manual"
        if status != row.get("certificate_status"):
            row["certificate_status"] = status
            canonicalized = True
    if canonicalized and not refresh_pending:
        save_json(RESULTS_PATH, report)
    inventory = inventory if isinstance(inventory, dict) else scan_inventory(ROOT)
    history = audit_records(ROOT)
    workbook_rows = [row for row in report['rows'] if row.get('source_kind') != 'manual']
    rows = current_pdf_rows(workbook_rows + manual_rows(ROOT, history=history), inventory=inventory, history=history)
    if recheck_obsolete and not refresh_pending:
        missing_pdf = any(row.get('certificate_status') in {'generated', 'generated_manual', 'completed'}
                          and row.get('outcome') != 'needs_review' and not row.get('pdf_filename')
                          for row in rows[:len(workbook_rows)])
        obsolete_template = any(
            obsolete_template_warning(row.get('detail', ''))
            or any(obsolete_template_warning(warning) for warning in (row.get('warnings') or []))
            for row in workbook_rows)
        if missing_pdf or obsolete_template:
            # Re-evaluate real data instead of just hiding the warning and leaving
            # stale attention flags (or incorrectly declaring the row valid).
            # Missing PDFs also invalidate cached Generated/Completed states.
            # Refresh is check-only; it cannot generate a certificate.
            return refresh(force=True)
    # Refresh already scanned and checked every saved PDF. Reuse that snapshot
    # when it is available instead of hashing the same files again.
    report['certificate_inventory'] = add_certificate_type_choices(inventory)
    report['rows'] = rows
    if refresh_pending:
        report['refresh_pending'] = True
    return report


def rebuild_history_state():
    """Rebuild the dashboard after cleanup without rewriting the Excel files."""
    if not all(value.strip() for value in load_settings(ROOT)['workbook_paths']):
        return refresh()
    paths = configured_paths(load_settings(ROOT))
    before = workbook_check_stamp(paths)
    snapshot = {}
    run(paths, ROOT, check_only=True, read_only_preflight=True,
        progress=None, report_sink=snapshot)
    save_json(RESULTS_PATH, snapshot)
    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    state = workspace_state(recheck_obsolete=False)
    if reusable_workbook_report(state) and workbook_check_stamp(paths) == before:
        WORKBOOK_CHECK_STAMPS[str(ROOT)] = before
    return state


def current_pdf_rows(rows, *, inventory=None, history=None):
    """Add validated current filenames to response rows without saving tracking."""
    issued = {'generated', 'generated_manual', 'completed'}
    records = {record['id']: record for record in read_register(ROOT)['records']} if any(
        row.get('source_kind') == 'manual' and status_value(row.get('certificate_status')) in issued
        for row in rows) else {}
    folder = (ROOT / 'Certificates').resolve()
    history = audit_records(ROOT) if history is None else history
    receipts = {(entry.get('source_kind'), entry.get('plant'),
                 str(entry.get('key')).strip().casefold() if entry.get('source_kind') == 'workbook' else str(entry.get('key'))): entry
                for entry in history if entry.get('state') == 'complete'}
    local_receipts = manual_receipts(ROOT, records.values(), history=history)
    intentional_review_cache = None
    result = []
    for original in rows:
        row = dict(original)
        row.pop('pdf_filename', None)  # Cached filenames are not current-file evidence.
        row.pop('pdf_relative_path', None)
        row.pop('has_issued_certificate', None)
        row.pop('last_certificate_details', None)
        row.pop('uploaded_pdf_sha256', None)
        source_kind = row.get('source_kind', 'workbook')
        if source_kind == 'workbook' and row.get('duplicate_serial_number'):
            # A shared Serial Number cannot safely identify either inward's
            # issuance history. Keep both rows available for workbook repair.
            result.append(row)
            continue
        key = row.get('record_id') if source_kind == 'manual' else row.get('sn')
        receipt = (local_receipts.get(key) if source_kind == 'manual'
                   else receipts.get((source_kind, row.get('plant'), str(key).strip().casefold())))
        conflict = workbook_receipt_conflict(row, receipt) if source_kind == 'workbook' else ''
        if conflict:
            row['record_identity_conflict'] = conflict
            row['issuance_blocked'] = True
            row['system_alert'] = True
            row['detail'] = conflict
            row['warnings'] = list(dict.fromkeys([*(row.get('warnings') or []), conflict]))
            result.append(row)
            continue
        # A receipt is historical evidence for restoring edit details. It does
        # not mean a current PDF still exists: recovery changes the workbook
        # status back to Ready after the file is deleted.
        if status_value(row.get('certificate_status')) in issued or receipt:
            row['has_issued_certificate'] = True
        if source_kind == 'workbook' and receipt and isinstance(receipt.get('after'), dict):
            row['last_certificate_details'] = editor_details(receipt['after'])
            changes = compare_details(row, row['last_certificate_details'])
            if intentional_review_cache is None:
                intentional_review_cache = intentional_reviews(ROOT)
            intentionally_kept = is_intentionally_kept(ROOT, receipt, row, reviews=intentional_review_cache)
            row['excel_details_changed_since_issue'] = bool(changes and not intentionally_kept)
            if changes and not intentionally_kept:
                row['excel_change_preview'] = changes
        result.append(row)
        if status_value(row.get('certificate_status')) not in issued:
            continue
        try:
            if source_kind == 'manual':
                record = records.get(key)
                if not record or record.get('plant') != row.get('plant'):
                    continue
                path = current_record_pdf(ROOT, record, receipt=local_receipts.get(key))
            else:
                path = find_certificate(ROOT / 'Certificates', {
                    'cod_no': row.get('certificate_no'), 'source_id': row.get('plant'), 'case_id': key,
                    'issued': row.get('certificate_status'), 'external_pdf': row.get('external_pdf')},
                    inventory=inventory, receipt=receipt)
            if not isinstance(path, (str, os.PathLike)):
                continue
            path = Path(path)
            current = path.resolve(strict=True)
            if (path.suffix.casefold() != '.pdf' or not current.is_relative_to(folder)
                    or current.name != path.name):
                continue
            relative = current.relative_to(folder).as_posix()
            if safe_certificate_path(ROOT, relative) != current:
                continue
            if receipt:
                recorded = Path(receipt.get('pdf_path', ''))
                if recorded.parts[:1] != ('Certificates',) or len(recorded.parts) < 2:
                    continue
                receipt_path = safe_certificate_path(ROOT, Path(*recorded.parts[1:]).as_posix(), expected_sha=receipt.get('pdf_sha256'))
                if current != receipt_path:
                    continue
            if not current_pdf_is_healthy(current, receipt.get('pdf_sha256') if receipt else None):
                continue
            row['pdf_filename'] = path.name
            row['pdf_relative_path'] = current.relative_to(folder).as_posix()
            uploaded_digest = ((receipt or {}).get('pdf_sha256') if (receipt or {}).get('external_upload')
                               else (row.get('external_pdf') or {}).get('sha256'))
            if uploaded_digest:
                row['uploaded_pdf_sha256'] = uploaded_digest
        except (OSError, ValueError):
            # Unavailable metadata must not turn a stale filename into an Open PDF link.
            continue
    return result


def editor_details(batch):
    """Only form values from the issued snapshot, retaining decimal precision."""
    details = {key: batch.get(key, '') for key in (
        'case_id', 'supplier', 'location', 'inward_date', 'form6', 'reference',
        'certificate_weight', 'cod_no', 'cod_date', 'certificate_type', 'certificate_fields',
        'comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale')}
    for key in ('inward_date', 'cod_date'):
        parsed = parse_date(details[key])
        if parsed:
            details[key] = parsed.isoformat()
    details['items'] = [{key: item.get(key, '') for key in ('item', 'quantity', 'per_piece_weight', 'weight')}
                        for item in batch.get('items', [])]
    if details['items']:
        details.update(details['items'][0])
    return clean_details(details)


def issued_workbook_cod_issue(source, plant, paths, target_batches, receipt=None):
    """Check readable ownership evidence without requiring another file to open."""
    if not receipt and status_value(source.get('issued')) not in {'generated', 'generated_manual', 'completed'}:
        return ''
    batches = [dict(batch, source_id=plant) for batch in target_batches]
    for index, path in enumerate(paths, 1):
        other_plant = f'plant{index}'
        if other_plant == plant:
            continue
        try:
            batches.extend(dict(batch, source_id=other_plant) for batch in read_batches(path))
        except (OSError, ValueError, BadZipFile, InvalidFileException):
            # Actual issuance still requires both numbering sources. Completion
            # can verify its existing receipt while another file is unavailable.
            continue
    ownership_view = dict(source, source_id=plant,
                          cod_no=source.get('cod_no') or (receipt or {}).get('cod'))
    return workbook_cod_ownership_issue(ownership_view, batches, read_register(ROOT)['records'])


def workbook_form(body, details, paths, *, saving=False):
    """Validate the inward identity and prepare the requested certificate edit scope."""
    if body.get('void_reissue'):
        raise ValueError('Certificate replacement with a new COD is no longer available. Corrections keep the current COD.')
    if details.get('certificate_type') == 'destruction':
        raise ValueError('Open Review & edit and choose Data Destruction needed. Its HDD certificate is created separately.')
    selected = selected_key(body.get('plant'), details.get('case_id'))
    if body.get('expected_revision') is not None:
        require_revision(ROOT, body['expected_revision'])
    if body.get('source_sn') is not None and (
            str(body['source_sn']) != selected[1] or body.get('source_plant') != selected[0]):
        raise ValueError('The selected Serial Number and plant cannot be changed. Open the correct record instead.')
    source_batches = read_batches(paths[int(selected[0][-1]) - 1])
    matches = [batch for batch in source_batches
               if str(batch['case_id']).strip().casefold() == selected[1].casefold()]
    if len(matches) > 1:
        rows = ', '.join(str(batch['source_row']) for batch in matches)
        raise ValueError(f'Serial Number {selected[1]} is repeated in Excel rows {rows}. '
                         'Open the workbook and give each record a different Serial Number before creating a certificate.')
    source = next(iter(matches), None)
    if body.get('expected_revision') is not None:
        require_revision(ROOT, body['expected_revision'])
    if source is None:
        if body.get('source_sn') is not None:
            raise ValueError('This inward no longer exists. Refresh the company list before continuing.')
        if details['certificate_type'] == 'standard' and not details.get('form6'):
            raise ValueError('Enter the Form 6 number to create this certificate.')
        if str((body.get('details') or {}).get('cod_no') or '').strip():
            raise ValueError('The COD is assigned automatically from the workbooks and certificate PDF history. It cannot be entered or changed here.')
        if str((body.get('details') or {}).get('cod_date') or '').strip():
            raise ValueError('The certificate date is assigned automatically for first generation.')
        details['cod_no'] = ''
        details['cod_date'] = ''
        return details, body.get('edit_scope', 'update_excel')
    if body.get('create_new'):
        raise ValueError('This Serial Number is already used in this plant. Choose the company from the list to edit it, or enter a different Serial Number.')
    if status_value(source.get('issued')) == 'na':
        raise ValueError('This inward is marked N/A and cannot be issued a certificate.')
    details['case_id'] = source['case_id']
    receipt = latest_issuance(ROOT, 'workbook', *selected)
    conflict = workbook_receipt_conflict(source, receipt)
    if conflict:
        raise ValueError(conflict)
    type_changed = details['certificate_type'] != source.get('certificate_type', 'standard')
    requested_scope = body.get('edit_scope', 'certificate_only')
    if type_changed and details['certificate_type'] in {'sita', 'siemens'} and body.get('review_edits') is True:
        details['form6'] = ''
    if saving and type_changed:
        if requested_scope == 'update_excel':
            if body.get('confirmed_excel') is not True:
                raise ValueError('Choose Yes, update Excel before changing this inward’s certificate type.')
        elif requested_scope == 'remove_excel':
            if (receipt or source.get('hdd_cod_no')
                    or status_value(source.get('issued')) in {'generated', 'generated_manual', 'completed'}):
                raise ValueError('A certificate has already been generated for this inward. It cannot be deleted.')
            if (body.get('confirmed_excel_removal') is not True or body.get('source_sn') is None
                    or not body.get('expected_revision')):
                raise ValueError('Confirm deleting this selected Excel inward and keeping the PDF before saving.')
        else:
            raise ValueError('The certificate type differs from Excel. Choose Update Excel or Delete Excel inward and keep PDF.')
    elif saving and requested_scope == 'remove_excel':
        raise ValueError('Deleting the inward while saving is only available for a certificate type change.')
    conflict = issued_workbook_cod_issue(source, selected[0], paths, source_batches, receipt)
    if conflict:
        raise ValueError(conflict)
    pdf_only_edits = bool(body.get('pdf_only_edits'))
    if pdf_only_edits and body.get('edit_scope', 'certificate_only') != 'certificate_only':
        raise ValueError('First-generation corrections can change the PDF only; Excel inward details stay unchanged.')
    if not str(details.get('reference') or '').strip():
        raise ValueError('Enter the reference number before creating the certificate.')
    if receipt or status_value(source.get('issued')) in {'generated', 'generated_manual', 'completed'}:
        if pdf_only_edits:
            raise ValueError('Refresh the certificate before editing its PDF.')
        # The issued record is authoritative for COD. The issue date may be
        # corrected after the first certificate exists.
        details['cod_no'] = source.get('cod_no') or (receipt or {}).get('cod', '')
        saved_date = source.get('cod_date') or (receipt or {}).get('certificate_date') or ''
        parsed = parse_date(saved_date)
        submitted_date = str(details.get('cod_date') or '').strip()
        if submitted_date:
            parsed_submitted = parse_date(submitted_date)
            if not parsed_submitted:
                raise ValueError('Enter a valid certificate date.')
            details['cod_date'] = parsed_submitted.isoformat()
        else:
            details['cod_date'] = parsed.isoformat() if parsed else str(saved_date)
        return details, body.get('edit_scope', 'certificate_only')
    if body.get('review_edits') is True:
        if source.get('duplicate_serial_number') or source.get('record_identity_conflict'):
            raise ValueError('Check this inward’s SN and certificate number in Excel, then refresh.')
        if status_value(source.get('issued')) not in {'blank', 'alert', 'ready'}:
            raise ValueError('This inward is not ready for a certificate. Check its status in Excel.')
        if details['certificate_type'] == 'standard' and not details.get('form6'):
            raise ValueError('Enter the Form 6 number to create this certificate.')
        if str(details.get('cod_no') or '').strip() != str(source.get('cod_no') or '').strip():
            raise ValueError('The COD is assigned automatically and cannot be changed here.')
        if parse_date(details.get('cod_date')) != parse_date(source.get('cod_date')):
            raise ValueError('The certificate date is assigned automatically for the first certificate.')
        # The same reviewed form supports edits before and after first issuance.
        # Preview stages them in memory; saving checks the chosen Excel scope.
        details['cod_no'] = source.get('cod_no', '')
        details['cod_date'] = str(source.get('cod_date') or '')
        return details, body.get('edit_scope', 'certificate_only')
    if source.get('blocking_issues'):
        raise ValueError(' '.join(source['blocking_issues']))
    if details['certificate_type'] == 'standard' and not source.get('form6'):
        if not details.get('form6'):
            raise ValueError('Enter the Form 6 number to create this certificate.')
    source_state = status_value(source.get('issued'))
    kind = details['certificate_type']
    if source_state in {'blank', 'alert'}:
        selected_source = dict(source, certificate_type=kind)
        if not inward_details_complete(selected_source):
            issues = source_value_issues(selected_source)
            message = ' '.join(issues) or 'Complete the company, received date and material weights in Excel.'
            raise ValueError(f'Excel details need checking: {message}')
    elif source_state != 'ready':
        raise ValueError('This inward is marked ' + (source.get('issued') or 'with an unknown status') + '. Choose Ready in Excel before creating its first certificate.')
    # Check normalized staged values so harmless date/number formatting is accepted.
    with WorkbookSnapshot(paths[int(selected[0][-1]) - 1]) as session:
        _, proposed = stage_certificate_details(session, details, validate_cells=False)
    protected = ('supplier', 'location', 'inward_date', 'form6', 'reference', 'cod_no', 'certificate_weight')
    for key in protected:
        if key in (body.get('details') or {}) and key in details and proposed.get(key) != source.get(key):
            deferred_form6 = key == 'form6' and not source.get('form6') and bool(details.get('form6'))
            deferred_reference = key == 'reference' and not source.get('reference') and bool(details.get('reference'))
            if not deferred_form6 and not deferred_reference and (not pdf_only_edits or key == 'cod_no'):
                raise ValueError('Excel details are locked for the first certificate. Correct the inward in Excel, or edit the certificate after generating it.')
    if 'cod_date' in (body.get('details') or {}) and parse_date(proposed.get('cod_date')) != parse_date(source.get('cod_date')):
        raise ValueError('The certificate date comes from Excel or is chosen automatically for first generation.')
    material_keys = ('item', 'quantity', 'per_piece_weight', 'weight')
    first_material = next(iter(source.get('items', [])), {})
    for key in material_keys:
        if key not in (body.get('details') or {}):
            continue
        proposed_value = details.get(key, '') if key == 'item' else number(details.get(key))
        if key == 'quantity' and proposed_value is None:
            proposed_value = text(details.get(key)) or None
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
    if not source.get('form6'):
        # Persist this deferred required field with issuance, independently of
        # any company, weight or material corrections intended for the PDF.
        canonical['form6'] = details['form6']
    if not source.get('reference'):
        canonical['reference'] = details['reference']
    return canonical, requested_scope if saving and type_changed else 'certificate_only'


def _write_preview_image(document, index, output):
    page = document[index]
    bitmap = None
    try:
        # Large imported pages must not require an unbounded bitmap allocation.
        scale = min(1.5, 2400 / max(page.get_size()))
        bitmap = page.render(scale=scale)
        bitmap.to_pil().save(output, format='PNG')
    finally:
        if bitmap is not None:
            bitmap.close()
        page.close()


def _render_preview_image(document, index, destination):
    temporary = destination.with_name(f'.{uuid.uuid4().hex}.png')
    try:
        _write_preview_image(document, index, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def uploaded_pdf_bytes(relative, digest):
    if not isinstance(digest, str) or not re.fullmatch(r'[a-f0-9]{64}', digest):
        raise ValueError('Refresh Companies and open the saved certificate again.')
    path = safe_certificate_path(ROOT, relative, expected_sha=digest)
    if not 0 < path.stat().st_size <= 50_000_000:
        raise ValueError('This PDF is missing or too large. Check the current file in Saved certificates.')
    payload = path.read_bytes()
    if digest_bytes(payload) != digest:
        raise ValueError('This PDF changed. Refresh and check the current file in Saved certificates.')
    return payload


def preview_page_image(identifier, page_number):
    if not re.fullmatch(r'[a-f0-9]{32}', str(identifier)) or type(page_number) is not int or page_number < 1:
        raise ValueError('This preview page was not found.')
    folder = ROOT / 'Data' / 'Previews'
    from preview_store import touch_preview
    touch_preview(ROOT, identifier)
    source = folder / f'{identifier}.pdf'
    if not source.is_file():
        raise ValueError('This preview is no longer available. Preview the certificate again.')
    destination = folder / f'{identifier}_page_{page_number}.png'
    if destination.is_file():
        return destination
    import pypdfium2 as pdfium
    with pdfium.PdfDocument(str(source)) as document:
        if page_number > len(document):
            raise ValueError('This preview page was not found.')
        _render_preview_image(document, page_number - 1, destination)
    return destination


def editor_preview(body):
    """Reuse a preview only for identical inputs and an unchanged workspace."""
    session = body.get('preview_session')
    if not isinstance(session, str) or not re.fullmatch(r'[A-Za-z0-9_-]{1,80}', session):
        return preview_certificate(body)
    try:
        revision = workspace_revision(ROOT)
    except (OSError, ValueError):
        return preview_certificate(body)
    key = (ROOT, session)
    inputs = {name: value for name, value in body.items() if name not in {'preview_client', 'preview_session'}}
    signature = (date.today(), revision, inputs)
    cached = PREVIEW_CACHE.get(key)
    if cached and cached[0] == signature:
        _, identifier, response = cached
        folder = ROOT / 'Data' / 'Previews'
        required = [folder / f'{identifier}.{suffix}' for suffix in ('pdf', 'png')]
        required.extend(folder / f'{identifier}_page_{index}.png'
                        for index in range(2, len(response['preview_detail_images']) + 2))
        if all(path.is_file() for path in required):
            from preview_store import retain_preview
            retain_preview(ROOT, identifier, session, body.get('preview_client'))
            PREVIEW_CACHE.move_to_end(key)
            return response
    response = preview_certificate(body)
    identifier = Path(response['preview_url']).stem
    try:
        unchanged = revision == workspace_revision(ROOT)
    except (OSError, ValueError):
        unchanged = False
    if unchanged:
        PREVIEW_CACHE[key] = (signature, identifier, response)
        PREVIEW_CACHE.move_to_end(key)
        while len(PREVIEW_CACHE) > 32:
            PREVIEW_CACHE.popitem(last=False)
    return response


def preview_certificate(body):
    source_kind = body.get("source_kind", "workbook")
    if source_kind not in ("workbook", "manual"):
        raise ValueError("Choose the plant workbook or local certificate register.")
    if source_kind == 'manual':
        validate_manual_identifiers(body)
    validate_certificate_weight_override(body, source_kind)
    details = clean_details(body.get("details"))
    if source_kind == "workbook" and not body.get("use_certificate_weight"):
        details.pop("certificate_weight", None)
    if source_kind == 'workbook':
        details, _ = workbook_form(body, details, configured_paths())
    batch, cod, issued_date, warnings = prepare_certificate(
        ROOT, body.get("plant"), details, record_id=body.get("record_id") if source_kind == "manual" else None,
        source_kind=source_kind,
        starting_cod=body.get('starting_cod'),
    )
    if source_kind == "workbook":
        paths = configured_paths()
        # Use the same staged material rows as final generation. Preview never
        # writes them, and corrections should not inherit stale source warnings.
        with WorkbookSnapshot(paths[int(body["plant"][-1]) - 1]) as session:
            _, staged = stage_certificate_details(session, details, validate_cells=False)
        batch.update(total_weight=staged['total_weight'], display_weight=staged['display_weight'],
                     certificate_weight=staged.get('certificate_weight'))
        if batch['display_weight'] and batch['display_weight'] > 0:
            warnings = [warning for warning in warnings if not warning.startswith('Weight is missing, invalid or not positive;')]
        warnings.extend(staged.get('issues', []))
    identifier = uuid.uuid4().hex
    destination = ROOT / "Data" / "Previews" / f"{identifier}.pdf"
    make_template_certificate(batch, cod, issued_date, destination, allow_warnings=True)
    import pypdfium2 as pdfium
    detail_images = []
    attached_preview = None
    with pdfium.PdfDocument(str(destination)) as document:
        preview_pages = 1
        if batch['certificate_type'] == 'destruction':
            from destruction_certificate import receipts
            try:
                preview_pages = 1 + max(1, (len(receipts(batch['certificate_fields'].get('device_receipts'))) + 8) // 9)
            except ValueError:
                preview_pages = 2
        for index in range(min(preview_pages, len(document))):
            name = f'{identifier}.png' if index == 0 else f'{identifier}_page_{index + 1}.png'
            _render_preview_image(document, index, destination.parent / name)
            if index:
                detail_images.append(f'/previews/{identifier}/page/{index + 1}.png')
        if batch['certificate_fields'].get('_annexure') and len(document) > preview_pages:
            from pdf_attachments import attachment_catalog
            attachment = attachment_catalog(ROOT).get(batch['certificate_fields']['_annexure'], {})
            attached_preview = {'name': attachment.get('name') or 'Attached PDF',
                                'pages': [f'/previews/{identifier}/page/{index + 1}.png'
                                          for index in range(preview_pages, len(document))]}
    from preview_store import retain_preview, cleanup_previews
    retain_preview(ROOT, identifier, body.get('preview_session'), body.get('preview_client'))
    cleanup_previews(ROOT)
    printed = display_values(batch, cod, issued_date)
    return {"preview_url": f"/previews/{identifier}.pdf", 'preview_image_url': f'/previews/{identifier}.png', "certificate_no": cod,
            "certificate_date": issued_date.isoformat(),
            'preview_detail_images': detail_images,
            'preview_attachment': attached_preview,
            "numbering": batch.get('_numbering'),
            "printed_details": {key: printed[key] for key in ('company_display', 'certificate_weight')},
            "warnings": list(dict.fromkeys(warnings))}


def selected_key(plant, case_id):
    if plant not in ("plant1", "plant2") or not str(case_id or "").strip():
        raise ValueError("Choose a plant and enter a Serial Number.")
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
                 if row.get('source_kind') != 'manual'
                 and row.get("plant") == selected[0]
                 and str(row.get("sn", "")) == selected[1]), None)


def workbook_check_error(report):
    problems = [row for row in report.get('rows', []) if row.get('source_kind') == 'system' and row.get('detail')]
    source_problem = next((row for row in problems if row.get('issue_code') == 'source_unavailable'
                           and row.get('plant') in {'plant1', 'plant2'}), None)
    if source_problem:
        return source_problem['detail']
    if problems:
        return problems[0]['detail']
    return 'Both plant files must be checked before creating a certificate. Refresh and review the workbook message.'


def open_selected_workbook_record(plant, case_id, source_sheet=None, source_row=None):
    """Select a unique live record, or open its workbook for manual repair."""
    selected = selected_key(plant, case_id)
    if source_row is not None:
        try:
            source_row = int(source_row)
        except (TypeError, ValueError) as exc:
            raise ValueError('Refresh the workbooks and select the row again.') from exc
        row = next((entry for entry in latest_result().get('rows', [])
                    if entry.get('plant') == selected[0]
                    and str(entry.get('sn', '')) == selected[1]
                    and entry.get('row') == source_row
                    and (not source_sheet or entry.get('sheet') == source_sheet)), None)
    else:
        row = target_row(latest_result(), selected)
    if not row or not row.get("file"):
        raise ValueError("Refresh the workbooks and select the row again.")
    paths = configured_paths(allow_unavailable=True)
    workbook = paths[int(plant[-1]) - 1]
    if not workbook.is_file():
        raise FileNotFoundError(f'Plant {plant[-1]} workbook could not be found. Check its path in Settings.')
    if Path(row["file"]).expanduser().resolve() != workbook:
        raise ValueError("The selected workbook changed. Refresh both plant workbooks and try again.")
    try:
        matches = [batch for batch in read_batches(workbook)
                   if str(batch.get("case_id")).strip().casefold() == selected[1].casefold()
                   and (not row.get("sheet") or batch.get("source_sheet") == row.get("sheet"))]
    except (ValueError, OSError, BadZipFile):
        open_excel_workbook(workbook)
        return workbook, row.get('sheet', ''), None
    if source_row is not None:
        matches = [batch for batch in matches if batch.get('source_row') == source_row]
    if source_sheet:
        matches = [batch for batch in matches if batch.get('source_sheet') == source_sheet]
    if not matches:
        open_excel_workbook(workbook)
        return workbook, row.get('sheet', ''), None
    if len(matches) > 1:
        return open_excel_record(workbook, matches[0]['source_sheet'], matches[0]['source_row'],
                                 case_id=selected[1], header_row=matches[0].get('source_header_row', 1),
                                 allow_duplicate_repair=source_row is None,
                                 preferred_row=source_row)
    batch = matches[0]
    return open_excel_record(workbook, batch["source_sheet"], batch["source_row"],
                             case_id=batch['case_id'], header_row=batch.get('source_header_row',1),
                             preferred_row=source_row)


def _linked_hdd_pending(plant, sn, source):
    """Whether this Excel inward still has a requested, unfinished HDD certificate."""
    key = str(sn).strip().casefold()
    records = [record for record in read_register(ROOT).get('records', [])
               if (link := record.get('details', {}).get('destruction_source'))
               and link.get('plant') == plant
               and str(link.get('case_id', '')).strip().casefold() == key]
    if not records:
        # An HDD number without its local record is unresolved history. Treat it
        # as pending instead of silently declaring the whole inward complete.
        return bool(source.get('hdd_cod_no'))
    for record in records:
        status = status_value(record.get('status'))
        issued = status in {'generated', 'generated_manual', 'completed'}
        if not record.get('destruction_requested', True) and not issued:
            continue
        if status not in {'completed', 'na'}:
            return True
    return False


def _sync_workbook_recycling_completion(plant, sn):
    """Keep Excel's inward status aligned with the two linked certificates."""
    try:
        if plant not in {'plant1', 'plant2'} or not str(sn or '').strip():
            return False
        path = configured_paths(allow_unavailable=True)[int(plant[-1]) - 1]
        if not path.is_file():
            return False
        with run_lock(ROOT):
            recover_issuances(ROOT)
            with WorkbookSession(path, BACKUP_DIR) as session:
                source = _selected_inward(session, sn)
                current = status_value(source.get('issued'))
                pending = _linked_hdd_pending(plant, sn, source)
                recycling_done = bool(source.get('recycling_completed'))
                target = source.get('issued')
                if current == 'completed' and pending:
                    receipt = latest_issuance(ROOT, 'workbook', plant, str(sn))
                    target = ('Generated Externally'
                              if source.get('external_pdf') or (receipt or {}).get('external_upload')
                              else 'Generated')
                    recycling_done = True  # Repair a completion saved before this coordination existed.
                elif recycling_done and not pending:
                    target = 'Completed'
                    recycling_done = False
                elif recycling_done and pending:
                    receipt = latest_issuance(ROOT, 'workbook', plant, str(sn))
                    target = ('Generated Externally'
                              if source.get('external_pdf') or (receipt or {}).get('external_upload')
                              else 'Generated')
                else:
                    return False
                session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                    source['cod_no'], source['issued'], target,
                    metadata={'_recycling_completed': recycling_done})
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        return True
    except (OSError, ValueError, LockBusyError):
        # The certificate status change has already been saved in its own
        # register. A later refresh can retry the Excel reconciliation.
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        return False


def _sync_linked_recycling_completion(record_id):
    """Reconcile an inward after its linked Data Destruction record changes."""
    try:
        record = next((item for item in read_register(ROOT).get('records', [])
                       if item.get('id') == record_id), None)
        if not record:
            return False
        link = record.get('details', {}).get('destruction_source') or {}
        return _sync_workbook_recycling_completion(link.get('plant'), link.get('case_id'))
    except (OSError, ValueError):
        return False


def _reconcile_linked_completions():
    """Repair stale aggregate states, including a removed unissued HDD request."""
    try:
        records = read_register(ROOT).get('records', [])
    except (OSError, ValueError):
        return False
    targets = set()
    for record in records:
        details = record.get('details', {})
        link = details.get('destruction_source') or {}
        if details.get('certificate_type') == 'destruction' and link.get('plant') in {'plant1', 'plant2'}:
            targets.add((link['plant'], str(link.get('case_id', '')).strip()))
    # A user may delete an unissued request from Settings. The workbook marker
    # still records that Recycling was sent, so revisit those rows even though
    # the now-deleted HDD record can no longer lead us back to them.
    try:
        for index, path in enumerate(configured_paths(allow_unavailable=True), 1):
            if not path.is_file():
                continue
            for batch in read_batches(BytesIO(path.read_bytes())):
                if batch.get('recycling_completed'):
                    targets.add((f'plant{index}', str(batch.get('case_id', '')).strip()))
    except (OSError, ValueError, InvalidFileException):
        pass
    changed = False
    for plant, sn in targets:
        changed = _sync_workbook_recycling_completion(plant, sn) or changed
    return changed


def complete_workbook_record(body):
    """Verify live record ownership while holding the workbook's write lock."""
    undo = body.get('undo_completed') is True
    with run_lock(ROOT):
        recover_issuances(ROOT)
        require_revision(ROOT, body.get('expected_revision'))
        selected = selected_key(body.get('plant'), body.get('sn'))
        completed_state = latest_result()
        row = target_row(completed_state, selected)
        if not row or not row.get('file'):
            raise ValueError('Refresh the workbooks and select the row again.')
        paths = configured_paths(allow_unavailable=True)
        path = paths[int(selected[0][-1]) - 1]
        if not path.is_file():
            raise FileNotFoundError(f'Plant {selected[0][-1]} workbook could not be found. Check its path in Settings.')
        if Path(row['file']).resolve() != path:
            raise ValueError('The selected workbook changed. Refresh the workbooks and try again.')
        with WorkbookSession(path, BACKUP_DIR) as session:
            require_revision(ROOT, body.get('expected_revision'), {str(path.resolve()): session.data})
            source_batches = read_batches(BytesIO(session.data))
            matches = [batch for batch in source_batches
                       if batch['case_id'].strip().casefold() == selected[1].casefold()]
            if len(matches) != 1:
                raise ValueError('The Serial Number is missing or repeated. Open Excel and correct it before marking Completed.')
            source = matches[0]
            receipt = latest_issuance(ROOT, 'workbook', *selected)
            conflict = workbook_receipt_conflict(source, receipt)
            if conflict:
                raise ValueError(conflict)
            conflict = issued_workbook_cod_issue(source, selected[0], paths, source_batches, receipt)
            if conflict:
                raise ValueError(conflict)
            if not undo and not find_certificate(ROOT / 'Certificates', dict(source, source_id=selected[0])):
                raise ValueError('The current PDF is missing, changed or unreadable. Refresh and recreate it before marking it Completed.')
            restored = 'Generated Externally' if source.get('external_pdf') or (receipt or {}).get('external_upload') else 'Generated'
            current_status = status_value(source.get('issued'))
            if undo:
                if source.get('recycling_completed') or current_status == 'completed':
                    target = restored
                    recycling_done = False
                else:
                    raise ValueError('This Recycling certificate is not marked Completed. Refresh and try again.')
            else:
                if current_status not in {'generated', 'generated_manual'}:
                    raise ValueError('Create the current certificate PDF before marking it Completed.')
                hdd_pending = _linked_hdd_pending(selected[0], selected[1], source)
                target = source['issued'] if hdd_pending else 'Completed'
                recycling_done = hdd_pending
            session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                                      source.get('cod_no', ''), source['issued'], target,
                                      metadata={'_recycling_completed': recycling_done})
        # The status has committed. Keep its acknowledgement even if rereading
        # another workbook or rebuilding the report fails afterwards.
        status = status_value(target)
        row.update(certificate_status=status, outcome=status, reminder_due=False, overdue=False,
                   recycling_completed_pending_hdd=recycling_done,
                   detail=('Recycling has been sent; Data Destruction is still pending.' if recycling_done else
                           'Completion undone. This certificate can be sent and marked Completed again.' if undo else
                           f'COD {source.get("cod_no", "")} is marked Completed.'))
        try:
            save_json(RESULTS_PATH, completed_state)
        except OSError:
            pass  # Return the committed state below even when the cache is unavailable.
        state = None if undo else merge_workbook_report(body.get('expected_revision'), completed_state)
        if state is None:
            WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        else:
            remember_workbook_stamp()
    if state is None:
        return post_commit_state(refresh, saved_message='Completion undone' if undo else 'Completed status saved successfully',
                                 fallback_state=completed_state)
    return state


def generate(*, selected=None, plants=None, confirmed=False, expected_revision=None, override_reason='',
             starting_cod=None, selections=None, review_only=False):
    selected_keys = None
    if selections is not None:
        if selected or plants is not None or not isinstance(selections, list) or not 1 <= len(selections) <= 500 or any(not isinstance(item, dict) for item in selections):
            raise ValueError('Choose between 1 and 500 certificates from the list.')
        selected_keys = [selected_key(item.get('plant'), item.get('sn')) for item in selections]
        if len(set(selected_keys)) != len(selected_keys):
            raise ValueError('A client was selected twice. Choose it only once.')
    settings = load_settings(ROOT)
    paths = configured_paths(settings)
    preflight = {}
    preflight_args = {"check_only": True, "read_only_preflight": True, "allow_warnings": True,
                      "progress": None, 'report_sink': preflight}
    if starting_cod is not None:
        preflight_args['starting_cod'] = starting_cod
    if expected_revision is not None:
        preflight_args['expected_revision'] = expected_revision
    if selected:
        preflight_args["selected"] = selected
    if selected_keys is not None:
        preflight_args['selected_keys'] = selected_keys
    if plants is not None:
        preflight_args["plants"] = plants
    run(paths, ROOT, **preflight_args)
    preflight = preflight or latest_result()
    if not preflight.get('scan_complete'):
        raise ValueError(workbook_check_error(preflight))
    eligible = [row for row in preflight.get('rows', [])
                if row.get('certificate_status') == 'ready' and row.get('outcome') == 'ready'
                and not row.get('issuance_blocked')]
    if selected_keys is not None:
        for key in selected_keys:
            target = target_row(preflight, key)
            if target is None or target not in eligible:
                detail = (target or {}).get('detail') or 'Refresh the list and choose the certificates again.'
                raise ValueError(f'SN {key[1]} in {key[0].replace("plant", "Plant ")} is not ready. {detail}')
        eligible = [row for row in eligible if (row.get('plant'), str(row.get('sn'))) in selected_keys]
    if selected:
        target = target_row(preflight, selected)
        if target is None:
            raise ValueError(f"Serial Number {selected[1]} was not found in {selected[0].replace('plant', 'Plant ')}.")
        if target not in eligible:
            raise ValueError(target.get('detail') or 'This certificate is not ready. Review its current details before creating it.')
    if not eligible:
        problem = next((row.get('detail') for row in preflight.get('rows', [])
                        if row.get('outcome') == 'needs_review'), None)
        raise ValueError(problem or 'No certificates are ready to create in this selection.')
    if (preflight.get('numbering', {}).get('needs_start') and starting_cod is None
            and any(not row.get('certificate_no') for row in eligible)):
        raise ValueError('Enter the starting COD number before creating certificates.')
    candidates = [row for row in eligible if row.get('warnings')]
    if review_only or (candidates and not confirmed):
        return {"requires_confirmation": True, "warnings": candidates, "created": 0}

    generated_report = {}
    arguments = {"progress": None, "allow_warnings": confirmed, 'override_reason': override_reason,
                 'report_sink': generated_report}
    if starting_cod is not None:
        arguments['starting_cod'] = starting_cod
    if expected_revision is not None:
        arguments['expected_revision'] = expected_revision
    if selected:
        arguments["selected"] = selected
    if selected_keys is not None:
        arguments['selected_keys'] = selected_keys
    if plants is not None:
        arguments["plants"] = plants
    result = run(paths, ROOT, **arguments)
    generated_report = generated_report or latest_result()
    if not result[0] and not generated_report.get('scan_complete'):
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        raise ValueError(workbook_check_error(generated_report))
    rows = generated_report.get('rows', [])
    reported_rows = {(row.get('plant'), str(row.get('sn'))): row for row in rows}
    generated_rows = {(row.get('plant'), str(row.get('sn'))): row for row in rows
                      if row.get('outcome') == 'generated_now'}
    unissued = []
    for row in eligible:
        key = (row.get('plant'), str(row.get('sn')))
        if key not in generated_rows:
            detail = reported_rows.get(key, {}).get('detail')
            if not detail and generated_report.get('refresh_pending'):
                detail = 'The batch stopped while finishing an earlier certificate. Refresh before retrying.'
            unissued.append({'plant': key[0], 'sn': key[1], 'supplier': row.get('supplier'),
                             'detail': detail or 'This certificate was not created. Refresh the list and review this company.'})
    partial_selection = selected is not None or selected_keys is not None or plants is not None
    state = merge_workbook_report(expected_revision, generated_report, preserve_unmentioned=partial_selection)
    if state is None:
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        state = post_commit_state(lambda: refresh(paths))
    elif not unissued:
        remember_workbook_stamp()
    else:
        # The last issuance report is an outcome, not a fresh check of what can be retried.
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    for row in state.get("rows", []):
        generated_row = generated_rows.get((row.get("plant"), str(row.get("sn"))))
        if generated_row:
            row["detail"] = generated_row.get("detail", row.get("detail", ""))
            row["warnings"] = generated_row.get("warnings", [])
            if generated_row.get("system_alert"):
                row["system_alert"] = True
    state["created"] = result[0]
    if partial_selection:
        try:
            save_json(RESULTS_PATH, state)
        except OSError:
            WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
            state['setup_warning'] = 'Refresh to reload the full company list.'
    return {"requires_confirmation": False, "state": state,
            "created": result[0], "review": result[1], "unissued": unissued,
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
    if scope not in {'certificate_only', 'update_excel', 'remove_excel'}:
        raise ValueError('Choose This certificate only or Also update Excel.')
    if body.get('review_edits') is True and scope == 'update_excel' and body.get('confirmed_excel') is not True:
        raise ValueError('Choose Yes, update Excel before saving these details to Excel.')
    if body.get('source_sn') is not None:
        if str(body['source_sn']) != selected[1] or body.get('source_plant') != plant:
            raise ValueError('The selected Serial Number and plant cannot be changed. Open the correct record instead.')
    paths = configured_paths()
    if body.get('create_new'):
        # The HTTP request lock covers this read; run() rechecks its revision.
        workbook = paths[int(plant[-1]) - 1]
        if any(str(batch['case_id']).strip().casefold() == selected[1].casefold()
               for batch in read_batches(workbook)):
            raise ValueError('This Serial Number is already used in this plant. '
                             'Choose the company from the list to edit it, or enter a different Serial Number.')
    details, scope = workbook_form(body, details, paths, saving=True)
    date_changed = False
    if scope == 'update_excel' and body.get('source_sn') is not None and 'inward_date' in details:
        source = next((batch for batch in read_batches(paths[int(plant[-1]) - 1])
                       if batch['case_id'] == selected[1]), None)
        date_changed = source is not None and source['inward_date'] != parse_date(details['inward_date'])
    snapshot = {}
    arguments = {'selected': selected,
                 'details': details, 'update_excel': scope == 'update_excel', 'progress': None, 'report_sink': snapshot}
    if scope == 'remove_excel':
        arguments.update(remove_inward=True, hdd_needed=body.get('data_destruction_needed'))
    if body.get('starting_cod') is not None:
        arguments['starting_cod'] = body['starting_cod']
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    run(paths, ROOT, check_only=True, read_only_preflight=True, **arguments)
    preflight = dict(snapshot) if snapshot else latest_result()
    if not preflight.get('scan_complete'):
        raise ValueError(workbook_check_error(preflight))
    row = target_row(preflight, selected)
    if row is None:
        error = next((item.get('detail') for item in preflight.get('rows', []) if item.get('system_alert')), None)
        raise ValueError(error or 'The certificate details could not be checked.')
    if row.get('issuance_blocked'):
        raise ValueError(row.get('detail') or 'Correct the Excel import issue before creating this certificate.')
    if (preflight.get('numbering', {}).get('needs_start') and body.get('starting_cod') is None
            and not details.get('cod_no')):
        raise ValueError('Enter the starting COD number before creating this certificate.')
    warnings = [row] if row.get('outcome') == 'needs_review' else []
    if body.get('review_only') is True or (warnings and not body.get('confirmed')):
        return {'requires_confirmation': True, 'warnings': warnings, 'created': 0}
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    result = run(paths, ROOT, allow_warnings=bool(body.get('confirmed')),
                 override_reason=body.get('override_reason', ''), **arguments)
    final_report = snapshot or latest_result()
    final_row = target_row(final_report, selected)
    if not result[0]:
        # A scoped run may have replaced the checked list with just its error.
        WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
        if not final_report.get('scan_complete'):
            raise ValueError(workbook_check_error(final_report))
        raise ValueError((final_row or {}).get('detail') or 'The certificate could not be created. Excel details were not saved.')
    if scope == 'remove_excel':
        state = post_commit_state(rebuild_history_state)
        identifier = final_row.get('converted_record_id')
        local_row = next((entry for entry in state.get('rows', []) if entry.get('record_id') == identifier), None)
        return {'requires_confirmation': False, 'state': state, 'created': result[0], 'review': result[1],
                'record_id': identifier, 'certificates': certificate_outputs([local_row] if local_row else []),
                'warnings': [state['setup_warning']] if state.get('setup_warning') else []}
    state = merge_workbook_report(body.get('expected_revision'), final_report, preserve_unmentioned=True)
    if state is None:
        # Save any companion HDD request against the current SN before a
        # date-order refresh can reassign that SN to another inward.
        state = (post_commit_state(lambda: workspace_state(recheck_obsolete=False), fallback_state=final_report)
                 if date_changed else post_commit_state(lambda: refresh(paths)))
    else:
        try:
            save_json(RESULTS_PATH, state)
        except OSError:
            WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
            state['setup_warning'] = 'The PDF was created. Refresh to reload the full company list.'
        else:
            remember_workbook_stamp()
    state_row = target_row(state, selected)
    if state_row and final_row:
        state_row['detail'] = final_row['detail'] + (' Certificate-only changes: original inward details were kept.'
                                                  if scope == 'certificate_only' else ' Edited details were also saved to Excel.')
        state_row['warnings'] = final_row.get('warnings', [])
    state = save_requested_destruction(body, state)
    if date_changed:
        warning = state.get('setup_warning')
        state = post_commit_state(lambda: refresh(paths, force=True), fallback_state=state)
        if warning and warning != state.get('setup_warning'):
            state['setup_warning'] = ' '.join(filter(None, [warning, state.get('setup_warning')]))
        if final_row:
            current = next((row for row in state.get('rows', [])
                            if row.get('source_kind') == 'workbook' and row.get('plant') == plant
                            and row.get('certificate_no') == final_row.get('certificate_no')), None)
            if current:
                final_row = current
    return {'requires_confirmation': False, 'state': state, 'created': result[0], 'review': result[1],
            'warnings': [state['setup_warning']] if state.get('setup_warning') else [],
            'certificates': certificate_outputs([final_row] if final_row else [])}


def save_requested_destruction(body, state):
    """Mark a separate HDD task after a successful save, without issuing a PDF."""
    if not isinstance(body.get('data_destruction_needed'), bool):
        return state
    try:
        local = None
        if body.get('source_kind') == 'manual':
            local = next((record for record in read_register(ROOT)['records']
                          if record['id'] == body.get('record_id')
                          and record['details'].get('certificate_type') != 'destruction'), None)
            if local is None:
                return state
        source_sn = body.get('source_sn') or (body.get('details') or {}).get('case_id')
        if body['data_destruction_needed'] is False:
            link = {'plant': body.get('plant'), 'case_id': str(source_sn or '')} if local is None else {
                'plant': None, 'case_id': str((local.get('removed_inward') or {}).get('case_id')
                                              or local['details'].get('case_id') or local['id']),
                'record_id': local['id']}
            if not any(record.get('destruction_requested', True)
                       and record['details'].get('destruction_source') == link
                       for record in read_register(ROOT)['records']):
                return state
        from destruction_certificate import request_for_inward
        identifier = request_for_inward(ROOT, body.get('plant'),
            source_sn, workspace_revision(ROOT),
            needed=body['data_destruction_needed'],
            record_id=local['id'] if local else None)
        if identifier:
            workbook_changed = _sync_linked_recycling_completion(identifier)
            if workbook_changed:
                state = post_commit_state(lambda: refresh(force=True),
                                          saved_message='Data Destruction request updated', fallback_state=state)
            return merge_manual_record_state(None, identifier, state) or state
        return state
    except (OSError, ValueError) as exc:
        state['setup_warning'] = ('Your changes were saved, but the Data Destruction request could not be updated. '
                                  f'Open Review & edit, set Data Destruction also needed, then choose Save request. Details: {exc}')
        return state


def save_inward_changes(body):
    """Save an unissued inward without allocating a COD or creating a PDF."""
    if body.get('confirmed_excel') is not True:
        raise ValueError('Choose Yes, update Excel before saving these changes.')
    plant, sn = selected_key(body.get('plant'), body.get('source_sn'))
    details = clean_details(body.get('details'))
    if str(details.get('case_id')) != sn or body.get('source_plant') != plant:
        raise ValueError('The SN and plant cannot be changed. Reopen this inward to try again.')
    if details.get('certificate_type') == 'destruction':
        raise ValueError('Choose Data Destruction needed to add a separate HDD task.')
    paths = configured_paths(allow_unavailable=True)
    path = paths[int(plant[-1]) - 1]
    with run_lock(ROOT):
        recover_issuances(ROOT)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path.resolve()): session.data}
            require_revision(ROOT, body.get('expected_revision'), overrides, force=False)
            matches = [row for row in read_batches(BytesIO(session.data)) if row['case_id'] == sn]
            if len(matches) != 1 or matches[0].get('duplicate_serial_number'):
                raise ValueError('The inward’s SN is missing or repeated. Check Excel, then refresh.')
            source = matches[0]
            current_status = status_value(source.get('issued'))
            if latest_issuance(ROOT, 'workbook', plant, sn) or current_status in {'generated', 'generated_manual', 'completed'}:
                raise ValueError('This certificate already has a PDF. Save a corrected PDF to update its Excel details too.')
            if current_status not in {'blank', 'alert', 'ready', 'na'}:
                raise ValueError('Check this inward’s status in Excel before saving changes.')
            if (details['certificate_type'] in {'sita', 'siemens'}
                    and details['certificate_type'] != source.get('certificate_type', 'standard')):
                details['form6'] = ''
            details['cod_no'], details['cod_date'] = source.get('cod_no', ''), source.get('cod_date', '')
            if not body.get('use_certificate_weight'):
                details.pop('certificate_weight', None)
            no_certificate = current_status == 'na' or supplier_is_always_not_applicable(details.get('supplier'))
            updated, saved = stage_certificate_details(session, details, inward_status='N/A' if no_certificate else 'Ready')
            ready = not no_certificate and inward_details_complete(saved)
            if not no_certificate and not ready:
                from workbook_store import _set_cell, _sheet_member
                from zipfile import ZipFile
                with ZipFile(BytesIO(updated)) as archive:
                    member = _sheet_member(archive, saved['source_sheet'])
                    xml = _set_cell(archive.read(member).decode('utf-8'),
                        f"{saved['column_mapping']['issued']['column']}{saved['source_row']}", 'Alert', 'inlineStr')
                    output = BytesIO()
                    with ZipFile(output, 'w') as target:
                        target.comment = archive.comment
                        for info in archive.infolist():
                            target.writestr(info, xml.encode('utf-8') if info.filename == member else archive.read(info.filename))
                updated = output.getvalue()
                saved = next(row for row in read_batches(BytesIO(updated)) if row['case_id'] == sn)
            from identifier_rules import require_unique, workspace_batches
            require_unique(saved, plant, 'workbook', sn,
                workspace_batches(ROOT, {str(path.resolve()): updated}), read_register(ROOT)['records'])
            serial_changes = {}
            if saved['inward_date'] != source['inward_date']:
                from inward_serials import resequence_date_ordered_inwards
                updated, serial_changes = resequence_date_ordered_inwards(updated, session.data, root=ROOT, plant=plant)
            saved_batches = read_batches(BytesIO(updated))
            saved = next(row for row in saved_batches if row['case_id'] == serial_changes.get(sn, sn))
            session.before_persist = lambda _: require_revision(ROOT, body.get('expected_revision'), overrides)
            session._persist(updated)
    row_batch = dict(saved, source_id=plant, source_path=str(path), _effective_certificate_status=saved['issued'])
    report = Report(ROOT, date.today(), 'Edit inward', publish_latest=False)
    outcome = 'not_applicable' if no_certificate else 'ready' if ready else 'needs_review'
    report.add_saved_inward(row_batch, outcome,
                           'Changes saved to Excel. No PDF has been created.')
    state = cached_workspace_state(body.get('expected_revision'))
    if state is not None:
        apply_inward_number_changes(state, plant, serial_changes, saved_batches)
        state = merge_added_workbook_row(body.get('expected_revision'), report.rows[0], state)
    try:
        latest = latest_result()
        apply_inward_number_changes(latest, plant, serial_changes, saved_batches)
        latest = merge_added_workbook_row(None, report.rows[0], latest)
        save_json(RESULTS_PATH, latest)
    except (OSError, ValueError):
        pass
    state = state if state is not None else post_commit_state(workspace_state,
        saved_message='Changes saved to Excel', fallback_state={'rows': report.rows, 'scan_complete': False})
    state = save_requested_destruction(dict(body, source_sn=saved['case_id']), state)
    remember_workbook_stamp()
    return {'saved': True, 'created': 0, 'state': state, 'sn': saved['case_id']}


def update_inward_status(body):
    action = body.get('action')
    if action not in {'mark_na', 'undo_na'}:
        raise ValueError('Choose Mark N/A or Undo N/A.')
    revision = body.get('expected_revision')
    if body.get('record_id'):
        from certificate_drafts import _manual_record_issued, _manual_status, write_register
        record_id = body['record_id']
        with run_lock(ROOT):
            recover_issuances(ROOT)
            require_revision(ROOT, revision)
            register = read_register(ROOT)
            record = next((item for item in register['records'] if item['id'] == record_id), None)
            if not record or _manual_record_issued(ROOT, record) or record.get('pdf_recovery_pending'):
                raise ValueError('Choose an unissued local certificate record.')
            current = _manual_status(record.get('status'))
            if action == 'mark_na' and current in {'blank', 'draft', 'ready', 'alert'}:
                record.update(status='na')
                record.pop('status_before_na', None)
                target = 'N/A'
            elif action == 'undo_na' and current == 'na':
                record.pop('status_before_na', None)
                record['status'] = _manual_status_from_current_details(record)
                target = {'ready':'Ready', 'alert':'Alert'}[record['status']]
            else:
                raise ValueError('The local certificate status has changed. Refresh and try again.')
            record['updated_at'] = datetime.now().isoformat()
            write_register(ROOT, register)
        linked_workbook_changed = (
            record.get('details', {}).get('certificate_type') == 'destruction'
            and _sync_linked_recycling_completion(record_id))
        state = merge_manual_record_state(revision, record_id)
        if linked_workbook_changed:
            state = post_commit_state(lambda: refresh(force=True),
                                      saved_message='Data Destruction status updated', fallback_state=state)
        return {'saved': True, 'status': target, 'state': state if state is not None else post_commit_state(workspace_state)}
    plant, sn = selected_key(body.get('plant'), body.get('sn'))
    path = configured_paths(allow_unavailable=True)[int(plant[-1]) - 1]
    with run_lock(ROOT):
        recover_issuances(ROOT)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path): session.data}
            require_revision(ROOT, revision, overrides, force=False)
            source = _selected_inward(session, sn)
            status = status_value(source['issued'])
            if (source.get('pdf_recovery_pending') or
                    source.get('external_pdf') or latest_issuance(ROOT, 'workbook', plant, sn)):
                raise ValueError('An existing certificate is linked to this inward. Review its certificate before changing this status.')
            if action == 'mark_na':
                if status not in {'blank', 'alert', 'ready'}:
                    raise ValueError('Only an unissued inward can be marked N/A.')
                target = 'N/A'
            else:
                if status != 'na':
                    raise ValueError('This inward is no longer marked N/A. Refresh and try again.')
                if supplier_is_always_not_applicable(source.get('supplier')):
                    raise ValueError('This company is always marked N/A. Correct the company name first if it was entered by mistake.')
                target = 'Ready' if inward_details_complete(source) else 'Alert'
            session.before_persist = lambda _: require_revision(ROOT, revision, overrides)
            session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                                      source['cod_no'], source['issued'], target)
    fresh = Report(ROOT, date.today(), 'Update inward status', publish_latest=False)
    fresh.add_saved_inward(dict(source, source_id=plant, source_path=str(path), issued=target),
                          {'N/A':'not_applicable', 'Ready':'ready', 'Alert':'needs_review'}[target],
                          f'Excel status saved as {target}.')
    state = cached_workspace_state(revision)
    if state is not None and not source.get('blocking_issues') and reusable_workbook_report(state):
        try:
            report = latest_result()
            if (reusable_workbook_report(report) and target_row(report, (plant, sn))
                    and target_row(state, (plant, sn))):
                for snapshot in (state, report):
                    row = target_row(snapshot, (plant, sn))
                    row.update(copy.deepcopy(fresh.rows[0]))
                    if not fresh.rows[0].get('system_alert'):
                        row.pop('system_alert', None)
                # Recheck issuance eligibility in the next background refresh.
                if target == 'Ready':
                    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
                save_json(RESULTS_PATH, report)
                if target != 'Ready':
                    remember_workbook_stamp()
                return {'saved': True, 'status': target, 'state': state}
        except (OSError, ValueError):
            pass
    return {'saved': True, 'status': target, 'state': post_commit_state(
        refresh, saved_message=f'Excel status saved as {target}',
        fallback_state=state or {'rows':fresh.rows, 'scan_complete':False})}


def _manual_status_from_current_details(record):
    """Recalculate an unissued local record's status from its current details."""
    details = record.get('details') if isinstance(record.get('details'), dict) else {}
    if details.get('certificate_type') == 'destruction':
        from destruction_certificate import content_issues
        return 'ready' if not content_issues(details) else 'alert'

    items, issues = material_items(details)
    inward_date = parse_date(details.get('inward_date'))
    batch = {
        'supplier': details.get('supplier', ''),
        'inward_date': inward_date,
        'items': items,
        'total_weight': sum(item.get('weight') or 0 for item in items),
        'issues': issues,
        'blocking_issues': [],
    }
    return 'ready' if inward_details_complete(batch) else 'alert'


def _selected_inward(session, sn):
    matches = [batch for batch in read_batches(BytesIO(session.data))
               if text(batch['case_id']).casefold() == sn.casefold()]
    if len(matches) != 1 or matches[0].get('duplicate_serial_number'):
        raise ValueError('The inward’s SN is missing or repeated. Check Excel, then refresh.')
    return matches[0]


def link_external_certificate(body):
    plant, sn = selected_key(body.get('plant'), body.get('sn'))
    revision = body.get('expected_revision')
    digest = body.get('sha256')
    path = safe_certificate_path(ROOT, body.get('relative_path'), expected_sha=digest)
    relative = path.relative_to((ROOT / 'Certificates').resolve()).as_posix()
    issued_date = parse_date(body.get('certificate_date'))
    if not issued_date or issued_date > date.today():
        raise ValueError('Enter the certificate’s issue date. It cannot be in the future.')
    paths = configured_paths()
    workbook = paths[int(plant[-1]) - 1]
    with run_lock(ROOT):
        recover_issuances(ROOT)
        require_revision(ROOT, revision)
        inventory = scan_inventory(ROOT)
        entry = next((item for item in inventory['entries']
                      if item.get('present') and item['path'] == f'Certificates/{relative}'), None)
        if (not entry or not entry.get('linkable_external')
                or not valid_pdf(path) or not digest or entry.get('sha256') != digest):
            raise ValueError('Choose a current external PDF from Saved certificates.')
        saved_type = certificate_type_choices().get(digest)
        if saved_type == 'destruction' or (not saved_type and re.search(r'_hdd\.pdf$', path.name, re.I)):
            raise ValueError('This is a Data Destruction PDF. Keep it under Saved certificates with its own certificate number.')
        cods = {cod_identity(value) for value in entry.get('current_cods', [])}
        if len(cods) != 1:
            raise ValueError('The PDF needs one identifiable certificate number. Check the printed number and filename.')
        cod = entry['current_cods'][0]
        if any(item.get('present') and item.get('sha256') != digest and
               cod_identity(cod) in {cod_identity(value) for value in item.get('current_cods', [])}
               for item in inventory['entries']):
            raise ValueError(f'COD {cod} appears on another PDF. Resolve the duplicate before linking.')
        with WorkbookSession(workbook, BACKUP_DIR) as session:
            overrides = {str(workbook): session.data}
            require_revision(ROOT, revision, overrides)
            source = _selected_inward(session, sn)
            status = status_value(source['issued'])
            if (status not in {'blank', 'alert', 'ready', 'generated_manual'} or source.get('pdf_recovery_pending')
                    or latest_issuance(ROOT, 'workbook', plant, sn)):
                raise ValueError('Choose an unissued inward or one marked Generated Externally.')
            if source['cod_no'] and cod_identity(source['cod_no']) != cod_identity(cod):
                raise ValueError(f'This inward already has COD {source["cod_no"]}. Check the PDF and Excel before linking.')
            from identifier_rules import workspace_batches
            batches = workspace_batches(ROOT, overrides)
            if any(cod_identity(row.get('hdd_cod_no')) == cod_identity(cod) for row in batches):
                raise ValueError(f'COD {cod} is already assigned to a Data Destruction certificate.')
            conflict = workbook_cod_ownership_issue(dict(source, source_id=plant, cod_no=cod),
                                                     batches, read_register(ROOT)['records'])
            if conflict:
                raise ValueError(conflict)
            external = {'cod': source['cod_no'] or cod, 'date': issued_date,
                        'link': {'path': relative, 'sha256': digest}}
            session.before_persist = lambda _: require_revision(ROOT, revision, overrides)
            session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                                      source['cod_no'], source['issued'], 'Generated Externally', external=external)
    fresh = Report(ROOT, date.today(), 'Link external certificate', publish_latest=False)
    fresh.add(dict(source, source_id=plant, source_path=str(workbook), issued='Generated Externally',
                   cod_no=external['cod'], cod_date=issued_date, external_pdf=external['link']),
              'generated_manual', f'COD {external["cod"]} was linked to the external PDF.')
    row = fresh.rows[0]
    row.update(pdf_filename=path.name, pdf_relative_path=relative, has_issued_certificate=True)
    fallback = merge_added_workbook_row(revision, row) or {'rows':fresh.rows, 'scan_complete':False}
    return {'saved': True, 'state': post_commit_state(
        refresh, saved_message='External certificate linked', fallback_state=fallback)}


def save_uploaded_certificate(body):
    from external_certificate_upload import save_upload
    result = save_upload(ROOT, BACKUP_DIR, body)
    if result.get('record_id'):
        _sync_linked_recycling_completion(result['record_id'])
    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    message = 'Certificate PDF replaced' if result['replaced'] else 'External certificate saved'
    result['state'] = post_commit_state(refresh, saved_message=message)
    if result['recovery_note']:
        result['state'].update(refresh_pending=True, recovery_note=result['recovery_note'])
    return result


def save_inward_followup(body):
    """Save transport/office details even after completion, without issuing a PDF."""
    from inward_followup import clean_followup, stage_followup
    if body.get('confirmed_excel') is not True:
        raise ValueError('Choose Yes, update Excel before saving these details.')
    plant, sn = selected_key(body.get('plant'), body.get('sn'))
    values = clean_followup(body.get('details'))
    path = configured_paths(allow_unavailable=True)[int(plant[-1]) - 1]
    with run_lock(ROOT):
        recover_issuances(ROOT)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path.resolve()): session.data}
            require_revision(ROOT, body.get('expected_revision'), overrides, force=False)
            updated, saved = stage_followup(session.data, sn, values)
            session.before_persist = lambda _: require_revision(ROOT, body.get('expected_revision'), overrides)
            if updated != session.data:
                session._persist(updated)
            else:
                require_revision(ROOT, body.get('expected_revision'), overrides)
                state = cached_workspace_state(body.get('expected_revision'))
                if state is not None and not isinstance(body.get('data_destruction_needed'), bool):
                    return {'saved': True, 'created': 0, 'state': state}
    status = status_value(saved.get('issued'))
    report = Report(ROOT, date.today(), 'Inward details', publish_latest=False)
    report.add_saved_inward(dict(saved, source_id=plant, source_path=str(path)),
                           'already_issued' if status in {'generated', 'generated_manual', 'completed'} else 'ready',
                           'Inward details saved to Excel.')
    fresh = report.rows[0]

    def merge_followup(state):
        row = next((item for item in state.setdefault('rows', []) if item.get('source_kind') != 'manual'
                    and item.get('plant') == plant and str(item.get('sn')) == sn), None)
        if row is None:
            state['rows'].append(copy.deepcopy(fresh))
        else:
            row.update({key:fresh[key] for key in (*values, 'sheet', 'row', 'column_mapping',
                       'certificate_status', 'certificate_no', 'certificate_date', 'reminder_due', 'overdue',
                       'age_days', 'inward_date', 'inward_date_quality')})
        return state

    state = cached_workspace_state(body.get('expected_revision'))
    if state is not None:
        merge_followup(state)
    try:
        latest = latest_result()
        merge_followup(latest)
        save_json(RESULTS_PATH, latest)
    except (OSError, ValueError):
        pass
    remember_workbook_stamp()
    if state is None:
        state = merge_followup(post_commit_state(workspace_state, saved_message='Inward details saved to Excel'))
    state = save_requested_destruction(dict(body, source_sn=sn), state)
    return {'saved': True, 'created': 0, 'state': state}


def resolve_issued_excel_changes(body):
    """Restore issued source details or acknowledge the exact current Excel values."""
    action = body.get('action')
    if action not in {'restore', 'keep'}:
        raise ValueError('Choose whether to restore the issued details or keep the Excel changes.')
    plant, sn = selected_key(body.get('plant'), body.get('sn'))
    path = configured_paths(allow_unavailable=True)[int(plant[-1]) - 1]
    if not path.is_file():
        raise FileNotFoundError(f'Plant {plant[-1]} workbook could not be found. Check its path in Settings.')

    current = restored = receipt = None
    with run_lock(ROOT):
        recover_issuances(ROOT)
        with WorkbookSession(path, BACKUP_DIR) as session:
            overrides = {str(path.resolve()): session.data}
            require_revision(ROOT, body.get('expected_revision'), overrides)
            matches = [batch for batch in read_batches(BytesIO(session.data))
                       if text(batch.get('case_id')).strip().casefold() == sn.casefold()]
            if len(matches) != 1 or matches[0].get('duplicate_serial_number'):
                raise ValueError('This SN is missing or repeated. Check Excel, then refresh before resolving the change.')
            source = matches[0]
            if status_value(source.get('issued')) not in {'generated', 'generated_manual', 'completed'}:
                raise ValueError('This inward no longer has an issued certificate. Refresh and review its current status.')
            receipt = latest_issuance(ROOT, 'workbook', plant, source['case_id'])
            if not receipt or receipt.get('state') != 'complete' or not isinstance(receipt.get('after'), dict):
                raise ValueError('The issued certificate details could not be verified. Refresh the workbooks and try again.')
            conflict = workbook_receipt_conflict(source, receipt)
            if conflict:
                raise ValueError(conflict)

            issued = editor_details(receipt['after'])
            current = editor_details(source)
            changes = compare_details(current, issued)
            if not changes:
                raise ValueError('Excel now matches the issued certificate. Refresh the company list to update this notice.')
            if action == 'keep':
                keep_intentional(ROOT, receipt, plant, source['case_id'], current)
            else:
                current_normalized = normalized_source_details(current)
                issued_normalized = normalized_source_details(issued)
                restore_values = {'case_id': source['case_id']}
                for key in ISSUED_SOURCE_FIELDS:
                    if current_normalized[key] != issued_normalized[key]:
                        restore_values[key] = issued.get(key, '')
                if current_normalized['items'] != issued_normalized['items']:
                    restore_values['items'] = issued.get('items', [])

                original_status = source.get('issued')
                original_cod = source.get('cod_no')
                original_cod_date = parse_date(source.get('cod_date'))
                original_metadata = {key: source.get(key) for key in (
                    'external_pdf',
                    'pdf_recovery_pending', 'recycling_completed', 'hdd_cod_no', 'hdd_cod_date')}
                updated, _ = stage_certificate_details(session, restore_values,
                    inward_status=original_status)
                verified = [batch for batch in read_batches(BytesIO(updated))
                            if text(batch.get('case_id')).strip().casefold() == sn.casefold()]
                if len(verified) != 1:
                    raise ValueError('Restored details could not be verified. Excel was left unchanged.')
                restored = verified[0]
                if (compare_details(restored, issued)
                        or status_value(restored.get('issued')) != status_value(original_status)
                        or text(restored.get('cod_no')) != text(original_cod)
                        or parse_date(restored.get('cod_date')) != original_cod_date
                        or any(restored.get(key) != value for key, value in original_metadata.items())):
                    raise ValueError('The restore did not preserve the issued certificate and its status. Excel was left unchanged.')

                # If the workbook write fails, the comparison simply returns.
                clear_review(ROOT, receipt['id'])
                revision = workspace_revision(ROOT, overrides, force=True)
                session.before_persist = lambda _: require_revision(ROOT, revision, overrides)
                session._persist(updated)

    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
    message = 'Issued details restored to Excel' if action == 'restore' else 'Excel changes kept intentionally'
    state = post_commit_state(refresh, saved_message=message)
    return {'saved': True, 'action': action, 'changes': changes,
            'source_details': editor_details(restored if action == 'restore' else source),
            'state': state}


def validate_manual_identifiers(body):
    """Keep COD fixed while allowing a valid date correction after issuance."""
    if body.get('void_reissue'):
        raise ValueError('Certificate replacement with a new COD is no longer available. Corrections keep the current COD.')
    supplied = body.get('details')
    if supplied is None:
        supplied = {}
    if not isinstance(supplied, dict):
        raise ValueError('Enter the certificate details first.')
    cod = str(supplied.get('cod_no') or '').strip()
    cod_date = str(supplied.get('cod_date') or '').strip()
    record_id = body.get('record_id')
    expected_cod = expected_date = ''
    issued = False
    if record_id:
        record = next((item for item in read_register(ROOT).get('records', [])
                       if item.get('id') == record_id), None)
        if record:
            expected_cod = str(record.get('details', {}).get('cod_no') or '').strip()
            expected_date = str(record.get('details', {}).get('cod_date') or '').strip()
            issued = status_value(record.get('status')) in {'generated_manual', 'completed'}
            issued = issued or bool(latest_issuance(ROOT, 'manual', record.get('plant'), record_id))
    if cod and cod != expected_cod:
        raise ValueError('The COD is assigned automatically from the workbook and certificate PDF history. It cannot be entered or changed here.')
    if cod_date:
        parsed = parse_date(cod_date)
        expected = parse_date(expected_date)
        if not parsed:
            raise ValueError('Enter a valid certificate date.')
        if not issued and parsed != expected:
            raise ValueError('The certificate date is assigned automatically until the first PDF is generated.')


def post_commit_state(read_state, *, saved_message='Certificate saved successfully', fallback_state=None):
    try:
        return read_state()
    except Exception as exc:
        if fallback_state is not None:
            state = dict(fallback_state, refresh_pending=True, scan_complete=False)
        else:
            try:
                state = latest_result()
            except Exception:
                state = {'rows': [], 'scan_complete': False}
        return dict(state, setup_warning=f'{saved_message}, but the dashboard could not refresh: {exc}. Do not repeat the save; refresh after correcting the issue.')


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
            destination = safe_certificate_path(ROOT, relative_path, expected_sha=receipt.get('pdf_sha256') if receipt else None)
            if not valid_pdf(destination):
                continue
            relative_path = destination.relative_to((ROOT / 'Certificates').resolve()).as_posix()
        except (OSError, ValueError, TypeError):
            continue
        filename = destination.name
        issued_details = row.get('last_certificate_details')
        issued_type = issued_details.get('certificate_type') if isinstance(issued_details, dict) else None
        outputs.append({'filename': filename, 'relative_path': relative_path,
                        'certificate_no': cod, 'plant': row.get('plant'), 'supplier': row.get('supplier'),
                        'certificate_type': issued_type or row.get('certificate_type'),
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
    ASSETS = {"/": ("index.html", "text/html; charset=utf-8"),
              "/analytics.css": ("analytics.css", "text/css; charset=utf-8"),
              "/analytics.js": ("analytics.js", "text/javascript; charset=utf-8"),
              "/styles.css": ("styles.css", "text/css; charset=utf-8"),
              "/app.js": ("app.js", "text/javascript; charset=utf-8")}

    def log_message(self, _format, *_args):
        return

    def send_json(self, value, status=200):
        global WORKSPACE_STATE_CACHE
        if isinstance(value, dict) and status < 400:
            value = dict(value)
            # Reads can fail while Excel is open; do not misreport an already
            # committed certificate as a failure just to obtain a UI snapshot.
            if value.get('changed') is not False:
                try:
                    revision = value.get('workspace_revision') if value.get('requires_confirmation') else workspace_revision(ROOT)
                    value['workspace_revision'] = revision
                    state = value.get('state')
                    if isinstance(state, dict) and isinstance(state.get('rows'), list):
                        if (ROOT / 'Data' / 'Reports' / 'Reminder action error.txt').is_file():
                            warning = 'The last reminder action could not finish. Check Settings → Reminder options and try again.'
                            if warning not in str(state.get('setup_warning', '')):
                                state['setup_warning'] = ' '.join(filter(None, [state.get('setup_warning'), warning]))
                        WORKSPACE_STATE_CACHE = (revision, copy.deepcopy(state))
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
        self.send_json({"error": operation_error_message(exc)}, status)

    def do_GET(self):
        path = urlsplit(self.path).path
        if path == '/api/dashboard-review':
            self.send_json({'application':'ecoreco-certificates', 'changed':False})
            return
        if path in self.ASSETS:
            self.send_asset(*self.ASSETS[path])
            return
        with OPERATION_LOCK:
            self._do_get()

    def send_asset(self, name, mime):
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
                        **({'next_cod': settings['next_cod']} if settings.get('next_cod') is not None else {}),
                    },
                    "certificate_folder": str(ROOT / "Certificates"),
                    "server_time": datetime.now().isoformat(timespec="minutes"),
                })
            except Exception as exc:
                self.send_error_json(exc, 500)
            return

        if path == "/api/templates":
            try:
                from pdf_attachments import attachment_catalog
                self.send_json({"templates": template_specs(), 'annexures': attachment_catalog(ROOT), 'changed': False})
            except Exception as exc:
                self.send_error_json(exc, 500)
            return
        if path == '/api/history':
            try:
                self.send_json({'history': audit_records(ROOT)[-50:]})
            except (OSError, ValueError) as exc:
                self.send_error_json(exc)
            return
        if path == '/api/analytics':
            try:
                from certificate_analytics import build_analytics
                query = parse_qs(urlsplit(self.path).query)
                report = latest_result()
                rows = [row for row in report.get('rows', []) if row.get('source_kind') != 'manual'] + manual_rows(ROOT)
                from analytics_pdf import folder_events
                history = audit_records(ROOT)
                # Also read PDFs saved outside the app, including standalone HDDs.
                # Keep this endpoint read-only; it must never allocate a number.
                known = [row.get('certificate_no') for row in rows
                         if row.get('certificate_date') and row.get('certificate_status') in {'generated', 'generated_manual', 'completed'}]
                known.extend(record.get('cod') for record in history if record.get('state') == 'complete')
                saved = folder_events(ROOT, scan_inventory(ROOT), known)
                self.send_json(build_analytics(rows, history, period=query.get('period', ['month'])[0],
                                               plant=query.get('plant', ['all'])[0], scan_complete=report.get('scan_complete', False),
                                               saved_pdfs=saved))
            except (OSError, ValueError) as exc:
                self.send_error_json(exc)
            return
        if path == '/api/certificate-preview-image':
            query = parse_qs(urlsplit(self.path).query)
            try:
                payload = uploaded_pdf_bytes(query.get('path', [''])[0], query.get('sha256', [''])[0])
                import pypdfium2 as pdfium
                rendered = BytesIO()
                with OPERATION_LOCK, pdfium.PdfDocument(payload) as document:
                    _write_preview_image(document, 0, rendered)
                image = rendered.getvalue()
            except (ValueError, OSError, RuntimeError, IndexError):
                self.send_error(404)
                return
            self.send_response(200)
            self.send_header('Content-Type', 'image/png')
            self.send_header('Content-Length', str(len(image)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(image)
            return
        if path.startswith('/api/certificates/'):
            relative_path = unquote(path[len('/api/certificates/'):])
            digest = parse_qs(urlsplit(self.path).query).get('sha256', [None])[0]
            try:
                if digest:
                    payload = uploaded_pdf_bytes(relative_path, digest)
                else:
                    destination = safe_certificate_path(ROOT, relative_path)
                    if not valid_pdf(destination):
                        raise ValueError('Invalid PDF')
                    payload = destination.read_bytes()
            except (ValueError, OSError):
                self.send_error_json('This PDF is missing or unreadable. Refresh and recreate it.', 404)
                return
            self.send_response(200)
            self.send_header('Content-Type', 'application/pdf')
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(payload)
            return
        preview_page_match = re.fullmatch(r'/previews/([a-f0-9]{32})/page/([1-9][0-9]{0,3})\.png', path)
        if preview_page_match:
            try:
                destination = preview_page_image(preview_page_match[1], int(preview_page_match[2]))
                payload = destination.read_bytes()
            except ValueError:
                self.send_error(404)
                return
            except (OSError, RuntimeError):
                self.send_error_json('This page could not load. Preview the certificate again.', 500)
                return
            self.send_response(200)
            self.send_header('Content-Type', 'image/png')
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(payload)
            return
        preview_match = re.fullmatch(r"/previews/([a-f0-9]{32})\.(pdf|png)", path)
        if preview_match:
            from preview_store import touch_preview
            touch_preview(ROOT, preview_match[1])
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

        self.send_error(404)

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
            if parsed.path == '/api/certificates/upload-file':
                from external_certificate_upload import UPLOAD_LIMIT, stage_upload
                content_type = self.headers.get('Content-Type')
                if not 0 < length <= UPLOAD_LIMIT or content_type not in {'application/pdf', 'image/png', 'image/jpeg'}:
                    raise ValueError('Choose a PDF, PNG, or JPEG file under 50 MB.')
                payload = self.rfile.read(length)
                if len(payload) != length:
                    raise ValueError('The upload was interrupted. Choose the file again.')
                self.send_json(stage_upload(ROOT, payload, content_type=content_type))
                return
            if length < 0 or length > 1_000_000:
                raise ValueError("The request is too large.")
            body = json.loads(self.rfile.read(length) or b"{}")
            if not isinstance(body, dict):
                raise ValueError("Expected an object of request values.")

            if parsed.path == '/api/dashboard-presence':
                request = DASHBOARD_REVIEW.presence(body.get('client_id'), body.get('visible'), body.get('acknowledged'))
                from preview_store import preview_presence
                preview_presence(ROOT, body.get('client_id'), body.get('preview_session'))
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
                    **({'next_cod': settings['next_cod']} if settings.get('next_cod') is not None else {}),
                }})
            elif parsed.path == "/api/preview":
                self.send_json(dict(editor_preview(body), changed=False))
            elif parsed.path == '/api/certificate-type':
                kind = save_certificate_type_choice(body)
                self.send_json({'saved': True, 'certificate_type': kind, 'changed': False})
            elif parsed.path == '/api/clients/add':
                self.send_json(add_client(body))
            elif parsed.path == '/api/clients/update':
                self.send_json(save_inward_changes(body))
            elif parsed.path == '/api/clients/resolve-issued-changes':
                self.send_json(resolve_issued_excel_changes(body))
            elif parsed.path == '/api/clients/status':
                self.send_json(update_inward_status(body))
            elif parsed.path == '/api/certificates/link-external':
                self.send_json(link_external_certificate(body))
            elif parsed.path == '/api/certificates/upload-saved-file':
                from external_certificate_upload import stage_saved_upload
                self.send_json(stage_saved_upload(ROOT, body))
            elif parsed.path == '/api/certificates/save-upload':
                self.send_json(save_uploaded_certificate(body))
            elif parsed.path == '/api/clients/followup':
                self.send_json(save_inward_followup(body))
            elif parsed.path in {'/api/entries/delete-preview', '/api/entries/delete'}:
                from entry_deletion import delete_entries
                result = delete_entries(ROOT, body, commit=parsed.path == '/api/entries/delete')
                if result.get('deleted'):
                    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
                    result['state'] = post_commit_state(rebuild_history_state, saved_message='Mistaken entries deleted')
                self.send_json(result)
            elif parsed.path in {'/api/clients/remove-preview', '/api/clients/remove'}:
                from inward_removal import remove_inward
                result = remove_inward(ROOT, body, commit=parsed.path == '/api/clients/remove')
                if result.get('removed'):
                    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
                    result['state'] = post_commit_state(rebuild_history_state, saved_message='Inward removed from Excel')
                self.send_json(result)
            elif parsed.path == '/api/entries/undo-delete':
                if body.get('kind') == 'inward':
                    from inward_removal import undo_inward_removal
                    result = undo_inward_removal(ROOT, body.get('request_id'), body.get('expected_revision'))
                elif body.get('kind') == 'entries':
                    from entry_deletion import undo_entry_deletion
                    result = undo_entry_deletion(ROOT, body.get('request_id'), body.get('expected_revision'))
                else:
                    raise ValueError('Choose a saved deletion to undo.')
                result['state'] = post_commit_state(rebuild_history_state, saved_message='Deletion undone')
                self.send_json(result)
            elif parsed.path == '/api/destruction/request':
                expected = body.get('expected_revision')
                if not isinstance(expected, str) or not expected:
                    require_revision(ROOT, expected)
                from destruction_certificate import request_for_inward
                identifier = request_for_inward(ROOT, body.get('plant'), body.get('sn'), expected,
                    needed=body.get('needed', True), confirmed=body.get('confirmed') is True,
                    record_id=body.get('record_id'))
                workbook_changed = _sync_linked_recycling_completion(identifier) if identifier else False
                state = merge_manual_record_state(expected, identifier)
                if workbook_changed:
                    state = post_commit_state(lambda: refresh(force=True),
                                              saved_message='Data Destruction request saved', fallback_state=state)
                self.send_json({'record_id': identifier, 'state': state if state is not None else post_commit_state(
                    workspace_state, saved_message='Data destruction request saved')})
            elif parsed.path == "/api/manual/save":
                require_revision(ROOT, body.get('expected_revision'))
                validate_manual_identifiers(body)
                identifier = save_draft(ROOT, body.get("plant"), body.get("details"), body.get("record_id"),
                                        expected_revision=body.get('expected_revision'))
                self.send_json({"record_id": identifier, "state": post_commit_state(
                    workspace_state, saved_message='Draft saved successfully')})
            elif parsed.path == "/api/manual/generate":
                validate_manual_identifiers(body)
                response = reviewed_operation(ROOT, parsed.path, body, lambda values: generate_draft(
                    ROOT, values.get('plant'), values.get('details'), values.get('record_id'),
                    confirmed=bool(values.get('confirmed')), override_reason=values.get('override_reason', ''),
                    expected_revision=values.get('expected_revision'), starting_cod=values.get('starting_cod'),
                    review_only=values.get('review_only') is True))
                if response.get('created'):
                    workbook_changed = _sync_linked_recycling_completion(response.get('record_id'))
                    response['state'] = post_commit_state(lambda: (
                        refresh(force=True) if workbook_changed else
                        None if response.get('workbook_updated') else
                        merge_manual_record_state(body.get('expected_revision'), response.get('record_id'))
                    ) or workspace_state())
                    if response.get('recovery_note'):
                        response['state'].update(refresh_pending=True, recovery_note=response['recovery_note'])
                    response['state'] = save_requested_destruction(dict(body, record_id=response.get('record_id')), response['state'])
                if response.get('created'):
                    if response['state'].get('setup_warning'):
                        response.setdefault('warnings', []).append(response['state']['setup_warning'])
                    response['certificates'] = certificate_outputs([{'source_kind': 'manual',
                        'record_id': response['record_id'], 'plant': body.get('plant'),
                        'pdf_filename': response.get('filename'), 'certificate_no': response.get('certificate_no'),
                        'pdf_relative_path': response.get('relative_path'),
                        'certificate_type': (body.get('details') or {}).get('certificate_type'),
                        'supplier': (body.get('details') or {}).get('supplier')}])
                self.send_json(response)
            elif parsed.path == "/api/manual/complete":
                require_revision(ROOT, body.get('expected_revision'))
                undo = body.get('undo_completed') is True
                complete_draft(ROOT, body.get("record_id"), expected_revision=body.get('expected_revision'), undo=undo)
                workbook_changed = _sync_linked_recycling_completion(body.get('record_id'))
                state = merge_manual_record_state(body.get('expected_revision'), body.get('record_id'))
                if workbook_changed:
                    state = post_commit_state(lambda: refresh(force=True),
                                              saved_message='Data Destruction completion updated', fallback_state=state)
                if state is None:
                    state = post_commit_state(workspace_state, saved_message='Completion undone' if undo else 'Completed status saved successfully')
                self.send_json({"state": state})
            elif parsed.path == "/api/generate":
                if str(body.get('custom_cod') or '').strip():
                    raise ValueError('The COD is assigned automatically from the workbook and certificate PDF history. Clear the custom COD and refresh to continue.')
                selected = None
                if body.get("plant") or body.get("sn"):
                    selected = selected_key(body.get("plant"), body.get("sn"))
                plants = body.get("plants")
                self.send_json(reviewed_operation(ROOT, parsed.path, body, lambda values: generate(
                    selected=selected, plants=plants, selections=body.get('selections'),
                    confirmed=bool(values.get('confirmed')), expected_revision=values.get('expected_revision'),
                    override_reason=values.get('override_reason', ''), starting_cod=values.get('starting_cod'),
                    review_only=values.get('review_only') is True)))
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
                    picked = filedialog.askopenfilename(parent=picker_root, title="Choose pages to add to the certificate",
                                                       filetypes=[("PDF files", "*.pdf")])
                finally:
                    picker_root.destroy()
                attachment = None
                if picked:
                    # A workbook can change while the native picker is open.
                    from pdf_attachments import store_attachment
                    attachment = store_attachment(ROOT, picked, body.get('expected_revision'))
                self.send_json({"annexure": attachment})
            elif parsed.path == "/api/complete":
                self.send_json({"state": complete_workbook_record(body)})
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
                path, sheet, row = open_selected_workbook_record(body.get("plant"), body.get("sn"),
                                                                  body.get('sheet'), body.get('row'))
                response = {"ok": True, 'selected':True, 'row_selected':row is not None,
                            'workbook':str(path), 'sheet':sheet, 'row':row,
                            'selection': 'repair' if row is None else 'record'}
                if row is None:
                    response['warning'] = (
                        'Excel opened, but this row changed. Refresh the list and try again.'
                        if body.get('row') is not None else
                        'Excel opened for repair. This Serial Number appears more than once; choose the correct row in Excel.')
                self.send_json(response)
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
            elif parsed.path == "/api/create-plant-workbook":
                plant = body.get("plant")
                if plant not in {"plant1", "plant2"}:
                    raise ValueError("Choose Plant 1 or Plant 2.")
                plant_label = "Plant 1" if plant == "plant1" else "Plant 2"
                import tkinter as tk
                from tkinter import filedialog

                settings = load_settings(ROOT)
                configured_path = settings["workbook_paths"][int(plant[-1]) - 1]
                current = Path(configured_path).expanduser() if configured_path else None
                initial_dir = (current.parent if current and current.parent.is_dir()
                               else ROOT / "Plant Workbooks" if (ROOT / "Plant Workbooks").is_dir()
                               else Path.home() / "Documents")
                if not initial_dir.is_dir():
                    initial_dir = Path.home()
                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes("-topmost", True)
                    picked = filedialog.asksaveasfilename(
                        parent=picker_root,
                        title=f"Create a new {plant_label} workbook",
                        initialdir=str(initial_dir),
                        initialfile=f"{plant_label}.xlsx",
                        defaultextension=".xlsx",
                        filetypes=(("Excel workbook", "*.xlsx"),),
                    )
                finally:
                    picker_root.destroy()
                if not picked:
                    self.send_json({"cancelled": True})
                    return
                destination = create_plant_workbook(picked, plant)
                self.send_json({"path": str(destination), "plant": plant,
                                "message": "New workbook created. Click Save in Settings to connect it."})
            elif parsed.path == "/api/settings":
                with run_lock(ROOT):
                    settings = load_settings(ROOT)
                    paths = body.get("workbook_paths", settings["workbook_paths"])
                    if not isinstance(paths, list) or len(paths) != 2 or any(not isinstance(item, str) for item in paths):
                        raise ValueError('Choose a file or leave each plant disconnected.')
                    resolved = [Path(item.strip()).expanduser().resolve() if item.strip() else None for item in paths]
                    previous = [Path(item.strip()).expanduser().resolve() if item.strip() else None
                                for item in settings['workbook_paths']]
                    for current, original in zip(resolved, previous):
                        if current != original:
                            if current is None:
                                continue
                            if not current.is_file():
                                raise ValueError("One of those workbooks could not be found.")
                    if (all(resolved) and (resolved[0] == resolved[1]
                            or (resolved[0].is_file() and resolved[1].is_file() and resolved[0].samefile(resolved[1])))):
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
                    settings.update(workbook_paths=[str(path) if path is not None else '' for path in resolved],
                                    reminder_time=reminder_time, reminders_enabled=enabled)
                    if 'next_cod' in body:
                        from certificate_core import certificate_number
                        from certificate_drafts import numbering_sources
                        from run_certificates import numbering_plan, starting_cod_sequence
                        raw_cod = body['next_cod']
                        if raw_cod is None or raw_cod == '':
                            settings.pop('next_cod', None)
                        else:
                            next_cod = starting_cod_sequence(raw_cod)
                            if next_cod != settings.get('next_cod'):
                                batches, _ = numbering_sources(ROOT, settings)
                                plan = numbering_plan(batches, settings, ROOT / 'Certificates')
                                if next_cod < plan['minimum_sequence']:
                                    earliest = certificate_number(plan['minimum_sequence'], settings)
                                    raise ValueError(f'Choose COD {earliest} or higher. Earlier numbers are reserved.')
                                settings['next_cod'] = next_cod
                    save_json(ROOT / "settings.local.json", settings)
                    WORKBOOK_CHECK_STAMPS.pop(str(ROOT), None)
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
                                                   timeout=60, check=False,
                                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
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
                    **({'next_cod': settings['next_cod']} if settings.get('next_cod') is not None else {}),
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
                def refresh_restored_workspace():
                    ensure_certificate_folders(ROOT)
                    return refresh()

                state = post_commit_state(refresh_restored_workspace, saved_message='Backup restored successfully')
                response = {'restored': result, 'state': state,
                            'message': 'Backup restored. The previous data was preserved in a recovery snapshot.'}
                try:
                    response['settings'] = load_settings(ROOT)
                except (OSError, ValueError) as exc:
                    state.setdefault('setup_warning', f'Backup restored, but Settings could not be read: {exc}. Reopen Settings before changing them.')
                self.send_json(response)
            elif parsed.path == "/api/stop":
                STOPPING = True
                try:
                    self.send_json({"ok": True})
                finally:
                    if SERVER:
                        threading.Thread(target=SERVER.shutdown, daemon=True).start()
            else:
                self.send_error_json("Unknown action.", 404)
        except ClientAddRequestEnded as exc:
            self.send_json({'error': str(exc), 'client_request_ended': True}, 400)
        except (ValueError, OSError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
            self.send_error_json(exc)
        except Exception:
            report = ROOT / 'Data' / 'Reports' / 'Last Dashboard Error.txt'
            try:
                report.parent.mkdir(parents=True, exist_ok=True)
                report.write_text(
                    f'{datetime.now().isoformat()} POST {parsed.path}\n{traceback.format_exc()}',
                    encoding='utf-8',
                )
            except OSError:
                self.send_error_json('An unexpected error interrupted the operation. Refresh to check what was saved.', 500)
            else:
                self.send_error_json(
                    'An unexpected error interrupted the operation. Refresh to check what was saved. '
                    'Technical details are in Data / Reports / Last Dashboard Error.txt.', 500)

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

    def dashboard_payload(response):
        payload = response.read()
        if response.status != 200 or 'application/json' not in response.getheader('Content-Type', ''):
            return None
        try:
            value = json.loads(payload)
        except (ValueError, TypeError, UnicodeError):
            return None
        return value if isinstance(value, dict) else None

    for port in range(8765, 8776):
        connection = None
        try:
            connection = HTTPConnection("127.0.0.1", port, timeout=0.6)
            connection.request('GET', '/api/dashboard-review')
            payload = dashboard_payload(connection.getresponse())
            is_existing_dashboard = bool(payload and payload.get('application') == 'ecoreco-certificates')
            if not is_existing_dashboard:
                # Older running dashboards do not yet expose the presence API.
                connection.request('GET', '/api/state')
                payload = dashboard_payload(connection.getresponse())
                state = payload.get('state') if payload else None
                settings = payload.get('settings') if payload else None
                paths = settings.get('workbook_paths') if isinstance(settings, dict) else None
                is_existing_dashboard = (isinstance(state, dict) and isinstance(state.get('rows'), list)
                                         and isinstance(state.get('scan_complete'), bool)
                                         and isinstance(paths, list) and len(paths) == 2
                                         and all(isinstance(path, str) for path in paths))
            if is_existing_dashboard:
                if attention and reuse_dashboard(port):
                    return
                if open_browser:
                    webbrowser.open(f"http://127.0.0.1:{port}/" + ('?view=attention' if attention else ''))
                return
        except (OSError, HTTPException):
            pass
        finally:
            if connection is not None:
                connection.close()
        try:
            SERVER = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if SERVER is None:
        raise RuntimeError("Could not start the local dashboard. Close any old certificate dashboard and try again.")
    from certificate_drafts import migrate_local_ownership
    migrate_local_ownership(ROOT)
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
