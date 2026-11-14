"""Keep the operator's original PDF and record its certificate details."""
from contextlib import ExitStack
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import json
import re
import uuid

from certificate_core import number, parse_date, status_value
from certificate_drafts import clean_details, read_register, register_path, write_register
from certificate_inventory import (CERTIFICATE_FOLDERS, cod_conflicts, identity, is_external_path,
                                   safe_certificate_path, scan_inventory, text_identifiers)
from certificate_templates import TEMPLATE_NAMES
from identifier_rules import workspace_batches
from issuance_safety import IssuanceTransaction, digest_bytes, file_digest, latest_issuance, recover_issuances
from pdf_attachments import checked_reader
from preview_store import cleanup_previews
from run_certificates import _safe_cod_name, certificate_filename, numbering_plan, workbook_cod_ownership_issue
from workbook_store import WorkbookSession
from workflow_support import load_settings, operation_error_message, run_lock
from workspace_review import require_revision


UPLOAD_LIMIT = 50_000_000


def _valid_cod(cod):
    return bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9 ._/-]{0,79}', cod)
                and any(ch.isdigit() for ch in cod))


def _inspect(root, payload):
    reader = checked_reader(payload)
    prefix = load_settings(root).get('certificate_number_prefix', '')
    cods, references = {}, set()
    # The front page identifies the certificate. Subsequent pages can be
    # annexures with identifiers from other documents.
    for mode in ('plain', 'layout'):
        found, refs = text_identifiers(reader.pages[0].extract_text(extraction_mode=mode) or '', number_prefix=prefix)
        for cod in found:
            if _valid_cod(cod):
                cods.setdefault(identity(cod), cod)
        references.update(refs)
    return list(cods.values()), sorted(references)


def stage_upload(root, payload):
    if not 0 < len(payload) <= UPLOAD_LIMIT:
        raise ValueError('Choose a PDF smaller than 50 MB.')
    cods, references = _inspect(root, payload)
    cleanup_previews(root)
    token = uuid.uuid4().hex
    folder = Path(root) / 'Data' / 'Previews'
    folder.mkdir(parents=True, exist_ok=True)
    with (folder / f'{token}.pdf').open('xb') as stream:
        stream.write(payload)
    return {'token':token, 'cod':cods[0] if len(cods) == 1 else '', 'changed':False,
            'ambiguous_cod':len(cods) > 1,
            'reference':references[0] if len(references) == 1 else ''}


def save_upload(root, backup_dir, body):
    root = Path(root).resolve()
    token = body.get('token')
    if not isinstance(token, str) or not re.fullmatch(r'[a-f0-9]{32}', token):
        raise ValueError('Choose the certificate PDF again.')
    stage = root / 'Data' / 'Previews' / f'{token}.pdf'
    if not stage.is_file():
        raise ValueError('The upload has expired. Choose the certificate PDF again.')
    submitted = body.get('details')
    kind = submitted.get('certificate_type') if isinstance(submitted, dict) else None
    if kind not in TEMPLATE_NAMES:
        raise ValueError('Choose the certificate type.')
    details = clean_details(submitted)
    reviewed_cod = details['cod_no']
    automatic = body.get('automatic_cod') is True or not reviewed_cod
    issued_date = parse_date(details['cod_date'])
    if not issued_date or issued_date > date.today():
        raise ValueError('Enter the certificate issue date. It cannot be in the future.')
    if details['inward_date'] and not parse_date(details['inward_date']):
        raise ValueError('Enter a valid received date.')
    if details['certificate_weight'] and (number(details['certificate_weight']) or 0) <= 0:
        raise ValueError('Enter a positive certificate weight in kg.')
    details['cod_date'] = issued_date.isoformat()
    revision = body.get('expected_revision')
    target = body.get('inward')
    if target is not None and (not isinstance(target, dict) or target.get('plant') not in {'plant1','plant2'}
                               or not isinstance(target.get('sn'), str) or not target['sn'].strip()):
        raise ValueError('Choose an Excel inward.')
    with run_lock(root), ExitStack() as stack:
        recover_issuances(root)
        require_revision(root, revision)
        payload = stage.read_bytes()
        if not 0 < len(payload) <= UPLOAD_LIMIT:
            raise ValueError('Choose a PDF smaller than 50 MB.')
        printed, _ = _inspect(root, payload)
        digest = digest_bytes(payload)
        session = None
        source = None
        overrides = {}
        plant = None
        register = read_register(root)
        if target:
            plant = target['plant']
            path = load_settings(root)['workbook_paths'][int(plant[-1]) - 1]
            if not path:
                raise ValueError('Connect the selected plant workbook in Settings.')
            session = stack.enter_context(WorkbookSession(Path(path).resolve(), backup_dir))
            overrides[str(session.path.resolve())] = session.data
        batches = workspace_batches(root, overrides)
        if target:
            matches = [row for row in batches if row['source_id'] == plant
                       and row['case_id'].strip().casefold() == target['sn'].strip().casefold()]
            if len(matches) != 1 or matches[0].get('duplicate_serial_number'):
                raise ValueError('This inward is missing or its SN is repeated. Check Excel and refresh.')
            source = matches[0]
            for field in ('case_id', 'supplier', 'location', 'reference', 'form6'):
                details[field] = source.get(field, '')
            details['inward_date'] = source['inward_date'].isoformat() if source['inward_date'] else ''
        if not details['supplier']:
            raise ValueError('Enter the company name.')

        is_hdd = bool(source and kind == 'destruction')
        workbook_record = bool(source and not is_hdd)
        existing_cod = source.get('cod_no') if workbook_record else ''
        existing = None
        if is_hdd:
            if status_value(source['issued']) == 'na':
                raise ValueError('This inward is marked N/A. Choose Undo N/A before attaching an HDD certificate.')
            details['destruction_source'] = link = {'plant':plant, 'case_id':source['case_id']}
            linked = [record for record in register['records'] if record['details'].get('destruction_source') == link]
            if len(linked) > 1 or source.get('hdd_cod_no'):
                raise ValueError('This inward already has a Data Destruction certificate. Open its record to review it.')
            existing = linked[0] if linked else None
            if existing and (existing['details'].get('cod_no') or status_value(existing.get('status')) == 'na'):
                raise ValueError('Review the existing Data Destruction record before uploading a certificate.')
            if existing:
                details['certificate_fields'] = existing['details'].get('certificate_fields', {})
        if workbook_record:
            if (status_value(source['issued']) not in {'blank', 'alert', 'ready', 'generated_manual'}
                    or source.get('external_pdf') or source.get('pdf_recovery_pending')
                    or latest_issuance(root, 'workbook', plant, source['case_id'])):
                raise ValueError('This inward already has a certificate or is marked N/A. Choose an inward available for attachment.')
        inventory = None
        cod = reviewed_cod
        if automatic:
            if len(printed) == 1:
                cod = printed[0]
            elif len(printed) > 1:
                raise ValueError('More than one certificate number was found on the first page. Check the PDF before saving.')
            elif existing_cod:
                cod = existing_cod
            else:
                inventory = scan_inventory(root, force=True)
                cod = numbering_plan(batches, load_settings(root), root / 'Certificates', inventory=inventory)['next_cod']
            if reviewed_cod and identity(reviewed_cod) != identity(cod):
                raise ValueError(f'The files changed and the suggested COD is now {cod}. Refresh and review the upload before saving.')
        if not _valid_cod(cod):
            raise ValueError('Enter the certificate number exactly as it appears on the PDF.')
        _safe_cod_name(cod)
        if printed and {identity(value) for value in printed} != {identity(cod)}:
            raise ValueError('The certificate number must match the number printed on the PDF. Check the file and try again.')
        if existing_cod and identity(existing_cod) != identity(cod):
            raise ValueError(f'This inward has COD {existing_cod}. Choose the corresponding PDF.')
        details['cod_no'] = cod
        conflict = workbook_cod_ownership_issue(
            {'source_id':plant if workbook_record else None, 'case_id':source['case_id'] if workbook_record else None,
             'cod_no':cod}, batches, register['records'])
        if conflict:
            raise ValueError(conflict)
        if any(identity(row.get('hdd_cod_no')) == identity(cod) for row in batches):
            raise ValueError(f'COD {cod} is already assigned to a Data Destruction certificate.')
        if inventory is None:
            inventory = scan_inventory(root, force=True)
        duplicates = [entry for entry in inventory['entries'] if entry.get('present') and entry.get('sha256') == digest]
        if any(entry.get('managed_owner') or not is_external_path(entry['path']) for entry in duplicates) or any(
                isinstance(row.get('external_pdf'), dict) and row['external_pdf'].get('sha256') == digest for row in batches):
            raise ValueError('This PDF already belongs to a saved certificate. Open it in Saved certificates.')
        allowed = [entry['path'] for entry in duplicates]
        if cod_conflicts(inventory, cod, allowed_paths=allowed):
            raise ValueError(f'COD {cod} is already used by another PDF. Check the certificate number.')
        require_revision(root, revision, overrides)
        relative = (allowed[0].removeprefix('Certificates/') if allowed else
                    f"{CERTIFICATE_FOLDERS['external']}/{certificate_filename(details, cod=cod)}")
        destination = safe_certificate_path(root, relative)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if not allowed and destination.exists():
            raise ValueError('A PDF with that filename already exists. Refresh and check Saved certificates.')
        relative = destination.relative_to(root / 'Certificates').as_posix()
        source_path = session.path if session else register_path(root)
        source_before = session.data if session else source_path.read_bytes() if source_path.is_file() else None
        identifier = source['case_id'] if workbook_record else existing['id'] if existing else uuid.uuid4().hex
        old_record = dict(existing or {})
        if not workbook_record:
            record = existing if existing else {'id':identifier, 'created_at':datetime.now().isoformat()}
            record.update(plant=plant, details=details, status='generated_manual', pdf_filename=destination.name,
                          pdf_relative_path=relative, destruction_requested=True,
                          pdf_recovery_pending=False, updated_at=datetime.now().isoformat())
            if not existing:
                register['records'].append(record)
        companion = None
        if is_hdd:
            from hdd_workbook import stage_hdd_tracking
            workbook_data = stage_hdd_tracking(session, source['case_id'], cod, issued_date)
            custom_path = register_path(root)
            companion = (custom_path, custom_path.read_bytes() if custom_path.is_file() else None,
                         json.dumps(register, indent=2, ensure_ascii=False, default=str).encode('utf-8'))
        transaction = IssuanceTransaction(root, destination, stage, source_path, source_before,
            source_kind='workbook' if workbook_record else 'manual', plant=plant, key=identifier,
            cod=cod, issued_date=issued_date, before=source if workbook_record else old_record,
            after=details, companion_source=companion, external_upload=True, retain_existing=bool(allowed))
        try:
            transaction.publish()
            if session:
                session.before_persist = transaction.expect_source
                if is_hdd:
                    transaction.sync_companion(True)
                    session._persist(workbook_data)
                else:
                    external = {'cod':cod, 'date':issued_date, 'link':{'path':relative, 'sha256':digest},
                                'certificate_type':kind}
                    session.set_issued_status(source['case_id'], source['source_sheet'], source['source_row'],
                                              source['cod_no'], source['issued'], 'Generated Externally', external=external)
            else:
                transaction.expect_source(json.dumps(register, indent=2, ensure_ascii=False, default=str).encode('utf-8'))
                write_register(root, register)
        except Exception:
            actual = digest_bytes(session.data) if session else file_digest(source_path)
            if actual == transaction.data['source_before']:
                transaction.rollback()
            raise
        note = ''
        try:
            transaction.complete()
        except (OSError, ValueError) as exc:
            note = operation_error_message(exc)
        try:
            stage.unlink(missing_ok=True)
        except OSError:
            pass
    return {'saved':True, 'certificate_no':cod, 'relative_path':relative, 'recovery_note':note,
            'record_id':None if workbook_record else identifier}
