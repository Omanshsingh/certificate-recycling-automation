"""Keep the operator's original PDF and record its certificate details."""
from contextlib import ExitStack
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import json
import re
import uuid

from certificate_core import batch_content, number, parse_date, status_value
from certificate_drafts import clean_details, current_record_pdf, read_register, register_path, write_register
from certificate_inventory import (cod_conflicts, filename_cod, identity, issuance_paths,
                                   safe_certificate_path, scan_inventory, text_identifiers)
from certificate_templates import TEMPLATE_NAMES, template_id
from identifier_rules import workspace_batches
from inward_followup import FIELDS
from issuance_safety import (IssuanceTransaction, audit_records, digest_bytes, file_digest, latest_issuance,
                             recover_issuances)
from pdf_attachments import checked_reader
from preview_store import cleanup_previews
from run_certificates import (_safe_cod_name, certificate_filename, find_certificate,
                              workbook_cod_ownership_issue)
from workbook_store import WorkbookSession
from workflow_support import load_settings, operation_error_message, run_lock
from workspace_review import require_revision


UPLOAD_LIMIT = 50_000_000
EDITABLE_CERTIFICATE_FIELDS = {'supplier', 'location', 'inward_date', 'reference', 'form6', 'certificate_weight'}


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


def _image_pdf(payload):
    from PIL import Image, ImageOps, UnidentifiedImageError
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen.canvas import Canvas

    try:
        with Image.open(BytesIO(payload)) as source:
            if (source.format not in {'PNG', 'JPEG'} or getattr(source, 'n_frames', 1) != 1
                    or source.width * source.height > 25_000_000):
                raise ValueError('Choose a single PNG or JPEG image under 25 megapixels.')
            image = ImageOps.exif_transpose(source)
            ratio = min(1, 842 / max(image.size))
            width, height = (dimension * ratio for dimension in image.size)
            output = BytesIO()
            page = Canvas(output, pagesize=(width, height), invariant=1)
            page.drawImage(ImageReader(image.convert('RGBA')), 0, 0, width, height, mask='auto')
            page.save()
            return output.getvalue()
    except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as exc:
        raise ValueError('This image could not be opened. Choose a readable PNG or JPEG.') from exc


def stage_upload(root, payload, *, content_type='application/pdf'):
    if not 0 < len(payload) <= UPLOAD_LIMIT:
        raise ValueError('Choose a certificate file under 50 MB.')
    if content_type in {'image/png', 'image/jpeg'}:
        payload = _image_pdf(payload)
    if len(payload) > UPLOAD_LIMIT:
        raise ValueError('This image is too large to save as a PDF. Choose a smaller image.')
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


def stage_saved_upload(root, body):
    digest = body.get('sha256')
    if not isinstance(digest, str) or not re.fullmatch(r'[a-f0-9]{64}', digest):
        raise ValueError('Refresh Saved certificates and choose the current file.')
    path = safe_certificate_path(root, body.get('relative_path'), expected_sha=digest)
    relative = path.relative_to((Path(root) / 'Certificates').resolve()).as_posix()
    entry = next((item for item in scan_inventory(root, force=True)['entries']
                  if item['path'] == f'Certificates/{relative}' and item.get('present')), None)
    if not entry or not entry.get('linkable_external') or entry.get('sha256') != digest:
        raise ValueError('This PDF changed or already belongs to a saved certificate. Refresh and choose the current file.')
    if not path.is_file() or path.stat().st_size > UPLOAD_LIMIT:
        raise ValueError('Choose an available certificate PDF under 50 MB.')
    payload = path.read_bytes()
    if digest_bytes(payload) != digest:
        raise ValueError('This PDF changed. Refresh Saved certificates and try again.')
    staged = stage_upload(root, payload)
    cod = filename_cod(path.name)
    if not staged['cod'] and not staged['ambiguous_cod'] and cod and _valid_cod(cod):
        staged['filename_cod'] = cod
    return staged


def _replacement_paths(root, inventory, current, cod, source_kind, plant, key):
    # Retire only PDFs with verified ownership, plus the selected current PDF
    # for older records created before issuance receipts were introduced.
    current_path = current.relative_to(root).as_posix()
    current_sha = file_digest(current)
    owned = {path.casefold() for path in issuance_paths(inventory, cod, source_kind, plant, key)}
    present = {entry['path']: entry['sha256'] for entry in inventory['entries']
               if entry['present'] and entry['path'].casefold() in owned}
    entry = next((row for row in inventory['entries'] if row['path'] == current_path and row['present']), None)
    if (not entry or not current_sha or entry.get('sha256') != current_sha
            or (entry.get('managed_owner') and current_path.casefold() not in owned)):
        raise ValueError('The current PDF could not be verified for this certificate. Refresh and check Saved certificates.')
    present[current_path] = current_sha
    # A retired legacy PDF may still reserve this COD in the inventory.
    retired = {(old['path'], old['sha256']) for receipt in audit_records(root)
               if receipt.get('state') == 'complete' and receipt.get('source_kind') == source_kind
               and receipt.get('plant') == plant and str(receipt.get('key')).casefold() == str(key).casefold()
               and identity(receipt.get('cod')) == identity(cod)
               for old in receipt.get('previous_pdfs', [])}
    allowed = [entry['path'] for entry in inventory['entries'] if entry['path'] in present
               or (not entry['present'] and (entry['path'].casefold() in owned
                   or (entry['path'], entry.get('sha256')) in retired))]
    paths = {safe_certificate_path(root, path.removeprefix('Certificates/'), expected_sha=sha): sha
             for path, sha in present.items()}
    return allowed, paths


def save_upload(root, backup_dir, body):
    root = Path(root).resolve()
    if body.get('certificate_edits'):
        raise ValueError('Client details are read-only during upload. Use Review & edit to change them.')
    token = body.get('token')
    if not isinstance(token, str) or not re.fullmatch(r'[a-f0-9]{32}', token):
        raise ValueError('Choose the certificate PDF again.')
    stage = root / 'Data' / 'Previews' / f'{token}.pdf'
    if not stage.is_file():
        raise ValueError('The upload has expired. Choose the certificate PDF again.')
    replacement = body.get('replacement')
    replacing = replacement is not None
    target = body.get('inward')
    record_id = None
    if replacing:
        if not isinstance(replacement, dict) or target is not None:
            raise ValueError('Choose an existing certificate to replace.')
        if 'record_id' in replacement:
            record_id = replacement['record_id']
            if not isinstance(record_id, str) or not re.fullmatch(r'[a-f0-9]{32}', record_id):
                raise ValueError('Choose an existing certificate to replace.')
        else:
            target = replacement
    submitted = body.get('details')
    if replacing and isinstance(submitted, dict):
        submitted = {'certificate_type':'standard', 'cod_no':submitted.get('cod_no'),
                     'cod_date':submitted.get('cod_date')}
    kind = submitted.get('certificate_type') if isinstance(submitted, dict) else None
    if not isinstance(kind, str) or kind not in TEMPLATE_NAMES:
        raise ValueError('Choose the certificate type.')
    details = clean_details(submitted)
    reviewed_cod = details['cod_no']
    issued_date = parse_date(details['cod_date'])
    if not issued_date or issued_date > date.today():
        raise ValueError('Enter the certificate issue date. It cannot be in the future.')
    if details['inward_date'] and not parse_date(details['inward_date']):
        raise ValueError('Enter a valid received date.')
    if details['inward_date']:
        details['inward_date'] = parse_date(details['inward_date']).isoformat()
    if details['certificate_weight'] and (number(details['certificate_weight']) or 0) <= 0:
        raise ValueError('Enter a positive certificate weight in kg.')
    details['cod_date'] = issued_date.isoformat()
    revision = body.get('expected_revision')
    local_parent_id = target.get('record_id') if isinstance(target, dict) and not replacing else None
    if local_parent_id is not None:
        if (kind != 'destruction' or not isinstance(local_parent_id, str)
                or not re.fullmatch(r'[a-f0-9]{32}', local_parent_id)):
            raise ValueError('Choose a Recycling record to attach its Data Destruction certificate.')
        target = None
    elif target is not None and (not isinstance(target, dict) or target.get('plant') not in ('plant1', 'plant2')
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
        existing = None
        if record_id:
            existing = next((record for record in register['records'] if record['id'] == record_id), None)
            if existing is None:
                raise ValueError('This certificate could not be found. Refresh and choose the current record.')
            plant = existing['plant']
            kind = existing['details']['certificate_type']
            link = existing['details'].get('destruction_source')
            if link and not link.get('record_id'):
                target = {'plant':link['plant'], 'sn':link['case_id']}
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
            if not replacing:
                fields = dict(source.get('certificate_fields') or {}) if kind == source.get('certificate_type') else {}
                if submitted.get('certificate_fields'):
                    fields.update(details['certificate_fields'])
                weight = details['certificate_weight']
                details.update(batch_content(source), **{key: source.get(key, '') for key in FIELDS})
                received = source['inward_date']
                details.update(certificate_type=kind, certificate_fields=fields,
                               inward_date=received.isoformat() if received else '',
                               certificate_weight=weight or source.get('certificate_weight') or '')
                if kind == 'destruction':
                    details['items'] = []
        is_hdd = bool(source and kind == 'destruction')
        workbook_record = bool(source and not is_hdd)
        source_kind = 'workbook' if workbook_record else 'manual'
        existing_cod = source.get('cod_no') if workbook_record else ''
        inventory = None
        previous_paths = {}
        permitted_paths = []
        if replacing:
            previous = source if workbook_record else existing
            if not previous or status_value(previous.get('issued' if workbook_record else 'status')) not in {
                    'generated', 'generated_manual', 'completed'}:
                raise ValueError('Choose an issued certificate with a current PDF.')
            previous_details = source if workbook_record else existing['details']
            existing_cod = previous_details['cod_no']
            if not existing_cod:
                raise ValueError('The current certificate needs a COD before its PDF can be replaced.')
            if is_hdd and identity(source.get('hdd_cod_no')) != identity(existing_cod):
                raise ValueError('The HDD certificate and Excel inward have different CODs. Check the linked entries.')
            key = source['case_id'] if workbook_record else existing['id']
            receipt = latest_issuance(root, source_kind, plant, key)
            if receipt and identity(receipt.get('cod')) != identity(existing_cod):
                raise ValueError('The current COD differs from the issuance history. Check the selected record.')
            inventory = scan_inventory(root, force=True)
            current = (find_certificate(root / 'Certificates', source, inventory=inventory, receipt=receipt)
                       if workbook_record else current_record_pdf(root, existing, receipt=receipt))
            if not current:
                raise ValueError('The current PDF is missing or changed. Refresh and check Saved certificates.')
            permitted_paths, previous_paths = _replacement_paths(
                root, inventory, current, existing_cod, source_kind, plant, key)
            saved_details = receipt.get('after') if receipt and workbook_record else None
            details = dict(previous_details, **(saved_details or {}))
            kind = template_id(details.get('certificate_type'))
            details.update(certificate_type=kind, cod_date=issued_date.isoformat())

        if (is_hdd or local_parent_id) and not replacing:
            if local_parent_id:
                parent = next((record for record in register['records'] if record['id'] == local_parent_id), None)
                if not parent or parent['details']['certificate_type'] == 'destruction':
                    raise ValueError('The linked certificate could not be found. Refresh and choose it again.')
                parent_details = parent['details']
                case_id = (parent.get('removed_inward') or {}).get('case_id') or parent_details.get('case_id') or local_parent_id
                link = {'plant':None, 'case_id':case_id, 'record_id':local_parent_id}
                details.update({key:parent_details.get(key, '') for key in EDITABLE_CERTIFICATE_FIELDS},
                               case_id=case_id, items=[])
            else:
                if source.get('certificate_type') == 'destruction' and source.get('cod_no'):
                    raise ValueError('This inward already has a Data Destruction certificate. Open its record to review it.')
                link = {'plant':plant, 'case_id':source['case_id']}
            details['destruction_source'] = link
            linked = [record for record in register['records']
                      if (candidate := record['details'].get('destruction_source'))
                      and (candidate.get('record_id') == local_parent_id if local_parent_id else candidate == link)]
            if len(linked) > 1 or (source and source.get('hdd_cod_no')):
                raise ValueError('This inward already has a Data Destruction certificate. Open its record to review it.')
            existing = linked[0] if linked else None
            if existing and (existing['details'].get('cod_no') or status_value(existing.get('status')) == 'na'):
                raise ValueError('Review the existing Data Destruction record before uploading a certificate.')
            if existing:
                details.update(existing['details'], cod_date=issued_date.isoformat())
                details['case_id'] = link['case_id']
        if not details['supplier']:
            raise ValueError('Enter the company name.')
        if workbook_record and not replacing:
            if (status_value(source['issued']) not in {'blank', 'alert', 'ready', 'generated_manual'}
                    or source.get('external_pdf') or source.get('pdf_recovery_pending')
                    or latest_issuance(root, 'workbook', plant, source['case_id'])):
                raise ValueError('This client already has a certificate. Select its current certificate in the client list to replace the PDF.')
        if replacing:
            if reviewed_cod and identity(reviewed_cod) != identity(existing_cod):
                raise ValueError(f'The replacement must keep COD {existing_cod}.')
            printed_numbers = {identity(value) for value in printed}
            if printed_numbers and identity(existing_cod) not in printed_numbers:
                raise ValueError(
                    f'The replacement PDF must match the number printed for existing COD {existing_cod}. '
                    'Choose the correct PDF and try again.'
                )
            cod = existing_cod
        else:
            printed_numbers = {identity(value) for value in printed}
            cod = reviewed_cod or existing_cod or (printed[0] if len(printed) == 1 else '')
            if not cod:
                if len(printed) > 1:
                    raise ValueError('The PDF text has more than one possible COD. Enter the number printed beside "Certificate No.".')
                raise ValueError('The app could not read a COD. Enter the number printed on this certificate.')
            if printed_numbers and identity(cod) not in printed_numbers:
                raise ValueError('The COD must match a number printed on the PDF. Check the file and try again.')
        if not _valid_cod(cod):
            raise ValueError('Enter the certificate number exactly as it appears on the PDF.')
        _safe_cod_name(cod)
        if existing_cod and identity(existing_cod) != identity(cod):
            raise ValueError(f'This inward has COD {existing_cod}. Choose the corresponding PDF.')
        details['cod_no'] = cod
        conflict = workbook_cod_ownership_issue(
            {'source_id':plant if workbook_record else None, 'case_id':source['case_id'] if workbook_record else None,
             'cod_no':cod}, batches, [record for record in register['records'] if record is not existing])
        if conflict:
            raise ValueError(conflict)
        if any(identity(row.get('hdd_cod_no')) == identity(cod) for row in batches
               if not (replacing and is_hdd and row is source)):
            raise ValueError(f'COD {cod} is already assigned to a Data Destruction certificate.')
        if inventory is None:
            inventory = scan_inventory(root, force=True)
        duplicates = [entry for entry in inventory['entries'] if entry.get('present') and entry.get('sha256') == digest]
        if any(not entry.get('linkable_external') for entry in duplicates) or any(
                isinstance(row.get('external_pdf'), dict) and row['external_pdf'].get('sha256') == digest for row in batches):
            raise ValueError('This PDF already belongs to a saved certificate. Open it in Saved certificates.')
        allowed = [entry['path'] for entry in duplicates]
        if cod_conflicts(inventory, cod, allowed_paths=allowed + permitted_paths):
            raise ValueError(f'COD {cod} is already used by another PDF. Check the certificate number.')
        require_revision(root, revision, overrides)
        relative = allowed[0].removeprefix('Certificates/') if allowed else certificate_filename(details, cod=cod)
        destination = safe_certificate_path(root, relative)
        if destination in previous_paths:
            stem = destination.stem.removesuffix('_HDD') if kind == 'destruction' else destination.stem
            suffix = '_HDD.pdf' if kind == 'destruction' else '.pdf'
            destination = destination.with_name(f'{stem}_{uuid.uuid4().hex[:8]}{suffix}')
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
            source_kind=source_kind, plant=plant, key=identifier,
            cod=cod, issued_date=issued_date, before=source if workbook_record else old_record,
            after=details, old_paths=previous_paths, companion_source=companion, external_upload=True,
            retain_existing=bool(allowed))
        try:
            if replacing and {root / item['path']: item['sha256'] for item in transaction.data['previous_pdfs']} != previous_paths:
                raise ValueError('The current PDF changed during the upload. Refresh and review it again.')
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
    return {'saved':True, 'replaced':replacing, 'certificate_no':cod, 'relative_path':relative, 'recovery_note':note,
            'record_id':None if workbook_record else identifier}
