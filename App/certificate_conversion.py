"""Prepare a local certificate and its HDD links before removing one Excel inward."""
from copy import deepcopy
from datetime import datetime
import json
import uuid


def prepare_local_certificate(root, plant, source, details, cod, issued_date, destination,
                              *, hdd_needed=None, warnings=()):
    from certificate_drafts import _manual_record_issued, clean_details, read_register, register_path
    from certificate_core import parse_date
    from issuance_safety import audit_records
    path = register_path(root)
    before = path.read_bytes() if path.is_file() else None
    data = read_register(root)
    identifier = uuid.uuid4().hex
    printed = clean_details(details)
    printed.update(case_id='', cod_no=str(cod), cod_date=issued_date.isoformat())
    if printed['certificate_type'] in {'sita', 'siemens'}:
        printed['form6'] = ''
    stamp = datetime.now().isoformat()
    record = {'id': identifier, 'plant': plant, 'details': printed, 'status': 'generated_manual',
              'pdf_filename': destination.name, 'pdf_relative_path': destination.relative_to(root / 'Certificates').as_posix(),
              'created_at': stamp, 'updated_at': stamp, 'warnings': list(warnings),
              'removed_inward': {'plant': plant, 'case_id': str(source['case_id']),
                                'source_sheet': source['source_sheet'], 'source_row': source['source_row']}}
    data['records'].append(record)
    old_link = {'plant': plant, 'case_id': str(source['case_id'])}
    new_link = dict(old_link, record_id=identifier)
    linked = [entry for entry in data['records'] if entry['details'].get('destruction_source') == old_link]
    for entry in linked:
        if hdd_needed is False and entry.get('destruction_requested', True) and _manual_record_issued(root, entry):
            raise ValueError('This Data Destruction certificate has been issued. Keep its request when saving this inward.')
        entry['details']['destruction_source'] = deepcopy(new_link)
        if isinstance(hdd_needed, bool):
            entry['destruction_requested'] = hdd_needed
    if hdd_needed is True and not linked:
        received = parse_date(printed.get('inward_date'))
        hdd = clean_details(dict(printed, certificate_type='destruction', cod_no='', cod_date='',
            reference='', items=[], destruction_source=new_link, certificate_fields={
                'device_receipts': json.dumps([{'date': received.isoformat() if received else '', 'count': ''}])}))
        data['records'].append({'id': uuid.uuid4().hex, 'plant': plant, 'details': hdd,
            'status': 'ready', 'created_at': stamp, 'updated_at': stamp, 'destruction_requested': True})
    receipt_ids = [entry['id'] for entry in audit_records(root) if entry.get('state') == 'complete'
        and entry.get('source_kind') == 'workbook' and entry.get('plant') == plant
        and str(entry.get('key')).strip().casefold() == str(source['case_id']).strip().casefold()]
    from manual_ownership import normalize_records
    normalize_records(data)
    return record, (path, before, json.dumps(data, indent=2, ensure_ascii=False, default=str).encode('utf-8')), {
        'plant': plant, 'key': str(source['case_id']), 'receipt_ids': receipt_ids}
