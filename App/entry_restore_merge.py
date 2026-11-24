"""Restore deleted records into current files using verified recovery copies."""
from contextlib import ExitStack
from copy import deepcopy
import hmac
from io import BytesIO
import json
from pathlib import Path
import re
from types import SimpleNamespace
from uuid import uuid4

from certificate_core import batch_content, norm, read_batches, text
from certificate_drafts import register_path
from client_intake import stage_new_client
from entry_deletion import _paths, _sync, require_deletion_within_restore_window
from identifier_rules import identifier, require_restored_identifiers, workspace_batches
from inward_serials import reserved_workbook_serials
from issuance_safety import _write_bytes, digest_bytes, recover_issuances
from workbook_store import WorkbookSession, WorkbookSnapshot
from workflow_support import load_settings, run_lock, save_json
from workspace_review import require_revision, review_token, workspace_revision


def source_journal(root, target):
    if target.get('kind') not in {'inward', 'entries'} or not re.fullmatch(r'[a-f0-9]{32}', str(target.get('request_id', ''))):
        raise ValueError('The saved deletion identifier is invalid.')
    folder = 'Inward Removals' if target['kind'] == 'inward' else 'Entry Deletions'
    return root / 'Data' / folder / (target['request_id'] + '.json')


def _verified_copy(path, digest):
    data = path.read_bytes()
    if digest_bytes(data) != digest:
        raise ValueError('A saved recovery copy failed verification.')
    return data


def _original_sources(root, target, journal):
    if target['kind'] == 'entries':
        _paths(root, journal)
        if journal.get('pdfs') or journal.get('receipt_ids'):
            raise ValueError('This deletion includes certificate history. Restore it separately from an unchanged recovery copy.')
        folder = root / 'Data' / 'Entry Deletion Recovery' / target['request_id']
        return {Path(item['path']).resolve():_verified_copy(folder/item['before_snapshot'], item['before'])
                for item in journal['sources']}
    path = Path(journal['path']).resolve()
    folder = (root / 'Data' / 'Workbook Backups').resolve()
    copies = [Path(journal['backup_path'])] if journal.get('backup_path') else folder.glob(f'{path.stem}_*{path.suffix}')
    for backup in copies:
        if backup.resolve().is_relative_to(folder) and backup.is_file():
            data = backup.read_bytes()
            if digest_bytes(data) == journal['before_sha256']:
                return {path:data}
    raise ValueError('The verified workbook copy from before this deletion is unavailable.')


def _stage_inward(path, current, original, entry, reserved):
    matches = [batch for batch in read_batches(BytesIO(original))
               if identifier(batch['case_id']) == identifier(entry['sn']) and batch['source_sheet'] == entry['sheet']
               and batch['source_row'] == entry['row']]
    if len(matches) != 1:
        raise ValueError('The deleted inward could not be found in its verified copy.')
    source = matches[0]
    if source.get('pdf_recovery_pending') or source.get('external_pdf') or source.get('hdd_cod_no') or source.get('recycling_completed'):
        raise ValueError('This inward contains issued certificate information and needs its original recovery copy.')
    if source.get('certificate_weight_invalid') or (source.get('raw_inward_date') and not source.get('inward_date')):
        raise ValueError('The recovery copy contains an invalid date or weight. Open it in Excel to recover this inward.')
    existing = read_batches(BytesIO(current))
    used = {identifier(batch['case_id']) for batch in existing}
    serial = source['case_id']
    raw = dict(source, sn='' if identifier(serial) in used else serial)
    staged, added = stage_new_client(SimpleNamespace(path=path, data=current), raw, reserved_serials=reserved, restoring=True)
    expected = batch_content(source)
    expected['case_id'] = added['case_id']
    expected['certificate_type'] = expected['certificate_type'] or 'standard'
    actual = batch_content(added)
    actual['certificate_type'] = actual['certificate_type'] or 'standard'
    if (actual != expected or any(added.get(key) != source.get(key) for key in
            ('issued', 'cod_no', 'cod_date', 'comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale'))):
        raise ValueError(f'The saved details for {source["supplier"]} could not be copied without changes. Use its workbook recovery copy.')
    return staged, added


def _safe_link(record, active_records, batches, own_inwards):
    details = record['details']
    link = details.get('destruction_source')
    if not link:
        return False
    replacement = None
    if link.get('record_id'):
        parent = next((item for item in active_records if item['id'] == link['record_id']), None)
        if parent and norm(parent['details'].get('supplier')) == norm(details.get('supplier')):
            replacement = link
    else:
        plant, sn = link.get('plant'), identifier(link.get('case_id'))
        parent = own_inwards.get((plant, sn))
        candidates = [row for row in batches if row['source_id'] == plant
                      and norm(row.get('supplier')) == norm(details.get('supplier'))]
        if not parent:
            form = identifier(details.get('form6'))
            date = text(details.get('inward_date'))
            matches = [row for row in candidates if (form and identifier(row.get('form6')) == form)
                       or (not form and date and text(row.get('inward_date')) == date)]
            if len(matches) == 1:
                parent = matches[0]
        if parent:
            replacement = dict(link, case_id=str(parent['case_id']))
    if replacement and not any(item['details'].get('destruction_source') == replacement for item in active_records):
        details['destruction_source'] = replacement
        if not replacement.get('record_id'):
            details.update(case_id=str(parent['case_id']), form6=parent.get('form6', ''))
        return False
    record['detached_inward'] = details.pop('destruction_source')
    details.update(case_id='', form6='')
    return True


def _prepare(root, rows, snapshots, paths, originals):
    register = register_path(root).resolve()
    original_register = register.read_bytes() if register.exists() else None
    current = json.loads(original_register) if original_register is not None else {'version':1, 'records':[]}
    records = current['records']
    workbooks = {key:session.data for key, session in snapshots.items()}
    groups, reviewed, kept = [], [], []
    for row, copies in zip(rows, originals):
        inwards = {}
        for entry in row['entries']:
            if entry['source_kind'] != 'workbook':
                continue
            plant = entry['plant']
            path = paths[plant]
            updated, added = _stage_inward(path, workbooks[str(path)], copies[path], entry,
                                          reserved_workbook_serials(root, plant))
            workbooks[str(path)] = updated
            inwards[(plant, identifier(entry['sn']))] = added
            reviewed.append(dict(entry, restore_sn=added['case_id']))
        groups.append((row, copies, inwards))

    batches = workspace_batches(root, workbooks)
    pending_records, pending_links = [], []
    existing_ids = {record['id'] for record in records}
    for row, copies, inwards in groups:
        if register not in copies:
            continue
        historic = {record['id']:record for record in json.loads(copies[register])['records']}
        for entry in row['entries']:
            if entry['source_kind'] != 'manual':
                continue
            key = entry['record_id']
            if key in existing_ids:
                raise ValueError(f'{entry["supplier"]} already exists in the current app register.')
            record = deepcopy(historic[key])
            pending_records.append((entry, record, inwards))
            existing_ids.add(key)
        for link in row.get('kept_hdd', []):
            old = historic.get(link['record_id'], {}).get('details', {}).get('destruction_source') or {}
            if old:
                pending_links.append((link, old, inwards))

    for entry, record, inwards in sorted(pending_records, key=lambda item: bool(item[1]['details'].get('destruction_source'))):
        detached = _safe_link(record, records, batches, inwards)
        records.append(record)
        reviewed.append(dict(entry, restore_standalone=detached))
    by_id = {record['id']:record for record in records}
    for link, old, inwards in pending_links:
        record = by_id.get(link['record_id'])
        if not record or record.get('detached_inward') != old or any(
                record['details'].get(key) for key in ('destruction_source', 'case_id', 'form6')):
            continue
        if old.get('record_id'):
            parent = by_id.get(old['record_id'], {}).get('details')
        else:
            parent = inwards.get((old.get('plant'), identifier(old.get('case_id'))))
        if not parent or not parent.get('case_id'):
            continue
        replacement = dict(old, case_id=parent['case_id'])
        if any(item['details'].get('destruction_source') == replacement for item in records):
            continue
        record['details'].update(destruction_source=replacement, case_id=parent['case_id'],
                                 form6=parent.get('form6', ''))
        record.pop('detached_inward')
        kept.append(link)
    identifiers = [dict(entry, sn=entry.get('restore_sn', entry.get('sn'))) for entry in reviewed]
    require_restored_identifiers(root, identifiers, workbooks, records)
    sources = {path:(snapshots[path].data, payload) for path, payload in workbooks.items() if snapshots[path].data != payload}
    if pending_records or kept:
        if original_register is None:
            raise ValueError('The current app register is missing. Recover that register before restoring deleted entries.')
        sources[str(register)] = original_register, json.dumps(current, indent=2, ensure_ascii=False).encode('utf-8')
    return {'entries':reviewed, 'pdfs':[], 'kept_hdd':kept}, sources


def finish_merge_restore(root, path, transaction):
    """Finish updating deletion markers after the merged files were committed."""
    for target in transaction['targets']:
        original_path = source_journal(root, target)
        data = json.loads(original_path.read_text(encoding='utf-8'))
        if target['kind'] == 'entries':
            data['state'] = 'undone'
        else:
            data['undone'] = True
        save_json(original_path, data)
    transaction['state'] = 'restored'
    save_json(path, transaction)


def _commit(root, rows, sources, sessions):
    identifier = uuid4().hex
    folder = root / 'Data' / 'Entry Deletion Recovery' / identifier
    folder.mkdir(parents=True)
    journal = root / 'Data' / 'Entry Deletions' / (identifier + '.json')
    data = {'version':1, 'id':identifier, 'state':'restoring', 'sources':[], 'pdfs':[],
            'targets':[{key:row[key] for key in ('kind', 'request_id')} for row in rows]}
    for index, (path, (before, after)) in enumerate(sources.items()):
        first, last = f'{index}-before.bin', f'{index}-after.bin'
        _write_bytes(folder/first, before)
        _write_bytes(folder/last, after)
        data['sources'].append({'path':path, 'before':digest_bytes(before), 'after':digest_bytes(after),
                                'before_snapshot':first, 'after_snapshot':last})
    save_json(journal, data)
    try:
        _sync(root, data, True, sessions)
        data['state'] = 'restore_committed'
        save_json(journal, data)
    except Exception:
        _sync(root, data, False, sessions)
        data['state'] = 'rolled_back'
        save_json(journal, data)
        raise
    try:
        finish_merge_restore(root, journal, data)
    except OSError as exc:
        raise ValueError('The entries were restored. Refresh the app to finish updating recovery history before trying again.') from exc


def merge_restores(root, rows, body, *, commit):
    root = Path(root).resolve()
    with run_lock(root), ExitStack() as stack:
        if commit:
            recover_issuances(root)
        journals = [json.loads(source_journal(root, row).read_text(encoding='utf-8')) for row in rows]
        for row, data in zip(rows, journals):
            if (row['kind'] == 'entries' and data.get('state') != 'complete') or data.get('undone') or data.get('undo_started'):
                raise ValueError('A selected deletion changed since the list opened. Reopen the restore search.')
        paths = {f'plant{i}':Path(value).expanduser().resolve() for i, value in enumerate(load_settings(root)['workbook_paths'], 1)
                 if value}
        selected_paths = {paths[entry['plant']] for row in rows for entry in row['entries'] if entry['source_kind'] == 'workbook'}
        configured = list(paths.values())
        if len(configured) == 2 and (configured[0] == configured[1] or
                (all(path.is_file() for path in configured) and configured[0].samefile(configured[1]))):
            raise ValueError('Plant 1 and Plant 2 must use different workbooks.')
        snapshots = {str(path):stack.enter_context(WorkbookSession(path, root/'Data'/'Workbook Backups') if commit
                                                   else WorkbookSnapshot(path)) for path in sorted(selected_paths)}
        overrides = {path:session.data for path, session in snapshots.items()}
        revision = workspace_revision(root, overrides, force=True)
        originals = [_original_sources(root, row, data) for row, data in zip(rows, journals)]
        review, sources = _prepare(root, rows, snapshots, paths, originals)
        token = review_token('/api/entries/restore', {'deletions':rows, 'review':review}, revision, [])
        if not commit:
            return {'review':review, 'review_token':token}
        supplied = body.get('review_token')
        if body.get('checked') is not True or not isinstance(supplied, str) or not supplied.isascii() or not hmac.compare_digest(supplied, token):
            raise ValueError('The current files or restore choices changed. Reopen the restore search and review the entries again.')
        require_revision(root, revision, overrides)
        for data in journals:
            require_deletion_within_restore_window(data)
        _commit(root, rows, sources, snapshots)
        return {'restored':len(review['entries']), 'errors':[], 'remaining':0}
