"""Reviewed removal of mistaken entries before certificate generation.

Active Excel rows, local records and verified PDFs are removed together. Recovery
copies and immutable issuance history remain outside the active certificate folder.
"""
from contextlib import ExitStack
from collections import Counter
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from io import BytesIO
import hmac
import json
import os
from pathlib import Path
import re
from zipfile import BadZipFile

from certificate_core import read_batches, status_value
from certificate_drafts import read_register, register_path
from certificate_inventory import scan_inventory
from inward_removal import stage_removal
from issuance_safety import audit_records, file_digest, digest_bytes, _replace_snapshot, _write_bytes, workbook_receipt_conflict
from workbook_store import WorkbookSession, WorkbookSnapshot
from workflow_support import load_settings, run_lock, save_json
from workspace_review import require_revision, review_token, workspace_revision


ENTRY_RESTORE_WINDOW = timedelta(days=30)
ENTRY_RESTORE_EXPIRED = 'The 30-day restore deadline has passed. To recover older entries, use a workspace backup from before the deletion.'


def deletion_restore_deadline(value):
    deleted_at = datetime.fromisoformat(value)
    if deleted_at.tzinfo is None:
        deleted_at = deleted_at.replace(tzinfo=timezone.utc)
    return deleted_at.astimezone(timezone.utc) + ENTRY_RESTORE_WINDOW


def require_deletion_within_restore_window(journal):
    try:
        deadline = deletion_restore_deadline(journal['at'])
    except (KeyError, ValueError, TypeError, OverflowError) as exc:
        raise ValueError('The date of this deletion could not be verified.') from exc
    if datetime.now(timezone.utc) >= deadline:
        raise ValueError(ENTRY_RESTORE_EXPIRED)


def _key(entry):
    return (entry['source_kind'], entry['plant'], str(entry.get('record_id') if entry['source_kind'] == 'manual' else entry.get('sn')).strip().casefold())


def _require_unissued(details, status, identity, receipts, inventory):
    issued = status_value(status) in {'generated', 'generated_manual', 'completed'}
    issued = issued or bool(details.get('hdd_cod_no'))
    issued = issued or any(row.get('state') != 'rolled_back' and
        (row.get('source_kind'), row.get('plant'), str(row.get('key', '')).strip().casefold()) == identity
        for row in receipts)
    issued = issued or any((owner.get('source_kind'), owner.get('plant'), str(owner.get('key', '')).strip().casefold()) == identity
        for pdf in inventory.get('entries', []) for owner in [pdf.get('managed_owner') or {}])
    if issued:
        raise ValueError('A certificate has already been generated for this entry. It cannot be deleted.')


def deletion_choices(root):
    """List saved records independently of dashboard filters and refresh state."""
    root = Path(root).resolve()
    rows, warnings, snapshots = [], [], {}
    with run_lock(root):
        paths = load_settings(root)['workbook_paths']
        receipts = audit_records(root)
        inventory = scan_inventory(root, persist=False)
        issued = {(row.get('source_kind'), row.get('plant'), str(row.get('key', '')).strip().casefold())
                  for row in receipts if row.get('state') != 'rolled_back'}
        issued.update((owner.get('source_kind'), owner.get('plant'), str(owner.get('key', '')).strip().casefold())
                      for pdf in inventory.get('entries', []) if (owner := pdf.get('managed_owner')))
        records = read_register(root)['records']
        linked_issued = set()

        def is_issued(details, status, identity):
            return (identity in issued or bool(details.get('hdd_cod_no'))
                    or status_value(status) in {'generated', 'generated_manual', 'completed'})

        def add(identity, details, status, **position):
            blocked = 'Certificate already generated' if is_issued(details, status, _key(identity)) else ''
            rows.append(dict(identity, **position, supplier=details.get('supplier'), location=details.get('location'),
                             certificate_type=details.get('certificate_type', 'standard'),
                             certificate_no=details.get('cod_no'), reference=details.get('reference'),
                             form6=details.get('form6'), certificate_status=status_value(status),
                             destruction_source=details.get('destruction_source'), deletion_blocked=blocked))
            return rows[-1]

        for record in records:
            details = record['details']
            identity = {'source_kind':'manual', 'plant':record['plant'], 'record_id':record['id']}
            row = add(identity, details, record.get('status'), removed_inward=record.get('removed_inward'))
            link = details.get('destruction_source') or {}
            if row['deletion_blocked'] and link:
                linked_issued.add(('manual', str(link['record_id']).casefold()) if link.get('record_id')
                                  else (link.get('plant'), str(link.get('case_id', '')).strip().casefold()))

        resolved = [Path(value).expanduser().resolve() if value else None for value in paths]
        shared = all(resolved) and (resolved[0] == resolved[1] or
                 (all(path.is_file() for path in resolved) and resolved[0].samefile(resolved[1])))
        for index, path in enumerate(resolved, 1):
            if path is None:
                warnings.append(f'Plant {index} has no connected workbook.')
                continue
            try:
                with WorkbookSnapshot(path) as snapshot:
                    batches = read_batches(BytesIO(snapshot.data))
                    snapshots[str(path)] = snapshot.data
            except (OSError, ValueError, KeyError, TypeError, RuntimeError, BadZipFile) as exc:
                warnings.append(f'Plant {index} could not be read: {exc}')
                continue
            plant = f'plant{index}'
            counts = Counter(str(batch['case_id']).strip().casefold() for batch in batches)
            for batch in batches:
                identity = {'source_kind':'workbook', 'plant':plant, 'sn':str(batch['case_id']),
                            'sheet':batch['source_sheet'], 'row':batch['source_row']}
                row = add(identity, batch, batch.get('issued'))
                if shared:
                    row['deletion_blocked'] = 'Connect a different workbook for each plant'
                elif counts[_key(identity)[2]] > 1:
                    row['deletion_blocked'] = 'Repeated serial number. Check Excel first.'
            for issue in getattr(batches, 'row_issues', ()):
                rows.append({**{key:issue.get(key) for key in ('sheet', 'row', 'supplier', 'sn', 'location', 'form6', 'reference')},
                             'source_kind':'workbook', 'plant':plant,
                             'certificate_no':', '.join(issue.get('reserved_cod_numbers', [])),
                             'deletion_blocked':'Check this row in Excel first'})

        for row in rows:
            source = ('manual', str(row['record_id']).casefold()) if row['source_kind'] == 'manual' else (
                row['plant'], str(row.get('sn') or '').strip().casefold())
            if not row['deletion_blocked'] and source in linked_issued:
                row['deletion_blocked'] = 'Linked Data Destruction certificate already generated'
        try:
            revision = workspace_revision(root, snapshots)
        except (OSError, ValueError) as exc:
            revision = None
            warnings.append(str(exc))
            for row in rows:
                row['deletion_blocked'] = row['deletion_blocked'] or 'Workspace files could not be checked'
        return {'rows':rows, 'warnings':warnings, 'workspace_revision':revision}


def deletion_history(root, *, selected=None):
    """List recent deletions, or look up specific choices even if they expired."""
    data_root = Path(root) / 'Data'
    now = datetime.now(timezone.utc)
    for kind, directory in (('entries', 'Entry Deletions'), ('inward', 'Inward Removals')):
        for path in (data_root / directory).glob('*.json'):
            if not re.fullmatch(r'[a-f0-9]{32}', path.stem):
                continue
            if selected is not None and (kind, path.stem) not in selected:
                continue
            try:
                data = json.loads(path.read_text(encoding='utf-8'))
                deadline = deletion_restore_deadline(data['at'])
            except (OSError, ValueError, KeyError, TypeError, OverflowError):
                continue
            if selected is None and now >= deadline:
                continue
            if kind == 'entries':
                if data.get('id') != path.stem or data.get('state') not in {'complete', 'undoing'}:
                    continue
                review = data.get('review')
                entries = review.get('entries') if isinstance(review, dict) else None
            else:
                if data.get('undone'):
                    continue
                if not data.get('saved'):
                    source, after = data.get('path'), data.get('after_sha256')
                    try:
                        if not source or not after or file_digest(Path(source)) != after:
                            continue
                    except (OSError, ValueError):
                        continue
                entries = [data.get('entry', {})]
            if isinstance(entries, list) and entries and all(isinstance(entry, dict) for entry in entries):
                yield deadline, {'kind':kind, 'request_id':path.stem, 'at':data['at'],
                                 'restore_until':deadline.isoformat()}, entries, data


def last_deletion(root):
    """Find the latest deletion for the standalone workspace recovery tool."""
    latest = max(deletion_history(root), key=lambda item: item[0], default=None)
    if latest is None:
        return None
    _, target, entries, _ = latest
    return dict(target, count=len(entries), supplier=entries[0].get('supplier'))


def undo_deletion(root, target):
    if target['kind'] == 'inward':
        from inward_removal import undo_inward_removal
        return undo_inward_removal(root, target['request_id'], None)
    return undo_entry_deletion(root, target['request_id'], None)


def _selected(body):
    choices = body.get('selected')
    if not isinstance(choices, list) or not 1 <= len(choices) <= 50:
        raise ValueError('Select between 1 and 50 entries to review.')
    result = []
    for choice in choices:
        if (not isinstance(choice, dict) or choice.get('source_kind') not in {'manual', 'workbook'}
                or choice.get('plant') not in ({'plant1','plant2',None} if choice.get('source_kind') == 'manual' else {'plant1','plant2'})):
            raise ValueError('Select an entry from the Settings search.')
        if choice['source_kind'] == 'workbook':
            if not isinstance(choice.get('sn'), str) or not choice['sn'].strip() or not isinstance(choice.get('sheet'), str) or type(choice.get('row')) is not int:
                raise ValueError('The selected Excel inward needs a fresh review.')
            result.append({key:choice[key] for key in ('source_kind','plant','sn','sheet','row')})
        else:
            if not isinstance(choice.get('record_id'), str) or not choice['record_id']:
                raise ValueError('Select a saved local record from the search.')
            result.append({key:choice.get(key) for key in ('source_kind','plant','record_id')})
    if len({(entry['source_kind'], None if entry['source_kind']=='manual' else entry['plant'],
             str(entry.get('record_id') if entry['source_kind']=='manual' else entry['sn']).strip().casefold()) for entry in result}) != len(result):
        raise ValueError('An entry was selected twice. Search and select it once.')
    return result


def _prepare(root, body, sessions):
    selected = _selected(body)
    settings = load_settings(root)
    registry = read_register(root)
    for choice in selected:
        if choice['source_kind'] == 'manual':
            record = next((row for row in registry['records'] if row['id'] == choice['record_id']), None)
            if record is not None:
                choice['plant'] = record['plant']
    originals = deepcopy(registry)
    inventory = scan_inventory(root, persist=False)
    receipts = audit_records(root)
    entries, sources = [], {}
    identities = set(map(_key, selected))
    manual_ids = {choice['record_id'] for choice in selected if choice['source_kind'] == 'manual'}
    workbook_links = {(choice['plant'], choice['sn'].strip().casefold()) for choice in selected if choice['source_kind'] == 'workbook'}
    linked = []
    for record in registry['records']:
        link = record['details'].get('destruction_source') or {}
        if (link.get('record_id') in manual_ids or (not link.get('record_id') and
                (link.get('plant'), str(link.get('case_id', '')).strip().casefold()) in workbook_links)):
            linked.append(record)
            _require_unissued(record['details'], record.get('status'),
                              ('manual', record['plant'], record['id'].casefold()), receipts, inventory)
    for choice in selected:
        plant = choice['plant']
        if choice['source_kind'] == 'workbook':
            value = settings['workbook_paths'][int(plant[-1])-1]
            path = Path(value).expanduser().resolve()
            session = sessions[str(path)]
            before = sources[str(path)][1] if str(path) in sources else session.data
            matches = [row for row in read_batches(BytesIO(before)) if str(row['case_id']).strip().casefold() == choice['sn'].strip().casefold()]
            if len(matches) != 1:
                raise ValueError('An SN is missing or repeated. Check Excel and search again.')
            source = matches[0]
            original_match = next((row for row in read_batches(BytesIO(session.data))
                                   if str(row['case_id']).strip().casefold() == choice['sn'].strip().casefold()), None)
            if not original_match or (original_match['source_sheet'], original_match['source_row']) != (choice['sheet'], choice['row']):
                raise ValueError('This inward moved in Excel. Search and review it again.')
            _require_unissued(source, source.get('issued'), _key(choice), receipts, inventory)
            previous = next((row for row in reversed(receipts) if row.get('state') == 'complete'
                and (row.get('source_kind'),row.get('plant'),str(row.get('key')).strip().casefold()) == _key(choice)), None)
            conflict = workbook_receipt_conflict(source, previous)
            if conflict:
                raise ValueError(conflict)
            updated, _, _ = stage_removal(before, choice['sn'], source['source_sheet'], source['source_row'])
            original = sources[str(path)][0] if str(path) in sources else session.data
            sources[str(path)] = original, updated
            cod = source.get('cod_no') or (previous or {}).get('cod', '')
            details = source
            status = source.get('issued','')
        else:
            record = next((row for row in registry['records'] if row['id'] == choice['record_id'] and row['plant'] == plant), None)
            if record is None:
                raise ValueError('A selected local record no longer exists. Search again.')
            details, cod = record['details'], record['details'].get('cod_no', '')
            status = record.get('status','')
            _require_unissued(details, status, _key(choice), receipts, inventory)
        entries.append(dict(choice, supplier=details.get('supplier', ''), location=details.get('location', ''),
                            certificate_type=details.get('certificate_type', 'standard'), cod=str(cod or ''),
                            status=status, received=str(details.get('inward_date') or ''),
                            form6=details.get('form6',''), reference=details.get('reference',''),
                            materials=len(details.get('items', []))))
    for path, (before, after) in list(sources.items()):
        from inward_serials import resequence_removed_inwards
        from workbook_store import order_register_for_save
        plant = next(f'plant{index}' for index, value in enumerate(settings['workbook_paths'], 1)
                     if value and str(Path(value).expanduser().resolve()) == path)
        after, _ = resequence_removed_inwards(after, before, root=root, plant=plant)
        from register_protection import remove_issued_inward_protection
        sources[path] = before, remove_issued_inward_protection(order_register_for_save(after, sessions[path]))
    kept_hdd = []
    for record in linked:
        if record['id'] in manual_ids:
            continue
        kept_hdd.append({'supplier':record['details'].get('supplier',''), 'cod':record['details'].get('cod_no',''), 'record_id':record['id']})
        record['detached_inward'] = deepcopy(record['details'].pop('destruction_source'))
        record['details'].update(case_id='', form6='')
    registry['records'] = [row for row in registry['records'] if row['id'] not in manual_ids]
    from manual_ownership import normalize_records
    normalize_records(registry)
    if registry != originals:
        path = register_path(root)
        sources[str(path.resolve())] = (path.read_bytes(), json.dumps(registry, indent=2, ensure_ascii=False, default=str).encode('utf-8'))
    pdfs = []
    for pdf in inventory.get('entries', []):
        owner = pdf.get('managed_owner') or {}
        identity = (owner.get('source_kind'), owner.get('plant'), str(owner.get('key','')).strip().casefold())
        if pdf.get('present') and identity in identities:
            path = (root/pdf['path']).resolve()
            if not path.is_relative_to(root/'Certificates') or file_digest(path) != pdf.get('sha256'):
                raise ValueError('A PDF changed or its path could not be verified. Search and review again.')
            pdfs.append({'path':pdf['path'], 'sha256':pdf['sha256']})
    ids = [row['id'] for row in receipts if row.get('state') == 'complete' and
        (row.get('source_kind'),row.get('plant'),str(row.get('key','')).strip().casefold()) in identities]
    review = {'entries': entries, 'pdfs': [pdf['path'] for pdf in pdfs], 'kept_hdd': kept_hdd,
              'include_linked_hdd':False}
    return review, sources, pdfs, ids


def _paths(root, data):
    configured = {str(Path(value).expanduser().resolve()) for value in load_settings(root)['workbook_paths'] if value}
    for entry in data['sources']:
        path = Path(entry['path']).resolve()
        if str(path) not in configured and path != register_path(root).resolve():
            raise ValueError('Reconnect the original workbook before recovering a mistaken-entry deletion.')
        if any(not re.fullmatch(r'[0-9]+-'+side+r'\.bin', str(entry.get(side+'_snapshot','')))
               or not re.fullmatch(r'[a-f0-9]{64}', str(entry.get(side,''))) for side in ('before','after')):
            raise ValueError('Invalid source deletion recovery snapshot.')
    for pdf in data['pdfs']:
        path, archive = (root/pdf['path']).resolve(), (root/pdf['archive']).resolve()
        if not path.is_relative_to(root/'Certificates') or not archive.is_relative_to(root/'Data'/'Archive'/'Deleted Entries'):
            raise ValueError('Invalid PDF deletion recovery path.')
        if not re.fullmatch(r'pdf-[0-9]+\.pdf', str(pdf.get('snapshot',''))):
            raise ValueError('Invalid PDF deletion recovery snapshot.')


def _sync(root, data, committed, sessions):
    _paths(root, data)
    folder = root/'Data'/'Entry Deletion Recovery'/data['id']
    # Check the whole set before changing any file during commit or recovery.
    for item in data['sources']:
        path = Path(item['path']).resolve()
        session = sessions.get(str(path))
        actual = digest_bytes(session.data) if session else file_digest(path)
        if actual not in {item['before'], item['after']}:
            raise ValueError('A source changed during deletion. Recovery copies were preserved; restore a backup before continuing.')
    for pdf in data['pdfs']:
        original, archived = file_digest(root/pdf['path']), file_digest(root/pdf['archive'])
        if original not in {None,pdf['sha256']} or archived not in {None,pdf['sha256']}:
            raise ValueError('A PDF changed during deletion. Recovery copies were preserved.')
        if file_digest(folder/pdf['snapshot']) != pdf['sha256']:
            raise ValueError('The PDF deletion recovery copy failed verification.')
    for item in data['sources']:
        side = 'after' if committed else 'before'
        path = Path(item['path']).resolve()
        session = sessions.get(str(path))
        actual = digest_bytes(session.data) if session else file_digest(path)
        if actual == item[side]:
            continue
        snapshot = folder/item[side+'_snapshot']
        payload = snapshot.read_bytes()
        if digest_bytes(payload) != item[side]:
            raise ValueError('The deletion recovery copy failed verification. Nothing was overwritten.')
        if session:
            session._persist(payload, exact=True)
        else:
            _replace_snapshot(snapshot, path, path.with_name('.'+data['id']+'.delete-recovery'), item[side])
    for pdf in data['pdfs']:
        source = root/pdf['path' if committed else 'archive']
        destination = root/pdf['archive' if committed else 'path']
        if source.is_file() and not destination.exists():
            destination.parent.mkdir(parents=True, exist_ok=True)
            os.replace(source, destination)
        elif not destination.exists():
            _replace_snapshot(folder/pdf['snapshot'], destination,
                              destination.with_name('.'+data['id']+'.recover.pdf'), pdf['sha256'])


def recover_entry_deletions(root):
    root = Path(root).resolve()
    recovered = []
    for path in sorted((root/'Data'/'Entry Deletions').glob('*.json')):
        data = json.loads(path.read_text(encoding='utf-8'))
        if data.get('state') in {'complete','rolled_back','undone','restored'}:
            continue
        if not re.fullmatch(r'[a-f0-9]{32}', str(data.get('id',''))) or data.get('version') != 1:
            raise ValueError('The entry deletion recovery record needs checking.')
        if data['state'] == 'restore_committed':
            from entry_restore_merge import finish_merge_restore
            finish_merge_restore(root, path, data)
            recovered.append('Finished restoring the deleted entries.')
            continue
        _paths(root, data)
        with ExitStack() as stack:
            sessions = {item['path']:stack.enter_context(WorkbookSession(Path(item['path']), root/'Data'/'Workbook Backups'))
                        for item in data['sources'] if Path(item['path']).resolve() != register_path(root).resolve()}
            _sync(root, data, False, sessions)
        undo = data['state'] == 'undoing'
        data['state'] = 'undone' if undo else 'rolled_back'
        save_json(path, data)
        recovered.append('Finished restoring the deleted entries.' if undo
                         else 'Restored an interrupted mistaken-entry deletion. Review and delete again if needed.')
    return recovered


def delete_entries(root, body, *, commit=False):
    root = Path(root).resolve()
    selected = _selected(body)
    identifier = body.get('request_id')
    if not re.fullmatch(r'[a-f0-9]{32}', str(identifier or '')):
        raise ValueError('Open the Settings deletion search and review the selected entries.')
    # Allow a retry to read its old committed journal, but require explicit
    # selection of each entry for any new deletion.
    legacy_include = body.get('include_linked_hdd', False)
    if type(legacy_include) is not bool:
        raise ValueError('Search and review the selected entries again.')
    request = {'selected':selected, 'include_linked_hdd':legacy_include, 'request_id':identifier}
    journal = root/'Data'/'Entry Deletions'/(identifier+'.json')
    with run_lock(root):
        from issuance_safety import recover_issuances
        recover_issuances(root)
        if journal.exists():
            existing = json.loads(journal.read_text(encoding='utf-8'))
            if existing.get('request') != request:
                raise ValueError('This deletion request belongs to different entries.')
            if existing.get('state') == 'complete' and commit:
                return {'deleted':True, 'review':existing['review']}
            raise ValueError('Review the selection again before retrying this deletion.')
        if legacy_include:
            raise ValueError('Select each Data Destruction entry you want to delete, then review the selection again.')
        settings = load_settings(root)
        paths = []
        for choice in selected:
            if choice['source_kind'] == 'workbook':
                value = settings['workbook_paths'][int(choice['plant'][-1])-1]
                if not value:
                    raise ValueError('Connect the selected plant workbook before deleting its inward.')
                paths.append(Path(value).expanduser().resolve())
        connected = [Path(value).expanduser().resolve() for value in settings['workbook_paths'] if value]
        if len(connected) == 2 and (connected[0] == connected[1] or (connected[0].exists() and connected[1].exists() and connected[0].samefile(connected[1]))):
            raise ValueError('Plant 1 and Plant 2 must use different Excel files.')
        with ExitStack() as stack:
            sessions = {str(path):stack.enter_context(WorkbookSession(path, root/'Data'/'Workbook Backups') if commit else WorkbookSnapshot(path)) for path in dict.fromkeys(paths)}
            overrides = {path:session.data for path,session in sessions.items()}
            require_revision(root, body.get('expected_revision'), overrides)
            review, sources, pdfs, receipt_ids = _prepare(root, body, sessions)
            token = review_token('/api/entries/delete', {'request':request,'review':review}, body['expected_revision'], [])
            if not commit:
                return {'review':review, 'review_token':token}
            if body.get('checked') is not True or not isinstance(body.get('review_token'),str) or not hmac.compare_digest(body['review_token'],token):
                raise ValueError('Check the selected entries and PDFs in the deletion review first.')
            folder = root/'Data'/'Entry Deletion Recovery'/identifier
            folder.mkdir(parents=True, exist_ok=False)
            data = {'version':1, 'id':identifier, 'state':'prepared', 'request':request, 'review':review,
                    'at':datetime.now(timezone.utc).isoformat(), 'receipt_ids':receipt_ids,
                    'cods':[entry['cod'] for entry in review['entries'] if entry['cod']], 'sources':[], 'pdfs':[]}
            for index,(path,(before,after)) in enumerate(sources.items()):
                first, last = f'{index}-before.bin', f'{index}-after.bin'
                _write_bytes(folder/first,before); _write_bytes(folder/last,after)
                data['sources'].append({'path':path,'before':digest_bytes(before),'after':digest_bytes(after),
                                        'before_snapshot':first,'after_snapshot':last})
            for index,pdf in enumerate(pdfs):
                snapshot = f'pdf-{index}.pdf'
                _write_bytes(folder/snapshot, (root/pdf['path']).read_bytes())
                data['pdfs'].append(dict(pdf, snapshot=snapshot,
                    archive=f'Data/Archive/Deleted Entries/{identifier}/{index}/{Path(pdf["path"]).name}'))
            save_json(journal, data)
            require_revision(root, body.get('expected_revision'), overrides)
            try:
                _sync(root, data, True, sessions)
                data['state'] = 'complete'; save_json(journal, data)
            except Exception:
                _sync(root, data, False, sessions)
                data['state'] = 'rolled_back'; save_json(journal, data)
                raise
        return {'deleted':True, 'review':review}


def undo_entry_deletion(root, identifier, revision):
    """Restore the reviewed deletion only while all affected files still match."""
    from issuance_safety import recover_issuances

    root = Path(root).resolve()
    if not isinstance(identifier, str) or not re.fullmatch(r'[a-f0-9]{32}', identifier):
        raise ValueError('Choose a saved deletion to undo.')
    journal = root/'Data'/'Entry Deletions'/f'{identifier}.json'
    with run_lock(root), ExitStack() as stack:
        recover_issuances(root)
        if not journal.is_file():
            raise ValueError('The deletion record was not found.')
        data = json.loads(journal.read_text(encoding='utf-8'))
        if data.get('state') == 'undone':
            return {'undone':True}
        if data.get('state') != 'complete' or data.get('id') != identifier or data.get('version') != 1:
            raise ValueError('This deletion has no completed change to undo.')
        _paths(root, data)
        sessions = {item['path']:stack.enter_context(WorkbookSession(Path(item['path']), root/'Data'/'Workbook Backups'))
                    for item in data['sources'] if Path(item['path']).resolve() != register_path(root).resolve()}
        if revision is not None:
            require_revision(root, revision, {path:session.data for path,session in sessions.items()})
        folder = root/'Data'/'Entry Deletion Recovery'/identifier
        restored_workbooks = {}
        restored_records = None
        for item in data['sources']:
            session = sessions.get(item['path'])
            actual = digest_bytes(session.data) if session else file_digest(Path(item['path']))
            if actual != item['after']:
                raise ValueError(f'{Path(item["path"]).name} changed after the deletion. '
                                 f'The pre-deletion copies are in {folder}. To roll back the workspace, use Restore backup.')
            before = (folder/item['before_snapshot']).read_bytes()
            if digest_bytes(before) != item['before']:
                raise ValueError('A pre-deletion recovery copy failed verification.')
            if session:
                restored_workbooks[item['path']] = before
            else:
                restored_records = json.loads(before)['records']
        from identifier_rules import require_restored_identifiers
        require_restored_identifiers(root, data['review']['entries'], restored_workbooks,
                                     read_register(root)['records'] if restored_records is None else restored_records)
        for pdf in data['pdfs']:
            if file_digest(root/pdf['path']) is not None:
                raise ValueError('Another PDF now uses a deleted certificate’s filename. The recovery copies were kept.')
            if (file_digest(folder/pdf['snapshot']) != pdf['sha256']
                    or file_digest(root/pdf['archive']) not in {None, pdf['sha256']}):
                raise ValueError('A PDF deletion recovery copy failed verification.')
        require_deletion_within_restore_window(data)
        data['state'] = 'undoing'
        save_json(journal, data)
        _sync(root, data, False, sessions)
        data['state'] = 'undone'
        try:
            save_json(journal, data)
        except OSError:
            pass  # Recovery can complete the journal using the restored hashes.
    return {'undone':True}
