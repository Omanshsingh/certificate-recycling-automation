"""Search, review, and restore saved entry deletions."""
from datetime import datetime, timezone
import hmac
from pathlib import Path
import re
from zipfile import BadZipFile

from entry_deletion import ENTRY_RESTORE_EXPIRED, _paths, deletion_history, undo_deletion
from file_fingerprints import fingerprint
from identifier_rules import identifier, workspace_batches
from workflow_support import load_settings
from workspace_review import review_token


def _availability(root, target, data, paths, digests):
    """Check recovery copies and decide whether to merge into the current files."""
    changed = False

    def probe(source, allowed):
        nonlocal changed
        path = Path(source)
        if path not in digests:
            digests[path] = fingerprint(path)[0]
        changed |= digests[path] not in allowed

    try:
        if target['kind'] == 'inward':
            plant = data['entry'].get('plant')
            if plant not in {'plant1', 'plant2'} or paths[int(plant[-1])-1] != Path(data['path']).resolve():
                return 'Reconnect the original workbook to restore this inward.', False
            allowed = {data['after_sha256']}
            if data.get('undo_started'):
                allowed.add(data['before_sha256'])
            probe(data['path'], allowed)
            backup = data.get('backup_path')
            folder = root / 'Data' / 'Workbook Backups'
            if backup:
                copy = Path(backup).resolve()
                if not copy.is_relative_to(folder.resolve()) or not copy.is_file():
                    return 'The saved recovery copy is unavailable.', False
            elif not folder.is_dir():
                return 'The saved recovery copy is unavailable.', False
        else:
            _paths(root, data)
            folder = root / 'Data' / 'Entry Deletion Recovery' / target['request_id']
            for source in data['sources']:
                allowed = {source['after']}
                if data['state'] == 'undoing':
                    allowed.add(source['before'])
                probe(source['path'], allowed)
                if not (folder / source['before_snapshot']).is_file():
                    return 'The saved recovery copy is unavailable.', False
            for pdf in data['pdfs']:
                if data['state'] != 'undoing' and (root / pdf['path']).exists():
                    return 'Another PDF now uses the same filename.', False
                if not (folder / pdf['snapshot']).is_file():
                    return 'The saved PDF recovery copy is unavailable.', False
    except FileNotFoundError:
        return 'An affected workbook or register is missing.', False
    except PermissionError:
        return 'An affected workbook could not be read. Close Excel and reopen this list.', False
    except (OSError, ValueError, TypeError, KeyError) as exc:
        return str(exc) or 'The recovery record could not be checked.', False
    if changed and (data.get('pdfs') or data.get('receipt_ids') or data.get('state') == 'undoing' or data.get('undo_started')):
        return 'This recovery includes certificate history or an interrupted restore. It needs its unchanged recovery files.', False
    return '', changed


def restore_choices(root, selected=None):
    root = Path(root).resolve()
    workbooks = load_settings(root)['workbook_paths']
    paths = [Path(path).expanduser().resolve() if path else None for path in workbooks]
    digests = {}
    rows = []
    now = datetime.now(timezone.utc)
    history = deletion_history(root, selected=selected)
    for deadline, target, entries, data in sorted(history, key=lambda item: item[0], reverse=True):
        review = data.get('review', {})
        if target['kind'] == 'inward':
            snapshot = data.get('snapshot') or {}
            entries = [dict(entries[0], source_kind='workbook', certificate_type=snapshot.get('certificate_type'),
                            location=snapshot.get('location'), reference=snapshot.get('reference'),
                            form6=snapshot.get('form6'), received=entries[0].get('inward_date'))]
        blocked, merge = (ENTRY_RESTORE_EXPIRED, False) if now >= deadline else _availability(root, target, data, paths, digests)
        rows.append(dict(target, entries=entries, pdfs=review.get('pdfs', []), kept_hdd=review.get('kept_hdd', []),
                         restore_blocked=blocked, merge_required=merge,
                         exact_only=bool(data.get('pdfs') or data.get('receipt_ids') or data.get('state') == 'undoing' or data.get('undo_started'))))
    if any(not row['restore_blocked'] and any(entry.get('source_kind') == 'workbook' and entry.get('form6')
                                               for entry in row['entries']) for row in rows):
        try:
            forms = {identifier(batch.get('form6')) for batch in workspace_batches(root)}
        except (OSError, ValueError, KeyError, RuntimeError, BadZipFile):
            forms = set()  # The locked review checks both workbooks again before a restore.
        for row in rows:
            if row['restore_blocked']:
                continue
            reused = next((entry['form6'] for entry in row['entries'] if entry.get('source_kind') == 'workbook'
                           and entry.get('form6') and identifier(entry['form6']) in forms), None)
            if reused:
                row['restore_blocked'] = f'Form 6 number {reused} is already assigned to an active entry.'
    return rows


def restore_entries(root, body, *, commit=False):
    choices = body.get('selected')
    if not isinstance(choices, list) or not 1 <= len(choices) <= 50:
        raise ValueError('Select between 1 and 50 deletions to review.')
    selected = set()
    for choice in choices:
        if not isinstance(choice, dict) or choice.get('kind') not in ('entries', 'inward') or not (
            isinstance(choice.get('request_id'), str) and re.fullmatch(r'[a-f0-9]{32}', choice['request_id'])
        ):
            raise ValueError('Choose deletions from the restore search.')
        selected.add((choice['kind'], choice['request_id']))
    if len(selected) != len(choices):
        raise ValueError('Select each deletion only once.')

    rows = restore_choices(root, selected)
    if len(rows) != len(selected):
        raise ValueError('A selected deletion was already restored or is missing. Reopen the restore search.')
    for row in rows:
        if row['restore_blocked']:
            raise ValueError(row['restore_blocked'])

    if len(rows) > 1 or rows[0]['merge_required']:
        if any(row['exact_only'] for row in rows):
            raise ValueError('Restore the selection with certificate history on its own, then restore the remaining entries.')
        from entry_restore_merge import merge_restores
        return merge_restores(root, rows, body, commit=commit)
    token = review_token('/api/entries/restore', {'deletions':rows}, '', [])
    if not commit:
        review = {key:[entry for row in rows for entry in row[key]] for key in ('entries', 'pdfs', 'kept_hdd')}
        return {'review':review, 'review_token':token}
    supplied = body.get('review_token')
    if body.get('checked') is not True or not isinstance(supplied, str) or not supplied.isascii() or not hmac.compare_digest(supplied, token):
        raise ValueError('Review and confirm the selected entries before restoring.')

    restored, errors = 0, []
    for row in rows:
        try:
            undo_deletion(root, row)
            restored += len(row['entries'])
        except (OSError, ValueError, KeyError, RuntimeError, BadZipFile) as exc:
            company = row['entries'][0].get('supplier') or 'Unnamed company'
            label = company if len(row['entries']) == 1 else f"{len(row['entries'])} entries including {company}"
            errors.append(f'{label}: {exc}')
            # An incomplete restore may require recovery before another write.
            break
    if errors and not restored:
        raise ValueError(errors[0])
    return {'restored':restored, 'errors':errors, 'remaining':sum(len(row['entries']) for row in rows) - restored}
