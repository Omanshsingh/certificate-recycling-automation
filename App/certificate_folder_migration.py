"""Move the former app folder without rewriting immutable issuance history."""
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
from urllib.parse import quote
import uuid

from workflow_support import run_lock, save_json

LEGACY = 'Created in App'
DESTINATION = 'Auto Generated'


def _ledger_path(root):
    return Path(root) / 'Data' / 'Certificate Folder Moves.json'


def _ledger(root):
    path = _ledger_path(root)
    if not path.exists():
        return {'version': 1, 'moves': [], 'state': 'complete'}
    value = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(value, dict) or value.get('version') != 1 or not isinstance(value.get('moves'), list):
        raise ValueError('The saved certificate folder move could not be read. Files were kept safe.')
    for move in value['moves']:
        if not isinstance(move, dict) or any(not isinstance(move.get(key), str) for key in ('old', 'new', 'sha256')):
            raise ValueError('The saved certificate folder move could not be read. Files were kept safe.')
        for key, folder in (('old', LEGACY), ('new', DESTINATION)):
            parts = move[key].replace('\\', '/').split('/')
            if len(parts) < 3 or parts[:2] != ['Certificates', folder] or any(p in {'', '.', '..'} or ':' in p for p in parts):
                raise ValueError('The saved certificate folder move has an invalid path. Files were kept safe.')
    return value


def moved_relative_path(root, relative, expected_sha=None):
    """Resolve an old link only when an actual recorded move proves its target."""
    raw = str(relative).replace('\\', '/')
    if not raw.startswith(f'{LEGACY}/'):
        return raw
    original = Path(root) / 'Certificates' / raw
    if original.exists():
        return raw  # A restored backup may still be in its original location.
    key = 'Certificates/' + raw
    for move in reversed(_ledger(root)['moves']):
        if move['old'].casefold() == key.casefold() and (expected_sha is None or move['sha256'] == expected_sha):
            return move['new'][len('Certificates/'):]
    return raw


def _digest(path):
    with Path(path).open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def _available(path, reserved):
    if not path.exists() and str(path).casefold() not in reserved:
        return path
    stem = path.stem
    hdd = '_hdd' if stem.casefold().endswith('_hdd') else ''
    if hdd:
        stem = stem[:-4]
    for index in range(2, 10000):
        suffix = f' ({index}){hdd}{path.suffix}'
        budget = 255 - len(suffix.encode('utf-16-le')) // 2
        kept = stem
        while len(kept.encode('utf-16-le')) // 2 > budget:
            kept = kept[:-1]
        candidate = path.with_name(kept.rstrip(' .') + suffix)
        if not candidate.exists() and str(candidate).casefold() not in reserved:
            return candidate
    raise ValueError('There are too many PDFs with the same name. Rename one and reopen the app.')


def _rewrite(value, replacements):
    if isinstance(value, list):
        return [_rewrite(item, replacements) for item in value]
    if not isinstance(value, dict):
        return value
    result = {key: _rewrite(item, replacements) for key, item in value.items()}
    changed_path = None
    for key in ('path', 'pdf_path', 'pdf_relative_path', 'relative_path', 'pdf_url'):
        original = result.get(key)
        if isinstance(original, str) and original.replace('\\', '/') in replacements:
            result[key] = replacements[original.replace('\\', '/')]
            if key in {'pdf_relative_path', 'relative_path', 'pdf_path'}:
                changed_path = result[key]
    if changed_path:
        for key in ('pdf_filename', 'filename'):
            if key in result:
                result[key] = Path(changed_path).name
    return result


def migrate_certificate_folder(root):
    root = Path(root).resolve()
    legacy = root / 'Certificates' / LEGACY
    if not legacy.exists() and not _ledger_path(root).exists():
        return
    if not legacy.exists() and _ledger(root).get('state') == 'complete':
        return
    with run_lock(root):
        from issuance_safety import recover_issuances
        recover_issuances(root)
        ledger = _ledger(root)
        if legacy.exists() and (legacy.is_symlink() or not legacy.resolve().is_relative_to(root / 'Certificates')):
            raise ValueError('The old certificate folder points outside Certificates. Files were kept safe.')
        items = sorted(legacy.rglob('*')) if legacy.is_dir() else []
        files = [path for path in items if path.is_file()]
        if any(path.is_symlink() or not path.resolve().is_relative_to(legacy.resolve()) for path in items):
            raise ValueError('The old certificate folder contains an external link. Files were kept safe.')
        unfinished = ledger.get('state') == 'moving'
        pending = ledger.get('pending', []) if unfinished else []
        if not isinstance(pending, list) or any(type(i) is not int or not 0 <= i < len(ledger['moves']) for i in pending):
            raise ValueError('The saved certificate folder move could not be read. Files were kept safe.')
        planned_sources = {ledger['moves'][i]['old'] for i in pending}
        additions, reserved = [], {str(root / move['new']).casefold() for move in ledger['moves']}
        for path in files:
            old = path.relative_to(root).as_posix()
            if old in planned_sources:
                continue
            target = root / 'Certificates' / DESTINATION / path.relative_to(legacy)
            target = _available(target, reserved)
            reserved.add(str(target).casefold())
            additions.append({'old': old, 'new': target.relative_to(root).as_posix(), 'sha256': _digest(path)})
        active = [root / 'Data' / name for name in ('Custom Certificates.json', 'Certificate Inventory.json', 'Reports/Latest Results.json')]
        # Validate every active JSON file before moving any PDF.
        original_bytes = {path: path.read_bytes() for path in active if path.is_file()}
        documents = {path: json.loads(payload.decode('utf-8')) for path, payload in original_bytes.items()}
        if additions:
            backup = root / 'Data' / 'Certificate Folder Backups' / f'{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}'
            for path in list(documents) + [root / move['old'] for move in additions]:
                destination = backup / path.relative_to(root)
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, destination)
            pending.extend(range(len(ledger['moves']), len(ledger['moves']) + len(additions)))
            ledger['moves'].extend(additions)
            ledger['pending'] = pending
            ledger['state'] = 'moving'
            save_json(_ledger_path(root), ledger)
        if additions or unfinished:
            for move in (ledger['moves'][i] for i in pending):
                source, destination = root / move['old'], root / move['new']
                if any(not path.resolve().is_relative_to(root / 'Certificates') for path in (source, destination)):
                    raise ValueError('A certificate folder points outside Certificates. Files were kept safe.')
                if source.exists():
                    if _digest(source) != move['sha256']:
                        raise ValueError('A certificate changed during its folder move. Files were kept safe; reopen the app to check it.')
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if not destination.exists():
                        # Windows rename fails if the target appears concurrently.
                        os.rename(source, destination)
                    elif _digest(destination) == move['sha256']:
                        source.unlink()  # The verified target and backup both exist.
                    else:
                        raise ValueError('Another PDF appeared during the folder move. Both files were kept safe.')
                if not destination.is_file() or _digest(destination) != move['sha256']:
                    raise ValueError('A moved certificate could not be checked. Its backup was kept safe.')
            replacements = {}
            for move in ledger['moves']:
                old, new = move['old'], move['new']
                replacements[old] = new
                replacements[old[len('Certificates/'):]] = new[len('Certificates/'):]
                replacements[(root / old).as_posix()] = (root / new).as_posix()
                replacements['/api/certificates/' + quote(old[len('Certificates/'):], safe='/')] = '/api/certificates/' + quote(new[len('Certificates/'):], safe='/')
            for path, value in documents.items():
                updated = _rewrite(value, replacements)
                if updated != value:
                    if path.read_bytes() != original_bytes[path]:
                        raise ValueError('Certificate details changed during the folder move. Reopen the app to finish safely.')
                    save_json(path, updated)
            ledger['state'] = 'complete'
            ledger['pending'] = []
            save_json(_ledger_path(root), ledger)
        if legacy.is_dir():
            for folder in sorted((p for p in legacy.rglob('*') if p.is_dir()), key=lambda p: len(p.parts), reverse=True):
                if folder.is_symlink() or not folder.resolve().is_relative_to(legacy.resolve()):
                    raise ValueError('The old certificate folder changed during the move. Files were kept safe.')
                folder.rmdir()
            legacy.rmdir()
