"""Consolidate legacy certificate folders while keeping issued PDF links valid."""
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
from urllib.parse import quote
import uuid

from workflow_support import run_lock, save_json

LEGACY_ORIGINS = {
    'created in app': 'app', 'auto generated': 'app',
    'manually added': 'external', 'externally created': 'external',
    'generated externally': 'external', 'external certificates': 'external',
}


def legacy_certificate_origin(path):
    parts = str(path).replace('\\', '/').casefold().removeprefix('certificates/').split('/')
    return LEGACY_ORIGINS.get(parts[0]) if len(parts) > 1 else None


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
        old, new = (move[key].replace('\\', '/').split('/') for key in ('old', 'new'))
        if (not all(parts[0] == 'Certificates' and all(p not in {'', '.', '..'} and ':' not in p for p in parts)
                    for parts in (old, new))
                or len(old) < 3 or old[1].casefold() not in LEGACY_ORIGINS
                or not (len(new) == 2 or (len(new) >= 3 and old[1] == 'Created in App' and new[1] == 'Auto Generated'))):
            raise ValueError('The saved certificate folder move has an invalid path. Files were kept safe.')
    return value


def migrated_certificate_origins(root):
    return {(move['new'].casefold(), move['sha256']): legacy_certificate_origin(move['old'])
            for move in _ledger(root)['moves']}


def moved_relative_path(root, relative, expected_sha=None):
    """Resolve an old link only when an actual recorded move proves its target."""
    raw = str(relative).replace('\\', '/')
    if not legacy_certificate_origin(raw):
        return raw
    base = Path(root) / 'Certificates'
    if (base / raw).exists():
        return raw  # A restored backup may still be in its original location.
    moves = {move['old'].casefold(): move['new'][len('Certificates/'):]
             for move in _ledger(root)['moves'] if expected_sha is None or move['sha256'] == expected_sha}
    while not (base / raw).exists():
        target = moves.get(f'Certificates/{raw}'.casefold())
        if target is None:
            break
        raw = target
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
    base = root / 'Certificates'
    if not base.is_dir():
        return
    if base.resolve() != base:
        raise ValueError('The Certificates folder points outside the workspace.')
    if not any(path.name.casefold() in LEGACY_ORIGINS and path.is_dir() for path in base.iterdir()):
        if _ledger(root).get('state') == 'complete':
            return
    with run_lock(root):
        from issuance_safety import recover_issuances
        recover_issuances(root)
        ledger = _ledger(root)
        files, directories = [], []
        for folder in sorted(path for path in base.iterdir()
                             if path.name.casefold() in LEGACY_ORIGINS and path.is_dir()):
            if folder.resolve() != folder:
                raise ValueError('An old certificate folder is a link. Files were kept safe.')
            items = sorted(folder.rglob('*'))
            if any(path.resolve() != path for path in items):
                raise ValueError('An old certificate folder contains a link. Files were kept safe.')
            files.extend(path for path in items if path.is_file())
            directories.extend(path for path in items if path.is_dir())
            directories.append(folder)
        unfinished = ledger.get('state') == 'moving'
        pending = ledger.get('pending', []) if unfinished else []
        if not isinstance(pending, list) or any(type(i) is not int or not 0 <= i < len(ledger['moves']) for i in pending):
            raise ValueError('The saved certificate folder move could not be read. Files were kept safe.')
        planned_sources = {ledger['moves'][i]['old'].casefold() for i in pending}
        # Finish an older pending Created in App → Auto Generated move first.
        # Its destination is flattened in the next pass under a new journal entry.
        nested_targets = {ledger['moves'][i]['new'].casefold() for i in pending
                          if len(Path(ledger['moves'][i]['new']).parts) > 2}
        planned_sources.update(nested_targets)
        inventory_path = root / 'Data' / 'Certificate Inventory.json'
        active = [inventory_path, root / 'Data' / 'Custom Certificates.json',
                  root / 'Data' / 'Reports' / 'Latest Results.json']
        # Read the active metadata before planning any move.
        original_bytes = {path: path.read_bytes() for path in active if path.is_file()}
        documents = {path: json.loads(payload.decode('utf-8-sig')) for path, payload in original_bytes.items()}
        reserved = {str(path).casefold() for path in base.iterdir()}
        reserved.update(str(root / move['new']).casefold() for move in ledger['moves'])
        reserved.update(str(root / entry['path']).casefold()
                        for entry in documents.get(inventory_path, {}).get('entries', []))
        additions = []
        for path in files:
            old = path.relative_to(root).as_posix()
            if old.casefold() in planned_sources:
                continue
            target = _available(base / path.name, reserved)
            reserved.add(str(target).casefold())
            additions.append({'old': old, 'new': target.relative_to(root).as_posix(), 'sha256': _digest(path)})
        if additions:
            backup = root / 'Data' / 'Certificate Folder Backups' / f'{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}'
            existing_ledger = [_ledger_path(root)] if _ledger_path(root).exists() else []
            for path in list(documents) + existing_ledger + [root / move['old'] for move in additions]:
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
                if any(path.resolve() != path or not path.is_relative_to(base) for path in (source, destination)):
                    raise ValueError('A certificate folder points outside Certificates. Files were kept safe.')
                if source.exists():
                    if _digest(source) != move['sha256']:
                        raise ValueError('A certificate changed during its folder move. Files were kept safe; reopen the app to check it.')
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    try:
                        if os.name == 'nt':
                            os.rename(source, destination)
                        else:
                            os.link(source, destination)
                            source.unlink()
                    except FileExistsError:
                        if _digest(destination) != move['sha256']:
                            raise ValueError('Another file appeared during the move. Both files were kept safe.')
                        source.unlink()  # The verified target and backup both exist.
                if not destination.is_file() or _digest(destination) != move['sha256']:
                    raise ValueError('A moved certificate could not be checked. Its backup was kept safe.')
            replacements = {}
            targets = {move['old']: move['new'] for move in ledger['moves']}
            for old, new in targets.items():
                new = targets.get(new, new)
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
        if not nested_targets:
            for folder in sorted(directories, key=lambda path: len(path.parts), reverse=True):
                if folder.resolve() != folder or not folder.is_relative_to(base):
                    raise ValueError('An old certificate folder changed during the move. Files were kept safe.')
                folder.rmdir()
    if nested_targets:
        migrate_certificate_folder(root)
