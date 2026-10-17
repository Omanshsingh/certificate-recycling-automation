"""Retire obsolete app history while preserving the operator's current files."""
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import uuid

from certificate_core import read_batches
from certificate_drafts import read_register, current_record_pdf
from certificate_inventory import scan_inventory
from issuance_safety import audit_records, file_digest
from workflow_support import load_settings, run_lock, save_json
from workspace_review import require_revision


def _plan(root):
    settings = load_settings(root)
    connections = [(index, Path(value).expanduser().resolve())
                   for index, value in enumerate(settings.get('workbook_paths', []), 1) if value.strip()]
    paths = [path for _, path in connections]
    if len(paths) == 2 and paths[0] == paths[1]:
        raise ValueError('Plant 1 and Plant 2 must use different workbooks.')
    connected_plants = {f'plant{index}' for index, _ in connections}
    batches = [dict(row, source_id=f'plant{index}')
               for index, path in connections for row in read_batches(path)]
    histories = audit_records(root)
    if any(row.get('state') not in {'complete', 'rolled_back'} for row in histories):
        raise ValueError('A certificate save is still pending. Reopen the app to finish recovery before clearing history.')
    register = read_register(root)
    # The current PDF must actually exist; a remembered filename is insufficient.
    records = [row for row in register['records'] if current_record_pdf(root, row)]
    manual_keys = {(row['plant'], row['id']) for row in records}
    current = {(row['source_id'], row['case_id'].strip().casefold()): row for row in batches
               if not row.get('duplicate_serial_number')}
    pdfs = sorted(path for path in (root / 'Certificates').rglob('*')
                  if path.is_file() and path.suffix.casefold() == '.pdf' and not path.name.startswith('.'))
    pdf_hashes = {}
    for path in pdfs:
        if not path.resolve().is_relative_to(root / 'Certificates'):
            raise ValueError('A certificate file points outside its folder. Move it into Certificates before clearing history.')
        pdf_hashes[path.relative_to(root).as_posix().casefold()] = file_digest(path)
    kept = {}
    for receipt in histories:
        if receipt.get('state') != 'complete' or not re.fullmatch(r'[a-f0-9]{32}', str(receipt.get('id', ''))):
            continue
        digest = pdf_hashes.get(str(receipt.get('pdf_path', '')).casefold())
        if not digest or digest != receipt.get('pdf_sha256'):
            continue
        kind, plant, key = (receipt.get(name) for name in ('source_kind', 'plant', 'key'))
        if kind == 'workbook' and plant in connected_plants:
            row = current.get((plant, str(key).strip().casefold()))
            previous = receipt.get('after') or receipt.get('before') or {}
            cod_key = lambda v: str(int(str(v))) if str(v).isascii() and str(v).isdecimal() else str(v or '').strip().casefold()
            if (not row or not row.get('cod_no')
                    or cod_key(row['cod_no']) != cod_key(receipt.get('cod'))
                    or str(row['supplier']).strip().casefold() != str(previous.get('supplier', '')).strip().casefold()):
                continue
        elif kind == 'manual' and (plant, key) not in manual_keys:
            continue
        elif kind not in {'manual', 'workbook'}:
            continue
        kept[(kind, plant, str(key).strip().casefold())] = receipt
    inventory_path = root / 'Data' / 'Certificate Inventory.json'
    old_inventory = json.loads(inventory_path.read_text(encoding='utf-8')) if inventory_path.exists() else {}
    removed_links = sum(str(entry.get('path', '')).casefold() not in pdf_hashes
                        for entry in old_inventory.get('entries', []))
    public = {'history_removed': len(histories) - len(kept),
              'pdf_links_removed': removed_links, 'local_records_removed': len(register['records']) - len(records),
              'history_kept': len(kept), 'excel_records_kept': len(batches), 'pdfs_kept': len(pdfs)}
    fingerprint = hashlib.sha256(json.dumps([histories, register, old_inventory],
                                sort_keys=True, default=str).encode('utf-8')).hexdigest()
    files = {str(path): file_digest(path) for path in paths + pdfs}
    attachments = set()
    for row in batches + [row.get('details', {}) for row in records] + [row.get('after', {}) for row in kept.values()]:
        value = (row.get('certificate_fields') or {}).get('_annexure')
        if value and re.fullmatch(r'[a-f0-9]{32}', str(value)):
            attachments.add(str(value))
    return dict(public=public, fingerprint=fingerprint, settings=settings, files=files,
                receipts=list(kept.values()), records=records, attachments=attachments)


def _inside(path, parent):
    resolved = path.resolve()
    if resolved == parent.resolve() or not resolved.is_relative_to(parent.resolve()):
        raise ValueError('The history cleanup path is outside its allowed folder.')
    return resolved


def clear_old_history(root, *, confirmed=False, expected_revision=None):
    root = Path(root).resolve()
    with run_lock(root):
        require_revision(root, expected_revision)
        plan = _plan(root)
        if not confirmed:
            return {'requires_confirmation': True, 'cleanup': plan['public'],
                    'warnings': [dict(plan['public'], history_revision=plan['fingerprint'])]}
        data = root / 'Data'
        if data.resolve() != data:
            raise ValueError('The app data folder is a linked location. Move it into this app before clearing history.')
        original_names = {path.name for path in data.iterdir()}
        candidates = [path for path in data.iterdir() if path.name not in {'run.lock', 'Backups'}]
        # Validate every absolute source before any directory move on Windows.
        for path in candidates:
            _inside(path, data)
            if path.is_symlink() or getattr(path.lstat(), 'st_file_attributes', 0) & 0x400:
                raise ValueError('A history folder is a linked location. Move it into this app before clearing history.')
        store = Path(os.environ.get('LOCALAPPDATA', str(Path.home() / '.local' / 'share'))).resolve() / 'Ecoreco' / 'History Backups'
        if store.resolve().is_relative_to(root):
            raise ValueError('The history recovery folder must be outside the app.')
        backup = store / (datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8])
        _inside(backup, store)
        backup.mkdir(parents=True, exist_ok=False)
        previous = backup / 'Previous Data'
        previous.mkdir()
        settings_path = root / 'settings.local.json'
        settings_before = settings_path.read_bytes() if settings_path.exists() else None
        if settings_before is not None:
            (backup / 'settings.local.json').write_bytes(settings_before)
        save_json(backup / 'manifest.json', {'version': 1, 'root': str(root), 'cleanup': plan['public'],
                                          'preserved_files': plan['files']})
        moved = []
        try:
            # Nothing in Excel or Certificates is moved or rewritten.
            for path in candidates:
                shutil.move(str(path), str(previous / path.name))
                moved.append(path.name)
            for receipt in plan['receipts']:
                save_json(data / 'Issuance History' / (receipt['id'] + '.json'), receipt)
            save_json(data / 'Custom Certificates.json', {'version': 1, 'records': plan['records']})
            save_json(data / 'History Reset.json', {'id': uuid.uuid4().hex,
                      'cleared_at': datetime.now().isoformat(timespec='seconds')})
            for identifier in plan['attachments']:
                for suffix in ('.pdf', '.json'):
                    source = previous / 'Attachments' / (identifier + suffix)
                    if source.is_file():
                        target = data / 'Attachments' / source.name
                        target.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(source, target)
            settings = dict(plan['settings'], certificate_number_start=1)
            save_json(settings_path, settings)
            fresh = scan_inventory(root)
            owners = {r['pdf_path'].casefold(): r for r in plan['receipts']}
            for entry in fresh['entries']:
                receipt = owners.get(entry['path'].casefold())
                if receipt and receipt['pdf_sha256'] == entry.get('sha256'):
                    entry['managed_owner'] = {key: receipt.get(key) for key in ('source_kind', 'plant', 'key', 'cod')}
            fresh.pop('warnings', None)
            save_json(data / 'Certificate Inventory.json', fresh)
            for path, digest in plan['files'].items():
                if file_digest(Path(path)) != digest:
                    raise ValueError('An Excel file or PDF changed during cleanup. The old history will be restored; try again.')
        except Exception:
            # Preserve failed staging as well; rollback never recursively deletes.
            failed = backup / 'Incomplete cleanup'
            for name in moved:
                destination = _inside(data / name, data)
                if destination.exists():
                    failed.mkdir(exist_ok=True)
                    shutil.move(str(destination), str(failed / name))
                shutil.move(str(previous / name), str(destination))
            # New files whose names were absent before must not keep test state alive.
            for path in list(data.iterdir()):
                if path.name not in original_names:
                    _inside(path, data)
                    failed.mkdir(exist_ok=True)
                    shutil.move(str(path), str(failed / path.name))
            if settings_before is not None:
                settings_path.write_bytes(settings_before)
            elif settings_path.exists():
                failed.mkdir(exist_ok=True)
                shutil.move(str(settings_path), str(failed / 'settings.local.json'))
            raise
        return {'requires_confirmation': False, 'cleared': True, 'cleanup': plan['public'],
                'recovery_folder': str(backup), 'message': 'Old history cleared. Current Excel files and PDFs were kept.'}
