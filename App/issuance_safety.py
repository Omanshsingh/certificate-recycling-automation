"""Durable PDF/source publication and immutable per-issuance history.

Call under run_lock. An interrupted issuance is resolved using source and PDF
hashes, not a potentially ambiguous COD/status match. Unknown edits are never
overwritten by recovery. Completed journal entries are retained as audit records.
"""
from datetime import datetime, timezone
import getpass
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import uuid

from pypdf import PdfReader
from workflow_support import save_json


def digest_bytes(data):
    return hashlib.sha256(data).hexdigest()


def file_digest(path):
    path = Path(path)
    if not path.is_file():
        return None
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def valid_pdf(path):
    """Structural health, not a claim that the certificate data is correct."""
    try:
        path = Path(path)
        if not path.is_file() or not 0 < path.stat().st_size <= 50 * 1024 * 1024:
            return False
        with path.open('rb') as stream:
            if not stream.read(1024).startswith(b'%PDF-'):
                return False
            reader = PdfReader(stream)
            if reader.is_encrypted or not 0 < len(reader.pages) <= 2000:
                return False
            for page in reader.pages:
                if page.mediabox.width <= 0 or page.mediabox.height <= 0:
                    return False
                contents = page.get_contents()
                if contents:
                    contents.get_data()
        return True
    except Exception:
        return False


def _write_bytes(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('xb') as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())


def _replace_snapshot(snapshot, destination, stage, expected):
    """Retry an interrupted stage safely; never trust a partial staged copy."""
    payload = Path(snapshot).read_bytes()
    if digest_bytes(payload) != expected:
        raise ValueError('The certificate recovery copy failed verification. Files were preserved; ask IT to review them.')
    if stage.exists() and file_digest(stage) != expected:
        os.replace(stage, stage.with_name(stage.name + '.interrupted-' + uuid.uuid4().hex))
    if not stage.exists():
        _write_bytes(stage, payload)
    os.replace(stage, destination)


def _certificate_path(root, value):
    root = Path(root).resolve()
    path = (root / value).resolve()
    if not path.is_relative_to(root / 'Certificates') or path.suffix.casefold() != '.pdf':
        raise ValueError('Invalid certificate recovery path. Ask IT to review the save history.')
    return path


def audit_records(root):
    folder = Path(root) / 'Data' / 'Issuance History'
    records = []
    for path in sorted(folder.glob('*.json')):
        data = json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(data, dict) or data.get('version') != 1:
            raise ValueError('The certificate save history is unreadable. Restore a backup before continuing.')
        records.append(data)
    return sorted(records, key=lambda row: (row.get('started_at', ''), row.get('id', '')))


def latest_issuance(root, source_kind, plant, key):
    identity = lambda value: str(value).strip().casefold() if source_kind == 'workbook' else str(value)
    return next((row for row in reversed(audit_records(root))
                 if row.get('state') == 'complete' and row.get('source_kind') == source_kind
                 and row.get('plant') == plant and identity(row.get('key')) == identity(key)), None)


def workbook_receipt_conflict(batch, receipt):
    """Reject reused Serial Numbers before restoring another inward's history."""
    if not receipt:
        return ''
    from certificate_core import parse_date

    def cod_key(value):
        value = str(value or '').strip()
        return str(int(value)) if value.isascii() and value.isdecimal() else value.casefold()

    current_cod = cod_key(batch.get('cod_no', batch.get('certificate_no')))
    previous_cod = cod_key(receipt.get('cod'))
    if current_cod and previous_cod and current_cod == previous_cod:
        return ''

    def source_facts(record):
        if not isinstance(record, dict) or not record.get('supplier') or not record.get('items'):
            return None
        received = parse_date(record.get('inward_date'))
        if received is None:
            return None
        facts = {key: record.get(key) for key in ('supplier', 'location', 'form6', 'reference', 'items')}
        facts['inward_date'] = received.isoformat()
        return json.dumps(facts, sort_keys=True, default=str)

    facts = source_facts(batch)
    if not current_cod and facts is not None and any(
            facts == source_facts(receipt.get(key)) for key in ('before', 'after')):
        return ''
    serial = batch.get('case_id', batch.get('sn', ''))
    return (f'Serial Number {serial} is linked to an earlier certificate, but the current Excel '
            'record does not match that history. Open Excel and check the Serial Number and COD. '
            'Give a new inward an unused Serial Number, or restore the correct identifiers for '
            'the existing inward, then save and refresh.')


class IssuanceTransaction:
    def __init__(self, root, destination, temporary, source_path, source_before,
                 *, source_kind, plant, key, cod, issued_date, before, after,
                 warnings=(), override_reason='', old_paths=(), template_path=None):
        self.root = Path(root).resolve()
        self.id = uuid.uuid4().hex
        self.folder = self.root / 'Data' / 'Issuance Recovery' / self.id
        self.path = self.root / 'Data' / 'Issuance History' / f'{self.id}.json'
        self.folder.mkdir(parents=True, exist_ok=True)
        if source_before is not None:
            _write_bytes(self.folder / 'source-before.bin', source_before)
        destination, temporary = Path(destination), Path(temporary)
        if not valid_pdf(temporary):
            raise ValueError('The new PDF could not be validated. Nothing was saved.')
        _write_bytes(self.folder / 'new.pdf', temporary.read_bytes())
        old = []
        for item in dict.fromkeys([str(destination)] + [str(path) for path in old_paths]):
            path = Path(item)
            if path.is_file():
                relative = path.resolve().relative_to(self.root).as_posix()
                _certificate_path(self.root, relative)
                backup = f'previous-{len(old)}.pdf'
                payload = path.read_bytes()
                _write_bytes(self.folder / backup, payload)
                old.append({'path': relative, 'snapshot': backup, 'sha256': digest_bytes(payload),
                            'archive': f'Data/Archive/Superseded Certificates/{self.id}/{len(old)}/{path.name}'})
        self.data = {
            'version': 1, 'id': self.id, 'state': 'prepared',
            'started_at': datetime.now(timezone.utc).isoformat(), 'operator': getpass.getuser(),
            'source_kind': source_kind, 'plant': plant, 'key': str(key), 'cod': str(cod),
            'certificate_date': str(issued_date), 'source_path': str(Path(source_path).resolve()),
            'source_before': digest_bytes(source_before) if source_before is not None else None,
            'source_expected': None, 'pdf_path': destination.resolve().relative_to(self.root).as_posix(),
            'pdf_sha256': file_digest(self.folder / 'new.pdf'), 'previous_pdfs': old,
            'before': before, 'after': after, 'accepted_warnings': list(warnings),
            'override_reason': str(override_reason or '').strip(),
            'template_sha256': file_digest(template_path) if template_path else None,
        }
        self._save()

    def _save(self):
        save_json(self.path, self.data)

    def expect_source(self, payload):
        self.data['source_expected'] = digest_bytes(payload)
        self._save()

    def publish(self):
        destination = _certificate_path(self.root, self.data['pdf_path'])
        stage = destination.with_name(f'.{self.id}.publish.pdf')
        _write_bytes(stage, (self.folder / 'new.pdf').read_bytes())
        os.replace(stage, destination)
        self.data['state'] = 'published'
        self._save()

    def mark_source_committed(self):
        self.data['state'] = 'source_committed'
        self._save()

    def complete(self):
        self.mark_source_committed()
        self._finish_committed()

    def _archive_destination(self, old, index, original):
        relative = old.get('archive')
        if not isinstance(relative, str):
            raise ValueError('Invalid certificate archive path.')
        legacy = f'Data/Archive/Superseded Certificates/{self.id}_{original.name}'
        if relative.replace('\\', '/') == legacy and len((self.id + '_' + original.name).encode('utf-16-le')) // 2 > 255:
            snapshot = old.get('snapshot')
            expected = old.get('sha256')
            if (not isinstance(snapshot, str) or not re.fullmatch(r'previous-[0-9]+\.pdf', snapshot)
                    or not isinstance(expected, str) or not re.fullmatch(r'[0-9a-f]{64}', expected)
                    or file_digest(self.folder / snapshot) != expected):
                raise ValueError('The previous PDF recovery copy failed verification. Files were preserved.')
            # Only the known, impossible legacy filename is migrated. Journal
            # the verified new destination before creating or moving a file.
            old['archive'] = f'Data/Archive/Superseded Certificates/{self.id}/{index}/{original.name}'
            self._save()
        archive = (self.root / old['archive']).resolve()
        if not archive.is_relative_to(self.root / 'Data' / 'Archive' / 'Superseded Certificates'):
            raise ValueError('Invalid certificate archive path.')
        return archive

    def _finish_committed(self):
        destination = _certificate_path(self.root, self.data['pdf_path'])
        if not destination.exists():
            # Completed receipts are never auto-restored. This is only for an
            # incomplete save whose source commit is proved by its hash.
            stage = destination.with_name(f'.{self.id}.recover.pdf')
            _replace_snapshot(self.folder / 'new.pdf', destination, stage, self.data['pdf_sha256'])
        # Do not hide a replacement made outside this app during recovery.
        if file_digest(destination) != self.data['pdf_sha256']:
            raise ValueError('The issued PDF changed during recovery. Its saved copy is retained; ask IT to review it.')
        for index, old in enumerate(self.data['previous_pdfs']):
            path = _certificate_path(self.root, old['path'])
            archive = self._archive_destination(old, index, path)
            archive.parent.mkdir(parents=True, exist_ok=True)
            if not archive.exists():
                _write_bytes(archive, (self.folder / old['snapshot']).read_bytes())
            elif file_digest(archive) != old['sha256']:
                # A crash during archive creation can leave a partial file.
                # Keep that file for diagnosis, then atomically publish the
                # verified original snapshot instead of trusting its presence.
                preserved = archive.with_name('.interrupted-' + uuid.uuid4().hex + '.pdf')
                os.replace(archive, preserved)
                _write_bytes(archive, (self.folder / old['snapshot']).read_bytes())
            if file_digest(archive) != old['sha256']:
                raise ValueError('The previous PDF archive could not be verified. The current and recovery copies were preserved.')
            if path != destination and path.exists():
                if file_digest(path) != old['sha256']:
                    raise ValueError('An older PDF changed outside this app. Both versions were preserved; ask IT to review them.')
                # Recoverable retirement: move the old file only after source commit.
                path.unlink()
        from certificate_inventory import record_issuance
        record_issuance(self.root, destination, self.data['cod'], self.data['source_kind'],
                        self.data['plant'], self.data['key'])
        self.data.update(state='complete', completed_at=datetime.now(timezone.utc).isoformat())
        self._save()

    def rollback(self):
        destination = _certificate_path(self.root, self.data['pdf_path'])
        current = file_digest(destination)
        previous = next((old for old in self.data['previous_pdfs'] if old['path'] == self.data['pdf_path']), None)
        allowed = {None, self.data['pdf_sha256'], previous['sha256'] if previous else None}
        if current not in allowed:
            raise ValueError('The PDF changed outside this app. Recovery did not overwrite it; ask IT to review the saved copies.')
        if previous:
            if current != previous['sha256']:
                stage = destination.with_name(f'.{self.id}.restore.pdf')
                _replace_snapshot(self.folder / previous['snapshot'], destination, stage, previous['sha256'])
        elif current == self.data['pdf_sha256']:
            # Quarantine, not permanent deletion: retains the uncommitted output.
            quarantine = self.root / 'Data' / 'Archive' / 'Uncommitted Certificates' / self.id
            quarantine.mkdir(parents=True, exist_ok=True)
            os.replace(destination, quarantine / destination.name)
        self.data.update(state='rolled_back', completed_at=datetime.now(timezone.utc).isoformat())
        self._save()


def recover_issuances(root):
    """Under run_lock, before reading/editing sources or allocating a new COD."""
    root = Path(root).resolve()
    from backup_service import check_restore_health
    check_restore_health(root)
    from workflow_support import load_settings
    configured = {str(Path(value).resolve()) for value in load_settings(root).get('workbook_paths', []) if value}
    recovered = []
    for data in audit_records(root):
        if data.get('state') in {'complete', 'rolled_back'}:
            continue
        identifier = data.get('id', '')
        if not isinstance(identifier, str) or not __import__('re').fullmatch(r'[a-f0-9]{32}', identifier):
            raise ValueError('Invalid certificate recovery identifier.')
        source = Path(data['source_path']).resolve()
        if str(source) not in configured and source != root / 'Data' / 'Custom Certificates.json':
            raise ValueError('A pending certificate uses another workbook. Reconnect its original workbook before recovery.')
        tx = object.__new__(IssuanceTransaction)
        tx.root, tx.id, tx.data = root, identifier, data
        tx.folder = root / 'Data' / 'Issuance Recovery' / identifier
        tx.path = root / 'Data' / 'Issuance History' / f'{identifier}.json'
        actual = file_digest(source)
        if data.get('source_expected') is not None and actual == data['source_expected']:
            tx._finish_committed()
            recovered.append(f"Completed interrupted certificate {data['cod']}.")
        elif actual == data.get('source_before'):
            tx.rollback()
            recovered.append(f"Rolled back interrupted certificate {data['cod']}; the original source was unchanged.")
        else:
            raise ValueError('A source changed during an interrupted save. Saved copies are preserved. Restore a backup or ask IT before creating another certificate.')
    return recovered
