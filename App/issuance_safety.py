"""Durable PDF/source publication and immutable per-issuance history.

Call under run_lock. An interrupted issuance is resolved using source and PDF
hashes, not a potentially ambiguous COD/status match. Unknown edits are never
overwritten by recovery. Completed journal entries are retained as audit records.
"""
from datetime import datetime, timezone
from copy import deepcopy
from functools import lru_cache
import getpass
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import uuid

from pypdf import PdfReader
from file_fingerprints import file_signature, fingerprint
from workflow_support import save_json


LOOKUP_RECEIPT = object()


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


@lru_cache(maxsize=256)
def _cached_pdf_health(path, signature):
    if not valid_pdf(path) or file_signature(path) != signature:
        raise OSError('PDF is unreadable or changed during validation')
    return True


def current_pdf_is_healthy(path, expected_sha=None):
    """Check an existing PDF for the dashboard, reusing work for unchanged files."""
    try:
        path = Path(path)
        signature = file_signature(path)
        if expected_sha is not None and fingerprint(path)[0] != expected_sha:
            return False
        healthy = valid_pdf(path) if signature[-1] is None else _cached_pdf_health(path, signature)
        return healthy and file_signature(path) == signature
    except OSError:
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


@lru_cache(maxsize=4)
def _raw_records(stamp):
    records = []
    for name, signature in stamp:
        path = Path(name)
        data = json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(data, dict) or data.get('version') != 1:
            raise ValueError('The certificate save history is unreadable. Restore a backup before continuing.')
        records.append(data)
    return sorted(records, key=lambda row: (row.get('started_at', ''), row.get('id', '')))


@lru_cache(maxsize=4)
def _projected_records(root, stamp, links_json, deletion_stamp=()):
    from inward_serials import project_receipts
    records = project_receipts(root, deepcopy(_raw_records(stamp)), links=json.loads(links_json))
    # Ownership follows an explicitly converted PDF; historical files stay immutable.
    transfers = {identifier: row for row in records if row.get('state') == 'complete'
                 for identifier in (row.get('converted_from') or {}).get('receipt_ids', [])}
    for row in records:
        owner = transfers.get(row.get('id'))
        if owner and row.get('source_kind') == 'workbook':
            row.update(recorded_source_kind='workbook', recorded_key=row['key'],
                       source_kind='manual', key=owner['key'])
    deleted = {}
    for name, signature in deletion_stamp:
        removal = json.loads(Path(name).read_text(encoding='utf-8'))
        if removal.get('state') == 'complete':
            deleted.update({identifier:removal['id'] for identifier in removal.get('receipt_ids', [])})
    for row in records:
        if row.get('id') in deleted:
            row.update(state='deleted', deleted_by=deleted[row['id']])
    index = {(row.get('source_kind'), row.get('plant'),
              str(row.get('key')).strip().casefold() if row.get('source_kind') == 'workbook' else str(row.get('key'))): row
             for row in records if row.get('state') == 'complete'}
    return records, index


def _receipt_snapshot(root):
    from inward_serials import workspace_links
    root = Path(root).resolve()
    stamp = tuple((str(path), file_signature(path)) for path in sorted((root / 'Data' / 'Issuance History').glob('*.json')))
    links_json = json.dumps(workspace_links(root), sort_keys=True)
    deletions = tuple((str(path), file_signature(path)) for path in sorted((root/'Data'/'Entry Deletions').glob('*.json')))
    return _projected_records(str(root), stamp, links_json, deletions)


def audit_records(root):
    records = deepcopy(_receipt_snapshot(root)[0])
    from certificate_drafts import read_register
    from manual_ownership import record_plant
    try:
        local = {row['id']: row for row in read_register(root)['records']}
    except (OSError, ValueError):
        # Recovery/audit must still inspect immutable history when the current
        # register is damaged. Normal edit/creation paths validate it separately.
        local = {}
    for row in records:
        if row.get('source_kind') == 'manual':
            row['plant'] = record_plant(local.get(str(row.get('key')), {'details':row.get('after') or {}}))
    return records


def manual_receipts(root, records, *, history=None):
    """Latest committed issuance for each selected local record and its owner."""
    owners = {row['id']: row['plant'] for row in records}
    latest = {}
    if not owners:
        return latest
    history = audit_records(root) if history is None else history
    for entry in history:
        key = str(entry.get('key'))
        if (entry.get('source_kind') == 'manual' and entry.get('state') == 'complete'
                and key in owners and entry.get('plant') in (None, owners[key])):
            latest[key] = entry
    return latest


def latest_issuance(root, source_kind, plant, key):
    if source_kind == 'manual':
        return next((row for row in reversed(audit_records(root)) if row.get('source_kind') == 'manual'
                     and row.get('state') == 'complete' and str(row.get('key')) == str(key)
                     and (row.get('plant') is None or row.get('plant') == plant)), None)
    identity = lambda value: str(value).strip().casefold() if source_kind == 'workbook' else str(value)
    return deepcopy(_receipt_snapshot(root)[1].get((source_kind, plant, identity(key))))


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
                 warnings=(), override_reason='', old_paths=(), template_path=None,
                 companion_source=None, converted_from=None, external_upload=False, retain_existing=False):
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
        for item in dict.fromkeys(([] if retain_existing else [str(destination)]) + [str(path) for path in old_paths]):
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
            'voided_certificate': None,
            'template_sha256': file_digest(template_path) if template_path else None,
        }
        if converted_from:
            self.data['converted_from'] = deepcopy(converted_from)
        if external_upload:
            self.data['external_upload'] = True
            self.data['retain_existing'] = retain_existing
        if companion_source:
            path, before_payload, after_payload = companion_source
            path = Path(path).resolve()
            if path != self.root / 'Data' / 'Custom Certificates.json':
                raise ValueError('Invalid companion certificate register.')
            if before_payload is not None:
                _write_bytes(self.folder / 'companion-before.bin', before_payload)
            _write_bytes(self.folder / 'companion-after.bin', after_payload)
            self.data['companion_source'] = {'path': str(path),
                'before': digest_bytes(before_payload) if before_payload is not None else None,
                'after': digest_bytes(after_payload)}
        self._save()

    def _save(self):
        save_json(self.path, self.data)

    def expect_source(self, payload):
        self.data['source_expected'] = digest_bytes(payload)
        self._save()

    def publish(self):
        destination = _certificate_path(self.root, self.data['pdf_path'])
        stage = destination.with_name(f'.{self.id}.publish.pdf')
        try:
            if self.data.get('retain_existing'):
                if file_digest(destination) != self.data['pdf_sha256']:
                    raise ValueError('The existing PDF changed. Refresh and select the file again.')
            else:
                _write_bytes(stage, (self.folder / 'new.pdf').read_bytes())
                if self.data.get('external_upload'):
                    if os.name == 'nt':
                        # Windows rename fails if a file appeared at the destination.
                        os.rename(stage, destination)
                    else:
                        os.link(stage, destination)
                        stage.unlink()
                else:
                    os.replace(stage, destination)
        except OSError as exc:
            try:
                stage.unlink(missing_ok=True)
            except OSError:
                pass
            if not isinstance(exc, PermissionError):
                raise
            if destination.is_file():
                raise PermissionError(
                    f'Could not replace {destination.name}. Close the PDF in your viewer or File Explorer preview, '
                    'then try again. If it still fails, check write access to the PDF and its folder.') from exc
            raise PermissionError(
                f'Could not save {destination.name}. Check write access to the Certificates folder, then try again.') from exc
        self.data['state'] = 'published'
        self._save()

    def sync_companion(self, committed):
        companion = self.data.get('companion_source')
        if not companion:
            return
        path = Path(companion['path']).resolve()
        if path != self.root / 'Data' / 'Custom Certificates.json':
            raise ValueError('Invalid companion certificate recovery path.')
        actual = file_digest(path)
        target = companion['after' if committed else 'before']
        if actual == target:
            return
        if actual not in {companion['before'], companion['after']}:
            raise ValueError('The local certificate register changed during saving. Recovery copies were preserved.')
        if target is None:
            # Retain the register created by an uncommitted conversion for recovery.
            saved = self.folder / ('uncommitted-register-' + uuid.uuid4().hex + '.json')
            os.replace(path, saved)
        else:
            stage = path.with_name('.' + self.id + '.register.json')
            _replace_snapshot(self.folder / ('companion-after.bin' if committed else 'companion-before.bin'),
                              path, stage, target)

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
        self.sync_companion(True)
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
                try:
                    path.unlink()
                except PermissionError as exc:
                    raise PermissionError(
                        f'Could not archive {path.name}. Close the PDF in your viewer or File Explorer preview, '
                        'then refresh to finish. If it still fails, check write access to the PDF and its folder.') from exc
        from certificate_inventory import record_issuance
        record_issuance(self.root, destination, self.data['cod'], self.data['source_kind'],
                        self.data['plant'], self.data['key'], external_upload=bool(self.data.get('external_upload')))
        self.data.update(state='complete', completed_at=datetime.now(timezone.utc).isoformat())
        self._save()

    def rollback(self):
        self.sync_companion(False)
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
        elif current == self.data['pdf_sha256'] and not self.data.get('retain_existing'):
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
    from entry_deletion import recover_entry_deletions
    recovered = recover_entry_deletions(root)
    for data in audit_records(root):
        if data.get('state') in {'complete', 'rolled_back', 'deleted'}:
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
