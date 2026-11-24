"""Observe manually added PDFs without adopting their contents into Excel.

The folder is another numbering source, not an instruction source. A durable
local ledger remembers observed identifiers even if a file is later removed.
"""
from pathlib import Path
from io import BytesIO
import json
import re
import hashlib

from pypdf import PdfReader
from certificate_folder_migration import legacy_certificate_origin, migrated_certificate_origins
from workflow_support import save_json, load_settings

MAX_PDF_BYTES = 50 * 1024 * 1024
MAX_PAGES = 8
INSPECTION_VERSION = 6
_RECORDED_COD_MISMATCH = 'Printed COD differs from the recorded certificate number. Check this PDF.'


def ensure_certificate_folders(root):
    """Prepare the shared output folder and consolidate older layouts."""
    base = Path(root) / 'Certificates'
    base.mkdir(parents=True, exist_ok=True)
    from certificate_folder_migration import migrate_certificate_folder
    migrate_certificate_folder(root)
    return base


def certificate_destination(root, filename):
    """Choose the shared output path for a new or revised certificate."""
    name = str(filename or '')
    if not name or Path(name).name != name or any(char in name for char in '/\\:') or not name.lower().endswith('.pdf'):
        raise ValueError('Invalid certificate filename.')
    path = safe_certificate_path(root, name)
    if path.name.casefold() != name.casefold():
        raise ValueError('A certificate filename points to another file.')
    return path


def safe_certificate_path(root, relative_path, *, expected_sha=None):
    """Resolve a path beneath Certificates, rejecting traversal and escaping symlinks."""
    base = (Path(root) / 'Certificates').resolve()
    raw = str(relative_path or '').replace('\\', '/')
    parts = raw.split('/')
    if (not raw or any(part in {'', '.', '..'} or ':' in part
                       or any(ord(char) < 32 for char in part) for part in parts)
            or not raw.lower().endswith('.pdf')):
        raise ValueError('Invalid certificate path.')
    from certificate_folder_migration import moved_relative_path
    raw = moved_relative_path(root, raw, expected_sha)
    destination = base.joinpath(*raw.split('/')).resolve()
    if not destination.is_relative_to(base) or destination.suffix.casefold() != '.pdf':
        raise ValueError('Invalid certificate path.')
    return destination


def identity(value):
    value = str(value or '').strip()
    if value.isascii() and value.isdecimal():
        return str(int(value))
    return re.sub(r'[^a-z0-9._-]+', '_', value.casefold()).strip(' ._')


def filename_cod(path):
    stem = re.sub(r' \(\d+\)(?=_hdd$|$)', '', Path(path).stem, flags=re.I)
    legacy = re.match(r'^(?:cod|certificate|hdd)[ _-]*(\d+)(?=_|$)', stem, re.I)
    if legacy:
        return legacy[1]
    if '_' not in stem:
        return stem if re.search(r'\d$', stem) else None
    first = stem.split('_', 1)[0]
    if first.casefold() not in {'cod', 'certificate', 'hdd'} and re.fullmatch(r'[A-Za-z0-9.-]+', first):
        return first
    return None


def text_identifiers(text, *, number_prefix=''):
    """Read labelled values, never an arbitrary date/weight/serial number."""
    cods, references = set(), set()
    cod_label = r'(?:certificate|cod|hdd)[ \t]*(?:no\.?|number)[ \t]*(?:&[ \t]*date)?'
    ref_label = r'(?:reference|ref)[ \t]*(?:no\.?|number)'
    for label, found in ((cod_label, cods), (ref_label, references)):
        for match in re.finditer(rf'(?im)\b{label}[ \t]*[:\-]?[ \t]*([^\n\r]*)', text):
            tail = match[1].strip()
            if label == cod_label and number_prefix:
                prefix_pattern = re.sub(r'(?:\\ )+', '[ \\t]+', re.escape(number_prefix.lstrip()))
                prefixed = re.match(prefix_pattern + r'\d+(?![A-Za-z0-9._/\-])', tail, re.I)
                if prefixed:
                    found.add(prefixed[0])
                    continue
            # Never assume a value on the following extracted line belongs to
            # this label: PowerPoint can emit all labels before all values.
            # The layout extraction pass associates columns on the same row.
            token = re.match(r'[A-Za-z0-9][A-Za-z0-9._/\-]*', tail)
            if token:
                value = token[0].rstrip('.')
                if (value.casefold() not in {'date', 'weight', 'reference', 'manifest', 'model', 'site', 'certificate', 'no', 'number', 'na', 'n/a', 'none'}
                        and not re.fullmatch(r'\d{1,4}[./-]\d{1,2}[./-]\d{2,4}', value)):
                    found.add(value)
    return sorted(cods), sorted(references)


def inspect_pdf(path, *, pdf_data=None, number_prefix=''):
    return _inspect_pdf(path, pdf_data=pdf_data, number_prefix=number_prefix)[:3]


def _inspect_pdf(path, *, pdf_data=None, number_prefix='', certificate_pages=None, issued_cod=None,
                 assigned_upload_cod=None):
    named = filename_cod(path)
    cods, references = (set() if issued_cod else set(filter(None, [named]))), set()
    warnings = []
    attachment_cods = set()
    try:
        if pdf_data is None:
            with path.open('rb') as stream:
                if b'%PDF-' not in stream.read(1024):
                    raise ValueError('file is not a readable PDF')
            if path.stat().st_size > MAX_PDF_BYTES:
                raise ValueError('PDF exceeds the 50 MB inspection limit')
            reader = PdfReader(path)
        else:
            if b'%PDF-' not in pdf_data[:1024]:
                raise ValueError('file is not a readable PDF')
            if len(pdf_data) > MAX_PDF_BYTES:
                raise ValueError('PDF exceeds the 50 MB inspection limit')
            reader = PdfReader(BytesIO(pdf_data))
        if reader.is_encrypted:
            raise ValueError('PDF is password protected')
        if not reader.pages:
            raise ValueError('PDF has no pages')
        printed = set()
        readable = False
        for index, page in enumerate(reader.pages[:MAX_PAGES]):
            for mode in ('plain', 'layout'):
                text = page.extract_text(extraction_mode=mode) or ''
                attached = certificate_pages is not None and index >= certificate_pages
                if not attached:
                    readable = readable or bool(text.strip())
                found, refs = text_identifiers(text, number_prefix=number_prefix)
                (attachment_cods if attached else printed).update(found)
                references.update(refs)
        if assigned_upload_cod:
            return [assigned_upload_cod], sorted(references), [], []
        cods.update(printed)
        if issued_cod:
            cods.add(str(issued_cod))
        # COD prefixes may themselves contain underscores. The first filename
        # token is then only a fragment, not an additional conflicting COD.
        filename_matches = [value for value in printed if identity(value) != identity(named) and (
            path.stem.casefold() == identity(value)
            or path.stem.casefold().startswith(identity(value) + '_'))]
        if not issued_cod and named and len(filename_matches) == 1:
            cods.discard(named)
            named = filename_matches[0]
            cods.add(named)
        if issued_cod and printed and any(identity(value) != identity(issued_cod) for value in printed):
            warnings.append(_RECORDED_COD_MISMATCH)
        elif not issued_cod and named and printed and any(identity(value) != identity(named) for value in printed):
            warnings.append('Filename COD and printed COD differ; both numbers are treated as used. Check this PDF.')
        if len({identity(value) for value in printed}) > 1:
            warnings.append('More than one printed COD was found; all are treated as used. Check this PDF.')
        if not readable:
            warnings.append('PDF has no readable text (possibly a scan); only its filename COD can be checked. References need manual verification.')
        elif not cods:
            warnings.append(cod_recovery_message(path.name))
        if len(reader.pages) > MAX_PAGES:
            warnings.append(f'Only the first {MAX_PAGES} PDF pages were inspected; check any later certificates manually.')
    except Exception as exc:
        warnings.append(f'Could not read PDF details: {exc}. The filename COD is still treated as used; verify other numbers manually.')
    return sorted(cods), sorted(references), warnings, sorted(attachment_cods)


def _issued_certificate_facts(root, *, history=None):
    """Read COD and attachment boundaries only from committed PDF receipts.

    Do not infer this from a filename, editable draft or PDF metadata. Older
    issued PDFs already have the template and receipt details in their journal,
    so they can be read correctly without rewriting the certificate.
    """
    from issuance_safety import audit_records
    try:
        records = audit_records(root) if history is None else history
    except (OSError, ValueError):
        return {}  # Without evidence, inspect every page and retain warnings.
    facts_by_digest = {}
    for record in records:
        if record.get('state') not in {'complete', 'source_committed'}:
            continue
        if record.get('external_replacement') or record.get('assigned_upload_cod'):
            continue  # Its selected COD belongs to the recorded path, not every copy of the PDF.
        details = record.get('after')
        if not isinstance(details, dict):
            continue
        digest = record.get('pdf_sha256')
        if not isinstance(digest, str) or not re.fullmatch(r'[a-f0-9]{64}', digest):
            continue
        cod = str(record.get('cod') or '').strip()
        if not cod:
            continue
        fields = details.get('certificate_fields') or {}
        pages = None
        if isinstance(fields, dict) and re.fullmatch(r'[a-f0-9]{32}', str(fields.get('_annexure') or '')):
            kind = details.get('certificate_type', 'standard')
            if kind == 'destruction':
                from destruction_certificate import receipts
                try:
                    # Older issued PDFs could contain continuation pages. Their
                    # immutable receipt proves where the attachment begins.
                    pages = 1 + max(1, (len(receipts(fields.get('device_receipts'), maximum=100)) + 8) // 9)
                except ValueError:
                    continue
            elif kind in {'standard', 'sita', 'siemens'}:
                pages = 1
            else:
                continue
        facts_by_digest.setdefault(digest, set()).add((cod, pages))
    return {digest: next(iter(facts)) for digest, facts in facts_by_digest.items() if len(facts) == 1}


def _issued_certificate_pages(root):
    """Return page boundaries for PDFs proven by exact committed receipts."""
    return {digest: pages for digest, (_cod, pages) in _issued_certificate_facts(root).items()
            if pages is not None}


def scan_inventory(root, persist=False, *, force=False):
    root = Path(root)
    number_prefix = load_settings(root).get('certificate_number_prefix', '')
    ledger_path = root / 'Data' / 'Certificate Inventory.json'
    previous = {'version': 1, 'entries': []}
    if ledger_path.exists():
        try:
            previous = json.loads(ledger_path.read_text(encoding='utf-8'))
            if not isinstance(previous, dict) or not isinstance(previous.get('entries'), list):
                raise ValueError('invalid inventory structure')
            for entry in previous['entries']:
                if (not isinstance(entry, dict) or not isinstance(entry.get('path'), str)
                        or not all(isinstance(entry.get(key), list) and all(isinstance(v, str) for v in entry[key]) for key in ('cods', 'references', 'warnings'))):
                    raise ValueError('invalid inventory record')
        except (OSError, ValueError, TypeError, KeyError) as exc:
            raise ValueError(f'The certificate number inventory is unreadable. Restore its backup before allocating a COD: {exc}') from exc
    entries = {entry['path'].casefold(): dict(entry, present=False) for entry in previous['entries']}
    from issuance_safety import audit_records
    history = audit_records(root)
    issued = _issued_certificate_facts(root, history=history)
    # A replacement associates those bytes with the selected certificate at
    # this path. The same bytes may also be filed under a different COD.
    replacements = {(record['pdf_path'].replace('\\', '/').casefold(), record['pdf_sha256']): str(record['cod'])
                    for record in history if record.get('state') in {'complete', 'source_committed'}
                    and record.get('external_replacement') and record.get('pdf_path')
                    and record.get('pdf_sha256') and record.get('cod')}
    assigned_uploads = {(record['pdf_path'].replace('\\', '/').casefold(), record['pdf_sha256']): str(record['cod'])
                        for record in history if record.get('state') in {'complete', 'source_committed'}
                        and record.get('assigned_upload_cod') and record.get('pdf_path')
                        and record.get('pdf_sha256') and record.get('cod')}
    origins = {record['pdf_sha256']: 'external' if record.get('external_upload') else 'app'
               for record in history if record.get('state') in {'complete', 'source_committed'} and record.get('pdf_sha256')}
    migrated = migrated_certificate_origins(root)
    owners_by_digest = {record.get('pdf_sha256'):{'source_kind':record.get('source_kind'),
                        'plant':record.get('plant'), 'key':str(record.get('key')), 'cod':str(record.get('cod'))}
                        for record in history if record.get('state') == 'complete'
                        and record.get('pdf_sha256') and not record.get('external_replacement')
                        and not record.get('assigned_upload_cod')}
    # A byte-identical PDF may now be intentionally saved for more than one
    # inward. Prefer the issuance tied to this exact path; digest-only ownership
    # remains a fallback for older files that were moved outside the app.
    owners_by_path = {}
    for record in history:
        if record.get('state') != 'complete' or not record.get('pdf_path') or not record.get('pdf_sha256'):
            continue
        owner = {'source_kind':record.get('source_kind'), 'plant':record.get('plant'),
                 'key':str(record.get('key')), 'cod':str(record.get('cod'))}
        owners_by_path[str(record['pdf_path']).replace('\\', '/').casefold()] = (record['pdf_sha256'], owner)
    for entry in entries.values():
        path_owner = owners_by_path.get(entry['path'].replace('\\', '/').casefold())
        if path_owner and path_owner[0] == entry.get('sha256'):
            entry['managed_owner'] = path_owner[1]
        elif entry.get('sha256') in owners_by_digest:
            entry['managed_owner'] = owners_by_digest[entry['sha256']]
    folder = root / 'Certificates'
    if folder.is_dir():
        for path in sorted(folder.rglob('*')):
            if not path.is_file() or path.suffix.casefold() != '.pdf' or path.name.startswith('.'):
                continue
            relative = path.relative_to(root).as_posix()
            try:
                stat = path.stat()
            except FileNotFoundError:
                # Explorer may remove/move a PDF after the directory listing.
                # Its prior identifiers remain reserved by the ledger above.
                continue
            old = entries.get(relative.casefold())
            pdf_data = None
            try:
                from file_fingerprints import fingerprint
                digest, pdf_data = fingerprint(path, force=force, payload_limit=MAX_PDF_BYTES)
            except OSError:
                digest = None
            issued_cod, certificate_pages = issued.get(digest, (None, None))
            replacement_cod = replacements.get((relative.casefold(), digest))
            assigned_upload_cod = assigned_uploads.get((relative.casefold(), digest))
            if replacement_cod:
                issued_cod, certificate_pages = replacement_cod, None
            if (old and old.get('inspection_version') == INSPECTION_VERSION
                    and old.get('inspection_prefix') == number_prefix
                    and old.get('inspection_certificate_pages') == certificate_pages
                    and old.get('inspection_issued_cod') == issued_cod
                    and old.get('inspection_replacement_cod') == replacement_cod
                    and old.get('inspection_assigned_upload_cod') == assigned_upload_cod
                    and digest is not None and old.get('sha256') == digest and old.get('size') == stat.st_size):
                # A restored byte-identical PDF remains the same issuance even
                # though atomic rollback changed its filesystem timestamp.
                entry = dict(old, present=True, mtime_ns=stat.st_mtime_ns)
            else:
                # Parse the same bytes we hashed. A PDF replaced in Explorer
                # between separate reads must not mix one hash with another COD.
                if digest is not None and pdf_data is None and stat.st_size <= MAX_PDF_BYTES:
                    digest, pdf_data = fingerprint(path, force=True, payload_limit=MAX_PDF_BYTES)
                cods, refs, warnings, attachment_cods = _inspect_pdf(
                    path, pdf_data=pdf_data, number_prefix=number_prefix, certificate_pages=certificate_pages,
                    issued_cod=issued_cod, assigned_upload_cod=assigned_upload_cod)
                observed_cods = set(cods)
                previous_cods = set((old or {}).get('certificate_cods', (old or {}).get('cods', [])))
                previous_numbers = set((old or {}).get('cods', []))
                if old and digest is not None and old.get('sha256') == digest:
                    # An older inspection counted attached certificates as this
                    # PDF's own numbers. Correct that attribution, but reserve all
                    # observed numbers below for future number allocation.
                    previous_cods.difference_update(set(attachment_cods) - set(cods))
                    if issued_cod:
                        # A filename changed on an exact copy cannot assign a
                        # second COD to the byte-identical issued certificate.
                        filename_number = filename_cod(path)
                        if filename_number and identity(filename_number) != identity(issued_cod) and not any(
                                identity(value) == identity(filename_number) for value in cods):
                            previous_cods.discard(filename_number)
                            previous_numbers.discard(filename_number)
                # Corrections do not make a previously observed number free.
                if replacement_cod:
                    cods, previous_cods = [replacement_cod], set()
                if assigned_upload_cod:
                    cods, previous_cods, previous_numbers = [assigned_upload_cod], set(), set()
                entry = {'path': relative, 'cods': sorted(observed_cods | set(attachment_cods) | previous_numbers),
                         'certificate_cods': sorted(set(cods) | previous_cods),
                         'attachment_cods': attachment_cods,
                         'references': sorted(set(refs) | set((old or {}).get('references', []))),
                         'current_cods': cods, 'warnings': warnings, 'present': True,
                         'size': len(pdf_data) if pdf_data is not None else stat.st_size,
                         'mtime_ns': stat.st_mtime_ns, 'sha256': digest,
                         'inspection_version': INSPECTION_VERSION, 'inspection_prefix': number_prefix,
                         'inspection_certificate_pages': certificate_pages, 'inspection_issued_cod': issued_cod,
                         'inspection_replacement_cod': replacement_cod,
                         'inspection_assigned_upload_cod': assigned_upload_cod}
                if old and digest is not None and old.get('sha256') == digest and old.get('managed_owner'):
                    entry['managed_owner'] = old['managed_owner']
            if replacement_cod:
                # The operator chose this PDF for this record. Apply to cached
                # inspections too, but only while the committed bytes match.
                entry['warnings'] = [warning for warning in entry['warnings']
                                     if warning != _RECORDED_COD_MISMATCH]
            path_owner = owners_by_path.get(relative.replace('\\', '/').casefold())
            if path_owner and path_owner[0] == digest:
                entry['managed_owner'] = path_owner[1]
            elif digest in owners_by_digest:
                entry['managed_owner'] = owners_by_digest[digest]
            origin = (origins.get(digest) or migrated.get((relative.casefold(), digest))
                      or legacy_certificate_origin(relative))
            if not origin and entry.get('managed_owner'):
                origin = 'external' if entry.get('external') else 'app'
            # A loose PDF has unknown origin. Offer explicit linking, but do
            # not use it as automatic evidence for an issued Excel row.
            entry.update(external=origin == 'external', app_generated=origin == 'app',
                         linkable_external=not entry.get('managed_owner')
                         and (origin == 'external' or (not origin and path.parent == folder)))
            entries[relative.casefold()] = entry
    result = {'version': 1, 'entries': sorted(entries.values(), key=lambda entry: entry['path'].casefold())}
    if persist and result != previous and (result['entries'] or ledger_path.exists()):
        if ledger_path.exists():
            # Keep the last valid ledger recoverable, independent of PDF files.
            import shutil
            from datetime import datetime
            import uuid
            backups = root / 'Data' / 'Certificate Inventory Backups'
            backups.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ledger_path, backups / f'Inventory_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.json')
        save_json(ledger_path, result)
    result['warnings'] = [f"{entry['path']}: {warning}" for entry in result['entries'] if entry['present'] for warning in entry['warnings']]
    owners = {}
    for entry in result['entries']:
        if entry['present']:
            content_id = entry.get('sha256') or f"path:{entry['path'].casefold()}"
            for cod in {identity(value) for value in entry.get('current_cods', entry['cods'])}:
                owners.setdefault(cod, {}).setdefault(content_id, entry['path'])
    for cod, paths in owners.items():
        if len(paths) > 1:
            result['warnings'].append(f"COD {cod} appears in different PDFs ({', '.join(Path(path).name for path in list(paths.values())[:3])}). Verify which certificate is current before sending.")
    return result


def cod_conflicts(inventory, cod, allowed_paths=(), owned=False):
    allowed = {str(path).replace('\\', '/').casefold() for path in allowed_paths}
    allowed_hashes = {entry.get('sha256') for entry in inventory['entries']
                      if entry['path'].casefold() in allowed and entry.get('sha256')
                      and all(identity(value) == identity(cod) for value in entry.get('current_cods', entry['cods']))}
    return [entry for entry in inventory['entries']
            if any(identity(value) == identity(cod) for value in entry.get('certificate_cods', entry['cods']))
            and entry.get('sha256') not in allowed_hashes
            and not (entry['path'].casefold() in allowed
                     and all(identity(value) == identity(cod) for value in entry.get('current_cods', entry['cods'])))]


def unresolved_cod_entries(inventory):
    """Present PDFs with no identifiable COD cannot safely inform numbering."""
    return [entry for entry in inventory.get('entries', [])
            if entry.get('present') and not entry.get('current_cods', entry.get('cods', []))]


def cod_recovery_message(relative_path):
    """Give the user a safe way to make an unlabelled PDF countable."""
    name = Path(str(relative_path or '')).name or 'the PDF'
    return (f"The app couldn't find a COD in {name}. Open it and read its COD, then rename the PDF "
            "so the COD comes first (for example, 0472_Client.pdf) and click Refresh. "
            "If it isn't an issued certificate, move it outside the Certificates folder.")


def reference_warnings(inventory, reference, allowed_paths=(), owned=False):
    """Compatibility API: references can be reused without a notice."""
    return []


def issuance_paths(inventory, cod, source_kind, plant, key):
    """Only committed issuance evidence establishes ownership, not a typed COD."""
    record_key = lambda value: str(value).strip().casefold() if source_kind == 'workbook' else str(value)
    return [entry['path'] for entry in inventory['entries']
            if isinstance(entry.get('managed_owner'), dict)
            and entry['managed_owner'].get('source_kind') == source_kind
            and (entry['managed_owner'].get('plant') == plant
                 or source_kind == 'manual' and entry['managed_owner'].get('plant') is None)
            and record_key(entry['managed_owner'].get('key')) == record_key(key)
            and identity(entry['managed_owner'].get('cod')) == identity(cod)]


def record_issuance(root, path, cod, source_kind, plant, key, *, external_upload=False):
    inventory = scan_inventory(root, persist=True)
    relative = Path(path).relative_to(Path(root)).as_posix()
    entry = next(entry for entry in inventory['entries'] if entry['path'].casefold() == relative.casefold())
    entry['managed_owner'] = {'source_kind': source_kind, 'plant': plant, 'key': str(key), 'cod': str(cod)}
    entry.update(external=external_upload, app_generated=not external_upload, linkable_external=False)
    inventory.pop('warnings', None)
    save_json(Path(root) / 'Data' / 'Certificate Inventory.json', inventory)
    return inventory


def verify_publication(root, snapshot, cod, reference=None, allowed_paths=(), owned=False, allow_warnings=False):
    """Recheck Explorer/PowerPoint changes made while the PDF was rendering."""
    current = scan_inventory(root, force=True)
    unknown = unresolved_cod_entries(current)
    if unknown:
        messages = [cod_recovery_message(entry['path']) for entry in unknown[:3]]
        if len(unknown) > 3:
            messages.append(f'There are {len(unknown) - 3} more PDFs to check.')
        raise ValueError(' '.join(messages))
    if cod_conflicts(current, cod, allowed_paths, owned):
        raise ValueError(f'COD {cod} is now present in another certificate PDF. Nothing was overwritten. Refresh and review the new automatic COD.')
    allowed = {str(path).replace('\\', '/').casefold() for path in allowed_paths}
    old = {entry['path'].casefold(): entry for entry in snapshot['entries']}
    for entry in current['entries']:
        previous = old.get(entry['path'].casefold())
        if entry['present'] and entry['path'].casefold() in allowed and previous:
            # A copied/restored byte-identical PDF can have a new timestamp.
            # Actual content identity, not mtime alone, determines a change.
            unchanged = (entry.get('sha256') is not None
                         and entry.get('sha256') == previous.get('sha256')
                         and entry.get('size') == previous.get('size'))
            if not unchanged:
                raise ValueError('The existing certificate changed while this PDF was being prepared. Nothing was overwritten. Refresh and review it first.')
