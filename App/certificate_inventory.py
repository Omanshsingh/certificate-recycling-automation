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
from workflow_support import save_json, load_settings

MAX_PDF_BYTES = 50 * 1024 * 1024
MAX_PAGES = 8
INSPECTION_VERSION = 3
CERTIFICATE_FOLDERS = {
    'workbook': 'Auto Generated',
    'manual': 'Created in App',
    'external': 'Manually Added',
}


def ensure_certificate_folders(root):
    """Create the three source folders while keeping the parent as the watched root."""
    base = Path(root) / 'Certificates'
    base.mkdir(parents=True, exist_ok=True)
    folders = {kind: base / name for kind, name in CERTIFICATE_FOLDERS.items()}
    for folder in folders.values():
        folder.mkdir(parents=True, exist_ok=True)
    return folders


def certificate_destination(root, source_kind, filename):
    """Choose an output folder for an app-issued PDF; external PDFs are never issued here."""
    if source_kind not in {'workbook', 'manual'}:
        raise ValueError('Choose the workbook or in-app certificate folder.')
    name = str(filename or '')
    if not name or Path(name).name != name or any(char in name for char in '/\\:') or not name.lower().endswith('.pdf'):
        raise ValueError('Invalid certificate filename.')
    return Path(root) / 'Certificates' / CERTIFICATE_FOLDERS[source_kind] / name


def safe_certificate_path(root, relative_path):
    """Resolve a path beneath Certificates, rejecting traversal and escaping symlinks."""
    base = (Path(root) / 'Certificates').resolve()
    raw = str(relative_path or '').replace('\\', '/')
    parts = raw.split('/')
    if (not raw or any(part in {'', '.', '..'} or ':' in part
                       or any(ord(char) < 32 for char in part) for part in parts)
            or not raw.lower().endswith('.pdf')):
        raise ValueError('Invalid certificate path.')
    destination = base.joinpath(*parts).resolve()
    if not destination.is_relative_to(base) or destination.suffix.casefold() != '.pdf':
        raise ValueError('Invalid certificate path.')
    return destination


def identity(value):
    value = str(value or '').strip()
    if value.isascii() and value.isdecimal():
        return str(int(value))
    return re.sub(r'[^a-z0-9._-]+', '_', value.casefold()).strip(' ._')


def filename_cod(path):
    stem = Path(path).stem
    legacy = re.match(r'^(?:cod|certificate)[ _-]*(\d+)(?=_|$)', stem, re.I)
    if legacy:
        return legacy[1]
    if '_' not in stem:
        return stem if re.search(r'\d$', stem) else None
    first = stem.split('_', 1)[0]
    if first.casefold() not in {'cod', 'certificate'} and re.fullmatch(r'[A-Za-z0-9.-]+', first):
        return first
    return None


def text_identifiers(text, *, number_prefix=''):
    """Read labelled values, never an arbitrary date/weight/serial number."""
    cods, references = set(), set()
    cod_label = r'(?:certificate|cod)[ \t]*(?:no\.?|number)[ \t]*(?:&[ \t]*date)?'
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
    named = filename_cod(path)
    cods, references = set(filter(None, [named])), set()
    warnings = []
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
        for page in reader.pages[:MAX_PAGES]:
            for mode in ('plain', 'layout'):
                text = page.extract_text(extraction_mode=mode) or ''
                readable = readable or bool(text.strip())
                found, refs = text_identifiers(text, number_prefix=number_prefix)
                printed.update(found)
                references.update(refs)
        cods.update(printed)
        # COD prefixes may themselves contain underscores. The first filename
        # token is then only a fragment, not an additional conflicting COD.
        filename_matches = [value for value in printed if identity(value) != identity(named) and (
            path.stem.casefold() == identity(value)
            or path.stem.casefold().startswith(identity(value) + '_'))]
        if named and len(filename_matches) == 1:
            cods.discard(named)
            named = filename_matches[0]
            cods.add(named)
        if named and printed and any(identity(value) != identity(named) for value in printed):
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
    return sorted(cods), sorted(references), warnings


def scan_inventory(root, persist=False):
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
                with path.open('rb') as stream:
                    # The parsing limit is not a file identity. Distinct large
                    # PDFs must never share a synthetic digest/cache entry.
                    if stat.st_size <= MAX_PDF_BYTES:
                        candidate = stream.read(MAX_PDF_BYTES + 1)
                        if len(candidate) <= MAX_PDF_BYTES:
                            pdf_data = candidate
                            digest = hashlib.sha256(candidate).hexdigest()
                        else:
                            checksum = hashlib.sha256(candidate)
                            for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                                checksum.update(chunk)
                            digest = checksum.hexdigest()
                    else:
                        digest = hashlib.file_digest(stream, 'sha256').hexdigest()
            except OSError:
                digest = None
            if (old and old.get('inspection_version') == INSPECTION_VERSION
                    and old.get('inspection_prefix') == number_prefix
                    and digest is not None and old.get('sha256') == digest and old.get('size') == stat.st_size):
                # A restored byte-identical PDF remains the same issuance even
                # though atomic rollback changed its filesystem timestamp.
                entry = dict(old, present=True, mtime_ns=stat.st_mtime_ns)
            else:
                # Parse the same bytes we hashed. A PDF replaced in Explorer
                # between separate reads must not mix one hash with another COD.
                cods, refs, warnings = inspect_pdf(path, pdf_data=pdf_data, number_prefix=number_prefix)
                # Corrections do not make a previously observed number free.
                entry = {'path': relative, 'cods': sorted(set(cods) | set((old or {}).get('cods', []))),
                         'references': sorted(set(refs) | set((old or {}).get('references', []))),
                         'current_cods': cods, 'warnings': warnings, 'present': True,
                         'size': len(pdf_data) if pdf_data is not None else stat.st_size,
                         'mtime_ns': stat.st_mtime_ns, 'sha256': digest,
                         'inspection_version': INSPECTION_VERSION, 'inspection_prefix': number_prefix}
                if old and digest is not None and old.get('sha256') == digest and old.get('managed_owner'):
                    entry['managed_owner'] = old['managed_owner']
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
            for cod in {identity(value) for value in entry.get('current_cods', entry['cods'])}:
                owners.setdefault(cod, []).append(entry['path'])
    for cod, paths in owners.items():
        if len(paths) > 1:
            result['warnings'].append(f"COD {cod} appears in multiple PDFs ({', '.join(Path(path).name for path in paths[:3])}). Verify which certificate is current before sending.")
    return result


def cod_conflicts(inventory, cod, allowed_paths=(), owned=False):
    allowed = {str(path).replace('\\', '/').casefold() for path in allowed_paths}
    return [entry for entry in inventory['entries']
            if any(identity(value) == identity(cod) for value in entry['cods'])
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
    if not str(reference or '').strip():
        return []
    allowed = {str(path).replace('\\', '/').casefold() for path in allowed_paths}
    found = [entry for entry in inventory['entries']
             if any(identity(value) == identity(reference) for value in entry['references'])
             and entry['path'].casefold() not in allowed]
    if not found:
        return []
    names = ', '.join(Path(entry['path']).name for entry in found[:3])
    return [f'Reference {reference} already appears in {names}. This may be the same inward or an already-created certificate. Review it; continuing keeps this reference unchanged.']


def issuance_paths(inventory, cod, source_kind, plant, key):
    """Only committed issuance evidence establishes ownership, not a typed COD."""
    record_key = lambda value: str(value).strip().casefold() if source_kind == 'workbook' else str(value)
    return [entry['path'] for entry in inventory['entries']
            if isinstance(entry.get('managed_owner'), dict)
            and entry['managed_owner'].get('source_kind') == source_kind
            and entry['managed_owner'].get('plant') == plant
            and record_key(entry['managed_owner'].get('key')) == record_key(key)
            and identity(entry['managed_owner'].get('cod')) == identity(cod)]


def record_issuance(root, path, cod, source_kind, plant, key):
    inventory = scan_inventory(root, persist=True)
    relative = Path(path).relative_to(Path(root)).as_posix()
    entry = next(entry for entry in inventory['entries'] if entry['path'].casefold() == relative.casefold())
    entry['managed_owner'] = {'source_kind': source_kind, 'plant': plant, 'key': str(key), 'cod': str(cod)}
    inventory.pop('warnings', None)
    save_json(Path(root) / 'Data' / 'Certificate Inventory.json', inventory)
    return inventory


def verify_publication(root, snapshot, cod, reference, allowed_paths=(), owned=False, allow_warnings=False):
    """Recheck Explorer/PowerPoint changes made while the PDF was rendering."""
    current = scan_inventory(root)
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
    def reference_owners(inventory):
        if not str(reference or '').strip():
            return {}
        return {entry['path'].casefold(): (entry.get('sha256'), entry.get('size'))
                for entry in inventory['entries']
                if entry['path'].casefold() not in allowed
                and any(identity(value) == identity(reference) for value in entry['references'])}
    before = reference_owners(snapshot)
    after = reference_owners(current)
    if any(path not in before or signature != before[path] for path, signature in after.items()):
        raise ValueError('Another PDF with this reference appeared while rendering. Nothing was overwritten. Refresh and confirm the duplicate-reference warning first.')
