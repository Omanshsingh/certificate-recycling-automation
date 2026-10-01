"""Observe manually added PDFs without adopting their contents into Excel.

The folder is another numbering source, not an instruction source. A durable
local ledger remembers observed identifiers even if a file is later removed.
"""
from pathlib import Path
import json
import re
import hashlib

from pypdf import PdfReader
from workflow_support import save_json

MAX_PDF_BYTES = 50 * 1024 * 1024
MAX_PAGES = 8


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


def text_identifiers(text):
    """Read labelled values, never an arbitrary date/weight/serial number."""
    cods, references = set(), set()
    cod_label = r'(?:certificate|cod)[ \t]*(?:no\.?|number)[ \t]*(?:&[ \t]*date)?'
    ref_label = r'(?:reference|ref)[ \t]*(?:no\.?|number)'
    for label, found in ((cod_label, cods), (ref_label, references)):
        for match in re.finditer(rf'(?im)\b{label}[ \t]*[:\-]?[ \t]*([^\n\r]*)', text):
            tail = match[1].strip()
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


def inspect_pdf(path):
    named = filename_cod(path)
    cods, references = set(filter(None, [named])), set()
    warnings = []
    try:
        with path.open('rb') as stream:
            if b'%PDF-' not in stream.read(1024):
                raise ValueError('file is not a readable PDF')
        if path.stat().st_size > MAX_PDF_BYTES:
            raise ValueError('PDF exceeds the 50 MB inspection limit')
        reader = PdfReader(path)
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
                found, refs = text_identifiers(text)
                printed.update(found)
                references.update(refs)
        cods.update(printed)
        if named and printed and any(identity(value) != identity(named) for value in printed):
            warnings.append('Filename COD and printed COD differ; both numbers are treated as used. Check this PDF.')
        if len({identity(value) for value in printed}) > 1:
            warnings.append('More than one printed COD was found; all are treated as used. Check this PDF.')
        if not readable:
            warnings.append('PDF has no readable text (possibly a scan); only its filename COD can be checked. References need manual verification.')
        elif not cods:
            warnings.append('No COD could be identified. Rename this PDF with its COD before creating another certificate.')
        if len(reader.pages) > MAX_PAGES:
            warnings.append(f'Only the first {MAX_PAGES} PDF pages were inspected; check any later certificates manually.')
    except Exception as exc:
        warnings.append(f'Could not read PDF details: {exc}. The filename COD is still treated as used; verify other numbers manually.')
    return sorted(cods), sorted(references), warnings


def scan_inventory(root, persist=False):
    root = Path(root)
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
            stat = path.stat()
            old = entries.get(relative.casefold())
            try:
                with path.open('rb') as stream:
                    digest = hashlib.file_digest(stream, 'sha256').hexdigest() if stat.st_size <= MAX_PDF_BYTES else 'over-limit'
            except OSError:
                digest = None
            if old and digest is not None and old.get('sha256') == digest and old.get('size') == stat.st_size:
                # A restored byte-identical PDF remains the same issuance even
                # though atomic rollback changed its filesystem timestamp.
                entry = dict(old, present=True, mtime_ns=stat.st_mtime_ns)
            else:
                cods, refs, warnings = inspect_pdf(path)
                # Corrections do not make a previously observed number free.
                entry = {'path': relative, 'cods': sorted(set(cods) | set((old or {}).get('cods', []))),
                         'references': sorted(set(refs) | set((old or {}).get('references', []))),
                         'current_cods': cods, 'warnings': warnings, 'present': True,
                         'size': stat.st_size, 'mtime_ns': stat.st_mtime_ns, 'sha256': digest}
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
    return [entry['path'] for entry in inventory['entries']
            if isinstance(entry.get('managed_owner'), dict)
            and entry['managed_owner'].get('source_kind') == source_kind
            and entry['managed_owner'].get('plant') == plant
            and str(entry['managed_owner'].get('key')) == str(key)
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
    if cod_conflicts(current, cod, allowed_paths, owned):
        raise ValueError(f'COD {cod} is now present in another certificate PDF. Nothing was overwritten. Refresh and review the new automatic COD.')
    allowed = {str(path).replace('\\', '/').casefold() for path in allowed_paths}
    old = {entry['path'].casefold(): entry for entry in snapshot['entries']}
    for entry in current['entries']:
        previous = old.get(entry['path'].casefold())
        if entry['present'] and entry['path'].casefold() in allowed and previous and (
                entry.get('size'), entry.get('mtime_ns'), entry.get('sha256')) != (previous.get('size'), previous.get('mtime_ns'), previous.get('sha256')):
            raise ValueError('The existing certificate changed while this PDF was being prepared. Nothing was overwritten. Refresh and review it first.')
    before = reference_warnings(snapshot, reference, allowed_paths, owned)
    after = reference_warnings(current, reference, allowed_paths, owned)
    if after and not before and not allow_warnings:
        raise ValueError('Another PDF with this reference appeared while rendering. Nothing was overwritten. Refresh and confirm the duplicate-reference warning first.')
