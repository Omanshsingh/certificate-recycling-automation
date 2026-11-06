"""Read dated certificate facts from saved PDFs, without changing their files.

Filesystem timestamps are deliberately not issuance dates. Imported PDFs with no
readable issue date stay visible in coverage rather than appearing as new today.
"""
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import hashlib
import re

from pypdf import PdfReader
from certificate_core import parse_date
from certificate_inventory import MAX_PDF_BYTES, filename_cod

DATE_TOKEN = r'(?:\d{4}-\d{2}-\d{2}|\d{1,2}[./-]\d{1,2}[./-]\d{4})'
TYPES = {'standard', 'destruction', 'siemens', 'sita'}


def _printed_date(text, kind, cod):
    lines = text.splitlines()
    dates = lambda line: [day for token in re.findall(DATE_TOKEN, line)
                          if (day := parse_date(token))]
    for index, line in enumerate(lines):
        if re.search(r'\b(?:certificate\s+(?:issue\s+)?date|date\s+of\s+issue|issued\s+(?:on|date)|certificate\s+no\.?\s*(?:&|and)\s*date)\b', line, re.I):
            found = dates(line)
            if found:
                return found[-1]
            # A label with its single value on the next line is unambiguous.
            if index + 1 < len(lines) and re.fullmatch(r'\s*' + DATE_TOKEN + r'\s*', lines[index + 1]):
                return dates(lines[index + 1])[0]
    if kind == 'destruction' and cod:
        # The supplied HDD design has one row: destruction date, HDD number,
        # issue date. Never use the receipt dates on the later details pages.
        for line in lines:
            found = dates(line)
            if len(found) == 2 and re.search(r'(?<![\w./-])' + re.escape(cod) + r'(?![\w./-])', line):
                return found[-1]
    return None


@lru_cache(maxsize=512)
def _inspect(path_text, digest):
    path = Path(path_text)
    if not path.is_file() or not 0 < path.stat().st_size <= MAX_PDF_BYTES:
        return None
    data = path.read_bytes()
    if hashlib.sha256(data).hexdigest() != digest:
        return None  # Replaced during the scan; the next refresh will retry.
    try:
        reader = PdfReader(BytesIO(data))
        if reader.is_encrypted or not reader.pages:
            return None
        meta = reader.metadata or {}
        text = reader.pages[0].extract_text(extraction_mode='layout') or ''
        subject = str(meta.get('/CertificateType') or meta.get('/Subject') or '').strip().casefold()
        kind = subject if subject in TYPES else None
        if re.search(r'_hdd$', path.stem, re.I):
            kind = 'destruction'
        if kind is None:
            title = str(meta.get('/Title') or '')
            if re.search(r'\b(?:secured\s+data\s+destruction|hdd)\b', title, re.I):
                kind = 'destruction'
            elif re.search(r'\bSITA\s+INFORMATION\b', text, re.I):
                kind = 'sita'
            elif re.search(r'\b(?:model\s+serial|x[- ]ray\s+tube)\b', text, re.I):
                kind = 'siemens'
            elif re.search(r'certificate\s+of\s+(?:e[- ]?waste\s+)?recycling', text, re.I):
                kind = 'standard'
        cod = filename_cod(path)
        day = parse_date(meta.get('/CertificateDate')) or _printed_date(text, kind, cod)
        plant = str(meta.get('/CertificatePlant') or '').strip()
        return {'type': kind, 'date': day, 'plant': plant if plant in {'plant1', 'plant2'} else None}
    except Exception:
        return None


def folder_events(root, inventory, known_codes=()):
    """Inspect only PDFs missing from issuance history/current app records."""
    from certificate_inventory import identity
    known = {identity(cod) for cod in known_codes}
    base = (Path(root) / 'Certificates').resolve()
    result = []
    seen_pdfs = set()
    for entry in inventory.get('entries', []):
        if not entry.get('present'):
            continue
        values = {identity(cod): cod for cod in entry.get('current_cods', entry.get('cods', [])) if identity(cod)}
        # A saved file that contains a known issued COD is a copy or an
        # attachment of that certificate, not another certificate event.
        if values.keys() & known:
            continue
        digest = entry.get('sha256')
        if digest and digest in seen_pdfs:
            continue
        path = (Path(root) / entry['path']).resolve()
        if not path.is_relative_to(base) or path.suffix.casefold() != '.pdf':
            continue
        if digest:
            seen_pdfs.add(digest)
        try:
            digest = entry.get('sha256')
            facts = _inspect(str(path), digest) if digest else None
        except OSError:
            facts = None
        if facts is None:
            facts = {'type': 'destruction' if path.stem.casefold().endswith('_hdd') else None,
                     'date': None, 'plant': None}
        if entry.get('managed_owner', {}).get('source_kind') == 'manual':
            facts = dict(facts, plant=entry['managed_owner'].get('plant'))
        result.append(dict(facts, cod=next(iter(values.values())) if len(values) == 1 else None, path=entry['path']))
    return result
