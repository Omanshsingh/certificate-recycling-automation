"""Keep optional certificate settings inside document properties, not columns."""
from io import BytesIO
import json
import re
import hashlib
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from copy import copy

NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
PROPERTY = 'EcorecoCertificateDetails'
PROPERTY_V2 = 'EcorecoCertificateDetailsV2_'
KEYS = {'form6', 'reference', 'certificate_type', 'certificate_fields', 'certificate_weight',
        '_pdf_recovery_pending', '_hdd_cod_no', '_hdd_cod_date', '_external_pdf'}
PROPERTY_CHUNK_SIZE = 180  # Excel limits each custom-property value to 255 characters.


def xml_document(root):
    """Use the default root namespace for Office package compatibility."""
    payload = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    match = re.search(rb'<([A-Za-z_][\w.-]*):[A-Za-z_]', payload)
    if match:
        prefix = match[1]
        payload = payload.replace(b'xmlns:' + prefix + b'=', b'xmlns=', 1)
        payload = payload.replace(b'<' + prefix + b':', b'<').replace(b'</' + prefix + b':', b'</')
    return payload


def read_details(source):
    if hasattr(source, 'getvalue'):
        source = BytesIO(source.getvalue())
    with ZipFile(source) as archive:
        if 'docProps/custom.xml' not in archive.namelist():
            return {}
        root = ET.fromstring(archive.read('docProps/custom.xml'))
        # Excel silently truncates a custom-property string after 255
        # characters. V2 stores each client's serialized details in numbered
        # short chunks so saving the workbook in Excel cannot corrupt the data.
        chunks = {}
        for prop in root:
            name = prop.get('name', '')
            match = re.fullmatch(re.escape(PROPERTY_V2) + r'([0-9a-f]{32})_(\d{4})', name)
            if match:
                try:
                    chunks.setdefault(match[1], {})[int(match[2])] = next(iter(prop)).text or ''
                except StopIteration as exc:
                    raise ValueError('Saved certificate settings need checking; the Excel file was not changed.') from exc
        value = {}
        for parts in chunks.values():
            if sorted(parts) != list(range(len(parts))):
                raise ValueError('Saved certificate settings need checking; the Excel file was not changed.')
            try:
                record = json.loads(''.join(parts[index] for index in range(len(parts))))
                if (not isinstance(record, dict) or not isinstance(record.get('key'), str)
                        or not isinstance(record.get('details'), dict)):
                    raise ValueError()
                value[record['key']] = record['details']
            except (ValueError, TypeError) as exc:
                raise ValueError('Saved certificate settings need checking; the Excel file was not changed.') from exc

        # Read v1 workbooks, and automatically migrate them on the next edit.
        prop = next((p for p in root if p.get('name') == PROPERTY), None)
        if prop is not None:
            try:
                legacy = json.loads(next(iter(prop)).text or '{}')
                if not isinstance(legacy, dict) or any(not isinstance(v, dict) for v in legacy.values()):
                    raise ValueError()
                value = {**legacy, **value}
            except (ValueError, StopIteration) as exc:
                # A truncated v1 value is recoverable if a complete v2
                # migration is present; otherwise fail closed.
                if not chunks:
                    raise ValueError('Saved certificate settings need checking; the Excel file was not changed.') from exc
        return value


def detail_key(sheet, case_id):
    return json.dumps([sheet, str(case_id)], ensure_ascii=False, separators=(',', ':'))


def _matching_detail_keys(saved, sheet, case_id, legacy_case_ids=()):
    """Match harmless casing changes without guessing across renamed registers."""
    identities = {str(value).strip().casefold() for value in (case_id, *legacy_case_ids)}
    matches = []
    for key in saved:
        try:
            identity = json.loads(key)
        except (ValueError, TypeError):
            continue
        if (isinstance(identity, list) and len(identity) == 2
                and isinstance(identity[0], str) and isinstance(identity[1], str)
                and identity[0].casefold() == str(sheet).casefold()
                and identity[1].strip().casefold() in identities):
            matches.append(key)
    if matches and any(saved[key] != saved[matches[0]] for key in matches[1:]):
        raise ValueError(f'Saved certificate settings for Serial Number {case_id} are ambiguous. '
                         'Keep its existing settings and resolve the conflicting Serial Numbers before issuing.')
    return matches


def lookup_details(saved, sheet, case_id, *, legacy_case_ids=()):
    keys = _matching_detail_keys(saved, sheet, case_id, legacy_case_ids)
    return saved[keys[0]] if keys else {}


def _physical_serial_aliases(data, sheet_name, case_id):
    """Recover old numeric metadata keys only from the matching Excel cell."""
    if not re.fullmatch(r'-?0\d+', str(case_id)):
        return ()
    from openpyxl import load_workbook
    from certificate_core import register_header, text
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        if sheet_name not in book.sheetnames:
            return ()
        sheet = book[sheet_name]
        header = register_header(sheet)
        if not header or 'sn' not in header[1]:
            return ()
        first, columns = header
        aliases = []
        for row in range(first + 1, sheet.max_row + 1):
            cell = sheet.cell(row, columns['sn'] + 1)
            raw = cell.value
            if (isinstance(raw, (int, float)) and not isinstance(raw, bool)
                    and re.fullmatch(r'0+', cell.number_format)
                    and float(raw).is_integer()
                    and f'{int(raw):0{len(cell.number_format)}d}' == str(case_id)):
                aliases.append(text(raw))
        return tuple(aliases)
    finally:
        book.close()


def write_details(data, sheet, case_id, details):
    updates = {key:value for key,value in details.items() if key in KEYS}
    if not updates:
        return data
    saved = read_details(BytesIO(data))
    key = detail_key(sheet, case_id)
    aliases = _physical_serial_aliases(data, sheet, case_id)
    previous = _matching_detail_keys(saved, sheet, case_id, aliases)
    existing = saved[previous[0]] if previous else {}
    for old_key in previous:
        saved.pop(old_key)
    saved[key] = {**existing, **updates}
    return _replace_details(data, saved)


def remove_details(data, sheet, case_id, *, original_data=None):
    """Forget this inward's private settings so a future entry cannot inherit them."""
    saved = read_details(BytesIO(data))
    aliases = _physical_serial_aliases(original_data or data, sheet, case_id)
    keys = _matching_detail_keys(saved, sheet, case_id, aliases)
    if not keys:
        return data
    for key in keys:
        saved.pop(key)
    return _replace_details(data, saved)


def _replace_details(data, saved):
    """Replace workbook detail properties with bounded, Excel-safe records."""
    with ZipFile(BytesIO(data)) as source:
        root = (ET.fromstring(source.read('docProps/custom.xml')) if 'docProps/custom.xml' in source.namelist()
                else ET.Element('{%s}Properties' % NS))
        # Replace v1 and v2 properties with bounded v2 values. Store the key
        # alongside each detail record so the split parts are self-describing.
        for prop in list(root):
            if prop.get('name') == PROPERTY or prop.get('name', '').startswith(PROPERTY_V2):
                root.remove(prop)
        pid = max([int(p.get('pid', '1')) for p in root] or [1]) + 1
        for record_key, record_details in sorted(saved.items()):
            payload = json.dumps({'key':record_key, 'details':record_details},
                                 ensure_ascii=False, separators=(',', ':'), allow_nan=False)
            digest = hashlib.sha256(record_key.encode('utf-8')).hexdigest()[:32]
            for index, start in enumerate(range(0, max(1, len(payload)), PROPERTY_CHUNK_SIZE)):
                name = f'{PROPERTY_V2}{digest}_{index:04d}'
                prop = ET.SubElement(root, '{%s}property' % NS, {
                    'fmtid':'{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                    'pid':str(pid), 'name':name})
                pid += 1
                ET.SubElement(prop, '{%s}lpwstr' % VT).text = payload[start:start + PROPERTY_CHUNK_SIZE]
        content = ET.fromstring(source.read('[Content_Types].xml'))
        ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
        if not any(p.get('PartName') == '/docProps/custom.xml' for p in content):
            ET.SubElement(content, '{%s}Override' % ct, {'PartName':'/docProps/custom.xml',
                'ContentType':'application/vnd.openxmlformats-officedocument.custom-properties+xml'})
        rels = ET.fromstring(source.read('_rels/.rels'))
        rn = 'http://schemas.openxmlformats.org/package/2006/relationships'
        rt = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
        if not any(p.get('Type') == rt for p in rels):
            used = {p.get('Id') for p in rels}; index = 1
            while f'rId{index}' in used:
                index += 1
            ET.SubElement(rels, '{%s}Relationship' % rn, {'Id':f'rId{index}', 'Type':rt, 'Target':'docProps/custom.xml'})
        changes = {'docProps/custom.xml':xml_document(root),
                   '[Content_Types].xml':xml_document(content),
                   '_rels/.rels':xml_document(rels)}
        output = BytesIO()
        with ZipFile(output, 'w') as target:
            target.comment = source.comment
            for info in source.infolist():
                target.writestr(copy(info), changes.pop(info.filename, source.read(info.filename)))
            for name, payload in changes.items():
                target.writestr(name, payload)
        return output.getvalue()
