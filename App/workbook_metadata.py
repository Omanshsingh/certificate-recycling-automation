"""Keep optional certificate settings inside document properties, not columns."""
from io import BytesIO
import json
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from copy import copy

NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
PROPERTY = 'EcorecoCertificateDetails'
KEYS = {'form6', 'reference', 'certificate_type', 'certificate_fields', 'certificate_weight'}


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
        prop = next((p for p in root if p.get('name') == PROPERTY), None)
        if prop is None:
            return {}
        try:
            value = json.loads(next(iter(prop)).text or '{}')
            if not isinstance(value, dict) or any(not isinstance(v, dict) for v in value.values()):
                raise ValueError()
            return value
        except (ValueError, StopIteration) as exc:
            raise ValueError('Saved certificate settings need checking; the Excel file was not changed.') from exc


def detail_key(sheet, case_id):
    return json.dumps([sheet, str(case_id)], ensure_ascii=False, separators=(',', ':'))


def write_details(data, sheet, case_id, details):
    updates = {key:value for key,value in details.items() if key in KEYS}
    if not updates:
        return data
    saved = read_details(BytesIO(data))
    key = detail_key(sheet, case_id)
    saved[key] = {**saved.get(key, {}), **updates}
    with ZipFile(BytesIO(data)) as source:
        root = (ET.fromstring(source.read('docProps/custom.xml')) if 'docProps/custom.xml' in source.namelist()
                else ET.Element('{%s}Properties' % NS))
        prop = next((p for p in root if p.get('name') == PROPERTY), None)
        if prop is None:
            prop = ET.SubElement(root, '{%s}property' % NS, {
                'fmtid':'{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                'pid':str(max([int(p.get('pid', '1')) for p in root] or [1]) + 1), 'name':PROPERTY})
        for child in list(prop):
            prop.remove(child)
        ET.SubElement(prop, '{%s}lpwstr' % VT).text = json.dumps(saved, ensure_ascii=False, allow_nan=False)
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
