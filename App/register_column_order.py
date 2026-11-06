"""Move the operational fields and tracking columns without rewriting XLSX data.

The original package is retained. Only cell addresses, dependent references and
column placement change; workbook properties and cached formula values survive.
"""
from copy import deepcopy
from io import BytesIO
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from register_protection import externally_protected
from certificate_core import batch_content, norm, read_batches, select_register

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DETAILS = ('Comments', 'Transporter/Vehicle Details', 'Docket No./Vehicle No', 'Weight Slip/Weight Scale')
TRACKING = ('cod_no', 'cod_date', 'issued')
LABELS = ('COD No', 'COD Date', 'Certificate Issued')
MIN_WIDTHS = (18, 26, 24, 24, 12, 13, 18)


class UnsupportedColumnOrderError(ValueError):
    """Keep the existing layout when Excel is needed to move its features."""


def _tag(name):
    return f'{{{NS}}}{name}'


def _xml(payload):
    for _, (prefix, uri) in ET.iterparse(BytesIO(payload), events=['start-ns']):
        if not re.fullmatch(r'ns\d+', prefix):
            ET.register_namespace(prefix, uri)
    return ET.fromstring(payload)


def _bytes(root):
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def normalize_register_columns(data):
    from workbook_store import _sheet_member
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        width = max(sheet.max_column, max(cell.column for cell in sheet[header] if cell.value is not None))
        headings = {norm(cell.value):cell.column for cell in sheet[header] if cell.value is not None}
        tail, missing = [], {}
        for key, name in zip(('comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale'), DETAILS):
            column = cols[key] + 1 if key in cols else headings.get(norm(name))
            if column is None:
                column = width + len(missing) + 1
                missing[column] = name
            tail.append(column)
        for key, name in zip(TRACKING, LABELS):
            column = cols[key] + 1 if key in cols else None
            if column is None:
                column = width + len(missing) + 1
                missing[column] = name
            tail.append(column)
        order = [column for column in range(1, width + 1) if column not in tail] + tail
        mapping = {old:new for new,old in enumerate(order,1)}
        if not missing and all(old == new for old,new in mapping.items()):
            return data
        if externally_protected(sheet):
            raise UnsupportedColumnOrderError('This Excel sheet is protected. Unprotect it before moving the COD, date and status columns.')
        if sheet._charts or sheet._images:
            raise UnsupportedColumnOrderError('Move the final register columns in Excel first: this register contains a drawing that needs Excel to preserve its references.')
        target_sheet = sheet.title

        def cell_ref(value):
            match = re.fullmatch(r'(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)?', value)
            if not match:
                raise UnsupportedColumnOrderError(f'The register column move cannot safely translate reference {value}.')
            absolute, letters, row_absolute, row = match.groups()
            old = column_index_from_string(letters)
            return absolute + get_column_letter(mapping.get(old, old)) + row_absolute + (row or '')

        def ranges(value, split=False, expand=False):
            if ':' not in value:
                return cell_ref(value)
            left,right = value.split(':',1)
            a = re.fullmatch(r'(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)?',left)
            b = re.fullmatch(r'(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)?',right)
            if not a or not b:
                if re.fullmatch(r'\$?\d+:\$?\d+',value):
                    return value
                raise UnsupportedColumnOrderError(f'The register column move cannot safely translate range {value}.')
            first,last = column_index_from_string(a[2]),column_index_from_string(b[2])
            columns = sorted(mapping.get(i,i) for i in range(first,last+1))
            if expand and first == 1 and last == width:
                columns = list(range(1,len(order)+1))
            groups = []
            for column in columns:
                if groups and column == groups[-1][-1] + 1:
                    groups[-1].append(column)
                else:
                    groups.append([column])
            if len(groups) > 1 and not split:
                raise UnsupportedColumnOrderError(f'Range {value} would be split by the column move. Move these columns in Excel to preserve that dependency.')
            return ' '.join(a[1]+get_column_letter(group[0])+a[3]+(a[4] or '')+':' +
                            b[1]+get_column_letter(group[-1])+b[3]+(b[4] or '') for group in groups)

        def formula(value, owner):
            if not value:
                return value
            equals = value.startswith('=')
            tokens = Tokenizer(value if equals else '='+value).items
            for token in tokens:
                if token.type != 'OPERAND' or token.subtype != 'RANGE':
                    continue
                prefix, ref = ('','')
                if '!' in token.value:
                    prefix,ref = token.value.rsplit('!',1)
                    if prefix.strip("'").replace("''", "'") != target_sheet:
                        continue
                    prefix += '!'
                elif owner == target_sheet:
                    ref = token.value
                else:
                    continue
                if re.fullmatch(r'\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+',ref):
                    token.value = prefix + ranges(ref)
            result = ''.join(token.value for token in tokens)
            return ('=' if equals else '') + result

        before = [(batch_content(row),row['cod_no'],row['cod_date'],row['issued']) for row in read_batches(BytesIO(data))]
        with ZipFile(BytesIO(data)) as source:
            member = _sheet_member(source,target_sheet)
            root = _xml(source.read(member))
            if root.find(_tag('drawing')) is not None or root.find(_tag('legacyDrawing')) is not None:
                raise UnsupportedColumnOrderError('Move the register columns in Excel first to preserve its drawing or cell comments.')
            for node in root.iter():
                name = node.tag.rsplit('}',1)[-1]
                if name in {'f','formula','formula1','formula2'}:
                    if name == 'f' and node.get('t') not in {None,'normal'}:
                        raise UnsupportedColumnOrderError('Move these columns in Excel first to preserve shared or array formulas.')
                    node.text = formula(node.text,target_sheet)
                for attr in ('ref','sqref','activeCell','topLeftCell'):
                    if node.get(attr):
                        value = node.attrib[attr]
                        if name == 'dimension':
                            node.attrib[attr] = f'A1:{get_column_letter(len(order))}{sheet.max_row}'
                        else:
                            node.attrib[attr] = ' '.join(ranges(ref,split=attr=='sqref',expand=name in {'autoFilter'})
                                                        for ref in value.split())
                if name == 'filterColumn' and node.get('colId') is not None:
                    raise UnsupportedColumnOrderError('Clear active register filter criteria in Excel before moving columns.')
            sheet_data = root.find(_tag('sheetData'))
            for row in sheet_data:
                row_number = int(row.attrib['r'])
                original_cells = {column_index_from_string(re.match(r'[A-Z]+',cell.attrib['r']).group()):cell
                                  for cell in row if cell.tag == _tag('c')}
                # Cell address r is separate from range references in validations/merges.
                for old,cell in original_cells.items():
                    cell.attrib['r'] = f'{get_column_letter(mapping.get(old,old))}{row_number}'
                guide = original_cells.get(cols.get('reference',cols['supplier']) + 1)
                for old,label in missing.items():
                    cell = ET.Element(_tag('c'), {'r':f'{get_column_letter(mapping[old])}{row_number}'})
                    if guide is not None and guide.get('s'):
                        cell.set('s',guide.get('s'))
                    if row_number == header:
                        cell.set('t','inlineStr')
                        ET.SubElement(ET.SubElement(cell,_tag('is')),_tag('t')).text = label
                    row.append(cell)
                cells = sorted([cell for cell in row if cell.tag == _tag('c')],key=lambda cell:column_index_from_string(re.match(r'[A-Z]+',cell.attrib['r']).group()))
                for cell in list(row):
                    if cell.tag == _tag('c'):
                        row.remove(cell)
                row[:0] = cells
                if 'spans' in row.attrib:
                    row.attrib['spans'] = f'1:{len(order)}'
            column_section = root.find(_tag('cols'))
            dimensions = {}
            if column_section is not None:
                for column in column_section:
                    for old in range(int(column.get('min')),int(column.get('max'))+1):
                        new = mapping.get(old,old)
                        clone = deepcopy(column);clone.set('min',str(new));clone.set('max',str(new));dimensions[new]=clone
                column_section.clear()
            else:
                column_section = ET.Element(_tag('cols'))
                root.insert(list(root).index(sheet_data),column_section)
            for old,minimum in zip(tail,MIN_WIDTHS):
                new = mapping[old]
                column = dimensions.setdefault(new,ET.Element(_tag('col'),{'min':str(new),'max':str(new)}))
                column.set('width',str(max(float(column.get('width','0')),minimum)))
                column.set('customWidth','1')
            column_section.extend(dimensions[column] for column in sorted(dimensions))
            changes = {member:_bytes(root)}
            for table in sheet.tables.values():
                if not table.ref.startswith(f'A{header}:') or column_index_from_string(re.search(r':([A-Z]+)',table.ref).group(1)) != width:
                    raise UnsupportedColumnOrderError('Move these columns in Excel first: the register contains a separate table.')
                table_member = next((name for name in source.namelist() if name.startswith('xl/tables/') and name.endswith('.xml')
                                     and _xml(source.read(name)).get('name') == table.name),None)
                if table_member is None:
                    raise UnsupportedColumnOrderError('The register table could not be checked before moving columns.')
                table_root=_xml(source.read(table_member))
                table_root.set('ref',ranges(table_root.get('ref'),expand=True))
                auto=table_root.find(_tag('autoFilter'))
                if auto is not None:
                    auto.set('ref',ranges(auto.get('ref'),expand=True))
                    for criterion in auto.findall(_tag('filterColumn')):
                        old=int(criterion.get('colId'))+1
                        criterion.set('colId',str(mapping[old]-1))
                table_columns=table_root.find(_tag('tableColumns'))
                existing={index:column for index,column in enumerate(table_columns,1)}
                next_id=max([int(column.get('id')) for column in existing.values()] or [0])
                for old,label in missing.items():
                    next_id+=1
                    existing[old]=ET.Element(_tag('tableColumn'),{'id':str(next_id),'name':label})
                table_columns[:]=[existing[old] for old in order]
                table_columns.set('count',str(len(order)))
                for node in table_root.iter():
                    if node.tag.rsplit('}',1)[-1] in {'calculatedColumnFormula','totalsRowFormula'}:
                        node.text=formula(node.text,target_sheet)
                changes[table_member]=_bytes(table_root)
            # References from other tabs and print/defined names follow the moved cells.
            for other in book.worksheets:
                if other.title == target_sheet:
                    continue
                name = _sheet_member(source,other.title)
                other_root = _xml(source.read(name));changed=False
                for f in other_root.iter(_tag('f')):
                    updated=formula(f.text,other.title)
                    if updated != f.text:
                        f.text=updated;changed=True
                if changed:
                    changes[name]=_bytes(other_root)
            workbook_root=_xml(source.read('xl/workbook.xml'));changed=False
            for defined in workbook_root.iter(_tag('definedName')):
                updated=formula(defined.text,'')
                if updated != defined.text:
                    defined.text=updated;changed=True
            if changed:
                changes['xl/workbook.xml']=_bytes(workbook_root)
            output=BytesIO()
            with ZipFile(output,'w') as target:
                target.comment=source.comment
                for info in source.infolist():
                    target.writestr(info,changes.get(info.filename,source.read(info.filename)))
            result=output.getvalue()
        after=[(batch_content(row),row['cod_no'],row['cod_date'],row['issued']) for row in read_batches(BytesIO(result))]
        if before != after:
            raise ValueError('The column move changed inward facts. The original workbook has been kept.')
        return result
    finally:
        book.close()
