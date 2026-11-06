"""Consistent dates and top-left details without reserializing an XLSX workbook."""
from io import BytesIO
import re
from xml.sax.saxutils import quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import to_excel

from register_protection import externally_protected
from certificate_core import parse_date, select_register
from workbook_store import _attributes, _element, _sheet_member, _patch_cell_styles, _set_cell, _date_style

DATE_FORMAT = 'dd/mm/yy'


def normalize_register_presentation(data):
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        if externally_protected(sheet):
            return data
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            styles = archive.read('xl/styles.xml').decode('utf-8')
            existing = next((_attributes(m.group())['numFmtId'] for m in _element(styles, 'numFmt').finditer(styles)
                             if _attributes(m.group()).get('formatCode') == DATE_FORMAT), None)
            if existing is None:
                styles, _ = _date_style(styles, 0, DATE_FORMAT)
                existing = next(_attributes(m.group())['numFmtId'] for m in _element(styles, 'numFmt').finditer(styles)
                                if _attributes(m.group()).get('formatCode') == DATE_FORMAT)
            section = _element(styles, 'cellXfs').search(styles)
            prefix = section.group('p')
            definitions = [m.group() for m in _element(section.group(), 'xf').finditer(section.group())]
            original_count = len(definitions)
            known = {value: index for index, value in enumerate(definitions)}
            ids = {}
            date_cols = {cols[key] + 1 for key in ('inward_date', 'cod_date') if key in cols}
            comments_col = cols.get('comments', -1) + 1
            for match in list(_element(xml, 'c').finditer(xml)):
                attrs = _attributes(match.group())
                cell = sheet[attrs['r']]
                if cell.row < header:
                    continue
                original = definitions[int(attrs.get('s', 0))]
                definition = original
                if definition.rstrip().endswith('/>'):
                    definition = definition.rstrip()[:-2] + f'></{prefix}xf>'
                alignment = _element(definition, 'alignment').search(definition)
                value = alignment.group() if alignment else f'<{prefix}alignment/>'
                for key, text in (('horizontal', 'left'), ('vertical', 'top')):
                    if re.search(rf'\b{key}=', value):
                        value = re.sub(rf'\b{key}=(["\x27]).*?\1', f'{key}="{text}"', value, count=1)
                    else:
                        end = value.index('>')
                        opening = value[:end]
                        opening = opening.rstrip('/') + f' {key}="{text}"' + ('/' if opening.endswith('/') else '')
                        value = opening + value[end:]
                if cell.column == comments_col and comments_col:
                    if re.search(r'\bwrapText=', value):
                        value = re.sub(r'\bwrapText=(["\x27]).*?\1', 'wrapText="1"', value, count=1)
                    else:
                        end = value.index('>')
                        opening = value[:end]
                        opening = opening.rstrip('/') + ' wrapText="1"' + ('/' if opening.endswith('/') else '')
                        value = opening + value[end:]
                definition = (definition[:alignment.start()] + value + definition[alignment.end():]
                              if alignment else definition[:definition.rfind('</')] + value + definition[definition.rfind('</'):])
                end = definition.index('>')
                opening = definition[:end]
                if re.search(r'\bapplyAlignment=', opening):
                    opening = re.sub(r'\bapplyAlignment=(["\x27]).*?\1', 'applyAlignment="1"', opening, count=1)
                else:
                    opening += ' applyAlignment="1"'
                definition = opening + definition[end:]
                is_date = cell.row > header and (cell.column in date_cols or is_date_format(cell.number_format))
                if is_date:
                    for key, value in (('numFmtId', existing), ('applyNumberFormat', '1')):
                        if re.search(rf'\b{key}=', definition):
                            definition = re.sub(rf'\b{key}=(["\x27]).*?\1', f'{key}={quoteattr(value)}', definition, count=1)
                        else:
                            end = definition.index('>')
                            definition = definition[:end] + f' {key}={quoteattr(value)}' + definition[end:]
                    if isinstance(cell.value, str) and cell.data_type != 'f':
                        parsed = parse_date(cell.value)
                        if parsed:
                            xml = _set_cell(xml, cell.coordinate, to_excel(parsed, book.epoch), 'n')
                if definition not in known:
                    known[definition] = len(definitions)
                    definitions.append(definition)
                ids[cell.coordinate] = known[definition]
            xml = _patch_cell_styles(xml, ids)
            block = section.group()
            block = block[:block.rfind('</')] + ''.join(definitions[original_count:]) + block[block.rfind('</'):]
            block = re.sub(r'\bcount=(["\x27]).*?\1', f'count="{len(definitions)}"', block, count=1)
            styles = styles[:section.start()] + block + styles[section.end():]
            changes = {member: xml.encode('utf-8'), 'xl/styles.xml': styles.encode('utf-8')}
            if all(archive.read(name) == value for name, value in changes.items()):
                return data
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, changes.get(info.filename, archive.read(info.filename)))
            return output.getvalue()
    finally:
        book.close()
