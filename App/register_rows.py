"""Move register rows in the XLSX package, retaining formulas and Office parts."""
from copy import copy
from io import BytesIO
import re
from xml.sax.saxutils import escape, unescape, quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries, get_column_letter

from register_protection import externally_protected
from certificate_core import select_register
from workbook_store import _attributes, _element, _sheet_member


def compact_register_gaps(data, *, source_rows=None):
    """Remove legacy cleared row boxes, preserving populated and merged blocks."""
    from certificate_core import read_batches, batch_content
    from workbook_store import register_content_end, UnsupportedReorderError
    before = read_batches(BytesIO(data))
    if not before:
        return data
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, _ = select_register(book)
        if externally_protected(sheet):
            return data
        end = register_content_end(sheet, header)
        gaps = []
        for row in range(header + 1, end + 1):
            if any(cell.value is not None or cell.comment or cell.hyperlink for cell in sheet[row]):
                continue
            if any(area.min_row <= row <= area.max_row
                   and sheet.cell(area.min_row, area.min_col).value is not None for area in sheet.merged_cells.ranges):
                continue
            gaps.append(row)
        if not gaps:
            return data
        groups = []
        for row in gaps:
            if groups and groups[-1][1] + 1 == row:
                groups[-1][1] = row
            else:
                groups.append([row, row])
        updated = data
        try:
            for first, last in reversed(groups):
                updated = shift_register_rows(updated, first, last - first + 1, delete=True)
        except ValueError as exc:
            raise UnsupportedReorderError(f'Empty row cleanup needs review in Excel: {exc}') from exc
        def facts(rows):
            return [(batch_content(row), *(row.get(key) for key in ('cod_no','cod_date','issued',
                     'comments','transporter_details','docket_vehicle_no','weight_slip_scale','hdd_cod_no','hdd_cod_date')))
                    for row in rows]
        if facts(read_batches(BytesIO(updated))) != facts(before):
            raise ValueError('The records after empty row cleanup could not be verified. The workbook was kept.')
        if source_rows is not None:
            for row in range(header + 1, end + 1):
                if row not in gaps:
                    source_rows[(sheet.title, row)] = row - sum(gap < row for gap in gaps)
        return updated
    finally:
        book.close()


def ensure_material_capacity(data, sn, required):
    """Make room inside an existing inward, keeping the next client intact."""
    from certificate_core import read_batches, text
    from workbook_store import register_content_end
    batches = read_batches(BytesIO(data))
    matches = [b for b in batches if text(b['case_id']).casefold() == text(sn).casefold()]
    if len(matches) != 1:
        return data
    batch = matches[0]
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        first = batch['source_row']
        following = [b['source_row'] for b in batches if b['source_row'] > first]
        summaries = sorted(batch.get('summary_rows', []))
        if not following and not summaries:
            return data  # The normal editor can append at the end without moving any rows.
        end = min(following or [register_content_end(sheet, header) + 1])
        insert_at = summaries[0] if summaries else end
        capacity = insert_at - first
        if required <= capacity:
            return data
        count = required - capacity
        scalar_cols = set(range(1, sheet.max_column + 1)) - {cols[k] + 1 for k in ('item', 'quantity', 'per_piece_weight', 'weight') if k in cols}
        if batch.get('hdd_cod_no'):
            scalar_cols -= {cols[k] + 1 for k in ('cod_no', 'cod_date') if k in cols}
        shifted = shift_register_rows(data, insert_at, count)
        with ZipFile(BytesIO(shifted)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            def expand_merge(match):
                block = match.group(); attrs = _attributes(block)
                left, top, right, bottom = range_boundaries(attrs['ref'])
                if top == first and bottom == insert_at - 1 and left == right and left in scalar_cols:
                    ref = f'{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom + count}'
                    return re.sub(r'\bref=(["\x27]).*?\1', f'ref={quoteattr(ref)}', block, count=1)
                return block
            xml = _element(xml, 'mergeCell').sub(expand_merge, xml)
            if 'weight' in cols:
                weight_col = get_column_letter(cols['weight'] + 1)
                def expand_total_row(match):
                    block = match.group(); row = int(_attributes(block)['r'])
                    if not first <= row < end + count:
                        return block
                    def expand_total(match):
                        fragment = match.group(); start, finish = fragment.index('>') + 1, fragment.rfind('</')
                        value = unescape(fragment[start:finish])
                        plain = re.sub(r'\s|\$', '', value).upper()
                        if plain == f'SUM({weight_col}{first}:{weight_col}{insert_at - 1})':
                            value = f'SUM({weight_col}{first}:{weight_col}{insert_at + count - 1})'
                        return fragment[:start] + escape(value) + fragment[finish:]
                    return _element(block, 'f').sub(expand_total, block)
                xml = _element(xml, 'row').sub(expand_total_row, xml)
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, xml.encode('utf-8') if info.filename == member else archive.read(info.filename))
            return output.getvalue()
    finally:
        book.close()


def shift_register_rows(data, first, count, *, delete=False):
    """Insert blank rows or remove a complete block and shift everything below it."""
    last = first + count - 1
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, _ = select_register(book)
        if first <= header or count < 1 or externally_protected(sheet):
            raise ValueError('The register rows cannot be moved. Check its headings and protection in Excel.')
        if (sheet._charts or sheet._images or any(c.comment or c.hyperlink for row in sheet for c in row)
                or any(name not in {'_xlnm.Print_Area', '_xlnm.Print_Titles'}
                       for name in list(book.defined_names) + list(sheet.defined_names))):
            raise ValueError('This register has objects or named calculations. Remove the entry in Excel instead.')

        def row_number(row):
            if delete:
                return row if row < first else row - count if row > last else None
            return row + count if row >= first else row

        def reference(value, *, rule=False):
            left, top, right, bottom = range_boundaries(value)
            if top is None or bottom is None:
                return value
            if delete:
                if bottom < first:
                    return value
                if top > last:
                    top, bottom = top - count, bottom - count
                else:
                    top = top if top < first else first
                    bottom = bottom - count if bottom > last else first - 1
                    if bottom < top:
                        return None
            else:
                top = top + count if top >= first else top
                bottom = min(1048576, bottom + count) if bottom >= first else bottom
            start = f'{get_column_letter(left)}{top}'
            return start if ':' not in value else f'{start}:{get_column_letter(right)}{bottom}'

        def formula(value, *, rule=False):
            parts = []
            for token in Tokenizer(value if value.startswith('=') else '=' + value).items:
                text = token.value
                if token.type == 'FUNC' and text.upper() in {'INDIRECT(', 'OFFSET(', 'ADDRESS(', 'ROW(', 'CELL('}:
                    raise ValueError('This register has calculations tied to row positions. Move its rows in Excel instead.')
                if token.type == 'OPERAND' and token.subtype == 'RANGE':
                    if not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?', text, re.I):
                        raise ValueError('This register has linked calculations. Move its rows in Excel instead.')
                    if delete and not rule and any(first <= int(n) <= last for n in re.findall(r'\d+', text)):
                        raise ValueError('Another calculation uses this inward. Remove the entry in Excel instead.')
                    text = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                                  lambda m: m[1] + str(row_number(int(m[2])) or first), text, flags=re.I)
                parts.append(text)
            return ''.join(parts)

        with ZipFile(BytesIO(data)) as archive:
            if any(name.casefold() == 'xl/vbaproject.bin' for name in archive.namelist()):
                raise ValueError('This workbook has macros. Move its rows in Excel instead.')
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            if any(_attributes(match.group()) for match in _element(xml, 'f').finditer(xml)):
                raise ValueError('This workbook has shared calculations. Move its rows in Excel instead.')
            if any(_element(xml, tag).search(xml) for tag in ('drawing', 'legacyDrawing', 'oleObjects', 'controls', 'rowBreaks')):
                raise ValueError('This register has objects anchored to rows. Move its rows in Excel instead.')
            for other in book:
                if other != sheet and any(cell.data_type == 'f' for row in other for cell in row):
                    raise ValueError('Another sheet has calculations. Move the register rows in Excel instead.')

            def patch_formula(match, *, rule=False):
                block = match.group()
                if _attributes(block):
                    raise ValueError('This workbook has shared calculations. Move its rows in Excel instead.')
                start, end = block.index('>') + 1, block.rfind('</')
                return block[:start] + escape(formula(unescape(block[start:end]), rule=rule)) + block[end:]

            def patch_row(match):
                block = match.group()
                old = int(_attributes(block)['r'])
                new = row_number(old)
                if new is None:
                    return ''
                block = re.sub(r'\br=(["\x27])\d+\1', f'r="{new}"', block, count=1)
                block = _element(block, 'c').sub(lambda m: re.sub(
                    r'\br=(["\x27])([A-Z]+)\d+\1', lambda c: f'r={quoteattr(c[2] + str(new))}', m.group(), count=1), block)
                return _element(block, 'f').sub(patch_formula, block)

            section = _element(xml, 'sheetData').search(xml)
            body = _element(section.group(), 'row').sub(patch_row, section.group())
            if not delete:
                template = next((m.group() for m in _element(section.group(), 'row').finditer(section.group())
                                 if int(_attributes(m.group())['r']) == first - 1), '')
                cells = [_attributes(m.group()) for m in _element(template, 'c').finditer(template)]
                prefix = section.group('p')
                rows = []
                for row in range(first, first + count):
                    payload = ''.join(f'<{prefix}c r="{re.match(r"[A-Z]+", c["r"])[0]}{row}" s="{c.get("s", 0)}"/>' for c in cells)
                    height = _attributes(template).get('ht') if template else None
                    height = f' ht="{height}" customHeight="1"' if height else ''
                    rows.append(f'<{prefix}row r="{row}"{height}>{payload}</{prefix}row>')
                existing = list(_element(body, 'row').finditer(body))
                position = next((m.start() for m in existing if int(_attributes(m.group())['r']) >= first), body.rfind('</'))
                body = body[:position] + ''.join(rows) + body[position:]
            xml = xml[:section.start()] + body + xml[section.end():]

            def patch_ref(match, attribute='ref'):
                block = match.group()
                old = _attributes(block).get(attribute)
                if not old:
                    return block
                refs = [reference(ref) for ref in old.split()]
                value = ' '.join(ref for ref in refs if ref)
                if not value:
                    return ''
                block = re.sub(rf'\b{attribute}=(["\x27]).*?\1', f'{attribute}={quoteattr(value)}', block, count=1)
                for tag in ('formula', 'formula1', 'formula2'):
                    block = _element(block, tag).sub(lambda m: patch_formula(m, rule=True), block)
                return block

            for tag in ('dimension', 'mergeCell', 'autoFilter'):
                xml = _element(xml, tag).sub(patch_ref, xml)
            for tag in ('conditionalFormatting', 'dataValidation'):
                xml = _element(xml, tag).sub(lambda m: patch_ref(m, 'sqref'), xml)
            for tag, child in (('mergeCells', 'mergeCell'), ('dataValidations', 'dataValidation')):
                def count_children(match):
                    block = match.group()
                    size = len(list(_element(block, child).finditer(block)))
                    return re.sub(r'\bcount=(["\x27]).*?\1', f'count="{size}"', block, count=1) if size else ''
                xml = _element(xml, tag).sub(count_children, xml)
            changes = {member: xml.encode('utf-8')}
            for name in archive.namelist():
                if name.startswith('xl/tables/') and name.endswith('.xml'):
                    table = archive.read(name).decode('utf-8')
                    opening = _element(table, 'table').search(table)
                    if opening and _attributes(opening.group()).get('name') in sheet.tables:
                        for tag in ('table', 'autoFilter'):
                            table = _element(table, tag).sub(patch_ref, table)
                        changes[name] = table.encode('utf-8')
            workbook = archive.read('xl/workbook.xml').decode('utf-8')
            def print_area(match):
                block = match.group()
                attrs = _attributes(block)
                if attrs.get('name') != '_xlnm.Print_Area' or attrs.get('localSheetId') != str(book.index(sheet)):
                    return block
                return re.sub(r'(\$[A-Z]+\$)(\d+)', lambda m: m[1] + str(row_number(int(m[2])) or first), block)
            changes['xl/workbook.xml'] = _element(workbook, 'definedName').sub(print_area, workbook).encode('utf-8')
            if 'xl/calcChain.xml' in archive.namelist():
                for name, tag, attribute, value in (
                    ('xl/_rels/workbook.xml.rels', 'Relationship', 'Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'),
                    ('[Content_Types].xml', 'Override', 'PartName', '/xl/calcChain.xml'),
                ):
                    fragment = archive.read(name).decode('utf-8')
                    changes[name] = _element(fragment, tag).sub(lambda m: '' if _attributes(m.group()).get(attribute) == value else m.group(), fragment).encode('utf-8')
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    if info.filename != 'xl/calcChain.xml':
                        target.writestr(copy(info), changes.get(info.filename, archive.read(info.filename)))
            return output.getvalue()
    finally:
        book.close()
