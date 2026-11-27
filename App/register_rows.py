"""Move register rows in the XLSX package, retaining formulas and Office parts."""
from copy import copy
from io import BytesIO
import re
from xml.sax.saxutils import escape, unescape, quoteattr
from zipfile import ZipFile

from register_reader import load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries, get_column_letter, coordinate_to_tuple

from register_protection import externally_protected
from certificate_core import norm, select_register
from workbook_store import (_attributes, _element, _inward_blocks, _register_occupied_cells,
                            _sheet_member, inward_client_columns)


def expanded_shared_formulas(xml, sheet):
    """Return openpyxl-expanded formulas for complete Excel shared-formula groups."""
    groups, ranges = {}, {}
    for cell_match in _element(xml, 'c').finditer(xml):
        cell_xml = cell_match.group()
        coordinate = _attributes(cell_xml).get('r')
        formula_match = _element(cell_xml, 'f').search(cell_xml)
        if not coordinate or not formula_match:
            continue
        attrs = _attributes(formula_match.group())
        if attrs.get('t') != 'shared':
            continue
        index = attrs.get('si')
        if index is None:
            raise ValueError('An Excel shared formula has no group number. Check it in Excel before moving rows.')
        groups.setdefault(index, []).append(coordinate)
        if attrs.get('ref'):
            ranges.setdefault(index, []).append(attrs['ref'])

    expanded = {}
    for index, members in groups.items():
        refs = ranges.get(index, [])
        if len(refs) != 1:
            raise ValueError('An Excel shared formula group is incomplete. Check it in Excel before moving rows.')
        left, top, right, bottom = range_boundaries(refs[0])
        expected = (right - left + 1) * (bottom - top + 1)
        positions = [coordinate_to_tuple(cell) for cell in members]
        if (len(set(members)) != expected or any(not (top <= row <= bottom and left <= column <= right)
                                                for row, column in positions)):
            raise ValueError('An Excel shared formula group is incomplete. Check it in Excel before moving rows.')
        for coordinate in members:
            value = sheet[coordinate].value
            if not isinstance(value, str) or not value.startswith('='):
                raise ValueError(f'Excel could not expand the shared formula at {coordinate}. Check it before moving rows.')
            expanded[coordinate] = value
    return expanded


def _material_totals(sheet, header, cols, records):
    """Identify simple totals over one client's own material-weight column."""
    if 'weight' not in cols:
        return set(), {}
    weight = get_column_letter(cols['weight'] + 1)
    pattern = re.compile(rf'=SUM\({weight}(\d+):{weight}(\d+)\)')
    total_cols = [cell.column for cell in sheet[header]
                  if norm(cell.value) in {'totalweight', 'totalweightkg', 'totalweightinkg'}]
    last_row = sheet.max_row
    formulas, ends = set(), {}
    for index, batch in enumerate(records):
        if not batch['items']:
            continue
        start = batch['source_row']
        limit = records[index + 1]['source_row'] - 1 if index + 1 < len(records) else last_row
        totals = [sheet.cell(start, column) for column in total_cols]
        totals.extend(sheet.cell(row, cols['weight'] + 1) for row in batch.get('summary_rows', []))
        for cell in totals:
            if cell.data_type != 'f' or not isinstance(cell.value, str):
                continue
            match = pattern.fullmatch(re.sub(r'\s|\$', '', cell.value).upper())
            if match and int(match[1]) == start and start < int(match[2]) <= limit:
                formulas.add(cell.coordinate)
                ends[start] = max(ends.get(start, start), int(match[2]))
    return formulas, ends


def compact_register_gaps(data, *, source_rows=None, book=None):
    """Remove cleared rows and shorten verified material totals and their merges."""
    from certificate_core import read_batches, batch_content
    from workbook_store import register_content_end, UnsupportedReorderError
    before = read_batches(BytesIO(data))
    if not before:
        return data
    owns_book = book is None
    if owns_book:
        book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        if externally_protected(sheet):
            return data
        records = [batch for batch in before if batch['source_sheet'] == sheet.title]
        totals, material_ends = _material_totals(sheet, header, cols, records)
        end = max(register_content_end(sheet, header), max(material_ends.values(), default=header))
        with ZipFile(BytesIO(data)) as archive:
            xml = archive.read(_sheet_member(archive, sheet.title)).decode('utf-8')
        occupied = {row for row, _ in _register_occupied_cells(sheet, xml)}
        client_cols = inward_client_columns(sheet, header, cols)
        blocks = dict(_inward_blocks(sheet, records))
        locked_rows = set()
        for area in sheet.merged_cells.ranges:
            if sheet.cell(area.min_row, area.min_col).value is None:
                continue
            can_shorten = (area.min_col == area.max_col and area.min_col in client_cols
                           and area.max_row <= blocks.get(area.min_row, 0))
            if not can_shorten:
                locked_rows.update(range(area.min_row, area.max_row + 1))
        gaps = []
        for row in range(header + 1, end + 1):
            if row not in occupied and row not in locked_rows:
                gaps.append(row)
        if not gaps:
            return data
        try:
            updated = delete_empty_register_rows(data, gaps, material_total_cells=totals)
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
        if owns_book:
            book.close()


def delete_empty_register_rows(data, rows, *, material_total_cells=()):
    """Delete verified empty rows and keep client totals aligned with their materials."""
    groups = []
    for row in sorted(set(rows)):
        if groups and groups[-1][1] + 1 == row:
            groups[-1][1] = row
        else:
            groups.append([row, row])
    for first, last in reversed(groups):
        count = last - first + 1
        data = shift_register_rows(data, first, count, delete=True, material_total_cells=material_total_cells)
        material_total_cells = {f'{get_column_letter(col)}{row - count if row > last else row}'
                                for row, col in map(coordinate_to_tuple, material_total_cells)}
    return data


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
        scalar_cols = inward_client_columns(sheet, header, cols)
        shifted = shift_register_rows(data, insert_at, count, expand_shared=True)
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


def shift_register_rows(data, first, count, *, delete=False, expand_shared=True, material_total_cells=()):
    """Insert or remove rows, expanding complete shared formulas without losing calculations."""
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

        def formula(value, *, rule=False, coordinate=None):
            # Only the gap cleaner supplies these verified material totals.
            shorten_total = delete and coordinate in material_total_cells
            parts = []
            for token in Tokenizer(value if value.startswith('=') else '=' + value).items:
                text = token.value
                if token.type == 'FUNC' and text.upper() in {'INDIRECT(', 'OFFSET(', 'ADDRESS(', 'ROW(', 'CELL('}:
                    raise ValueError('This register has calculations tied to row positions. Move its rows in Excel instead.')
                if token.type == 'OPERAND' and token.subtype == 'RANGE':
                    if not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?', text, re.I):
                        raise ValueError('This register has linked calculations. Move its rows in Excel instead.')
                    if delete and not rule and not shorten_total and any(first <= int(n) <= last for n in re.findall(r'\d+', text)):
                        raise ValueError('Another calculation uses this inward. Remove the entry in Excel instead.')
                    clipped_row = first - 1 if shorten_total else first
                    text = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                                  lambda m: m[1] + str(row_number(int(m[2])) or clipped_row), text, flags=re.I)
                parts.append(text)
            return ''.join(parts)

        with ZipFile(BytesIO(data)) as archive:
            if any(name.casefold() == 'xl/vbaproject.bin' for name in archive.namelist()):
                raise ValueError('This workbook has macros. Move its rows in Excel instead.')
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            shared_values = expanded_shared_formulas(xml, sheet) if expand_shared else {}
            if any((attrs := _attributes(match.group())) and not (expand_shared and attrs.get('t') == 'shared')
                   for match in _element(xml, 'f').finditer(xml)):
                raise ValueError('This workbook has shared calculations or other special formulas. Move its rows in Excel instead.')
            if any(_element(xml, tag).search(xml) for tag in ('drawing', 'legacyDrawing', 'oleObjects', 'controls', 'rowBreaks')):
                raise ValueError('This register has objects anchored to rows. Move its rows in Excel instead.')
            for other in book:
                if other != sheet and any(cell.data_type == 'f' for row in other for cell in row):
                    raise ValueError('Another sheet has calculations. Move the register rows in Excel instead.')

            def patch_formula(match, *, rule=False, coordinate=None):
                block = match.group()
                if _attributes(block):
                    raise ValueError('This workbook has shared calculations. Move its rows in Excel instead.')
                start, end = block.index('>') + 1, block.rfind('</')
                value = unescape(block[start:end])
                return block[:start] + escape(formula(value, rule=rule, coordinate=coordinate)) + block[end:]

            def patch_row(match):
                block = match.group()
                old = int(_attributes(block)['r'])
                new = row_number(old)
                if new is None:
                    return ''
                block = re.sub(r'\br=(["\x27])\d+\1', f'r="{new}"', block, count=1)

                def patch_cell(cell_match):
                    cell_xml = cell_match.group()
                    coordinate = _attributes(cell_xml)['r']

                    def patch_cell_formula(formula_match):
                        formula_xml = formula_match.group()
                        formula_attrs = _attributes(formula_xml)
                        if formula_attrs.get('t') != 'shared':
                            return patch_formula(formula_match, coordinate=coordinate)
                        # Split a complete shared group into independent,
                        # row-adjusted formulas so Excel retains its calculations.
                        value = shared_values.get(coordinate)
                        if not value:
                            raise ValueError(f'Excel could not expand the shared formula at {coordinate}. Check it before moving rows.')
                        prefix = formula_match.group('p')
                        return f'<{prefix}f>{escape(formula(value, coordinate=coordinate))}</{prefix}f>'

                    cell_xml = _element(cell_xml, 'f').sub(patch_cell_formula, cell_xml)
                    return re.sub(r'\br=(["\x27])([A-Z]+)\d+\1',
                                  lambda c: f'r={quoteattr(c[2] + str(new))}', cell_xml, count=1)

                return _element(block, 'c').sub(patch_cell, block)

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

            def patch_ref(match, attribute='ref', *, merge=False):
                block = match.group()
                old = _attributes(block).get(attribute)
                if not old:
                    return block
                refs = [reference(ref) for ref in old.split()]
                value = ' '.join(ref for ref in refs if ref)
                if not value:
                    return ''
                if merge:
                    left, top, right, bottom = range_boundaries(value)
                    if left == right and top == bottom:
                        return ''
                block = re.sub(rf'\b{attribute}=(["\x27]).*?\1', f'{attribute}={quoteattr(value)}', block, count=1)
                for tag in ('formula', 'formula1', 'formula2'):
                    block = _element(block, tag).sub(lambda m: patch_formula(m, rule=True), block)
                return block

            for tag in ('dimension', 'autoFilter'):
                xml = _element(xml, tag).sub(patch_ref, xml)
            xml = _element(xml, 'mergeCell').sub(lambda match: patch_ref(match, merge=True), xml)
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
