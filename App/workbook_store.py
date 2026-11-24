"""Lock source files and atomically update certificate fields in their XLSX package."""
from __future__ import annotations

from contextlib import contextmanager
from copy import copy
from datetime import datetime
from io import BytesIO
from html import unescape
import errno
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.formula.tokenizer import Tokenizer, TokenizerError
from openpyxl.styles import Side
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import to_excel
from openpyxl.styles.stylesheet import write_stylesheet
from openpyxl.xml.functions import tostring

from register_protection import externally_protected, remove_issued_inward_protection
from certificate_core import (batch_content, columns, inward_details_complete, material_items, number,
                              parse_date, read_batches, register_header, select_register, status_value,
                              source_value_issues, supplier_is_always_not_applicable, text)
from workbook_metadata import write_details


def _certificate_fields_json(value):
    """Store template fields as a JSON object without altering their strings."""
    if value is None or value == "":
        value = {}
    elif isinstance(value, str):
        value = json.loads(value)
    if not isinstance(value, dict):
        raise ValueError("Certificate Fields must be a JSON object.")
    return json.dumps(value, ensure_ascii=False, allow_nan=False)


def formatted_client_template(sheet, header_row, cols, multi_line=False):
    """Choose an existing client row that carries the workbook's separator style.

    Imported registers often contain later rows with only default formatting,
    while the original client blocks have a bottom border and merged detail
    cells. Prefer that original bordered pattern and a plain supplier style.
    """
    sn_col = cols.get('sn', 0) + 1
    supplier_col = cols.get('supplier', 0) + 1
    right = max(sheet.max_column, max(cols.values(), default=0) + 1)
    rows = [row for row in range(header_row + 1, sheet.max_row + 1)
            if text(sheet.cell(row, sn_col).value)
            or text(sheet.cell(row, supplier_col).value)]
    if not rows:
        rows = [row for row in range(header_row + 1, sheet.max_row + 1)
                if any(cell.has_style for cell in sheet[row])
                or any(area.min_row == row for area in sheet.merged_cells.ranges)]
    if not rows:
        return header_row + 1

    def score(row):
        cells = [sheet.cell(row, col) for col in range(1, right + 1)]
        bottom = sum(bool(getattr(cell.border.bottom, 'style', None)) for cell in cells)
        supplier = sheet.cell(row, supplier_col)
        fill = supplier.fill
        has_fill = fill.fill_type is not None and fill.fill_type != 'none'
        plain_supplier = not supplier.font.bold and not has_fill
        block_merges = sum(area.min_row == row and area.min_col == area.max_col
                           and area.max_row > row for area in sheet.merged_cells.ranges)
        styled = sum(cell.has_style for cell in cells)
        return bottom, plain_supplier, block_merges, styled, row

    # Border count is the key signal. The final tuple entries make the chosen
    # template deterministic when several original client blocks match.
    bordered = [row for row in rows if score(row)[0] > 0]
    candidates = bordered or rows
    if multi_line:
        merged = [row for row in candidates if score(row)[2] > 0]
        candidates = merged or candidates
    return max(candidates, key=score)


CLIENT_TEXT_FIELDS = ('supplier', 'location', 'item', 'comments', 'transporter_details',
                      'docket_vehicle_no', 'weight_slip_scale')
INWARD_ROW_FIELDS = ('item', 'quantity', 'per_piece_weight', 'weight', 'cod_no', 'cod_date')


def inward_row_columns(cols):
    return {cols[key] + 1 for key in INWARD_ROW_FIELDS if key in cols}


def inward_client_columns(sheet, header, cols):
    return {cell.column for cell in sheet[header] if text(cell.value)} - inward_row_columns(cols)


def copy_client_row_styles(sheet, source_row, target_rows, width=None, *, wrap_columns=()):
    """Copy cell formatting and row height, left-aligning a newly appended block."""
    width = width or sheet.max_column
    wrap_columns = set(wrap_columns)
    source_height = sheet.row_dimensions[source_row].height
    for target_row in target_rows:
        for column in range(1, width + 1):
            target = sheet.cell(target_row, column)
            source = sheet.cell(source_row, column)
            target._style = copy(source._style)
            alignment = copy(source.alignment)
            alignment.horizontal = 'left'
            alignment.vertical = 'top'
            if column in wrap_columns:
                alignment.wrap_text = True
            target.alignment = alignment
        if source_height is not None:
            sheet.row_dimensions[target_row].height = source_height


def left_align_client_rows(sheet, rows, width=None, *, wrap_columns=()):
    """Left-align copied rows while preserving wrapping and vertical alignment."""
    width = width or sheet.max_column
    wrap_columns = set(wrap_columns)
    for row in set(rows):
        for column in range(1, width + 1):
            cell = sheet.cell(row, column)
            alignment = copy(cell.alignment)
            alignment.horizontal = 'left'
            alignment.vertical = 'top'
            if column in wrap_columns:
                alignment.wrap_text = True
            cell.alignment = alignment


def register_content_end(sheet, header_row):
    """Ignore style-only tail rows while preserving notes and existing merge groups."""
    last = max([header_row] + [cell.row for cell in sheet._cells.values()
                              if cell.value is not None or cell.comment or cell.hyperlink])
    while True:
        extended = max([last] + [area.max_row for area in sheet.merged_cells.ranges
                                if area.min_row <= last < area.max_row])
        if extended == last:
            return last
        last = extended


def next_register_row(sheet, header_row, count=1):
    """Reuse plain blank rows; leave prepared blank merge groups intact."""
    first = register_content_end(sheet, header_row) + 1
    while True:
        overlapping = [area.max_row for area in sheet.merged_cells.ranges
                       if area.min_row <= first + count - 1 and area.max_row >= first]
        if not overlapping:
            return first
        first = max(overlapping) + 1


def checked_register_column(column):
    """Never serialize a cell outside Excel's actual column grid."""
    if not 1 <= column <= 16384:
        raise ValueError('Excel has no free column for the missing field. '
                         'Move or remove unused final columns in Excel before saving; the original workbook was kept.')
    return column


def register_content_width(sheet, header_row):
    """Keep populated notes and occupied body merges; ignore formatting-only columns."""
    last = max([1] + [cell.column for cell in sheet._cells.values()
                      if cell.value is not None or cell.comment or cell.hyperlink])
    end_row = register_content_end(sheet, header_row)
    while True:
        extended = max([last] + [area.max_col for area in sheet.merged_cells.ranges
                                if area.max_row >= header_row and area.min_row <= end_row
                                and area.min_col <= last < area.max_col])
        if extended == last:
            return last
        last = extended


def next_register_column(sheet, header_row):
    return checked_register_column(register_content_width(sheet, header_row) + 1)


def _inward_blocks(sheet, batches):
    """Locate complete record blocks without treating formatted empty rows as assets."""
    starts = sorted(batch['source_row'] for batch in batches)
    if not starts:
        return []
    last = max((cell.row for row in sheet for cell in row
                if cell.value is not None), default=starts[-1])
    last = max([last] + [area.max_row for area in sheet.merged_cells.ranges
                        if area.min_row >= starts[-1]])
    return [(start, starts[index + 1] - 1 if index + 1 < len(starts) else last)
            for index, start in enumerate(starts)]


def _style_inward_dividers(sheet, blocks):
    """Give every physical edge a style, including merged-cell placeholders."""
    light = Side(style='thin', color='FFD1D5DB')
    dark = Side(style='medium', color='FF000000')
    no_border = Side(style=None)
    merged_cells = {(row, col): area for area in sheet.merged_cells.ranges
                    for row in range(area.min_row, area.max_row + 1)
                    for col in range(area.min_col, area.max_col + 1)}
    for first, last in blocks:
        for row in range(first, last + 1):
            for cell in sheet[row]:
                merged = merged_cells.get((row, cell.column))
                border = copy(cell.border)
                border.left = dark if not merged or cell.column == merged.min_col else no_border
                border.right = dark if not merged or cell.column == merged.max_col else no_border
                border.top = (dark if row == first else light) if not merged or row == merged.min_row else no_border
                border.bottom = (dark if row == last else light) if not merged or row == merged.max_row else no_border
                if merged and row == merged.min_row and cell.column == merged.min_col:
                    # Excel also stores the complete rectangle on its anchor.
                    border.bottom = dark if merged.max_row == last else light
                cell.border = border


def _patch_register_merges(xml, removed, added):
    if not removed and not added:
        return xml
    section = _element(xml, 'mergeCells').search(xml)
    prefix = re.search(r'<(?P<p>(?:[\w.-]+:)?)worksheet\b', xml).group('p')
    additions = ''.join(f'<{prefix}mergeCell ref="{area}"/>' for area in added)
    if section:
        block = _element(xml, 'mergeCell').sub(
            lambda match: '' if _attributes(match.group())['ref'] in removed else match.group(), section.group())
        position = block.rfind('</')
        block = block[:position] + additions + block[position:]
        count = len(list(_element(block, 'mergeCell').finditer(block)))
        replacement = re.sub(r'\bcount=(["\x27]).*?\1', f'count="{count}"', block, count=1) if count else ''
        return xml[:section.start()] + replacement + xml[section.end():]
    following = ('phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
                 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties',
                 'cellWatches', 'ignoredErrors', 'drawing', 'legacyDrawing', 'tableParts', 'extLst')
    position = min((match.start() for tag in following if (match := _element(xml, tag).search(xml))), default=xml.rfind('</'))
    return xml[:position] + f'<{prefix}mergeCells count="{len(added)}">{additions}</{prefix}mergeCells>' + xml[position:]


def _register_occupied_cells(sheet, xml):
    cells = {(cell.row, cell.column) for cell in sheet._cells.values()
             if cell.value is not None or cell.comment or cell.hyperlink}
    if sheet.merged_cells.ranges:
        # A workbook can contain values underneath a merge. openpyxl hides
        # them, but changing the layout must neither hide nor delete data.
        for match in _element(xml, 'c').finditer(xml):
            fragment = match.group()
            if fragment.endswith('/>'):
                continue
            cell = sheet[_attributes(fragment)['r']]
            if isinstance(cell, MergedCell) and (
                    _element(fragment, 'f').search(fragment) or _element(fragment, 'is').search(fragment)
                    or re.search(r'<(?:[\w.-]+:)?v\b[^>]*>[^<]+<', fragment)):
                cells.add((cell.row, cell.column))
    return cells


def _client_block_merges(sheet, blocks, columns, occupied_cells, unassigned_rows):
    """Give each client-level column one cell, preserving independent values."""
    removed, added = set(), []
    table_bounds = [range_boundaries(table.ref) for table in sheet.tables.values()]
    merges_by_column = {}
    for area in sheet.merged_cells.ranges:
        for column in range(area.min_col, area.max_col + 1):
            merges_by_column.setdefault(column, []).append(area)
    occupied_rows = {row for row, _ in occupied_cells}
    for first, last in blocks:
        limit = min((row - 1 for row in unassigned_rows if first < row <= last), default=last)
        end = max((row for row in range(first + 1, limit + 1) if row in occupied_rows), default=first)
        if end == first:
            continue
        for column in sorted(columns):
            if any(left <= column <= right and top <= end and bottom >= first
                   for left, top, right, bottom in table_bounds):
                continue
            if any((row, column) in occupied_cells for row in range(first + 1, end + 1)):
                continue
            overlaps = [area for area in merges_by_column.get(column, ())
                        if area.min_row <= end and area.max_row >= first]
            if any(area.min_col != column or area.max_col != column or area.min_row != first
                   or area.max_row >= end for area in overlaps):
                continue
            for area in overlaps:
                removed.add(str(area))
                sheet.unmerge_cells(str(area))
            letter = get_column_letter(column)
            merged = f'{letter}{first}:{letter}{end}'
            sheet.merge_cells(merged)
            added.append(merged)
    return removed, added


def _register_border_styles(book, sheet, blocks, xml, styles_xml, unassigned_rows):
    _, header, cols = select_register(book)
    line_columns = inward_row_columns(cols)
    occupied = _register_occupied_cells(sheet, xml)
    line_merges = {str(area) for area in sheet.merged_cells.ranges
                   if area.min_col == area.max_col and area.min_col in line_columns
                   and area.min_row < area.max_row
                   and not any((row, area.min_col) in occupied for row in range(area.min_row + 1, area.max_row + 1))
                   and any(first <= area.min_row <= area.max_row <= last for first, last in blocks)}
    for area in line_merges:
        sheet.unmerge_cells(area)
    removed, added = _client_block_merges(sheet, blocks, inward_client_columns(sheet, header, cols), occupied, unassigned_rows)
    xml = _patch_register_merges(xml, line_merges | removed, added)
    # openpyxl discards non-border formatting on merged placeholders when
    # reading. Start from their original styles so a border repair cannot
    # change fonts, fills, number formats, alignment or protection.
    originals = {}
    for match in _element(xml, 'c').finditer(xml):
        attrs = _attributes(match.group())
        cell = sheet[attrs['r']]
        originals[cell.coordinate] = int(attrs.get('s', 0))
        cell._style = copy(book._cell_styles[originals[cell.coordinate]])
    _style_inward_dividers(sheet, blocks)
    formats = _element(styles_xml, 'cellXfs').search(styles_xml)
    definitions = [m.group() for m in _element(formats.group(), 'xf').finditer(formats.group())]
    original_count = len(definitions)
    known = {value: index for index, value in enumerate(definitions)}
    ids = {}
    for first, last in blocks:
        for row in sheet.iter_rows(min_row=first, max_row=last):
            for cell in row:
                definition = definitions[originals.get(cell.coordinate, 0)]
                # Clone the actual XML format, changing only borderId. Some
                # Excel flags (e.g. applyNumberFormat) are lost on an openpyxl
                # round trip even though its style objects appear identical.
                if re.search(r'\bborderId=', definition):
                    definition = re.sub(r'\bborderId=(["\x27]).*?\1', f'borderId="{cell._style.borderId}"', definition, count=1)
                else:
                    definition = re.sub(r'(<(?:[\w.-]+:)?xf)\b', rf'\1 borderId="{cell._style.borderId}"', definition, count=1)
                if definition not in known:
                    known[definition] = len(definitions)
                    definitions.append(definition)
                ids[cell.coordinate] = known[definition]
    if len(definitions) == original_count:
        return ids, styles_xml, xml
    emitted = sync_generated_styles(book, styles_xml)
    emitted_formats = _element(emitted, 'cellXfs').search(emitted)
    block = formats.group()
    block = block[:block.rfind('</')] + ''.join(definitions[original_count:]) + block[block.rfind('</'):]
    block = re.sub(r'\bcount=(["\x27]).*?\1', f'count="{len(definitions)}"', block, count=1)
    return ids, emitted[:emitted_formats.start()] + block + emitted[emitted_formats.end():], xml


def _patch_cell_styles(xml, style_ids):
    """Change only cell style references; emit missing blank border-edge cells."""
    pending = {}
    for coordinate, style in style_ids.items():
        row_number = int(re.search(r'\d+$', coordinate).group())
        pending.setdefault(row_number, {})[coordinate] = style

    def patch_row(match):
        fragment = match.group()
        row_number = int(_attributes(fragment)['r'])
        row_styles = pending.pop(row_number, {})

        def patch_cell(cell_match):
            cell_xml = cell_match.group()
            coordinate = _attributes(cell_xml)['r']
            style = row_styles.pop(coordinate, None)
            if style is None:
                return cell_xml
            end = cell_xml.index('>')
            opening = cell_xml[:end]
            if re.search(r'\bs=', opening):
                opening = re.sub(r'\bs=(["\x27]).*?\1', f's="{style}"', opening, count=1)
            else:
                opening = opening.rstrip('/') + f' s="{style}"' + ('/' if opening.endswith('/') else '')
            return opening + cell_xml[end:]

        fragment = _element(fragment, 'c').sub(patch_cell, fragment)
        missing = row_styles
        if missing:
            if fragment.rstrip().endswith('/>'):
                fragment = fragment.rstrip()[:-2] + f'></{match.group("p")}row>'
            # Insert backwards so original fragments, including formulas and
            # cached values, remain byte-for-byte intact.
            for coord in sorted(missing, key=lambda c: column_index_from_string(re.match(r'[A-Z]+', c).group()), reverse=True):
                col = column_index_from_string(re.match(r'[A-Z]+', coord).group())
                following = next((c for c in _element(fragment, 'c').finditer(fragment)
                                  if column_index_from_string(re.match(r'[A-Z]+', _attributes(c.group())['r']).group()) > col), None)
                position = following.start() if following else fragment.rfind('</')
                fragment = fragment[:position] + f'<{match.group("p")}c r="{coord}" s="{missing.pop(coord)}"/>' + fragment[position:]
        return fragment

    xml = _element(xml, 'row').sub(patch_row, xml)
    if pending:
        section = _element(xml, 'sheetData').search(xml)
        rows = {int(_attributes(m.group())['r']): m.group() for m in _element(section.group(), 'row').finditer(section.group())}
        for row in pending:
            rows.setdefault(row, f'<{section.group("p")}row r="{row}"></{section.group("p")}row>')
        block = section.group()
        block = block[:block.index('>') + 1] + ''.join(rows[row] for row in sorted(rows)) + block[block.rfind('</'):]
        xml = xml[:section.start()] + block + xml[section.end():]
        return _patch_cell_styles(xml, {coord: style for cells in pending.values() for coord, style in cells.items()})
    return xml


def repair_register_borders(data, *, book=None, batches=None):
    """Repair client merges and borders without moving rows or changing cell values."""
    batches = read_batches(BytesIO(data)) if batches is None else batches
    owns_book = book is None
    if owns_book:
        book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, _ = select_register(book)
        if externally_protected(sheet):
            return data
        blocks = _inward_blocks(sheet, [b for b in batches if b['source_sheet'] == sheet.title])
        if not blocks:
            return data
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            unassigned = {issue['row'] for issue in getattr(batches, 'row_issues', ())}
            style_ids, styles, unmerged = _register_border_styles(
                book, sheet, blocks, xml, archive.read('xl/styles.xml').decode('utf-8'), unassigned)
            patched = _patch_cell_styles(unmerged, style_ids)
            if patched == xml:
                return data
            changes = {member: patched.encode('utf-8'),
                       'xl/styles.xml': styles.encode('utf-8')}
            if all(archive.read(name) == value for name, value in changes.items()):
                return data
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, changes.get(info.filename, archive.read(info.filename)))
            return output.getvalue()
    finally:
        if owns_book:
            book.close()


class UnsupportedReorderError(ValueError):
    """The register is usable, but an Excel feature cannot safely move here."""


def order_register_for_save(data, session, *, fit_inward=None):
    """Keep a verified record save usable when only automatic sorting is unsafe."""
    from register_column_order import normalize_register_columns, UnsupportedColumnOrderError
    try:
        data = normalize_register_columns(data)
    except UnsupportedColumnOrderError as exc:
        session.order_warning = str(exc)
    from register_presentation import normalize_register_presentation
    from register_rows import compact_register_gaps
    try:
        data = compact_register_gaps(data)
    except UnsupportedReorderError as exc:
        session.order_warning = str(exc)
    try:
        data = sort_register_by_inward_date(data)
    except UnsupportedReorderError as exc:
        session.order_warning = str(exc)
        data = repair_register_borders(data)
    return normalize_register_presentation(data, fit_inward=fit_inward)


def sort_register_by_inward_date(data, *, source_rows=None):
    """Sort whole inward blocks, preserving the XLSX package and record numbers.

    Only block-local cell formulas can travel automatically. Reject ambiguous
    dependencies rather than silently changing formulas, notes, or objects.
    Dates missing from the source stay at the end in their original order.
    """
    batches = read_batches(BytesIO(data))
    if not batches:
        return data
    # Sorting moves physical blocks, not Serial Number identities. Duplicates
    # elsewhere in the register must not prevent an unrelated inward save.
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header_row, cols = select_register(book)
        records = [batch for batch in batches if batch['source_sheet'] == sheet.title]
        blocks = _inward_blocks(sheet, records)
        by_start = {batch['source_row']: batch for batch in records}
        ordered = sorted(blocks, key=lambda block: (
            by_start[block[0]]['inward_date'] is None,
            by_start[block[0]]['inward_date'] or datetime.max.date()))
        moved = ordered != blocks
        if not moved and source_rows is None:
            return repair_register_borders(data, book=book, batches=batches)
        row_map = {}
        destination = blocks[0][0]
        for first, last in ordered:
            for old in range(first, last + 1):
                row_map[old] = destination
                destination += 1

        def unsafe(reason):
            raise UnsupportedReorderError('Automatic date sorting is unavailable because ' + reason +
                                          '. Row order was kept. Sort dates in Excel if needed.')

        def move_formula(formula, first, last):
            if not isinstance(formula, str):
                unsafe('a shared or array formula needs Excel to move it')
            try:
                tokens = Tokenizer(formula if formula.startswith('=') else '=' + formula).items
            except TokenizerError:
                unsafe('an Excel formula cannot be interpreted safely')
            result = []
            for token in tokens:
                value = token.value
                if (token.type == 'FUNC' and token.subtype == 'OPEN'
                        and value.rstrip('(').upper().split('.')[-1]
                        in {'INDIRECT', 'OFFSET', 'ADDRESS', 'ROW', 'COLUMN', 'CELL', 'HYPERLINK'}):
                    unsafe('a formula depends on physical row positions or constructs a cell reference from text')
                if token.type == 'OPERAND' and token.subtype == 'RANGE':
                    if not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?', value, re.I):
                        unsafe('a formula uses another sheet, a named range, or a table reference')
                    refs = re.findall(r'(\$?[A-Z]{1,3}\$?)(\d+)', value, re.I)
                    if any(not first <= int(row) <= last for _, row in refs):
                        unsafe('a formula refers to a different inward or a fixed worksheet row')
                    value = re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                                   lambda match: match[1] + str(row_map[int(match[2])]), value, flags=re.I)
                result.append(value)
            return '=' + ''.join(result)

        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            if moved:
                from register_rows import expanded_shared_formulas
                try:
                    shared_values = expanded_shared_formulas(xml, sheet)
                except ValueError as exc:
                    unsafe('a shared formula group could not be expanded safely')
                if any(name.casefold() == 'xl/vbaproject.bin' for name in archive.namelist()):
                    unsafe('the workbook contains VBA macros whose row references cannot be verified')
                if externally_protected(sheet):
                    unsafe('the register worksheet is protected')
                # These features carry row anchors/dependencies outside sheetData.
                # A normal plant register does not need any of them.
                if (sheet._charts or sheet._images or any(cell.comment or cell.hyperlink
                                                         for row in sheet for cell in row)):
                    unsafe('the register contains a drawing, comment, or hyperlink anchored to rows')
                if any(_element(xml, tag).search(xml) for tag in
                       ('drawing', 'legacyDrawing', 'oleObjects', 'controls', 'rowBreaks')):
                    unsafe('the register contains an object or page break anchored to rows')
                if any(name not in {'_xlnm.Print_Area', '_xlnm.Print_Titles'}
                       for name in list(book.defined_names) + list(sheet.defined_names)):
                    unsafe('the workbook contains a custom named range')
                for other in book:
                    if other.title == sheet.title:
                        continue
                    if any(cell.data_type == 'f' for row in other for cell in row):
                        unsafe('another worksheet contains formulas that may depend on these rows')
                covered = {row: (first, last) for first, last in blocks for row in range(first, last + 1)}
                for row in sheet:
                    for cell in row:
                        if cell.data_type == 'f':
                            if cell.row not in covered:
                                unsafe('a formula outside the inward blocks may depend on the register')
                            move_formula(shared_values.get(cell.coordinate, cell.value), *covered[cell.row])
                for area in sheet.merged_cells.ranges:
                    overlaps = area.max_row >= blocks[0][0] and area.min_row <= blocks[-1][1]
                    if overlaps and (area.min_row not in row_map or area.max_row not in row_map
                                     or covered[area.min_row] != covered[area.max_row]):
                        unsafe(f'merged cells {area} span different inward records')
                for row in range(blocks[0][0], blocks[-1][1] + 1):
                    serial = text(sheet.cell(row, cols['sn'] + 1).value)
                    if serial and row not in by_start and not any(row in batch['summary_rows'] for batch in records):
                        unsafe(f'Excel row {row} contains a note or unrecognised Serial Number')

            unassigned = {issue['row'] for issue in getattr(batches, 'row_issues', ())}
            style_ids, styles, xml = _register_border_styles(
                book, sheet, blocks, xml, archive.read('xl/styles.xml').decode('utf-8'), unassigned)
            xml = _patch_cell_styles(xml, style_ids)
            rows_section = _element(xml, 'sheetData').search(xml)
            rows = []
            for match in _element(rows_section.group(), 'row').finditer(rows_section.group()):
                fragment = match.group()
                old = int(_attributes(fragment)['r'])
                target = row_map.get(old, old)
                if old in row_map:
                    fragment = re.sub(r'\br=(["\x27])\d+\1', f'r="{target}"', fragment, count=1)
                    def relocate_cell(match):
                        cell_xml = match.group()
                        coordinate = _attributes(cell_xml)['r']
                        updated_coord = re.sub(r'\d+$', str(target), coordinate)
                        cell_xml = re.sub(r'\br=(["\x27]).*?\1', 'r=' + quoteattr(updated_coord), cell_xml, count=1)
                        opening_end = cell_xml.index('>')
                        opening = cell_xml[:opening_end]
                        if re.search(r'\bs=', opening):
                            opening = re.sub(r'\bs=(["\x27]).*?\1', f's="{style_ids.get(coordinate, 0)}"', opening)
                        else:
                            opening = opening.rstrip('/') + f' s="{style_ids.get(coordinate, 0)}"' + ('/' if opening.endswith('/') else '')
                        cell_xml = opening + cell_xml[opening_end:]
                        if moved:
                            def relocate_formula(formula_match):
                                formula_xml = formula_match.group()
                                formula_attrs = _attributes(formula_xml)
                                if formula_attrs:
                                    if formula_attrs.get('t') != 'shared':
                                        unsafe('an array or special formula needs Excel to move it')
                                    value = shared_values.get(coordinate)
                                    if not value:
                                        unsafe('a shared formula group could not be expanded safely')
                                    expanded = move_formula(value, *covered[old])[1:]
                                    prefix = formula_match.group('p')
                                    return f'<{prefix}f>{escape(expanded)}</{prefix}f>'
                                start, end = formula_xml.index('>') + 1, formula_xml.rfind('</')
                                value = move_formula(unescape(formula_xml[start:end]), *covered[old])[1:]
                                return formula_xml[:start] + escape(value) + formula_xml[end:]
                            cell_xml = _element(cell_xml, 'f').sub(relocate_formula, cell_xml)
                        return cell_xml
                    fragment = _element(fragment, 'c').sub(relocate_cell, fragment)
                rows.append((target, fragment))
            section = rows_section.group()
            section = section[:section.index('>') + 1] + ''.join(value for _, value in sorted(rows)) + section[section.rfind('</'):]
            xml = xml[:rows_section.start()] + section + xml[rows_section.end():]

            if moved:
                def relocate_merge(match):
                    fragment = match.group()
                    ref = _attributes(fragment)['ref']
                    left, first, right, last = range_boundaries(ref)
                    if first not in row_map:
                        return fragment
                    ref = f'{get_column_letter(left)}{row_map[first]}:{get_column_letter(right)}{row_map[last]}'
                    return re.sub(r'\bref=(["\x27]).*?\1', 'ref=' + quoteattr(ref), fragment, count=1)
                xml = _element(xml, 'mergeCell').sub(relocate_merge, xml)
                # Whole-register rules remain positional. Rules confined to one
                # inward move with it. Partial ranges spanning blocks are unsafe.
                def relocate_rule(match):
                    fragment = match.group()
                    refs = _attributes(fragment).get('sqref', '')
                    relocated = []
                    local_blocks = set()
                    stationary = False
                    for ref in refs.split():
                        left, first, right, last = range_boundaries(ref)
                        if first not in row_map and (last is None or last < blocks[0][0]
                                                     or first > blocks[-1][1]):
                            relocated.append(ref)
                            stationary = True
                        elif first is not None and last is not None and first <= blocks[0][0] and last >= blocks[-1][1]:
                            relocated.append(ref)
                            stationary = True
                        elif first in row_map and last in row_map and covered[first] == covered[last]:
                            relocated.append(f'{get_column_letter(left)}{row_map[first]}:{get_column_letter(right)}{row_map[last]}')
                            local_blocks.add(covered[first])
                        else:
                            unsafe('an Excel formatting or validation range spans parts of different inwards')
                    if local_blocks and (stationary or len(local_blocks) != 1):
                        unsafe('an Excel rule mixes different inward blocks')
                    fragment = _range_fragment(fragment, ' '.join(relocated))
                    if local_blocks:
                        first, last = next(iter(local_blocks))
                        def move_rule_formula(match):
                            block = match.group()
                            start, end = block.index('>') + 1, block.rfind('</')
                            value = unescape(block[start:end])
                            value = move_formula(value, first, last)
                            return block[:start] + escape(value[1:]) + block[end:]
                        for tag in ('formula', 'formula1', 'formula2'):
                            fragment = _element(fragment, tag).sub(move_rule_formula, fragment)
                    return fragment
                for tag in ('conditionalFormatting', 'dataValidation'):
                    xml = _element(xml, tag).sub(relocate_rule, xml)
            ET.fromstring(xml)
            changes = {member: xml.encode('utf-8'), 'xl/styles.xml': styles.encode('utf-8')}
            # Excel rebuilds its calculation cache after formula cells move.
            # Keeping its old row references would leave a stale calc chain.
            if moved and 'xl/calcChain.xml' in archive.namelist():
                for filename, tag, attribute, value in (
                    ('xl/_rels/workbook.xml.rels', 'Relationship', 'Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain'),
                    ('[Content_Types].xml', 'Override', 'PartName', '/xl/calcChain.xml'),
                ):
                    source_xml = archive.read(filename).decode('utf-8')
                    source_xml = _element(source_xml, tag).sub(
                        lambda match: '' if _attributes(match.group()).get(attribute) == value else match.group(), source_xml)
                    changes[filename] = source_xml.encode('utf-8')
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    if moved and info.filename == 'xl/calcChain.xml':
                        continue
                    target.writestr(info, changes.get(info.filename, archive.read(info.filename)))
            updated = output.getvalue()
        after = read_batches(BytesIO(updated))
        relocated_rows = {(sheet.title, old): new for old, new in row_map.items()}
        # Verify every original block at its new position. Keying this check by
        # Serial Number would collapse two manually duplicated records.
        before_content = {(batch['source_sheet'], relocated_rows.get(
            (batch['source_sheet'], batch['source_row']), batch['source_row'])):
            (batch_content(batch), batch['cod_no'], batch['cod_date'], batch['issued']) for batch in batches}
        after_content = {(batch['source_sheet'], batch['source_row']):
            (batch_content(batch), batch['cod_no'], batch['cod_date'], batch['issued']) for batch in after}
        if before_content != after_content:
            raise ValueError('The inward order could not be verified; the original workbook was kept.')
        if source_rows is not None:
            source_rows.update(relocated_rows)
        return updated
    finally:
        book.close()


def sync_generated_styles(book, styles_xml):
    """Save styles synthesized when openpyxl reads merged border edges.

    Style IDs registered by Cell.style_id may exceed the original XLSX style
    table. Append their definitions before the app writes cells using those IDs.
    Original style IDs and definitions remain untouched.
    """
    generated = tostring(write_stylesheet(book)).decode('utf-8')
    for section_name, child_name in (('borders', 'border'), ('cellXfs', 'xf')):
        original = _element(styles_xml, section_name).search(styles_xml)
        emitted = _element(generated, section_name).search(generated)
        if original is None or emitted is None:
            raise ValueError('Workbook formatting is missing; the file was not changed.')
        before = list(_element(original.group(), child_name).finditer(original.group()))
        after = list(_element(emitted.group(), child_name).finditer(emitted.group()))
        extras = ''.join(item.group() for item in after[len(before):])
        if not extras:
            continue
        prefix = original.group('p')
        if prefix:
            extras = re.sub(r'<(/?)([A-Za-z_][\w.-]*)(?=[\s/>])',
                            lambda m:'<' + m[1] + prefix + m[2], extras)
        block = original.group()
        position = block.rfind('</')
        block = block[:position] + extras + block[position:]
        block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{len(after)}"', block, count=1)
        styles_xml = styles_xml[:original.start()] + block + styles_xml[original.end():]
    return styles_xml


def _check_batch_merge(sheet, cell, source_row, cols):
    """Only the top cell of a vertical merge confined to this inward is writable."""
    from certificate_core import norm
    summary_labels = {'total', 'subtotal', 'grandtotal', 'totalweight', 'totalweightkg', 'totalweightinkg'}

    def starts_inward(row):
        serial = text(sheet.cell(row, cols['sn'] + 1).value)
        if not serial:
            return False
        item = text(sheet.cell(row, cols['item'] + 1).value) if 'item' in cols else ''
        has_primary_facts = any(text(sheet.cell(row, cols[key] + 1).value)
                                for key in ('supplier', 'inward_date') if key in cols)
        if (norm(serial) in summary_labels or norm(item) in summary_labels) and not has_primary_facts:
            return False
        if re.match(r'^(?:note|notes|remark|remarks)\s*:', serial, re.I) and not any(
                text(sheet.cell(row, column + 1).value) for key, column in cols.items() if key != 'sn'):
            return False
        return True

    for area in sheet.merged_cells.ranges:
        if cell.coordinate not in area:
            continue
        next_inward = any(starts_inward(row) for row in range(source_row + 1, area.max_row + 1))
        if (area.min_col != area.max_col or area.min_row != source_row
                or cell.row != area.min_row or cell.column != area.min_col
                or next_inward):
            raise ValueError(
                f'Excel cells {area} are shared across columns or inward records. '
                'Keep the correction in the PDF only, or separate those cells in Excel before saving there.'
            )


class LockBusyError(PermissionError):
    """A real file-sharing/lock conflict, distinct from missing write access."""


@contextmanager
def exclusive_file(path, *, wait_timeout=5):
    """Windows denies other readers/writers but permits our atomic replacement."""
    path = Path(path)
    if os.name == "nt":
        import ctypes
        import msvcrt
        from ctypes import wintypes
        api = ctypes.WinDLL("kernel32", use_last_error=True)
        api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        api.CreateFileW.restype = wintypes.HANDLE
        # Windows reports a sharing violation while Excel, Explorer's preview
        # handler, or a sync/backup scanner briefly has the file open. Give
        # short-lived system handles time to clear before reporting a lock.
        deadline = time.monotonic() + max(0, wait_timeout)
        while True:
            handle = api.CreateFileW(str(path.resolve()), 0xC0000000, 4, None, 3, 0x80, None)
            if handle != ctypes.c_void_p(-1).value:
                break
            code = ctypes.get_last_error()
            if code not in (32, 33) or time.monotonic() >= deadline:
                if code in (32, 33):
                    raise LockBusyError(
                        f"Windows still reports {path.name} in use after waiting {wait_timeout:g} seconds. "
                        "Excel may be closed while another app, preview pane, sync, or backup service holds it. "
                        f"Close that handle or wait for it to finish, then retry. (Windows error {code})"
                    )
                raise PermissionError(f"Cannot update {path.name}. Check that the workbook exists and is writable. (Windows error {code})")
            time.sleep(0.2)
        stream = os.fdopen(msvcrt.open_osfhandle(handle, os.O_RDWR | os.O_BINARY), "r+b")
    else:
        import fcntl
        stream = path.open("r+b")
        try:
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            stream.close()
            if exc.errno in (errno.EAGAIN, errno.EWOULDBLOCK):
                raise LockBusyError(f"Cannot update {path.name}; it is already locked by another process.") from exc
            raise
        except Exception:
            stream.close()
            raise
    try:
        yield stream
    finally:
        stream.close()


def _element(xml, name):
    return re.compile(rf'<(?P<p>(?:[A-Za-z_][\w.-]*:)?)' + name + r'\b[^>]*?(?:/>|>.*?</(?P=p)' + name + r'\s*>)', re.S)


def _attributes(fragment):
    opening = fragment[:fragment.index(">")]
    return {match[0]: unescape(match[2]) for match in re.findall(r'([\w:.-]+)\s*=\s*(["\x27])(.*?)\2', opening)}


def checked_excel_text(value, label, *, normalize_paste=False):
    """Validate XML text before a pasted field reaches Excel serialization."""
    value = str(value)
    if normalize_paste:
        value = value.replace('\v', ' ').replace('\f', ' ')
    if any((ord(character) < 32 and character not in '\t\n\r')
           or 0xD800 <= ord(character) <= 0xDFFF or ord(character) in {0xFFFE, 0xFFFF}
           for character in value):
        raise ValueError(f'{label} contains an unsupported control character. '
                         'Paste plain text or remove that character before saving.')
    return value


def _set_cell(xml, coordinate, value, kind, style=None):
    checked_register_column(column_index_from_string(re.match(r'[A-Z]+', coordinate).group()))
    if kind == 'inlineStr':
        value = checked_excel_text(value, f'Excel cell {coordinate}')
    row_number = re.search(r"\d+$", coordinate).group()
    row_match = next((m for m in _element(xml, "row").finditer(xml) if _attributes(m.group()).get("r") == row_number), None)
    if not row_match:
        raise ValueError(f"Workbook row {row_number} is missing; no cells were changed.")
    row = row_match.group()
    if row.rstrip().endswith('/>'):
        row = row.rstrip()[:-2] + f'></{row_match.group("p")}row>'
    cell_matches = list(_element(row, "c").finditer(row))
    match = next((m for m in cell_matches if _attributes(m.group()).get("r") == coordinate), None)
    attrs = _attributes(match.group()) if match else {"r": coordinate}
    if kind == "blank":
        attrs.pop("t", None)
    else:
        attrs["t"] = kind
    if style is not None:
        attrs["s"] = str(style)
    prefix = match.group("p") if match else row_match.group("p")
    payload = "" if kind == "blank" else f"<{prefix}is><{prefix}t>{escape(str(value))}</{prefix}t></{prefix}is>" if kind == "inlineStr" else f"<{prefix}v>{value}</{prefix}v>"
    replacement = "<" + prefix + "c " + " ".join(f"{k}={quoteattr(v)}" for k, v in attrs.items()) + ">" + payload + f"</{prefix}c>"
    if match:
        row = row[:match.start()] + replacement + row[match.end():]
    else:
        col = column_index_from_string(re.match(r"[A-Z]+", coordinate).group())
        following = next((m for m in cell_matches if column_index_from_string(re.match(r"[A-Z]+", _attributes(m.group())["r"]).group()) > col), None)
        position = following.start() if following else row.rfind("</")
        row = row[:position] + replacement + row[position:]
    return xml[:row_match.start()] + row + xml[row_match.end():]


def _without_column(sqref, column):
    """Split multi-column ranges without losing cells outside the managed column."""
    target = column_index_from_string(column)
    kept = []
    for reference in sqref.split():
        match = re.fullmatch(r'\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?', reference, re.I)
        if not match:
            raise ValueError(f'Cannot safely preserve Excel rule range {reference!r}; the workbook was not changed.')
        first, last = column_index_from_string(match[1].upper()), column_index_from_string((match[3] or match[1]).upper())
        top, bottom = int(match[2]), int(match[4] or match[2])
        if not first <= target <= last:
            kept.append(reference)
            continue
        for left, right in ((first, target - 1), (target + 1, last)):
            if left <= right:
                start = f'{get_column_letter(left)}{top}'
                end = f'{get_column_letter(right)}{bottom}'
                kept.append(start if start == end else f'{start}:{end}')
    return ' '.join(kept)


def _range_fragment(fragment, refs, original_refs=None):
    """Keep relative formulas anchored to the same cells after range splitting."""
    fragment = re.sub(r'\bsqref=(["\']).*?\1', 'sqref=' + quoteattr(refs), fragment, count=1)
    if original_refs:
        origin = original_refs.split()[0].split(':')[0].replace('$', '')
        destination = refs.split()[0].split(':')[0].replace('$', '')
        if origin != destination:
            def move_formula(match):
                block = match.group()
                start, end = block.index('>') + 1, block.rfind('</')
                value = unescape(block[start:end])
                if value and not value.startswith('"'):
                    try:
                        translated = Translator(value if value.startswith('=') else '=' + value,
                                                origin=origin).translate_formula(destination)
                    except TranslatorError as exc:
                        raise ValueError('An Excel rule cannot be moved safely; the workbook was not changed.') from exc
                    value = translated if value.startswith('=') else translated[1:]
                return block[:start] + escape(value) + block[end:]
            for name in ('formula', 'formula1', 'formula2'):
                fragment = _element(fragment, name).sub(move_formula, fragment)
    return fragment


def _ensure_status_validation(xml, status_column, row_limit, first_row=2):
    """Give the entire status column the same simple, complete dropdown."""
    def includes_column(sqref):
        target = column_index_from_string(status_column)
        for reference in sqref.split():
            area = re.fullmatch(r"\$?([A-Z]+)\$?\d+(?::\$?([A-Z]+)\$?\d+)?", reference, re.I)
            if area:
                first = column_index_from_string(area.group(1).upper())
                last = column_index_from_string((area.group(2) or area.group(1)).upper())
                if min(first, last) <= target <= max(first, last):
                    return True
        return False

    match = _element(xml, "dataValidations").search(xml)
    root_prefix_match = re.search(r"<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b", xml)
    prefix = match.group("p") if match else (root_prefix_match.group("p") if root_prefix_match else "")
    validation = (f'<{prefix}dataValidation type="list" allowBlank="1" showErrorMessage="1" '
                  f'sqref="{status_column}{first_row}:{status_column}{row_limit}">'
                  f'<{prefix}formula1>",Ready,Alert,N/A,Generated,Completed"</{prefix}formula1>'
                  f'</{prefix}dataValidation>')
    if match:
        block = match.group()
        validations = list(_element(block, "dataValidation").finditer(block))
        remaining = []
        inserted = False
        for item in validations:
            fragment = item.group()
            if includes_column(_attributes(fragment).get("sqref", "")):
                original_refs = _attributes(fragment)['sqref']
                retained_refs = _without_column(original_refs, status_column)
                if retained_refs:
                    remaining.append(_range_fragment(fragment, retained_refs, original_refs))
                if not inserted:
                    remaining.append(validation)
                    inserted = True
            else:
                remaining.append(fragment)
        if not inserted:
            remaining.append(validation)
        opening = block[:block.index(">") + 1]
        if opening.rstrip().endswith("/>"):
            opening = opening.rstrip()[:-2] + ">"
        closing = f"</{prefix}dataValidations>"
        new_block = opening + "".join(remaining) + closing
        new_block = re.sub(r'\bcount=["\']\d+["\']', f'count="{len(remaining)}"', new_block, count=1)
        return xml[:match.start()] + new_block + xml[match.end():]
    insert_at = len(xml)
    for name in ("hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"):
        found = _element(xml, name).search(xml)
        if found:
            insert_at = min(insert_at, found.start())
    if insert_at == len(xml):
        insert_at = xml.rfind("</")
    return xml[:insert_at] + f'<{prefix}dataValidations count="1">{validation}</{prefix}dataValidations>' + xml[insert_at:]


def _update_issued_status_formatting(xml, coordinate, first_row=2):
    """Keep old plant-workbook dropdowns compatible with the PDF workflow."""
    match = re.fullmatch(r"([A-Z]+)(\d+)", coordinate)
    if not match:
        return xml
    status_column = match.group(1)
    xml = _ensure_status_validation(xml, status_column, 10000, first_row)

    def includes_status_column(sqref):
        target = column_index_from_string(status_column)
        for reference in sqref.split():
            area = re.fullmatch(r"\$?([A-Z]+)\$?\d+(?::\$?([A-Z]+)\$?\d+)?", reference, re.I)
            if area:
                first = column_index_from_string(area.group(1).upper())
                last = column_index_from_string((area.group(2) or area.group(1)).upper())
                if min(first, last) <= target <= max(first, last):
                    return True
        return False

    formula_pattern = re.compile(
        r"(<(?:[A-Za-z_][\w.-]*:)?formula\b[^>]*>)(.*?)(</(?:[A-Za-z_][\w.-]*:)?formula\s*>)",
        flags=re.DOTALL,
    )
    def rewrite_conditional_formatting(block_match):
        block = block_match.group()
        attributes = _attributes(block)
        if not includes_status_column(attributes.get("sqref", "")):
            return block
        if _without_column(attributes.get('sqref', ''), status_column):
            return block  # Broad user rules retain their original formulas.

        def rewrite_formula(formula_match):
            formula = formula_match.group(2)
            if not re.search(rf"\$?{re.escape(status_column)}\$?\d+", formula, flags=re.IGNORECASE):
                return formula_match.group()

            for old, new in (("DRAFTED", "GENERATED"), ("CREATED", "GENERATED"),
                             ("SENT", "COMPLETED"), ("CHECK DETAILS", "ALERT"),
                             ("WORKING DRAFT", "ALERT")):
                formula = re.sub(r'(["\'])' + re.escape(old) + r'(["\'])',
                                 lambda match: match.group(1) + new + match.group(2),
                                 formula, flags=re.IGNORECASE)
            return formula_match.group(1) + formula + formula_match.group(3)

        return formula_pattern.sub(rewrite_formula, block)

    return _element(xml, "conditionalFormatting").sub(rewrite_conditional_formatting, xml)


def _add_alert_dxf(styles_xml):
    """Append a shared red/white differential style for overdue client cells."""
    match = _element(styles_xml, "dxfs").search(styles_xml)
    root_match = re.search(r"<(?P<p>(?:[A-Za-z_][\w.-]*:)?)styleSheet\b", styles_xml)
    prefix = match.group("p") if match else (root_match.group("p") if root_match else "")
    block = match.group() if match else ""
    dxf_matches = list(_element(block, "dxf").finditer(block)) if block else []
    for index, item in enumerate(dxf_matches):
        if _dxf_matches_colors(item.group(), "FFC00000", "FFFFFFFF"):
            return styles_xml, index
    dxf = (f'<{prefix}dxf><{prefix}font><{prefix}b/><{prefix}color rgb="FFFFFFFF"/>'
           f'</{prefix}font><{prefix}fill><{prefix}patternFill patternType="solid">'
           f'<{prefix}fgColor rgb="FFC00000"/><{prefix}bgColor rgb="FFC00000"/>'
           f'</{prefix}patternFill></{prefix}fill></{prefix}dxf>')
    if not match:
        block = f'<{prefix}dxfs count="1">{dxf}</{prefix}dxfs>'
        insert_at = len(styles_xml)
        for name in ("tableStyles", "colors", "extLst"):
            found = _element(styles_xml, name).search(styles_xml)
            if found:
                insert_at = min(insert_at, found.start())
        if insert_at == len(styles_xml):
            insert_at = styles_xml.rfind("</")
        return styles_xml[:insert_at] + block + styles_xml[insert_at:], 0
    if block.rstrip().endswith("/>"):
        block = block.rstrip()[:-2] + "></" + prefix + "dxfs>"
    close = block.rfind("</")
    block = block[:close] + dxf + block[close:]
    block = re.sub(r'\bcount=["\']\d+["\']', f'count="{len(dxf_matches) + 1}"', block, count=1)
    return styles_xml[:match.start()] + block + styles_xml[match.end():], len(dxf_matches)


def _add_color_dxf(styles_xml, fill_rgb, font_rgb="FF000000"):
    """Append/reuse a status style with explicit, matching fill colors."""
    match = _element(styles_xml, "dxfs").search(styles_xml)
    if not match:
        raise ValueError("Workbook styles could not store the status colors.")
    block = match.group()
    items = list(_element(block, "dxf").finditer(block))
    for index, item in enumerate(items):
        if _dxf_matches_colors(item.group(), fill_rgb, font_rgb):
            return styles_xml, index
    prefix = match.group("p")
    dxf = (f'<{prefix}dxf><{prefix}font><{prefix}b/><{prefix}color rgb="{font_rgb}"/>'
           f'</{prefix}font><{prefix}fill><{prefix}patternFill patternType="solid">'
           f'<{prefix}fgColor rgb="{fill_rgb}"/><{prefix}bgColor rgb="{fill_rgb}"/>'
           f'</{prefix}patternFill></{prefix}fill></{prefix}dxf>')
    close = block.rfind("</")
    block = block[:close] + dxf + block[close:]
    block = re.sub(r'\bcount=["\']\d+["\']', f'count="{len(items) + 1}"', block, count=1)
    return styles_xml[:match.start()] + block + styles_xml[match.end():], len(items)


def _dxf_matches_colors(fragment, fill_rgb, font_rgb):
    """Only reuse a dxf when Excel has an explicit fill color on both channels."""
    try:
        root = ET.fromstring(fragment)
    except ET.ParseError:
        return False
    parts = {element.tag.rsplit("}", 1)[-1]: element for element in root.iter()}
    fill = parts.get("patternFill")
    font = parts.get("font")
    if fill is None or font is None:
        return False
    fill_colors = {element.tag.rsplit("}", 1)[-1]: element.attrib.get("rgb", "").upper()
                   for element in fill}
    font_color = next((element.attrib.get("rgb", "").upper()
                       for element in font if element.tag.rsplit("}", 1)[-1] == "color"), "")
    wanted_fill = fill_rgb.upper()
    return (fill.attrib.get("patternType", "").casefold() == "solid"
            and fill_colors.get("fgColor") == wanted_fill
            and fill_colors.get("bgColor") == wanted_fill
            and font_color == font_rgb.upper())


def _status_column_is_in_range(sqref, status_column):
    target = column_index_from_string(status_column)
    for reference in sqref.split():
        area = re.fullmatch(r"\$?([A-Z]+)\$?\d+(?::\$?([A-Z]+)\$?\d+)?", reference, re.I)
        if area:
            first = column_index_from_string(area.group(1).upper())
            last = column_index_from_string((area.group(2) or area.group(1)).upper())
            if min(first, last) <= target <= max(first, last):
                return True
    return False


def _is_status_color_rule(fragment, status_column):
    formula = unescape(fragment).upper()
    if "ECORECO_STATUS_COLOR" in formula:
        return True
    if not re.search(rf"\$?{re.escape(status_column)}\$?\d+", formula):
        return False
    return any(f'"{status}"' in formula for status in (
        "READY", "ALERT", "CHECK DETAILS", "WORKING DRAFT", "N/A", "BLANK",
        "GENERATED", "GENERATED MANUAL", "GENERATED EXTERNALLY", "CREATED", "DRAFTED", "COMPLETED", "SENT",
    ))


def _replace_status_color_rules(sheet_xml, status_column):
    """Remove app/legacy status-color rules while preserving unrelated Excel rules."""
    def clean(block_match):
        block = block_match.group()
        original_refs = _attributes(block).get('sqref', '')
        if not _status_column_is_in_range(original_refs, status_column):
            return block
        rules = list(_element(block, "cfRule").finditer(block))
        remove = [rule for rule in rules if _is_status_color_rule(rule.group(), status_column)]
        if not remove:
            return block
        retained_refs = _without_column(original_refs, status_column)
        retained_block = ''
        if retained_refs:
            retained_block = block
            for rule in reversed(rules):
                if rule not in remove:
                    retained_block = retained_block[:rule.start()] + retained_block[rule.end():]
            retained_block = _range_fragment(retained_block, retained_refs, original_refs)
        for rule in reversed(remove):
            block = block[:rule.start()] + block[rule.end():]
        if not _element(block, "cfRule").search(block):
            return retained_block
        return block + retained_block

    return _element(sheet_xml, "conditionalFormatting").sub(clean, sheet_xml)


def _renumber_conditional_priorities(sheet_xml):
    """Give every remaining conditional-formatting rule a unique Excel priority."""
    priority = 0
    pattern = re.compile(r"(<(?:[A-Za-z_][\w.-]*:)?cfRule\b[^>]*\bpriority=[\"'])(\d+)([\"'])")

    def renumber(match):
        nonlocal priority
        priority += 1
        return f"{match.group(1)}{priority}{match.group(3)}"

    return pattern.sub(renumber, sheet_xml)


def _add_status_conditional_formatting(sheet_xml, status_column, sn_column, dxf_ids, row_limit, first_row=2):
    """Keep one simple, authoritative color rule per Certificate Issued value."""
    sheet_xml = _replace_status_color_rules(sheet_xml, status_column)
    values = (
        ("READY", f'UPPER(TRIM(${status_column}2))="READY"'),
        ("ALERT", f'UPPER(TRIM(${status_column}2))="ALERT"'),
        ("NA", f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")="N/A"'),
        ("BLANK", f'AND(${sn_column}2<>"",OR(${status_column}2="",UPPER(TRIM(${status_column}2))="BLANK"))'),
        ("GENERATED", f'UPPER(TRIM(${status_column}2))="GENERATED"'),
        ("GENERATED EXTERNALLY", f'UPPER(TRIM(${status_column}2))="GENERATED EXTERNALLY"'),
        ("COMPLETED", f'UPPER(TRIM(${status_column}2))="COMPLETED"'),
    )
    prefix_match = re.search(r"<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b", sheet_xml)
    prefix = prefix_match.group("p") if prefix_match else ""
    rules = []
    for label, formula in values:
        if first_row != 2:
            formula = re.sub(r'(\$[A-Z]+)2\b', lambda match: match[1] + str(first_row), formula)
        rules.append(
            f'<{prefix}cfRule type="expression" dxfId="{dxf_ids[label]}" '
            f'priority="1" stopIfTrue="1"><{prefix}formula>{escape(formula)}'
            f'</{prefix}formula></{prefix}cfRule>'
        )
    block = (f'<{prefix}conditionalFormatting sqref="{status_column}{first_row}:{status_column}{row_limit}">'
             + "".join(rules) + f'</{prefix}conditionalFormatting>')
    existing = _element(sheet_xml, "conditionalFormatting").search(sheet_xml)
    insert_at = existing.start() if existing else len(sheet_xml)
    if not existing:
        for name in ("dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"):
            found = _element(sheet_xml, name).search(sheet_xml)
            if found:
                insert_at = min(insert_at, found.start())
        if insert_at == len(sheet_xml):
            insert_at = sheet_xml.rfind("</")
    return _renumber_conditional_priorities(sheet_xml[:insert_at] + block + sheet_xml[insert_at:])


def _add_alert_conditional_formatting(sheet_xml, *, sn_column, date_column,
                                      status_column, overdue_columns,
                                      date_attention_columns, overdue_dxf_id,
                                      attention_dxf_id, row_limit, first_row=2):
    """Show overdue and date-quality alerts in red beside the received date."""
    if not overdue_columns or not date_attention_columns:
        return sheet_xml
    overdue_refs = " ".join(f"{column}{first_row}:{column}{row_limit}" for column in overdue_columns)
    attention_refs = " ".join(f"{column}{first_row}:{column}{row_limit}" for column in date_attention_columns)
    date_ref = f"${date_column}2"
    trimmed_date = f"TRIM({date_ref})"
    parsed_text_date = f"DATE(VALUE(RIGHT({trimmed_date},4)),VALUE(MID({trimmed_date},4,2)),VALUE(LEFT({trimmed_date},2)))"
    parsed_iso_text_date = f"DATE(VALUE(LEFT({trimmed_date},4)),VALUE(MID({trimmed_date},6,2)),VALUE(RIGHT({trimmed_date},2)))"
    short_year = (f'IF(VALUE(RIGHT({trimmed_date},2))>=69,'
                  f'1900+VALUE(RIGHT({trimmed_date},2)),2000+VALUE(RIGHT({trimmed_date},2)))')
    parsed_short_text_date = f"DATE({short_year},VALUE(MID({trimmed_date},4,2)),VALUE(LEFT({trimmed_date},2)))"
    separators = (f'OR(MID({trimmed_date},3,1)=".",MID({trimmed_date},3,1)="/",'
                  f'MID({trimmed_date},3,1)="-")')
    matching_separator = f'MID({trimmed_date},6,1)=MID({trimmed_date},3,1)'
    full_date_is_valid = (
        f'AND(DAY({parsed_text_date})=VALUE(LEFT({trimmed_date},2)),'
        f'MONTH({parsed_text_date})=VALUE(MID({trimmed_date},4,2)),'
        f'YEAR({parsed_text_date})=VALUE(RIGHT({trimmed_date},4)))'
    )
    short_date_is_valid = (
        f'AND(DAY({parsed_short_text_date})=VALUE(LEFT({trimmed_date},2)),'
        f'MONTH({parsed_short_text_date})=VALUE(MID({trimmed_date},4,2)),'
        f'YEAR({parsed_short_text_date})={short_year})'
    )
    iso_date_is_valid = (
        f'AND(DAY({parsed_iso_text_date})=VALUE(RIGHT({trimmed_date},2)),'
        f'MONTH({parsed_iso_text_date})=VALUE(MID({trimmed_date},6,2)),'
        f'YEAR({parsed_iso_text_date})=VALUE(LEFT({trimmed_date},4)))'
    )
    parsed_mdy = f'DATE(VALUE(RIGHT({trimmed_date},4)),VALUE(LEFT({trimmed_date},2)),VALUE(MID({trimmed_date},4,2)))'
    mdy_is_valid = (f'AND(VALUE(MID({trimmed_date},4,2))>12,'
                    f'DAY({parsed_mdy})=VALUE(MID({trimmed_date},4,2)),'
                    f'MONTH({parsed_mdy})=VALUE(LEFT({trimmed_date},2)),'
                    f'YEAR({parsed_mdy})=VALUE(RIGHT({trimmed_date},4)))')
    mdy_overdue = f'IFERROR(IF({mdy_is_valid},AND({parsed_mdy}<=TODAY(),TODAY()-{parsed_mdy}>=25),FALSE),FALSE)'
    mdy_attention = f'IFERROR(IF({mdy_is_valid},{parsed_mdy}>TODAY(),TRUE),TRUE)'
    # English month names are ordinary Excel date text. DATEVALUE is only a
    # display rule; the source parser independently validates every issued date.
    named_overdue = (f'IFERROR(IF(ISNUMBER(SEARCH(" ",{trimmed_date})),AND(DATEVALUE({trimmed_date})<=TODAY(),'
                     f'TODAY()-DATEVALUE({trimmed_date})>=25),FALSE),FALSE)')
    named_attention = f'IFERROR(IF(ISNUMBER(SEARCH(" ",{trimmed_date})),DATEVALUE({trimmed_date})>TODAY(),TRUE),TRUE)'
    full_text_date_overdue = (
        f'IFERROR(IF({full_date_is_valid},AND({parsed_text_date}<=TODAY(),'
        f'TODAY()-{parsed_text_date}>=25),{mdy_overdue}),FALSE)'
    )
    short_text_date_overdue = (
        f'IFERROR(IF({short_date_is_valid},AND({parsed_short_text_date}<=TODAY(),'
        f'TODAY()-{parsed_short_text_date}>=25),FALSE),FALSE)'
    )
    iso_text_date_overdue = (
        f'IFERROR(IF({iso_date_is_valid},AND({parsed_iso_text_date}<=TODAY(),'
        f'TODAY()-{parsed_iso_text_date}>=25),FALSE),FALSE)'
    )
    dmy_full_shape = f'AND(LEN({trimmed_date})=10,{separators},{matching_separator})'
    dmy_short_shape = f'AND(LEN({trimmed_date})=8,{separators},{matching_separator})'
    iso_shape = (f'AND(LEN({trimmed_date})=10,MID({trimmed_date},5,1)="-",'
                 f'MID({trimmed_date},8,1)="-")')
    # Choose the matching shape before parsing. Keep data-quality attention
    # separate from the 25-day overdue color so future dates are not mistaken
    # for overdue receipts. Invalid calendar dates remain attention items.
    text_date_overdue = (
        f'IF({trimmed_date}="",FALSE,IF(LEN({trimmed_date})=10,'
        f'IF({iso_shape},{iso_text_date_overdue},IF({dmy_full_shape},{full_text_date_overdue},{named_overdue})),'
        f'IF(LEN({trimmed_date})=8,IF({dmy_short_shape},{short_text_date_overdue},{named_overdue}),{named_overdue})))'
    )
    text_date_attention = (
        f'IF({trimmed_date}="",TRUE,IF(LEN({trimmed_date})=10,'
        f'IF({iso_shape},IFERROR(OR(NOT({iso_date_is_valid}),{parsed_iso_text_date}>TODAY()),TRUE),'
        f'IF({dmy_full_shape},IFERROR(IF({full_date_is_valid},{parsed_text_date}>TODAY(),{mdy_attention}),TRUE),{named_attention})),'
        f'IF(LEN({trimmed_date})=8,IF({dmy_short_shape},'
        f'IFERROR(OR(NOT({short_date_is_valid}),{parsed_short_text_date}>TODAY()),TRUE),{named_attention}),{named_attention})))'
    )
    overdue = (
        f'IF(ISNUMBER({date_ref}),AND(INT({date_ref})<=TODAY(),'
        f'TODAY()-INT({date_ref})>=25),{text_date_overdue})'
    )
    attention = f'IF(ISNUMBER({date_ref}),INT({date_ref})>TODAY(),{text_date_attention})'
    unfinished = (
        f'UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    overdue_formula = f'AND(${sn_column}2<>"",{overdue},{unfinished}'
    attention_formula = (
        f'AND(${sn_column}2<>"",{attention},'
        f'UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    previous_full_text_date_due = (
        f'IFERROR(IF({full_date_is_valid},OR({parsed_text_date}>TODAY(),'
        f'TODAY()-{parsed_text_date}>=25),TRUE),TRUE)'
    )
    previous_short_text_date_due = (
        f'IFERROR(IF({short_date_is_valid},OR({parsed_short_text_date}>TODAY(),'
        f'TODAY()-{parsed_short_text_date}>=25),TRUE),TRUE)'
    )
    previous_iso_text_date_due = (
        f'IFERROR(IF({iso_date_is_valid},OR({parsed_iso_text_date}>TODAY(),'
        f'TODAY()-{parsed_iso_text_date}>=25),TRUE),TRUE)'
    )
    previous_text_date_due = (
        f'IF({trimmed_date}="",TRUE,IF(LEN({trimmed_date})=10,'
        f'IF({iso_shape},{previous_iso_text_date_due},'
        f'IF({dmy_full_shape},{previous_full_text_date_due},TRUE)),'
        f'IF(LEN({trimmed_date})=8,IF({dmy_short_shape},'
        f'{previous_short_text_date_due},TRUE),TRUE)))'
    )
    previous_age_due = (
        f'IF(ISNUMBER({date_ref}),OR(INT({date_ref})>TODAY(),'
        f'TODAY()-INT({date_ref})>=25),{previous_text_date_due})'
    )
    previous_formula = (
        f'AND(${sn_column}2<>"",{previous_age_due},'
        f'UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    legacy_formula = (
        f'AND(${sn_column}2<>"",OR(NOT(ISNUMBER({date_ref})),{date_ref}>TODAY(),'
        f'TODAY()-{date_ref}>=25),UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    managed_formulas = {overdue_formula, attention_formula, previous_formula, legacy_formula}
    if first_row != 2:
        move = lambda formula: re.sub(r'(\$[A-Z]+)2\b', lambda match: match[1] + str(first_row), formula)
        overdue_formula, attention_formula = move(overdue_formula), move(attention_formula)
        managed_formulas.update((overdue_formula, attention_formula))

    def is_managed_alert_formula(rule_formula):
        normalized = rule_formula.upper().replace(" ", "")
        row_match = re.match(rf'AND\(\${re.escape(sn_column)}(\d+)<>"",', normalized)
        if rule_formula in managed_formulas:
            return True
        if not row_match:
            return False
        row_number = row_match.group(1)
        return (
            f'${date_column}{row_number}' in normalized
            and f'${status_column}{row_number}' in normalized
            and "TODAY()" in normalized and ">=25" in normalized
            and "COMPLETED" in normalized and '"N/A"' in normalized
        )

    current_matches = []
    for block_match in _element(sheet_xml, "conditionalFormatting").finditer(sheet_xml):
        block = block_match.group()
        for rule_match in _element(block, "cfRule").finditer(block):
            rule = rule_match.group()
            formula_match = _element(rule, "formula").search(rule)
            if formula_match is None:
                continue
            formula_xml = formula_match.group()
            rule_formula = unescape(formula_xml[formula_xml.index(">") + 1:formula_xml.rfind("</")])
            if is_managed_alert_formula(rule_formula):
                current_matches.append((rule_formula, _attributes(rule).get("dxfId"),
                                        _attributes(block).get("sqref")))
    expected_matches = sorted((
        (overdue_formula, str(overdue_dxf_id), overdue_refs),
        (attention_formula, str(attention_dxf_id), attention_refs),
    ))
    if sorted(current_matches) == expected_matches:
        return _renumber_conditional_priorities(sheet_xml)

    # Replace this automation's previous 25-day rule rather than layering a
    # second rule over its ranges. Preserve any unrelated user-authored rules.
    def remove_managed_rules(block_match):
        block = block_match.group()
        rules = list(_element(block, "cfRule").finditer(block))
        managed = []
        for rule_match in rules:
            rule = rule_match.group()
            formula_match = _element(rule, "formula").search(rule)
            if formula_match is None:
                continue
            formula_xml = formula_match.group()
            rule_formula = unescape(formula_xml[formula_xml.index(">") + 1:formula_xml.rfind("</")])
            if is_managed_alert_formula(rule_formula):
                managed.append(rule_match)
        for rule_match in reversed(managed):
            block = block[:rule_match.start()] + block[rule_match.end():]
        if managed and not _element(block, "cfRule").search(block):
            return ""
        return block

    sheet_xml = _element(sheet_xml, "conditionalFormatting").sub(remove_managed_rules, sheet_xml)
    prefix_match = re.search(r"<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b", sheet_xml)
    prefix = prefix_match.group("p") if prefix_match else ""

    def append_rule(xml, refs, formula, dxf_id):
        escaped_formula = escape(formula)
        existing_priorities = [int(value) for value in re.findall(
            r'<(?:[\w.-]+:)?cfRule\b[^>]*\bpriority=["\'](\d+)', xml)]
        priority = max(existing_priorities, default=0) + 1
        block = (f'<{prefix}conditionalFormatting sqref="{refs}">'
                 f'<{prefix}cfRule type="expression" dxfId="{dxf_id}" priority="{priority}">'
                 f'<{prefix}formula>{escaped_formula}</{prefix}formula>'
                 f'</{prefix}cfRule></{prefix}conditionalFormatting>')
        insert_at = len(xml)
        for name in ("dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"):
            found = _element(xml, name).search(xml)
            if found:
                insert_at = min(insert_at, found.start())
        if insert_at == len(xml):
            insert_at = xml.rfind("</")
        return xml[:insert_at] + block + xml[insert_at:]

    sheet_xml = append_rule(sheet_xml, overdue_refs, overdue_formula, overdue_dxf_id)
    sheet_xml = append_rule(sheet_xml, attention_refs, attention_formula, attention_dxf_id)
    return _renumber_conditional_priorities(sheet_xml)


def _date_style(styles_xml, existing_index, format_code=None):
    num_format = '14'
    if format_code:
        formats = _element(styles_xml, 'numFmts').search(styles_xml)
        prefix = re.search(r'<(?P<p>(?:[A-Za-z_][\w.-]*:)?)styleSheet\b', styles_xml).group('p')
        existing = list(_element(formats.group(), 'numFmt').finditer(formats.group())) if formats else []
        match = next((m for m in existing if _attributes(m.group()).get('formatCode') == format_code), None)
        if match:
            num_format = _attributes(match.group())['numFmtId']
        else:
            num_format = str(max([int(_attributes(m.group())['numFmtId']) for m in existing] or [163]) + 1)
            fragment = f'<{prefix}numFmt numFmtId="{num_format}" formatCode={quoteattr(format_code)}/>'
            if formats:
                block = formats.group()
                if block.rstrip().endswith('/>'):
                    block = block.rstrip()[:-2] + f'></{prefix}numFmts>'
                pos = block.rfind('</')
                block = block[:pos] + fragment + block[pos:]
                block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{len(existing)+1}"', block, count=1)
                styles_xml = styles_xml[:formats.start()] + block + styles_xml[formats.end():]
            else:
                pos = _element(styles_xml, 'fonts').search(styles_xml).start()
                styles_xml = styles_xml[:pos] + f'<{prefix}numFmts count="1">{fragment}</{prefix}numFmts>' + styles_xml[pos:]
    section = _element(styles_xml, "cellXfs").search(styles_xml)
    if section is None:
        raise ValueError("Workbook cell styles are missing.")
    fragments = list(_element(section.group(), "xf").finditer(section.group()))
    original = fragments[existing_index].group()
    opening_end = original.index(">")
    opening = original[:opening_end]
    for key, value in (("numFmtId", num_format), ("applyNumberFormat", "1")):
        if re.search(rf'\b{key}=', opening):
            opening = re.sub(rf'\b{key}=["\x27][^"\x27]*["\x27]', f'{key}="{value}"', opening)
        else:
            opening = opening.rstrip("/") + f' {key}="{value}"' + ("/" if opening.endswith("/") else "")
    new_style = opening + original[opening_end:]
    block = section.group()
    block = block[:block.rfind("</")] + new_style + block[block.rfind("</"):]
    block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{len(fragments) + 1}"', block, count=1)
    return styles_xml[:section.start()] + block + styles_xml[section.end():], len(fragments)


def _sheet_member(archive, sheet_name):
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = next(s for s in workbook.findall("s:sheets/s:sheet", ns) if s.attrib["name"] == sheet_name)
    rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(r.attrib["Target"] for r in rels if r.attrib["Id"] == rel_id)
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))


def add_tracking_column(data, *, on_order_warning=None):
    """Keep operational fields before the final COD/date/status columns."""
    from register_column_order import normalize_register_columns, UnsupportedColumnOrderError
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header_row, cols = select_register(book)
        if 'issued' in cols:
            try:
                return normalize_register_columns(data)
            except UnsupportedColumnOrderError as exc:
                # Reading/refreshing a valid register must remain possible
                # when only its automatic column move needs Excel.
                if on_order_warning:
                    on_order_warning(str(exc))
                return data
        if externally_protected(sheet):
            raise ValueError('Unprotect the register before adding Certificate Issued.')
        last_column = max(cell.column for cell in sheet[header_row] if text(cell.value))
        column = checked_register_column(last_column + 1)
        if any(text(sheet.cell(row, column).value) for row in range(header_row + 1, sheet.max_row + 1)):
            raise ValueError('The column after the last heading contains data. Name that column in Excel before connecting the file.')
        source_header = sheet.cell(header_row, cols['supplier'] + 1)
        heading = sheet.cell(header_row, column)
        heading._style = copy(source_header._style)
        alignment = copy(source_header.alignment)
        alignment.wrapText = True
        heading.alignment = alignment
        batches = read_batches(BytesIO(data))
        with ZipFile(BytesIO(data)) as source:
            member = _sheet_member(source, sheet.title)
            xml = source.read(member).decode('utf-8')
            xml = _set_cell(xml, f'{get_column_letter(column)}{header_row}', 'Certificate Issued', 'inlineStr', heading.style_id)
            prefix = re.search(r'<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b', xml).group('p')
            merges = []
            for index, batch in enumerate(batches):
                start = batch['source_row']
                end = (batches[index + 1]['source_row'] - 1 if index + 1 < len(batches)
                       else register_content_end(sheet, header_row))
                for row in range(start, end + 1):
                    style = sheet.cell(row, cols['sn'] + 1).style_id
                    xml = _set_cell(xml, f'{get_column_letter(column)}{row}', '', 'blank', style)
                if end > start:
                    merges.append(f'{get_column_letter(column)}{start}:{get_column_letter(column)}{end}')
            xml = _patch_register_merges(xml, set(), merges)
            column_xml = f'<{prefix}col width="18" customWidth="1" min="{column}" max="{column}"/>'
            col_section = _element(xml, 'cols').search(xml)
            if col_section:
                block = col_section.group(); pos = block.rfind('</')
                block = block[:pos] + column_xml + block[pos:]
                xml = xml[:col_section.start()] + block + xml[col_section.end():]
            else:
                pos = _element(xml, 'sheetData').search(xml).start()
                xml = xml[:pos] + f'<{prefix}cols>{column_xml}</{prefix}cols>' + xml[pos:]
            dimension = _element(xml, 'dimension').search(xml)
            if dimension:
                block = re.sub(r'\bref=["\x27][^"\x27]*["\x27]', f'ref="A1:{get_column_letter(max(sheet.max_column, column))}{sheet.max_row}"', dimension.group())
                xml = xml[:dimension.start()] + block + xml[dimension.end():]
            styles = sync_generated_styles(book, source.read('xl/styles.xml').decode('utf-8'))
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = source.comment
                for info in source.infolist():
                    payload = xml.encode('utf-8') if info.filename == member else styles.encode('utf-8') if info.filename == 'xl/styles.xml' else source.read(info.filename)
                    target.writestr(copy(info), payload)
            return normalize_register_columns(output.getvalue())
    finally:
        book.close()


class WorkbookSnapshot:
    """Read a workbook without excluding Excel or other readers."""

    def __init__(self, path):
        self.path = Path(path)
        self.data = None

    def __enter__(self):
        self.data = self.path.read_bytes()
        return self

    def __exit__(self, *args):
        return False


class WorkbookSession:
    def __init__(self, path, backup_dir, *, lock_timeout=5):
        self.lock_timeout = lock_timeout
        self.path = Path(path)
        self.backup_dir = Path(backup_dir)
        self.backup_path = None
        self.context = None
        self.before_persist = None
        self.order_warning = None

    def __enter__(self):
        self.context = exclusive_file(self.path, wait_timeout=self.lock_timeout)
        self.stream = self.context.__enter__()
        try:
            self.data = self.stream.read()
            from inward_serials import bind_workbook
            bind_workbook(self.path, self.data)
            return self
        except Exception:
            self.context.__exit__(None, None, None)
            raise

    def __exit__(self, *args):
        if self.context:
            try:
                self.context.__exit__(*args)
            finally:
                self.context = None
                from inward_serials import unbind_workbook
                unbind_workbook(self.path)

    def ensure_alert_formatting(self, sheet_name, batches=None, as_of=None, *, root=None, plant=None):
        """Refresh colors and promote complete, unselected inwards to Ready."""
        check = load_workbook(BytesIO(self.data), data_only=False)
        try:
            if sheet_name in check and externally_protected(check[sheet_name]):
                self.order_warning = 'This Excel sheet is protected. Unprotect it to update its status or layout.'
                return None
        finally:
            check.close()
        self.order_warning = None
        from register_presentation import normalize_register_presentation
        prepared = normalize_register_presentation(add_tracking_column(self.data,
            on_order_warning=lambda warning: setattr(self, 'order_warning', warning)))
        source_rows = {}
        from register_rows import compact_register_gaps
        try:
            prepared = compact_register_gaps(prepared, source_rows=source_rows)
        except UnsupportedReorderError as exc:
            self.order_warning = str(exc)
        sorted_rows = {}
        serial_changes = {}
        before_sort, rows_before_sort = prepared, source_rows
        try:
            prepared = sort_register_by_inward_date(prepared, source_rows=sorted_rows)
            source_rows = ({identity: sorted_rows.get((identity[0], row), row) for identity, row in source_rows.items()}
                           if source_rows else sorted_rows)
            if root is not None and plant is not None:
                from inward_serials import resequence_date_ordered_inwards
                try:
                    prepared, serial_changes = resequence_date_ordered_inwards(prepared, self.data, root=root, plant=plant)
                except ValueError as exc:
                    prepared, source_rows = before_sort, rows_before_sort
                    self.order_warning = str(exc)
        except UnsupportedReorderError as exc:
            # Manual edits can introduce cross-record formulas or row-anchored
            # objects. Keep their order intact and let the app explain why.
            self.order_warning = str(exc)
            prepared = repair_register_borders(prepared)
        if prepared != self.data:
            self._persist(prepared)
            identity_field = 'case_id' if serial_changes else 'source_row'
            fresh = {(b['source_sheet'], b[identity_field]): b for b in read_batches(BytesIO(self.data))}
            for batch in batches or []:
                if batch['source_sheet'] == sheet_name:
                    previous = (batch['source_sheet'], batch['source_row'])
                    current = ((batch['source_sheet'], serial_changes.get(batch['case_id'], batch['case_id']))
                               if serial_changes else (batch['source_sheet'], source_rows.get(previous, batch['source_row'])))
                    if current in fresh:
                        history_conflict = batch.get('record_identity_conflict')
                        batch.update(fresh[current])
                        if history_conflict and history_conflict not in batch['blocking_issues']:
                            batch['blocking_issues'].append(history_conflict)
        batches = batches if batches is not None else read_batches(BytesIO(self.data))
        sheet_batches = {batch["source_row"]: batch for batch in batches
                         if batch["source_sheet"] == sheet_name}
        book = load_workbook(BytesIO(self.data), data_only=False, read_only=False)
        try:
            if sheet_name not in book.sheetnames:
                return None
            sheet = book[sheet_name]
            header = register_header(sheet)
            if header is None:
                return None
            header_row, cols = header
            first_row = header_row + 1
            if not {"sn", "inward_date", "issued", "supplier"}.issubset(cols):
                return None
            # Red identifies a 25-day overdue supplier. Amber marks a date
            # that needs checking, directly on the date cell.
            overdue_columns = [get_column_letter(cols["supplier"] + 1)]
            date_attention_columns = [get_column_letter(cols["inward_date"] + 1)]
            sn_column = get_column_letter(cols["sn"] + 1)
            date_column = get_column_letter(cols["inward_date"] + 1)
            status_column = get_column_letter(cols["issued"] + 1)
            row_limit = max(10000, sheet.max_row + 2000)
            status_labels = {
                "ready": "Ready", "alert": "Alert", "na": "N/A", "blank": "",
                "generated": "Generated", "generated_manual": "Generated Externally",
                "completed": "Completed",
            }
            status_updates = []
            for row_number in range(first_row, sheet.max_row + 1):
                cell = sheet.cell(row_number, cols["issued"] + 1)
                if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                    continue
                batch = sheet_batches.get(row_number)
                # A history conflict belongs to this inward. Preserve its
                # status while allowing other complete records to become Ready.
                if batch and batch.get('record_identity_conflict'):
                    continue
                if any(cell.coordinate in area for area in sheet.merged_cells.ranges):
                    if not batch:
                        continue
                    try:
                        _check_batch_merge(sheet, cell, row_number, cols)
                    except ValueError:
                        continue
                current_status = status_value(cell.value)
                canonical = status_labels.get(current_status)
                if batch and supplier_is_always_not_applicable(batch.get("supplier")):
                    canonical = "N/A"
                elif batch and current_status in {"blank", "ready", "alert"}:
                    if source_value_issues(batch, as_of):
                        canonical = "Alert"
                    elif inward_details_complete(batch, as_of):
                        canonical = "Ready"
                    else:
                        canonical = "Alert"
                if canonical is not None and cell.value != (canonical or None):
                    status_updates.append((cell.coordinate, canonical))
        finally:
            book.close()

        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, sheet_name)
            styles_xml = source.read("xl/styles.xml").decode("utf-8")
            updated_styles, red_dxf = _add_alert_dxf(styles_xml)
            updated_styles, date_attention_dxf = _add_color_dxf(
                updated_styles, "FFC00000", "FFFFFFFF")
            status_dxf_ids = {"READY": None, "ALERT": None, "NA": None,
                              "BLANK": None, "GENERATED": None,
                              "GENERATED EXTERNALLY": None, "COMPLETED": None}
            for label, fill, font in (
                ("READY", "FFB6D7A8", "FF006100"),
                ("ALERT", "FFC00000", "FFFFFFFF"),
                ("NA", "FFFFFFFF", "FF000000"),
                ("BLANK", "FFFFFFFF", "FF000000"),
                ("GENERATED", "FFFFE699", "FF7F6000"),
                ("GENERATED EXTERNALLY", "FFDDEBF7", "FF1F4E78"),
                ("COMPLETED", "FFFFE699", "FF7F6000"),
            ):
                updated_styles, status_dxf_ids[label] = _add_color_dxf(updated_styles, fill, font)
            original_sheet_xml = source.read(member).decode("utf-8")
            sheet_xml = original_sheet_xml
            for coordinate, canonical in status_updates:
                sheet_xml = _set_cell(sheet_xml, coordinate, canonical, "inlineStr" if canonical else "blank")
            sheet_xml = _update_issued_status_formatting(sheet_xml, f"{status_column}{first_row}", first_row)
            sheet_xml = _add_status_conditional_formatting(
                sheet_xml, status_column, sn_column, status_dxf_ids, row_limit, first_row,
            )
            updated_sheet = _add_alert_conditional_formatting(
                sheet_xml, sn_column=sn_column, date_column=date_column,
                status_column=status_column, overdue_columns=overdue_columns,
                date_attention_columns=date_attention_columns,
                overdue_dxf_id=red_dxf, attention_dxf_id=date_attention_dxf,
                row_limit=row_limit, first_row=first_row,
            )
            updated = self.data
            if updated_styles != styles_xml or updated_sheet != original_sheet_xml:
                ET.fromstring(updated_styles)
                ET.fromstring(updated_sheet)
                changed = {"xl/styles.xml": updated_styles.encode("utf-8"),
                           member: updated_sheet.encode("utf-8")}
                buffer = BytesIO()
                with ZipFile(buffer, "w") as target:
                    target.comment = source.comment
                    for info in source.infolist():
                        target.writestr(info, changed.get(info.filename, source.read(info.filename)))
                updated = buffer.getvalue()
        fitted = normalize_register_presentation(updated, wrap_only=True, fit_register=True)
        fitted = remove_issued_inward_protection(fitted)
        backup = self._persist(fitted, exact=True) if fitted != self.data else self.backup_path
        for coordinate, canonical in status_updates:
            batch = sheet_batches.get(int(re.search(r"\d+$", coordinate).group()))
            if batch is not None:
                batch["issued"] = canonical
        return backup

    def prepare(self, batch, certificate_no=None, **_legacy_options):
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and run again.")
        matches = [b for b in read_batches(BytesIO(self.data))
                   if text(b['case_id']).casefold() == text(batch['case_id']).casefold()
                   and b['source_sheet'] == batch['source_sheet']]
        if len(matches) > 1:
            raise ValueError(f"Serial Number {batch['case_id']} appears more than once. Give each inward a unique Serial Number before generating.")
        if len(matches) != 1 or batch_content(matches[0]) != batch_content(batch):
            raise ValueError("The source batch changed; check the worksheet before continuing.")
        current = matches[0]
        if status_value(current["issued"]) != "ready":
            raise ValueError("The row is no longer marked Ready. Choose Ready in Excel before generating its PDF.")
        book = load_workbook(BytesIO(self.data), data_only=False)
        try:
            sheet = book[current["source_sheet"]]
            cols = columns([c.value for c in sheet[current.get('source_header_row', 1)]])
            # Missing output columns are added by commit(). This lets the
            # operator recover old registers without manually editing their
            # headers first.
            if "cod_no" not in cols or "cod_date" not in cols or "issued" not in cols:
                return (current, None, None, None, 0, False, book.epoch)
            cells = [sheet.cell(current["source_row"], cols[key] + 1) for key in ("cod_no", "cod_date", "issued")]
            for cell in cells:
                _check_batch_merge(sheet, cell, current['source_row'], cols)
                if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            return (current, cells[0].coordinate, cells[1].coordinate, cells[2].coordinate,
                    cells[1].style_id, is_date_format(cells[1].number_format), book.epoch)
        finally:
            book.close()

    def commit(self, batch, certificate_no, certificate_date, target_status="Generated", overrides=None,
               staged_data=None, allow_reissue=False, new_batch=False, deferred_form6=None, deferred_reference=None):
        if status_value(target_status) != "generated":
            raise ValueError("The app writes Generated after it creates the PDF. Set Completed only after the certificate has been sent.")
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and run again.")
        original_batches = read_batches(BytesIO(self.data))
        matches = [b for b in original_batches
                   if text(b['case_id']).casefold() == text(batch['case_id']).casefold()
                   and b['source_sheet'] == batch['source_sheet']]
        if new_batch:
            if any(text(b['case_id']).casefold() == text(batch['case_id']).casefold() for b in original_batches):
                raise ValueError(f"Serial Number {batch['case_id']} already exists. Choose a different Serial Number for this new inward.")
            if staged_data is None:
                raise ValueError("A new inward row requires staged certificate details.")
            current = batch
        else:
            if len(matches) > 1:
                raise ValueError(f"Serial Number {batch['case_id']} appears more than once. Give each inward a unique Serial Number before saving.")
            if len(matches) != 1 or batch_content(matches[0]) != batch_content(batch):
                raise ValueError("The source batch changed; check the worksheet before continuing.")
            current = matches[0]
        if not allow_reissue and status_value(current["issued"]) != "ready":
            raise ValueError("The row is no longer marked Ready. Choose Ready in Excel before generating its PDF.")

        # Staging may replace formula/protected values for a render-only view.
        # Existing records still require their original output-cell checks.
        # A verified new Serial Number has no original record: its staged row
        # can now occupy an unrelated original record's position after sorting.
        if not new_batch:
            original = load_workbook(BytesIO(self.data), data_only=False,
                                     keep_vba=self.path.suffix.casefold() == ".xlsm")
            try:
                sheet = original[current["source_sheet"]]
                cols = columns([cell.value for cell in sheet[current.get('source_header_row', 1)]])
                for key in ("cod_no", "cod_date", "issued"):
                    if key not in cols:
                        continue
                    cell = sheet.cell(current["source_row"], cols[key] + 1)
                    _check_batch_merge(sheet, cell, current['source_row'], cols)
                    if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                        raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            finally:
                original.close()

        working_data = self.data if staged_data is None else staged_data
        if staged_data is not None:
            staged_matches = [b for b in read_batches(BytesIO(staged_data))
                              if b["case_id"] == current["case_id"]
                              and b["source_sheet"] == current["source_sheet"]]
            if (len(staged_matches) != 1
                    or (new_batch and batch_content(staged_matches[0]) != batch_content(current))):
                raise ValueError("The staged inward row changed; the original workbook was kept.")
            # Backdating moves the complete inward block. Original output cells
            # were checked above; tracking writes now use its staged position.
            current = dict(current, source_row=staged_matches[0]['source_row'])
        # Explicit form issuance without staged data is certificate-only.
        # Preserve the existing override/header contract for normal Ready runs.
        certificate_only = allow_reissue and staged_data is None
        persisted_overrides = None if certificate_only else overrides
        if deferred_form6 is not None:
            if (text(current.get('form6')) or not text(deferred_form6)
                    or status_value(current.get('issued')) in {'generated', 'generated_manual', 'completed'}):
                raise ValueError('The Form 6 number changed. Refresh and review this certificate again.')
            persisted_overrides = dict(persisted_overrides or {}, form6=text(deferred_form6))
        if deferred_reference is not None:
            if (text(current.get('reference')) or not text(deferred_reference)
                    or status_value(current.get('issued')) in {'generated', 'generated_manual', 'completed'}):
                raise ValueError('The reference number changed. Refresh and review this certificate again.')
            persisted_overrides = dict(persisted_overrides or {}, reference=text(deferred_reference))

        # Use openpyxl for this one transactional update so missing headers
        # and operator overrides can be added together, then persist the
        # resulting package atomically through the existing session.
        book = load_workbook(BytesIO(working_data), data_only=False,
                             keep_vba=self.path.suffix.casefold() == ".xlsm")
        try:
            sheet = book[current["source_sheet"]]
            header_row = current.get('source_header_row', 1)
            cols = columns([c.value for c in sheet[header_row]])
            labels = {
                "cod_no": "COD No", "cod_date": "COD Date", "issued": "Certificate Issued",
                "supplier": "Supplier Name", "location": "Location", "inward_date": "Inward Date",
                "form6": "Form 6 No", "reference": "Reference No.",
                "certificate_weight": "Certificate Weight (KG)",
                "certificate_type": "Certificate Type", "certificate_fields": "Certificate Fields",
            }
            required = {"cod_no", "cod_date", "issued"}
            for key in required:
                if key not in cols:
                    if key != 'issued':
                        raise ValueError(f'Excel is missing {labels[key]}. Add that existing register heading before generating a certificate.')
                    column = next_register_column(sheet, header_row)
                    sheet.cell(header_row, column, labels[key])
                    cols[key] = column - 1

            def check_cell(cell):
                _check_batch_merge(sheet, cell, current['source_row'], cols)
                if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")

            row = current["source_row"]
            output_values = {
                "cod_no": certificate_no,
                "cod_date": certificate_date,
                "issued": target_status,
            }
            for key, value in (persisted_overrides or {}).items():
                if key in cols and key in labels and key not in {"cod_no", "cod_date", "issued"}:
                    if key == "inward_date":
                        value = parse_date(value)
                    elif key == "certificate_weight":
                        value = number(value)
                    elif key == "certificate_type":
                        value = text(value) or "standard"
                    elif key == "certificate_fields":
                        value = _certificate_fields_json(value)
                    elif value is not None:
                        value = str(value).strip()
                    output_values[key] = value
            for key, value in output_values.items():
                cell = sheet.cell(row, cols[key] + 1)
                check_cell(cell)
                if key == "certificate_weight" and value is None:
                    cell.value = ""
                else:
                    cell.value = value
                if key == "cod_date" and not is_date_format(cell.number_format):
                    cell.number_format = "dd/mm/yy"
            buffer = BytesIO()
            book.save(buffer)
            updated = buffer.getvalue()
        finally:
            book.close()

        metadata = {key:value for key,value in (persisted_overrides or {}).items() if key not in cols}
        metadata['_pdf_recovery_pending'] = False
        metadata['_recycling_completed'] = False
        if current.get('external_pdf'):
            metadata['_external_pdf'] = None
        updated = write_details(updated, current['source_sheet'], current['case_id'], metadata)
        updated = self._move_primary_hdd_after_recycling_write(updated, current)
        updated = order_register_for_save(updated, self, fit_inward=(current['source_sheet'], current['case_id']))

        verified_matches = [b for b in read_batches(BytesIO(updated))
                            if b["case_id"] == current["case_id"]
                            and b["source_sheet"] == current["source_sheet"]]
        if len(verified_matches) != 1:
            raise ValueError("The updated workbook could not be read back; the original workbook was kept.")
        verified = verified_matches[0]
        if (verified["cod_no"] != certificate_no
                or parse_date(verified["cod_date"]) != certificate_date
                or status_value(verified["issued"]) != status_value(target_status)):
            raise ValueError(f"Excel COD, date, or {target_status} status could not be verified; the original workbook was kept.")
        for key, value in (persisted_overrides or {}).items():
            if key in {"supplier", "location", "form6", "reference"} and text(value) != verified.get(key, ""):
                raise ValueError(f"Excel {key} could not be verified; the original workbook was kept.")
            if key == "certificate_type" and (text(value) or "standard") != verified.get(key, "standard"):
                raise ValueError("Excel certificate_type could not be verified; the original workbook was kept.")
            if key == "certificate_fields" and json.loads(_certificate_fields_json(value)) != verified.get(key, {}):
                raise ValueError("Excel certificate_fields could not be verified; the original workbook was kept.")
        return self._persist(updated)

    def set_issued_status(self, case_id, source_sheet, source_row, expected_certificate_no,
                          expected_status, new_status, *, external=None, metadata=None):
        """Update status, with optional tracking details for a linked external PDF."""
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and try again.")
        if isinstance(expected_status, str):
            expected_status = (expected_status,)
        expected = {status_value(value) for value in expected_status}
        target = status_value(new_status)
        target_labels = {"ready": "Ready", "alert": "Alert", "na": "N/A", "generated": "Generated",
                         "generated_manual": "Generated Externally", "completed": "Completed"}
        if target not in target_labels:
            raise ValueError("Choose a supported certificate status.")
        new_status = target_labels[target]
        matches = [batch for batch in read_batches(BytesIO(self.data))
                   if text(batch['case_id']).casefold() == text(case_id).casefold()
                   and batch["source_sheet"] == source_sheet]
        if len(matches) > 1:
            raise ValueError(f"Serial Number {case_id} appears more than once. Give each inward a unique Serial Number before changing its status.")
        if len(matches) != 1:
            raise ValueError("The selected inward row changed or could not be found. Check both files again before updating its status.")
        current = matches[0]
        if current["cod_no"] != str(expected_certificate_no):
            raise ValueError("The selected row's COD No changed. Check both files again before updating its status.")
        if status_value(current["issued"]) not in expected:
            raise ValueError("The selected row's Certificate Issued status changed. Check both files again before updating it.")
        book = load_workbook(BytesIO(self.data), data_only=False)
        try:
            sheet = book[current["source_sheet"]]
            cols = columns([cell.value for cell in sheet[current.get('source_header_row', 1)]])
            if "issued" not in cols:
                raise ValueError("Add a Certificate Issued column before updating the completion status.")
            cell = sheet.cell(current["source_row"], cols["issued"] + 1)
            _check_batch_merge(sheet, cell, current['source_row'], cols)
            if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                raise ValueError(f"{cell.coordinate} is a formula or protected cell; its status was not changed.")
            coordinate = cell.coordinate
            tracking = {}
            external_type = external.get('certificate_type') if external else None
            if external:
                for key, value, kind in (
                    ('cod_no', external['cod'], 'inlineStr'),
                    ('cod_date', to_excel(external['date'], book.epoch), 'n'),
                ):
                    if key not in cols:
                        raise ValueError('The workbook needs its COD No and COD Date columns. Refresh it and try again.')
                    tracking_cell = sheet.cell(current['source_row'], cols[key] + 1)
                    _check_batch_merge(sheet, tracking_cell, current['source_row'], cols)
                    if tracking_cell.data_type == 'f' or (externally_protected(sheet) and tracking_cell.protection.locked):
                        raise ValueError(f'{tracking_cell.coordinate} uses a formula or is protected in Excel.')
                    tracking[tracking_cell.coordinate] = (value, kind, tracking_cell.style_id,
                                                          key == 'cod_date' and not is_date_format(tracking_cell.number_format))
                if external_type and external_type != current['certificate_type'] and 'certificate_type' in cols:
                    cell = sheet.cell(current['source_row'], cols['certificate_type'] + 1)
                    _check_batch_merge(sheet, cell, current['source_row'], cols)
                    if cell.data_type == 'f' or (externally_protected(sheet) and cell.protection.locked):
                        raise ValueError(f'{cell.coordinate} uses a formula or is protected in Excel.')
                    tracking[cell.coordinate] = (external_type, 'inlineStr', cell.style_id, False)
            generated_styles = None
            if tracking:
                with ZipFile(BytesIO(self.data)) as archive:
                    generated_styles = sync_generated_styles(book, archive.read('xl/styles.xml').decode('utf-8'))
        finally:
            book.close()
        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, current["source_sheet"])
            xml = source.read(member).decode("utf-8")
            xml = _set_cell(xml, coordinate, new_status, "inlineStr")
            styles = generated_styles
            for address, (value, kind, style, need_date_style) in tracking.items():
                if need_date_style:
                    styles, style = _date_style(styles, style, 'dd/mm/yy')
                xml = _set_cell(xml, address, value, kind, style=style)
            xml = _update_issued_status_formatting(xml, coordinate, current.get('source_header_row', 1) + 1)
            ET.fromstring(xml)
            buffer = BytesIO()
            with ZipFile(buffer, "w") as target_archive:
                target_archive.comment = source.comment
                for info in source.infolist():
                    content = (xml.encode('utf-8') if info.filename == member else styles.encode('utf-8')
                               if info.filename == 'xl/styles.xml' and tracking else source.read(info.filename))
                    target_archive.writestr(info, content)
        updated = buffer.getvalue()
        expected_cod = current['cod_no']
        expected_date = parse_date(current['cod_date'])
        metadata_updates = dict(metadata or {})
        metadata_updates.setdefault('_recycling_completed', False)
        if external:
            expected_cod, expected_date = external['cod'], external['date']
            metadata_updates.update(_external_pdf=external['link'], _pdf_recovery_pending=False)
            if external_type and 'certificate_type' not in cols:
                metadata_updates['certificate_type'] = external_type
        if target == 'ready' and status_value(current['issued']) in {'generated', 'generated_manual', 'completed'}:
            metadata_updates['_pdf_recovery_pending'] = True
        if metadata_updates:
            updated = write_details(updated, current['source_sheet'], current['case_id'], metadata_updates)
        if external:
            updated = self._move_primary_hdd_after_recycling_write(updated, current)
        verified = next(batch for batch in read_batches(BytesIO(updated))
                        if batch["case_id"] == current['case_id']
                        and batch["source_sheet"] == source_sheet
                        and batch["source_row"] == current['source_row'])
        if (status_value(verified["issued"]) != target
                or verified["cod_no"] != expected_cod
                or parse_date(verified["cod_date"]) != expected_date
                or (external and verified['external_pdf'] != external['link'])
                or (external_type and verified['certificate_type'] != external_type)):
            raise ValueError("The workbook status update could not be verified; the original file was kept.")
        from register_presentation import normalize_register_presentation
        updated = normalize_register_presentation(updated, wrap_only=True, fit_inward=(source_sheet, current['case_id']))
        return self._persist(updated)

    def _move_primary_hdd_after_recycling_write(self, data, previous):
        if not previous.get('hdd_on_main_row'):
            return data
        from hdd_workbook import stage_hdd_tracking
        staged = WorkbookSnapshot(self.path)
        staged.data = data
        return stage_hdd_tracking(staged, previous['case_id'], previous['hdd_cod_no'],
                                  parse_date(previous['hdd_cod_date']), format_register=False)

    def _persist(self, updated, *, exact=False):
        if not exact:
            from register_presentation import normalize_register_presentation
            updated = normalize_register_presentation(updated, wrap_only=True)
            updated = remove_issued_inward_protection(updated)
        if self.before_persist:
            self.before_persist(updated)
        if not self.backup_path:
            self.backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = self.backup_dir / f"{self.path.stem}_{stamp}_{uuid.uuid4().hex[:8]}{self.path.suffix}"
            created = False
            try:
                with backup_path.open("xb") as backup:
                    created = True
                    backup.write(self.data)
                    backup.flush()
                    os.fsync(backup.fileno())
            except Exception:
                if created:
                    backup_path.unlink(missing_ok=True)
                raise
            self.backup_path = backup_path
        fd, name = tempfile.mkstemp(prefix=".certificate_", suffix=self.path.suffix, dir=self.path.parent)
        temporary = Path(name)
        new_context = None
        try:
            with os.fdopen(fd, "wb") as output:
                output.write(updated)
                output.flush()
                os.fsync(output.fileno())
            new_context = exclusive_file(temporary)
            new_stream = new_context.__enter__()
            # Windows cannot replace the target while our target handle is
            # open. The replacement is already locked, so readers cannot
            # observe a partially written workbook after the atomic swap.
            self.context.__exit__(None, None, None)
            self.context = None
            try:
                os.replace(temporary, self.path)
            except Exception:
                restore = exclusive_file(self.path)
                self.stream = restore.__enter__()
                self.context = restore
                self.data = self.stream.read()
                raise
            self.context, self.stream, self.data = new_context, new_stream, updated
            from inward_serials import bind_workbook
            bind_workbook(self.path, self.data)
            new_context = None
        finally:
            if new_context:
                new_context.__exit__(None, None, None)
            if temporary.exists():
                temporary.unlink()
        return self.backup_path


def update_issued_status(path, backup_dir, case_id, source_sheet, source_row,
                         expected_certificate_no, expected_status, new_status):
    """Safely update only the issued-status cell after checking its current values."""
    with WorkbookSession(path, backup_dir) as session:
        return session.set_issued_status(
            case_id, source_sheet, source_row, expected_certificate_no,
            expected_status, new_status,
        )


def stage_certificate_details(session, details, validate_cells=True, *, inward_status="Ready"):
    """Build a Ready form view without writing or changing session.data.

    Skip source-cell checks only for a certificate-only render. Commit still
    checks the original tracking cells before updating COD, date and status.
    """
    import math
    from excel_values import formula_number

    details = dict(details or {})
    source_case_id = text(details.get('case_id'))
    text_labels = {
        'case_id': 'Serial Number', 'supplier': 'Company name', 'location': 'Location',
        'inward_date': 'Inward Date', 'form6': 'Form 6 number', 'reference': 'Reference number',
        'item': 'Material description', 'quantity': 'Quantity', 'per_piece_weight': 'Weight per piece',
        'weight': 'Weight', 'certificate_weight': 'Certificate weight',
        'cod_no': 'COD number', 'cod_date': 'COD date', 'certificate_type': 'Certificate type',
    }
    for key, label in text_labels.items():
        if isinstance(details.get(key), str):
            details[key] = checked_excel_text(details[key], label, normalize_paste=True)
    if 'certificate_fields' in details:
        def normalize_field_values(value):
            if isinstance(value, str):
                return checked_excel_text(value, 'Certificate fields', normalize_paste=True)
            if isinstance(value, dict):
                return {checked_excel_text(key, 'Certificate field name'): normalize_field_values(item)
                        for key, item in value.items()}
            if isinstance(value, list):
                return [normalize_field_values(item) for item in value]
            return value
        raw_fields = details['certificate_fields']
        if isinstance(raw_fields, str) and raw_fields:
            raw_fields = json.loads(raw_fields)
        details['certificate_fields'] = normalize_field_values(raw_fields)
    if isinstance(details.get('items'), list):
        normalized_items = []
        for index, item in enumerate(details['items'], 1):
            if isinstance(item, dict):
                item = dict(item)
                for key in ('item', 'quantity', 'per_piece_weight', 'weight'):
                    if isinstance(item.get(key), str):
                        item[key] = checked_excel_text(item[key], f'Material {index}: {text_labels[key]}', normalize_paste=True)
            normalized_items.append(item)
        details['items'] = normalized_items
    case_id = text(details.get("case_id"))
    if not case_id:
        raise ValueError("Serial Number is required.")
    source_batches = read_batches(BytesIO(session.data))
    if validate_cells and 'items' in details:
        from register_rows import ensure_material_capacity
        from types import SimpleNamespace
        required = len(material_items(details)[0])
        current = next((b for b in source_batches if text(b['case_id']).casefold() == source_case_id.casefold()), None)
        expanded = (ensure_material_capacity(session.data, source_case_id, required)
                    if current and required > len(current['items']) else session.data)
        if expanded != session.data:
            shadow = SimpleNamespace(data=expanded, path=session.path, order_warning=None)
            result = stage_certificate_details(shadow, details, validate_cells=True, inward_status=inward_status)
            if shadow.order_warning:
                session.order_warning = shadow.order_warning
            return result
    if source_case_id != case_id and any(
            text(batch['case_id']).casefold() == source_case_id.casefold() for batch in source_batches):
        raise ValueError('The existing Serial Number contains unsupported control characters. '
                         'Correct that Serial Number in Excel, save, then refresh before editing this inward.')
    book = load_workbook(BytesIO(session.data), data_only=False,
                         keep_vba=session.path.suffix.casefold() == ".xlsm")
    try:
        sheet, header_row, cols = select_register(book)
        original_width = register_content_width(sheet, header_row)
        register_tables = [(table, range_boundaries(table.ref)) for table in sheet.tables.values()
                           if range_boundaries(table.ref)[1] == header_row
                           and range_boundaries(table.ref)[2] == original_width
                           and all(range_boundaries(table.ref)[0] <= column + 1
                                   <= range_boundaries(table.ref)[2] for column in cols.values())]
        if not validate_cells:
            # A preview/PDF-only correction is an in-memory view. Removing body
            # merges here makes placeholder cells writable without touching Excel.
            for area in list(sheet.merged_cells.ranges):
                if area.min_row > header_row:
                    sheet.unmerge_cells(str(area))
        labels = {
            "sn": "SN", "inward_date": "Inward Date", "supplier": "Supplier Name",
            "location": "Location", "item": "Item", "quantity": "Quantity in Pcs",
            "per_piece_weight": "Per Pcs Weight in (KG)", "weight": "Weight (KG)",
            "certificate_weight": "Certificate Weight (KG)", "form6": "Form 6 No",
            "reference": "Reference No.", "cod_no": "COD No", "cod_date": "COD Date",
            "issued": "Certificate Issued",
            "certificate_type": "Certificate Type", "certificate_fields": "Certificate Fields",
        }
        material_override = material_items(details)[0] if "items" in details else None
        entered_columns = {'location': bool(text(details.get('location')))}
        for key in ('quantity', 'per_piece_weight', 'weight'):
            entered_columns[key] = (
                any(item.get(key) is not None for item in material_override)
                if material_override is not None else details.get(key) not in (None, ''))
        if validate_cells:
            for key, entered in entered_columns.items():
                if key in cols or not entered:
                    continue
                if externally_protected(sheet):
                    raise ValueError(f'Unprotect the register before adding the missing {labels[key]} column.')
                column = next_register_column(sheet, header_row)
                heading = sheet.cell(header_row, column, labels[key])
                heading._style = copy(sheet.cell(header_row, cols['supplier'] + 1)._style)
                cols[key] = column - 1
        for key, value in details.items():
            if key in labels and key not in cols and (
                    not validate_cells or key == 'issued' or (key == 'form6' and text(value))
            ) and (value not in (None, "") or key in {"form6", "issued", "certificate_type", "certificate_fields"}):
                column = next_register_column(sheet, header_row)
                sheet.cell(header_row, column, labels[key])
                cols[key] = column - 1
        for key in (("item", "form6", "issued") if not validate_cells else ('issued',)):
            if key not in cols:
                column = next_register_column(sheet, header_row)
                sheet.cell(header_row, column, labels[key])
                cols[key] = column - 1

        if "items" in details and not validate_cells:
            for key in ("quantity", "per_piece_weight", "weight"):
                if key not in cols:
                    column = next_register_column(sheet, header_row)
                    sheet.cell(header_row, column, labels[key])
                    cols[key] = column - 1

        serial_rows = {}
        sheet_batches = [batch for batch in source_batches if batch['source_sheet'] == sheet.title]
        for batch in sheet_batches:
            serial_rows.setdefault(text(batch['case_id']).casefold(), []).append(batch['source_row'])
        rows = serial_rows.get(case_id.casefold(), [])
        if len(rows) > 1:
            serial = text(sheet.cell(rows[0], cols['sn'] + 1).value)
            raise ValueError(f'Serial Number {serial} appears more than once in Excel rows '
                             + ', '.join(map(str, rows)) + '. Give each inward a unique Serial Number before saving.')
        original_batch = next((batch for batch in sheet_batches
                               if text(batch['case_id']).casefold() == case_id.casefold()), None)
        is_new_record = original_batch is None
        target_row = (original_batch['source_row'] if original_batch else next_register_row(
            sheet, header_row, max(1, len(material_override or []))))
        if original_batch:
            case_id = original_batch['case_id']
        if validate_cells and is_new_record and any(table.totalsRowCount for table, _ in register_tables):
            raise ValueError('This Excel table has a separate total row. Add the new inward before that total in Excel, then refresh.')

        # A PDF-only edit is a batch projection. Its material count must not
        # depend on the space available before the next Excel inward.
        render_material_only = material_override is not None and not validate_cells
        summary_rows = set((original_batch or {}).get('summary_rows', []))
        material_rows = []
        if material_override is not None and not render_material_only:
            material_rows = [target_row]
            following_primary = False
            for row_number in range(target_row + 1, register_content_end(sheet, header_row) + 1):
                if row_number in summary_rows:
                    if not validate_cells:
                        material_rows.append(row_number)
                    continue
                if text(sheet.cell(row_number, cols["sn"] + 1).value):
                    following_primary = True
                    break
                material_rows.append(row_number)
            if len(material_override) > len(material_rows):
                if following_primary or (validate_cells and summary_rows):
                    raise ValueError(
                        f"This Excel record has room for {len(material_rows)} material row(s) before the next client. "
                        "Choose PDF only, or add the extra material row in Excel before the next client or total row."
                    )
                next_row = max(register_content_end(sheet, header_row) + 1, target_row + 1)
                material_rows.extend(range(next_row, next_row + len(material_override) - len(material_rows)))

        shrink_material_rows = []
        shrink_retained_end = None
        shrink_total_cells = ()
        if validate_cells and material_override is not None and original_batch:
            material_columns = {cols[key] + 1 for key in ('item', 'quantity', 'per_piece_weight', 'weight') if key in cols}
            following_rows = [batch['source_row'] for batch in sheet_batches if batch['source_row'] > target_row]
            block_end = min(following_rows) - 1 if following_rows else register_content_end(sheet, header_row)
            original_material_rows = [row_number for row_number in range(target_row, block_end + 1)
                if row_number not in summary_rows and any(
                    sheet.cell(row_number, column).value is not None for column in material_columns)]
            if (len(original_material_rows) == len(original_batch.get('items', []))
                    and len(material_override) < len(original_material_rows)
                    and material_rows[:len(original_material_rows)] == original_material_rows):
                surplus_rows = original_material_rows[len(material_override):]
                retained_end = original_material_rows[len(material_override) - 1]
                occupied_other_cells = any(
                    cell.value is not None or cell.comment or cell.hyperlink
                    for row_number in surplus_rows for cell in sheet[row_number]
                    if cell.column not in material_columns)
                affected_merges = [area for area in sheet.merged_cells.ranges
                    if any(area.min_row <= row_number <= area.max_row for row_number in surplus_rows)]
                safe_merges = all(area.min_col == area.max_col and area.min_col not in material_columns
                    and area.min_row == target_row and area.max_row == original_material_rows[-1]
                    for area in affected_merges)
                if not occupied_other_cells and safe_merges:
                    from register_rows import _material_totals
                    shrink_total_cells, _ = _material_totals(sheet, header_row, cols, sheet_batches)
                    for area in affected_merges:
                        sheet.unmerge_cells(str(area))
                        if retained_end > target_row:
                            sheet.merge_cells(start_row=area.min_row, start_column=area.min_col,
                                              end_row=retained_end, end_column=area.max_col)
                    shrink_material_rows = surplus_rows
                    shrink_retained_end = retained_end

        if is_new_record:
            template_row = formatted_client_template(sheet, header_row, cols, len(material_rows) > 1)
            continuation_row = (template_row + 1 if any(
                area.min_row == template_row and area.max_row > template_row
                for area in sheet.merged_cells.ranges) else template_row)
            style_rows = material_rows or [target_row]
            wrap_columns = {cols[key] + 1 for key in CLIENT_TEXT_FIELDS if key in cols}
            for index, row_number in enumerate(style_rows):
                source_row = template_row if index == 0 else continuation_row
                copy_client_row_styles(sheet, source_row, [row_number], sheet.max_column,
                                       wrap_columns=wrap_columns)
        values = {
            "sn": case_id,
            "inward_date": parse_date(details.get("inward_date")),
            "supplier": text(details.get("supplier")),
            "location": text(details.get("location")),
            "item": material_override[0]["item"] if material_override else text(details.get("item")),
            "quantity": material_override[0]["quantity"] if material_override else (
                number(details.get("quantity")) if number(details.get("quantity")) is not None
                else text(details.get("quantity")) or None),
            "per_piece_weight": material_override[0]["per_piece_weight"] if material_override else number(details.get("per_piece_weight")),
            "weight": material_override[0]["weight"] if material_override else number(details.get("weight")),
            "certificate_weight": number(details.get("certificate_weight")),
            "form6": text(details.get("form6")),
            "reference": text(details.get("reference")),
            "cod_no": text(details.get("cod_no")),
            "cod_date": parse_date(details.get("cod_date")),
            "issued": inward_status,
            "certificate_type": text(details.get("certificate_type")) or "standard",
            "certificate_fields": _certificate_fields_json(details.get("certificate_fields")),
        }

        def write_material_weight(cell, material, *, removing=False):
            """Keep equivalent formulas; never claim a manual weight was saved when it was ignored."""
            desired = material.get('weight')
            if cell.data_type == 'f':
                refs = {f'{get_column_letter(cols[key] + 1)}{cell.row}': number(material.get(key))
                        for key in ('quantity', 'per_piece_weight') if key in cols}
                computed = formula_number(cell.value, refs)
                if (not removing and desired is not None and computed is not None
                        and math.isclose(computed, desired, rel_tol=1e-9, abs_tol=1e-8)):
                    return
                if externally_protected(sheet) and cell.protection.locked:
                    raise ValueError(f'{cell.coordinate} is a protected weight formula. '
                                     'Keep this correction in the PDF only, or unlock that weight cell in Excel before saving.')
                # Prove that this is the same bounded, same-row arithmetic the
                # reader supports, even when its new inputs are text/blank.
                supported = computed is not None
                for quantity, per_piece in ((1, 2), (2, 3), (3, 5)):
                    probes = {f'{get_column_letter(cols[key] + 1)}{cell.row}': value
                              for key, value in (('quantity', quantity), ('per_piece_weight', per_piece))
                              if key in cols}
                    supported = supported or formula_number(cell.value, probes) is not None
                if not supported:
                    raise ValueError(f'{cell.coordinate} has a weight formula the app cannot safely replace. '
                                     'Keep this correction in the PDF only, or enter a confirmed weight in that Excel cell before saving.')
                cell.value = None if removing else desired
                return
            if externally_protected(sheet) and cell.protection.locked:
                raise ValueError(f'{cell.coordinate} is a protected weight cell; it was not changed.')
            cell.value = None if removing or desired in (None, '') else desired

        for key, value in values.items():
            # Omitted item fields and continuation material rows stay intact.
            if key not in details and key not in {"sn", "issued"}:
                continue
            if key not in cols:
                continue
            if key == 'sn' and not is_new_record:
                # The displayed identity can contain zero padding supplied by
                # Excel's number format. Keep its original numeric cell and
                # legacy metadata key when editing an existing inward.
                continue
            if (key in {'cod_no', 'cod_date'} and original_batch
                    and original_batch.get('hdd_on_main_row')):
                continue  # Only a successful recycling write may move the HDD tracking.
            if material_override is not None and key in {'item', 'quantity', 'per_piece_weight', 'weight'}:
                # Real saves write all material rows together below; PDF-only
                # projections are applied directly to the returned batch.
                continue
            cell = sheet.cell(target_row, cols[key] + 1)
            if key == 'weight' and cell.data_type == 'f' and validate_cells:
                _check_batch_merge(sheet, cell, target_row, cols)
                material = {'weight': value, **{
                    field: sheet.cell(target_row, cols[field] + 1).value
                    for field in ('quantity', 'per_piece_weight') if field in cols}}
                write_material_weight(cell, material)
                continue
            if validate_cells:
                if key in {'item', 'quantity', 'per_piece_weight', 'weight'}:
                    if any(cell.coordinate in area for area in sheet.merged_cells.ranges):
                        raise ValueError('Excel material lines share merged cells. Choose PDF only, or separate the material cells in Excel before saving there.')
                else:
                    _check_batch_merge(sheet, cell, target_row, cols)
                if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            cell.value = value
        if material_override is not None and not render_material_only:
            for index, row_number in enumerate(material_rows):
                material = material_override[index] if index < len(material_override) else {}
                for key in ("item", "quantity", "per_piece_weight", "weight"):
                    if key not in cols:
                        continue
                    cell = sheet.cell(row_number, cols[key] + 1)
                    if validate_cells:
                        if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                            raise ValueError('Excel material lines share merged cells. Choose PDF only, or separate the material cells in Excel before saving there.')
                        if key == 'weight':
                            write_material_weight(cell, material, removing=index >= len(material_override))
                            continue
                        if cell.data_type == "f" or (externally_protected(sheet) and cell.protection.locked):
                            raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
                    value = material.get(key)
                    cell.value = value if value not in (None, "") else None
            for row_number in summary_rows:
                if not validate_cells:
                    # Render-only staging may remove summaries in memory. The
                    # original workbook and its formula cells remain untouched.
                    if row_number not in material_rows[:len(material_override)]:
                        for key in ('item', 'quantity', 'per_piece_weight', 'weight'):
                            sheet.cell(row_number, cols[key] + 1).value = None
                    if text(sheet.cell(row_number, cols['sn'] + 1).value):
                        sheet.cell(row_number, cols['sn'] + 1).value = None
                elif 'weight' in cols:
                    cell = sheet.cell(row_number, cols['weight'] + 1)
                    if cell.data_type != 'f':
                        if externally_protected(sheet) and cell.protection.locked:
                            raise ValueError(f'{cell.coordinate} is protected; its total was not changed.')
                        cell.value = sum(material['weight'] or 0 for position, material in zip(material_rows, material_override)
                                         if position < row_number)
                    elif shrink_material_rows and shrink_retained_end is not None:
                        weight_column = get_column_letter(cols['weight'] + 1)
                        formula = re.sub(r'\s|\$', '', text(cell.value)).upper()
                        original_end = max(shrink_material_rows + [shrink_retained_end])
                        if formula == f'=SUM({weight_column}{target_row}:{weight_column}{original_end})':
                            cell.value = f'=SUM({weight_column}{target_row}:{weight_column}{shrink_retained_end})'
        if validate_cells:
            from openpyxl.worksheet.table import TableColumn
            final_width = register_content_width(sheet, header_row)
            for table, bounds in register_tables:
                # Extend simple register tables without shrinking prepared
                # blank table rows or moving an existing totals row.
                if final_width > bounds[2]:
                    next_id = max([column.id for column in table.tableColumns] or [0]) + 1
                    table.tableColumns.extend(TableColumn(
                        id=next_id + offset, name=text(sheet.cell(header_row, column).value))
                        for offset, column in enumerate(range(bounds[2] + 1, final_width + 1)))
                end_row = bounds[3] if table.totalsRowCount else max(
                    bounds[3], register_content_end(sheet, header_row))
                table.ref = (f'{get_column_letter(bounds[0])}{header_row}:'
                             f'{get_column_letter(final_width)}{end_row}')
                if table.autoFilter is not None and not table.totalsRowCount:
                    table.autoFilter.ref = table.ref
        buffer = BytesIO()
        book.save(buffer)
        updated = buffer.getvalue()
    finally:
        book.close()

    if validate_cells:
        updated = write_details(updated, sheet.title, case_id,
                                {key:value for key,value in details.items() if key not in cols})
        shrink_warning = None
        if shrink_material_rows:
            from register_rows import delete_empty_register_rows
            try:
                updated = delete_empty_register_rows(updated, shrink_material_rows, material_total_cells=shrink_total_cells)
            except ValueError:
                shrink_warning = ('The material change was saved, but Excel kept an extra blank row because its layout '
                                  'could not be safely compacted automatically. Review the workbook before reusing that row.')
        updated = order_register_for_save(updated, session, fit_inward=(sheet.title, case_id))
        if shrink_warning:
            session.order_warning = ' '.join(filter(None, (session.order_warning, shrink_warning)))

    parsed = read_batches(BytesIO(updated))
    matches = [batch for batch in parsed if batch["case_id"] == case_id]
    if len(matches) != 1:
        raise ValueError("The entered certificate details could not be verified in the workbook.")
    if status_value(matches[0]["issued"]) != status_value(inward_status):
        raise ValueError(f"The entered inward details were not marked {inward_status} in the workbook.")
    if material_override is not None:
        if validate_cells:
            persisted = matches[0]['items']
            if len(persisted) != len(material_override):
                raise ValueError('The saved Excel material rows do not match the entered assets. The original workbook was kept.')
            for actual, intended in zip(persisted, material_override):
                for key in ('item', 'quantity', 'per_piece_weight', 'weight'):
                    if key not in cols:
                        continue
                    left, right = actual.get(key), intended.get(key)
                    left_number, right_number = number(left), number(right)
                    same = (math.isclose(left_number, right_number, rel_tol=1e-9, abs_tol=1e-8)
                            if key != 'item' and left_number is not None and right_number is not None
                            else text(left) == text(right))
                    if not same:
                        raise ValueError('The saved Excel material weights or quantities do not match the entered assets. '
                                         'The original workbook was kept; choose PDF only or correct those material cells in Excel.')
        matches[0]["items"] = material_override
        matches[0]["total_weight"] = sum(item["weight"] or 0 for item in material_override)
        certificate_weight = matches[0].get("certificate_weight")
        if certificate_weight is not None and certificate_weight > 0:
            matches[0]["display_weight"] = certificate_weight
            matches[0]["weight_source"] = "certificate"
        else:
            matches[0]["display_weight"] = matches[0]["total_weight"]
            matches[0]["weight_source"] = "inward"
        material_issue_prefixes = ("no item rows", "missing item description", "invalid or negative quantity", "invalid or negative per_piece_weight",
                                   "invalid weight", "one or more item weights", "total weight must be", "an item has an invalid")
        matches[0]["issues"] = [issue for issue in matches[0].get("issues", [])
                                 if not issue.startswith(material_issue_prefixes)]
        _, material_issues = material_items(details)
        matches[0]["issues"].extend(material_issues)
    for key in ("certificate_type", "certificate_fields"):
        if key not in details:
            continue
        expected = values[key]
        if key == "certificate_fields":
            expected = json.loads(expected)
        default = "standard" if key == "certificate_type" else {}
        if matches[0].get(key, default) != expected:
            raise ValueError(f"Excel {key} could not be verified; the original workbook was kept.")
    if validate_cells:
        from inward_followup import FIELDS, stage_followup
        if any(key in details for key in FIELDS):
            values = {key:details.get(key, matches[0].get(key, '')) for key in FIELDS}
            updated, matches[0] = stage_followup(updated, case_id, values)
    return updated, matches[0]


def upsert_certificate_details(path, backup_dir, details):
    """Persist a detail-form row using the same staging logic as issuance."""
    with WorkbookSession(path, backup_dir) as session:
        updated, batch = stage_certificate_details(session, details)
        session._persist(updated)
        return batch
