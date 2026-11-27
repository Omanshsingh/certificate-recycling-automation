"""Format the plant register without reserializing the workbook."""
from functools import lru_cache
from io import BytesIO
import math
import re
from unicodedata import east_asian_width
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr
from zipfile import ZipFile

from PIL import ImageFont
from register_reader import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

from register_protection import externally_protected
from certificate_core import parse_date, read_batches, select_register
from workbook_store import (_attributes, _date_style, _element, _ensure_status_validation, _inward_blocks,
                            _patch_cell_styles, _set_cell, _sheet_member)

DATE_FORMAT = 'dd/mm/yy'
MAX_ROW_HEIGHT = 409
FONT_FILES = {
    'calibri': ('calibri', 'calibrib', 'calibrii', 'calibriz'),
    'arial': ('arial', 'arialbd', 'ariali', 'arialbi'),
    'times new roman': ('times', 'timesbd', 'timesi', 'timesbi'),
    'segoe ui': ('segoeui', 'segoeuib', 'segoeuii', 'segoeuiz'),
}


def _set_attribute(fragment, key, value):
    end = fragment.index('>')
    opening = fragment[:end]
    attribute = f'{key}={quoteattr(str(value))}'
    pattern = rf'\s{key}\s*=\s*(["\x27]).*?\1'
    if re.search(pattern, opening):
        opening = re.sub(pattern, lambda _: ' ' + attribute, opening, count=1)
    else:
        opening = opening.rstrip('/') + ' ' + attribute + ('/' if opening.endswith('/') else '')
    return opening + fragment[end:]


def _wrapped_style(definition, prefix, *, top_left=False):
    if not top_left:
        return definition
    if definition.rstrip().endswith('/>'):
        definition = definition.rstrip()[:-2] + f'></{prefix}xf>'
    alignment = _element(definition, 'alignment').search(definition)
    value = alignment.group() if alignment else f'<{prefix}alignment/>'
    value = _set_attribute(value, 'horizontal', 'left')
    value = _set_attribute(value, 'vertical', 'top')
    if alignment:
        definition = definition[:alignment.start()] + value + definition[alignment.end():]
    else:
        end = definition.rfind('</')
        definition = definition[:end] + value + definition[end:]
    return _set_attribute(definition, 'applyAlignment', 1)


@lru_cache(maxsize=128)
def _font_measure(name, size, bold=False, italic=False):
    family = name.casefold()
    variants = FONT_FILES.get(family, FONT_FILES['calibri'])
    filename = variants[int(bold) + 2 * int(italic)]
    for candidate in (filename + '.ttf', 'DejaVuSans.ttf'):
        try:
            return ImageFont.truetype(candidate, max(1, round(size * 96 / 72))).getlength
        except OSError:
            continue
    return lambda value: sum(1 if east_asian_width(char) in 'WF' else .6 for char in value) * size * 96 / 72


@lru_cache(maxsize=1024)
def _wrapped_lines(value, width, measure):
    lines = 0
    for paragraph in value.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        lines += 1
        if measure(paragraph) <= width:
            continue
        line = ''
        for token in re.findall(r'\s+|\S+', paragraph.expandtabs(4)):
            if not line and token.isspace():
                continue
            if measure(line + token) <= width:
                line += token
                continue
            if line.strip():
                lines += 1
            line = ''
            if token.isspace():
                continue
            # Split long identifiers as well as ordinary words.
            for char in token:
                if line and measure(line + char) > width:
                    lines += 1
                    line = ''
                line += char
    return lines


def _fit_wrapped_rows(xml, sheet, display_cells, normal_font, fit_rows, bold_columns):
    if not fit_rows:
        return xml
    default_height = sheet.sheet_format.defaultRowHeight or 15
    measure_digit = _font_measure(normal_font.name or 'Calibri', normal_font.sz or 11)
    digit_width = max(1, round(max(measure_digit(digit) for digit in '0123456789')))
    default_width = sheet.sheet_format.defaultColWidth or sheet.sheet_format.baseColWidth or 8
    widths = {}
    for dimension in sheet.column_dimensions.values():
        first = dimension.min or column_index_from_string(dimension.index)
        for col in range(first, (dimension.max or first) + 1):
            widths[col] = 0 if dimension.hidden else dimension.width

    merges = {(area.min_row, area.min_col): area for area in sheet.merged_cells.ranges}
    current = sheet.row_dimensions
    heights = {row: default_height for row in fit_rows
               if not (row in current and current[row].hidden)}
    # A merge crossing the area being fitted may depend on rows outside it.
    for area in merges.values():
        affected = heights.keys() & range(area.min_row, area.max_row + 1)
        if affected and (area.min_row not in fit_rows or area.max_row not in fit_rows):
            for row in affected:
                heights.pop(row)
    requirements = []
    for cell, value in display_cells:
        area = merges.get((cell.row, cell.column))
        last_col = area.max_col if area else cell.column
        last_row = area.max_row if area else cell.row
        width = sum(widths.get(col, default_width) for col in range(cell.column, last_col + 1))
        if width <= 0:
            continue
        if cell.alignment.text_rotation:
            for row in range(cell.row, last_row + 1):
                heights.pop(row, None)
            continue
        size = cell.font.sz or 11
        lines = 1
        if isinstance(value, str) and cell.alignment.wrap_text:
            measure = _font_measure(cell.font.name or 'Calibri', size,
                                    bool(cell.font.b) or cell.column in bold_columns, bool(cell.font.i))
            indent = (cell.alignment.indent or 0) * measure('   ')
            lines = _wrapped_lines(value.rstrip(), max(1, width * digit_width - 6 - indent), measure)
        height = math.ceil((max(default_height, size * 1.35) * lines + (2 if lines > 1 else 0)) * 2) / 2
        requirements.append((cell.row, last_row, height))

    def existing_height(row):
        dimension = current.get(row)
        return (dimension.height if dimension and dimension.height is not None else default_height)

    # Fit ordinary cells first. Add only the remaining height to a vertical
    # merge, accounting for all visible rows it already spans.
    for first, last, required in sorted(requirements, key=lambda item: item[1] - item[0]):
        rows = [row for row in range(first, last + 1)
                if not (row in current and current[row].hidden)]
        if not rows:
            continue
        adjustable = [row for row in rows if row in heights]
        deficit = required - sum(heights.get(row, existing_height(row)) for row in rows)
        for index, row in enumerate(adjustable):
            if deficit <= 0:
                break
            height = heights[row]
            target = min(MAX_ROW_HEIGHT, height + math.ceil(deficit * 2 / (len(adjustable) - index)) / 2)
            if target > height:
                heights[row] = target
                deficit -= target - height

    heights = {row: height for row, height in heights.items() if height != existing_height(row)}
    if not heights:
        return xml

    def fit_row(match):
        fragment = match.group()
        row = int(_attributes(fragment)['r'])
        if row in heights:
            if heights[row] == default_height:
                end = fragment.index('>')
                opening = re.sub(r'\s(?:ht|customHeight)=(["\x27]).*?\1', '', fragment[:end])
                return opening + fragment[end:]
            fragment = _set_attribute(fragment, 'ht', f'{heights[row]:g}')
            fragment = _set_attribute(fragment, 'customHeight', 1)
        return fragment

    section = _element(xml, 'sheetData').search(xml)
    block = section.group()
    prefix = section.group('p')
    rows = {int(_attributes(row.group())['r']): row for row in _element(block, 'row').finditer(block)}
    missing = sorted(heights.keys() - rows.keys(), reverse=True)
    pieces = []
    position = block.index('>') + 1
    pieces.append(block[:position])

    def append_missing_row(row):
        pieces.append(f'<{prefix}row r="{row}" ht="{heights[row]:g}" customHeight="1"/>')

    for row, match in rows.items():
        pieces.append(block[position:match.start()])
        while missing and missing[-1] < row:
            append_missing_row(missing.pop())
        pieces.append(fit_row(match))
        position = match.end()
    for row in reversed(missing):
        append_missing_row(row)
    pieces.append(block[position:])
    return xml[:section.start()] + ''.join(pieces) + xml[section.end():]


def normalize_register_presentation(data, *, wrap_only=False, include_validation=False, fit_inward=None,
                                    fit_register=False):
    if wrap_only and not (fit_inward or fit_register or include_validation):
        return data
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        if externally_protected(sheet):
            return data
        fit_rows = set()
        fit_one = isinstance(fit_inward, (tuple, list)) and len(fit_inward) == 2
        if fit_register or fit_one:
            batches = read_batches(BytesIO(data))
            records = [batch for batch in batches if batch['source_sheet'] == sheet.title]
            blocks = _inward_blocks(sheet, records)
            if not fit_register:
                fit_sheet, fit_case_id = fit_inward
                matches = [batch['source_row'] for batch in records if fit_sheet == sheet.title
                           and str(batch['case_id']).casefold() == str(fit_case_id).casefold()]
                blocks = [block for block in blocks if len(matches) == 1 and block[0] == matches[0]]
            fit_rows = {row for first, last in blocks for row in range(first, last + 1)}
            fit_rows.difference_update(issue['row'] for issue in batches.row_issues)
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            if wrap_only:
                display_cells = []
                for row in _element(xml, 'row').finditer(xml):
                    fragment = row.group()
                    if int(_attributes(fragment)['r']) not in fit_rows:
                        continue
                    for match in _element(fragment, 'c').finditer(fragment):
                        cell_xml = match.group()
                        attrs = _attributes(cell_xml)
                        cell = sheet[attrs['r']]
                        value = _display_cell_value(cell, cell_xml, attrs)
                        if value is not None and value != '':
                            display_cells.append((cell, value))
                bold_columns = {cols[key] + 1 for key in ('issued', 'supplier') if key in cols}
                fitted = _fit_wrapped_rows(xml, sheet, display_cells, book._fonts[0], fit_rows, bold_columns)
                if include_validation and 'issued' in cols:
                    fitted = _ensure_status_validation(fitted, get_column_letter(cols['issued'] + 1),
                                                       max(10000, sheet.max_row + 2000), header + 1)
                if fitted == xml:
                    return data
                return _presentation_package(archive, {member: fitted.encode('utf-8')})

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
            display_cells = []
            style_ids = {}
            for match in _element(xml, 'c').finditer(xml):
                attrs = _attributes(match.group())
                cell = sheet[attrs['r']]
                if cell.row < header:
                    continue
                is_date = cell.row > header and (cell.column in date_cols or is_date_format(cell.number_format))
                value = _display_cell_value(cell, match.group(), attrs)
                if is_date and isinstance(value, str) and cell.data_type != 'f':
                    parsed = parse_date(value)
                    if parsed:
                        xml = _set_cell(xml, cell.coordinate, to_excel(parsed, book.epoch), 'n')
                        value = None
                if cell.row in fit_rows and value is not None and value != '':
                    display_cells.append((cell, value))
                style_key = (int(attrs.get('s', 0)), is_date)
                if style_key in style_ids:
                    ids[cell.coordinate] = style_ids[style_key]
                    continue
                definition = _wrapped_style(definitions[style_key[0]], prefix, top_left=True)
                if is_date:
                    definition = _set_attribute(definition, 'numFmtId', existing)
                    definition = _set_attribute(definition, 'applyNumberFormat', 1)
                if definition not in known:
                    known[definition] = len(definitions)
                    definitions.append(definition)
                ids[cell.coordinate] = style_ids[style_key] = known[definition]
            xml = _patch_cell_styles(xml, ids)
            # The register's status and overdue rules can display bold text.
            bold_columns = {cols[key] + 1 for key in ('issued', 'supplier') if key in cols}
            xml = _fit_wrapped_rows(xml, sheet, display_cells, book._fonts[0], fit_rows, bold_columns)
            if include_validation and 'issued' in cols:
                status = get_column_letter(cols['issued'] + 1)
                xml = _ensure_status_validation(xml, status, max(10000, sheet.max_row + 2000), header + 1)
            block = section.group()
            block = block[:block.rfind('</')] + ''.join(definitions[original_count:]) + block[block.rfind('</'):]
            block = re.sub(r'\bcount=(["\x27]).*?\1', f'count="{len(definitions)}"', block, count=1)
            styles = styles[:section.start()] + block + styles[section.end():]
            changes = {member: xml.encode('utf-8'), 'xl/styles.xml': styles.encode('utf-8')}
            if all(archive.read(name) == value for name, value in changes.items()):
                return data
            return _presentation_package(archive, changes)
    finally:
        book.close()


def _display_cell_value(cell, fragment, attrs):
    if cell.data_type != 'f':
        return cell.value
    cached = _element(fragment, 'v').search(fragment) if attrs.get('t') == 'str' else None
    if cached:
        inner = cached.group().partition('>')[2].rpartition('<')[0]
        return ET.fromstring(f'<v>{inner}</v>').text
    return None


def _presentation_package(archive, changes):
    output = BytesIO()
    with ZipFile(output, 'w') as target:
        target.comment = archive.comment
        for info in archive.infolist():
            content = changes[info.filename] if info.filename in changes else archive.read(info.filename)
            target.writestr(info, content)
    return output.getvalue()
