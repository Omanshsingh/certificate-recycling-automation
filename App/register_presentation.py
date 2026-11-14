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
from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

from register_protection import externally_protected
from certificate_core import parse_date, select_register
from workbook_store import (_attributes, _date_style, _element, _ensure_status_validation,
                            _patch_cell_styles, _set_cell, _sheet_member, register_content_width)

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
    if definition.rstrip().endswith('/>'):
        definition = definition.rstrip()[:-2] + f'></{prefix}xf>'
    alignment = _element(definition, 'alignment').search(definition)
    value = alignment.group() if alignment else f'<{prefix}alignment/>'
    if top_left:
        value = _set_attribute(value, 'horizontal', 'left')
        value = _set_attribute(value, 'vertical', 'top')
    value = _set_attribute(value, 'wrapText', 1)
    if _attributes(value).get('shrinkToFit') in ('true', '1'):
        value = _set_attribute(value, 'shrinkToFit', 0)
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


def _fit_wrapped_rows(xml, sheet, text_cells, normal_font):
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
    heights = {}
    requirements = []
    for cell, value in text_cells:
        if cell.alignment.text_rotation:
            continue
        area = merges.get((cell.row, cell.column))
        last_col = area.max_col if area else cell.column
        width = sum(widths.get(col, default_width) for col in range(cell.column, last_col + 1))
        if width <= 0:
            continue
        size = cell.font.sz or 11
        measure = _font_measure(cell.font.name or 'Calibri', size, bool(cell.font.b), bool(cell.font.i))
        indent = (cell.alignment.indent or 0) * measure('   ')
        lines = _wrapped_lines(value, max(1, width * digit_width - 6 - indent), measure)
        if lines > 1:
            height = math.ceil((max(default_height, size * 1.35) * lines + 2) * 2) / 2
            requirements.append((cell.row, area.max_row if area else cell.row, height))

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
        available = [heights.get(row, existing_height(row)) for row in rows]
        deficit = required - sum(available)
        for index, (row, height) in enumerate(zip(rows, available)):
            if deficit <= 0:
                break
            target = min(MAX_ROW_HEIGHT, height + math.ceil(deficit * 2 / (len(rows) - index)) / 2)
            if target > height:
                heights[row] = target
                deficit -= target - height

    if not heights:
        return xml

    def fit_row(match):
        fragment = match.group()
        row = int(_attributes(fragment)['r'])
        if row in heights:
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


def normalize_register_presentation(data, *, wrap_only=False, include_validation=False):
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
            if existing is None and not wrap_only:
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
            text_cells = []
            style_ids = {}
            for match in _element(xml, 'c').finditer(xml):
                attrs = _attributes(match.group())
                cell = sheet[attrs['r']]
                if cell.row < header:
                    continue
                is_date = not wrap_only and cell.row > header and (cell.column in date_cols or is_date_format(cell.number_format))
                value = cell.value
                if is_date and isinstance(value, str) and cell.data_type != 'f':
                    parsed = parse_date(value)
                    if parsed:
                        xml = _set_cell(xml, cell.coordinate, to_excel(parsed, book.epoch), 'n')
                        value = None
                if cell.data_type == 'f':
                    cached = _element(match.group(), 'v').search(match.group()) if attrs.get('t') == 'str' else None
                    value = None
                    if cached:
                        # Parse the cached text, including numeric character references.
                        inner = cached.group().partition('>')[2].rpartition('<')[0]
                        value = ET.fromstring(f'<v>{inner}</v>').text
                if isinstance(value, str) and value:
                    text_cells.append((cell, value))
                style_key = (int(attrs.get('s', 0)), is_date)
                if style_key in style_ids:
                    ids[cell.coordinate] = style_ids[style_key]
                    continue
                definition = _wrapped_style(definitions[style_key[0]], prefix, top_left=not wrap_only)
                if is_date:
                    definition = _set_attribute(definition, 'numFmtId', existing)
                    definition = _set_attribute(definition, 'applyNumberFormat', 1)
                if definition not in known:
                    known[definition] = len(definitions)
                    definitions.append(definition)
                ids[cell.coordinate] = style_ids[style_key] = known[definition]
            xml = _patch_cell_styles(xml, ids)
            xml = _fit_wrapped_rows(xml, sheet, text_cells, book._fonts[0])
            xml = _wrap_blank_columns(xml, sheet, header, definitions, known, prefix)
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
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, changes.get(info.filename, archive.read(info.filename)))
            return output.getvalue()
    finally:
        book.close()


def _wrap_blank_columns(xml, sheet, header, definitions, known, style_prefix):
    """Use a wrapped column default for future direct entry into blank cells."""
    width = register_content_width(sheet, header)
    section = _element(xml, 'cols').search(xml)
    prefix = _element(xml, 'sheetData').search(xml).group('p')
    columns = {}
    if section:
        for match in _element(section.group(), 'col').finditer(section.group()):
            fragment = match.group()
            attrs = _attributes(fragment)
            columns[int(attrs['min'])] = (int(attrs['max']), fragment)
    additions = []
    index = 1
    for first, (last, fragment) in sorted(columns.items()):
        if index < first and index <= width:
            additions.append((index, min(first - 1, width), None))
        additions.append((first, last, fragment))
        index = max(index, last + 1)
    if index <= width:
        additions.append((index, width, None))
    output = []
    for first, last, fragment in additions:
        if first > width:
            output.append(fragment)
            continue
        if fragment and last > width:
            output_fragment = _set_attribute(fragment, 'max', width)
        else:
            output_fragment = fragment
        if not output_fragment:
            output_fragment = f'<{prefix}col min="{first}" max="{min(last, width)}"/>'
        original = definitions[int(_attributes(output_fragment).get('style', 0))]
        definition = _wrapped_style(original, style_prefix)
        if definition not in known:
            known[definition] = len(definitions)
            definitions.append(definition)
        output.append(_set_attribute(output_fragment, 'style', known[definition]))
        if fragment and last > width:
            output.append(_set_attribute(fragment, 'min', width + 1))
    if section and not section.group().rstrip().endswith('/>'):
        original = section.group()
        block = original[:original.index('>') + 1] + ''.join(output) + original[original.rfind('</'):]
    else:
        block = f'<{prefix}cols>{"".join(output)}</{prefix}cols>'
    if section:
        return xml[:section.start()] + block + xml[section.end():]
    pos = _element(xml, 'sheetData').search(xml).start()
    return xml[:pos] + block + xml[pos:]
