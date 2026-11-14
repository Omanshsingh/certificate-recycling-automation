"""Apply Excel's editable ranges around issued inward blocks."""

from base64 import b64encode
from copy import copy
from hashlib import sha512
from io import BytesIO
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import column_index_from_string, get_column_letter

from certificate_core import read_batches, select_register, status_value


# A recognizable protection record lets the app maintain its own sheet lock
# without treating a password chosen by the workbook owner as app-owned.
_SALT = b'EcorecoIssuedV1!'
_SPINS = 10_000


def _protection_hash():
    value = sha512(_SALT + 'Ecoreco issued inwards'.encode('utf-16le')).digest()
    for index in range(_SPINS):
        value = sha512(value + index.to_bytes(4, 'little')).digest()
    return b64encode(value).decode('ascii')


_SALT64 = b64encode(_SALT).decode('ascii')
_HASH = _protection_hash()
_ISSUED = {'generated', 'generated_manual', 'completed'}
_FOLLOWUP_COLUMNS = ('comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale')


def _managed(sheet):
    protection = sheet.protection
    return (protection.algorithmName == 'SHA-512' and protection.saltValue == _SALT64
            and protection.spinCount == _SPINS and protection.hashValue == _HASH)


def externally_protected(sheet):
    return sheet.protection.sheet and not _managed(sheet)


def protect_issued_inwards(data):
    """Patch the register sheet only, preserving worksheet values and package parts."""
    from workbook_store import (_element, _patch_cell_styles, _sheet_member,
                                register_content_end, sync_generated_styles)

    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, columns = select_register(book)
        if externally_protected(sheet):
            return data
        batches = read_batches(BytesIO(data))
        end = register_content_end(sheet, header)
        locked = []
        for index, batch in enumerate(batches):
            if (status_value(batch['issued']) in _ISSUED or batch['hdd_cod_no']
                    or batch['pdf_recovery_pending']):
                first = batch['source_row']
                last = batches[index + 1]['source_row'] - 1 if index + 1 < len(batches) else end
                locked.append((first, last))
        if not locked and not _managed(sheet):
            return data

        with ZipFile(BytesIO(data)) as source:
            member = _sheet_member(source, sheet.title)
            original = source.read(member).decode('utf-8')
            old_styles = source.read('xl/styles.xml').decode('utf-8')
            xml = _element(original, 'sheetProtection').sub('', original)
            xml = _element(xml, 'protectedRanges').sub('', xml)
            styles = old_styles
            if locked:
                # Blank cells can inherit an unlocked row or column style.
                # Apply locked styles wherever the template opted out of the default.
                replacements = {}
                column_styles = {}
                width = sheet.max_column
                for dimension in sheet.column_dimensions.values():
                    if dimension.has_style:
                        first_column = dimension.min or column_index_from_string(dimension.index)
                        last_column = min(dimension.max or first_column, width)
                        for column in range(first_column, last_column + 1):
                            column_styles[column] = dimension
                for first, last in locked:
                    for row in sheet.iter_rows(min_row=first, max_row=last, max_col=width):
                        for cell in row:
                            if isinstance(cell, MergedCell):
                                continue
                            base = cell
                            if not cell.has_style:
                                row_style = sheet.row_dimensions.get(cell.row)
                                col_style = column_styles.get(cell.column)
                                base = next((style for style in (row_style, col_style)
                                             if style is not None and style.has_style), cell)
                            if not base.protection.locked:
                                cell._style = copy(base._style)
                                protection = copy(base.protection)
                                protection.locked = True
                                cell.protection = protection
                                replacements[cell.coordinate] = cell.style_id
                if replacements:
                    xml = _patch_cell_styles(xml, replacements)
                    styles = sync_generated_styles(book, styles)

                prefix = _element(xml, 'sheetData').search(xml).group('p')
                protection = (f'<{prefix}sheetProtection sheet="1" algorithmName="SHA-512" '
                              f'hashValue="{_HASH}" saltValue="{_SALT64}" spinCount="{_SPINS}" '
                              'objects="1" scenarios="1" autoFilter="0" formatRows="0"/>')
                editable = []
                start = header + 1
                for first, last in [*locked, (1048577, 1048577)]:
                    if start < first:
                        editable.append(f'<{prefix}protectedRange name="EcorecoEditable{len(editable) + 1}" '
                                        f'sqref="A{start}:XFD{first - 1}"/>')
                    start = last + 1
                for key in _FOLLOWUP_COLUMNS:
                    if key not in columns:
                        continue
                    column = columns[key] + 1
                    letter = get_column_letter(column)
                    merges = [area for area in sheet.merged_cells.ranges
                              if area.min_col <= column <= area.max_col]
                    for first, last in locked:
                        blocked = {row for row in range(first, last + 1)
                                   if sheet.cell(row, column).data_type == 'f'}
                        for area in merges:
                            if (area.min_col != area.max_col or area.min_row < first or area.max_row > last
                                    or area.min_row in blocked):
                                blocked.update(range(max(first, area.min_row), min(last, area.max_row) + 1))
                        start = first
                        for stop in sorted(blocked | {last + 1}):
                            if stop > start:
                                editable.append(
                                    f'<{prefix}protectedRange name="EcorecoDetails{len(editable) + 1}" '
                                    f'sqref="{letter}{start}:{letter}{stop - 1}"/>')
                            start = stop + 1
                if editable:
                    protection += f'<{prefix}protectedRanges>{"".join(editable)}</{prefix}protectedRanges>'
                # Sheet protection and editable ranges precede filters, merges,
                # formatting and other worksheet features in the XLSX schema.
                preceding = [_element(xml, tag).search(xml) for tag in
                             ('sheetData', 'sheetCalcPr')]
                position = max(match.end() for match in preceding if match)
                xml = xml[:position] + protection + xml[position:]
            if xml == original and styles == old_styles:
                return data
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = source.comment
                for info in source.infolist():
                    content = (xml.encode('utf-8') if info.filename == member else
                               styles.encode('utf-8') if info.filename == 'xl/styles.xml' else source.read(info.filename))
                    target.writestr(info, content)
            return output.getvalue()
    finally:
        book.close()
