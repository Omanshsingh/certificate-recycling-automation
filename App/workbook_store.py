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
import uuid
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.datetime import to_excel
from openpyxl.styles.stylesheet import write_stylesheet
from openpyxl.xml.functions import tostring

from certificate_core import batch_content, columns, inward_details_complete, material_items, number, parse_date, read_batches, register_header, select_register, status_value, text
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
        bottom = sum(bool(cell.border.bottom.style) for cell in cells)
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


def copy_client_row_styles(sheet, source_row, target_rows, width=None):
    """Copy cell formatting and row height for a newly appended client block."""
    width = width or sheet.max_column
    source_height = sheet.row_dimensions[source_row].height
    for target_row in target_rows:
        for column in range(1, width + 1):
            sheet.cell(target_row, column)._style = copy(sheet.cell(source_row, column)._style)
        if source_height is not None:
            sheet.row_dimensions[target_row].height = source_height


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
    for area in sheet.merged_cells.ranges:
        if cell.coordinate not in area:
            continue
        next_inward = any(text(sheet.cell(row, cols['sn'] + 1).value)
                          for row in range(source_row + 1, area.max_row + 1))
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
def exclusive_file(path):
    """Windows denies other readers/writers but permits our atomic replacement."""
    path = Path(path)
    if os.name == "nt":
        import ctypes
        import msvcrt
        from ctypes import wintypes
        api = ctypes.WinDLL("kernel32", use_last_error=True)
        api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        api.CreateFileW.restype = wintypes.HANDLE
        handle = api.CreateFileW(str(path.resolve()), 0xC0000000, 4, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            code = ctypes.get_last_error()
            if code in (32, 33):
                raise LockBusyError(f"Cannot update {path.name}. Save and close it in Excel, then try again. (Windows error {code})")
            raise PermissionError(f"Cannot update {path.name}. Save and close it in Excel, and check write access. (Windows error {code})")
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


def _set_cell(xml, coordinate, value, kind, style=None):
    row_number = re.search(r"\d+$", coordinate).group()
    row_match = next((m for m in _element(xml, "row").finditer(xml) if _attributes(m.group()).get("r") == row_number), None)
    if not row_match:
        raise ValueError(f"Workbook row {row_number} is missing; no cells were changed.")
    row = row_match.group()
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
                  f'<{prefix}formula1>"Ready,Alert,N/A,Generated,Generated Manual,Completed"</{prefix}formula1>'
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
        "GENERATED", "GENERATED MANUAL", "CREATED", "DRAFTED", "COMPLETED", "SENT",
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
        ("GENERATED MANUAL", f'UPPER(TRIM(${status_column}2))="GENERATED MANUAL"'),
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
    """Show overdue rows in red and date-quality issues beside the date in amber."""
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


def add_tracking_column(data):
    """Add only Certificate Issued, preserving the existing register package."""
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header_row, cols = select_register(book)
        if 'issued' in cols:
            return data
        if sheet.protection.sheet:
            raise ValueError('Unprotect the register before adding Certificate Issued.')
        last_column = max(cell.column for cell in sheet[header_row] if text(cell.value))
        column = last_column + 1
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
                end = batches[index + 1]['source_row'] - 1 if index + 1 < len(batches) else sheet.max_row
                for row in range(start, end + 1):
                    style = sheet.cell(row, cols['sn'] + 1).style_id
                    xml = _set_cell(xml, f'{get_column_letter(column)}{row}', '', 'blank', style)
                if end > start:
                    merges.append(f'<{prefix}mergeCell ref="{get_column_letter(column)}{start}:{get_column_letter(column)}{end}"/>')
            # Append the new status merges without normalizing existing cells.
            merge_section = _element(xml, 'mergeCells').search(xml)
            if merges:
                if merge_section:
                    block = merge_section.group(); pos = block.rfind('</')
                    block = block[:pos] + ''.join(merges) + block[pos:]
                    count = len(list(_element(block, 'mergeCell').finditer(block)))
                    block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{count}"', block, count=1)
                    xml = xml[:merge_section.start()] + block + xml[merge_section.end():]
                else:
                    pos = min([m.start() for tag in ('phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'drawing', 'extLst')
                               if (m := _element(xml, tag).search(xml))] or [xml.rfind('</')])
                    xml = xml[:pos] + f'<{prefix}mergeCells count="{len(merges)}">' + ''.join(merges) + f'</{prefix}mergeCells>' + xml[pos:]
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
                block = re.sub(r'\bref=["\x27][^"\x27]*["\x27]', f'ref="A1:{get_column_letter(column)}{sheet.max_row}"', dimension.group())
                xml = xml[:dimension.start()] + block + xml[dimension.end():]
            styles = sync_generated_styles(book, source.read('xl/styles.xml').decode('utf-8'))
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = source.comment
                for info in source.infolist():
                    payload = xml.encode('utf-8') if info.filename == member else styles.encode('utf-8') if info.filename == 'xl/styles.xml' else source.read(info.filename)
                    target.writestr(copy(info), payload)
            return output.getvalue()
    finally:
        book.close()


class WorkbookSession:
    def __init__(self, path, backup_dir):
        self.path = Path(path)
        self.backup_dir = Path(backup_dir)
        self.backup_path = None
        self.context = None
        self.before_persist = None

    def __enter__(self):
        self.context = exclusive_file(self.path)
        self.stream = self.context.__enter__()
        try:
            self.data = self.stream.read()
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

    def ensure_alert_formatting(self, sheet_name, batches=None, as_of=None):
        """Refresh colors and promote complete, unselected inwards to Ready."""
        prepared = add_tracking_column(self.data)
        if prepared != self.data:
            self._persist(prepared)
            fresh = {b['case_id']:b for b in read_batches(BytesIO(self.data)) if b['source_sheet'] == sheet_name}
            for batch in batches or []:
                if batch['source_sheet'] == sheet_name and batch['case_id'] in fresh:
                    batch.update(fresh[batch['case_id']])
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
                "generated": "Generated", "generated_manual": "Generated Manual",
                "completed": "Completed",
            }
            status_updates = []
            for row_number in range(first_row, sheet.max_row + 1):
                cell = sheet.cell(row_number, cols["issued"] + 1)
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    continue
                batch = sheet_batches.get(row_number)
                if any(cell.coordinate in area for area in sheet.merged_cells.ranges):
                    if not batch:
                        continue
                    try:
                        _check_batch_merge(sheet, cell, row_number, cols)
                    except ValueError:
                        continue
                canonical = status_labels.get(status_value(cell.value))
                if canonical == "" and batch and inward_details_complete(batch, as_of):
                    canonical = "Ready"
                if canonical is not None and cell.value != (canonical or None):
                    status_updates.append((cell.coordinate, canonical))
        finally:
            book.close()

        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, sheet_name)
            styles_xml = source.read("xl/styles.xml").decode("utf-8")
            updated_styles, red_dxf = _add_alert_dxf(styles_xml)
            updated_styles, date_attention_dxf = _add_color_dxf(
                updated_styles, "FFFFF2CC", "FF7F6000")
            status_dxf_ids = {"READY": None, "ALERT": None, "NA": None,
                              "BLANK": None, "GENERATED": None,
                              "GENERATED MANUAL": None, "COMPLETED": None}
            for label, fill, font in (
                ("READY", "FFB6D7A8", "FF006100"),
                ("ALERT", "FFFCE4D6", "FF7F6000"),
                ("NA", "FFFFFFFF", "FF000000"),
                ("BLANK", "FFFFFFFF", "FF000000"),
                ("GENERATED", "FFFFE699", "FF7F6000"),
                ("GENERATED MANUAL", "FFDDEBF7", "FF1F4E78"),
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
            if updated_styles == styles_xml and updated_sheet == original_sheet_xml:
                return self.backup_path
            ET.fromstring(updated_styles)
            ET.fromstring(updated_sheet)
            changed = {"xl/styles.xml": updated_styles.encode("utf-8"),
                       member: updated_sheet.encode("utf-8")}
            buffer = BytesIO()
            with ZipFile(buffer, "w") as target:
                target.comment = source.comment
                for info in source.infolist():
                    target.writestr(info, changed.get(info.filename, source.read(info.filename)))
        backup = self._persist(buffer.getvalue())
        for coordinate, canonical in status_updates:
            batch = sheet_batches.get(int(re.search(r"\d+$", coordinate).group()))
            if batch is not None:
                batch["issued"] = canonical
        return backup

    def prepare(self, batch, certificate_no=None, **_legacy_options):
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and run again.")
        matches = [b for b in read_batches(BytesIO(self.data)) if b["case_id"] == batch["case_id"] and b["source_sheet"] == batch["source_sheet"]]
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
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            return (current, cells[0].coordinate, cells[1].coordinate, cells[2].coordinate,
                    cells[1].style_id, is_date_format(cells[1].number_format), book.epoch)
        finally:
            book.close()

    def commit(self, batch, certificate_no, certificate_date, target_status="Generated", overrides=None,
               staged_data=None, allow_reissue=False, new_batch=False):
        if status_value(target_status) != "generated":
            raise ValueError("The app writes Generated after it creates the PDF. Set Completed only after the certificate has been sent.")
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and run again.")
        original_batches = read_batches(BytesIO(self.data))
        matches = [b for b in original_batches
                   if b["case_id"] == batch["case_id"] and b["source_sheet"] == batch["source_sheet"]]
        if new_batch:
            if any(b["case_id"] == batch["case_id"] for b in original_batches):
                raise ValueError("This Client ID already exists; refresh before creating a new inward row.")
            if staged_data is None:
                raise ValueError("A new inward row requires staged certificate details.")
            current = batch
        else:
            if len(matches) != 1 or batch_content(matches[0]) != batch_content(batch):
                raise ValueError("The source batch changed; check the worksheet before continuing.")
            current = matches[0]
        if not allow_reissue and status_value(current["issued"]) != "ready":
            raise ValueError("The row is no longer marked Ready. Choose Ready in Excel before generating its PDF.")

        # Staging may replace formula/protected values for a render-only view.
        # Never let that remove the original output-cell protection checks.
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
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
        finally:
            original.close()

        working_data = self.data if staged_data is None else staged_data
        if staged_data is not None:
            staged_matches = [b for b in read_batches(BytesIO(staged_data))
                              if b["case_id"] == current["case_id"]
                              and b["source_sheet"] == current["source_sheet"]]
            if (len(staged_matches) != 1
                    or staged_matches[0]["source_row"] != current["source_row"]
                    or (new_batch and batch_content(staged_matches[0]) != batch_content(current))):
                raise ValueError("The staged inward row changed; the original workbook was kept.")
        # Explicit form issuance without staged data is certificate-only.
        # Preserve the existing override/header contract for normal Ready runs.
        certificate_only = allow_reissue and staged_data is None
        persisted_overrides = None if certificate_only else overrides

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
                    column = sheet.max_column + 1
                    sheet.cell(header_row, column, labels[key])
                    cols[key] = column - 1

            def check_cell(cell):
                _check_batch_merge(sheet, cell, current['source_row'], cols)
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
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
                    cell.number_format = "dd.mm.yyyy"
            buffer = BytesIO()
            book.save(buffer)
            updated = buffer.getvalue()
        finally:
            book.close()

        updated = write_details(updated, current['source_sheet'], current['case_id'],
                                {key:value for key,value in (persisted_overrides or {}).items() if key not in cols})

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
                          expected_status, new_status):
        """Change only Certificate Issued for an operator action or PDF recovery."""
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and try again.")
        if isinstance(expected_status, str):
            expected_status = (expected_status,)
        expected = {status_value(value) for value in expected_status}
        target = status_value(new_status)
        target_labels = {"ready": "Ready", "generated": "Generated",
                         "generated_manual": "Generated Manual", "completed": "Completed"}
        if target not in target_labels:
            raise ValueError("Only Ready, Generated, Generated Manual or Completed can be recorded by this action.")
        new_status = target_labels[target]
        matches = [batch for batch in read_batches(BytesIO(self.data))
                   if batch["case_id"] == str(case_id)
                   and batch["source_sheet"] == source_sheet
                   and batch["source_row"] == int(source_row)]
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
            if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                raise ValueError(f"{cell.coordinate} is a formula or protected cell; its status was not changed.")
            coordinate = cell.coordinate
        finally:
            book.close()
        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, current["source_sheet"])
            xml = source.read(member).decode("utf-8")
            xml = _set_cell(xml, coordinate, new_status, "inlineStr")
            xml = _update_issued_status_formatting(xml, coordinate, current.get('source_header_row', 1) + 1)
            ET.fromstring(xml)
            buffer = BytesIO()
            with ZipFile(buffer, "w") as target_archive:
                target_archive.comment = source.comment
                for info in source.infolist():
                    target_archive.writestr(info, xml.encode("utf-8") if info.filename == member
                                             else source.read(info.filename))
        updated = buffer.getvalue()
        verified = next(batch for batch in read_batches(BytesIO(updated))
                        if batch["case_id"] == str(case_id)
                        and batch["source_sheet"] == source_sheet
                        and batch["source_row"] == int(source_row))
        if (status_value(verified["issued"]) != target
                or verified["cod_no"] != current["cod_no"]
                or parse_date(verified["cod_date"]) != parse_date(current["cod_date"])):
            raise ValueError("The workbook status update could not be verified; the original file was kept.")
        return self._persist(updated)

    def _persist(self, updated):
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


def stage_certificate_details(session, details, validate_cells=True):
    """Build a Ready form view without writing or changing session.data.

    Skip source-cell checks only for a certificate-only render. Commit still
    checks the original tracking cells before updating COD, date and status.
    """
    details = dict(details or {})
    case_id = text(details.get("case_id"))
    if not case_id:
        raise ValueError("Client ID / SN is required.")
    book = load_workbook(BytesIO(session.data), data_only=False,
                         keep_vba=session.path.suffix.casefold() == ".xlsm")
    try:
        sheet, header_row, cols = select_register(book)
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
        for key, value in details.items():
            if key in labels and key not in cols and (not validate_cells or key == 'issued') and (value not in (None, "") or key in {"form6", "issued", "certificate_type", "certificate_fields"}):
                column = sheet.max_column + 1
                sheet.cell(header_row, column, labels[key])
                cols[key] = column - 1
        for key in (("item", "form6", "issued") if not validate_cells else ('issued',)):
            if key not in cols:
                column = sheet.max_column + 1
                sheet.cell(header_row, column, labels[key])
                cols[key] = column - 1

        if "items" in details and not validate_cells:
            for key in ("quantity", "per_piece_weight", "weight"):
                if key not in cols:
                    column = sheet.max_column + 1
                    sheet.cell(header_row, column, labels[key])
                    cols[key] = column - 1

        target_row = None
        for row_number in range(header_row + 1, sheet.max_row + 1):
            if text(sheet.cell(row_number, cols["sn"] + 1).value) == case_id:
                if target_row is not None:
                    raise ValueError(f"SN {case_id} appears more than once; correct the workbook before using the detail form.")
                target_row = row_number
        is_new_record = target_row is None
        target_row = target_row or sheet.max_row + 1

        material_override = material_items(details)[0] if "items" in details else None
        original_batch = next((batch for batch in read_batches(BytesIO(session.data))
                               if batch['case_id'] == case_id and batch['source_sheet'] == sheet.title), None)
        summary_rows = set((original_batch or {}).get('summary_rows', []))
        material_rows = []
        if material_override is not None:
            material_rows = [target_row]
            following_primary = False
            for row_number in range(target_row + 1, sheet.max_row + 1):
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
                next_row = max(sheet.max_row + 1, target_row + 1)
                material_rows.extend(range(next_row, next_row + len(material_override) - len(material_rows)))

        if is_new_record:
            template_row = formatted_client_template(sheet, header_row, cols, len(material_rows) > 1)
            continuation_row = (template_row + 1 if any(
                area.min_row == template_row and area.max_row > template_row
                for area in sheet.merged_cells.ranges) else template_row)
            style_rows = material_rows or [target_row]
            for index, row_number in enumerate(style_rows):
                source_row = template_row if index == 0 else continuation_row
                copy_client_row_styles(sheet, source_row, [row_number], sheet.max_column)
            if len(style_rows) > 1 and validate_cells:
                material_columns = {cols[key] + 1 for key in
                                    ('item', 'quantity', 'per_piece_weight', 'weight') if key in cols}
                for area in list(sheet.merged_cells.ranges):
                    if (area.min_row == template_row and area.min_col == area.max_col
                            and area.max_row > template_row and area.min_col not in material_columns):
                        column = get_column_letter(area.min_col)
                        sheet.merge_cells(f'{column}{style_rows[0]}:{column}{style_rows[-1]}')

        values = {
            "sn": case_id,
            "inward_date": parse_date(details.get("inward_date")),
            "supplier": text(details.get("supplier")),
            "location": text(details.get("location")),
            "item": material_override[0]["item"] if material_override else text(details.get("item")),
            "quantity": material_override[0]["quantity"] if material_override else number(details.get("quantity")),
            "per_piece_weight": material_override[0]["per_piece_weight"] if material_override else number(details.get("per_piece_weight")),
            "weight": material_override[0]["weight"] if material_override else number(details.get("weight")),
            "certificate_weight": number(details.get("certificate_weight")),
            "form6": text(details.get("form6")),
            "reference": text(details.get("reference")),
            "cod_no": text(details.get("cod_no")),
            "cod_date": parse_date(details.get("cod_date")),
            "issued": "Ready",
            "certificate_type": text(details.get("certificate_type")) or "standard",
            "certificate_fields": _certificate_fields_json(details.get("certificate_fields")),
        }
        for key, value in values.items():
            # Omitted item fields and continuation material rows stay intact.
            if key not in details and key not in {"sn", "issued"}:
                continue
            if key not in cols:
                continue
            cell = sheet.cell(target_row, cols[key] + 1)
            # Weight (KG) is commonly an Excel formula such as =F32*G32.
            # The form already writes the quantity and per-piece inputs, so
            # preserve the formula instead of trying to replace its result.
            if key == "weight" and cell.data_type == "f":
                continue
            if validate_cells:
                if key in {'item', 'quantity', 'per_piece_weight', 'weight'}:
                    if any(cell.coordinate in area for area in sheet.merged_cells.ranges):
                        raise ValueError('Excel material lines share merged cells. Choose PDF only, or separate the material cells in Excel before saving there.')
                else:
                    _check_batch_merge(sheet, cell, target_row, cols)
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            cell.value = value
        if material_override is not None:
            for index, row_number in enumerate(material_rows):
                material = material_override[index] if index < len(material_override) else {}
                for key in ("item", "quantity", "per_piece_weight", "weight"):
                    if key not in cols:
                        continue
                    cell = sheet.cell(row_number, cols[key] + 1)
                    # Let the workbook calculate formula-backed line weights
                    # from the item inputs written above.
                    if key == "weight" and cell.data_type == "f" and row_number not in summary_rows:
                        continue
                    if validate_cells:
                        if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                            raise ValueError('Excel material lines share merged cells. Choose PDF only, or separate the material cells in Excel before saving there.')
                        if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
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
                        if sheet.protection.sheet and cell.protection.locked:
                            raise ValueError(f'{cell.coordinate} is protected; its total was not changed.')
                        cell.value = sum(material['weight'] or 0 for position, material in zip(material_rows, material_override)
                                         if position < row_number)
        buffer = BytesIO()
        book.save(buffer)
        updated = buffer.getvalue()
    finally:
        book.close()

    if validate_cells:
        updated = write_details(updated, sheet.title, case_id,
                                {key:value for key,value in details.items() if key not in cols})

    parsed = read_batches(BytesIO(updated))
    matches = [batch for batch in parsed if batch["case_id"] == case_id]
    if len(matches) != 1:
        raise ValueError("The entered certificate details could not be verified in the workbook.")
    if status_value(matches[0]["issued"]) != "ready":
        raise ValueError("The entered certificate details were not marked Ready in the workbook.")
    if material_override is not None:
        matches[0]["items"] = material_override
        matches[0]["total_weight"] = sum(item["weight"] or 0 for item in material_override)
        certificate_weight = matches[0].get("certificate_weight")
        if certificate_weight is not None and certificate_weight > 0:
            matches[0]["display_weight"] = certificate_weight
            matches[0]["weight_source"] = "certificate"
        else:
            matches[0]["display_weight"] = matches[0]["total_weight"]
            matches[0]["weight_source"] = "inward"
        material_issue_prefixes = ("missing item description", "invalid or negative quantity", "invalid or negative per_piece_weight",
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
    return updated, matches[0]


def upsert_certificate_details(path, backup_dir, details):
    """Persist a detail-form row using the same staging logic as issuance."""
    with WorkbookSession(path, backup_dir) as session:
        updated, batch = stage_certificate_details(session, details)
        session._persist(updated)
        return batch
