"""Lock source files and atomically update certificate fields in their XLSX package."""
from __future__ import annotations

from contextlib import contextmanager
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
from openpyxl.utils.datetime import to_excel

from certificate_core import batch_content, columns, number, parse_date, read_batches, status_value, text


def _certificate_fields_json(value):
    """Store template fields as a JSON object without altering their strings."""
    if value is None or value == "":
        value = {}
    elif isinstance(value, str):
        value = json.loads(value)
    if not isinstance(value, dict):
        raise ValueError("Certificate Fields must be a JSON object.")
    return json.dumps(value, ensure_ascii=False, allow_nan=False)


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
    attrs["t"] = kind
    if style is not None:
        attrs["s"] = str(style)
    prefix = match.group("p") if match else row_match.group("p")
    payload = f"<{prefix}is><{prefix}t>{escape(str(value))}</{prefix}t></{prefix}is>" if kind == "inlineStr" else f"<{prefix}v>{value}</{prefix}v>"
    replacement = "<" + prefix + "c " + " ".join(f"{k}={quoteattr(v)}" for k, v in attrs.items()) + ">" + payload + f"</{prefix}c>"
    if match:
        row = row[:match.start()] + replacement + row[match.end():]
    else:
        col = column_index_from_string(re.match(r"[A-Z]+", coordinate).group())
        following = next((m for m in cell_matches if column_index_from_string(re.match(r"[A-Z]+", _attributes(m.group())["r"]).group()) > col), None)
        position = following.start() if following else row.rfind("</")
        row = row[:position] + replacement + row[position:]
    return xml[:row_match.start()] + row + xml[row_match.end():]


def _ensure_status_validation(xml, status_column, row_limit):
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
                  f'sqref="{status_column}2:{status_column}{row_limit}">'
                  f'<{prefix}formula1>"Ready,Alert,N/A,Blank,Generated,Generated Manual,Completed"</{prefix}formula1>'
                  f'</{prefix}dataValidation>')
    if match:
        block = match.group()
        validations = list(_element(block, "dataValidation").finditer(block))
        remaining = []
        inserted = False
        for item in validations:
            fragment = item.group()
            if includes_column(_attributes(fragment).get("sqref", "")):
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


def _update_issued_status_formatting(xml, coordinate):
    """Keep old plant-workbook dropdowns compatible with the PDF workflow."""
    match = re.fullmatch(r"([A-Z]+)(\d+)", coordinate)
    if not match:
        return xml
    status_column = match.group(1)
    xml = _ensure_status_validation(xml, status_column, 10000)

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
        if not _status_column_is_in_range(_attributes(block).get("sqref", ""), status_column):
            return block
        rules = list(_element(block, "cfRule").finditer(block))
        remove = [rule for rule in rules if _is_status_color_rule(rule.group(), status_column)]
        if not remove:
            return block
        for rule in reversed(remove):
            block = block[:rule.start()] + block[rule.end():]
        if not _element(block, "cfRule").search(block):
            return ""
        return block

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


def _add_status_conditional_formatting(sheet_xml, status_column, sn_column, dxf_ids, row_limit):
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
        rules.append(
            f'<{prefix}cfRule type="expression" dxfId="{dxf_ids[label]}" '
            f'priority="1" stopIfTrue="1"><{prefix}formula>{escape(formula)}'
            f'</{prefix}formula></{prefix}cfRule>'
        )
    block = (f'<{prefix}conditionalFormatting sqref="{status_column}2:{status_column}{row_limit}">'
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
                                      status_column, alert_columns, dxf_id,
                                      row_limit):
    """Highlight identifying cells for overdue or invalid unfinished inward rows."""
    if not alert_columns:
        return sheet_xml
    refs = " ".join(f"{column}2:{column}{row_limit}" for column in alert_columns)
    date_ref = f"${date_column}2"
    trimmed_date = f"TRIM({date_ref})"
    parsed_text_date = f"DATE(VALUE(RIGHT({trimmed_date},4)),VALUE(MID({trimmed_date},4,2)),VALUE(LEFT({trimmed_date},2)))"
    text_date_due = (
        f'IF({trimmed_date}="",TRUE,IFERROR(AND(LEN({trimmed_date})=10,'
        f'OR(MID({trimmed_date},3,1)=".",MID({trimmed_date},3,1)="/",'
        f'MID({trimmed_date},3,1)="-"),MID({trimmed_date},6,1)=MID({trimmed_date},3,1),'
        f'DAY({parsed_text_date})=VALUE(LEFT({trimmed_date},2)),'
        f'MONTH({parsed_text_date})=VALUE(MID({trimmed_date},4,2)),'
        f'YEAR({parsed_text_date})=VALUE(RIGHT({trimmed_date},4)),'
        f'OR({parsed_text_date}>TODAY(),TODAY()-{parsed_text_date}>=25)),TRUE))'
    )
    age_due = (
        f'IF(ISNUMBER({date_ref}),OR(INT({date_ref})>TODAY(),'
        f'TODAY()-INT({date_ref})>=25),{text_date_due})'
    )
    formula = (
        f'AND(${sn_column}2<>"",{age_due},UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    legacy_formula = (
        f'AND(${sn_column}2<>"",OR(NOT(ISNUMBER({date_ref})),{date_ref}>TODAY(),'
        f'TODAY()-{date_ref}>=25),UPPER(TRIM(${status_column}2))<>"COMPLETED",'
        f'SUBSTITUTE(UPPER(TRIM(${status_column}2))," ","")<>"N/A")'
    )
    managed_formulas = {formula, legacy_formula}

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
            if rule_formula == formula:
                current_matches.append((rule, _attributes(rule).get("dxfId"),
                                        _attributes(block).get("sqref")))
    if (len(current_matches) == 1 and current_matches[0][1] == str(dxf_id)
            and current_matches[0][2] == refs):
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
            if rule_formula in managed_formulas:
                managed.append(rule_match)
        for rule_match in reversed(managed):
            block = block[:rule_match.start()] + block[rule_match.end():]
        if managed and not _element(block, "cfRule").search(block):
            return ""
        return block

    sheet_xml = _element(sheet_xml, "conditionalFormatting").sub(remove_managed_rules, sheet_xml)
    escaped_formula = escape(formula)
    existing_priorities = [int(value) for value in re.findall(r'<(?:[\w.-]+:)?cfRule\b[^>]*\bpriority=["\'](\d+)', sheet_xml)]
    priority = max(existing_priorities, default=0) + 1
    prefix_match = re.search(r"<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b", sheet_xml)
    prefix = prefix_match.group("p") if prefix_match else ""
    block = (f'<{prefix}conditionalFormatting sqref="{refs}">'
             f'<{prefix}cfRule type="expression" dxfId="{dxf_id}" priority="{priority}">'
             f'<{prefix}formula>{escaped_formula}</{prefix}formula>'
             f'</{prefix}cfRule></{prefix}conditionalFormatting>')
    insert_at = len(sheet_xml)
    for name in ("dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"):
        found = _element(sheet_xml, name).search(sheet_xml)
        if found:
            insert_at = min(insert_at, found.start())
    if insert_at == len(sheet_xml):
        insert_at = sheet_xml.rfind("</")
    return _renumber_conditional_priorities(sheet_xml[:insert_at] + block + sheet_xml[insert_at:])


def _date_style(styles_xml, existing_index):
    section = _element(styles_xml, "cellXfs").search(styles_xml)
    if section is None:
        raise ValueError("Workbook cell styles are missing.")
    fragments = list(_element(section.group(), "xf").finditer(section.group()))
    original = fragments[existing_index].group()
    opening_end = original.index(">")
    opening = original[:opening_end]
    for key, value in (("numFmtId", "14"), ("applyNumberFormat", "1")):
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

    def ensure_alert_formatting(self, sheet_name):
        """Add a dynamic Excel red highlight for late, unfinished inward rows."""
        book = load_workbook(BytesIO(self.data), data_only=False, read_only=False)
        try:
            if sheet_name not in book.sheetnames:
                return None
            sheet = book[sheet_name]
            cols = columns([cell.value for cell in sheet[1]])
            if not {"sn", "inward_date", "issued", "supplier"}.issubset(cols):
                return None
            # Only the supplier name is the user-facing alert indicator.
            # SN and Reference No. remain normal identifiers in the sheet.
            alert_columns = [get_column_letter(cols["supplier"] + 1)]
            sn_column = get_column_letter(cols["sn"] + 1)
            date_column = get_column_letter(cols["inward_date"] + 1)
            status_column = get_column_letter(cols["issued"] + 1)
            row_limit = max(10000, sheet.max_row + 2000)
            status_labels = {
                "ready": "Ready", "alert": "Alert", "na": "N/A", "blank": "Blank",
                "generated": "Generated", "generated_manual": "Generated Manual",
                "completed": "Completed",
            }
            status_updates = []
            for row_number in range(2, sheet.max_row + 1):
                cell = sheet.cell(row_number, cols["issued"] + 1)
                if cell.data_type == "f" or cell.value is None or not str(cell.value).strip():
                    continue
                canonical = status_labels.get(status_value(cell.value))
                if canonical and cell.value != canonical:
                    status_updates.append((cell.coordinate, canonical))
        finally:
            book.close()

        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, sheet_name)
            styles_xml = source.read("xl/styles.xml").decode("utf-8")
            updated_styles, red_dxf = _add_alert_dxf(styles_xml)
            status_dxf_ids = {"READY": None, "ALERT": None, "NA": None,
                              "BLANK": red_dxf, "GENERATED": None,
                              "GENERATED MANUAL": None, "COMPLETED": None}
            for label, fill, font in (
                ("READY", "FFB6D7A8", "FF006100"),
                ("ALERT", "FFFCE4D6", "FF7F6000"),
                ("NA", "FFFFFFFF", "FF000000"),
                ("GENERATED", "FFFFE699", "FF7F6000"),
                ("GENERATED MANUAL", "FFDDEBF7", "FF1F4E78"),
                ("COMPLETED", "FFFFE699", "FF7F6000"),
            ):
                updated_styles, status_dxf_ids[label] = _add_color_dxf(updated_styles, fill, font)
            original_sheet_xml = source.read(member).decode("utf-8")
            sheet_xml = original_sheet_xml
            for coordinate, canonical in status_updates:
                sheet_xml = _set_cell(sheet_xml, coordinate, canonical, "inlineStr")
            sheet_xml = _update_issued_status_formatting(sheet_xml, f"{status_column}2")
            sheet_xml = _add_status_conditional_formatting(
                sheet_xml, status_column, sn_column, status_dxf_ids, row_limit,
            )
            updated_sheet = _add_alert_conditional_formatting(
                sheet_xml, sn_column=sn_column, date_column=date_column,
                status_column=status_column, alert_columns=alert_columns,
                dxf_id=red_dxf, row_limit=row_limit,
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
        return self._persist(buffer.getvalue())

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
            cols = columns([c.value for c in sheet[1]])
            # Missing output columns are added by commit(). This lets the
            # operator recover old registers without manually editing their
            # headers first.
            if "cod_no" not in cols or "cod_date" not in cols or "issued" not in cols:
                return (current, None, None, None, 0, False, book.epoch)
            cells = [sheet.cell(current["source_row"], cols[key] + 1) for key in ("cod_no", "cod_date", "issued")]
            for cell in cells:
                if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge the COD output cells before issuing.")
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
            cols = columns([cell.value for cell in sheet[1]])
            for key in ("cod_no", "cod_date", "issued"):
                if key not in cols:
                    continue
                cell = sheet.cell(current["source_row"], cols[key] + 1)
                if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge it before issuing.")
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
            cols = columns([c.value for c in sheet[1]])
            labels = {
                "cod_no": "COD No", "cod_date": "COD Date", "issued": "Certificate Issued",
                "supplier": "Supplier Name", "location": "Location", "inward_date": "Inward Date",
                "form6": "Form 6 No", "reference": "Reference No.",
                "certificate_weight": "Certificate Weight (KG)",
                "certificate_type": "Certificate Type", "certificate_fields": "Certificate Fields",
            }
            required = {"cod_no", "cod_date", "issued"}
            if not certificate_only:
                required.add("form6")
            if persisted_overrides:
                required.update(key for key in persisted_overrides if key in labels)
            for key in required:
                if key not in cols:
                    column = sheet.max_column + 1
                    sheet.cell(1, column, labels[key])
                    cols[key] = column - 1

            def check_cell(cell):
                if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge it before issuing.")
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")

            row = current["source_row"]
            output_values = {
                "cod_no": certificate_no,
                "cod_date": certificate_date,
                "issued": target_status,
            }
            for key, value in (persisted_overrides or {}).items():
                if key in labels and key not in {"cod_no", "cod_date", "issued"}:
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
            cols = columns([cell.value for cell in sheet[1]])
            if "issued" not in cols:
                raise ValueError("Add a Certificate Issued column before updating the completion status.")
            cell = sheet.cell(current["source_row"], cols["issued"] + 1)
            if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge it before changing the status.")
            if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                raise ValueError(f"{cell.coordinate} is a formula or protected cell; its status was not changed.")
            coordinate = cell.coordinate
        finally:
            book.close()
        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, current["source_sheet"])
            xml = source.read(member).decode("utf-8")
            xml = _set_cell(xml, coordinate, new_status, "inlineStr")
            xml = _update_issued_status_formatting(xml, coordinate)
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
        candidates = []
        for sheet in book.worksheets:
            cols = columns([cell.value for cell in sheet[1]])
            if {"sn", "supplier", "inward_date"}.issubset(cols):
                candidates.append((sheet, cols))
        if len(candidates) != 1:
            raise ValueError("The workbook must contain one inward register with SN, Supplier Name and Inward Date headings.")
        sheet, cols = candidates[0]
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
            if key in labels and key not in cols and (value not in (None, "") or key in {"form6", "issued", "certificate_type", "certificate_fields"}):
                column = sheet.max_column + 1
                sheet.cell(1, column, labels[key])
                cols[key] = column - 1
        for key in ("item", "form6", "issued"):
            if key not in cols:
                column = sheet.max_column + 1
                sheet.cell(1, column, labels[key])
                cols[key] = column - 1

        target_row = None
        for row_number in range(2, sheet.max_row + 1):
            if text(sheet.cell(row_number, cols["sn"] + 1).value) == case_id:
                if target_row is not None:
                    raise ValueError(f"SN {case_id} appears more than once; correct the workbook before using the detail form.")
                target_row = row_number
        target_row = target_row or sheet.max_row + 1

        values = {
            "sn": case_id,
            "inward_date": parse_date(details.get("inward_date")),
            "supplier": text(details.get("supplier")),
            "location": text(details.get("location")),
            "item": text(details.get("item")),
            "quantity": number(details.get("quantity")),
            "per_piece_weight": number(details.get("per_piece_weight")),
            "weight": number(details.get("weight")),
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
            if validate_cells:
                if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge it before using the detail form.")
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            cell.value = value
        buffer = BytesIO()
        book.save(buffer)
        updated = buffer.getvalue()
    finally:
        book.close()

    parsed = read_batches(BytesIO(updated))
    matches = [batch for batch in parsed if batch["case_id"] == case_id]
    if len(matches) != 1:
        raise ValueError("The entered certificate details could not be verified in the workbook.")
    if status_value(matches[0]["issued"]) != "ready":
        raise ValueError("The entered certificate details were not marked Ready in the workbook.")
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
