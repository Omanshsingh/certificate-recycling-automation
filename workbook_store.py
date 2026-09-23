"""Lock source files and update only the two COD cells in their XLSX package."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime
from io import BytesIO
from html import unescape
import errno
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
from openpyxl.utils.cell import column_index_from_string
from openpyxl.utils.datetime import to_excel

from certificate_core import batch_content, columns, parse_date, read_batches, status_value, text


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

    def prepare(self, batch, certificate_no=None):
        if self.context is None:
            raise PermissionError("The workbook became unavailable during its update. Close it in Excel and run again.")
        matches = [b for b in read_batches(BytesIO(self.data)) if b["case_id"] == batch["case_id"] and b["source_sheet"] == batch["source_sheet"]]
        if len(matches) != 1 or batch_content(matches[0]) != batch_content(batch):
            raise ValueError("The source batch changed; check the worksheet before continuing.")
        current = matches[0]
        if status_value(current["issued"]) != "ready":
            raise ValueError("Certificate Issued must contain Ready before generating a certificate.")
        if current["cod_no"] and current["cod_no"] != certificate_no:
            raise ValueError("This batch already has a different COD number; it was not overwritten.")
        book = load_workbook(BytesIO(self.data), data_only=False)
        try:
            sheet = book[current["source_sheet"]]
            cols = columns([c.value for c in sheet[1]])
            if not {"cod_no", "cod_date"}.issubset(cols):
                raise ValueError("Add COD No and COD Date columns before preparing drafts.")
            cells = [sheet.cell(current["source_row"], cols[key] + 1) for key in ("cod_no", "cod_date")]
            for cell in cells:
                if any(cell.coordinate in merged for merged in sheet.merged_cells.ranges):
                    raise ValueError(f"{cell.coordinate} belongs to merged cells; unmerge the COD output cells before issuing.")
                if cell.data_type == "f" or (sheet.protection.sheet and cell.protection.locked):
                    raise ValueError(f"{cell.coordinate} is a formula or protected cell; it was not overwritten.")
            if text(current["cod_date"]) and not current["cod_no"]:
                raise ValueError("COD Date is already filled but COD No is blank; check this batch first.")
            return current, cells[0].coordinate, cells[1].coordinate, cells[1].style_id, is_date_format(cells[1].number_format), book.epoch
        finally:
            book.close()

    def commit(self, batch, certificate_no, certificate_date):
        current, cod_cell, date_cell, style_id, has_date_style, epoch = self.prepare(batch, certificate_no)
        if text(current["cod_date"]) and parse_date(current["cod_date"]) != certificate_date:
            raise ValueError("COD Date differs from the saved certificate date; it was not overwritten.")
        with ZipFile(BytesIO(self.data)) as source:
            member = _sheet_member(source, current["source_sheet"])
            xml = source.read(member).decode("utf-8")
            changed = {}
            if not has_date_style:
                styles, style_id = _date_style(source.read("xl/styles.xml").decode("utf-8"), style_id)
                changed["xl/styles.xml"] = styles.encode("utf-8")
            xml = _set_cell(xml, cod_cell, certificate_no, "inlineStr")
            xml = _set_cell(xml, date_cell, to_excel(certificate_date, epoch), "n", style_id)
            ET.fromstring(xml)
            changed[member] = xml.encode("utf-8")
            buffer = BytesIO()
            with ZipFile(buffer, "w") as target:
                target.comment = source.comment
                for info in source.infolist():
                    target.writestr(info, changed.get(info.filename, source.read(info.filename)))
        updated = buffer.getvalue()
        verified = next(b for b in read_batches(BytesIO(updated)) if b["case_id"] == current["case_id"])
        if verified["cod_no"] != certificate_no or parse_date(verified["cod_date"]) != certificate_date:
            raise ValueError("Excel update could not be verified; the original workbook was kept.")
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
