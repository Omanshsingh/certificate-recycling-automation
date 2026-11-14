"""Source parsing, eligibility, and exact-layout certificate generation."""
from __future__ import annotations

from collections import Counter
from copy import deepcopy
from datetime import date, datetime
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import re
import math
import os
import json

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from excel_values import excel_date, formula_number, weight_number, quantity_number
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

PDF_WIDTH, PDF_HEIGHT = 960, 540
TEMPLATE_DESIGN_WIDTH, TEMPLATE_DESIGN_HEIGHT = 1534, 867
MAX_SNAPSHOT_CACHE_BYTES = 4 * 1024 * 1024
# Vertical grid lines in the supplied certificate template, in source-image pixels.
CERTIFICATE_NUMBER_CELL = (797, 878)
CERTIFICATE_DATE_CELL = (878, 990)
DETAIL_VALUE_CELL = (797, 990)


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


NO_CERTIFICATE_SUPPLIERS = {"reverseecommercepvtltdbookmyjunk"}


def supplier_is_always_not_applicable(value):
    """Known inward supplier that never requires a recycling certificate."""
    return norm(value) in NO_CERTIFICATE_SUPPLIERS


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y",
                "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
                "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y",
                "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    # Preserve the established day-first convention for ambiguous dates.
    # Month-first text is accepted only when it cannot be a valid day-first date.
    match = re.fullmatch(r"(\d{1,2})([/.-])(\d{1,2})\2(\d{4})", text)
    if match and int(match[3]) > 12:
        try:
            return date(int(match[4]), int(match[1]), int(match[3]))
        except ValueError:
            pass
    return None


def number(value):
    """Accept complete finite numbers; never extract a number from bad input."""
    if value is None or isinstance(value, bool):
        return None
    numeric = str(value).strip()
    # Both Western and Indian thousands grouping occur in the inward register.
    integer = r"(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{1,2}(?:,[0-9]{2})+,[0-9]{3})"
    if not re.fullmatch(rf"[+-]?(?:{integer}(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", numeric):
        return None
    result = float(numeric.replace(",", ""))
    return result if math.isfinite(result) else None


def material_items(details):
    """Normalize the material editor's rows and calculate missing per-row weights."""
    raw_items = details.get("items")
    source = raw_items if isinstance(raw_items, list) else [details]
    items, issues = [], []
    for index, raw in enumerate(source, 1):
        if not isinstance(raw, dict):
            issues.append(f"material {index} is not a valid row")
            continue
        item = text(raw.get("item"))
        raw_quantity = text(raw.get("quantity"))
        raw_per_piece = text(raw.get("per_piece_weight"))
        raw_weight = text(raw.get("weight"))
        if not any((item, raw_quantity, raw_per_piece, raw_weight)):
            continue
        numeric_quantity = number(raw_quantity)
        # Text quantities such as "1 lot" are valid labels when the line's
        # total weight is entered directly. Only numeric quantities can be
        # multiplied by a per-piece weight.
        quantity = numeric_quantity if numeric_quantity is not None else raw_quantity or None
        per_piece_weight = number(raw_per_piece)
        weight = number(raw_weight)
        for label, raw_value, parsed in (("quantity", raw_quantity, numeric_quantity),
                                         ("per-piece weight", raw_per_piece, per_piece_weight),
                                         ("weight", raw_weight, weight)):
            if label != "quantity" and raw_value and parsed is None:
                issues.append(f"material {index} has an invalid {label}")
            elif parsed is not None and parsed < 0:
                issues.append(f"material {index} has a negative {label}")
        if not item:
            issues.append(f"material {index} is missing its description")
        if not raw_weight and numeric_quantity is not None and per_piece_weight is not None:
            weight = numeric_quantity * per_piece_weight
        if weight is None:
            issues.append(f"material {index} weight is missing or cannot be calculated")
        items.append({"item": item, "quantity": quantity, "per_piece_weight": per_piece_weight, "weight": weight})
    if not items:
        issues.append("no item rows")
    if any(item["weight"] is None for item in items):
        issues.append("one or more item weights cannot be calculated")
    total = sum(item["weight"] or 0 for item in items)
    if not math.isfinite(total) or total <= 0:
        issues.append("total weight must be a positive number")
    if any(item["weight"] is not None and (not math.isfinite(item["weight"]) or item["weight"] < 0) for item in items):
        issues.append("an item has an invalid or negative weight")
    return items, list(dict.fromkeys(issues))


def text(value):
    return str(value).strip() if value is not None else ""


ALIASES = {
    "sn": {"sn", "sno", "sr", "srno", "srnumber", "serialno", "serialnumber", "batchnumber", "inwardnumber", "inwardno", "inwardid"},
    "inward_date": {"inwarddate", "dateofmaterialreceipt", "materialreceiptdate", "receiveddate", "receiptdate"},
    "supplier": {"suppliername", "supplier", "clientname", "client", "companyname", "company", "customer", "customername"},
    "location": {"location", "city", "sitelocation", "siteaddress", "clientlocation", "clientcity"},
    "item": {"item", "itemdescription", "material", "description", "materialdescription", "assetdescription"},
    "quantity": {"quantityinpcs", "quantity", "qty"},
    "per_piece_weight": {"perpcsweightinkg", "perpieceweightinkg", "perpieceweight"},
    "weight": {"weightkg", "weightinkg", "weight"},
    # Optional batch-level weight requested by clients who want net weight
    # after the box/packing weight has been removed.
    "certificate_weight": {
        "certificateweightkg", "certificateweightinkg", "certificateweight",
        "certweightkg", "certweight", "netweightkg", "netweightinkg", "netweight",
        "netweightleavingboxkg", "netweightleavingbox", "weightforcertificate",
        "certificateweightnetweight", "certificateweightnetweightkg",
        "certificatenetweight", "certificatenetweightkg",
    },
    "form6": {"form6", "form6no", "form6number", "manifestno", "manifestnumber"},
    "reference": {"reference", "referenceno", "referencenumber", "refno", "refnumber", "certificatereferenceno"},
    "cod_no": {"codno", "codnumber", "certificateissuedno", "certificateno"},
    "cod_date": {"coddate", "certificateissueddate", "certificatedate"},
    "issued": {"certificateissued", "certificatestatus", "codstatus", "status"},
    "certificate_type": {"certificatetype", "codtype", "certificatetemplate"},
    "certificate_fields": {"certificatefields", "certificatedetails"},
    "comments": {"comments", "comment", "remarks", "remark"},
    "transporter_details": {"transportervehicledetails", "transporterdetails"},
    "docket_vehicle_no": {"docketnovehicleno", "docketnumbervehiclenumber"},
    "weight_slip_scale": {"weightslipweightscale", "weightslipscaleweight"},
}


def columns(headers):
    found = {}
    normalized = [norm(h) for h in headers]
    for key, choices in ALIASES.items():
        matches = [i for i, heading in enumerate(normalized) if heading in choices]
        if len(matches) > 1:
            labels = ", ".join(get_column_letter(i + 1) for i in matches)
            label = 'Serial Number' if key == 'sn' else key
            raise ValueError(f"Ambiguous duplicate headings for {label} in columns {labels}; keep one authoritative column.")
        if matches:
            found[key] = matches[0]
    return found


def is_mail_tracking_header(heading):
    """Mail timestamps/status are operational notes, never certificate values."""
    return re.fullmatch(
        r'(?:reference|refno|refnumber|site|location|city|manifest|form6|certificateweight|netweight)'
        r'(?:pog)?(?:copy|copies)?(?:e?mail)(?:sent|received)(?:date|time|on|status)?',
        norm(heading),
    ) is not None


def register_header(sheet):
    """Find one authoritative heading row within the first 20 rows."""
    candidates = []
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 20)):
        normalized = [norm(cell.value) for cell in row]
        if all(ALIASES[key].intersection(normalized) for key in ('sn', 'supplier', 'inward_date')):
            candidates.append(row[0].row)
    if not candidates:
        return None
    if len(candidates) != 1:
        raise ValueError(f'Worksheet "{sheet.title}" has more than one heading row. Keep one register header; remove repeated headers from the data.')
    row_number = candidates[0]
    return row_number, columns([cell.value for cell in sheet[row_number]])


def select_register(workbook):
    """Use an unambiguous register; an identical hidden copy is harmless."""
    candidates = []
    for sheet in workbook.worksheets:
        header = register_header(sheet)
        if header:
            candidates.append((sheet, *header))
    visible = [entry for entry in candidates if entry[0].sheet_state == 'visible']
    if len(visible) == 1:
        selected = visible[0]
        # An archived register can contain CODs that are absent from PDFs. Never
        # silently ignore those numbering sources. Only an exact hidden copy is safe.
        source_rows = list(selected[0].iter_rows(min_row=selected[1], values_only=True))
        for sheet, header_row, _ in candidates:
            if sheet is selected[0]:
                continue
            if list(sheet.iter_rows(min_row=header_row, values_only=True)) != source_rows:
                raise ValueError(f'Hidden register "{sheet.title}" contains different records. Move the archive to a separate workbook so its COD history can be checked; keep one current register here.')
        return selected
    if len(candidates) > 1:
        names = ', '.join(f'"{entry[0].title}"' for entry in candidates)
        raise ValueError(f'More than one inward register was found: {names}. Keep one current register in this workbook; archive the others separately. No rows were combined.')
    if candidates:
        raise ValueError(f'The inward register "{candidates[0][0].title}" is hidden. Unhide the current register in Excel, then refresh.')
    raise ValueError('No inward register found. Put Serial Number, Company / Supplier and Inward Date headings within the first 20 rows, then refresh.')


def source_value_issues(batch, as_of=None, certificate_type=None):
    """One validation policy for automatic Ready, preview and final generation.

    Form 6 is requested when a certificate is created, not when an inward is
    added. Missing values are checked by the certificate issuance workflow.
    """
    as_of = as_of or date.today()
    issues = list(batch.get('issues', []))
    received = batch.get('inward_date')
    if received and received > as_of:
        issues.append('Inward Date is in the future; correct it before generating the certificate')
    raw_cod_date = batch.get('cod_date')
    cod_date = parse_date(raw_cod_date)
    if text(raw_cod_date) and cod_date is None:
        issues.append('COD Date is not a valid date; correct it or clear it to use today')
    elif cod_date and cod_date > as_of:
        issues.append('COD Date is in the future; check this date before creating the certificate')
    if text(batch.get('cod_no')).startswith('='):
        issues.append('COD No contains a formula; enter a fixed COD or clear the cell')
    labels = {'cod_no':'COD No', 'cod_date':'COD Date', 'issued':'Certificate Issued'}
    for key in batch.get('missing_headers', []):
        issues.append(f'Add the {labels.get(key, key)} column to the plant workbook')
    return list(dict.fromkeys(issues))


def read_batches(path):
    # A save or refresh may inspect the same immutable workbook snapshot more
    # than once. Give each caller independent rows, since they annotate them.
    if isinstance(path, BytesIO) and path.getbuffer().nbytes <= MAX_SNAPSHOT_CACHE_BYTES:
        return deepcopy(_snapshot_batches(path.getvalue()))
    if isinstance(path, (str, os.PathLike)) and Path(path).suffix.casefold() in {'.xlsx', '.xlsm', '.xltx', '.xltm'}:
        with Path(path).open('rb') as source:
            data = source.read(MAX_SNAPSHOT_CACHE_BYTES + 1)
        if len(data) <= MAX_SNAPSHOT_CACHE_BYTES:
            return deepcopy(_snapshot_batches(data))
    return _read_batches(path)


@lru_cache(maxsize=8)
def _snapshot_batches(data):
    return _read_batches(BytesIO(data))


def _read_batches(path):
    from workbook_metadata import read_details, lookup_details
    saved_details = read_details(path)
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        sheet, header_row, header_columns = select_register(workbook)
        sheet_name = sheet.title
        rows = list(sheet.iter_rows(min_row=header_row, values_only=True))
        epoch = workbook.epoch
        mapping = {key: {'column': get_column_letter(index + 1), 'heading': text(rows[0][index])}
                   for key, index in header_columns.items()}
        header_issues = []
        for area in sheet.merged_cells.ranges:
            if area.min_row <= header_row <= area.max_row and area.max_col > area.min_col:
                unmapped = [col for col in range(area.min_col, area.max_col + 1)
                            if col - 1 not in header_columns.values()]
                if any(text(sheet.cell(row, col).value) for col in unmapped
                       for row in range(header_row + 1, sheet.max_row + 1)):
                    header_issues.append(f'Merged header {area} hides a populated column. Unmerge the headings and name each column before generating.')
        for index, heading in enumerate(rows[0]):
            if index in header_columns.values() or not text(heading) or is_mail_tracking_header(heading):
                continue
            if re.match(r'^(?:reference|refno|refnumber|site|location|city|manifest|form6|certificateweight|netweight)', norm(heading)):
                if any(text(row[index]) for row in rows[1:]):
                    header_issues.append(f'Column {get_column_letter(index + 1)} "{text(heading)}" has data but is not recognised. Rename its heading to Reference No, Location, Form 6 No or Certificate Weight as appropriate.')
        identity_formats = {
            key: {row_number: sheet.cell(row_number, header_columns[key] + 1).number_format
                  for row_number in range(header_row + 1, sheet.max_row + 1)}
            for key in ('sn', 'cod_no') if key in header_columns}
    finally:
        workbook.close()
    if not rows:
        raise ValueError("The workbook is empty.")
    cols = columns(rows[0])
    required = ["sn", "inward_date", "supplier", "item"]
    missing = [key for key in required if key not in cols]
    if "weight" not in cols and not {"quantity", "per_piece_weight"}.issubset(cols):
        missing.append("weight or quantity/per-piece weight")
    if missing:
        raise ValueError("Missing required workbook columns: " + ", ".join(missing))

    def read_cell_text(key, value, excel_row):
        if key in {'sn', 'cod_no'} and isinstance(value, (int, float)) and not isinstance(value, bool):
            number_format = identity_formats.get(key, {}).get(excel_row, "")
            if re.fullmatch(r"0+", number_format) and float(value).is_integer():
                return f"{int(value):0{len(number_format)}d}"
        return text(value)

    batches = []
    current = None
    summary_labels = {'total', 'subtotal', 'grandtotal', 'totalweight', 'totalweightkg', 'totalweightinkg'}
    for excel_row, raw in enumerate(rows[1:], header_row + 1):
        case_id = read_cell_text('sn', raw[cols['sn']], excel_row)
        item_label = text(raw[cols['item']])
        # Notes with no certificate/material data are not new company records.
        if re.match(r'^(?:note|notes|remark|remarks)\s*:', case_id, re.I) and not any(
                text(raw[index]) for key, index in cols.items() if key != 'sn'):
            continue
        summary_label = item_label if norm(item_label) in summary_labels else case_id if norm(case_id) in summary_labels else ''
        # A client can legitimately use Total as a Serial Number or material
        # description. Company/date facts distinguish a new inward from a
        # summary following the preceding inward's assets.
        starts_primary = bool(case_id) and any(
            text(raw[cols[key]]) for key in ('supplier', 'inward_date') if key in cols)
        if summary_label and not starts_primary:
            if current is None:
                raise ValueError(f'Row {excel_row} is a total without an inward. Place totals after their material rows.')
            raw_total = raw[cols['weight']] if 'weight' in cols else None
            summary_weight = weight_number(raw_total, number)
            expected_total = sum(material['weight'] or 0 for material in current['items'])
            has_inputs = any(text(raw[cols[key]]) for key in ('quantity', 'per_piece_weight') if key in cols)
            conflicting_facts = any(text(raw[cols[key]]) for key in ('supplier', 'inward_date', 'form6', 'reference', 'cod_no', 'cod_date', 'issued') if key in cols)
            if text(raw_total).startswith('=') and 'weight' in cols:
                weight_col = get_column_letter(cols['weight'] + 1)
                formula = re.sub(r'\s+|\$', '', text(raw_total)).upper()
                if formula == f'=SUM({weight_col}{current["source_row"]}:{weight_col}{excel_row - 1})' and not current['summary_rows']:
                    summary_weight = expected_total
            verified = (current['items'] and all(material['weight'] is not None for material in current['items'])
                        and not has_inputs and not conflicting_facts and summary_weight is not None
                        and math.isclose(summary_weight, expected_total, rel_tol=1e-9, abs_tol=1e-8))
            if verified:
                current['summary_rows'].append(excel_row)
                current['summary_records'].append({'row':excel_row, 'label':norm(summary_label), 'weight':summary_weight})
                current['import_notes'].append(f'Excel row {excel_row}: total checked against the materials; counted once.')
            else:
                issue = f'Excel row {excel_row}: total row needs checking. It does not clearly match the materials; correct or remove the summary row before generating.'
                current['issues'].append(issue)
                current['blocking_issues'].append(issue)
            continue
        if case_id:
            saved_issues = []
            try:
                saved = lookup_details(saved_details, sheet_name, case_id,
                                       legacy_case_ids=(text(raw[cols['sn']]),))
            except ValueError as exc:
                saved = {}
                saved_issues.append(str(exc))
            current = {
                "case_id": case_id,
                "source_row": excel_row,
                "inward_date": excel_date(raw[cols["inward_date"]], epoch, parse_date),
                "raw_inward_date": raw[cols["inward_date"]],
                "supplier": text(raw[cols["supplier"]]),
                "location": text(raw[cols["location"]]) if "location" in cols else "",
                "form6": text(raw[cols["form6"]]) if "form6" in cols else text(saved.get('form6')),
                "reference": text(raw[cols["reference"]]) if "reference" in cols else text(saved.get('reference')),
                "pdf_recovery_pending": bool(saved.get('_pdf_recovery_pending')),
                "external_pdf": saved.get('_external_pdf'),
                "hdd_cod_no": text(saved.get('_hdd_cod_no')),
                "hdd_cod_date": saved.get('_hdd_cod_date', ''),
                **{key: text(raw[cols[key]]) if key in cols else '' for key in
                   ('comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale')},
                "certificate_weight": None if 'certificate_weight' in cols else number(saved.get('certificate_weight')),
                "certificate_weight_invalid": False,
                "certificate_type": text(raw[cols["certificate_type"]]) if "certificate_type" in cols else text(saved.get('certificate_type')) or "standard",
                "certificate_fields": {} if 'certificate_fields' in cols else saved.get('certificate_fields', {}),
                "cod_no": read_cell_text("cod_no", raw[cols["cod_no"]], excel_row) if "cod_no" in cols else "",
                "cod_date": (excel_date(raw[cols['cod_date']], epoch, parse_date) or raw[cols['cod_date']]) if 'cod_date' in cols else None,
                "issued": ("N/A" if supplier_is_always_not_applicable(raw[cols["supplier"]])
                           else text(raw[cols["issued"]]) if "issued" in cols else ""),
                "source_sheet": sheet_name,
                "source_header_row": header_row,
                "column_mapping": mapping,
                "summary_rows": [],
                'summary_records': [],
                "import_notes": [],
                "blocking_issues": [*header_issues, *saved_issues],
                # Reference can be entered at certificate creation; only output
                # columns are reported here as missing headers.
                "missing_headers": [key for key in ("issued", "cod_no", "cod_date") if key not in cols],
                "items": [],
                "issues": [*header_issues, *saved_issues],
            }
            from certificate_templates import template_id
            try:
                current["certificate_type"] = template_id(current["certificate_type"])
            except ValueError:
                current["issues"].append("Unknown certificate type; choose a template in Details")
            if "certificate_fields" in cols and text(raw[cols["certificate_fields"]]):
                try:
                    fields = json.loads(text(raw[cols["certificate_fields"]]))
                    if not isinstance(fields, dict):
                        raise ValueError("expected an object")
                    current["certificate_fields"] = fields
                except (ValueError, TypeError):
                    current["issues"].append("Certificate Fields is not valid saved certificate data; open Details and correct it")
            batches.append(current)
        if current is None:
            if any(text(raw[index]) for index in cols.values()):
                raise ValueError(f"Row {excel_row} has data but no preceding Serial Number; enter its Serial Number or place it under the correct inward before issuing.")
            continue
        if not case_id:
            for key in ("inward_date", "supplier", "location", "form6", "reference", "cod_no", "cod_date", "issued"):
                if key not in cols or not text(raw[cols[key]]):
                    continue
                if current.get('hdd_cod_no') and excel_row == current['source_row'] + 1:
                    if key == 'cod_no' and read_cell_text(key, raw[cols[key]], excel_row) == current['hdd_cod_no']:
                        continue
                    if key == 'cod_date' and excel_date(raw[cols[key]], epoch, parse_date) == parse_date(current.get('hdd_cod_date')):
                        continue
                actual = raw[cols[key]]
                expected = current[key]
                if key in {"inward_date", "cod_date"}:
                    actual_date = excel_date(actual, epoch, parse_date)
                    same = actual_date is not None and actual_date == excel_date(expected, epoch, parse_date)
                else:
                    same = read_cell_text(key, actual, excel_row) == text(expected)
                if not same:
                    if key == "cod_no":
                        raise ValueError(f"COD No on continuation row {excel_row} differs from its inward; correct the missing Serial Number or misplaced COD before assigning new numbers.")
                    current["issues"].append(f"conflicting {key} on continuation row {excel_row}; check the missing Serial Number")
        for key in ("sn", "inward_date", "supplier", "location", "form6", "reference", "issued"):
            if key in cols and text(raw[cols[key]]).startswith("="):
                current["issues"].append(f"formula in {key} on row {excel_row}; enter a confirmed value")

        # Certificate/net weight is a batch-level value.  It is normally
        # entered on the first row, but accepting it on a continuation row
        # makes the import tolerant of the way users maintain their sheets.
        if "certificate_weight" in cols:
            raw_certificate_weight = raw[cols["certificate_weight"]]
            if text(raw_certificate_weight):
                parsed_certificate_weight = weight_number(raw_certificate_weight, number)
                if parsed_certificate_weight is None or parsed_certificate_weight <= 0:
                    current["certificate_weight_invalid"] = True
                    current["issues"].append(
                        f"invalid or negative certificate/net weight on row {excel_row}"
                    )
                elif current["certificate_weight"] is None:
                    current["certificate_weight"] = parsed_certificate_weight
                elif current["certificate_weight"] != parsed_certificate_weight:
                    current["issues"].append(
                        f"conflicting certificate/net weight on continuation row {excel_row}"
                    )
        item = text(raw[cols["item"]]) if "item" in cols else ""
        if not item:
            has_material_values = any(text(raw[cols[key]]) for key in ("quantity", "per_piece_weight", "weight") if key in cols)
            if not has_material_values:
                continue
            current["issues"].append(f"missing item description on row {excel_row}")
        raw_quantity = raw[cols["quantity"]] if "quantity" in cols else None
        quantity_is_numeric = number(raw_quantity) is not None
        # Text quantities such as "1 lot" are meaningful worksheet inputs,
        # not numeric amounts to multiply. Preserve them as written.
        qty = (number(raw_quantity) if quantity_is_numeric else text(raw_quantity) or None) if "quantity" in cols else None
        per = weight_number(raw[cols["per_piece_weight"]], number) if "per_piece_weight" in cols else None
        for key, parsed in (("quantity", qty), ("per_piece_weight", per)):
            if key in cols and text(raw[cols[key]]) and (parsed is None or isinstance(parsed, (int, float)) and parsed < 0):
                current["issues"].append(f"invalid or negative {key} on row {excel_row}")
        raw_weight = raw[cols["weight"]] if "weight" in cols else None
        weight = weight_number(raw_weight, number)
        may_calculate = not text(raw_weight)
        if text(raw_weight).startswith("="):
            references = {f"{get_column_letter(cols[key] + 1)}{excel_row}": parsed
                          for key, parsed in (('quantity', qty), ('per_piece_weight', per)) if key in cols}
            weight = formula_number(text(raw_weight), references)
            may_calculate = False
        if may_calculate:
            weight = qty * per if quantity_is_numeric and qty is not None and per is not None else None
        elif weight is None:
            current["issues"].append(f"invalid weight or unsupported weight formula on row {excel_row}; enter a confirmed weight in kg, or use same-row quantity/per-piece arithmetic, SUM or ROUND. External and other-row formulas are not read.")
        current["items"].append({"item": item, "quantity": qty, "per_piece_weight": per, "weight": weight})

    serial_rows = {}
    for batch in batches:
        serial_rows.setdefault(batch['case_id'].strip().casefold(), []).append(batch['source_row'])
    form6_counts = Counter(batch['form6'].strip().casefold() for batch in batches if batch['form6'])
    for batch in batches:
        if not batch["inward_date"]:
            batch["issues"].append("missing inward date")
        if not batch["supplier"]:
            batch["issues"].append("missing supplier name")
        if not batch["items"]:
            batch["issues"].append("no item rows")
        if any(item["weight"] is None for item in batch["items"]):
            batch["issues"].append("one or more item weights cannot be calculated")
        batch["total_weight"] = sum(item["weight"] or 0 for item in batch["items"])
        for summary in batch['summary_records']:
            if summary['label'] != 'subtotal' and not math.isclose(summary['weight'], batch['total_weight'], rel_tol=1e-9, abs_tol=1e-8):
                issue = f'Excel row {summary["row"]}: total row needs checking. More materials follow this total; move the total below all materials and check its value.'
                batch['issues'].append(issue)
                batch['blocking_issues'].append(issue)
        if not math.isfinite(batch["total_weight"]) or batch["total_weight"] <= 0:
            batch["issues"].append("total weight must be a positive number")
        if any(item["weight"] is not None and (not math.isfinite(item["weight"]) or item["weight"] < 0) for item in batch["items"]):
            batch["issues"].append("an item has an invalid or negative weight")
        if (batch["certificate_weight"] is not None
                and math.isfinite(batch["certificate_weight"])
                and batch["certificate_weight"] > 0):
            batch["display_weight"] = batch["certificate_weight"]
            batch["weight_source"] = "certificate"
        else:
            # A bad optional value remains a visible warning, but the
            # operator can still create a certificate using the inward total
            # after explicitly accepting that consequence.
            batch["display_weight"] = batch["total_weight"]
            batch["weight_source"] = "inward"
        duplicate_rows = serial_rows[batch['case_id'].strip().casefold()]
        if len(duplicate_rows) > 1:
            rows_label = ', '.join(str(row) for row in duplicate_rows)
            issue = (f'Serial Number "{batch["case_id"]}" is repeated in sheet "{sheet_name}", '
                     f'rows {rows_label}. Open Excel and give each inward a different Serial Number, '
                     'then save and refresh before creating a certificate.')
            batch['duplicate_serial_number'] = True
            batch['issues'].append(issue)
            batch['blocking_issues'].append(issue)
        if batch['form6'] and form6_counts[batch['form6'].strip().casefold()] > 1:
            issue = f'Form 6 number {batch["form6"]} is repeated on more than one inward. Correct the duplicate in Excel before creating a certificate.'
            batch['issues'].append(issue)
            batch['blocking_issues'].append(issue)
    return batches


def inward_details_complete(batch, as_of=None):
    """Approve an unselected status only when its Excel source facts are valid.

    Manual identifiers and the required reference can be entered on the preview
    page. COD and date are assigned on creation.
    """
    as_of = as_of or date.today()
    if source_value_issues(batch, as_of):
        return False
    received = batch.get("inward_date")
    if not received or received > as_of:
        return False
    return True


def status_value(value):
    cleaned = norm(value)
    if cleaned == "ready":
        return "ready"
    if cleaned in {"na", "notapplicable"}:
        return "na"
    if cleaned == "blank":
        return "blank"
    if cleaned in {"created", "drafted", "generated"}:
        return "generated"
    if cleaned in {"generatedmanual", "manualgenerated", "generatedexternally", "externalgenerated"}:
        return "generated_manual"
    if cleaned in {"completed", "sent"}:
        return "completed"
    if cleaned in {"alert", "checkdetails", "workingdraft"}:
        # Legacy warning/draft labels must never become eligible for creation.
        return "alert"
    return "blank" if not cleaned else "unknown"


def effective_certificate_status(batch, record=None, completion_ack=None):
    """Normalize the worksheet status, including legacy labels from old installs."""
    return status_value(batch.get("issued"))


def batch_content(batch):
    return {key: batch.get(key) for key in (
        "case_id", "inward_date", "supplier", "location", "form6", "reference",
        "certificate_weight", "items", "certificate_type", "certificate_fields",
    )}


def eligible(batch, as_of, days):
    inward_date = batch.get("inward_date")
    return bool(inward_date and (as_of - inward_date).days >= days)


def certificate_reminder_due(batch, as_of, days=25):
    """Remind from day 25 until the user marks the certificate Completed or N/A."""
    issued = batch.get("_effective_certificate_status", batch.get("issued"))
    return bool(
        batch.get("case_id")
        and eligible(batch, as_of, days)
        and status_value(issued) not in {"na", "completed"}
    )


def certificate_number(sequence, settings):
    return f"{settings.get('certificate_number_prefix', '')}{sequence:0{int(settings.get('certificate_number_width', 4))}d}"


def certificate_sequence(value, settings):
    value = text(value)
    match = re.search(r"(\d+)$", value)
    return int(match.group(1)) if match else 0


def fit(value, font, size, max_width, minimum_size=5, allow_overflow=False):
    value = str(value)
    if any(ord(character) < 32 for character in value):
        if not allow_overflow:
            raise ValueError("Certificate field contains a line break or control character; enter a single-line value.")
        value = re.sub(r"[\x00-\x1f]+", " ", value).strip()
    while size > minimum_size and stringWidth(value, font, size) > max_width:
        size = max(minimum_size, size - 0.5)
    if stringWidth(value, font, size) > max_width:
        if not allow_overflow:
            raise ValueError("Certificate field is too long to fit; shorten the display value before issuing.")
        # The operator has accepted the warning. Keep the PDF valid and
        # readable by clipping the display value to the available field.
        while value and stringWidth(value + "...", font, size) > max_width:
            value = value[:-1]
        value = value.rstrip() + ("..." if value else "")
    return value, size


def vendor_font():
    name = "CertificateVendor"
    if name in pdfmetrics.getRegisteredFontNames():
        return name
    path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / "calibrib.ttf"
    if path.is_file():
        pdfmetrics.registerFont(TTFont(name, str(path)))
        return name
    return "Helvetica-Bold"


def _template_cell_geometry(cell, template_size=(TEMPLATE_DESIGN_WIDTH, TEMPLATE_DESIGN_HEIGHT)):
    """Return the center and usable width of a source-template cell in PDF points."""
    template_width, template_height = template_size
    scale = min(PDF_WIDTH / template_width, PDF_HEIGHT / template_height)
    x_offset = (PDF_WIDTH - template_width * scale) / 2
    design_scale = template_width / TEMPLATE_DESIGN_WIDTH
    left, right = (edge * design_scale for edge in cell)
    return x_offset + ((left + right) / 2) * scale, (right - left) * scale


def supplier_lines(batch, max_width=620, allow_overflow=False):
    """Keep the vendor block centred and readable inside the template's gap."""
    supplier = " ".join(batch["supplier"].upper().split())
    # Workbook test data may prefix a supplier with its plant; that is not part
    # of the client name printed on the certificate.
    supplier = re.sub(r"^PLANT\s*[12]\s*(?:[-:]\s*)?", "", supplier).strip()
    # Print the complete location supplied by Excel or the PDF editor.
    location = " ".join(str(batch.get("location") or "").upper().split())
    full = f"{supplier}, {location}" if location else supplier
    font = vendor_font()
    for size in (28, 26, 24, 22):
        if stringWidth(full, font, size) <= max_width:
            return [full], size
    words = full.split()
    for size in (20, 19, 18, 17, 16):
        choices = []
        for split in range(1, len(words)):
            lines = [" ".join(words[:split]), " ".join(words[split:])]
            widths = [stringWidth(line, font, size) for line in lines]
            if max(widths) <= max_width:
                choices.append((abs(widths[0] - widths[1]), lines))
        if choices:
            return min(choices, key=lambda item: item[0])[1], size
    if allow_overflow:
        clipped, _ = fit(full, font, 16, max_width, minimum_size=5, allow_overflow=True)
        return [clipped], 16
    raise ValueError("Supplier name/location is too long for the certificate; shorten the display name before issuing.")


def certificate_layout(batch, certificate_no, certificate_date, allow_warnings=False,
                       template_size=(TEMPLATE_DESIGN_WIDTH, TEMPLATE_DESIGN_HEIGHT)):
    """Validate overlays before reserving a COD number or opening a PDF file."""
    inward_date = batch.get("inward_date") or certificate_date
    display_weight = batch.get("display_weight")
    if display_weight is None:
        display_weight = batch.get("total_weight")
    if display_weight is None or not math.isfinite(display_weight):
        display_weight = 0
    cod_center, cod_cell_width = _template_cell_geometry(CERTIFICATE_NUMBER_CELL, template_size)
    date_center, date_cell_width = _template_cell_geometry(CERTIFICATE_DATE_CELL, template_size)
    detail_center, detail_cell_width = _template_cell_geometry(DETAIL_VALUE_CELL, template_size)
    fields = [
        (certificate_no, cod_center, 54, 7.5, cod_cell_width - 2),
        (certificate_date.strftime("%d.%m.%Y"), date_center, 54, 7.0, date_cell_width - 10),
        (inward_date.strftime("%d.%m.%Y"), detail_center, 69, 8.5, detail_cell_width - 12),
        (f"{display_weight:.1f} KG", detail_center, 84, 8.5, detail_cell_width - 12),
        (batch.get("reference", ""), detail_center, 99, 7.5, detail_cell_width - 12),
        (batch.get("form6", ""), detail_center, 114, 7.5, detail_cell_width - 12),
    ]
    fitted = [(fit(raw, "Helvetica-Bold", size, width, allow_overflow=allow_warnings), x, y)
              for raw, x, y, size, width in fields]
    lines, supplier_size = supplier_lines(batch, allow_overflow=allow_warnings)
    return fitted, lines, supplier_size


def make_certificate(batch, certificate_no, certificate_date, template, destination, allow_warnings=False):
    if Path(template).suffix.casefold() == ".pdf":
        from certificate_templates import make_template_certificate
        return make_template_certificate(batch, certificate_no, certificate_date, destination, allow_warnings=allow_warnings)
    background = ImageReader(str(template))
    template_size = background.getSize()
    fitted, lines, supplier_size = certificate_layout(
        batch, certificate_no, certificate_date, allow_warnings=allow_warnings,
        template_size=template_size,
    )
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    W, H = PDF_WIDTH, PDF_HEIGHT
    pdf = canvas.Canvas(str(destination), pagesize=(W, H))
    background_scale = min(W / template_size[0], H / template_size[1])
    background_width = template_size[0] * background_scale
    background_height = template_size[1] * background_scale
    pdf.drawImage(background, (W - background_width) / 2, (H - background_height) / 2,
                  width=background_width, height=background_height, mask="auto")

    for (raw, size), x, y in fitted:
        pdf.setFont("Helvetica-Bold", size)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawCentredString(x, H - y, raw)
    pdf.setFont(vendor_font(), supplier_size)
    leading = supplier_size + 4
    first_baseline = 262 - (len(lines) - 1) * leading / 2
    for index, line in enumerate(lines):
        pdf.drawCentredString(W / 2, H - first_baseline - index * leading, line)
    pdf.showPage()
    pdf.save()
    return destination
