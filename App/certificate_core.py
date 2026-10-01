"""Source parsing, eligibility, and exact-layout certificate generation."""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from pathlib import Path
import re
import math
import os
import json

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

PDF_WIDTH, PDF_HEIGHT = 960, 540
TEMPLATE_DESIGN_WIDTH, TEMPLATE_DESIGN_HEIGHT = 1534, 867
# Vertical grid lines in the supplied certificate template, in source-image pixels.
CERTIFICATE_NUMBER_CELL = (797, 878)
CERTIFICATE_DATE_CELL = (878, 990)
DETAIL_VALUE_CELL = (797, 990)


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(text, fmt).date()
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


def text(value):
    return str(value).strip() if value is not None else ""


ALIASES = {
    "sn": {"sn", "serialnumber", "batchnumber", "inwardnumber"},
    "inward_date": {"inwarddate", "dateofmaterialreceipt", "materialreceiptdate"},
    "supplier": {"suppliername", "clientname", "companyname"},
    "location": {"location", "city"},
    "item": {"item", "itemdescription", "material"},
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
    "form6": {"form6no", "form6number"},
    "reference": {"referenceno", "referencenumber", "certificatereferenceno"},
    "cod_no": {"codno", "codnumber", "certificateissuedno", "certificateno"},
    "cod_date": {"coddate", "certificateissueddate", "certificatedate"},
    "issued": {"certificateissued", "certificatestatus", "codstatus"},
    "certificate_type": {"certificatetype", "codtype", "certificatetemplate"},
    "certificate_fields": {"certificatefields", "certificatedetails"},
}


def columns(headers):
    found = {}
    normalized = [norm(h) for h in headers]
    for key, choices in ALIASES.items():
        matches = [i for i, heading in enumerate(normalized) if heading in choices]
        if len(matches) > 1:
            labels = ", ".join(get_column_letter(i + 1) for i in matches)
            raise ValueError(f"Ambiguous duplicate headings for {key} in columns {labels}; keep one authoritative column.")
        if matches:
            found[key] = matches[0]
    return found


def read_batches(path):
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        # Unrelated sheets may use repeated headings; validate the register only.
        matches = [s for s in workbook.worksheets if all(
            ALIASES[key].intersection(norm(c.value) for c in s[1])
            for key in ("sn", "supplier", "inward_date"))]
        if len(matches) != 1:
            raise ValueError("Each plant workbook must contain one inward register with SN, Supplier Name and Inward Date headings.")
        sheet = matches[0]
        sheet_name = sheet.title
        rows = list(sheet.iter_rows(values_only=True))
        header_columns = columns(rows[0]) if rows else {}
        cod_formats = {}
        if "cod_no" in header_columns:
            cod_column = header_columns["cod_no"] + 1
            cod_formats = {row_number: sheet.cell(row_number, cod_column).number_format
                           for row_number in range(2, sheet.max_row + 1)}
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
        if key == "cod_no" and isinstance(value, (int, float)) and not isinstance(value, bool):
            number_format = cod_formats.get(excel_row, "")
            if re.fullmatch(r"0+", number_format) and float(value).is_integer():
                return f"{int(value):0{len(number_format)}d}"
        return text(value)

    batches = []
    current = None
    for excel_row, raw in enumerate(rows[1:], 2):
        case_id = text(raw[cols["sn"]])
        if case_id:
            current = {
                "case_id": case_id,
                "source_row": excel_row,
                "inward_date": parse_date(raw[cols["inward_date"]]),
                "supplier": text(raw[cols["supplier"]]),
                "location": text(raw[cols["location"]]) if "location" in cols else "",
                "form6": text(raw[cols["form6"]]) if "form6" in cols else "",
                "reference": text(raw[cols["reference"]]) if "reference" in cols else "",
                "certificate_weight": None,
                "certificate_weight_invalid": False,
                "certificate_type": text(raw[cols["certificate_type"]]) if "certificate_type" in cols else "standard",
                "certificate_fields": {},
                "cod_no": read_cell_text("cod_no", raw[cols["cod_no"]], excel_row) if "cod_no" in cols else "",
                "cod_date": raw[cols["cod_date"]] if "cod_date" in cols else None,
                "issued": text(raw[cols["issued"]]) if "issued" in cols else "",
                "source_sheet": sheet_name,
                # Reference No. is optional; only certificate output columns
                # are reported here as missing headers.
                "missing_headers": [key for key in ("issued", "cod_no", "cod_date") if key not in cols],
                "items": [],
                "issues": [],
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
                raise ValueError(f"Row {excel_row} has data but no preceding SN; correct this orphan row before issuing.")
            continue
        if not case_id:
            for key in ("inward_date", "supplier", "location", "form6", "reference", "cod_no", "cod_date", "issued"):
                if key not in cols or not text(raw[cols[key]]):
                    continue
                actual = raw[cols[key]]
                expected = current[key]
                if key in {"inward_date", "cod_date"}:
                    same = parse_date(actual) is not None and parse_date(actual) == parse_date(expected)
                else:
                    same = read_cell_text(key, actual, excel_row) == text(expected)
                if not same:
                    if key == "cod_no":
                        raise ValueError(f"COD No on continuation row {excel_row} differs from its batch; correct the missing SN or misplaced COD before assigning new numbers.")
                    current["issues"].append(f"conflicting {key} on continuation row {excel_row}; check the missing SN")
        for key in ("sn", "inward_date", "supplier", "location", "form6", "reference", "issued"):
            if key in cols and text(raw[cols[key]]).startswith("="):
                current["issues"].append(f"formula in {key} on row {excel_row}; enter a confirmed value")

        # Certificate/net weight is a batch-level value.  It is normally
        # entered on the first row, but accepting it on a continuation row
        # makes the import tolerant of the way users maintain their sheets.
        if "certificate_weight" in cols:
            raw_certificate_weight = raw[cols["certificate_weight"]]
            if text(raw_certificate_weight):
                parsed_certificate_weight = number(raw_certificate_weight)
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
            if any(text(raw[cols[key]]) for key in ("quantity", "per_piece_weight", "weight") if key in cols):
                current["issues"].append(f"missing item description on row {excel_row}")
            continue
        qty = number(raw[cols["quantity"]]) if "quantity" in cols else None
        per = number(raw[cols["per_piece_weight"]]) if "per_piece_weight" in cols else None
        for key, parsed in (("quantity", qty), ("per_piece_weight", per)):
            if key in cols and text(raw[cols[key]]) and (parsed is None or parsed < 0):
                current["issues"].append(f"invalid or negative {key} on row {excel_row}")
        raw_weight = raw[cols["weight"]] if "weight" in cols else None
        weight = number(raw_weight)
        may_calculate = not text(raw_weight)
        if text(raw_weight).startswith("=") and {"quantity", "per_piece_weight"}.issubset(cols):
            formula = re.sub(r"\s+|\$", "", text(raw_weight)).upper()
            qty_ref = f"{get_column_letter(cols['quantity'] + 1)}{excel_row}"
            per_ref = f"{get_column_letter(cols['per_piece_weight'] + 1)}{excel_row}"
            may_calculate = formula in {f"={qty_ref}*{per_ref}", f"={per_ref}*{qty_ref}"}
        if may_calculate:
            weight = qty * per if qty is not None and per is not None else None
        elif weight is None:
            current["issues"].append(f"invalid weight or unsupported weight formula on row {excel_row}; enter a confirmed weight or a same-row quantity times per-piece formula")
        current["items"].append({"item": item, "quantity": qty, "per_piece_weight": per, "weight": weight})

    case_id_counts = Counter(batch["case_id"] for batch in batches)
    form6_counts = Counter(batch["form6"] for batch in batches if batch["form6"])
    for batch in batches:
        if not batch["inward_date"]:
            batch["issues"].append("missing inward date")
        if not batch["supplier"]:
            batch["issues"].append("missing supplier name")
        if not batch["form6"]:
            batch["issues"].append("missing Form 6 number / manifest number")
        if not batch["items"]:
            batch["issues"].append("no item rows")
        if any(item["weight"] is None for item in batch["items"]):
            batch["issues"].append("one or more item weights cannot be calculated")
        batch["total_weight"] = sum(item["weight"] or 0 for item in batch["items"])
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
        if case_id_counts[batch["case_id"]] > 1:
            batch["issues"].append("SN is repeated on more than one batch in this plant workbook")
        if batch["form6"] and form6_counts[batch["form6"]] > 1:
            batch["issues"].append("Form 6 number is repeated on more than one batch in this plant workbook")
    return batches


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
    if cleaned in {"generatedmanual", "manualgenerated"}:
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
    # The certificate reference displays the city only, not state/country text.
    location = " ".join(batch.get("location", "").split(",", 1)[0].upper().split())
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
