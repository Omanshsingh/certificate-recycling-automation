"""Source parsing, eligibility, and exact-layout certificate generation."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from email.headerregistry import Address
from email.errors import HeaderParseError
import re
import math
import os

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


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


def valid_email(value):
    """One mailbox only; never pass address-list syntax or header text to Outlook."""
    if not value or len(value) > 254 or any(c.isspace() for c in value):
        return False
    try:
        parsed = Address(addr_spec=value)
    except (ValueError, HeaderParseError):
        return False
    return bool(parsed.username and parsed.domain and "." in parsed.domain and parsed.addr_spec == value)


ALIASES = {
    "sn": {"sn", "serialnumber", "batchnumber", "inwardnumber"},
    "inward_date": {"inwarddate", "dateofmaterialreceipt", "materialreceiptdate"},
    "supplier": {"suppliername", "clientname", "companyname"},
    "location": {"location", "city"},
    "item": {"item", "itemdescription", "material"},
    "quantity": {"quantityinpcs", "quantity", "qty"},
    "per_piece_weight": {"perpcsweightinkg", "perpieceweightinkg", "perpieceweight"},
    "weight": {"weightkg", "weightinkg", "weight"},
    "form6": {"form6no", "form6number"},
    "email": {"clientemail", "clientemailid", "email", "emailid", "customeremail"},
    "reference": {"referenceno", "referencenumber", "certificatereferenceno"},
    "cod_no": {"codno", "codnumber", "certificateissuedno", "certificateno"},
    "cod_date": {"coddate", "certificateissueddate", "certificatedate"},
    "issued": {"certificateissued", "certificatestatus", "codstatus"},
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
                "email": text(raw[cols["email"]]) if "email" in cols else "",
                "reference": text(raw[cols["reference"]]) if "reference" in cols else "",
                "cod_no": text(raw[cols["cod_no"]]) if "cod_no" in cols else "",
                "cod_date": raw[cols["cod_date"]] if "cod_date" in cols else None,
                "issued": text(raw[cols["issued"]]) if "issued" in cols else "",
                "source_sheet": sheet_name,
                "missing_headers": [key for key in ("issued", "cod_no", "cod_date", "reference") if key not in cols],
                "items": [],
                "issues": [],
            }
            batches.append(current)
        if current is None:
            if any(text(raw[index]) for index in cols.values()):
                raise ValueError(f"Row {excel_row} has data but no preceding SN; correct this orphan row before issuing.")
            continue
        if not case_id:
            for key in ("inward_date", "supplier", "location", "form6", "email", "reference", "cod_no", "cod_date", "issued"):
                if key not in cols or not text(raw[cols[key]]):
                    continue
                actual = raw[cols[key]]
                expected = current[key]
                if key in {"inward_date", "cod_date"}:
                    same = parse_date(actual) is not None and parse_date(actual) == parse_date(expected)
                else:
                    same = text(actual) == text(expected)
                if not same:
                    if key == "cod_no":
                        raise ValueError(f"COD No on continuation row {excel_row} differs from its batch; correct the missing SN or misplaced COD before assigning new numbers.")
                    current["issues"].append(f"conflicting {key} on continuation row {excel_row}; check the missing SN")
        for key in ("sn", "inward_date", "supplier", "location", "form6", "email", "reference", "issued"):
            if key in cols and text(raw[cols[key]]).startswith("="):
                current["issues"].append(f"formula in {key} on row {excel_row}; enter a confirmed value")
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

    for batch in batches:
        if not batch["inward_date"]:
            batch["issues"].append("missing inward date")
        if not batch["supplier"]:
            batch["issues"].append("missing supplier name")
        if not valid_email(batch["email"]):
            batch["issues"].append("missing or invalid client email")
        if not batch["reference"]:
            batch["issues"].append("Reference No. column missing: ask Chahna to add it" if "reference" not in cols else "missing reference number")
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
        if sum(b["case_id"] == batch["case_id"] for b in batches) > 1:
            batch["issues"].append("SN is repeated on more than one batch in this plant workbook")
    return batches


def status_value(value):
    cleaned = norm(value)
    if cleaned == "ready":
        return "ready"
    if cleaned in {"na", "notapplicable"}:
        return "na"
    return "blank" if not cleaned else "unknown"


def batch_content(batch):
    return {key: batch.get(key) for key in ("case_id", "inward_date", "supplier", "location", "form6", "email", "reference", "items")}


def eligible(batch, as_of, days):
    return bool(batch["inward_date"] and (as_of - batch["inward_date"]).days >= days)


def certificate_number(sequence, settings):
    return f"{settings.get('certificate_number_prefix', '')}{sequence:0{int(settings.get('certificate_number_width', 4))}d}"


def certificate_sequence(value, settings):
    value = text(value)
    prefix = settings.get("certificate_number_prefix", "")
    match = re.fullmatch(re.escape(prefix) + r"(\d+)", value)
    if not match:
        raise ValueError(f"COD No {value!r} does not match the configured number format. Check it before assigning new numbers.")
    return int(match.group(1))


def fit(value, font, size, max_width):
    value = str(value)
    if any(ord(character) < 32 for character in value):
        raise ValueError("Certificate field contains a line break or control character; enter a single-line value.")
    while size > 7 and stringWidth(value, font, size) > max_width:
        size = max(7, size - 0.5)
    if stringWidth(value, font, size) > max_width:
        raise ValueError("Certificate field is too long to fit; shorten the display value before issuing.")
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


def supplier_lines(batch, max_width=620):
    """Keep the vendor block centred and readable inside the template's gap."""
    supplier = " ".join(batch["supplier"].upper().split())
    location = " ".join(batch.get("location", "").upper().split())
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
    raise ValueError("Supplier name/location is too long for the certificate; shorten the display name before issuing.")


def certificate_layout(batch, certificate_no, certificate_date):
    """Validate overlays before reserving a COD number or opening a PDF file."""
    fields = [
        (certificate_no, 506, 54, 7.5, 40),
        (certificate_date.strftime("%d.%m.%Y"), 566, 54, 7.0, 48),
        (batch["inward_date"].strftime("%d.%m.%Y"), 506, 69, 8.5, 108),
        (f"{batch['total_weight']:.1f} KG", 506, 84, 8.5, 108),
        (batch["reference"], 506, 99, 7.5, 108),
        (batch["form6"], 506, 114, 7.5, 108),
    ]
    fitted = [(fit(raw, "Helvetica-Bold", size, width), x, y) for raw, x, y, size, width in fields]
    lines, supplier_size = supplier_lines(batch)
    return fitted, lines, supplier_size


def make_certificate(batch, certificate_no, certificate_date, template, destination):
    fitted, lines, supplier_size = certificate_layout(batch, certificate_no, certificate_date)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    W, H = 960, 540
    pdf = canvas.Canvas(str(destination), pagesize=(W, H))
    pdf.drawImage(ImageReader(str(template)), 0, 0, width=W, height=H, preserveAspectRatio=True, mask="auto")

    for (raw, size), x, y in fitted:
        pdf.setFont("Helvetica-Bold", size)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawString(x, H - y, raw)
    pdf.setFont(vendor_font(), supplier_size)
    leading = supplier_size + 4
    first_baseline = 262 - (len(lines) - 1) * leading / 2
    for index, line in enumerate(lines):
        pdf.drawCentredString(W / 2, H - first_baseline - index * leading, line)
    pdf.showPage()
    pdf.save()
    return destination
