"""Source parsing, eligibility, and exact-layout certificate generation."""
from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import re

from openpyxl import load_workbook
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
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
    if value is None or value == "":
        return None
    if isinstance(value, str) and value.strip().startswith("="):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else None


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
    "form6": {"form6no", "form6number"},
    "email": {"clientemail", "clientemailid", "email", "emailid", "customeremail"},
    "reference": {"referenceno", "referencenumber", "certificatereferenceno"},
}


def columns(headers):
    found = {}
    normalized = {norm(h): i for i, h in enumerate(headers)}
    for key, choices in ALIASES.items():
        for choice in choices:
            if choice in normalized:
                found[key] = normalized[choice]
                break
    return found


def read_batches(path):
    workbook = load_workbook(path, data_only=False, read_only=False)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("The workbook is empty.")
    cols = columns(rows[0])
    required = ["sn", "inward_date", "supplier", "item", "weight", "form6", "email", "reference"]
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
                "form6": text(raw[cols["form6"]]),
                "email": text(raw[cols["email"]]),
                "reference": text(raw[cols["reference"]]),
                "items": [],
                "issues": [],
            }
            batches.append(current)
        if current is None:
            continue
        item = text(raw[cols["item"]]) if "item" in cols else ""
        if not item:
            continue
        qty = number(raw[cols["quantity"]]) if "quantity" in cols else None
        per = number(raw[cols["per_piece_weight"]]) if "per_piece_weight" in cols else None
        explicit = number(raw[cols["weight"]]) if "weight" in cols else None
        weight = explicit if explicit is not None else (qty * per if qty is not None and per is not None else None)
        current["items"].append({"item": item, "quantity": qty, "per_piece_weight": per, "weight": weight})

    for batch in batches:
        if not batch["inward_date"]:
            batch["issues"].append("missing inward date")
        if not batch["supplier"]:
            batch["issues"].append("missing supplier name")
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", batch["email"]):
            batch["issues"].append("missing or invalid client email")
        if not batch["reference"]:
            batch["issues"].append("missing reference number")
        if not batch["form6"]:
            batch["issues"].append("missing Form 6 number / manifest number")
        if not batch["items"]:
            batch["issues"].append("no item rows")
        if any(item["weight"] is None for item in batch["items"]):
            batch["issues"].append("one or more item weights cannot be calculated")
        batch["total_weight"] = sum(item["weight"] or 0 for item in batch["items"])
    return batches


def eligible(batch, as_of, days):
    return bool(batch["inward_date"] and as_of >= batch["inward_date"] + timedelta(days=days))


def certificate_number(sequence, settings):
    return f"{settings.get('certificate_number_prefix', '')}{sequence:0{int(settings.get('certificate_number_width', 4))}d}"


def fit(value, font, size, max_width):
    value = str(value)
    while size > 7 and stringWidth(value, font, size) > max_width:
        size -= 0.5
    return value, size


def make_certificate(batch, certificate_no, certificate_date, template, destination):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    W, H = 960, 540
    pdf = canvas.Canvas(str(destination), pagesize=(W, H))
    pdf.drawImage(ImageReader(str(template)), 0, 0, width=W, height=H, preserveAspectRatio=True, mask="auto")

    def value(raw, x, y, size=8.5, max_width=180):
        raw, size = fit(raw, "Helvetica-Bold", size, max_width)
        pdf.setFont("Helvetica-Bold", size)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawString(x, H - y, raw)

    value(certificate_no, 506, 54, 7.5, 62)
    value(certificate_date.strftime("%d.%m.%Y"), 566, 54, 7.0, 56)
    value(batch["inward_date"].strftime("%d.%m.%Y"), 506, 69, 8.5, 180)
    value(f"{batch['total_weight']:.1f} KG", 506, 84, 8.5, 180)
    value(batch["reference"], 506, 99, 7.5, 180)
    value(batch["form6"], 506, 114, 7.5, 180)
    supplier = f"{batch['supplier']}, {batch['location']}" if batch["location"] else batch["supplier"]
    supplier, supplier_size = fit(supplier.upper(), "Helvetica-Bold", 15, 500)
    pdf.setFont("Helvetica-Bold", supplier_size)
    pdf.drawCentredString(W / 2, H - 251, supplier)
    pdf.showPage()
    pdf.save()
    return destination
