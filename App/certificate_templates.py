"""Three explicit certificate designs, backed by the supplied PowerPoint PDFs.

Original text is reused as vector PDF content when unchanged. Only operator
edits and certificate values are typeset, so logos and artwork keep their
original proportions and defaults keep the original fonts and formatting.
"""
from __future__ import annotations

from copy import copy
from datetime import date
from io import BytesIO
import json
import os
from pathlib import Path
import re

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, DecodedStreamObject, NameObject, DictionaryObject
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


TEMPLATE_NAMES = {"standard": "Standard COD", "sita": "Sita International", "siemens": "Siemens Medical"}
TEMPLATE_ROOT = Path(__file__).resolve().parent / "templates" / "certificate-types"
SHAPE_KEYS = {5: ("certificate_title", "Certificate title"), 6: ("intro", "Introduction"),
              7: ("from_line", "From line"), 8: ("statement", "Recycling statement"),
              9: ("services", "Services line"), 10: ("registrations", "Registration details"),
              11: ("footer", "Office, facility and certification details"),
              14: ("company_display", "Company line on certificate")}
TABLE_KEYS = {
    "standard": {(1, 2): "cod_no", (1, 3): "cod_date", (2, 2): "inward_date", (3, 2): "certificate_weight",
                 (4, 2): "reference", (5, 2): "form6"},
    "sita": {(1, 2): "cod_date", (2, 2): "cod_no", (3, 2): "inward_date", (4, 2): "reference", (5, 2): "certificate_weight"},
    "siemens": {(1, 2): "cod_no", (1, 3): "cod_date", (2, 2): "model_name", (3, 2): "model_serial",
                (4, 2): "tube_serial", (5, 2): "certificate_weight", (6, 2): "site_location"},
}
DETAIL_LABELS = {"model_name": "Model name", "model_serial": "Model serial number",
                 "tube_serial": "X-ray tube serial number", "site_location": "Site location"}


def clean_text(value):
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def template_id(value):
    key = str(value or "standard").strip().casefold()
    # Accept full labels entered in Excel, without inferring from supplier names.
    aliases = {"standard cod": "standard", "sita international": "sita", "siemens medical": "siemens"}
    key = aliases.get(key, key)
    if key not in TEMPLATE_NAMES:
        raise ValueError("Choose a certificate type: Standard COD, Sita International, or Siemens Medical.")
    return key


def geometry(kind):
    kind = template_id(kind)
    data = json.loads((TEMPLATE_ROOT / f"{kind}.geometry.json").read_text(encoding="utf-8-sig"))
    if kind == "standard":
        # The certificate format ends at Manifest No.; keep the supplied
        # source intact but omit its extra disposal row from generated PDFs.
        data["cells"] = [cell for cell in data["cells"] if cell["row"] <= 5]
    return data


def regions(kind):
    data = geometry(kind)
    result = []
    for shape in data["shapes"]:
        if shape["id"] == 12:
            lines = clean_text(shape["text"]).split("\n")
            for key, label, value, top in (("signatory_heading", "Signatory heading", lines[0], shape["top"]),
                                           ("signatory_label", "Signatory label", lines[-1], shape["top"] + shape["height"] - 26)):
                result.append(dict(shape, key=key, label=label, text=value, top=top, height=26, group="wording"))
        elif shape["id"] in SHAPE_KEYS:
            key, label = SHAPE_KEYS[shape["id"]]
            field = dict(shape, key=key, label=label, text=clean_text(shape["text"]), group="wording")
            if template_id(kind) == "standard" and key == "company_display":
                field.update(font="Calibri", bold=-1)
            result.append(field)
    for cell in data["cells"]:
        position = cell["row"], cell["column"]
        if cell["column"] == 1:
            result.append(dict(cell, key=f"label_{cell['row']}", label=f"Table label: {clean_text(cell['text'])}",
                               text=clean_text(cell["text"]), group="wording", table=True))
        else:
            key = TABLE_KEYS[template_id(kind)].get(position)
            if key:
                field = dict(cell, key=key, label=DETAIL_LABELS.get(key, key), text="", group="details", table=True)
                if template_id(kind) == "siemens":
                    field["bold"] = -1
                result.append(field)
    return result


def template_specs():
    common = {"cod_no", "cod_date", "inward_date", "certificate_weight", "reference", "form6"}
    return [{"id": kind, "name": name,
             "default_company": next(field["text"].strip() for field in regions(kind) if field["key"] == "company_display"), "fields": [
        {"key": "weight_display", "label": "Weight as printed (optional)", "default": "", "group": "details", "multiline": False},
    ] + [
        {"key": field["key"], "label": field["label"], "default": field["text"], "group": field["group"],
         "multiline": field["key"] in {"statement", "footer", "registrations", "services"},
         "type": "text"}
        for field in regions(kind) if field["key"] not in common
        and (field["group"] == "details" or field["key"] == "company_display")
    ]} for kind, name in TEMPLATE_NAMES.items()]


def obsolete_template_warning(message):
    """Retire display warnings for controls that no longer exist, not data warnings."""
    message = str(message).casefold()
    if "disposal date" in message or "date of e-waste disposal" in message:
        return True
    if "extends beyond the page" in message or "may overlap " in message:
        return True
    fixed_labels = [label for key, label in SHAPE_KEYS.values() if key != "company_display"]
    fixed_labels += ["signatory heading", "signatory label", "table label:"]
    return "is long for its space;" in message and any(label.casefold() in message for label in fixed_labels)


def current_warnings(warnings):
    if not isinstance(warnings, list):
        return []
    return [warning for warning in warnings if not obsolete_template_warning(warning)]


def layout_defaults(kind):
    data = geometry(kind)
    company = next(field for field in regions(kind) if field["key"] == "company_display")
    return {"table": {"left": min(cell["left"] for cell in data["cells"]),
                      "top": min(cell["top"] for cell in data["cells"]), "scale": 1},
            "company_display": {"left": company["left"], "top": company["top"], "width": company["width"],
                                "height": company["height"], "font_size": company["size"] or 24}}


def validate_fields(kind, values):
    if values is None:
        return {}
    if not isinstance(values, dict):
        raise ValueError("Certificate fields must contain text values.")
    allowed = {field["key"] for field in regions(kind)
               if field["group"] == "details" or field["key"] == "company_display"} | {"weight_display"}
    result = {}
    for key, value in values.items():
        if key == "_annexure":
            if value and not re.fullmatch(r"[a-f0-9]{32}", str(value)):
                raise ValueError("Choose the annexure PDF again.")
            if value:
                result[key] = str(value)
            continue
        if key not in allowed:
            continue
        if not isinstance(value, (str, int, float)) or isinstance(value, bool):
            raise ValueError(f"Enter text for {key}.")
        value = clean_text(value)
        if len(value) > 6000 or any(ord(ch) < 32 and ch not in "\n\t" for ch in value):
            raise ValueError(f"The {key} text is too long or contains an unsupported character.")
        result[key] = value
    return result


def font_name(region):
    family = region.get("font", "Calibri")
    bold = region.get("bold") == -1
    font = f"COD-{family}-{int(bold)}"
    if font in pdfmetrics.getRegisteredFontNames():
        return font
    windows = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    candidates = [windows / ("calibrib.ttf" if bold else "calibri.ttf")]
    if "sitka" in family.casefold():
        candidates = [windows / ("SitkaB.ttc" if bold else "Sitka.ttc")]
    for path in candidates:
        if path.is_file():
            try:
                pdfmetrics.registerFont(TTFont(font, str(path)))
                return font
            except Exception:
                continue
    if "sitka" in family.casefold():
        # Current Windows provides Sitka as a variable font. Freeze its Small
        # optical size and requested weight before ReportLab embeds it. This
        # keeps replacement client names in the same serif face as the PPT.
        variable_path = windows / "SitkaVF.ttf"
        if variable_path.is_file():
            try:
                from fontTools.ttLib import TTFont as VariableFont
                from fontTools.varLib.instancer import instantiateVariableFont
                variable = VariableFont(variable_path)
                static = instantiateVariableFont(variable, {"opsz": 11, "wght": 700 if bold else 400}, inplace=True)
                output = BytesIO()
                static.save(output)
                static.close()
                output.seek(0)
                pdfmetrics.registerFont(TTFont(font, output))
                return font
            except Exception:
                pass
        return "Times-Bold" if bold else "Times-Roman"
    return "Helvetica-Bold" if bold else "Helvetica"


def display_values(batch, cod, issued_date):
    from certificate_core import parse_date
    values = validate_fields(batch.get("certificate_type"), batch.get("certificate_fields"))
    if template_id(batch.get("certificate_type")) == "siemens" and not values.get("site_location", "").strip():
        values["site_location"] = str(batch.get("location") or "").strip()
    supplier = str(batch.get("supplier") or "").strip()
    supplier = re.sub(r"^plant[\s_-]*[12](?=$|[\s_-])[\s_-]*", "", supplier, flags=re.I)
    city = str(batch.get("location") or "").strip()
    company = f"{supplier}, {city}" if city else supplier
    # Empty company override means use the freely entered company/city fields.
    if not values.get("company_display", "").strip():
        values["company_display"] = company.upper()
    weight = batch.get("display_weight", batch.get("certificate_weight", batch.get("total_weight")))
    weight_text = f"{float(weight):g} KG" if weight is not None else ""
    weight_text = values.get("weight_display", "").strip() or weight_text
    receipt = parse_date(batch.get("inward_date"))
    values.update(cod_no=str(cod), cod_date=issued_date.strftime("%d.%m.%Y"),
                  inward_date=receipt.strftime("%d.%m.%Y") if receipt else "",
                  certificate_weight=weight_text, reference=str(batch.get("reference") or ""),
                  form6=str(batch.get("form6") or ""))
    return values


def positioned_regions(kind, values):
    kind = template_id(kind)
    layout = values.get("_layout", {})
    defaults = layout_defaults(kind)
    table = dict(defaults["table"], **layout.get("table", {}))
    table_changed = table != defaults["table"]
    for source in regions(kind):
        field = dict(source)
        if field["key"] == "company_display":
            # The source placeholder is offset and, for Standard, intersects
            # the processing paragraph. Centre the actual name on the page
            # in the clear space between the introduction and statement.
            top, height, margin = {"standard": (212, 44, 2),
                                   "sita": (202, 44, 2),
                                   "siemens": (233.5, 38, 1)}[kind]
            field.update(left=90, width=780, top=top, height=height,
                         align=2, margins=[0, margin, 0, margin])
            if kind == "standard":
                field["size"] = 26
        if field.get("table") and table_changed:
            scale = table["scale"]
            field["left"] = table["left"] + (field["left"] - defaults["table"]["left"]) * scale
            field["top"] = table["top"] + (field["top"] - defaults["table"]["top"]) * scale
            field["width"] *= scale
            field["height"] *= scale
            field["size"] = (field["size"] or 11) * scale
            field["margins"] = [margin * scale for margin in field["margins"]]
        for key, raw in layout.get(field["key"], {}).items():
            if key in {"left", "top", "width", "height"}:
                field[key] = raw
            elif key == "font_size":
                field["size"] = raw
        field["moved"] = field != source
        yield field


def _wrap(text, font, size, width):
    lines = []
    for paragraph in text.replace("\t", "    ").split("\n"):
        line = ""
        for word in paragraph.split():
            candidate = f"{line} {word}".strip()
            if line and pdfmetrics.stringWidth(candidate, font, size) > width:
                lines.append(line)
                line = word
            else:
                line = candidate
        lines.append(line)
    return lines or [""]


def _layout(region, value):
    font = font_name(region)
    margins = region.get("margins", [2, 2, 2, 2])
    width = region["width"] - margins[0] - margins[2]
    height = region["height"] - margins[1] - margins[3]
    preferred = region.get("size") or 11
    if preferred <= 0:
        preferred = 12
    size = float(preferred)
    company = region["key"] == "company_display"
    minimum = 16 if company else 6
    while size >= minimum:
        lines = _wrap(value, font, size, width)
        leading = size * 1.2
        ascent, descent = pdfmetrics.getAscentDescent(font, size)
        if (not company or len(lines) <= 2) and (len(lines)-1)*leading + ascent-descent <= height and all(pdfmetrics.stringWidth(line, font, size) <= width for line in lines):
            return font, size, lines, leading, width, height, size < 8
        size -= .5
    return font, minimum, _wrap(value, font, minimum, width), minimum * 1.2, width, height, True


def layout_warnings(batch, cod, issued_date):
    values = display_values(batch, cod, issued_date)
    warnings = []
    placed = list(positioned_regions(batch.get("certificate_type"), values))
    for region in placed:
        value = values.get(region["key"], region["text"])
        if value and value != region["text"] and _layout(region, value)[-1]:
            warnings.append(f"{region['label']} is long for its space; the PDF may use small text or clip it. Shorten it and preview before sending.")
        if (region["left"] < 0 or region["top"] < 0 or region["left"] + region["width"] > 960
                or region["top"] + region["height"] > 540):
            warnings.append(f"{region['label']} extends beyond the page; move it inside and preview before sending.")
    table_cells = [field for field in placed if field.get("table")]
    for field in placed:
        if field.get("table") or not field.get("moved"):
            continue
        for other in placed:
            if other is field or other.get("table") or not values.get(other["key"], other["text"]):
                continue
            if (max(field["left"], other["left"]) < min(field["left"]+field["width"], other["left"]+other["width"])
                    and max(field["top"], other["top"]) < min(field["top"]+field["height"], other["top"]+other["height"])):
                warnings.append(f"{field['label']} may overlap {other['label']}; check the preview.")
                break
    if table_cells and values.get("_layout", {}).get("table"):
        bounds = (min(c["left"] for c in table_cells), min(c["top"] for c in table_cells),
                  max(c["left"]+c["width"] for c in table_cells), max(c["top"]+c["height"] for c in table_cells))
        for field in placed:
            if field.get("table") or not values.get(field["key"], field["text"]):
                continue
            if (max(bounds[0], field["left"]) < min(bounds[2], field["left"]+field["width"])
                    and max(bounds[1], field["top"]) < min(bounds[3], field["top"]+field["height"])):
                warnings.append(f"The information box may overlap {field['label']}; check the preview.")
                break
    return list(dict.fromkeys(warnings))


def _original_region(source, region, height):
    layer = copy(source)
    left, top, width, box_height = (float(region[k]) for k in ("left", "top", "width", "height"))
    # PowerPoint marks each text run as a content group. Keep only the runs
    # in this field, rather than merely hiding all the other original text
    # behind a clip (which would leave old client names in PDF text searches).
    groups, group, depth = [], [], 0
    for operands, operator in source.get_contents().operations:
        if operator in {b"BDC", b"BMC"}:
            if depth == 0:
                group = []
            depth += 1
        if depth:
            group.append((operands, operator))
        if operator == b"EMC" and depth:
            depth -= 1
            if depth == 0:
                groups.append(group)
    kept = []
    for group in groups:
        positions = [args for args, op in group if op == b"Tm"]
        matches = any(left-.2 <= float(pos[4]) <= left+width+.2
                      and height-top-box_height-.2 <= float(pos[5]) <= height-top+.2 for pos in positions)
        if not positions and region["key"] == "certificate_title":
            # The native title shadow is a small image tagged alongside its
            # text runs. Logos and watermark groups carry no language tag.
            lang = any(op == b"BDC" and len(args) > 1 and isinstance(args[1], dict)
                       and args[1].get("/Lang") for args, op in group)
            matches = lang and any(op == b"cm" and left <= float(args[4]) <= left+width
                                  and height-top-box_height-30 <= float(args[5]) <= height-top
                                  for args, op in group)
        if matches:
            # Neighbouring text boxes can have intersecting bounds in the
            # supplied PPT. Exclude another field's text even if its baseline
            # falls inside this box (for example the signatory heading).
            if positions:
                probe = copy(source)
                probe_stream = ContentStream(None, source.pdf)
                probe_stream.operations = group
                probe[NameObject("/Contents")] = probe_stream
                fragment = re.sub(r"[^a-z0-9]", "", clean_text(probe.extract_text() or "").casefold())
                expected = re.sub(r"[^a-z0-9]", "", clean_text(region["text"]).casefold())
                if fragment and fragment[:40] not in expected:
                    continue
            kept.extend(group)
    content = ContentStream(None, source.pdf)
    content.operations = kept
    # Import only resources referenced by these runs. Unused transparency
    # masks from the complete PowerPoint page must not replace artwork masks.
    resources = DictionaryObject()
    source_resources = source['/Resources']
    for category, operator in (('/XObject', b'Do'), ('/Font', b'Tf'), ('/ExtGState', b'gs')):
        names = {args[0] for args, op in kept if op == operator and args}
        if names and category in source_resources:
            resources[NameObject(category)] = DictionaryObject({name: source_resources[category].raw_get(name) for name in names})
    for category in ('/ColorSpace', '/Pattern', '/Shading', '/Properties', '/ProcSet'):
        if category in source_resources:
            resources[NameObject(category)] = source_resources[category]
    layer[NameObject('/Resources')] = resources
    contents = content.get_data()
    stream = DecodedStreamObject()
    stream.set_data(f"q\n{left} {height-top-box_height} {width} {box_height} re W n\n".encode() + contents + b"\nQ\n")
    layer[NameObject("/Contents")] = stream
    return layer


def make_template_certificate(batch, cod, issued_date, destination, allow_warnings=False):
    kind = template_id(batch.get("certificate_type"))
    data = geometry(kind)
    values = display_values(batch, cod, issued_date)
    warnings = layout_warnings(batch, cod, issued_date)
    if warnings and not allow_warnings:
        raise ValueError("; ".join(warnings))
    original = PdfReader(TEMPLATE_ROOT / f"{kind}.pdf").pages[0]
    writer = PdfWriter()
    writer.add_page(PdfReader(TEMPLATE_ROOT / f"{kind}.background.pdf").pages[0])
    background = writer.pages[0]
    overlay = BytesIO()
    pdf = canvas.Canvas(overlay, pagesize=(data["width"], data["height"]))
    placed = list(positioned_regions(kind, values))
    # Table borders are structural vectors. Drawing the merged cell geometry
    # lets operators move/resize the complete box while keeping the artwork.
    pdf.setLineWidth(.75)
    pdf.setStrokeColorRGB(0, 0, 0)
    edges = set()
    for field in placed:
        if not field.get("table"):
            continue
        x, y, width, height = (field[k] for k in ("left", "top", "width", "height"))
        for edge in ((x,y,x+width,y), (x,y+height,x+width,y+height), (x,y,x,y+height), (x+width,y,x+width,y+height)):
            edges.add(tuple(round(point, 4) for point in edge))
    for x1, y1, x2, y2 in sorted(edges):
        pdf.line(x1, data["height"]-y1, x2, data["height"]-y2)
    for region in placed:
        value = values.get(region["key"], region["text"])
        if value == region["text"] and value and not region.get("moved"):
            background.merge_page(_original_region(original, region, data["height"]))
            continue
        if not value:
            continue
        font, size, lines, leading, width, height, _ = _layout(region, value)
        ml, mt, mr, mb = region.get("margins", [2, 2, 2, 2])
        x = region["left"] + ml
        top = region["top"] + mt
        ascent, descent = pdfmetrics.getAscentDescent(font, size)
        # Fit uses actual glyph height plus baseline spacing, rather than
        # rounding height/leading down and dropping a fitted second line.
        available_lines = max(1, int((height - (ascent - descent)) / leading + 1e-7) + 1)
        if region["key"] == "company_display":
            available_lines = min(2, available_lines)
        if len(lines) > available_lines:
            lines = lines[:available_lines]
            lines[-1] = lines[-1].rstrip() + "..."
        text_height = (len(lines) - 1) * leading + ascent - descent
        baseline = data["height"] - top - (height-text_height)/2 - ascent
        color = int(region.get("color") or 0)
        if color < 0:
            color = 0
        pdf.saveState()
        clip = pdf.beginPath()
        clip.rect(x, data["height"] - top - height, width, height)
        pdf.clipPath(clip, stroke=0)
        pdf.setFont(font, size)
        pdf.setFillColorRGB((color & 255)/255, ((color >> 8) & 255)/255, ((color >> 16) & 255)/255)
        align = region.get("align", 2)
        for index, line in enumerate(lines):
            y = baseline - index * leading
            if region.get("table") and region.get("group") == "details":
                pdf.drawCentredString(x + width/2, y, line)
            elif align == 1:
                pdf.drawString(x, y, line)
            elif align == 3:
                pdf.drawRightString(x + width, y, line)
            else:
                pdf.drawCentredString(x + width/2, y, line)
        pdf.restoreState()
    pdf.save()
    background.merge_page(PdfReader(overlay).pages[0])
    writer.add_metadata({"/Title": f"{TEMPLATE_NAMES[kind]} - {cod}", "/Subject": kind})
    annexure = values.get("_annexure")
    if annexure:
        from app_paths import DATA_ROOT
        annexure_path = Path(batch.get("_data_root", DATA_ROOT)) / "Data" / "Attachments" / f"{annexure}.pdf"
        if not annexure_path.is_file():
            raise ValueError("The attached annexure is missing. Choose its PDF again before creating the certificate.")
        attached = PdfReader(annexure_path)
        if attached.is_encrypted:
            raise ValueError("The annexure PDF is password protected. Choose an unlocked PDF.")
        for page in attached.pages:
            writer.add_page(page)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as stream:
        writer.write(stream)
    return destination
