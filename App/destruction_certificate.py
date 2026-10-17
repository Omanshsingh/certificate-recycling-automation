"""Additional data destruction certificates using the supplied two-page artwork."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


SPEC = {'id': 'destruction', 'name': 'Secured Data Destruction', 'default_company': '', 'fields': [
    {'key': 'destruction_date', 'label': 'Date devices were destroyed', 'type': 'date',
     'default': '', 'group': 'details', 'multiline': False},
    {'key': 'device_receipts', 'label': 'Device receipts', 'type': 'receipts',
     'default': '', 'group': 'details', 'multiline': True},
]}


def receipts(value):
    from certificate_core import parse_date
    if not value:
        return []
    try:
        entries = json.loads(value)
    except (TypeError, ValueError) as exc:
        raise ValueError('Enter a received date and whole device count for each receipt.') from exc
    if not isinstance(entries, list) or len(entries) > 100:
        raise ValueError('Enter no more than 100 device receipts.')
    result = []
    for index, entry in enumerate(entries, 1):
        if not isinstance(entry, dict):
            raise ValueError(f'Device receipt {index} is invalid.')
        raw_date, raw_count = str(entry.get('date') or '').strip(), str(entry.get('count') or '').strip()
        if not raw_date and not raw_count:
            continue
        received = parse_date(raw_date)
        if not received or not raw_count.isascii() or not raw_count.isdigit() or not 0 < int(raw_count) <= 1000000000:
            raise ValueError(f'Device receipt {index}: enter a valid received date and a positive whole device count.')
        result.append((received, int(raw_count)))
    return result


def content_issues(details, as_of=None):
    """Required facts cannot be replaced by a warning override on final issuance."""
    from certificate_core import parse_date
    as_of = as_of or date.today()
    fields = details.get('certificate_fields') or {}
    issues = []
    if not str(details.get('supplier') or '').strip():
        issues.append('Enter the company name for the data destruction certificate.')
    destroyed = parse_date(fields.get('destruction_date'))
    if not destroyed:
        issues.append('Enter the actual date the devices were destroyed.')
    elif destroyed > as_of:
        issues.append('The destruction date cannot be in the future.')
    try:
        rows = receipts(fields.get('device_receipts'))
    except ValueError as exc:
        issues.append(str(exc))
        rows = []
    if not rows:
        issues.append('Add at least one device receipt with a received date and device count.')
    if any(received > as_of for received, _ in rows):
        issues.append('Device received dates cannot be in the future.')
    if destroyed and any(received > destroyed for received, _ in rows):
        issues.append('The destruction date cannot be before a device received date.')
    return issues


def layout_warnings(batch):
    warnings = []
    for key, width in [('supplier', 590), ('location', 153)]:
        try:
            fitted_text(str(batch.get(key) or ''), width, minimum=8)
        except ValueError:
            warnings.append(f'Shorten the {"company name" if key == "supplier" else "city/location"} to fit the data destruction certificate.')
    return warnings


def fitted_text(value, width, minimum=8, preferred=13):
    from certificate_templates import font_name
    from reportlab.pdfbase import pdfmetrics
    font = font_name({'font': 'Calibri'})
    value = ' '.join(str(value).split())
    size = preferred
    while size > minimum and pdfmetrics.stringWidth(value, font, size) > width:
        size -= .5
    if pdfmetrics.stringWidth(value, font, size) > width:
        raise ValueError('Text is too long for the supplied data destruction design.')
    return font, size, value


def make_certificate(batch, cod, issued_date, destination):
    from certificate_core import parse_date
    from certificate_templates import TEMPLATE_ROOT
    template = PdfReader(TEMPLATE_ROOT / 'destruction.pdf')
    fields = batch.get('certificate_fields') or {}
    # An incomplete preview shows blanks and its validation messages, never invented dates/counts.
    try:
        rows = receipts(fields.get('device_receipts'))
    except ValueError:
        rows = []
    destroyed = parse_date(fields.get('destruction_date'))
    writer = PdfWriter()

    def page(index, draw):
        writer.add_page(template.pages[index])
        overlay = BytesIO()
        pdf = canvas.Canvas(overlay, pagesize=(960, 540))
        draw(pdf)
        pdf.save()
        writer.pages[-1].merge_page(PdfReader(overlay).pages[0])

    def text(pdf, value, left, baseline, width, preferred=13):
        font, size, value = fitted_text(value, width, preferred=preferred)
        pdf.setFont(font, size)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawCentredString(left + width / 2, 540 - baseline, value)

    def front(pdf):
        text(pdf, batch.get('supplier', '').upper(), 184, 264, 590)
        text(pdf, batch.get('location', ''), 204, 298, 153, 11)
        text(pdf, destroyed.strftime('%d.%m.%Y') if destroyed else '', 242, 332, 127, 11)
        text(pdf, cod, 487, 332, 90, 11)
        text(pdf, issued_date.strftime('%d.%m.%Y'), 619, 332, 115, 11)

    page(0, front)
    groups = [rows[index:index + 9] for index in range(0, len(rows), 9)] or [[]]
    for index, group in enumerate(groups):
        def details(pdf):
            # A searchable identifier ties every details page to its certificate.
            text(pdf, f'Certificate No. {cod}   |   Details {index + 1} of {len(groups)}', 310, 25, 340, 9)
            for row_index, (received, count) in enumerate(group):
                baseline = 99 + row_index * 20.65
                text(pdf, received.strftime('%d.%m.%Y'), 314, baseline, 165, 11)
                text(pdf, str(count), 482, baseline, 165, 11)
            text(pdf, str(sum(count for _, count in group)) if group else '', 482, 283, 165, 11)
            if len(groups) > 1:
                text(pdf, f'Page subtotal above. Total devices: {sum(count for _, count in rows)}', 310, 307, 340, 10)
        page(1, details)
    annexure = fields.get('_annexure')
    if annexure:
        from app_paths import DATA_ROOT
        attachment = Path(batch.get('_data_root', DATA_ROOT)) / 'Data' / 'Attachments' / f'{annexure}.pdf'
        if not attachment.is_file():
            raise ValueError('The attached annexure is missing. Choose its PDF again.')
        attached = PdfReader(attachment)
        if attached.is_encrypted:
            raise ValueError('Choose an unlocked annexure PDF.')
        for attached_page in attached.pages:
            writer.add_page(attached_page)
    writer.add_metadata({'/Title': f'Secured Data Destruction - {cod}', '/Subject': 'destruction'})
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open('wb') as stream:
        writer.write(stream)
    return destination


def request_for_inward(root, plant, case_id, expected_revision=None):
    """Save an idempotent additional request without writing to the inward workbook."""
    from certificate_core import read_batches
    from certificate_drafts import read_register, save_draft
    from workflow_support import load_settings
    from workspace_review import require_revision
    if expected_revision is not None:
        require_revision(root, expected_revision)
    if plant not in {'plant1', 'plant2'} or not str(case_id or '').strip():
        raise ValueError('Choose an inward client first.')
    case_id = str(case_id).strip()
    paths = load_settings(root)['workbook_paths']
    path = paths[int(plant[-1]) - 1]
    if not path or not Path(path).is_file():
        raise ValueError('Connect this plant workbook and refresh before adding its data destruction request.')
    matches = [batch for batch in read_batches(Path(path))
               if str(batch['case_id']).strip().casefold() == case_id.casefold()]
    if len(matches) != 1:
        raise ValueError('The inward Serial Number is missing or repeated. Correct Excel and refresh.')
    source = matches[0]
    link = {'plant': plant, 'case_id': str(source['case_id'])}
    for record in read_register(root)['records']:
        if record['details'].get('destruction_source') == link:
            return record['id']
    from certificate_core import parse_date
    received = parse_date(source.get('inward_date'))
    details = dict(source, certificate_type='destruction', certificate_fields={
        'device_receipts': json.dumps([{'date': received.isoformat() if received else '', 'count': ''}])},
                   destruction_source=link, cod_no='', cod_date='', items=[])
    details.pop('certificate_weight', None)
    return save_draft(root, plant, details, expected_revision=expected_revision)
