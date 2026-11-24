"""Additional data destruction certificates using the supplied two-page artwork."""
from datetime import date
from io import BytesIO
import json
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


SPEC = {'id': 'destruction', 'name': 'Data Destruction', 'default_company': '', 'fields': [
    {'key': 'destruction_date', 'label': 'Date devices were destroyed', 'type': 'date',
     'default': '', 'group': 'details', 'multiline': False,
     'help': 'The actual destruction date printed on page 1. Receipt dates are entered separately below.'},
    {'key': 'device_receipts', 'label': 'Device receipts', 'type': 'receipts',
     'default': '', 'group': 'details', 'multiline': True},
]}


def receipts(value, *, maximum=9):
    from certificate_core import parse_date
    if not value:
        return []
    try:
        entries = json.loads(value)
    except (TypeError, ValueError) as exc:
        raise ValueError('Enter a received date and whole device count for each receipt.') from exc
    if not isinstance(entries, list) or len(entries) > maximum:
        raise ValueError(f'Only {maximum} device rows fit on the HDD certificate. Remove the extra rows.')
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
            fitted_text(str(batch.get(key) or '').upper(), width, minimum=8)
        except ValueError:
            warnings.append(f'Shorten the {"company name" if key == "supplier" else "city/location"} to fit the data destruction certificate.')
    return warnings


def fitted_text(value, width, minimum=8, preferred=14, bold=True):
    from certificate_templates import font_name
    from reportlab.pdfbase import pdfmetrics
    font = font_name({'font': 'Calibri', 'bold': -1 if bold else 0})
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
    if len(template.pages) < 2:
        raise ValueError('The Data Destruction template needs both its certificate and device receipt pages.')
    fields = batch.get('certificate_fields') or {}
    # An incomplete preview shows blanks and its validation messages, never invented dates/counts.
    try:
        rows = receipts(fields.get('device_receipts'))
    except ValueError:
        rows = []
    destroyed = parse_date(fields.get('destruction_date'))
    writer = PdfWriter()

    def page(index, draw):
        # A repeated details page needs independent content/resources. Reusing
        # pypdf's clone map makes later overlays appear on earlier pages too.
        writer.reset_translation(template)
        writer.add_page(template.pages[index])
        overlay = BytesIO()
        pdf = canvas.Canvas(overlay, pagesize=(960, 540))
        draw(pdf)
        pdf.save()
        writer.pages[-1].merge_page(PdfReader(overlay).pages[0])

    def text(pdf, value, left, baseline, width, preferred=14, bold=True):
        font, size, value = fitted_text(value, width, preferred=preferred, bold=bold)
        pdf.setFont(font, size)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawCentredString(left + width / 2, 540 - baseline, value)

    def front(pdf):
        text(pdf, batch.get('supplier', '').upper(), 184, 258, 590)
        text(pdf, str(batch.get('location') or '').upper(), 204, 295, 153)
        text(pdf, destroyed.strftime('%d/%m/%Y') if destroyed else '', 242, 332, 127)
        text(pdf, cod, 487, 332, 90)
        text(pdf, issued_date.strftime('%d/%m/%Y'), 619, 332, 115)

    page(0, front)
    groups = [rows[index:index + 9] for index in range(0, len(rows), 9)] or [[]]
    for index, group in enumerate(groups):
        def details(pdf):
            for row_index, (received, count) in enumerate(group):
                baseline = 99 + row_index * 20.65
                text(pdf, received.strftime('%d/%m/%Y'), 314, baseline, 165)
                text(pdf, str(count), 482, baseline, 165)
            text(pdf, str(sum(count for _, count in group)) if group else '', 482, 283, 165)
            if len(groups) > 1:
                text(pdf, f'Page subtotal above. Total devices: {sum(count for _, count in rows)}', 310, 307, 340, 10, bold=False)
        page(1, details)
    annexure = fields.get('_annexure')
    if annexure:
        from app_paths import DATA_ROOT
        from pdf_attachments import append_attachment
        append_attachment(writer, batch.get('_data_root', DATA_ROOT), annexure)
    writer.add_metadata({'/Title': f'Secured Data Destruction - {cod}', '/Subject': 'destruction',
                         '/CertificateDate': issued_date.isoformat(),
                         '/CertificatePlant': str(batch.get('source_id') or '')})
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open('wb') as stream:
        writer.write(stream)
    return destination


def _save_linked_request_locked(root, plant, source, link, data, *, needed):
    """Save a linked request using an inward already checked under the run lock."""
    from certificate_core import parse_date
    from certificate_drafts import write_register, clean_details, _manual_record_issued, _save_draft_locked

    current = next((record for record in data['records']
                    if record['details'].get('destruction_source') == link), None)
    if current:
        if current.get('destruction_requested', True) != needed:
            if not needed and _manual_record_issued(root, current):
                raise ValueError('This Data Destruction certificate has been issued. Open its row to review it or mark it Completed.')
            current['destruction_requested'] = needed
            write_register(root, data)
        return current['id']
    if not needed:
        return None
    received = parse_date(source.get('inward_date'))
    details = dict(source, certificate_type='destruction', certificate_fields={
        'device_receipts': json.dumps([{'date': received.isoformat() if received else '', 'count': ''}])},
                   destruction_source=link, cod_no='', cod_date='', reference='', items=[])
    details.pop('certificate_weight', None)
    return _save_draft_locked(root, plant, clean_details(details), data)


def request_for_inward(root, plant, case_id, expected_revision=None, *, needed=True, confirmed=False, record_id=None):
    """Save an idempotent additional request without writing to the inward workbook."""
    from certificate_core import read_batches
    from certificate_drafts import read_register
    from issuance_safety import recover_issuances
    from workflow_support import load_settings, run_lock
    from workspace_review import require_revision
    if not isinstance(needed, bool):
        raise ValueError('Choose whether Data Destruction is needed.')
    if (plant not in ('plant1', 'plant2') and not record_id) or (not record_id and not str(case_id or '').strip()):
        raise ValueError('Choose an inward client first.')
    case_id = str(case_id).strip()
    with run_lock(root):
        recover_issuances(root)
        data = read_register(root)
        if record_id:
            record = next((entry for entry in data['records'] if entry['id'] == record_id
                           and entry['plant'] == plant
                           and entry['details'].get('certificate_type') != 'destruction'), None)
            if not record:
                raise ValueError('The linked local certificate was not found. Refresh and try again.')
            source = record['details']
            local_sn = (record.get('removed_inward') or {}).get('case_id') or source.get('case_id') or record_id
            link = {'plant': None, 'case_id': str(local_sn), 'record_id': record_id}
        else:
            path = load_settings(root)['workbook_paths'][int(plant[-1]) - 1]
            if not path or not Path(path).is_file():
                raise ValueError('Connect this plant workbook and refresh before adding its data destruction request.')
            path = Path(path).resolve()
            source_bytes = path.read_bytes()
            matches = [batch for batch in read_batches(BytesIO(source_bytes))
                       if str(batch['case_id']).strip().casefold() == case_id.casefold()]
            if len(matches) != 1:
                raise ValueError('The inward Serial Number is missing or repeated. Correct Excel and refresh.')
            source = matches[0]
            link = {'plant': plant, 'case_id': str(source['case_id'])}

        if expected_revision is not None:
            require_revision(root, expected_revision)
            if not record_id and path.read_bytes() != source_bytes:
                raise ValueError('The Excel files changed while opening Data Destruction. Refresh and try again.')
        return _save_linked_request_locked(root, plant, source, link, data, needed=needed)
