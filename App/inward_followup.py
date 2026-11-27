"""Lower-priority inward details, kept separate from certificate readiness."""
from io import BytesIO
from zipfile import ZipFile

from register_reader import load_workbook

from register_protection import externally_protected
from certificate_core import batch_content, read_batches, select_register, text
from workbook_store import _check_batch_merge, _set_cell, _sheet_member, checked_excel_text

FIELDS = {
    'comments': ('Comments', 1000),
    'transporter_details': ('Transporter / vehicle details', 1000),
    'docket_vehicle_no': ('Docket no. / vehicle no.', 500),
    'weight_slip_scale': ('Weight slip / scale weight', 500),
}


def clean_followup(raw):
    if not isinstance(raw, dict) or set(raw) != set(FIELDS):
        raise ValueError('Send only the four inward details shown in this form.')
    result = {}
    for key, (label, limit) in FIELDS.items():
        if not isinstance(raw[key], str):
            raise ValueError(f'Enter text for {label}.')
        value = checked_excel_text(raw[key].strip(), label, normalize_paste=True)
        if len(value) > limit:
            raise ValueError(f'{label} is too long. Use {limit} characters or fewer.')
        result[key] = value
    return result


def stage_followup(data, sn, values):
    """Patch only the four cells; never change certificate status or content."""
    values = clean_followup(values)
    original = read_batches(BytesIO(data))
    matches = [row for row in original if row['case_id'] == sn]
    if len(matches) != 1 or matches[0].get('duplicate_serial_number'):
        raise ValueError('This SN is missing or repeated. Check Excel, then refresh.')
    source = matches[0]
    if all(key in source['column_mapping'] and source[key] == value for key, value in values.items()):
        return data, source
    if any(key not in source['column_mapping'] for key in FIELDS):
        from register_column_order import normalize_register_columns
        data = normalize_register_columns(data)
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, cols = select_register(book)
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            for key, value in values.items():
                cell = sheet.cell(source['source_row'], cols[key] + 1)
                if text(cell.value) == value:
                    continue
                _check_batch_merge(sheet, cell, source['source_row'], cols)
                if cell.data_type == 'f':
                    raise ValueError(f'{FIELDS[key][0]} uses an Excel formula. Enter its value in Excel first.')
                if externally_protected(sheet) and cell.protection.locked:
                    raise ValueError(f'{FIELDS[key][0]} is locked in Excel. Unlock it before saving.')
                xml = _set_cell(xml, cell.coordinate, value, 'inlineStr' if value else 'blank')
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, xml.encode('utf-8') if info.filename == member else archive.read(info.filename))
        updated = output.getvalue()
    finally:
        book.close()
    verified = read_batches(BytesIO(updated))
    if len(verified) != len(original):
        raise ValueError('The inward details could not be checked. Excel was kept unchanged.')
    for before, after in zip(original, verified):
        if (batch_content(before) != batch_content(after)
                or any(before.get(key) != after.get(key) for key in ('cod_no', 'cod_date', 'issued'))):
            raise ValueError('Certificate details changed unexpectedly. Excel was kept unchanged.')
        if before['case_id'] != sn and any(before.get(key) != after.get(key) for key in FIELDS):
            raise ValueError('Another inward changed unexpectedly. Excel was kept unchanged.')
    saved = next(row for row in verified if row['case_id'] == sn)
    if any(saved.get(key) != value for key, value in values.items()):
        raise ValueError('The inward details could not be checked. Excel was kept unchanged.')
    from register_presentation import normalize_register_presentation
    updated = normalize_register_presentation(updated, wrap_only=True, fit_inward=(sheet.title, sn))
    return updated, saved
