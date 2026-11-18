"""Record linked HDD tracking without adding a row for an HDD-only inward."""
from io import BytesIO
import re
from types import SimpleNamespace
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel

from register_protection import externally_protected
from register_rows import shift_register_rows
from certificate_core import parse_date, read_batches, select_register, status_value, text
from certificate_inventory import identity as cod_identity
from excel_values import excel_date
from workbook_metadata import _replace_details, read_details, write_details
from workbook_store import (_attributes, _date_style, _element, _sheet_member, _set_cell,
                            _patch_cell_styles, order_register_for_save, sync_generated_styles)


def hdd_uses_main_row(batch):
    return (not batch['cod_no'] and not batch['cod_date']
            and (batch.get('hdd_on_main_row') or status_value(batch['issued']) == 'na'))

def hdd_repair_matches_inward(batch, details):
    """Require evidence beyond a serial number before repairing an issued HDD link."""
    tracked = cod_identity(batch.get('hdd_cod_no'))
    if tracked and tracked != cod_identity(details.get('cod_no')):
        return False
    form6 = text(details.get('form6')).casefold()
    source_form6 = text(batch.get('form6')).casefold()
    if form6 and source_form6:
        return form6 == source_form6
    received = parse_date(details.get('inward_date'))
    supplier = text(details.get('supplier')).casefold()
    reference = text(details.get('reference')).casefold()
    return bool(received and received == batch.get('inward_date')
                and supplier and supplier == text(batch.get('supplier')).casefold()
                and (batch.get('hdd_on_main_row')
                     or (reference and reference == text(batch.get('reference')).casefold())))


def recover_hdd_tracking_link(session, batches, record, issued_date):
    """Relink a missing SN only when Excel still proves the HDD's identity."""
    from inward_serials import LINK_KEY, links_from
    details = record['details']
    cod = text(details['cod_no'])
    form6 = text(details.get('form6')).casefold()
    inward_date = parse_date(details.get('inward_date'))
    if not form6 or not inward_date or not text(details.get('supplier')):
        return None
    claims = [batch for batch in batches if cod in batch.get('reserved_cod_numbers', ())
              or cod in (batch.get('cod_no'), batch.get('hdd_cod_no'))]
    if len(claims) != 1:
        return None
    batch = claims[0]
    if (not hdd_uses_main_row(batch) or batch.get('duplicate_serial_number')
            or text(batch['form6']).casefold() != form6 or batch['inward_date'] != inward_date
            or text(batch['supplier']).casefold() != text(details['supplier']).casefold()
            or text(batch['reference']).casefold() != text(details.get('reference')).casefold()):
        return None
    book = load_workbook(BytesIO(session.data), data_only=False)
    try:
        sheet, _, cols = select_register(book)
        continuation = batch['source_row'] + 1
        if (any(other['source_row'] == continuation for other in batches)
                or not {'cod_no', 'cod_date'}.issubset(cols)
                or text(sheet.cell(continuation, cols['cod_no'] + 1).value) != cod
                or parse_date(sheet.cell(continuation, cols['cod_date'] + 1).value) != issued_date):
            return None
    finally:
        book.close()
    data = write_details(session.data, batch['source_sheet'], batch['case_id'],
                         {'_hdd_cod_no': cod, '_hdd_cod_date': issued_date.isoformat()})
    saved = read_details(BytesIO(data))
    links = links_from(data)
    links.setdefault('manual_sources', {})[record['id']] = batch['case_id']
    saved[LINK_KEY] = links
    staged = SimpleNamespace(data=_replace_details(data, saved))
    return stage_hdd_tracking(staged, batch['case_id'], cod, issued_date)


def stage_hdd_tracking(session, sn, cod, issued_date, *, format_register=True):
    from workbook_store import add_tracking_column
    batches = read_batches(BytesIO(session.data))
    data = add_tracking_column(session.data)
    if data != session.data:
        batches = read_batches(BytesIO(data))
    matches = [b for b in batches if text(b['case_id']).casefold() == text(sn).casefold()]
    if len(matches) != 1:
        raise ValueError('The linked inward is missing or its SN is repeated. Refresh Excel before creating the HDD certificate.')
    batch = matches[0]
    tracked_cod = cod_identity(batch.get('hdd_cod_no'))
    if tracked_cod and tracked_cod != cod_identity(cod):
        raise ValueError('This inward already tracks a different HDD number. Check the existing certificate before changing Excel.')
    first = batch['source_row']
    continuation = first + 1
    use_main = hdd_uses_main_row(batch)
    row = first if use_main else continuation
    following = min((b['source_row'] for b in batches if b['source_row'] > first), default=1048577)
    if not use_main and (following == row or row in batch.get('summary_rows', [])):
        data = shift_register_rows(data, row, 1, expand_shared=True)
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, cols = select_register(book)
        if externally_protected(sheet) or not {'cod_no', 'cod_date'}.issubset(cols):
            raise ValueError('The linked inward needs writable COD No and COD Date columns before creating its HDD certificate.')
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            date_style_id = None
            style_ids = {}
            remove_old_row = False
            old_number = cod_identity(sheet.cell(row, cols['cod_no'] + 1).value)
            old_date = parse_date(batch.get('hdd_cod_date'))
            for key, value in (('cod_no', cod), ('cod_date', to_excel(issued_date, book.epoch))):
                column = cols[key] + 1
                current = sheet.cell(row, column)
                if current.data_type == 'f':
                    raise ValueError('The HDD tracking line contains a formula. Check Excel before creating the certificate.')
                if key == 'cod_no':
                    matches_existing = old_number == cod_identity(cod)
                else:
                    date_in_cell = excel_date(current.value, book.epoch, parse_date)
                    matches_existing = (date_in_cell == issued_date
                                        or (tracked_cod and old_number == tracked_cod and date_in_cell == old_date))
                if current.value not in (None, '') and not matches_existing:
                    raise ValueError(f'{current.coordinate} already contains tracking data. Check the existing HDD number and date in Excel.')
                coordinate = f'{get_column_letter(column)}{row}'
                style_ids[coordinate] = sheet.cell(batch['source_row'], column).style_id
                if key == 'cod_date' and not is_date_format(sheet.cell(batch['source_row'], column).number_format):
                    date_style_id = coordinate
                xml = _patch_cell_styles(xml, {coordinate: style_ids[coordinate]})
                xml = _set_cell(xml, coordinate, value, 'inlineStr' if key == 'cod_no' else 'n')
            if use_main and not batch.get('hdd_on_main_row') and batch.get('hdd_cod_no') and continuation < following:
                old_cod = sheet.cell(continuation, cols['cod_no'] + 1)
                old_date = sheet.cell(continuation, cols['cod_date'] + 1)
                if (text(old_cod.value) == batch['hdd_cod_no']
                        and parse_date(old_date.value) == parse_date(batch['hdd_cod_date'])
                        and not any(cell.comment or cell.hyperlink for cell in (old_cod, old_date))):
                    remove_old_row = not any(
                        cell.value is not None or cell.comment or cell.hyperlink
                        for cell in sheet[continuation] if cell.column not in {old_cod.column, old_date.column}
                    )
                    for cell in (old_cod, old_date):
                        xml = _set_cell(xml, cell.coordinate, '', 'blank')
                elif text(old_cod.value) == batch['hdd_cod_no']:
                    raise ValueError('The HDD tracking row was edited in Excel. Check its date and notes before moving it.')
            xml = _patch_cell_styles(xml, style_ids)
            section = _element(xml, 'mergeCells').search(xml)
            if section:
                def unmerge(match):
                    area = _attributes(match.group())['ref']
                    return '' if any(area == str(merged) and merged.min_col in {cols['cod_no'] + 1, cols['cod_date'] + 1}
                                     and merged.min_row <= row <= merged.max_row for merged in sheet.merged_cells.ranges) else match.group()
                block = _element(section.group(), 'mergeCell').sub(unmerge, section.group())
                count = len(list(_element(block, 'mergeCell').finditer(block)))
                block = re.sub(r'\bcount=(["\x27]).*?\1', f'count="{count}"', block, count=1) if count else ''
                xml = xml[:section.start()] + block + xml[section.end():]
            output = BytesIO()
            styles = sync_generated_styles(book, archive.read('xl/styles.xml').decode('utf-8'))
            if date_style_id:
                styles, date_style = _date_style(styles, style_ids[date_style_id], 'dd/mm/yy')
                xml = _patch_cell_styles(xml, {date_style_id: date_style})
            updates = {member: xml.encode('utf-8'), 'xl/styles.xml': styles.encode('utf-8')}
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, updates.get(info.filename, archive.read(info.filename)))
        updated = write_details(output.getvalue(), sheet.title, batch['case_id'],
                                {'_hdd_cod_no': str(cod), '_hdd_cod_date': issued_date.isoformat(),
                                 '_hdd_primary': use_main})
        if remove_old_row:
            try:
                updated = shift_register_rows(updated, continuation, 1, delete=True)
            except ValueError:
                pass  # Excel formulas or objects can require keeping an empty row.
        if format_register:
            updated = order_register_for_save(updated, session)
        actual = next(b for b in read_batches(BytesIO(updated)) if b['case_id'] == batch['case_id'])
        if ((actual['cod_no'], actual['cod_date'], actual['hdd_cod_no'], parse_date(actual['hdd_cod_date']))
                != (batch['cod_no'], batch['cod_date'], str(cod), issued_date)
                or bool(actual.get('hdd_on_main_row')) != bool(use_main)):
            raise ValueError('The recycling and HDD numbers could not be verified. Excel was kept unchanged.')
        return updated
    finally:
        book.close()
