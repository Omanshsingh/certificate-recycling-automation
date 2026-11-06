"""Record a linked HDD certificate below recycling COD/date in the same inward."""
from io import BytesIO
import re
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import to_excel

from register_protection import externally_protected
from certificate_core import read_batches, select_register, text
from workbook_metadata import write_details
from workbook_store import _attributes, _element, _sheet_member, _set_cell, _patch_cell_styles, order_register_for_save


def stage_hdd_tracking(session, sn, cod, issued_date):
    data = session.data
    matches = [b for b in read_batches(BytesIO(data)) if text(b['case_id']).casefold() == text(sn).casefold()]
    if len(matches) != 1:
        raise ValueError('The linked inward is missing or its SN is repeated. Refresh Excel before creating the HDD certificate.')
    batch = matches[0]
    row = batch['source_row'] + 1
    following = min([b['source_row'] for b in read_batches(BytesIO(data)) if b['source_row'] > batch['source_row']] or [1048576])
    if following == row or row in batch.get('summary_rows', []):
        from register_rows import shift_register_rows
        data = shift_register_rows(data, row, 1)
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, cols = select_register(book)
        if externally_protected(sheet) or not {'cod_no', 'cod_date'}.issubset(cols):
            raise ValueError('The linked inward needs writable COD No and COD Date columns before creating its HDD certificate.')
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            style_ids = {}
            for key, value in (('cod_no', cod), ('cod_date', to_excel(issued_date, book.epoch))):
                column = cols[key] + 1
                current = sheet.cell(row, column)
                if current.data_type == 'f':
                    raise ValueError('The HDD tracking line contains a formula. Check Excel before creating the certificate.')
                coordinate = f'{get_column_letter(column)}{row}'
                style_ids[coordinate] = sheet.cell(batch['source_row'], column).style_id
                xml = _patch_cell_styles(xml, {coordinate: style_ids[coordinate]})
                xml = _set_cell(xml, coordinate, value, 'inlineStr' if key == 'cod_no' else 'n')
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
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, xml.encode('utf-8') if info.filename == member else archive.read(info.filename))
        updated = write_details(output.getvalue(), sheet.title, batch['case_id'],
                                {'_hdd_cod_no': str(cod), '_hdd_cod_date': issued_date.isoformat()})
        updated = order_register_for_save(updated, session)
        actual = next(b for b in read_batches(BytesIO(updated)) if b['case_id'] == batch['case_id'])
        if (actual['cod_no'], actual['cod_date'], actual['hdd_cod_no']) != (batch['cod_no'], batch['cod_date'], str(cod)):
            raise ValueError('The recycling and HDD numbers could not be verified. Excel was kept unchanged.')
        return updated
    finally:
        book.close()
