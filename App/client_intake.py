"""Append a client block while keeping the original XLSX package intact."""
from copy import copy
from datetime import datetime
from io import BytesIO
import math
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import quoteattr
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.datetime import to_excel

from certificate_core import (batch_content, inward_details_complete, norm, number,
                              parse_date, read_batches, select_register, text)
from workbook_store import (_attributes, _date_style, _element, _range_fragment,
                            _set_cell, _sheet_member, _ensure_status_validation,
                            formatted_client_template, sync_generated_styles,
                            add_tracking_column)
from workbook_metadata import write_details


LABELS = {'location':'Location', 'quantity':'Quantity in Pcs',
          'per_piece_weight':'Per Pcs Weight in (KG)', 'weight':'Weight (KG)',
          'form6':'Form 6 No', 'reference':'Reference No.', 'cod_no':'COD No',
          'cod_date':'COD Date', 'issued':'Certificate Issued'}


def intake_details(raw):
    if not isinstance(raw, dict):
        raise ValueError('Enter the new client details.')
    details = {}
    for key in ('supplier', 'location', 'form6', 'reference'):
        details[key] = text(raw.get(key))
        if len(details[key]) > (500 if key in {'supplier', 'location'} else 100):
            raise ValueError(f'The {key} value is too long.')
    if not details['supplier']:
        raise ValueError('Enter the company name.')
    details['certificate_type'] = text(raw.get('certificate_type')) or 'standard'
    if details['certificate_type'] not in {'standard','sita','siemens'}:
        raise ValueError('Choose a valid certificate type.')
    received = text(raw.get('inward_date'))
    details['inward_date'] = parse_date(received)
    if received and details['inward_date'] is None:
        raise ValueError('Choose a valid received date.')
    source = raw.get('items', [])
    if not isinstance(source, list) or len(source) > 100:
        raise ValueError('Add up to 100 material lines at a time.')
    items = []
    for index, raw_item in enumerate(source, 1):
        if not isinstance(raw_item, dict):
            raise ValueError(f'Check material {index}.')
        item = {'item':text(raw_item.get('item'))}
        if len(item['item']) > 500:
            raise ValueError(f'Shorten the description for material {index}.')
        for key, label in (('quantity','quantity'), ('per_piece_weight','weight per piece'), ('weight','weight')):
            value = text(raw_item.get(key))
            item[key] = number(value) if value else None
            if value and (item[key] is None or not math.isfinite(item[key]) or item[key] < 0):
                raise ValueError(f'Enter a valid, non-negative {label} for material {index}.')
        if item['item'] or any(item[key] is not None for key in ('quantity','per_piece_weight','weight')):
            items.append(item)
    details['items'] = items or [{'item':'', 'quantity':None, 'per_piece_weight':None, 'weight':None}]
    return details


def _package(source, changes):
    buffer = BytesIO()
    with ZipFile(buffer, 'w') as target:
        target.comment = source.comment
        for info in source.infolist():
            target.writestr(copy(info), changes.get(info.filename, source.read(info.filename)))
    return buffer.getvalue()


def _append_block(xml, name, fragments, before):
    if not fragments:
        return xml
    found = _element(xml, name).search(xml)
    prefix = re.search(r'<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b', xml).group('p')
    if found:
        block = found.group()
        if block.rstrip().endswith('/>'):
            block = block.rstrip()[:-2] + f'></{prefix}{name}>'
        end = block.rfind('</')
        block = block[:end] + ''.join(fragments) + block[end:]
        if name in {'mergeCells', 'dataValidations'}:
            child = 'mergeCell' if name == 'mergeCells' else 'dataValidation'
            count = len(list(_element(block, child).finditer(block)))
            block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{count}"', block, count=1)
        return xml[:found.start()] + block + xml[found.end():]
    position = min([match.start() for tag in before if (match := _element(xml, tag).search(xml))]
                   or [xml.rfind('</')])
    count = f' count="{len(fragments)}"' if name in {'mergeCells','dataValidations'} else ''
    return xml[:position] + f'<{prefix}{name}{count}>' + ''.join(fragments) + f'</{prefix}{name}>' + xml[position:]


def _copy_rules(xml, source_rows, destination_rows, status_column, width):
    """Copy rules only for new cells outside an existing rule's coverage."""
    additions = {'dataValidation':[], 'conditionalFormatting':[]}
    priority = max([int(value) for value in re.findall(r'\bpriority="(\d+)"', xml)] or [0])
    for name in additions:
        for match in list(_element(xml, name).finditer(xml)):
            fragment = match.group()
            original = _attributes(fragment).get('sqref', '')
            ranges = [(left or 1, top or 1, right or width, bottom or 1048576)
                      for left,top,right,bottom in (range_boundaries(ref.replace('$','')) for ref in original.split())]
            targets = []
            for template, destination in zip(source_rows, destination_rows):
                for left, top, right, bottom in ranges:
                    if top <= template <= bottom:
                        for col in range(left, min(right, width) + 1):
                            if name == 'dataValidation' and col == status_column:
                                continue
                            if any(a <= col <= c and b <= destination <= d for a,b,c,d in ranges):
                                continue
                            targets.append(f'{get_column_letter(col)}{destination}')
            if not targets:
                continue
            copied = _range_fragment(fragment, ' '.join(dict.fromkeys(targets)), original)
            if name == 'conditionalFormatting':
                def new_priority(_match):
                    nonlocal priority
                    priority += 1
                    return f'priority="{priority}"'
                copied = re.sub(r'\bpriority=["\x27]\d+["\x27]', new_priority, copied)
            additions[name].append(copied)
    xml = _append_block(xml, 'dataValidations', additions['dataValidation'],
                        ('hyperlinks','printOptions','pageMargins','pageSetup','drawing','extLst'))
    position = min([m.start() for tag in ('dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','drawing','extLst')
                    if (m := _element(xml, tag).search(xml))] or [xml.rfind('</')])
    return xml[:position] + ''.join(additions['conditionalFormatting']) + xml[position:]


def stage_new_client(session, raw):
    details = intake_details(raw)
    source_data = add_tracking_column(session.data)
    original = read_batches(BytesIO(source_data))
    book = load_workbook(BytesIO(source_data), data_only=False)
    try:
        sheet, header_row, cols = select_register(book)
        if sheet.protection.sheet:
            raise ValueError('This Excel sheet is protected. Unprotect it before adding a client.')
        empty_template = not original and all(cell.value is None for row in
            sheet.iter_rows(min_row=header_row + 1) for cell in row)
        first = header_row + 1 if empty_template else sheet.max_row + 1
        last = first + len(details['items']) - 1
        if last > 1048576:
            raise ValueError('This Excel sheet has no room for another client.')
        tables = list(sheet.tables.values())
        if tables and (len(tables) != 1 or tables[0].totalsRowCount
                       or range_boundaries(tables[0].ref)[1:] != (header_row, sheet.max_column, first - 1)
                       or any(key not in cols for key in LABELS)):
            raise ValueError('This register has a separate table total or an unusual table layout. Add its entry in Excel for now; the file has not been changed.')
        ids = {batch['case_id'] for batch in original}
        sequence = max([int(value) for value in ids if re.fullmatch(r'\d{1,12}', value)] or [0]) + 1
        while str(sequence) in ids:
            sequence += 1
        case_id = str(sequence)
        # Copy the register's real client separator and plain supplier style.
        # Later imported rows may have no borders, so don't select by recency.
        template = formatted_client_template(sheet, header_row, cols, len(details['items']) > 1)
        template_end = max([area.max_row for area in sheet.merged_cells.ranges
                            if area.min_row == template and area.max_row > template] or [template])
        continuation = template + 1 if template_end > template else template
        count = len(details['items'])
        source_rows = ([template] if count == 1 else
                       [template] + [continuation] * (count - 2) + [template_end])
        destinations = list(range(first, last + 1))
        with ZipFile(BytesIO(source_data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            styles = archive.read('xl/styles.xml').decode('utf-8')
            prefix = re.search(r'<(?P<p>(?:[A-Za-z_][\w.-]*:)?)worksheet\b', xml).group('p')
            if empty_template:
                # Reuse an empty formatted register instead of leaving its
                # old blank client blocks above the first app-created client.
                xml = _element(xml, 'row').sub(lambda m: '' if int(_attributes(m.group())['r']) > header_row else m.group(), xml)
                def keep_header_merges(match):
                    block = _element(match.group(), 'mergeCell').sub(
                        lambda m: '' if range_boundaries(_attributes(m.group())['ref'])[1] > header_row else m.group(), match.group())
                    count = len(list(_element(block, 'mergeCell').finditer(block)))
                    return re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{count}"', block, count=1) if count else ''
                xml = _element(xml, 'mergeCells').sub(keep_header_merges, xml)
            labels = {'issued':'Certificate Issued'}
            for key, label in labels.items():
                if key not in cols:
                    cols[key] = max(sheet.max_column, max(cols.values()) + 1)
                    style = sheet.cell(header_row, cols['supplier'] + 1).style_id
                    xml = _set_cell(xml, f'{get_column_letter(cols[key] + 1)}{header_row}', label, 'inlineStr', style)
            width = max(sheet.max_column, max(cols.values()) + 1)
            fragments = []
            for source_row, destination in zip(source_rows, destinations):
                attributes = {'r':str(destination)}
                row_dimension = sheet.row_dimensions.get(source_row)
                if row_dimension and row_dimension.height is not None:
                    attributes.update(ht=str(row_dimension.height), customHeight='1')
                cells = ''.join(f'<{prefix}c r="{get_column_letter(col)}{destination}" s="{sheet.cell(source_row, min(col, sheet.max_column)).style_id}"/>'
                                for col in range(1, width + 1))
                fragments.append(f'<{prefix}row ' + ' '.join(f'{k}={quoteattr(v)}' for k,v in attributes.items()) + f'>{cells}</{prefix}row>')
            xml = _append_block(xml, 'sheetData', fragments, ())
            styles = sync_generated_styles(book, styles)
            for key in ('supplier','location','form6','reference'):
                if key not in cols:
                    continue
                xml = _set_cell(xml, f'{get_column_letter(cols[key] + 1)}{first}', details[key], 'inlineStr' if details[key] else 'blank')
            if 'certificate_type' in cols:
                xml = _set_cell(xml, f'{get_column_letter(cols["certificate_type"] + 1)}{first}', details['certificate_type'], 'inlineStr')
            xml = _set_cell(xml, f'{get_column_letter(cols["sn"] + 1)}{first}', case_id, 'inlineStr')
            if details['inward_date']:
                source_cell = sheet.cell(template, cols['inward_date'] + 1)
                style = source_cell.style_id
                if not is_date_format(source_cell.number_format):
                    styles, style = _date_style(styles, style, 'dd.mm.yyyy')
                value = to_excel(datetime.combine(details['inward_date'], datetime.min.time()), book.epoch)
                xml = _set_cell(xml, f'{get_column_letter(cols["inward_date"] + 1)}{first}', value, 'n', style)
            for row, material in zip(destinations, details['items']):
                xml = _set_cell(xml, f'{get_column_letter(cols["item"] + 1)}{row}', material['item'], 'inlineStr' if material['item'] else 'blank')
                for key in ('quantity','per_piece_weight','weight'):
                    if key not in cols:
                        continue
                    value = material[key]
                    if key == 'weight' and value is None and 'per_piece_weight' not in cols:
                        value = (material['quantity'] * material['per_piece_weight']
                                 if material['quantity'] is not None and material['per_piece_weight'] is not None else None)
                    xml = _set_cell(xml, f'{get_column_letter(cols[key] + 1)}{row}', value, 'n' if value is not None else 'blank')
                if material['weight'] is None and {'quantity','per_piece_weight','weight'}.issubset(cols):
                    # These are the same bounded inputs the certificate reader recalculates.
                    coordinate = f'{get_column_letter(cols["weight"] + 1)}{row}'
                    formula = f'{get_column_letter(cols["quantity"] + 1)}{row}*{get_column_letter(cols["per_piece_weight"] + 1)}{row}'
                    cell = next(m for m in _element(xml, 'c').finditer(xml) if _attributes(m.group()).get('r') == coordinate)
                    fragment = cell.group().replace(f'</{prefix}c>', f'<{prefix}f>{formula}</{prefix}f></{prefix}c>')
                    xml = xml[:cell.start()] + fragment + xml[cell.end():]
            total_cols = [index for index, cell in enumerate(sheet[header_row], 1)
                          if norm(cell.value) in {'totalweightinkg','totalweightkg','totalweight'}]
            if total_cols:
                weight_col = get_column_letter(cols['weight'] + 1)
                coordinate = f'{get_column_letter(total_cols[0])}{first}'
                cell = next(m for m in _element(xml, 'c').finditer(xml) if _attributes(m.group()).get('r') == coordinate)
                fragment = cell.group().rstrip()
                if fragment.endswith('/>'):
                    fragment = fragment[:-2] + f'></{prefix}c>'
                fragment = fragment.replace(f'</{prefix}c>', f'<{prefix}f>SUM({weight_col}{first}:{weight_col}{last})</{prefix}f></{prefix}c>')
                xml = xml[:cell.start()] + fragment + xml[cell.end():]
            merges = []
            for area in sheet.merged_cells.ranges:
                if len(destinations) > 1 and area.min_row == template and area.min_col == area.max_col and area.max_row > template:
                    column = get_column_letter(area.min_col)
                    if area.min_col - 1 not in {cols[key] for key in ('item','quantity','per_piece_weight','weight') if key in cols}:
                        merges.append(f'<{prefix}mergeCell ref="{column}{first}:{column}{last}"/>')
            xml = _append_block(xml, 'mergeCells', merges, ('phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','drawing','extLst'))
            xml = _copy_rules(xml, source_rows, destinations, cols['issued'] + 1, width)
            xml = _ensure_status_validation(xml, get_column_letter(cols['issued'] + 1), max(10000,last), header_row + 1)
            dimension = _element(xml, 'dimension').search(xml)
            if dimension:
                fragment = re.sub(r'\bref=["\x27][^"\x27]*["\x27]', f'ref="A1:{get_column_letter(width)}{last}"', dimension.group())
                xml = xml[:dimension.start()] + fragment + xml[dimension.end():]
            autofilter = _element(xml, 'autoFilter').search(xml)
            if autofilter:
                bounds = range_boundaries(_attributes(autofilter.group())['ref'])
                if bounds[1] == header_row and bounds[3] == first - 1:
                    ref = f'{get_column_letter(bounds[0])}{header_row}:{get_column_letter(max(bounds[2],width))}{last}'
                    fragment = re.sub(r'\bref=["\x27][^"\x27]*["\x27]', 'ref=' + quoteattr(ref), autofilter.group(), count=1)
                    xml = xml[:autofilter.start()] + fragment + xml[autofilter.end():]
            changes = {member:xml.encode('utf-8'), 'xl/styles.xml':styles.encode('utf-8')}
            if tables:
                table = tables[0]
                table_member = next(name for name in archive.namelist() if name.startswith('xl/tables/') and name.endswith('.xml')
                                    and ET.fromstring(archive.read(name)).attrib.get('name') == table.name)
                table_xml = archive.read(table_member).decode('utf-8')
                left = range_boundaries(table.ref)[0]
                ref = f'{get_column_letter(left)}{header_row}:{get_column_letter(width)}{last}'
                table_xml = re.sub(r'\bref=["\x27]' + re.escape(table.ref) + r'["\x27]', 'ref=' + quoteattr(ref), table_xml)
                if width > sheet.max_column:
                    section = _element(table_xml, 'tableColumns').search(table_xml)
                    block = section.group()
                    elements = list(_element(block, 'tableColumn').finditer(block))
                    next_id = max(int(_attributes(column.group())['id']) for column in elements) + 1
                    table_prefix = section.group('p')
                    names = {position + 1:labels[key] for key,position in cols.items() if key in labels}
                    columns_xml = ''.join(f'<{table_prefix}tableColumn id="{next_id + offset}" name={quoteattr(names[col])}/>'
                                          for offset,col in enumerate(range(sheet.max_column + 1,width + 1)))
                    end = block.rfind('</')
                    block = block[:end] + columns_xml + block[end:]
                    block = re.sub(r'\bcount=["\x27]\d+["\x27]', f'count="{len(elements) + width - sheet.max_column}"', block, count=1)
                    table_xml = table_xml[:section.start()] + block + table_xml[section.end():]
                changes[table_member] = table_xml.encode('utf-8')
            workbook_xml = archive.read('xl/workbook.xml').decode('utf-8')
            calculation = _element(workbook_xml, 'calcPr').search(workbook_xml)
            if calculation:
                fragment = calculation.group()
                for key in ('fullCalcOnLoad','forceFullCalc'):
                    if re.search(rf'\b{key}=', fragment):
                        fragment = re.sub(rf'\b{key}=["\x27][^"\x27]*["\x27]', f'{key}="1"', fragment)
                    else:
                        fragment = fragment.replace('/>', f' {key}="1"/>')
                workbook_xml = workbook_xml[:calculation.start()] + fragment + workbook_xml[calculation.end():]
                changes['xl/workbook.xml'] = workbook_xml.encode('utf-8')
            metadata = {key:details[key] for key in ('form6','reference','certificate_type')
                        if key not in cols and details[key] and (key != 'certificate_type' or details[key] != 'standard')}
            updated = write_details(_package(archive, changes), sheet.title, case_id, metadata)
            batches = read_batches(BytesIO(updated))
            added = next(batch for batch in batches if batch['case_id'] == case_id)
            ready = inward_details_complete(added)
            xml = _set_cell(xml, f'{get_column_letter(cols["issued"] + 1)}{first}', 'Ready' if ready else '', 'inlineStr' if ready else 'blank')
            changes[member] = xml.encode('utf-8')
            ET.fromstring(xml)
            updated = write_details(_package(archive, changes), sheet.title, case_id, metadata)
            checked = read_batches(BytesIO(updated))
            if len(checked) != len(original) + 1 or any(batch_content(a) != batch_content(b) for a,b in zip(original,checked)):
                raise ValueError('The new Excel entry could not be verified. The original workbook was kept.')
            added = checked[-1]
            if added['supplier'] != details['supplier'] or added['case_id'] != case_id:
                raise ValueError('The client details could not be verified. The original workbook was kept.')
            return updated, added
    finally:
        book.close()
