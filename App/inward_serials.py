"""Keep date-ordered SNs and their historical links in the same atomic workbook save."""
from copy import deepcopy
from functools import lru_cache
from io import BytesIO
import json
from pathlib import Path
import re
import threading
from zipfile import BadZipFile, ZipFile

from register_protection import externally_protected
from workbook_metadata import read_details, detail_key, _replace_details

LINK_KEY = '__EcorecoInwardNumberLinks__'
_LOCAL = threading.local()


def links_from(source):
    if isinstance(source, (str, Path)):
        from file_fingerprints import file_signature
        path = Path(source).resolve()
        return deepcopy(_cached_links(str(path), file_signature(path)))
    if isinstance(source, bytes):
        source = BytesIO(source)
    links = read_details(source).get(LINK_KEY, {})
    if not isinstance(links, dict) or any(not isinstance(links.get(key, {}), dict)
                                          for key in ('receipts', 'manual_sources', 'additions')):
        raise ValueError('The saved inward number links need checking. The Excel file was kept.')
    return links


@lru_cache(maxsize=128)
def _cached_links(path, signature):
    return links_from(BytesIO(Path(path).read_bytes()))


def bind_workbook(path, data):
    active = getattr(_LOCAL, 'workbooks', None)
    if active is None:
        active = _LOCAL.workbooks = {}
    active[str(Path(path).resolve())] = links_from(data)


def unbind_workbook(path):
    getattr(_LOCAL, 'workbooks', {}).pop(str(Path(path).resolve()), None)


def workspace_links(root):
    from workflow_support import load_settings
    active = getattr(_LOCAL, 'workbooks', {})
    result = {}
    for index, value in enumerate(load_settings(root).get('workbook_paths', []), 1):
        if not value:
            continue
        path = Path(value).expanduser().resolve()
        try:
            result[f'plant{index}'] = active[str(path)] if str(path) in active else links_from(path)
        except (OSError, BadZipFile):
            # Completion of an available plant can still proceed when the
            # other configured file is unavailable.
            continue
    return result


def project_receipts(root, records, *, links=None):
    links = workspace_links(root) if links is None else links
    for record in records:
        if record.get('state') != 'complete':
            continue
        plant_links = links.get(record.get('plant'), {})
        if record.get('source_kind') == 'workbook':
            current = plant_links.get('receipts', {}).get(record.get('id'))
            if current is not None:
                record['recorded_key'] = record.get('key')
                record['key'] = current
                for field in ('before', 'after'):
                    if isinstance(record.get(field), dict) and 'case_id' in record[field]:
                        record[field]['case_id'] = current
        elif record.get('source_kind') == 'manual':
            for source_plant, source_links in links.items():
                current = source_links.get('manual_sources', {}).get(record.get('key'))
                if current is None:
                    continue
                for field in ('before', 'after'):
                    details = record.get(field, {})
                    if isinstance(details, dict):
                        details = details.get('details', details)
                        link = details.get('destruction_source') if isinstance(details, dict) else None
                        if isinstance(link, dict) and not link.get('record_id') and link.get('plant') == source_plant:
                            link['case_id'] = current
    return records


def project_manual_sources(root, data):
    for plant, links in workspace_links(root).items():
        for record in data.get('records', []):
            current = links.get('manual_sources', {}).get(record.get('id'))
            link = record.get('details', {}).get('destruction_source')
            if current is not None and isinstance(link, dict) and not link.get('record_id') and link.get('plant') == plant:
                link['case_id'] = current
    return data


def resequence_added_inward(data, added, original_data, *, root, plant, request_id):
    """Insert an automatic numeric SN in date order, shifting following numeric SNs."""
    from certificate_core import read_batches, batch_content, select_register, text
    from workbook_store import _set_cell, _sheet_member, sort_register_by_inward_date
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from issuance_safety import audit_records, workbook_receipt_conflict
    from certificate_drafts import read_register

    before = read_batches(BytesIO(data))
    target_rows = [row for row in before if row['source_sheet'] == added['source_sheet']]
    if target_rows != sorted(target_rows, key=lambda row:(row['inward_date'] is None, row['inward_date'])):
        data = sort_register_by_inward_date(data)
        before = read_batches(BytesIO(data))
    original = read_batches(BytesIO(original_data))
    if len({row['case_id'].strip().casefold() for row in original}) != len(original):
        raise ValueError('Some SNs are repeated in this plant. Give each existing entry its own SN before adding another.')
    numeric = [row for row in before if row['source_sheet'] == added['source_sheet']
               and re.fullmatch(r'\d{1,12}', row['case_id'])]
    if len({int(row['case_id']) for row in numeric}) != len(numeric):
        raise ValueError('Some SNs are repeated in this plant. Give each existing entry its own SN before adding another.')
    old_numeric = [row for row in original if row['source_sheet'] == added['source_sheet']
                   and re.fullmatch(r'\d{1,12}', row['case_id'])]
    first = min([int(row['case_id']) for row in old_numeric] or [1])
    renumber = {}
    for number, row in enumerate(numeric, first):
        old = row['case_id']
        new = old if int(old) == number else str(number).zfill(len(old)) if old.startswith('0') else str(number)
        if new != old:
            renumber[old] = new
    final_sn = renumber.get(added['case_id'], added['case_id'])
    if not renumber:
        return data, next(row for row in before if row['source_sheet'] == added['source_sheet']
                          and row['case_id'] == final_sn), {}
    original_ids = {row['case_id'] for row in original}
    changes = {old:new for old,new in renumber.items() if old in original_ids}
    saved = read_details(BytesIO(data))
    links = deepcopy(saved.get(LINK_KEY, {}))
    for field in ('receipts', 'manual_sources', 'additions'):
        links.setdefault(field, {})
    original_by_sn = {row['case_id'].strip().casefold():row for row in original}
    for receipt in audit_records(root):
        old = str(receipt.get('key', ''))
        if receipt.get('state') == 'complete' and receipt.get('source_kind') == 'workbook' and receipt.get('plant') == plant and old in changes:
            conflict = workbook_receipt_conflict(original_by_sn[old.strip().casefold()], receipt)
            if conflict:
                raise ValueError(conflict)
            links['receipts'][receipt['id']] = changes[old]
    for record in read_register(root)['records']:
        link = record['details'].get('destruction_source')
        if isinstance(link, dict) and not link.get('record_id') and link.get('plant') == plant and link.get('case_id') in changes:
            links['manual_sources'][record['id']] = changes[link['case_id']]
    for journal_path in (Path(root) / 'Data' / 'Client Additions').glob('*.json'):
        journal = json.loads(journal_path.read_text(encoding='utf-8'))
        current = links['additions'].get(journal_path.stem, journal.get('sn'))
        if journal.get('plant') == plant and current in changes:
            links['additions'][journal_path.stem] = changes[current]
    if request_id is not None:
        links['additions'][request_id] = final_sn
    rebuilt = {}
    numeric_aliases = {int(old):new for old,new in renumber.items()}
    for key, value in saved.items():
        try:
            identity = json.loads(key)
        except ValueError:
            identity = None
        if isinstance(identity, list) and len(identity) == 2 and identity[0] == added['source_sheet']:
            old_key = str(identity[1])
            identity[1] = renumber.get(old_key, numeric_aliases.get(int(old_key), old_key)
                                       if re.fullmatch(r'\d{1,12}', old_key) else old_key)
            key = detail_key(*identity)
        if key in rebuilt and rebuilt[key] != value:
            raise ValueError('The saved details for these SNs overlap. The Excel file was kept.')
        rebuilt[key] = value
    if any(links.values()):
        rebuilt[LINK_KEY] = links
    else:
        rebuilt.pop(LINK_KEY, None)
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, _, cols = select_register(book)
        if externally_protected(sheet):
            raise ValueError('Unprotect the sheet before changing its SNs.')
        with ZipFile(BytesIO(data)) as archive:
            member = _sheet_member(archive, sheet.title)
            xml = archive.read(member).decode('utf-8')
            for row in numeric:
                if row['case_id'] not in renumber:
                    continue
                for row_number in [row['source_row'], *row.get('summary_rows', [])]:
                    cell = sheet.cell(row_number, cols['sn'] + 1)
                    if row_number != row['source_row']:
                        raw_sn = text(cell.value)
                        if raw_sn != row['case_id'] and not (re.fullmatch(r'\d{1,12}', raw_sn)
                                                             and int(raw_sn) == int(row['case_id'])):
                            continue
                    if cell.data_type == 'f':
                        raise ValueError('An SN is calculated by an Excel formula. Enter plain SN values before adding a backdated entry.')
                    xml = _set_cell(xml, f'{get_column_letter(cols["sn"] + 1)}{row_number}', renumber[row['case_id']], 'inlineStr')
            output = BytesIO()
            with ZipFile(output, 'w') as target:
                target.comment = archive.comment
                for info in archive.infolist():
                    target.writestr(info, xml.encode('utf-8') if info.filename == member else archive.read(info.filename))
    finally:
        book.close()
    updated = _replace_details(output.getvalue(), rebuilt)
    after = read_batches(BytesIO(updated))
    for previous in before:
        current = next((row for row in after if row['source_sheet'] == previous['source_sheet']
                        and row['case_id'] == renumber.get(previous['case_id'], previous['case_id'])), None)
        expected = dict(batch_content(previous), case_id=renumber.get(previous['case_id'], previous['case_id']))
        if current is None or batch_content(current) != expected or any(current.get(key) != previous.get(key) for key in ('cod_no', 'cod_date', 'issued')):
            raise ValueError('The updated SNs could not be verified. The original Excel file was kept.')
    return updated, next(row for row in after if row['source_sheet'] == added['source_sheet'] and row['case_id'] == final_sn), changes


def resequence_date_ordered_inwards(data, original_data, *, root, plant):
    """Reassign numeric SNs when their date order no longer matches their number order."""
    from certificate_core import read_batches
    rows = read_batches(BytesIO(data))
    numeric = [row for row in rows if re.fullmatch(r'\d{1,12}', row['case_id'])]
    numbers = [int(row['case_id']) for row in numeric]
    if numbers == sorted(numbers):
        return data, {}
    updated, _, changes = resequence_added_inward(
        data, numeric[0], original_data, root=root, plant=plant, request_id=None)
    return updated, changes


def resequence_removed_inwards(data, original_data, *, root, plant):
    """Close numeric SN gaps after deletion, keeping issued records and HDD links attached."""
    from certificate_core import read_batches
    from workflow_support import load_settings

    remaining = read_batches(BytesIO(data))
    remaining_ids = {row['case_id'] for row in remaining}
    removed_ids = {row['case_id'] for row in read_batches(BytesIO(original_data))} - remaining_ids
    saved = read_details(BytesIO(data))
    links = links_from(data)
    additions = links.setdefault('additions', {})
    path = Path(load_settings(root)['workbook_paths'][int(plant[-1]) - 1]).expanduser().resolve()
    changed = False
    for journal_path in (Path(root) / 'Data' / 'Client Additions').glob('*.json'):
        journal = json.loads(journal_path.read_text(encoding='utf-8'))
        request_id = journal_path.stem
        current = additions.get(request_id, journal.get('sn'))
        if journal.get('plant') == plant and journal.get('path') == str(path) and current in removed_ids:
            # The number can immediately be assigned to the next inward.
            additions[request_id] = None
            changed = True
    if changed:
        saved[LINK_KEY] = links
        data = _replace_details(data, saved)
    if not remaining:
        return data, {}
    updated, _, changes = resequence_added_inward(
        data, remaining[0], original_data, root=root, plant=plant, request_id=None)
    return updated, changes
