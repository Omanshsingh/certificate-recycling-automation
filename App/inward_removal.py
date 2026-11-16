"""Reviewed removal of an unissued inward, closing its row gap safely."""
from copy import copy
from datetime import datetime, timezone
from io import BytesIO
import hashlib
import hmac
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries

from register_protection import externally_protected, protect_issued_inwards
from certificate_core import batch_content, read_batches, select_register, status_value, text
from certificate_inventory import scan_inventory
from certificate_drafts import read_register
from issuance_safety import audit_records
from workbook_metadata import remove_details
from workbook_store import WorkbookSession, _attributes, _element, _sheet_member
from workflow_support import load_settings, run_lock, save_json
from workspace_review import require_revision, review_token


def _key(value):
    return str(value or '').strip().casefold()


def stage_removal(data, sn, sheet_name, source_row, *, certificate_conversion=False):
    """Remove this complete record block, retaining package parts and remaining facts."""
    batches = read_batches(BytesIO(data))
    matches = [b for b in batches if _key(b['case_id']) == _key(sn)]
    if len(matches) != 1:
        raise ValueError('This SN is missing or appears more than once. Check Excel and refresh.')
    batch = matches[0]
    if batch['source_sheet'] != sheet_name or batch['source_row'] != source_row:
        raise ValueError('This inward moved in Excel. Refresh and check the entry again.')
    if not certificate_conversion and status_value(batch.get('issued')) in {'generated', 'generated_manual', 'completed'}:
        raise ValueError('This inward already has certificate details, so it cannot be removed here.')
    if any('total row needs checking' in issue for issue in batch.get('blocking_issues', [])):
        raise ValueError('Check the total rows in Excel before removing this inward.')
    book = load_workbook(BytesIO(data), data_only=False)
    try:
        sheet, header, cols = select_register(book)
        if externally_protected(sheet):
            raise ValueError('This Excel sheet is protected. Unprotect it before removing an inward.')
        if any(other._pivots for other in book) or any(table.totalsRowCount or any(
                column.calculatedColumnFormula for column in table.tableColumns) for table in sheet.tables.values()):
            raise ValueError('This workbook has table calculations. Remove the entry in Excel instead.')
        first = batch['source_row']
        if sheet.cell(first, cols['sn'] + 1).data_type == 'f':
            raise ValueError('This SN uses an Excel formula. Check the entry in Excel first.')
        limit = min([b['source_row'] for b in batches if b['source_row'] > first] or [sheet.max_row + 1])
        material_cols = [cols[k] + 1 for k in ('item', 'quantity', 'per_piece_weight', 'weight') if k in cols]
        last = max([first, *batch.get('summary_rows', [])] + [r for r in range(first, limit)
                   if any(sheet.cell(r, c).value is not None for c in material_cols)])
        # Include all merged detail edges, while rejecting a merge shared by two entries.
        for area in sorted(sheet.merged_cells.ranges, key=lambda area: area.min_row):
            if area.min_row <= last and area.max_row >= first:
                if area.min_row < first or area.max_row >= limit:
                    raise ValueError('Two entries share merged cells. Separate them in Excel before removing one.')
                last = max(last, area.max_row)
        for r in ([] if certificate_conversion else range(first, last + 1)):
            if 'issued' in cols and status_value(sheet.cell(r, cols['issued'] + 1).value) in {'generated', 'generated_manual', 'completed'}:
                raise ValueError('This inward already has a certificate, so it cannot be removed here.')
        for r in range(first + 1, last + 1):
            if text(sheet.cell(r, cols['sn'] + 1).value) and r not in batch.get('summary_rows', []):
                raise ValueError('There is a note or another SN inside this entry. Check it in Excel first.')
        if (sheet._charts or sheet._images or any(sheet.cell(r, c).comment or sheet.cell(r, c).hyperlink
                for r in range(first, last + 1) for c in range(1, sheet.max_column + 1))):
            raise ValueError('This entry has an Excel object or note. Remove the entry in Excel instead.')
        if any(name not in {'_xlnm.Print_Area', '_xlnm.Print_Titles'}
               for name in list(book.defined_names) + list(sheet.defined_names)):
            raise ValueError('This workbook has named calculations. Remove the entry in Excel instead.')
        for other in book:
            for row in other:
                for cell in row:
                    if cell.data_type != 'f':
                        continue
                    if not isinstance(cell.value, str):
                        raise ValueError('This workbook has shared calculations. Remove the entry in Excel instead.')
                    if other == sheet and first <= cell.row <= last:
                        continue
                    # Unknown/dynamic references could depend on the removed entry.
                    if other != sheet:
                        raise ValueError('Another sheet has calculations. Remove the entry in Excel instead.')
                    for token in Tokenizer(cell.value).items:
                        if token.type == 'FUNC' and token.value.upper() in {'INDIRECT(', 'OFFSET('}:
                            raise ValueError('This workbook has linked calculations. Remove the entry in Excel instead.')
                        if token.type == 'OPERAND' and token.subtype == 'RANGE':
                            if not re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?', token.value, re.I):
                                raise ValueError('This workbook has linked calculations. Remove the entry in Excel instead.')
                            _, top, _, bottom = range_boundaries(token.value.replace('$', '').upper())
                            if top <= last and bottom >= first:
                                raise ValueError('Another calculation uses this inward. Remove the entry in Excel instead.')
        from register_rows import shift_register_rows
        updated = shift_register_rows(data, first, last - first + 1, delete=True)
        updated = remove_details(updated, sheet.title, batch['case_id'], original_data=data)
        remaining = read_batches(BytesIO(updated))
        expected = [b for b in batches if b is not batch]
        def facts(b):
            return (batch_content(b), *(b.get(k) for k in ('source_sheet', 'cod_no', 'cod_date', 'issued')))
        if (list(map(facts, remaining)) != list(map(facts, expected))
                or any(new['source_row'] != old['source_row'] - (last - first + 1 if old['source_row'] > last else 0)
                       for new, old in zip(remaining, expected))):
            raise ValueError('The other entries could not be verified. Excel has not been changed.')
        entry = {'plant': None, 'sn': batch['case_id'], 'supplier': batch['supplier'],
                 'sheet': sheet.title, 'row': first, 'last_row': last,
                 'inward_date': str(batch.get('inward_date') or ''), 'materials': len(batch['items'])}
        return updated, entry, batch
    finally:
        book.close()


def _prepare(root, body, data):
    plant, sn = body.get('plant'), body.get('sn')
    # A cleared COD/status must not hide a previously created certificate.
    if any(r.get('source_kind') == 'workbook' and r.get('plant') == plant and _key(r.get('key')) == _key(sn)
           and r.get('state') != 'rolled_back' for r in audit_records(root)):
        raise ValueError('A certificate was already created for this inward, so it cannot be removed here.')
    for pdf in scan_inventory(root, persist=False).get('entries', []):
        owner = pdf.get('managed_owner') or {}
        if owner.get('source_kind') == 'workbook' and owner.get('plant') == plant and _key(owner.get('key')) == _key(sn):
            raise ValueError('A certificate was already created for this inward, so it cannot be removed here.')
    for record in read_register(root)['records']:
        source = record.get('details', {}).get('destruction_source') or {}
        if not source.get('record_id') and source.get('plant') == plant and _key(source.get('case_id')) == _key(sn):
            raise ValueError('This inward has a data destruction request, so it cannot be removed here.')
    updated, entry, batch = stage_removal(data, sn, body.get('sheet'), body.get('row'))
    entry['plant'] = plant
    return updated, entry, batch


def remove_inward(root, body, *, commit=False):
    root = Path(root).resolve()
    plant = body.get('plant')
    request_id = body.get('request_id')
    if (plant not in {'plant1', 'plant2'} or not isinstance(body.get('sn'), str) or not body['sn'].strip()
            or not isinstance(body.get('sheet'), str) or type(body.get('row')) is not int
            or not re.fullmatch(r'[a-f0-9]{32}', str(request_id or ''))):
        raise ValueError('Reopen the inward and try again.')
    settings = load_settings(root)
    configured = settings['workbook_paths'][int(plant[-1]) - 1]
    if not configured:
        raise ValueError('Connect this plant Excel file in Settings first.')
    path = Path(configured).expanduser().resolve()
    other = settings['workbook_paths'][1 if plant == 'plant1' else 0]
    if other:
        other_path = Path(other).expanduser().resolve()
        if path == other_path or (path.is_file() and other_path.is_file() and path.samefile(other_path)):
            raise ValueError('Plant 1 and Plant 2 must use different Excel files.')
    identity = {key: body[key] for key in ('plant', 'sn', 'sheet', 'row', 'request_id')}
    journal_path = root / 'Data' / 'Inward Removals' / f'{request_id}.json'
    with run_lock(root):
        journal = json.loads(journal_path.read_text(encoding='utf-8')) if journal_path.is_file() else None
        if journal and (journal.get('identity') != identity or journal.get('path') != str(path)):
            raise ValueError('This removal request belongs to a different entry. Reopen the inward.')
        if journal and journal.get('undo_started'):
            raise ValueError('This removal was undone or is being restored. Review the inward again before removing it.')
        if commit and journal and journal.get('saved'):
            return {'removed': True, 'entry': journal['entry']}
        with WorkbookSession(path, root / 'Data' / 'Workbook Backups') as session:
            overrides = {str(path): session.data}
            if commit and journal and hashlib.sha256(session.data).hexdigest() == journal['after_sha256']:
                journal['saved'] = True
                try:
                    save_json(journal_path, journal)
                except OSError:
                    pass
                return {'removed': True, 'entry': journal['entry']}
            if commit and journal:
                raise ValueError('This removal attempt could not be confirmed. Keep the entry, then reopen Remove inward to check it again.')
            require_revision(root, body.get('expected_revision'), overrides)
            updated, entry, batch = _prepare(root, body, session.data)
            from inward_serials import resequence_removed_inwards
            updated, _ = resequence_removed_inwards(updated, session.data, root=root, plant=plant)
            token = review_token('/api/clients/remove', {'identity': identity, 'entry': entry}, body['expected_revision'], [])
            if not commit:
                return {'entry': entry, 'review_token': token, 'expected_revision': body['expected_revision']}
            if body.get('checked') is not True or not isinstance(body.get('review_token'), str) or not hmac.compare_digest(body['review_token'], token):
                raise ValueError('Check the entry in the removal message before choosing Remove entry.')
            updated = protect_issued_inwards(updated)
            journal = {'identity': identity, 'path': str(path), 'entry': entry,
                       'snapshot': json.loads(json.dumps(batch_content(batch), default=str)),
                       'at': datetime.now(timezone.utc).isoformat(), 'saved': False,
                       'before_sha256': hashlib.sha256(session.data).hexdigest(),
                       'after_sha256': hashlib.sha256(updated).hexdigest()}
            save_json(journal_path, journal)
            session.before_persist = lambda _: require_revision(root, body['expected_revision'], overrides)
            session._persist(updated, exact=True)  # Durable automatic backup precedes atomic replacement.
            journal.update(saved=True, backup_path=str(session.backup_path))
            try:
                save_json(journal_path, journal)
            except OSError:
                pass  # The prepared hashes prove a successful write after a lost response.
            return {'removed': True, 'entry': entry}


def undo_inward_removal(root, request_id, revision):
    """Restore one removed inward from its automatic workbook backup."""
    from issuance_safety import recover_issuances

    root = Path(root).resolve()
    if not isinstance(request_id, str) or not re.fullmatch(r'[a-f0-9]{32}', request_id):
        raise ValueError('Choose a saved deletion to undo.')
    journal_path = root / 'Data' / 'Inward Removals' / f'{request_id}.json'
    backups = root / 'Data' / 'Workbook Backups'
    with run_lock(root):
        recover_issuances(root)
        if not journal_path.is_file():
            raise ValueError('The inward removal record was not found.')
        journal = json.loads(journal_path.read_text(encoding='utf-8'))
        if journal.get('undone'):
            return {'undone':True}
        path = Path(journal['path']).resolve()
        plant = journal.get('entry', {}).get('plant')
        settings = load_settings(root)
        if plant not in {'plant1', 'plant2'} or path != Path(settings['workbook_paths'][int(plant[-1]) - 1]).expanduser().resolve():
            raise ValueError('Reconnect the original workbook to its plant before restoring this inward.')
        with WorkbookSession(path, backups) as session:
            overrides = {str(path):session.data}
            current = hashlib.sha256(session.data).hexdigest()
            if journal.get('undo_started') and current == journal['before_sha256']:
                journal['undone'] = True
                save_json(journal_path, journal)
                return {'undone':True}
            require_revision(root, revision, overrides)
            saved_backup = journal.get('backup_path')
            candidates = [Path(saved_backup)] if saved_backup else backups.glob(f'{path.stem}_*{path.suffix}')
            if current != journal['after_sha256']:
                location = str(saved_backup or backups)
                raise ValueError(f'{path.name} changed after the removal. The pre-deletion copy is at {location}. '
                                 'Open that copy in Excel and bring the missing inward into the current workbook.')
            for candidate in candidates:
                if not candidate.resolve().is_relative_to(backups.resolve()) or not candidate.is_file():
                    continue
                before = candidate.read_bytes()
                if hashlib.sha256(before).hexdigest() == journal['before_sha256']:
                    break
            else:
                raise ValueError('The verified copy from before the removal is unavailable. Check your workbook backups.')
            from identifier_rules import require_restored_identifiers
            require_restored_identifiers(root, [dict(journal['entry'], source_kind='workbook')],
                                         {str(path):before}, read_register(root)['records'])
            journal['undo_started'] = True
            save_json(journal_path, journal)
            session.before_persist = lambda _: require_revision(root, revision, overrides)
            session._persist(before, exact=True)
            journal['undone'] = True
            try:
                save_json(journal_path, journal)
            except OSError:
                pass  # The before hash proves the undo already committed.
        return {'undone':True}
