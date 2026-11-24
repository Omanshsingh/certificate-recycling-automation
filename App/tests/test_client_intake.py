"""Client intake tests use temporary workbooks, including copies of plant layouts."""
from datetime import date
from copy import copy
from io import BytesIO
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter

from certificate_core import batch_content, read_batches, select_register
from client_intake import stage_new_client
from workbook_store import WorkbookSession, _element, _sheet_member
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_workflow import make_workbook
import web_app

APP = Path(__file__).resolve().parents[1]


def column_for(sheet, heading):
    return next(cell.column for cell in sheet[1] if cell.value == heading)


class ClientIntakeTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='client-intake-')))
        self.paths = [self.root / 'Plant 1.xlsx', self.root / 'Plant 2.xlsx']
        for index,path in enumerate(self.paths, 1):
            make_workbook(path, [{'sn':str(index * 1000 + 1), 'supplier':'Existing client', 'cod':'0011', 'status':'Generated'}])
        save_json(self.root / 'settings.local.json', {'workbook_paths':[str(path) for path in self.paths], 'reminders_enabled':False})
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'Data' / 'Workbook Backups'))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.root / 'Data' / 'Reports' / 'Latest Results.json'))
        self.enterContext(patch.object(web_app, 'refresh', return_value={'rows':[], 'scan_complete':True}))

    def body(self, **details):
        return {'plant':'plant1', 'request_id':'a' * 32, 'expected_revision':workspace_revision(self.root),
                'details':{'supplier':'New client', 'form6':'FORM-NEW', **details}}

    def test_certificate_scope_creates_only_the_selected_requests(self):
        from certificate_drafts import read_register

        for index, (scope, expected_status) in enumerate(
                [('recycling', 'Ready'), ('destruction', 'N/A'), ('both', 'Ready')], start=1):
            body = dict(self.body(sn=str(1001 + index), supplier=f'Client {index}',
                                  form6=f'FORM-{index}', reference=f'REF-{index}',
                                  inward_date='2026-10-01', items=[{'item':'Laptop','weight':5}]),
                        certificate_scope=scope, request_id=f'{index:032x}')
            result = web_app.add_client(body)
            self.assertEqual(result['client']['certificate_scope'], scope)
            self.assertEqual(web_app.add_client(body)['client'], result['client'])
            sn = result['client']['sn']
            batch = next(row for row in read_batches(self.paths[0]) if row['case_id'] == sn)
            self.assertEqual(batch['issued'], expected_status)
            self.assertEqual((batch['cod_no'], batch['hdd_cod_no']), ('', ''))
            records = [row for row in read_register(self.root)['records']
                       if row['details'].get('destruction_source') == {'plant':'plant1', 'case_id':sn}]
            self.assertEqual(len(records), int(scope != 'recycling'))
            if records:
                hdd = next(row for row in result['state']['rows'] if row.get('record_id') == records[0]['id'])
                self.assertEqual(hdd['certificate_type'], 'destruction')
            if scope == 'destruction':
                response = web_app.update_inward_status({'plant':'plant1', 'sn':sn, 'action':'undo_na',
                                                         'expected_revision':workspace_revision(self.root)})
                self.assertEqual(response['status'], 'Ready')
                self.assertEqual(read_register(self.root)['records'][0]['id'], records[0]['id'])

    def test_retry_finishes_an_interrupted_hdd_only_add_without_another_inward(self):
        from certificate_drafts import read_register

        body = dict(self.body(), certificate_scope='destruction')
        with patch('destruction_certificate._save_linked_request_locked', side_effect=OSError('Local register is busy')):
            with self.assertRaisesRegex(ValueError, 'client is in Excel.*again to finish'):
                web_app.add_client(body)
        before = self.paths[0].read_bytes()
        result = web_app.add_client(body)
        self.assertEqual(len(read_batches(self.paths[0])), 2)
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(result['client']['status'], 'na')
        self.assertEqual(len(read_register(self.root)['records']), 1)

    def test_medical_intake_can_queue_a_separate_hdd_request_on_retry(self):
        from certificate_drafts import read_register

        body = dict(self.body(form6='', reference='NEW-REF', inward_date='2026-10-01',
                              items=[{'item':'Equipment', 'weight':10}]),
                    certificate_scope='siemens', data_destruction_needed=True)
        result = web_app.add_client(body)
        self.assertEqual(web_app.add_client(body)['client'], result['client'])
        sn = result['client']['sn']
        self.assertEqual(len(read_batches(self.paths[0])), 2)
        inward = next(row for row in read_batches(self.paths[0]) if row['case_id'] == sn)
        self.assertEqual((inward['certificate_type'], inward['issued'], inward['cod_no']), ('siemens', 'Ready', ''))
        hdd, = read_register(self.root)['records']
        self.assertEqual(hdd['details']['destruction_source'], {'plant':'plant1', 'case_id':sn})
        self.assertEqual((hdd['details']['certificate_type'], hdd['details']['cod_no']), ('destruction', ''))

    def test_inward_can_be_saved_before_form6_is_entered(self):
        body = self.body(form6='', reference='NEW-REFERENCE', inward_date='2026-10-01',
                         items=[{'item':'Asset','weight':5}])
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (body['expected_revision'], {'rows': [], 'scan_complete': True})):
            result = web_app.add_client(body)
        added = next(row for row in read_batches(self.paths[0])
                     if row['case_id'] == result['client']['sn'])
        self.assertEqual(added['form6'], '')
        self.assertEqual(added['issued'], 'Ready')
        row = web_app.target_row(result['state'], ('plant1', added['case_id']))
        self.assertEqual(row['outcome'], 'needs_review')
        self.assertIn('Form 6 number is required to create the certificate.', row['warnings'])

    def test_backdated_client_can_use_chronological_sr_and_excel_edits_remain_discoverable(self):
        result = web_app.add_client(self.body(sn='1000', inward_date='2026-06-15',
            form6='FORM-BACKDATED', items=[{'item':'Laptop','quantity':'1','weight':'4.5'}]))
        self.assertEqual(result['client']['sn'], '1000')
        batch = next(row for row in read_batches(self.paths[0]) if row['case_id'] == '1000')
        self.assertEqual(batch['inward_date'], date(2026, 6, 15))
        self.assertEqual(batch['issued'], 'Ready')

        # Model the user later changing the SR/date directly in Excel. Refresh
        # parsing must continue to identify the inward by its current SR.
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.cell(batch['source_row'], 1, '997')
        sheet.cell(batch['source_row'], 2, date(2026, 6, 12))
        book.save(self.paths[0]); book.close()
        edited = next(row for row in read_batches(self.paths[0]) if row['case_id'] == '997')
        self.assertEqual(edited['inward_date'], date(2026, 6, 12))
        self.assertEqual(edited['supplier'], 'New client')

    def test_add_client_updates_the_checked_dashboard_without_a_full_refresh(self):
        body = self.body(inward_date='2026-10-01', items=[{'item':'Laptop','weight':5}])
        expected = body['expected_revision']
        state = {'scan_complete':True, 'rows':[{'source_kind':'workbook','plant':'plant1','sn':'1001',
                'supplier':'Existing client','certificate_status':'generated'}]}
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (expected, state)), \
                patch.object(web_app, 'refresh', side_effect=AssertionError('A full refresh should not run after adding a client')):
            result = web_app.add_client(body)
        self.assertTrue(result['client_added'])
        self.assertEqual(web_app.target_row(result['state'], ('plant1','1001'))['certificate_status'], 'generated')
        added = web_app.target_row(result['state'], ('plant1',result['client']['sn']))
        self.assertEqual(added['supplier'], 'New client')
        self.assertEqual(added['certificate_status'], 'ready')
        self.assertEqual(added['outcome'], 'needs_review')
        self.assertIn('Reference number is required to create the certificate.', added['warnings'])

    def test_duplicate_custom_sr_is_rejected_without_modifying_workbook(self):
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'already used'):
            web_app.add_client(self.body(sn='1001'))
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_multiple_materials_ready_and_old_rows_and_package_parts_preserved(self):
        details = {'inward_date':'2026-10-01', 'form6':'FORM-NEW', 'location':'Pune, Maharashtra',
                   'items':[{'item':'Laptop','quantity':'2','per_piece_weight':'3.5'}, {'item':'Printer','weight':'8'}]}
        original = self.paths[0].read_bytes()
        original_batches = read_batches(BytesIO(original))
        web_app.add_client(self.body(**details))
        current_batches = read_batches(self.paths[0])
        batch = next(row for row in current_batches if row['supplier'] == 'New client')
        self.assertEqual(batch['issued'], 'Ready'); self.assertEqual(batch['total_weight'], 15)
        self.assertEqual(len(batch['items']), 2)
        self.assertCountEqual([batch_content(row) for row in original_batches],
                              [batch_content(row) for row in current_batches if row is not batch])
        for previous in original_batches:
            current = next(row for row in current_batches if row['case_id'] == previous['case_id'])
            for key in ('cod_no', 'cod_date', 'issued'):
                self.assertEqual(current[key], previous[key])
        with ZipFile(BytesIO(original)) as source, ZipFile(self.paths[0]) as after:
            member = _sheet_member(source, 'Inward Register')
            # Date ordering can move complete sections and standardize their
            # separators; unrelated XLSX parts must remain byte-for-byte intact.
            for name in source.namelist():
                if name not in {member,'xl/styles.xml','xl/workbook.xml'}:
                    self.assertEqual(source.read(name), after.read(name), name)
        book = load_workbook(self.paths[0])
        try:
            sheet, _, cols = select_register(book)
            for key in ('supplier', 'location', 'item', 'comments', 'transporter_details',
                        'docket_vehicle_no', 'weight_slip_scale'):
                if key not in cols:
                    continue
                column = cols[key] + 1
                anchor = sheet.cell(batch['source_row'], column)
                continuation = sheet.cell(batch['source_row'] + 1, column)
                self.assertTrue(anchor.alignment.wrap_text, key)
                if key == 'item':
                    self.assertTrue(continuation.alignment.wrap_text, key)
                else:
                    self.assertIn(f'{anchor.coordinate}:{continuation.coordinate}',
                                  {str(area) for area in sheet.merged_cells.ranges}, key)
        finally:
            book.close()

    def test_new_material_blocks_show_cod_lines_even_when_source_layout_was_merged(self):
        book = load_workbook(self.paths[0]); sheet, _, cols = select_register(book)
        sheet.cell(3, cols['item'] + 1, 'Older second asset')
        sheet.cell(3, cols['weight'] + 1, 2)
        for key in ('cod_no', 'cod_date'):
            col = cols[key] + 1
            sheet.merge_cells(start_row=2, end_row=3, start_column=col, end_column=col)
        book.save(self.paths[0]); book.close()
        web_app.add_client(self.body(items=[{'item':f'New asset {i}', 'weight':i} for i in (1, 2, 3)]))
        batches = read_batches(self.paths[0])
        book = load_workbook(self.paths[0]); sheet, _, cols = select_register(book)
        added = next(b for b in batches if b['supplier'] == 'New client')
        for key in ('cod_no', 'cod_date'):
            col = cols[key] + 1
            self.assertFalse(any(a.min_col <= col <= a.max_col for a in sheet.merged_cells.ranges))
            for row in range(added['source_row'], added['source_row'] + 2):
                self.assertEqual(sheet.cell(row, col).border.bottom.style, 'thin')
                self.assertEqual(sheet.cell(row + 1, col).border.top.style, 'thin')
        book.close()

    def test_layout_merges_totals_and_dropdowns_are_cloned_without_previous_values(self):
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.append([None,None,None,None,'Battery',1,2,'=F3*G3'])
        for column in ('A','B','C','D','I','M'):
            sheet.merge_cells(f'{column}2:{column}3')
        sheet['N1'] = 'Total Weight in KG'; sheet['N2'] = '=SUM(H2:H3)'; sheet.merge_cells('N2:N3')
        sheet['O1'] = 'Comments'; sheet['O2'] = 'Previous private note'; sheet.merge_cells('O2:O3')
        for cell in sheet[2]:
            cell.fill = PatternFill('solid', fgColor='EAF4EE'); cell.alignment = Alignment(vertical='center')
        sheet.row_dimensions[2].height = 32
        validation = DataValidation(type='list', formula1='"Mumbai,Pune"'); validation.add('D2:D3'); sheet.add_data_validation(validation)
        book.save(self.paths[0]); book.close()
        web_app.add_client(self.body(items=[{'item':'Laptop','weight':5},{'item':'Printer','quantity':2,'per_piece_weight':3},{'item':'UPS','weight':2}]))
        book = load_workbook(self.paths[0]); sheet = book.active
        self.assertIn('A4:A6', {str(area) for area in sheet.merged_cells.ranges})
        headers = {cell.value:cell.column for cell in sheet[1] if cell.value}
        total_col = headers['Total Weight in KG']
        total_letter = get_column_letter(total_col)
        self.assertIn(f'{total_letter}4:{total_letter}6', {str(area) for area in sheet.merged_cells.ranges})
        self.assertEqual(sheet.cell(4, total_col).value, '=SUM(H4:H6)')
        self.assertEqual(sheet['H5'].value, '=F5*G5')
        self.assertEqual(sheet['C4'].alignment.horizontal, 'left')
        self.assertEqual(sheet['E4'].alignment.horizontal, 'left')
        self.assertEqual(sheet['H5'].alignment.horizontal, 'left')
        self.assertIsNone(sheet.cell(4, headers['Weight Slip/Weight Scale']).value)
        self.assertIsNone(sheet.cell(4, headers['Comments']).value)
        self.assertEqual(copy(sheet['C4'].fill), copy(sheet['C2'].fill))
        self.assertEqual(sheet.row_dimensions[4].height, 32)
        self.assertTrue(any('D4' in rule for rule in sheet.data_validations.dataValidation))
        status_letter = get_column_letter(headers['Certificate Issued'])
        self.assertTrue(any(f'{status_letter}6' in rule for rule in sheet.data_validations.dataValidation))
        book.close()

    def test_retry_after_lost_response_does_not_add_twice_even_with_stale_revision(self):
        body = self.body(items=[{'item':'Laptop','weight':5}])
        first = web_app.add_client(body)
        before = self.paths[0].read_bytes()
        second = web_app.add_client(body)
        self.assertEqual(first['client'], second['client'])
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(len(read_batches(self.paths[0])), 2)
        self.assertEqual(len(list(web_app.BACKUP_DIR.glob('*.xlsx'))), 1)

    def test_deleted_client_retry_does_not_return_the_reused_serial(self):
        from inward_removal import remove_inward

        body = self.body(inward_date='2026-10-01', items=[{'item':'Laptop','weight':5}])
        sn = web_app.add_client(body)['client']['sn']
        web_app.add_client(dict(self.body(supplier='Next client', inward_date='2026-10-02',
                                          form6='FORM-NEXT', items=[{'item':'Printer','weight':7}]),
                                request_id='b' * 32))
        row = next(batch for batch in read_batches(self.paths[0]) if batch['case_id'] == sn)
        removal = {'plant':'plant1', 'sn':sn, 'sheet':row['source_sheet'], 'row':row['source_row'],
                   'request_id':'c' * 32, 'expected_revision':workspace_revision(self.root)}
        preview = remove_inward(self.root, removal)
        remove_inward(self.root, dict(removal, review_token=preview['review_token'], checked=True), commit=True)
        replacement = next(batch for batch in read_batches(self.paths[0]) if batch['case_id'] == sn)
        self.assertEqual(replacement['supplier'], 'Next client')
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(web_app.ClientAddRequestEnded, 'later removed'):
            web_app.add_client(body)
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_quantity_only_draft_leaves_weight_cell_blank(self):
        result = web_app.add_client(self.body(items=[{'item':'Laptop','quantity':'2'}]))
        row = next(batch for batch in read_batches(self.paths[0])
                   if batch['case_id'] == result['client']['sn'])
        book = load_workbook(self.paths[0])
        try:
            sheet, _, cols = select_register(book)
            self.assertIsNone(sheet.cell(row['source_row'], cols['weight'] + 1).value)
        finally:
            book.close()

    def test_prepared_journal_recovers_if_final_journal_write_failed(self):
        real_save = web_app.save_json
        def fail_saved(path, data):
            if data.get('saved'):
                raise OSError('Simulated storage failure')
            return real_save(path, data)
        body = self.body()
        with patch.object(web_app, 'save_json', side_effect=fail_saved):
            first = web_app.add_client(body)
        second = web_app.add_client(body)
        self.assertEqual(first['client']['sn'], second['client']['sn'])
        self.assertEqual(len(read_batches(self.paths[0])), 2)

    def test_changed_request_is_rejected_and_stale_new_request_makes_no_changes(self):
        body = self.body(); web_app.add_client(body)
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'different details'):
            web_app.add_client(dict(body, details={'supplier':'Different', 'form6':'FORM-NEW'}))
        with self.assertRaisesRegex(ValueError, 'files changed'):
            web_app.add_client(dict(body, request_id='b' * 32))
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_validation_and_protected_sheets_never_write(self):
        before = self.paths[0].read_bytes()
        for fields in ({'supplier':''}, {'inward_date':'bad'}, {'items':[{'weight':'bad'}]},
                       {'items':[{'quantity':'-1'}]}, {'items':[{'weight':'Infinity'}]}):
            with self.subTest(fields=fields), self.assertRaises(ValueError):
                web_app.add_client(self.body(**fields))
            self.assertEqual(self.paths[0].read_bytes(), before)
        book = load_workbook(self.paths[0]); book.active.protection.sheet = True; book.save(self.paths[0]); book.close()
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'protected'):
            web_app.add_client(self.body())
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertFalse(list(web_app.BACKUP_DIR.glob('*.xlsx')))

    def test_table_range_and_filter_expand(self):
        book = load_workbook(self.paths[0]); sheet = book.active
        table = Table(displayName='Clients', ref='A1:M2'); table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9')
        sheet.add_table(table); book.save(self.paths[0]); book.close()
        web_app.add_client(self.body(items=[{'item':'Printer','weight':5},{'item':'Laptop','weight':3}]))
        book = load_workbook(self.paths[0]); table = book.active.tables['Clients']
        last_column = get_column_letter(book.active.max_column)
        self.assertEqual(table.ref, f'A1:{last_column}4')
        self.assertEqual(table.autoFilter.ref, f'A1:{last_column}4')
        book.close()

    def test_missing_optional_details_are_saved_without_extra_columns(self):
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.delete_cols(9,5); book.save(self.paths[0]); book.close()
        before = [batch_content(batch) for batch in read_batches(self.paths[0])]
        web_app.add_client(self.body(form6='FORM-NEW',reference='REF-NEW',location='Pune'))
        after = read_batches(self.paths[0])
        self.assertEqual([batch_content(batch) for batch in after[:-1]], before)
        self.assertEqual(after[-1]['form6'],'FORM-NEW'); self.assertEqual(after[-1]['reference'],'REF-NEW')
        self.assertEqual(after[-1]['missing_headers'], [])
        book = load_workbook(self.paths[0])
        headers = [cell.value for cell in book.active[1]]
        self.assertEqual(headers[-3:], ['COD No', 'COD Date', 'Certificate Issued'])
        self.assertEqual(len(headers), len(set(headers)))
        book.close()

    def test_four_operational_details_are_saved_and_missing_columns_are_added(self):
        details = {
            'comments':'Collection completed',
            'transporter_details':'ABC Logistics / KA-01-AA-1234',
            'docket_vehicle_no':'DCK-204',
            'weight_slip_scale':'Scale ticket 8891',
        }
        web_app.add_client(self.body(**details, items=[{'item':'Laptop','weight':5}]))
        book = load_workbook(self.paths[0]); sheet = book.active
        headers = {sheet.cell(1, column).value:column for column in range(1, sheet.max_column + 1)}
        for label, value in zip(('Comments','Transporter/Vehicle Details','Docket No./Vehicle No','Weight Slip/Weight Scale'), details.values()):
            self.assertEqual(sheet.cell(3, headers[label]).value, value)
        book.close()

    def test_text_quantity_does_not_multiply_but_numeric_quantity_does(self):
        web_app.add_client(self.body(inward_date='2026-10-01', form6='FORM-NEW', items=[
            {'item':'Lot of cables','quantity':'1 lot','per_piece_weight':'3.5'},
            {'item':'Laptop','quantity':'2','per_piece_weight':'3.5'},
        ]))
        book = load_workbook(self.paths[0], data_only=False); sheet = book.active
        self.assertEqual(sheet['F3'].value, '1 lot')
        self.assertEqual(sheet['G3'].value, 3.5)
        self.assertEqual(sheet['H3'].value, None)
        self.assertEqual(sheet['F4'].value, 2)
        self.assertEqual(sheet['H4'].value, '=F4*G4')
        book.close()
        batch = read_batches(self.paths[0])[-1]
        self.assertEqual(batch['items'][0]['quantity'], '1 lot')
        self.assertIsNone(batch['items'][0]['weight'])
        self.assertEqual(batch['items'][1]['weight'], 7)

    def test_plant2_and_text_ids_get_an_automatic_unique_id(self):
        book = load_workbook(self.paths[1]); book.active['A2'] = 'CLIENT-2001'; book.save(self.paths[1]); book.close()
        other = self.paths[0].read_bytes()
        body = self.body(); body['plant'] = 'plant2'
        result = web_app.add_client(body)
        self.assertEqual(result['client']['sn'], '1')
        self.assertEqual(self.paths[0].read_bytes(), other)

    def test_real_plant_layout_copies_support_new_blocks(self):
        for index in (1,2):
            source = APP.parent / 'Plant Workbooks' / f'Plant {index}.xlsx'
            if not source.is_file():
                continue
            with self.subTest(plant=index):
                target = self.root / f'layout{index}.xlsx'; shutil.copy2(source,target)
                original = read_batches(target)
                with WorkbookSession(target, self.root/'layout-backups') as session:
                    updated, batch = stage_new_client(session, {'supplier':'Sandbox new client', 'inward_date':'2026-10-01',
                        'form6':'SANDBOX-F6', 'items':[{'item':'Laptop','quantity':2,'per_piece_weight':3.5},{'item':'Printer','weight':8}]})
                    session._persist(updated)
                after = read_batches(target)
                self.assertCountEqual([batch_content(row) for row in original],
                                      [batch_content(row) for row in after if row['case_id'] != batch['case_id']])
                self.assertEqual(batch['total_weight'],15); self.assertEqual(batch['issued'],'Ready')

    def test_older_add_requests_default_to_recycling_without_a_certificate_choice(self):
        body = self.body(certificate_type='sita', inward_date='2026-10-01', items=[{'item':'Asset','weight':10}])
        result = web_app.add_client(body)
        self.assertEqual(read_batches(self.paths[0])[-1]['certificate_type'],'standard')
        self.assertEqual(read_batches(self.paths[0])[-1]['form6'],'FORM-NEW')
        body = self.body(certificate_type='siemens', inward_date='2026-10-01', form6='FORM-NEW-2',
                         items=[{'item':'Asset','weight':10}])
        body['request_id'] = 'e' * 32
        web_app.add_client(body)
        self.assertEqual(read_batches(self.paths[0])[-1]['certificate_type'],'standard')
        self.assertEqual(read_batches(self.paths[0])[-1]['form6'],'FORM-NEW-2')

    def test_medical_and_international_choices_survive_workbook_reload_without_creating_hdd(self):
        from certificate_drafts import read_register

        book = load_workbook(self.paths[1])
        sheet = book.active
        for heading in ('Certificate Type', 'Certificate Weight (KG)', 'Certificate Fields'):
            sheet.cell(1, sheet.max_column + 1, heading)
        book.save(self.paths[1])
        book.close()
        for index, kind in enumerate(('siemens', 'sita')):
            plant = f'plant{index + 1}'
            fields = {'model_name': 'Scanner'} if kind == 'siemens' else {'company_display': 'International Client'}
            body = dict(self.body(supplier=f'New {kind} client', form6='', reference='NEW-REF',
                                  certificate_weight='12.5', certificate_fields=fields,
                                  inward_date='2026-10-01', items=[{'item':'Equipment', 'weight':10}]),
                        plant=plant, certificate_scope=kind, request_id=f'{index + 1:032x}')
            cached = (body['expected_revision'], {'rows': [], 'scan_complete': True})
            with patch.object(web_app, 'WORKSPACE_STATE_CACHE', cached):
                result = web_app.add_client(body)
                self.assertEqual(web_app.add_client(body)['client'], result['client'])
            sn = result['client']['sn']
            batches = read_batches(self.paths[index])
            batch = next(row for row in batches if row['case_id'] == sn)
            self.assertEqual(len(batches), 2)
            self.assertEqual(batch['certificate_type'], kind)
            self.assertEqual(batch['certificate_fields'], fields)
            self.assertEqual(batch['certificate_weight'], 12.5)
            self.assertEqual((batch['form6'], batch['cod_no'], batch['hdd_cod_no']), ('', '', ''))
            shown = web_app.target_row(result['state'], (plant, sn))
            self.assertEqual(shown['certificate_type'], kind)
            self.assertEqual(shown['outcome'], 'needs_review' if kind == 'siemens' else 'ready')
        self.assertEqual(read_register(self.root)['records'], [])

    def test_book_my_junk_client_is_added_as_na_and_cannot_be_issued(self):
        result = web_app.add_client(self.body(
            supplier='Reverse E - Commerce Pvt. Ltd. (Book My Junk)',
            inward_date='2026-10-01', items=[{'item':'Asset','weight':10}],
        ))
        added = read_batches(self.paths[0])[-1]
        self.assertEqual(added['issued'], 'N/A')
        book = load_workbook(self.paths[0])
        try:
            status_column = column_for(book.active, 'Certificate Issued')
            self.assertEqual(book.active.cell(3, status_column).value, 'N/A')
        finally:
            book.close()
        details = web_app.editor_details(added)
        with self.assertRaisesRegex(ValueError, 'marked N/A'):
            web_app.create_certificate({
                'plant':'plant1', 'source_sn':added['case_id'], 'source_plant':'plant1',
                'details':details,
            })

    def test_table_keeps_its_columns_and_adds_standard_client(self):
        book = load_workbook(self.paths[0]); book.active.add_table(Table(displayName='Clients',ref='A1:M2'))
        book.save(self.paths[0]); book.close()
        web_app.add_client(self.body(certificate_type='sita', comments='Pickup arranged',
                                     transporter_details='ABC Logistics', docket_vehicle_no='DCK-42',
                                     weight_slip_scale='Scale 778'))
        book = load_workbook(self.paths[0]); table = book.active.tables['Clients']
        self.assertEqual(table.ref,'A1:Q3')
        self.assertEqual([column.name for column in table.tableColumns[-4:]],
                         ['Weight Slip/Weight Scale','COD No','COD Date','Certificate Issued'])
        self.assertEqual([column.name for column in table.tableColumns[10:14]],
                         ['Comments','Transporter/Vehicle Details','Docket No./Vehicle No','Weight Slip/Weight Scale'])
        self.assertEqual(read_batches(self.paths[0])[-1]['certificate_type'], 'standard')
        book.close()

    def test_form6_can_be_entered_when_creating_a_certificate(self):
        book = load_workbook(self.paths[0]); book.active['I2'] = ''; book.save(self.paths[0]); book.close()
        source = read_batches(self.paths[0])[0]
        from certificate_core import source_value_issues
        self.assertEqual(source['form6'], '')
        self.assertNotIn('missing Form 6 number / manifest number',
                         source_value_issues(source, certificate_type='sita'))
        self.assertNotIn('missing Form 6 number / manifest number', source['blocking_issues'])

    def test_future_date_keeps_blank_status_and_text_is_not_an_excel_formula(self):
        web_app.add_client(self.body(supplier='=HYPERLINK("https://example.invalid")', inward_date='2099-01-01',
                                    form6='FORM-NEW', items=[{'item':'Asset','weight':10}]))
        book = load_workbook(self.paths[0]); self.assertEqual(book.active['C3'].data_type,'s')
        status_column = column_for(book.active, 'Certificate Issued')
        self.assertIsNone(book.active.cell(3, status_column).value); book.close()

    def test_locked_excel_and_backup_failure_keep_original_file(self):
        from workbook_store import exclusive_file
        body = self.body(); before = self.paths[0].read_bytes()
        with exclusive_file(self.paths[0]):
            with self.assertRaises((OSError, ValueError)):
                web_app.add_client(body)
        self.assertEqual(self.paths[0].read_bytes(),before)
        with patch.object(WorkbookSession,'_persist',side_effect=OSError('Backup is unavailable')):
            with self.assertRaises(OSError):
                web_app.add_client(body)
        self.assertEqual(self.paths[0].read_bytes(),before)
        result = web_app.add_client(body)
        self.assertEqual(result['client']['sn'],'1002')

    def test_stale_external_change_at_save_boundary_is_rejected(self):
        before = self.paths[0].read_bytes(); body = self.body()
        original_stage = web_app.stage_new_client
        def stage_then_other_change(session, details, **kwargs):
            staged = original_stage(session, details, **kwargs)
            book = load_workbook(self.paths[1]); book.active['C2'] = 'External change'; book.save(self.paths[1]); book.close()
            return staged
        with patch.object(web_app,'stage_new_client',side_effect=stage_then_other_change):
            with self.assertRaisesRegex(ValueError,'files changed'):
                web_app.add_client(body)
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_newly_added_formula_block_can_generate_a_certificate(self):
        shutil.copytree(APP/'templates', self.root/'App'/'templates')
        result = web_app.add_client(self.body(inward_date='2026-10-01', form6='NEW-F6', reference='NEW-REF',
                        items=[{'item':'Laptop','quantity':2,'per_piece_weight':3.5},{'item':'Printer','weight':8}]))
        from run_certificates import run
        created,_,_ = run(self.paths,self.root,selected=('plant1',result['client']['sn']),progress=None)
        self.assertEqual(created,1)
        added = read_batches(self.paths[0])[-1]
        self.assertEqual(added['total_weight'],15); self.assertEqual(added['issued'],'Generated')
        self.assertEqual(len(list((self.root/'Certificates').rglob('*.pdf'))),1)


if __name__ == '__main__':
    unittest.main()
