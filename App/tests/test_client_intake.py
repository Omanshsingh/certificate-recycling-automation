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

from certificate_core import batch_content, read_batches
from client_intake import stage_new_client
from workbook_store import WorkbookSession, _element, _sheet_member
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_workflow import make_workbook
import web_app

APP = Path(__file__).resolve().parents[1]


class ClientIntakeTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='client-intake-')))
        self.paths = [self.root / 'Plant 1.xlsx', self.root / 'Plant 2.xlsx']
        for index,path in enumerate(self.paths, 1):
            make_workbook(path, [{'sn':str(index * 1000 + 1), 'supplier':'Existing client', 'cod':'0011', 'status':'Generated'}])
        save_json(self.root / 'settings.local.json', {'workbook_paths':[str(path) for path in self.paths], 'reminders_enabled':False})
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'Data' / 'Workbook Backups'))
        self.enterContext(patch.object(web_app, 'refresh', return_value={'rows':[], 'scan_complete':True}))

    def body(self, **details):
        return {'plant':'plant1', 'request_id':'a' * 32, 'expected_revision':workspace_revision(self.root),
                'details':{'supplier':'New client', **details}}

    def test_company_only_creates_blank_status_and_can_be_completed_in_excel(self):
        before = self.paths[0].read_bytes()
        other = self.paths[1].read_bytes()
        result = web_app.add_client(self.body())
        self.assertEqual(result['client']['sn'], '1002')
        self.assertEqual(result['client']['status'], 'blank')
        book = load_workbook(self.paths[0]); sheet = book.active
        self.assertIsNone(sheet['M3'].value)
        self.assertIsNone(sheet['K3'].value); self.assertIsNone(sheet['L3'].value)
        self.assertEqual(sheet['H3'].value, '=F3*G3')
        sheet['B3'] = date(2026,10,1); sheet['E3'] = 'Laptop'; sheet['F3'] = 2; sheet['G3'] = 3; sheet['I3'] = 'FORM-NEW'
        book.save(self.paths[0]); book.close()
        from certificate_core import inward_details_complete
        self.assertTrue(inward_details_complete(read_batches(self.paths[0])[-1]))
        self.assertEqual(self.paths[1].read_bytes(), other)
        backups = list(web_app.BACKUP_DIR.glob('*.xlsx'))
        self.assertEqual(len(backups), 1); self.assertEqual(backups[0].read_bytes(), before)
        self.assertFalse(list(self.root.rglob('*.pdf')))

    def test_multiple_materials_ready_and_old_rows_and_package_parts_preserved(self):
        details = {'inward_date':'2026-10-01', 'form6':'FORM-NEW', 'location':'Pune, Maharashtra',
                   'items':[{'item':'Laptop','quantity':'2','per_piece_weight':'3.5'}, {'item':'Printer','weight':'8'}]}
        original = self.paths[0].read_bytes()
        web_app.add_client(self.body(**details))
        batch = read_batches(self.paths[0])[-1]
        self.assertEqual(batch['issued'], 'Ready'); self.assertEqual(batch['total_weight'], 15)
        self.assertEqual(len(batch['items']), 2)
        with ZipFile(BytesIO(original)) as source, ZipFile(self.paths[0]) as after:
            member = _sheet_member(source, 'Inward Register')
            old_xml = source.read(member).decode(); new_xml = after.read(member).decode()
            old_rows = list(_element(old_xml, 'row').finditer(old_xml))
            new_rows = list(_element(new_xml, 'row').finditer(new_xml))
            self.assertEqual([m.group() for m in old_rows], [m.group() for m in new_rows[:len(old_rows)]])
            for name in source.namelist():
                if name not in {member,'xl/styles.xml','xl/workbook.xml'}:
                    self.assertEqual(source.read(name), after.read(name), name)

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
        self.assertIn('N4:N6', {str(area) for area in sheet.merged_cells.ranges})
        self.assertEqual(sheet['N4'].value, '=SUM(H4:H6)')
        self.assertEqual(sheet['H5'].value, '=F5*G5')
        self.assertIsNone(sheet['O4'].value); self.assertIsNone(sheet['K4'].value)
        self.assertEqual(copy(sheet['C4'].fill), copy(sheet['C2'].fill))
        self.assertEqual(sheet.row_dimensions[4].height, 32)
        self.assertTrue(any('D4' in rule for rule in sheet.data_validations.dataValidation))
        self.assertTrue(any('M6' in rule for rule in sheet.data_validations.dataValidation))
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
            web_app.add_client(dict(body, details={'supplier':'Different'}))
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
        self.assertEqual(table.ref, 'A1:M4'); self.assertEqual(table.autoFilter.ref, 'A1:M4'); book.close()

    def test_missing_optional_details_are_saved_without_extra_columns(self):
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.delete_cols(9,5); book.save(self.paths[0]); book.close()
        before = [batch_content(batch) for batch in read_batches(self.paths[0])]
        web_app.add_client(self.body(form6='FORM-NEW',reference='REF-NEW',location='Pune'))
        after = read_batches(self.paths[0])
        self.assertEqual([batch_content(batch) for batch in after[:-1]], before)
        self.assertEqual(after[-1]['form6'],'FORM-NEW'); self.assertEqual(after[-1]['reference'],'REF-NEW')
        self.assertEqual(after[-1]['missing_headers'], ['cod_no','cod_date'])
        book = load_workbook(self.paths[0])
        self.assertEqual(book.active.max_column, 9)
        self.assertEqual(book.active.cell(1,9).value, 'Certificate Issued')
        book.close()

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
                self.assertEqual([batch_content(row) for row in original], [batch_content(row) for row in after[:-1]])
                self.assertEqual(batch['total_weight'],15); self.assertEqual(batch['issued'],'Ready')

    def test_special_types_need_no_form6_and_type_is_stored(self):
        for kind in ('sita','siemens'):
            with self.subTest(kind=kind):
                body = self.body(certificate_type=kind, inward_date='2026-10-01', items=[{'item':'Asset','weight':10}])
                body['request_id'] = ('c' if kind == 'sita' else 'd') * 32
                result = web_app.add_client(body)
                self.assertEqual(result['client']['status'],'ready')
                self.assertEqual(read_batches(self.paths[0])[-1]['certificate_type'],kind)

    def test_table_keeps_its_columns_and_saves_sita_type_in_properties(self):
        book = load_workbook(self.paths[0]); book.active.add_table(Table(displayName='Clients',ref='A1:M2'))
        book.save(self.paths[0]); book.close()
        web_app.add_client(self.body(certificate_type='sita'))
        book = load_workbook(self.paths[0]); table = book.active.tables['Clients']
        self.assertEqual(table.ref,'A1:M3'); self.assertEqual(table.tableColumns[-1].name,'Certificate Issued')
        self.assertEqual(read_batches(self.paths[0])[-1]['certificate_type'], 'sita')
        book.close()

    def test_future_date_keeps_blank_status_and_text_is_not_an_excel_formula(self):
        web_app.add_client(self.body(supplier='=HYPERLINK("https://example.invalid")', inward_date='2099-01-01',
                                    form6='FORM-NEW', items=[{'item':'Asset','weight':10}]))
        book = load_workbook(self.paths[0]); self.assertEqual(book.active['C3'].data_type,'s')
        self.assertIsNone(book.active['M3'].value); book.close()

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
        def stage_then_other_change(session, details):
            staged = original_stage(session,details)
            book = load_workbook(self.paths[1]); book.active['C2'] = 'External change'; book.save(self.paths[1]); book.close()
            return staged
        with patch.object(web_app,'stage_new_client',side_effect=stage_then_other_change):
            with self.assertRaisesRegex(ValueError,'files changed'):
                web_app.add_client(body)
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_newly_added_formula_block_can_generate_a_certificate(self):
        shutil.copytree(APP/'templates', self.root/'App'/'templates')
        result = web_app.add_client(self.body(inward_date='2026-10-01', form6='NEW-F6',
                        items=[{'item':'Laptop','quantity':2,'per_piece_weight':3.5},{'item':'Printer','weight':8}]))
        from run_certificates import run
        created,_,_ = run(self.paths,self.root,selected=('plant1',result['client']['sn']),progress=None)
        self.assertEqual(created,1)
        added = read_batches(self.paths[0])[-1]
        self.assertEqual(added['total_weight'],15); self.assertEqual(added['issued'],'Generated')
        self.assertEqual(len(list((self.root/'Certificates').rglob('*.pdf'))),1)


if __name__ == '__main__':
    unittest.main()
