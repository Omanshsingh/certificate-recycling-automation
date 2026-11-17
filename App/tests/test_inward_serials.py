from datetime import date
from copy import copy
from io import BytesIO
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook

from certificate_core import read_batches
from inward_serials import resequence_added_inward
from certificate_drafts import read_register
from certificate_inventory import scan_inventory, issuance_paths
from destruction_certificate import request_for_inward
from issuance_safety import latest_issuance
from run_certificates import run, find_certificate
from workbook_metadata import write_details
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_workflow import make_workbook
import web_app


class InwardSerialTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='inward-serials-')))
        self.paths = [self.root / 'Plant 1.xlsx', self.root / 'Plant 2.xlsx']
        make_workbook(self.paths[0], [dict(sn=str(sn), supplier=f'Company {sn}', date=date(2026,9,day),
                                         form6=f'FORM-{sn}', reference=f'REF-{sn}')
                                    for sn, day in [(4,1),(5,3),(6,5)]])
        make_workbook(self.paths[1], [dict(sn='1', status='N/A', cod='0040', form6='OTHER-FORM')])
        self.paths[0].write_bytes(write_details(self.paths[0].read_bytes(), 'Inward Register', '5',
                                               {'certificate_fields':{'company_display':'Company 5'}}))
        shutil.copytree(Path(__file__).resolve().parents[1] / 'templates', self.root / 'App' / 'templates')
        save_json(self.root / 'settings.local.json', {'workbook_paths':[str(path) for path in self.paths],
                                                    'reminders_enabled':False})
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'Data' / 'Workbook Backups'))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.root / 'Data' / 'Reports' / 'Latest Results.json'))
        run(self.paths, self.root, check_only=True, progress=None)

    def add(self, day, request='a', plant='plant1'):
        state = web_app.workspace_state(recheck_obsolete=False)
        body = {'plant':plant, 'request_id':request * 32, 'expected_revision':workspace_revision(self.root),
                'details':{'supplier':f'Backdated {request}', 'inward_date':f'2026-09-{day:02}' if day else '',
                           'form6':f'NEW-{request}', 'reference':f'NEWREF-{request}',
                           'items':[{'item':'Laptop','quantity':2,'per_piece_weight':3}]}}
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (body['expected_revision'], state)), \
                patch.object(web_app, 'refresh', side_effect=AssertionError('Adding an entry must use the checked state')):
            result = web_app.add_client(body)
        return result, body

    def add_shared_weight_formula(self):
        book = load_workbook(self.paths[0])
        book.active['H2'] = '=F2*G2'
        book.active['H3'] = '=F3*G3'
        book.save(self.paths[0])
        book.close()
        output = BytesIO()
        with ZipFile(self.paths[0]) as source, ZipFile(output, 'w') as target:
            for member in source.infolist():
                payload = source.read(member.filename)
                if member.filename == 'xl/worksheets/sheet1.xml':
                    xml = ET.fromstring(payload)
                    first = xml.find('.//{*}c[@r="H2"]/{*}f')
                    first.attrib.update(t='shared', si='0', ref='H2:H3')
                    second = xml.find('.//{*}c[@r="H3"]/{*}f')
                    second.attrib.update(t='shared', si='0')
                    second.text = None
                    payload = ET.tostring(xml)
                target.writestr(member, payload)
        self.paths[0].write_bytes(output.getvalue())

    def test_saving_a_late_inward_date_moves_the_entry_and_updates_its_sn(self):
        result, original_request = self.add(None)
        self.assertEqual(result['client']['sn'], '7')
        source = next(row for row in read_batches(self.paths[0]) if row['case_id'] == '7')
        hdd_id = request_for_inward(self.root, 'plant1', '7', workspace_revision(self.root))
        details = dict(web_app.editor_details(source), inward_date='2026-09-02')
        revision = workspace_revision(self.root)
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (revision, web_app.workspace_state(False))):
            saved = web_app.save_inward_changes({'plant':'plant1', 'source_plant':'plant1', 'source_sn':'7',
                'source_kind':'workbook', 'confirmed_excel':True, 'expected_revision':revision, 'details':details})
        self.assertEqual(saved['sn'], '5')
        rows = read_batches(self.paths[0])
        self.assertEqual([(r['case_id'], r['supplier']) for r in rows],
                         [('4','Company 4'), ('5','Backdated a'), ('6','Company 5'), ('7','Company 6')])
        hdd = next(r for r in saved['state']['rows'] if r.get('record_id') == hdd_id)
        self.assertEqual(hdd['destruction_source']['case_id'], '5')
        self.assertEqual(web_app.add_client(original_request)['client']['sn'], '5')

    def test_refresh_renumbers_a_late_date_entered_directly_in_excel(self):
        run(self.paths, self.root, progress=None, allow_warnings=True)
        original_receipt = latest_issuance(self.root, 'workbook', 'plant1', '5')
        self.add(None)
        book = load_workbook(self.paths[0])
        sheet = book.active
        row = next(r for r in read_batches(self.paths[0]) if r['case_id'] == '7')
        sheet.cell(row['source_row'], 2).value = date(2026, 9, 2)
        book.save(self.paths[0]); book.close()
        state = web_app.refresh(force=True)
        rows = read_batches(self.paths[0])
        self.assertEqual([(r['case_id'], r['supplier']) for r in rows],
                         [('4','Company 4'), ('5','Backdated a'), ('6','Company 5'), ('7','Company 6')])
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '6')['id'], original_receipt['id'])
        self.assertFalse(next(r for r in state['rows'] if r.get('plant') == 'plant1' and r['sn'] == '6').get('record_identity_conflict'))

    def test_creating_a_pdf_and_saving_the_late_date_returns_the_new_sn(self):
        self.add(None)
        source = next(row for row in read_batches(self.paths[0]) if row['case_id'] == '7')
        response = web_app.create_certificate({'plant':'plant1', 'source_plant':'plant1', 'source_sn':'7',
            'source_kind':'workbook', 'edit_scope':'update_excel', 'review_edits':True,
            'confirmed_excel':True, 'confirmed':True,
            'data_destruction_needed':True, 'expected_revision':workspace_revision(self.root),
            'details':dict(web_app.editor_details(source), inward_date='2026-09-02')})
        self.assertEqual(response['created'], 1)
        self.assertEqual(response['certificates'][0]['sn'], '5')
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '5')['cod'],
                         response['certificates'][0]['certificate_no'])
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'], '5')

    def test_existing_out_of_order_numbers_can_be_fixed_without_adding_a_client(self):
        make_workbook(self.paths[0], [dict(sn=str(sn), supplier=f'Company {sn}',
                                         date=date(2026,9,day), form6=f'FORM-{sn}')
                                    for sn,day in [(4,1),(25,3),(24,5)]])
        original = self.paths[0].read_bytes()
        rows = read_batches(BytesIO(original))
        updated, _, changes = resequence_added_inward(original, rows[0], original,
            root=self.root, plant='plant1', request_id=None)
        fixed = read_batches(BytesIO(updated))
        self.assertEqual([row['case_id'] for row in fixed], ['4','5','6'])
        self.assertEqual(changes, {'25':'5','24':'6'})
        self.assertEqual([row['supplier'] for row in fixed], [row['supplier'] for row in rows])
        self.assertEqual(self.paths[0].read_bytes(), original)

    def test_backdated_add_with_shared_formula_sorts_and_resequences(self):
        self.add_shared_weight_formula()
        original = read_batches(self.paths[0])
        result, _ = self.add(2, 'c')

        self.assertEqual(result['client']['sn'], '5')
        self.assertFalse(result['order_warning'])
        saved = read_batches(self.paths[0])
        self.assertEqual([row['case_id'] for row in saved], ['4', '5', '6', '7'])
        self.assertEqual([row['supplier'] for row in saved],
                         [original[0]['supplier'], 'Backdated c', original[1]['supplier'], original[2]['supplier']])
        with ZipFile(self.paths[0]) as archive:
            xml = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            formulas = [formula for formula in xml.findall('.//{*}f') if formula.attrib.get('t') == 'shared']
            self.assertEqual(formulas, [])
        book = load_workbook(self.paths[0], data_only=False)
        self.assertEqual(book.active['H2'].value, '=F2*G2')
        self.assertEqual(book.active['H4'].value, '=F4*G4')
        book.close()

    def test_backdated_insert_shifts_later_numbers_and_keeps_certificates_and_hdd_links(self):
        created, _, _ = run(self.paths, self.root, progress=None, allow_warnings=True)
        self.assertEqual(created, 3)
        original = {row['case_id']:row for row in read_batches(self.paths[0])}
        hdd_id = request_for_inward(self.root, 'plant1', '5', workspace_revision(self.root))
        register_bytes = (self.root / 'Data' / 'Custom Certificates.json').read_bytes()
        pdfs = {path:path.read_bytes() for path in (self.root / 'Certificates').rglob('*.pdf')}
        history = {path:path.read_bytes() for path in (self.root / 'Data' / 'Issuance History').glob('*.json')}
        other = self.paths[1].read_bytes()

        result, _ = self.add(2)
        rows = read_batches(self.paths[0])
        self.assertEqual([row['case_id'] for row in rows], ['4','5','6','7'])
        self.assertEqual([row['supplier'] for row in rows], ['Company 4','Backdated a','Company 5','Company 6'])
        self.assertEqual(result['client']['sn'], '5')
        moved = rows[2]
        self.assertEqual(moved['certificate_fields'], {'company_display':'Company 5'})
        self.assertEqual(moved['cod_no'], original['5']['cod_no'])
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '6')['cod'], original['5']['cod_no'])
        self.assertIsNone(latest_issuance(self.root, 'workbook', 'plant1', '5'))
        self.assertTrue(find_certificate(self.root / 'Certificates', dict(moved, source_id='plant1')))
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'], '6')
        self.assertEqual(request_for_inward(self.root, 'plant1', '6', workspace_revision(self.root)), hdd_id)
        inventory = scan_inventory(self.root)
        self.assertTrue(issuance_paths(inventory, moved['cod_no'], 'workbook', 'plant1', '6'))
        self.assertEqual(self.paths[1].read_bytes(), other)
        self.assertEqual((self.root / 'Data' / 'Custom Certificates.json').read_bytes(), register_bytes)
        self.assertTrue(all(path.read_bytes() == data for path, data in pdfs.items()))
        self.assertTrue(all(path.read_bytes() == data for path, data in history.items()))
        expected = workspace_revision(self.root)
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (expected, result['state'])):
            completion_state = web_app.complete_workbook_record({'plant':'plant1','sn':'6','expected_revision':expected})
        after = {row['case_id']:row for row in read_batches(self.paths[0])}
        self.assertEqual(after['6']['issued'], 'Generated')
        self.assertTrue(next(row for row in completion_state['rows']
                             if row.get('plant') == 'plant1' and row.get('sn') == '6')
                        ['recycling_completed_pending_hdd'])
        self.assertEqual(after['5']['issued'], 'Ready')
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '6')
        correction = web_app.create_certificate({'plant':'plant1','source_sn':'6','source_plant':'plant1',
            'details':web_app.editor_details(receipt['after']), 'edit_scope':'certificate_only',
            'confirmed':True, 'expected_revision':workspace_revision(self.root)})
        self.assertEqual(correction['created'], 1)
        self.assertEqual(next(row for row in read_batches(self.paths[0]) if row['case_id'] == '6')['cod_no'], original['5']['cod_no'])
        self.add(1, 'd')
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '7')['cod'], original['5']['cod_no'])
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'], '7')

    def test_repeated_backdated_additions_and_lost_response_follow_the_same_client(self):
        first, first_body = self.add(2)
        second, _ = self.add(1, 'b')
        self.assertEqual(second['client']['sn'], '5')  # Same date goes after the existing first entry.
        retried = web_app.add_client(first_body)
        self.assertEqual(retried['client']['sn'], '6')
        self.assertEqual(len(read_batches(self.paths[0])), 5)
        self.assertEqual([row['case_id'] for row in read_batches(self.paths[0])], ['4','5','6','7','8'])
        self.assertEqual(next(row for row in read_batches(self.paths[0]) if row['case_id'] == '6')['supplier'], 'Backdated a')

    def test_failed_save_leaves_numbers_and_links_unchanged_and_new_rows_are_left_aligned(self):
        before = self.paths[0].read_bytes()
        with patch.object(web_app.WorkbookSession, '_persist', side_effect=OSError('Simulated save failure')):
            with self.assertRaisesRegex(OSError, 'Simulated'):
                self.add(2)
        self.assertEqual(self.paths[0].read_bytes(), before)
        result, _ = self.add(2, 'c')
        row = next(item for item in read_batches(self.paths[0]) if item['case_id'] == result['client']['sn'])
        old_book = load_workbook(BytesIO(before))
        book = load_workbook(self.paths[0])
        try:
            sheet = book[row['source_sheet']]
            for cell in sheet[row['source_row']]:
                self.assertEqual(cell.alignment.horizontal, 'left')
            self.assertEqual(copy(sheet['A1'].alignment), copy(old_book.active['A1'].alignment))
        finally:
            book.close(); old_book.close()


if __name__ == '__main__':
    unittest.main()
