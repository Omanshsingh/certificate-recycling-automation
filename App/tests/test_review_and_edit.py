"""Review edits before first issuance and separate HDD requests, in disposable files."""
import json
import unittest
from unittest.mock import patch

from certificate_drafts import read_register
from issuance_safety import latest_issuance
from pypdf import PdfReader
import web_app
from workspace_review import workspace_revision
from tests import test_simple_certificate_flow as fixtures


class ReviewAndEditTests(unittest.TestCase):
    setUp = fixtures.SimpleCertificateFlowTests.setUp
    source = fixtures.SimpleCertificateFlowTests.source
    body = fixtures.SimpleCertificateFlowTests.body
    create = fixtures.SimpleCertificateFlowTests.create

    def reviewed(self, **changes):
        payload = self.body('standard', **changes)
        payload.update(review_edits=True, edit_scope='certificate_only')
        return payload

    def test_retired_workbook_replacement_cannot_preview_or_save_a_new_cod(self):
        self.assertEqual(self.create(self.reviewed())['created'], 1)
        payload = self.reviewed(supplier='Must not be saved')
        payload.update(void_reissue=True, void_reason='Old browser page')
        before = {path.relative_to(self.root): path.read_bytes()
                  for path in self.root.rglob('*') if path.is_file()}
        for operation in (web_app.preview_certificate, web_app.create_certificate):
            with self.subTest(operation=operation.__name__):
                with self.assertRaisesRegex(ValueError, 'no longer available.*current COD'):
                    operation(payload)
                self.assertEqual({path.relative_to(self.root): path.read_bytes()
                                  for path in self.root.rglob('*') if path.is_file()}, before)

    def test_adding_material_saves_before_the_next_client_without_overwriting_it(self):
        from openpyxl import load_workbook
        from certificate_core import read_batches
        book=load_workbook(self.paths[0]);sheet=book.active
        sheet.append([cell.value for cell in sheet[2]])
        sheet['A3']='1002';sheet['C3']='Keep this client';sheet['I3']='F6-KEEP';sheet['J3']='REF-KEEP';sheet['M3']='N/A'
        book.save(self.paths[0]);book.close()
        before=read_batches(self.paths[0])[1]
        payload=self.reviewed(items=[{'item':'Laptops','quantity':'2','per_piece_weight':'3','weight':'6'},
                                     {'item':'Monitors','quantity':'1','per_piece_weight':'4','weight':'4'}])
        payload.update(edit_scope='update_excel',confirmed_excel=True)
        self.assertEqual(self.create(payload)['created'],1)
        after=read_batches(self.paths[0]);self.assertEqual(len(after[0]['items']),2)
        self.assertEqual(after[0]['total_weight'],10)
        self.assertEqual((after[1]['case_id'],after[1]['supplier'],after[1]['form6'],after[1]['source_row']),
                         (before['case_id'],before['supplier'],before['form6'],4))

    def test_adding_material_extends_the_existing_total_formula(self):
        from openpyxl import load_workbook
        book=load_workbook(self.paths[0]);sheet=book.active
        sheet.append([cell.value for cell in sheet[2]])
        sheet['A3']='1002';sheet['I3']='F6-KEEP';sheet['J3']='REF-KEEP';sheet['M3']='N/A'
        sheet.insert_rows(3);sheet['F2']=2;sheet['G2']=3;sheet['H2']='=F2*G2'
        sheet['E3']='Total';sheet['H3']='=SUM(H2:H2)'
        for column in ('A','B','C','D','I','J','K','L','M'):sheet.merge_cells(f'{column}2:{column}3')
        book.save(self.paths[0]);book.close()
        payload=self.reviewed(items=[{'item':'Laptops','quantity':'2','per_piece_weight':'3','weight':'6'},
                                     {'item':'Monitors','quantity':'1','per_piece_weight':'4','weight':'4'}])
        payload.update(edit_scope='update_excel',confirmed_excel=True)
        self.assertEqual(self.create(payload)['created'],1)
        self.assertEqual(self.source()['total_weight'],10)
        book=load_workbook(self.paths[0]);sheet=book.active
        self.assertEqual(sheet['H4'].value,'=SUM(H2:H3)');self.assertEqual(sheet['A5'].value,'1002');book.close()

    def test_retired_manual_replacement_cannot_preview_or_pass_save_validation(self):
        from certificate_drafts import generate_draft
        details = dict(self.reviewed()['details'], case_id='MANUAL-1',
                       form6='MANUAL-F6', reference='MANUAL-REF', cod_no='', cod_date='')
        result = generate_draft(self.root, 'plant1', details, starting_cod='0001')
        self.assertEqual(result['created'], 1)
        record = read_register(self.root)['records'][0]
        payload = {'source_kind': 'manual', 'plant': 'plant1', 'record_id': record['id'],
                   'details': dict(record['details'], supplier='Must not be saved'),
                   'void_reissue': True, 'void_reason': 'Old browser page'}
        before = {path.relative_to(self.root): path.read_bytes()
                  for path in self.root.rglob('*') if path.is_file()}
        for operation in (web_app.preview_certificate, web_app.validate_manual_identifiers):
            with self.subTest(operation=operation.__name__):
                with self.assertRaisesRegex(ValueError, 'no longer available.*current COD'):
                    operation(payload)
                self.assertEqual({path.relative_to(self.root): path.read_bytes()
                                  for path in self.root.rglob('*') if path.is_file()}, before)

    def test_first_preview_and_no_choice_print_edits_and_keep_excel_details(self):
        original = self.source()
        payload = self.reviewed(supplier='Reviewed company', location='Mumbai')
        before = self.paths[0].read_bytes()
        web_app.preview_certificate(payload)
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(self.create(payload)['created'], 1)
        saved = self.source()
        self.assertEqual(saved['supplier'], original['supplier'])
        self.assertEqual(saved['location'], original['location'])
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        self.assertIn('REVIEWED COMPANY', PdfReader(self.root / receipt['pdf_path']).pages[0].extract_text())

    def test_first_yes_choice_requires_consent_then_saves_reviewed_details_with_pdf(self):
        payload = self.reviewed(supplier='Reviewed company')
        payload['edit_scope'] = 'update_excel'
        before = [path.read_bytes() for path in self.paths]
        with self.assertRaisesRegex(ValueError, 'Yes, update Excel'):
            self.create(payload)
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        payload['confirmed_excel'] = True
        self.assertEqual(self.create(payload)['created'], 1)
        self.assertEqual(self.source()['supplier'], 'Reviewed company')
        self.assertEqual(self.source()['issued'], 'Generated')
        self.assertEqual(self.paths[1].read_bytes(), before[1])

    def test_save_before_first_pdf_requires_consent_and_does_not_allocate_cod(self):
        payload = self.reviewed(supplier='Saved inward name', form6='', reference='')
        before = [path.read_bytes() for path in self.paths]
        with self.assertRaisesRegex(ValueError, 'Yes, update Excel'):
            web_app.save_inward_changes(payload)
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        payload['confirmed_excel'] = True
        result = web_app.save_inward_changes(payload)
        self.assertTrue(result['saved'])
        self.assertEqual(result['created'], 0)
        self.assertEqual(self.source()['supplier'], 'Saved inward name')
        self.assertFalse(self.source()['cod_no'])
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertIsNone(latest_issuance(self.root, 'workbook', 'plant1', '1001'))
        self.assertEqual(self.paths[1].read_bytes(), before[1])

    def test_stale_review_and_failed_save_leave_inward_unchanged(self):
        payload = self.reviewed(supplier='Not saved')
        payload['confirmed_excel'] = True
        original = self.paths[0].read_bytes()
        with patch.object(web_app.WorkbookSession, '_persist', side_effect=OSError('Excel is busy')):
            with self.assertRaisesRegex(OSError, 'Excel is busy'):
                web_app.save_inward_changes(payload)
        self.assertEqual(self.paths[0].read_bytes(), original)
        payload['expected_revision'] = 'old-review'
        with self.assertRaises(ValueError):
            web_app.save_inward_changes(payload)
        self.assertEqual(self.paths[0].read_bytes(), original)

    def test_incomplete_saved_details_show_alert_without_creating_a_pdf(self):
        payload = self.reviewed(inward_date='')
        payload['confirmed_excel'] = True
        result = web_app.save_inward_changes(payload)
        self.assertTrue(result['saved'])
        self.assertEqual(self.source()['issued'], 'Alert')
        self.assertEqual(self.source()['cod_no'], '')
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))

    def test_checking_hdd_marks_one_pending_task_and_does_not_create_its_pdf_or_cod(self):
        payload = self.reviewed()
        payload['data_destruction_needed'] = True
        result = self.create(payload)
        self.assertEqual(result['created'], 1)
        records = read_register(self.root)['records']
        self.assertEqual(len(records), 1)
        request = records[0]
        self.assertEqual(request['details']['certificate_type'], 'destruction')
        self.assertEqual(request['details']['destruction_source'], {'plant':'plant1','case_id':'1001'})
        self.assertEqual(request['details']['cod_no'], '')
        self.assertEqual(request['details']['reference'], '')
        self.assertEqual(len(list((self.root / 'Certificates').rglob('*.pdf'))), 1)
        from destruction_certificate import request_for_inward
        self.assertEqual(request_for_inward(self.root,'plant1','1001',workspace_revision(self.root)), request['id'])
        self.assertEqual(len(read_register(self.root)['records']), 1)
        self.assertTrue(any(row.get('record_id') == request['id'] for row in result['state']['rows']))

    def test_linked_hdd_can_be_issued_before_recycling(self):
        from certificate_drafts import generate_draft
        from destruction_certificate import request_for_inward

        original = self.source()
        record_id = request_for_inward(self.root, 'plant1', '1001', workspace_revision(self.root))
        details = dict(read_register(self.root)['records'][0]['details'], certificate_fields={
            'destruction_date': '2026-09-02',
            'device_receipts': json.dumps([{'date': '2026-09-01', 'count': '4'}]),
        })
        hdd = generate_draft(self.root, 'plant1', details, record_id, starting_cod='0001',
                             expected_revision=workspace_revision(self.root))
        self.assertEqual(hdd['created'], 1)
        inward = self.source()
        self.assertEqual((inward['cod_no'], inward['issued']), ('', original['issued']))
        self.assertEqual(inward['hdd_cod_no'], hdd['certificate_no'])
        self.assertEqual(len(list((self.root / 'Certificates').rglob('*.pdf'))), 1)

        self.assertEqual(self.create(self.reviewed())['created'], 1)
        inward = self.source()
        self.assertTrue(inward['cod_no'])
        self.assertNotEqual(inward['cod_no'], hdd['certificate_no'])
        self.assertEqual(inward['hdd_cod_no'], hdd['certificate_no'])
        self.assertEqual(inward['issued'], 'Generated')
        self.assertEqual(len(list((self.root / 'Certificates').rglob('*.pdf'))), 2)

    def test_save_to_excel_can_mark_hdd_without_creating_any_pdf(self):
        payload = self.reviewed(supplier='Saved before PDF')
        payload.update(confirmed_excel=True, data_destruction_needed=True)
        result = web_app.save_inward_changes(payload)
        self.assertEqual(result['created'], 0)
        self.assertEqual(len(read_register(self.root)['records']), 1)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertEqual(read_register(self.root)['records'][0]['details']['supplier'], 'Saved before PDF')

    def test_issued_inward_requires_corrected_pdf_instead_of_silent_excel_only_change(self):
        self.create(self.reviewed())
        payload = self.reviewed(supplier='Would mismatch the PDF')
        payload['confirmed_excel'] = True
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'already has a PDF'):
            web_app.save_inward_changes(payload)
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_request_failure_keeps_successful_primary_certificate(self):
        payload = self.reviewed()
        payload['data_destruction_needed'] = True
        with patch('destruction_certificate.request_for_inward', side_effect=ValueError('Try the HDD request again')):
            result = self.create(payload)
        self.assertEqual(result['created'], 1)
        self.assertEqual(self.source()['issued'], 'Generated')
        self.assertIn('request could not be updated', result['state']['setup_warning'])
        self.assertEqual(len(list((self.root / 'Certificates').rglob('*.pdf'))), 1)


if __name__ == '__main__':
    unittest.main()
