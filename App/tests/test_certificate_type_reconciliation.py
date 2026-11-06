"""Type-changing saves respect the user's Excel/PDF choice, in temporary files."""
from copy import deepcopy
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from certificate_core import read_batches
from certificate_drafts import read_register
from issuance_safety import latest_issuance
from issuance_safety import recover_issuances, IssuanceTransaction
from workflow_support import run_lock
from workspace_review import workspace_revision
import run_certificates as runner
import web_app
from tests import test_simple_certificate_flow as fixtures


class CertificateTypeReconciliationTests(unittest.TestCase):
    setUp = fixtures.SimpleCertificateFlowTests.setUp
    source = fixtures.SimpleCertificateFlowTests.source
    body = fixtures.SimpleCertificateFlowTests.body
    create = fixtures.SimpleCertificateFlowTests.create

    def reviewed(self, kind='standard', **changes):
        payload = self.body(kind, **changes)
        payload.update(review_edits=True, edit_scope='certificate_only')
        payload.pop('confirmed_excel', None)
        return payload

    def test_preview_allows_each_type_without_excel_consent_or_writes(self):
        before = [path.read_bytes() for path in self.paths]
        for kind in ('standard', 'siemens', 'sita'):
            with self.subTest(kind=kind):
                self.assertTrue(web_app.preview_certificate(self.reviewed(kind))['preview_url'])
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        self.assertIsNone(latest_issuance(self.root, 'workbook', 'plant1', '1001'))

    def test_international_can_delete_original_inward_and_keep_a_local_pdf(self):
        other = self.paths[1].read_bytes()
        original = self.paths[0].read_bytes()
        payload = self.reviewed('sita', supplier='International PDF company',
            items=[{'item':'PDF equipment', 'quantity':'1', 'weight':'7'}])
        payload.update(edit_scope='remove_excel', confirmed_excel_removal=True)
        response = self.create(payload)
        self.assertEqual(response['created'], 1)
        self.assertEqual(read_batches(self.paths[0]), [])
        self.assertEqual(self.paths[1].read_bytes(), other)
        record = next(row for row in read_register(self.root)['records'] if row['id'] == response['record_id'])
        self.assertEqual(record['details']['certificate_type'], 'sita')
        self.assertEqual(record['details']['supplier'], 'International PDF company')
        self.assertEqual(record['details']['form6'], '')
        self.assertEqual(record['details']['case_id'], '')
        self.assertTrue((self.root / 'Certificates' / record['pdf_relative_path']).is_file())
        self.assertFalse(any(row.get('source_kind') == 'workbook' and row.get('sn') == '1001' for row in response['state']['rows']))
        self.assertEqual(response['certificates'][0]['source_kind'], 'manual')
        self.assertTrue(any(path.read_bytes() == original for path in (self.root/'Data'/'Workbook Backups').glob('*.xlsx')))

    def test_excel_update_still_requires_explicit_consent(self):
        payload = self.reviewed('siemens')
        payload['edit_scope'] = 'update_excel'
        before = [path.read_bytes() for path in self.paths]
        with self.assertRaisesRegex(ValueError, 'Yes, update Excel'):
            self.create(payload)
        self.assertEqual([path.read_bytes() for path in self.paths], before)

    def test_confirmed_type_changes_save_selected_inward_and_keep_same_type_pdf_only_edits(self):
        other = self.paths[1].read_bytes()
        cod = None
        for kind in ('siemens', 'sita', 'standard'):
            with self.subTest(kind=kind):
                payload = self.reviewed(kind, supplier=f'{kind} reviewed inward')
                if kind == 'standard':
                    payload['details']['form6'] = 'RECYCLING-F6'
                payload.update(edit_scope='update_excel', confirmed_excel=True)
                self.assertEqual(self.create(payload)['created'], 1)
                source = self.source()
                self.assertEqual(source['certificate_type'], kind)
                self.assertEqual(source['supplier'], f'{kind} reviewed inward')
                if kind in {'sita', 'siemens'}:
                    self.assertEqual(source['form6'], '')
                if cod is None:
                    cod = source['cod_no']
                self.assertEqual(source['cod_no'], cod)
                self.assertEqual(self.paths[1].read_bytes(), other)
                pdf_only = self.reviewed(kind, supplier=f'{kind} PDF correction')
                self.assertEqual(self.create(pdf_only)['created'], 1)
                self.assertEqual(self.source()['supplier'], source['supplier'])
                receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
                self.assertEqual(receipt['after']['supplier'], f'{kind} PDF correction')
                self.assertEqual(receipt['after']['edit_scope'], 'certificate_only')

    def test_failed_type_change_keeps_previous_pdf_excel_and_history(self):
        self.assertEqual(self.create(self.reviewed())['created'], 1)
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        pdf = self.root / receipt['pdf_path']
        before = ([path.read_bytes() for path in self.paths], pdf.read_bytes(), deepcopy(receipt))
        payload = self.reviewed('siemens')
        payload.update(edit_scope='update_excel', confirmed_excel=True)
        with patch.object(runner.WorkbookSession, '_persist', side_effect=PermissionError('Excel is busy')):
            with self.assertRaisesRegex(ValueError, 'Excel is busy'):
                self.create(payload)
        self.assertEqual([path.read_bytes() for path in self.paths], before[0])
        self.assertEqual(pdf.read_bytes(), before[1])
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '1001'), before[2])

    def test_existing_pdf_only_type_difference_requires_a_choice_on_next_save(self):
        payload = self.body('siemens')
        details, _ = web_app.workbook_form(payload, payload['details'], self.paths)
        self.assertEqual(runner.run(self.paths, self.root, selected=('plant1', '1001'),
            details=details, update_excel=False, starting_cod='0001', progress=None)[0], 1)
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        payload = self.reviewed('siemens')
        payload['details'] = web_app.editor_details(receipt['after'])
        payload['details']['supplier'] = 'Corrected PDF-only medical company'
        with self.assertRaisesRegex(ValueError, 'Choose Update Excel or Delete Excel inward'):
            self.create(payload)
        payload.update(edit_scope='update_excel', confirmed_excel=True)
        self.assertEqual(self.create(payload)['created'], 1)
        self.assertEqual(self.source()['certificate_type'], 'siemens')
        self.assertEqual(self.source()['cod_no'], receipt['cod'])
        saved = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        self.assertEqual(saved['after']['supplier'], 'Corrected PDF-only medical company')
        self.assertEqual(saved['after']['certificate_type'], 'siemens')

    def test_type_change_with_hdd_still_creates_a_separate_unnumbered_request(self):
        payload = self.reviewed('siemens')
        payload.update(edit_scope='remove_excel', confirmed_excel_removal=True, data_destruction_needed=True)
        response = self.create(payload)
        self.assertEqual(response['created'], 1)
        self.assertEqual(read_batches(self.paths[0]), [])
        records = read_register(self.root)['records']
        self.assertEqual(len(records), 2)
        request = next(row for row in records if row['details']['certificate_type'] == 'destruction')
        self.assertEqual(request['details']['cod_no'], '')
        self.assertEqual(request['details']['destruction_source'], {'plant': None, 'case_id': '1001', 'record_id':response['record_id']})

    def test_reconciliation_does_not_change_another_inward_in_the_same_workbook(self):
        book = load_workbook(self.paths[0])
        try:
            sheet = book.active
            sheet.append([cell.value for cell in sheet[2]])
            sheet.cell(3, 1, '1002')
            sheet.cell(3, 3, 'Another inward')
            sheet.cell(3, 9, 'FORM6-1002')
            sheet.cell(3, 10, 'REF-1002')
            book.save(self.paths[0])
        finally:
            book.close()
        other = web_app.editor_details(read_batches(self.paths[0])[1])
        payload = self.reviewed('siemens', supplier='Selected inward corrected')
        payload.update(edit_scope='update_excel', confirmed_excel=True)
        self.assertEqual(self.create(payload)['created'], 1)
        batches = read_batches(self.paths[0])
        self.assertEqual(web_app.editor_details(next(row for row in batches if row['case_id'] == '1002')), other)
        self.assertEqual(next(row for row in batches if row['case_id'] == '1001')['certificate_type'], 'siemens')

    def removal(self, kind='siemens', **changes):
        payload = self.reviewed(kind, **changes)
        payload.update(edit_scope='remove_excel', confirmed_excel_removal=True)
        return payload

    def test_deletion_requires_confirmation_type_change_and_fresh_review(self):
        before = self.paths[0].read_bytes()
        payload = self.removal()
        payload['confirmed_excel_removal'] = False
        with self.assertRaisesRegex(ValueError, 'Confirm deleting'):
            self.create(payload)
        with self.assertRaisesRegex(ValueError, 'only available for a certificate type change'):
            self.create(self.removal('standard'))
        payload = self.removal()
        payload['expected_revision'] = 'stale'
        with self.assertRaises(ValueError):
            self.create(payload)
        with self.assertRaisesRegex(ValueError, 'Choose Update Excel or Delete Excel inward'):
            self.create(self.reviewed('siemens'))
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(read_register(self.root)['records'], [])

    def test_issued_type_change_refuses_deletion_and_preserves_history_and_hdd(self):
        payload=self.reviewed();payload['data_destruction_needed']=True
        self.assertEqual(self.create(payload)['created'],1)
        old=latest_issuance(self.root,'workbook','plant1','1001')
        before={path.relative_to(self.root):path.read_bytes() for path in self.root.rglob('*') if path.is_file()}
        with self.assertRaisesRegex(ValueError,'generated.*cannot be deleted'):
            self.create(self.removal())
        self.assertEqual({path.relative_to(self.root):path.read_bytes() for path in self.root.rglob('*') if path.is_file()},before)
        self.assertEqual(latest_issuance(self.root,'workbook','plant1','1001')['id'],old['id'])

    def test_failed_first_generation_removal_preserves_inward_and_pending_hdd(self):
        from destruction_certificate import request_for_inward
        request_for_inward(self.root,'plant1','1001',workspace_revision(self.root))
        registry=self.root/'Data'/'Custom Certificates.json'
        before=self.paths[0].read_bytes(),registry.read_bytes()
        with patch.object(runner.WorkbookSession,'_persist',side_effect=PermissionError('Excel is busy')):
            with self.assertRaisesRegex(ValueError,'Excel is busy'):
                self.create(self.removal())
        self.assertEqual(self.paths[0].read_bytes(),before[0]);self.assertEqual(registry.read_bytes(),before[1])
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_interruption_before_first_excel_commit_rolls_back_pdf_and_local_register(self):
        class Interrupted(BaseException):pass
        before=self.paths[0].read_bytes()
        with patch.object(runner.WorkbookSession,'_persist',side_effect=Interrupted):
            with self.assertRaises(Interrupted):self.create(self.removal())
        with run_lock(self.root):recover_issuances(self.root)
        self.assertEqual(self.paths[0].read_bytes(),before)
        self.assertEqual(read_register(self.root)['records'],[])
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_interruption_after_excel_commit_finishes_local_certificate_on_recovery(self):
        payload = self.removal('sita')
        details, _ = web_app.workbook_form(payload, payload['details'], self.paths)
        with patch.object(IssuanceTransaction, 'complete', side_effect=OSError('Interrupted history update')):
            self.assertEqual(runner.run(self.paths, self.root, selected=('plant1','1001'),
                details=details, update_excel=False, remove_inward=True, starting_cod='0001', progress=None)[0], 1)
        with run_lock(self.root):
            recover_issuances(self.root)
        self.assertEqual(read_batches(self.paths[0]), [])
        kept = read_register(self.root)['records'][0]
        self.assertTrue((self.root/'Certificates'/kept['pdf_relative_path']).is_file())
        self.assertEqual(latest_issuance(self.root,'manual','plant1',kept['id'])['state'], 'complete')

    def test_standalone_medical_still_creates_pdf_without_changing_excel(self):
        from certificate_drafts import generate_draft
        before = [path.read_bytes() for path in self.paths]
        details = self.body('siemens')['details']
        details.update(case_id='', form6='')
        response = generate_draft(self.root, 'plant1', details, starting_cod='0001')
        self.assertEqual(response['created'], 1)
        self.assertEqual([path.read_bytes() for path in self.paths], before)

    def test_deleted_excel_identity_is_not_reused_by_local_hdd_or_history(self):
        payload = self.removal()
        payload['data_destruction_needed'] = True
        response = self.create(payload)
        from tests.test_certificate_workflow import make_workbook
        make_workbook(self.paths[0], [{'sn':'1001','supplier':'New unrelated inward','form6':'F6-NEW'}])
        state = web_app.rebuild_history_state()
        row = web_app.target_row(state, ('plant1','1001'))
        self.assertFalse(row.get('has_issued_certificate'))
        self.assertFalse(row.get('record_identity_conflict'))
        from inward_removal import remove_inward
        check = dict(plant='plant1', sn='1001', sheet=row['sheet'], row=row['row'],
            request_id='b'*32, expected_revision=workspace_revision(self.root))
        self.assertEqual(remove_inward(self.root, check)['entry']['supplier'], 'New unrelated inward')
        self.assertEqual(len(read_register(self.root)['records']), 2)
        self.assertTrue(any(entry.get('record_id') == response['record_id'] for entry in state['rows']))

    def test_local_certificate_and_linked_hdd_remain_editable_after_excel_removal(self):
        payload = self.removal()
        payload['data_destruction_needed'] = True
        response = self.create(payload)
        from certificate_drafts import generate_draft, complete_draft
        records = read_register(self.root)['records']
        base = next(record for record in records if record['id'] == response['record_id'])
        hdd = next(record for record in records if record['details']['certificate_type'] == 'destruction')
        before = [path.read_bytes() for path in self.paths]
        details = dict(base['details'], supplier='Corrected local Medical company')
        corrected = generate_draft(self.root, 'plant1', details, record_id=base['id'],
                                  expected_revision=workspace_revision(self.root))
        self.assertEqual(corrected['created'], 1)
        self.assertEqual(corrected['certificate_no'], base['details']['cod_no'])
        hdd_details = dict(hdd['details'], certificate_fields={
            'destruction_date':'2026-10-29', 'device_receipts':'[{"date":"2026-10-27","count":"2"}]'})
        result = generate_draft(self.root, 'plant1', hdd_details, record_id=hdd['id'],
                                expected_revision=workspace_revision(self.root))
        self.assertEqual(result['created'], 1)
        complete_draft(self.root, base['id'], expected_revision=workspace_revision(self.root))
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        self.assertEqual(next(record for record in read_register(self.root)['records'] if record['id'] == base['id'])['status'], 'completed')

    def test_first_generation_conversion_removes_materials_and_compacts_other_inward(self):
        book = load_workbook(self.paths[0])
        try:
            sheet = book.active
            values = [cell.value for cell in sheet[2]]
            sheet.append([None,None,None,None,'Monitors',2,3,6])
            sheet.append(values)
            sheet.cell(4, 1, '1002'); sheet.cell(4, 3, 'Keep this inward')
            sheet.cell(4, 9, 'F6-KEEP'); sheet.cell(4, 10, 'REF-KEEP')
            for column in ('A','B','C','D','I','J','K','L','M'):
                sheet.merge_cells(f'{column}2:{column}3')
            book.save(self.paths[0])
        finally:
            book.close()
        source = read_batches(self.paths[0])
        other = next(row for row in source if row['case_id'] == '1002')
        other_facts = web_app.editor_details(other)
        other_plant = self.paths[1].read_bytes()
        self.assertEqual(len(self.source()['items']), 2)
        self.assertEqual(self.create(self.removal())['created'], 1)
        remaining = read_batches(self.paths[0])
        self.assertEqual(len(remaining), 1)
        self.assertEqual(web_app.editor_details(remaining[0]), dict(other_facts,case_id='1001'))
        self.assertEqual(remaining[0]['source_row'], 2)
        self.assertEqual(self.paths[1].read_bytes(), other_plant)


if __name__ == '__main__':
    unittest.main()
