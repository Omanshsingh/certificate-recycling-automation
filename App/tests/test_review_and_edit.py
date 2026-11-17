"""Review edits before first issuance and separate HDD requests, in disposable files."""
import json
from pathlib import Path
import unittest
from unittest.mock import patch

from certificate_drafts import read_register
from issuance_safety import latest_issuance
from pypdf import PdfReader
import web_app
from web_app import refresh as refresh_workbooks
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
        row = web_app.target_row(result['state'], ('plant1', '1001'))
        self.assertEqual(row['outcome'], 'needs_review')
        self.assertIn('Form 6 number is required to create the certificate.', row['warnings'])
        self.assertIn('Reference number is required to create the certificate.', row['warnings'])
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

    def test_save_company_name_keeps_na_status_and_blocks_issuance(self):
        from openpyxl import load_workbook

        book = load_workbook(self.paths[0])
        book.active['M2'] = 'N/A'
        book.save(self.paths[0])
        book.close()
        self.assertEqual(self.source()['issued'], 'N/A')
        payload = self.reviewed(supplier='Renamed N/A client', form6='', reference='')
        payload['confirmed_excel'] = True
        result = web_app.save_inward_changes(payload)
        self.assertEqual(result['created'], 0)
        self.assertEqual(self.source()['supplier'], 'Renamed N/A client')
        self.assertEqual(self.source()['issued'], 'N/A')
        self.assertEqual(web_app.target_row(result['state'], ('plant1', '1001'))['certificate_status'], 'na')
        payload['expected_revision'] = workspace_revision(self.root)
        with self.assertRaisesRegex(ValueError, 'marked N/A'):
            web_app.preview_certificate(payload)

    def test_mark_na_and_undo_restore_status_from_inward_details(self):
        for complete, expected in ((True, 'Ready'), (False, 'Alert')):
            with self.subTest(complete=complete):
                if not complete:
                    from openpyxl import load_workbook
                    with self.paths[0].open('rb') as stream:
                        book = load_workbook(stream)
                    book.active['B2'] = None
                    book.save(self.paths[0])
                    book.close()
                def change(action):
                    return web_app.update_inward_status({
                        'plant': 'plant1', 'sn': '1001', 'action': action,
                        'expected_revision': workspace_revision(self.root)})
                self.assertEqual(change('mark_na')['status'], 'N/A')
                self.assertEqual(self.source()['issued'], 'N/A')
                self.assertEqual(change('undo_na')['status'], expected)
                self.assertEqual(self.source()['issued'], expected)
                self.assertFalse(self.source()['cod_no'])
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))

    def test_external_pdf_links_writes_tracking_and_keeps_pdf_bytes(self):
        from datetime import date
        from issuance_safety import file_digest
        from run_certificates import find_certificate
        from tests.test_certificate_inventory import external_pdf

        folder = self.root / 'Certificates'
        relative = 'Supplier Certificate.pdf'
        path = external_pdf(folder / relative, cod='0419')
        original = path.read_bytes()
        result = web_app.link_external_certificate({
            'plant':'plant1', 'sn':'1001', 'relative_path':relative, 'sha256':file_digest(path),
            'certificate_date':'2026-10-14', 'expected_revision':workspace_revision(self.root),
        })
        self.assertTrue(result['saved'])
        source = self.source()
        self.assertEqual((source['issued'], source['cod_no'], source['cod_date']),
                         ('Generated Externally', '0419', date(2026, 10, 14)))
        self.assertEqual(find_certificate(folder, dict(source, source_id='plant1')), path)
        self.assertEqual(path.read_bytes(), original)
        self.assertEqual(list(folder.rglob('*.pdf')), [path])

    def test_app_only_na_blocks_issuance_and_undo_restores_local_status(self):
        from certificate_drafts import generate_draft, save_draft

        details = dict(self.reviewed()['details'], case_id='LOCAL', form6='LOCAL-FORM6', reference='LOCAL-REF')
        record_id = save_draft(self.root, None, details)
        def change(action):
            return web_app.update_inward_status({'record_id':record_id, 'action':action,
                'expected_revision':workspace_revision(self.root)})
        self.assertEqual(change('mark_na')['status'], 'N/A')
        with self.assertRaisesRegex(ValueError, 'Undo N/A'):
            generate_draft(self.root, record_id=record_id)
        self.assertEqual(change('undo_na')['status'], 'Ready')
        self.assertEqual(read_register(self.root)['records'][0]['status'], 'ready')

    def test_app_only_undo_na_returns_alert_when_material_details_are_incomplete(self):
        from certificate_drafts import save_draft

        details = dict(self.reviewed()['details'], case_id='LOCAL-ALERT',
                       form6='LOCAL-FORM6', reference='LOCAL-REF')
        details['items'] = [{'item':'Laptop', 'quantity':'1', 'per_piece_weight':'', 'weight':''}]
        record_id = save_draft(self.root, None, details)

        def change(action):
            return web_app.update_inward_status({
                'record_id':record_id, 'action':action,
                'expected_revision':workspace_revision(self.root),
            })

        self.assertEqual(change('mark_na')['status'], 'N/A')
        undone = change('undo_na')
        self.assertEqual(undone['status'], 'Alert')
        self.assertEqual(read_register(self.root)['records'][0]['status'], 'alert')
        local_row = next(row for row in undone['state']['rows'] if row.get('record_id') == record_id)
        self.assertEqual(local_row['certificate_status'], 'alert')

    def test_uploaded_pdf_attaches_to_excel_or_creates_an_app_only_record(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf
        from run_certificates import find_certificate
        from certificate_drafts import current_record_pdf
        from certificate_inventory import scan_inventory

        source_file = self.root / 'operator.pdf'
        other_before = self.paths[1].read_bytes()
        for cod, inward in (('0421', {'plant':'plant1', 'sn':'1001'}), ('0422', None)):
            original = external_pdf(source_file, cod=cod).read_bytes()
            stage = stage_upload(self.root, original)
            self.assertEqual(stage['cod'], cod)
            result = save_upload(self.root, web_app.BACKUP_DIR, {
                'token':stage['token'], 'inward':inward, 'expected_revision':workspace_revision(self.root),
                'details':{'supplier':'Independent client', 'certificate_type':'standard',
                           'cod_date':'2026-10-14', 'location':'Pune', 'reference':'EXT-REF'},
            })
            self.assertEqual(result['certificate_no'], cod)
            saved = self.root / 'Certificates' / result['relative_path']
            self.assertEqual(saved.read_bytes(), original)
            self.assertEqual(Path(result['relative_path']).parent, Path('.'))
            entry = next(item for item in scan_inventory(self.root)['entries']
                         if item.get('present') and item['path'] == f"Certificates/{result['relative_path']}")
            self.assertTrue(entry['external'])
            if inward:
                batch = self.source()
                self.assertEqual((batch['cod_no'], batch['issued']), (cod, 'Generated Externally'))
                self.assertEqual(find_certificate(self.root / 'Certificates', dict(batch, source_id='plant1')), saved)
                linked_before = self.paths[0].read_bytes()
            else:
                record = read_register(self.root)['records'][0]
                self.assertEqual(record['details']['supplier'], 'Independent client')
                self.assertEqual(record['status'], 'generated_manual')
                self.assertEqual(current_record_pdf(self.root, record), saved)
                self.assertEqual(self.paths[0].read_bytes(), linked_before)
        self.assertEqual(self.paths[1].read_bytes(), other_before)

    def test_upload_keeps_client_details_read_only_and_requires_review_edit_for_changes(self):
        from external_certificate_upload import stage_upload
        from tests.test_certificate_inventory import external_pdf

        self.enterContext(patch.object(web_app, 'refresh', new=refresh_workbooks))
        original = self.source()
        case = {'plant':'plant1', 'sn':'1001'}
        kind_date = {'certificate_type':'standard', 'cod_date':'2026-10-14'}

        pdf = external_pdf(self.root / 'upload.pdf', cod='0421', reference='External certificate').read_bytes()
        staged = stage_upload(self.root, pdf)
        with self.assertRaisesRegex(ValueError, 'read-only during upload'):
            web_app.save_uploaded_certificate({
                'token':staged['token'], 'inward':case, 'details':kind_date,
                'certificate_edits':{'case_id':'other inward'},
                'expected_revision':workspace_revision(self.root),
            })
        self.assertEqual(self.source()['cod_no'], '')

        first = web_app.save_uploaded_certificate({
            'token':staged['token'], 'inward':case, 'details':kind_date,
            'expected_revision':workspace_revision(self.root),
        })
        row = web_app.target_row(first['state'], ('plant1', '1001'))
        self.assertEqual(row['last_certificate_details']['supplier'], original['supplier'])
        self.assertEqual(row['last_certificate_details']['location'], original['location'])
        self.assertEqual(row['last_certificate_details']['reference'], original['reference'])
        uploaded_source = self.source()
        self.assertEqual((uploaded_source['cod_no'], uploaded_source['issued']),
                         ('0421', 'Generated Externally'))
        self.assertEqual(uploaded_source['supplier'], original['supplier'])
        self.assertEqual(uploaded_source['location'], original['location'])

        replacement_pdf = external_pdf(self.root / 'replacement.pdf', cod=first['certificate_no'],
                                       reference='Replacement PDF').read_bytes()
        replacement = stage_upload(self.root, replacement_pdf)
        second = web_app.save_uploaded_certificate({
            'token':replacement['token'], 'replacement':case, 'details':kind_date,
            'expected_revision':workspace_revision(self.root),
        })
        row = web_app.target_row(second['state'], ('plant1', '1001'))
        self.assertEqual(row['last_certificate_details']['supplier'], original['supplier'])
        self.assertEqual(row['last_certificate_details']['reference'], original['reference'])
        self.assertEqual(second['certificate_no'], first['certificate_no'])
        current = self.source()
        for field in ('supplier', 'location', 'inward_date', 'reference', 'form6', 'items', 'certificate_weight'):
            self.assertEqual(current[field], original[field])

    def test_replacement_upload_keeps_cod_archives_previous_versions_and_reopens_completed(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf
        from workbook_store import WorkbookSession

        self.assertEqual(self.create(self.reviewed(supplier='Name used on the certificate'))['created'], 1)
        original = self.source()
        cod = original['cod_no']
        with WorkbookSession(self.paths[0], web_app.BACKUP_DIR) as session:
            session.set_issued_status('1001', original['source_sheet'], original['source_row'],
                                      cod, 'Generated', 'Completed')
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        old_path = self.root / receipt['pdf_path']
        old_pdf = old_path.read_bytes()
        before = self.paths[0].read_bytes()
        other = self.paths[1].read_bytes()
        body = {'replacement':{'plant':'plant1', 'sn':'1001'},
                'details':{'cod_no':cod, 'cod_date':'2026-10-14'}}
        for index, number in enumerate(('0428', cod, cod, cod)):
            payload = external_pdf(self.root / f'replacement-{index}.pdf', cod=number,
                                   reference=f'Revision {index}').read_bytes()
            body.update(token=stage_upload(self.root, payload)['token'],
                        expected_revision=workspace_revision(self.root))
            if index == 0:
                with self.assertRaisesRegex(ValueError, 'match the number printed'):
                    save_upload(self.root, web_app.BACKUP_DIR, body)
                self.assertEqual(self.paths[0].read_bytes(), before)
                self.assertEqual(old_path.read_bytes(), old_pdf)
                continue
            if index == 1:
                with patch.object(WorkbookSession, '_persist', side_effect=OSError('Excel is busy')):
                    with self.assertRaisesRegex(OSError, 'Excel is busy'):
                        save_upload(self.root, web_app.BACKUP_DIR, body)
                self.assertEqual(self.paths[0].read_bytes(), before)
                self.assertEqual(old_path.read_bytes(), old_pdf)
                self.assertEqual(list((self.root / 'Certificates').rglob('*.pdf')), [old_path])
            result = save_upload(self.root, web_app.BACKUP_DIR, body)
            self.assertTrue(result['replaced'])
            self.assertEqual(result['certificate_no'], cod)
            current = self.root / 'Certificates' / result['relative_path']
            self.assertEqual(current.read_bytes(), payload)
            self.assertEqual(list((self.root / 'Certificates').rglob('*.pdf')), [current])
            self.assertFalse(old_path.exists())
            receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
            self.assertEqual((self.root / receipt['previous_pdfs'][0]['archive']).read_bytes(), old_pdf)
            self.assertEqual(receipt['after']['supplier'], 'Name used on the certificate')
            row = self.source()
            self.assertEqual((row['cod_no'], row['issued'], row['supplier']),
                             (cod, 'Generated Externally', original['supplier']))
            old_path, old_pdf = current, payload
        self.assertEqual(self.paths[1].read_bytes(), other)

    def test_replace_app_only_and_linked_hdd_pdfs(self):
        from certificate_drafts import current_record_pdf
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf

        for cod, kind, inward in (
            ('0431', 'standard', None),
            ('0432', 'destruction', {'plant':'plant1', 'sn':'1001'}),
        ):
            first = external_pdf(self.root / 'first.pdf', cod=cod).read_bytes()
            issued = save_upload(self.root, web_app.BACKUP_DIR, {
                'token':stage_upload(self.root, first)['token'], 'inward':inward,
                'expected_revision':workspace_revision(self.root),
                'details':{'certificate_type':kind, 'supplier':'Upload client', 'cod_no':cod,
                           'cod_date':'2026-10-14'},
            })
            identifier = issued['record_id']
            old_path = self.root / 'Certificates' / issued['relative_path']
            revised = external_pdf(self.root / 'revised.pdf', cod=cod, reference='New design').read_bytes()
            result = save_upload(self.root, web_app.BACKUP_DIR, {
                'token':stage_upload(self.root, revised)['token'], 'replacement':{'record_id':identifier},
                'expected_revision':workspace_revision(self.root), 'details':{'cod_date':'2026-10-15'},
            })
            record = next(record for record in read_register(self.root)['records'] if record['id'] == identifier)
            self.assertEqual(result['record_id'], identifier)
            self.assertEqual(record['details']['certificate_type'], kind)
            self.assertEqual(current_record_pdf(self.root, record).read_bytes(), revised)
            self.assertFalse(old_path.exists())
            if inward:
                workbook = self.source()
                self.assertEqual((workbook['cod_no'], workbook['hdd_cod_no']), ('', cod))
                self.assertEqual(workbook['hdd_cod_date'], '2026-10-15')

    def test_uploaded_hdd_preserves_the_existing_recycling_certificate(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf
        from certificate_drafts import current_record_pdf

        self.assertEqual(self.create(self.reviewed())['created'], 1)
        original = self.source()
        payload = external_pdf(self.root/'operator.pdf', cod='0423').read_bytes()
        staged = stage_upload(self.root, payload)
        result = save_upload(self.root, web_app.BACKUP_DIR, {
            'token':staged['token'], 'inward':{'plant':'plant1','sn':'1001'},
            'expected_revision':workspace_revision(self.root),
            'details':{'certificate_type':'destruction', 'cod_no':'0423', 'cod_date':'2026-10-14'},
        })
        saved = self.source()
        self.assertEqual((saved['cod_no'], saved['issued'], saved['certificate_type']),
                         (original['cod_no'], original['issued'], original['certificate_type']))
        self.assertEqual(saved['hdd_cod_no'], '0423')
        self.assertEqual(current_record_pdf(self.root, read_register(self.root)['records'][0]).read_bytes(), payload)
        self.assertTrue(result['saved'])

    def test_external_upload_rejects_wrong_cod_and_rolls_back_failed_excel_save(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf

        staged = stage_upload(self.root, external_pdf(self.root/'operator.pdf', cod='0424').read_bytes())
        body = {'token':staged['token'], 'inward':{'plant':'plant1','sn':'1001'},
                'expected_revision':workspace_revision(self.root),
                'details':{'certificate_type':'standard', 'cod_no':'0425', 'cod_date':'2026-10-14'}}
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'match a number printed'):
            save_upload(self.root, web_app.BACKUP_DIR, body)
        body['details']['cod_no'] = '0424'
        with patch.object(web_app.WorkbookSession, '_persist', side_effect=OSError('Excel is busy')):
            with self.assertRaisesRegex(OSError, 'Excel is busy'):
                save_upload(self.root, web_app.BACKUP_DIR, body)
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_invalid_upload_fields_allow_retrying_the_same_pdf(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf

        payload = external_pdf(self.root/'operator.pdf', cod='0424').read_bytes()
        staged = stage_upload(self.root, payload)
        body = {'token':staged['token'], 'expected_revision':workspace_revision(self.root),
                'details':{'supplier':'Independent client', 'certificate_type':'standard',
                           'cod_date':'2026-10-14', 'reference':'EXT-REF'}}
        with self.assertRaisesRegex(ValueError, 'Choose the certificate type'):
            save_upload(self.root, web_app.BACKUP_DIR,
                        dict(body, details=dict(body['details'], certificate_type=[])))
        with self.assertRaisesRegex(ValueError, 'Choose an Excel inward'):
            save_upload(self.root, web_app.BACKUP_DIR, dict(body, inward={'plant':{}, 'sn':'1001'}))
        result = save_upload(self.root, web_app.BACKUP_DIR, body)
        self.assertTrue(result['saved'])
        self.assertEqual((self.root/'Certificates'/result['relative_path']).read_bytes(), payload)

    def test_upload_without_readable_cod_uses_numbering_and_preserves_original_pdf(self):
        from external_certificate_upload import save_upload, stage_upload
        from tests.test_certificate_inventory import external_pdf
        from workflow_support import load_settings, save_json

        save_json(self.root/'settings.local.json', dict(load_settings(self.root), next_cod=451))
        payload = external_pdf(self.root/'operator.pdf').read_bytes()
        staged = stage_upload(self.root, payload)
        self.assertEqual(staged['cod'], '')
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))
        body = {'token':staged['token'], 'inward':{'plant':'plant1','sn':'1001'},
                'expected_revision':workspace_revision(self.root), 'automatic_cod':True,
                'details':{'certificate_type':'standard', 'cod_no':'0450', 'cod_date':'2026-10-14'}}
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'suggested COD is now 0451'):
            save_upload(self.root, web_app.BACKUP_DIR, body)
        self.assertEqual(self.paths[0].read_bytes(), before)
        body['details']['cod_no'] = '0451'
        result = save_upload(self.root, web_app.BACKUP_DIR, body)
        self.assertEqual(self.source()['cod_no'], '0451')
        self.assertEqual((self.root/'Certificates'/result['relative_path']).read_bytes(), payload)

        payload = external_pdf(self.root/'another.pdf', reference='Another certificate').read_bytes()
        body.update(token=stage_upload(self.root, payload)['token'], inward=None,
                    expected_revision=workspace_revision(self.root))
        del body['automatic_cod'], body['details']['cod_no']
        body['details']['supplier'] = 'App-only client'
        result = save_upload(self.root, web_app.BACKUP_DIR, body)
        self.assertEqual(result['certificate_no'], '0452')
        self.assertEqual((self.root/'Certificates'/result['relative_path']).read_bytes(), payload)

    def test_upload_without_readable_cod_uses_existing_inward_cod(self):
        from external_certificate_upload import save_upload, stage_upload
        from openpyxl import load_workbook
        from tests.test_certificate_inventory import external_pdf

        book = load_workbook(self.paths[0])
        book.active['K2'] = '0426'
        book.active['M2'] = 'Generated Externally'
        book.save(self.paths[0])
        book.close()
        payload = external_pdf(self.root/'operator.pdf').read_bytes()
        result = save_upload(self.root, web_app.BACKUP_DIR, {
            'token':stage_upload(self.root, payload)['token'], 'inward':{'plant':'plant1','sn':'1001'},
            'expected_revision':workspace_revision(self.root),
            'details':{'certificate_type':'standard', 'cod_date':'2026-10-14'},
        })
        self.assertEqual(result['certificate_no'], '0426')
        self.assertEqual(self.source()['cod_no'], '0426')

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

    def issue_two_material_certificate(self):
        from openpyxl import load_workbook
        book = load_workbook(self.paths[0])
        sheet = book.active
        sheet['E3'] = 'Monitors'
        sheet['F3'] = 1
        sheet['G3'] = 4
        sheet['H3'] = 4
        book.save(self.paths[0]); book.close()
        issued_details = web_app.editor_details(self.source())
        self.assertEqual(len(issued_details['items']), 2)
        self.assertEqual(self.create(self.reviewed())['created'], 1)
        return latest_issuance(self.root, 'workbook', 'plant1', '1001')

    def edit_workbook_after_issue(self):
        from datetime import date
        from openpyxl import load_workbook
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet['B2'] = date(2026, 8, 18)
        sheet['C2'] = 'Edited company name'
        sheet['D2'] = 'Edited site'
        sheet['E2'] = 'Edited material'
        sheet['F2'] = 4
        sheet['G2'] = 2
        sheet['H2'] = 8
        sheet['I2'] = 'F6-EDITED'
        sheet['J2'] = 'REF-EDITED'
        for column in 'EFGH': sheet[f'{column}3'] = None
        book.save(self.paths[0]); book.close()

    def issued_change_payload(self, action, *, revision=None):
        return {'action':action, 'plant':'plant1', 'sn':'1001',
                'expected_revision':revision or workspace_revision(self.root)}

    def test_change_comparison_normalizes_dates_and_numbers_but_shows_zero_and_all_material_changes(self):
        from issued_detail_changes import compare_details
        current = {'inward_date':'01/09/26', 'supplier':'  Acme   Ltd ',
                   'items':[{'item':'Asset', 'quantity':'2.0', 'per_piece_weight':'0', 'weight':'0'}]}
        issued = {'inward_date':'2026-09-01', 'supplier':'Acme Ltd',
                  'items':[{'item':'Asset', 'quantity':2, 'per_piece_weight':1, 'weight':2}]}
        changes = compare_details(current, issued)
        self.assertEqual([change['field'] for change in changes], ['Materials and weights'])
        self.assertIn('0 kg each', changes[0]['current'])
        self.assertIn('Qty 2', changes[0]['current'])
        self.assertEqual(compare_details({'inward_date':'1/9/26'}, {'inward_date':'2026-09-01'}), [])

    def test_keep_marks_only_the_reviewed_values_and_does_not_touch_excel_status_or_pdf(self):
        from certificate_core import read_batches
        from issued_detail_changes import is_intentionally_kept
        from openpyxl import load_workbook
        from web_app import current_pdf_rows

        receipt = self.issue_two_material_certificate()
        web_app.complete_workbook_record({'plant':'plant1', 'sn':'1001',
            'expected_revision':workspace_revision(self.root)})
        self.edit_workbook_after_issue()
        workbook_before = self.paths[0].read_bytes()
        pdf_path = self.root / receipt['pdf_path']
        pdf_before = pdf_path.read_bytes()
        source = read_batches(self.paths[0])[0]
        payload = self.issued_change_payload('keep')
        result = web_app.resolve_issued_excel_changes(payload)

        self.assertTrue(result['saved'])
        self.assertEqual(result['action'], 'keep')
        self.assertEqual(self.paths[0].read_bytes(), workbook_before)
        self.assertEqual(pdf_path.read_bytes(), pdf_before)
        saved = read_batches(self.paths[0])[0]
        self.assertEqual((saved['issued'], saved['cod_no']), ('Completed', source['cod_no']))
        self.assertTrue(is_intentionally_kept(self.root, receipt, web_app.editor_details(saved)))
        row = dict(saved, plant='plant1', source_kind='workbook', sn='1001', certificate_status='Completed')
        self.assertFalse(current_pdf_rows([row])[0]['excel_details_changed_since_issue'])

        book = load_workbook(self.paths[0]); book.active['C2'] = 'Changed once more'; book.save(self.paths[0]); book.close()
        changed_again = read_batches(self.paths[0])[0]
        row = dict(changed_again, plant='plant1', source_kind='workbook', sn='1001', certificate_status='Completed')
        self.assertTrue(current_pdf_rows([row])[0]['excel_details_changed_since_issue'])

    def test_restore_puts_all_issued_values_back_and_preserves_completed_status_and_pdf(self):
        from certificate_core import read_batches
        from web_app import current_pdf_rows

        receipt = self.issue_two_material_certificate()
        web_app.complete_workbook_record({'plant':'plant1', 'sn':'1001',
            'expected_revision':workspace_revision(self.root)})
        self.edit_workbook_after_issue()
        issued = web_app.editor_details(receipt['after'])
        changed = web_app.editor_details(self.source())
        self.assertGreaterEqual(len(web_app.compare_details(changed, issued)), 6)
        pdf_path = self.root / receipt['pdf_path']
        pdf_before = pdf_path.read_bytes()

        result = web_app.resolve_issued_excel_changes(self.issued_change_payload('restore'))
        restored = read_batches(self.paths[0])[0]
        self.assertEqual(web_app.compare_details(restored, issued), [])
        self.assertEqual(restored['issued'], 'Completed')
        self.assertEqual(restored['cod_no'], receipt['cod'])
        self.assertEqual(restored['cod_date'].isoformat(), issued['cod_date'])
        self.assertEqual(len(restored['items']), 2)
        self.assertEqual(pdf_path.read_bytes(), pdf_before)
        self.assertEqual(result['action'], 'restore')
        row = dict(restored, plant='plant1', source_kind='workbook', sn='1001', certificate_status='Completed')
        self.assertFalse(current_pdf_rows([row])[0]['excel_details_changed_since_issue'])

    def test_restore_preserves_pending_hdd_lifecycle_metadata(self):
        from certificate_core import read_batches
        from workbook_metadata import write_details

        receipt = self.issue_two_material_certificate()
        self.paths[0].write_bytes(write_details(self.paths[0].read_bytes(), 'Inward Register', '1001', {
            '_recycling_completed':True, '_hdd_cod_no':'0777', '_hdd_cod_date':'2026-10-12'}))
        self.edit_workbook_after_issue()
        self.assertTrue(self.source()['recycling_completed'])

        web_app.resolve_issued_excel_changes(self.issued_change_payload('restore'))
        restored = read_batches(self.paths[0])[0]
        self.assertEqual(restored['issued'], 'Generated')
        self.assertTrue(restored['recycling_completed'])
        self.assertEqual(restored['hdd_cod_no'], '0777')
        self.assertEqual(restored['hdd_cod_date'], '2026-10-12')
        self.assertEqual(web_app.compare_details(restored, web_app.editor_details(receipt['after'])), [])

    def test_stale_review_and_duplicate_serial_fail_without_writing(self):
        from openpyxl import load_workbook

        self.issue_two_material_certificate()
        stale_revision = workspace_revision(self.root)
        self.edit_workbook_after_issue()
        current_bytes = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'files changed'):
            web_app.resolve_issued_excel_changes(self.issued_change_payload('restore', revision=stale_revision))
        self.assertEqual(self.paths[0].read_bytes(), current_bytes)

        book = load_workbook(self.paths[0]); book.active['A4'] = '1001'; book.save(self.paths[0]); book.close()
        duplicate_bytes = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'missing or repeated'):
            web_app.resolve_issued_excel_changes(self.issued_change_payload('keep'))
        self.assertEqual(self.paths[0].read_bytes(), duplicate_bytes)

    def test_restore_refuses_formula_source_cells_without_partial_write(self):
        from openpyxl import load_workbook

        receipt = self.issue_two_material_certificate()
        book = load_workbook(self.paths[0]); book.active['C2'] = '=1+1'; book.save(self.paths[0]); book.close()
        workbook_before = self.paths[0].read_bytes()
        pdf_before = (self.root / receipt['pdf_path']).read_bytes()
        with self.assertRaisesRegex(ValueError, 'formula'):
            web_app.resolve_issued_excel_changes(self.issued_change_payload('restore'))
        self.assertEqual(self.paths[0].read_bytes(), workbook_before)
        self.assertEqual((self.root / receipt['pdf_path']).read_bytes(), pdf_before)
        self.assertFalse((self.root / 'Data' / 'Issued Excel Change Reviews.json').exists())

    def test_damaged_keep_notice_file_fails_open_and_can_be_replaced(self):
        from certificate_core import read_batches
        from issued_detail_changes import REVIEW_FILE
        from web_app import current_pdf_rows

        self.issue_two_material_certificate()
        self.edit_workbook_after_issue()
        review_path = self.root / REVIEW_FILE
        review_path.parent.mkdir(parents=True, exist_ok=True)
        review_path.write_text('{damaged json', encoding='utf-8')
        source = read_batches(self.paths[0])[0]
        row = dict(source, plant='plant1', source_kind='workbook', sn='1001', certificate_status='Generated')
        self.assertTrue(current_pdf_rows([row])[0]['excel_details_changed_since_issue'])

        web_app.resolve_issued_excel_changes(self.issued_change_payload('keep'))
        saved = json.loads(review_path.read_text(encoding='utf-8'))
        self.assertEqual(saved['version'], 1)
        self.assertTrue(saved['reviews'])


if __name__ == '__main__':
    unittest.main()
