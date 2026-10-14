"""First-generation locks and later corrections, with disposable workbooks/PDFs."""
from copy import deepcopy
import ctypes
from datetime import date
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch, Mock

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter

from certificate_core import read_batches
from issuance_safety import latest_issuance
import run_certificates as runner
import web_app
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_workflow import make_workbook

APP = Path(__file__).resolve().parents[1]


class SimpleCertificateFlowTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='simple-certificate-')))
        self.paths = [self.root / 'plant1.xlsx', self.root / 'plant2.xlsx']
        make_workbook(self.paths[0], [{'sn':'1001', 'supplier':'Different inward name',
            'location':'Pune site', 'date':date(2026, 9, 1), 'weight':12.345678901,
            'quantity':1, 'per_piece_weight':12.345678901, 'form6':''}])
        make_workbook(self.paths[1], [{'sn':'2001', 'status':'N/A'}])
        shutil.copytree(APP / 'templates', self.root / 'App' / 'templates')
        save_json(self.root / 'settings.local.json', {'workbook_paths':[str(path) for path in self.paths],
            'certificate_number_prefix':'', 'certificate_number_width':4,
            'reminders_enabled':False, 'reminder_time':'11:00'})
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'Data' / 'Workbook Backups'))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.root / 'Data' / 'Reports' / 'Latest Results.json'))
        self.enterContext(patch.object(web_app, 'refresh', side_effect=lambda *_: web_app.workspace_state(False)))

    def source(self):
        return read_batches(self.paths[0])[0]

    def body(self, kind='sita', **changes):
        details = web_app.editor_details(self.source())
        details['certificate_type'] = kind
        details['certificate_fields'] = {'model_name':'Scanner', 'model_serial':'M-123', 'tube_serial':'T-456'} if kind == 'siemens' else {}
        details.update(changes)
        return {'plant':'plant1', 'source_sn':'1001', 'source_plant':'plant1', 'source_kind':'workbook',
                'details':details, 'edit_scope':'update_excel', 'expected_revision':workspace_revision(self.root)}

    def create(self, body=None):
        return web_app.create_certificate(self.body() if body is None else body)

    def cells(self):
        book = load_workbook(self.paths[0])
        try:
            return list(book.active.values)
        finally:
            book.close()

    def test_first_sita_uses_excel_facts_no_form6_and_never_saves_source_edits(self):
        before, other = self.cells(), self.paths[1].read_bytes()
        response = self.create()
        self.assertEqual(response['created'], 1)
        after = self.cells()
        for row, original in enumerate(before):
            for column, value in enumerate(original):
                if row == 1 and column in {10, 11, 12}:
                    continue
                self.assertEqual(after[row][column], value)
        self.assertEqual(self.paths[1].read_bytes(), other)
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        self.assertEqual(receipt['after']['certificate_type'], 'sita')
        self.assertEqual(receipt['after']['edit_scope'], 'certificate_only')
        self.assertFalse(any('Form 6' in warning for warning in receipt['accepted_warnings']))
        text = PdfReader(self.root / receipt['pdf_path']).pages[0].extract_text()
        self.assertIn('SITA INFORMATION', text)
        row = web_app.target_row(response['state'], ('plant1', '1001'))
        self.assertTrue(row['has_issued_certificate'])
        self.assertEqual(row['last_certificate_details']['certificate_type'], 'sita')
        self.assertEqual(self.source()['certificate_type'], 'standard')

    def test_preview_and_forged_first_generation_edits_are_readonly_and_rejected(self):
        original = self.paths[0].read_bytes()
        for edits in ({'supplier':'Changed'}, {'weight':'99'}, {'cod_no':'0099'}, {'certificate_weight':'99'}):
            body = self.body(**edits)
            if 'certificate_weight' in edits:
                body['use_certificate_weight'] = True
            with self.subTest(edits=edits):
                with self.assertRaisesRegex(ValueError, 'locked|first generation'):
                    self.create(body)
                with self.assertRaises(ValueError):
                    web_app.preview_certificate(body)
                self.assertEqual(self.paths[0].read_bytes(), original)
        preview = web_app.preview_certificate(self.body())
        self.assertTrue(preview['preview_image_url'])
        self.assertEqual(self.paths[0].read_bytes(), original)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertIsNone(latest_issuance(self.root, 'workbook', 'plant1', '1001'))

    def test_first_generation_pdf_edits_are_printed_without_changing_inward_facts(self):
        before = self.cells()
        body = self.body('standard', supplier='Corrected PDF company', reference='PDF-REF-42',
                         location='Pune, Maharashtra, India')
        body.update(edit_scope='certificate_only', pdf_only_edits=True,
                    confirmed=True, override_reason='Checked PDF-only correction.')
        preview = web_app.preview_certificate(body)
        self.assertEqual(preview['printed_details']['company_display'], 'CORRECTED PDF COMPANY, PUNE, MAHARASHTRA, INDIA')
        preview_pdf = self.root / 'Data' / 'Previews' / Path(preview['preview_url']).name
        self.assertIn('PUNE, MAHARASHTRA, INDIA', ' '.join(PdfReader(preview_pdf).pages[0].extract_text().split()))
        response = self.create(body)
        self.assertEqual(response['created'], 1)
        after = self.cells()
        for row, original in enumerate(before):
            for column, value in enumerate(original):
                if row == 1 and column in {10, 11, 12}:
                    continue
                self.assertEqual(after[row][column], value)
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        text = PdfReader(self.root / receipt['pdf_path']).pages[0].extract_text()
        self.assertIn('CORRECTED PDF COMPANY', text)
        self.assertIn('PDF-REF-42', text)
        self.assertIn('PUNE, MAHARASHTRA, INDIA', ' '.join(text.split()))
        self.assertEqual(receipt['after']['supplier'], 'Corrected PDF company')
        self.assertEqual(self.source()['supplier'], 'Different inward name')
        self.assertEqual(self.source()['location'], 'Pune site')

    def test_special_type_name_can_be_edited_through_main_company_field_without_an_override_input(self):
        for kind in ('sita', 'siemens'):
            with self.subTest(kind=kind):
                before = self.paths[0].read_bytes()
                body = self.body(kind, supplier='Updated company', location='Pune, Maharashtra, India',
                                 certificate_fields={'company_display':''})
                body.update(pdf_only_edits=True, edit_scope='certificate_only')
                preview = web_app.preview_certificate(body)
                self.assertEqual(preview['printed_details']['company_display'], 'UPDATED COMPANY, PUNE, MAHARASHTRA, INDIA')
                self.assertEqual(self.paths[0].read_bytes(), before)

    def test_local_pdf_open_and_reveal_accept_only_valid_certificate_paths(self):
        folder = self.root / 'Certificates' / 'Created in App'
        folder.mkdir(parents=True)
        pdf = folder / '0001_Test.pdf'
        writer = PdfWriter(); writer.add_blank_page(width=200, height=200)
        with pdf.open('wb') as stream:
            writer.write(stream)
        with (patch.object(web_app, 'ROOT', self.root),
              patch.object(web_app, 'open_path') as open_path,
              patch.object(web_app, 'reveal_file_in_explorer') as select_file):
            self.assertEqual(web_app.open_certificate_locally('Created in App/0001_Test.pdf'), pdf)
            select_file.assert_called_once_with(pdf)
            open_path.assert_called_once_with(pdf)
            select_file.reset_mock(); open_path.reset_mock()
            web_app.open_certificate_locally('Created in App/0001_Test.pdf', open_document=False)
            select_file.assert_called_once_with(pdf)
            open_path.assert_not_called()
            select_file.reset_mock()
            with self.assertRaises(ValueError):
                web_app.open_certificate_locally('../outside.pdf')
            select_file.assert_not_called(); open_path.assert_not_called()
        with self.assertRaises(ValueError):
            web_app.open_certificate_locally('../outside.pdf')

    @unittest.skipUnless(hasattr(ctypes, 'WinDLL'), 'Windows Shell API')
    def test_native_selection_preserves_comma_paths_and_reports_shell_failures(self):
        path = self.root / 'Folder, Mumbai' / '0001_Company, India.pdf'
        for result in (0, -2147467259):
            with self.subTest(result=result):
                ole, shell = Mock(), Mock()
                ole.CoInitializeEx.return_value = 0
                def parse_file(_name, _binding, pointer, _flags, _attributes):
                    ctypes.cast(pointer, ctypes.POINTER(ctypes.c_void_p))[0] = ctypes.c_void_p(1234)
                    return 0
                shell.SHParseDisplayName.side_effect = parse_file
                shell.SHOpenFolderAndSelectItems.return_value = result
                with patch('ctypes.WinDLL', side_effect=lambda library: ole if library == 'ole32' else shell):
                    if result < 0:
                        with self.assertRaisesRegex(OSError, 'could not select this PDF'):
                            web_app.reveal_file_in_explorer(path)
                    else:
                        web_app.reveal_file_in_explorer(path)
                self.assertEqual(shell.SHParseDisplayName.call_args.args[0], str(path))
                selection = shell.SHOpenFolderAndSelectItems.call_args.args
                self.assertEqual(selection[0].value, 1234)
                self.assertEqual(selection[1:], (0, None, 0))
                ole.CoTaskMemFree.assert_called_once()
                ole.CoUninitialize.assert_called_once()

    def test_siemens_autofills_excel_site_and_preserves_three_manual_fields(self):
        body = self.body('siemens')
        body['details']['certificate_fields']['site_location'] = 'Forged site'
        self.assertEqual(self.create(body)['created'], 1)
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        fields = receipt['after']['certificate_fields']
        self.assertEqual(fields['site_location'], 'Pune site')
        for key, expected in [('model_name','Scanner'), ('model_serial','M-123'), ('tube_serial','T-456')]:
            self.assertEqual(fields[key], expected)
        text = PdfReader(self.root / receipt['pdf_path']).pages[0].extract_text()
        for value in ['Pune site', 'Scanner', 'M-123', 'T-456']:
            self.assertIn(value, text)

    def test_preview_metadata_matches_pdf_with_saved_and_automatic_cod_date(self):
        for saved in (False, True):
            with self.subTest(saved=saved):
                if saved:
                    book = load_workbook(self.paths[0])
                    book.active['K2'] = '0362'
                    book.active['L2'] = date(2026, 9, 30)
                    book.save(self.paths[0]); book.close()
                original = self.paths[0].read_bytes()
                preview = web_app.preview_certificate(self.body())
                issued = date.fromisoformat(preview['certificate_date'])
                if saved:
                    self.assertEqual(preview['certificate_no'], '0362')
                    self.assertEqual(issued, date(2026, 9, 30))
                else:
                    self.assertEqual(issued, date.today())
                pdf = self.root / 'Data' / 'Previews' / Path(preview['preview_url']).name
                text = PdfReader(pdf).pages[0].extract_text()
                for expected in [preview['certificate_no'], issued.strftime('%d.%m.%Y'),
                                 *preview['printed_details'].values(), self.source()['reference']]:
                    self.assertIn(expected, text)
                self.assertEqual(self.paths[0].read_bytes(), original)
                self.assertIsNone(latest_issuance(self.root, 'workbook', 'plant1', '1001'))

    def test_first_standard_cannot_change_company_through_a_printed_name_override(self):
        body = self.body('standard')
        body['details']['certificate_fields'] = {'company_display':'Forged company'}
        before = self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError, 'company name is locked'):
            self.create(body)
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_missing_excel_site_allows_manual_siemens_location(self):
        book = load_workbook(self.paths[0])
        book.active['D2'] = None
        book.save(self.paths[0]); book.close()
        body = self.body('siemens')
        body['details']['certificate_fields']['site_location'] = 'Manual site'
        self.assertEqual(self.create(body)['created'], 1)
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '1001')['after']['certificate_fields']['site_location'], 'Manual site')

    def correction(self):
        first = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        return {'plant':'plant1', 'source_sn':'1001', 'source_plant':'plant1',
                'details':web_app.editor_details(first['after']), 'expected_revision':workspace_revision(self.root)}

    def test_pdf_only_correction_reopens_latest_details_keeps_cod_and_archives_old_pdf(self):
        self.create()
        first = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        original = deepcopy(self.source())
        body = self.correction()
        body['details'].update(supplier='Corrected name', location='Mumbai', cod_no='', cod_date='')
        response = self.create(body)  # Omitted scope defaults to PDF-only.
        self.assertEqual(response['created'], 1)
        self.assertEqual(self.source()['supplier'], original['supplier'])
        last = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        self.assertEqual(last['cod'], first['cod'])
        self.assertEqual(last['certificate_date'], first['certificate_date'])
        row = web_app.target_row(response['state'], ('plant1', '1001'))
        self.assertEqual(row['last_certificate_details']['supplier'], 'Corrected name')
        self.assertEqual(row['last_certificate_details']['location'], 'Mumbai')
        self.assertTrue(list((self.root / 'Data' / 'Archive' / 'Superseded Certificates').glob('*.pdf')))

    def test_yes_updates_only_selected_inward_and_no_correction_save_on_failure(self):
        self.create()
        body = self.correction()
        body['edit_scope'] = 'update_excel'
        body['details']['supplier'] = 'Corrected Excel name'
        original, other = self.paths[0].read_bytes(), self.paths[1].read_bytes()
        with patch.object(runner, 'make_certificate', side_effect=OSError('render failed')):
            with self.assertRaisesRegex(ValueError, 'render failed'):
                self.create(body)
        self.assertEqual(self.paths[0].read_bytes(), original)
        self.assertEqual(self.create(body)['created'], 1)
        self.assertEqual(self.source()['supplier'], 'Corrected Excel name')
        self.assertEqual(self.source()['certificate_type'], 'sita')
        self.assertEqual(self.paths[1].read_bytes(), other)

    def test_missing_pdf_still_has_editable_issued_snapshot(self):
        self.create(self.body('siemens'))
        receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        (self.root / receipt['pdf_path']).unlink()
        runner.run(self.paths, self.root, check_only=True, progress=None)
        row = web_app.target_row(web_app.workspace_state(False), ('plant1', '1001'))
        self.assertEqual(row['certificate_status'], 'ready')
        self.assertTrue(row['has_issued_certificate'])
        self.assertEqual(row['last_certificate_details']['certificate_type'], 'siemens')

    def test_bulk_recreation_retains_app_only_type_and_siemens_identifiers(self):
        self.create(self.body('siemens'))
        first = latest_issuance(self.root, 'workbook', 'plant1', '1001')
        (self.root / first['pdf_path']).unlink()
        runner.run(self.paths, self.root, check_only=True, progress=None)
        for attempt in range(2):
            with self.subTest(attempt=attempt):
                self.assertEqual(runner.run(self.paths, self.root, plants=['plant1'], progress=None)[0], 1)
                receipt = latest_issuance(self.root, 'workbook', 'plant1', '1001')
                self.assertEqual(receipt['after']['certificate_type'], 'siemens')
                self.assertEqual(receipt['after']['certificate_fields']['model_serial'], 'M-123')
                self.assertEqual(receipt['after']['edit_scope'], 'certificate_only')
                self.assertEqual(self.source()['certificate_type'], 'standard')
                (self.root / receipt['pdf_path']).unlink()
                runner.run(self.paths, self.root, check_only=True, progress=None)


if __name__ == '__main__':
    unittest.main()
