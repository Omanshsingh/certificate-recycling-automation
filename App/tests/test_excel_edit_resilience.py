"""Exercise manual-Excel workflows with disposable workbooks and Windows handles."""
from contextlib import contextmanager
from datetime import date
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from certificate_core import read_batches
from issuance_safety import workbook_receipt_conflict
import run_certificates as runner
import web_app
from workbook_store import LockBusyError, WorkbookSession
from workflow_support import save_json
from workspace_review import workspace_revision
from tests.test_certificate_workflow import make_workbook


AS_OF = date(2026, 10, 15)
APP = Path(__file__).resolve().parents[1]


@contextmanager
def saved_workbook_open_for_reading(path):
    """An Excel/preview-style handle: other readers allowed, writers denied."""
    import ctypes
    from ctypes import wintypes
    api = ctypes.WinDLL('kernel32', use_last_error=True)
    api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                               ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    api.CreateFileW.restype = wintypes.HANDLE
    api.CloseHandle.argtypes = [wintypes.HANDLE]
    api.CloseHandle.restype = wintypes.BOOL
    handle = api.CreateFileW(str(path.resolve()), 0x80000000, 1, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        yield
    finally:
        api.CloseHandle(handle)


class ExcelEditResilienceTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='excel-edit-resilience-')))
        self.paths = [self.root / 'Plant 1.xlsx', self.root / 'Plant 2.xlsx']
        make_workbook(self.paths[0], [{'sn': '1001', 'status': '', 'date': date(2026, 10, 1)}])
        make_workbook(self.paths[1], [{'sn': '2001', 'status': 'N/A', 'cod': '0040', 'reference':'REF-OTHER-PLANT'}])
        shutil.copytree(APP / 'templates', self.root / 'App' / 'templates')
        save_json(self.root / 'settings.local.json', {
            'workbook_paths': [str(path) for path in self.paths],
            'certificate_number_prefix': '', 'certificate_number_width': 4,
            'reminders_enabled': False,
        })
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'backups'))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.root / 'Data' / 'Reports' / 'Latest Results.json'))

    def check(self, **kwargs):
        report = {}
        runner.run(self.paths, self.root, check_only=True, as_of=AS_OF, progress=None,
                   report_sink=report, **kwargs)
        return report

    def test_background_refresh_picks_up_saved_excel_edits_and_inserted_or_deleted_rows(self):
        with patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}):
            initial = web_app.refresh()
            self.assertTrue(initial['scan_complete'])
            self.assertTrue(any(row['sn'] == '1001' for row in initial['rows']))

            book = load_workbook(self.paths[0])
            sheet = book['Inward Register']
            sheet['C2'] = 'Updated directly in Excel'
            sheet.insert_rows(3)
            for column, value in enumerate((
                    '1002', date(2026, 10, 5), 'Added directly in Excel', 'Pune',
                    'Printer', 1, 4, 4, 'FORM6-NEW', 'REF-NEW', '', '', 'Ready'), 1):
                sheet.cell(3, column, value)
            book.save(self.paths[0])
            book.close()

            added = web_app.check_workbook_changes()
            self.assertTrue(added['changed'])
            self.assertTrue(added['state']['scan_complete'])
            self.assertTrue(any(row['sn'] == '1001' and row['supplier'] == 'Updated directly in Excel'
                                for row in added['state']['rows']))
            self.assertTrue(any(row['sn'] == '1002' and row['supplier'] == 'Added directly in Excel'
                                for row in added['state']['rows']))

            added_batch = next(batch for batch in read_batches(self.paths[0]) if batch['case_id'] == '1002')
            book = load_workbook(self.paths[0])
            book['Inward Register'].delete_rows(added_batch['source_row'], 1)
            book.save(self.paths[0])
            book.close()

            removed = web_app.check_workbook_changes()
            self.assertTrue(removed['changed'])
            self.assertTrue(removed['state']['scan_complete'])
            self.assertFalse(any(row['sn'] == '1002' for row in removed['state']['rows']))
            self.assertTrue(any(row['sn'] == '1001' and row['supplier'] == 'Updated directly in Excel'
                                for row in removed['state']['rows']))

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_refresh_reads_locked_excel_and_shows_effective_ready_without_saving_it(self):
        before = self.paths[0].read_bytes()
        with saved_workbook_open_for_reading(self.paths[0]):
            report = self.check()
        self.assertTrue(report['scan_complete'])
        row = next(row for row in report['rows'] if row['sn'] == '1001')
        self.assertTrue(row['workbook_read_only'])
        self.assertEqual(row['certificate_status'], 'ready')
        self.assertTrue(any('Showing its saved contents' in warning for warning in row['warnings']))
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertEqual(read_batches(self.paths[0])[0]['issued'], '')
        self.assertEqual(report['created'], 0)

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_invalid_saved_source_shows_alert_even_when_excel_prevents_status_write(self):
        make_workbook(self.paths[0], [{'sn': '1001', 'status': 'Ready', 'weight': -7}])
        before = self.paths[0].read_bytes()
        with saved_workbook_open_for_reading(self.paths[0]):
            report = self.check()
        row = next(row for row in report['rows'] if row['sn'] == '1001')
        self.assertEqual(row['certificate_status'], 'alert')
        self.assertTrue(row['workbook_read_only'])
        self.assertTrue(any('weight' in warning for warning in row['warnings']))
        self.assertEqual(self.paths[0].read_bytes(), before)

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_closing_excel_retries_pending_updates_even_without_new_file_changes(self):
        with patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}):
            with saved_workbook_open_for_reading(self.paths[0]):
                opened = web_app.refresh(self.paths)
                self.assertTrue(any(row.get('workbook_read_only') for row in opened['rows']))
                self.assertNotIn(str(self.root), web_app.WORKBOOK_CHECK_STAMPS)
                saved_stamp = web_app.workbook_check_stamp(self.paths)
            self.assertEqual(web_app.workbook_check_stamp(self.paths), saved_stamp)
            retried = web_app.check_workbook_changes()
            self.assertTrue(retried['changed'])
            self.assertFalse(any(row.get('workbook_read_only') for row in retried['state']['rows']))
            self.assertIn(str(self.root), web_app.WORKBOOK_CHECK_STAMPS)
        self.assertEqual(read_batches(self.paths[0])[0]['issued'], 'Ready')

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_preview_works_with_excel_reader_open_without_reserving_or_writing(self):
        # Excel can still contain an old automatic Alert while saved source
        # corrections are already complete and its status write is pending.
        make_workbook(self.paths[0], [{'sn': '1001', 'status': 'Alert', 'date': date(2026, 10, 1)}])
        before = [path.read_bytes() for path in self.paths]
        details = web_app.editor_details(read_batches(self.paths[0])[0])
        with saved_workbook_open_for_reading(self.paths[0]):
            response = web_app.preview_certificate({'plant': 'plant1', 'source_kind': 'workbook',
                'source_sn': '1001', 'source_plant': 'plant1', 'details': details})
        self.assertTrue(response['preview_url'].endswith('.pdf'))
        self.assertTrue((self.root / 'Data' / 'Previews' / response['preview_url'].split('/')[-1]).is_file())
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_generation_still_requires_write_access_and_cannot_publish_when_locked(self):
        before = self.paths[0].read_bytes()
        report = {}
        with saved_workbook_open_for_reading(self.paths[0]), \
                patch.object(runner, 'WorkbookSession', side_effect=lambda path, backup: WorkbookSession(path, backup, lock_timeout=0)):
            runner.run(self.paths, self.root, as_of=AS_OF, progress=None, allow_warnings=True, report_sink=report)
        self.assertFalse(report['scan_complete'])
        self.assertEqual(report['created'], 0)
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_failed_issuance_names_locked_workbook_and_refresh_recovers_after_close(self):
        def open_without_wait(path, backup_dir, **_):
            return WorkbookSession(path, backup_dir, lock_timeout=0)

        with patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}), \
                patch.object(runner, 'WorkbookSession', side_effect=open_without_wait):
            self.assertTrue(web_app.refresh()['scan_complete'])
            before = self.paths[0].read_bytes()
            stamp = web_app.workbook_check_stamp(self.paths)
            details = web_app.editor_details(read_batches(self.paths[0])[0])
            body = {'plant': 'plant1', 'source_kind': 'workbook', 'source_sn': '1001',
                    'source_plant': 'plant1', 'details': details, 'expected_revision': workspace_revision(self.root)}
            with saved_workbook_open_for_reading(self.paths[0]):
                with self.assertRaisesRegex(ValueError, r'Plant 1\.xlsx is unavailable: .* in use'):
                    web_app.create_certificate(body)
                with self.assertRaisesRegex(ValueError, r'Plant 1\.xlsx is unavailable: .* in use'):
                    web_app.generate(selected=('plant1', '1001'), expected_revision=body['expected_revision'])
            self.assertEqual(self.paths[0].read_bytes(), before)
            self.assertEqual(web_app.workbook_check_stamp(self.paths), stamp)
            self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
            self.assertFalse(web_app.latest_result()['scan_complete'])

            background = web_app.check_workbook_changes()
            self.assertTrue(background['changed'])
            self.assertTrue(background['state']['scan_complete'])
            self.assertFalse(web_app.check_workbook_changes()['changed'])
            body['expected_revision'] = workspace_revision(self.root)
            self.assertEqual(web_app.create_certificate(body)['created'], 1)

    @unittest.skipUnless(os.name == 'nt', 'Windows preserves the workbook timestamp on rename')
    def test_refresh_recovers_when_missing_workbook_returns_to_the_same_path(self):
        workbook = self.paths[1]
        moved = workbook.with_suffix('.moved')
        with patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}):
            self.assertTrue(web_app.refresh()['scan_complete'])
            stamp = web_app.workbook_check_stamp(self.paths)
            workbook.rename(moved)
            try:
                self.assertFalse(web_app.refresh()['scan_complete'])
            finally:
                moved.rename(workbook)
            self.assertEqual(web_app.workbook_check_stamp(self.paths), stamp)
            recovered = web_app.refresh()
            self.assertTrue(recovered['scan_complete'])
            self.assertFalse(recovered.get('already_up_to_date'))
            self.assertTrue(web_app.refresh()['already_up_to_date'])

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_locked_pdf_reissue_keeps_original_and_can_retry_after_closing_viewer(self):
        web_app.refresh()
        details = web_app.editor_details(read_batches(self.paths[0])[0])
        body = {'plant': 'plant1', 'source_kind': 'workbook', 'source_sn': '1001',
                'source_plant': 'plant1', 'details': details}
        self.assertEqual(web_app.create_certificate(body)['created'], 1)
        pdf, = (self.root / 'Certificates').rglob('*.pdf')
        before = {path: path.read_bytes() for path in (pdf, *self.paths)}
        details = web_app.editor_details(read_batches(self.paths[0])[0])
        correction = {**body, 'details': {**details, 'reference': 'CORRECTED-REFERENCE'}}
        with saved_workbook_open_for_reading(pdf):
            with self.assertRaisesRegex(ValueError, r'Could not replace .*\.pdf\. Close the PDF .* then try again'):
                web_app.create_certificate(correction)
        self.assertEqual({path: path.read_bytes() for path in before}, before)
        self.assertEqual(list((self.root / 'Certificates').rglob('*.pdf')), [pdf])
        self.assertEqual(web_app.create_certificate(correction)['created'], 1)
        self.assertNotEqual(pdf.read_bytes(), before[pdf])

    @unittest.skipUnless(os.name == 'nt', 'Windows file sharing behavior')
    def test_refresh_archives_old_named_pdf_after_viewer_releases_it(self):
        with patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}):
            web_app.refresh()
            details = web_app.editor_details(read_batches(self.paths[0])[0])
            body = {'plant': 'plant1', 'source_kind': 'workbook', 'source_sn': '1001',
                    'source_plant': 'plant1', 'details': details}
            self.assertEqual(web_app.create_certificate(body)['created'], 1)
            old_pdf, = (self.root / 'Certificates').rglob('*.pdf')
            checked = web_app.refresh()
            revision = workspace_revision(self.root)
            details = web_app.editor_details(read_batches(self.paths[0])[0])
            correction = {**body, 'expected_revision': revision,
                          'details': {**details, 'supplier': 'Corrected PDF name'}}
            with saved_workbook_open_for_reading(old_pdf), \
                    patch.object(web_app, 'WORKSPACE_STATE_CACHE', (revision, checked)):
                result = web_app.create_certificate(correction)
                self.assertEqual(result['created'], 1)
                self.assertTrue(result['state']['refresh_pending'])
                self.assertIn('Close the PDF', result['state']['recovery_note'])
                self.assertTrue(old_pdf.exists())
            self.assertTrue(web_app.refresh()['scan_complete'])
            self.assertFalse(old_pdf.exists())
            current_pdf, = (self.root / 'Certificates').rglob('*.pdf')
            self.assertIn('Corrected PDF name', current_pdf.name)
            self.assertTrue(list((self.root / 'Data' / 'Archive').rglob(old_pdf.name)))

    def test_duplicate_serials_cannot_be_confirmed_into_a_certificate_or_source_write(self):
        make_workbook(self.paths[0], [{'sn': 'ABC', 'status': 'Generated', 'cod': '0001'},
                                    {'sn': 'abc', 'status': 'Ready'}])
        before = self.paths[0].read_bytes()
        report = {}
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None, allow_warnings=True,
                   selected=('plant1', 'ABC'), details=web_app.editor_details(read_batches(self.paths[0])[0]),
                   report_sink=report)
        self.assertEqual(report['created'], 0)
        self.assertTrue(any('Serial Number ABC is repeated' in row['detail'] for row in report['rows']))
        self.assertEqual(self.paths[0].read_bytes(), before)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))

    def test_unreadable_register_can_still_be_opened_for_manual_repair(self):
        cached = {'plant': 'plant1', 'sn': '1001', 'file': str(self.paths[0]), 'sheet': 'Inward Register'}
        with patch.object(web_app, 'latest_result', return_value={'rows': [cached]}), \
                patch.object(web_app, 'read_batches', side_effect=ValueError('Invalid column layout')), \
                patch.object(web_app, 'open_excel_workbook') as native:
            opened = web_app.open_selected_workbook_record('plant1', '1001')
        self.assertEqual(opened, (self.paths[0], 'Inward Register', None))
        native.assert_called_once_with(self.paths[0])

    def test_duplicate_records_do_not_inherit_each_others_issuance_receipt(self):
        rows = [{'source_kind': 'workbook', 'plant': 'plant1', 'sn': '1001',
                 'certificate_status': 'generated', 'duplicate_serial_number': True,
                 'last_certificate_details': {'supplier': 'Wrong inherited company'}}]
        receipt = {'source_kind': 'workbook', 'plant': 'plant1', 'key': '1001',
                   'state': 'complete', 'after': {'supplier': 'First company'}}
        with patch.object(web_app, 'audit_records', return_value=[receipt]):
            result = web_app.current_pdf_rows(rows)
        self.assertNotIn('last_certificate_details', result[0])
        self.assertNotIn('has_issued_certificate', result[0])
        self.assertNotIn('pdf_relative_path', result[0])

    def test_reused_serial_cannot_restore_a_different_inwards_history(self):
        source = read_batches(self.paths[0])[0]
        prior = dict(source, supplier='Unrelated old company')
        receipt = {'version': 1, 'id': 'b' * 32, 'state': 'complete', 'source_kind': 'workbook',
                   'plant': 'plant1', 'key': '1001', 'cod': '0001',
                   'pdf_path': 'Certificates/Auto Generated/0001_Old_Pune_1.pdf',
                   'pdf_sha256': '0' * 64, 'before': prior, 'after': prior}
        save_json(self.root / 'Data' / 'Issuance History' / 'old.json', receipt)
        before = self.paths[0].read_bytes()
        details = web_app.editor_details(source)
        with self.assertRaisesRegex(ValueError, 'Serial Number 1001 is linked to an earlier certificate'):
            web_app.workbook_form({'plant': 'plant1'}, details, self.paths)
        self.assertEqual(self.paths[0].read_bytes(), before)
        report = {}
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None, allow_warnings=True, report_sink=report)
        row = next(row for row in report['rows'] if row['sn'] == '1001')
        self.assertTrue(row['record_identity_conflict'])
        self.assertTrue(row['issuance_blocked'])
        self.assertEqual(report['created'], 0)
        # Refresh may maintain borders and conditional formatting. Every inward
        # value, including its status, must stay unchanged on the conflict.
        from certificate_core import batch_content
        current=read_batches(self.paths[0])[0]
        self.assertEqual(batch_content(current), batch_content(source))
        for key in ('issued', 'cod_no', 'cod_date'):
            self.assertEqual(current[key], source[key])
        before = self.paths[0].read_bytes()
        enriched = web_app.current_pdf_rows([row])[0]
        self.assertNotIn('last_certificate_details', enriched)
        self.assertNotIn('has_issued_certificate', enriched)
        with self.assertRaisesRegex(ValueError, 'linked to an earlier certificate'):
            web_app.complete_workbook_record({'plant': 'plant1', 'sn': '1001',
                                              'expected_revision': workspace_revision(self.root)})
        self.assertEqual(self.paths[0].read_bytes(), before)

    def test_receipt_requires_same_cod_or_matching_source_facts_when_cod_is_blank(self):
        source = read_batches(self.paths[0])[0]
        receipt = {'cod': '0001', 'before': source, 'after': dict(source, supplier='Printed company')}
        self.assertEqual(workbook_receipt_conflict(source, receipt), '')
        self.assertEqual(workbook_receipt_conflict(dict(source, cod_no='1', supplier='Excel corrected name'), receipt), '')
        self.assertTrue(workbook_receipt_conflict(dict(source, cod_no='0002'), receipt))
        self.assertTrue(workbook_receipt_conflict(dict(source, supplier='Completely different inward'), receipt))

    def test_generated_record_can_be_completed_after_live_ownership_check(self):
        report = {}
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None, report_sink=report)
        self.assertEqual(report['created'], 1)
        source = read_batches(self.paths[0])[0]
        result = web_app.complete_workbook_record({'plant': 'plant1', 'sn': '1001',
                                                  'expected_revision': workspace_revision(self.root)})
        completed = read_batches(self.paths[0])[0]
        self.assertEqual(completed['issued'], 'Completed')
        self.assertEqual(completed['cod_no'], source['cod_no'])
        self.assertTrue(result['scan_complete'])
        self.assertEqual(web_app.target_row(result, ('plant1','1001'))['certificate_status'], 'completed')

    def test_completion_updates_the_checked_dashboard_without_a_second_workbook_scan(self):
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None)
        state = web_app.workspace_state(recheck_obsolete=False)
        state['rows'].append({'source_kind':'manual', 'record_id':'keep-manual-row', 'plant':'plant2',
                              'sn':'Custom 123', 'certificate_status':'ready'})
        expected = workspace_revision(self.root)
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (expected, state)), \
                patch.object(web_app, 'refresh', side_effect=AssertionError('A full workbook refresh should not run after completion')):
            result = web_app.complete_workbook_record({'plant':'plant1', 'sn':'1001', 'expected_revision':expected})
        self.assertEqual(web_app.target_row(result, ('plant1','1001'))['certificate_status'], 'completed')
        self.assertTrue(web_app.target_row(result, ('plant1','1001'))['pdf_filename'])
        self.assertTrue(any(row.get('record_id') == 'keep-manual-row' for row in result['rows']))

    def test_completion_refresh_failure_preserves_success_and_actual_completed_status(self):
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None)
        expected = workspace_revision(self.root)
        with patch.object(web_app, 'refresh', side_effect=OSError('Refresh connection interrupted')):
            result = web_app.complete_workbook_record({'plant': 'plant1', 'sn': '1001',
                                                      'expected_revision': expected})
        self.assertEqual(read_batches(self.paths[0])[0]['issued'], 'Completed')
        self.assertIn('Completed status saved successfully', result['setup_warning'])
        self.assertEqual(web_app.target_row(result, ('plant1','1001'))['certificate_status'], 'completed')
        self.assertFalse(web_app.target_row(result, ('plant1','1001'))['reminder_due'])
        self.assertEqual(web_app.target_row(web_app.latest_result(), ('plant1','1001'))['certificate_status'], 'completed')

    def test_completion_acknowledgement_survives_both_cache_and_refresh_failure(self):
        runner.run(self.paths, self.root, as_of=AS_OF, progress=None)
        expected = workspace_revision(self.root)
        with patch.object(web_app, 'save_json', side_effect=OSError('Cache unavailable')), \
                patch.object(web_app, 'refresh', side_effect=OSError('Refresh unavailable')):
            result = web_app.complete_workbook_record({'plant':'plant1', 'sn':'1001', 'expected_revision':expected})
        self.assertEqual(read_batches(self.paths[0])[0]['issued'], 'Completed')
        self.assertEqual(web_app.target_row(result, ('plant1','1001'))['certificate_status'], 'completed')
        self.assertTrue(result['refresh_pending'])
        self.assertFalse(result['scan_complete'])
        self.assertIn('Completed status saved successfully',result['setup_warning'])


if __name__ == '__main__':
    unittest.main()
