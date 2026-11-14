"""Current-PDF response metadata, using disposable files and no issuance."""
from copy import deepcopy
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from urllib.parse import quote

from pypdf import PdfWriter

import certificate_drafts as drafts
from issuance_safety import file_digest
import web_app
from workflow_support import run_lock, save_json


class CurrentPdfMetadataTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='certificate-metadata-test-')
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.folder = self.root / 'Certificates'
        self.folder.mkdir()
        self.results = self.root / 'Data' / 'Reports' / 'Latest Results.json'
        self.enterContext(patch.object(web_app, 'ROOT', self.root))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.results))
        self.enterContext(patch.object(web_app, 'BACKUP_DIR', self.root / 'Data' / 'Workbook Backups'))
        # These tests must never issue a certificate or open an actual workbook.
        self.enterContext(patch.object(web_app, 'run', side_effect=AssertionError('No workbook operation')))
        self.enterContext(patch.object(web_app, 'make_template_certificate', side_effect=AssertionError('No issuance')))
        # Match the existing workspace lifecycle before checking durable bytes.
        with run_lock(self.root):
            pass
        self.records = []
        save_json(self.root / 'settings.local.json', {'workbook_paths': [str(self.root/'plant1.xlsx'), str(self.root/'plant2.xlsx')]})

    def pdf(self, name, *, pages=1, encrypted=False):
        path = self.folder / name
        path.parent.mkdir(parents=True, exist_ok=True)
        writer = PdfWriter()
        for _ in range(pages):
            writer.add_blank_page(width=300, height=400)
        if encrypted:
            writer.encrypt('fixture-password')
        writer.write(path)
        return path

    def workbook_row(self, *, status='generated', cod='0001', sn='CLIENT', **extra):
        return dict(source_kind='workbook', plant='plant1', sn=sn,
                    certificate_status=status, certificate_no=cod,
                    certificate_date='2026-09-18', supplier='Renamed company', **extra)

    def manual_record(self, filename, *, status='generated_manual', record_id='manual-1', cod='0002',
                      relative_path=None):
        record = dict(id=record_id, plant='plant2', status=status,
                      details=dict(case_id='MANUAL', supplier='Manual company', certificate_type='sita',
                                   certificate_fields={}, cod_no=cod, cod_date='2026-09-18',
                                   certificate_weight='8.5', inward_date='2026-09-01'))
        if filename is not None:
            record['pdf_filename'] = filename
        if relative_path is not None:
            record['pdf_relative_path'] = relative_path
        self.records = [entry for entry in self.records if entry['id'] != record_id] + [record]
        save_json(drafts.register_path(self.root), dict(version=1, records=self.records))
        return record

    def manual_row(self, record):
        return next(row for row in drafts.manual_rows(self.root) if row['record_id'] == record['id'])

    def receipt(self, path, *, source_kind='workbook', plant='plant1', key='CLIENT',
                cod='0001', order=1, state='complete'):
        identifier = f'{source_kind}-{order}'
        value = dict(version=1, id=identifier, state=state, source_kind=source_kind,
                     plant=plant, key=key, cod=cod, started_at=f'2026-09-{order:02}T10:00:00Z',
                     pdf_path=path.relative_to(self.root).as_posix(), pdf_sha256=file_digest(path))
        save_json(self.root / 'Data' / 'Issuance History' / f'{identifier}.json', value)
        return value

    def cache(self, rows):
        save_json(self.results, dict(rows=rows, scan_complete=True, created=0, review=0))

    def snapshot(self):
        return {path.relative_to(self.root).as_posix(): path.read_bytes()
                for path in self.root.rglob('*') if path.is_file()}

    def test_workspace_enriches_workbook_and_manual_issued_statuses_without_tracking_writes(self):
        workbook = self.pdf('0001_Previous_company_Pune_1.pdf')
        manual = self.pdf('0002_Manual company #1%_2.pdf')
        self.receipt(workbook)
        self.receipt(manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        for workbook_status in ('generated', 'generated_manual', 'completed'):
            for manual_status in ('generated_manual', 'completed'):
                with self.subTest(workbook=workbook_status, manual=manual_status):
                    self.manual_record(manual.name, status=manual_status)
                    self.cache([self.workbook_row(status=workbook_status, pdf_filename='obsolete.pdf')])
                    before = self.snapshot()
                    state = web_app.workspace_state()
                    rows = state['rows']
                    self.assertEqual([row['pdf_filename'] for row in rows], [workbook.name, manual.name])
                    expected_manual = 'generated' if manual_status == 'generated_manual' else manual_status
                    self.assertEqual([row['certificate_status'] for row in rows], [workbook_status, expected_manual])
                    self.assertEqual(self.snapshot(), before)
                    # Validated row metadata also works with the unchanged success formatter.
                    outputs = web_app.certificate_outputs(rows)
                    self.assertEqual([entry['filename'] for entry in outputs], [workbook.name, manual.name])
                    self.assertEqual(outputs[1]['pdf_url'], '/api/certificates/' + quote(manual.name, safe=''))

    def test_category_subfolders_keep_saved_pdf_links_and_manual_register_in_sync(self):
        workbook = self.pdf('Auto Generated/0001_Workbook_Client_Pune_1.pdf')
        manual = self.pdf('Auto Generated/0002_Manual company_Pune_2.pdf')
        self.receipt(workbook)
        self.receipt(manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        record = self.manual_record(manual.name, relative_path='Auto Generated/' + manual.name)
        self.cache([self.workbook_row()])

        rows = web_app.workspace_state()['rows']
        outputs = web_app.certificate_outputs(rows)

        self.assertEqual([row['pdf_relative_path'] for row in rows], [
            'Auto Generated/' + workbook.name,
            'Auto Generated/' + manual.name,
        ])
        self.assertEqual([output['relative_path'] for output in outputs], [
            'Auto Generated/' + workbook.name,
            'Auto Generated/' + manual.name,
        ])
        self.assertEqual(outputs[0]['pdf_url'], '/api/certificates/' + quote('Auto Generated/' + workbook.name, safe='/'))
        self.assertEqual(outputs[1]['pdf_url'], '/api/certificates/' + quote('Auto Generated/' + manual.name, safe='/'))
        self.assertEqual(drafts.read_register(self.root)['records'][0], dict(record, plant=None))

    def test_latest_reissue_wins_over_obsolete_files_and_cached_filenames(self):
        old_workbook = self.pdf('0001_Old_company_1.pdf')
        current_workbook = self.pdf('0001_New_company_1.pdf', pages=2)
        old_manual = self.pdf('0002_Old_manual_2.pdf')
        current_manual = self.pdf('0002_New_manual_2.pdf', pages=2)
        self.receipt(old_workbook)
        self.receipt(current_workbook, order=2)
        self.receipt(old_manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        self.receipt(current_manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002', order=2)
        record = self.manual_record(current_manual.name)
        self.cache([self.workbook_row(pdf_filename=old_workbook.name)])
        before = self.snapshot()
        rows = web_app.workspace_state()['rows']
        self.assertEqual([row['pdf_filename'] for row in rows], [current_workbook.name, current_manual.name])
        self.assertEqual(self.snapshot(), before)
        self.assertEqual(drafts.read_register(self.root)['records'][0], dict(record, plant=None))

    def test_missing_latest_is_not_replaced_by_an_obsolete_same_cod_file(self):
        old_workbook = self.pdf('0001_Old_company_1.pdf')
        current_workbook = self.pdf('0001_New_company_1.pdf')
        old_manual = self.pdf('0002_Old_manual_2.pdf')
        current_manual = self.pdf('0002_New_manual_2.pdf')
        self.receipt(old_workbook)
        self.receipt(current_workbook, order=2)
        self.receipt(old_manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        self.receipt(current_manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002', order=2)
        record = self.manual_record(current_manual.name)
        manual_row = self.manual_row(record)
        manual_row['pdf_filename'] = old_manual.name
        rows = [self.workbook_row(pdf_filename=old_workbook.name), manual_row]
        current_workbook.unlink()
        current_manual.unlink()
        before, original = self.snapshot(), deepcopy(rows)
        enriched = web_app.current_pdf_rows(rows)
        self.assertTrue(all('pdf_filename' not in row for row in enriched))
        self.assertEqual(rows, original)
        self.assertEqual(self.snapshot(), before)

    def test_record_pointing_to_obsolete_manual_file_cannot_override_latest_receipt(self):
        old = self.pdf('0002_Old_manual_2.pdf')
        current = self.pdf('0002_New_manual_2.pdf')
        self.receipt(old, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        self.receipt(current, source_kind='manual', plant='plant2', key='manual-1', cod='0002', order=2)
        row = self.manual_row(self.manual_record(old.name))
        before = self.snapshot()
        self.assertNotIn('pdf_filename', web_app.current_pdf_rows([row])[0])
        self.assertEqual(self.snapshot(), before)

    def test_legacy_manual_without_saved_filename_still_checks_receipt_hash(self):
        current = self.pdf('0002_Legacy_manual_2.pdf')
        self.receipt(current, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        row = self.manual_row(self.manual_record(None))
        self.assertEqual(web_app.current_pdf_rows([row])[0]['pdf_filename'], current.name)
        self.pdf(current.name, pages=2)
        self.assertNotIn('pdf_filename', web_app.current_pdf_rows([row])[0])

    def test_changed_current_pdf_hash_does_not_fall_back_to_obsolete_files(self):
        current_workbook = self.pdf('0001_Current_1.pdf')
        current_manual = self.pdf('0002_Current_2.pdf')
        self.pdf('0001_Obsolete_1.pdf')
        self.pdf('0002_Obsolete_2.pdf')
        self.receipt(current_workbook)
        self.receipt(current_manual, source_kind='manual', plant='plant2', key='manual-1', cod='0002')
        rows = [self.workbook_row(), self.manual_row(self.manual_record(current_manual.name))]
        self.pdf(current_workbook.name, pages=2)
        self.pdf(current_manual.name, pages=2)
        before = self.snapshot()
        self.assertTrue(all('pdf_filename' not in row for row in web_app.current_pdf_rows(rows)))
        self.assertEqual(self.snapshot(), before)

    def test_missing_workbook_keeps_existing_refresh_recovery_path(self):
        self.cache([self.workbook_row(status='completed', pdf_filename='missing.pdf')])
        recovered = dict(rows=[self.workbook_row(status='ready')], scan_complete=True)
        before = self.snapshot()
        with patch.object(web_app, 'refresh', return_value=recovered) as refresh:
            self.assertIs(web_app.workspace_state(), recovered)
        refresh.assert_called_once_with(force=True)
        self.assertEqual(self.snapshot(), before)

    def test_missing_manual_keeps_recovery_identity_and_returns_no_pdf_metadata(self):
        record = self.manual_record('missing.pdf', status='completed')
        self.cache([])
        state = web_app.workspace_state()
        row = state['rows'][0]
        self.assertEqual((row['record_id'], row['certificate_no'], row['certificate_date']),
                         (record['id'], '0002', '2026-09-18'))
        self.assertEqual(row['certificate_status'], 'ready')
        self.assertNotIn('pdf_filename', row)
        self.assertEqual(drafts.read_register(self.root)['records'][0]['status'], 'ready')
        before = self.snapshot()
        web_app.workspace_state()
        self.assertEqual(self.snapshot(), before, 'Repeated recovery must remain idempotent')

    def test_refresh_with_disconnected_workbooks_keeps_valid_manual_pdf_metadata(self):
        current = self.pdf('0002_Manual_2.pdf')
        self.manual_record(current.name)
        with patch.object(web_app, 'configured_paths', side_effect=ValueError('Excel disconnected')):
            state = web_app.refresh()
        self.assertFalse(state['scan_complete'])
        self.assertEqual(state['setup_warning'], 'Excel disconnected')
        self.assertEqual(state['rows'][0]['pdf_filename'], current.name)

    def test_wrong_plant_or_missing_manual_record_cannot_supply_a_link(self):
        self.pdf('0001_Other_plant_2.pdf')
        self.pdf('0002_Manual_2.pdf')
        record = self.manual_record('0002_Manual_2.pdf')
        manual_row = self.manual_row(record)
        for extra in ({'record_id': 'absent'}, {'plant': 'plant1'}):
            with self.subTest(extra=extra):
                rows = [self.workbook_row(), dict(manual_row, **extra)]
                self.assertTrue(all('pdf_filename' not in row for row in web_app.current_pdf_rows(rows)))

    def test_unissued_rows_discard_cached_metadata_even_when_pdf_exists(self):
        self.pdf('0001_Current_1.pdf')
        rows = [self.workbook_row(status=status, pdf_filename='0001_Current_1.pdf')
                for status in ('ready', 'alert', 'blank', 'na', 'draft', 'unknown')]
        before, original = self.snapshot(), deepcopy(rows)
        self.assertTrue(all('pdf_filename' not in row for row in web_app.current_pdf_rows(rows)))
        self.assertEqual(rows, original)
        self.assertEqual(self.snapshot(), before)

    def test_non_path_lookup_results_omit_metadata_without_mutating_rows_or_tracking(self):
        current = self.pdf('0002_Current_2.pdf')
        record = self.manual_record(current.name)
        rows = [self.workbook_row(pdf_filename='stale.pdf'),
                dict(self.manual_row(record), pdf_filename='stale.pdf')]
        original, before = deepcopy(rows), self.snapshot()
        for lookup in (True, False, None, 123, object(), {'path': current}, [current], b'current.pdf'):
            with self.subTest(lookup=lookup), \
                    patch.object(web_app, 'find_certificate', return_value=lookup), \
                    patch.object(web_app, 'current_record_pdf', return_value=lookup):
                enriched = web_app.current_pdf_rows(rows)
                self.assertTrue(all('pdf_filename' not in row for row in enriched))
                self.assertEqual([row['certificate_status'] for row in enriched],
                                 ['generated', 'generated_manual'])
                self.assertEqual(rows, original)
                self.assertEqual(self.snapshot(), before)

    def test_unsafe_manual_saved_paths_are_not_exposed(self):
        outside = self.pdf('../Outside.pdf')
        nested = self.pdf('nested/Inside.pdf')
        for filename in ('../Outside.pdf', str(outside.resolve()),
                         '..\\Outside.pdf', 123):
            with self.subTest(filename=filename):
                record = self.manual_record(filename)
                row = dict(source_kind='manual', plant='plant2', record_id=record['id'],
                           certificate_status='generated_manual', certificate_no='0002',
                           pdf_filename='stale.pdf')
                before = self.snapshot()
                self.assertNotIn('pdf_filename', web_app.current_pdf_rows([row])[0])
                self.assertEqual(self.snapshot(), before)
        self.assertTrue(nested.is_file())

    def test_escaped_or_non_pdf_receipt_paths_are_not_exposed(self):
        for name in ('../Outside.pdf', '0001_Current_1.txt'):
            with self.subTest(name=name):
                path = self.pdf(name)
                receipt = self.receipt(path)
                # A receipt is still untrusted path data, even with a matching hash.
                if name.startswith('../'):
                    receipt['pdf_path'] = str(path.resolve())
                    save_json(self.root / 'Data' / 'Issuance History' / 'workbook-1.json', receipt)
                before = self.snapshot()
                rows = web_app.current_pdf_rows([self.workbook_row(pdf_filename='stale.pdf')])
                self.assertNotIn('pdf_filename', rows[0])
                self.assertEqual(self.snapshot(), before)

    def test_safe_nested_pdf_paths_remain_available(self):
        current=self.pdf('nested/0001_Current_1.pdf')
        self.receipt(current)
        before=self.snapshot()
        row=web_app.current_pdf_rows([self.workbook_row()])[0]
        self.assertEqual(row['pdf_relative_path'],'nested/0001_Current_1.pdf')
        self.assertEqual(self.snapshot(),before)

    def test_missing_unreadable_and_encrypted_pdfs_have_no_metadata(self):
        for kind in ('missing', 'unreadable', 'encrypted'):
            with self.subTest(kind=kind):
                filename = f'0001_{kind}_1.pdf'
                if kind == 'unreadable':
                    (self.folder / filename).write_bytes(b'not a readable PDF')
                elif kind == 'encrypted':
                    self.pdf(filename, encrypted=True)
                record = self.manual_record(filename, cod='0001')
                rows = [self.workbook_row(pdf_filename=filename),
                        dict(source_kind='manual', plant='plant2', record_id=record['id'],
                             certificate_status='completed', certificate_no='0001', pdf_filename=filename)]
                before = self.snapshot()
                self.assertTrue(all('pdf_filename' not in row for row in web_app.current_pdf_rows(rows)))
                self.assertEqual(self.snapshot(), before)

    def test_success_outputs_still_report_saved_filename_without_current_file(self):
        row = self.workbook_row(pdf_filename='0001_Actual #1%.pdf')
        self.assertEqual(web_app.certificate_outputs([row]), [], 'A missing PDF must not be exposed as an openable saved file.')
        self.assertNotIn('pdf_filename', web_app.current_pdf_rows([row])[0])

    def test_success_outputs_keep_audit_fallback_including_source_committed_receipts(self):
        for state in ('complete', 'source_committed'):
            with self.subTest(state=state):
                current = self.pdf('0001_Actual_1.pdf')
                self.receipt(current, state=state)
                current.unlink()
                outputs = web_app.certificate_outputs([self.workbook_row()])
                self.assertEqual(outputs, [], 'Audit history does not make a missing PDF appear saved.')


if __name__ == '__main__':
    unittest.main()
