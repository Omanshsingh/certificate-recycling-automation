"""Duplicate numbers remain repairable without allowing ambiguous issuance."""
from datetime import date
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from certificate_core import batch_content, read_batches
from client_intake import stage_new_client
from tests.test_certificate_workflow import make_workbook
from workbook_store import WorkbookSession
from workflow_support import Report, open_alert_workbook, open_excel_record, open_excel_workbook


class SerialNumberSafetyTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='serial-number-safety-')))
        self.path = self.root / 'Plant 1.xlsx'
        make_workbook(self.path, [{'sn': 'AbC', 'supplier': 'First company'},
                                  {'sn': 'abc', 'supplier': 'Second company'}])

    def test_manual_duplicates_are_readable_but_both_block_issuance_with_row_diagnostics(self):
        before = self.path.read_bytes()
        batches = read_batches(self.path)
        self.assertEqual([batch['supplier'] for batch in batches], ['First company', 'Second company'])
        for batch in batches:
            self.assertTrue(batch['duplicate_serial_number'])
            self.assertTrue(any('Serial Number' in issue and 'rows 2, 3' in issue
                                and 'Inward Register' in issue for issue in batch['blocking_issues']))
        self.assertEqual(self.path.read_bytes(), before)

    def test_intake_duplicate_explains_workbook_sheet_and_rows_without_saving(self):
        before = self.path.read_bytes()
        with WorkbookSession(self.path, self.root / 'backups') as session:
            with self.assertRaisesRegex(ValueError, 'Serial Number.*already used.*Plant 1.xlsx.*row 2.*row 3'):
                stage_new_client(session, {'sn': ' ABC ', 'supplier': 'New company', 'form6': 'F6-new'})
        self.assertEqual(self.path.read_bytes(), before)

    def test_repair_open_does_not_parse_or_claim_a_selected_record(self):
        before = self.path.read_bytes()
        with patch('workflow_support.os.startfile', create=True) as native, \
                patch('workflow_support.read_batches', side_effect=ValueError('Broken register')):
            opened = open_excel_record(self.path, 'Inward Register', 2, case_id='AbC',
                                       header_row=1, allow_duplicate_repair=True)
        self.assertEqual(opened, (self.path.resolve(), 'Inward Register', None))
        native.assert_called_once_with(str(self.path.resolve()))
        self.assertEqual(self.path.read_bytes(), before)

    def test_existing_manual_duplicates_allow_unrelated_chronological_intake_without_modifying_source(self):
        before = self.path.read_bytes()
        original = read_batches(self.path)
        with WorkbookSession(self.path, self.root / 'backups') as session:
            staged, added = stage_new_client(session, {'sn': 'NEW', 'supplier': 'New company',
                                                      'form6': 'F6-new', 'inward_date': '2026-01-01'})
        current = read_batches(BytesIO(staged))
        self.assertEqual(current[0]['case_id'], 'NEW')
        self.assertEqual(added['source_row'], current[0]['source_row'])
        existing = [batch for batch in current if batch['case_id'] != 'NEW']
        self.assertCountEqual([batch_content(batch) for batch in original],
                              [batch_content(batch) for batch in existing])
        self.assertTrue(all(batch['duplicate_serial_number'] for batch in existing))
        self.assertEqual(self.path.read_bytes(), before)

    def test_plain_open_rejects_missing_paths(self):
        with patch('workflow_support.os.startfile', create=True) as native:
            with self.assertRaises(FileNotFoundError):
                open_excel_workbook(self.root / 'Missing.xlsx')
        native.assert_not_called()

    def test_duplicate_alert_opens_workbook_for_repair_instead_of_rejecting_it(self):
        row = {'plant': 'plant1', 'sn': 'AbC', 'sheet': 'Inward Register',
               'file': str(self.path), 'system_alert': True}
        with patch('workflow_support.open_excel_workbook') as native, \
                patch('workflow_support.open_excel_record') as select:
            self.assertEqual(open_alert_workbook({'rows': [row]}, row), self.path)
        native.assert_called_once_with(self.path)
        select.assert_not_called()

    def test_read_only_report_is_honest_about_pending_status_update(self):
        batch = read_batches(self.path)[0]
        batch.update(_workbook_read_only=True, _missing_pdf=True)
        report = Report(self.root, date(2026, 10, 15), 'check')
        report.add(batch, 'ready', 'Create when ready.')
        row = report.rows[0]
        self.assertTrue(row['workbook_read_only'])
        self.assertIn('Excel status update is pending', row['detail'])
        self.assertNotIn('Excel was changed', row['detail'])
        self.assertTrue(any('Showing its saved contents' in warning for warning in row['warnings']))


if __name__ == '__main__':
    unittest.main()
