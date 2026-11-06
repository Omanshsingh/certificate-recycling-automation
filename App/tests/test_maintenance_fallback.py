"""Optional layout maintenance must not block a readable Excel register."""
from io import BytesIO
from pathlib import Path
import tempfile
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import patch

from openpyxl import load_workbook
from certificate_core import read_batches
from workbook_store import WorkbookSession, order_register_for_save
from tests.test_certificate_templates import write_register


class MaintenanceFallbackTests(TestCase):
    def setUp(self):
        self.root=Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.path=self.root/'plant1.xlsx'
        write_register(self.path)

    def test_protected_refresh_is_read_only_including_formatting_and_backups(self):
        book=load_workbook(self.path);book.active.protection.sheet=True
        book.save(self.path);book.close();before=self.path.read_bytes()
        with WorkbookSession(self.path,self.root/'backups') as session:
            session.ensure_alert_formatting('Inward Register')
            self.assertIn('protected',session.order_warning)
        self.assertEqual(self.path.read_bytes(),before)
        self.assertFalse((self.root/'backups').exists())

    def test_shared_status_merge_keeps_layout_and_facts_and_reports_why(self):
        book=load_workbook(self.path);book.active.merge_cells('M2:N2')
        book.save(self.path);book.close();before=read_batches(self.path)
        with WorkbookSession(self.path,self.root/'backups') as session:
            session.ensure_alert_formatting('Inward Register')
            self.assertIn('M2:N2',session.order_warning)
        book=load_workbook(self.path)
        self.assertIn('M2:N2',{str(area) for area in book.active.merged_cells.ranges})
        self.assertEqual(book.active['M1'].value,'Certificate Issued')
        book.close()
        self.assertEqual(read_batches(self.path),before)

    def test_failed_data_verification_is_never_treated_as_optional_layout_work(self):
        with patch('register_column_order.normalize_register_columns',side_effect=ValueError('Changed inward facts')):
            with self.assertRaisesRegex(ValueError,'Changed inward facts'):
                order_register_for_save(self.path.read_bytes(),SimpleNamespace())
