from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles import Protection
from openpyxl.utils.cell import range_boundaries

from certificate_core import read_batches
from client_intake import stage_new_client
from register_protection import protect_issued_inwards
from tests.test_certificate_workflow import make_workbook
from workbook_store import WorkbookSession


class RegisterProtectionTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(self.enterContext(TemporaryDirectory()))
        self.path = self.folder / 'Plant 1.xlsx'
        make_workbook(self.path, [{'sn':'1', 'status':'Ready'}, {'sn':'2', 'status':'Ready'}])

    def editable_rows(self, data):
        with ZipFile(BytesIO(data)) as archive:
            root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        self.assertIsNotNone(root.find('{*}sheetProtection'))
        return [(top, bottom) for _, top, _, bottom in
                (range_boundaries(item.get('sqref')) for item in root.findall('.//{*}protectedRange')
                 if item.get('name').startswith('EcorecoEditable'))]

    def test_issuance_locks_the_inward_and_its_materials_and_app_can_keep_writing(self):
        book = load_workbook(self.path)
        sheet = book.active
        sheet.insert_rows(3)
        sheet['E3'] = 'More equipment'
        sheet['H3'] = 2
        sheet['C2'].protection = Protection(locked=False)
        book.save(self.path)
        book.close()
        issued_on = date(2026, 11, 6)
        with WorkbookSession(self.path, self.folder / 'Backups') as session:
            batch = read_batches(BytesIO(session.data))[0]
            session.commit(batch, '0001', issued_on)
            self.assertEqual(self.editable_rows(session.data), [(4, 1048576)])
            book = load_workbook(BytesIO(session.data))
            self.assertTrue(book.active['C2'].protection.locked)
            self.assertFalse(book.active.protection.formatCells)
            self.assertFalse(book.active.protection.formatRows)
            book.close()
            batch = read_batches(BytesIO(session.data))[0]
            with ZipFile(BytesIO(session.data)) as archive:
                root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            editable_details = {item.get('sqref') for item in root.findall('.//{*}protectedRange')
                                if item.get('name').startswith('EcorecoDetails')}
            for key in ('comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale'):
                column = batch['column_mapping'][key]['column']
                self.assertIn(f'{column}2:{column}3', editable_details)
            self.assertEqual(protect_issued_inwards(session.data), session.data)
            session.set_issued_status('1', batch['source_sheet'], batch['source_row'], '0001', 'Generated', 'Completed')
            updated, _ = stage_new_client(session, {'supplier':'New intake'})
            session._persist(updated)
            self.assertEqual(self.editable_rows(session.data), [(4, 1048576)])
            self.assertEqual([row['issued'] for row in read_batches(BytesIO(session.data))][:2], ['Completed','Ready'])

    def test_existing_managed_register_allows_cell_formatting_after_refresh(self):
        make_workbook(self.path, [{'sn':'1', 'status':'Generated', 'cod':'0001'},
                                  {'sn':'2', 'status':'Ready'}])
        managed = protect_issued_inwards(self.path.read_bytes())
        older = BytesIO()
        with ZipFile(BytesIO(managed)) as source, ZipFile(older, 'w') as target:
            for entry in source.infolist():
                content = source.read(entry)
                if entry.filename == 'xl/worksheets/sheet1.xml':
                    content = content.replace(b' formatCells="0"', b'')
                target.writestr(entry, content)
        self.path.write_bytes(older.getvalue())
        with WorkbookSession(self.path, self.folder / 'Backups') as session:
            session.ensure_alert_formatting('Inward Register')
            self.assertEqual(self.editable_rows(session.data), [(3, 1048576)])
            book = load_workbook(BytesIO(session.data))
            try:
                self.assertFalse(book.active.protection.formatCells)
                self.assertTrue(book.active['C2'].protection.locked)
            finally:
                book.close()

    def test_refresh_locks_external_issuance_without_replacing_owner_protection(self):
        make_workbook(self.path, [{'sn':'1', 'status':'Generated Externally', 'cod':'0001'}])
        with WorkbookSession(self.path, self.folder / 'Backups') as session:
            session.ensure_alert_formatting('Inward Register')
            self.assertEqual(self.editable_rows(session.data), [(3, 1048576)])
        book = load_workbook(self.path)
        book.active.protection.set_password('workbook-owner')
        book.active.protection.hashValue = None
        book.save(self.path)
        book.close()
        before = self.path.read_bytes()
        self.assertEqual(protect_issued_inwards(before), before)
