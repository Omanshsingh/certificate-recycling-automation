"""Old cleared inward boxes are compacted without losing records or merge groups."""
from io import BytesIO
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

from certificate_core import read_batches, batch_content, select_register
from register_rows import compact_register_gaps
from tests.test_certificate_workflow import make_workbook
from workbook_store import WorkbookSession, UnsupportedReorderError


class RegisterGapTests(unittest.TestCase):
    def setUp(self):
        self.root=Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.path=self.root/'plant.xlsx'
        make_workbook(self.path,[{'sn':'1'},{'sn':'2'},{'sn':'3','status':'Completed','cod':'0042'}])

    def insert_gaps(self):
        book=load_workbook(self.path);book.active.insert_rows(3,2);book.save(self.path);book.close()

    def test_cleanup_removes_only_empty_boxes_and_keeps_primary_facts(self):
        self.insert_gaps();before=read_batches(self.path);mapping={}
        updated=compact_register_gaps(self.path.read_bytes(),source_rows=mapping)
        after=read_batches(BytesIO(updated))
        self.assertEqual([batch_content(b) for b in after],[batch_content(b) for b in before])
        self.assertEqual([b['source_row'] for b in after],[2,3,4])
        self.assertEqual((after[-1]['cod_no'],after[-1]['issued']),('0042','Completed'))
        self.assertEqual(mapping[('Inward Register',5)],3)
        self.assertEqual(compact_register_gaps(updated),updated)

    def test_blank_rows_under_a_client_field_merge_are_removed_without_a_total(self):
        self.insert_gaps();book=load_workbook(self.path)
        book.active.merge_cells('A2:A4');book.save(self.path);book.close()
        before=self.path.read_bytes()
        updated = compact_register_gaps(before)
        self.assertEqual([b['source_row'] for b in read_batches(BytesIO(updated))], [2, 3, 4])
        self.assertEqual(compact_register_gaps(updated), updated)

    def test_cleared_material_tail_shrinks_its_total_and_merges(self):
        self.insert_gaps()
        book = load_workbook(self.path)
        sheet = book.active
        sheet['N1'] = 'Total Weight in KG'
        sheet['N2'] = '=SUM($H$2:$H$4)'
        sheet.merge_cells('A2:A4')
        sheet.merge_cells('M2:M4')
        book.save(self.path)
        book.close()
        before = self.path.read_bytes()
        mapping = {}
        updated = compact_register_gaps(before, source_rows=mapping)
        after = read_batches(BytesIO(updated))
        self.assertEqual([batch_content(b) for b in after],
                         [batch_content(b) for b in read_batches(BytesIO(before))])
        self.assertEqual([batch['source_row'] for batch in after], [2, 3, 4])
        self.assertEqual(mapping[('Inward Register', 5)], 3)
        book = load_workbook(BytesIO(updated))
        self.assertEqual(book.active['N2'].value, '=SUM($H$2:$H$2)')
        self.assertFalse(book.active.merged_cells.ranges)
        book.close()
        self.assertEqual(compact_register_gaps(updated), updated)

    def test_refresh_drops_a_cleared_material_row_but_keeps_a_cod_only_row(self):
        self.insert_gaps()
        book = load_workbook(self.path)
        sheet = book.active
        sheet['K4'] = 'DD-0081'
        sheet['L4'] = sheet['B2'].value
        sheet.merge_cells('I2:I4')
        book.save(self.path)
        book.close()
        before = read_batches(self.path)
        with WorkbookSession(self.path, self.root/'backups') as session:
            session.ensure_alert_formatting('Inward Register')
        self.assertEqual([b['case_id'] for b in read_batches(self.path)], [b['case_id'] for b in before])
        book = load_workbook(self.path)
        sheet, _, cols = select_register(book)
        cod = sheet.cell(3, cols['cod_no'] + 1)
        cod_date = sheet.cell(3, cols['cod_date'] + 1)
        self.assertEqual((cod.value, cod_date.value), ('DD-0081', sheet['B2'].value))
        self.assertEqual(sheet['A4'].value, '2')
        self.assertIn('I2:I3', {str(area) for area in sheet.merged_cells.ranges})
        self.assertTrue(all(cod.coordinate not in area and cod_date.coordinate not in area
                            for area in sheet.merged_cells.ranges))
        book.close()

    def test_refresh_remaps_its_in_memory_rows_and_backs_up_original_bytes(self):
        self.insert_gaps();before=self.path.read_bytes();batches=read_batches(self.path)
        with WorkbookSession(self.path,self.root/'backups') as session:
            session.ensure_alert_formatting('Inward Register',batches)
        self.assertEqual([b['source_row'] for b in batches],[2,3,4])
        self.assertEqual([b['case_id'] for b in batches],['1','2','3'])
        self.assertTrue(any(p.read_bytes()==before for p in (self.root/'backups').glob('*.xlsx')))

    def test_a_formula_using_a_gap_is_not_silently_changed(self):
        self.insert_gaps();book=load_workbook(self.path)
        book.active['N2']='=H3';book.save(self.path);book.close()
        before=self.path.read_bytes()
        with self.assertRaisesRegex(UnsupportedReorderError,'cleanup needs review'):
            compact_register_gaps(before)
        self.assertEqual(self.path.read_bytes(),before)


if __name__=='__main__':unittest.main()
