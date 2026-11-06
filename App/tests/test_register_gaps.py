"""Old cleared inward boxes are compacted without losing records or merge groups."""
from io import BytesIO
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

from certificate_core import read_batches, batch_content
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

    def test_blank_rows_within_a_populated_merged_block_are_kept(self):
        self.insert_gaps();book=load_workbook(self.path)
        book.active.merge_cells('A2:A4');book.save(self.path);book.close()
        before=self.path.read_bytes()
        self.assertEqual(compact_register_gaps(before),before)

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
