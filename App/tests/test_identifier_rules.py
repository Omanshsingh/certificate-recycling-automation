"""Form 6 is global, SN is per plant, and neither is an overridable warning."""
from datetime import date
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from certificate_core import read_batches
from certificate_drafts import prepare_certificate, save_draft
from identifier_rules import conflicts, require_unique, workspace_batches
from run_certificates import run
from workflow_support import save_json
from tests.test_certificate_templates import write_register
from workspace_review import workspace_revision
import web_app


class IdentifierRuleTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root/f'Plant {i}.xlsx' for i in (1,2)]
        for i,path in enumerate(self.paths,1):
            write_register(path,case_id='1001',cod=f'{300+i:04}')
            self.change(i,form6=f'F6-{i}')
        save_json(self.root/'settings.local.json',{'workbook_paths':list(map(str,self.paths))})

    def change(self,plant,**values):
        path = self.paths[plant-1]
        book = load_workbook(path)
        columns = {'form6':9,'cod':11,'sn':1}
        for key,value in values.items():
            book.active.cell(2,columns[key]).value = value
        book.save(path);book.close()

    def test_same_sn_is_allowed_in_different_plants_and_own_record_is_excluded(self):
        batches = workspace_batches(self.root)
        require_unique(batches[0],'plant1','workbook','1001',batches)
        require_unique(batches[1],'plant2','workbook','1001',batches)
        with self.assertRaisesRegex(ValueError,'SN.*already used'):
            require_unique({'case_id':' 1001 '},'plant1','workbook',None,batches)
        with self.assertRaisesRegex(ValueError,'Form 6.*already used'):
            require_unique({'case_id':'NEW','form6':' f6-2 '},'plant1','workbook',None,batches)

    def test_excel_duplicates_in_both_plants_block_confirmed_generation_without_pdfs(self):
        self.change(2,form6=' f6-1 ')
        before = [path.read_bytes() for path in self.paths]
        report = {}
        result = run(self.paths,self.root,selected=('plant1','1001'),allow_warnings=True,
                     read_only_preflight=False,report_sink=report,progress=None)
        self.assertEqual(result[0],0)
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))
        row = next(row for row in report['rows'] if row.get('plant')=='plant1' and row.get('sn')=='1001')
        self.assertTrue(row['issuance_blocked'])
        self.assertIn('Form 6',row['detail'])
        self.assertIn('Plant 2',row['detail'])

    def test_manual_preview_and_save_reject_reused_manifest_but_keep_reference_manual(self):
        details = {'certificate_type':'standard','supplier':'New company','certificate_weight':'10',
                   'inward_date':'2026-10-01','form6':'F6-2','reference':'MANUAL / R-007'}
        for operation in (lambda: prepare_certificate(self.root,'plant1',details),
                          lambda: save_draft(self.root,'plant1',details)):
            with self.assertRaisesRegex(ValueError,'Form 6.*already used'):
                operation()
        details['form6']='F6-NEW'
        identifier = save_draft(self.root,'plant1',details)
        from certificate_drafts import read_register
        self.assertEqual(read_register(self.root)['records'][0]['details']['reference'],'MANUAL / R-007')
        batch,_,_,_ = prepare_certificate(self.root,'plant1',details,record_id=identifier)
        self.assertEqual(batch['reference'],'MANUAL / R-007')

    def test_supplemental_destruction_can_reference_its_original_inward(self):
        details = {'certificate_type':'destruction','case_id':'1001','form6':'F6-1',
                   'destruction_source':{'plant':'plant1','case_id':'1001'}}
        require_unique(details,'plant1','manual','extra-certificate',workspace_batches(self.root))

    def test_add_client_api_rejects_cross_plant_form6_without_saving(self):
        before = [path.read_bytes() for path in self.paths]
        body={'plant':'plant1','request_id':'a'*32,'expected_revision':workspace_revision(self.root),
              'details':{'sn':'NEW','supplier':'New Client','form6':'F6-2'}}
        with patch.object(web_app,'ROOT',self.root), patch.object(web_app,'BACKUP_DIR',self.root/'Backups'):
            with self.assertRaisesRegex(ValueError,'Form 6.*already used'):
                web_app.add_client(body)
        self.assertEqual(before,[path.read_bytes() for path in self.paths])

    def test_form6_duplicate_in_same_workbook_is_blocking(self):
        path=self.paths[0];book=load_workbook(path)
        second=[cell.value for cell in book.active[2]];second[0]='1002';second[8]=' f6-1 ';second[10]='0305'
        book.active.append(second);book.save(path);book.close()
        for batch in read_batches(path):
            self.assertTrue(any('Form 6' in issue for issue in batch['blocking_issues']))

    def test_duplicate_cod_on_unissued_inwards_is_not_an_overridable_warning(self):
        self.change(2,cod='0301')
        report = {}
        result = run(self.paths,self.root,selected=('plant1','1001'),allow_warnings=True,
                     report_sink=report,progress=None)
        self.assertEqual(result[0],0)
        row = next(row for row in report['rows'] if row.get('plant')=='plant1' and row.get('sn')=='1001')
        self.assertTrue(row['issuance_blocked'])
        self.assertIn('COD 0301',row['detail'])


if __name__=='__main__':
    unittest.main()
