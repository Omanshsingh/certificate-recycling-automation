"""A batch must issue exactly the reviewed plant/SN pairs, never the whole plant."""
from pathlib import Path
import tempfile
import shutil
import unittest
from unittest.mock import patch
from datetime import date

from openpyxl import load_workbook
from certificate_core import read_batches
from tests.test_certificate_templates import write_register
from workflow_support import Report, save_json
from workspace_review import workspace_revision
import web_app


class ChosenCertificatesTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root / f'plant{i}.xlsx' for i in (1,2)]
        for index,path in enumerate(self.paths,1):
            write_register(path, case_id='1', cod=f'{300+index:04}', status='Completed')
            book=load_workbook(path);sheet=book.active
            sheet.append(['2',date(2026,9,1),f'Choose {index}','Pune','Laptop',1,10,10,f'CHOOSE-F6-{index}',f'CHOOSE-REF-{index}',None,None,'Ready'])
            sheet.append(['3',date(2026,9,1),f'Leave {index}','Pune','Laptop',1,10,10,f'LEAVE-F6-{index}',f'LEAVE-REF-{index}',None,None,'Ready'])
            book.save(path);book.close()
        save_json(self.root/'settings.local.json',{'workbook_paths':list(map(str,self.paths))})
        shutil.copytree(Path(web_app.__file__).parent/'templates',self.root/'App'/'templates')
        self.enterContext(patch.object(web_app,'ROOT',self.root))
        self.enterContext(patch.object(web_app,'RESULTS_PATH',self.root/'Data'/'Reports'/'Latest Results.json'))
        self.enterContext(patch.object(web_app,'BACKUP_DIR',self.root/'Backups'))

    def test_only_checked_clients_across_both_plants_receive_pdfs(self):
        result=web_app.generate(selections=[{'plant':'plant1','sn':'2'},{'plant':'plant2','sn':'2'}],
                                confirmed=True,expected_revision=workspace_revision(self.root))
        self.assertEqual(result['created'],2)
        self.assertEqual(len(list((self.root/'Certificates').rglob('*.pdf'))),2)
        for path in self.paths:
            rows={batch['case_id']:batch for batch in read_batches(path)}
            self.assertEqual(rows['2']['issued'],'Generated')
            self.assertEqual(rows['3']['issued'],'Ready')
            self.assertEqual(rows['3']['cod_no'],'')

    def test_empty_duplicate_unknown_and_changed_selections_create_nothing(self):
        before=[path.read_bytes() for path in self.paths]
        selections=[[],[{'plant':'plant1','sn':'2'}]*2,[{'plant':'plant1','sn':'2'},{'plant':'plant2','sn':'missing'}]]
        for chosen in selections:
            with self.subTest(chosen=chosen),self.assertRaises(ValueError):
                web_app.generate(selections=chosen,confirmed=True,expected_revision=workspace_revision(self.root))
        with self.assertRaises(ValueError):
            web_app.generate(selections=[{'plant':'plant1','sn':'2'}],confirmed=True,expected_revision='outdated')
        self.assertEqual(before,[path.read_bytes() for path in self.paths])
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_selected_creation_keeps_other_clients_in_cached_and_saved_company_list(self):
        report = Report(self.root, date.today(), 'Check', publish_latest=False)
        for index,path in enumerate(self.paths,1):
            for batch in read_batches(path):
                report.add(dict(batch, source_id=f'plant{index}', source_path=str(path)), 'ready', 'Checked')
        state = {'rows':report.rows, 'scan_complete':True}
        revision = workspace_revision(self.root)
        with patch.object(web_app,'WORKSPACE_STATE_CACHE',(revision,state)):
            result = web_app.generate(selections=[{'plant':'plant1','sn':'2'}],confirmed=True,expected_revision=revision)
        self.assertEqual(result['created'],1)
        self.assertEqual(len(result['state']['rows']),6)
        rows={(row['plant'],row['sn']):row for row in result['state']['rows']}
        self.assertEqual(rows['plant1','2']['certificate_status'],'generated')
        self.assertEqual(rows['plant1','3']['certificate_status'],'ready')
        self.assertEqual(rows['plant2','2']['certificate_status'],'ready')
        self.assertEqual(len(web_app.latest_result()['rows']),6)
        self.assertEqual(len(state['rows']),6)


if __name__=='__main__':unittest.main()
