"""Native Excel launch guards; workbooks are disposable and COM is mocked here."""
import base64
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from tests.test_certificate_workflow import make_workbook
from workflow_support import open_excel_record
import web_app


class ExcelSelectionTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix='excel-selection-unit-')))
        self.path = self.root/'Plant workbook.xlsx'
        make_workbook(self.path,[{'sn':'1001'},{'sn':'1002'}])

    def test_selection_uses_live_identity_and_returns_the_moved_row_without_saving(self):
        before = self.path.read_bytes()
        complete = subprocess.CompletedProcess([],0,'{"selected":true,"row":9,"last_row":11}','')
        with patch('workflow_support.shutil.which',return_value='powershell.exe'), patch('workflow_support.subprocess.run',return_value=complete) as native:
            result = open_excel_record(self.path,'Inward Register',2,case_id='1001',header_row=1)
        script = base64.b64decode(native.call_args.args[0][-1]).decode('utf-16le')
        self.assertEqual(result,(self.path.resolve(),'Inward Register',9))
        self.assertIn('StringComparison]::Ordinal',script)
        self.assertIn('$matches.Count -ne 1',script)
        self.assertIn('$target=$sheet.Range',script)
        self.assertIn('$endRow',script)
        self.assertNotIn('.Save(',script); self.assertNotIn('.Interior.',script)
        self.assertEqual(self.path.read_bytes(),before)

    def test_caller_without_metadata_resolves_the_client_before_launch(self):
        complete = subprocess.CompletedProcess([],0,'{"selected":true,"row":3}','')
        with patch('workflow_support.shutil.which',return_value='powershell.exe'), patch('workflow_support.subprocess.run',return_value=complete) as native:
            self.assertEqual(open_excel_record(self.path,'Inward Register',3)[2],3)
        script = base64.b64decode(native.call_args.args[0][-1]).decode('utf-16le')
        self.assertIn(base64.b64encode(b'1002').decode('ascii'),script)

    def test_unavailable_or_ambiguous_live_selection_is_an_error(self):
        before = self.path.read_bytes()
        with patch('workflow_support.shutil.which',return_value='powershell.exe'), patch('workflow_support.subprocess.run',return_value=subprocess.CompletedProcess([],1,'','Client ID is repeated')):
            with self.assertRaisesRegex(OSError,'repeated'):
                open_excel_record(self.path,'Inward Register',2,case_id='1001',header_row=1)
        self.assertEqual(self.path.read_bytes(),before)

    def test_unconfirmed_selection_is_not_reported_as_success(self):
        with patch('workflow_support.shutil.which',return_value='powershell.exe'), patch('workflow_support.subprocess.run',return_value=subprocess.CompletedProcess([],0,'{"selected":false,"row":2}','')):
            with self.assertRaisesRegex(OSError,'confirm'):
                open_excel_record(self.path,'Inward Register',2,case_id='1001',header_row=1)

    def test_unsafe_or_missing_client_locations_do_not_launch_excel(self):
        with patch('workflow_support.subprocess.run') as native:
            for case_id,header,row in (('',1,2),('1001',0,2),('1001',1,1)):
                with self.subTest(case_id=case_id,header=header,row=row), self.assertRaises(ValueError):
                    open_excel_record(self.path,'Inward Register',row,case_id=case_id,header_row=header)
            with self.assertRaises(ValueError):
                open_excel_record(self.path,'Wrong worksheet',2)
        native.assert_not_called()

    def test_excel_dialog_timeout_has_a_short_actionable_message(self):
        with patch('workflow_support.shutil.which',return_value='powershell.exe'), patch('workflow_support.subprocess.run',side_effect=subprocess.TimeoutExpired('encoded command',30)):
            with self.assertRaisesRegex(OSError,'Finish any open message in Excel'):
                open_excel_record(self.path,'Inward Register',2,case_id='1001',header_row=1)

    def test_web_endpoint_rejects_duplicate_ids_before_launching_excel(self):
        make_workbook(self.path,[{'sn':'1001'},{'sn':'1001'}])
        row = {'plant':'plant1','sn':'1001','file':str(self.path),'sheet':'Inward Register'}
        with patch.object(web_app,'latest_result',return_value={'rows':[row]}), patch.object(web_app,'configured_paths',return_value=[self.path,self.root/'Other.xlsx']), patch.object(web_app,'open_excel_record') as native:
            with self.assertRaisesRegex(ValueError,'changed'):
                web_app.open_selected_workbook_record('plant1','1001')
        native.assert_not_called()


if __name__ == '__main__':
    unittest.main()
