"""Removal guards and package preservation, using disposable workbooks only."""
from io import BytesIO
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook

from certificate_core import read_batches
from inward_removal import remove_inward, stage_removal
from tests.test_certificate_workflow import make_workbook
from workbook_metadata import read_details, write_details
from workbook_store import exclusive_file
from workflow_support import run_lock, save_json
from workspace_review import workspace_revision


class InwardRemovalTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root/'Plant 1.xlsx', self.root/'Plant 2.xlsx']
        make_workbook(self.paths[0], [{'sn':'1'}, {'sn':'2'}, {'sn':'3','status':'N/A'}])
        make_workbook(self.paths[1], [{'sn':'2'}])  # Same SN in the other plant is valid.
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.insert_rows(4)
        sheet['E4']='Printer';sheet['F4']=2;sheet['G4']=3;sheet['H4']='=F4*G4'
        for column in ('A','B','C','D','I','J','K','L','M'):
            sheet.merge_cells(f'{column}3:{column}4')
        book.save(self.paths[0]);book.close()
        self.paths[0].write_bytes(write_details(self.paths[0].read_bytes(),'Inward Register','2',{'reference':'PRIVATE-REMOVED'}))
        save_json(self.root/'settings.local.json', {'workbook_paths':list(map(str,self.paths)), 'reminders_enabled':False})
        with run_lock(self.root):
            pass

    def body(self):
        return dict(plant='plant1',sn='2',sheet='Inward Register',row=3,request_id='a'*32,
                    expected_revision=workspace_revision(self.root))

    def reviewed(self):
        body=self.body();preview=remove_inward(self.root,body)
        return dict(body,review_token=preview['review_token'],checked=True)

    def mutate(self, action, path=None):
        path=path or self.paths[0];book=load_workbook(path)
        action(book.active);book.save(path);book.close()

    def test_preview_and_cancel_change_no_files(self):
        before={p:p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        result=remove_inward(self.root,self.body())
        self.assertEqual(result['entry']['materials'],2)
        self.assertEqual(result['entry']['sn'],'2')
        self.assertEqual(before,{p:p.read_bytes() for p in self.root.rglob('*') if p.is_file()})

    def test_removal_keeps_neighbouring_entries_other_plant_and_original_backup(self):
        before=self.paths[0].read_bytes();other=self.paths[1].read_bytes()
        rows=read_batches(self.paths[0]);result=remove_inward(self.root,self.reviewed(),commit=True)
        self.assertTrue(result['removed'])
        self.assertEqual(read_batches(self.paths[0]),[rows[0],dict(rows[2],case_id='2',source_row=3)])
        self.assertEqual(self.paths[1].read_bytes(),other)
        backups=list((self.root/'Data/Workbook Backups').glob('*.xlsx'))
        self.assertEqual([p.read_bytes() for p in backups],[before])
        self.assertEqual(read_details(self.paths[0]),{})
        book=load_workbook(self.paths[0]);sheet=book.active
        self.assertEqual(sheet['A3'].value,'2');self.assertIsNone(sheet['E4'].value)
        self.assertIsNone(sheet['A5'].value)
        self.assertFalse(any(a.min_row==3 for a in sheet.merged_cells.ranges));book.close()

    def test_package_only_changes_selected_sheet_and_its_private_metadata(self):
        before=self.paths[0].read_bytes();body=self.reviewed()
        remove_inward(self.root,body,commit=True)
        with ZipFile(BytesIO(before)) as original,ZipFile(self.paths[0]) as current:
            for name in original.namelist():
                if name not in {'xl/worksheets/sheet1.xml','docProps/custom.xml','[Content_Types].xml','_rels/.rels'}:
                    self.assertEqual(original.read(name),current.read(name),name)

    def test_bare_or_unchecked_confirmation_is_rejected_without_writes(self):
        body=self.reviewed();before=self.paths[0].read_bytes()
        for changes in ({'checked':False},{'checked':'true'},{'review_token':''},{'review_token':'forged'}):
            with self.subTest(changes=changes),self.assertRaises(ValueError):
                remove_inward(self.root,dict(body,**changes),commit=True)
            self.assertEqual(self.paths[0].read_bytes(),before)
        self.assertFalse((self.root/'Data/Workbook Backups').exists())

    def test_changed_or_moved_entry_invalidates_confirmation(self):
        body=self.reviewed();self.mutate(lambda s:setattr(s['C3'],'value','Changed company'))
        before=self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError,'files changed'):
            remove_inward(self.root,body,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)
        body['expected_revision']=workspace_revision(self.root)
        with self.assertRaisesRegex(ValueError,'Check the entry'):
            remove_inward(self.root,body,commit=True)
        with self.assertRaisesRegex(ValueError,'moved'):
            remove_inward(self.root,dict(self.body(),row=5))

    def test_other_plant_changes_also_invalidate_confirmation(self):
        body=self.reviewed();before=self.paths[0].read_bytes()
        self.mutate(lambda s:setattr(s['C2'],'value','Changed other plant'),self.paths[1])
        with self.assertRaisesRegex(ValueError,'files changed'):
            remove_inward(self.root,body,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_generated_and_completed_certificates_block_removal(self):
        for column,value in [('M3','Generated'),('M3','Completed')]:
            with self.subTest(column=column,value=value):
                self.mutate(lambda s:setattr(s[column],'value',value))
                before=self.paths[0].read_bytes()
                with self.assertRaisesRegex(ValueError,'certificate'):
                    remove_inward(self.root,self.body())
                self.assertEqual(self.paths[0].read_bytes(),before)
                self.mutate(lambda s:setattr(s[column],'value',None))

    def test_cleared_excel_tracking_does_not_hide_created_certificate_history(self):
        save_json(self.root/'Data/Issuance History'/('b'*32+'.json'),
                  dict(version=1,source_kind='workbook',plant='plant1',key='2',state='complete'))
        with self.assertRaisesRegex(ValueError,'already created'):
            remove_inward(self.root,self.body())

    def test_linked_destruction_request_protects_payment_followup(self):
        save_json(self.root/'Data/Custom Certificates.json',dict(version=1,records=[dict(id='request',plant='plant1',status='draft',
            details=dict(certificate_type='destruction',destruction_source=dict(plant='plant1',case_id='2')))]))
        with self.assertRaisesRegex(ValueError,'data destruction request'):
            remove_inward(self.root,self.body())

    def test_duplicate_sn_is_never_guessed(self):
        self.mutate(lambda s:setattr(s['A5'],'value','2'))
        with self.assertRaisesRegex(ValueError,'more than once'):
            remove_inward(self.root,self.body())

    def test_shared_merges_and_dependent_formulas_block_removal(self):
        original=self.paths[0].read_bytes()
        for change in [lambda s:s.merge_cells('N3:N5'),lambda s:setattr(s['H5'],'value','=SUM(H3:H4)'),
                       lambda s:setattr(s['H5'],'value','=INDIRECT("H3")')]:
            self.paths[0].write_bytes(original);self.mutate(change)
            before=self.paths[0].read_bytes()
            with self.assertRaises(ValueError):
                remove_inward(self.root,self.body())
            self.assertEqual(self.paths[0].read_bytes(),before)

    def test_protection_lock_and_backup_failure_preserve_original(self):
        body=self.reviewed();before=self.paths[0].read_bytes()
        with patch('workbook_store.Path.mkdir',side_effect=OSError('Backup unavailable')):
            with self.assertRaises(OSError):
                remove_inward(self.root,body,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)
        self.mutate(lambda s:setattr(s.protection,'sheet',True))
        with self.assertRaisesRegex(ValueError,'protected'):
            remove_inward(self.root,self.body())
        with exclusive_file(self.paths[0]):
            with self.assertRaises(OSError):
                remove_inward(self.root,dict(body,expected_revision='old'))

    def test_lost_response_retry_does_not_remove_a_replacement_entry(self):
        body=self.reviewed();remove_inward(self.root,body,commit=True)
        self.mutate(lambda s:(setattr(s['A3'],'value','2'),setattr(s['C3'],'value','New replacement')))
        before=self.paths[0].read_bytes()
        self.assertTrue(remove_inward(self.root,body,commit=True)['removed'])
        self.assertEqual(self.paths[0].read_bytes(),before)
        self.assertEqual(len(list((self.root/'Data/Workbook Backups').glob('*.xlsx'))),1)

    def test_prepared_journal_proves_success_when_final_journal_write_fails(self):
        from inward_removal import save_json as save
        def fail_final(path,value):
            if value.get('saved'):
                raise OSError('Journal write failed after Excel was saved')
            save(path,value)
        body=self.reviewed()
        with patch('inward_removal.save_json',side_effect=fail_final):
            self.assertTrue(remove_inward(self.root,body,commit=True)['removed'])
        before=self.paths[0].read_bytes()
        self.assertTrue(remove_inward(self.root,body,commit=True)['removed'])
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_unfinished_request_never_replays_against_an_unverified_file(self):
        body=self.reviewed();before=self.paths[0].read_bytes()
        with patch('workbook_store.WorkbookSession._persist',side_effect=OSError('Disk unavailable')):
            with self.assertRaises(OSError):
                remove_inward(self.root,body,commit=True)
        with self.assertRaisesRegex(ValueError,'reopen Remove inward'):
            remove_inward(self.root,body,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)
        fresh=dict(self.body(),request_id='c'*32)
        preview=remove_inward(self.root,fresh)
        self.assertTrue(remove_inward(self.root,dict(fresh,checked=True,review_token=preview['review_token']),commit=True)['removed'])

    def test_external_change_at_commit_boundary_stops_before_backup_or_write(self):
        from inward_removal import save_json as save
        body=self.reviewed();before=self.paths[0].read_bytes()
        def change_other_plant(path,value):
            save(path,value)
            if value.get('saved') is False:
                self.mutate(lambda s:setattr(s['C2'],'value','External change'),self.paths[1])
        with patch('inward_removal.save_json',side_effect=change_other_plant):
            with self.assertRaisesRegex(ValueError,'files changed'):
                remove_inward(self.root,body,commit=True)
        self.assertEqual(self.paths[0].read_bytes(),before)
        self.assertFalse((self.root/'Data/Workbook Backups').exists())

    def test_one_file_cannot_be_removed_as_two_different_plants(self):
        save_json(self.root/'settings.local.json',{'workbook_paths':[str(self.paths[0])]*2})
        with self.assertRaisesRegex(ValueError,'different Excel files'):
            remove_inward(self.root,self.body())

    def test_shared_formula_anchor_is_never_removed_under_other_entries(self):
        self.mutate(lambda s:setattr(s['H3'],'value','=F3*G3'))
        source=self.paths[0].read_bytes();output=BytesIO()
        with ZipFile(BytesIO(source)) as original,ZipFile(output,'w') as target:
            for info in original.infolist():
                data=original.read(info.filename)
                if info.filename=='xl/worksheets/sheet1.xml':
                    data=data.replace(b'<f>F3*G3</f>',b'<f t="shared" si="0" ref="H3:H5">F3*G3</f>')
                target.writestr(info,data)
        self.paths[0].write_bytes(output.getvalue());before=self.paths[0].read_bytes()
        with self.assertRaisesRegex(ValueError,'shared formula group is incomplete'):
            remove_inward(self.root,self.body())
        self.assertEqual(self.paths[0].read_bytes(),before)

    def test_complete_shared_formula_group_survives_removing_an_inward(self):
        book=load_workbook(self.paths[0]);sheet=book.active
        for row in range(2,6): sheet[f'H{row}']=f'=F{row}*G{row}'
        book.save(self.paths[0]);book.close()
        output=BytesIO()
        with ZipFile(self.paths[0]) as source,ZipFile(output,'w') as target:
            for member in source.infolist():
                payload=source.read(member.filename)
                if member.filename=='xl/worksheets/sheet1.xml':
                    xml=ET.fromstring(payload)
                    for row in range(2,6):
                        formula=xml.find(f'.//{{*}}c[@r="H{row}"]/{{*}}f')
                        formula.attrib.update(t='shared',si='0')
                        if row==2: formula.attrib['ref']='H2:H5'
                        else: formula.text=None
                    payload=ET.tostring(xml)
                target.writestr(member,payload)
        self.paths[0].write_bytes(output.getvalue())

        result=remove_inward(self.root,self.reviewed(),commit=True)

        self.assertTrue(result['removed'])
        self.assertEqual([row['case_id'] for row in read_batches(self.paths[0])],['1','2'])
        book=load_workbook(self.paths[0],data_only=False)
        self.assertEqual(book.active['H2'].value,'=F2*G2')
        self.assertEqual(book.active['H3'].value,'=F3*G3')
        book.close()


if __name__=='__main__':
    unittest.main()
