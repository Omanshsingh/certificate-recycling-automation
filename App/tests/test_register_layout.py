"""Original plant schemas, merged borders and hidden certificate settings."""
from copy import copy
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Font, PatternFill

from certificate_core import read_batches
from client_intake import stage_new_client
from workbook_store import WorkbookSession, add_tracking_column, stage_certificate_details


class RegisterLayoutTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))

    def source(self, plant):
        path = self.root / f'Plant {plant}.xlsx'
        w = Workbook(); s = w.active; s.title = 'Sheet1'
        headers = ['SN','Inward Date','Supplier Name','Location','Item','Quantity in Pcs']
        if plant == 2:
            headers += ['Per Pcs Weight in (KG)']
        headers += ['Weight (KG)','Form 6 No','Total Weight in KG','Comments']
        if plant == 1:
            headers += ['Reference No.']
        headers += ['COD No','COD Date']
        s.append(headers)
        for cell in s[1]:
            cell.font = Font(name='Calibri',size=11,bold=True)
            cell.fill = PatternFill('solid',fgColor='4BC8D6')
        for r in (2,3):
            for c in range(1,len(headers)+1):
                cell = s.cell(r,c)
                cell.font = Font(name='Calibri',size=11)
                cell.border = Border(left=Side(style='medium'),right=Side(style='medium'),bottom=Side(style='medium'))
        s['A2']='1';s['B2']='01.10.2026';s['C2']='Existing';s['D2']='Pune';s['E2']='Laptop';s['E3']='CPU';s['F2']=2;s['F3']=1
        weight = 7 if plant == 1 else 8
        s.cell(2,weight,10);s.cell(3,weight,6);s.cell(2,weight+1,'FORM-OLD')
        s.cell(2,weight+2,16)
        for col in [1,2,3,4,*range(weight+1,len(headers)+1)]:
            s.merge_cells(start_row=2,end_row=3,start_column=col,end_column=col)
        # Only the lower merge edge is saved. Loading synthesizes the anchor's
        # combined border and a new style ID, the old corruption trigger.
        s['A2'].border = Border(top=Side(style='medium'),left=Side(style='medium'),right=Side(style='medium'))
        s['A3'].border = Border(bottom=Side(style='medium'))
        s.row_dimensions[2].height=25;s.row_dimensions[3].height=25
        w.save(path);w.close()
        return path, headers

    def test_each_plant_adds_only_status_and_keeps_formulas_and_complete_borders(self):
        for plant in (1,2):
            with self.subTest(plant=plant):
                path, headers = self.source(plant)
                before = path.read_bytes()
                updated,batch=stage_new_client(SimpleNamespace(data=before),{
                    'supplier':'New client','location':'Mumbai','inward_date':'2026-10-01',
                    'form6':'FORM-NEW','reference':'REF-NEW','items':[
                        {'item':'Monitor','quantity':2,'per_piece_weight':3},
                        {'item':'Printer','quantity':1,'per_piece_weight':8}]})
                w=load_workbook(BytesIO(updated));s=w.active
                expected = headers + ['Certificate Issued']
                if plant == 1:
                    expected += ['Per Pcs Weight in (KG)']
                self.assertEqual([c.value for c in s[1]], expected)
                self.assertEqual(batch['total_weight'],14);self.assertEqual(batch['issued'],'Ready')
                self.assertEqual(batch['reference'],'REF-NEW')
                self.assertEqual(s['B4'].number_format,'dd.mm.yyyy')
                self.assertEqual(s['C4'].font.name,'Calibri');self.assertFalse(s['C4'].font.bold)
                self.assertIn('A4:A5',{str(a) for a in s.merged_cells.ranges})
                self.assertEqual(s['E5'].border.bottom.style,'medium')
                if plant==2:self.assertEqual(s['H4'].value,'=F4*G4')
                else:
                    self.assertEqual(s['G4'].value, '=F4*O4')
                    self.assertEqual(s['O4'].value, 3)
                w.close()
                with ZipFile(BytesIO(updated)) as z:
                    styles=ET.fromstring(z.read('xl/styles.xml')); ns={'x':styles.tag.split('}')[0][1:]}
                    count=len(styles.find('x:cellXfs',ns))
                    sheet=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
                    self.assertTrue(all(int(c.get('s','0'))<count for c in sheet.findall('.//x:c',ns)))
                self.assertEqual(path.read_bytes(),before)

    def test_certificate_corrections_round_trip_without_extra_visible_columns(self):
        path,headers=self.source(2)
        with WorkbookSession(path,self.root/'backups') as session:
            updated,batch=stage_certificate_details(session,{'case_id':'1','certificate_type':'siemens',
                'certificate_fields':{'model_serial':'SERIAL-01'},'reference':'REF-01','certificate_weight':20})
            session._persist(updated)
        self.assertEqual(batch['certificate_type'],'siemens')
        self.assertEqual(batch['reference'],'REF-01');self.assertEqual(batch['display_weight'],20)
        w=load_workbook(path);self.assertEqual([c.value for c in w.active[1]],headers+['Certificate Issued']);w.close()
        self.assertEqual(read_batches(path)[0]['certificate_fields'],{'model_serial':'SERIAL-01'})

    def test_tracking_setup_is_idempotent_and_reuses_the_trailing_empty_column(self):
        path,headers=self.source(1)
        w=load_workbook(path);w.active.cell(1,len(headers)+1).font=Font(name='Calibri');w.save(path);w.close()
        updated=add_tracking_column(path.read_bytes())
        self.assertEqual(add_tracking_column(updated),updated)
        w=load_workbook(BytesIO(updated));self.assertEqual(w.active.max_column,len(headers)+1);w.close()

    def test_real_quantity_labels_are_supported_without_extracting_arbitrary_text(self):
        from certificate_core import number
        from excel_values import quantity_number
        self.assertEqual(quantity_number('1 Lot',number),1)
        self.assertEqual(quantity_number('12 pcs',number),12)
        self.assertIsNone(quantity_number('2 broken entries',number))
        self.assertIsNone(quantity_number('1 kg',number))

    def test_empty_formatted_template_starts_at_row_two_without_old_blank_merges(self):
        path,headers=self.source(2)
        w=load_workbook(path);s=w.active
        for row in s.iter_rows(min_row=2):
            for cell in row:
                if cell.__class__.__name__ != 'MergedCell':cell.value=None
        w.save(path);w.close()
        updated,batch=stage_new_client(SimpleNamespace(data=path.read_bytes()),{
            'supplier':'First client','location':'Pune','inward_date':'2026-10-13','form6':'F01',
            'items':[{'item':'Laptop','quantity':2,'per_piece_weight':3},
                     {'item':'CPU','quantity':1,'per_piece_weight':8}]})
        self.assertEqual(batch['source_row'],2);self.assertEqual(batch['case_id'],'1')
        w=load_workbook(BytesIO(updated));s=w.active
        self.assertEqual(s.max_row,3);self.assertEqual(len(read_batches(BytesIO(updated))),1)
        self.assertEqual(s['J2'].value,'=SUM(H2:H3)');self.assertEqual(s['H2'].value,'=F2*G2')
        self.assertIn('A2:A3',{str(a) for a in s.merged_cells.ranges});w.close()

    def test_metadata_keeps_office_package_root_namespaces_and_unicode_values(self):
        from workbook_metadata import write_details,read_details,detail_key
        path,_=self.source(2)
        updated=write_details(path.read_bytes(),'Sheet1','1',{'reference':'REF & \u0915','certificate_fields':{'note':'a < b'}})
        self.assertEqual(read_details(BytesIO(updated))[detail_key('Sheet1','1')]['reference'],'REF & \u0915')
        with ZipFile(BytesIO(updated)) as z:
            for name,tag in [('[Content_Types].xml','Types'),('_rels/.rels','Relationships'),('docProps/custom.xml','Properties')]:
                payload=z.read(name)
                self.assertIn(('<'+tag+' xmlns=').encode(),payload)
                ET.fromstring(payload)

    def test_metadata_survives_excel_custom_property_limit_for_many_clients(self):
        from workbook_metadata import write_details,read_details,detail_key,PROPERTY_V2
        from io import BytesIO
        from zipfile import ZipFile
        path,_=self.source(2)
        updated=path.read_bytes()
        expected={}
        for index in range(1,11):
            details={'reference':f'REF-{index:02d}',
                     'certificate_fields':{'note':'client field ' + ('x'*500)}}
            expected[detail_key('Sheet1',str(index))]=details
            updated=write_details(updated,'Sheet1',str(index),details)
        self.assertEqual(read_details(BytesIO(updated)),expected)
        with ZipFile(BytesIO(updated)) as archive:
            props=ET.fromstring(archive.read('docProps/custom.xml'))
            chunks=[p for p in props if p.get('name','').startswith(PROPERTY_V2)]
            self.assertGreater(len(chunks),len(expected))
            self.assertTrue(all(len(next(iter(prop)).text or '')<=180 for prop in chunks))


if __name__=='__main__':
    unittest.main()
