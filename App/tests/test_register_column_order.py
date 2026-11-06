"""Column moves retain records, formulas, merges, formatting and package data."""
from copy import copy
from datetime import date
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from certificate_core import read_batches, batch_content
from register_column_order import normalize_register_columns, DETAILS
from tests.test_certificate_templates import write_register
from workbook_metadata import write_details
from client_intake import stage_new_client
from types import SimpleNamespace
from workbook_store import add_tracking_column


class ColumnOrderTests(unittest.TestCase):
    def setUp(self):
        self.root=Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.path=self.root/'Plant.xlsx'
        write_register(self.path,continuation=True,cod='0301',status='Generated')

    def legacy(self,details=True,table=False):
        book=load_workbook(self.path);sheet=book.active
        sheet.merge_cells('K2:K3');sheet.merge_cells('L2:L3');sheet.merge_cells('M2:M3')
        if details:
            for index,label in enumerate(DETAILS,14):
                sheet.cell(1,index).value=label;sheet.cell(2,index).value=f'Value {label}'
        sheet['K2'].number_format='0000'
        validation=DataValidation(type='list',formula1='"Ready,Generated,Completed"')
        sheet.add_data_validation(validation);validation.add('M2:M10000')
        sheet.conditional_formatting.add('M2:M10000',FormulaRule(formula=['$M2="Generated"'],fill=PatternFill('solid',fgColor='FFFF0000')))
        # Keep the formula in a different sheet so its reference follows the move.
        other=book.create_sheet('Summary');other['A1']="='Inward Register'!M2";other['B1']='=2+2'
        if table:
            last='Q' if details else 'M'
            sheet.add_table(Table(displayName='RegisterTable',ref=f'A1:{last}3'))
        output=BytesIO();book.save(output);book.close()
        return write_details(output.getvalue(),'Inward Register','CLIENT',{'certificate_type':'sita'})

    def test_existing_fields_tracking_and_values_move_without_changing_records_or_metadata(self):
        data=self.legacy();before=read_batches(BytesIO(data));updated=normalize_register_columns(data)
        after=read_batches(BytesIO(updated))
        self.assertEqual([batch_content(b) for b in before],[batch_content(b) for b in after])
        self.assertEqual([(b['cod_no'],b['cod_date'],b['issued']) for b in before],[(b['cod_no'],b['cod_date'],b['issued']) for b in after])
        book=load_workbook(BytesIO(updated));sheet=book.active
        self.assertEqual([c.value for c in sheet[1]][-7:],[*DETAILS,'COD No','COD Date','Certificate Issued'])
        self.assertEqual([sheet.cell(2,i).value for i in range(11,15)],[f'Value {label}' for label in DETAILS])
        self.assertIn('Q2:Q3',{str(a) for a in sheet.merged_cells.ranges})
        self.assertEqual(str(sheet.data_validations.dataValidation[0].sqref),'Q2:Q10000')
        self.assertEqual(book['Summary']['A1'].value,"='Inward Register'!Q2")
        self.assertEqual(book['Summary']['B1'].value,'=2+2')
        self.assertEqual(sheet['O2'].number_format,'0000')
        self.assertTrue(any(rule.formula==['$Q2="Generated"'] for key in sheet.conditional_formatting for rule in sheet.conditional_formatting[key]))
        book.close()
        with ZipFile(BytesIO(data)) as old, ZipFile(BytesIO(updated)) as new:
            for name in old.namelist():
                if name not in {'xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/workbook.xml'}:
                    self.assertEqual(old.read(name),new.read(name),name)
        self.assertEqual(normalize_register_columns(updated),updated)

    def test_missing_four_columns_are_added_before_tracking(self):
        data=self.path.read_bytes();updated=normalize_register_columns(data)
        book=load_workbook(BytesIO(updated));sheet=book.active
        self.assertEqual([c.value for c in sheet[1]][-7:],[*DETAILS,'COD No','COD Date','Certificate Issued'])
        self.assertEqual([sheet.cell(2,i).value for i in range(11,15)],[None]*4)
        self.assertEqual(sheet.cell(2,15).value,'0301')
        book.close()
        self.assertEqual(normalize_register_columns(updated),updated)

    def test_register_tables_keep_names_columns_and_range(self):
        data=self.legacy(table=True);updated=normalize_register_columns(data)
        book=load_workbook(BytesIO(updated));table=book.active.tables['RegisterTable']
        self.assertEqual(table.ref,'A1:Q3')
        self.assertEqual([column.name for column in table.tableColumns][-7:],[*DETAILS,'COD No','COD Date','Certificate Issued'])
        book.close()

    def test_app_intake_does_not_append_fields_after_status(self):
        data=add_tracking_column(self.path.read_bytes())
        updated,_=stage_new_client(SimpleNamespace(data=data),{'sn':'NEW','supplier':'New company','form6':'NEW-F6',
             'comments':'Call first','transporter_details':'Vehicle 1','docket_vehicle_no':'D1','weight_slip_scale':'W1'})
        book=load_workbook(BytesIO(updated));headings=[c.value for c in book.active[1]]
        self.assertEqual(headings[-7:],[*DETAILS,'COD No','COD Date','Certificate Issued'])
        book.close()

    def test_workbook_issuance_keeps_tracking_at_end_and_commits_history(self):
        import shutil
        from run_certificates import run
        from workflow_support import save_json
        from issuance_safety import audit_records, recover_issuances
        shutil.copytree(Path(__file__).resolve().parents[1]/'templates',self.root/'App'/'templates')
        other=self.root/'Other.xlsx'
        write_register(other,case_id='OTHER',cod='0302')
        book=load_workbook(other);book.active['I2']='FORM-OTHER';book.save(other);book.close()
        save_json(self.root/'settings.local.json',{'workbook_paths':[str(self.path),str(other)]})
        result=run([self.path,other],self.root,selected=('plant1','CLIENT'),allow_warnings=True,progress=None)
        self.assertEqual(result[0],1)
        book=load_workbook(self.path)
        self.assertEqual([c.value for c in book.active[1]][-7:],[*DETAILS,'COD No','COD Date','Certificate Issued'])
        book.close()
        self.assertEqual(read_batches(self.path)[0]['issued'],'Generated')
        recover_issuances(self.root)
        self.assertEqual(audit_records(self.root)[-1]['state'],'complete')


if __name__=='__main__':unittest.main()
