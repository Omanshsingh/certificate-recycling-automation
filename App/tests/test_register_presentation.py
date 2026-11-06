"""Display-only Excel date and alignment changes preserve underlying records."""
from copy import copy
from datetime import datetime
from io import BytesIO
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from register_presentation import normalize_register_presentation


class RegisterPresentationTests(unittest.TestCase):
    def source(self):
        book=Workbook();sheet=book.active
        sheet.append(['SN','Inward Date','Supplier Name','Location','Item','Weight (KG)','Form 6 No','COD No','COD Date','Certificate Issued'])
        sheet.append([1,'02.11.2026','Company','Pune','Laptop',5,'F6-1','0042',datetime(2026,11,3),'Generated'])
        sheet['F2']='=2+3';sheet['A2'].font=Font(name='Arial',bold=True,color='112233')
        sheet['A2'].fill=PatternFill('solid',fgColor='EEFFAA');sheet['A2'].alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        sheet['I2'].number_format='dd.mm.yyyy'
        buffer=BytesIO();book.save(buffer);book.close();return buffer.getvalue()

    def test_dates_and_details_use_requested_format_without_changing_numbers_or_formulas(self):
        updated=normalize_register_presentation(self.source());book=load_workbook(BytesIO(updated));sheet=book.active
        self.assertEqual(sheet['B2'].value,datetime(2026,11,2));self.assertEqual(sheet['I2'].value,datetime(2026,11,3))
        self.assertEqual(sheet['B2'].number_format,'dd/mm/yy');self.assertEqual(sheet['I2'].number_format,'dd/mm/yy')
        self.assertEqual(sheet['H2'].value,'0042');self.assertEqual(sheet['F2'].value,'=2+3')
        for row in sheet:
            for cell in row:
                self.assertEqual((cell.alignment.horizontal,cell.alignment.vertical),('left','top'))
        self.assertEqual(sheet['A2'].font.name,'Arial');self.assertTrue(sheet['A2'].font.bold)
        self.assertEqual(sheet['A2'].fill.fgColor.rgb,'00EEFFAA');self.assertTrue(sheet['A2'].alignment.wrap_text)
        book.close()

    def test_repeated_formatting_is_idempotent(self):
        updated=normalize_register_presentation(self.source())
        self.assertEqual(normalize_register_presentation(updated),updated)


if __name__=='__main__':unittest.main()
