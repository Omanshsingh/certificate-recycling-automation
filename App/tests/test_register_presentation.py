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

    def test_wrap_all_register_fields_and_expand_fixed_and_merged_rows(self):
        book = load_workbook(BytesIO(self.source()))
        sheet = book.active
        sheet.column_dimensions['C'].width = 17
        sheet.column_dimensions['E'].width = 14
        sheet.merge_cells('C2:C4')
        sheet['C2'] = '\n'.join(['Company branch name'] * 8)
        sheet['E2'] = 'EquipmentID' * 12
        sheet['E2'].alignment = Alignment(horizontal='right', shrink_to_fit=True)
        sheet.row_dimensions[2].height = 15
        sheet.row_dimensions[3].height = 15
        book.create_sheet('Unrelated')['A1'] = 'Other sheet'
        buffer = BytesIO()
        book.save(buffer)
        book.close()

        updated = normalize_register_presentation(buffer.getvalue(), wrap_only=True, include_validation=True)
        self.assertEqual(normalize_register_presentation(updated, wrap_only=True, include_validation=True), updated)
        result = load_workbook(BytesIO(updated))
        self.addCleanup(result.close)
        sheet = result.worksheets[0]
        self.assertIn('C2:C4', sheet.merged_cells)
        self.assertEqual(sheet['E2'].value, 'EquipmentID' * 12)
        self.assertEqual(sheet['F2'].value, '=2+3')
        self.assertEqual(sheet['B2'].value, '02.11.2026')
        self.assertEqual(sheet['E2'].alignment.horizontal, 'right')
        self.assertFalse(sheet['E2'].alignment.shrink_to_fit)
        self.assertTrue(all(sheet[coordinate].alignment.wrap_text for coordinate in ('B2', 'C2', 'E2', 'F2')))
        self.assertGreater(sheet.row_dimensions[2].height, 30)
        self.assertGreaterEqual(sum(sheet.row_dimensions[row].height for row in range(2, 5)), 120)
        self.assertIsNone(result['Unrelated']['A1'].alignment.wrap_text)
        self.assertTrue(sheet.column_dimensions['C'].alignment.wrap_text)
        self.assertTrue(all('Generated Externally' not in rule.formula1
                            for rule in sheet.data_validations.dataValidation))


if __name__=='__main__':unittest.main()
