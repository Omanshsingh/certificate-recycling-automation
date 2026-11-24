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
        self.assertEqual(sheet['A2'].fill.fgColor.rgb,'00EEFFAA')
        self.assertTrue(sheet['A2'].alignment.wrap_text);self.assertFalse(sheet['B2'].alignment.wrap_text)
        book.close()

    def test_repeated_formatting_is_idempotent(self):
        updated=normalize_register_presentation(self.source())
        self.assertEqual(normalize_register_presentation(updated),updated)

    def test_existing_wrap_settings_and_row_heights_survive_status_save(self):
        book = load_workbook(BytesIO(self.source()))
        sheet = book.active
        sheet.column_dimensions['C'].width = 17
        sheet.column_dimensions['C'].alignment = Alignment(wrap_text=True)
        sheet.column_dimensions['E'].width = 14
        sheet.merge_cells('C2:C4')
        sheet['C2'] = '\n'.join(['Company branch name'] * 8)
        sheet['C2'].alignment = Alignment(vertical='center', wrap_text=True)
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
        self.assertTrue(sheet['E2'].alignment.shrink_to_fit)
        self.assertTrue(sheet['C2'].alignment.wrap_text)
        self.assertTrue(all(not sheet[coordinate].alignment.wrap_text for coordinate in ('B2', 'E2', 'F2')))
        self.assertEqual(sheet.row_dimensions[2].height, 15)
        self.assertEqual(sheet.row_dimensions[3].height, 15)
        self.assertIsNone(result['Unrelated']['A1'].alignment.wrap_text)
        self.assertTrue(sheet.column_dimensions['C'].alignment.wrap_text)
        self.assertTrue(all('Generated Externally' not in rule.formula1
                            for rule in sheet.data_validations.dataValidation))

    def test_auto_fit_only_touches_the_new_inward_and_its_wrapped_cells(self):
        book = Workbook()
        sheet = book.active
        sheet.append(['SN','Inward Date','Supplier Name','Location','Item','Weight (KG)','Form 6 No','COD No','COD Date','Certificate Issued'])
        sheet.append([1,'02.11.2026','A company name long enough to wrap across the narrow column','Pune','Laptop',5,'F6-1','0042','03/11/26','Ready'])
        sheet.append([2,'03.11.2026','Another company name long enough to wrap across the narrow column','Mumbai','Monitor',8,'F6-2','0043','04/11/26','Ready'])
        sheet.column_dimensions['C'].width = 15
        sheet['C2'].alignment = Alignment(wrap_text=True, vertical='top')
        sheet['C3'].alignment = Alignment(wrap_text=False, vertical='top')
        sheet.row_dimensions[2].height = 15
        sheet.row_dimensions[3].height = 15
        buffer = BytesIO()
        book.save(buffer)
        book.close()

        updated = normalize_register_presentation(buffer.getvalue(), fit_inward=('Sheet', '1'))
        result = load_workbook(BytesIO(updated))
        self.addCleanup(result.close)
        sheet = result.active
        self.assertTrue(sheet['C2'].alignment.wrap_text)
        self.assertFalse(sheet['C3'].alignment.wrap_text)
        self.assertGreater(sheet.row_dimensions[2].height, 15)
        self.assertEqual(sheet.row_dimensions[3].height, 15)

    def test_refit_shrinks_replaced_text_and_preserves_current_wrapping(self):
        book = load_workbook(BytesIO(self.source()))
        sheet = book.active
        sheet['C2'] = 'Short client name'
        sheet['C2'].alignment = Alignment(wrap_text=True)
        sheet.column_dimensions['C'].width = 28
        sheet['J2'] = 'Generated Externally'
        sheet['J2'].alignment = Alignment(wrap_text=True)
        sheet.column_dimensions['J'].width = 13
        sheet.row_dimensions[1].height = 55
        sheet.row_dimensions[2].height = 80
        other = book.create_sheet('Other')
        other['A2'] = 'Leave this sheet as authored'
        other.row_dimensions[2].height = 80

        def fit():
            buffer = BytesIO()
            book.save(buffer)
            return normalize_register_presentation(buffer.getvalue(), wrap_only=True, fit_register=True)

        wrapped = fit()
        with_height = load_workbook(BytesIO(wrapped))
        self.assertGreater(with_height.active.row_dimensions[2].height, 15)
        self.assertLess(with_height.active.row_dimensions[2].height, 40)
        with_height.close()
        sheet['J2'] = 'Completed'
        shortened = fit()
        book.close()
        result = load_workbook(BytesIO(shortened))
        self.addCleanup(result.close)
        self.assertIsNone(result.active.row_dimensions[2].height)
        self.assertEqual(result.active.row_dimensions[1].height, 55)
        self.assertEqual(result['Other'].row_dimensions[2].height, 80)
        self.assertEqual(normalize_register_presentation(shortened, wrap_only=True, fit_register=True), shortened)

    def test_refit_recomputes_the_combined_height_of_merged_cells(self):
        book = load_workbook(BytesIO(self.source()))
        sheet = book.active
        sheet['E3'] = 'Printer'
        sheet['F3'] = 5
        sheet.merge_cells('C2:C3')
        sheet['C2'] = 'First line\nSecond line\nThird line'
        sheet['C2'].alignment = Alignment(wrap_text=True)
        sheet.column_dimensions['C'].width = 24
        for row in (2, 3):
            sheet.row_dimensions[row].height = 70
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        fitted = normalize_register_presentation(buffer.getvalue(), fit_inward=('Sheet', '1'))
        book = load_workbook(BytesIO(fitted))
        self.assertGreaterEqual(sum(book.active.row_dimensions[row].height for row in (2, 3)), 45)
        book.active['C2'] = 'Company'
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        shortened = normalize_register_presentation(buffer.getvalue(), fit_inward=('Sheet', '1'))
        resized = load_workbook(BytesIO(shortened))
        self.addCleanup(resized.close)
        self.assertIn('C2:C3', resized.active.merged_cells)
        self.assertTrue(all(resized.active.row_dimensions[row].height is None for row in (2, 3)))


if __name__=='__main__':unittest.main()
