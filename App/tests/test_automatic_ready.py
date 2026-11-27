"""Automatic Ready and genuinely empty statuses, using disposable workbooks."""
from datetime import timedelta
from unittest.mock import patch

from openpyxl import load_workbook

from certificate_core import read_batches, register_header
from workbook_store import WorkbookSession
import web_app
from tests.test_certificate_workflow import AutomationTestBase, AS_OF, make_workbook


class AutomaticReadyTests(AutomationTestBase):
    def refresh_sheet(self):
        with WorkbookSession(self.plant1, self.base / 'Data' / 'Workbook Backups') as session:
            session.ensure_alert_formatting('Inward Register', as_of=AS_OF)

    def statuses(self):
        book = load_workbook(self.plant1)
        try:
            column=register_header(book.active)[1]['issued']+1
            return [book.active.cell(row, column).value for row in range(2, book.active.max_row + 1)]
        finally:
            book.close()

    def test_incomplete_rows_auto_alert_then_become_ready_after_details_are_filled(self):
        make_workbook(self.plant1, [{'sn':'1001', 'status':'', 'supplier':''}])
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Alert'])
        book = load_workbook(self.plant1)
        book.active['C2'] = 'Completed company'
        book.active['I2'] = 'FORM-1001'
        book.save(self.plant1)
        book.close()
        original = read_batches(self.plant1)[0]
        self.refresh_sheet()
        current = read_batches(self.plant1)[0]
        self.assertEqual(current, dict(original, issued='Ready', _recycling_status='Ready'))
        self.assertEqual(self.statuses(), ['Ready'])
        after = self.plant1.read_bytes()
        self.refresh_sheet()
        self.assertEqual(self.plant1.read_bytes(), after, 'Unchanged refresh must not rewrite the workbook.')
        self.assertTrue(list((self.base / 'Data' / 'Workbook Backups').glob('*.xlsx')))

    def test_legacy_blank_is_cleared_and_optional_output_values_do_not_block_ready(self):
        make_workbook(self.plant1, [
            {'sn':'A', 'status':'Blank', 'supplier':''},
            {'sn':'B', 'status':'Blank', 'reference':'', 'cod':'', 'cod_date':None},
            {'sn':'C', 'status':'   ', 'supplier':''},
        ])
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Alert', 'Ready', 'Alert'])
        book = load_workbook(self.plant1)
        try:
            column=book.active.cell(1,register_header(book.active)[1]['issued']+1).column_letter
            validation = next(v for v in book.active.data_validations.dataValidation if f'{column}2:{column}10000' in str(v.sqref))
            self.assertNotIn('Blank', validation.formula1)
            self.assertTrue(validation.allowBlank)
        finally:
            book.close()

    def test_missing_or_invalid_source_details_never_get_auto_approved(self):
        changes = [{'supplier':''}, {'date':None},
                   {'weight':-1}, {'weight':0}, {'item':''},
                   {'date':AS_OF + timedelta(days=1)}, {'cod_date':'bad date'},
                   {'cod_date':AS_OF + timedelta(days=1)}, {'cod':'=1+1'}]
        make_workbook(self.plant1, [dict(change, sn=str(i), status='') for i, change in enumerate(changes)])
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Alert'] * len(changes))

    def test_explicit_statuses_are_preserved(self):
        statuses = ['Ready', 'Alert', 'N/A', 'Generated', 'Generated Externally', 'Completed', 'User value']
        make_workbook(self.plant1, [dict(sn=str(i), status=status,
                                         supplier='' if status == 'Alert' else 'Client')
                                    for i, status in enumerate(statuses)])
        self.refresh_sheet()
        self.assertEqual(self.statuses(), statuses)

    def test_book_my_junk_is_always_na_and_never_auto_ready(self):
        make_workbook(self.plant1, [{
            'sn':'B-MJ', 'supplier':'Reverse E - Commerce Pvt. Ltd. (Book My Junk)',
            'form6':'FORM-MJ', 'status':'',
        }])
        self.refresh_sheet()
        self.assertEqual(read_batches(self.plant1)[0]['issued'], 'N/A')
        self.assertEqual(self.statuses(), ['N/A'])

        book = load_workbook(self.plant1)
        book.active.cell(2,register_header(book.active)[1]['issued']+1).value = 'Ready'
        book.save(self.plant1)
        book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['N/A'])
        self.assertEqual(read_batches(self.plant1)[0]['issued'], 'N/A')

    def test_form6_can_be_entered_at_creation_for_each_template(self):
        make_workbook(self.plant1, [dict(sn=kind, status='', form6='') for kind in ('sita', 'siemens', 'standard')])
        book = load_workbook(self.plant1)
        book.active['N1'] = 'Certificate Type'
        for row, kind in enumerate(('sita', 'siemens', 'standard'), 2):
            book.active.cell(row, 14, kind)
        book.save(self.plant1)
        book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Ready', 'Ready', 'Ready'])
        book = load_workbook(self.plant1)
        book.active['I2'] = 'SITA-F6'
        book.active['I3'] = 'SIEMENS-F6'
        book.save(self.plant1)
        book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Ready', 'Ready', 'Ready'])

    def test_protected_merged_and_formula_status_cells_are_kept(self):
        for kind in ('protected', 'merged', 'formula'):
            with self.subTest(kind=kind):
                make_workbook(self.plant1, [dict(sn='A', status='')])
                book = load_workbook(self.plant1)
                if kind == 'protected':
                    book.active.protection.sheet = True
                elif kind == 'merged':
                    book.active.merge_cells('M2:N2')
                else:
                    book.active['M2'] = '=IF(A2="","","Ready")'
                before = book.active['M2'].value
                book.save(self.plant1)
                book.close()
                self.refresh_sheet()
                self.assertEqual(self.statuses(), [before])

    def test_vertical_shared_status_becomes_ready_without_unmerging(self):
        make_workbook(self.plant1, [dict(sn='A', status='')])
        book = load_workbook(self.plant1)
        try:
            book.active.append([None, None, None, None, 'Extra material', 1, 3, 3])
            book.active.merge_cells('M2:M3')
            book.save(self.plant1)
        finally:
            book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Ready', None])
        book = load_workbook(self.plant1)
        try:
            column=book.active.cell(1,register_header(book.active)[1]['issued']+1).column_letter
            self.assertIn(f'{column}2:{column}3', {str(area) for area in book.active.merged_cells.ranges})
        finally:
            book.close()

    def test_continuation_materials_and_unused_rows_are_respected(self):
        make_workbook(self.plant1, [dict(sn='A', status='')])
        book = load_workbook(self.plant1)
        book.active.append([None, None, None, None, 'Extra material', None, None, None])
        book.active.cell(5, 13).value = None
        book.active.cell(5, 13).number_format = '@'
        book.save(self.plant1)
        book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Alert', None, None, None])
        book = load_workbook(self.plant1)
        book.active['H3'] = 3
        book.save(self.plant1)
        book.close()
        self.refresh_sheet()
        self.assertEqual(self.statuses(), ['Ready', None, None, None])

    def test_refresh_reports_ready_without_creating_pdf_and_only_checks_changed_files(self):
        make_workbook(self.plant1, [dict(sn='A', status='', supplier='')])
        make_workbook(self.plant2, [dict(sn='B', status='N/A')])
        self.enterContext(patch.object(web_app, 'ROOT', self.base))
        self.enterContext(patch.object(web_app, 'RESULTS_PATH', self.base / 'Data' / 'Reports' / 'Latest Results.json'))
        self.enterContext(patch.object(web_app, 'WORKBOOK_CHECK_STAMPS', {}))
        self.assertTrue(web_app.check_workbook_changes()['changed'])
        self.assertFalse(web_app.check_workbook_changes()['changed'])

        book = load_workbook(self.plant1)
        book.active['C2'] = 'Filled in Excel'
        book.save(self.plant1)
        book.close()
        changed = web_app.check_workbook_changes()
        self.assertTrue(changed['changed'])
        row = next(row for row in changed['state']['rows'] if row['sn'] == 'A')
        self.assertEqual(row['certificate_status'], 'ready')
        self.assertEqual(self.statuses(), ['Ready'])
        self.assertFalse(list((self.base / 'Certificates').rglob('*.pdf')))
        self.assertFalse(web_app.check_workbook_changes()['changed'])

    def test_creation_uses_the_auto_ready_source_snapshot_and_skips_unfinished_rows(self):
        make_workbook(self.plant1, [dict(sn='A', status=''), dict(sn='B', status='', supplier='')])
        make_workbook(self.plant2, [dict(sn='C', status='N/A')])
        created, review, _ = self.execute()
        self.assertEqual((created, review), (1, 1))
        self.assertEqual(self.statuses(), ['Generated', 'Alert'])
