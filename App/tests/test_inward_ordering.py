"""Backdated register writes preserve whole inwards and their certificate identity."""
from copy import copy
from datetime import date
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from certificate_core import batch_content, read_batches
from plant_workbook_template import plant_workbook_bytes
from workbook_store import WorkbookSession, sort_register_by_inward_date, stage_certificate_details


class InwardOrderingTests(unittest.TestCase):
    def fixture(self):
        book = load_workbook(BytesIO(plant_workbook_bytes('plant1')))
        sheet = book.active
        for serial, first, last, inward, cod in (
            ('late', 2, 3, date(2026, 10, 10), '00101'),
            ('early', 4, 4, date(2026, 10, 1), '00102'),
            ('same', 5, 6, date(2026, 10, 1), '00103'),
        ):
            for row in range(first, last + 1):
                for col in range(1, 19):
                    sheet.cell(row, col)._style = copy(sheet.cell(2, col)._style)
                    sheet.cell(row, col).border = Border(top=Side(style='medium', color='FF000000'), bottom=Side(style='medium', color='FF000000'))
                sheet.cell(row, 5, f'{serial} asset {row - first + 1}')
                sheet.cell(row, 6, 2)
                sheet.cell(row, 7, row)
                sheet.cell(row, 8, f'=F{row}*G{row}')
            for col, value in {1: serial, 2: inward, 3: serial + ' client', 4: 'Pune', 9: serial + '-F6',
                               10: f'=SUM(H{first}:H{last})', 16: cod, 17: date(2026, 10, 12), 18: 'Generated'}.items():
                sheet.cell(first, col, value)
                if last > first:
                    sheet.merge_cells(start_row=first, end_row=last, start_column=col, end_column=col)
            sheet.row_dimensions[first].height = first + 30
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        return buffer.getvalue()

    def test_sort_preserves_assets_merges_formulas_cod_and_package_members(self):
        data = self.fixture()
        output = BytesIO()
        with ZipFile(BytesIO(data)) as source, ZipFile(output, 'w') as target:
            for info in source.infolist():
                target.writestr(info, source.read(info.filename))
            target.writestr('custom/preserved.bin', b'preserve exactly')
        data = output.getvalue()
        original = {batch['case_id']: batch for batch in read_batches(BytesIO(data))}
        updated = sort_register_by_inward_date(data)
        parsed = read_batches(BytesIO(updated))
        self.assertEqual([batch['case_id'] for batch in parsed], ['early', 'same', 'late'])
        for batch in parsed:
            self.assertEqual(batch_content(batch), batch_content(original[batch['case_id']]))
            self.assertEqual(batch['cod_no'], original[batch['case_id']]['cod_no'])
            self.assertEqual(batch['cod_date'], original[batch['case_id']]['cod_date'])
        book = load_workbook(BytesIO(updated))
        sheet = book.active
        self.assertEqual(sheet['H5'].value, '=F5*G5')
        self.assertEqual(sheet['J5'].value, '=SUM(H5:H6)')
        self.assertEqual(sheet.row_dimensions[5].height, 32)
        self.assertIn('A5:A6', {str(area) for area in sheet.merged_cells.ranges})
        self.assertEqual(sheet['E5'].border.bottom.color.rgb, 'FFD1D5DB')
        self.assertEqual(sheet['E6'].border.top.color.rgb, 'FFD1D5DB')
        self.assertEqual(sheet['E6'].border.bottom.color.rgb, 'FF000000')
        self.assertEqual(sheet['E6'].border.bottom.style, 'medium')
        book.close()
        with ZipFile(BytesIO(updated)) as archive:
            self.assertEqual(archive.read('custom/preserved.bin'), b'preserve exactly')

    def test_absolute_references_move_with_the_inward(self):
        book = load_workbook(BytesIO(self.fixture()))
        book.active['J2'] = '=SUM($H$2:$H$3)'
        output = BytesIO()
        book.save(output)
        book.close()
        sorted_data = sort_register_by_inward_date(output.getvalue())
        book = load_workbook(BytesIO(sorted_data))
        self.assertEqual(book.active['J5'].value, '=SUM($H$5:$H$6)')
        book.close()

    def test_date_correction_reorders_then_commit_updates_the_correct_record(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'register.xlsx'
            path.write_bytes(sort_register_by_inward_date(self.fixture()))
            original = next(batch for batch in read_batches(path) if batch['case_id'] == 'late')
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                staged, edited = stage_certificate_details(session, {'case_id': 'LATE', 'inward_date': '2026-09-20'})
                self.assertEqual(edited['case_id'], 'late')
                self.assertEqual(edited['source_row'], 2)
                self.assertEqual(original['source_row'], 5)
                session.commit(original, '00101', date(2026, 10, 14), staged_data=staged, allow_reissue=True)
            records = read_batches(path)
            self.assertEqual([batch['case_id'] for batch in records], ['late', 'early', 'same'])
            late = records[0]
            self.assertEqual(late['cod_no'], '00101')
            self.assertEqual(late['cod_date'], date(2026, 10, 14))
            self.assertEqual(records[1]['cod_no'], '00102')
            self.assertEqual(records[1]['cod_date'], date(2026, 10, 12))

    def test_cross_inward_formula_keeps_date_edit_without_reordering_or_touching_source(self):
        book = load_workbook(BytesIO(self.fixture()))
        book.active['K2'] = '=H4'
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        data = buffer.getvalue()
        session = SimpleNamespace(data=data, path=Path('fixture.xlsx'))
        staged, edited = stage_certificate_details(session, {'case_id': 'late', 'inward_date': '2026-10-09'})
        self.assertEqual(edited['inward_date'], date(2026, 10, 9))
        self.assertEqual(edited['source_row'], 2)
        self.assertIn('different inward', session.order_warning)
        parsed = read_batches(BytesIO(staged))
        self.assertEqual([batch['case_id'] for batch in parsed], ['late', 'early', 'same'])
        self.assertEqual([batch['cod_no'] for batch in parsed], ['00101', '00102', '00103'])
        staged_book = load_workbook(BytesIO(staged))
        self.assertEqual(staged_book.active['K2'].value, '=H4')
        staged_book.close()
        for original, saved in zip(read_batches(BytesIO(data))[1:], parsed[1:]):
            self.assertEqual(batch_content(saved), batch_content(original))
            self.assertEqual(saved['cod_date'], original['cod_date'])
        self.assertEqual(session.data, data)

    def test_writable_refresh_reorders_manual_edits_and_refreshes_source_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'register.xlsx'
            path.write_bytes(self.fixture())
            batches = read_batches(path)
            late = batches[0]
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                session.ensure_alert_formatting('Sheet1', batches)
                self.assertIsNone(session.order_warning)
            self.assertEqual(late['source_row'], 5)
            self.assertEqual([batch['case_id'] for batch in read_batches(path)], ['early', 'same', 'late'])

    def test_sort_is_byte_stable_after_first_update(self):
        data = sort_register_by_inward_date(self.fixture())
        self.assertEqual(sort_register_by_inward_date(data), data)

    def test_render_only_preview_does_not_reorder_unsupported_source(self):
        book = load_workbook(BytesIO(self.fixture()))
        book.active['K2'] = '=H4'
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        session = SimpleNamespace(data=buffer.getvalue(), path=Path('fixture.xlsx'))
        staged, batch = stage_certificate_details(session, {'case_id': 'late', 'reference': 'preview'}, validate_cells=False)
        self.assertEqual(batch['source_row'], 2)
        self.assertEqual([row['case_id'] for row in read_batches(BytesIO(staged))], ['late', 'early', 'same'])
        self.assertEqual(read_batches(BytesIO(session.data))[0]['reference'], '')

    def test_dynamic_formula_cannot_keep_a_hidden_reference_to_the_old_row(self):
        book = load_workbook(BytesIO(self.fixture()))
        book.active['K2'] = '=INDIRECT("H4")'
        buffer = BytesIO()
        book.save(buffer)
        book.close()
        with self.assertRaisesRegex(ValueError, 'physical row positions'):
            sort_register_by_inward_date(buffer.getvalue())

    def test_status_update_follows_unique_serial_after_its_excel_row_moves(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'register.xlsx'
            path.write_bytes(sort_register_by_inward_date(self.fixture()))
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                # Row 2 was the old location of late, and now belongs to early.
                session.set_issued_status('LATE', 'Sheet1', 2, '00101', 'Generated', 'Completed')
            batches = {batch['case_id']: batch for batch in read_batches(path)}
            self.assertEqual(batches['late']['issued'], 'Completed')
            self.assertEqual(batches['early']['issued'], 'Generated')
            self.assertEqual(batches['early']['cod_no'], '00102')

    def test_casefold_duplicate_refuses_status_update_before_persistence(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'register.xlsx'
            book = load_workbook(BytesIO(self.fixture()))
            book.active['A4'] = 'LATE'
            book.save(path)
            book.close()
            before = path.read_bytes()
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                with self.assertRaisesRegex(ValueError, 'Serial Number late appears more than once'):
                    session.set_issued_status('late', 'Sheet1', 2, '00101', 'Generated', 'Completed')
            self.assertEqual(path.read_bytes(), before)

    def test_repeated_total_serial_cells_are_summaries_not_duplicate_inwards(self):
        book = load_workbook(BytesIO(plant_workbook_bytes('plant1')))
        sheet = book.active
        for serial, first, weight in [('FIRST', 2, 3), ('SECOND', 4, 7)]:
            sheet.cell(first, 1, serial)
            sheet.cell(first, 2, date(2026, 10, 1))
            sheet.cell(first, 3, serial + ' client')
            sheet.cell(first, 5, 'Laptop')
            sheet.cell(first, 8, weight)
            sheet.cell(first, 9, serial + '-F6')
            sheet.cell(first, 18, 'Ready')
            sheet.cell(first + 1, 1, 'Total')
            sheet.cell(first + 1, 8, weight)
        output = BytesIO()
        book.save(output)
        book.close()
        data = output.getvalue()
        parsed = read_batches(BytesIO(data))
        self.assertEqual([batch['summary_rows'] for batch in parsed], [[3], [5]])
        session = SimpleNamespace(data=data, path=Path('fixture.xlsx'))
        staged, batch = stage_certificate_details(session, {'case_id': 'first', 'reference': 'preview'}, validate_cells=False)
        self.assertEqual(batch['case_id'], 'FIRST')
        self.assertEqual(batch['source_row'], 2)
        self.assertEqual(batch['reference'], 'preview')
        self.assertEqual([row['case_id'] for row in read_batches(BytesIO(staged))], ['FIRST', 'SECOND'])
        # Actual new-batch creation must also distinguish summary labels from
        # real record identifiers and append without modifying either inward.
        staged, added = stage_certificate_details(session, {
            'case_id': 'THIRD', 'supplier': 'Third client', 'inward_date': '2026-10-02',
            'item': 'Monitor', 'weight': 2, 'form6': 'THIRD-F6',
        })
        self.assertEqual(added['case_id'], 'THIRD')
        self.assertEqual([row['case_id'] for row in read_batches(BytesIO(staged))], ['FIRST', 'SECOND', 'THIRD'])

    def test_actual_casefold_duplicate_still_blocks_staging(self):
        book = load_workbook(BytesIO(self.fixture()))
        book.active['A4'] = 'LATE'
        output = BytesIO()
        book.save(output)
        book.close()
        session = SimpleNamespace(data=output.getvalue(), path=Path('fixture.xlsx'))
        with self.assertRaisesRegex(ValueError, 'Serial Number late appears more than once in Excel rows 2, 4'):
            stage_certificate_details(session, {'case_id': 'late', 'reference': 'preview'}, validate_cells=False)

    def test_backdated_new_record_does_not_check_unrelated_original_output_cells(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'register.xlsx'
            book = load_workbook(BytesIO(plant_workbook_bytes('plant1')))
            sheet = book.active
            for column, value in {1: 'EXISTING', 2: date(2026, 10, 1), 3: 'Existing client',
                                  5: 'Laptop', 8: 5, 9: 'OLD-F6', 16: '00101',
                                  17: '=DATE(2026,10,12)', 18: 'Generated'}.items():
                sheet.cell(2, column, value)
            book.save(path)
            book.close()
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                original = read_batches(BytesIO(session.data))[0]
                staged, added = stage_certificate_details(session, {
                    'case_id': 'NEW', 'supplier': 'New client', 'inward_date': '2026-09-20',
                    'item': 'Monitor', 'weight': 2, 'form6': 'NEW-F6',
                })
                self.assertEqual(added['source_row'], original['source_row'])
                session.commit(added, '00102', date(2026, 10, 15), staged_data=staged, new_batch=True)
            saved = read_batches(path)
            self.assertEqual([batch['case_id'] for batch in saved], ['NEW', 'EXISTING'])
            self.assertEqual(saved[0]['cod_no'], '00102')
            self.assertEqual(saved[0]['cod_date'], date(2026, 10, 15))
            self.assertEqual(saved[1]['cod_no'], '00101')
            self.assertEqual(saved[1]['cod_date'], '=DATE(2026,10,12)')
            book = load_workbook(path)
            self.assertEqual(book.active['Q3'].value, '=DATE(2026,10,12)')
            self.assertTrue(book.active['Q3'].protection.locked)
            book.close()
            # The original-cell protection check still applies to that existing
            # record even if a render-only stage removes its formula in memory.
            with WorkbookSession(path, Path(folder) / 'backups') as session:
                existing = next(batch for batch in read_batches(BytesIO(session.data)) if batch['case_id'] == 'EXISTING')
                staged, _ = stage_certificate_details(session, {
                    'case_id': 'EXISTING', 'cod_date': '2026-10-15',
                }, validate_cells=False)
                with self.assertRaisesRegex(ValueError, 'formula or protected cell'):
                    session.commit(existing, '00101', date(2026, 10, 15), staged_data=staged, allow_reissue=True)


if __name__ == '__main__':
    unittest.main()
