"""Linked HDD tracking, failure recovery, and recycling identity in disposable files."""
from datetime import date
from io import BytesIO
import unittest
from unittest.mock import patch

from openpyxl import load_workbook

from certificate_core import read_batches, batch_content, select_register
from certificate_drafts import generate_draft, read_register, current_record_pdf
from destruction_certificate import request_for_inward
from issuance_safety import IssuanceTransaction, latest_issuance, recover_issuances
from tests import test_destruction_certificate as fixtures

AS_OF = fixtures.AS_OF
from workbook_store import WorkbookSession
from workflow_support import run_lock
from workspace_review import workspace_revision


class HddWorkbookTests(unittest.TestCase):
    setUp = fixtures.DestructionTests.setUp
    details = fixtures.DestructionTests.details

    def generate(self):
        identifier = request_for_inward(self.root, 'plant1', 'CLIENT1')
        record = next(r for r in read_register(self.root)['records'] if r['id'] == identifier)
        details = dict(record['details'], certificate_fields=self.details()['certificate_fields'])
        return generate_draft(self.root, 'plant1', details, identifier, as_of=AS_OF,
                              expected_revision=workspace_revision(self.root))

    def test_single_row_inward_inserts_hdd_line_before_next_client(self):
        book = load_workbook(self.paths[0]); sheet = book.active
        sheet.append([cell.value for cell in sheet[2]])
        sheet['A3'] = 'NEXT'; sheet['I3'] = 'F6-NEXT'; sheet['K3'] = '0304'
        book.save(self.paths[0]); book.close()
        before = read_batches(self.paths[0]); other = self.paths[1].read_bytes()
        result = self.generate(); after = read_batches(self.paths[0])
        self.assertEqual(after[0]['cod_no'], before[0]['cod_no'])
        self.assertEqual(after[0]['issued'], before[0]['issued'])
        self.assertEqual(batch_content(after[1]), batch_content(before[1]))
        self.assertEqual(after[1]['source_row'], 4)
        self.assertEqual((after[1]['cod_no'], after[1]['issued']), (before[1]['cod_no'], before[1]['issued']))
        self.assertEqual(after[0]['hdd_cod_no'], result['certificate_no'])
        self.assertEqual(self.paths[1].read_bytes(), other)
        book = load_workbook(self.paths[0]); sheet = book.active
        _, _, cols = select_register(book)
        hdd = sheet.cell(3, cols['cod_no'] + 1)
        self.assertEqual(hdd.value, result['certificate_no'])
        self.assertEqual(sheet.cell(3, cols['cod_date'] + 1).number_format, 'dd/mm/yy')
        self.assertEqual(hdd.border.top.style, 'thin')
        self.assertEqual(hdd.border.bottom.style, 'medium')
        self.assertEqual(sheet['A2'].alignment.vertical, 'top')
        book.close()

    def test_hdd_correction_keeps_both_numbers_and_updates_its_date(self):
        result = self.generate(); before = read_batches(self.paths[0])[0]
        record = read_register(self.root)['records'][0]
        details = dict(record['details'], cod_date='2026-10-18')
        corrected = generate_draft(self.root, 'plant1', details, result['record_id'], as_of=date(2026, 10, 18))
        after = read_batches(self.paths[0])[0]
        self.assertEqual(corrected['certificate_no'], result['certificate_no'])
        self.assertEqual((after['cod_no'], after['cod_date']), (before['cod_no'], before['cod_date']))
        self.assertEqual(after['hdd_cod_date'], '2026-10-18')

    def test_hdd_uses_styles_synthesized_from_merged_tracking_borders(self):
        from openpyxl.styles import Border, Side

        book = load_workbook(self.paths[0])
        sheet, _, cols = select_register(book)
        cod_col = cols['cod_no'] + 1
        sheet.merge_cells(start_row=2, end_row=3, start_column=cod_col, end_column=cod_col)
        sheet.cell(3, cod_col).border = Border(bottom=Side(style='thick'))
        book.save(self.paths[0])
        book.close()
        result = self.generate()
        batch = read_batches(self.paths[0])[0]
        self.assertEqual(batch['hdd_cod_no'], result['certificate_no'])
        self.assertNotEqual(batch['cod_no'], result['certificate_no'])

    def test_hdd_can_issue_first_without_a_recycling_number_or_date(self):
        book = load_workbook(self.paths[0])
        sheet, _, cols = select_register(book)
        for key in ('cod_no', 'cod_date'):
            cell = sheet.cell(2, cols[key] + 1)
            cell.value = None
            cell.number_format = 'General'
        sheet.cell(2, cols['issued'] + 1).value = 'N/A'
        book.save(self.paths[0])
        book.close()
        result = self.generate()
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch['issued'], batch['cod_no'], batch['cod_date']), ('N/A', '', None))
        self.assertEqual(batch['hdd_cod_no'], result['certificate_no'])
        self.assertEqual(batch['hdd_cod_date'], AS_OF.isoformat())
        from certificate_drafts import manual_rows
        self.assertEqual(manual_rows(self.root)[0]['certificate_status'], 'generated')

    def test_app_only_recycling_can_add_one_separate_hdd_record(self):
        from certificate_drafts import save_draft

        recycling = save_draft(self.root, None, dict(self.details(), certificate_type='standard',
                                                     certificate_fields={}, reference='LOCAL-REC'))
        identifier = request_for_inward(self.root, None, '', record_id=recycling)
        self.assertEqual(request_for_inward(self.root, None, '', record_id=recycling), identifier)
        records = {record['id']: record for record in read_register(self.root)['records']}
        self.assertEqual(records[recycling]['details']['certificate_type'], 'standard')
        self.assertEqual(records[identifier]['details']['destruction_source']['record_id'], recycling)
        self.assertEqual(len(records), 2)
        details = dict(records[identifier]['details'], certificate_fields=self.details()['certificate_fields'])
        with self.assertRaisesRegex(ValueError, 'already has a Data Destruction record'):
            generate_draft(self.root, None, details, as_of=AS_OF)
        result = generate_draft(self.root, None, details, identifier, as_of=AS_OF)
        self.assertEqual(result['created'], 1)
        self.assertFalse(result['workbook_updated'])
        self.assertEqual(len(read_register(self.root)['records']), 2)

    def test_recycling_correction_keeps_the_hdd_tracking_line(self):
        result = self.generate(); batch = read_batches(self.paths[0])[0]
        with WorkbookSession(self.paths[0], self.root / 'Data' / 'Workbook Backups') as session:
            session.commit(batch, batch['cod_no'], AS_OF, allow_reissue=True)
        after = read_batches(self.paths[0])[0]
        self.assertEqual(after['hdd_cod_no'], result['certificate_no'])
        book = load_workbook(self.paths[0]); sheet, _, cols = select_register(book)
        self.assertEqual(sheet.cell(3, cols['cod_no'] + 1).value, result['certificate_no']); book.close()

    def test_workbook_failure_rolls_back_pdf_and_local_record(self):
        identifier = request_for_inward(self.root, 'plant1', 'CLIENT1')
        before = self.paths[0].read_bytes(), (self.root / 'Data' / 'Custom Certificates.json').read_bytes()
        with patch.object(WorkbookSession, '_persist', side_effect=OSError('Excel is busy')):
            with self.assertRaisesRegex(OSError, 'Excel is busy'):
                self.generate()
        self.assertEqual(self.paths[0].read_bytes(), before[0])
        self.assertEqual((self.root / 'Data' / 'Custom Certificates.json').read_bytes(), before[1])
        self.assertIsNone(latest_issuance(self.root, 'manual', 'plant1', identifier))
        self.assertIsNone(current_record_pdf(self.root, read_register(self.root)['records'][0]))

    def test_linked_revision_check_still_rejects_external_changes_during_render(self):
        import certificate_drafts
        request_for_inward(self.root, 'plant1', 'CLIENT1')
        before = self.paths[0].read_bytes(), (self.root / 'Data' / 'Custom Certificates.json').read_bytes()
        render = certificate_drafts.make_template_certificate
        def change_other_plant(*args, **kwargs):
            render(*args, **kwargs)
            book = load_workbook(self.paths[1]); book.active['C2'] = 'External edit'; book.save(self.paths[1]); book.close()
        with patch.object(certificate_drafts, 'make_template_certificate', side_effect=change_other_plant):
            with self.assertRaisesRegex(ValueError, 'files changed'):
                self.generate()
        self.assertEqual(self.paths[0].read_bytes(), before[0])
        self.assertEqual((self.root / 'Data' / 'Custom Certificates.json').read_bytes(), before[1])
        self.assertIsNone(current_record_pdf(self.root, read_register(self.root)['records'][0]))

    def test_interruption_before_workbook_save_restores_companion_register(self):
        class Interrupted(BaseException): pass
        request_for_inward(self.root, 'plant1', 'CLIENT1')
        before = self.paths[0].read_bytes(), (self.root / 'Data' / 'Custom Certificates.json').read_bytes()
        with patch.object(WorkbookSession, '_persist', side_effect=Interrupted()):
            with self.assertRaises(Interrupted): self.generate()
        with run_lock(self.root): recover_issuances(self.root)
        self.assertEqual(self.paths[0].read_bytes(), before[0])
        self.assertEqual((self.root / 'Data' / 'Custom Certificates.json').read_bytes(), before[1])

    def test_interruption_after_workbook_save_finishes_pdf_and_local_tracking(self):
        class Interrupted(BaseException): pass
        with patch.object(IssuanceTransaction, 'mark_source_committed', side_effect=Interrupted()):
            with self.assertRaises(Interrupted): self.generate()
        with run_lock(self.root): recover_issuances(self.root)
        record = read_register(self.root)['records'][0]
        self.assertIsNotNone(current_record_pdf(self.root, record))
        self.assertEqual(read_batches(self.paths[0])[0]['hdd_cod_no'], record['details']['cod_no'])

    def test_refresh_adds_tracking_for_an_older_hdd_without_allocating_a_number(self):
        result = self.generate()
        book = load_workbook(self.paths[0]); sheet, _, cols = select_register(book)
        sheet.cell(3, cols['cod_no'] + 1).value = None
        sheet.cell(3, cols['cod_date'] + 1).value = None
        book.save(self.paths[0]); book.close()
        from workbook_metadata import read_details, _replace_details
        saved = read_details(self.paths[0])
        for value in saved.values():
            value.pop('_hdd_cod_no', None); value.pop('_hdd_cod_date', None)
        self.paths[0].write_bytes(_replace_details(self.paths[0].read_bytes(), saved))
        before = {str(path): path.read_bytes() for path in (self.root / 'Certificates').rglob('*.pdf')}
        from run_certificates import run
        self.assertEqual(run(self.paths, self.root, check_only=True, as_of=AS_OF, progress=None)[0], 0)
        self.assertEqual(read_batches(self.paths[0])[0]['hdd_cod_no'], result['certificate_no'])
        self.assertEqual({str(path): path.read_bytes() for path in (self.root / 'Certificates').rglob('*.pdf')}, before)


if __name__ == '__main__': unittest.main()
