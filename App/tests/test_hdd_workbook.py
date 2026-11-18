"""Linked HDD tracking, failure recovery, and recycling identity in disposable files."""
from contextlib import closing
from datetime import date
from io import BytesIO
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

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

    def test_hdd_only_uses_existing_row_and_preserves_the_hdd_when_recycling_is_uploaded(self):
        path = self.paths[0]
        fixtures.write_register(path, case_id='CLIENT1', cod='', status='N/A')
        result = self.generate()
        hdd_cod = result['certificate_no']
        with closing(load_workbook(path)) as book:
            sheet, _, cols = select_register(book)
            self.assertEqual(sheet.max_row, 2)
            self.assertEqual(sheet.cell(2, cols['cod_no'] + 1).value, hdd_cod)
        before = read_batches(path)[0]
        self.assertEqual((before['cod_no'], before['hdd_cod_no'], before['issued']), ('', hdd_cod, 'N/A'))
        with WorkbookSession(path, self.root / 'backups') as session:
            session.set_issued_status('CLIENT1', before['source_sheet'], 2, '', 'N/A', 'Generated Externally',
                                      external={'cod': '0499', 'date': AS_OF, 'link': {'path': 'recycling.pdf'}})
        after = read_batches(path)[0]
        self.assertEqual((after['cod_no'], after['hdd_cod_no']), ('0499', hdd_cod))
        with closing(load_workbook(path)) as book:
            sheet, _, cols = select_register(book)
            self.assertEqual(sheet.cell(3, cols['cod_no'] + 1).value, hdd_cod)

    def test_existing_hdd_only_tracking_row_is_compacted_without_moving_next_clients_data(self):
        from types import SimpleNamespace
        from hdd_workbook import stage_hdd_tracking
        path = self.paths[0]
        fixtures.write_register(path, case_id='CLIENT1', cod='', status='Ready')
        initial = stage_hdd_tracking(SimpleNamespace(data=path.read_bytes()), 'CLIENT1', '0303', AS_OF)
        with closing(load_workbook(BytesIO(initial))) as book:
            sheet, _, cols = select_register(book)
            self.assertEqual(sheet.cell(3, cols['cod_no'] + 1).value, '0303')
            sheet.cell(2, cols['issued'] + 1, 'N/A')
            sheet.append([cell.value for cell in sheet[2]])
            sheet['A4'] = 'NEXT'
            sheet.cell(4, cols['issued'] + 1, 'Ready')
            book.save(path)
        before = read_batches(path)
        migrated = stage_hdd_tracking(SimpleNamespace(data=path.read_bytes()), 'CLIENT1', '0303', AS_OF)
        after = read_batches(BytesIO(migrated))
        self.assertEqual(after[1]['source_row'], 3)
        self.assertEqual(batch_content(after[1]), batch_content(before[1]))
        self.assertEqual(len(after[0]['items']), 1)
        self.assertEqual(after[0]['hdd_cod_no'], '0303')
        self.assertTrue(after[0]['hdd_on_main_row'])

        from hdd_workbook import recover_hdd_tracking_link
        from inward_serials import links_from
        from workbook_metadata import write_details
        raw = write_details(path.read_bytes(), before[0]['source_sheet'], 'CLIENT1',
                            {'_hdd_cod_no': '', '_hdd_cod_date': ''})
        rows = read_batches(BytesIO(raw))
        details = {key: rows[0][key] for key in ('supplier', 'form6', 'reference', 'inward_date')}
        record = {'id': 'stale-hdd', 'details': dict(details, cod_no='0303')}
        stale = SimpleNamespace(data=raw)
        self.assertIsNone(recover_hdd_tracking_link(
            stale, rows, {'id': 'stale-hdd', 'details': dict(record['details'], form6='Different')}, AS_OF))
        restored = recover_hdd_tracking_link(stale, rows, record, AS_OF)
        self.assertEqual(links_from(restored)['manual_sources']['stale-hdd'], 'CLIENT1')
        self.assertTrue(read_batches(BytesIO(restored))[0]['hdd_on_main_row'])

    def test_hdd_correction_keeps_both_numbers_and_updates_its_date(self):
        result = self.generate(); before = read_batches(self.paths[0])[0]
        record = read_register(self.root)['records'][0]
        details = dict(record['details'], cod_date='2026-10-18')
        corrected = generate_draft(self.root, 'plant1', details, result['record_id'], as_of=date(2026, 10, 18))
        after = read_batches(self.paths[0])[0]
        self.assertEqual(corrected['certificate_no'], result['certificate_no'])
        self.assertEqual((after['cod_no'], after['cod_date']), (before['cod_no'], before['cod_date']))
        self.assertEqual(after['hdd_cod_date'], '2026-10-18')

    def test_hdd_row_insertion_preserves_excel_shared_weight_formulas(self):
        path = self.paths[0]
        with closing(load_workbook(path)) as book:
            sheet = book.active
            sheet.append([cell.value for cell in sheet[2]])
            sheet['A3'] = 'NEXT'
            sheet['I3'] = 'F6-NEXT'
            sheet['K3'] = '0304'
            sheet['H2'] = '=F2*G2'
            sheet['H3'] = '=F3*G3'
            book.save(path)

        shared = BytesIO()
        with ZipFile(path) as source, ZipFile(shared, 'w') as target:
            for member in source.infolist():
                payload = source.read(member)
                if member.filename == 'xl/worksheets/sheet1.xml':
                    xml = ET.fromstring(payload)
                    first = xml.find('.//{*}c[@r="H2"]/{*}f')
                    first.attrib.update(t='shared', si='0', ref='H2:H3')
                    second = xml.find('.//{*}c[@r="H3"]/{*}f')
                    second.attrib.update(t='shared', si='0')
                    second.text = None
                    payload = ET.tostring(xml)
                target.writestr(member, payload)
        path.write_bytes(shared.getvalue())
        before = read_batches(path)
        result = self.generate()
        after = read_batches(path)
        self.assertEqual(after[0]['hdd_cod_no'], result['certificate_no'])
        self.assertEqual(batch_content(after[1]), batch_content(before[1]))
        with closing(load_workbook(path)) as book:
            self.assertEqual(book.active['H2'].value, '=F2*G2')
            self.assertEqual(book.active['H4'].value, '=F4*G4')

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

    def test_refresh_does_not_attach_an_old_hdd_number_to_a_reused_serial(self):
        from run_certificates import run
        self.generate()
        fixtures.write_register(self.paths[0], case_id='CLIENT1', cod='', status='N/A')
        with closing(load_workbook(self.paths[0])) as book:
            sheet, _, cols = select_register(book)
            sheet.cell(2, cols['form6'] + 1, 'UNRELATED-FORM6')
            sheet.cell(2, cols['supplier'] + 1, 'Unrelated inward')
            book.save(self.paths[0])
        report = {}
        run(self.paths, self.root, check_only=True, as_of=AS_OF, progress=None, report_sink=report)
        self.assertEqual(read_batches(self.paths[0])[0]['hdd_cod_no'], '')
        self.assertTrue(any(note.get('kind') == 'hdd_tracking' and 'CLIENT1' in note['detail']
                            for note in report['maintenance_notes']))

    def test_hdd_update_keeps_unrelated_tracking_cells_intact(self):
        from types import SimpleNamespace
        from hdd_workbook import stage_hdd_tracking
        result = self.generate()
        path = self.paths[0]
        for key, value in (('cod_no', 'UNRELATED-COD'), ('cod_date', date(2025, 12, 1))):
            with closing(load_workbook(path)) as book:
                sheet, _, cols = select_register(book)
                sheet.cell(3, cols['cod_no'] + 1, result['certificate_no'])
                sheet.cell(3, cols['cod_date'] + 1, AS_OF)
                sheet.cell(3, cols[key] + 1, value)
                book.save(path)
            before = path.read_bytes()
            with self.assertRaisesRegex(ValueError, 'already contains tracking data'):
                stage_hdd_tracking(SimpleNamespace(data=before), 'CLIENT1', result['certificate_no'], AS_OF)
            self.assertEqual(path.read_bytes(), before)


if __name__ == '__main__': unittest.main()
