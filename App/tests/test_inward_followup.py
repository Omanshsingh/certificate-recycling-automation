"""Later inward details must never alter an issued certificate or its reminders."""
from copy import deepcopy
from io import BytesIO
from unittest import TestCase, main
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import load_workbook

from certificate_core import batch_content, read_batches
from inward_followup import FIELDS, clean_followup, stage_followup
from issuance_safety import latest_issuance
from tests import test_review_and_edit as review_fixtures
from workflow_support import Report
from workspace_review import workspace_revision
import web_app


class InwardFollowupTests(TestCase):
    setUp = review_fixtures.ReviewAndEditTests.setUp
    source = review_fixtures.ReviewAndEditTests.source
    body = review_fixtures.ReviewAndEditTests.body
    create = review_fixtures.ReviewAndEditTests.create
    reviewed = review_fixtures.ReviewAndEditTests.reviewed

    def payload(self, **values):
        details = dict(comments='', transporter_details='', docket_vehicle_no='', weight_slip_scale='')
        details.update(values)
        return dict(plant='plant1', sn='1001', expected_revision=workspace_revision(self.root), confirmed_excel=True, details=details)

    def test_partial_save_exposes_only_four_fields_and_keeps_certificate_facts_and_other_plant(self):
        before = deepcopy(self.source())
        other = self.paths[1].read_bytes()
        result = web_app.save_inward_followup(self.payload(comments='Pickup checked', transporter_details='Eco Recycling'))
        self.assertTrue(result['saved'])
        self.assertEqual(result['created'], 0)
        after = self.source()
        self.assertEqual(batch_content(before), batch_content(after))
        for key in ('cod_no', 'cod_date', 'issued'):
            self.assertEqual(after[key], before[key])
        self.assertEqual(after['comments'], 'Pickup checked')
        self.assertEqual(after['docket_vehicle_no'], '')
        self.assertEqual(self.paths[1].read_bytes(), other)
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertTrue(list(web_app.BACKUP_DIR.glob('*.xlsx')))
        row = next(row for row in result['state']['rows'] if row.get('plant') == 'plant1')
        self.assertEqual(row['comments'], 'Pickup checked')

    def test_completed_certificate_pdf_history_and_status_are_unchanged(self):
        self.create(self.reviewed())
        source = self.source()
        from workbook_store import WorkbookSession
        with WorkbookSession(self.paths[0], web_app.BACKUP_DIR) as session:
            session.set_issued_status('1001', source['source_sheet'], source['source_row'], source['cod_no'], 'generated', 'Completed')
        before = self.source()
        receipt = deepcopy(latest_issuance(self.root, 'workbook', 'plant1', '1001'))
        pdf = self.root / receipt['pdf_path']
        original_pdf = pdf.read_bytes()
        result = web_app.save_inward_followup(self.payload(comments='Not applicable', transporter_details='Truck MH 01',
            docket_vehicle_no='D-123', weight_slip_scale='SLIP-456'))
        after = self.source()
        self.assertEqual(result['created'], 0)
        self.assertEqual(after['issued'], 'Completed')
        self.assertEqual(after['cod_no'], before['cod_no'])
        self.assertEqual(after['cod_date'], before['cod_date'])
        self.assertEqual(pdf.read_bytes(), original_pdf)
        self.assertEqual(latest_issuance(self.root, 'workbook', 'plant1', '1001'), receipt)
        row = next(row for row in result['state']['rows'] if row.get('plant') == 'plant1')
        self.assertEqual(row['certificate_status'], 'completed')
        self.assertFalse(row['reminder_due'])

    def test_cache_and_latest_report_update_without_losing_pdf_or_warning_state(self):
        report = Report(self.root, __import__('datetime').date.today(), 'Check', publish_latest=False)
        report.add(dict(self.source(), source_id='plant1'), 'ready', 'Existing report')
        state = {'rows':report.rows, 'scan_complete':True}
        state['rows'][0].update(pdf_filename='existing.pdf', has_issued_certificate=True)
        payload = self.payload(comments='New comment')
        with patch.object(web_app, 'WORKSPACE_STATE_CACHE', (payload['expected_revision'], state)):
            result = web_app.save_inward_followup(payload)
        self.assertEqual(result['state']['rows'][0]['comments'], 'New comment')
        self.assertEqual(result['state']['rows'][0]['pdf_filename'], 'existing.pdf')
        self.assertTrue(result['state']['rows'][0]['has_issued_certificate'])
        self.assertEqual(state['rows'][0]['comments'], '')

    def test_missing_consent_stale_revision_bad_fields_and_failed_save_preserve_original(self):
        original = self.paths[0].read_bytes()
        for change in ({'confirmed_excel':False}, {'expected_revision':'outdated'},
                       {'details':dict(self.payload()['details'], supplier='Not allowed')}):
            with self.assertRaises(ValueError):
                web_app.save_inward_followup(dict(self.payload(), **change))
            self.assertEqual(self.paths[0].read_bytes(), original)
        with patch.object(web_app.WorkbookSession, '_persist', side_effect=OSError('Excel is busy')):
            with self.assertRaisesRegex(OSError, 'Excel is busy'):
                web_app.save_inward_followup(self.payload(comments='Not saved'))
        self.assertEqual(self.paths[0].read_bytes(), original)

    def test_native_package_and_existing_cell_styles_are_preserved(self):
        data, _ = stage_followup(self.paths[0].read_bytes(), '1001', self.payload()['details'])
        values = dict(self.payload()['details'], comments='Changed', docket_vehicle_no='Not applicable')
        updated, _ = stage_followup(data, '1001', values)
        with ZipFile(BytesIO(data)) as before, ZipFile(BytesIO(updated)) as after:
            changed = [name for name in before.namelist() if before.read(name) != after.read(name)]
            self.assertEqual(changed, ['xl/worksheets/sheet1.xml'])
        old_book, new_book = load_workbook(BytesIO(data)), load_workbook(BytesIO(updated))
        try:
            for old_row in old_book.active:
                for cell in old_row:
                    self.assertEqual(cell.style_id, new_book.active[cell.coordinate].style_id)
            self.assertEqual(list(old_book.active.merged_cells.ranges), list(new_book.active.merged_cells.ranges))
        finally:
            old_book.close(); new_book.close()

    def test_clearing_details_returns_them_to_followup_and_formula_or_locked_cells_are_not_overwritten(self):
        values = {key:'Not applicable' for key in FIELDS}
        data, saved = stage_followup(self.paths[0].read_bytes(), '1001', values)
        empty, cleared = stage_followup(data, '1001', {key:'' for key in FIELDS})
        self.assertTrue(all(cleared[key] == '' for key in FIELDS))
        for formula, protected in ((True, False), (False, True)):
            book = load_workbook(BytesIO(data))
            cell = book.active[f"{saved['column_mapping']['comments']['column']}{saved['source_row']}"]
            cell.value = '=1+1' if formula else 'Existing'
            book.active.protection.sheet = protected
            output = BytesIO(); book.save(output); book.close()
            with self.assertRaisesRegex(ValueError, 'formula|locked'):
                stage_followup(output.getvalue(), '1001', dict(values, comments='Changed'))

    def test_field_validation_rejects_unsupported_characters_and_overlong_text(self):
        for value in ('bad\x01text', 'x' * 1001):
            with self.assertRaises(ValueError):
                clean_followup(dict(self.payload()['details'], comments=value))
        data, saved = stage_followup(self.paths[0].read_bytes(), '1001', dict(self.payload()['details'], comments='=1+1'))
        book = load_workbook(BytesIO(data))
        try:
            cell = book.active[f"{saved['column_mapping']['comments']['column']}{saved['source_row']}"]
            self.assertEqual(cell.data_type, 's')
            self.assertEqual(cell.value, '=1+1')
        finally:
            book.close()

    def test_inline_unissued_save_keeps_all_four_details_and_core_changes_together(self):
        payload = self.reviewed(supplier='Updated name', comments='Pickup checked',
            transporter_details='Eco Recycling', docket_vehicle_no='DC-20', weight_slip_scale='WS-20')
        payload['confirmed_excel'] = True
        result = web_app.save_inward_changes(payload)
        self.assertEqual(result['created'], 0)
        self.assertEqual(self.source()['supplier'], 'Updated name')
        self.assertEqual(self.source()['comments'], 'Pickup checked')
        self.assertEqual(self.source()['weight_slip_scale'], 'WS-20')
        self.assertFalse(list((self.root/'Certificates').rglob('*.pdf')))

    def test_pdf_only_keeps_inward_fields_and_yes_saves_them_with_corrected_pdf(self):
        payload = self.reviewed(comments='Pickup checked', transporter_details='Eco Recycling',
            docket_vehicle_no='DC-20', weight_slip_scale='WS-20')
        self.assertEqual(self.create(payload)['created'], 1)
        self.assertEqual(self.source()['comments'], '')
        payload = self.reviewed(comments='Pickup checked', transporter_details='Eco Recycling',
            docket_vehicle_no='DC-20', weight_slip_scale='WS-20')
        payload.update(confirmed_excel=True, edit_scope='update_excel')
        self.assertEqual(self.create(payload)['created'], 1)
        self.assertEqual(self.source()['comments'], 'Pickup checked')
        self.assertEqual(self.source()['weight_slip_scale'], 'WS-20')


if __name__ == '__main__':
    main()
