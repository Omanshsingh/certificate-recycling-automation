"""References are required outside HDD and can be reused without a notice."""
from datetime import date
import json
import shutil
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from certificate_core import read_batches
from certificate_drafts import generate_draft, read_register, save_draft, manual_rows
from destruction_certificate import request_for_inward
from identifier_rules import shared_reference_warnings
from run_certificates import run
from tests.test_certificate_templates import write_register
from workflow_support import save_json
from workspace_review import workspace_revision
import web_app


class ReferencePolicyTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory()))
        self.paths = [self.root / f'Plant {i}.xlsx' for i in (1, 2)]
        for index, path in enumerate(self.paths, 1):
            write_register(path, case_id=f'CLIENT{index}', cod='0302' if index == 2 else '',
                           status='Completed' if index == 2 else 'Ready')
            self.set_reference(index, f'PLANT-{index}-REF')
        save_json(self.root / 'settings.local.json', {'workbook_paths':list(map(str, self.paths))})
        shutil.copytree(Path(web_app.__file__).parent / 'templates', self.root / 'App' / 'templates')

    def set_reference(self, plant, value):
        book = load_workbook(self.paths[plant - 1])
        book.active.cell(2, 10).value = value
        book.save(self.paths[plant - 1]); book.close()

    def details(self, kind='standard', reference='MANUAL REF / 0007'):
        return {'certificate_type':kind, 'supplier':'Reference Policy Client', 'location':'Mumbai',
                'inward_date':'2026-10-01', 'certificate_weight':'10', 'form6':f'MANUAL-F6-{kind}',
                'reference':reference, 'certificate_fields':
                {'destruction_date':'2026-10-16', 'device_receipts':json.dumps([{'date':'2026-10-01','count':'12'}])}
                if kind == 'destruction' else {}}

    def test_blank_reference_blocks_other_manual_certificates_even_after_confirmation(self):
        for kind in ('standard','sita','siemens'):
            for value in ('', '   '):
                for confirmed in (False, True):
                    with self.subTest(kind=kind, value=value, confirmed=confirmed):
                        with self.assertRaisesRegex(ValueError, 'reference number before creating'):
                            generate_draft(self.root, 'plant1', self.details(kind, value),
                                           confirmed=confirmed, as_of=date(2026,10,19))
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertFalse(read_register(self.root)['records'])
        # A draft can still be saved while the operator is gathering details.
        self.assertTrue(save_draft(self.root, 'plant1', self.details(reference='')))

    def test_missing_excel_reference_blocks_single_and_bulk_runs_even_when_warnings_accepted(self):
        self.set_reference(1, '')
        for selected in (None, ('plant1','CLIENT1')):
            report = {}
            result = run(self.paths, self.root, selected=selected, plants=['plant1'] if selected is None else None,
                         allow_warnings=True, report_sink=report, progress=None)
            self.assertEqual(result[0], 0)
            row = next(row for row in report['rows'] if row.get('sn') == 'CLIENT1')
            self.assertIn('reference number', row['detail'])
        self.assertFalse(list((self.root / 'Certificates').rglob('*.pdf')))
        self.assertEqual(read_batches(self.paths[0])[0]['issued'], 'Ready')

    def test_entering_missing_reference_at_creation_saves_it_without_changing_other_inward_facts(self):
        self.set_reference(1, '')
        original = read_batches(self.paths[0])[0]
        other = self.paths[1].read_bytes()
        details = web_app.editor_details(original)
        details['reference'] = 'CLIENT REF / 0007'
        body = {'source_kind':'workbook','plant':'plant1','source_plant':'plant1','source_sn':'CLIENT1',
                'edit_scope':'certificate_only','details':details, 'expected_revision':workspace_revision(self.root)}
        with patch.object(web_app,'ROOT',self.root), patch.object(web_app,'BACKUP_DIR',self.root/'Backups'), \
                patch.object(web_app,'RESULTS_PATH',self.root/'Results.json'), \
                patch.object(web_app,'refresh',side_effect=lambda *_:web_app.workspace_state(False)):
            preview = web_app.preview_certificate(body)
            self.assertTrue(preview['preview_url'])
            self.assertEqual(read_batches(self.paths[0])[0]['reference'], '')
            result = web_app.create_certificate(body)
        self.assertEqual(result['created'], 1)
        saved = read_batches(self.paths[0])[0]
        self.assertEqual(saved['reference'], 'CLIENT REF / 0007')
        for key in ('supplier','location','inward_date','form6','items'):
            self.assertEqual(saved[key], original[key], key)
        self.assertEqual(self.paths[1].read_bytes(), other)

    def test_shared_reference_creates_without_confirmation(self):
        for kind in ('standard',):
            details = self.details(kind, 'PLANT-2-REF')
            result = generate_draft(self.root,'plant1',details,as_of=date(2026,10,19))
            self.assertFalse(result['requires_confirmation'])
            self.assertEqual(result['created'],1)
            record = next(row for row in read_register(self.root)['records'] if row['id'] == result['record_id'])
            self.assertEqual(record['details']['reference'],'PLANT-2-REF')

    def test_standalone_and_linked_hdd_preview_issue_and_reissue_without_a_reference(self):
        self.set_reference(1, '')
        before = [path.read_bytes() for path in self.paths]
        original = read_batches(self.paths[0])[0]
        for linked in (False, True):
            with self.subTest(linked=linked):
                record_id = None
                details = self.details('destruction', '')
                if linked:
                    record_id = request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root))
                    saved = next(record for record in read_register(self.root)['records'] if record['id'] == record_id)
                    self.assertEqual(saved['details']['reference'], '')
                    details = dict(saved['details'], certificate_fields=details['certificate_fields'])
                body = {'plant':'plant1','source_kind':'manual','record_id':record_id,'details':details}
                with patch.object(web_app, 'ROOT', self.root):
                    self.assertTrue(web_app.preview_certificate(body)['preview_url'])
                result = generate_draft(self.root, 'plant1', details, record_id, as_of=date(2026,10,19),
                                        expected_revision=workspace_revision(self.root))
                self.assertEqual(result['created'], 1)
                corrected = generate_draft(self.root, 'plant1', details, result['record_id'], as_of=date(2026,10,19),
                                           expected_revision=workspace_revision(self.root))
                self.assertEqual(corrected['certificate_no'], result['certificate_no'])
                if not linked:
                    self.assertEqual(before, [path.read_bytes() for path in self.paths])
        linked = read_batches(self.paths[0])[0]
        self.assertEqual(linked['hdd_cod_no'], result['certificate_no'])
        for key in ('cod_no','cod_date','reference','supplier','inward_date','items','issued'):
            self.assertEqual(linked[key], original[key])
        self.assertEqual(before[1], self.paths[1].read_bytes())

    def test_legacy_hdd_references_do_not_warn_or_claim_another_certificates_reference(self):
        details = self.details('destruction', 'PLANT-2-REF')
        result = generate_draft(self.root, 'plant1', details, as_of=date(2026,10,19))
        self.assertEqual(result['created'], 1)
        self.assertFalse(any('Reference ' in warning for warning in result['warnings']))
        records = read_register(self.root)
        record = next(record for record in records['records'] if record['id'] == result['record_id'])
        record['warnings'] = ['Reference PLANT-2-REF is also used by Plant 2. You can keep the same reference and continue.',
                              'Check the destruction date.']
        save_json(self.root / 'Data' / 'Custom Certificates.json', records)
        row = next(row for row in manual_rows(self.root) if row['record_id'] == record['id'])
        self.assertEqual(row['warnings'], ['Check the destruction date.'])
        self.assertEqual(read_register(self.root)['records'][0]['details']['reference'], 'PLANT-2-REF')
        self.assertEqual(shared_reference_warnings({'reference':'PLANT-2-REF'}, 'plant1', 'manual', None, [], records['records']), [])

    def test_own_reference_is_excluded_and_matching_ignores_case_and_outer_spaces(self):
        batches=[dict(read_batches(path)[0],source_id=f'plant{i}') for i,path in enumerate(self.paths,1)]
        self.assertEqual(shared_reference_warnings(batches[0],'plant1','workbook','CLIENT1',batches),[])
        self.assertEqual(shared_reference_warnings({'reference':' plant-2-ref '},'plant1','manual',None,batches), [])


if __name__ == '__main__':
    unittest.main()
