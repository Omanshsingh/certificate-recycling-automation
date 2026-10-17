"""Destruction is supplemental to an inward; all issuance tests use disposable data."""
from datetime import date
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader
from certificate_drafts import (save_draft, generate_draft, manual_rows, read_register,
                                complete_draft, reconcile_manual, prepare_certificate)
from certificate_templates import make_template_certificate, template_specs, template_id
from destruction_certificate import receipts, request_for_inward, content_issues
from tests.test_certificate_templates import write_register
from workflow_support import save_json
from workspace_review import workspace_revision
import web_app

AS_OF = date(2026, 10, 17)


class DestructionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.paths = [self.root / f'plant{i}.xlsx' for i in (1, 2)]
        for i, path in enumerate(self.paths, 1):
            write_register(path, case_id=f'CLIENT{i}', cod=f'{300+i:04}', status='Completed')
        save_json(self.root / 'settings.local.json', {'workbook_paths': list(map(str, self.paths))})

    def details(self, count=12):
        return {'certificate_type': 'destruction', 'supplier': 'Example International Ltd',
                'location': 'London, UK', 'certificate_fields': {'destruction_date': '2026-10-16',
                'device_receipts': json.dumps([{'date': '2026-10-01', 'count': str(count)}])}}

    def test_request_is_idempotent_and_does_not_mutate_either_inward(self):
        before = [path.read_bytes() for path in self.paths]
        record = request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root))
        self.assertEqual(record, request_for_inward(self.root, 'plant1', 'CLIENT1', workspace_revision(self.root)))
        self.assertEqual(len(read_register(self.root)['records']), 1)
        row = manual_rows(self.root)[0]
        self.assertTrue(row['destruction_pending'])
        self.assertEqual(row['destruction_source'], {'plant': 'plant1', 'case_id': 'CLIENT1'})
        self.assertFalse(row['certificate_no'])
        self.assertEqual(before, [path.read_bytes() for path in self.paths])

    def test_international_client_can_issue_without_an_inward_or_weight(self):
        result = generate_draft(self.root, 'plant2', self.details(), as_of=AS_OF)
        self.assertEqual(result['created'], 1)
        self.assertEqual(result['certificate_no'], '0303')
        path = self.root / 'Certificates' / result['relative_path']
        pdf = PdfReader(path)
        self.assertEqual(len(pdf.pages), 2)
        self.assertIn('EXAMPLE INTERNATIONAL LTD', pdf.pages[0].extract_text())
        self.assertIn('16.10.2026', pdf.pages[0].extract_text())
        self.assertIn('12', pdf.pages[1].extract_text())
        self.assertTrue(manual_rows(self.root)[0]['destruction_pending'])
        complete_draft(self.root, result['record_id'])
        self.assertFalse(manual_rows(self.root)[0]['destruction_pending'])
        path.unlink()
        reconcile_manual(self.root)
        self.assertTrue(manual_rows(self.root)[0]['destruction_pending'])

    def test_linked_issuance_keeps_recycling_cod_status_and_workbook_bytes(self):
        before = [path.read_bytes() for path in self.paths]
        recycling = self.root / 'Certificates' / 'Auto Generated' / '0301_Independent Client Ltd_Pune.pdf'
        make_template_certificate({'certificate_type':'standard', 'supplier':'Independent Client Ltd',
                                  'location':'Pune', 'inward_date':date(2026,9,1),
                                  'display_weight':10, 'form6':'FORM-1', 'certificate_fields':{}},
                                  '0301', AS_OF, recycling)
        recycling_before = recycling.read_bytes()
        record = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = read_register(self.root)['records'][0]['details']
        details['certificate_fields'] = self.details()['certificate_fields']
        result = generate_draft(self.root, 'plant1', details, record, as_of=AS_OF)
        self.assertEqual(result['created'], 1)
        self.assertEqual(result['certificate_no'], '0303')
        self.assertEqual(before, [path.read_bytes() for path in self.paths])
        self.assertEqual(recycling_before, recycling.read_bytes())
        # PDF-only corrections cannot detach the original inward.
        details.pop('destruction_source')
        save_draft(self.root, 'plant1', details, record)
        self.assertEqual(read_register(self.root)['records'][0]['details']['destruction_source']['case_id'], 'CLIENT1')

    def test_required_facts_cannot_be_bypassed_and_preview_is_non_mutating(self):
        details = self.details()
        details['certificate_fields']['destruction_date'] = ''
        with self.assertRaisesRegex(ValueError, 'actual date'):
            generate_draft(self.root, 'plant1', details, confirmed=True, as_of=AS_OF)
        before = workspace_revision(self.root)
        _, _, _, warnings = prepare_certificate(self.root, 'plant1', details, as_of=AS_OF)
        self.assertTrue(any('actual date' in warning for warning in warnings))
        self.assertEqual(before, workspace_revision(self.root))
        for count in ['-1', '0', '1.5', '1 Lot', 'NaN']:
            with self.subTest(count=count), self.assertRaises(ValueError):
                receipts(json.dumps([{'date':'2026-10-01', 'count':count}]))
        details = self.details()
        details['certificate_fields']['destruction_date'] = '2026-09-01'
        self.assertTrue(any('before' in issue for issue in content_issues(details, AS_OF)))

    def test_more_than_nine_receipts_continue_without_dropping_devices(self):
        details = self.details()
        details['certificate_fields']['device_receipts'] = json.dumps([
            {'date': '2026-10-01', 'count': str(i)} for i in range(1, 12)])
        path = self.root / 'continuation.pdf'
        make_template_certificate(details, '0303', AS_OF, path)
        pdf = PdfReader(path)
        self.assertEqual(len(pdf.pages), 3)
        self.assertIn('Total devices: 66', pdf.pages[2].extract_text())
        self.assertIn('Details 2 of 2', pdf.pages[2].extract_text())

    def test_cannot_replace_recycling_or_move_link(self):
        with self.assertRaisesRegex(ValueError, 'additional certificate'):
            web_app.workbook_form({}, self.details(), self.paths)
        record = request_for_inward(self.root, 'plant1', 'CLIENT1')
        details = self.details()
        details['certificate_type'] = 'standard'
        with self.assertRaisesRegex(ValueError, 'own certificate record'):
            save_draft(self.root, 'plant1', details, record)
        with self.assertRaisesRegex(ValueError, 'linked inward'):
            save_draft(self.root, 'plant2', self.details(), record)

    def test_spec_and_two_page_preview(self):
        self.assertEqual(template_id('HDD COD'), 'destruction')
        self.assertIn('destruction', {spec['id'] for spec in template_specs()})
        with patch.object(web_app, 'ROOT', self.root):
            response = web_app.preview_certificate({'source_kind':'manual', 'plant':'plant1', 'details':self.details()})
        self.assertEqual(len(response['preview_detail_images']), 1)
        self.assertFalse(read_register(self.root)['records'])


if __name__ == '__main__':
    unittest.main()
