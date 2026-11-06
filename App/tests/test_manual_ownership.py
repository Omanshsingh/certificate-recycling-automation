"""App-only certificates own their global COD without claiming an Excel plant."""
import json
import unittest
from datetime import date

from certificate_drafts import (save_draft, read_register, register_path, generate_draft,
                                current_record_pdf, migrate_local_ownership, prepare_certificate)
from certificate_inventory import scan_inventory
from issuance_safety import latest_issuance, audit_records
from workspace_review import workspace_revision
from entry_deletion import delete_entries
from certificate_analytics import build_analytics
from tests import test_certificate_inventory as fixtures
AS_OF = fixtures.AS_OF


class ManualOwnershipTests(unittest.TestCase):
    def setUp(self):
        fixture = fixtures.InventoryWorkflowTests()
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        self.root, self.paths = fixture.root, fixture.paths
    details = fixtures.InventoryWorkflowTests.details

    def test_app_only_save_preview_and_issue_have_no_plant_even_from_a_stale_client(self):
        identifier = save_draft(self.root, 'plant2', self.details())
        self.assertIsNone(json.loads(register_path(self.root).read_text())['records'][0]['plant'])
        batch, _, _, _ = prepare_certificate(self.root, None, self.details(), record_id=identifier, as_of=AS_OF)
        self.assertIsNone(batch['source_id'])
        result = generate_draft(self.root, None, self.details(), identifier, as_of=AS_OF)
        self.assertEqual(result['created'], 1)
        record = read_register(self.root)['records'][0]
        self.assertIsNone(record['plant'])
        self.assertIsNotNone(current_record_pdf(self.root, record))
        self.assertIsNone(latest_issuance(self.root, 'manual', None, identifier)['plant'])
        self.assertIsNone(next(e for e in scan_inventory(self.root)['entries'] if e.get('managed_owner'))['managed_owner']['plant'])

    def test_legacy_migration_preserves_issued_pdf_cod_and_history_for_correction(self):
        issued = generate_draft(self.root, 'plant1', self.details(), as_of=AS_OF)
        path = register_path(self.root);data = json.loads(path.read_text());data['records'][0]['plant']='plant2'
        path.write_text(json.dumps(data))
        journal = next((self.root/'Data/Issuance History').glob('*.json'))
        old = json.loads(journal.read_text());old['plant']='plant2';journal.write_text(json.dumps(old))
        before_history=journal.read_bytes();record=read_register(self.root)['records'][0]
        pdf=current_record_pdf(self.root,record);before_pdf=pdf.read_bytes()
        migrate_local_ownership(self.root)
        self.assertIsNone(json.loads(path.read_text())['records'][0]['plant'])
        self.assertEqual(pdf.read_bytes(),before_pdf)
        self.assertEqual(journal.read_bytes(),before_history)
        self.assertTrue(list((self.root/'Data/Custom Register Backups').glob('*.json')))
        corrected=generate_draft(self.root,None,dict(record['details'],supplier='Corrected local company'),record['id'],as_of=AS_OF)
        self.assertEqual(corrected['certificate_no'],issued['certificate_no'])
        self.assertEqual(journal.read_bytes(),before_history)

    def test_unissued_local_record_can_be_deleted_without_a_plant(self):
        identifier=save_draft(self.root,None,self.details())
        body=dict(selected=[dict(source_kind='manual',plant=None,record_id=identifier)],include_linked_hdd=False,
                  expected_revision=workspace_revision(self.root),request_id='f'*32)
        preview=delete_entries(self.root,body)
        self.assertIsNone(preview['review']['entries'][0]['plant'])
        delete_entries(self.root,dict(body,checked=True,review_token=preview['review_token']),commit=True)
        self.assertEqual(read_register(self.root)['records'],[])

    def test_local_certificates_count_in_totals_but_not_either_plant(self):
        row=dict(source_kind='manual',plant=None,certificate_no='0042',certificate_date='2026-10-15',certificate_status='generated_manual')
        all_data=build_analytics([row],[],today=date(2026,10,15))
        self.assertEqual(all_data['totals']['generated'],1)
        self.assertEqual(all_data['coverage']['no_plant'],1)
        for plant in ('plant1','plant2'):
            self.assertEqual(build_analytics([row],[],plant=plant,today=date(2026,10,15))['totals']['generated'],0)

    def test_legacy_local_serials_do_not_require_an_invented_plant_namespace(self):
        first=save_draft(self.root,None,self.details(form6=''))
        second=save_draft(self.root,None,self.details(form6='',supplier='Another local company'))
        self.assertNotEqual(first,second)
        self.assertTrue(all(record['plant'] is None for record in read_register(self.root)['records']))
