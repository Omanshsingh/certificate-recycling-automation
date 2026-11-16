"""Shared certificate output and recoverable migration of existing issued certificates."""
from datetime import date
import json
from pathlib import Path
import tempfile
from unittest import TestCase
from unittest.mock import patch

from pypdf import PdfWriter
import certificate_folder_migration as migration
from certificate_inventory import (certificate_destination, ensure_certificate_folders,
                                   issuance_paths, safe_certificate_path, scan_inventory)
from certificate_drafts import (current_record_pdf, generate_draft, manual_rows,
                                complete_draft, read_register, register_path)
from issuance_safety import file_digest
from workflow_support import save_json
import web_app


class CertificateFolderMigrationTests(TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name); self.base=self.root/'Certificates'
        save_json(self.root/'settings.local.json',{'workbook_paths':['','']})

    def pdf(self,relative,width=300):
        path=self.base/relative; path.parent.mkdir(parents=True,exist_ok=True)
        pdf=PdfWriter(); pdf.add_blank_page(width=width,height=400); pdf.write(path)
        return path

    def issued_legacy(self):
        pdf=self.pdf('Created in App/0012_Client_Mumbai.pdf')
        record={'id':'manual-one','plant':'plant1','status':'completed',
                'details':{'certificate_type':'sita','supplier':'Client','location':'Mumbai','reference':'REF-12',
                           'certificate_weight':'10','inward_date':'2026-09-01','cod_no':'0012','cod_date':'2026-10-01','certificate_fields':{}},
                'pdf_filename':pdf.name,'pdf_relative_path':'Created in App/'+pdf.name}
        save_json(register_path(self.root),{'version':1,'records':[record]})
        receipt={'version':1,'id':'a'*32,'state':'complete','source_kind':'manual','plant':'plant1','key':record['id'],
                 'cod':'0012','pdf_path':pdf.relative_to(self.root).as_posix(),'pdf_sha256':file_digest(pdf),
                 'started_at':'2026-10-01T12:00:00Z'}
        history=self.root/'Data'/'Issuance History'/('a'*32+'.json'); save_json(history,receipt)
        inventory=scan_inventory(self.root)
        inventory['entries'][0]['managed_owner']={'source_kind':'manual','plant':'plant1','key':record['id'],'cod':'0012'}
        save_json(self.root/'Data'/'Certificate Inventory.json',inventory)
        save_json(self.root/'Data'/'Reports'/'Latest Results.json',{'rows':[dict(source_kind='manual',record_id=record['id'],plant='plant1',certificate_status='completed',certificate_no='0012',pdf_filename=pdf.name,pdf_relative_path=record['pdf_relative_path'])]})
        return pdf,history

    def test_new_workspace_uses_only_the_shared_certificate_folder(self):
        self.assertEqual(ensure_certificate_folders(self.root), self.base)
        self.assertEqual(list(self.base.iterdir()), [])
        self.assertEqual(certificate_destination(self.root, '0012_Client.pdf'), self.base / '0012_Client.pdf')

    def test_every_app_certificate_type_is_issued_into_the_shared_folder(self):
        self.pdf('Manually Added/0099_Previous_Client.pdf')
        for kind in ('standard','sita','siemens','destruction'):
            details={'certificate_type':kind,'supplier':'Client '+kind,'location':'Mumbai','reference':'REF-'+kind,
                     'inward_date':'2026-10-01','certificate_weight':'10','weight':'10','form6':'FORM-'+kind,
                     'certificate_fields':{}}
            if kind=='destruction':
                details['certificate_fields']={'destruction_date':'2026-10-01','device_receipts':json.dumps([{'date':'2026-10-01','count':'2'}])}
            response=generate_draft(self.root,'plant1',details,confirmed=True,as_of=date(2026,10,19))
            self.assertEqual(response['created'],1)
            self.assertEqual(Path(response['relative_path']).parent, Path('.'))
            self.assertTrue((self.base / response['relative_path']).is_file())
        self.assertFalse((self.base/'Created in App').exists())

    def test_migration_preserves_pdf_history_status_ownership_and_old_links(self):
        old,history=self.issued_legacy(); pdf_before=old.read_bytes(); history_before=history.read_bytes()
        ensure_certificate_folders(self.root)
        target=self.base/old.name
        self.assertEqual(target.read_bytes(),pdf_before); self.assertEqual(history.read_bytes(),history_before)
        self.assertFalse(old.parent.exists())
        record=read_register(self.root)['records'][0]
        self.assertEqual(record['pdf_relative_path'],old.name)
        self.assertEqual(current_record_pdf(self.root,record),target)
        self.assertEqual(manual_rows(self.root)[0]['certificate_status'],'completed')
        self.assertEqual(safe_certificate_path(self.root,'Created in App/'+old.name),target)
        inventory=scan_inventory(self.root)
        self.assertEqual(issuance_paths(inventory,'12','manual','plant1',record['id']),['Certificates/'+old.name])
        with patch.object(web_app,'ROOT',self.root):
            rows=web_app.current_pdf_rows(manual_rows(self.root))
            self.assertEqual(rows[0]['pdf_relative_path'],old.name)
            output=web_app.certificate_outputs(rows)[0]
            self.assertEqual(output['relative_path'],old.name)
        cached=json.loads((self.root/'Data'/'Reports'/'Latest Results.json').read_text())
        self.assertEqual(cached['rows'][0]['pdf_relative_path'],old.name)
        backups=list((self.root/'Data'/'Certificate Folder Backups').rglob(old.name))
        self.assertEqual(len(backups),1); self.assertEqual(backups[0].read_bytes(),pdf_before)
        before={p:p.read_bytes() for p in self.root.rglob('*') if p.is_file()}
        ensure_certificate_folders(self.root)
        self.assertEqual(before,{p:p.read_bytes() for p in self.root.rglob('*') if p.is_file()})

    def test_same_filename_keeps_both_different_pdfs_and_correct_record_link(self):
        old,history=self.issued_legacy(); payload=old.read_bytes(); history_before=history.read_bytes()
        existing=self.pdf(old.name, width=500); foreign=existing.read_bytes()
        ensure_certificate_folders(self.root)
        moved=self.base/'0012_Client_Mumbai (2).pdf'
        self.assertEqual(existing.read_bytes(),foreign); self.assertEqual(moved.read_bytes(),payload)
        record=read_register(self.root)['records'][0]
        self.assertEqual(record['pdf_filename'],moved.name)
        self.assertEqual(current_record_pdf(self.root,record),moved)
        self.assertEqual(safe_certificate_path(self.root,'Created in App/'+old.name),moved)
        self.assertEqual(history.read_bytes(),history_before)

    def test_migrated_certificate_can_be_corrected_and_completed(self):
        _,history=self.issued_legacy(); old_history=history.read_bytes(); ensure_certificate_folders(self.root)
        response=generate_draft(self.root,record_id='manual-one',confirmed=True,as_of=date(2026,10,19))
        self.assertEqual(response['certificate_no'],'0012')
        self.assertEqual(response['relative_path'],'0012_Client_Mumbai.pdf')
        complete_draft(self.root,'manual-one')
        self.assertEqual(manual_rows(self.root)[0]['certificate_status'],'completed')
        self.assertEqual(history.read_bytes(),old_history)
        self.assertFalse((self.base/'Created in App').exists())

    def test_interrupted_register_update_resumes_without_duplicate_moves(self):
        old,_=self.issued_legacy(); original_save=migration.save_json
        def fail_register(path,value):
            if Path(path)==register_path(self.root): raise OSError('Register temporarily in use')
            return original_save(path,value)
        with patch.object(migration,'save_json',side_effect=fail_register), self.assertRaises(OSError):
            ensure_certificate_folders(self.root)
        self.assertEqual(migration._ledger(self.root)['state'],'moving')
        ensure_certificate_folders(self.root)
        self.assertEqual(migration._ledger(self.root)['state'],'complete')
        self.assertEqual(len(list(self.base.glob('*.pdf'))),1)
        self.assertIsNotNone(current_record_pdf(self.root,read_register(self.root)['records'][0]))

    def test_later_old_folder_files_move_even_after_previous_pdf_is_archived(self):
        old,_=self.issued_legacy(); ensure_certificate_folders(self.root)
        (self.base/old.name).unlink() # Simulate retirement after a later correction.
        new=self.pdf('Created in App/0013_New_Client_hdd.pdf')
        new_before=new.read_bytes(); ensure_certificate_folders(self.root)
        self.assertEqual((self.base/new.name).read_bytes(),new_before)
        self.assertFalse((self.base/'Created in App').exists())

    def test_unreadable_register_does_not_move_or_remove_original_files(self):
        old,_=self.issued_legacy(); before=old.read_bytes(); register_path(self.root).write_text('{broken')
        with self.assertRaises(ValueError): ensure_certificate_folders(self.root)
        self.assertEqual(old.read_bytes(),before)
        self.assertFalse((self.base/old.name).exists())

    def test_pending_old_migration_completes_and_both_old_links_resolve(self):
        old, _ = self.issued_legacy()
        intermediate = f'Auto Generated/{old.name}'
        digest = file_digest(old)
        save_json(migration._ledger_path(self.root), {'version':1, 'state':'moving', 'pending':[0],
            'moves':[{'old':old.relative_to(self.root).as_posix(), 'new':f'Certificates/{intermediate}',
                      'sha256':digest}]})
        ensure_certificate_folders(self.root)
        target = self.base / old.name
        self.assertEqual(file_digest(target), digest)
        self.assertEqual(safe_certificate_path(self.root, f'Created in App/{old.name}'), target)
        self.assertEqual(safe_certificate_path(self.root, intermediate), target)
        self.assertFalse((self.base / 'Created in App').exists())
        self.assertFalse((self.base / 'Auto Generated').exists())
        self.assertEqual(current_record_pdf(self.root, read_register(self.root)['records'][0]), target)

    def test_same_name_in_both_folders_keeps_both_sources(self):
        app_pdf = self.pdf('Auto Generated/0004.pdf')
        external = self.pdf('Manually Added/0004.pdf', width=500)
        app_digest, external_digest = file_digest(app_pdf), file_digest(external)
        ensure_certificate_folders(self.root)
        entries = {entry['sha256']:entry for entry in scan_inventory(self.root)['entries']}
        self.assertEqual(entries[app_digest]['path'], 'Certificates/0004.pdf')
        self.assertTrue(entries[app_digest]['app_generated'])
        self.assertEqual(entries[external_digest]['path'], 'Certificates/0004 (2).pdf')
        self.assertTrue(entries[external_digest]['external'])
        self.assertEqual(entries[external_digest]['current_cods'], ['0004'])
        self.assertEqual(safe_certificate_path(self.root, 'Manually Added/0004.pdf'),
                         self.base / '0004 (2).pdf')

    def test_hdd_collision_keeps_suffix_and_old_backup_paths_still_resolve(self):
        old=self.pdf('Created in App/0013_Client_hdd.pdf'); self.pdf(old.name, width=500)
        ensure_certificate_folders(self.root)
        self.assertTrue((self.base/'0013_Client (2)_hdd.pdf').is_file())
        restored=self.pdf('Created in App/'+old.name)
        self.assertEqual(safe_certificate_path(self.root,'Created in App/'+old.name),restored)
