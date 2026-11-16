"""Backup/restore tests: disposable workspaces only, no live Excel or project data."""
from contextlib import closing
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import backup_service as backup
from workbook_store import exclusive_file
from workflow_support import run_lock


class BackupServiceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        self.root = self.base / "workspace"
        self.root.mkdir()
        self.books = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsm"]
        for index, book in enumerate(self.books):
            # ZIP bytes, including a VBA member, must survive without re-saving Excel.
            with ZipFile(book, "w") as archive:
                archive.writestr("xl/workbook.xml", f"plant {index}")
                if book.suffix == ".xlsm":
                    archive.writestr("xl/vbaProject.bin", b"macro bytes")
        self.settings = {"workbook_paths": [str(p) for p in self.books]}
        self.write_settings()
        self.put("Certificates/0001_client_1.pdf", b"issued PDF")
        self.put("Data/Certificate Inventory.json", b'{"entries":[]}')
        self.put("Data/Custom Certificates.json", b'{"records":[]}')
        self.put("Data/Audit/events.jsonl", b'{"event":"issued"}\n')
        self.put("Data/Issuance Journal.json", b'{"committed":true}')
        self.put("Data/Issuance History/complete.json", b'{"version":1,"state":"complete"}')
        self.put("Data/Issuance History/rolled-back.json", b'{"version":1,"state":"rolled_back"}')
        self.put("Data/Issuance Recovery/completed/new.pdf", b"new issued PDF")
        self.put("Data/Issuance Recovery/completed/previous-0.pdf", b"previous issued PDF")
        self.put("Data/Issuance Recovery/completed/source-before.bin", b"original workbook bytes")
        self.put("Data/Working Drafts/a.json", b'{"draft":true}')
        self.put("Data/Attachments/a.pdf", b"annexure")
        self.put("Data/Attachments/a.json", b'{"pages":1}')
        self.put("Data/Archive/Superseded Certificates/old.pdf", b"historic COD")
        self.put("Output/Certificates/legacy.pdf", b"legacy")
        self.put("App/templates/certificate-types/standard.pdf", b"template PDF")
        self.put("App/templates/certificate-types/standard.pptx", b"original template")
        self.put("App/templates/certificate-types/standard.geometry.json", b"{}")
        for excluded in ("Data/Previews/example.pdf", "Data/Reports/results.json",
                         "Data/Workbook Backups/old.xlsx", "Data/Custom Register Backups/old.json",
                         "Data/Archive/old.xlsx", "Data/run.lock", "Data/random.tmp",
                         "App/templates/no-execution.py", "Output/other.pdf"):
            self.put(excluded, b"excluded")

    def put(self, relative, data):
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def write_settings(self):
        (self.root / "settings.local.json").write_text(json.dumps(self.settings), encoding="utf-8")

    def create(self, destination=None):
        return backup.create_backup(self.root, destination)

    def restore(self, item):
        return backup.restore_backup(self.root, item["backup_id"], confirmed=True)

    def test_backup_restores_after_a_completed_deletion_undo(self):
        identifier = 'a' * 32
        data = json.dumps({'id': identifier, 'version': 1, 'state': 'undone'}).encode()
        journal = self.put(f'Data/Entry Deletions/{identifier}.json', data)
        saved = self.create()
        self.put('Data/Working Drafts/later.json', b'{"draft":true}')
        restored = self.restore(saved)
        self.assertEqual(restored['backup_id'], saved['backup_id'])
        self.assertEqual(journal.read_bytes(), data)

    def manifest(self, item):
        with ZipFile(item["path"]) as archive:
            return json.loads(archive.read("manifest.json"))

    def rewrite_archive(self, item, mutate, *, sign=True, update_digest=True):
        """Simulate an imported/malformed signed artifact to exercise path validation."""
        path = Path(item["path"])
        with ZipFile(path) as archive:
            contents = {info.filename: archive.read(info) for info in archive.infolist()}
        manifest = json.loads(contents["manifest.json"])
        mutate(manifest, contents)
        contents["manifest.json"] = backup._json_bytes(manifest)
        key = backup._key(self.root)
        if sign:
            contents["manifest.hmac"] = backup._sign(manifest, key).encode("ascii")
        with ZipFile(path, "w") as archive:
            for name, data in contents.items():
                archive.writestr(name, data)
        if update_digest:
            catalog = backup._catalog(self.root, key)
            next(r for r in catalog["backups"] if r["backup_id"] == item["backup_id"])["sha256"] = backup._hash(path)
            backup._save_catalog(self.root, catalog, key)

    def test_includes_durable_data_templates_archived_pdfs_and_both_workbook_formats(self):
        item = self.create()
        manifest = self.manifest(item)
        paths = {f.get("relative_path") for f in manifest["files"]}
        self.assertEqual(manifest["original_workbook_paths"], [str(p) for p in self.books])
        self.assertEqual({f["workbook_index"] for f in manifest["files"] if f["kind"] == "workbook"}, {0, 1})
        for included in ("settings.local.json", "Certificates/0001_client_1.pdf",
                         "Data/Certificate Inventory.json", "Data/Custom Certificates.json",
                         "Data/Audit/events.jsonl", "Data/Issuance Journal.json",
                         "Data/Issuance History/complete.json", "Data/Issuance History/rolled-back.json",
                         "Data/Issuance Recovery/completed/new.pdf",
                         "Data/Issuance Recovery/completed/previous-0.pdf",
                         "Data/Issuance Recovery/completed/source-before.bin",
                         "Data/Working Drafts/a.json", "Data/Attachments/a.pdf",
                         "Data/Archive/Superseded Certificates/old.pdf",
                         "Output/Certificates/legacy.pdf",
                         "App/templates/certificate-types/standard.pptx"):
            self.assertIn(included, paths)
        self.assertNotIn("Data/Previews/example.pdf", paths)
        self.assertNotIn("Data/Archive/old.xlsx", paths)
        self.assertFalse(any(p and ("Backups" in p or "Reports" in p) for p in paths))
        with ZipFile(item["path"]) as archive:
            for entry in manifest["files"]:
                data = archive.read(entry["member"])
                self.assertEqual(entry["sha256"], hashlib.sha256(data).hexdigest())
                self.assertEqual(entry["size"], len(data))
        self.assertTrue(backup.list_backups(self.root)[0]["valid"])

    def test_list_is_read_only_when_no_backups_exist(self):
        before = sorted(p.relative_to(self.root).as_posix() for p in self.root.rglob("*"))
        self.assertEqual(backup.list_backups(self.root), [])
        self.assertEqual(before, sorted(p.relative_to(self.root).as_posix() for p in self.root.rglob("*")))

    def test_persisted_destination_and_explicit_override_are_registered(self):
        chosen = self.base / "External Disk" / "Backups"
        self.settings["backup_destination"] = str(chosen)
        self.write_settings()
        first = self.create()
        second = self.create(self.base / "other")
        self.assertEqual(Path(first["path"]).parent, chosen)
        self.assertEqual(Path(second["path"]).parent, self.base / "other")
        self.assertEqual([r["backup_id"] for r in backup.list_backups(self.root)],
                         [second["backup_id"], first["backup_id"]])
        self.assertEqual(json.loads((self.root / "settings.local.json").read_text())["backup_destination"], str(chosen))

    def test_destination_cannot_be_inside_included_paths_or_contain_a_workbook(self):
        for destination in (self.root, self.root / "Certificates" / "Backups",
                            self.root / "Data" / "Attachments" / "Backups",
                            self.root / "App" / "random", self.base):
            with self.subTest(destination=destination), self.assertRaises(backup.BackupError):
                self.create(destination)
        self.assertFalse((self.root / "Data" / "Backups").exists())

    def test_unsupported_or_missing_workbook_aborts_without_zip(self):
        self.books[1].unlink()
        with self.assertRaisesRegex(backup.BackupError, "missing"):
            self.create()
        self.settings["workbook_paths"][1] = str(self.root / "other.xls")
        self.write_settings()
        with self.assertRaisesRegex(backup.BackupError, "XLSX and XLSM"):
            self.create()
        self.assertFalse((self.root / "Data" / "Backups").exists())

    def test_confirmation_requires_boolean_true(self):
        item = self.create()
        for confirmed in (False, None, "true", 1):
            with self.subTest(confirmed=confirmed), self.assertRaisesRegex(backup.BackupError, "confirmation"):
                backup.restore_backup(self.root, item["backup_id"], confirmed=confirmed)

    def test_id_never_accepts_a_path_or_unknown_id(self):
        for value in ("../file.zip", "C:\\backups\\file.zip", "backup-" + "a" * 32, None):
            with self.subTest(value=value), self.assertRaises(backup.BackupError):
                backup.restore_backup(self.root, value, confirmed=True)

    def test_restore_preserves_pre_restore_snapshot_and_archives_only_issuance_extras(self):
        original = [p.read_bytes() for p in self.books]
        item = self.create(self.base / "separate")
        for book in self.books:
            book.write_bytes(b"new workbook")
        self.put("Certificates/0001_client_1.pdf", b"new PDF")
        extra = self.put("Certificates/0002_client_2.pdf", b"retain newer PDF")
        unrelated = self.put("Data/unrelated-note.txt", b"unrelated")
        preview = self.put("Data/Previews/new.pdf", b"preview")
        attachment = self.put("Data/Attachments/new.pdf", b"reusable attachment")
        result = self.restore(item)
        self.assertEqual([p.read_bytes() for p in self.books], original)
        self.assertEqual((self.root / "Certificates/0001_client_1.pdf").read_bytes(), b"issued PDF")
        self.assertFalse(extra.exists())
        self.assertEqual((Path(result["extras_path"]) / "Certificates/0002_client_2.pdf").read_bytes(), b"retain newer PDF")
        for path, data in ((unrelated, b"unrelated"), (preview, b"preview"), (attachment, b"reusable attachment")):
            self.assertEqual(path.read_bytes(), data)
        self.assertTrue(Path(result["pre_restore_path"]).is_file())
        previous = self.manifest({"path": result["pre_restore_path"]})
        with ZipFile(result["pre_restore_path"]) as archive:
            entry = next(f for f in previous["files"] if f["kind"] == "workbook" and f["workbook_index"] == 0)
            self.assertEqual(archive.read(entry["member"]), b"new workbook")
        journal = json.loads((Path(result["recovery_path"]) / "journal.json").read_text())
        self.assertEqual(journal["state"], "committed")

    def test_missing_file_can_be_recovered_if_original_mapping_is_still_configured(self):
        item = self.create()
        original = self.books[0].read_bytes()
        self.books[0].unlink()
        (self.root / "Data/Attachments/a.pdf").unlink()
        self.restore(item)
        self.assertEqual(self.books[0].read_bytes(), original)
        self.assertEqual((self.root / "Data/Attachments/a.pdf").read_bytes(), b"annexure")

    def test_reconfigured_or_missing_original_workbook_mapping_is_rejected(self):
        item = self.create()
        unrelated = self.base / "another project.xlsx"
        unrelated.write_bytes(b"do not touch")
        for replacement in (str(unrelated), ""):
            self.settings["workbook_paths"][0] = replacement
            self.write_settings()
            with self.assertRaisesRegex(backup.BackupError, "mapping"):
                self.restore(item)
            self.assertEqual(unrelated.read_bytes(), b"do not touch")
        self.assertFalse((self.root / "Data/Backups/Restore Transactions").exists())

    def test_exact_configured_external_workbook_is_allowed_but_relocation_is_not(self):
        external = self.base / "approved external.xlsm"
        shutil.copy2(self.books[1], external)
        self.settings["workbook_paths"][1] = str(external)
        self.write_settings()
        item = self.create()
        original = external.read_bytes()
        external.write_bytes(b"changed")
        self.restore(item)
        self.assertEqual(external.read_bytes(), original)

    def test_another_workspace_cannot_restore_even_with_copied_trust_store(self):
        item = self.create()
        another = self.base / "different workspace"
        shutil.copytree(self.root, another)
        copied = backup.list_backups(another)
        self.assertFalse(copied[0]["valid"])
        with self.assertRaisesRegex(backup.BackupError, "different workspace"):
            backup.restore_backup(another, item["backup_id"], confirmed=True)

    def test_unregistered_archives_are_not_listed_or_imported(self):
        item = self.create()
        copy_id = "backup-" + "f" * 32
        shutil.copy2(item["path"], Path(item["path"]).with_name(copy_id + ".zip"))
        self.assertEqual(len(backup.list_backups(self.root)), 1)
        with self.assertRaisesRegex(backup.BackupError, "Unknown"):
            backup.restore_backup(self.root, copy_id, confirmed=True)

    def test_tampered_zip_and_missing_zip_are_reported_and_refused(self):
        item = self.create()
        path = Path(item["path"])
        path.write_bytes(path.read_bytes() + b"tampered")
        self.assertFalse(backup.list_backups(self.root)[0]["valid"])
        with self.assertRaisesRegex(backup.BackupError, "SHA256"):
            self.restore(item)
        path.unlink()
        self.assertIn("missing", backup.list_backups(self.root)[0]["error"])

    def test_tampered_catalog_and_missing_trust_key_fail_closed(self):
        self.create()
        catalog = self.root / "Data/Backups/catalog.json"
        original = catalog.read_bytes()
        value = json.loads(original)
        value["catalog"]["backups"][0]["directory"] = str(self.base / "untrusted")
        catalog.write_text(json.dumps(value))
        with self.assertRaises(backup.BackupError):
            backup.list_backups(self.root)
        catalog.write_bytes(original)
        (self.root / "Data/Backups/trust.key").unlink()
        with self.assertRaisesRegex(backup.BackupError, "trust key"):
            backup.list_backups(self.root)
        with self.assertRaisesRegex(backup.BackupError, "trust key"):
            self.create()

    def test_manifest_signature_failure_is_rejected_even_if_catalog_digest_matches(self):
        item = self.create()
        self.rewrite_archive(item, lambda manifest, _: manifest.update(purpose="forged"), sign=False)
        with self.assertRaisesRegex(backup.BackupError, "authentication"):
            self.restore(item)

    def test_payload_checksum_failure_is_detected_before_overwrites(self):
        item = self.create()
        def damage(manifest, contents):
            contents[manifest["files"][0]["member"]] = b"X" * manifest["files"][0]["size"]
        self.rewrite_archive(item, damage)
        original = self.books[0].read_bytes()
        with self.assertRaisesRegex(backup.BackupError, "payload SHA256"):
            self.restore(item)
        self.assertEqual(self.books[0].read_bytes(), original)

    def test_traversal_absolute_ads_reserved_and_unapproved_targets_are_rejected(self):
        for relative in ("../other.xlsx", "/absolute.xlsx", "Data/../../other.xlsx",
                         "Data/file.json:stream", "Data/CON.json", "Data/name. ",
                         "App/web_app.py", "Data/Backups/trust.key", "Data/run.lock"):
            with self.subTest(relative=relative):
                item = self.create()
                def change(manifest, _, relative=relative):
                    entry = next(f for f in manifest["files"] if f["kind"] == "workspace")
                    entry["relative_path"] = relative
                self.rewrite_archive(item, change)
                with self.assertRaises(backup.BackupError):
                    self.restore(item)

    def test_extra_traversal_member_and_case_duplicate_targets_are_rejected(self):
        item = self.create()
        self.rewrite_archive(item, lambda _, contents: contents.update({"../outside": b"bad"}))
        with self.assertRaises(backup.BackupError):
            self.restore(item)
        item = self.create()
        def duplicate(manifest, _):
            entries = [f for f in manifest["files"] if f["kind"] == "workspace"]
            entries[1]["relative_path"] = entries[0]["relative_path"].upper()
        self.rewrite_archive(item, duplicate)
        with self.assertRaises(backup.BackupError):
            self.restore(item)

    def test_archive_expansion_limits_abort_without_writes(self):
        item = self.create()
        with patch.object(backup, "MAX_TOTAL_BYTES", 1), self.assertRaises(backup.BackupError):
            self.restore(item)
        self.assertFalse((self.root / "Data/Backups/Restore Transactions").exists())

    def test_restored_settings_cannot_redirect_workbook_mapping(self):
        item = self.create()
        unrelated = self.base / "other project.xlsx"
        unrelated.write_bytes(b"another project")
        def redirect(manifest, contents):
            entry = next(e for e in manifest["files"] if e.get("relative_path") == "settings.local.json")
            data = backup._json_bytes({"workbook_paths": [str(unrelated), str(self.books[1])]})
            contents[entry["member"]] = data
            entry.update(size=len(data), sha256=hashlib.sha256(data).hexdigest())
        self.rewrite_archive(item, redirect)
        before = self.books[0].read_bytes()
        with self.assertRaisesRegex(backup.BackupError, "Saved settings"):
            self.restore(item)
        self.assertEqual(self.books[0].read_bytes(), before)
        self.assertEqual(unrelated.read_bytes(), b"another project")

    def pending_issuance(self):
        identifier = "a" * 32
        recovery = f"Data/Issuance Recovery/{identifier}"
        before, new = b"previous source", b"pending PDF"
        self.put(f"{recovery}/source-before.bin", before)
        self.put(f"{recovery}/new.pdf", new)
        journal = {"version": 1, "id": identifier, "state": "published",
                   "source_path": str(self.books[0]), "source_before": hashlib.sha256(before).hexdigest(),
                   "pdf_path": "Certificates/pending.pdf", "pdf_sha256": hashlib.sha256(new).hexdigest(),
                   "previous_pdfs": []}
        path = self.put(f"Data/Issuance History/{identifier}.json", backup._json_bytes(journal))
        return path, journal

    def issued(self, identifier, *, cod="0001", key="CASE1", relative="Certificates/0001_client_1.pdf", text="original"):
        from reportlab.pdfgen.canvas import Canvas
        from certificate_inventory import record_issuance
        content = BytesIO()
        canvas = Canvas(content)
        canvas.drawString(40, 800, f"COD No: {cod}")
        canvas.drawString(40, 770, text)
        canvas.save()
        pdf = self.put(relative, content.getvalue())
        receipt = {"version": 1, "id": identifier, "state": "complete", "cod": cod,
                   "source_kind": "workbook", "plant": "plant1", "key": key,
                   "source_path": str(self.books[0]), "started_at": f"2026-10-{int(identifier[0], 16):02d}",
                   "pdf_path": relative, "pdf_sha256": hashlib.sha256(content.getvalue()).hexdigest()}
        self.put(f"Data/Issuance History/{identifier}.json", backup._json_bytes(receipt))
        self.put(f"Data/Issuance Recovery/{identifier}/new.pdf", content.getvalue())
        self.put(f"Data/Issuance Recovery/{identifier}/previous-0.pdf", b"previous output")
        self.put(f"Data/Issuance Recovery/{identifier}/source-before.bin", self.books[0].read_bytes())
        record_issuance(self.root, pdf, cod, "workbook", "plant1", key)
        return pdf, receipt

    def test_restore_same_cod_reissue_uses_old_receipt_pdf_and_ownership(self):
        from certificate_inventory import cod_conflicts, issuance_paths, scan_inventory
        from issuance_safety import latest_issuance, file_digest
        from run_certificates import find_certificate
        for renamed in (False, True):
            with self.subTest(renamed=renamed):
                old_pdf, old = self.issued("1" * 32)
                original_pdf = old_pdf.read_bytes()
                item = self.create()
                new_relative = "Certificates/0001_renamed_client_1.pdf" if renamed else old["pdf_path"]
                if renamed:
                    self.put("Data/Archive/Superseded Certificates/old.pdf", original_pdf)
                    old_pdf.unlink()
                new_pdf, new = self.issued("2" * 32, relative=new_relative, text="reissued")
                latest_pdf_bytes = new_pdf.read_bytes()
                result = self.restore(item)
                self.assertEqual(old_pdf.read_bytes(), original_pdf)
                latest = latest_issuance(self.root, "workbook", "plant1", "CASE1")
                self.assertEqual(latest["id"], old["id"])
                self.assertEqual(file_digest(old_pdf), latest["pdf_sha256"])
                self.assertEqual(find_certificate(self.root / "Certificates", {
                    "cod_no": "0001", "source_id": "plant1", "case_id": "CASE1"}), old_pdf)
                archived = Path(result["extras_path"])
                self.assertEqual(json.loads((archived / f"Data/Issuance History/{new['id']}.json").read_bytes()), new)
                self.assertEqual((archived / f"Data/Issuance Recovery/{new['id']}/new.pdf").read_bytes(), latest_pdf_bytes)
                self.assertFalse((self.root / f"Data/Issuance Recovery/{new['id']}/source-before.bin").exists())
                inventory = scan_inventory(self.root)
                allowed = issuance_paths(inventory, "1", "workbook", "plant1", "CASE1")
                self.assertEqual(allowed, [old["pdf_path"]])
                self.assertEqual(cod_conflicts(inventory, "1", allowed), [])
                if renamed:
                    self.assertFalse(new_pdf.exists())
                    self.assertEqual((archived / new_relative).read_bytes(), latest_pdf_bytes)

    def test_new_certificate_is_archived_but_cod_reserved_across_repeated_restores(self):
        from certificate_inventory import cod_conflicts, issuance_paths, scan_inventory
        from issuance_safety import latest_issuance
        from run_certificates import _highest_sequence
        self.issued("1" * 32)
        item = self.create()
        new_pdf, new = self.issued("3" * 32, cod="0099", key="NEWCASE", relative="Certificates/0099_new_client_1.pdf")
        original_pdf = new_pdf.read_bytes()
        result = self.restore(item)
        self.assertFalse(new_pdf.exists())
        self.assertEqual((Path(result["extras_path"]) / new["pdf_path"]).read_bytes(), original_pdf)
        self.assertIsNone(latest_issuance(self.root, "workbook", "plant1", "NEWCASE"))
        inventory = scan_inventory(self.root)
        conflicts = cod_conflicts(inventory, "99")
        self.assertTrue(conflicts)
        self.assertTrue(all(not entry["present"] and not entry.get("managed_owner") for entry in conflicts))
        self.assertEqual(issuance_paths(inventory, "99", "workbook", "plant1", "NEWCASE"), [])
        self.assertGreaterEqual(_highest_sequence([], self.settings, self.root / "Certificates", inventory=inventory), 99)
        archived_before = list((self.root / "Data/Archive/Restore Extras").rglob("*"))
        self.restore(item)
        self.assertTrue(cod_conflicts(scan_inventory(self.root), "99"))
        self.assertTrue(all(path.exists() for path in archived_before))
        # Subsequent snapshots preserve recovery receipts/binary sources in the extras archive.
        later = self.create()
        relatives = {e.get("relative_path") for e in self.manifest(later)["files"]}
        self.assertIn((Path(result["extras_path"]) / f"Data/Issuance History/{new['id']}.json").relative_to(self.root).as_posix(), relatives)
        self.assertIn((Path(result["extras_path"]) / f"Data/Issuance Recovery/{new['id']}/source-before.bin").relative_to(self.root).as_posix(), relatives)

    def test_deleted_certificate_and_receipt_only_cod_remain_reserved(self):
        from certificate_inventory import cod_conflicts, scan_inventory
        item = self.create()
        pdf, receipt = self.issued("3" * 32, cod="0098", relative="Certificates/0098_deleted_1.pdf")
        pdf.unlink()
        self.put("Data/Issuance History/receipt-only.json", backup._json_bytes({"version": 1, "state": "complete", "cod": "0097"}))
        self.restore(item)
        current = scan_inventory(self.root)
        self.assertTrue(cod_conflicts(current, "98"))
        self.assertTrue(cod_conflicts(current, "97"))

    def test_failed_restore_rolls_back_extras_and_exact_original_ledger(self):
        item = self.create()
        pdf, receipt = self.issued("3" * 32, cod="0099", relative="Certificates/0099_new_1.pdf")
        ledger = (self.root / "Data/Certificate Inventory.json").read_bytes()
        expected = {p: p.read_bytes() for p in [pdf, self.root / f"Data/Issuance History/{receipt['id']}.json",
                    *list((self.root / f"Data/Issuance Recovery/{receipt['id']}").iterdir())]}
        original_replace = backup._replace_file
        calls = 0
        def fail_once(source, target):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise PermissionError("injected restore replacement failure")
            return original_replace(source, target)
        with patch.object(backup, "_replace_file", side_effect=fail_once), self.assertRaises(backup.RestoreError) as failure:
            self.restore(item)
        self.assertTrue(failure.exception.rolled_back)
        for path, content in expected.items():
            self.assertEqual(path.read_bytes(), content)
        self.assertEqual((self.root / "Data/Certificate Inventory.json").read_bytes(), ledger)
        self.assertEqual(backup.check_restore_health(self.root), {"healthy": True})

    def test_partial_archive_move_failure_rolls_back_all_extras(self):
        item = self.create()
        pdf, receipt = self.issued("3" * 32, cod="0099", relative="Certificates/0099_new_1.pdf")
        original = backup._move_extra
        calls = 0
        def fail_after_move(source, destination, expected):
            nonlocal calls
            calls += 1
            original(source, destination, expected)
            if calls == 2:
                raise OSError("archive move completed before simulated failure")
        with patch.object(backup, "_move_extra", side_effect=fail_after_move), self.assertRaises(backup.RestoreError) as failure:
            self.restore(item)
        self.assertTrue(failure.exception.rolled_back)
        self.assertTrue(pdf.exists())
        self.assertEqual(json.loads((self.root / f"Data/Issuance History/{receipt['id']}.json").read_bytes()), receipt)

    def test_pending_issuance_and_recovery_snapshots_are_preserved(self):
        path, journal = self.pending_issuance()
        item = self.create()
        self.restore(item)
        self.assertEqual(json.loads(path.read_bytes()), journal)
        self.assertEqual((self.root / f"Data/Issuance Recovery/{journal['id']}/source-before.bin").read_bytes(), b"previous source")

    def test_pending_journal_cannot_direct_startup_recovery_to_another_project(self):
        path, journal = self.pending_issuance()
        unrelated = self.base / "other project.xlsx"
        unrelated.write_bytes(b"do not touch")
        journal["source_path"] = str(unrelated)
        path.write_bytes(backup._json_bytes(journal))
        item = self.create()
        with self.assertRaisesRegex(backup.BackupError, "unapproved original source"):
            self.restore(item)
        self.assertEqual(unrelated.read_bytes(), b"do not touch")

    def test_pending_journal_missing_snapshots_is_refused_before_overwrites(self):
        _, journal = self.pending_issuance()
        (self.root / f"Data/Issuance Recovery/{journal['id']}/source-before.bin").unlink()
        item = self.create()
        with self.assertRaisesRegex(backup.BackupError, "recovery snapshot"):
            self.restore(item)

    def test_missing_workbook_payload_is_refused_even_with_original_mapping(self):
        item = self.create()
        def remove(manifest, contents):
            entry = manifest["files"].pop(0)
            contents.pop(entry["member"])
            for index, remaining in enumerate(manifest["files"]):
                old = remaining["member"]
                contents[f"payload/{index:06d}"] = contents.pop(old)
                remaining["member"] = f"payload/{index:06d}"
        self.rewrite_archive(item, remove)
        with self.assertRaisesRegex(backup.BackupError, "missing an original workbook"):
            self.restore(item)

    def test_size_limited_backup_leaves_no_partial_archive(self):
        with patch.object(backup, "MAX_FILE_BYTES", 1), self.assertRaises(backup.BackupError):
            self.create()
        self.assertEqual(list((self.root / "Data/Backups").glob("*.partial")), [])
        self.assertEqual(backup.list_backups(self.root), [])

    def test_sqlite_is_preserved_and_active_sidecars_are_rejected(self):
        import sqlite3
        database = self.root / "Data/certificate-history.sqlite"
        with closing(sqlite3.connect(database)) as connection, connection:
            connection.execute("CREATE TABLE batches (cod TEXT)")
            connection.execute("INSERT INTO batches VALUES ('0017')")
        item = self.create()
        database.write_bytes(b"corrupt database")
        self.restore(item)
        with closing(sqlite3.connect(database)) as connection:
            self.assertEqual(connection.execute("SELECT cod FROM batches").fetchall(), [("0017",)])
        Path(str(database) + "-wal").write_bytes(b"active WAL")
        with self.assertRaisesRegex(backup.BackupError, "checkpoint"):
            self.create()
        with self.assertRaisesRegex(backup.BackupError, "checkpoint"):
            self.restore(item)

    def test_run_lock_conflicts_and_optional_existing_lock(self):
        item = self.create()
        with run_lock(self.root):
            with self.assertRaises(PermissionError):
                self.create()
            with self.assertRaises(PermissionError):
                self.restore(item)
            nested = backup.create_backup(self.root, locked=True)
            restored = backup.restore_backup(self.root, nested["backup_id"], confirmed=True, locked=True)
            self.assertGreater(restored["restored_files"], 0)

    def test_excel_workbook_lock_prevents_backup_and_restore(self):
        item = self.create()
        with exclusive_file(self.books[0]):
            with self.assertRaises(PermissionError):
                self.create()
            with self.assertRaises(PermissionError):
                self.restore(item)

    def test_failure_rolls_back_every_attempted_file_and_retains_originals(self):
        item = self.create()
        self.books[0].write_bytes(b"current plant 1")
        self.books[1].write_bytes(b"current plant 2")
        before = [p.read_bytes() for p in self.books]
        replace = backup._replace_file
        calls = 0
        def fail_once(source, target):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise PermissionError("simulated blocked replacement")
            return replace(source, target)
        with patch.object(backup, "_replace_file", side_effect=fail_once):
            with self.assertRaises(backup.RestoreError) as failure:
                self.restore(item)
        self.assertTrue(failure.exception.rolled_back)
        self.assertEqual([p.read_bytes() for p in self.books], before)
        recovery = Path(failure.exception.recovery_path)
        self.assertEqual((recovery / "originals/0").read_bytes(), before[0])
        self.assertEqual(json.loads((recovery / "journal.json").read_text())["state"], "rolled_back")

    def test_rollback_of_new_files_moves_them_to_recoverable_siblings(self):
        item = self.create()
        self.books[0].unlink()
        replace = backup._replace_file
        calls = 0
        def fail_once(source, target):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError("injected failure")
            return replace(source, target)
        with patch.object(backup, "_replace_file", side_effect=fail_once), self.assertRaises(backup.RestoreError) as failure:
            self.restore(item)
        self.assertTrue(failure.exception.rolled_back)
        self.assertFalse(self.books[0].exists())
        preserved = list(self.root.glob("Plant 1.xlsx.restore-recovery-*"))
        self.assertEqual(len(preserved), 1)
        with ZipFile(preserved[0]) as archive:
            self.assertEqual(archive.read("xl/workbook.xml"), b"plant 0")

    def test_failed_rollback_blocks_further_mutation_and_keeps_original(self):
        item = self.create()
        self.books[0].write_bytes(b"recover me")
        replace = backup._replace_file
        calls = 0
        def fail_restore_and_rollback(source, target):
            nonlocal calls
            calls += 1
            if calls >= 2:
                raise OSError("disk unavailable")
            return replace(source, target)
        with patch.object(backup, "_replace_file", side_effect=fail_restore_and_rollback), self.assertRaises(backup.RestoreError) as failure:
            self.restore(item)
        self.assertFalse(failure.exception.rolled_back)
        self.assertEqual((Path(failure.exception.recovery_path) / "originals/0").read_bytes(), b"recover me")
        with self.assertRaisesRegex(backup.BackupError, "interrupted restore"):
            self.create()
        with self.assertRaisesRegex(backup.BackupError, "interrupted restore"):
            self.restore(item)

    def test_process_interruption_journal_blocks_backup_and_restore(self):
        item = self.create()
        self.put("Data/Backups/Restore Transactions/interrupted/journal.json", b'{"state":"in_progress"}')
        with self.assertRaisesRegex(backup.BackupError, "supervised recovery"):
            self.create()
        with self.assertRaisesRegex(backup.BackupError, "supervised recovery"):
            self.restore(item)
        self.assertTrue(backup.list_backups(self.root)[0]["valid"])

    def test_public_restore_health_guard_is_read_only_and_blocks_interrupted_restore(self):
        before = sorted(p.relative_to(self.root).as_posix() for p in self.root.rglob("*"))
        self.assertEqual(backup.check_restore_health(self.root), {"healthy": True})
        self.assertEqual(before, sorted(p.relative_to(self.root).as_posix() for p in self.root.rglob("*")))
        path = self.put("Data/Backups/Restore Transactions/test/journal.json", b'{"state":"in_progress"}')
        original = path.read_bytes()
        with self.assertRaisesRegex(backup.BackupError, "supervised recovery"):
            backup.check_restore_health(self.root)
        self.assertEqual(path.read_bytes(), original)
        for state in ("prepared", "committed", "rolled_back"):
            path.write_bytes(backup._json_bytes({"state": state}))
            self.assertEqual(backup.check_restore_health(self.root), {"healthy": True})
        for content in (b"broken", b"null", b"{}", b'{"state":"rollback_failed"}'):
            path.write_bytes(content)
            with self.assertRaises(backup.BackupError):
                backup.check_restore_health(self.root)
            self.assertEqual(path.read_bytes(), content)

    def test_empty_persisted_destination_selects_local_default(self):
        self.settings["backup_destination"] = ""
        self.write_settings()
        item = self.create()
        self.assertEqual(Path(item["path"]).parent, self.root / "Data/Backups")

    def test_read_only_template_and_archive_can_be_backed_up_and_inspected(self):
        template = self.root / "App/templates/certificate-types/standard.pdf"
        template.chmod(0o444)
        self.addCleanup(lambda: template.chmod(0o666))
        item = self.create()
        archive = Path(item["path"])
        archive.chmod(0o444)
        self.addCleanup(lambda: archive.chmod(0o666))
        self.assertTrue(backup.list_backups(self.root)[0]["valid"])

    def test_hardlinks_are_rejected_for_sources_and_restore_targets(self):
        original = self.root / "Certificates/0001_client_1.pdf"
        os.link(original, self.base / "hardlink.pdf")
        with self.assertRaisesRegex(backup.BackupError, "Hard-linked"):
            self.create()
        (self.base / "hardlink.pdf").unlink()
        item = self.create()
        os.link(self.books[0], self.base / "linked.xlsx")
        with self.assertRaisesRegex(backup.BackupError, "Hard-linked"):
            self.restore(item)

    def test_symlink_sources_destinations_and_restore_parent_are_rejected(self):
        link = self.root / "Data/Attachments/link.pdf"
        try:
            link.symlink_to(self.books[0])
        except (OSError, NotImplementedError):
            self.skipTest("Creating symlinks requires Windows developer mode or privileges.")
        with self.assertRaisesRegex(backup.BackupError, "Links"):
            self.create()
        link.unlink()
        item = self.create()
        external = self.base / "outside"
        external.mkdir()
        destination = self.base / "linked-destination"
        destination.symlink_to(external, target_is_directory=True)
        with self.assertRaisesRegex(backup.BackupError, "Links"):
            self.create(destination)
        attachments = self.root / "Data/Attachments"
        shutil.rmtree(attachments)
        attachments.symlink_to(external, target_is_directory=True)
        with self.assertRaisesRegex(backup.BackupError, "Links"):
            self.restore(item)
        self.assertEqual(list(external.iterdir()), [])

    @unittest.skipUnless(os.name == "nt", "Windows junction test")
    def test_windows_junctions_are_rejected_without_following_them(self):
        import subprocess
        external = self.base / "external"
        external.mkdir()
        external.joinpath("historic.pdf").write_bytes(b"outside content")
        junction = self.root / "Certificates/junction"
        result = subprocess.run(["cmd", "/c", "mklink", "/J", str(junction), str(external)],
                                capture_output=True, check=False)
        if result.returncode:
            self.skipTest("Junction creation unavailable")
        self.addCleanup(lambda: junction.rmdir() if junction.exists() else None)
        with self.assertRaisesRegex(backup.BackupError, "reparse"):
            self.create()
        self.assertEqual(external.joinpath("historic.pdf").read_bytes(), b"outside content")


if __name__ == "__main__":
    unittest.main()
