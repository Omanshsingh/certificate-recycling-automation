"""Protect observed external PDF identifiers using disposable workspaces."""
from __future__ import annotations

from datetime import date
from io import BytesIO
import json
import os
from pathlib import Path
import tempfile
import unittest
import zlib
from unittest.mock import patch

from openpyxl import Workbook
from reportlab.lib.pdfencrypt import StandardEncryption
from reportlab.pdfgen import canvas

from certificate_core import read_batches
import certificate_drafts as drafts
from certificate_inventory import (cod_conflicts, ensure_certificate_folders,
                                   inspect_pdf, issuance_paths, record_issuance, reference_warnings,
                                   safe_certificate_path, scan_inventory, text_identifiers, verify_publication)
import run_certificates as runner
import web_app
from run_certificates import certificate_filename, run
from workflow_support import save_json


AS_OF = date(2026, 10, 1)
HEADERS = ["SN", "Inward Date", "Supplier Name", "Location", "Item",
           "Quantity in Pcs", "Per Pcs Weight in (KG)", "Weight (KG)",
           "Form 6 No", "Reference No.", "COD No", "COD Date", "Certificate Issued",
           "Certificate Type", "Certificate Fields"]


def external_pdf(path, *, cod=None, reference=None, layout=False, password=None):
    """A readable certificate, including the native two-column label layout."""
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = canvas.Canvas(str(path), pagesize=(600, 800), pageCompression=0, invariant=True,
                        encrypt=StandardEncryption(password) if password else None)
    pdf.setFont("Helvetica", 12)
    pdf.drawString(40, 760, "External Certificate of Recycling")
    if cod is not None:
        if layout:
            pdf.drawString(40, 720, "Certificate No. & Date")
            pdf.drawString(290, 720, cod)
            pdf.drawString(410, 720, "01.10.2026")
        else:
            pdf.drawString(40, 720, f"Certificate No.: {cod}")
    if reference is not None:
        if layout:
            pdf.drawString(40, 680, "Reference No.")
            pdf.drawString(290, 680, reference)
        else:
            pdf.drawString(40, 680, f"Reference No.: {reference}")
    pdf.drawString(40, 640, "Weight: 7.5 KG")
    pdf.drawString(40, 610, "Model Serial Number: 9999")
    pdf.showPage()
    pdf.save()
    return path


def register(path, *, case_id, cod="", status="Ready", reference="WORKBOOK-REF"):
    book = Workbook()
    try:
        sheet = book.active
        sheet.title = "Inward Register"
        sheet.append(HEADERS)
        sheet.append([case_id, date(2026, 9, 1), "Independent Client Ltd", "Pune", "Servers",
                      2, 5, 10, f"FORM-{zlib.crc32((str(path) + case_id).encode()):08X}", reference, cod, AS_OF if cod else None,
                      status, "sita", "{}"])
        book.save(path)
    finally:
        book.close()


class TemporaryInventoryTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.folder = self.root / "Certificates"
        self.folder.mkdir()
        self.ledger = self.root / "Data" / "Certificate Inventory.json"

    def issued_pdf(self, response):
        return self.folder / response.get("relative_path", response["filename"])

    def entry(self, inventory, path):
        relative = path.relative_to(self.root).as_posix()
        return next(entry for entry in inventory["entries"] if entry["path"] == relative)

    def replace_preserving_metadata(self, path, old, new):
        before = path.stat()
        original = path.read_bytes()
        self.assertIn(old, original)
        self.assertEqual(len(old), len(new))
        path.write_bytes(original.replace(old, new))
        os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
        after = path.stat()
        self.assertEqual((after.st_size, after.st_mtime_ns), (before.st_size, before.st_mtime_ns))


class CertificateFolderStructureTests(TemporaryInventoryTest):
    def test_pdfs_placed_in_the_shared_folder_are_available_as_external(self):
        self.assertEqual(ensure_certificate_folders(self.root), self.folder)
        pdf = external_pdf(self.folder / '0001_Client_City.pdf', cod='0001')
        entry, = scan_inventory(self.root)['entries']
        self.assertEqual(entry['path'], 'Certificates/' + pdf.name)
        self.assertFalse(entry['external'])
        self.assertTrue(entry['linkable_external'])
        record_issuance(self.root, pdf, '0001', 'workbook', 'plant1', '1')
        entry, = scan_inventory(self.root)['entries']
        self.assertFalse(entry['external'])
        self.assertTrue(entry['app_generated'])
        self.assertFalse(entry['linkable_external'])

    def test_read_only_scan_does_not_create_destination_folders(self):
        inventory = scan_inventory(self.root)
        self.assertEqual(inventory['entries'], [])
        self.assertEqual(list(self.folder.iterdir()), [])

    def test_pdf_path_validation_allows_nested_certificate_folders_but_not_traversal(self):
        self.assertEqual(safe_certificate_path(self.root, 'legacy.pdf'), self.folder / 'legacy.pdf')
        self.assertEqual(safe_certificate_path(self.root, 'Auto Generated/new.pdf'),
                         self.folder / 'Auto Generated' / 'new.pdf')
        for relative in ('Unknown/nested.pdf', 'Auto Generated/deeper/file.pdf'):
            self.assertEqual(safe_certificate_path(self.root, relative), self.folder / relative)
        for unsafe in ('../outside.pdf', '/outside.pdf', 'not-a-pdf.txt',
                       'Manually Added\\..\\outside.pdf'):
            with self.subTest(path=unsafe), self.assertRaises(ValueError):
                safe_certificate_path(self.root, unsafe)


class InventoryApiTests(TemporaryInventoryTest):
    def test_labelled_text_ignores_unlabelled_dates_weights_and_model_numbers(self):
        cods, references = text_identifiers(
            "Certificate No. & Date: 0319 01.10.2026\nReference No.: REF-EXTERNAL\n"
            "Weight: 7.5 KG\nModel Serial Number: 9999\nDate: 01.10.2026\n")
        self.assertEqual((cods, references), (["0319"], ["REF-EXTERNAL"]))
        self.assertEqual(text_identifiers("Certificate Date: 01.10.2026\nWeight: 70\nSerial: 9999"), ([], []))
        self.assertEqual(text_identifiers("COD No.: 01.10.2026\nReference No.: N/A"), ([], []))

    def test_all_labels_before_values_cannot_claim_next_line_as_cod_or_reference(self):
        self.assertEqual(text_identifiers(
            "Certificate No. & Date\nReference No.\nWeight\nModel Serial Number\n"
            "0319 01.10.2026\nREF-EXTERNAL\n7.5 KG\n9999\n"), ([], []))

    def test_printed_identifiers_are_read_from_plain_and_table_layout_pdfs(self):
        for layout in (False, True):
            with self.subTest(layout=layout):
                path = external_pdf(self.folder / f"external-{layout}.pdf", cod="0319",
                                    reference="REF-EXTERNAL", layout=layout)
                cods, references, warnings = inspect_pdf(path)
                self.assertEqual((cods, references, warnings), (["0319"], ["REF-EXTERNAL"], []))

    def test_scan_is_recursive_case_insensitive_and_ignores_other_folders_and_hidden_work(self):
        valid = external_pdf(self.folder / "nested" / "0319_External_Client_1.PDF", cod="0319")
        external_pdf(self.root / "QA" / "9999_Probe_1.pdf", cod="9999")
        external_pdf(self.root / "Data" / "Previews" / "9998.pdf", cod="9998")
        for name in (".9997.tmp.pdf", ".9996.rollback.pdf", ".9995.pdf"):
            external_pdf(self.folder / name, cod="9997")
        inventory = scan_inventory(self.root)
        self.assertEqual([entry["path"] for entry in inventory["entries"]], [valid.relative_to(self.root).as_posix()])
        self.assertEqual(inventory["entries"][0]["cods"], ["0319"])
        self.assertFalse(self.ledger.exists())

    def test_filename_cod_supports_leading_numbers_and_legacy_names(self):
        for name, cod in (("0319_Client_City_1.pdf", "0319"), ("0042.pdf", "0042"),
                          ("COD_0043.pdf", "0043"), ("Certificate_0044.pdf", "0044"),
                          ("COD_0040_clientcity_1.pdf", "0040")):
            with self.subTest(name=name):
                path = external_pdf(self.folder / name)
                cods, _, _ = inspect_pdf(path)
                self.assertIn(cod, cods)

    def test_filename_and_printed_cod_mismatch_reserves_both_and_warns(self):
        path = external_pdf(self.folder / "0319_External_Client_1.pdf", cod="0499", reference="EXT-REF")
        inventory = scan_inventory(self.root)
        entry = self.entry(inventory, path)
        self.assertEqual(set(entry["cods"]), {"0319", "0499"})
        self.assertTrue(any("differ" in warning.casefold() for warning in entry["warnings"]))
        self.assertTrue(inventory["warnings"])
        self.assertEqual(cod_conflicts(inventory, "319"), [entry])
        self.assertEqual(cod_conflicts(inventory, "0499"), [entry])

    def test_unreadable_and_encrypted_pdfs_warn_but_reserve_filename_numbers(self):
        unreadable = self.folder / "0701_Broken_1.pdf"
        unreadable.write_bytes(b"This is not a PDF package.")
        encrypted = external_pdf(self.folder / "0702_Locked_2.pdf", cod="0999", password="test-password")
        inventory = scan_inventory(self.root)
        for path, cod in ((unreadable, "0701"), (encrypted, "0702")):
            with self.subTest(path=path.name):
                entry = self.entry(inventory, path)
                self.assertEqual(entry["cods"], [cod])
                self.assertTrue(entry["warnings"])
                self.assertTrue(cod_conflicts(inventory, cod))
        self.assertFalse(cod_conflicts(inventory, "0999"))
        self.assertEqual(len(inventory["warnings"]), 2)

    def test_persisted_inventory_remembers_deleted_identifiers_without_rescanning_them(self):
        path = external_pdf(self.folder / "0319_External_1.pdf", cod="0319", reference="REF-OLD")
        original = path.read_bytes()
        first = scan_inventory(self.root, persist=True)
        self.assertTrue(self.ledger.exists())
        self.assertEqual(path.read_bytes(), original)
        entry = self.entry(first, path)
        self.assertEqual((entry["present"], entry["size"], entry["mtime_ns"]),
                         (True, path.stat().st_size, path.stat().st_mtime_ns))
        path.unlink()
        with patch("certificate_inventory.inspect_pdf", side_effect=AssertionError("deleted file was inspected")):
            second = scan_inventory(self.root, persist=True)
        deleted = self.entry(second, path)
        self.assertFalse(deleted["present"])
        self.assertEqual((deleted["cods"], deleted["references"]), (["0319"], ["REF-OLD"]))
        self.assertEqual(cod_conflicts(second, "319"), [deleted])
        self.assertFalse(reference_warnings(second, "REF-OLD"))
        # An ownership flag alone cannot exempt an unrelated historical PDF.
        self.assertEqual(cod_conflicts(second, "319", owned=True), [deleted])
        self.assertFalse(reference_warnings(second, "REF-OLD", owned=True))
        self.assertFalse(cod_conflicts(second, "319", allowed_paths=[deleted["path"]], owned=True))
        self.assertFalse(reference_warnings(second, "REF-OLD", allowed_paths=[deleted["path"]], owned=True))
        self.assertTrue(list((self.root / "Data" / "Certificate Inventory Backups").glob("*.json")))

    def test_read_only_scan_does_not_create_or_change_ledger(self):
        path = external_pdf(self.folder / "0319_External_1.pdf", cod="0319")
        scan_inventory(self.root, persist=False)
        self.assertFalse(self.ledger.exists())
        scan_inventory(self.root, persist=True)
        ledger = self.ledger.read_bytes()
        external_pdf(self.folder / "0320_Another_1.pdf", cod="0320")
        inventory = scan_inventory(self.root, persist=False)
        self.assertEqual(len(inventory["entries"]), 2)
        self.assertEqual(self.ledger.read_bytes(), ledger)
        self.assertTrue(path.exists())

    def test_correcting_pdf_keeps_all_previously_observed_numbers(self):
        path = external_pdf(self.folder / "0319_External_1.pdf", cod="0419", reference="REF-OLD")
        scan_inventory(self.root, persist=True)
        external_pdf(path, cod="0520", reference="REFERENCE-NEW-LONGER")
        inventory = scan_inventory(self.root, persist=True)
        entry = self.entry(inventory, path)
        self.assertEqual(set(entry["cods"]), {"0319", "0419", "0520"})
        self.assertEqual(set(entry["references"]), {"REF-OLD", "REFERENCE-NEW-LONGER"})

    def test_content_digest_refreshes_identifiers_when_size_and_mtime_are_unchanged(self):
        path = external_pdf(self.folder / "external-document.pdf", cod="0319", reference="REF-A")
        first = scan_inventory(self.root, persist=True)
        self.replace_preserving_metadata(path, b"0319", b"0320")
        self.replace_preserving_metadata(path, b"REF-A", b"REF-B")
        second = scan_inventory(self.root, persist=True)
        entry = self.entry(second, path)
        self.assertEqual(set(entry["cods"]), {"0319", "0320"})
        self.assertEqual(entry["current_cods"], ["0320"])
        self.assertEqual(set(entry["references"]), {"REF-A", "REF-B"})
        self.assertNotEqual(entry["sha256"], self.entry(first, path)["sha256"])

    def test_publication_detects_changed_owned_content_even_with_identical_metadata(self):
        path = external_pdf(self.folder / "0319_Owned_1.pdf", cod="0319", reference="REF-A")
        snapshot = scan_inventory(self.root, persist=True)
        ledger = self.ledger.read_bytes()
        self.replace_preserving_metadata(path, b"Model Serial Number: 9999", b"Model Serial Number: 8888")
        changed = path.read_bytes()
        with self.assertRaisesRegex(ValueError, "existing certificate changed"):
            verify_publication(self.root, snapshot, "0319", "REF-A",
                               allowed_paths=[path.relative_to(self.root).as_posix()], owned=True)
        self.assertEqual(path.read_bytes(), changed)
        self.assertEqual(self.ledger.read_bytes(), ledger)

    def test_committed_owner_tag_only_exempts_that_records_observed_paths(self):
        own = external_pdf(self.folder / "0016_Owned_1.pdf", cod="0016", reference="SHARED-REF")
        foreign = external_pdf(self.folder / "0319_Foreign_2.pdf", cod="0319", reference="SHARED-REF")
        inventory = record_issuance(self.root, own, "0016", "manual", "plant1", "record-1")
        allowed = issuance_paths(inventory, "16", "manual", "plant1", "record-1")
        self.assertEqual(allowed, [own.relative_to(self.root).as_posix()])
        self.assertEqual(issuance_paths(inventory, "16", "manual", "plant2", "record-1"), [])
        self.assertEqual(issuance_paths(inventory, "16", "manual", "plant1", "other-record"), [])
        own.unlink()
        foreign.unlink()
        inventory = scan_inventory(self.root, persist=True)
        self.assertFalse(cod_conflicts(inventory, "16", allowed_paths=allowed, owned=True))
        self.assertTrue(cod_conflicts(inventory, "319", allowed_paths=allowed, owned=True))
        warnings = reference_warnings(inventory, "SHARED-REF", allowed_paths=allowed, owned=True)
        self.assertEqual(warnings, [])

    def test_allowed_owned_pdf_is_exempt_but_mismatched_printed_number_is_not(self):
        path = external_pdf(self.folder / "0319_External_1.pdf", cod="0319", reference="REF-OLD")
        inventory = scan_inventory(self.root)
        allowed = (path.relative_to(self.root).as_posix().upper().replace("/", "\\"),)
        self.assertFalse(cod_conflicts(inventory, "319", allowed_paths=allowed))
        self.assertFalse(reference_warnings(inventory, "ref-old", allowed_paths=allowed))
        external_pdf(path, cod="0520", reference="REF-OLD")
        inventory = scan_inventory(self.root)
        self.assertTrue(cod_conflicts(inventory, "319", allowed_paths=allowed, owned=True))

    def test_duplicate_reference_is_silent_without_reserving_an_unrelated_cod(self):
        path = external_pdf(self.folder / "0319_External_1.pdf", cod="0319", reference="REF-EXTERNAL")
        inventory = scan_inventory(self.root)
        warnings = reference_warnings(inventory, "ref-external")
        self.assertEqual(warnings, [])
        self.assertFalse(cod_conflicts(inventory, "0199"))
        self.assertFalse(reference_warnings(inventory, ""))
        self.assertFalse(reference_warnings(inventory, "DIFFERENT-REF"))

    def test_corrupt_ledger_is_a_hard_error_and_is_not_overwritten(self):
        self.ledger.parent.mkdir(parents=True)
        for content in ("{broken", json.dumps({"entries": [{"path": "Certificates/a.pdf", "cods": [319]}]})):
            with self.subTest(content=content):
                self.ledger.write_text(content, encoding="utf-8")
                with self.assertRaisesRegex(ValueError, "inventory is unreadable"):
                    scan_inventory(self.root, persist=True)
                self.assertEqual(self.ledger.read_text(encoding="utf-8"), content)


class InventoryWorkflowTests(TemporaryInventoryTest):
    def setUp(self):
        super().setUp()
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        # Numbering-only markers do not claim a current certificate exists.
        register(self.paths[0], case_id="EXCEL-1", cod="0010", status="N/A")
        register(self.paths[1], case_id="EXCEL-2", cod="0015", status="N/A")
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "certificate_number_prefix": "", "certificate_number_start": 1,
            "certificate_number_width": 4, "reminder_time": "11:00", "reminders_enabled": False,
        })

    def details(self, **changes):
        details = {"case_id": "MANUAL-1", "supplier": "Independent Client Ltd", "location": "Pune",
                   "certificate_type": "sita", "certificate_fields": {}, "inward_date": "01.09.2026",
                   "certificate_weight": "8.5", "weight": "10", "reference": "MANUAL-REF"}
        details.update(changes)
        return details

    def generate_manual(self, **options):
        return drafts.generate_draft(self.root, plant="plant1", details=self.details(), as_of=AS_OF, **options)

    def results(self):
        return json.loads((self.root / "Data" / "Reports" / "Latest Results.json").read_text(encoding="utf-8"))

    def test_manual_automatic_number_reserves_printed_cod_from_external_named_pdf(self):
        external_pdf(self.folder / "Manually Added" / "external-document.pdf", cod="0319", reference="EXT-REF", layout=True)
        response = self.generate_manual()
        self.assertEqual((response["created"], response["certificate_no"]), (1, "0320"))

    def test_workbook_automatic_number_reserves_printed_cod_from_external_pdf(self):
        external_pdf(self.folder / "Manually Added" / "external-document.pdf", cod="0319", reference="EXT-REF")
        register(self.paths[0], case_id="NEW-EXCEL")
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual((created, review), (1, 0))
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0320")

    def test_manual_cod_and_date_inputs_are_ignored_and_values_are_assigned_from_history(self):
        external_pdf(self.folder / "Manually Added" / "0319_External.pdf", cod="0319")
        response = drafts.generate_draft(self.root, plant="plant1",
            details=self.details(cod_no="0999", cod_date="2020-01-01"), as_of=AS_OF)
        self.assertEqual((response["certificate_no"], drafts.read_register(self.root)["records"][0]["details"]["cod_date"]),
                         ("0320", AS_OF.isoformat()))

    def test_pdf_without_identifiable_cod_blocks_generation_with_a_clear_recovery_path(self):
        from pypdf import PdfWriter
        unknown = self.folder / "Manually Added" / "unknown-scan.pdf"
        unknown.parent.mkdir(parents=True, exist_ok=True)
        writer = PdfWriter(); writer.add_blank_page(width=300, height=400)
        with unknown.open("wb") as stream:
            writer.write(stream)
        register(self.paths[0], case_id="NEW-EXCEL")
        created, _, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual(created, 0)
        row = next(row for row in self.results()["rows"] if row.get("sn") == "NEW-EXCEL")
        self.assertTrue(row["issuance_blocked"])
        self.assertIn("rename the PDF", row["detail"])
        self.assertIn("outside the Certificates folder", row["detail"])
        with self.assertRaisesRegex(ValueError, "rename the PDF"):
            self.generate_manual()

    def test_manual_external_duplicate_requires_confirmation_and_never_overwrites(self):
        external = external_pdf(self.folder / "0319_Independent_Client_Ltd_Pune_1.pdf",
                                cod="0319", reference="EXT-REF")
        original = external.read_bytes()
        details = self.details(cod_no="0319")
        response = drafts.generate_draft(self.root, plant="plant1", details=details, as_of=AS_OF)
        self.assertEqual((response["created"], response["certificate_no"]), (1, "0320"))
        self.assertEqual(external.read_bytes(), original)
        self.assertNotEqual(response["filename"], external.name)

    def test_workbook_external_duplicate_requires_confirmation_and_never_overwrites(self):
        register(self.paths[0], case_id="NEW-EXCEL")
        batch = read_batches(self.paths[0])[0]
        external = external_pdf(self.folder / certificate_filename(batch, 1, "0319"), cod="0319", reference="EXT-REF")
        original = external.read_bytes()
        options = {"as_of": AS_OF, "progress": lambda _: None, "selected": ("plant1", "NEW-EXCEL"), "custom_cod": "0319"}
        created, _, _ = run(self.paths, self.root, **options)
        self.assertEqual(created, 0)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "")
        self.assertEqual(external.read_bytes(), original)
        created, _, _ = run(self.paths, self.root, allow_warnings=True, **options)
        self.assertEqual(created, 1)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0320")
        self.assertEqual(external.read_bytes(), original)

    def test_create_api_rejects_user_entered_cod_even_when_a_foreign_pdf_matches(self):
        external = external_pdf(self.folder / "0319_Independent_Client_Ltd_Pune_1.pdf",
                                cod="0319", reference="EXT-REF")
        original = external.read_bytes()
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "BACKUP_DIR", self.root / "Data" / "Workbook Backups"))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", self.root / "Data" / "Reports" / "Latest Results.json"))

        def create(confirmed):
            from workspace_review import workspace_revision
            body = json.dumps({"plant": "plant1", "confirmed": confirmed,
                               'expected_revision': workspace_revision(self.root),
                               'review_token': response.get('review_token') if confirmed else None,
                               'override_reason': 'Manager approved replacement COD' if confirmed else '',
                               "details": self.details(case_id="API-NEW", cod_no="0319", form6="FORM-API-NEW", item="Servers",
                                                       quantity="2", per_piece_weight="5", weight="10")}).encode()
            handler = object.__new__(web_app.Handler)
            handler.path = "/api/create"
            handler.headers = {"Content-Length": str(len(body))}
            handler.rfile = BytesIO(body)
            with patch.object(handler, "_same_origin", return_value=True), \
                    patch.object(handler, "send_json") as sent, \
                    patch.object(handler, "send_error_json") as errors:
                handler.do_POST()
            sent.assert_not_called()
            errors.assert_called_once()
            return str(errors.call_args.args[0])

        message = create(False)
        self.assertIn("assigned automatically", message)
        self.assertEqual(external.read_bytes(), original)
        inventory = scan_inventory(self.root)
        self.assertEqual(issuance_paths(inventory, "0319", "workbook", "plant1", "API-NEW"), [])
        self.assertFalse(any(batch["case_id"] == "API-NEW" for batch in read_batches(self.paths[0])))
        self.assertEqual(external.read_bytes(), original)

    def test_saved_manual_draft_ignores_typed_cod_without_committed_pdf(self):
        external = external_pdf(self.folder / "0319_Independent_Client_Ltd_Pune_1.pdf",
                                cod="0319", reference="EXT-REF")
        original = external.read_bytes()
        record_id = drafts.save_draft(self.root, "plant1", self.details(cod_no="0319"))
        self.assertEqual(drafts.read_register(self.root)["records"][0]["details"]["cod_no"], "")
        response = drafts.generate_draft(self.root, record_id=record_id, as_of=AS_OF)
        self.assertEqual((response["created"], response["certificate_no"]), (1, "0320"))
        self.assertEqual(external.read_bytes(), original)
        self.assertEqual(response["record_id"], record_id)
        self.assertTrue(external.is_file(), "A never-issued draft must not archive a foreign PDF found by typed COD")
        self.assertEqual(external.read_bytes(), original)

    def test_reference_collision_is_silent_and_cod_stays_automatic(self):
        external = external_pdf(self.folder / "0319_External_2.pdf", cod="0319", reference="SHARED-REF")
        original = external.read_bytes()
        details = self.details(cod_no="0199", reference="SHARED-REF")
        response = drafts.generate_draft(self.root, plant="plant1", details=details, as_of=AS_OF)
        self.assertFalse(response["requires_confirmation"])
        self.assertFalse(any(warning.startswith('Reference ') for warning in response['warnings']))
        self.assertEqual(response["certificate_no"], "0320")
        self.assertEqual(drafts.read_register(self.root)["records"][0]["details"]["reference"], "SHARED-REF")
        self.assertEqual(external.read_bytes(), original)

    def test_reference_collision_is_silent_and_workbook_cod_stays_unchanged(self):
        external = external_pdf(self.folder / "0319_External_2.pdf", cod="0319", reference="SHARED-REF")
        original = external.read_bytes()
        register(self.paths[0], case_id="NEW-EXCEL", reference="SHARED-REF")
        options = {"as_of": AS_OF, "progress": lambda _: None, "selected": ("plant1", "NEW-EXCEL"), "custom_cod": "0199"}
        created, _, _ = run(self.paths, self.root, **options)
        self.assertEqual(created, 1)
        self.assertFalse(any(warning.startswith('Reference ') for row in self.results()['rows'] for warning in row.get('warnings', [])))
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch["cod_no"], batch["reference"]), ("0199", "SHARED-REF"))
        self.assertEqual(external.read_bytes(), original)

    def test_manual_owned_regeneration_and_deleted_owned_pdf_keep_cod(self):
        first = self.generate_manual()
        scan_inventory(self.root, persist=True)
        regenerated = drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertFalse(regenerated["requires_confirmation"])
        self.assertEqual(regenerated["certificate_no"], first["certificate_no"])
        self.issued_pdf(regenerated).unlink()
        scan_inventory(self.root, persist=True)
        drafts.reconcile_manual(self.root)
        recreated = drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertFalse(recreated["requires_confirmation"])
        self.assertEqual(recreated["certificate_no"], first["certificate_no"])

    def test_owned_regeneration_is_silent_for_deleted_foreign_reference(self):
        first = self.generate_manual()
        foreign = external_pdf(self.folder / "0319_Foreign_Client_2.pdf", cod="0319", reference="MANUAL-REF")
        scan_inventory(self.root, persist=True)
        foreign.unlink()
        response = drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertFalse(response["requires_confirmation"])
        self.assertEqual((response["created"], response["certificate_no"]), (1, first["certificate_no"]))
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual(record["details"]["reference"], "MANUAL-REF")
        self.assertFalse(foreign.exists())

    def test_owned_manual_recreation_with_foreign_same_cod_cannot_change_issued_cod(self):
        first = self.generate_manual()
        self.issued_pdf(first).unlink()
        drafts.reconcile_manual(self.root)
        external = external_pdf(self.folder / f"{first['certificate_no']}_External_Client_2.pdf",
                                cod=first["certificate_no"], reference="EXT-REF")
        original = external.read_bytes()
        for confirmed in (False, True):
            with self.assertRaisesRegex(ValueError, 'already used by another PDF'):
                drafts.generate_draft(self.root, record_id=first['record_id'], confirmed=confirmed, as_of=AS_OF)
        self.assertEqual(drafts.read_register(self.root)["records"][0]["details"]["cod_no"], first["certificate_no"])
        self.assertEqual(external.read_bytes(), original)

    def test_owned_workbook_regeneration_keeps_cod_without_duplicate_warning(self):
        register(self.paths[0], case_id="NEW-EXCEL")
        created, _, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual(created, 1)
        first = read_batches(self.paths[0])[0]
        register(self.paths[0], case_id="NEW-EXCEL", cod=first["cod_no"])
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual((created, review), (1, 0))
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], first["cod_no"])

    def test_replaced_issued_manual_pdf_before_preflight_is_foreign_and_never_archived(self):
        first = self.generate_manual()
        replaced = self.issued_pdf(first)
        inventory = scan_inventory(self.root)
        self.assertEqual(issuance_paths(inventory, first["certificate_no"], "manual", "plant1", first["record_id"]),
                         [replaced.relative_to(self.root).as_posix()])
        external_pdf(replaced, cod=first["certificate_no"], reference="EXT-EDIT")
        replacement = replaced.read_bytes()
        inventory = scan_inventory(self.root)
        self.assertEqual(issuance_paths(inventory, first["certificate_no"], "manual", "plant1", first["record_id"]), [])
        register_path = self.root / "Data" / "Custom Certificates.json"
        original_register = register_path.read_bytes()
        for confirmed in (False, True):
            with self.assertRaisesRegex(ValueError, 'already used by another PDF'):
                drafts.generate_draft(self.root, record_id=first['record_id'], confirmed=confirmed, as_of=AS_OF)
        self.assertEqual(register_path.read_bytes(), original_register)
        self.assertEqual(replaced.read_bytes(), replacement)
        self.assertTrue(replaced.is_file(), "An externally replaced issued PDF must not be archived as an owned revision")
        self.assertEqual(replaced.read_bytes(), replacement)
        self.assertFalse(list((self.root / "Data" / "Archive" / "Superseded Certificates").glob("*.pdf")))
        self.assertFalse(list(self.folder.glob(".*")))

    def test_generated_workbook_refresh_does_not_retrust_replaced_pdf_before_regeneration(self):
        register(self.paths[0], case_id="NEW-EXCEL")
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual((created, review), (1, 0))
        first = read_batches(self.paths[0])[0]
        self.assertEqual(first["issued"], "Generated")
        replaced = self.folder / certificate_filename(first, 1, first["cod_no"])
        self.assertTrue(issuance_paths(scan_inventory(self.root), first["cod_no"], "workbook", "plant1", "NEW-EXCEL"))
        external_pdf(replaced, cod=first["cod_no"], reference="EXT-EDIT")
        replacement = replaced.read_bytes()
        run(self.paths, self.root, as_of=AS_OF, check_only=True, progress=lambda _: None)
        inventory = scan_inventory(self.root)
        self.assertEqual(issuance_paths(inventory, first["cod_no"], "workbook", "plant1", "NEW-EXCEL"), [],
                         "Generated refresh must not restore ownership dropped by a content change")
        self.assertEqual(replaced.read_bytes(), replacement)
        register(self.paths[0], case_id="NEW-EXCEL", cod=first["cod_no"], status="Ready")
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual((created, review), (0, 1))
        unchanged = read_batches(self.paths[0])[0]
        self.assertEqual((unchanged["cod_no"], unchanged["issued"]), (first["cod_no"], "Ready"))
        self.assertEqual(replaced.read_bytes(), replacement)
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, allow_warnings=True, progress=lambda _: None)
        self.assertEqual(created, 1)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0017")
        self.assertEqual(replaced.read_bytes(), replacement)
        self.assertFalse(list((self.root / "Data" / "Archive" / "Superseded Certificates").glob("*.pdf")))
        self.assertFalse(list(self.folder.glob(".*")))

    def test_deleted_external_number_is_reserved_for_new_manual_certificates(self):
        external = external_pdf(self.folder / "external-document.pdf", cod="0319")
        scan_inventory(self.root, persist=True)
        external.unlink()
        response = self.generate_manual()
        self.assertEqual(response["certificate_no"], "0320")

    def test_corrupt_inventory_blocks_manual_and_workbook_generation(self):
        self.ledger.parent.mkdir(parents=True, exist_ok=True)
        self.ledger.write_text("{broken", encoding="utf-8")
        register(self.paths[0], case_id="NEW-EXCEL")
        with self.assertRaisesRegex(ValueError, "inventory"):
            self.generate_manual()
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual(created, 0)
        self.assertGreater(review, 0)
        self.assertEqual(list(self.folder.iterdir()), [])
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "")
        self.assertEqual(self.ledger.read_text(encoding="utf-8"), "{broken")

    def test_archive_uuid_prefix_reserves_original_cod_without_inflating_sequence(self):
        archive = self.root / "Data" / "Archive" / "Superseded Certificates"
        external_pdf(archive / f"{'9' * 32}_0031_Former_Client_1.pdf", cod="0031")
        external_pdf(self.root / "QA" / "9999_Probe_1.pdf", cod="9999")
        response = self.generate_manual()
        self.assertEqual(response["certificate_no"], "0032")

    def test_external_cod_arriving_during_manual_render_blocks_publication_even_when_confirmed(self):
        external = self.folder / "0016_Independent_Client_Ltd_Pune_1.pdf"
        renderer = drafts.make_template_certificate
        captured = []

        def render_then_add_external(*args, **kwargs):
            result = renderer(*args, **kwargs)
            external_pdf(external, cod="0016", reference="EXT-RACE")
            captured.append(external.read_bytes())
            return result

        before = [path.read_bytes() for path in self.paths]
        with patch("certificate_drafts.make_template_certificate", side_effect=render_then_add_external):
            with self.assertRaisesRegex(ValueError, "Nothing was overwritten"):
                self.generate_manual(confirmed=True)
        self.assertEqual(external.read_bytes(), captured[0])
        self.assertEqual(list(self.folder.iterdir()), [external])
        self.assertFalse((self.root / "Data" / "Custom Certificates.json").exists())
        self.assertEqual([path.read_bytes() for path in self.paths], before)

    def test_external_cod_arriving_during_workbook_render_keeps_pdf_and_unissued_cells(self):
        register(self.paths[0], case_id="NEW-EXCEL")
        batch = read_batches(self.paths[0])[0]
        external = self.folder / certificate_filename(batch, 1, "0016")
        renderer = runner.make_certificate
        captured = []

        def render_then_add_external(*args, **kwargs):
            result = renderer(*args, **kwargs)
            external_pdf(external, cod="0016", reference="EXT-RACE")
            captured.append(external.read_bytes())
            return result

        with patch("run_certificates.make_certificate", side_effect=render_then_add_external):
            created, review, _ = run(self.paths, self.root, as_of=AS_OF, allow_warnings=True, progress=lambda _: None)
        self.assertEqual((created, review), (0, 1))
        updated = read_batches(self.paths[0])[0]
        self.assertEqual((updated["cod_no"], updated["cod_date"], updated["issued"]), ("", None, "Ready"))
        self.assertEqual(external.read_bytes(), captured[0])
        self.assertEqual(list(self.folder.iterdir()), [external])
        self.assertTrue(any("Nothing was overwritten" in row["detail"] for row in self.results()["rows"]))

    def test_owned_pdf_changed_during_manual_render_is_not_overwritten(self):
        first = self.generate_manual()
        existing = self.issued_pdf(first)
        register_path = self.root / "Data" / "Custom Certificates.json"
        original_register = register_path.read_bytes()
        renderer = drafts.make_template_certificate
        captured = []

        def render_then_modify_owned_pdf(*args, **kwargs):
            result = renderer(*args, **kwargs)
            external_pdf(existing, cod=first["certificate_no"], reference="EXT-EDIT")
            captured.append(existing.read_bytes())
            return result

        with patch("certificate_drafts.make_template_certificate", side_effect=render_then_modify_owned_pdf):
            with self.assertRaisesRegex(ValueError, "existing certificate changed"):
                drafts.generate_draft(self.root, record_id=first["record_id"], confirmed=True, as_of=AS_OF)
        self.assertEqual(register_path.read_bytes(), original_register)
        self.assertEqual(existing.read_bytes(), captured[0])
        self.assertEqual(list(self.folder.rglob("*.pdf")), [existing])

    def test_new_duplicate_reference_during_render_needs_no_confirmation(self):
        external = self.folder / "0319_External_Client_2.pdf"
        renderer = drafts.make_template_certificate
        captured = []

        def render_then_add_reference(*args, **kwargs):
            result = renderer(*args, **kwargs)
            external_pdf(external, cod="0319", reference="MANUAL-REF")
            captured.append(external.read_bytes())
            return result

        with patch("certificate_drafts.make_template_certificate", side_effect=render_then_add_reference):
            response = self.generate_manual()
        self.assertEqual(response['created'], 1)
        self.assertEqual(external.read_bytes(), captured[0])
        self.assertTrue(self.issued_pdf(response).is_file())
        self.assertEqual(drafts.read_register(self.root)['records'][0]['details']['reference'], 'MANUAL-REF')

    def test_confirmed_reference_race_keeps_requested_cod_and_reference(self):
        external = self.folder / "0319_External_Client_2.pdf"
        renderer = drafts.make_template_certificate
        captured = []

        def render_then_add_reference(*args, **kwargs):
            result = renderer(*args, **kwargs)
            external_pdf(external, cod="0319", reference="MANUAL-REF")
            captured.append(external.read_bytes())
            return result

        with patch("certificate_drafts.make_template_certificate", side_effect=render_then_add_reference):
            response = self.generate_manual(confirmed=True)
        self.assertEqual((response["created"], response["certificate_no"]), (1, "0016"))
        self.assertEqual(drafts.read_register(self.root)["records"][0]["details"]["reference"], "MANUAL-REF")
        self.assertEqual(external.read_bytes(), captured[0])
        self.assertTrue(self.issued_pdf(response).is_file())
        self.assertEqual(sorted(path.name for path in self.folder.rglob("*.pdf")),
                         sorted([external.name, response["filename"]]))
        self.assertFalse(list(self.folder.glob(".*")))


if __name__ == "__main__":
    unittest.main()
