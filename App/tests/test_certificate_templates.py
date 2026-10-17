"""Native template, metadata and recovery checks using disposable files only."""
from __future__ import annotations

from collections import Counter
from datetime import date
from hashlib import sha256
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch
import warnings

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader, PdfWriter

from certificate_core import parse_date, read_batches
import certificate_drafts as drafts
import web_app
from certificate_templates import (
    _original_region, geometry, layout_warnings, make_template_certificate, regions, template_specs,
    validate_fields,
)
from run_certificates import run
from workbook_store import WorkbookSession, upsert_certificate_details
from workflow_support import Report, save_json


APP = Path(__file__).resolve().parents[1]
AS_OF = date(2026, 10, 1)
HEADERS = [
    "SN", "Inward Date", "Supplier Name", "Location", "Item",
    "Quantity in Pcs", "Per Pcs Weight in (KG)", "Weight (KG)",
    "Form 6 No", "Reference No.", "COD No", "COD Date", "Certificate Issued",
]


def write_register(path, *, case_id="CLIENT", cod="", status="Ready", continuation=False):
    book = Workbook()
    try:
        sheet = book.active
        sheet.title = "Inward Register"
        sheet.append(HEADERS)
        sheet.append([case_id, date(2026, 9, 1), "Independent Client Ltd", "Pune",
                      "Servers", 2, 5, 10, "FORM-1", "REF-1", cod,
                      AS_OF if cod else None, status])
        if continuation:
            sheet.append([None, None, None, None, "Monitors", 3, 4, 12])
        book.save(path)
    finally:
        book.close()


def pdf_text(path):
    return "\n".join(page.extract_text() or "" for page in PdfReader(path).pages)


class TemporaryCertificateTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.backups = self.root / "Data" / "Workbook Backups"

    def manual_pdf(self, result):
        return self.root / "Certificates" / result.get("relative_path", result["filename"])

    def batch(self, kind="standard", **changes):
        batch = {
            "case_id": "CLIENT", "supplier": "Independent Client Ltd", "location": "Pune",
            "inward_date": date(2026, 9, 1), "certificate_weight": 8.5,
            "display_weight": 8.5, "total_weight": 10, "reference": "REF-1", "form6": "FORM-1",
            "certificate_type": kind, "certificate_fields": {},
        }
        batch.update(changes)
        return batch


class NativeTemplateTests(TemporaryCertificateTest):
    def test_full_location_is_printed_for_all_certificate_types(self):
        for kind in ('standard', 'sita', 'siemens'):
            with self.subTest(kind=kind):
                output = self.root / f'{kind}-full-location.pdf'
                make_template_certificate(self.batch(kind, location='Pune, Maharashtra, India'),
                                          '0001', AS_OF, output, allow_warnings=True)
                self.assertIn('PUNE, MAHARASHTRA, INDIA', ' '.join(pdf_text(output).split()))

    def test_native_title_imports_only_referenced_resources_without_unused_transparency(self):
        for kind in ("standard", "sita", "siemens"):
            with self.subTest(kind=kind):
                source = PdfReader(APP / "templates" / "certificate-types" / f"{kind}.pdf").pages[0]
                title = next(region for region in regions(kind) if region["key"] == "certificate_title")
                original_names = {category: set(source["/Resources"].get(category, {}))
                                  for category in ("/XObject", "/Font", "/ExtGState")}
                layer = _original_region(source, title, geometry(kind)["height"])
                operations = layer.get_contents().operations
                self.assertIn("RECYCLING", layer.extract_text())
                for category, operator in (("/XObject", b"Do"), ("/Font", b"Tf"), ("/ExtGState", b"gs")):
                    referenced = {args[0] for args, op in operations if op == operator}
                    actual = set(layer["/Resources"].get(category, {}))
                    self.assertEqual(actual, referenced, f"{kind} imported unused or missing {category} resources")
                    for name in referenced:
                        self.assertEqual(layer["/Resources"][category].raw_get(name),
                                         source["/Resources"][category].raw_get(name))
                    self.assertEqual(set(source["/Resources"].get(category, {})), original_names[category])
                # The native pages contain artwork masks, but their title runs
                # do not use them. Importing these unused states hid Sita's logo.
                self.assertTrue(original_names["/ExtGState"])
                self.assertFalse(any(op == b"gs" for _, op in operations))
                self.assertFalse(layer["/Resources"].get("/ExtGState"))

    def test_native_background_image_bytes_survive_title_merges_for_every_design(self):
        def image_fingerprints(page):
            images = Counter()
            for reference in page["/Resources"].get("/XObject", {}).values():
                image = reference.get_object()
                if image.get("/Subtype") == "/Image":
                    identity = (image.get("/Width"), image.get("/Height"), image.get("/BitsPerComponent"),
                                sha256(image.get_data()).hexdigest())
                    images[identity] += 1
            return images

        for kind in ("standard", "sita", "siemens"):
            with self.subTest(kind=kind):
                background = PdfReader(APP / "templates" / "certificate-types" / f"{kind}.background.pdf").pages[0]
                native_images = image_fingerprints(background)
                self.assertTrue(native_images)
                destination = self.root / f"artwork-{kind}.pdf"
                with warnings.catch_warnings():
                    warnings.filterwarnings("error", message=r"Calling `PageObject\.replace_contents\(\)`", category=DeprecationWarning)
                    make_template_certificate(self.batch(kind), "ARTWORK-1", AS_OF, destination)
                output = PdfReader(destination).pages[0]
                self.assertFalse(native_images - image_fingerprints(output),
                                 f"{kind} lost or changed a native background image")

    def test_specs_expose_four_explicit_designs_and_siemens_serial_fields(self):
        specs = {spec["id"]: spec for spec in template_specs()}
        self.assertEqual(set(specs), {"standard", "sita", "siemens", "destruction"})
        keys = {field["key"] for field in specs["siemens"]["fields"]}
        self.assertTrue({"model_name", "model_serial", "tube_serial", "site_location"} <= keys)
        for spec in specs.values():
            keys = [field["key"] for field in spec["fields"]]
            self.assertEqual(len(keys), len(set(keys)))
            expected = ({"destruction_date", "device_receipts"} if spec["id"] == "destruction" else {"company_display", "weight_display"})
            if spec["id"] == "siemens":
                expected |= {"model_name", "model_serial", "tube_serial", "site_location"}
            self.assertEqual(set(keys), expected)
            self.assertNotIn("layout_defaults", spec)

    def test_all_native_designs_render_custom_data_without_company_inference(self):
        for kind in ("standard", "sita", "siemens"):
            with self.subTest(kind=kind):
                destination = self.root / f"{kind}.pdf"
                fields = {"model_name": "Scanner Alpha", "model_serial": "MODEL-123",
                          "tube_serial": "TUBE-456", "site_location": "Pune Hospital"} if kind == "siemens" else {}
                make_template_certificate(self.batch(kind, certificate_fields=fields),
                                          "COD-0199", AS_OF, destination)
                output = PdfReader(destination)
                source = PdfReader(APP / "templates" / "certificate-types" / f"{kind}.pdf")
                self.assertEqual(len(output.pages), 1)
                self.assertEqual(tuple(output.pages[0].mediabox), tuple(source.pages[0].mediabox))
                self.assertEqual(output.metadata.subject, kind)
                extracted = pdf_text(destination)
                for expected in ("INDEPENDENT CLIENT LTD, PUNE", "COD-0199", "01.10.2026", "8.5 KG"):
                    self.assertIn(expected, extracted)
                if kind == "siemens":
                    for expected in fields.values():
                        self.assertIn(expected, extracted)
                    self.assertNotIn("SIEMENS HEALTHCARE PRIVATE LIMITED", extracted)
                elif kind == "sita":
                    self.assertNotIn("SITA INFORMATION NETWORKING COMPUTING", extracted)
        destination = self.root / "explicit-standard.pdf"
        make_template_certificate(self.batch(supplier="Siemens Alternate Ltd"), "COD-0200", AS_OF, destination)
        self.assertEqual(PdfReader(destination).metadata.subject, "standard")
        self.assertIn("SIEMENS ALTERNATE LTD, PUNE", pdf_text(destination))

    def test_custom_company_weight_and_serials_are_used_but_wording_is_fixed(self):
        batch = self.batch("siemens", certificate_fields={
            "company_display": "OTHER MEDICAL COMPANY",
            "intro": "Custom equipment received for recycling",
            "weight_display": "8.50 KG NET",
            "model_name": "Scanner Beta", "site_location": "Surat Hospital",
            "model_serial": "NEW-MODEL-SERIAL", "tube_serial": "NEW-TUBE-SERIAL",
        })
        destination = self.root / "custom.pdf"
        make_template_certificate(batch, "CUSTOM-77", AS_OF, destination)
        extracted = pdf_text(destination)
        for key, value in batch["certificate_fields"].items():
            if key != "intro":
                self.assertIn(value, extracted)
        # Check searchable stale content as well as the replacement values.
        # This relies on the renderer filtering original runs before clipping.
        self.assertNotIn("SIEMENS HEALTHCARE PRIVATE LIMITED", extracted)
        self.assertNotIn(batch["certificate_fields"]["intro"], extracted)
        self.assertIn("This is to certify that", extracted)

    def test_reserved_wording_and_style_fields_are_dropped_for_every_design(self):
        for kind in ("standard", "sita", "siemens"):
            with self.subTest(kind=kind):
                allowed = {"company_display": "OTHER COMPANY", "weight_display": "8.50 KG NET",
                           "_annexure": "a" * 32}
                if kind == "siemens":
                    allowed.update(model_name="Scanner Beta", model_serial="MODEL-NEW",
                                   tube_serial="TUBE-NEW", site_location="Surat Hospital")
                reserved = {region["key"]: {"unsupported": "wording must stay fixed"}
                            for region in regions(kind)
                            if region["group"] == "wording" and region["key"] != "company_display"}
                reserved.update(_layout={"company_display": {"left": -100, "font_size": 2000},
                                         "table": {"top": -100, "scale": 2000}},
                                _style={"font": "Courier", "color": "red"}, font_size=2000,
                                disposal_date="01.01.2099")
                self.assertEqual(validate_fields(kind, dict(allowed, **reserved)), allowed)

    def test_unknown_design_fails_before_output_is_created(self):
        destination = self.root / "unknown.pdf"
        with self.assertRaises(ValueError):
            make_template_certificate(self.batch("future-template"), "COD-1", AS_OF, destination)
        self.assertFalse(destination.exists())

    def test_long_field_warns_and_requires_explicit_acceptance(self):
        batch = self.batch(certificate_fields={"company_display": "Very long company name " * 100})
        warnings = layout_warnings(batch, "COD-1", AS_OF)
        self.assertTrue(warnings)
        destination = self.root / "long.pdf"
        with self.assertRaises(ValueError):
            make_template_certificate(batch, "COD-1", AS_OF, destination)
        self.assertFalse(destination.exists())
        make_template_certificate(batch, "COD-1", AS_OF, destination, allow_warnings=True)
        self.assertEqual(len(PdfReader(destination).pages), 1)

    def test_supplied_wording_and_layout_do_not_change_native_text_positions_or_fonts(self):
        def text_positions(path):
            collected = []
            def capture(text, cm, tm, _font, size):
                if text.strip():
                    collected.append((text, tuple(cm), tuple(tm), size))
            PdfReader(path).pages[0].extract_text(visitor_text=capture)
            return collected

        for kind in ("standard", "sita", "siemens"):
            with self.subTest(kind=kind):
                fixed = self.batch(kind)
                fields = {region["key"]: "REPLACED FIXED WORDING"
                          for region in regions(kind)
                          if region["group"] == "wording" and region["key"] != "company_display"}
                fields["_layout"] = {"company_display": {"left": -100, "font_size": 100},
                                     "table": {"top": -100, "scale": 2},
                                     "model_serial": {"left": -100}}
                supplied = self.batch(kind, certificate_fields=fields)
                self.assertEqual(layout_warnings(supplied, "COD-1", AS_OF),
                                 layout_warnings(fixed, "COD-1", AS_OF))
                original = self.root / f"fixed-{kind}.pdf"
                attempted = self.root / f"attempted-style-{kind}.pdf"
                make_template_certificate(fixed, "COD-1", AS_OF, original)
                make_template_certificate(supplied, "COD-1", AS_OF, attempted)
                self.assertEqual(text_positions(attempted), text_positions(original))
                self.assertNotIn("REPLACED FIXED WORDING", pdf_text(attempted))

    def test_standard_has_five_information_rows_and_ignores_legacy_disposal_data(self):
        asset = APP / "templates" / "certificate-types" / "standard.geometry.json"
        original = asset.read_bytes()
        self.assertEqual({cell["row"] for cell in geometry("standard")["cells"]}, {1, 2, 3, 4, 5})
        table_regions = [region for region in regions("standard") if region.get("table")]
        self.assertEqual({region["key"] for region in table_regions if region["column"] == 1},
                         {"label_1", "label_2", "label_3", "label_4", "label_5"})
        self.assertNotIn("disposal_date", {region["key"] for region in table_regions})
        destination = self.root / "standard-five-rows.pdf"
        make_template_certificate(self.batch(disposal_date="31.12.2099",
                                             certificate_fields={"disposal_date": "31.12.2099"}),
                                  "COD-1", AS_OF, destination)
        extracted = " ".join(pdf_text(destination).split()).casefold()
        self.assertNotIn("date of e-waste disposal", extracted)
        self.assertNotIn("31.12.2099", extracted)
        self.assertIn("form-1", extracted)
        self.assertEqual(asset.read_bytes(), original)

    def test_annexure_appends_native_pages_in_order(self):
        token = "a" * 32
        attachment = self.root / "Data" / "Attachments" / f"{token}.pdf"
        attachment.parent.mkdir(parents=True)
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        writer.add_blank_page(width=500, height=600)
        writer.write(attachment)
        batch = self.batch("siemens", certificate_fields={"_annexure": token}, _data_root=self.root)
        destination = self.root / "annexure.pdf"
        make_template_certificate(batch, "COD-1", AS_OF, destination)
        output = PdfReader(destination)
        self.assertEqual(len(output.pages), 3)
        self.assertIn("INDEPENDENT CLIENT LTD", output.pages[0].extract_text())
        self.assertEqual(tuple(output.pages[1].mediabox), (0, 0, 300, 400))
        self.assertEqual(tuple(output.pages[2].mediabox), (0, 0, 500, 600))

    def test_missing_annexure_does_not_publish_partial_certificate(self):
        batch = self.batch(certificate_fields={"_annexure": "b" * 32}, _data_root=self.root)
        destination = self.root / "missing.pdf"
        with self.assertRaisesRegex(ValueError, "annexure is missing"):
            make_template_certificate(batch, "COD-1", AS_OF, destination)
        self.assertFalse(destination.exists())


class WorkbookTemplateTests(TemporaryCertificateTest):
    def test_upsert_round_trips_json_metadata_and_preserves_multiple_items(self):
        path = self.root / "plant.xlsx"
        write_register(path, cod="CUSTOM-77", continuation=True)
        original = path.read_bytes()
        before = read_batches(path)[0]
        fields = {"model_serial": "SERIAL-77", "note": 'Unicode \u0915 & "quoted"\nline',
                  "_annexure": "a" * 32}
        updated = upsert_certificate_details(path, self.backups, {
            "case_id": "CLIENT", "certificate_type": "siemens", "certificate_fields": fields,
        })
        self.assertEqual((updated["certificate_type"], updated["certificate_fields"]), ("siemens", fields))
        self.assertEqual(updated["items"], before["items"])
        self.assertEqual(updated["total_weight"], 22)
        self.assertEqual(updated["cod_no"], "CUSTOM-77")
        self.assertEqual(next(self.backups.glob("*.xlsx")).read_bytes(), original)
        updated = upsert_certificate_details(path, self.backups, {"case_id": "CLIENT", "form6": "UPDATED"})
        self.assertEqual(updated["certificate_fields"], fields)
        book = load_workbook(path)
        try:
            headers = [cell.value for cell in book.active[1]]
            self.assertNotIn("Certificate Type", headers)
            self.assertNotIn("Certificate Fields", headers)
            from workbook_metadata import read_details, detail_key
            saved = read_details(path)[detail_key(book.active.title, 'CLIENT')]
            self.assertEqual(saved['certificate_fields'], fields)
        finally:
            book.close()

    def test_commit_persists_metadata_and_custom_cod_with_verification(self):
        path = self.root / "plant.xlsx"
        write_register(path, continuation=True)
        original = path.read_bytes()
        batch = read_batches(path)[0]
        fields = {"model_serial": "S-88", "weight_display": "22 KG NET"}
        with WorkbookSession(path, self.backups) as session:
            session.commit(batch, "CUSTOM-88", AS_OF,
                           overrides={"certificate_type": "siemens", "certificate_fields": fields})
        updated = read_batches(path)[0]
        self.assertEqual((updated["cod_no"], updated["issued"]), ("CUSTOM-88", "Generated"))
        self.assertEqual((updated["certificate_type"], updated["certificate_fields"]), ("siemens", fields))
        self.assertEqual(updated["items"], batch["items"])
        self.assertEqual(next(self.backups.glob("*.xlsx")).read_bytes(), original)

    def test_blank_type_defaults_and_unknown_type_is_preserved(self):
        path = self.root / "plant.xlsx"
        write_register(path)
        batch = upsert_certificate_details(path, self.backups, {
            "case_id": "CLIENT", "certificate_type": "future-template", "certificate_fields": {"key": "value"},
        })
        self.assertEqual(batch["certificate_type"], "future-template")
        batch = upsert_certificate_details(path, self.backups, {
            "case_id": "CLIENT", "certificate_type": "", "certificate_fields": {},
        })
        self.assertEqual((batch["certificate_type"], batch["certificate_fields"]), ("standard", {}))

    def test_invalid_fields_reject_write_and_keep_original_workbook(self):
        path = self.root / "plant.xlsx"
        write_register(path)
        original = path.read_bytes()
        with self.assertRaises(ValueError):
            upsert_certificate_details(path, self.backups, {"case_id": "CLIENT", "certificate_fields": "[]"})
        self.assertEqual(path.read_bytes(), original)
        self.assertFalse(self.backups.exists())

    def test_generated_to_ready_recovery_keeps_cod_date_and_backup(self):
        path = self.root / "plant.xlsx"
        write_register(path, cod="CUSTOM-88", status="Generated", continuation=True)
        original = path.read_bytes()
        batch = read_batches(path)[0]
        with WorkbookSession(path, self.backups) as session:
            backup = session.set_issued_status("CLIENT", "Inward Register", 2, "CUSTOM-88", "Generated", "Ready")
        updated = read_batches(path)[0]
        self.assertEqual(updated["issued"], "Ready")
        self.assertEqual(updated["cod_no"], batch["cod_no"])
        self.assertEqual(parse_date(updated["cod_date"]), parse_date(batch["cod_date"]))
        self.assertEqual(updated["items"], batch["items"])
        self.assertEqual(Path(backup).read_bytes(), original)

    def test_recovery_refuses_stale_cod_or_status_without_changes(self):
        path = self.root / "plant.xlsx"
        write_register(path, cod="CUSTOM-88", status="Generated")
        original = path.read_bytes()
        for cod, status in (("WRONG-COD", "Generated"), ("CUSTOM-88", "Ready")):
            with self.subTest(cod=cod, status=status):
                with WorkbookSession(path, self.backups) as session:
                    with self.assertRaises(ValueError):
                        session.set_issued_status("CLIENT", "Inward Register", 2, cod, status, "Ready")
                self.assertEqual(path.read_bytes(), original)
        self.assertFalse(self.backups.exists())

    def test_recovery_refuses_formula_status_without_overwriting_workbook(self):
        path = self.root / "plant.xlsx"
        write_register(path, cod="CUSTOM-88", status="Generated")
        book = load_workbook(path)
        try:
            book.active["M2"] = '= "Generated"'
            book.save(path)
        finally:
            book.close()
        original = path.read_bytes()
        # Supply the cached status expected from Excel so the test reaches the
        # formula safeguard rather than failing the preceding status check.
        cached = read_batches(path)[0]
        cached["issued"] = "Generated"
        with patch("workbook_store.read_batches", return_value=[cached]):
            with WorkbookSession(path, self.backups) as session:
                with self.assertRaisesRegex(ValueError, "formula or protected"):
                    session.set_issued_status("CLIENT", "Inward Register", 2,
                                              "CUSTOM-88", "Generated", "Ready")
        self.assertEqual(path.read_bytes(), original)
        self.assertFalse(self.backups.exists())

    def test_metadata_readback_mismatch_keeps_original_before_persistence(self):
        path = self.root / "plant.xlsx"
        write_register(path)
        original = path.read_bytes()
        real_reader = read_batches

        def mismatched_reader(source):
            batches = real_reader(source)
            for batch in batches:
                batch["certificate_fields"] = {"model_serial": "CORRUPTED"}
            return batches

        with patch("workbook_store.read_batches", side_effect=mismatched_reader):
            with self.assertRaisesRegex(ValueError, "certificate_fields could not be verified"):
                upsert_certificate_details(path, self.backups, {
                    "case_id": "CLIENT", "certificate_type": "siemens",
                    "certificate_fields": {"model_serial": "EXPECTED"},
                })
        self.assertEqual(path.read_bytes(), original)
        self.assertFalse(self.backups.exists())

    def test_report_includes_workbook_template_metadata(self):
        report = Report(self.root, AS_OF, "test")
        report.add(self.batch("future-template", certificate_fields={"model_serial": "S-88"}),
                   "needs_review", "Unknown template")
        self.assertEqual(report.rows[0]["source_kind"], "workbook")
        self.assertEqual(report.rows[0]["certificate_type"], "future-template")
        self.assertEqual(report.rows[0]["certificate_fields"], {"model_serial": "S-88"})


class DraftWorkflowTests(TemporaryCertificateTest):
    def setUp(self):
        super().setUp()
        self.paths = [self.root / "plant1.xlsx", self.root / "plant2.xlsx"]
        # Numbering-only markers do not claim a current certificate exists.
        write_register(self.paths[0], case_id="EXCEL-1", cod="0010", status="N/A")
        write_register(self.paths[1], case_id="EXCEL-2", cod="0015", status="N/A")
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in self.paths],
            "certificate_number_prefix": "", "certificate_number_start": 1,
            "certificate_number_width": 4, "reminder_time": "11:00", "reminders_enabled": False,
        })
        native = self.root / "App" / "templates" / "certificate-types"
        native.mkdir(parents=True)
        shutil.copy2(APP / "templates" / "certificate-types" / "standard.pdf", native / "standard.pdf")

    def details(self, kind="sita", **changes):
        details = {
            "case_id": "MANUAL-1", "supplier": "Independent Client Ltd", "location": "Pune",
            "certificate_type": kind, "certificate_fields": {}, "inward_date": "01.09.2026",
            "weight": "10", "certificate_weight": "8.5", "item": "Servers",
            "form6": "FORM-1", "reference": "REF-1",
        }
        if kind == "siemens":
            details["certificate_fields"] = {
                "model_name": "Scanner Alpha", "model_serial": "MODEL-123",
                "tube_serial": "TUBE-456", "site_location": "Pune Hospital",
            }
        details.update(changes)
        if "reference" not in changes:
            details["reference"] = f"R-{details['case_id']}"
        return details

    def generate(self, plant="plant1", details=None, **options):
        return drafts.generate_draft(self.root, plant=plant, details=self.details() if details is None else details,
                                     as_of=AS_OF, **options)

    def test_sita_missing_inward_requires_confirmation_then_generates_with_warnings(self):
        original_workbooks = [path.read_bytes() for path in self.paths]
        details = self.details(inward_date="")
        result = self.generate(details=details)
        self.assertTrue(result["requires_confirmation"])
        self.assertEqual(result["created"], 0)
        self.assertTrue(any("receipt date" in warning.casefold()
                            for row in result["warnings"] for warning in row["warnings"]))
        self.assertFalse((self.root / "Data" / "Custom Certificates.json").exists())
        self.assertFalse(list(self.root.rglob("*.tmp.pdf")))
        result = self.generate(details=details, confirmed=True)
        self.assertEqual(result["created"], 1)
        self.assertFalse(result["requires_confirmation"])
        self.assertTrue(result["warnings"])
        output = self.manual_pdf(result)
        self.assertIn("INDEPENDENT CLIENT LTD, PUNE", pdf_text(output))
        rows = drafts.manual_rows(self.root)
        self.assertEqual((len(rows), rows[0]["source_kind"], rows[0]["certificate_type"]), (1, "manual", "sita"))
        self.assertEqual(rows[0]["certificate_status"], "generated_manual")
        self.assertEqual([path.read_bytes() for path in self.paths], original_workbooks)

    def test_siemens_alternate_company_and_serials_do_not_require_inward(self):
        details = self.details("siemens", supplier="Alternate Medical Ltd", inward_date="")
        result = self.generate(plant="plant2", details=details)
        self.assertFalse(result["requires_confirmation"])
        self.assertEqual(result["created"], 1)
        output = self.manual_pdf(result)
        extracted = pdf_text(output)
        self.assertIn("ALTERNATE MEDICAL LTD, PUNE", extracted)
        for value in details["certificate_fields"].values():
            self.assertIn(value, extracted)
        self.assertNotIn("SIEMENS HEALTHCARE PRIVATE LIMITED", extracted)
        row = drafts.manual_rows(self.root)[0]
        self.assertEqual(row["certificate_fields"]["model_serial"], "MODEL-123")
        self.assertEqual(row["plant"], "plant2")

    def test_draft_edit_retains_record_id_and_assigned_cod(self):
        result = self.generate()
        record_id, cod = result["record_id"], result["certificate_no"]
        edited = self.details(supplier="Renamed Client Ltd", certificate_fields={"weight_display": "8.50 KG NET"})
        edited_id = drafts.save_draft(self.root, "plant1", edited, record_id=record_id)
        self.assertEqual(edited_id, record_id)
        rows = drafts.manual_rows(self.root)
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0]["record_id"], rows[0]["certificate_no"], rows[0]["certificate_status"]),
                         (record_id, cod, "ready"))
        regenerated = drafts.generate_draft(self.root, record_id=record_id, as_of=AS_OF)
        self.assertEqual((regenerated["record_id"], regenerated["certificate_no"]), (record_id, cod))
        self.assertIn("RENAMED CLIENT LTD, PUNE", pdf_text(self.manual_pdf(regenerated)))

    def test_shared_sequence_accounts_for_both_workbooks_manual_and_pdf_floor(self):
        first = self.generate()
        second = self.generate(plant="plant2", details=self.details(case_id="MANUAL-2"))
        self.assertEqual((first["certificate_no"], second["certificate_no"]), ("0016", "0017"))
        floor = self.root / "Certificates" / "0030_Previous_Client_2.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        writer.write(floor)
        next_result = self.generate(details=self.details(case_id="MANUAL-3"))
        self.assertEqual(next_result["certificate_no"], "0031")
        # Generation through Excel must continue the same sequence.
        write_register(self.paths[0], case_id="NEW-EXCEL")
        upsert_certificate_details(self.paths[0], self.backups, {
            "case_id": "NEW-EXCEL", "certificate_type": "standard",
            "certificate_fields": {},
        })
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual(created, 1)
        self.assertEqual(review, 0)
        self.assertEqual(read_batches(self.paths[0])[0]["cod_no"], "0032")

    def test_native_standard_without_disposal_date_generates_without_warning(self):
        write_register(self.paths[0], case_id="STANDARD-NO-DISPOSAL")
        created, review, _ = run(self.paths, self.root, as_of=AS_OF, progress=lambda _: None)
        self.assertEqual((created, review), (1, 0))
        report = json.loads((self.root / "Data" / "Reports" / "Latest Results.json").read_text(encoding="utf-8"))
        row = next(row for row in report["rows"] if row["sn"] == "STANDARD-NO-DISPOSAL")
        self.assertEqual(row["warnings"], [])
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch["issued"], batch["cod_no"]), ("Generated", "0016"))
        self.assertEqual(len(list((self.root / "Certificates").rglob("*.pdf"))), 1)

    def test_manual_standard_without_disposal_date_generates_without_warning(self):
        result = self.generate(details=self.details("standard"))
        self.assertFalse(result["requires_confirmation"])
        self.assertEqual((result["created"], result["warnings"]), (1, []))

    def test_clean_details_ignores_obsolete_top_level_and_json_disposal_date(self):
        details = self.details("standard", disposal_date="obsolete malformed date " * 1000,
                               certificate_fields={"disposal_date": "32.09.2026",
                                                   "company_display": "OTHER COMPANY"})
        cleaned = drafts.clean_details(details)
        self.assertNotIn("disposal_date", cleaned)
        self.assertEqual(cleaned["certificate_fields"], {"company_display": "OTHER COMPANY"})
        self.assertEqual(cleaned["inward_date"], details["inward_date"])

    def test_manual_rows_hide_obsolete_disposal_warnings_without_rewriting_history(self):
        record_id = drafts.save_draft(self.root, "plant1", self.details("standard"))
        data = drafts.read_register(self.root)
        record = next(record for record in data["records"] if record["id"] == record_id)
        history = ["Disposal date is missing or invalid; it will be blank.",
                   "Material receipt date needs review.", "Missing or invalid DISPOSAL DATE",
                   "Reference already appears in another certificate."]
        record["warnings"] = history
        record["details"]["disposal_date"] = "32.09.2026"
        drafts.write_register(self.root, data)
        register = self.root / "Data" / "Custom Certificates.json"
        original = register.read_bytes()
        row = drafts.manual_rows(self.root)[0]
        self.assertEqual(row["warnings"], [history[1], history[3]])
        self.assertNotIn("disposal_date", row)
        self.assertEqual(register.read_bytes(), original)
        saved = drafts.read_register(self.root)["records"][0]
        self.assertEqual(saved["warnings"], history)
        self.assertEqual(saved["details"]["disposal_date"], "32.09.2026")

    def test_manual_rows_drop_fixed_wording_and_layout_warnings_but_keep_company_and_data(self):
        record_id = drafts.save_draft(self.root, "plant1", self.details("siemens"))
        data = drafts.read_register(self.root)
        record = next(record for record in data["records"] if record["id"] == record_id)
        obsolete = [
            "Certificate title is long for its space; the PDF may use small text or clip it.",
            "Table label: COD No. is long for its space; shorten it and preview before sending.",
            "Introduction extends beyond the page; move it inside and preview before sending.",
            "Company line on certificate extends beyond the page; move it inside and preview before sending.",
            "Company line on certificate may overlap Recycling statement; check the preview.",
            "The information box may overlap Services line; check the preview.",
            "Missing or invalid disposal date; this value will be blank on the certificate",
        ]
        retained = [
            "Company line on certificate is long for its space; the PDF may use small text or clip it.",
            "Model serial number is long for its space; the PDF may use small text or clip it.",
            "Certificate/net weight is invalid; inward weight will be used if available.",
            "Material receipt date is missing or invalid; it will be blank on the certificate.",
            "Model serial number is missing; enter N/A if it does not apply.",
            "Reference REF-1 already appears in another certificate. Review it.",
        ]
        record["warnings"] = obsolete + retained
        drafts.write_register(self.root, data)
        register = self.root / "Data" / "Custom Certificates.json"
        original = register.read_bytes()
        row = drafts.manual_rows(self.root)[0]
        self.assertEqual(row["warnings"], retained)
        self.assertEqual(register.read_bytes(), original)
        self.assertEqual(drafts.read_register(self.root)["records"][0]["warnings"], obsolete + retained)

    def test_deleted_manual_pdf_preserves_sequence_and_can_recreate_same_cod(self):
        first = self.generate()
        self.manual_pdf(first).unlink()
        drafts.reconcile_manual(self.root)
        rows = drafts.manual_rows(self.root)
        self.assertEqual((rows[0]["certificate_status"], rows[0]["certificate_no"]), ("ready", "0016"))
        self.assertEqual(rows[0]["certificate_date"], AS_OF.isoformat())
        second = self.generate(plant="plant2", details=self.details(case_id="MANUAL-2"))
        self.assertEqual(second["certificate_no"], "0017")
        recreated = drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertEqual(recreated["certificate_no"], first["certificate_no"])
        self.assertEqual(recreated["record_id"], first["record_id"])

    def test_preflight_does_not_mutate_register_workbooks_or_allocate_number(self):
        record_id = drafts.save_draft(self.root, "plant1", self.details())
        before = {str(path.relative_to(self.root)): path.read_bytes()
                  for path in self.root.rglob("*") if path.is_file()}
        batch, cod, issued_date, warnings = drafts.prepare_certificate(
            self.root, "plant1", self.details(), record_id=record_id, as_of=AS_OF)
        after = {str(path.relative_to(self.root)): path.read_bytes()
                 for path in self.root.rglob("*") if path.is_file()}
        self.assertEqual(before, after)
        self.assertEqual((cod, issued_date, warnings), ("0016", AS_OF, []))
        self.assertEqual(batch["certificate_type"], "sita")
        self.assertEqual(drafts.manual_rows(self.root)[0]["certificate_no"], "")

    def test_missing_pdf_recreates_with_original_automatic_cod(self):
        first = self.generate(plant="plant2", details=self.details(cod_no="SITA-CUSTOM-77"))
        folder = self.root / "Certificates"
        self.manual_pdf(first).unlink()
        # An unrelated other-plant certificate must not count as this output.
        # Same-COD external conflicts have their own inventory regression test.
        unrelated = folder / "OTHER-COD-88_Other_Company_1.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        writer.write(unrelated)
        drafts.reconcile_manual(self.root)
        register = drafts.read_register(self.root)
        record = register["records"][0]
        self.assertEqual(record["status"], "ready")
        self.assertEqual(record["details"]["cod_no"], first["certificate_no"])
        self.assertEqual(record["details"]["cod_date"], AS_OF.isoformat())
        regenerated = drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertEqual((regenerated["record_id"], regenerated["certificate_no"]),
                         (first["record_id"], first["certificate_no"]))
        self.assertTrue(self.manual_pdf(regenerated).exists())
        self.assertTrue(unrelated.exists())

    def test_workbook_preflight_keeps_own_cod_and_manual_source_uses_automatic_cod(self):
        details = self.details(case_id="EXCEL-1", cod_no="0010")
        before = [path.read_bytes() for path in self.paths]
        _, cod, _, warnings = drafts.prepare_certificate(
            self.root, "plant1", details, as_of=AS_OF, source_kind="workbook")
        self.assertEqual((cod, warnings), ("0010", []))
        _, cod, _, warnings = drafts.prepare_certificate(
            self.root, "plant1", details, as_of=AS_OF, source_kind="manual")
        self.assertEqual(cod, "0016")
        self.assertFalse(any("belongs to another certificate" in warning for warning in warnings))
        self.assertEqual([path.read_bytes() for path in self.paths], before)
        self.assertFalse((self.root / "Data" / "Custom Certificates.json").exists())

    def test_issued_manual_certificate_date_can_be_corrected_with_cod_preserved(self):
        first = self.generate(plant="plant1")
        row = drafts.manual_rows(self.root)[0]
        details = row["raw_details"]
        details["cod_no"] = first["certificate_no"]
        details["cod_date"] = "2026-09-28"
        with patch.object(web_app, "ROOT", self.root):
            web_app.validate_manual_identifiers({"record_id": first["record_id"], "details": details})
        corrected = drafts.generate_draft(self.root, plant="plant1", details=details,
                                           record_id=first["record_id"], as_of=AS_OF)
        self.assertEqual(corrected["certificate_no"], first["certificate_no"])
        self.assertEqual(drafts.manual_rows(self.root)[0]["certificate_date"], "2026-09-28")
        self.assertIn("28.09.2026", pdf_text(self.manual_pdf(corrected)))

    def test_disconnected_source_requires_explicit_numbering_confirmation(self):
        settings_path = self.root / "settings.local.json"
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
        settings["workbook_paths"][1] = ""
        save_json(settings_path, settings)
        result = self.generate()
        self.assertTrue(result["requires_confirmation"])
        self.assertEqual(result["created"], 0)
        warnings = [warning for row in result["warnings"] for warning in row["warnings"]]
        self.assertTrue(any("Plant 2 workbook is not connected" in warning for warning in warnings))
        self.assertFalse((self.root / "Data" / "Custom Certificates.json").exists())

    def test_manual_rows_preserve_invalid_raw_weight_and_dates_for_editing(self):
        details = self.details(certificate_weight="not-a-weight", inward_date="31.02.2026",
                               cod_date="not-a-date")
        record_id = drafts.save_draft(self.root, "plant1", details)
        register = self.root / "Data" / "Custom Certificates.json"
        before = register.read_bytes()
        row = drafts.manual_rows(self.root)[0]
        self.assertIsNone(row["certificate_weight"])
        self.assertEqual(row["display_weight"], 10)
        for key in ("certificate_weight", "inward_date"):
            self.assertEqual(row["raw_details"][key], details[key])
        self.assertEqual(row["raw_details"]["cod_date"], "")
        self.assertEqual(register.read_bytes(), before)
        edited = dict(row["raw_details"], supplier="Renamed Client Ltd")
        self.assertEqual(drafts.save_draft(self.root, "plant1", edited, record_id=record_id), record_id)
        saved = drafts.read_register(self.root)["records"][0]["details"]
        for key in ("certificate_weight", "inward_date"):
            self.assertEqual(saved[key], details[key])
        self.assertEqual(saved["cod_date"], "")

    def test_corrupted_xlsx_keeps_available_numbering_source_and_warns_before_generation(self):
        corrupted = b"This is not an XLSX ZIP package."
        self.paths[1].write_bytes(corrupted)
        settings = {"workbook_paths": [str(path) for path in self.paths]}
        batches, warnings = drafts.numbering_sources(self.root, settings)
        self.assertEqual([(batch["source_id"], batch["cod_no"]) for batch in batches], [("plant1", "0010")])
        self.assertTrue(any("Plant 2 workbook could not be checked" in warning for warning in warnings))
        response = self.generate()
        self.assertTrue(response["requires_confirmation"])
        self.assertEqual(response["created"], 0)
        self.assertFalse((self.root / "Data" / "Custom Certificates.json").exists())
        response = self.generate(confirmed=True)
        self.assertEqual((response["created"], response["certificate_no"]), (1, "0011"))
        self.assertTrue(any("Plant 2 workbook could not be checked" in warning for warning in response["warnings"]))
        self.assertTrue(self.manual_pdf(response).is_file())
        self.assertEqual(self.paths[1].read_bytes(), corrupted)

    def test_unsupported_workbook_extension_keeps_other_source_and_returns_warning(self):
        unsupported = self.root / "unsupported.xls"
        shutil.copy2(self.paths[1], unsupported)
        settings = {"workbook_paths": [str(self.paths[0]), str(unsupported)]}
        batches, warnings = drafts.numbering_sources(self.root, settings)
        self.assertEqual([(batch["source_id"], batch["cod_no"]) for batch in batches], [("plant1", "0010")])
        self.assertTrue(any("Plant 2 workbook could not be checked" in warning for warning in warnings))

    def test_complete_draft_keeps_id_cod_but_returns_ready_after_pdf_deletion(self):
        before = [path.read_bytes() for path in self.paths]
        result = self.generate()
        drafts.complete_draft(self.root, result["record_id"])
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual((record["id"], record["status"], record["details"]["cod_no"]),
                         (result["record_id"], "completed", result["certificate_no"]))
        self.assertEqual(drafts.manual_rows(self.root)[0]["certificate_status"], "completed")
        self.manual_pdf(result).unlink()
        drafts.reconcile_manual(self.root)
        recovered = drafts.read_register(self.root)["records"][0]
        self.assertEqual((recovered["id"], recovered["status"], recovered["details"]["cod_no"]),
                         (result["record_id"], "ready", result["certificate_no"]))
        self.assertEqual([path.read_bytes() for path in self.paths], before)

    def test_legacy_custom_record_statuses_are_migrated_to_canonical_values(self):
        result = self.generate()
        expectations = (("Created", "generated_manual"), ("Sent", "completed"),
                        ("Check Details", "alert"), ("Working Draft", "alert"),
                        ("Generated Manual", "generated_manual"))
        for old_status, expected in expectations:
            with self.subTest(old_status=old_status):
                data = drafts.read_register(self.root)
                data["records"][0]["status"] = old_status
                drafts.write_register(self.root, data)
                drafts.reconcile_manual(self.root)
                record = drafts.read_register(self.root)["records"][0]
                self.assertEqual(record["status"], expected)
                self.assertEqual(drafts.manual_rows(self.root)[0]["certificate_status"], expected)
                self.assertTrue(drafts.current_record_pdf(self.root, record).is_file())

    def test_completion_refuses_ready_draft_or_missing_generated_pdf(self):
        record_id = drafts.save_draft(self.root, "plant1", self.details())
        register = self.root / "Data" / "Custom Certificates.json"
        original = register.read_bytes()
        with self.assertRaisesRegex(ValueError, "Create this certificate PDF"):
            drafts.complete_draft(self.root, record_id)
        self.assertEqual(register.read_bytes(), original)
        result = drafts.generate_draft(self.root, record_id=record_id, as_of=AS_OF)
        self.manual_pdf(result).unlink()
        original = register.read_bytes()
        with self.assertRaisesRegex(ValueError, "PDF is missing"):
            drafts.complete_draft(self.root, record_id)
        self.assertEqual(register.read_bytes(), original)

    def test_completion_persistence_failure_retains_generated_record_and_pdf(self):
        result = self.generate()
        output = self.manual_pdf(result)
        register = self.root / "Data" / "Custom Certificates.json"
        original_pdf, original_register = output.read_bytes(), register.read_bytes()
        with patch("certificate_drafts.write_register", side_effect=PermissionError("register blocked")):
            with self.assertRaisesRegex(PermissionError, "register blocked"):
                drafts.complete_draft(self.root, result["record_id"])
        self.assertEqual(register.read_bytes(), original_register)
        self.assertEqual(output.read_bytes(), original_pdf)
        self.assertEqual(drafts.read_register(self.root)["records"][0]["status"], "generated_manual")

    def test_register_persistence_failure_removes_new_pdf_and_keeps_draft(self):
        record_id = drafts.save_draft(self.root, "plant1", self.details())
        register = self.root / "Data" / "Custom Certificates.json"
        original = register.read_bytes()
        with patch("certificate_drafts.write_register", side_effect=PermissionError("register blocked")):
            with self.assertRaisesRegex(PermissionError, "register blocked"):
                drafts.generate_draft(self.root, record_id=record_id, as_of=AS_OF)
        self.assertEqual(register.read_bytes(), original)
        self.assertEqual(list((self.root / "Certificates").rglob("*.pdf")), [])
        retried = drafts.generate_draft(self.root, record_id=record_id, as_of=AS_OF)
        self.assertEqual((retried["record_id"], retried["certificate_no"]), (record_id, "0016"))

    def test_register_persistence_failure_restores_previous_same_cod_pdf(self):
        first = self.generate()
        output = self.manual_pdf(first)
        original_pdf = output.read_bytes()
        edited = self.details(certificate_fields={"weight_display": "8.50 KG NET"})
        drafts.save_draft(self.root, "plant1", edited, record_id=first["record_id"])
        register = self.root / "Data" / "Custom Certificates.json"
        original_register = register.read_bytes()
        with patch("certificate_drafts.write_register", side_effect=PermissionError("register blocked")):
            with self.assertRaises(PermissionError):
                drafts.generate_draft(self.root, record_id=first["record_id"], as_of=AS_OF)
        self.assertEqual(output.read_bytes(), original_pdf)
        self.assertEqual(register.read_bytes(), original_register)
        self.assertEqual(list((self.root / "Certificates").rglob("*.pdf")), [output])

    def test_renamed_revision_archives_old_pdf_and_old_copy_cannot_mask_missing_latest(self):
        first = self.generate()
        folder = self.root / "Certificates"
        old_pdf = self.manual_pdf(first)
        original = old_pdf.read_bytes()
        renamed = self.details(supplier="Renamed Client Ltd", location="Surat")
        regenerated = drafts.generate_draft(self.root, details=renamed,
                                           record_id=first["record_id"], as_of=AS_OF)
        latest = self.manual_pdf(regenerated)
        self.assertNotEqual(latest, old_pdf)
        self.assertEqual((regenerated["record_id"], regenerated["certificate_no"]),
                         (first["record_id"], first["certificate_no"]))
        self.assertEqual(list(folder.rglob("*.pdf")), [latest])
        archive = self.root / "Data" / "Archive" / "Superseded Certificates"
        archived = list(archive.rglob(first['filename']))
        self.assertEqual(len(archived), 1)
        self.assertEqual(archived[0].read_bytes(), original)
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual(record["pdf_filename"], regenerated["filename"])
        self.assertEqual(drafts.current_record_pdf(self.root, record), latest)

        # Simulate a surviving legacy copy alongside the newly recorded revision.
        # Its matching COD and plant must not satisfy the latest-file check.
        shutil.copy2(archived[0], old_pdf)
        latest.unlink()
        row = drafts.manual_rows(self.root)[0]
        self.assertEqual((row["certificate_status"], row["certificate_no"]), ("ready", first["certificate_no"]))
        self.assertTrue(row["system_alert"])
        self.assertIsNone(drafts.current_record_pdf(self.root, record))
        with self.assertRaisesRegex(ValueError, "PDF is missing"):
            drafts.complete_draft(self.root, first["record_id"])
        drafts.reconcile_manual(self.root)
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual((record["status"], record["pdf_filename"], record["details"]["cod_no"]),
                         ("ready", regenerated["filename"], first["certificate_no"]))
        self.assertEqual(old_pdf.read_bytes(), original)
        self.assertEqual(archived[0].read_bytes(), original)

    def test_renamed_revision_register_failure_restores_archived_old_pdf(self):
        first = self.generate()
        folder = self.root / "Certificates"
        old_pdf = self.manual_pdf(first)
        register = self.root / "Data" / "Custom Certificates.json"
        original_pdf, original_register = old_pdf.read_bytes(), register.read_bytes()
        archive = self.root / "Data" / "Archive" / "Superseded Certificates"

        def fail_after_archive(_root, _data):
            self.assertTrue(old_pdf.exists(), 'Retire the previous PDF only after the source is committed')
            archived = list(archive.rglob("*.pdf"))
            self.assertEqual(len(archived), 0)
            self.assertEqual(old_pdf.read_bytes(), original_pdf)
            self.assertEqual(len([p for p in folder.rglob('*.pdf') if not p.name.startswith('.')]), 2)
            raise PermissionError("register blocked after old PDF was archived")

        renamed = self.details(supplier="Renamed Client Ltd", location="Surat")
        with patch("certificate_drafts.write_register", side_effect=fail_after_archive):
            with self.assertRaisesRegex(PermissionError, "register blocked"):
                drafts.generate_draft(self.root, details=renamed,
                                      record_id=first["record_id"], as_of=AS_OF)
        self.assertEqual(register.read_bytes(), original_register)
        self.assertEqual(old_pdf.read_bytes(), original_pdf)
        self.assertEqual(list(folder.rglob("*.pdf")), [old_pdf])
        self.assertEqual(list(archive.glob('*.pdf')), [])
        record = drafts.read_register(self.root)["records"][0]
        self.assertEqual(record["status"], "generated_manual")
        self.assertEqual(drafts.current_record_pdf(self.root, record), old_pdf)

    def test_missing_workbook_pdf_resets_ready_but_wrong_plant_pdf_does_not_match(self):
        write_register(self.paths[0], cod="CUSTOM-77", status="Generated")
        folder = self.root / "Certificates"
        folder.mkdir()
        wrong_plant = folder / "CUSTOM-77_Previous_Company_2.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        writer.write(wrong_plant)
        run(self.paths, self.root, as_of=AS_OF, check_only=True, progress=lambda _: None)
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch["issued"], batch["cod_no"], parse_date(batch["cod_date"])),
                         ("Ready", "CUSTOM-77", AS_OF))
        self.assertTrue(wrong_plant.exists())
        self.assertTrue(list(self.backups.glob("*.xlsx")))

    def test_matching_workbook_pdf_survives_company_rename(self):
        write_register(self.paths[0], cod="CUSTOM-77", status="Generated")
        folder = self.root / "Certificates"
        folder.mkdir()
        existing = folder / "CUSTOM-77_Old_Name_Old_City_1.pdf"
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=400)
        writer.write(existing)
        run(self.paths, self.root, as_of=AS_OF, check_only=True, progress=lambda _: None)
        batch = read_batches(self.paths[0])[0]
        self.assertEqual((batch["issued"], batch["cod_no"]), ("Generated", "CUSTOM-77"))


class WebGenerationTests(TemporaryCertificateTest):
    def test_cached_obsolete_workbook_warnings_trigger_check_only_refresh_not_generation(self):
        obsolete = "Certificate title is long for its space; the PDF may use small text or clip it."
        legitimate = "Company line on certificate is long for its space; shorten it before sending."
        paths = [self.root / "dummy1.xlsx", self.root / "dummy2.xlsx"]
        results = self.root / "Data" / "Reports" / "Latest Results.json"
        manual_row = {"plant": "plant2", "sn": "MANUAL", "source_kind": "manual", "record_id": "manual-id"}
        refreshed_row = {"plant": "plant1", "sn": "CLIENT", "source_kind": "workbook",
                         "certificate_status": "ready", "outcome": "needs_review", "system_alert": True,
                         "warnings": [legitimate], "detail": legitimate}
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "RESULTS_PATH", results))
        self.enterContext(patch.object(web_app, "configured_paths", return_value=paths))
        self.enterContext(patch.object(web_app, "load_settings", return_value={}))
        self.enterContext(patch.object(web_app, "manual_rows", return_value=[manual_row]))
        self.enterContext(patch.object(web_app, "find_certificate", return_value=True))
        self.enterContext(patch.object(web_app, "generate", side_effect=AssertionError("workspace read must not generate")))

        def check_only_refresh(actual_paths, root, **options):
            self.assertEqual((actual_paths, root), (paths, self.root))
            self.assertTrue(options.get("check_only"), "Cached warning recovery must never issue a certificate")
            save_json(results, {"rows": [refreshed_row], "scan_complete": True, "created": 0, "review": 1})
            return 0, 1, self.root

        for stale_warnings in ([obsolete], []):
            with self.subTest(warning_embedded_only_in_detail=not stale_warnings):
                cached_row = dict(refreshed_row, warnings=stale_warnings, detail=f"Old check: {obsolete}")
                save_json(results, {"rows": [cached_row], "scan_complete": True, "created": 0, "review": 1})
                with patch.object(web_app, "run", side_effect=check_only_refresh) as check:
                    state = web_app.workspace_state()
                check.assert_called_once()
                self.assertEqual(state["rows"], [refreshed_row, manual_row])
                self.assertEqual((state["created"], state["review"]), (0, 1))
                self.assertEqual(state["rows"][0]["outcome"], "needs_review")
                self.assertTrue(state["rows"][0]["system_alert"])
                self.assertFalse(list((self.root / "Certificates").rglob("*.pdf")))

    def test_selected_generation_warning_keeps_workspace_unchanged_until_confirmation(self):
        selected = ("plant1", "SELECTED")
        warning_row = {"plant": "plant1", "sn": "SELECTED", "source_kind": "workbook",
                       "certificate_status": "ready", "outcome": "ready",
                       "warnings": ["Material receipt date is missing"]}
        other_row = {"plant": "plant2", "sn": "OTHER", "source_kind": "workbook",
                     "certificate_status": "completed", "outcome": "completed"}
        manual_row = {"plant": "plant2", "sn": "MANUAL", "source_kind": "manual", "record_id": "manual-id"}
        targeted_report = {"rows": [warning_row], "scan_complete": True}
        full_report = {"rows": [warning_row, other_row], "scan_complete": True}
        paths = [self.root / "dummy1.xlsx", self.root / "dummy2.xlsx"]
        self.enterContext(patch.object(web_app, "ROOT", self.root))
        self.enterContext(patch.object(web_app, "load_settings", return_value={"workbook_paths": [str(path) for path in paths]}))
        self.enterContext(patch.object(web_app, "configured_paths", return_value=paths))
        runner = self.enterContext(patch.object(web_app, "run", return_value=(0, 1, self.root)))
        self.enterContext(patch.object(web_app, "latest_result", side_effect=[targeted_report, full_report]))
        self.enterContext(patch.object(web_app, "manual_rows", return_value=[manual_row]))
        # This test covers whole-workspace warning composition, not PDF recovery.
        self.enterContext(patch.object(web_app, "find_certificate", return_value=True))
        refresh = self.enterContext(patch.object(web_app, "refresh", side_effect=lambda _paths: web_app.workspace_state()))
        response = web_app.generate(selected=selected)
        self.assertTrue(response["requires_confirmation"])
        self.assertEqual(response["warnings"], [warning_row])
        self.assertNotIn('state', response)
        refresh.assert_not_called()
        runner.assert_called_once_with(paths, self.root, check_only=True, read_only_preflight=True,
                                       allow_warnings=True, progress=None, report_sink={}, selected=selected)
        self.assertEqual(list(self.root.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
