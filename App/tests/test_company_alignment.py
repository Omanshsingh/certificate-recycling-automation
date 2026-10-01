"""Measure client-name placement in disposable native-template PDFs."""
from __future__ import annotations

from copy import deepcopy
from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfReader
from reportlab.pdfbase import pdfmetrics

import certificate_templates as templates


KINDS = ("standard", "sita", "siemens")
ISSUED = date(2026, 10, 1)
COD = "CUSTOM-077"
LONG_COMPANY = (
    "INDEPENDENT INDUSTRIAL TECHNOLOGY AND ENVIRONMENTAL EQUIPMENT RECYCLING "
    "SERVICES INDIA PRIVATE LIMITED, BENGALURU"
)
MULTILINE_COMPANY = (
    "GLOBAL ENVIRONMENTAL EQUIPMENT RECYCLING INDIA PRIVATE LIMITED\n"
    "INDUSTRIAL TECHNOLOGY SERVICES DIVISION, BENGALURU"
)


def drawn_text(page):
    """Capture baselines before Tj; visitor_text sees ReportLab's later T*."""
    records = []
    font_size = 0

    def operand(operator, args, cm, tm):
        nonlocal font_size
        if operator == b"Tf":
            font_size = float(args[1])
        elif operator == b"Tj":
            raw = args[0]
            text = raw.decode("latin-1") if isinstance(raw, bytes) else str(raw)
            if not text.strip():
                return
            x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
            y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
            records.append({"text": text, "x": x, "y": y, "size": font_size,
                            "transform": tuple(cm[:4])})

    page.extract_text(visitor_operand_before=operand)
    return records


class CompanyAlignmentTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.counter = 0

    def batch(self, kind, company):
        fields = {"company_display": company}
        if kind == "siemens":
            fields.update(model_name="Scanner Alpha", model_serial="MODEL-123",
                          tube_serial="TUBE-456", site_location="Pune Hospital")
        return {"case_id": "CLIENT-077", "supplier": "Independent Client Ltd", "location": "Pune",
                "inward_date": date(2026, 9, 1), "certificate_weight": 8.5, "display_weight": 8.5,
                "total_weight": 10, "reference": "REF-077", "form6": "FORM-077",
                "cod_no": COD, "cod_date": ISSUED, "issued": "Generated",
                "certificate_type": kind, "certificate_fields": fields}

    def company_region(self, batch):
        values = templates.display_values(batch, COD, ISSUED)
        region = next(region for region in templates.positioned_regions(batch["certificate_type"], values)
                      if region["key"] == "company_display")
        return region, values["company_display"]

    def render(self, batch, *, allow_warnings=False):
        self.counter += 1
        destination = self.root / f"{self.counter}-{batch['certificate_type']}.pdf"
        before = deepcopy(batch)
        templates.make_template_certificate(batch, COD, ISSUED, destination, allow_warnings=allow_warnings)
        self.assertEqual(batch, before, "Alignment must not mutate issued COD/date/status or caller metadata")
        reader = PdfReader(destination)
        self.assertEqual(len(reader.pages), 1)
        self.assertEqual(reader.metadata.subject, batch["certificate_type"])
        self.assertIn(COD, reader.metadata.title)
        extracted = reader.pages[0].extract_text()
        self.assertIn(COD, extracted)
        self.assertIn("01.10.2026", extracted)
        self.assertFalse((self.root / "Data").exists(), "The renderer must not write issuance state")
        return reader.pages[0]

    def assert_company_lines(self, page, batch, expected_count):
        region, company = self.company_region(batch)
        font, size, expected, leading, width, height, warning = templates._layout(region, company)
        self.assertFalse(warning, "A realistic client name must fit without confirmation")
        self.assertEqual(len(expected), expected_count)
        wanted = set(expected)
        records = [record for record in drawn_text(page) if record["text"] in wanted]
        self.assertEqual([record["text"] for record in records], expected,
                         "All fitted lines must be drawn in full, without ellipsis or stale native text")
        page_height = float(page.mediabox.height)
        tops, bottoms = [], []
        for record in records:
            self.assertEqual(record["transform"], (1, 0, 0, 1))
            self.assertAlmostEqual(record["size"], size, places=4)
            self.assertGreaterEqual(record["size"], 16)
            actual_width = pdfmetrics.stringWidth(record["text"], font, record["size"])
            self.assertAlmostEqual(record["x"] + actual_width / 2, 480, delta=0.002)
            self.assertGreaterEqual(record["x"], 90 - 0.002)
            self.assertLessEqual(record["x"] + actual_width, 870 + 0.002)
            ascent, descent = pdfmetrics.getAscentDescent(font, record["size"])
            tops.append(page_height - (record["y"] + ascent))
            bottoms.append(page_height - (record["y"] + descent))
        ml, mt, mr, mb = region["margins"]
        self.assertGreaterEqual(min(tops), region["top"] + mt - 0.002)
        self.assertLessEqual(max(bottoms), region["top"] + region["height"] - mb + 0.002)
        # Bounding the actual company ink between its neighbours' full boxes
        # is stronger than merely comparing their extracted baselines.
        others = templates.regions(batch["certificate_type"])
        above = [item for item in others if item["key"] in {"intro", "from_line"}]
        self.assertGreater(min(tops), max(item["top"] + item["height"] for item in above))
        below = [item["top"] for item in others
                 if item["key"] != "company_display" and item["top"] >= region["top"]]
        self.assertLess(max(bottoms), min(below))
        if expected_count == 2:
            self.assertAlmostEqual(records[0]["y"] - records[1]["y"], leading, delta=0.002)
        return records, (font, size, leading, height)

    def test_fixed_company_boxes_center_page_and_force_native_defaults_to_move(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                region, _ = self.company_region(self.batch(kind, "ACME INDIA LTD"))
                self.assertEqual((region["left"], region["width"], region["align"]), (90, 780, 2))
                self.assertEqual(region["left"] + region["width"] / 2, 480)
                self.assertTrue(region["moved"], "Even an exact native company name must be retyped at page centre")
                expected_top = {"standard": 212, "sita": 202, "siemens": 233.5}[kind]
                self.assertAlmostEqual(region["top"], expected_top, delta=0.01)
                self.assertEqual(region["size"], {"standard": 26, "sita": 19, "siemens": 24}[kind])
                self.assertEqual(region["margins"], [0, 1, 0, 1] if kind == "siemens" else [0, 2, 0, 2])
                if kind != "siemens":
                    self.assertEqual(region["height"], 44)

    def test_short_names_are_drawn_centered_at_original_preferred_sizes(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                batch = self.batch(kind, "ACME INDIA LTD")
                self.assertEqual(templates.layout_warnings(batch, COD, ISSUED), [])
                records, _ = self.assert_company_lines(self.render(batch), batch, 1)
                self.assertEqual(records[0]["size"], {"standard": 26, "sita": 19, "siemens": 24}[kind])

    def test_realistic_long_names_have_two_centered_lines_without_overlap_or_truncation(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                batch = self.batch(kind, LONG_COMPANY)
                self.assertEqual(templates.layout_warnings(batch, COD, ISSUED), [])
                self.assert_company_lines(self.render(batch), batch, 2)

    def test_explicit_two_line_name_survives_glyph_height_fit_not_old_line_cap(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                batch = self.batch(kind, MULTILINE_COMPANY)
                records, (_, _, leading, height) = self.assert_company_lines(self.render(batch), batch, 2)
                self.assertLess(height, 2 * leading, "This fixture must exercise the old height/leading truncation bug")
                self.assertEqual([record["text"] for record in records], MULTILINE_COMPANY.splitlines())

    def test_exact_native_company_names_are_not_reused_at_old_offcentre_positions(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                native = next(region for region in templates.regions(kind) if region["key"] == "company_display")
                batch = self.batch(kind, native["text"])
                # Standard's supplied placeholder is empty; it must still use
                # the same centred fallback company/city geometry.
                with patch.object(templates, "_original_region", wraps=templates._original_region) as reused:
                    page = self.render(batch)
                self.assertFalse(any(call.args[1]["key"] == "company_display" for call in reused.call_args_list))
                self.assert_company_lines(page, batch, 1)

    def test_extreme_names_warn_and_require_confirmation_with_at_most_two_ellipsis_lines(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                batch = self.batch(kind, "EXTREMELY LONG INTERNATIONAL RECYCLING COMPANY " * 100)
                warnings = templates.layout_warnings(batch, COD, ISSUED)
                self.assertTrue(any("Company line on certificate is long" in warning for warning in warnings))
                rejected = self.root / f"rejected-{kind}.pdf"
                with self.assertRaisesRegex(ValueError, "Company line on certificate is long"):
                    templates.make_template_certificate(batch, COD, ISSUED, rejected)
                self.assertFalse(rejected.exists())
                page = self.render(batch, allow_warnings=True)
                region, _ = self.company_region(batch)
                records = [record for record in drawn_text(page)
                           if region["top"] <= float(page.mediabox.height) - record["y"] <= region["top"] + region["height"]]
                self.assertTrue(records)
                self.assertLessEqual(len(records), 2)
                self.assertTrue(records[-1]["text"].endswith("..."))
                self.assertTrue(all(record["size"] == 16 for record in records))

    def test_fixed_alignment_and_wording_cannot_be_overridden_by_user_metadata(self):
        for kind in KINDS:
            with self.subTest(kind=kind):
                baseline = self.batch(kind, "ACME INDIA LTD")
                attempted = deepcopy(baseline)
                attempted["certificate_fields"].update(
                    _layout={"company_display": {"left": -100, "width": 100, "top": 0, "font_size": 6},
                             "table": {"top": 0, "scale": 2}},
                    _style={"align": 1, "font": "Courier"},
                    intro="UNAUTHORIZED INTRODUCTION", certificate_title="UNAUTHORIZED TITLE",
                    statement="UNAUTHORIZED STATEMENT", label_1="UNAUTHORIZED LABEL")
                self.assertEqual(templates.validate_fields(kind, attempted["certificate_fields"]), baseline["certificate_fields"])
                original = self.render(baseline)
                output = self.render(attempted)
                self.assertEqual(drawn_text(output), drawn_text(original))
                self.assertEqual(output.extract_text(), original.extract_text())
                self.assertNotIn("UNAUTHORIZED", output.extract_text())
                self.assert_company_lines(output, attempted, 1)

    def test_company_minimum_font_does_not_change_generic_data_fit_floor(self):
        batch = self.batch("siemens", "LONG COMPANY NAME " * 100)
        company, value = self.company_region(batch)
        generic = next(region for region in templates.regions("siemens") if region["key"] == "model_serial")
        self.assertEqual(templates._layout(company, value)[1], 16)
        self.assertEqual(templates._layout(generic, "VERY LONG SERIAL " * 100)[1], 6)


if __name__ == "__main__":
    unittest.main()
