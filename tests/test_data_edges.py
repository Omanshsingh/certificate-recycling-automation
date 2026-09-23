"""Local-only parsing, layout, and package-write regression tests."""
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import re
import ctypes
import os
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill, Protection
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from reportlab.pdfbase.pdfmetrics import stringWidth

from certificate_core import certificate_layout, columns, eligible, fit, make_certificate, number, parse_date, read_batches, valid_email
from workbook_store import WorkbookSession, _set_cell, exclusive_file, LockBusyError


TODAY = date(2026, 9, 23)
HEADERS = ["SN", "Inward Date", "Supplier Name", "Location", "Item", "Quantity", "Per Piece Weight", "Weight", "Form 6 No", "Client Email", "Reference No", "COD No", "COD Date", "Certificate Issued"]


class DataEdges(unittest.TestCase):
    def test_overlays_stay_inside_actual_template_value_cells(self):
        batch = {"inward_date": TODAY, "total_weight": 135.6, "reference": "REFERENCE-1234567890", "form6": "MANIFEST-1234567890", "supplier": "Dummy Client Ltd", "location": "Mumbai"}
        fields, _, _ = certificate_layout(batch, "0062", TODAY)
        for index, ((value, size), x, _) in enumerate(fields):
            # Original template dividers are approximately x550 and x620 points.
            self.assertLessEqual(x + stringWidth(value, "Helvetica-Bold", size), 546 if index == 0 else 614)
        for name in ("reference", "form6"):
            with self.subTest(field=name), self.assertRaisesRegex(ValueError, "too long"):
                certificate_layout(dict(batch, **{name: "W" * 20}), "0062", TODAY)

    def test_email_rejects_lists_header_text_and_malformed_dots(self):
        for value in ("client@example.test;", "client@example.test,other", "first@example.test;second@example.test", "name <client@example.test>", "a..b@example.test", "a@bad..test", "client@example.test\nBcc: other@example.test"):
            with self.subTest(value=value):
                self.assertFalse(valid_email(value))
        self.assertTrue(valid_email("client+tag@example.test"))
        self.assertTrue(valid_email("First.Last@example.test"))

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.path = self.root / "Dummy.xlsx"
        self.backups = self.root / "backups"
        self.addCleanup(self.temp.cleanup)

    def book(self, changes=None, continuation=None, edit=None):
        book = Workbook()
        sheet = book.active
        sheet.title = "Register"
        sheet.append(HEADERS)
        row = ["1", date(2026, 8, 1), "Dummy Recycling Ltd", "Mumbai", "Computers", 2, 3.5, "=F2*G2", "FORM-1", "dummy@example.test", "REF-1", None, None, "Ready"]
        for key, value in (changes or {}).items():
            row[HEADERS.index(key)] = value
        sheet.append(row)
        if continuation:
            sheet.append([continuation.get(key) for key in HEADERS])
        if edit:
            edit(book, sheet)
        book.save(self.path)
        book.close()
        return self.path

    def batch(self, **kwargs):
        return read_batches(self.book(**kwargs))[0]

    def test_numeric_values_require_complete_finite_numbers(self):
        for value in [True, False, "12kg garbage", "1/2", "1,2", "--2", "3.2.1", float("nan"), float("inf"), "=1+2"]:
            with self.subTest(value=value):
                self.assertIsNone(number(value))
        for value, expected in [("1,234.50", 1234.5), ("1,23,456.5", 123456.5), ("1e3", 1000), (".5", .5), ("-0.5", -.5), ("0", 0)]:
            with self.subTest(value=value):
                self.assertEqual(number(value), expected)

    def test_duplicate_and_alias_headers_are_rejected(self):
        for extra in ["Weight", "Weight (KG)", "Client Email ID", "SN"]:
            with self.subTest(extra=extra):
                with self.assertRaisesRegex(ValueError, "[Aa]mbiguous|[Dd]uplicate"):
                    columns(HEADERS + [extra])

    def test_ambiguous_register_headers_are_rejected(self):
        self.book(edit=lambda _, sheet: sheet.cell(1, 15, "Weight (KG)"))
        with self.assertRaisesRegex(ValueError, "[Aa]mbiguous|[Dd]uplicate"):
            read_batches(self.path)

    def test_two_register_sheets_are_rejected_as_ambiguous(self):
        def add_second_register(book, _):
            book.create_sheet("Archived Register").append(HEADERS)
        self.book(edit=add_second_register)
        with self.assertRaisesRegex(ValueError, "one inward register"):
            read_batches(self.path)

    def test_only_unambiguous_same_row_product_formulas_are_calculated(self):
        for formula in ["=$F$2 * G2", "=G2*F2"]:
            with self.subTest(formula=formula):
                batch = self.batch(changes={"Weight": formula})
                self.assertFalse(batch["issues"], batch["issues"])
                self.assertEqual(batch["total_weight"], 7)
        for value in ["=F2*G2+100", "=F3*G3", "=SUM(F2:G2)", "bad", "#VALUE!"]:
            with self.subTest(value=value):
                self.assertTrue(self.batch(changes={"Weight": value})["issues"])

    def test_blank_weight_can_use_quantity_but_invalid_factors_block(self):
        self.assertEqual(self.batch(changes={"Weight": None})["total_weight"], 7)
        for changes in [{"Quantity": -2, "Per Piece Weight": -3.5}, {"Quantity": "two", "Weight": 7}, {"Quantity": "=1+1"}, {"Per Piece Weight": True}]:
            with self.subTest(changes=changes):
                self.assertTrue(self.batch(changes=changes)["issues"])

    def test_continuation_rows_keep_items_and_flag_conflicting_metadata(self):
        item = {"Item": "Screens", "Quantity": 3, "Per Piece Weight": 2, "Weight": "=F3*G3"}
        batch = self.batch(continuation=item)
        self.assertEqual(batch["total_weight"], 13)
        self.assertFalse(batch["issues"])
        for key, value in [("Supplier Name", "Different client"), ("Certificate Issued", "N/A"), ("Reference No", "REF-2")]:
            with self.subTest(key=key):
                batch = self.batch(continuation={**item, key: value})
                self.assertTrue(batch["issues"])

    def test_continuation_cod_cannot_disappear_from_global_numbering(self):
        with self.assertRaisesRegex(ValueError, "COD No on continuation row"):
            self.batch(continuation={"COD No": "0099"})
        batch = self.batch(changes={"COD No": "0099"}, continuation={"COD No": "0099", "Item": "Screens", "Weight": 2})
        self.assertEqual(batch["cod_no"], "0099")
        self.assertFalse(batch["issues"])

    def test_orphan_item_rows_are_not_silently_ignored(self):
        self.book(changes={"SN": None})
        with self.assertRaisesRegex(ValueError, "SN|orphan"):
            read_batches(self.path)

    def test_formula_metadata_cannot_be_printed_as_literal_reference(self):
        for key in ["Reference No", "Form 6 No", "Supplier Name"]:
            with self.subTest(key=key):
                self.assertTrue(self.batch(changes={key: '="value"'})["issues"])

    def test_dates_and_boundary_do_not_guess_numeric_serials(self):
        self.assertEqual(parse_date(datetime(2024, 2, 29, 18)), date(2024, 2, 29))
        self.assertEqual(parse_date("29.02.2024"), date(2024, 2, 29))
        for value in ["29.02.2023", "31/04/2026", 45200, "=TODAY()"]:
            self.assertIsNone(parse_date(value))
        self.assertFalse(eligible({"inward_date": date.max}, TODAY, 30))
        self.assertTrue(eligible({"inward_date": date(2026, 8, 24)}, TODAY, 30))

    def test_field_fitting_rejects_overflow_at_minimum_size(self):
        value, size = fit("Reference 12345", "Helvetica-Bold", 8.5, 180)
        self.assertLessEqual(stringWidth(value, "Helvetica-Bold", size), 180)
        with self.assertRaisesRegex(ValueError, "too long|fit"):
            fit("R" * 200, "Helvetica-Bold", 7.5, 180)

    def test_invalid_layout_is_rejected_before_creating_pdf(self):
        batch = self.batch(changes={"Reference No": "R" * 200})
        destination = self.root / "out" / "certificate.pdf"
        with patch("certificate_core.canvas.Canvas") as create:
            with self.assertRaisesRegex(ValueError, "too long|fit"):
                make_certificate(batch, "0001", TODAY, "unused.png", destination)
            create.assert_not_called()
        self.assertFalse(destination.exists())

    def test_merged_output_anchor_and_child_are_rejected_before_write(self):
        for cell_range in ["L2:M2", "L2:L3", "K2:L2"]:
            with self.subTest(cell_range=cell_range):
                batch = self.batch(edit=lambda _, sheet: sheet.merge_cells(cell_range))
                before = self.path.read_bytes()
                with WorkbookSession(self.path, self.backups) as session:
                    with self.assertRaisesRegex(ValueError, "merged"):
                        session.prepare(batch)
                self.assertEqual(self.path.read_bytes(), before)

    def test_protection_and_formula_outputs_are_rejected(self):
        def protect(_, sheet):
            sheet.protection.sheet = True
        batch = self.batch(edit=protect)
        with WorkbookSession(self.path, self.backups) as session:
            with self.assertRaisesRegex(ValueError, "protected"):
                session.prepare(batch)
        batch = self.batch(changes={"COD Date": "=TODAY()"})
        with WorkbookSession(self.path, self.backups) as session:
            with self.assertRaisesRegex(ValueError, "formula"):
                session.prepare(batch)

    def test_absent_output_xml_cells_are_inserted_and_other_parts_preserved(self):
        def features(book, sheet):
            book.create_sheet("Untouched")["A1"] = "=1+2"
            sheet["A2"].alignment = Alignment(wrap_text=True)
            sheet.freeze_panes = "C2"
            sheet.merge_cells("O2:P2")
        batch = self.batch(edit=features)
        with ZipFile(self.path, "a") as archive:
            archive.writestr("customXml/preserve.xml", b"<native feature='unchanged'/>")
            archive.comment = b"Keep this archive comment"
        original = self.path.read_bytes()
        with ZipFile(BytesIO(original)) as archive:
            self.assertNotRegex(archive.read("xl/worksheets/sheet1.xml").decode(), r'<c r="[LM]2"')
        with WorkbookSession(self.path, self.backups) as session:
            backup = session.commit(batch, "0001", TODAY)
        self.assertEqual(backup.read_bytes(), original)
        with ZipFile(BytesIO(original)) as before, ZipFile(self.path) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertEqual(before.comment, after.comment)
            for name in before.namelist():
                if name not in {"xl/worksheets/sheet1.xml", "xl/styles.xml"}:
                    self.assertEqual(before.read(name), after.read(name), name)
            old_xml = before.read("xl/worksheets/sheet1.xml").decode()
            new_xml = after.read("xl/worksheets/sheet1.xml").decode()
            self.assertEqual(old_xml, re.sub(r'<c r="[LM]2".*?</c>', '', new_xml))
        book = load_workbook(self.path)
        self.addCleanup(book.close)
        self.assertEqual(book.active["L2"].value, "0001")
        self.assertEqual(book.active["M2"].value.date(), TODAY)

    def test_date_epoch_and_existing_date_style_survive(self):
        def edit(book, sheet):
            book.epoch = CALENDAR_MAC_1904
            sheet["M2"].number_format = "dd.mm.yyyy"
            sheet["M2"].fill = PatternFill("solid", fgColor="AAFFAA")
            sheet.protection.sheet = True
            for coord in ["L2", "M2"]:
                sheet[coord].protection = Protection(locked=False)
        batch = self.batch(edit=edit)
        with ZipFile(self.path) as archive:
            styles = archive.read("xl/styles.xml")
        with WorkbookSession(self.path, self.backups) as session:
            session.commit(batch, "0001", TODAY)
        book = load_workbook(self.path)
        self.addCleanup(book.close)
        self.assertEqual(book.epoch, CALENDAR_MAC_1904)
        self.assertEqual(book.active["M2"].value.date(), TODAY)
        self.assertEqual(book.active["M2"].number_format, "dd.mm.yyyy")
        self.assertEqual(book.active["M2"].fill.fgColor.rgb, "00AAFFAA")
        with ZipFile(self.path) as archive:
            self.assertEqual(archive.read("xl/styles.xml"), styles)

    def test_atomic_replace_failure_keeps_original_and_allows_retry(self):
        batch = self.batch()
        original = self.path.read_bytes()
        with WorkbookSession(self.path, self.backups) as session:
            with patch("workbook_store.os.replace", side_effect=OSError("replace failed")):
                with self.assertRaisesRegex(OSError, "replace failed"):
                    session.commit(batch, "0001", TODAY)
            self.assertEqual(session.data, original)
            self.assertEqual(session.backup_path.read_bytes(), original)
            self.assertFalse(list(self.root.glob(".certificate_*")))
            session.commit(batch, "0001", TODAY)
        self.assertEqual(read_batches(self.path)[0]["cod_no"], "0001")

    def test_failed_backup_is_not_reused_on_retry(self):
        batch = self.batch()
        original = self.path.read_bytes()
        original_open = Path.open
        class FailingBackup:
            def __enter__(self):
                return self
            def __exit__(self, *args):
                pass
            def write(self, data):
                raise OSError("backup disk full")
        def open_with_failure(path, mode="r", *args, **kwargs):
            if path.parent == self.backups and mode == "xb":
                return FailingBackup()
            return original_open(path, mode, *args, **kwargs)
        with WorkbookSession(self.path, self.backups) as session:
            with patch.object(Path, "open", open_with_failure):
                with self.assertRaisesRegex(OSError, "backup disk full"):
                    session.commit(batch, "0001", TODAY)
            self.assertIsNone(session.backup_path)
            self.assertEqual(session.data, original)
            backup = session.commit(batch, "0001", TODAY)
        self.assertEqual(backup.read_bytes(), original)

    def test_sync_failure_preserves_original_and_cleans_partial_files(self):
        for failed_sync in [1, 2]:
            with self.subTest(failed_sync=failed_sync):
                batch = self.batch()
                original = self.path.read_bytes()
                with WorkbookSession(self.path, self.backups) as session:
                    outcomes = [None] * (failed_sync - 1) + [OSError("disk sync failed")]
                    with patch("workbook_store.os.fsync", side_effect=outcomes):
                        with self.assertRaisesRegex(OSError, "disk sync failed"):
                            session.commit(batch, "0001", TODAY)
                    self.assertEqual(session.data, original)
                    self.assertFalse(list(self.root.glob(".certificate_*")))
                    if failed_sync == 1:
                        self.assertIsNone(session.backup_path)
                        self.assertFalse(list(self.backups.glob("*.xlsx")))
                    else:
                        self.assertEqual(session.backup_path.read_bytes(), original)
                self.assertEqual(self.path.read_bytes(), original)

    def test_new_date_format_retains_output_cell_appearance(self):
        def edit(_, sheet):
            sheet["M2"].number_format = "General"
            sheet["M2"].fill = PatternFill("solid", fgColor="AAFFAA")
            sheet["M2"].alignment = Alignment(horizontal="center", wrap_text=True)
        batch = self.batch(edit=edit)
        with WorkbookSession(self.path, self.backups) as session:
            session.commit(batch, "0001", TODAY)
        book = load_workbook(self.path)
        self.addCleanup(book.close)
        cell = book.active["M2"]
        self.assertEqual(cell.value.date(), TODAY)
        self.assertEqual(cell.fill.fgColor.rgb, "00AAFFAA")
        self.assertEqual(cell.alignment.horizontal, "center")
        self.assertTrue(cell.alignment.wrap_text)

    def test_closed_session_cannot_write(self):
        batch = self.batch()
        with WorkbookSession(self.path, self.backups) as session:
            pass
        with self.assertRaisesRegex(PermissionError, "unavailable"):
            session.commit(batch, "0001", TODAY)

    @unittest.skipUnless(os.name == "nt", "Windows sharing-error classification")
    def test_lock_contention_is_distinct_from_access_denied(self):
        for error in [32, 33, 5, 2]:
            with self.subTest(error=error):
                with patch("ctypes.WinDLL") as api, patch("ctypes.get_last_error", return_value=error):
                    api.return_value.CreateFileW.return_value = ctypes.c_void_p(-1).value
                    with self.assertRaises(PermissionError) as caught:
                        with exclusive_file(self.path):
                            self.fail("Lock must fail")
                    self.assertEqual(isinstance(caught.exception, LockBusyError), error in (32, 33))

    def test_actual_double_lock_raises_busy_error(self):
        self.book()
        with exclusive_file(self.path):
            with self.assertRaises(LockBusyError):
                with exclusive_file(self.path):
                    self.fail("Second exclusive lock must not succeed")

    def test_zero_cod_date_is_not_an_empty_cell(self):
        batch = self.batch(changes={"COD Date": 0})
        original = self.path.read_bytes()
        with WorkbookSession(self.path, self.backups) as session:
            with self.assertRaisesRegex(ValueError, "COD Date is already filled"):
                session.commit(batch, "0001", TODAY)
        self.assertEqual(self.path.read_bytes(), original)

    def test_xml_attribute_entities_are_not_double_escaped(self):
        xml = '<worksheet><sheetData><row r="2"><c r="L2" custom="a&amp;b"/></row></sheetData></worksheet>'
        changed = _set_cell(xml, "L2", "0001", "inlineStr")
        self.assertIn('custom="a&amp;b"', changed)
        self.assertNotIn("&amp;amp;", changed)


if __name__ == "__main__":
    unittest.main()
