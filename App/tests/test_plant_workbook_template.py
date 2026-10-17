from io import BytesIO
from pathlib import Path
import json
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import load_workbook

from certificate_core import read_batches
from client_intake import stage_new_client
from plant_workbook_template import HEADERS, create_plant_workbook, plant_workbook_bytes
import web_app
from workflow_support import save_json
from workspace_review import workspace_revision


class PlantWorkbookTemplateTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(self.enterContext(tempfile.TemporaryDirectory(prefix="plant-template-")))

    def test_blank_template_has_the_shared_headers_and_a_fully_ruled_entry_row(self):
        for plant in ("plant1", "plant2"):
            with self.subTest(plant=plant):
                data = plant_workbook_bytes(plant)
                book = load_workbook(BytesIO(data), data_only=False)
                sheet = book["Sheet1"]
                self.assertEqual(tuple(cell.value for cell in sheet[1]), HEADERS)
                self.assertEqual(sheet.max_column, len(HEADERS))
                self.assertEqual(sheet["G1"].value, "Per Pcs Weight in (KG)")
                self.assertEqual(sheet["H1"].value, "Weight (KG)")
                self.assertEqual(sheet.max_row, 2)
                self.assertEqual(len(sheet.merged_cells.ranges), 0)
                self.assertEqual(sheet.freeze_panes, "E2")
                self.assertTrue(sheet.sheet_view.showGridLines)
                self.assertEqual(sheet["A1"].font.name, "Calibri")
                self.assertTrue(sheet["A1"].font.bold)
                self.assertEqual(sheet["A1"].fill.fgColor.rgb, "FF4DD0E1")
                self.assertEqual(sheet["A1"].alignment.horizontal, "center")
                self.assertEqual(sheet["A1"].border.left.style, "thin")
                self.assertEqual(sheet["H2"].border.top.style, "medium")
                self.assertEqual(sheet["H2"].border.bottom.style, "medium")
                self.assertEqual(sheet["A2"].border.left.style, "medium")
                self.assertEqual(sheet["A2"].border.bottom.style, "medium")
                self.assertEqual(sheet["A2"].alignment.horizontal, "center")
                self.assertEqual(sheet["C2"].alignment.horizontal, "center")
                self.assertEqual(sheet["O2"].alignment.horizontal, "center")
                self.assertTrue(any("R2:R10000" in str(rule.sqref)
                                    for rule in sheet.conditional_formatting))
                self.assertEqual(len(sheet.data_validations.dataValidation), 1)
                self.assertEqual(str(sheet.data_validations.dataValidation[0].sqref), "R2:R10000")
                self.assertEqual(read_batches(BytesIO(data)), [])
                book.close()

    def test_app_added_client_gets_total_formula_and_keeps_register_readable(self):
        data = plant_workbook_bytes("plant1")
        updated, batch = stage_new_client(SimpleNamespace(data=data), {
            "supplier": "Template client", "location": "Pune, Maharashtra, India",
            "inward_date": "2026-10-01", "form6": "F6-1", "reference": "REF-1",
            "items": [
                {"item": "Laptop", "quantity": 2, "weight": 12.5},
                {"item": "Monitor", "quantity": 1, "weight": 8},
            ],
        })
        book = load_workbook(BytesIO(updated), data_only=False)
        sheet = book["Sheet1"]
        self.assertEqual(sheet["J2"].value, "=SUM(H2:H3)")
        self.assertEqual(sheet["R2"].value, "Ready")
        self.assertEqual(sheet["H2"].value, 12.5)
        self.assertEqual(sheet["H3"].value, 8)
        self.assertEqual(sheet["A2"].alignment.horizontal, "center")
        self.assertEqual(sheet["E2"].alignment.horizontal, "center")
        self.assertEqual(sheet["E3"].alignment.horizontal, "center")
        self.assertIn("A2:A3", {str(area) for area in sheet.merged_cells.ranges})
        book.close()

        saved = read_batches(BytesIO(updated))
        self.assertEqual(len(saved), 1)
        self.assertEqual(saved[0]["case_id"], batch["case_id"])
        self.assertEqual(saved[0]["total_weight"], 20.5)
        self.assertEqual(saved[0]["reference"], "REF-1")

    def test_single_item_client_keeps_a_complete_dark_border(self):
        updated, _ = stage_new_client(SimpleNamespace(data=plant_workbook_bytes("plant1")), {
            "supplier": "One item", "location": "Pune", "form6": "F6-ONE",
            "items": [{"item": "Laptop", "quantity": 1, "weight": 2.5}],
        })
        book = load_workbook(BytesIO(updated), data_only=False)
        sheet = book["Sheet1"]
        for coordinate in ("A2", "E2", "H2", "R2"):
            with self.subTest(cell=coordinate):
                self.assertEqual(sheet[coordinate].border.top.style, "medium")
                self.assertEqual(sheet[coordinate].border.bottom.style, "medium")
        self.assertEqual(sheet["H2"].value, 2.5)
        book.close()

    def test_new_template_saves_operational_fields_and_supports_numeric_or_text_quantities(self):
        data = plant_workbook_bytes("plant1")
        updated, _ = stage_new_client(SimpleNamespace(data=data), {
            "supplier":"Template client", "location":"Pune", "inward_date":"2026-10-01",
            "form6":"F6-1", "comments":"Call before pickup",
            "transporter_details":"ABC Logistics", "docket_vehicle_no":"DCK-101",
            "weight_slip_scale":"Scale ticket 44", "items":[
                {"item":"Laptop", "quantity":"2", "per_piece_weight":"3.5"},
                {"item":"Cable lot", "quantity":"1 lot", "per_piece_weight":"3.5"},
            ],
        })
        book = load_workbook(BytesIO(updated), data_only=False)
        sheet = book["Sheet1"]
        self.assertEqual([sheet.cell(2, column).value for column in range(11, 15)],
                         ["Call before pickup", "ABC Logistics", "DCK-101", "Scale ticket 44"])
        self.assertEqual(sheet["F2"].value, 2)
        self.assertEqual(sheet["G2"].value, 3.5)
        self.assertEqual(sheet["H2"].value, "=F2*G2")
        self.assertEqual(sheet["F3"].value, "1 lot")
        self.assertEqual(sheet["G3"].value, 3.5)
        self.assertIsNone(sheet["H3"].value)
        book.close()
        batches = read_batches(BytesIO(updated))
        self.assertEqual(batches[0]["items"][0]["weight"], 7)
        self.assertEqual(batches[0]["items"][1]["quantity"], "1 lot")
        self.assertIsNone(batches[0]["items"][1]["weight"])

    def test_create_protects_existing_workbook_and_only_writes_xlsx(self):
        destination = self.root / "Plant 1.xlsx"
        destination.write_bytes(b"keep this existing workbook")
        with self.assertRaisesRegex(ValueError, "already exists"):
            create_plant_workbook(destination, "plant1")
        self.assertEqual(destination.read_bytes(), b"keep this existing workbook")
        with self.assertRaisesRegex(ValueError, r"\.xlsx"):
            create_plant_workbook(self.root / "Plant 2.xlsm", "plant2")
        with self.assertRaisesRegex(ValueError, "Choose Plant 1 or Plant 2"):
            plant_workbook_bytes("plant3")

    def test_create_writes_a_valid_empty_register(self):
        destination = create_plant_workbook(self.root / "Plant 2.xlsx", "plant2")
        self.assertTrue(destination.is_file())
        self.assertEqual(read_batches(destination), [])
        book = load_workbook(destination, read_only=True)
        self.assertEqual(book.sheetnames, ["Sheet1"])
        book.close()

    def post_template(self, destination, plant="plant1"):
        handler = object.__new__(web_app.Handler)
        handler.path = "/api/create-plant-workbook"
        payload = json.dumps({"plant": plant}).encode("utf-8")
        handler.headers = {"Host": "127.0.0.1", "Content-Length": str(len(payload))}
        handler.rfile = BytesIO(payload)
        responses = []
        handler.send_json = lambda value, status=200: responses.append((status, value))
        with patch.object(web_app, "ROOT", self.root), \
                patch("tkinter.Tk"), \
                patch("tkinter.filedialog.asksaveasfilename", return_value=str(destination)):
            handler.do_POST()
        self.assertEqual(len(responses), 1)
        return responses[0]

    def test_create_template_api_returns_path_without_connecting_until_settings_are_saved(self):
        paths = [self.root / "Plant 1.xlsx", self.root / "Plant 2.xlsx"]
        save_json(self.root / "settings.local.json", {"workbook_paths": [str(path) for path in paths]})
        settings_before = (self.root / "settings.local.json").read_bytes()
        destination = self.root / "New Plant 1.xlsx"
        status, response = self.post_template(destination)
        self.assertEqual(status, 200)
        self.assertEqual(response["plant"], "plant1")
        self.assertEqual(Path(response["path"]), destination.resolve())
        self.assertIn("Click Save in Settings", response["message"])
        self.assertEqual((self.root / "settings.local.json").read_bytes(), settings_before)
        self.assertEqual(read_batches(destination), [])

    def test_create_template_cancel_does_not_create_a_file(self):
        destination = self.root / "Cancelled.xlsx"
        status, response = self.post_template("")
        self.assertEqual(status, 200)
        self.assertTrue(response["cancelled"])
        self.assertFalse(destination.exists())

    def test_created_plant_templates_accept_clients_through_the_app_without_manual_header_work(self):
        paths = [create_plant_workbook(self.root / f"Plant {index}.xlsx", f"plant{index}")
                 for index in (1, 2)]
        save_json(self.root / "settings.local.json", {
            "workbook_paths": [str(path) for path in paths], "reminders_enabled": False,
        })
        with patch.object(web_app, "ROOT", self.root), \
                patch.object(web_app, "BACKUP_DIR", self.root / "Data" / "Workbook Backups"), \
                patch.object(web_app, "refresh", return_value={"rows": [], "scan_complete": True}):
            for index, plant in enumerate(("plant1", "plant2"), 1):
                result = web_app.add_client({
                    "plant": plant, "request_id": f"{index:032x}",
                    "expected_revision": workspace_revision(self.root),
                    "details": {
                        "supplier": f"New client {index}", "location": "Pune, Maharashtra, India",
                        "inward_date": "2026-10-01", "form6": f"F6-{index}",
                        "reference": f"REF-{index}",
                        "items": [
                            {"item": "Laptop", "quantity": 2, "weight": 12.5},
                            {"item": "Monitor", "quantity": 1, "weight": 8},
                        ],
                    },
                })
                self.assertTrue(result["client_added"])
                self.assertEqual(result["client"], {
                    "plant": plant, "sn": "1", "supplier": f"New client {index}", "status": "ready",
                })
                saved = read_batches(paths[index - 1])
                self.assertEqual(len(saved), 1)
                self.assertEqual(saved[0]["total_weight"], 20.5)


if __name__ == "__main__":
    unittest.main()
