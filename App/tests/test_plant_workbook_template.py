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

    def test_blank_template_has_the_shared_register_headers_and_a_styled_client_block(self):
        for plant in ("plant1", "plant2"):
            with self.subTest(plant=plant):
                data = plant_workbook_bytes(plant)
                book = load_workbook(BytesIO(data), data_only=False)
                sheet = book["Sheet1"]
                self.assertEqual(tuple(cell.value for cell in sheet[1]), HEADERS)
                self.assertEqual(sheet.max_column, len(HEADERS))
                self.assertIn("A2:A4", {str(area) for area in sheet.merged_cells.ranges})
                self.assertIn("I2:I4", {str(area) for area in sheet.merged_cells.ranges})
                self.assertIn("Q2:Q4", {str(area) for area in sheet.merged_cells.ranges})
                self.assertEqual(sheet.freeze_panes, "E2")
                self.assertEqual(sheet["A1"].font.name, "Calibri")
                self.assertTrue(sheet["A1"].font.bold)
                self.assertEqual(sheet["A2"].border.left.style, "medium")
                self.assertEqual(sheet["A4"].border.bottom.style, "medium")
                self.assertTrue(any("Q2:Q10000" in str(rule.sqref)
                                    for rule in sheet.conditional_formatting))
                self.assertEqual(len(sheet.data_validations.dataValidation), 1)
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
        self.assertEqual(sheet["I2"].value, "=SUM(G2:G3)")
        self.assertEqual(sheet["Q2"].value, "Ready")
        self.assertEqual(sheet["G2"].value, 12.5)
        self.assertEqual(sheet["G3"].value, 8)
        self.assertIn("A2:A3", {str(area) for area in sheet.merged_cells.ranges})
        book.close()

        saved = read_batches(BytesIO(updated))
        self.assertEqual(len(saved), 1)
        self.assertEqual(saved[0]["case_id"], batch["case_id"])
        self.assertEqual(saved[0]["total_weight"], 20.5)
        self.assertEqual(saved[0]["reference"], "REF-1")

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
