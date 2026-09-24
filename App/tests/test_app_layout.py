"""Installed layout must keep the operator's existing data outside App."""
from pathlib import Path
import tempfile
import unittest
from app_paths import DATA_ROOT, PROGRAM_ROOT, resource
from workflow_support import load_settings, save_json


class LayoutTests(unittest.TestCase):
    def test_installed_layout_keeps_data_root_and_resources_distinct(self):
        self.assertEqual(PROGRAM_ROOT.name, 'App')
        self.assertEqual(DATA_ROOT, PROGRAM_ROOT.parent)
        self.assertEqual(resource(DATA_ROOT, 'templates'), PROGRAM_ROOT / 'templates')
        self.assertTrue((resource(DATA_ROOT, 'templates') / 'Certificate of Recycling.png').is_file())
        self.assertTrue((DATA_ROOT / 'Prepare Certificates.bat').is_file())

    def test_existing_local_settings_override_packaged_defaults(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            save_json(root / 'App' / 'settings.example.json', {'certificate_number_start': 1})
            save_json(root / 'settings.local.json', {'certificate_number_start': 42, 'workbook_paths': ['plant1.xlsx', 'plant2.xlsx']})
            settings = load_settings(root)
            self.assertEqual(settings['certificate_number_start'], 42)
            self.assertEqual(settings['workbook_paths'], ['plant1.xlsx', 'plant2.xlsx'])
