"""Explicit browser QA server. Never uses the operator's workbooks or PDFs."""
from pathlib import Path
import shutil
import tempfile

import web_app
from tests.test_certificate_workflow import make_workbook
from workflow_support import save_json


def main():
    with tempfile.TemporaryDirectory(prefix='certificate-browser-qa-') as name:
        root = Path(name)
        paths = [root / f'Plant {plant}.xlsx' for plant in (1, 2)]
        for plant, path in enumerate(paths, 1):
            make_workbook(path, [{'sn':f'{plant}0{index:02}',
                'supplier': ('Northstar Electronics Recovery and Secure Asset Lifecycle Services' if index == 1 else f'Acme Client {index}'),
                'reference':f'REF-{plant}-{index}', 'form6':f'F6-{plant}-{index}'} for index in range(1, 7)])
        templates = root / 'App' / 'templates'
        shutil.copytree(Path(web_app.__file__).parent / 'templates', templates)
        save_json(root / 'settings.local.json', {'workbook_paths':[str(p) for p in paths],
                    'reminders_enabled':False, 'reminder_time':'11:00', 'backup_destination':str(root / 'Backups Separate')})
        web_app.ROOT = root
        web_app.BACKUP_DIR = root / 'Data' / 'Workbook Backups'
        web_app.RESULTS_PATH = root / 'Data' / 'Reports' / 'Latest Results.json'
        web_app.STOPPING = False
        web_app.SERVER = web_app.ThreadingHTTPServer(('127.0.0.1', 8766), web_app.Handler)
        web_app.SERVER.daemon_threads = False
        print(f'QA workspace: {root}', flush=True)
        print('QA server: http://127.0.0.1:8766/', flush=True)
        try:
            web_app.SERVER.serve_forever(poll_interval=0.25)
        finally:
            web_app.SERVER.server_close()


if __name__ == '__main__':
    main()
