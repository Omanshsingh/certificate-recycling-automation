"""Serve the real dashboard against disposable data for browser workflow checks.

Run from the repository: python App/tests/serve_ui_sandbox.py
No operator workbooks, settings or certificate files are used or modified.
"""
from datetime import date, timedelta
from http.server import ThreadingHTTPServer
from pathlib import Path
import argparse
import shutil
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import web_app
from run_certificates import run
from tests.test_certificate_workflow import make_workbook
from tests.test_certificate_inventory import external_pdf
from workflow_support import save_json


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8875)
    parser.add_argument('--sample-pdfs', type=int, choices=range(0, 11), default=0,
                        help='Add independently created sample PDFs for saved-file testing.')
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix='certificate-ui-sandbox-') as temporary:
        root = Path(temporary)
        templates = root / 'App' / 'templates'
        shutil.copytree(Path(web_app.__file__).parent / 'templates', templates)
        shutil.copy2(Path(web_app.__file__).parent / 'settings.example.json', root / 'App' / 'settings.example.json')
        paths = [root / 'Plant 1.xlsx', root / 'Plant 2.xlsx']
        today = date.today()
        make_workbook(paths[0], [
            {'sn':'1001','supplier':'Acme Recycling Ltd','date':today-timedelta(days=5),'weight':12.5},
            {'sn':'1002','supplier':'Harbor Electronics','date':today-timedelta(days=24),'weight':8},
            {'sn':'1003','supplier':'Generated Example','date':today-timedelta(days=3),'weight':7},
            {'sn':'1004','supplier':'Inward Awaiting Details','status':'','form6':'','date':today-timedelta(days=26)},
        ])
        make_workbook(paths[1], [
            {'sn':'2001','supplier':'Northstar Services','date':today-timedelta(days=25),'weight':20},
            {'sn':'2002','supplier':'','date':today-timedelta(days=27),'weight':4},
            {'sn':'2003','supplier':'No Certificate Needed','status':'N/A','date':today-timedelta(days=30)},
        ])
        save_json(root / 'settings.local.json', {'workbook_paths':[str(path) for path in paths],
            'certificate_number_prefix':'','certificate_number_start':1,'certificate_number_width':4,
            'reminders_enabled':False,'reminder_time':'11:00'})
        web_app.ROOT = root
        web_app.BACKUP_DIR = root / 'Data' / 'Workbook Backups'
        web_app.RESULTS_PATH = root / 'Data' / 'Reports' / 'Latest Results.json'
        run(paths, root, selected=('plant1','1003'), progress=None)
        companies = ['Seaside Electronics Chennai', 'Riverside Laptops Pune',
                     'Northstar Electronics Recovery and Secure Asset Lifecycle Services Jaipur',
                     'Sita Metals Delhi', 'Siemens Medical Bengaluru', 'Harbor Assets Kochi',
                     'Acme Services Mumbai', 'Westgate Equipment Surat',
                     'Summit Technology Hyderabad', 'Metro Computers Ahmedabad']
        for index, company in enumerate(companies[:args.sample_pdfs], 1):
            cod = f'{9000 + index:04d}'
            filename = f'{cod}_{company}_{1 + index % 2}.pdf'
            external_pdf(root / 'Certificates' / filename, cod=cod, reference=f'MANUAL-{index}')
        web_app.refresh()
        server = ThreadingHTTPServer(('127.0.0.1', args.port), web_app.Handler)
        web_app.SERVER = server
        print(f'Sandbox: {root}\nDashboard: http://127.0.0.1:{args.port}/', flush=True)
        try:
            server.serve_forever(poll_interval=0.25)
        except KeyboardInterrupt:
            pass
        finally:
            server.server_close()


if __name__ == '__main__':
    main()
