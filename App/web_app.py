"""Loopback-only browser interface for the Ecoreco certificate workflow."""
from __future__ import annotations

from datetime import datetime
from http.client import HTTPConnection
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import subprocess
import threading
from urllib.parse import urlsplit, unquote, quote
import webbrowser
import re
import uuid
import shutil

from app_paths import DATA_ROOT
from certificate_core import status_value
from run_certificates import run, find_certificate
from workbook_store import update_issued_status, WorkbookSession, stage_certificate_details
from workflow_support import load_settings, save_json
from certificate_drafts import (clean_details, save_draft, generate_draft, complete_draft,
                                prepare_certificate, manual_rows, reconcile_manual)
from certificate_templates import template_specs, make_template_certificate, obsolete_template_warning
from pypdf import PdfReader
from certificate_inventory import scan_inventory
from workflow_support import run_lock
from issuance_safety import recover_issuances, valid_pdf, latest_issuance, audit_records
from workspace_review import workspace_revision, require_revision, reviewed_operation
from backup_service import create_backup, list_backups, restore_backup


ROOT = Path(DATA_ROOT).resolve()
WEB_ROOT = Path(__file__).resolve().parent / "web"
BACKUP_DIR = ROOT / "Data" / "Workbook Backups"
RESULTS_PATH = ROOT / "Data" / "Reports" / "Latest Results.json"
SERVER: ThreadingHTTPServer | None = None
OPERATION_LOCK = threading.RLock()
STOPPING = False


def configured_paths(settings=None):
    settings = settings or load_settings(ROOT)
    paths = [Path(value).expanduser().resolve() for value in settings["workbook_paths"]]
    if len(paths) != 2 or not all(path.is_file() for path in paths):
        raise ValueError("Choose the Plant 1 and Plant 2 workbooks in Settings first.")
    if paths[0] == paths[1] or paths[0].samefile(paths[1]):
        raise ValueError("Plant 1 and Plant 2 must use different workbooks.")
    return paths


def latest_result():
    if not RESULTS_PATH.is_file():
        return {"rows": [], "scan_complete": False, "created": 0, "review": 0}
    value = json.loads(RESULTS_PATH.read_text(encoding="utf-8"))
    if not isinstance(value, dict) or not isinstance(value.get("rows"), list):
        raise ValueError("The latest work list is unreadable. Refresh the workbooks to rebuild it.")
    return value


def refresh(paths=None):
    settings = load_settings(ROOT)
    reconcile_manual(ROOT)
    with run_lock(ROOT):
        recover_issuances(ROOT)
        inventory = scan_inventory(ROOT, persist=True)
    try:
        paths = paths or configured_paths(settings)
    except ValueError as exc:
        return {"rows": manual_rows(ROOT), "scan_complete": False, "created": 0,
                "review": 0, "setup_warning": str(exc), 'certificate_inventory': inventory}
    run(paths, ROOT, check_only=True, progress=None)
    return workspace_state(recheck_obsolete=False)


def workspace_state(recheck_obsolete=True):
    reconcile_manual(ROOT)
    report = latest_result()
    canonicalized = False
    for row in report.get("rows", []):
        status = status_value(row.get("certificate_status"))
        if row.get("source_kind") == "manual" and status == "generated":
            status = "generated_manual"
        if status != row.get("certificate_status"):
            row["certificate_status"] = status
            canonicalized = True
    if canonicalized:
        save_json(RESULTS_PATH, report)
    missing_pdf = any(
        row.get('certificate_status') in {'generated', 'generated_manual', 'completed'}
        and not find_certificate(ROOT / 'Certificates', {
            'cod_no': row.get('certificate_no'), 'source_id': row.get('plant'), 'case_id': row.get('sn')})
        for row in report.get('rows', []) if row.get('source_kind') != 'manual')
    if recheck_obsolete and (missing_pdf or any(
            obsolete_template_warning(row.get('detail', ''))
            or any(obsolete_template_warning(warning) for warning in (row.get('warnings') or []))
            for row in report.get('rows', []) if row.get('source_kind') != 'manual')):
        # Re-evaluate real data instead of just hiding the warning and leaving
        # stale attention flags (or incorrectly declaring the row valid).
        # Missing PDFs also invalidate cached Generated/Completed states.
        # Refresh is check-only; it cannot generate a certificate.
        return refresh()
    report['certificate_inventory'] = scan_inventory(ROOT)
    report["rows"] = [row for row in report.get("rows", []) if row.get("source_kind") != "manual"] + manual_rows(ROOT)
    return report


def preview_certificate(body):
    source_kind = body.get("source_kind", "workbook")
    if source_kind not in {"workbook", "manual"}:
        raise ValueError("Choose the plant workbook or local certificate register.")
    details = clean_details(body.get("details"))
    batch, cod, issued_date, warnings = prepare_certificate(
        ROOT, body.get("plant"), details, record_id=body.get("record_id") if source_kind == "manual" else None,
        source_kind=source_kind,
    )
    if source_kind == "workbook":
        paths = configured_paths()
        # Use the same staged material rows as final generation. Preview never
        # writes them, and corrections should not inherit stale source warnings.
        with WorkbookSession(paths[int(body["plant"][-1]) - 1], BACKUP_DIR) as session:
            _, staged = stage_certificate_details(session, details, validate_cells=False)
        batch.update(total_weight=staged['total_weight'], display_weight=staged['display_weight'],
                     certificate_weight=staged.get('certificate_weight'))
        if batch['display_weight'] and batch['display_weight'] > 0:
            warnings = [warning for warning in warnings if not warning.startswith('Weight is missing, invalid or not positive;')]
        warnings.extend(issue for issue in staged.get('issues', [])
                        if not (batch['certificate_type'] in {'sita', 'siemens'}
                                and issue == 'missing Form 6 number / manifest number'))
    identifier = uuid.uuid4().hex
    destination = ROOT / "Data" / "Previews" / f"{identifier}.pdf"
    make_template_certificate(batch, cod, issued_date, destination, allow_warnings=True)
    import pypdfium2 as pdfium
    with pdfium.PdfDocument(str(destination)) as document:
        page = document[0]
        bitmap = page.render(scale=1.5)
        try:
            bitmap.to_pil().save(destination.with_suffix('.png'))
        finally:
            bitmap.close()
            page.close()
    return {"preview_url": f"/previews/{identifier}.pdf", 'preview_image_url': f'/previews/{identifier}.png', "certificate_no": cod,
            "warnings": list(dict.fromkeys(warnings))}


def selected_key(plant, case_id):
    if plant not in {"plant1", "plant2"} or not str(case_id or "").strip():
        raise ValueError("Choose a plant and enter a Client ID / SN.")
    return plant, str(case_id).strip()


def target_row(report, selected):
    return next((row for row in report.get("rows", [])
                 if row.get("plant") == selected[0]
                 and str(row.get("sn", "")) == selected[1]), None)


def generate(*, selected=None, plants=None, custom_cod=None, confirmed=False, expected_revision=None, override_reason=''):
    settings = load_settings(ROOT)
    paths = configured_paths(settings)
    preflight = {}
    preflight_args = {"check_only": True, "progress": None, 'report_sink': preflight}
    if expected_revision is not None:
        preflight_args['expected_revision'] = expected_revision
    if selected:
        preflight_args["selected"] = selected
    if plants is not None:
        preflight_args["plants"] = plants
    if custom_cod:
        preflight_args["custom_cod"] = custom_cod
    run(paths, ROOT, **preflight_args)
    preflight = preflight or latest_result()
    candidates = [row for row in preflight.get("rows", [])
                  if row.get("certificate_status") == "ready"
                  and row.get("outcome") == "needs_review"]
    if selected and target_row(preflight, selected) is None:
        raise ValueError(f"Client ID {selected[1]} was not found in {selected[0].replace('plant', 'Plant ')}.")
    if candidates and not confirmed:
        return {"requires_confirmation": True, "warnings": candidates,
                "state": refresh(paths)}

    generated_report = {}
    arguments = {"progress": None, "allow_warnings": confirmed, 'override_reason': override_reason,
                 'report_sink': generated_report}
    if expected_revision is not None:
        arguments['expected_revision'] = expected_revision
    if selected:
        arguments["selected"] = selected
    if plants is not None:
        arguments["plants"] = plants
    if custom_cod:
        arguments["custom_cod"] = custom_cod
    result = run(paths, ROOT, **arguments)
    generated_report = generated_report or latest_result()
    generated_rows = {
        (row.get("plant"), str(row.get("sn"))): row
        for row in generated_report.get("rows", [])
        if row.get("outcome") == "generated_now"
    }
    state = post_commit_state(lambda: refresh(paths))
    for row in state.get("rows", []):
        generated_row = generated_rows.get((row.get("plant"), str(row.get("sn"))))
        if generated_row:
            row["detail"] = generated_row.get("detail", row.get("detail", ""))
            row["warnings"] = generated_row.get("warnings", [])
            if generated_row.get("system_alert"):
                row["system_alert"] = True
    state["created"] = result[0]
    return {"requires_confirmation": False, "state": state,
            "created": result[0], "review": result[1],
            'warnings': [state['setup_warning']] if state.get('setup_warning') else [],
            'certificates': certificate_outputs(list(generated_rows.values()))}


def create_certificate(body):
    """Preflight form edits without writing Excel; commit only with the final PDF."""
    plant = body.get('plant')
    details = clean_details(body.get('details'))
    selected = selected_key(plant, details.get('case_id'))
    scope = body.get('edit_scope', 'update_excel')
    if scope not in {'certificate_only', 'update_excel'}:
        raise ValueError('Choose This certificate only or Also update Excel.')
    if body.get('source_sn') is not None:
        if str(body['source_sn']) != selected[1] or body.get('source_plant') != plant:
            raise ValueError('The selected inward ID and plant cannot be changed. Open the correct inward instead.')
    paths = configured_paths()
    snapshot = {}
    arguments = {'selected': selected, 'custom_cod': details.get('cod_no') or None,
                 'details': details, 'update_excel': scope == 'update_excel', 'progress': None, 'report_sink': snapshot}
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    run(paths, ROOT, check_only=True, **arguments)
    preflight = dict(snapshot) if snapshot else latest_result()
    row = target_row(preflight, selected)
    if row is None:
        error = next((item.get('detail') for item in preflight.get('rows', []) if item.get('system_alert')), None)
        raise ValueError(error or 'The certificate details could not be checked.')
    if row.get('outcome') == 'needs_review' and not body.get('confirmed'):
        return {'requires_confirmation': True, 'warnings': [row], 'created': 0}
    if body.get('expected_revision') is not None:
        arguments['expected_revision'] = body['expected_revision']
    result = run(paths, ROOT, allow_warnings=bool(body.get('confirmed')),
                 override_reason=body.get('override_reason', ''), **arguments)
    final_report = snapshot or latest_result()
    final_row = target_row(final_report, selected)
    if not result[0]:
        raise ValueError((final_row or {}).get('detail') or 'The certificate could not be created. Excel details were not saved.')
    state = post_commit_state(lambda: refresh(paths))
    state_row = target_row(state, selected)
    if state_row and final_row:
        state_row['detail'] = final_row['detail'] + (' Certificate-only changes: original inward details were kept.'
                                                  if scope == 'certificate_only' else ' Edited details were also saved to Excel.')
        state_row['warnings'] = final_row.get('warnings', [])
    return {'requires_confirmation': False, 'state': state, 'created': result[0], 'review': result[1],
            'warnings': [state['setup_warning']] if state.get('setup_warning') else [],
            'certificates': certificate_outputs([final_row] if final_row else [])}


def post_commit_state(read_state):
    try:
        return read_state()
    except Exception as exc:
        try:
            state = latest_result()
        except Exception:
            state = {'rows': [], 'scan_complete': False}
        return dict(state, setup_warning=f'Certificate saved successfully, but the dashboard could not refresh: {exc}. Do not create it again; refresh after correcting the issue.')


def certificate_outputs(rows):
    outputs = []
    for row in rows:
        receipt = None
        if row.get('pdf_filename'):
            receipt = {'pdf_path': 'Certificates/' + row['pdf_filename'], 'cod': row.get('certificate_no')}
        else:
            try:
                receipt = next((entry for entry in reversed(audit_records(ROOT)) if entry.get('state') in {'complete', 'source_committed'}
                                and entry.get('source_kind') == row.get('source_kind', 'workbook')
                                and entry.get('plant') == row.get('plant')
                                and str(entry.get('key')) == str(row.get('record_id') or row.get('sn'))), None)
            except (OSError, ValueError):
                pass
        if receipt:
            filename = Path(receipt['pdf_path']).name
            outputs.append({'filename': filename, 'certificate_no': receipt['cod'],
                            'plant': row.get('plant'), 'supplier': row.get('supplier'),
                            'pdf_url': '/api/certificates/' + quote(filename, safe='')})
    return outputs


def open_path(path):
    opener = getattr(os, "startfile", None)
    if opener is None:
        raise RuntimeError("Opening local folders is available from Windows Explorer on this computer.")
    opener(str(path))


class Handler(BaseHTTPRequestHandler):
    server_version = "EcorecoLocal/1.0"

    def log_message(self, _format, *_args):
        return

    def send_json(self, value, status=200):
        if isinstance(value, dict) and status < 400:
            value = dict(value)
            # Reads can fail while Excel is open; do not misreport an already
            # committed certificate as a failure just to obtain a UI snapshot.
            try:
                value['workspace_revision'] = workspace_revision(ROOT)
            except (OSError, ValueError):
                value.setdefault('workspace_revision', None)
        payload = json.dumps(value, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(payload)

    def send_error_json(self, exc, status=400):
        self.send_json({"error": str(exc)}, status)

    def do_GET(self):
        with OPERATION_LOCK:
            self._do_get()

    def _do_get(self):
        path = urlsplit(self.path).path
        if path == "/api/state":
            try:
                settings = load_settings(ROOT)
                self.send_json({
                    "state": workspace_state(),
                    "settings": {
                        "workbook_paths": settings["workbook_paths"],
                        "reminders_enabled": settings["reminders_enabled"],
                        "reminder_time": settings["reminder_time"],
                        'backup_destination': settings.get('backup_destination', ''),
                    },
                    "certificate_folder": str(ROOT / "Certificates"),
                    "server_time": datetime.now().isoformat(timespec="minutes"),
                })
            except Exception as exc:
                self.send_error_json(exc, 500)
            return

        if path == "/api/templates":
            try:
                self.send_json({"templates": template_specs()})
            except Exception as exc:
                self.send_error_json(exc, 500)
            return
        if path == '/api/history':
            try:
                self.send_json({'history': audit_records(ROOT)[-50:]})
            except (OSError, ValueError) as exc:
                self.send_error_json(exc)
            return
        if path.startswith('/api/certificates/'):
            name = unquote(path[len('/api/certificates/'):])
            if Path(name).name != name or '/' in name or '\\' in name or not name.lower().endswith('.pdf'):
                self.send_error(404)
                return
            destination = ROOT / 'Certificates' / name
            if not valid_pdf(destination):
                self.send_error_json('This PDF is missing or unreadable. Refresh and recreate it.', 404)
                return
            payload = destination.read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'application/pdf')
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(payload)
            return
        preview_match = re.fullmatch(r"/previews/([a-f0-9]{32})\.(pdf|png)", path)
        if preview_match:
            extension = preview_match[2]
            destination = ROOT / "Data" / "Previews" / f"{preview_match[1]}.{extension}"
            if not destination.is_file():
                self.send_error(404)
                return
            payload = destination.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", 'application/pdf' if extension == 'pdf' else 'image/png')
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(payload)
            return

        assets = {"/": ("index.html", "text/html; charset=utf-8"),
                  "/styles.css": ("styles.css", "text/css; charset=utf-8"),
                  "/app.js": ("app.js", "text/javascript; charset=utf-8")}
        asset = assets.get(path)
        if not asset:
            self.send_error(404)
            return
        name, mime = asset
        try:
            payload = (WEB_ROOT / name).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            self.wfile.write(payload)
        except OSError as exc:
            self.send_error_json(exc, 500)

    def do_POST(self):
        with OPERATION_LOCK:
            if STOPPING:
                self.send_error_json('The dashboard is closing. Reopen it before starting another action.', 503)
                return
            self._do_post()

    def _do_post(self):
        global SERVER
        global STOPPING
        parsed = urlsplit(self.path)
        if not self._same_origin():
            self.send_error_json("This local app only accepts requests from its own page.", 403)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 0 or length > 1_000_000:
                raise ValueError("The request is too large.")
            body = json.loads(self.rfile.read(length) or b"{}")
            if not isinstance(body, dict):
                raise ValueError("Expected an object of request values.")

            if parsed.path == "/api/check":
                result = refresh()
                self.send_json({"state": result})
            elif parsed.path == "/api/preview":
                self.send_json(preview_certificate(body))
            elif parsed.path == "/api/manual/save":
                require_revision(ROOT, body.get('expected_revision'))
                identifier = save_draft(ROOT, body.get("plant"), body.get("details"), body.get("record_id"),
                                        expected_revision=body.get('expected_revision'))
                self.send_json({"record_id": identifier, "state": workspace_state()})
            elif parsed.path == "/api/manual/generate":
                response = reviewed_operation(ROOT, parsed.path, body, lambda values: generate_draft(
                    ROOT, values.get('plant'), values.get('details'), values.get('record_id'),
                    confirmed=bool(values.get('confirmed')), override_reason=values.get('override_reason', ''),
                    expected_revision=values.get('expected_revision')))
                response["state"] = post_commit_state(workspace_state) if response.get('created') else workspace_state()
                if response.get('created'):
                    if response['state'].get('setup_warning'):
                        response.setdefault('warnings', []).append(response['state']['setup_warning'])
                    response['certificates'] = certificate_outputs([{'source_kind': 'manual',
                        'record_id': response['record_id'], 'plant': body.get('plant'),
                        'pdf_filename': response.get('filename'), 'certificate_no': response.get('certificate_no'),
                        'supplier': (body.get('details') or {}).get('supplier')}])
                self.send_json(response)
            elif parsed.path == "/api/manual/complete":
                require_revision(ROOT, body.get('expected_revision'))
                complete_draft(ROOT, body.get("record_id"), expected_revision=body.get('expected_revision'))
                self.send_json({"state": workspace_state()})
            elif parsed.path == "/api/generate":
                selected = None
                if body.get("plant") or body.get("sn"):
                    selected = selected_key(body.get("plant"), body.get("sn"))
                plants = body.get("plants")
                self.send_json(reviewed_operation(ROOT, parsed.path, body, lambda values: generate(
                    selected=selected, plants=plants, custom_cod=str(values.get('custom_cod') or '').strip() or None,
                    confirmed=bool(values.get('confirmed')), expected_revision=values.get('expected_revision'),
                    override_reason=values.get('override_reason', ''))))
            elif parsed.path == "/api/create":
                self.send_json(reviewed_operation(ROOT, parsed.path, body, create_certificate))
            elif parsed.path == "/api/pick-annexure":
                require_revision(ROOT, body.get('expected_revision'))
                import tkinter as tk
                from tkinter import filedialog
                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes("-topmost", True)
                    picked = filedialog.askopenfilename(parent=picker_root, title="Choose an annexure PDF",
                                                       filetypes=[("PDF files", "*.pdf")])
                finally:
                    picker_root.destroy()
                attachment = None
                if picked:
                    # A workbook can change while the native picker is open.
                    require_revision(ROOT, body.get('expected_revision'))
                    source = Path(picked)
                    if source.suffix.casefold() != ".pdf" or source.stat().st_size > 50_000_000:
                        raise ValueError("Choose a PDF smaller than 50 MB.")
                    reader = PdfReader(source)
                    if reader.is_encrypted or not reader.pages:
                        raise ValueError("Choose an unlocked PDF with at least one page.")
                    identifier = uuid.uuid4().hex
                    folder = ROOT / "Data" / "Attachments"
                    folder.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, folder / f"{identifier}.pdf")
                    attachment = {"id": identifier, "name": source.name, "pages": len(reader.pages)}
                    save_json(folder / f"{identifier}.json", attachment)
                self.send_json({"annexure": attachment})
            elif parsed.path == "/api/complete":
                with run_lock(ROOT):
                    recover_issuances(ROOT)
                    require_revision(ROOT, body.get('expected_revision'))
                    selected = selected_key(body.get("plant"), body.get("sn"))
                    row = target_row(latest_result(), selected)
                    if not row or not row.get("file"):
                        raise ValueError("Refresh the workbooks and select the row again.")
                    if not find_certificate(ROOT / 'Certificates', {'cod_no': row.get('certificate_no'),
                            'source_id': row['plant'], 'case_id': row['sn']}):
                        raise ValueError('The current PDF is missing, changed or unreadable. Refresh and recreate it before marking it Completed.')
                    update_issued_status(
                        row["file"], BACKUP_DIR, row["sn"], row["sheet"], row["row"],
                        row.get("certificate_no", ""), row.get("certificate_status", "Generated"), "Completed",
                    )
                self.send_json({"state": refresh()})
            elif parsed.path == "/api/open-folder":
                folder = ROOT / "Certificates"
                folder.mkdir(parents=True, exist_ok=True)
                open_path(folder)
                self.send_json({"ok": True})
            elif parsed.path == "/api/open-workbook":
                selected = selected_key(body.get("plant"), body.get("sn"))
                row = target_row(latest_result(), selected)
                if not row or not row.get("file"):
                    raise ValueError("Refresh the workbooks and select the row again.")
                open_path(row["file"])
                self.send_json({"ok": True})
            elif parsed.path == "/api/pick-workbook":
                import tkinter as tk
                from tkinter import filedialog

                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes("-topmost", True)
                    picked = filedialog.askopenfilename(
                        parent=picker_root,
                        title="Choose a plant Excel workbook",
                        filetypes=(("Excel workbooks", "*.xlsx *.xlsm"),),
                    )
                finally:
                    picker_root.destroy()
                if not picked:
                    self.send_json({"cancelled": True})
                    return
                picked_path = Path(picked).expanduser().resolve()
                settings = load_settings(ROOT)
                configured = [Path(value).expanduser().resolve() for value in settings["workbook_paths"]]
                plant = next((f"plant{index + 1}" for index, path in enumerate(configured)
                              if path.is_file() and path.samefile(picked_path)), None)
                if not plant and not body.get("for_settings"):
                    raise ValueError("That workbook is not configured yet. Add its path in Settings first.")
                self.send_json({"path": str(picked_path), "plant": plant})
            elif parsed.path == "/api/settings":
                settings = load_settings(ROOT)
                paths = body.get("workbook_paths", settings["workbook_paths"])
                if not isinstance(paths, list) or len(paths) != 2 or any(not isinstance(item, str) or not item.strip() for item in paths):
                    raise ValueError("Enter both Plant 1 and Plant 2 workbook paths.")
                resolved = [Path(item).expanduser().resolve() for item in paths]
                if not all(path.is_file() for path in resolved):
                    raise ValueError("One of those workbooks could not be found.")
                if resolved[0] == resolved[1] or resolved[0].samefile(resolved[1]):
                    raise ValueError("Plant 1 and Plant 2 must use different workbooks.")
                reminder_time = str(body.get("reminder_time", settings["reminder_time"]))
                if len(reminder_time) != 5 or reminder_time[2] != ":":
                    raise ValueError("Choose a valid reminder time.")
                try:
                    datetime.strptime(reminder_time, "%H:%M")
                except ValueError as exc:
                    raise ValueError("Choose a valid reminder time.") from exc
                enabled = body.get("reminders_enabled", settings["reminders_enabled"])
                if not isinstance(enabled, bool):
                    raise ValueError("Reminder enabled must be on or off.")
                settings.update(workbook_paths=[str(path) for path in resolved],
                                reminder_time=reminder_time, reminders_enabled=enabled)
                save_json(ROOT / "settings.local.json", settings)
                message = "Settings saved."
                if body.get("update_reminder"):
                    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                               "-File", str(ROOT / "App" / "Install Daily Task.ps1"),
                               "-Mode", "Install" if enabled else "Disable"]
                    if enabled:
                        command.extend(["-ReminderTime", reminder_time])
                    completed = subprocess.run(command, capture_output=True, text=True,
                                               timeout=60, check=False)
                    if completed.returncode:
                        raise RuntimeError((completed.stdout or completed.stderr).strip()
                                            or "Windows could not update the daily reminder.")
                    message = "Settings and daily reminder updated."
                self.send_json({"settings": {
                    "workbook_paths": settings["workbook_paths"],
                    "reminders_enabled": enabled, "reminder_time": reminder_time,
                    'backup_destination': settings.get('backup_destination', ''),
                }, "message": message})
            elif parsed.path == '/api/backup/pick-destination':
                import tkinter as tk
                from tkinter import filedialog
                picker_root = tk.Tk()
                picker_root.withdraw()
                try:
                    picker_root.attributes('-topmost', True)
                    picked = filedialog.askdirectory(parent=picker_root, title='Choose a separate backup folder')
                finally:
                    picker_root.destroy()
                if not picked:
                    self.send_json({'cancelled': True})
                    return
                chosen = Path(picked).resolve()
                if chosen == ROOT or chosen.is_relative_to(ROOT / 'Certificates') or chosen.is_relative_to(ROOT / 'Data') or chosen.is_relative_to(ROOT / 'App'):
                    raise ValueError('Choose a separate folder, outside this app and its certificate data.')
                settings = load_settings(ROOT)
                settings['backup_destination'] = str(chosen)
                with run_lock(ROOT):
                    save_json(ROOT / 'settings.local.json', settings)
                self.send_json({'backup_destination': str(chosen), 'settings': settings,
                                'message': 'Backup folder saved. A separate drive is recommended.'})
            elif parsed.path == '/api/backup/create':
                result = create_backup(ROOT)
                self.send_json({'backup': result, 'backup_destination': load_settings(ROOT).get('backup_destination', ''),
                                'message': 'Verified backup created. Keep the backup and its trust files in a safe separate location.',
                                'health': '\n'.join(result.get('warnings', []))})
            elif parsed.path == '/api/backup/list':
                self.send_json({'backups': list_backups(ROOT),
                                'backup_destination': load_settings(ROOT).get('backup_destination', '')})
            elif parsed.path == '/api/backup/restore':
                with run_lock(ROOT):
                    require_revision(ROOT, body.get('expected_revision'))
                    result = restore_backup(ROOT, body.get('backup_id'), confirmed=body.get('confirmed') is True, locked=True)
                self.send_json({'restored': result, 'state': refresh(), 'settings': load_settings(ROOT),
                                'message': 'Backup restored. The previous data was preserved in a recovery snapshot.'})
            elif parsed.path == "/api/stop":
                STOPPING = True
                self.send_json({"ok": True})
                if SERVER:
                    threading.Thread(target=SERVER.shutdown, daemon=True).start()
            else:
                self.send_error_json("Unknown action.", 404)
        except (ValueError, OSError, RuntimeError, KeyError, json.JSONDecodeError) as exc:
            self.send_error_json(exc)
        except Exception as exc:
            self.send_error_json(exc, 500)

    def _same_origin(self):
        host = self.headers.get("Host", "").split(":", 1)[0].casefold()
        if host not in {"127.0.0.1", "localhost"}:
            return False
        origin = self.headers.get("Origin")
        if not origin:
            return True
        parsed = urlsplit(origin)
        return parsed.hostname in {"127.0.0.1", "localhost"} and parsed.port == self.server.server_port


def main():
    global SERVER
    for port in range(8765, 8776):
        try:
            connection = HTTPConnection("127.0.0.1", port, timeout=0.6)
            connection.request("GET", "/api/state")
            response = connection.getresponse()
            is_existing_dashboard = response.status == 200 and "application/json" in response.getheader("Content-Type", "")
            response.read()
            connection.close()
            if is_existing_dashboard:
                webbrowser.open(f"http://127.0.0.1:{port}/")
                return
        except OSError:
            pass
        try:
            SERVER = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if SERVER is None:
        raise RuntimeError("Could not start the local dashboard. Close any old certificate dashboard and try again.")
    # Wait for request workers before exit: a second tab cannot terminate an
    # in-progress publication halfway through a workbook/register update.
    SERVER.daemon_threads = False
    url = f"http://127.0.0.1:{SERVER.server_port}/"
    threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    SERVER.serve_forever(poll_interval=0.25)
    SERVER.server_close()


if __name__ == "__main__":
    main()
