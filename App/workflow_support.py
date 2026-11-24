"""Settings, safe run locking, reports and focused user reminders."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime, timedelta
import base64
from html import unescape
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox
import uuid

from app_paths import resource
from certificate_core import ALIASES, certificate_reminder_due, parse_date, read_batches, status_value
from workbook_store import LockBusyError, exclusive_file


def operation_error_message(error):
    """Keep resource failures and exceptions without text useful to the operator."""
    if isinstance(error, MemoryError):
        return ('Not enough available memory to finish. Close unused apps, '
                'then refresh and try again.')
    return str(error).strip() or 'The operation could not finish. Refresh and try again.'


def load_settings(root):
    root = Path(root)
    defaults_path = resource(root, "settings.example.json")
    local_path = root / "settings.local.json"

    def read(path):
        if not path.exists():
            return {}
        value = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(value, dict):
            raise ValueError(f"{path.name} must contain a settings object.")
        return value

    merged = read(defaults_path)
    merged.update(read(local_path))
    allowed = {"workbook_paths", "certificate_number_prefix", "certificate_number_start",
               "certificate_number_width", "next_cod", "reminder_time", "reminders_enabled", "backup_destination"}
    settings = {key: value for key, value in merged.items() if key in allowed}
    paths = settings.get("workbook_paths", ["", ""])
    if not isinstance(paths, list) or len(paths) > 2 or any(not isinstance(item, str) for item in paths):
        raise ValueError("Settings need the two plant Excel file paths.")
    settings["workbook_paths"] = (paths + ["", ""])[:2]
    for key, default, low, high in (("certificate_number_start", 1, 1, 999999999999),
                                    ("certificate_number_width", 4, 1, 18)):
        value = settings.get(key, default)
        if type(value) is not int or not low <= value <= high:
            raise ValueError(f"{key} must be a whole number from {low} to {high}.")
        settings[key] = value
    if settings.get('next_cod') is not None and (type(settings['next_cod']) is not int
                                               or not 1 <= settings['next_cod'] <= 999999999999):
        raise ValueError('The next COD must be a whole number from 1 to 999999999999.')
    prefix = settings.get("certificate_number_prefix", "")
    if not isinstance(prefix, str) or any(ord(ch) < 32 for ch in prefix):
        raise ValueError("The COD prefix must be plain text.")
    settings["certificate_number_prefix"] = prefix
    reminder_time = settings.get("reminder_time", "11:00")
    if not isinstance(reminder_time, str) or not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", reminder_time):
        raise ValueError("Choose a valid reminder time, such as 11:00 AM.")
    settings["reminder_time"] = reminder_time
    enabled = settings.get("reminders_enabled", True)
    if not isinstance(enabled, bool):
        raise ValueError("reminders_enabled must be true or false.")
    settings["reminders_enabled"] = enabled
    destination = settings.get('backup_destination', '')
    if not isinstance(destination, str) or len(destination) > 32000:
        raise ValueError('Choose a valid backup folder.')
    settings['backup_destination'] = destination
    return settings


def save_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temp.open('w', encoding='utf-8', newline='') as stream:
            stream.write(json.dumps(value, indent=2, ensure_ascii=False, default=str))
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


@contextmanager
def run_lock(root):
    lock = Path(root) / "Data" / "run.lock"
    lock.parent.mkdir(parents=True, exist_ok=True)
    try:
        with lock.open("xb"):
            pass
    except FileExistsError:
        pass
    context = exclusive_file(lock, wait_timeout=0)
    try:
        context.__enter__()
    except LockBusyError as exc:
        raise LockBusyError('Another certificate or backup operation is running. '
                            'Wait for it to finish, then try again.') from exc
    try:
        yield
    finally:
        context.__exit__(None, None, None)


class Report:
    def __init__(self, root, as_of, mode, publish_latest=True):
        self.root = Path(root)
        self.as_of = as_of
        self.mode = mode
        self.publish_latest = publish_latest
        self.output = self.root / "Data" / "Reports"
        self.output.mkdir(parents=True, exist_ok=True)
        self.run_id = uuid.uuid4().hex
        self.observed_sources = []
        self.scan_complete = False
        self.pending_recovery = ''
        self.rows = []
        self.created = 0
        self.maintenance_notes = []
        self.reserved_cod_numbers = []
        self.unassigned_rows = []

    def add(self, batch, outcome, detail, issue_code=None, system_alert=False, warnings=None, **_legacy):
        is_system = not str(batch.get('case_id') or '').strip()
        if batch.get("_missing_pdf") and outcome != "generated_now":
            cod = str(batch.get("cod_no") or "").strip()
            recovery = (f"Keep COD {cod} when recreating it." if cod else
                        "A new COD will be assigned when you create it.")
            status_note = ('Ready to recreate; Excel status update is pending.'
                           if batch.get('_workbook_read_only') or batch.get('_preflight_read_only') else 'Excel was changed to Ready.')
            if batch.get('pdf_recovery_pending'):
                status_note = 'Still Ready to recreate.'
            detail = f"PDF missing from Certificates: {status_note} {recovery} {detail}"
            system_alert = True
        raw_date = batch.get('raw_inward_date', batch.get("inward_date"))
        if batch.get('inward_date') is not None:
            raw_date = batch['inward_date']
        inward_date = parse_date(raw_date)
        if raw_date is None or (isinstance(raw_date, str) and not raw_date.strip()):
            quality = "missing"
        elif inward_date is None:
            quality = "invalid"
        elif inward_date > self.as_of:
            quality = "future"
        else:
            quality = "valid"
        age = (self.as_of - inward_date).days if inward_date else None
        state = batch.get("_effective_certificate_status", batch.get("issued", ""))
        reminder = dict(batch, inward_date=inward_date, _effective_certificate_status=state)
        due = (False if batch.get('recycling_completed') else
               certificate_reminder_due(reminder, self.as_of, days=25))
        status = status_value(state)
        if batch.get('recycling_completed') and status in {'generated', 'generated_manual'}:
            detail = 'Recycling has been sent; the requested Data Destruction certificate is still pending.'
        date_warning = {
            "missing": "Missing inward date; check the date in Excel.",
            "invalid": "Invalid inward date; check the date in Excel.",
            "future": "Inward date is in the future; check the date in Excel.",
        }.get(quality)
        date_needs_attention = bool(not is_system and date_warning and status not in {"na", "completed"})
        row_warnings = list(warnings or [])
        if batch.get('_workbook_read_only'):
            row_warnings.append('Excel is open or in use. Showing its saved contents; save and close it before the app writes changes.')
        if date_needs_attention and date_warning not in row_warnings:
            row_warnings.append(date_warning)
        row = {
            "source_kind": "system" if is_system else "workbook",
            'converted_record_id': batch.get('converted_record_id'),
            "plant": batch.get("source_id", "Setup"),
            "file": batch.get("source_path", ""),
            "sheet": batch.get("source_sheet", ""),
            "row": batch.get("source_row", ""),
            "sn": str(batch.get("case_id", "") or ""),
            "supplier": str(batch.get("supplier", "") or ""),
            "location": str(batch.get("location", "") or ""),
            "form6": str(batch.get("form6", "") or ""),
            "reference": str(batch.get("reference", "") or ""),
            **{key: str(batch.get(key, '') or '') for key in
               ('comments', 'transporter_details', 'docket_vehicle_no', 'weight_slip_scale')},
            "certificate_weight": batch.get("certificate_weight"),
            "certificate_type": str(batch.get("certificate_type") or "").strip() or "standard",
            "certificate_fields": dict(batch.get("certificate_fields") or {}),
            "display_weight": batch.get("display_weight", batch.get("total_weight")),
            "weight_source": batch.get("weight_source", "inward"),
            "items": list(batch.get("items", [])),
            'source_header_row': batch.get('source_header_row', 1),
            'column_mapping': dict(batch.get('column_mapping', {})),
            'import_notes': list(batch.get('import_notes', [])),
            'issuance_blocked': bool(batch.get('_issuance_blocked') or batch.get('blocking_issues')),
            'duplicate_serial_number': bool(batch.get('duplicate_serial_number')),
            'record_identity_conflict': batch.get('record_identity_conflict', ''),
            'workbook_read_only': bool(batch.get('_workbook_read_only')),
            "outcome": outcome,
            "detail": detail,
            "certificate_no": str(batch.get("cod_no", "") or ""),
            "external_pdf": batch.get('external_pdf'),
            "hdd_cod_no": batch.get('hdd_cod_no', ''),
            "certificate_date": parse_date(batch.get("cod_date")).isoformat() if parse_date(batch.get("cod_date")) else "",
            "certificate_status": status,
            "recycling_completed_pending_hdd": bool(batch.get('recycling_completed'))
                and status_value(state) in {'generated', 'generated_manual'},
            'pdf_filename': batch.get('pdf_filename', ''),
            'pdf_relative_path': batch.get('pdf_relative_path', ''),
            "inward_date": inward_date.isoformat() if inward_date else "",
            "inward_date_quality": quality,
            "age_days": age,
            "reminder_due": due,
            "overdue": bool(due and age is not None and age > 30),
            "pdf_missing": bool(batch.get('_missing_pdf')),
            "warnings": row_warnings,
        }
        if issue_code:
            row["issue_code"] = issue_code
        if system_alert or date_needs_attention or is_system:
            row["system_alert"] = True
        if is_system:
            row.update(inward_date_quality='unavailable', warnings=list(warnings or [detail]))
        self.rows.append(row)

    def add_saved_inward(self, batch, outcome, detail):
        # Saving an inward does not require the identifiers needed to issue its PDF.
        warnings = []
        if outcome == 'ready':
            if batch.get('certificate_type', 'standard') == 'standard' and not batch.get('form6'):
                warnings.append('Form 6 number is required to create the certificate.')
            if not str(batch.get('reference') or '').strip():
                warnings.append('Reference number is required to create the certificate.')
            if batch.get('certificate_type') == 'siemens':
                fields = batch.get('certificate_fields') or {}
                required = (fields.get('model_name'), fields.get('model_serial'), fields.get('tube_serial'),
                            fields.get('site_location') or batch.get('location'))
                if any(not str(value or '').strip() for value in required):
                    warnings.append('Medical equipment details are required to create the certificate.')
        if warnings:
            outcome = 'needs_review'
            detail += ' Open Review & edit to finish the certificate details.'
        self.add(batch, outcome, detail, warnings=warnings)

    def finish(self):
        needs_review = sum(row["outcome"] == "needs_review" for row in self.rows)
        result = {
            "run_id": self.run_id,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "as_of": self.as_of.isoformat(),
            "mode": self.mode,
            "observed_sources": self.observed_sources,
            "scan_complete": self.scan_complete,
            "created": self.created,
            "review": needs_review,
            "output": str(self.output),
            "certificate_folder": str(self.root / "Certificates"),
            "rows": self.rows,
            "maintenance_notes": self.maintenance_notes,
        }
        if hasattr(self, 'certificate_inventory'):
            result['certificate_inventory'] = self.certificate_inventory
        if hasattr(self, 'numbering'):
            result['numbering'] = self.numbering
        if self.pending_recovery:
            result.update(refresh_pending=True, recovery_note=self.pending_recovery)
        if hasattr(self, 'result_sink') and self.result_sink is not None:
            self.result_sink.clear()
            self.result_sink.update(result)
        if not self.publish_latest:
            return self.created, needs_review, self.output
        lines = [
            f"Mode: {self.mode}",
            f"Completed: {result['completed_at']}",
            f"Both plant files checked: {'Yes' if self.scan_complete else 'No'}",
            f"PDF certificates generated: {self.created}",
            f"Rows needing attention: {needs_review}",
            "",
        ]
        lines.extend(
            f"{row['plant']} | Serial Number {row['sn']} | {row['supplier']} | {row['outcome']}: {row['detail']}"
            for row in self.rows
        )
        json_path = self.output / "Latest Results.json"
        text_path = self.output / "Latest Results.txt"
        try:
            save_json(json_path, result)
            text_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        except OSError as exc:
            if self.created <= 0:
                raise
            # PDFs/source commits are already durable. A report-cache failure
            # must not turn them into a failed save or invite duplicate work.
            message = f'Certificates saved, but the work list report could not be written: {exc}. Refresh the app before creating more.'
            result['setup_warning'] = message
            if hasattr(self, 'result_sink') and self.result_sink is not None:
                self.result_sink['setup_warning'] = message
        return self.created, needs_review, self.output


def _actionable_rows(result):
    rows = result.get("rows", []) if isinstance(result, dict) else []
    return [row for row in rows if row.get("reminder_due") or row.get("system_alert")
            or row.get('attention_pending') or row.get('destruction_pending')]


def alert_workbook_targets(result, workbook_paths):
    """Use configured plant files, opening each affected plant only once."""
    paths = {f'plant{index}': Path(value).expanduser().resolve()
             for index, value in enumerate(workbook_paths, 1) if index <= 2 and value}
    affected = set()
    for row in _actionable_rows(result):
        if row.get('source_kind') == 'manual':
            link = row.get('destruction_source')
            if isinstance(link, dict) and not link.get('record_id') and link.get('plant') in paths:
                affected.add(link['plant'])
            continue
        if row.get('plant') in paths:
            affected.add(row['plant'])
        elif row.get('file'):
            source = Path(row['file']).expanduser().resolve()
            affected.update(plant for plant, path in paths.items() if source == path)
    return [(plant, path) for plant, path in paths.items() if plant in affected]


def open_alert_workbooks(result, workbook_paths):
    """Open affected Excel files normally, without reading or editing cells."""
    targets = alert_workbook_targets(result, workbook_paths)
    if not targets:
        raise FileNotFoundError('No plant Excel file is linked to these alerts.')
    opener = getattr(os, 'startfile', None)
    if opener is None:
        raise RuntimeError('Opening Excel files is available on Windows.')
    opened = set()
    missing = []
    for plant, path in targets:
        if not path.is_file():
            missing.append(f'Plant {plant[-1]}')
            continue
        if path not in opened:
            opener(str(path))
            opened.add(path)
    if missing:
        prefix = 'Available Excel files opened. ' if opened else ''
        raise FileNotFoundError(prefix + ', '.join(missing) + ' could not be found. Check its path in Settings.')
    return [path for _, path in targets]


def open_excel_workbook(workbook):
    """Open a workbook for manual repair without reading or changing its cells."""
    path = Path(workbook).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f'The plant workbook could not be found: {path}')
    opener = getattr(os, 'startfile', None)
    if opener is None:
        raise RuntimeError('Opening Excel files is available on Windows.')
    opener(str(path))
    return path


def open_excel_record(workbook, sheet_name, row_number, *, case_id=None, header_row=None,
                      allow_duplicate_repair=False, preferred_row=None):
    """Find the client's live ID and select its whole section without saving Excel."""
    path = Path(workbook).expanduser().resolve()
    if allow_duplicate_repair:
        # A repeated number cannot identify one inward reliably. Allow the
        # operator to repair the workbook without claiming a row was selected.
        return open_excel_workbook(path), str(sheet_name), None
    try:
        row_number = int(row_number)
    except (TypeError, ValueError) as exc:
        raise ValueError("The selected Excel row is not available. Refresh the work list and try again.") from exc
    if not path.is_file() or not str(sheet_name or "").strip() or row_number < 2:
        raise ValueError("The selected Excel record is no longer available. Refresh the work list and try again.")
    if case_id is None or header_row is None:
        matches = [batch for batch in read_batches(path) if batch['source_sheet'] == sheet_name
                   and batch['source_row'] == row_number
                   and (case_id is None or batch['case_id'] == str(case_id))]
        if len(matches) != 1:
            raise ValueError('The client moved or changed in Excel. Refresh and select it again.')
        case_id = matches[0]['case_id']
        header_row = matches[0].get('source_header_row',1)
    case_id = str(case_id).strip()
    if not case_id or not isinstance(header_row, int) or not 1 <= header_row < 1048576:
        raise ValueError('The selected Excel client is not available. Refresh and select it again.')
    try:
        preferred_row = int(preferred_row) if preferred_row is not None else None
    except (TypeError, ValueError) as exc:
        raise ValueError('The selected Excel row is no longer available. Refresh and select it again.') from exc
    if preferred_row is not None and preferred_row <= header_row:
        raise ValueError('The selected Excel row is no longer available. Refresh and select it again.')
    powershell = shutil.which("powershell.exe")
    if not powershell:
        raise OSError("Excel could not be opened at the selected row. Ask IT to check this computer's Excel setup.")
    encode = lambda value: base64.b64encode(str(value).encode("utf-8")).decode("ascii")
    aliases = json.dumps({key:sorted(ALIASES[key]) for key in ('sn','item','quantity','per_piece_weight','weight')})
    script = (
        "$ErrorActionPreference='Stop'; "
        f"$path=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{encode(path)}')); "
        f"$sheetName=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{encode(sheet_name)}')); "
        f"$clientId=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{encode(case_id)}')); "
        f"$aliases=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{encode(aliases)}')) | ConvertFrom-Json; "
        f"$headerRow={header_row}; "
        f"$preferredRow={preferred_row or 0}; "
        "try {$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')} "
        "catch {$excel=New-Object -ComObject Excel.Application}; "
        "Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class CertificateExcelFocus { "
        "[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int c); "
        "[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h); }'; "
        "$windowHandle=[IntPtr]::Zero; $deadline=[DateTime]::UtcNow.AddSeconds(18); "
        "while($true){try{ "
        "$excel.Visible=$true; $book=$null; "
        "if($windowHandle -eq [IntPtr]::Zero){$windowHandle=[IntPtr]$excel.Hwnd}; "
        "foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}; "
        "if($null -eq $book){$book=$excel.Workbooks.Open($path)}; "
        "$sheet=$book.Worksheets.Item($sheetName); "
        "$book.Activate() | Out-Null; $sheet.Activate() | Out-Null; "
        "$cells=$sheet.Cells; "
        "$lastColumn=$cells.Item($headerRow,16384).End(-4159).Column; "
        "if($lastColumn -lt 2){throw 'The Excel headings changed. Save the workbook and refresh the dashboard.'}; "
        "$headings=$sheet.Range($cells.Item($headerRow,1),$cells.Item($headerRow,$lastColumn)).Value2; "
        "$mapping=@{}; foreach($property in $aliases.PSObject.Properties){ "
        "$found=@(); for($col=1;$col -le $lastColumn;$col++){ "
        "$heading=([string]$headings[1,$col]).ToLowerInvariant() -replace '[^a-z0-9]',''; "
        "if($property.Value -contains $heading){$found+= $col}}; "
        "if($found.Count -gt 1){throw 'Duplicate Excel headings were found. Save the workbook and refresh.'}; "
        "if($found.Count -eq 1){$mapping[$property.Name]=$found[0]}}; "
        "if(-not $mapping.ContainsKey('sn')){throw 'The Serial Number heading moved. Save the workbook and refresh.'}; "
        "$lastRow=$sheet.UsedRange.Row+$sheet.UsedRange.Rows.Count-1; "
        "if($lastRow -le $headerRow){throw 'This client is no longer in the workbook.'}; "
        "$values=$sheet.Range($cells.Item($headerRow+1,1),$cells.Item($lastRow,$lastColumn)).Value2; "
        "$matchingRows=@(); for($index=1;$index -le $lastRow-$headerRow;$index++){ "
        "$serialRaw=$values[$index,$mapping['sn']]; $value=([string]$serialRaw).Trim(); "
        "if($serialRaw -is [double] -and [Math]::Truncate($serialRaw) -eq $serialRaw){ "
        "$serialFormat=[string]$cells.Item($headerRow+$index,$mapping['sn']).NumberFormat; "
        "if($serialFormat -match '^0+$'){ "
        "$value=$serialRaw.ToString('0',[Globalization.CultureInfo]::InvariantCulture); "
        "if($value.StartsWith('-')){$value='-'+$value.Substring(1).PadLeft([Math]::Max(0,$serialFormat.Length-1),'0')} "
        "else{$value=$value.PadLeft($serialFormat.Length,'0')}}}; "
        "if([string]::Equals($value,$clientId,[StringComparison]::OrdinalIgnoreCase)){$matchingRows+=$index}}; "
        "if($preferredRow -gt $headerRow){$preferredIndex=$preferredRow-$headerRow; "
        "if($matchingRows -notcontains $preferredIndex){throw 'The selected client moved in Excel. Refresh the work list and try again.'}; "
        "$firstIndex=$preferredIndex} elseif($matchingRows.Count -eq 1){$firstIndex=$matchingRows[0]} "
        "else{throw 'This Serial Number appears more than once. Open a specific list entry or correct the Serial Number in Excel.'}; "
        "$rowNumber=$headerRow+$firstIndex; "
        "$limit=$lastRow-$headerRow; for($index=$firstIndex+1;$index -le $limit;$index++){ "
        "if(([string]$values[$index,$mapping['sn']]).Trim()){$limit=$index-1;break}}; "
        "$idCell=$cells.Item($rowNumber,$mapping['sn']); "
        "$endRow=$idCell.MergeArea.Row+$idCell.MergeArea.Rows.Count-1; "
        "foreach($key in @('item','quantity','per_piece_weight')){if($mapping.ContainsKey($key)){ "
        "for($index=$firstIndex+1;$index -le $limit;$index++){ "
        "if(([string]$values[$index,$mapping[$key]]).Trim()){$endRow=[Math]::Max($endRow,$headerRow+$index)}}}}; "
        "if($mapping.ContainsKey('weight') -and $limit -gt $firstIndex){ "
        "$weights=$sheet.Range($cells.Item($headerRow+1,$mapping['weight']),$cells.Item($lastRow,$mapping['weight'])).Formula; "
        "for($index=$firstIndex+1;$index -le $limit;$index++){ "
        "$rawWeight=([string]$weights[$index,1]).Trim(); "
        "if($rawWeight -and -not $rawWeight.StartsWith('=')){$endRow=[Math]::Max($endRow,$headerRow+$index)}}}; "
        "$endRow=[Math]::Min([Math]::Max($endRow,$rowNumber),$headerRow+$limit); "
        "$target=$sheet.Range($cells.Item($rowNumber,1),$cells.Item($endRow,$lastColumn)); "
        "if($sheet.FilterMode -and $idCell.EntireRow.Hidden){$sheet.ShowAllData() | Out-Null}; "
        "if($idCell.EntireRow.Hidden){throw 'This client is in a hidden Excel row. Unhide it in Excel and try again.'}; "
        "$excel.Goto($target,$true) | Out-Null; $target.Select() | Out-Null; "
        "$excel.ActiveWindow.ScrollRow=[Math]::Max($headerRow,$rowNumber-3); $excel.ActiveWindow.ScrollColumn=1; "
        "if($excel.Selection.Address() -ne $target.Address() -or $excel.ActiveWorkbook.FullName -ine $path "
        "-or $excel.ActiveSheet.Name -ine $sheetName){ "
        "throw 'Excel could not confirm the selected client. Try Open in Excel again.'}; "
        "[CertificateExcelFocus]::ShowWindow([IntPtr]$excel.Hwnd,9) | Out-Null; "
        "[CertificateExcelFocus]::SetForegroundWindow([IntPtr]$excel.Hwnd) | Out-Null; "
        "if($book.ReadOnly){throw 'Excel opened this workbook read-only. Close the other copy or release the file lock, then reopen it to edit.'}; "
        "@{row=$rowNumber;last_row=$endRow;selected=$true} | ConvertTo-Json -Compress; break "
        "} catch { "
        # Excel can briefly return a null COM object while switching worksheets.
        "$busy=$_.FullyQualifiedErrorId -eq 'InvokeMethodOnNull'; "
        "for($cause=$_.Exception;$null -ne $cause;$cause=$cause.InnerException){ "
        "if($cause.HResult -in @(-2147418111,-2147417846,-2146777998)){$busy=$true;break}}; "
        "if(-not $busy){throw}; "
        "if($windowHandle -ne [IntPtr]::Zero){ "
        "[CertificateExcelFocus]::ShowWindow($windowHandle,9) | Out-Null; "
        "[CertificateExcelFocus]::SetForegroundWindow($windowHandle) | Out-Null}; "
        "if([DateTime]::UtcNow -ge $deadline){throw 'Excel is not ready to highlight this entry. Close any Excel message or finish editing the cell, then click Open in Excel again.'}; "
        "Start-Sleep -Milliseconds 350}}"
    )
    command = [powershell, "-NoLogo", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden",
               "-EncodedCommand", base64.b64encode(script.encode("utf-16le")).decode("ascii")]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=30,
                                   check=False, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired as exc:
        raise OSError('Excel is taking too long to respond. Finish any open message in Excel, then try Open in Excel again.') from exc
    if completed.returncode:
        message = (completed.stderr or completed.stdout or "Excel could not open the selected worksheet row.").strip()
        errors = re.findall(r'<S S="Error">(.*?)</S>', message, re.S)
        if errors:
            message = unescape(errors[0]).replace('_x000D__x000A_', '\n').strip()
        raise OSError(message)
    try:
        result = json.loads(completed.stdout.strip())
        selected_row = int(result['row'])
    except (ValueError, KeyError, TypeError) as exc:
        raise OSError('Excel did not confirm the selected client. Try Open in Excel again.') from exc
    if result.get('selected') is not True or not header_row < selected_row <= 1048576:
        raise OSError('Excel could not confirm the highlighted client. Try again.')
    return path, str(sheet_name), selected_row


def open_alert_workbook(result, selected_row=None):
    """Open the selected alert's plant workbook at its matching inward row."""
    actionable = _actionable_rows(result)
    row = selected_row if selected_row is not None else next((item for item in actionable if item.get("file")), None)
    if row is None:
        raise FileNotFoundError("No plant workbook is linked to the certificates in this reminder.")
    if selected_row is not None and not any(
            item.get("plant") == row.get("plant") and str(item.get("sn")) == str(row.get("sn"))
            and str(item.get("file")) == str(row.get("file")) for item in actionable):
        raise ValueError("That alert is no longer in the current work list. Refresh and choose it again.")
    workbook = Path(row["file"])
    if not workbook.is_file():
        raise FileNotFoundError(f"The plant workbook could not be found:\n{workbook}")
    matches = [batch for batch in read_batches(workbook)
               if str(batch.get("case_id")).strip().casefold() == str(row.get("sn")).strip().casefold()
               and (not row.get("sheet") or batch.get("source_sheet") == row.get("sheet"))]
    if not matches:
        raise ValueError("The selected client row changed in Excel. Refresh both plant workbooks and try again.")
    if len(matches) > 1:
        open_excel_workbook(workbook)
        return workbook
    open_excel_record(workbook, matches[0]["source_sheet"], matches[0]["source_row"],
                      case_id=matches[0]['case_id'], header_row=matches[0].get('source_header_row',1))
    return workbook


def _notice_text(rows):
    due = [row for row in rows if row.get("reminder_due")]
    urgent = [row for row in rows if not row.get("reminder_due")
              and (row.get('system_alert') or row.get('attention_pending') or row.get('destruction_pending'))]
    parts = []
    if due:
        parts.append(f"{len(due)} certificate{'s' if len(due) != 1 else ''} pending.")
    if urgent:
        parts.append(f"{len(urgent)} record{'s have' if len(urgent) != 1 else ' has'} details to check.")
    return "\n".join(parts + ["", "Choose Review now to see the list and what to do next."])


def notify_once(root, result, notifier, today=None, enabled=True, repeat_pending=False,
                reminder_time="11:00", record_only=False, reminder_only=False):
    """Show one centered daily summary for actionable reminders, not every row."""
    today = today or date.today()
    rows = _actionable_rows(result)
    if not rows:
        return False
    state_path = Path(root) / "Data" / "reminder-state.json"
    state = {}
    try:
        loaded = json.loads(state_path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            state = loaded
    except (OSError, ValueError):
        pass
    if (state.get('dismissed_day') == today.isoformat()
            and state.get('dismissed_time', reminder_time) == reminder_time):
        return False
    from certificate_templates import current_warnings
    signature = json.dumps(sorted(json.dumps({
        'source_kind': row.get('source_kind'), 'plant': row.get('plant'),
        'key': row.get('record_id') or row.get('sn'), 'status': row.get('certificate_status'),
        'age': row.get('age_days'), 'issue_code': row.get('issue_code'),
        'warnings': sorted(set(current_warnings(row.get('warnings', [])))),
        'detail': row.get('detail') if row.get('system_alert') and not row.get('warnings') and not row.get('pdf_missing') else None,
        'system_alert': bool(row.get('system_alert')), 'pdf_missing': bool(row.get('pdf_missing')),
        'pending_hdd': bool(row.get('attention_pending') or row.get('destruction_pending')),
    }, sort_keys=True, default=str) for row in rows))
    if (not repeat_pending and state.get("last_day") == today.isoformat()
            and state.get('last_time') == reminder_time and state.get("signature") == signature):
        return False
    if enabled and not record_only:
        notifier("Certificates needing attention", _notice_text(rows))
        # Record the alert only after it was successfully shown. If the desktop
        # notification fails, a later scheduled/sign-in check must retry it.
        # The modal's Dismiss callback may have just saved dismissal state.
        # Preserve it instead of writing the snapshot from before the dialog.
        try:
            latest = json.loads(state_path.read_text(encoding='utf-8'))
            if isinstance(latest, dict):
                state = latest
        except (OSError, ValueError):
            pass
        state.update(last_day=today.isoformat(), last_time=reminder_time, signature=signature)
        save_json(state_path, state)
        return True
    return False


def _powershell(root, mode, extra=()):
    script = Path(root) / "App" / "Install Daily Task.ps1"
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-Mode", mode]
    for key, value in extra:
        command.extend([key, str(value)])
    return subprocess.run(command, capture_output=True, text=True, timeout=30, check=False,
                          creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))


def schedule_one_hour_reminder(root, now=None):
    run_at = (now or datetime.now()) + timedelta(hours=1)
    completed = _powershell(root, "Snooze", (("-SnoozeAt", run_at.strftime("%Y-%m-%dT%H:%M:%S")),))
    if completed.returncode:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Windows could not schedule the reminder.")
    save_json(Path(root) / 'Data' / 'Snoozed Reminder.json', {'run_at': run_at.isoformat(timespec='seconds')})


def cancel_snoozed_reminder(root, force=False):
    marker = Path(root) / 'Data' / 'Snoozed Reminder.json'
    if not marker.is_file():
        # Dismissal state also suppresses any untracked legacy one-off task.
        # Ordinary daily alerts do not need a PowerShell process just to close.
        return False
    completed = _powershell(root, "CancelSnooze")
    if completed.returncode and not force:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Windows could not clear the reminder.")
    if completed.returncode == 0:
        marker.unlink(missing_ok=True)
        return True
    return False


def launch_review_app(root):
    script = Path(root) / "App" / "web_app.py"
    subprocess.Popen([sys.executable, str(script), '--attention'], cwd=str(script.parent), close_fds=True)


def dismiss_reminders_today(root, today=None):
    """Hide today's reminders without marking any certificate completed."""
    path = Path(root) / 'Data' / 'reminder-state.json'
    try:
        state = json.loads(path.read_text(encoding='utf-8'))
        if not isinstance(state, dict):
            state = {}
    except (OSError, ValueError):
        state = {}
    state['dismissed_day'] = (today or date.today()).isoformat()
    # Choosing a new reminder time is an explicit request to resume reminders.
    state['dismissed_time'] = load_settings(root)['reminder_time']
    save_json(path, state)
    # Persist dismissal even if Windows cannot remove a scheduled snooze.
    # That task also consults this state, so it will stay quiet for today.
    try:
        cancel_snoozed_reminder(root, force=True)
    except (OSError, RuntimeError, subprocess.SubprocessError):
        pass


def show_centered_reminder(title, message, parent=None, *, on_snooze=None,
                           on_dismiss=None, on_open_workbook=None,
                           on_open_automation=None, workbook_rows=None, workbook_plants=None,
                           error_root=None):
    own_root = parent is None
    window = tk.Tk() if own_root else tk.Toplevel(parent)
    window.title(title)
    has_excel = bool(on_open_workbook and workbook_plants)
    window.geometry('500x265' if has_excel else '500x220')
    window.resizable(False, False)
    window.transient(parent) if parent else None
    outer = tk.Frame(window, padx=18, pady=16)
    outer.pack(fill="both", expand=True)
    tk.Label(outer, text=title, font=("Segoe UI", 13, "bold"), anchor="w").pack(fill="x", pady=(0, 8))
    tk.Label(outer, text=message, font=("Segoe UI", 10), justify="left", anchor="nw", wraplength=460).pack(fill="both", expand=True)
    actions = tk.Frame(outer)
    actions.pack(fill="x", pady=(18, 0))
    result = {"action": "dismiss", "closed": False}

    def perform(callback, args, action):
        try:
            callback(*args)
            if error_root:
                try:
                    (Path(error_root) / 'Data' / 'Reports' / 'Reminder action error.txt').unlink(missing_ok=True)
                except OSError:
                    pass
        except Exception as exc:
            # The reminder has already closed. Keep a failed task update
            # recoverable without putting another modal in the user's way.
            import traceback
            if error_root:
                path = Path(error_root) / 'Data' / 'Reports' / 'Reminder action error.txt'
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(f'{action}: {exc}\n{traceback.format_exc()}', encoding='utf-8')
            else:
                if sys.stderr:
                    sys.stderr.write(f'Reminder {action} failed: {exc}\n')

    def finish(action, callback, *args):
        if result['closed']:
            return
        result.update(action=action, closed=True)
        window.destroy()
        if callback:
            # A non-daemon worker survives the scheduled process returning,
            # while the UI thread and window close immediately.
            import threading
            result['worker'] = threading.Thread(target=perform, args=(callback, args, action), daemon=False)
            result['worker'].start()

    tk.Button(actions, text="Review now", font=("Segoe UI", 10, "bold"), background='#29343c', foreground='white', command=lambda: finish("open_automation", on_open_automation), padx=10, pady=6).pack(side="left")
    tk.Button(actions, text="Remind me in 1 hour", font=("Segoe UI", 10), command=lambda: finish("snooze", on_snooze), padx=8, pady=6).pack(side="left", padx=8)
    tk.Button(actions, text="Dismiss for today", font=("Segoe UI", 10), command=lambda: finish("dismiss", on_dismiss), padx=8, pady=6).pack(side="right")
    if has_excel:
        excel_actions = tk.Frame(outer)
        excel_actions.pack(fill='x', pady=(12, 0))
        label = 'Open both Excel files' if len(set(workbook_plants)) > 1 else f'Open Plant {workbook_plants[0][-1]} Excel'
        tk.Button(excel_actions, text=label, font=('Segoe UI', 10),
                  command=lambda: finish('open_workbook', on_open_workbook), padx=10, pady=6).pack(side='left')
    window.protocol("WM_DELETE_WINDOW", lambda: finish("dismiss", on_dismiss))
    window.bind("<Escape>", lambda _: finish("dismiss", on_dismiss))
    window.update_idletasks()
    x = max(0, (window.winfo_screenwidth() - window.winfo_width()) // 2)
    y = max(0, (window.winfo_screenheight() - window.winfo_height()) // 2)
    window.geometry(f"+{x}+{y}")
    if own_root:
        window.mainloop()
    else:
        parent.wait_window(window)
    if result.get('worker'):
        # The window is gone; finish the hidden action before notify_once
        # records its summary, preserving a dismissal saved by the callback.
        result['worker'].join()
    return result["action"]
