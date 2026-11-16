"""Generate recycling certificates as PDFs from the two plant workbooks.

This workflow has no email or Outlook integration. Complete unselected rows become Ready;
generates its PDF, sends it through their normal process, and marks it Completed.
"""
from __future__ import annotations

from contextlib import ExitStack
from datetime import date
from io import BytesIO
import re
from pathlib import Path
import shutil
import sqlite3
import uuid
import json

from certificate_core import (
    certificate_layout,
    certificate_number,
    certificate_reminder_due,
    certificate_sequence,
    make_certificate,
    inward_details_complete,
    number,
    parse_date,
    read_batches,
    source_value_issues,
    status_value,
)
from workbook_store import LockBusyError, WorkbookSession, WorkbookSnapshot, stage_certificate_details
from workflow_support import Report, load_settings, run_lock, operation_error_message
from app_paths import resource
from certificate_templates import TEMPLATE_ROOT, template_id, validate_fields, layout_warnings
from certificate_inventory import (scan_inventory, cod_conflicts, verify_publication,
                                   issuance_paths, record_issuance, certificate_destination,
                                   unresolved_cod_entries, cod_recovery_message)
from issuance_safety import (IssuanceTransaction, recover_issuances, latest_issuance,
                             audit_records, workbook_receipt_conflict,
                             LOOKUP_RECEIPT, current_pdf_is_healthy)


def _safe_cod_name(value: str) -> str:
    """Make a COD safe as a Windows filename without changing the workbook COD."""
    raw = str(value).strip()
    raw_device_name = raw.split(".", 1)[0].rstrip(" .").upper()
    if raw_device_name in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", raw_device_name):
        raise ValueError("This COD is a reserved Windows filename. Enter a different COD number.")
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", raw).strip(" ._")
    if not safe or safe in {".", ".."}:
        raise ValueError("Enter a COD number that can be used as a PDF filename.")
    device_name = safe.split(".", 1)[0].upper()
    if device_name in {"CON", "PRN", "AUX", "NUL"} or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", device_name):
        raise ValueError("This COD is a reserved Windows filename. Enter a different COD number.")
    if len(safe) > 240:
        raise ValueError("This COD is too long to use as a PDF filename. Shorten it to 240 characters or fewer.")
    return safe


def _safe_name_part(value, fallback):
    """Keep the requested human-readable PDF name safe on Windows."""
    raw = re.sub(r"\s+", "_", str(value or "").strip())
    raw = re.sub(r"[^A-Za-z0-9._-]+", "_", raw)
    raw = re.sub(r"_+", "_", raw).strip(" ._")
    return raw or fallback


def _human_filename_section(value, fallback):
    """Keep names readable while removing Windows-forbidden filename characters."""
    raw = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", " ", str(value or ""))
    raw = re.sub(r"[\s_]+", " ", raw).strip(" .")
    return raw or fallback


def cod_identity(value):
    raw = str(value or "").strip()
    return str(int(raw)) if raw.isdecimal() and raw.isascii() else _safe_name_part(raw, "").casefold()


def workbook_cod_ownership_issue(batch, batches, manual_records=()):
    """An issued inward cannot adopt another current record's COD."""
    cod = str(batch.get('cod_no') or '').strip()
    if not cod:
        return ''
    identity = cod_identity(cod)
    own = (batch.get('source_id'), str(batch.get('case_id') or '').strip().casefold())
    owners = []
    for other in batches:
        key = (other.get('source_id'), str(other.get('case_id') or '').strip().casefold())
        if key != own and cod_identity(other.get('cod_no')) == identity:
            plant = str(other.get('source_id') or '').replace('plant', 'Plant ')
            owners.append(f'{plant}, sheet "{other.get("source_sheet")}", row {other.get("source_row")} '
                          f'(Serial Number {other.get("case_id")})')
    for record in manual_records:
        details = record.get('details') or {}
        if cod_identity(details.get('cod_no')) == identity:
            owners.append(f'the local certificate for {details.get("supplier") or "an unnamed company"}')
    if not owners:
        return ''
    return (f'COD {cod} is also assigned to ' + '; '.join(owners) +
            '. Each certificate needs a unique COD. Check the COD values before creating, correcting or completing a certificate.')


def certificate_filename(batch, plant_number=None, cod=None):
    """Return a PDF name with COD, client, and city/country, without a plant suffix."""
    cod = str(cod if cod is not None else batch.get("cod_no") or "").strip()
    if not cod:
        raise ValueError("Assign a COD number before naming the certificate PDF.")
    raw_company = str(batch.get("supplier") or "").strip()
    raw_company = re.sub(r"^plant[\s_-]*[12](?=$|[\s_-])[\s_-]*", "", raw_company, flags=re.I)
    company = _human_filename_section(raw_company, "Unknown Company")
    city = _human_filename_section(batch.get("location"), "Unknown City")
    cod_part = _safe_cod_name(cod)
    def units(value):
        return sum(2 if ord(character) > 0xffff else 1 for character in value)

    suffix = '_HDD' if batch.get('certificate_type') == 'destruction' else ''
    available = 240 - units(cod_part) - 2 - len(suffix)
    if available < 2:
        raise ValueError("This COD is too long to include the company and city in the PDF filename.")

    def shorten(value, limit):
        kept = []
        used = 0
        for character in value:
            cost = 2 if ord(character) > 0xffff else 1
            if used + cost > limit:
                break
            kept.append(character)
            used += cost
        return ''.join(kept).rstrip(' .')

    if units(company) + units(city) > available:
        city_budget = min(units(city), max(1, available // 2))
        company_budget = min(units(company), available - city_budget)
        city_budget = available - company_budget
        # The source and printed values remain complete. Only the filename's
        # descriptive sections share the Windows component-length budget.
        company = shorten(company, company_budget) or 'C'
        city = shorten(city, city_budget) or 'L'
    return f"{cod_part}_{company}_{city}{suffix}.pdf"


def _read_sources(paths, sessions, report):
    batches = []
    for index, path in enumerate(paths, 1):
        source_id = f"plant{index}"
        if sessions[index - 1] is None:
            continue
        try:
            source_batches = list(read_batches(BytesIO(sessions[index - 1].data)))
        except Exception as exc:
            report.add({"source_id": source_id, "source_path": str(path)}, "needs_review",
                       f"{path.name} could not be read: {exc}", issue_code="source_unavailable", system_alert=True)
            continue
        report.observed_sources.append(source_id)
        for batch in source_batches:
            batch.update(source_id=source_id, source_path=str(path))
            if isinstance(sessions[index - 1], WorkbookSnapshot):
                batch['_workbook_read_only'] = True
            batches.append(batch)
    report.scan_complete = len(report.observed_sources) == 2
    return batches


def _highest_sequence(batches, settings, output_folder, legacy_history=None, inventory=None):
    highest = max(0, int(settings.get("certificate_number_start", 1)) - 1)
    for batch in batches:
        for key in ('cod_no', 'hdd_cod_no'):
            if batch.get(key):
                highest = max(highest, certificate_sequence(batch[key], settings))
    # Existing flat output PDFs are the durable numbering floor if an old
    # workbook row is archived or its COD cell is cleared by the user.
    output_folder = Path(output_folder).resolve()
    output_root = output_folder.parent
    inventory = inventory if inventory is not None else scan_inventory(output_root)
    for path in (output_root/'Data'/'Entry Deletions').glob('*.json'):
        deletion = json.loads(path.read_text(encoding='utf-8'))
        if deletion.get('state') == 'complete':
            for value in deletion.get('cods', []):
                highest = max(highest, certificate_sequence(value, settings))
    for entry in inventory['entries']:
        for value in entry['cods']:
            highest = max(highest, certificate_sequence(value, settings))
    custom_register = output_root / "Data" / "Custom Certificates.json"
    if custom_register.is_file():
        custom_data = json.loads(custom_register.read_text(encoding="utf-8"))
        for record in custom_data.get("records", []):
            highest = max(highest, certificate_sequence(record.get("details", {}).get("cod_no", ""), settings))
    legacy_certificate_folder = output_root / "Output" / "Certificates"
    # Only durable certificate/history locations, never sample templates,
    # preview files or development/test directories elsewhere in the project.
    durable_folders = (output_folder, legacy_certificate_folder, output_root / 'Data' / 'Archive')
    for durable_folder in durable_folders:
        for pdf in durable_folder.rglob('*.pdf') if durable_folder.is_dir() else ():
            if pdf.name.startswith('.'):
                continue
            if pdf.is_relative_to(output_root / 'Data' / 'Archive' / 'Uncommitted Certificates'):
                # A rollback quarantine is diagnostic evidence, not an issued
                # certificate or a reason to consume its never-committed COD.
                continue
            stem = pdf.stem
            if durable_folder == output_root / 'Data' / 'Archive':
                # Recovery archives prepend a random ID, not a COD.
                stem = re.sub(r'^[a-f0-9]{32}_', '', stem, flags=re.I)
            # New COD-led PDFs have COD, company, and city fields. Older
            # generated PDFs may still have a final _1/_2 plant marker.
            new_client_name = re.fullmatch(r"[^_]+_.+_.+", stem, re.I)
            if new_client_name:
                cod_prefix = stem.split("_", 1)[0]
                highest = max(highest, certificate_sequence(cod_prefix, settings))
                continue
            # Do not treat older client-name PDFs as a number source.
            legacy_client_name = re.fullmatch(r"cod_.+_[12]", stem, re.I)
            legacy_cod_name = (
                bool(re.fullmatch(r"(?:cod_|certificate_)\d+", stem, re.I))
                or ("_" not in stem and stem.casefold().startswith(("cod", "certificate")))
            )
            in_flat_certificate_folder = pdf.parent.resolve() in {
                output_folder,
                legacy_certificate_folder.resolve(),
            }
            if not (legacy_cod_name or (in_flat_certificate_folder and not legacy_client_name)):
                continue
            match = re.search(r"(\d+)$", stem)
            if match:
                highest = max(highest, int(match.group(1)))
    # Read only legacy COD values for sequence continuity; no mail state is used.
    if legacy_history and Path(legacy_history).is_file():
        try:
            connection = sqlite3.connect(f"file:{Path(legacy_history).resolve().as_posix()}?mode=ro", uri=True)
            try:
                tables = {row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                for table in tables.intersection({"batches", "archived_issuances"}):
                    columns = {row[1] for row in connection.execute(f"PRAGMA table_info({table})")}
                    if "certificate_no" not in columns:
                        continue
                    for (value,) in connection.execute(f"SELECT certificate_no FROM {table}"):
                        highest = max(highest, certificate_sequence(value, settings))
            finally:
                connection.close()
        except sqlite3.Error:
            pass
    return highest


def starting_cod_sequence(value):
    """Accept a positive sequence number, including zero-padded input."""
    raw = str(value).strip()
    if not re.fullmatch(r'[0-9]{1,12}', raw) or int(raw) < 1:
        raise ValueError('Enter a starting COD number from 1 to 999999999999.')
    return int(raw)


def numbering_plan(batches, settings, output_folder, *, inventory=None, starting_cod=None):
    """Use the same numbering rules for batches, single previews and saves."""
    folder = Path(output_folder)
    highest = _highest_sequence(batches, dict(settings, certificate_number_start=1), folder,
                                folder.parent / 'Data' / 'certificate-history.sqlite', inventory=inventory)
    start = starting_cod_sequence(starting_cod) if starting_cod is not None else None
    configured = settings.get('next_cod')
    if start is not None and highest:
        raise ValueError('Existing CODs were found. Refresh to continue numbering automatically.')
    if start is not None and configured is not None:
        raise ValueError('The next COD is set in Settings. Refresh to use that sequence.')
    minimum = max(int(settings.get('certificate_number_start', 1)), highest + 1)
    sequence = start if start is not None else max(configured or 1, minimum)
    return {'needs_start': not highest and configured is None, 'next_cod': certificate_number(sequence, settings),
            'minimum_sequence': minimum,
            'prefix': settings.get('certificate_number_prefix', ''),
            'width': int(settings.get('certificate_number_width', 4)), 'sequence': sequence}


def _apply_overrides(batch, overrides):
    """Build the certificate view without mutating the parsed source batch."""
    if not overrides:
        return dict(batch)
    view = dict(batch)
    view["issues"] = list(batch.get("issues", []))
    for key in ("supplier", "location", "form6", "reference"):
        if key in overrides:
            view[key] = str(overrides[key] or "").strip()
    for key in ("certificate_type", "certificate_fields"):
        if key in overrides:
            view[key] = overrides[key]
    if "inward_date" in overrides:
        view["inward_date"] = parse_date(overrides["inward_date"])
    if "certificate_weight" in overrides:
        raw = overrides["certificate_weight"]
        parsed = raw if isinstance(raw, (int, float)) else number(raw)
        view["certificate_weight"] = parsed
        view["certificate_weight_invalid"] = parsed is None or parsed <= 0
    # An override is an explicit correction for the corresponding warning.
    if view.get("supplier"):
        view["issues"] = [issue for issue in view["issues"] if issue != "missing supplier name"]
    if view.get("inward_date"):
        view["issues"] = [issue for issue in view["issues"] if issue != "missing inward date"]
    if "reference" in overrides:
        view["issues"] = [issue for issue in view["issues"] if "reference" not in issue.casefold()]
    if "certificate_weight" in overrides:
        view["issues"] = [issue for issue in view["issues"] if "certificate/net weight" not in issue]
    certificate_weight = view.get("certificate_weight")
    if certificate_weight is not None and certificate_weight > 0:
        view["display_weight"] = certificate_weight
        view["weight_source"] = "certificate"
    else:
        view["display_weight"] = view.get("total_weight") or 0
        view["weight_source"] = "inward"
        if "certificate_weight" in overrides:
            view["issues"].append("certificate/net weight override must be a positive number")
    return view


def find_certificate(output_folder, batch, *, inventory=None, receipt=LOOKUP_RECEIPT):
    """Match stable COD + plant, even if the company/city was renamed in Excel."""
    folder = Path(output_folder)
    cod = str(batch.get("cod_no") or "").strip()
    if not cod or not folder.is_dir():
        return None
    try:
        safe = _safe_cod_name(cod)
    except ValueError:
        return None
    plant = str(batch.get("source_id") or "")[-1:]
    if receipt is LOOKUP_RECEIPT:
        receipt = latest_issuance(folder.parent, 'workbook', batch.get('source_id'), batch.get('case_id'))
    if receipt and cod_identity(receipt.get('cod')) == cod_identity(cod):
        from certificate_inventory import safe_certificate_path
        relative = receipt['pdf_path'].removeprefix('Certificates/')
        current = safe_certificate_path(folder.parent, relative, expected_sha=receipt.get('pdf_sha256'))
        return current if current_pdf_is_healthy(current, receipt['pdf_sha256']) else None
    if batch.get('external_pdf'):
        from certificate_inventory import safe_certificate_path
        link = batch['external_pdf']
        try:
            current = safe_certificate_path(folder.parent, link['path'], expected_sha=link['sha256'])
            if current_pdf_is_healthy(current, link['sha256']):
                return current
        except (KeyError, TypeError, ValueError, OSError):
            pass
        return None
    from certificate_inventory import safe_certificate_path
    expected_name = certificate_filename(batch, cod=cod).casefold()
    valid_matches = []
    for path in folder.rglob("*"):
        if path.suffix.casefold() != '.pdf':
            continue
        prefix = path.stem.split("_", 1)[0]
        same_prefix = path.stem.casefold().startswith((safe + "_").casefold())
        if safe.isdecimal() and prefix.isdecimal():
            same_prefix = cod_identity(safe) == cod_identity(prefix)
        # Keep resolving certificates created before plant numbers were
        # removed from filenames. Identify candidates before parsing PDF data;
        # unrelated certificates cannot match and can be expensive to open.
        candidate = (same_prefix and (path.name.casefold() == expected_name
                                     or plant and path.stem.endswith(f"_{plant}")))
        candidate = candidate or path.stem.casefold() in {safe.casefold(), f"cod_{safe}".casefold(), f"certificate_{safe}".casefold()}
        candidate = candidate or (safe.isdecimal() and path.stem.isdecimal()
                                  and cod_identity(safe) == cod_identity(path.stem))
        if candidate and current_pdf_is_healthy(path):
            valid_matches.append(path)
    # A unique filename/COD match is safe to reuse. Duplicate PDFs with the
    # same COD need their issuance receipt to identify the owning inward.
    if valid_matches:
        return valid_matches[0] if len(valid_matches) == 1 else None
    if status_value(batch.get('issued')) not in {'generated_manual', 'completed'}:
        return None
    inventory = inventory if inventory is not None else scan_inventory(folder.parent)
    matches = [entry for entry in inventory['entries']
               if entry.get('present') and not entry.get('managed_owner') and not entry.get('warnings')
               and entry.get('external')
               and {cod_identity(value) for value in entry.get('current_cods', [])} == {cod_identity(cod)}]
    if len(matches) == 1:
        current = safe_certificate_path(folder.parent, matches[0]['path'].removeprefix('Certificates/'))
        if current_pdf_is_healthy(current):
            return current
    return None


def run(workbook_paths, root, *, as_of=None, progress=print, check_only=False,
        publish_latest=True, limit=None, selected=None, custom_cod=None,
        plants=None, allow_warnings=False, overrides=None, details=None, update_excel=True,
        override_reason='', expected_revision=None, report_sink=None, read_only_preflight=False,
        starting_cod=None, selected_keys=None,
        remove_inward=False, hdd_needed=None, hdd_removal_confirmed=False):
    """Check both workbooks or generate PDFs for rows explicitly marked Ready."""
    root = Path(root)
    if remove_inward and (not selected or details is None or update_excel):
        raise ValueError('Select one inward and review its removal before keeping the PDF only.')
    if read_only_preflight and not check_only:
        raise ValueError('Read-only preflight cannot create certificates.')
    settings = load_settings(root)
    as_of = as_of or date.today()
    manual_cod = str(custom_cod or "").strip()
    if selected_keys is not None:
        if selected or plants is not None or not isinstance(selected_keys, (list, tuple, set, frozenset)) or not 1 <= len(selected_keys) <= 500:
            raise ValueError('Choose between 1 and 500 certificates from the list.')
        if any(not isinstance(key, (list, tuple)) or len(key) != 2 or key[0] not in {'plant1','plant2'} or not str(key[1]).strip() for key in selected_keys):
            raise ValueError('Choose valid clients from the certificate list.')
        selected_keys = {(key[0], str(key[1]).strip()) for key in selected_keys}
    start_sequence = None
    if starting_cod is not None:
        if manual_cod:
            raise ValueError('Enter either a starting COD number or a custom COD, not both.')
        start_sequence = starting_cod_sequence(starting_cod)
    if plants is not None:
        if not isinstance(plants, (list, tuple)) or not plants:
            raise ValueError("Choose Plant 1, Plant 2, or All plants.")
        plants = tuple(dict.fromkeys(plants))
        if any(plant not in {"plant1", "plant2"} for plant in plants):
            raise ValueError("Choose Plant 1, Plant 2, or All plants.")
        if selected:
            raise ValueError("Choose one inward or a plant selection, not both.")
    if manual_cod and not selected:
        raise ValueError("A custom COD can only be used when generating one selected Ready row.")
    if details is not None and (not selected or str(details.get('case_id')) != str(selected[1])):
        raise ValueError("The selected Serial Number cannot be changed when editing its certificate.")
    mode = "Check plant files" if check_only else "Generate certificates"
    with run_lock(root):
        if read_only_preflight:
            from backup_service import check_restore_health
            check_restore_health(root)
            if any(entry.get('state') not in {'complete', 'rolled_back', 'deleted'} for entry in audit_records(root)):
                raise ValueError('A previous certificate save needs recovery. Refresh the app, then try again.')
        else:
            recover_issuances(root)
        if expected_revision is not None:
            from workspace_review import require_revision
            require_revision(root, expected_revision)
        report = Report(root, as_of, mode, publish_latest=publish_latest and not read_only_preflight)
        report.result_sink = report_sink
        try:
            paths = [Path(path).expanduser().resolve() for path in workbook_paths]
            if len(paths) != 2 or any(not str(path) for path in paths):
                raise ValueError("Select the Plant 1 and Plant 2 Excel files in Settings.")
            if paths[0] == paths[1] or (paths[0].is_file() and paths[1].is_file() and paths[0].samefile(paths[1])):
                raise ValueError("Plant 1 and Plant 2 must be two different Excel files.")

            backup_dir = root / "Data" / "Workbook Backups"
            output_folder = root / "Certificates"
            template = resource(root, "templates") / "Certificate of Recycling.png"
            with ExitStack() as stack:
                sessions = []
                for index, path in enumerate(paths, 1):
                    try:
                        if read_only_preflight:
                            session = stack.enter_context(WorkbookSnapshot(path))
                        elif check_only:
                            try:
                                session = stack.enter_context(WorkbookSession(path, backup_dir, lock_timeout=0))
                            except LockBusyError:
                                # Browsing saved Excel data must remain possible
                                # while Excel is open. Only real saves need an
                                # exclusive workbook handle.
                                session = stack.enter_context(WorkbookSnapshot(path))
                        else:
                            session = stack.enter_context(WorkbookSession(path, backup_dir))
                        sessions.append(session)
                    except Exception as exc:
                        sessions.append(None)
                        report.add({'source_id': f'plant{index}', 'source_path': str(path)}, 'needs_review',
                                   f'{path.name} is unavailable: {exc}', issue_code='source_unavailable', system_alert=True)
                if expected_revision is not None:
                    require_revision(root, expected_revision, {str(session.path.resolve()): session.data for session in sessions if session is not None})
                inventory = scan_inventory(root, persist=not read_only_preflight)
                report.certificate_inventory = inventory
                if check_only and not read_only_preflight and details is None:
                    from certificate_drafts import read_register
                    from hdd_workbook import stage_hdd_tracking
                    for record in read_register(root)['records']:
                        hdd = record.get('details', {})
                        link = hdd.get('destruction_source') or {}
                        issued_date = parse_date(hdd.get('cod_date'))
                        if (hdd.get('certificate_type') != 'destruction' or not hdd.get('cod_no') or not issued_date
                                or not link or link.get('record_id') or link.get('plant') not in {'plant1', 'plant2'}):
                            continue
                        session = sessions[int(link['plant'][-1]) - 1]
                        if session is None or isinstance(session, WorkbookSnapshot):
                            continue
                        try:
                            source = next((b for b in read_batches(BytesIO(session.data))
                                           if str(b['case_id']) == str(link.get('case_id'))), None)
                            if source and (source.get('hdd_cod_no'), parse_date(source.get('hdd_cod_date'))) != (hdd['cod_no'], issued_date):
                                session._persist(stage_hdd_tracking(session, link['case_id'], hdd['cod_no'], issued_date))
                        except (OSError, ValueError) as exc:
                            report.maintenance_notes.append({'plant': link['plant'], 'detail': f'HDD number could not be added to Excel: {exc}'})
                batches = _read_sources(paths, sessions, report)
                if read_only_preflight:
                    for batch in batches:
                        batch.pop('_workbook_read_only', None)
                        batch['_preflight_read_only'] = True
                history = {(entry.get('plant'), str(entry.get('key')).strip().casefold()): entry
                           for entry in audit_records(root)
                           if entry.get('source_kind') == 'workbook' and entry.get('state') == 'complete'}
                from certificate_drafts import read_register
                manual_records = read_register(root)['records']
                for batch in batches:
                    receipt = history.get((batch['source_id'], batch['case_id'].strip().casefold()))
                    conflict = workbook_receipt_conflict(batch, receipt)
                    if not conflict:
                        ownership_view = dict(batch, cod_no=batch.get('cod_no') or (receipt or {}).get('cod'))
                        conflict = workbook_cod_ownership_issue(ownership_view, batches, manual_records)
                    if conflict:
                        batch['record_identity_conflict'] = conflict
                        batch['blocking_issues'].append(conflict)
                if not report.scan_complete:
                    report.add({}, "needs_review", "Certificate creation is paused. Connect and check both plant files in Settings, then refresh. This prevents duplicate COD numbers.",
                               issue_code="source_unavailable", system_alert=True)
                    for batch in batches:
                        batch['_issuance_blocked'] = True
                        receipt = history.get((batch['source_id'], batch['case_id'].strip().casefold()))
                        existing = find_certificate(output_folder, batch, inventory=inventory, receipt=receipt)
                        if existing:
                            batch.update(pdf_filename=existing.name, pdf_relative_path=existing.relative_to(output_folder).as_posix())
                        report.add(batch, 'source_check_incomplete',
                                   'This company was loaded. Certificate creation is paused until both plant files can be checked.',
                                   warnings=source_value_issues(batch, as_of))
                    return report.finish()

                from identifier_rules import mark_conflicts
                mark_conflicts(batches, manual_records)

                original_form_batch = None
                staged_data = None
                removed_data = None
                numbering_batches = list(batches)
                if details is not None:
                    session = sessions[int(selected[0][-1]) - 1]
                    matches = [batch for batch in batches if batch['source_id'] == selected[0]
                               and str(batch['case_id']).strip().casefold() == str(selected[1]).strip().casefold()]
                    if len(matches) > 1:
                        rows = ', '.join(str(batch['source_row']) for batch in matches)
                        raise ValueError(f'Serial Number {selected[1]} is repeated in Excel rows {rows}. '
                                         'Open the workbook and give each record a different Serial Number before creating a certificate.')
                    original_form_batch = next(iter(matches), None)
                    if remove_inward:
                        if not original_form_batch or original_form_batch.get('certificate_type', 'standard') == details.get('certificate_type'):
                            raise ValueError('Excel removal is only available when changing this inward’s certificate type.')
                        if (original_form_batch.get('hdd_cod_no')
                                or status_value(original_form_batch.get('issued')) in {'generated', 'generated_manual', 'completed'}
                                or latest_issuance(root, 'workbook', selected[0], selected[1])):
                            raise ValueError('A certificate has already been generated for this inward. It cannot be deleted.')
                        from inward_removal import stage_removal
                        removed_data, _, _ = stage_removal(session.data, original_form_batch['case_id'],
                            original_form_batch['source_sheet'], original_form_batch['source_row'], certificate_conversion=True)
                        from inward_serials import resequence_removed_inwards
                        removed_data, _ = resequence_removed_inwards(removed_data, session.data, root=root, plant=selected[0])
                    if original_form_batch and original_form_batch.get('record_identity_conflict'):
                        raise ValueError(original_form_batch['record_identity_conflict'])
                    if original_form_batch is None and not update_excel:
                        raise ValueError("Certificate-only changes require an existing inward. Choose Also update Excel or create a manual certificate.")
                    staged_data, staged_batch = stage_certificate_details(session, details, validate_cells=update_excel)
                    staged_batch.update(source_id=selected[0], source_path=str(session.path))
                    if isinstance(session, WorkbookSnapshot):
                        staged_batch['_preflight_read_only' if read_only_preflight else '_workbook_read_only'] = True
                    batches = [batch for batch in batches if (batch['source_id'], str(batch['case_id'])) != selected]
                    batches.append(staged_batch)
                    numbering_batches.append(staged_batch)

                mark_conflicts(batches, manual_records)
                renumbered = False
                for index, batch in enumerate(batches):
                    if plants is not None and batch["source_id"] not in plants:
                        continue
                    if (details is None and not read_only_preflight and not batch.get('_workbook_read_only')
                            and (index == 0 or batch["source_id"] != batches[index - 1]["source_id"])):
                        plant = batch['source_id']
                        plant_batches = [row for row in batches if row['source_id'] == plant]
                        before_numbers = [row['case_id'] for row in plant_batches]
                        session = sessions[int(plant[-1]) - 1]
                        can_renumber = check_only and selected is None and selected_keys is None
                        session.ensure_alert_formatting(batch['source_sheet'], plant_batches, as_of,
                                                        root=root if can_renumber else None, plant=plant)
                        renumbered |= before_numbers != [row['case_id'] for row in plant_batches]
                        order_warning = session.order_warning
                        if order_warning:
                            report.maintenance_notes.append({'plant': plant, 'detail': order_warning})

                if renumbered:
                    history = {(entry.get('plant'), str(entry.get('key')).strip().casefold()): entry
                               for entry in audit_records(root)
                               if entry.get('source_kind') == 'workbook' and entry.get('state') == 'complete'}
                    inventory = report.certificate_inventory = scan_inventory(root, persist=True)

                plan = numbering_plan(numbering_batches, settings, output_folder,
                                      inventory=inventory, starting_cod=start_sequence)
                next_sequence = plan['sequence']
                report.numbering = {key: value for key, value in plan.items() if key != 'sequence'}
                ready_count = 0
                for batch in batches:
                    if plants is not None and batch["source_id"] not in plants:
                        continue
                    if selected_keys is not None and (batch['source_id'], str(batch['case_id'])) not in selected_keys:
                        continue
                    if selected and (batch["source_id"], str(batch["case_id"])) != (selected[0], str(selected[1])):
                        continue
                    if limit is not None and ready_count >= limit and not check_only:
                        break
                    state = status_value(batch.get("issued"))
                    if batch.get('duplicate_serial_number') or batch.get('record_identity_conflict'):
                        batch['_issuance_blocked'] = True
                        report.add(batch, 'needs_review', ' '.join(batch['blocking_issues']),
                                   warnings=batch['blocking_issues'], system_alert=True)
                        continue
                    if (batch.get('_workbook_read_only') or read_only_preflight) and state in {'blank', 'ready', 'alert'}:
                        # Effective status only: the saved Excel cell remains
                        # untouched until a write can safely acquire its lock.
                        state = 'ready' if inward_details_complete(batch, as_of) else 'alert'
                    batch["_effective_certificate_status"] = state
                    age = (as_of - batch["inward_date"]).days if batch.get("inward_date") else None
                    due = certificate_reminder_due(batch, as_of, days=25)
                    overdue = bool(due and age is not None and age > 30)

                    if state == "na":
                        report.add(batch, "not_applicable", "N/A — no certificate is required.")
                        continue
                    if state == "alert":
                        issues = source_value_issues(batch, as_of)
                        detail = ("This row has issues to fix before it can be created: " + "; ".join(issues)
                                  if issues else
                                  "This row is marked Alert. Review the inward, then choose Ready or N/A in Excel.")
                        report.add(batch, "needs_review", detail, warnings=issues, system_alert=True)
                        continue
                    if state == "unknown":
                        report.add(batch, "needs_review", f"Unknown Certificate Issued value {batch.get('issued')!r}; review it, then choose Ready, Alert or N/A, or clear the cell.")
                        continue
                    receipt = history.get((batch['source_id'], batch['case_id'].strip().casefold()))
                    if state == 'ready':
                        previously_issued = batch.get('pdf_recovery_pending') or (
                            receipt and cod_identity(receipt.get('cod')) == cod_identity(batch.get('cod_no')))
                        if previously_issued and not find_certificate(output_folder, batch, inventory=inventory, receipt=receipt):
                            batch['_missing_pdf'] = True
                    if state in {"generated", "generated_manual", "completed"}:
                        existing_pdf = find_certificate(output_folder, batch, inventory=inventory, receipt=receipt)
                        if existing_pdf:
                            if state == "completed":
                                report.add(batch, "completed", f"COD {batch.get('cod_no')} is marked Completed.")
                            elif state == "generated_manual":
                                report.add(batch, "generated_manual", f"COD {batch.get('cod_no')} has a manually generated PDF. Send it using your normal process, then mark Completed.")
                            else:
                                report.add(batch, "generated", f"COD {batch.get('cod_no')} has a generated PDF. Send it using your normal process, then mark Completed.")
                            continue
                        if state == 'generated_manual' or batch.get('external_pdf'):
                            detail = ('The external PDF could not be matched. Open Saved certificates and link the correct PDF.'
                                      if state == 'generated_manual' else
                                      'The externally created PDF linked to this completed inward is missing or has changed.')
                            report.add(batch, 'needs_review', detail, system_alert=True)
                            continue
                        try:
                            if not (batch.get('_workbook_read_only') or read_only_preflight):
                                sessions[int(batch["source_id"][-1]) - 1].set_issued_status(
                                    batch["case_id"], batch["source_sheet"], batch["source_row"],
                                    batch["cod_no"], batch["issued"], "Ready",
                                )
                            batch["issued"] = "Ready"
                            batch["_effective_certificate_status"] = state = "ready"
                            batch["_missing_pdf"] = True
                        except Exception as exc:
                            report.add(batch, "needs_review", f"The PDF is missing, but Excel could not be changed to Ready: {exc}", system_alert=True)
                            continue
                    if state == "blank":
                        detail = "Fill the required inward details in Excel. The empty status will change to Ready when they are complete."
                        report.add(batch, "needs_review", detail, warnings=source_value_issues(batch, as_of))
                        continue
                    if state != "ready":
                        report.add(batch, "needs_review", "Choose Ready or N/A, or clear the Certificate Issued cell.")
                        continue

                    if batch.get('certificate_type', 'standard') == 'standard' and not batch.get('form6'):
                        report.add(batch, "needs_review",
                                   "Enter the Form 6 number when creating this certificate.",
                                   warnings=["Form 6 number is required to create the certificate."])
                        continue

                    if batch.get('blocking_issues'):
                        batch['_issuance_blocked'] = True
                        report.add(batch, 'needs_review', ' '.join(batch['blocking_issues']),
                                   warnings=batch['blocking_issues'], system_alert=True)
                        continue

                    unknown_pdfs = unresolved_cod_entries(inventory)
                    if unknown_pdfs:
                        batch['_issuance_blocked'] = True
                        recovery = [cod_recovery_message(entry['path']) for entry in unknown_pdfs[:3]]
                        if len(unknown_pdfs) > 3:
                            recovery.append(f"There are {len(unknown_pdfs) - 3} more PDFs to check.")
                        report.add(batch, 'needs_review', ' '.join(recovery), warnings=recovery, system_alert=True)
                        continue

                    source_batch = batch
                    remembered_choice = False
                    if details is None:
                        previous = (receipt or {}).get('after', {})
                        if isinstance(previous, dict) and previous.get('edit_scope') == 'certificate_only':
                            # A type chosen on the first preview is retained locally, not in Excel.
                            # Bulk recreation must use that template and its manually entered identifiers.
                            batch = _apply_overrides(batch, {key: previous[key] for key in
                                ('certificate_type', 'certificate_fields') if key in previous})
                            remembered_choice = True
                    batch = _apply_overrides(batch, overrides if selected and
                                             (batch["source_id"], str(batch["case_id"])) == selected else None)
                    if not str(batch.get('reference') or '').strip():
                        report.add(batch, 'needs_review', 'Enter the reference number before creating the certificate.',
                                   warnings=['Reference number is required to create the certificate.'])
                        continue
                    ready_count += 1
                    issues = source_value_issues(batch, as_of)
                    try:
                        kind = template_id(batch.get("certificate_type"))
                        if kind == 'destruction':
                            raise ValueError('Data Destruction is an additional certificate. Open Review & edit and choose Data Destruction needed.')
                        batch["certificate_type"] = kind
                        batch["certificate_fields"] = validate_fields(kind, batch.get("certificate_fields"))
                    except ValueError as exc:
                        report.add(batch, "needs_review", str(exc), system_alert=True)
                        continue
                    if kind == "siemens":
                        if not batch["certificate_fields"].get("site_location", "").strip():
                            batch["certificate_fields"]["site_location"] = str(batch.get("location") or "").strip()
                        for key, label in (("model_name", "model name"), ("model_serial", "model serial number"),
                                           ("tube_serial", "X-ray tube serial number"), ("site_location", "site location")):
                            if not str(batch["certificate_fields"].get(key) or "").strip():
                                issues.append(f"Missing {label}; enter N/A if it does not apply")
                    owned_code = manual_cod or source_batch.get('cod_no')
                    allowed_paths = issuance_paths(inventory, owned_code, 'workbook', source_batch['source_id'], source_batch['case_id'])
                    owned = bool(allowed_paths)
                    if issues and not allow_warnings:
                        report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)),
                                   warnings=list(dict.fromkeys(issues)), system_alert=True)
                        continue

                    row_cod = manual_cod or str(batch.get("cod_no") or "").strip()
                    try:
                        cod = row_cod or certificate_number(next_sequence, settings)
                        cod_owners = []
                        for owner in batches:
                            if cod_identity(owner.get("cod_no")) == cod_identity(cod):
                                cod_owners.append(owner)
                        custom_register = root / "Data" / "Custom Certificates.json"
                        if custom_register.is_file():
                            for record in json.loads(custom_register.read_text(encoding="utf-8")).get("records", []):
                                if cod_identity(record.get("details", {}).get("cod_no")) == cod_identity(cod):
                                    cod_owners.append({"source_id": "custom", "case_id": record["id"]})
                        if any(owner["source_id"] != batch["source_id"] or str(owner["case_id"]) != str(batch["case_id"])
                               for owner in cod_owners):
                            message = f"COD {cod} is already assigned to another plant row. Give each certificate a unique COD."
                            if not allow_warnings:
                                raise ValueError(message)
                            issues.append(message + " A new automatic COD was used after confirmation.")
                            cod = certificate_number(next_sequence, settings)
                        conflicts = cod_conflicts(inventory, cod, allowed_paths, owned=owned and cod_identity(cod) == cod_identity(source_batch.get('cod_no')))
                        if conflicts:
                            names = ', '.join(Path(entry['path']).name for entry in conflicts[:3])
                            replacement = certificate_number(next_sequence, settings)
                            message = f"COD {cod} is already recorded in certificate PDF history ({names}). Continuing will use new automatic COD {replacement}; existing PDFs will not be overwritten."
                            if not allow_warnings:
                                report.add(batch, 'needs_review', message, warnings=[message], system_alert=True)
                                continue
                            issues.append(message)
                            cod = replacement
                        try:
                            safe_cod = _safe_cod_name(cod)
                        except ValueError as exc:
                            if not allow_warnings:
                                raise
                            issues.append(str(exc) + " A new automatic COD was used after confirmation.")
                            cod = certificate_number(next_sequence, settings)
                            safe_cod = _safe_cod_name(cod)
                        issued_date = parse_date(batch.get("cod_date")) or as_of
                        destination = certificate_destination(root, certificate_filename(batch, cod=cod))
                        batch['pdf_relative_path'] = destination.relative_to(output_folder).as_posix()
                        native_template = TEMPLATE_ROOT / f"{kind}.pdf"
                        batch["_data_root"] = str(root)
                        if native_template.is_file() and (kind != "standard" or (resource(root, "templates") / "certificate-types" / "standard.pdf").is_file()):
                            active_template = native_template
                            issues.extend(layout_warnings(batch, cod, issued_date))
                            if issues and not allow_warnings:
                                report.add(batch, "needs_review", "; ".join(dict.fromkeys(issues)), warnings=list(dict.fromkeys(issues)), system_alert=True)
                                continue
                        else:
                            active_template = template
                            certificate_layout(batch, cod, issued_date, allow_warnings=allow_warnings)
                        if check_only:
                            detail = f"Ready to generate COD {cod}."
                            if issues:
                                detail += " Warnings: " + "; ".join(dict.fromkeys(issues))
                            report.add(batch, "ready", detail, system_alert=bool(issues),
                                       warnings=list(dict.fromkeys(issues)))
                            if not row_cod:
                                next_sequence += 1
                            else:
                                next_sequence = max(next_sequence, certificate_sequence(cod, settings) + 1)
                            continue

                        temporary = root / "Data" / "Previews" / f".{safe_cod}.{uuid.uuid4().hex}.tmp.pdf"
                        transaction = None
                        workbook_committed = False
                        session = sessions[int(batch["source_id"][-1]) - 1]
                        try:
                            make_certificate(batch, cod, issued_date, active_template, temporary,
                                             allow_warnings=allow_warnings)
                            previous_cod = source_batch.get('cod_no')
                            old_paths = issuance_paths(inventory, previous_cod,
                                                       'workbook', batch['source_id'], batch['case_id'])
                            verify_publication(root, inventory, cod, allowed_paths=list(set(allowed_paths + old_paths)),
                                               owned=owned and cod_identity(cod) == cod_identity(source_batch.get('cod_no')),
                                               allow_warnings=allow_warnings)
                            local_record, companion, converted_from = None, None, None
                            if remove_inward:
                                from certificate_conversion import prepare_local_certificate
                                local_record, companion, converted_from = prepare_local_certificate(
                                    root, batch['source_id'], original_form_batch, batch, cod, issued_date, destination,
                                    hdd_needed=hdd_needed, hdd_removal_confirmed=hdd_removal_confirmed,
                                    warnings=list(dict.fromkeys(issues)))
                            transaction = IssuanceTransaction(
                                root, destination, temporary, session.path, session.data,
                                source_kind='manual' if remove_inward else 'workbook', plant=batch['source_id'],
                                key=local_record['id'] if local_record else batch['case_id'],
                                cod=cod, issued_date=issued_date, before=original_form_batch or source_batch,
                                after=local_record['details'] if local_record else dict(batch, cod_no=cod, cod_date=issued_date,
                                           edit_scope='certificate_only' if remembered_choice or not update_excel else 'update_excel'),
                                warnings=list(dict.fromkeys(issues)), override_reason=override_reason,
                                old_paths=[root / path for path in old_paths], template_path=active_template,
                                companion_source=companion, converted_from=converted_from)
                            transaction.publish()
                            session.before_persist = transaction.expect_source
                            if remove_inward:
                                transaction.expect_source(removed_data)
                                transaction.sync_companion(True)
                                backup = session._persist(removed_data)
                                batch['converted_record_id'] = local_record['id']
                            elif details is not None:
                                deferred_form6 = None
                                deferred_reference = None
                                if (not update_excel and original_form_batch
                                        and not original_form_batch.get('form6')
                                        and status_value(original_form_batch.get('issued'))
                                            not in {'generated', 'generated_manual', 'completed'}):
                                    deferred_form6 = batch.get('form6')
                                if (not update_excel and original_form_batch
                                        and not original_form_batch.get('reference')
                                        and status_value(original_form_batch.get('issued'))
                                            not in {'generated', 'generated_manual', 'completed'}):
                                    deferred_reference = batch.get('reference')
                                backup = session.commit(original_form_batch or source_batch, cod, issued_date,
                                                        staged_data=staged_data if update_excel else None,
                                                        deferred_form6=deferred_form6,
                                                        deferred_reference=deferred_reference,
                                                        allow_reissue=True, new_batch=original_form_batch is None)
                            else:
                                backup = session.commit(source_batch, cod, issued_date, target_status="Generated",
                                                        overrides=overrides if selected and
                                                        (source_batch["source_id"], str(source_batch["case_id"])) == selected else None)
                            workbook_committed = True
                            try:
                                transaction.mark_source_committed()
                                transaction.complete()
                            except (OSError, ValueError) as recovery_error:
                                report.pending_recovery = operation_error_message(recovery_error)
                                issues.append(f'Certificate and Excel saved; final history update needs recovery on refresh: {recovery_error}')
                        except Exception as exc:
                            if transaction and not workbook_committed:
                                try:
                                    transaction.rollback()
                                except (OSError, ValueError) as recovery_error:
                                    raise RuntimeError(
                                        f"{exc}. The Excel update failed and the previous PDF could not be restored; "
                                        f"saved recovery copies were preserved. Ask for help before retrying: {recovery_error}"
                                    ) from exc
                            raise
                        finally:
                            session.before_persist = None
                            try:
                                temporary.unlink(missing_ok=True)
                            except OSError:
                                pass

                        batch["cod_no"] = cod
                        batch["cod_date"] = issued_date
                        batch["issued"] = "Generated"
                        batch['pdf_filename'] = destination.name
                        batch["_effective_certificate_status"] = "generated"
                        report.created += 1
                        detail = f"Generated {destination.name}; Excel was updated to COD {cod}, {issued_date:%d.%m.%Y}, Generated. Send it manually, then mark Completed."
                        if remove_inward:
                            detail = f'Generated {destination.name}; the selected inward was removed from Excel. The PDF is kept in the local certificate register. Send it manually, then mark Completed.'
                        if issues:
                            detail += " Warnings accepted: " + "; ".join(dict.fromkeys(issues))
                        if backup:
                            detail += f" Workbook backup: {backup.name}."
                        report.add(batch, "generated_now", detail,
                                   system_alert=bool(issues), warnings=list(dict.fromkeys(issues)))
                        if report.pending_recovery:
                            # Do not commit another row to this source until the
                            # pending transaction's exact source hash is recovered.
                            break
                        # Reserve this successful issuance immediately for any
                        # following row in the same bulk operation.
                        # A one-certificate request has no following row that
                        # needs this refreshed reservation. The final scan
                        # below records it once after the operation.
                        if not (selected or (selected_keys is not None and len(selected_keys) == 1)):
                            try:
                                inventory = scan_inventory(root, persist=True)
                            except Exception as exc:
                                report.pending_recovery = ('The certificate list could not be updated after saving. '
                                                           f'Refresh before retrying the rest of the batch. {operation_error_message(exc)}')
                                break
                        if not row_cod:
                            next_sequence += 1
                        else:
                            next_sequence = max(next_sequence, certificate_sequence(cod, settings) + 1)
                        if progress:
                            progress(f"{batch['source_id']} SN {batch['case_id']}: saved {destination.name}")
                    except Exception as exc:
                        report.add(batch, "needs_review", operation_error_message(exc), issue_code="certificate_generation_failure", system_alert=True)
            # A check-only pass cannot change PDFs, so keep the inventory read
            # at the start of this locked operation instead of hashing every
            # saved certificate a second time before returning.
            if not check_only:
                report.certificate_inventory = scan_inventory(root, persist=not read_only_preflight)
            return report.finish()
        except Exception as exc:
            report.add({}, "needs_review", operation_error_message(exc), issue_code="run_failure", system_alert=True)
            return report.finish()
