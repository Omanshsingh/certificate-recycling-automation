"""Local, authenticated snapshots and recoverable restores.

The backend owns native folder selection and persistence of backup_destination.
Call with locked=True ONLY while already holding workflow_support.run_lock(root).
Never expose that option, a destination path, or an archive path to restore clients.
Archives are authenticated, not encrypted. The local trust store is not backed up
into itself; disaster recovery must separately preserve it and the archives.
"""
from __future__ import annotations

from contextlib import ExitStack, contextmanager, nullcontext
from datetime import datetime, timezone
import hashlib
import hmac
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import uuid
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

from workflow_support import load_settings, run_lock
from workbook_store import LockBusyError, exclusive_file

VERSION = 1
MAX_FILES = 100_000
MAX_FILE_BYTES = 2 * 1024**3
MAX_TOTAL_BYTES = 20 * 1024**3
MAX_MANIFEST_BYTES = 32 * 1024**2
ID_PATTERN = re.compile(r"backup-[0-9a-f]{32}\Z")
SKIP_DIRS = {"preview", "previews", "report", "reports", "cache", "__pycache__",
             "backup", "backups", "archive", "archives"}
TEMPLATE_SUFFIXES = {".pdf", ".pptx", ".png", ".jpg", ".jpeg", ".json", ".ttf", ".otf"}


class BackupError(ValueError):
    """Unsafe, incomplete, corrupt, or incompatible backup operation."""


class RestoreError(RuntimeError):
    """Restore failed; recovery_path contains originals and a durable journal."""

    def __init__(self, message, *, recovery_path, rolled_back):
        super().__init__(message)
        self.recovery_path = str(recovery_path)
        self.rolled_back = rolled_back


def _json_bytes(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"),
                      ensure_ascii=False, allow_nan=False).encode("utf-8")


def _safe_path(path):
    """Inspect lexical ancestors BEFORE resolving; reject junctions and hardlinks."""
    path = Path(os.path.abspath(Path(path).expanduser()))
    for part in [*reversed(path.parents), path]:
        try:
            info = part.lstat()
        except FileNotFoundError:
            continue
        if stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400:
            raise BackupError(f"Links and Windows reparse points are not allowed: {part}")
        if stat.S_ISREG(info.st_mode) and info.st_nlink > 1:
            raise BackupError(f"Hard-linked files are not allowed: {part}")
        if not (stat.S_ISDIR(info.st_mode) or stat.S_ISREG(info.st_mode)):
            raise BackupError(f"Not a regular file or directory: {part}")
    return path


def _root(root):
    root = _safe_path(root)
    if not root.is_dir():
        raise BackupError("The workspace must be an existing directory.")
    _safe_path(root / "Data" / "run.lock")
    _safe_path(root / "settings.local.json")
    _safe_path(root / "App" / "settings.example.json")
    _safe_path(root / "settings.example.json")
    return root


def _identity(path):
    # Case folding also rejects Windows aliases when tested on another platform.
    return str(path).casefold()


def _inside(path, parent):
    return _identity(path) == _identity(parent) or _identity(path).startswith(
        _identity(parent) + os.sep)


def _relative(value):
    if not isinstance(value, str) or not value or "\\" in value or ":" in value:
        raise BackupError("Unsafe relative archive path.")
    parts = value.split("/")
    for part in parts:
        if (part in {"", ".", ".."} or part.endswith((".", " "))
                or any(ord(c) < 32 or c in '<>"|?*' for c in part)
                or re.fullmatch(r"(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\..*)?",
                                part, re.I)):
            raise BackupError("Unsafe relative archive path.")
    if PurePosixPath(value).is_absolute():
        raise BackupError("Absolute archive paths are not allowed.")
    return parts


def _approved(relative):
    parts = _relative(relative)
    lower = [p.casefold() for p in parts]
    name = lower[-1]
    if name.startswith(("~$", ".")) or name.endswith((".lock", ".tmp", "-wal", "-shm", "-journal")):
        return False
    if relative in {"settings.local.json", "settings.example.json", "App/settings.example.json"}:
        return True
    if lower[0] in {"templates", "app"}:
        offset = 1 if lower[0] == "templates" else 2
        if lower[0] == "app" and (len(lower) < 3 or lower[1] != "templates"):
            return False
        return (not any(p in SKIP_DIRS or "backup" in p for p in lower[offset:-1])
                and Path(name).suffix in TEMPLATE_SUFFIXES)
    if lower[0] == "certificates" or lower[:2] == ["output", "certificates"]:
        folders = lower[1:-1] if lower[0] == "certificates" else lower[2:-1]
        return not any(p in SKIP_DIRS or "backup" in p for p in folders)
    if lower[0] != "data" or len(parts) < 2:
        return False
    folders = lower[1:-1]
    if any((p in SKIP_DIRS or "backup" in p) and p not in {"archive", "archives"}
           for p in folders):
        return False
    if any(p in {"archive", "archives"} for p in folders):
        if (lower[:3] == ["data", "archive", "restore extras"] and len(parts) > 4
                and re.fullmatch(r"[0-9a-f]{32}", lower[3])):
            return _active_restore_file("/".join(parts[4:]))
        return Path(name).suffix == ".pdf"
    return True


def _settings(root):
    settings = load_settings(root)
    local = root / "settings.local.json"
    if local.exists() and local.stat().st_size > MAX_MANIFEST_BYTES:
        raise BackupError("Local settings exceed the backup safety limit.")
    raw = json.loads(local.read_text(encoding="utf-8-sig")) if local.exists() else {}
    destination = raw.get("backup_destination", settings.get("backup_destination"))
    if destination is not None and not isinstance(destination, str):
        raise BackupError("backup_destination must be a folder path.")
    if isinstance(destination, str) and not destination.strip():
        destination = None  # Settings defaults may represent the local default as "".
    return settings, destination


def _workbooks(root, settings):
    paths = []
    for raw in settings["workbook_paths"]:
        if not raw.strip():
            paths.append(None)
            continue
        # Match the backend: relative configured paths are relative to its cwd.
        path = _safe_path(raw)
        if path.suffix.casefold() not in {".xlsx", ".xlsm"}:
            raise BackupError("Only configured XLSX and XLSM workbooks can be backed up.")
        if _inside(path, root / "Data" / "Backups"):
            raise BackupError("A configured workbook cannot be in the backup trust store.")
        paths.append(path)
    populated = [p for p in paths if p is not None]
    if len({_identity(p) for p in populated}) != len(populated):
        raise BackupError("Plant workbooks must have distinct paths.")
    return paths


def _store(root):
    return _safe_path(root / "Data" / "Backups")


def _atomic_bytes(path, content):
    path = _safe_path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = _safe_path(path.with_name("." + path.name + "." + uuid.uuid4().hex + ".tmp"))
    try:
        with temporary.open("xb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        _safe_path(path)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)  # Only our disposable staging file.


def _key(root, create=False):
    store = _store(root)
    path = _safe_path(store / "trust.key")
    if not path.exists():
        if not create:
            raise BackupError("The local backup trust key is missing; restore is refused.")
        if (store / "catalog.json").exists() or any(store.glob("backup-*.zip")):
            raise BackupError("Existing backups have no trust key; do not generate a replacement.")
        store.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(os.urandom(32))
            stream.flush()
            os.fsync(stream.fileno())
        if os.name != "nt":
            path.chmod(0o600)
    key = path.read_bytes()
    if len(key) != 32:
        raise BackupError("The local backup trust key is invalid.")
    return key


def _sign(value, key):
    return hmac.new(key, _json_bytes(value), hashlib.sha256).hexdigest()


def _catalog(root, key):
    path = _safe_path(_store(root) / "catalog.json")
    if not path.exists():
        return {"version": VERSION, "backups": []}
    try:
        if path.stat().st_size > MAX_MANIFEST_BYTES:
            raise BackupError("Backup catalog is too large.")
        envelope = json.loads(path.read_text(encoding="utf-8"))
        catalog = envelope["catalog"]
        if (not hmac.compare_digest(envelope["signature"], _sign(catalog, key))
                or catalog["version"] != VERSION or not isinstance(catalog["backups"], list)):
            raise BackupError("The local backup catalog failed authentication.")
        ids = [item["backup_id"] for item in catalog["backups"]]
        if len(ids) != len(set(ids)) or any(not ID_PATTERN.fullmatch(i) for i in ids):
            raise BackupError("The local backup catalog contains invalid IDs.")
        return catalog
    except (KeyError, TypeError, ValueError) as exc:
        raise BackupError(f"The local backup catalog is invalid: {exc}") from exc


def _save_catalog(root, catalog, key):
    _atomic_bytes(_store(root) / "catalog.json",
                  _json_bytes({"catalog": catalog, "signature": _sign(catalog, key)}))


def _destination(root, requested, persisted, workbooks):
    selected = requested if requested is not None else persisted
    if selected is None:
        path = _store(root)
    else:
        if not isinstance(selected, (str, os.PathLike)) or not str(selected).strip():
            raise BackupError("Choose a valid backup folder.")
        path = _safe_path(selected)
    # Do not permit recursion, restore-target overlap, or writing into program files.
    included = [root / p for p in ("Certificates", "Output", "templates", "App", "Data")]
    if path != _store(root) and any(_inside(path, p) for p in included):
        raise BackupError("The backup destination cannot be inside an included data or program folder.")
    if path == root or any(p is not None and _inside(p, path) for p in workbooks):
        raise BackupError("Choose a separate folder that does not contain a configured workbook.")
    if path.exists() and not path.is_dir():
        raise BackupError("The backup destination must be a folder.")
    return path


def _sources(root, workbooks, allow_missing=False):
    result, seen = [], set()
    for index, path in enumerate(workbooks):
        if path is None:
            continue
        if not path.is_file():
            if allow_missing:
                continue
            raise BackupError(f"Configured workbook is missing: {path}")
        result.append((path, {"kind": "workbook", "workbook_index": index}))
        seen.add(_identity(path))
    for name in ("settings.local.json", "settings.example.json", "App/settings.example.json"):
        path = _safe_path(root / name)
        if path.is_file() and _identity(path) not in seen:
            result.append((path, {"kind": "workspace", "relative_path": name}))
            seen.add(_identity(path))
    for name in ("Certificates", "Data", "Output/Certificates", "templates", "App/templates"):
        base = _safe_path(root / name)
        if not base.exists():
            continue
        for folder, directories, files in os.walk(base, followlinks=False):
            # Prune excluded folders without following a link into any of them.
            retained = []
            for directory in sorted(directories):
                low = directory.casefold()
                if low in SKIP_DIRS and low not in {"archive", "archives"} or "backup" in low:
                    continue
                candidate = _safe_path(Path(folder) / directory)
                retained.append(candidate.name)
            directories[:] = retained
            for filename in sorted(files):
                path = Path(folder) / filename
                relative = path.relative_to(root).as_posix()
                if not _approved(relative):
                    continue
                path = _safe_path(path)
                if _identity(path) in seen:
                    continue
                if not path.is_file():
                    raise BackupError(f"A backup source disappeared: {path}")
                seen.add(_identity(path))
                result.append((path, {"kind": "workspace", "relative_path": relative}))
    if len(result) > MAX_FILES:
        raise BackupError("Too many files for a local backup.")
    for path, _ in result:
        _no_sqlite_sidecars(path)
    return result


def _no_sqlite_sidecars(path):
    if path.suffix.casefold() in {".sqlite", ".sqlite3", ".db"}:
        if any(Path(str(path) + suffix).exists() for suffix in ("-wal", "-shm", "-journal")):
            raise BackupError(f"Close database users and checkpoint SQLite before backup/restore: {path}")


def _pending(root):
    directory = _safe_path(_store(root) / "Restore Transactions")
    if not directory.exists():
        return
    for transaction in directory.iterdir():
        _safe_path(transaction)
        journal = _safe_path(transaction / "journal.json")
        if journal.exists():
            try:
                state = json.loads(journal.read_text(encoding="utf-8"))["state"]
            except (ValueError, KeyError, TypeError) as exc:
                raise BackupError(f"Unusable restore journal requires recovery: {journal}") from exc
            if state not in {"committed", "rolled_back", "prepared"}:
                raise BackupError(f"An interrupted restore requires supervised recovery: {journal}")


def check_restore_health(root):
    """Read-only startup/write guard; raise BackupError for interrupted restore.

    Call inside the same run_lock as every issuance or other workspace mutation,
    and before issuance recovery. Never automatically replay unknown journal paths.
    Prepared journals have not begun replacement; committed/rolled_back are safe.
    """
    root = _root(root)
    _pending(root)
    return {"healthy": True}


def _hash(path):
    digest = hashlib.sha256()
    with _safe_path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@contextmanager
def _read_source(path):
    """Lock capture without asking for write access to templates/read-only media."""
    path = _safe_path(path)
    if os.name == "nt":
        import ctypes
        import msvcrt
        from ctypes import wintypes
        api = ctypes.WinDLL("kernel32", use_last_error=True)
        api.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                   ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        api.CreateFileW.restype = wintypes.HANDLE
        handle = api.CreateFileW(str(path), 0x80000000, 4, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            code = ctypes.get_last_error()
            if code in (32, 33):
                raise LockBusyError(f"Close Excel or other file users before backup/restore: {path}")
            raise PermissionError(f"Cannot read backup file: {path} (Windows error {code})")
        stream = os.fdopen(msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY), "rb")
    else:
        import fcntl
        stream = path.open("rb")
        try:
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            stream.close()
            raise LockBusyError(f"Close other file users before backup/restore: {path}") from exc
    try:
        yield stream
    finally:
        stream.close()


def _create(root, destination, *, allow_missing=False, held=None, purpose="manual", configuration=None):
    _pending(root)
    settings, persisted = configuration if configuration is not None else _settings(root)
    workbooks = _workbooks(root, settings)
    destination = _destination(root, destination, persisted, workbooks)
    sources = _sources(root, workbooks, allow_missing)
    key = _key(root, create=True)
    catalog = _catalog(root, key)
    destination.mkdir(parents=True, exist_ok=True)
    backup_id = "backup-" + uuid.uuid4().hex
    path = _safe_path(destination / (backup_id + ".zip"))
    temporary = path.with_suffix(".partial")
    manifest = {"version": VERSION, "backup_id": backup_id,
                "created_at": datetime.now(timezone.utc).isoformat(),
                "workspace_root": str(root), "purpose": purpose,
                "original_workbook_paths": [str(p) if p is not None else None for p in workbooks],
                "files": []}
    total = 0
    try:
        with ExitStack() as locks, ZipFile(temporary, "x", compression=ZIP_DEFLATED) as archive:
            streams = dict(held or {})
            for source, _ in sources:
                _safe_path(source)
                if source not in streams:
                    streams[source] = locks.enter_context(_read_source(source))
            defaults = root / ("settings.example.json" if (root / "settings.example.json").exists()
                               else "App/settings.example.json")
            merged = {}
            for config_path in (defaults, root / "settings.local.json"):
                if config_path in streams:
                    streams[config_path].seek(0)
                    content = streams[config_path].read(MAX_MANIFEST_BYTES + 1)
                    if len(content) > MAX_MANIFEST_BYTES:
                        raise BackupError("Settings exceed the backup safety limit.")
                    value = json.loads(content.decode("utf-8-sig"))
                    if not isinstance(value, dict):
                        raise BackupError("Settings must be an object.")
                    merged.update(value)
            raw_paths = merged.get("workbook_paths", ["", ""])
            if not isinstance(raw_paths, list) or len(raw_paths) > 2 or any(not isinstance(p, str) for p in raw_paths):
                raise BackupError("Settings changed during backup; retry after reviewing them.")
            actual = _workbooks(root, {"workbook_paths": (raw_paths + ["", ""])[:2]})
            if actual != workbooks:
                raise BackupError("Configured workbook mapping changed during backup.")
            for index, (source, target) in enumerate(sources):
                _no_sqlite_sidecars(source)
                stream = streams[source]
                stream.seek(0)
                before = os.fstat(stream.fileno())
                if before.st_size > MAX_FILE_BYTES:
                    raise BackupError(f"File exceeds backup size limit: {source}")
                member = f"payload/{index:06d}"
                digest, size = hashlib.sha256(), 0
                with archive.open(member, "w", force_zip64=True) as output:
                    for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                        size += len(chunk)
                        total += len(chunk)
                        if size > MAX_FILE_BYTES or total > MAX_TOTAL_BYTES:
                            raise BackupError("Backup exceeds the configured safety size limits.")
                        digest.update(chunk)
                        output.write(chunk)
                after = os.fstat(stream.fileno())
                current = _safe_path(source).stat()
                if (before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns
                        or before.st_ino != current.st_ino or before.st_dev != current.st_dev):
                    raise BackupError(f"A source changed during backup: {source}")
                manifest["files"].append({**target, "member": member, "size": size,
                                          "sha256": digest.hexdigest()})
            encoded_manifest = _json_bytes(manifest)
            if len(encoded_manifest) > MAX_MANIFEST_BYTES:
                raise BackupError("Manifest exceeds the backup safety limit.")
            archive.writestr("manifest.json", encoded_manifest)
            archive.writestr("manifest.hmac", _sign(manifest, key))
        with temporary.open("r+b") as stream:
            os.fsync(stream.fileno())
        _safe_path(path)
        if path.exists():
            raise BackupError("The generated backup ID already exists; no archive was replaced.")
        os.replace(temporary, path)
        record = {"backup_id": backup_id, "directory": str(destination), "sha256": _hash(path)}
        catalog["backups"].append(record)
        _save_catalog(root, catalog, key)
        return _summary(manifest, path)
    finally:
        temporary.unlink(missing_ok=True)


def _summary(manifest, path):
    return {"backup_id": manifest["backup_id"], "path": str(path),
            "created_at": manifest["created_at"], "purpose": manifest["purpose"],
            "file_count": len(manifest["files"]),
            "total_bytes": sum(f["size"] for f in manifest["files"]),
            "original_workbook_paths": manifest["original_workbook_paths"],
            "trust_files": [str(Path(manifest["workspace_root"]) / "Data" / "Backups" / name)
                            for name in ("trust.key", "catalog.json")],
            "warnings": ["A backup on the same disk does not protect against disk failure.",
                         "Archives contain private data and are not encrypted."]}


def create_backup(root, destination=None, *, locked=False):
    """Create a ZIP and return backup_id/path/created_at/file_count/total_bytes.

    destination overrides persisted backup_destination for this call; it does not
    change settings. Missing configured workbooks fail the entire backup.
    """
    root = _root(root)
    with nullcontext() if locked else run_lock(root):
        return _create(root, destination)


def _inspect(root, record, key, *, stage=None):
    try:
        return _inspect_archive(root, record, key, stage=stage)
    except BackupError:
        raise
    except (BadZipFile, KeyError, TypeError, ValueError, RuntimeError) as exc:
        raise BackupError(f"Invalid backup archive: {exc}") from exc


def _inspect_archive(root, record, key, *, stage=None):
    backup_id = record["backup_id"]
    if not isinstance(backup_id, str) or not ID_PATTERN.fullmatch(backup_id):
        raise BackupError("Invalid backup ID.")
    directory = _safe_path(record["directory"])
    path = _safe_path(directory / (backup_id + ".zip"))
    if not path.is_file():
        raise BackupError("The registered backup file is missing.")
    with _read_source(path) as source:
        digest = hashlib.sha256()
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
        if not hmac.compare_digest(digest.hexdigest(), record["sha256"]):
            raise BackupError("Backup archive SHA256 failed verification.")
        source.seek(0)
        with ZipFile(source) as archive:
            infos = archive.infolist()
            names = [i.filename for i in infos]
            if len(names) > MAX_FILES + 2 or len(set(n.casefold() for n in names)) != len(names):
                raise BackupError("Duplicate archive members or too many files.")
            for info in infos:
                _relative(info.filename)
                mode = info.external_attr >> 16
                if (info.is_dir() or stat.S_ISLNK(mode)
                        or stat.S_IFMT(mode) not in {0, stat.S_IFREG}
                        or info.flag_bits & 1 or info.file_size > MAX_FILE_BYTES):
                    raise BackupError("Unsupported or unsafe archive member.")
            if (archive.getinfo("manifest.json").file_size > MAX_MANIFEST_BYTES
                    or archive.getinfo("manifest.hmac").file_size != 64):
                raise BackupError("Invalid manifest size.")
            manifest = json.loads(archive.read("manifest.json"))
            signature = archive.read("manifest.hmac").decode("ascii")
            if not hmac.compare_digest(signature, _sign(manifest, key)):
                raise BackupError("Backup manifest authentication failed.")
            if (manifest["version"] != VERSION or manifest["backup_id"] != backup_id
                    or _identity(manifest["workspace_root"]) != _identity(root)):
                raise BackupError("Backup belongs to a different workspace or format.")
            entries = manifest["files"]
            if not isinstance(entries, list) or len(entries) > MAX_FILES:
                raise BackupError("Invalid file manifest.")
            expected = {"manifest.json", "manifest.hmac"}
            targets, total = set(), 0
            for index, entry in enumerate(entries):
                if entry["member"] != f"payload/{index:06d}":
                    raise BackupError("Invalid payload mapping.")
                expected.add(entry["member"])
                if entry["kind"] == "workspace":
                    if not _approved(entry["relative_path"]):
                        raise BackupError("Unapproved workspace restore path.")
                    target = "root/" + entry["relative_path"].casefold()
                elif entry["kind"] == "workbook" and type(entry["workbook_index"]) is int and entry["workbook_index"] in (0, 1):
                    target = f"workbook/{entry['workbook_index']}"
                else:
                    raise BackupError("Invalid restore target type.")
                if target in targets:
                    raise BackupError("Duplicate restore targets.")
                targets.add(target)
                info = archive.getinfo(entry["member"])
                if (type(entry["size"]) is not int or entry["size"] < 0
                        or entry["size"] != info.file_size
                        or not re.fullmatch(r"[0-9a-f]{64}", entry["sha256"])):
                    raise BackupError("Invalid payload size or checksum.")
                total += info.file_size
                if total > MAX_TOTAL_BYTES:
                    raise BackupError("Expanded backup exceeds the safety limit.")
                # Read and hash every payload, even when merely listing backups.
                with archive.open(info) as payload, ExitStack() as outputs:
                    output = outputs.enter_context((stage / str(index)).open("xb")) if stage else None
                    digest, size = hashlib.sha256(), 0
                    for chunk in iter(lambda: payload.read(1024 * 1024), b""):
                        size += len(chunk)
                        if size > entry["size"]:
                            raise BackupError("Expanded payload exceeds declared size.")
                        digest.update(chunk)
                        if output:
                            output.write(chunk)
                    if output:
                        output.flush()
                        os.fsync(output.fileno())
                if size != entry["size"] or digest.hexdigest() != entry["sha256"]:
                    raise BackupError("Backup payload SHA256 failed verification.")
            if expected != set(names):
                raise BackupError("Unexpected or missing archive members.")
    return manifest, path


def list_backups(root):
    """Return newest-first verified registered backups; invalid entries have valid=False.

    No filesystem glob/import or caller-supplied archive path is trusted. Listing
    performs no writes and deliberately reports damaged/missing registered ZIPs.
    """
    root = _root(root)
    if not (_store(root) / "catalog.json").exists():
        return []
    key = _key(root)
    records = _catalog(root, key)["backups"]
    result = []
    for record in reversed(records):
        try:
            manifest, path = _inspect(root, record, key)
            result.append({**_summary(manifest, path), "valid": True})
        except (OSError, ValueError, KeyError, TypeError, RuntimeError, BadZipFile) as exc:
            result.append({"backup_id": record["backup_id"], "valid": False, "error": str(exc)})
    return result


def _targets(root, manifest):
    settings, _ = _settings(root)
    workbooks = _workbooks(root, settings)
    original = manifest["original_workbook_paths"]
    current = [str(p) if p is not None else None for p in workbooks]
    if (not isinstance(original, list) or len(original) != len(current)
            or any((a is None) != (b is None) or
                   (a is not None and (not isinstance(a, str) or _identity(a) != _identity(b)))
                   for a, b in zip(original, current))):
        raise BackupError("Original workbook mapping is missing or differs from the configured workbooks.")
    present = {e["workbook_index"] for e in manifest["files"] if e["kind"] == "workbook"}
    if manifest["purpose"] != "pre-restore" and present != {i for i, p in enumerate(workbooks) if p is not None}:
        raise BackupError("The backup is missing an original workbook payload.")
    result, seen = [], set()
    for entry in manifest["files"]:
        if entry["kind"] == "workbook":
            target = workbooks[entry["workbook_index"]]
            if target is None:
                raise BackupError("Original workbook mapping is missing.")
        else:
            target = root.joinpath(*_relative(entry["relative_path"]))
            if any(p is not None and _identity(target) == _identity(p) for p in workbooks):
                raise BackupError("Workbook cannot be replaced through a workspace payload.")
        target = _safe_path(target)
        if target.exists() and not target.is_file():
            raise BackupError(f"Restore target is not a regular file: {target}")
        if _identity(target) in seen:
            raise BackupError("Restore targets collide.")
        seen.add(_identity(target))
        _no_sqlite_sidecars(target)
        result.append(target)
    # Reject parent/child aliases before staging any replacement.
    identities = sorted(seen)
    if any(b.startswith(a + os.sep) for a, b in zip(identities, identities[1:])):
        raise BackupError("Restore targets overlap.")
    return result


def _validate_staged(root, manifest, stage):
    """Restored configuration and executable recovery journals cannot redirect writes."""
    entries = {entry.get("relative_path"): (index, entry)
               for index, entry in enumerate(manifest["files"]) if entry["kind"] == "workspace"}

    def read_json(relative, fallback=False):
        if relative in entries:
            index, entry = entries[relative]
            if entry["size"] > MAX_MANIFEST_BYTES:
                raise BackupError("Settings or issuance journal exceed the safety limit.")
            content = (stage / str(index)).read_bytes()
        elif fallback and (root / relative).is_file():
            content = _safe_path(root / relative).read_bytes()
        else:
            return {}
        try:
            value = json.loads(content.decode("utf-8-sig"))
            if not isinstance(value, dict):
                raise ValueError("expected an object")
            return value
        except (ValueError, UnicodeError) as exc:
            raise BackupError(f"Invalid saved configuration or journal: {relative}") from exc

    # resource(root, ...) prefers the direct default settings file when present.
    defaults = "settings.example.json" if (root / "settings.example.json").exists() or "settings.example.json" in entries else "App/settings.example.json"
    merged = read_json(defaults, fallback=True)
    merged.update(read_json("settings.local.json", fallback=True))
    paths = merged.get("workbook_paths", ["", ""])
    if not isinstance(paths, list) or len(paths) > 2 or any(not isinstance(p, str) for p in paths):
        raise BackupError("Saved settings contain an invalid workbook mapping.")
    configured = _workbooks(root, {"workbook_paths": (paths + ["", ""])[:2]})
    mapping = [str(p) if p is not None else None for p in configured]
    if [_identity(p) if p else None for p in mapping] != [_identity(p) if p else None for p in manifest["original_workbook_paths"]]:
        raise BackupError("Saved settings would change the original configured workbook mapping.")

    allowed_sources = {_identity(p) for p in configured if p is not None}
    allowed_sources.add(_identity(root / "Data" / "Custom Certificates.json"))
    for relative in entries:
        if relative.startswith('Data/Entry Deletions/') and relative.endswith('.json'):
            deletion = read_json(relative)
            identifier = deletion.get('id')
            if (not isinstance(identifier,str) or not re.fullmatch(r'[a-f0-9]{32}',identifier)
                    or relative != f'Data/Entry Deletions/{identifier}.json'
                    or deletion.get('state') not in {'complete', 'rolled_back', 'undone'}):
                raise BackupError('Recover the interrupted entry deletion before backing up or restoring it.')
    for relative in entries:
        if not relative.startswith("Data/Issuance History/") or not relative.endswith(".json"):
            continue
        journal = read_json(relative)
        if journal.get("state") in {"complete", "rolled_back"}:
            continue  # Historical absolute source paths are inert audit evidence.
        identifier = journal.get("id")
        if (journal.get("version") != 1 or not isinstance(identifier, str)
                or not re.fullmatch(r"[a-f0-9]{32}", identifier)
                or relative != f"Data/Issuance History/{identifier}.json"
                or journal.get("state") not in {"prepared", "published", "source_committed"}):
            raise BackupError("An invalid pending issuance journal cannot be restored.")
        source = journal.get("source_path")
        if (not isinstance(source, str) or not Path(source).is_absolute()
                or _identity(_safe_path(source)) not in allowed_sources):
            raise BackupError("A pending issuance journal references an unapproved original source.")

        def certificate_path(value, archive=False):
            parts = _relative(value)
            prefix = ["Data", "Archive", "Superseded Certificates"] if archive else ["Certificates"]
            if parts[:len(prefix)] != prefix or Path(value).suffix.casefold() != ".pdf":
                raise BackupError("A pending issuance journal references an unapproved PDF path.")
            _safe_path(root.joinpath(*parts))

        def snapshot(name, expected):
            relative_snapshot = f"Data/Issuance Recovery/{identifier}/{name}"
            if relative_snapshot not in entries or entries[relative_snapshot][1]["sha256"] != expected:
                raise BackupError("A pending issuance journal is missing its verified recovery snapshot.")

        certificate_path(journal.get("pdf_path"))
        snapshot("new.pdf", journal.get("pdf_sha256"))
        if journal.get("source_before") is not None:
            snapshot("source-before.bin", journal["source_before"])
        previous = journal.get("previous_pdfs")
        if not isinstance(previous, list):
            raise BackupError("Invalid pending issuance PDF history.")
        for old in previous:
            if not isinstance(old, dict) or not isinstance(old.get("snapshot"), str) or not re.fullmatch(r"previous-[0-9]+\.pdf", old["snapshot"]):
                raise BackupError("Unsafe pending issuance recovery snapshot path.")
            certificate_path(old.get("path"))
            certificate_path(old.get("archive"), archive=True)
            snapshot(old["snapshot"], old.get("sha256"))


def _replace_file(source, target):
    """Copy beside the target so replacement works across disks; fsync first."""
    _safe_path(source)
    target = _safe_path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = _safe_path(target.with_name("." + target.name + "." + uuid.uuid4().hex + ".tmp"))
    try:
        with source.open("rb") as input_file, temporary.open("xb") as output:
            shutil.copyfileobj(input_file, output, 1024 * 1024)
            output.flush()
            os.fsync(output.fileno())
        _safe_path(target)
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


def _active_restore_file(relative):
    """Only issuance state can be retired; unrelated files are never swept up."""
    parts = _relative(relative)
    lower = [p.casefold() for p in parts]
    return (_approved(relative) and
            ((lower[0] == "certificates" or lower[:2] == ["output", "certificates"])
             and Path(relative).suffix.casefold() == ".pdf"
             or lower[:2] == ["data", "issuance history"] and Path(relative).suffix.casefold() == ".json"
             or lower[:2] in (["data", "issuance recovery"], ["data", "working drafts"],
                              ["data", "entry deletions"], ["data", "entry deletion recovery"])
             or lower == ["data", "custom certificates.json"]))


def _restore_extras(root, manifest, configuration, extras_root):
    restored = {e["relative_path"].casefold() for e in manifest["files"] if e["kind"] == "workspace"}
    result = []
    for source, entry in _sources(root, _workbooks(root, configuration[0]), allow_missing=True):
        relative = entry.get("relative_path")
        if relative and _active_restore_file(relative) and relative.casefold() not in restored:
            destination = _safe_path(extras_root.joinpath(*_relative(relative)))
            if destination.exists():
                raise BackupError("Restore extras archive path is occupied.")
            result.append({"target": str(source), "archive": str(destination), "sha256": _hash(source)})
    return result


def _inventory_entries(data):
    try:
        value = json.loads(data.decode("utf-8-sig"))
        entries = value["entries"]
        if not isinstance(value, dict) or not isinstance(entries, list):
            raise ValueError("expected an inventory object")
        for entry in entries:
            if (not isinstance(entry, dict) or not isinstance(entry.get("path"), str)
                    or not all(isinstance(entry.get(key), list) and all(isinstance(v, str) for v in entry[key])
                               for key in ("cods", "references", "warnings"))):
                raise ValueError("invalid inventory record")
        if len({e["path"].casefold() for e in entries}) != len(entries):
            raise ValueError("duplicate inventory paths")
        return entries
    except (ValueError, TypeError, KeyError, UnicodeError) as exc:
        raise BackupError("The saved COD inventory is invalid; review its reservation history before restore.") from exc


def _merge_reservations(root, manifest, stage, targets, recovery):
    """Keep all observed CODs, but only the snapshot's ownership and PDF hashes."""
    from certificate_inventory import identity, scan_inventory
    try:
        observed = scan_inventory(root)  # Read-only: includes deleted-file reservations and untracked PDFs.
    except ValueError as exc:
        raise BackupError("The current COD reservation ledger needs review before a safe restore.") from exc
    relative = "Data/Certificate Inventory.json"
    ledger_index = next((i for i, e in enumerate(manifest["files"])
                         if (e.get("relative_path") or "").casefold() == relative.casefold()), None)
    entries = _inventory_entries((stage / str(ledger_index)).read_bytes()) if ledger_index is not None else []
    restored = {e["path"].casefold(): e for e in entries}
    for current in observed["entries"]:
        old = restored.get(current["path"].casefold())
        if old is not None:
            # Never copy managed_owner, current_cods, sha256 or presence from later issuances.
            for field in ("cods", "references"):
                old[field] = sorted(set(old[field]) | set(current[field]))
    codes = {v for e in observed["entries"] for v in e["cods"]}
    for path in (root/'Data'/'Entry Deletions').glob('*.json'):
        deletion = json.loads(_safe_path(path).read_text(encoding='utf-8'))
        if deletion.get('state') == 'complete':
            codes.update(deletion.get('cods', []))
    references = {v for e in observed["entries"] for v in e["references"]}
    # A committed receipt can outlive a missing PDF or an unfinished inventory update.
    history = _safe_path(root / "Data" / "Issuance History")
    for path in history.glob("*.json") if history.is_dir() else ():
        try:
            receipt = json.loads(_safe_path(path).read_text(encoding="utf-8-sig"))
        except (ValueError, UnicodeError) as exc:
            raise BackupError("Current issuance receipts need review before preserving COD reservations.") from exc
        if isinstance(receipt, dict) and isinstance(receipt.get("cod"), str) and receipt["cod"].strip():
            codes.add(receipt["cod"])
    reserved = {identity(v) for e in entries for v in e["cods"]}
    new_codes = sorted(v for v in codes if identity(v) not in reserved)
    new_references = sorted(references - {v for e in entries for v in e["references"]})
    if new_codes or new_references:
        # This is a ledger-only path: never an active PDF, receipt or source owner.
        reservation_path = "Data/Restore COD Reservations.json"
        reservation = restored.get(reservation_path.casefold())
        if reservation is None:
            reservation = {"path": reservation_path, "cods": [], "references": [], "warnings": []}
            entries.append(reservation)
        reservation.update(present=False, current_cods=[], restore_reservation=True)
        reservation.pop("managed_owner", None)
        reservation["cods"] = sorted(set(reservation["cods"]) | set(new_codes))
        reservation["references"] = sorted(set(reservation["references"]) | set(new_references))
    payload = _json_bytes({"version": 1, "entries": sorted(entries, key=lambda e: e["path"].casefold())})
    if len(payload) > MAX_MANIFEST_BYTES:
        raise BackupError("Merged COD reservation ledger exceeds the safety limit.")
    # Preserve the authenticated ledger before replacing its staged copy with the derived union.
    if ledger_index is None:
        ledger_index = len(manifest["files"])
        targets.append(_safe_path(root / relative))
        manifest["files"].append({"kind": "workspace", "relative_path": relative,
                                   "member": f"payload/{ledger_index:06d}"})
    else:
        shutil.copyfile(stage / str(ledger_index), recovery / "inventory-from-backup.json")
    _atomic_bytes(stage / str(ledger_index), payload)
    manifest["files"][ledger_index].update(size=len(payload), sha256=hashlib.sha256(payload).hexdigest())
    return len({identity(v) for v in new_codes})


def _move_extra(source, destination, expected):
    source, destination = _safe_path(source), _safe_path(destination)
    if not source.is_file() or _hash(source) != expected:
        raise BackupError("An issuance extra changed after review; no unknown file was moved.")
    if destination.exists():
        raise BackupError("An issuance extras destination is occupied; no file was overwritten.")
    destination.parent.mkdir(parents=True, exist_ok=True)
    _safe_path(destination)
    os.replace(source, destination)
    if _hash(destination) != expected:
        raise BackupError("Archived issuance extra failed checksum verification.")


def restore_backup(root, backup_id, confirmed=False, *, locked=False):
    """Restore a registered ID after confirmed is literally True.

    Returns restored_files, pre_restore_backup_id/path, recovery_path and warnings.
    Later active PDFs/receipts/recovery snapshots are recoverably archived and
    observed COD reservations are unioned without importing later ownership.
    Unrelated files are kept. Exceptions roll back replacements and archive moves;
    a process kill/power loss leaves a journal and blocks further backup/restore.
    """
    if confirmed is not True:
        raise BackupError("Explicit confirmation is required before restoring a backup.")
    if not isinstance(backup_id, str) or not ID_PATTERN.fullmatch(backup_id):
        raise BackupError("Choose a backup ID from list_backups, never an archive path.")
    root = _root(root)
    with nullcontext() if locked else run_lock(root):
        _pending(root)
        key = _key(root)
        record = next((r for r in _catalog(root, key)["backups"] if r["backup_id"] == backup_id), None)
        if record is None:
            raise BackupError("Unknown backup ID for this workspace.")
        # Inspection and mapping validation happen before creating any transaction.
        manifest, _ = _inspect(root, record, key)
        targets = _targets(root, manifest)
        configuration = _settings(root)  # Capture before Windows denies reads on locked settings.
        recovery = _safe_path(_store(root) / "Restore Transactions" / uuid.uuid4().hex)
        stage, originals = recovery / "staged", recovery / "originals"
        stage.mkdir(parents=True)
        originals.mkdir()
        manifest, _ = _inspect(root, record, key, stage=stage)
        _validate_staged(root, manifest, stage)
        extras_root = _safe_path(root / "Data" / "Archive" / "Restore Extras" / recovery.name)
        extras = _restore_extras(root, manifest, configuration, extras_root)
        reserved_cods = _merge_reservations(root, manifest, stage, targets, recovery)
        journal = {"state": "prepared", "backup_id": backup_id, "files": [], "extras": extras,
                   "extras_path": str(extras_root), "reserved_cods": reserved_cods}
        attempted, moved = [], []
        with ExitStack() as locks:
            held = {}
            for index, target in enumerate(targets):
                _safe_path(target)
                existed = target.exists()
                if existed:
                    held[target] = locks.enter_context(exclusive_file(target))
                    with (originals / str(index)).open("xb") as saved:
                        held[target].seek(0)
                        shutil.copyfileobj(held[target], saved, 1024 * 1024)
                        saved.flush()
                        os.fsync(saved.fileno())
                journal["files"].append({"target": str(target), "existed": existed,
                                          "original": str(originals / str(index)),
                                          "staged": str(stage / str(index))})
            for extra in extras:
                target = _safe_path(extra["target"])
                held[target] = locks.enter_context(exclusive_file(target))
            snapshot = _create(root, _store(root), allow_missing=True, held=held,
                               purpose="pre-restore", configuration=configuration)
            journal["pre_restore_backup_id"] = snapshot["backup_id"]
            _atomic_bytes(recovery / "journal.json", _json_bytes(journal))
            try:
                journal["state"] = "in_progress"
                _atomic_bytes(recovery / "journal.json", _json_bytes(journal))
                for extra in extras:
                    held[Path(extra["target"])].close()
                    moved.append(extra)  # Include a move which fails after rename in rollback.
                    _move_extra(extra["target"], extra["archive"], extra["sha256"])
                for index, target in enumerate(targets):
                    # Recheck links immediately before each replacement.
                    _safe_path(target)
                    if target.exists() != journal["files"][index]["existed"]:
                        raise BackupError(f"A restore target appeared or disappeared: {target}")
                    if target in held:
                        # Windows atomic replacement requires closing our write handle.
                        # run_lock still serializes the backend; a newly opened Excel
                        # handle makes replacement fail and invokes rollback.
                        held[target].close()
                        if _hash(target) != _hash(originals / str(index)):
                            raise BackupError(f"A restore target changed after staging: {target}")
                    attempted.append(index)
                    _replace_file(stage / str(index), target)
                    if _hash(target) != manifest["files"][index]["sha256"]:
                        raise BackupError(f"Restored file failed verification: {target}")
                journal["state"] = "committed"
                _atomic_bytes(recovery / "journal.json", _json_bytes(journal))
            except BaseException as exc:
                failures = []
                for index in reversed(attempted):
                    target = targets[index]
                    try:
                        if journal["files"][index]["existed"]:
                            _replace_file(originals / str(index), target)
                            if _hash(target) != _hash(originals / str(index)):
                                raise BackupError("Rollback checksum mismatch.")
                        elif target.exists():
                            # Recoverable move, never delete a newly restored material file.
                            _safe_path(target)
                            preserved = target.with_name(target.name + ".restore-recovery-" + recovery.name)
                            if preserved.exists():
                                raise BackupError("Rollback recovery filename is occupied.")
                            os.replace(target, preserved)
                            journal["files"][index]["preserved_restored_file"] = str(preserved)
                    except BaseException as rollback_exc:
                        failures.append(f"{target}: {rollback_exc}")
                for extra in reversed(moved):
                    try:
                        archived, target = _safe_path(extra["archive"]), _safe_path(extra["target"])
                        if archived.exists():
                            _move_extra(archived, target, extra["sha256"])
                        elif not target.is_file() or _hash(target) != extra["sha256"]:
                            raise BackupError("An issuance extra could not be located for rollback.")
                    except BaseException as rollback_exc:
                        failures.append(f"{extra['target']}: {rollback_exc}")
                journal.update(state="rollback_failed" if failures else "rolled_back", errors=failures)
                try:
                    _atomic_bytes(recovery / "journal.json", _json_bytes(journal))
                except OSError as journal_exc:
                    failures.append(f"Could not save rollback journal: {journal_exc}")
                raise RestoreError(f"Restore failed: {exc}. " +
                                   ("Rollback needs supervised recovery: " + "; ".join(failures)
                                    if failures else "All attempted replacements were rolled back."),
                                   recovery_path=recovery, rolled_back=not failures) from exc
        return {"backup_id": backup_id, "restored_files": len(targets),
                "archived_files": len(extras), "extras_path": str(extras_root), "reserved_cods": reserved_cods,
                "pre_restore_backup_id": snapshot["backup_id"],
                "pre_restore_path": snapshot["path"], "recovery_path": str(recovery),
                "warnings": ["Later active certificate PDFs and issuance receipts/snapshots were archived recoverably.",
                             "Historical COD reservations were retained without adopting later source ownership.",
                             "Refresh the dashboard and verify COD continuity before issuing."]}
