"""Retain active editor previews and expire disposable PDF/image groups."""
from pathlib import Path
import re
import threading
import time

_ACTIVE = {}
_CLIENTS = {}
_LOCK = threading.RLock()
_LAST_CLEANUP = {}
PATTERN = re.compile(r'^([a-f0-9]{32})(?:_page_[0-9]+)?\.(?:pdf|png)$')


def retain_preview(root, identifier, session=None, client=None):
    root = str(Path(root).resolve())
    session = session if isinstance(session, str) and re.fullmatch(r'[A-Za-z0-9_-]{1,80}', session) else identifier
    with _LOCK:
        now = time.monotonic()
        _ACTIVE[(root, session)] = (identifier, now)
        if isinstance(client, str) and 1 <= len(client) <= 64:
            previous = _CLIENTS.get((root, client))
            if previous and previous != session:
                _ACTIVE.pop((root, previous), None)
            _CLIENTS[(root, client)] = session
            while len(_CLIENTS) > 256:
                _CLIENTS.pop(next(iter(_CLIENTS)))
        for key, (_identifier, seen) in list(_ACTIVE.items()):
            if now - seen > 4 * 3600:
                _ACTIVE.pop(key, None)
        while len(_ACTIVE) > 128:
            _ACTIVE.pop(min(_ACTIVE, key=lambda key: _ACTIVE[key][1]), None)


def preview_presence(root, client, session=None):
    root = str(Path(root).resolve())
    with _LOCK:
        previous = _CLIENTS.get((root, client))
        if previous and previous != session:
            _ACTIVE.pop((root, previous), None)
            _CLIENTS.pop((root, client), None)
        if isinstance(session, str) and (root, session) in _ACTIVE:
            identifier, _seen = _ACTIVE[(root, session)]
            _ACTIVE[(root, session)] = (identifier, time.monotonic())
            _CLIENTS[(root, client)] = session


def touch_preview(root, identifier):
    root = str(Path(root).resolve())
    with _LOCK:
        for key, (current, _seen) in list(_ACTIVE.items()):
            if key[0] == root and current == identifier:
                _ACTIVE[key] = (current, time.monotonic())


def cleanup_previews(root, *, max_bytes=256 * 1024 * 1024, max_age=24 * 3600, force=False):
    resolved_root = Path(root).resolve()
    intended = resolved_root / 'Data' / 'Previews'
    folder = intended.resolve()
    if folder != intended:
        return  # Never follow a redirected preview folder into another store.
    root_key = str(resolved_root)
    with _LOCK:
        tick = time.monotonic()
        if not force and tick - _LAST_CLEANUP.get(root_key, -float('inf')) < 30:
            return
        _LAST_CLEANUP[root_key] = tick
        active = {identifier for (key, _session), (identifier, seen) in _ACTIVE.items()
                  if key == root_key and tick - seen <= 4 * 3600}
        now = time.time()
        groups = {}
        for path in folder.glob('*'):
            match = PATTERN.fullmatch(path.name)
            if not match or path.is_symlink() or path.resolve().parent != folder or not path.is_file():
                continue
            try:
                stat = path.stat()
            except FileNotFoundError:
                continue
            group = groups.setdefault(match[1], {'files': [], 'size': 0, 'mtime': 0})
            group['files'].append(path)
            group['size'] += stat.st_size
            group['mtime'] = max(group['mtime'], stat.st_mtime)
        total = sum(group['size'] for group in groups.values())
        for identifier, group in sorted(groups.items(), key=lambda item: item[1]['mtime']):
            age = now - group['mtime']
            if identifier in active or age < 60:
                continue
            if age <= max_age and total <= max_bytes:
                continue
            for path in group['files']:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass  # A browser or scanner may still hold this artifact.
            total -= group['size']
