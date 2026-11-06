"""Bounded digest reuse for display scans; publication can force fresh reads."""
from collections import OrderedDict
from functools import lru_cache
import hashlib
import os
from pathlib import Path
import threading
import time

_CACHE = OrderedDict()
_LOCK = threading.RLock()
MAX_ENTRIES = 2048
MAX_AGE = 60


@lru_cache(maxsize=1)
def _windows_api():
    import ctypes
    from ctypes import wintypes
    class BasicInfo(ctypes.Structure):
        _fields_ = [('creation', ctypes.c_longlong), ('access', ctypes.c_longlong),
                    ('write', ctypes.c_longlong), ('change', ctypes.c_longlong),
                    ('attributes', wintypes.DWORD)]
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    query = kernel.GetFileInformationByHandleEx
    query.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
    query.restype = wintypes.BOOL
    close = kernel.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    return BasicInfo, create, query, close


def _windows_change_time(path):
    """Windows st_ctime is creation time; use the filesystem's change time."""
    import ctypes
    BasicInfo, create, query, close = _windows_api()
    handle = create(str(path), 0x80, 7, None, 3, 0, None)
    if handle == ctypes.c_void_p(-1).value:
        return None  # Unsupported metadata must disable digest reuse.
    try:
        data = BasicInfo()
        if not query(handle, 0, ctypes.byref(data), ctypes.sizeof(data)):
            return None
        return data.change
    finally:
        close(handle)


def file_signature(path):
    path = Path(path)
    stat = path.stat()
    change = _windows_change_time(path.resolve()) if os.name == 'nt' else stat.st_ctime_ns
    return (stat.st_dev, stat.st_ino, stat.st_size, stat.st_mtime_ns, change)


def fingerprint(path, *, force=False, payload_limit=None):
    """Return a digest and optional exact bytes; never retain PDF bytes in cache."""
    path = Path(path)
    key = str(path.resolve())
    before = file_signature(path)
    now = time.monotonic()
    with _LOCK:
        cached = _CACHE.get(key)
        if not force and before[-1] is not None and cached and cached[0] == before and now - cached[2] < MAX_AGE:
            _CACHE.move_to_end(key)
            return cached[1], None
    payload = None
    with path.open('rb') as stream:
        if payload_limit is not None and before[2] <= payload_limit:
            candidate = stream.read(payload_limit + 1)
            checksum = hashlib.sha256(candidate)
            if len(candidate) <= payload_limit:
                payload = candidate
            else:
                for chunk in iter(lambda: stream.read(1024 * 1024), b''):
                    checksum.update(chunk)
            digest = checksum.hexdigest()
        else:
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
    if file_signature(path) != before:
        raise OSError(f'{path.name} changed while it was being checked. Retry the operation.')
    with _LOCK:
        _CACHE[key] = (before, digest, now)
        _CACHE.move_to_end(key)
        while len(_CACHE) > MAX_ENTRIES:
            _CACHE.popitem(last=False)
    return digest, payload
