"""Cross-process workbook lock.

The lock file lives in the system temp directory. Exclusion is an exclusive
create plus a handle kept open for the critical section. Other processes may
read the record. A live owner is never taken over because its heartbeat is
stale; takeover requires the recorded process to be gone or its start time to
differ (PID reuse).
"""

from __future__ import annotations

import ctypes
import json
import os
import threading
import time
from ctypes import wintypes
from hashlib import sha256
from pathlib import Path

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.CreateFileW.restype = wintypes.HANDLE
kernel32.CreateFileW.argtypes = [
    wintypes.LPCWSTR,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.LPVOID,
    wintypes.DWORD,
    wintypes.DWORD,
    wintypes.HANDLE,
]
kernel32.GetFileSize.restype = wintypes.DWORD
kernel32.ReadFile.argtypes = [wintypes.HANDLE, wintypes.LPVOID, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID]
kernel32.WriteFile.argtypes = [wintypes.HANDLE, wintypes.LPCVOID, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD), wintypes.LPVOID]

GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
FILE_SHARE_READ = 0x00000001
FILE_SHARE_WRITE = 0x00000002
CREATE_NEW = 1
OPEN_EXISTING = 3
FILE_ATTRIBUTE_NORMAL = 0x80
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

_locks_guard = threading.Lock()
_retained: dict[str, "FileLock"] = {}


class FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", wintypes.DWORD),
        ("dwHighDateTime", wintypes.DWORD),
    ]


class LockTimeout(TimeoutError):
    pass


def canonical_path(path: str | os.PathLike[str]) -> str:
    return os.path.normcase(os.path.abspath(os.fspath(path)))


def _lock_dir() -> Path:
    directory = Path(os.environ.get("TEMP") or os.environ.get("TMP") or ".") / "excel-sovereign-mcp-locks"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def _lock_path(workbook: str) -> Path:
    digest = sha256(canonical_path(workbook).encode("utf-8")).hexdigest()
    return _lock_dir() / f"{digest}.lock"


def process_start_time(pid: int) -> int | None:
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    creation = FILETIME()
    exit_time = FILETIME()
    kernel_time = FILETIME()
    user_time = FILETIME()
    try:
        ok = kernel32.GetProcessTimes(
            handle,
            ctypes.byref(creation),
            ctypes.byref(exit_time),
            ctypes.byref(kernel_time),
            ctypes.byref(user_time),
        )
        if not ok:
            return None
        return (int(creation.dwHighDateTime) << 32) | int(creation.dwLowDateTime)
    finally:
        kernel32.CloseHandle(handle)


def _owner_alive(record: dict) -> bool:
    pid = int(record.get("pid") or 0)
    recorded = int(record.get("start") or 0)
    if pid <= 0:
        return False
    actual = process_start_time(pid)
    if actual is None:
        return False
    return actual == recorded


def _create_file(path: str, disposition: int, access: int, share: int) -> int | None:
    handle = kernel32.CreateFileW(path, access, share, None, disposition, FILE_ATTRIBUTE_NORMAL, None)
    if handle == INVALID_HANDLE_VALUE or handle is None:
        return None
    return int(handle)


def _read_record(path: Path) -> dict | None:
    handle = _create_file(str(path), OPEN_EXISTING, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE)
    if handle is None:
        return None
    try:
        size = kernel32.GetFileSize(handle, None)
        if size in (0, 0xFFFFFFFF):
            return None
        buffer = ctypes.create_string_buffer(size)
        read = wintypes.DWORD()
        ok = kernel32.ReadFile(handle, buffer, size, ctypes.byref(read), None)
        if not ok:
            return None
        return json.loads(buffer.raw[: read.value].decode("utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return None
    finally:
        kernel32.CloseHandle(handle)


def _write_handle(handle: int, payload: bytes) -> None:
    kernel32.SetFilePointer(handle, 0, None, 0)
    written = wintypes.DWORD()
    if not kernel32.WriteFile(handle, payload, len(payload), ctypes.byref(written), None):
        raise OSError("failed to write lock record")
    kernel32.SetEndOfFile(handle)


class FileLock:
    def __init__(self, workbook: str, timeout: float = 1800.0):
        self.workbook = canonical_path(workbook)
        self.timeout = timeout
        self._path = _lock_path(self.workbook)
        self._handle: int | None = None
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self.retained = False

    def acquire(self) -> None:
        deadline = time.monotonic() + self.timeout
        while True:
            handle = _create_file(
                str(self._path),
                CREATE_NEW,
                GENERIC_READ | GENERIC_WRITE,
                FILE_SHARE_READ,
            )
            if handle is not None:
                self._handle = handle
                self._write_record()
                self._thread = threading.Thread(target=self._heartbeat, name="excel-sovereign-mcp-lock", daemon=True)
                self._thread.start()
                return
            record = _read_record(self._path)
            if record is None or not _owner_alive(record):
                try:
                    self._path.unlink(missing_ok=True)
                except OSError:
                    pass
            if time.monotonic() >= deadline:
                raise LockTimeout(self.workbook)
            time.sleep(0.2)

    def _payload(self) -> bytes:
        record = {
            "pid": os.getpid(),
            "start": process_start_time(os.getpid()) or 0,
            "heartbeat": time.time(),
            "path": self.workbook,
        }
        return json.dumps(record, ensure_ascii=False).encode("utf-8")

    def _write_record(self) -> None:
        if self._handle is None:
            return
        _write_handle(self._handle, self._payload())

    def _heartbeat(self) -> None:
        while not self._stop.wait(2.0):
            try:
                self._write_record()
            except OSError:
                return

    def release(self) -> None:
        if self.retained:
            return
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=3)
        if self._handle is not None:
            kernel32.CloseHandle(self._handle)
            self._handle = None
        try:
            self._path.unlink(missing_ok=True)
        except OSError:
            pass

    def retain_until_file_free(self) -> None:
        """Keep this lock until the workbook can be opened exclusively."""
        self.retained = True
        with _locks_guard:
            _retained[self.workbook] = self
        threading.Thread(target=self._watch_free, name="excel-sovereign-mcp-lock-retain", daemon=True).start()

    def _watch_free(self) -> None:
        while True:
            if not file_is_locked(self.workbook):
                self.retained = False
                self.release()
                with _locks_guard:
                    _retained.pop(self.workbook, None)
                return
            time.sleep(1.0)

    def __enter__(self) -> "FileLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.release()


def file_is_locked(path: str) -> bool:
    """True when the file cannot be opened with FileShare.None."""
    if not os.path.exists(path):
        return False
    handle = kernel32.CreateFileW(
        canonical_path(path),
        GENERIC_READ | GENERIC_WRITE,
        0,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None,
    )
    if handle == INVALID_HANDLE_VALUE or not handle:
        return True
    kernel32.CloseHandle(handle)
    return False
