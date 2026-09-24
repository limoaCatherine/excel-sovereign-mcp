"""excelcli sessions. This module is the only caller of excelcli."""

from __future__ import annotations

import json
import os
import subprocess
import threading
import time
from pathlib import Path

from excel_sovereign.lock import file_is_locked

COM_QUEUE = threading.Lock()
ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TIMEOUT = int(os.environ.get("EXCEL_MCP_COM_TIMEOUT", "180"))


class ExcelCliError(RuntimeError):
    def __init__(self, message: str, *, results: list[dict] | None = None, pid_known: bool = True):
        super().__init__(message)
        self.results = results or []
        self.pid_known = pid_known


def excelcli_path() -> Path:
    env = os.environ.get("EXCELCLI")
    candidates = []
    if env:
        candidates.append(Path(env))
    base = ROOT / "vendor" / "mcp-server-excel" / "src" / "ExcelMcp.CLI" / "bin"
    candidates.extend(
        [
            base / "Release" / "net10.0-windows" / "excelcli.exe",
            base / "Debug" / "net10.0-windows" / "excelcli.exe",
        ]
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError("excelcli.exe was not found. Build vendor/mcp-server-excel or set EXCELCLI.")


def excel_pids() -> set[int]:
    script = "Get-Process excel -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id"
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-Command", script],
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )
    pids: set[int] = set()
    for line in completed.stdout.splitlines():
        line = line.strip()
        if line.isdigit():
            pids.add(int(line))
    return pids


def _short_error(text: str) -> str:
    lines = []
    for line in str(text).splitlines():
        lowered = line.strip()
        if lowered.startswith("at ") or "StackTrace" in lowered:
            continue
        lines.append(lowered)
        if len(lines) >= 4:
            break
    message = " ".join(lines).strip() or "excel command failed"
    return message[:500]


def run_batch(commands: list[dict], timeout: int) -> list[dict]:
    cli = excelcli_path()
    payload = json.dumps(commands, ensure_ascii=False)
    directory = Path(os.environ.get("TEMP") or ".") / "excel-sovereign-mcp-batch"
    directory.mkdir(parents=True, exist_ok=True)
    batch_path = directory / f"batch-{os.getpid()}-{time.time_ns()}.json"
    batch_path.write_text(payload, encoding="utf-8")
    try:
        completed = subprocess.run(
            [str(cli), "-q", "batch", "--stop-on-error", "-i", str(batch_path)],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.TimeoutExpired as exc:
        raise ExcelCliError("excelcli timed out", pid_known=False) from exc
    finally:
        try:
            batch_path.unlink(missing_ok=True)
        except OSError:
            pass
    results = []
    for line in (completed.stdout or "").splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            results.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    if completed.returncode != 0 and not results:
        message = _short_error(completed.stderr or completed.stdout or "excelcli failed")
        raise ExcelCliError(message, results=results)
    return results


def _session_id(results: list[dict]) -> str | None:
    for item in results:
        result = item.get("result")
        if isinstance(result, dict) and result.get("sessionId"):
            return str(result["sessionId"])
        if isinstance(result, str):
            try:
                parsed = json.loads(result)
            except json.JSONDecodeError:
                continue
            if isinstance(parsed, dict) and parsed.get("sessionId"):
                return str(parsed["sessionId"])
    return None


def _failed(results: list[dict]) -> dict | None:
    for item in results:
        if item.get("command") == "session.open" and item.get("success"):
            continue
        if not item.get("success", True):
            return item
    return None


def kill_pids(pids: set[int]) -> bool:
    """Kill only the Excel processes started for this call. Returns False if none were known."""
    if not pids:
        return False
    for pid in pids:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    return True


def close_session(session_id: str, save: bool, timeout: int = 60) -> None:
    run_batch(
        [{"command": "session.close", "sessionId": session_id, "args": {"save": save}}],
        timeout=timeout,
    )


class ComSession:
    """One Excel session. The caller holds COM_QUEUE and the workbook lock."""

    def __init__(self, path: str, *, allow_macros: bool = False, timeout: int | None = None):
        self.path = path
        self.allow_macros = allow_macros
        self.timeout = timeout or DEFAULT_TIMEOUT
        self.session_id: str | None = None
        self.pids: set[int] = set()
        self.pid_known = False

    def open(self) -> None:
        before = excel_pids()
        results = run_batch(
            [
                {
                    "command": "session.open",
                    "args": {
                        "filePath": self.path,
                        "allowMacros": self.allow_macros,
                        "timeoutSeconds": max(10, min(self.timeout, 3600)),
                    },
                }
            ],
            timeout=self.timeout,
        )
        self.pids = excel_pids() - before
        self.pid_known = bool(self.pids)
        failure = _failed(results)
        if failure is not None:
            raise ExcelCliError(_short_error(str(failure.get("error") or failure)), results=results, pid_known=self.pid_known)
        self.session_id = _session_id(results)
        if not self.session_id:
            raise ExcelCliError("session id missing", results=results, pid_known=self.pid_known)

    def call(self, commands: list[dict]) -> list[dict]:
        if not self.session_id:
            raise ExcelCliError("session is not open", pid_known=self.pid_known)
        wrapped = []
        for command in commands:
            item = dict(command)
            item["sessionId"] = self.session_id
            wrapped.append(item)
        results = run_batch(wrapped, timeout=self.timeout)
        failure = _failed(results)
        if failure is not None:
            raise ExcelCliError(_short_error(str(failure.get("error") or failure)), results=results, pid_known=self.pid_known)
        return results

    def close(self, save: bool) -> None:
        if not self.session_id:
            return
        session_id = self.session_id
        self.session_id = None
        close_session(session_id, save=save, timeout=self.timeout)

    def abort(self) -> bool:
        """Close without saving. Kill only PIDs started for this session.

        Returns True when the workbook is no longer locked. False means the PID
        was unknown and the caller must keep the file lock.
        """
        try:
            self.close(save=False)
        except ExcelCliError:
            self.session_id = None
        if not file_is_locked(self.path):
            return True
        if not self.pid_known:
            return False
        kill_pids(self.pids)
        for _ in range(40):
            if not file_is_locked(self.path):
                return True
            time.sleep(0.25)
        return False
