"""Six tools and the write pipeline. Models do not choose an engine or hold a session."""

from __future__ import annotations

import json
import os
from typing import Any

from mcp.server.mcpserver import Image, MCPServer

from excel_sovereign.book import (
    apply_ops,
    create_empty,
    discard,
    load_for_edit,
    read_page,
    replace_target,
    save_atomic,
    trim_ops,
)
from excel_sovereign.excel_cli import COM_QUEUE, ComSession, ExcelCliError
from excel_sovereign.inspect import is_encrypted_or_irm, inspect_workbook
from excel_sovereign.layout import expand_layouts
from excel_sovereign.lock import FileLock, LockTimeout, canonical_path, file_is_locked
from excel_sovereign.route import (
    COM_COMMANDS,
    MAX_WRITE_CELLS,
    QUERY_ACTIONS,
    RouteDecision,
    com_args,
    count_cells,
    decide,
    action_outside,
    normalize_ops,
    shift_address,
    validate_known,
)
from excel_sovereign.verify import (
    VerifyOutcome,
    _object_anchors,
    calculate,
    capture_ranges,
    plan_sheets,
    spill_addresses,
    union_ranges,
)

mcp = MCPServer(
    "excel-sovereign-mcp",
    title="excelMCP",
    instructions=(
        "Local Excel server with six tools and one ops list. "
        "The server chooses openpyxl or Excel, saves once, and screenshots. "
        "Close the workbook in Excel before calling."
    ),
)

READ_LIMIT = 4000
ALLOWED = {".xlsx", ".xlsm"}


def _error(code: str, message: str, **extra) -> dict:
    body = {
        "ok": False,
        "committed": False,
        "saved": False,
        "verified": False,
        "calculated": False,
        "calc": "skipped",
        "engine": None,
        "completedOps": 0,
        "failedAt": None,
        "wrote": [],
        "screenshots": [],
        "omitted": [],
        "suggestedNextActions": [],
        "error": {"code": code, "message": message},
    }
    body.update(extra)
    return body


def _extension(path: str) -> str | None:
    suffix = os.path.splitext(path)[1].lower()
    if suffix not in ALLOWED:
        return suffix or "(none)"
    return None


def _response(
    *,
    path: str,
    engine: str,
    tool: str,
    committed: bool,
    calc: str,
    calc_required: bool,
    wrote: list[dict],
    outcome: VerifyOutcome | None,
    completed: int,
    failed_at: int | None,
    error: dict | None,
) -> tuple[dict, list]:
    shots = outcome.shots if outcome else []
    omitted = outcome.omitted if outcome else []
    spill_unknown = bool(outcome and outcome.spill_unknown)
    cropped = any(shot.cropped for shot in shots)
    verified = bool(
        outcome
        and outcome.error is None
        and (shots or outcome.nothing_to_show)
        and not omitted
        and not spill_unknown
        and not cropped
        and (not calc_required or calc == "done")
    )
    calculated = calc == "done"
    ok = committed and verified and (not calc_required or calculated)
    if error:
        ok = False
        verified = False
    ranges = {shot.sheet: shot.range for shot in shots}
    suggestions = []
    if committed and not verified:
        suggestions.append({"tool": tool, "ops": [{"action": "verify", "ranges": ranges}]})
    body = {
        "ok": ok,
        "committed": committed,
        "saved": committed,
        "verified": verified,
        "calculated": calculated,
        "calc": calc,
        "engine": engine,
        "path": path,
        "completedOps": completed,
        "failedAt": failed_at,
        "wrote": wrote,
        "screenshots": [
            {
                "sheet": shot.sheet,
                "range": shot.range,
                "cropped": shot.cropped,
                "spillUnknown": shot.spill_unknown or spill_unknown,
            }
            for shot in shots
        ],
        "omitted": omitted,
        "suggestedNextActions": suggestions,
    }
    if error:
        body["error"] = error
    if outcome and outcome.error and "error" not in body:
        body["error"] = {"code": "verify_failed", "message": outcome.error}
    return body, shots


def _track(ops: list[dict]) -> tuple[list[dict], list[dict], list[str], dict[str, str]]:
    wrote: list[dict] = []
    formulas: list[dict] = []
    touched: list[str] = []
    renamed: dict[str, str] = {}

    def remember(sheet: str) -> None:
        if sheet and sheet not in touched:
            touched.append(sheet)

    def forget(name: str) -> None:
        wrote[:] = [item for item in wrote if item["sheet"] != name]
        formulas[:] = [item for item in formulas if item["sheet"] != name]
        if name in touched:
            touched.remove(name)

    for op in ops:
        action = op["action"]
        args = op.get("args") if isinstance(op.get("args"), dict) else {}
        sheet = str(op.get("sheet") or args.get("sheetName") or "")
        target = op.get("range") or op.get("cell") or args.get("rangeAddress") or args.get("cellAddress")
        if action in {"insert_rows", "delete_rows", "insert_columns", "delete_columns"} and sheet:
            for item in wrote:
                if item["sheet"] == sheet:
                    item["range"] = shift_address(item["range"], op)
            for item in formulas:
                if item["sheet"] == sheet:
                    item["cell"] = shift_address(item["cell"], op).split(":")[0]
        if action == "rename_sheet":
            old = str(op.get("oldName") or sheet)
            new = str(op.get("newName"))
            renamed[new] = old
            for item in wrote:
                if item["sheet"] == old:
                    item["sheet"] = new
            for item in formulas:
                if item["sheet"] == old:
                    item["sheet"] = new
            if old in touched:
                touched[touched.index(old)] = new
            else:
                remember(new)
            continue
        if action == "set_values":
            address = _block_range(str(op.get("range")), op.get("values") or [])
            wrote.append({"sheet": sheet, "range": address})
            remember(sheet)
            for cell, formula in _formula_cells(sheet, str(op.get("range")), op.get("values") or []):
                formulas.append({"sheet": sheet, "cell": cell, "formula": formula})
        elif action == "set_formulas":
            address = _block_range(str(op.get("range")), op.get("formulas") or [])
            wrote.append({"sheet": sheet, "range": address})
            remember(sheet)
            for cell, formula in _formula_cells(sheet, str(op.get("range")), op.get("formulas") or []):
                formulas.append({"sheet": sheet, "cell": cell, "formula": formula})
        elif action in {"delete_sheet", "sheet_hide"}:
            # Nothing left to photograph on a deleted or hidden sheet.
            forget(str(sheet or op.get("name") or ""))
        elif action in {"create_sheet", "copy_sheet", "move_sheet", "sheet_show", "freeze", "unfreeze"}:
            name = str(op.get("newName") or op.get("name") or args.get("targetName") or sheet)
            if action == "copy_sheet" and not op.get("newName") and not args.get("targetName"):
                name = str(op.get("source") or args.get("sourceName") or sheet)
            if name:
                wrote.append({"sheet": name, "range": "A1:L40"})
                remember(name)
        elif target and sheet:
            wrote.append({"sheet": sheet, "range": str(target)})
            remember(sheet)
        elif action in {"insert_rows", "delete_rows"}:
            row = int(op["row"])
            count = int(op.get("count") or 1)
            wrote.append({"sheet": sheet, "range": f"A{row}:A{row + count - 1}"})
            remember(sheet)
        elif action in {"insert_columns", "delete_columns"}:
            wrote.append({"sheet": sheet, "range": f"{op['column']}1"})
            remember(sheet)
    return wrote, formulas, touched, renamed


def _block_range(start: str, rows: list) -> str:
    from openpyxl.utils import get_column_letter, range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(start)
    height = max(len(rows), 1)
    width = max((len(row) for row in rows if isinstance(row, list)), default=1)
    end_row = max(max_row, min_row + height - 1)
    end_col = max(max_col, min_col + width - 1)
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(end_col)}{end_row}"


def _formula_cells(sheet: str, start: str, rows: list) -> list[tuple[str, str]]:
    from openpyxl.utils import get_column_letter, range_boundaries

    min_col, min_row, _, _ = range_boundaries(start)
    found = []
    for r_index, row in enumerate(rows):
        if not isinstance(row, list):
            continue
        for c_index, value in enumerate(row):
            if isinstance(value, str) and value.startswith("="):
                found.append((f"{get_column_letter(min_col + c_index)}{min_row + r_index}", value))
    del sheet
    return found


def _run_verify(
    session: ComSession,
    path: str,
    *,
    calc_required: bool,
    wrote: list[dict],
    formula_cells: list[dict],
    touched: list[str],
    renamed: dict[str, str],
    explicit_ranges: dict[str, str] | None = None,
) -> VerifyOutcome:
    outcome = VerifyOutcome(calc="skipped", calculated=False)
    try:
        if calc_required:
            calculate(session)
            outcome.calc = "done"
            outcome.calculated = True
        ranges, omitted, _ = plan_sheets(path, wrote, formula_cells, touched, renamed)
        outcome.omitted = omitted
        for sheet, pieces in _object_anchors(session).items():
            if touched and sheet not in touched and sheet not in ranges:
                continue
            merged = union_ranges([ranges.get(sheet, ""), *pieces])
            if merged:
                ranges[sheet] = merged
        if explicit_ranges:
            ranges.update(explicit_ranges)
        spill_unknown = False
        probes = [(item["sheet"], item["cell"]) for item in formula_cells]
        for (sheet, _), (supported, address) in zip(probes, spill_addresses(session, probes)):
            if not supported:
                spill_unknown = True
                continue
            if address:
                ranges[sheet] = union_ranges([ranges.get(sheet, address), address]) or address
        outcome.spill_unknown = spill_unknown
        if not ranges and touched:
            for shot in capture_ranges(session, [(sheet, "A1:L40") for sheet in touched]):
                shot.cropped = True
                shot.spill_unknown = spill_unknown
                outcome.shots.append(shot)
        for shot in capture_ranges(session, [(sheet, address) for sheet, address in ranges.items() if address]):
            shot.spill_unknown = spill_unknown
            outcome.shots.append(shot)
        reasons = []
        if spill_unknown:
            reasons.append("spill range unavailable")
        if omitted:
            reasons.append("dependent ranges omitted")
        if any(shot.cropped for shot in outcome.shots):
            reasons.append("screenshot cropped")
        if not outcome.shots and not ranges and not touched:
            outcome.nothing_to_show = True
        elif not outcome.shots:
            reasons.append("no screenshot range")
        if reasons:
            outcome.error = "; ".join(reasons)
    except Exception as exc:
        if calc_required and outcome.calc != "done":
            outcome.calc = "failed"
            outcome.calculated = False
            outcome.error = str(exc)
        else:
            outcome.error = str(exc)
    return outcome


def apply_workbook(path: str, ops: list[dict], tool: str = "workbook_apply") -> tuple[dict, list]:
    try:
        ops = normalize_ops(ops)
    except ValueError as exc:
        return _error("invalid_ops", str(exc), path=path), []
    foreign = action_outside(tool, ops)
    if foreign:
        action, owner = foreign
        return _error(
            "wrong_tool",
            f"{action} belongs to {owner}",
            path=path,
            suggestedNextActions=[{"tool": owner, "ops": [{"action": action}]}],
        ), []
    try:
        ops = expand_layouts(ops)
    except ValueError as exc:
        return _error("invalid_ops", str(exc), path=path), []
    unknown = validate_known(ops)
    if unknown:
        return _error("unsupported", f"unsupported action: {unknown}", path=path), []
    if any(op["action"] == "verify" for op in ops) and len(ops) != 1:
        return _error("verify_must_be_alone", "verify cannot be combined with other ops", path=path), []
    if count_cells(ops) > MAX_WRITE_CELLS:
        return _error("too_large", f"a single call can write at most {MAX_WRITE_CELLS} cells", path=path), []
    path = canonical_path(path)
    bad = _extension(path)
    if bad:
        return _error("unsupported_format", f"only .xlsx and .xlsm are accepted, got {bad}", path=path), []

    creating = not os.path.exists(path)
    if creating:
        if not ops or ops[0]["action"] != "create_workbook":
            return _error("not_found", "workbook does not exist", path=path), []
        create_empty(path)
        ops = [op for op in ops if op["action"] != "create_workbook"]
        if not ops:
            ops = [{"action": "verify", "ranges": {"Sheet1": "A1"}}]
    elif is_encrypted_or_irm(path):
        return _error("encrypted_or_irm", "encrypted or IRM workbooks are rejected before modification", path=path), []

    try:
        lock = FileLock(path)
        lock.acquire()
    except LockTimeout:
        return _error("lock_timeout", "timed out waiting for the workbook lock", path=path), []

    try:
        if file_is_locked(path):
            return _error("file_locked", "workbook is open in Excel or another process", path=path), []
        if any(op["action"] == "trim_sheet" for op in ops):
            try:
                ops = _expand_trims(path, ops)
            except KeyError as exc:
                return _error("invalid_ops", str(exc).strip("'\""), path=path), []
            if not ops:
                return _nothing_to_do(path, "nothing lies past the data and the ranges formulas, names, or rules still reference"), []
        info = inspect_workbook(path)
        if ops and ops[0]["action"] == "verify" and len(ops) == 1:
            decision = RouteDecision("com", True, False, "verify")
        else:
            decision = decide(ops, info)
        if decision.engine == "openpyxl":
            body, shots = _apply_openpyxl(path, ops, tool, decision, lock)
        else:
            body, shots = _apply_com(path, ops, tool, decision, lock)
        return body, shots
    finally:
        lock.release()


def _expand_trims(path: str, ops: list[dict]) -> list[dict]:
    expanded = []
    for op in ops:
        if op["action"] != "trim_sheet":
            expanded.append(op)
            continue
        sheet = str(op.get("sheet") or "")
        if not sheet:
            raise KeyError("trim_sheet needs sheet")
        expanded.extend(trim_ops(path, sheet))
    return expanded


def _nothing_to_do(path: str, note: str) -> dict:
    return {
        "ok": True,
        "committed": False,
        "saved": False,
        "verified": False,
        "calculated": False,
        "calc": "skipped",
        "engine": None,
        "path": path,
        "completedOps": 0,
        "failedAt": None,
        "note": note,
    }


def _apply_openpyxl(path: str, ops: list[dict], tool: str, decision: RouteDecision, lock: FileLock):
    workbook = load_for_edit(path)
    applied = apply_ops(workbook, ops)
    if applied.error:
        workbook.close()
        return _response(
            path=path,
            engine="openpyxl",
            tool=tool,
            committed=False,
            calc="skipped",
            calc_required=decision.calc_required,
            wrote=applied.wrote,
            outcome=None,
            completed=applied.completed,
            failed_at=applied.failed_at,
            error={"code": "op_failed", "message": applied.error},
        )
    temp_path = None
    try:
        temp_path = save_atomic(workbook, path)
    except Exception as exc:
        return _response(
            path=path,
            engine="openpyxl",
            tool=tool,
            committed=False,
            calc="failed" if decision.calc_required else "skipped",
            calc_required=decision.calc_required,
            wrote=applied.wrote,
            outcome=None,
            completed=applied.completed,
            failed_at=applied.failed_at,
            error={"code": "save_failed", "message": str(exc)},
        )
    finally:
        workbook.close()

    outcome = VerifyOutcome(calc="skipped", calculated=False)
    committed = False
    try:
        if decision.calc_required:
            with COM_QUEUE:
                session = ComSession(temp_path, allow_macros=False)
                try:
                    session.open()
                    outcome = _run_verify(
                        session,
                        temp_path,
                        calc_required=True,
                        wrote=applied.wrote,
                        formula_cells=applied.formula_cells,
                        touched=applied.touched_sheets,
                        renamed={},
                    )
                    if outcome.calc != "done":
                        freed = session.abort()
                        if not freed:
                            lock.retain_until_file_free()
                        discard(temp_path)
                        return _response(
                            path=path,
                            engine="openpyxl",
                            tool=tool,
                            committed=False,
                            calc="failed",
                            calc_required=True,
                            wrote=applied.wrote,
                            outcome=outcome,
                            completed=applied.completed,
                            failed_at=applied.failed_at,
                            error={"code": "calc_failed", "message": outcome.error or "calculation failed"},
                        )
                    session.close(save=True)
                except Exception as exc:
                    freed = session.abort()
                    if not freed:
                        lock.retain_until_file_free()
                    discard(temp_path)
                    return _response(
                        path=path,
                        engine="openpyxl",
                        tool=tool,
                        committed=False,
                        calc="failed",
                        calc_required=True,
                        wrote=applied.wrote,
                        outcome=None,
                        completed=applied.completed,
                        failed_at=applied.failed_at,
                        error={"code": "excel_busy" if not session.pid_known else "calc_failed", "message": str(exc)},
                    )
            replace_target(temp_path, path)
            committed = True
            temp_path = None
        else:
            replace_target(temp_path, path)
            committed = True
            temp_path = None
            with COM_QUEUE:
                session = ComSession(path, allow_macros=False)
                try:
                    session.open()
                    outcome = _run_verify(
                        session,
                        path,
                        calc_required=False,
                        wrote=applied.wrote,
                        formula_cells=[],
                        touched=applied.touched_sheets,
                        renamed={},
                    )
                    session.close(save=False)
                except Exception as exc:
                    freed = session.abort()
                    if not freed:
                        lock.retain_until_file_free()
                    outcome = VerifyOutcome(calc="skipped", calculated=False, error=str(exc))
    except Exception as exc:
        discard(temp_path)
        return _response(
            path=path,
            engine="openpyxl",
            tool=tool,
            committed=committed,
            calc=outcome.calc,
            calc_required=decision.calc_required,
            wrote=applied.wrote,
            outcome=outcome,
            completed=applied.completed,
            failed_at=applied.failed_at,
            error={"code": "save_failed", "message": str(exc), "tempPath": temp_path},
        )
    error = None
    if outcome.error:
        error = {"code": "verify_failed", "message": outcome.error}
    return _response(
        path=path,
        engine="openpyxl",
        tool=tool,
        committed=committed,
        calc=outcome.calc,
        calc_required=decision.calc_required,
        wrote=applied.wrote,
        outcome=outcome,
        completed=applied.completed,
        failed_at=None,
        error=error,
    )


def _apply_com(path: str, ops: list[dict], tool: str, decision: RouteDecision, lock: FileLock):
    if ops and ops[0]["action"] == "verify" and len(ops) == 1:
        return _verify_only(path, ops[0], tool, lock)
    wrote, formulas, touched, renamed = _track(ops)
    commands = []
    op_index = []
    for index, op in enumerate(ops):
        if op["action"] == "create_workbook":
            continue
        command = COM_COMMANDS.get(op["action"])
        if not command:
            return _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc="skipped",
                calc_required=decision.calc_required,
                wrote=[],
                outcome=None,
                completed=0,
                failed_at=index,
                error={"code": "unsupported", "message": op["action"]},
            )
        commands.append({"command": command, "args": com_args(op)})
        op_index.append(index)
    query_only = decision.reason == "query"
    slow = {"table_add_to_data_model", "table_create_from_dax"}
    timeout = 600 if any(op["action"].startswith(("powerquery", "datamodel")) or op["action"] in slow for op in ops) else None
    with COM_QUEUE:
        session = ComSession(path, allow_macros=decision.allow_macros, timeout=timeout)
        try:
            session.open()
        except ExcelCliError as exc:
            freed = session.abort()
            if not freed:
                lock.retain_until_file_free()
            code = "excel_busy" if not exc.pid_known else "op_failed"
            return _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc="skipped",
                calc_required=decision.calc_required,
                wrote=[],
                outcome=None,
                completed=0,
                failed_at=0,
                error={"code": code, "message": str(exc)},
            )
        try:
            results = session.call(commands) if commands else []
        except ExcelCliError as exc:
            freed = session.abort()
            if not freed:
                lock.retain_until_file_free()
            failed = exc.index if exc.index is not None and exc.index < len(op_index) else 0
            body, shots = _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc="skipped",
                calc_required=decision.calc_required,
                wrote=[] if query_only else wrote,
                outcome=None,
                completed=failed,
                failed_at=op_index[failed] if op_index else 0,
                error={"code": "op_failed", "message": str(exc)},
            )
            return _with_results(body, ops, op_index, exc.results), shots
        completed = len(commands)
        if query_only:
            try:
                session.close(save=False)
            except ExcelCliError:
                if not session.abort():
                    lock.retain_until_file_free()
            return _query_response(path, completed, _results_of(ops, op_index, results)), []
        outcome = _run_verify(
            session,
            path,
            calc_required=decision.calc_required,
            wrote=wrote,
            formula_cells=formulas,
            touched=touched,
            renamed=renamed,
        )
        if decision.calc_required and outcome.calc != "done":
            freed = session.abort()
            if not freed:
                lock.retain_until_file_free()
            return _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc="failed",
                calc_required=True,
                wrote=wrote,
                outcome=outcome,
                completed=completed,
                failed_at=None,
                error={"code": "calc_failed", "message": outcome.error or "calculation failed"},
            )
        try:
            session.close(save=True)
        except ExcelCliError as exc:
            freed = session.abort()
            if not freed:
                lock.retain_until_file_free()
            return _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc=outcome.calc,
                calc_required=decision.calc_required,
                wrote=wrote,
                outcome=outcome,
                completed=completed,
                failed_at=None,
                error={"code": "save_failed" if session.pid_known else "excel_busy", "message": str(exc)},
            )
    error = {"code": "verify_failed", "message": outcome.error} if outcome.error else None
    if decision.calc_required and outcome.calc != "done":
        error = error or {"code": "calc_failed", "message": outcome.error or "calculation failed"}
    body, shots = _response(
        path=path,
        engine="com",
        tool=tool,
        committed=True,
        calc=outcome.calc,
        calc_required=decision.calc_required,
        wrote=wrote,
        outcome=outcome,
        completed=completed,
        failed_at=None,
        error=error,
    )
    return _with_results(body, ops, op_index, results), shots


_RESULT_NOISE = {"success", "filePath", "sessionId"}


def _results_of(ops: list[dict], op_index: list[int], results: list[dict]) -> list[dict]:
    """Payloads of query ops, keyed back to their position in the caller's ops."""
    found = []
    for item in results:
        position = item.get("index")
        if not isinstance(position, int) or position >= len(op_index):
            continue
        index = op_index[position]
        action = ops[index]["action"]
        if action not in QUERY_ACTIONS or not item.get("success", True):
            continue
        payload = item.get("result")
        if isinstance(payload, str):
            try:
                payload = json.loads(payload)
            except json.JSONDecodeError:
                pass
        if isinstance(payload, dict):
            payload = {key: value for key, value in payload.items() if key not in _RESULT_NOISE}
        found.append({"op": index, "action": action, "result": payload})
    return found


def _with_results(body: dict, ops: list[dict], op_index: list[int], results: list[dict]) -> dict:
    found = _results_of(ops, op_index, results)
    if found:
        body["results"] = found
    return body


def _query_response(path: str, completed: int, results: list[dict]) -> dict:
    return {
        "ok": True,
        "committed": False,
        "saved": False,
        "verified": False,
        "calculated": False,
        "calc": "skipped",
        "engine": "com",
        "path": path,
        "completedOps": completed,
        "failedAt": None,
        "readOnly": True,
        "results": results,
    }


def _verify_only(path: str, op: dict, tool: str, lock: FileLock):
    explicit = op.get("ranges") or {}
    sheets = list(op.get("sheets") or explicit.keys())
    wrote = [{"sheet": sheet, "range": explicit.get(sheet, "A1:L40")} for sheet in sheets]
    with COM_QUEUE:
        session = ComSession(path, allow_macros=False)
        try:
            session.open()
            outcome = _run_verify(
                session,
                path,
                calc_required=True,
                wrote=wrote,
                formula_cells=[],
                touched=sheets,
                renamed={},
                explicit_ranges=explicit or None,
            )
            if outcome.calc != "done":
                freed = session.abort()
                if not freed:
                    lock.retain_until_file_free()
                return _response(
                    path=path,
                    engine="com",
                    tool=tool,
                    committed=False,
                    calc="failed",
                    calc_required=True,
                    wrote=wrote,
                    outcome=outcome,
                    completed=0,
                    failed_at=0,
                    error={"code": "calc_failed", "message": outcome.error or "calculation failed"},
                )
            session.close(save=True)
        except ExcelCliError as exc:
            freed = session.abort()
            if not freed:
                lock.retain_until_file_free()
            return _response(
                path=path,
                engine="com",
                tool=tool,
                committed=False,
                calc="failed",
                calc_required=True,
                wrote=wrote,
                outcome=None,
                completed=0,
                failed_at=0,
                error={"code": "verify_failed", "message": str(exc)},
            )
    error = {"code": "verify_failed", "message": outcome.error} if outcome.error else None
    return _response(
        path=path,
        engine="com",
        tool=tool,
        committed=True,
        calc=outcome.calc,
        calc_required=True,
        wrote=wrote,
        outcome=outcome,
        completed=1,
        failed_at=None,
        error=error,
    )


def _compact(body: dict) -> str:
    # The MCP fallback serializer indents by 2, which puts every null of a dense page on its own line.
    return json.dumps(body, ensure_ascii=False, separators=(",", ":"), default=str)


def _with_images(body: dict, shots: list) -> list:
    content: list[Any] = [_compact(body)]
    for shot in shots:
        if not getattr(shot, "data", None):
            continue
        fmt = "png" if "png" in shot.mime else "jpeg"
        content.append(Image(data=shot.data, format=fmt))
    return content


@mcp.tool(structured_output=False)
def workbook_read(
    path: str,
    sheet: str | None = None,
    range: str | None = None,
    includeStyles: bool = False,
    limit: int = READ_LIMIT,
    mode: str = "sparse",
    preview: int = 0,
    maxText: int | None = None,
    find: str | None = None,
) -> list:
    """Read values, formulas, and cached values. Takes the workbook lock. Does not screenshot.

    mode=overview: every sheet (or `sheet`) with usedRange, cell/formula/merged counts, freeze, state.
      preview=N adds the first N non-empty rows per sheet (text clipped to 40 unless maxText).
      Start here on an unfamiliar workbook.
    mode=sparse (default): only non-empty cells, grouped by row: rows {"5": {"A": v, "C": {"f": "=..", "v": cached}}}.
      Plain cells are bare values; formula cells are {f, v}; includeStyles adds "s". Also lists merged ranges.
      limit counts non-empty cells; pages end on a row boundary.
    mode=dense: row-major 2D values (+formulas only if the page has any, +styles), the shape set_values takes.
      limit counts every cell in the rectangle.
    mode=find: cells whose value or formula text contains `find` (case-insensitive), across every sheet
      or just `sheet`, optionally inside `range`. limit caps matches; truncated says there were more.
    Without range the page runs from A1 to the last non-empty cell. maxText clips long strings.
    Follow nextRange until it is null. Page cap is 4000 cells.
    A workbook open in Excel is read from its last saved state and the page says openInExcel: true.
    """
    return [_compact(read_workbook(path, sheet, range, includeStyles, limit, mode, preview, maxText, find))]


def read_workbook(
    path: str,
    sheet: str | None = None,
    range: str | None = None,
    includeStyles: bool = False,
    limit: int = READ_LIMIT,
    mode: str = "sparse",
    preview: int = 0,
    maxText: int | None = None,
    find: str | None = None,
) -> dict:
    full = canonical_path(path)
    bad = _extension(full)
    if bad:
        return _error("unsupported_format", f"only .xlsx and .xlsm are accepted, got {bad}", path=full)
    if not os.path.exists(full):
        return _error("not_found", "workbook does not exist", path=full)
    if is_encrypted_or_irm(full):
        return _error("encrypted_or_irm", "encrypted or IRM workbooks are rejected", path=full)
    try:
        with FileLock(full):
            open_elsewhere = file_is_locked(full)
            try:
                page = read_page(
                    full,
                    sheet,
                    range,
                    includeStyles,
                    max(1, min(limit, READ_LIMIT)),
                    mode=mode,
                    preview=max(0, preview),
                    max_text=maxText,
                    find=find,
                )
            except PermissionError:
                return _error("file_locked", "workbook is open in another process that blocks reading", path=full)
    except LockTimeout:
        return _error("lock_timeout", "timed out waiting for the workbook lock", path=full)
    except Exception as exc:
        return _error("op_failed", str(exc), path=full)
    page["ok"] = True
    page["path"] = full
    if open_elsewhere:
        # Unsaved edits in Excel are not visible here.
        page["openInExcel"] = True
    return page


def _apply(path: str, ops: list[dict[str, Any]], tool: str) -> list:
    body, shots = apply_workbook(path, ops, tool)
    return _with_images(body, shots)


@mcp.tool(structured_output=False)
def workbook_apply(path: str, ops: list[dict[str, Any]]) -> list:
    """Write values, formulas, names, formatting, sheet structure, and layout. Saves once and screenshots.

    Actions: set_values, set_formulas, clear_contents, clear_all, format, merge, unmerge,
    set_row_height, set_column_width, define_name, define_name_update, create_workbook,
    create_sheet, rename_sheet, copy_sheet, move_sheet, delete_sheet, insert_rows, delete_rows,
    insert_columns, delete_columns, insert_cells, delete_cells, trim_sheet, layout, verify.
    One file per call. Tables, pivots, charts, queries, the data model, and VBA have their own tools;
    those tools also accept the cell, format, and sheet actions above, so seed data and build on it in one call.
    English formulas. Excel number formats. Close the file in Excel first. Do not repeat a committed insert.
    layout starts a sheet. profile is finance, analytics, or general. A fact block becomes an Excel table.
    trim_sheet {sheet}: deletes leftover rows and columns past the last non-empty cell, while keeping
    anything formulas, names, validations, or conditional formats still reference. If nothing is past
    that, the call does not save and note says why. overview shows leftover formatting as extent.
    """
    return _apply(path, ops, "workbook_apply")


@mcp.tool(structured_output=False)
def excel_table(path: str, ops: list[dict[str, Any]]) -> list:
    """Excel tables, pivot tables, charts, and slicers. Saves once and screenshots.

    Actions: table_list, table_read, table_create, table_append, table_resize, table_rename,
    table_delete, table_set_style, table_apply_filter, table_clear_filters, pivot_list,
    pivot_create_from_range, pivot_create_from_table, pivot_refresh, pivot_delete, chart_list,
    chart_create_from_range, chart_create_from_table, chart_move, chart_fit, chart_delete,
    slicer_list, slicer_create, slicer_delete, verify. Also the workbook_apply cell/format/sheet actions.
    Put command fields on the op. Do not paint a table body with format.
    *_list and table_read return data under results. A call made only of them does not save or screenshot.
    """
    return _apply(path, ops, "excel_table")


@mcp.tool(structured_output=False)
def excel_model(path: str, ops: list[dict[str, Any]]) -> list:
    """Power Query and the data model. Saves once and screenshots.

    Actions: powerquery_list, powerquery_view, powerquery_create, powerquery_update,
    powerquery_refresh, powerquery_refresh_all, powerquery_delete, powerquery_rename,
    datamodel_list_tables, datamodel_list_measures, datamodel_create_measure,
    datamodel_update_measure, datamodel_delete_measure, datamodel_evaluate, datamodel_refresh,
    table_add_to_data_model, table_create_from_dax, verify. Also the workbook_apply cell/format/sheet actions.
    *_list, powerquery_view, and datamodel_evaluate return data under results.
    A call made only of them does not save or screenshot.
    """
    return _apply(path, ops, "excel_model")


@mcp.tool(structured_output=False)
def excel_view(path: str, ops: list[dict[str, Any]]) -> list:
    """Conditional formats, validation, comments, hyperlinks, freeze panes, and sheet visibility.

    Actions: conditional_format_add, conditional_format_clear, conditional_format_list,
    validation_add, validation_get, validation_remove, comment_set, comment_get, comment_clear,
    threaded_comment_add, hyperlink_add, hyperlink_remove, freeze, unfreeze, sheet_hide,
    sheet_show, verify. Also the workbook_apply cell/format/sheet actions.
    conditional_format_list, validation_get, and comment_get return data under results.
    A call made only of them does not save or screenshot.
    """
    return _apply(path, ops, "excel_view")


@mcp.tool(structured_output=False)
def excel_vba(path: str, ops: list[dict[str, Any]]) -> list:
    """List, view, import, update, run, or delete VBA. Macros are enabled only on this tool.

    Actions: vba_list, vba_view, vba_import, vba_update, vba_run, vba_delete, verify.
    Also the workbook_apply cell/format/sheet actions.
    vba_list and vba_view return data under results and, alone, do not save or screenshot.
    Excel must trust access to the VBA project object model.
    """
    return _apply(path, ops, "excel_vba")


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
