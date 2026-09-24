"""Recalculation, spill extent, and screenshots at the end of a write."""

from __future__ import annotations

import base64
from dataclasses import dataclass, field

from excel_sovereign.excel_cli import ComSession
from excel_sovereign.inspect import bounding_box, dependent_ranges, range_size, top_left_window

IMAGE_BUDGET = int(1.5 * 1024 * 1024)
DEPENDENT_ROWS = 40
DEPENDENT_COLS = 16


@dataclass
class Shot:
    sheet: str
    range: str
    cropped: bool = False
    spill_unknown: bool = False
    mime: str = "image/jpeg"
    data: bytes = b""


@dataclass
class VerifyOutcome:
    calc: str
    calculated: bool
    shots: list[Shot] = field(default_factory=list)
    omitted: list[dict] = field(default_factory=list)
    spill_unknown: bool = False
    error: str | None = None
    # Committed ops touched no cells or sheets (e.g. a connection-only query).
    nothing_to_show: bool = False


def _payload(item: dict) -> dict:
    result = item.get("result")
    if isinstance(result, str):
        import json

        try:
            result = json.loads(result)
        except json.JSONDecodeError:
            return {}
    return result if isinstance(result, dict) else {}


def calculate(session: ComSession) -> None:
    session.call([{"command": "calculation.calculate", "args": {"scope": "Workbook"}}])


def _spill_of(payload: dict) -> tuple[bool, str | None]:
    supported = bool(payload.get("supported", payload.get("Supported", False)))
    has_spill = bool(payload.get("hasSpill", payload.get("HasSpill", False)))
    address = payload.get("address") or payload.get("Address")
    if not supported:
        return False, None
    if has_spill and address:
        return True, str(address).replace("$", "")
    return True, None


def spill_address(session: ComSession, sheet: str, cell: str) -> tuple[bool, str | None]:
    """Returns (supported, address). Address is set only when the cell spills."""
    return spill_addresses(session, [(sheet, cell)])[0]


def spill_addresses(session: ComSession, cells: list[tuple[str, str]]) -> list[tuple[bool, str | None]]:
    """One excelcli run for every probe. Same order as `cells`."""
    if not cells:
        return []
    results = session.call(
        [
            {"command": "range.get-spill", "args": {"sheetName": sheet, "rangeAddress": cell}}
            for sheet, cell in cells
        ]
    )
    found = [_spill_of(_payload(item)) for item in results[: len(cells)]]
    found.extend([(False, None)] * (len(cells) - len(found)))
    return found


def _image_of(item: dict) -> tuple[bytes, str, str]:
    payload = _payload(item)
    encoded = payload.get("imageBase64") or payload.get("ImageBase64") or ""
    mime = payload.get("mimeType") or payload.get("MimeType") or "image/jpeg"
    message = str(payload.get("message") or payload.get("Message") or "")
    if not encoded:
        raise RuntimeError(message or "screenshot returned no image")
    return base64.b64decode(encoded), mime, message


def _capture_command(sheet: str, address: str, quality: str) -> dict:
    return {
        "command": "screenshot.capture",
        "args": {"sheetName": sheet, "rangeAddress": address, "quality": quality},
    }


def _capture_once(session: ComSession, sheet: str, address: str, quality: str) -> tuple[bytes, str, str]:
    results = session.call([_capture_command(sheet, address, quality)])
    return _image_of(results[-1] if results else {})


def _shrink(session: ComSession, sheet: str, address: str, data: bytes, mime: str, message: str) -> Shot:
    cropped = "truncat" in message.lower()
    used = address
    if len(data) > IMAGE_BUDGET:
        data, mime, message = _capture_once(session, sheet, address, "Low")
        cropped = cropped or "truncat" in message.lower()
    if len(data) > IMAGE_BUDGET:
        used = top_left_window(address)
        data, mime, message = _capture_once(session, sheet, used, "Low")
        cropped = True
    return Shot(sheet=sheet, range=used, cropped=cropped, mime=mime, data=data)


def capture_range(session: ComSession, sheet: str, address: str) -> Shot:
    return capture_ranges(session, [(sheet, address)])[0]


def capture_ranges(session: ComSession, targets: list[tuple[str, str]]) -> list[Shot]:
    """Medium-quality captures in one run. Only oversize images are retaken one by one."""
    if not targets:
        return []
    results = session.call([_capture_command(sheet, address, "Medium") for sheet, address in targets])
    shots = []
    for (sheet, address), item in zip(targets, results):
        data, mime, message = _image_of(item)
        shots.append(_shrink(session, sheet, address, data, mime, message))
    if len(shots) != len(targets):
        raise RuntimeError("screenshot batch returned fewer images than requested")
    return shots


def union_ranges(ranges: list[str]) -> str | None:
    return bounding_box([item for item in ranges if item])


def _object_anchors(session: ComSession) -> dict[str, list[str]]:
    """Chart and pivot anchors currently visible in the session."""
    found: dict[str, list[str]] = {}

    def walk(node, sheet: str | None = None) -> None:
        if isinstance(node, dict):
            current = node.get("sheetName") or node.get("SheetName") or sheet
            top = node.get("topLeftCell") or node.get("TopLeftCell")
            bottom = node.get("bottomRightCell") or node.get("BottomRightCell")
            occupied = node.get("range") or node.get("Range")
            if current and top and bottom:
                found.setdefault(str(current), []).append(f"{str(top).replace('$', '')}:{str(bottom).replace('$', '')}")
            elif current and isinstance(occupied, str) and ":" in occupied and "!" not in occupied:
                found.setdefault(str(current), []).append(occupied.replace("$", ""))
            for value in node.values():
                walk(value, str(current) if current else sheet)
        elif isinstance(node, list):
            for value in node:
                walk(value, sheet)

    commands = [{"command": name, "args": {}} for name in ("chart.list", "pivottable.list")]
    try:
        batches = [session.call(commands)]
    except Exception:
        batches = []
        for command in commands:
            try:
                batches.append(session.call([command]))
            except Exception:
                continue
    for results in batches:
        for item in results:
            walk(_payload(item))
    return found


def plan_sheets(
    path: str,
    wrote: list[dict],
    formula_cells: list[dict],
    touched: list[str],
    renamed_from: dict[str, str],
) -> tuple[dict[str, str], list[dict], bool]:
    """Build per-sheet capture ranges.

    Returns ranges, omitted dependents, and whether any spill probe is required.
    """
    by_sheet: dict[str, list[str]] = {}
    for item in wrote:
        by_sheet.setdefault(item["sheet"], []).append(item["range"])
    for sheet in touched:
        by_sheet.setdefault(sheet, [])
    names = list(dict.fromkeys([*touched, *renamed_from.values(), *renamed_from.keys()]))
    omitted: list[dict] = []
    try:
        dependents = dependent_ranges(path, [name for name in names if name])
    except Exception:
        dependents = {}
    for sheet, cells in dependents.items():
        box = bounding_box(cells)
        if not box:
            continue
        rows, cols, _ = range_size(box)
        if rows > DEPENDENT_ROWS or cols > DEPENDENT_COLS:
            omitted.append({"sheet": sheet, "reason": "dependents exceed capture budget"})
            continue
        by_sheet.setdefault(sheet, []).append(box)
    ranges = {}
    for sheet, pieces in by_sheet.items():
        if pieces:
            box = union_ranges(pieces)
            if box:
                ranges[sheet] = box
    return ranges, omitted, bool(formula_cells)
