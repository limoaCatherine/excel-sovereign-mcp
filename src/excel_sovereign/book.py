"""openpyxl writes. One load, one save, keep_vba. This module is the only openpyxl caller."""

from __future__ import annotations

import os
import re
import shutil
import tempfile
import threading
from collections import OrderedDict
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

from excel_sovereign.lock import canonical_path


@dataclass
class BookResult:
    wrote: list[dict] = field(default_factory=list)
    formula_cells: list[dict] = field(default_factory=list)
    completed: int = 0
    failed_at: int | None = None
    error: str | None = None
    touched_sheets: list[str] = field(default_factory=list)


def load_for_edit(path: str):
    return load_workbook(path, keep_vba=path.lower().endswith(".xlsm"))


def create_empty(path: str) -> None:
    workbook = Workbook()
    workbook.active.title = "Sheet1"
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    workbook.save(path)
    workbook.close()


def save_atomic(workbook, path: str) -> str:
    """Save to a sibling temp file. Caller replaces the target after verification prep."""
    directory = os.path.dirname(path) or "."
    suffix = ".xlsm" if path.lower().endswith(".xlsm") else ".xlsx"
    fd, temp_path = tempfile.mkstemp(prefix=".excel-sovereign-mcp-", suffix=suffix, dir=directory)
    os.close(fd)
    try:
        workbook.save(temp_path)
    except Exception:
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise
    return temp_path


def replace_target(temp_path: str, path: str) -> None:
    os.replace(temp_path, path)


def apply_ops(workbook, ops: list[dict]) -> BookResult:
    result = BookResult()
    for index, op in enumerate(ops):
        try:
            spans = _apply_one(workbook, op)
        except Exception as exc:  # operation error before save
            result.failed_at = index
            result.error = str(exc)
            return result
        result.completed = index + 1
        for sheet, address, formulas in spans:
            if sheet not in result.touched_sheets:
                result.touched_sheets.append(sheet)
            if address:
                result.wrote.append({"sheet": sheet, "range": address})
            for cell, formula in formulas:
                result.formula_cells.append({"sheet": sheet, "cell": cell, "formula": formula})
    return result


def _sheet(workbook, name: str) -> Worksheet:
    if name not in workbook.sheetnames:
        raise KeyError(f"sheet not found: {name}")
    return workbook[name]


def _apply_one(workbook, op: dict) -> list[tuple[str, str | None, list[tuple[str, str]]]]:
    action = op["action"]
    if action == "create_workbook":
        return []
    if action == "create_sheet":
        name = str(op.get("name") or op.get("sheet"))
        if name in workbook.sheetnames:
            raise ValueError(f"sheet already exists: {name}")
        workbook.create_sheet(name)
        return [(name, "A1", [])]
    if action == "set_values":
        return [_write_block(workbook, op, formulas=False)]
    if action == "set_formulas":
        return [_write_block(workbook, op, formulas=True)]
    if action == "define_name":
        _define_name(workbook, op, replace=False)
        return []
    if action == "define_name_update":
        _define_name(workbook, op, replace=True)
        return []
    if action == "format":
        _format(workbook, op)
        return [(str(op["sheet"]), str(op["range"]), [])]
    if action == "merge":
        _sheet(workbook, op["sheet"]).merge_cells(str(op["range"]))
        return [(str(op["sheet"]), str(op["range"]), [])]
    if action == "unmerge":
        _sheet(workbook, op["sheet"]).unmerge_cells(str(op["range"]))
        return [(str(op["sheet"]), str(op["range"]), [])]
    if action == "set_row_height":
        ws = _sheet(workbook, op["sheet"])
        row = int(op["row"])
        count = int(op.get("count") or 1)
        height = float(op["height"])
        for offset in range(count):
            ws.row_dimensions[row + offset].height = height
        return [(str(op["sheet"]), f"A{row}:A{row + count - 1}", [])]
    if action == "set_column_width":
        ws = _sheet(workbook, op["sheet"])
        column = str(op["column"])
        ws.column_dimensions[column].width = float(op["width"])
        return [(str(op["sheet"]), f"{column}1", [])]
    if action == "clear_contents":
        _clear(workbook, op, formats=False)
        return [(str(op["sheet"]), str(op["range"]), [])]
    if action == "clear_all":
        _clear(workbook, op, formats=True)
        return [(str(op["sheet"]), str(op["range"]), [])]
    raise ValueError(f"openpyxl cannot apply {action}")


def _write_block(workbook, op: dict, formulas: bool):
    ws = _sheet(workbook, op["sheet"])
    rows = op["formulas"] if formulas else op["values"]
    if not isinstance(rows, list) or not rows:
        raise ValueError("values or formulas must be a non-empty 2D array")
    min_col, min_row, _, _ = range_boundaries(str(op["range"]))
    width = max(len(row) for row in rows)
    _reject_merged_write(ws, min_row, min_col, min_row + len(rows) - 1, min_col + width - 1)
    found: list[tuple[str, str]] = []
    for r_index, row in enumerate(rows):
        if not isinstance(row, list):
            raise ValueError("each row must be a list")
        for c_index in range(width):
            value = row[c_index] if c_index < len(row) else None
            cell = ws.cell(min_row + r_index, min_col + c_index, value)
            if isinstance(value, str) and value.startswith("="):
                found.append((cell.coordinate, value))
    end = ws.cell(min_row + len(rows) - 1, min_col + width - 1).coordinate
    start = ws.cell(min_row, min_col).coordinate
    address = start if start == end else f"{start}:{end}"
    return (str(op["sheet"]), address, found)


def _reject_merged_write(ws: Worksheet, min_row: int, min_col: int, max_row: int, max_col: int) -> None:
    for merged in ws.merged_cells.ranges:
        overlaps = not (
            merged.max_row < min_row
            or merged.min_row > max_row
            or merged.max_col < min_col
            or merged.min_col > max_col
        )
        if not overlaps:
            continue
        only_top_left = (
            min_row == max_row == merged.min_row and min_col == max_col == merged.min_col
        )
        if not only_top_left:
            raise ValueError(f"write intersects merged range {merged.coord}")


def _define_name(workbook, op: dict, replace: bool) -> None:
    name = str(op["name"])
    formula = str(op.get("formula") or op.get("reference") or "")
    if formula.startswith("="):
        formula = formula[1:]
    existing = name in workbook.defined_names
    if existing and not replace:
        raise ValueError(f"name already exists: {name}")
    if existing:
        del workbook.defined_names[name]
    workbook.defined_names.add(DefinedName(name=name, attr_text=formula))


def _format(workbook, op: dict) -> None:
    ws = _sheet(workbook, op["sheet"])
    min_col, min_row, max_col, max_row = range_boundaries(str(op["range"]))
    font_kwargs = {}
    if "bold" in op:
        font_kwargs["bold"] = bool(op["bold"])
    if "italic" in op:
        font_kwargs["italic"] = bool(op["italic"])
    if "underline" in op:
        font_kwargs["underline"] = "single" if op["underline"] else None
    if op.get("fontSize") is not None:
        font_kwargs["size"] = op["fontSize"]
    if op.get("fontName"):
        font_kwargs["name"] = op["fontName"]
    if op.get("fontColor"):
        font_kwargs["color"] = _hex(op["fontColor"])
    fill = None
    if op.get("fillColor"):
        fill = PatternFill("solid", fgColor=_hex(op["fillColor"]))
    border = None
    if op.get("borderStyle"):
        side = Side(style=_border_style(op["borderStyle"]), color=_hex(op.get("borderColor") or "000000"))
        if op.get("borderEdges") == "top":
            border = Border(top=side)
        else:
            border = Border(left=side, right=side, top=side, bottom=side)
    alignment_kwargs = {}
    if op.get("horizontalAlignment"):
        alignment_kwargs["horizontal"] = op["horizontalAlignment"]
    if op.get("verticalAlignment"):
        alignment_kwargs["vertical"] = "center" if op["verticalAlignment"] == "middle" else op["verticalAlignment"]
    if "wrapText" in op:
        alignment_kwargs["wrap_text"] = bool(op["wrapText"])
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if font_kwargs:
                cell.font = Font(
                    name=font_kwargs.get("name", cell.font.name),
                    size=font_kwargs.get("size", cell.font.size),
                    bold=font_kwargs.get("bold", cell.font.bold),
                    italic=font_kwargs.get("italic", cell.font.italic),
                    underline=font_kwargs.get("underline", cell.font.underline),
                    color=font_kwargs.get("color", cell.font.color),
                )
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment_kwargs:
                cell.alignment = Alignment(
                    horizontal=alignment_kwargs.get("horizontal", cell.alignment.horizontal),
                    vertical=alignment_kwargs.get("vertical", cell.alignment.vertical),
                    wrap_text=alignment_kwargs.get("wrap_text", cell.alignment.wrap_text),
                )
            if op.get("numberFormat"):
                cell.number_format = op["numberFormat"]


def _clear(workbook, op: dict, formats: bool) -> None:
    ws = _sheet(workbook, op["sheet"])
    min_col, min_row, max_col, max_row = range_boundaries(str(op["range"]))
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None
            if formats:
                cell.font = Font()
                cell.fill = PatternFill()
                cell.border = Border()
                cell.alignment = Alignment()
                cell.number_format = "General"


def _hex(value: str) -> str:
    text = str(value).strip().lstrip("#")
    if len(text) == 6:
        return "FF" + text.upper()
    return text.upper()


def _border_style(name: str) -> str:
    mapping = {
        "continuous": "thin",
        "thin": "thin",
        "medium": "medium",
        "thick": "thick",
        "double": "double",
        "dash": "dashed",
        "dot": "dotted",
        "hairline": "hair",
    }
    return mapping.get(str(name), str(name))


READ_MODES = ("sparse", "dense", "overview", "find")
READ_PAGE_MAX = 4000
MAX_COLUMN = 16384
_CACHE_SIZE = 2
_cache: OrderedDict[str, "_ReadBook"] = OrderedDict()
_cache_guard = threading.Lock()


class _ReadBook:
    """Loaded once per file version. Cached values load only when a page holds a formula."""

    def __init__(self, path: str, stamp: tuple[int, int]):
        self.path = path
        self.stamp = stamp
        self.book = load_workbook(path, data_only=False, read_only=False)
        self._cached = None
        self._guard = threading.Lock()

    def cached_sheet(self, name: str):
        with self._guard:
            if self._cached is None:
                self._cached = load_workbook(self.path, data_only=True, read_only=False)
        return self._cached[name]


def _read_book(path: str) -> _ReadBook:
    stat = os.stat(path)
    stamp = (stat.st_mtime_ns, stat.st_size)
    with _cache_guard:
        hit = _cache.get(path)
        if hit is not None and hit.stamp == stamp:
            _cache.move_to_end(path)
            return hit
    fresh = _ReadBook(path, stamp)
    with _cache_guard:
        _cache[path] = fresh
        _cache.move_to_end(path)
        while len(_cache) > _CACHE_SIZE:
            _cache.popitem(last=False)
    return fresh


def _formula_text(raw) -> str | None:
    if isinstance(raw, ArrayFormula):
        return raw.text
    if isinstance(raw, str) and raw.startswith("="):
        return raw
    return None


def _present(cell) -> bool:
    return cell.value is not None and cell.value != ""


def _stored(ws):
    # ws.cell() creates missing cells, which would grow a cached sheet on every read.
    return ws._cells


def _used_bounds(ws) -> tuple[int, int, int, int] | None:
    min_row = min_col = max_row = max_col = None
    for (row, col), cell in _stored(ws).items():
        if not _present(cell):
            continue
        min_row = row if min_row is None else min(min_row, row)
        max_row = row if max_row is None else max(max_row, row)
        min_col = col if min_col is None else min(min_col, col)
        max_col = col if max_col is None else max(max_col, col)
    if min_row is None:
        return None
    return min_col, min_row, max_col, max_row


def _address(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return start if start == end else f"{start}:{end}"


def _clip(value, max_text: int | None):
    if max_text and isinstance(value, str) and len(value) > max_text:
        return value[:max_text] + "…"
    return value


def _page_bounds(ws, address: str | None) -> tuple[int, int, int, int] | None:
    if address:
        return range_boundaries(address)
    used = _used_bounds(ws)
    if used is None:
        return None
    return 1, 1, used[2], used[3]


def _merged_within(ws, bounds: tuple[int, int, int, int]) -> list[str]:
    min_col, min_row, max_col, max_row = bounds
    hits = []
    for merged in ws.merged_cells.ranges:
        if merged.max_row < min_row or merged.min_row > max_row:
            continue
        if merged.max_col < min_col or merged.min_col > max_col:
            continue
        hits.append(merged.coord)
    return hits


def _sparse_rows(
    loaded: _ReadBook,
    ws,
    bounds: tuple[int, int, int, int],
    include_styles: bool,
    limit: int,
    max_text: int | None,
    max_rows: int | None = None,
) -> tuple[dict, str | None]:
    min_col, min_row, max_col, max_row = bounds
    by_row: dict[int, list] = {}
    for (row, col), cell in _stored(ws).items():
        if min_row <= row <= max_row and min_col <= col <= max_col and _present(cell):
            by_row.setdefault(row, []).append((col, cell))
    rows: dict[str, dict] = {}
    consumed = 0
    next_range = None
    for row in sorted(by_row):
        if consumed >= limit or (max_rows is not None and len(rows) >= max_rows):
            next_range = _address(min_col, row, max_col, max_row)
            break
        out = {}
        for col, cell in sorted(by_row[row], key=lambda item: item[0]):
            formula = _formula_text(cell.value)
            if formula is None and not include_styles:
                out[get_column_letter(col)] = _clip(cell.value, max_text)
            else:
                entry: dict = {}
                if formula is not None:
                    entry["f"] = _clip(formula, max_text)
                    entry["v"] = _clip(loaded.cached_sheet(ws.title).cell(row, col).value, max_text)
                else:
                    entry["v"] = _clip(cell.value, max_text)
                if include_styles:
                    entry["s"] = _style_of(cell)
                out[get_column_letter(col)] = entry
            consumed += 1
        rows[str(row)] = out
    return rows, next_range


def _dense_page(
    loaded: _ReadBook,
    ws,
    bounds: tuple[int, int, int, int],
    include_styles: bool,
    limit: int,
    max_text: int | None,
) -> dict:
    min_col, min_row, max_col, max_row = bounds
    stored = _stored(ws)
    values, formulas, styles = [], [], []
    has_formula = False
    next_range = None
    consumed = 0
    stop_row, stop_col = min_row, min_col
    for row in range(min_row, max_row + 1):
        if consumed >= limit:
            next_range = _address(min_col, row, max_col, max_row)
            break
        value_row, formula_row, style_row = [], [], []
        for col in range(min_col, max_col + 1):
            if consumed >= limit:
                next_range = f"{get_column_letter(col)}{row}:{get_column_letter(max_col)}{max_row}"
                break
            cell = stored.get((row, col))
            raw = cell.value if cell is not None else None
            formula = _formula_text(raw)
            if formula is not None:
                has_formula = True
                raw = loaded.cached_sheet(ws.title).cell(row, col).value
            formula_row.append(formula)
            value_row.append(_clip(raw, max_text))
            if include_styles:
                style_row.append(_style_of(cell if cell is not None else Cell(ws, row=row, column=col)))
            consumed += 1
            stop_row, stop_col = row, col
        values.append(value_row)
        formulas.append(formula_row)
        if include_styles:
            styles.append(style_row)
        if next_range:
            break
    page = {
        "range": _address(min_col, min_row, stop_col, stop_row),
        "values": values,
    }
    if has_formula:
        page["formulas"] = formulas
    if include_styles:
        page["styles"] = styles
    page["nextRange"] = next_range
    return page


def _extent(ws) -> tuple[int, int]:
    """Last row and column that carry anything: a value, a style, or a row/column setting."""
    max_row = max_col = 0
    for row, col in _stored(ws):
        max_row = max(max_row, row)
        max_col = max(max_col, col)
    for key, dim in ws.column_dimensions.items():
        # A span to the last column is the sheet-wide column style. Excel rewrites it after any delete.
        if (dim.max or 0) >= MAX_COLUMN:
            continue
        max_col = max(max_col, dim.max or 0, column_index_from_string(key))
    for row, dim in ws.row_dimensions.items():
        if dim.ht is not None or dim.hidden or dim.outline_level or dim.customFormat:
            max_row = max(max_row, row)
    return max_row, max_col


_STRING_LITERAL = re.compile(r'"(?:[^"]|"")*"')
_REFERENCE = re.compile(
    r"(?<![A-Za-z0-9_.\u4e00-\u9fff])"
    r"(?:(?:'((?:[^']|'')+)'|([^\s!'(),=+\-*/&<>:;{}\"^%]+))!)?"
    r"(?:\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?"
    r"|\$?([A-Z]{1,3}):\$?([A-Z]{1,3})"
    r"|\$?(\d+):\$?(\d+))"
    r"(?![A-Za-z0-9_(\u4e00-\u9fff])"
)


def _reach(text: str, home: str | None, sheet: str) -> tuple[int, int]:
    """Furthest row and column of `sheet` that one formula or name refers to."""
    max_row = max_col = 0
    for match in _REFERENCE.finditer(_STRING_LITERAL.sub('""', text)):
        target = (match.group(1) or "").replace("''", "'") or match.group(2) or home
        if target != sheet:
            continue
        if match.group(3):
            max_col = max(max_col, column_index_from_string(match.group(3)))
            max_row = max(max_row, int(match.group(4)))
            if match.group(5):
                max_col = max(max_col, column_index_from_string(match.group(5)))
                max_row = max(max_row, int(match.group(6)))
        elif match.group(7):
            max_col = max(max_col, column_index_from_string(match.group(7)), column_index_from_string(match.group(8)))
        elif match.group(9):
            max_row = max(max_row, int(match.group(9)), int(match.group(10)))
    return max_row, max_col


def _referenced_extent(book, sheet: str) -> tuple[int, int]:
    """Area of `sheet` that formulas, names, conditional formats, or validations still point at.

    Deleting inside it would shrink those references, so trim keeps it.
    """
    max_row = max_col = 0

    def grow(found: tuple[int, int]) -> None:
        nonlocal max_row, max_col
        max_row = max(max_row, found[0])
        max_col = max(max_col, found[1])

    for ws in book.worksheets:
        for cell in _stored(ws).values():
            formula = _formula_text(cell.value)
            if formula and (ws.title == sheet or sheet in formula):
                grow(_reach(formula, ws.title, sheet))
        for defined in ws.defined_names.values():
            grow(_reach(defined.attr_text or "", ws.title, sheet))
    for defined in book.defined_names.values():
        grow(_reach(defined.attr_text or "", None, sheet))
    ws = book[sheet]
    areas = [str(rule.sqref) for rule in ws.conditional_formatting]
    areas += [str(rule.sqref) for rule in ws.data_validations.dataValidation]
    for area in areas:
        for piece in area.split():
            bounds = range_boundaries(piece)
            grow((bounds[3] or 0, bounds[2] or 0))
    return max_row, max_col


def trim_ops(path: str, sheet: str) -> list[dict]:
    """delete_columns / delete_rows past the last non-empty cell and past anything still referenced."""
    loaded = _read_book(canonical_path(path))
    if sheet not in loaded.book.sheetnames:
        raise KeyError(f"sheet not found: {sheet}")
    ws = loaded.book[sheet]
    used = _used_bounds(ws)
    used_row, used_col = (used[3], used[2]) if used else (0, 0)
    ref_row, ref_col = _referenced_extent(loaded.book, sheet)
    used_row, used_col = max(used_row, ref_row), max(used_col, ref_col)
    ext_row, ext_col = _extent(ws)
    ops = []
    if ext_col > used_col:
        ops.append({"action": "delete_columns", "sheet": sheet, "column": get_column_letter(used_col + 1), "count": ext_col - used_col})
    if ext_row > used_row:
        ops.append({"action": "delete_rows", "sheet": sheet, "row": used_row + 1, "count": ext_row - used_row})
    return ops


def _overview(loaded: _ReadBook, names: list[str], preview: int, max_text: int | None) -> list[dict]:
    sheets = []
    for name in names:
        ws = loaded.book[name]
        used = _used_bounds(ws)
        cells = [cell for cell in _stored(ws).values() if _present(cell)]
        info = {
            "name": name,
            "usedRange": _address(*used) if used else None,
            "cells": len(cells),
            "formulas": sum(1 for cell in cells if _formula_text(cell.value) is not None),
            "merged": len(ws.merged_cells.ranges),
        }
        ext_row, ext_col = _extent(ws)
        if used and (ext_row > used[3] or ext_col > used[2]):
            # Formatting past the data. trim_sheet removes it.
            info["extent"] = _address(1, 1, max(ext_col, 1), max(ext_row, 1))
        if ws.freeze_panes:
            info["freeze"] = ws.freeze_panes
        if ws.sheet_state != "visible":
            info["state"] = ws.sheet_state
        if preview and used:
            rows, _ = _sparse_rows(loaded, ws, used, False, READ_PAGE_MAX, max_text or 40, max_rows=preview)
            info["preview"] = rows
        sheets.append(info)
    return sheets


def _find(
    loaded: _ReadBook,
    names: list[str],
    address: str | None,
    needle: str,
    limit: int,
    max_text: int | None,
) -> dict:
    target = needle.casefold()
    matches = []
    truncated = False
    for name in names:
        ws = loaded.book[name]
        bounds = range_boundaries(address) if address else None
        for (row, col), cell in sorted(_stored(ws).items()):
            if not _present(cell):
                continue
            if bounds and not (bounds[1] <= row <= bounds[3] and bounds[0] <= col <= bounds[2]):
                continue
            formula = _formula_text(cell.value)
            text = formula if formula is not None else str(cell.value)
            if target not in text.casefold():
                continue
            if len(matches) >= limit:
                truncated = True
                break
            hit = {"sheet": name, "cell": f"{get_column_letter(col)}{row}"}
            if formula is not None:
                hit["f"] = _clip(formula, max_text)
            else:
                hit["v"] = _clip(cell.value, max_text)
            matches.append(hit)
        if truncated:
            break
    return {"mode": "find", "find": needle, "matches": matches, "truncated": truncated}


def read_page(
    path: str,
    sheet: str | None,
    address: str | None,
    include_styles: bool,
    limit: int,
    mode: str = "sparse",
    preview: int = 0,
    max_text: int | None = None,
    find: str | None = None,
) -> dict:
    if mode not in READ_MODES:
        raise ValueError(f"mode must be one of {', '.join(READ_MODES)}, got {mode}")
    path = canonical_path(path)
    loaded = _read_book(path)
    names = loaded.book.sheetnames
    if sheet is not None and sheet not in names:
        raise KeyError(f"sheet not found: {sheet}")
    if mode == "overview":
        return {"mode": mode, "sheets": _overview(loaded, [sheet] if sheet else names, preview, max_text)}
    if mode == "find":
        if not find:
            raise ValueError("mode=find needs find")
        return _find(loaded, [sheet] if sheet else names, address, find, limit, max_text or 80)
    sheet_name = sheet or names[0]
    ws = loaded.book[sheet_name]
    bounds = _page_bounds(ws, address)
    if bounds is None:
        return {"sheet": sheet_name, "mode": mode, "range": None, "empty": True, "nextRange": None, "sheets": names}
    if mode == "dense":
        page = _dense_page(loaded, ws, bounds, include_styles, limit, max_text)
        return {"sheet": sheet_name, "mode": mode, **page, "sheets": names}
    rows, next_range = _sparse_rows(loaded, ws, bounds, include_styles, limit, max_text)
    payload = {"sheet": sheet_name, "mode": mode, "range": _address(*bounds), "rows": rows}
    merged = _merged_within(ws, bounds)
    if merged:
        payload["merged"] = merged
    payload["nextRange"] = next_range
    payload["sheets"] = names
    return payload


def _style_of(cell) -> dict:
    font = cell.font
    fill = cell.fill
    alignment = cell.alignment
    return {
        "fontName": font.name,
        "fontSize": font.size,
        "bold": bool(font.bold),
        "italic": bool(font.italic),
        "underline": bool(font.underline),
        "fontColor": _color_out(font.color),
        "fillColor": _color_out(getattr(fill.fgColor, "rgb", None) and fill.fgColor or fill.fgColor),
        "numberFormat": cell.number_format,
        "horizontalAlignment": alignment.horizontal,
        "verticalAlignment": alignment.vertical,
        "wrapText": bool(alignment.wrap_text),
    }


def _color_out(color) -> str | dict | None:
    if color is None:
        return None
    color_type = getattr(color, "type", None)
    if color_type == "rgb" and color.rgb and color.rgb != "00000000":
        rgb = str(color.rgb)
        return "#" + rgb[-6:]
    if color_type == "theme":
        return {"theme": color.theme, "tint": color.tint}
    if color_type == "indexed":
        return {"indexed": color.indexed}
    return None


def discard(path: str | None) -> None:
    if path and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass


def file_hash(path: str) -> str:
    import hashlib

    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def copy_file(src: str, dest: str) -> None:
    shutil.copy2(src, dest)
