"""SafeXL read-only compact cell viewer.

Lets a small-context model read a workbook without dumping raw sheet XML:
``view()`` returns a bounded window of cells (value, type, formula text and
cached value), ``render_text()`` turns that into one terse line per cell.

Like ``safexl_mcp.inspect`` it walks the OOXML zip parts directly (zip + XML,
no spreadsheet library), never writes, and never evaluates a formula or runs a
macro: a formula cell shows its formula text and the value Excel cached.
Array formulas are marked, not expanded. A shared-formula child shows its own
formula text, the master's shifted by the row/column offset.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .engine import SafeXLError
from .inspect import NS_MAIN, _col_to_num, _num_to_col, _sheet_list

__all__ = ["ViewError", "view", "render_text"]

_CELL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_RANGE_RE = re.compile(r"^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$")
_STR_MAX = 200
# Quoted text (string literal or quoted sheet name) is skipped; otherwise a cell ref, a
# whole-column range (A:C) or a whole-row range (2:5), each not glued to a name or function.
_SHIFT_RE = re.compile(
    r"""("(?:[^"]|"")*"|'(?:[^']|'')*')"""
    r"""|(?<![A-Za-z0-9_.])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])"""
    r"""|(?<![A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])"""
    r"""|(?<![A-Za-z0-9_.$])(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_(])"""
)

_M = f"{{{NS_MAIN}}}"


class ViewError(SafeXLError, ValueError):
    """Unknown sheet or malformed range (also a SafeXLError for CLI/server callers)."""


def _parse_range(rng: str) -> tuple[int, int, int, int]:
    """-> (col1, row1, col2, row2), normalized so col1<=col2, row1<=row2."""
    m = _RANGE_RE.match(rng.strip())
    if not m:
        raise ViewError(f"bad range {rng!r}; use e.g. A1:D20 or B7")
    c1, r1 = _col_to_num(m.group(1).upper()), int(m.group(2))
    c2, r2 = (_col_to_num(m.group(3).upper()), int(m.group(4))) if m.group(3) else (c1, r1)
    if min(r1, r2) < 1:
        raise ViewError(f"bad range {rng!r}; rows start at 1")
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def _text_of(si: ET.Element) -> str:
    """Concatenated <t> text of a shared-string / inline-string item (skips phonetic runs)."""
    parts: list[str] = []
    for child in si:
        if child.tag == f"{_M}t":
            parts.append(child.text or "")
        elif child.tag == f"{_M}r":
            t = child.find(f"{_M}t")
            if t is not None:
                parts.append(t.text or "")
    return "".join(parts)


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [_text_of(si) for si in root.findall(f"{_M}si")]


def _number(text: str) -> int | float | str:
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return text


def _cell_from_el(c: ET.Element, sst: list[str]) -> dict[str, Any] | None:
    """One <c> -> {ref, col, row, type, value, formula?, ...}, or None if empty."""
    ref = c.get("r", "")
    m = _CELL_RE.match(ref)
    if not m:
        return None
    t = c.get("t", "n")
    v_el = c.find(f"{_M}v")
    f_el = c.find(f"{_M}f")
    raw = v_el.text if v_el is not None else None

    value: Any = None
    typ = "n"
    if t == "s":
        typ = "s"
        if raw is not None:
            try:
                value = sst[int(raw)]
            except (ValueError, IndexError):
                value = None
    elif t == "inlineStr":
        typ = "s"
        is_el = c.find(f"{_M}is")
        value = _text_of(is_el) if is_el is not None else None
    elif t == "str":
        typ, value = "s", raw
    elif t == "b":
        typ, value = "b", (raw == "1") if raw is not None else None
    elif t == "e":
        typ, value = "e", raw
    elif t == "d":
        typ, value = "d", raw
    else:
        value = _number(raw) if raw is not None else None

    out: dict[str, Any] = {
        "ref": f"{m.group(1).upper()}{m.group(2)}",
        "col": _col_to_num(m.group(1).upper()),
        "row": int(m.group(2)),
        "type": typ,
        "value": value,
    }
    if f_el is not None:
        ftype = f_el.get("t")
        text = (f_el.text or "").strip()
        if ftype == "shared":
            out["shared"] = f_el.get("si")
        elif ftype == "array":
            out["array"] = f_el.get("ref") or True
        if text:
            out["formula"] = "=" + text
        elif ftype not in ("shared", "array"):
            out["formula"] = "="
    elif value is None:
        return None  # styled-but-empty cell
    return out


def _shift_formula(text: str, d_col: int, d_row: int) -> str:
    """Translate a shared-formula master body by an offset: relative refs move, ``$`` refs stay."""

    def col(abs_: str, letters: str) -> str:
        n = _col_to_num(letters.upper())
        if not abs_:
            n += d_col
        return f"{abs_}{_num_to_col(n)}" if n >= 1 else "#REF!"

    def row(abs_: str, digits: str) -> str:
        n = int(digits)
        if not abs_:
            n += d_row
        return f"{abs_}{n}" if n >= 1 else "#REF!"

    def sub(m: re.Match) -> str:
        if m.group(1) is not None:
            return m.group(1)
        if m.group(3) is not None:
            a, b = col(m.group(2), m.group(3)), row(m.group(4), m.group(5))
            return "#REF!" if "#REF!" in (a, b) else a + b
        if m.group(7) is not None:
            a, b = col(m.group(6), m.group(7)), col(m.group(8), m.group(9))
            return "#REF!" if "#REF!" in (a, b) else f"{a}:{b}"
        a, b = row(m.group(10), m.group(11)), row(m.group(12), m.group(13))
        return "#REF!" if "#REF!" in (a, b) else f"{a}:{b}"

    return _SHIFT_RE.sub(sub, text)


def _iter_cells(zf: zipfile.ZipFile, target: str, sst: list[str]):
    masters: dict[str, tuple[str, int, int]] = {}  # shared-formula group id -> (body, col, row)
    with zf.open(target) as fh:
        for _ev, el in ET.iterparse(fh, events=("end",)):
            if el.tag == f"{_M}c":
                cell = _cell_from_el(el, sst)
                el.clear()
                if cell is not None:
                    si = cell.get("shared")
                    if si is not None:
                        if cell.get("formula", "=") != "=":
                            masters[si] = (cell["formula"][1:], cell["col"], cell["row"])
                        elif si in masters:
                            body, mc, mr = masters[si]
                            cell["formula"] = "=" + _shift_formula(body, cell["col"] - mc, cell["row"] - mr)
                            cell["shared_from"] = si
                    yield cell


def _matches(cell: dict[str, Any], needle: str) -> bool:
    v = cell["value"]
    if v is not None and needle in str(v).lower():
        return True
    f = cell.get("formula")
    return bool(f) and needle in f.lower()


def _public(cell: dict[str, Any], formulas: bool, sheet: str | None) -> dict[str, Any]:
    out: dict[str, Any] = {"ref": cell["ref"], "type": cell["type"], "value": cell["value"]}
    if sheet is not None:
        out["sheet"] = sheet
    if formulas:
        for k in ("formula", "shared", "array", "shared_from"):
            if k in cell:
                out[k] = cell[k]
    return out


def _used_range(min_c: int, min_r: int, max_c: int, max_r: int) -> str | None:
    if max_r == 0:
        return None
    return f"{_num_to_col(min_c)}{min_r}:{_num_to_col(max_c)}{max_r}"


def view(
    path: str | Path,
    sheet: str | None = None,
    cell_range: str | None = None,
    find: str | None = None,
    max_cells: int = 400,
    formulas: bool = True,
) -> dict[str, Any]:
    """Read-only compact view of ``path``.

    Returns ``{sheets, sheet, dims, cells, truncated, next_range?}``. ``cells``
    holds at most ``max_cells`` non-empty cells in row-major order, each
    ``{ref, type, value, formula?}`` (type: n number, s string, b bool, e error,
    d date). Formula cells carry the formula text and the cached value; a shared
    formula's master is marked (``shared`` = group id) and each child carries its own
    formula text (the master's, shifted by the offset; ``shared_from`` = group id);
    array formulas are marked (``array`` = ref). Nothing is evaluated. With ``find`` (case-insensitive substring over
    value or formula text) the matches come back, each with its ``sheet``; all
    sheets are searched when ``sheet`` is None. When ``truncated``, ``next_range``
    (non-find views) is the range to pass as ``cell_range`` for the next window.

    Raises ViewError (a ValueError and a SafeXLError) for an unknown sheet or a
    bad range; SafeXLError if the file is missing or not an OOXML workbook.
    """
    if max_cells < 1:
        raise ViewError("max_cells must be >= 1")
    path = Path(path)
    if not path.is_file():
        raise SafeXLError(f"file not found: {path}")
    rng = _parse_range(cell_range) if cell_range else None

    try:
        with zipfile.ZipFile(path, "r") as zf:
            try:
                wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            except KeyError as exc:
                raise SafeXLError(f"{path}: not a valid OOXML workbook (missing xl/workbook.xml)") from exc
            metas = _sheet_list(zf, wb_root)
            names = [m["name"] for m in metas]
            by_name = {m["name"]: m for m in metas}
            if sheet is not None and sheet not in by_name:
                raise ViewError(f"no sheet {sheet!r}; sheets: {', '.join(names)}")
            sst = _shared_strings(zf)
            members = set(zf.namelist())

            def scan(name: str):
                target = by_name[name]["target"]
                if not target or target not in members:
                    return iter(())
                return _iter_cells(zf, target, sst)

            if find is not None:
                return _do_find(scan, names, sheet, rng, find, max_cells, formulas)

            sel = sheet if sheet is not None else (names[0] if names else None)
            if sel is None:
                return {"sheets": names, "sheet": None, "dims": None, "cells": [], "truncated": False}
            return _do_window(scan(sel), names, sel, rng, max_cells, formulas)
    except zipfile.BadZipFile as exc:
        raise SafeXLError(f"{path}: not a valid OOXML workbook (not a zip)") from exc
    except ET.ParseError as exc:
        raise SafeXLError(f"{path}: malformed XML: {exc}") from exc


def _in_rng(cell: dict[str, Any], rng: tuple[int, int, int, int] | None) -> bool:
    if rng is None:
        return True
    c1, r1, c2, r2 = rng
    return c1 <= cell["col"] <= c2 and r1 <= cell["row"] <= r2


def _do_window(cells_iter, names, sel, rng, max_cells, formulas) -> dict[str, Any]:
    min_c = min_r = 10**9
    max_c = max_r = 0
    window: list[dict[str, Any]] = []
    truncated = False
    nxt: dict[str, Any] | None = None
    for cell in cells_iter:
        min_c, max_c = min(min_c, cell["col"]), max(max_c, cell["col"])
        min_r, max_r = min(min_r, cell["row"]), max(max_r, cell["row"])
        if not _in_rng(cell, rng):
            continue
        if len(window) < max_cells:
            window.append(cell)
        elif nxt is None:
            truncated, nxt = True, cell
    skip = None
    if nxt is not None and window and window[-1]["row"] == nxt["row"]:
        # The cut fell mid-row. A range cannot start mid-row, so end the window at the previous
        # row (next_range then repeats nothing); a single row wider than max_cells is cut and the
        # rest of that row is named in the hint, so paging carries on from the next row.
        if window[0]["row"] != nxt["row"]:
            while window and window[-1]["row"] == nxt["row"]:
                window.pop()
        else:
            skip = nxt
    out: dict[str, Any] = {
        "sheets": names,
        "sheet": sel,
        "dims": _used_range(min_c, min_r, max_c, max_r),
        "cells": [_public(c, formulas, None) for c in window],
        "truncated": truncated,
    }
    if truncated and nxt is not None:
        c1, _r1, c2, r2 = rng if rng else (1, 1, max_c, max_r)
        if skip is not None:
            if skip["row"] < r2:  # a range with r1 > r2 would be normalized back over this row
                out["next_range"] = f"{_num_to_col(c1)}{skip['row'] + 1}:{_num_to_col(c2)}{r2}"
            out["hint"] = (f"row {skip['row']} has more than max_cells cells; the rest starts at "
                           f"{skip['ref']} (view it with cell_range={skip['ref']}:{_num_to_col(c2)}{skip['row']})")
        else:
            out["next_range"] = f"{_num_to_col(c1)}{nxt['row']}:{_num_to_col(c2)}{r2}"
    return out


def _do_find(scan, names, sheet, rng, find, max_cells, formulas) -> dict[str, Any]:
    needle = find.lower()
    hits: list[dict[str, Any]] = []
    truncated = False
    for name in [sheet] if sheet is not None else names:
        for cell in scan(name):
            if not _in_rng(cell, rng) or not _matches(cell, needle):
                continue
            if len(hits) >= max_cells:
                truncated = True
                break
            hits.append(_public(cell, formulas, name))
        if truncated:
            break
    out: dict[str, Any] = {
        "sheets": names,
        "sheet": sheet,
        "dims": None,
        "cells": hits,
        "truncated": truncated,
    }
    if truncated:
        out["hint"] = "more matches exist; pass sheet= or cell_range= to narrow, or raise max_cells"
    return out


def _fmt_value(v: Any, typ: str) -> str:
    if v is None:
        return "-"
    if typ == "b":
        return "TRUE" if v else "FALSE"
    if typ == "s":
        s = v if len(v) <= _STR_MAX else v[:_STR_MAX] + "..."
        return json.dumps(s, ensure_ascii=False)
    return str(v)


def render_text(view_result: dict[str, Any]) -> str:
    """Terse text form: ``B12  n  1250.5  =SUM(B2:B11)``, one line per cell.

    Strings are JSON-quoted so numbers stored as text stay visible. Find hits
    are prefixed ``Sheet!``. Header and footer lines start with ``#``.
    """
    lines: list[str] = []
    sheet = view_result.get("sheet")
    if view_result.get("dims") is not None:
        lines.append(f"# {sheet}  used range {view_result['dims']}")
    for c in view_result["cells"]:
        ref = f"{c['sheet']}!{c['ref']}" if "sheet" in c else c["ref"]
        parts = [ref, c["type"], _fmt_value(c["value"], c["type"])]
        f = c.get("formula")
        if "shared_from" in c:
            parts.append(f)
        elif "shared" in c:
            tag = f"[shared {c['shared']}]"
            parts.append(f"{f} {tag}" if f else tag)
        elif "array" in c:
            parts.append(f"{f} [array]" if f else "[array]")
        elif f:
            parts.append(f)
        lines.append("  ".join(parts))
    if not view_result["cells"]:
        lines.append("# (no cells)")
    if view_result.get("truncated"):
        if view_result.get("next_range"):
            lines.append(f"# truncated: call again with cell_range={view_result['next_range']!r}")
            if view_result.get("hint"):
                lines.append(f"# {view_result['hint']}")
        else:
            lines.append(f"# truncated: {view_result.get('hint', 'narrow with sheet/cell_range')}")
    return "\n".join(lines)
