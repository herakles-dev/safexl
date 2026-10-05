"""SafeXL salvage: rescue a workbook another writer damaged by re-saving it.

A model that re-saves a workbook with openpyxl/pandas loses parts the writer
does not model (charts, pivots, VBA, extLst, ...). ``salvage`` recovers the
*intended* edits from that damaged copy and re-applies them to the ORIGINAL
with the SafeXL engine, so every part the edits do not touch stays
byte-identical to the original.

Intended edit vs re-save artefact
---------------------------------
Cells are compared by resolved semantic content (``safexl_mcp.diff``'s cell
representation), never raw XML. A cell is an *intended edit* only when the
damaged cell differs from the original (after structural steps, see below)
AND is not explained by the other writer's re-save. These differences are
artefacts and are ignored:

- an empty cell vs a missing cell (style-only cells are dropped/added);
- numbers equal within 1e-12 relative (``1`` vs ``1.0``, float re-encoding);
- shared strings vs inline strings (resolved to text before comparing);
- formulas equal after stripping a leading ``=``, whitespace and sheet-name
  quotes; a formula-cell's cached value is never compared at all;
- a formula in the damaged file where the original held a *shared-formula
  child* (openpyxl expands shared formulas; the child text cannot be verified
  without a translator, so it is treated as untouched);
- a cell whose damaged content equals the ORIGINAL content at the cell it
  maps back to (a formula another writer left stale after a row insert or
  sheet rename is not an edit).

Not applied, reported in ``skipped`` (never silently dropped): a cell that was
cleared (the engine cannot clear), a formula replaced by a plain value (the
signature of a ``data_only`` re-save, ambiguous with a real overwrite), and an
error value.

Do no harm (v2)
---------------
``salvage`` first runs ``verify(original, damaged)``. Only a ``damaged`` or
``invalid`` copy proceeds; ``intact`` and ``rewritten`` with no lost feature
return ``nothing_to_do`` and write nothing (``force=True`` overrides). After
writing, ``verify(original, output)`` must say ``intact`` or the output is
deleted and the result is ``unsupported``. Any engine refusal inside salvage
is ``unsupported`` with the reason, never an exception.

Table growth: a table whose ref grew in the damaged copy, where an applied
cell edit sits in the added area, is resized with ``resize_table`` (carried
in ``tables``); anything else is reported in ``skipped``.

Structural changes
------------------
Detected on every call and listed in ``structural.detected``:
sheet added / renamed / removed, rows inserted or deleted, columns shifted.
Handled (``handled: true``): a sheet rename (``rename_sheet``), an added sheet
(``add_sheet`` then cell edits), ONE uniform row-insert block per sheet
(``insert_rows`` then cell edits), and column-scoped insert blocks
(one per group of columns that shifted together: ``insert_cells`` over its
span, then cell edits). Everything else (removed sheet, deleted rows,
shifted columns, several uniform row-insert blocks in one sheet, ambiguous
shifts) returns ``unsupported`` and writes NO output file.

Never executes macros; never writes to either input.
"""
from __future__ import annotations

import math
import re
import shutil
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from .diff import (
    _CELL_REF_RE,
    _read_shared_strings,
    _sheet_cells,
    _sheet_target_map,
    diff_workbooks,
)
from .engine import SafeXLError, edit
from .structural import _fmt_range, _load_tables, add_sheet, insert_cells, insert_rows, rename_sheet, resize_table
from .verify import verify

__all__ = ["salvage"]

_EMPTY: dict[str, Any] = {"kind": "empty"}


# ---------------------------------------------------------------------------
# Loading and comparison
# ---------------------------------------------------------------------------


def _load(path: Path) -> tuple[list[str], dict[str, dict[str, dict[str, Any]]]]:
    with zipfile.ZipFile(path, "r") as zf:
        targets = _sheet_target_map(zf)
        shared = _read_shared_strings(zf)
        cells = {name: _sheet_cells(zf, tgt, shared) for name, tgt in targets.items()}
    return list(targets), cells


def _split(ref: str) -> tuple[int, int]:
    m = _CELL_REF_RE.match(ref)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return col, int(m.group(2))


def _norm_formula(f: str) -> str:
    f = f.strip()
    if f.startswith("="):
        f = f[1:]
    return re.sub(r"\s+", "", f.replace("'", ""))


def _is_shared_child(rep: dict[str, Any]) -> bool:
    return rep.get("kind") == "formula" and rep.get("formula", "").startswith("<shared formula")


def _same(a: dict[str, Any], b: dict[str, Any]) -> bool:
    if a.get("kind") != b.get("kind"):
        return False
    if a["kind"] == "empty":
        return True
    if a["kind"] == "formula":
        return _norm_formula(a["formula"]) == _norm_formula(b["formula"])
    va, vb = a["value"], b["value"]
    if isinstance(va, bool) or isinstance(vb, bool):
        return type(va) is type(vb) and va == vb
    if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
        return math.isclose(va, vb, rel_tol=1e-12, abs_tol=1e-300)
    return va == vb


def _valkey(rep: dict[str, Any]) -> Any:
    if rep["kind"] == "formula":
        return "<f>"
    v = rep.get("value")
    if isinstance(v, bool):
        return ("b", v)
    if isinstance(v, (int, float)):
        return ("n", round(float(v), 9))
    return ("s", v)


# ---------------------------------------------------------------------------
# Structural detection
# ---------------------------------------------------------------------------


def _sequence(cells: dict[str, dict[str, Any]], by_row: bool) -> list[tuple]:
    groups: dict[int, list[tuple]] = {}
    for ref, rep in cells.items():
        if rep["kind"] == "empty":
            continue
        col, row = _split(ref)
        key, other = (row, col) if by_row else (col, row)
        groups.setdefault(key, []).append((other, _valkey(rep)))
    top = max(groups, default=0)
    return [tuple(sorted(groups.get(i, []), key=lambda t: t[0])) for i in range(1, top + 1)]


def _shift_ops(so: list[tuple], sd: list[tuple]) -> list[tuple[str, int, int, int, int]]:
    """Opcodes that cannot be explained by cell edits: pure inserts/deletes and
    unequal replaces that are not at the tail of either sequence."""
    out = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, so, sd, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        if tag == "replace" and (i2 - i1) == (j2 - j1):
            continue
        if tag == "insert" and i1 >= len(so):
            continue
        if tag == "delete" and j1 >= len(sd):
            continue
        if tag == "replace" and (i2 >= len(so) or j2 >= len(sd)):
            continue
        out.append((tag, i1, i2, j1, j2))
    return out


def _similarity(a: dict[str, dict[str, Any]], b: dict[str, dict[str, Any]]) -> float:
    ka = {r for r, v in a.items() if v["kind"] != "empty"}
    if not ka:
        return 1.0
    hit = sum(1 for r in ka if r in b and _same(a[r], b[r]))
    return hit / len(ka)


# ---------------------------------------------------------------------------
# Column-scoped shifts and table growth
# ---------------------------------------------------------------------------


def _column_scoped(oc: dict[str, dict[str, Any]], dc: dict[str, dict[str, Any]]) -> list[tuple[int, int, int, int]] | None:
    """Blocks ``(first_row, count, first_col, last_col)`` when every shifted
    column shows one pure insert and nothing else shifted, else None. Columns
    with the same insert form one block; blocks must not overlap and a block's
    span may not cross a non-empty column that did not shift."""
    def cols(cells):
        out: dict[int, dict[int, Any]] = {}
        for ref, rep in cells.items():
            if rep["kind"] != "empty":
                c, r = _split(ref)
                out.setdefault(c, {})[r] = _valkey(rep)
        return out

    co, cd = cols(oc), cols(dc)
    groups: dict[tuple[int, int], list[int]] = {}
    for c in sorted(set(co) | set(cd)):
        a, b = co.get(c, {}), cd.get(c, {})
        so = [(a[r],) if r in a else () for r in range(1, max(a, default=0) + 1)]
        sd = [(b[r],) if r in b else () for r in range(1, max(b, default=0) + 1)]
        ops = _shift_ops(so, sd)
        if not ops:
            continue
        if len(ops) != 1 or ops[0][0] != "insert":
            return None
        groups.setdefault((ops[0][3], ops[0][4]), []).append(c)
    if not groups:
        return None
    blocks = []
    for (j1, j2), cs in sorted(groups.items(), key=lambda kv: min(kv[1])):
        lo, hi = min(cs), max(cs)
        if any(lo <= c <= hi and c not in cs and (c in co or c in cd) for c in set(co) | set(cd)):
            return None
        blocks.append((j1 + 1, j2 - j1, lo, hi))
    spans = sorted((b[2], b[3]) for b in blocks)
    if any(spans[i][1] >= spans[i + 1][0] for i in range(len(spans) - 1)):
        return None
    return blocks


def _col_letters(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _table_growth(cur: Path, damaged: Path, applied: list[dict[str, Any]], skipped: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Tables whose ref grew in ``damaged`` and whose added area holds an applied edit."""
    with zipfile.ZipFile(cur) as zc, zipfile.ZipFile(damaged) as zd:
        have = {(t["name"].lower(), t["sheet"]): t for t in _load_tables(zc)}
        want = _load_tables(zd)
    out = []
    for t in want:
        h = have.get((t["name"].lower(), t["sheet"]))
        if h is None or h["bounds"] == t["bounds"]:
            continue
        (c1, r1, c2, r2), (oc1, or1, oc2, or2) = t["bounds"], h["bounds"]
        label = {"sheet": t["sheet"], "ref": _fmt_range(t["bounds"])}
        if (c1, r1) != (oc1, or1) or c2 < oc2 or r2 < or2:
            skipped.append({**label, "reason": f"table {t['name']!r} changed {_fmt_range(h['bounds'])} -> {_fmt_range(t['bounds'])} (not a pure growth); not carried over"})
            continue
        explained = False
        for a in applied:
            if a["sheet"] != t["sheet"]:
                continue
            c, r = _split(a["ref"])
            if c1 <= c <= c2 and r1 <= r <= r2 and (c > oc2 or r > or2):
                explained = True
                break
        if not explained:
            skipped.append({**label, "reason": f"table {t['name']!r} grew to {_fmt_range(t['bounds'])} but no cell edit explains the added area; not carried over"})
            continue
        out.append({"table": t["name"], "sheet": t["sheet"], "old_ref": _fmt_range(h["bounds"]), "ref": _fmt_range(t["bounds"])})
    return out


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def salvage(original: str | Path, damaged: str | Path, output: str | Path, *, force: bool = False) -> dict[str, Any]:
    """Re-apply the intended cell edits found in ``damaged`` to ``original``.

    Returns ``{status, applied, skipped, structural, parts, tables, verify}``.
    ``status`` is ``salvaged`` (every difference explained: applied or
    artefact), ``partial`` (some skipped, output written), ``unsupported``
    (unhandled structural change, an engine refusal, or a failed self-check;
    NO output file) or ``nothing_to_do`` (no output written). ``force=True``
    also salvages a copy that verify calls ``rewritten`` or ``intact``.

    Raises ``SafeXLError`` only when an input file is missing.
    """
    original, damaged, output = Path(original), Path(damaged), Path(output)
    for p in (original, damaged):
        if not p.is_file():
            raise SafeXLError(f"file not found: {p}")
    result: dict[str, Any] = {
        "status": "nothing_to_do",
        "applied": [],
        "skipped": [],
        "structural": {"detected": []},
        "parts": {"unchanged": 0, "changed": []},
        "tables": [],
        "verify": None,
    }
    try:
        gate = verify(original, damaged)
        result["gate"] = {"verdict": gate["verdict"], "lost_features": [x["feature"] for x in gate["lost_features"]]}
        if not force and (gate["verdict"] == "intact" or (gate["verdict"] == "rewritten" and not gate["lost_features"])):
            return result
        _salvage(original, damaged, output, result)
    except (SafeXLError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["status"] = "unsupported"
        result["skipped"].append({"sheet": "*", "ref": "*", "reason": f"refused, no output written: {exc}"})
    if result["status"] in ("unsupported", "nothing_to_do") and output.exists() and result.get("_wrote"):
        output.unlink(missing_ok=True)
    result.pop("_wrote", None)
    return result


def _salvage(original: Path, damaged: Path, output: Path, result: dict[str, Any]) -> None:
    applied, skipped = result["applied"], result["skipped"]
    detected = result["structural"]["detected"]

    o_order, o_cells = _load(original)
    d_order, d_cells = _load(damaged)

    # ---- sheets: rename / add / remove --------------------------------
    removed = [n for n in o_order if n not in d_order]
    added = [n for n in d_order if n not in o_order]
    renames: dict[str, str] = {}  # original name -> damaged name
    for old in removed:  # a rename keeps the sheet's position and most of its content
        idx = o_order.index(old)
        new = d_order[idx] if idx < len(d_order) else None
        if new in added and new not in renames.values() and _similarity(o_cells[old], d_cells[new]) >= 0.5:
            renames[old] = new
    for old, new in renames.items():
        detected.append({"kind": "sheet_renamed", "from": old, "to": new, "handled": True})
    new_sheets = [n for n in added if n not in renames.values()]
    for n in new_sheets:
        detected.append({"kind": "sheet_added", "sheet": n, "handled": True})
    gone = [n for n in removed if n not in renames]
    for n in gone:
        detected.append({"kind": "sheet_removed", "sheet": n, "handled": False})

    # ---- rows / columns on common sheets -------------------------------
    # damaged sheet name -> blocks (first_row, count, first_col, last_col); full rows span every column
    inserts: dict[str, list[tuple[int, int, int, int]]] = {}
    for old in o_order:
        new = renames.get(old, old)
        if new not in d_cells or old in gone:
            continue
        oc, dc = o_cells[old], d_cells[new]
        col_ops = _shift_ops(_sequence(oc, False), _sequence(dc, False))
        if col_ops:
            detected.append({"kind": "columns_shifted", "sheet": new, "ops": [list(o) for o in col_ops], "handled": False})
            continue
        row_ops = _shift_ops(_sequence(oc, True), _sequence(dc, True))
        if not row_ops and (scoped := _column_scoped(oc, dc)) is None:
            continue
        if row_ops and len(row_ops) == 1 and row_ops[0][0] == "insert":
            _, _, _, j1, j2 = row_ops[0]
            inserts[new] = [(j1 + 1, j2 - j1, 1, 1 << 14)]
            detected.append({"kind": "rows_inserted", "sheet": new, "first_row": j1 + 1, "count": j2 - j1, "handled": True})
            continue
        scoped = scoped if not row_ops else _column_scoped(oc, dc)
        if scoped:
            inserts[new] = scoped
            for first, count, ca, cb in scoped:
                detected.append({"kind": "cells_inserted_down", "sheet": new, "range": f"{_col_letters(ca)}{first}:{_col_letters(cb)}{first + count - 1}", "handled": True})
            continue
        kind = "rows_deleted" if all(o[0] == "delete" for o in row_ops) else "rows_shifted_ambiguous"
        detected.append({"kind": kind, "sheet": new, "ops": [list(o) for o in row_ops], "handled": False})

    if any(not d["handled"] for d in detected):
        result["status"] = "unsupported"
        reasons = [d["kind"] + (f" ({d['sheet']})" if "sheet" in d else "") for d in detected if not d["handled"]]
        skipped.append({"sheet": "*", "ref": "*", "reason": "unhandled structural change, no output written: " + ", ".join(reasons)})
        return

    # ---- apply structural steps to a private chain of temp files -------
    with tempfile.TemporaryDirectory(prefix="safexl-salvage-") as tmp_s:
        tmp = Path(tmp_s)
        cur = tmp / f"step0{original.suffix}"
        shutil.copyfile(original, cur)
        n = 0

        def step(fn, *args, **kw) -> None:
            nonlocal cur, n
            n += 1
            nxt = tmp / f"step{n}{original.suffix}"
            fn(cur, nxt, *args, **kw)
            cur = nxt

        structural_done = False
        for old, new in renames.items():
            step(rename_sheet, old, new)
            structural_done = True
        for name in new_sheets:
            step(add_sheet, name)
            structural_done = True
        for name, blocks in inserts.items():
            for first, count, ca, cb in blocks:
                if ca == 1 and cb == 1 << 14:
                    step(insert_rows, name, first, count)
                else:
                    step(insert_cells, sheet=name, cell_range=f"{_col_letters(ca)}{first}:{_col_letters(cb)}{first + count - 1}")
                structural_done = True

        # ---- cell-level diff against the structurally aligned file ----
        _, i_cells = _load(cur)
        back = {v: k for k, v in renames.items()}
        changes: list[dict[str, Any]] = []
        for name in d_order:
            if name not in i_cells:
                continue
            ic, dc = i_cells[name], d_cells[name]
            oc = o_cells.get(back.get(name, name), {})
            blocks = inserts.get(name, [])
            for ref in sorted(set(ic) | set(dc), key=_split):
                d, i = dc.get(ref, _EMPTY), ic.get(ref, _EMPTY)
                if _same(d, i):
                    continue
                col, row = _split(ref)
                blk = next((b for b in blocks if b[2] <= col <= b[3]), None)
                if blk and blk[0] <= row < blk[0] + blk[1]:
                    o = _EMPTY
                else:
                    orow = row - blk[1] if blk and row >= blk[0] + blk[1] else row
                    o = oc.get(f"{ref[: len(ref) - len(str(row))]}{orow}", _EMPTY)
                if _same(d, o):
                    continue  # untouched by the model; differs only because of our structural step
                if _is_shared_child(i) and d["kind"] == "formula":
                    continue  # openpyxl expanded a shared formula; unverifiable, treated as untouched
                if d["kind"] == "empty":
                    skipped.append({"sheet": name, "ref": ref, "reason": "cell cleared in damaged file; the engine cannot clear cells"})
                elif d["kind"] == "formula":
                    changes.append({"sheet": name, "cell": ref, "formula": d["formula"]})
                    applied.append({"sheet": name, "ref": ref, "kind": "formula"})
                elif d.get("error"):
                    skipped.append({"sheet": name, "ref": ref, "reason": f"error value {d['value']!r} not applied"})
                elif i["kind"] == "formula":
                    skipped.append({"sheet": name, "ref": ref, "reason": "formula replaced by a plain value (data_only re-save or real overwrite; ambiguous)"})
                else:
                    changes.append({"sheet": name, "cell": ref, "value": d["value"]})
                    applied.append({"sheet": name, "ref": ref, "kind": "value"})

        if changes:
            step(edit, changes)
            structural_done = True
        # ---- tables the model grew ---------------------------------------
        tables = _table_growth(cur, damaged, applied, skipped)
        for t in tables:
            try:
                step(resize_table, table=t["table"], ref=t["ref"], sheet=t["sheet"])
            except SafeXLError as exc:
                skipped.append({"sheet": t["sheet"], "ref": t["ref"], "reason": f"table {t['table']!r} resize refused: {exc}"})
                continue
            result["tables"].append(t)
            structural_done = True

        if not structural_done:
            result["status"] = "nothing_to_do" if not skipped else "unsupported"
            return

        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(cur, output)
        result["_wrote"] = True

    # ---- self-check: the output must be certifiably intact -------------
    receipt = verify(original, output)
    result["verify"] = {"verdict": receipt["verdict"], "lost_features": [x["feature"] for x in receipt["lost_features"]]}
    if receipt["verdict"] != "intact":
        output.unlink(missing_ok=True)
        result["status"] = "unsupported"
        skipped.append({"sheet": "*", "ref": "*", "reason": "self-check failed, output deleted: " + "; ".join(receipt["reasons"][:3])})
        return
    result["status"] = "partial" if skipped else "salvaged"
    parts = diff_workbooks(original, output).parts
    result["parts"] = {"unchanged": parts["unchanged_count"], "changed": [c["path"] for c in parts["changed"]]}
