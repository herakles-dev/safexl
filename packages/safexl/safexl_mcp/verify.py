"""safexl_mcp.verify -- deterministic "is this edited workbook intact?" receipt.

``verify(before, after)`` compares an edited workbook against its original and returns a plain dict
(the receipt): a verdict, whether the package is still valid, which features were lost, whether a
non-SafeXL writer rebuilt the package, and the part-level diff. Read-only; never raises on a bad
file (a corrupt, non-zip or over-cap file is verdict "invalid"); never executes macros or fetches
anything. Resource caps (MAX_PARTS, MAX_PART_BYTES, MAX_TOTAL_BYTES) live in safexl_mcp.features.

    intact    = no feature lost and package not rebuilt
    rewritten = another writer rebuilt the package but no tracked feature was lost; parts that should
                be byte-identical are not, so nothing can be certified
    damaged   = at least one tracked feature was lost (each is named in lost_features)
    invalid   = the original or the edited file is not a valid package
Precedence: invalid > damaged > rewritten > intact.

The verdict is loss-based and only says what was LOST. Three further fields say what was ADDED or
CHANGED, because a workbook can be weaponised or silently altered without losing anything:

    added_risk   = named additions the original lacked: a VBA project, an external link part, a new
                   external relationship target, sheet protection added, a sheet that became hidden or
                   veryHidden, changed calculation settings, hidden rows/columns added, formulas calling
                   WEBSERVICE / HYPERLINK or DDE-style links. Each entry is {"risk": name, "detail": text}.
    cell_changes = per sheet (matched by sheetId, so a rename is not a change): cells whose value differs,
                   whose formula differs, whose style index differs, that were added, that were removed,
                   rows removed, and ``cells_changed`` (distinct cells with any difference). Positional:
                   an insert-rows edit reports the shifted cells.
    summary      = one plain sentence saying exactly what the receipt covers.

Cell comparison is a streamed merge-join over both sheets in document order: memory is one row per side
plus the shared-string tables, never the whole sheet. CAPS: at most MAX_CELLS_COMPARED (5,000,000) cells
per sheet are compared and at most MAX_SHARED_STRINGS (2,000,000) shared strings are resolved per
workbook; past a cap the sheet's entry carries ``"complete": false`` and its counts are lower bounds. A
sheet whose rows are not in ascending order is also reported ``"complete": false``.

Row inserts: when a sheet differs, rows are aligned by signature (values and styles, not formula text;
at most 1,100,000 rows kept per side) to find whole-row insert blocks. The sheet is then re-compared with
the shift undone and reports ``rows_inserted`` ({"count", "blocks": [{"row", "count"}]}); formulas whose
references merely moved are counted in ``formulas_shifted``, not ``formulas_changed``. A column-scoped insert
is not aligned: when the changed cells mostly reappear elsewhere in the same columns the sheet is marked
``shift_detected: false, positional: true`` and the summary says its counts are positional. Anything
uncertain stays positional; no shift is ever guessed.
"""
from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Iterator
from xml.etree import ElementTree as ET

from .features import fingerprint, lost_features, untouchable_changed
from .validate import FileReport, check_content_types, check_relationships, check_well_formed_xml

__all__ = ["verify"]

_VALIDITY_CHECKS = {"zip", "well_formed_xml", "content_types", "relationships"}

MAX_CELLS_COMPARED = 5_000_000
MAX_SHARED_STRINGS = 2_000_000
_MAX_LISTED = 20

_MAIN_NS = {
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main",
}
_REL_NS = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
}
_RISKY_FORMULAS = {
    "WEBSERVICE": re.compile(r"\bWEBSERVICE\s*\(", re.I),
    "HYPERLINK": re.compile(r"\bHYPERLINK\s*\(", re.I),
    "DDE": re.compile(r"\bDDE(?:AUTO)?\b|[A-Za-z_][\w.]*\|\s*'[^']*'\s*!", re.I),
}
_RE_REL_TAG = re.compile(rb"<(?:\w+:)?Relationship\b[^>]*>")
_RE_EXTERNAL = re.compile(rb'TargetMode\s*=\s*"External"', re.I)
_RE_TARGET = re.compile(rb'\bTarget\s*=\s*"([^"]*)"')


def _file_sha256(path: str | Path) -> str | None:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return None


def _validity_problems(path: str | Path) -> list[str]:
    # Only the four checks that feed the receipt's validity. The other validator checks (row/cell order,
    # dimension, calc chain, shared strings) walk every cell of every sheet again for findings verify ignores.
    report = FileReport(path=str(path))
    try:
        with zipfile.ZipFile(path, "r") as z:
            parsed = check_well_formed_xml(z, report)
            check_content_types(z, parsed, report)
            check_relationships(z, parsed, report)
    except zipfile.BadZipFile as e:
        report.add("zip", "CRITICAL", f"not a valid zip/OOXML package: {e}")
    return [f"[{f.check}] {f.detail}" for f in report.findings if f.check in _VALIDITY_CHECKS]


# ---------------------------------------------------------------------------
# added_risk / cell_changes / summary
# ---------------------------------------------------------------------------


def _split_tag(tag: str) -> tuple[str, str]:
    if tag.startswith("{"):
        ns, _, local = tag[1:].partition("}")
        return ns, local
    return "", tag


def _main(el_tag: str, name: str) -> bool:
    ns, local = _split_tag(el_tag)
    return local == name and ns in _MAIN_NS


def _workbook_info(zf: zipfile.ZipFile) -> tuple[dict[str, dict], dict[str, str]]:
    """({sheetId: {"name", "state", "part"}}, calcPr attributes) from xl/workbook.xml."""
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels: dict[str, str] = {}
    if "xl/_rels/workbook.xml.rels" in zf.namelist():
        for rel in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")):
            target = rel.get("Target") or ""
            if rel.get("TargetMode") == "External" or not rel.get("Id"):
                continue
            rels[rel.get("Id")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    sheets: dict[str, dict] = {}
    calc: dict[str, str] = {}
    for el in root.iter():
        if _main(el.tag, "sheet"):
            rid = next((v for k, v in el.attrib.items() if _split_tag(k)[1] == "id" and _split_tag(k)[0] in _REL_NS), None)
            sid = el.get("sheetId") or f"name:{el.get('name')}"
            sheets[sid] = {"name": el.get("name") or sid, "state": el.get("state") or "visible", "part": rels.get(rid or "")}
        elif _main(el.tag, "calcPr"):
            calc = dict(el.attrib)
    return sheets, calc


def _shared_strings(zf: zipfile.ZipFile) -> tuple[list[str] | None, bool]:
    """(strings, complete). ``strings`` is None when absent; complete False past MAX_SHARED_STRINGS."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return [], True
    out: list[str] = []
    with zf.open("xl/sharedStrings.xml") as fh:
        for _ev, el in ET.iterparse(fh, events=("end",)):
            if _main(el.tag, "si"):
                parts: list[str] = []

                def walk(node: ET.Element, skip: bool = False) -> None:
                    for ch in node:
                        if _main(ch.tag, "rPh"):
                            continue
                        if _main(ch.tag, "t"):
                            parts.append(ch.text or "")
                        else:
                            walk(ch)

                walk(el)
                out.append("".join(parts))
                el.clear()
                if len(out) > MAX_SHARED_STRINGS:
                    return None, False
    return out, True


def _value_sig(t: str | None, v: str | None, inline: str | None, sst: list[str] | None):
    if inline is not None:
        return ("s", inline)
    if v is None:
        return None
    if t == "s":
        try:
            return ("s", sst[int(v)]) if sst is not None else ("sidx", v)
        except (ValueError, IndexError):
            return ("sidx", v)
    if t in ("str",):
        return ("s", v)
    if t in ("b", "e"):
        return (t, v.strip())
    try:
        return ("n", float(v))
    except ValueError:
        return ("s", v)


def _sheet_rows(zf: zipfile.ZipFile, part: str, sst: list[str] | None, stats: dict) -> Iterator[tuple[int, bool, dict]]:
    """Yield (row number, hidden, {col: (value_sig, formula_sig, style)}) in document order.
    ``stats`` accumulates protected / hidden_cols / risky formula counts."""
    from .engine import _col_to_num

    # Only "end" events, and a dict lookup on the full tag instead of a namespace split per element:
    # this loop runs once per element of the sheet, so its constant factor is the whole cost of verify.
    kind: dict[str, str] = {}
    for ns in _MAIN_NS:
        for nm in ("row", "c", "col", "sheetProtection", "f", "v", "is", "t"):
            kind[f"{{{ns}}}{nm}"] = nm
    get = kind.get
    cols: dict[str, int] = {}
    row_no, last_col, last_row = 0, 0, 0
    cells: dict = {}
    with zf.open(part) as fh:
        for _ev, el in ET.iterparse(fh, events=("end",)):
            nm = get(el.tag)
            if nm == "c":
                ref = el.get("r") or ""
                letters = ref.rstrip("0123456789")
                col = cols.get(letters) if letters else None
                if col is None:
                    col = cols[letters] = _col_to_num(letters.upper()) if letters else last_col + 1
                last_col = col
                f_el = v_el = is_el = None
                for ch in el:
                    cn = get(ch.tag)
                    if cn == "f":
                        f_el = ch
                    elif cn == "v":
                        v_el = ch
                    elif cn == "is":
                        is_el = ch
                inline = None
                if is_el is not None:
                    inline = "".join((t.text or "") for t in is_el.iter() if get(t.tag) == "t")
                fsig = None
                if f_el is not None:
                    ftext = (f_el.text or "").strip()
                    fsig = (ftext, f_el.get("t"), f_el.get("ref"), f_el.get("si"))
                    up = ftext.upper()
                    if "WEBSERVICE" in up or "HYPERLINK" in up or "DDE" in up or "|" in up:
                        for name, rx in _RISKY_FORMULAS.items():
                            if rx.search(ftext):
                                stats["risky"][name] = stats["risky"].get(name, 0) + 1
                cells[col] = (
                    _value_sig(el.get("t"), None if v_el is None else (v_el.text or ""), inline, sst),
                    fsig,
                    el.get("s") or "0",
                )
                el.clear()
            elif nm == "row":
                row_no = int(el.get("r") or last_row + 1)
                last_row = row_no
                hidden = (el.get("hidden") or "").lower() in ("1", "true")
                yield row_no, hidden, cells
                if hidden:
                    stats["hidden_rows"] += 1
                cells, last_col = {}, 0
                el.clear()
            elif nm == "sheetProtection":
                stats["protected"] = True
            elif nm == "col" and (el.get("hidden") or "").lower() in ("1", "true"):
                stats["hidden_cols"] += 1


def _new_stats() -> dict:
    return {"protected": False, "hidden_cols": 0, "hidden_rows": 0, "risky": {}}


MAX_ROWS_ALIGNED = 1_100_000
_ALIGN_WINDOW = 64
_ALIGN_LOOKAHEAD = 3
_MAX_MOVED_KEYS = 100_000


def _row_hash(cells: dict) -> int:
    """Hash of a row's values and styles (formula text excluded: an insert rewrites it)."""
    return hash(tuple(sorted((c, v[0], v[1] is not None, v[2]) for c, v in cells.items())))


def _detect_row_inserts(B: list[tuple[int, int]], A: list[tuple[int, int]]) -> list[tuple[int, int, int]]:
    """Align before rows B to after rows A by row signature. Returns whole-row insert blocks
    [(row in before coordinates, count, first row in after coordinates)]. A block is only
    claimed when an after row with the same signature as the current before row sits k > 0 rows
    lower (within _ALIGN_WINDOW after rows) and the next _ALIGN_LOOKAHEAD rows match under the
    same shift; anything else is left as ordinary changes, so detection never invents a shift."""
    blocks: list[tuple[int, int, int]] = []
    i = j = s = 0
    budget = 8 * (len(B) + len(A)) + 1000  # window scans are bounded overall; out of budget -> no claim at all
    while i < len(B) and j < len(A):
        if budget <= 0:
            return []
        budget -= 1
        rb, hb = B[i]
        ra, ha = A[j]
        tb = rb + s
        if tb == ra and hb == ha:
            i += 1
            j += 1
            continue
        found = None
        if ra >= tb:
            for jj in range(j, min(len(A), j + _ALIGN_WINDOW + 1)):
                budget -= 1
                r2, h2 = A[jj]
                if r2 > tb and h2 == hb:
                    k = r2 - tb
                    if all(
                        i + d >= len(B) or (jj + d < len(A) and A[jj + d] == (B[i + d][0] + s + k, B[i + d][1]))
                        for d in range(1, _ALIGN_LOOKAHEAD + 1)
                    ):
                        found = (jj, k)
                        break
        if found:
            blocks.append((rb, found[1], tb))
            s += found[1]
            j = found[0]
        elif tb == ra:
            i += 1
            j += 1
        elif tb < ra:
            i += 1
        else:
            j += 1
    return blocks


def _compare_sheets(zb, pb, sst_b, zaft, pa, sst_a, st_b: dict, st_a: dict, blocks=(), fshift=None) -> dict:
    """Merge-join two sheets. ``blocks`` (from _detect_row_inserts) map before rows to their
    shifted positions; ``fshift(formula_text)`` rewrites a before formula for the same shift."""
    out = {"values_changed": 0, "formulas_changed": 0, "styles_changed": 0, "added": 0, "removed": 0,
           "rows_removed": 0, "cells_changed": 0, "complete": True}
    if blocks:
        out["rows_inserted"] = {"count": sum(k for _p, k, _a in blocks),
                                "blocks": [{"row": pa_, "count": k} for _p, k, pa_ in blocks[:_MAX_LISTED]]}
    if fshift:
        out["formulas_shifted"] = 0

    def shifted_rows(rows):
        # rows ascend, so one pointer over the (ascending) blocks gives each row's shift in O(1)
        bi, add = 0, 0
        for r, h, c in rows:
            while bi < len(blocks) and r >= blocks[bi][0]:
                add += blocks[bi][1]
                bi += 1
            yield r + add, h, c

    gb = _sheet_rows(zb, pb, sst_b, st_b) if pb else iter(())
    if blocks:
        gb = shifted_rows(gb)
    ga = _sheet_rows(zaft, pa, sst_a, st_a) if pa else iter(())
    from collections import Counter

    moved_b: Counter = Counter()
    moved_a: Counter = Counter()
    moved_n = [0, 0]  # running totals of the two counters (summing a Counter per call is quadratic)
    rows_b: list = []  # (row, hash) of rows that differ positionally, kept for the insert alignment
    rows_a: list = []
    keep_rows = not blocks

    def note(counter: Counter, which: int, col: int, cell) -> None:
        if cell is not None and cell[0] is not None and moved_n[which] < _MAX_MOVED_KEYS:
            counter[(col, cell[0])] += 1
            moved_n[which] += 1

    rb, ra = next(gb, None), next(ga, None)
    seen = 0
    last_b = last_a = 0
    seen_b = seen_a = None
    while rb is not None or ra is not None:
        # A row object stays current across iterations while the other side advances, so check order once per row.
        if rb is not None and rb is not seen_b:
            if rb[0] <= last_b:
                out["complete"] = False  # rows not ascending: cannot merge-join
                break
            last_b, seen_b = rb[0], rb
        if ra is not None and ra is not seen_a:
            if ra[0] <= last_a:
                out["complete"] = False
                break
            last_a, seen_a = ra[0], ra
        if rb is not None and (ra is None or rb[0] < ra[0]):
            out["rows_removed"] += 1
            out["removed"] += len(rb[2])
            out["cells_changed"] += len(rb[2])
            for col, cell in rb[2].items():
                note(moved_b, 0, col, cell)
            if keep_rows and len(rows_b) < MAX_ROWS_ALIGNED:
                rows_b.append((rb[0], _row_hash(rb[2])))
            seen += len(rb[2])
            rb = next(gb, None)
        elif ra is not None and (rb is None or ra[0] < rb[0]):
            out["added"] += len(ra[2])
            out["cells_changed"] += len(ra[2])
            for col, cell in ra[2].items():
                note(moved_a, 1, col, cell)
            if keep_rows and len(rows_a) < MAX_ROWS_ALIGNED:
                rows_a.append((ra[0], _row_hash(ra[2])))
            seen += len(ra[2])
            ra = next(ga, None)
        else:
            cb, ca = rb[2], ra[2]
            if cb == ca:  # identical row (C-speed dict comparison): nothing to count
                seen += len(cb)
                rb, ra = next(gb, None), next(ga, None)
                continue
            if keep_rows and len(rows_b) < MAX_ROWS_ALIGNED:
                rows_b.append((rb[0], _row_hash(cb)))
                rows_a.append((ra[0], _row_hash(ca)))
            for col in cb.keys() | ca.keys():
                x, y = cb.get(col), ca.get(col)
                if x is None:
                    out["added"] += 1
                    out["cells_changed"] += 1
                    note(moved_a, 1, col, y)
                elif y is None:
                    out["removed"] += 1
                    out["cells_changed"] += 1
                    note(moved_b, 0, col, x)
                else:
                    diff = False
                    if x[0] != y[0]:
                        out["values_changed"] += 1
                        diff = True
                        note(moved_b, 0, col, x)
                        note(moved_a, 1, col, y)
                    if x[1] != y[1]:
                        if fshift and x[1] is not None and y[1] is not None and (fshift(x[1][0]),) + x[1][1:] == y[1]:
                            out["formulas_shifted"] += 1  # same formula, references moved with the insert
                        else:
                            out["formulas_changed"] += 1
                            diff = True
                    if x[2] != y[2]:
                        out["styles_changed"] += 1
                        diff = True
                    out["cells_changed"] += diff
            seen += len(cb) + len(ca)
            rb, ra = next(gb, None), next(ga, None)
        if seen > MAX_CELLS_COMPARED:
            out["complete"] = False
            break
    if not blocks:
        # Column-scoped insert (insert_cells) is not aligned. Say so only with evidence: the cells that
        # "changed" or "were removed" mostly reappear, same column and value, among those "added" or "changed".
        overlap = sum((moved_b & moved_a).values())
        if overlap >= 2 and overlap * 2 >= min(moved_n):
            out["shift_detected"] = False
            out["positional"] = True
        if (out["added"] or out["removed"]) and len(rows_b) < MAX_ROWS_ALIGNED and len(rows_a) < MAX_ROWS_ALIGNED:
            out["_rows"] = (rows_b, rows_a)  # consumed (and removed) by the caller
    return out


# ---------------------------------------------------------------------------
# Text-level row scan (the fast path)
# ---------------------------------------------------------------------------
#
# Parsing every cell of two large sheets is what makes verify slow, and almost every row of an edited
# sheet is byte-identical to its original. So rows are LOCATED by one C-speed regex over the part text
# with comments/CDATA/PIs blanked (offsets preserved), identical rows (same number, same text) are
# dropped without parsing, and only the remaining rows are parsed with a real XML parser. Rows that moved
# because of an insert are matched by a text normalisation (cell/row refs and formula text removed) to
# align them, and compared after rewriting the before row's refs by the detected shift; a row that equals
# its shifted original is dropped too. Only rows that still differ are parsed. Applicable when both sheets
# carry byte-identical shared strings and unprefixed rows with ascending r; otherwise the streamed
# element-by-element comparison above is used. Memory: the two part texts plus the differing rows' offsets
# (at most MAX_ROWS_ALIGNED per side).

_TEXT_ROW_RE = re.compile(r'<row\b(?:[^>]*?\sr="(\d+)")?[^>]*?(?:/>|>.*?</row>)', re.DOTALL)
_TEXT_ROWNUM_RE = re.compile(r'\sr="(\d+)"')
_TEXT_REF_RE = re.compile(r'(\sr=")([A-Za-z]*)(\d+)(")')
_TEXT_NORM_RE = re.compile(r'\sr="[A-Za-z]*\d+"|<f\b[^>]*?/>|<f\b[^>]*>.*?</f>', re.DOTALL)
_TEXT_F_RE = re.compile(r"(<f\b[^>]*>)(.*?)(</f>)", re.DOTALL)


class _TextSheet:
    __slots__ = ("text", "root_open", "rows", "rest")

    def __init__(self, text: str, root_open: str, rows: list, rest: str):
        self.text, self.root_open, self.rows, self.rest = text, root_open, rows, rest


def _text_sheet(zf: zipfile.ZipFile, part: str) -> "_TextSheet | None":
    from .engine import _ROOT_OPEN_RE, mask_spans_same_length

    text = zf.read(part).decode("utf-8", errors="replace")
    masked = mask_spans_same_length(text)
    root = _ROOT_OPEN_RE.search(masked)
    if not root or "<sheetData" not in masked:
        return None
    rows, last = [], 0
    for m in _TEXT_ROW_RE.finditer(masked):
        g = m.group(1)  # row number, None when the <row> has no r attribute
        if g is None or int(g) <= last or len(rows) >= MAX_ROWS_ALIGNED:
            return None
        last = int(g)
        rows.append((last, m.start(), m.end()))
    # Everything outside the rows (<cols>, protection, merges, ...) is the text before the first row and after
    # the last one: rows are contiguous inside <sheetData>.
    rest = text[: rows[0][1]] + text[rows[-1][2]:] if rows else text
    return _TextSheet(text, text[root.start():root.end()], rows, rest)


def _text_stats(ts: "_TextSheet", stats: dict) -> None:
    """protected / hidden columns from everything outside the rows (small)."""
    for el in ET.fromstring(ts.rest).iter():
        if _main(el.tag, "sheetProtection"):
            stats["protected"] = True
        elif _main(el.tag, "col") and (el.get("hidden") or "").lower() in ("1", "true"):
            stats["hidden_cols"] += 1


def _parse_text_row(ts: "_TextSheet", row_text: str, sst, stats: dict) -> tuple[int, dict]:
    """(row number, {col: (value_sig, formula_sig, style)}) of one row's text, by a real XML parse."""
    from .engine import _col_to_num

    row = ET.fromstring(ts.root_open + row_text + "</worksheet>")[0]
    if (row.get("hidden") or "").lower() in ("1", "true"):
        stats["hidden_rows"] += 1
    cells: dict = {}
    last_col = 0
    for el in row:
        if not _main(el.tag, "c"):
            continue
        letters = (el.get("r") or "").rstrip("0123456789")
        col = _col_to_num(letters.upper()) if letters else last_col + 1
        last_col = col
        f_el = v_el = is_el = None
        for ch in el:
            if _main(ch.tag, "f"):
                f_el = ch
            elif _main(ch.tag, "v"):
                v_el = ch
            elif _main(ch.tag, "is"):
                is_el = ch
        inline = None if is_el is None else "".join((t.text or "") for t in is_el.iter() if _main(t.tag, "t"))
        fsig = None
        if f_el is not None:
            ftext = (f_el.text or "").strip()
            fsig = (ftext, f_el.get("t"), f_el.get("ref"), f_el.get("si"))
            for name, rx in _RISKY_FORMULAS.items():
                if rx.search(ftext):
                    stats["risky"][name] = stats["risky"].get(name, 0) + 1
        cells[col] = (_value_sig(el.get("t"), None if v_el is None else (v_el.text or ""), inline, sst), fsig, el.get("s") or "0")
    return int(row.get("r") or 0), cells


def _text_scan(tb: "_TextSheet", ta: "_TextSheet"):
    """Rows that are not byte-identical at the same number: ([(r, start, end, norm_hash)], same for after)."""
    db: list = []
    da: list = []
    i = j = 0
    rb, ra = tb.rows, ta.rows
    while i < len(rb) or j < len(ra):
        if j >= len(ra) or (i < len(rb) and rb[i][0] < ra[j][0]):
            r, s, e = rb[i]
            db.append((r, s, e, hash(_TEXT_NORM_RE.sub("", tb.text[s:e]))))
            i += 1
        elif i >= len(rb) or ra[j][0] < rb[i][0]:
            r, s, e = ra[j]
            da.append((r, s, e, hash(_TEXT_NORM_RE.sub("", ta.text[s:e]))))
            j += 1
        else:
            (r, s, e), (r2, s2, e2) = rb[i], ra[j]
            if e - s != e2 - s2 or tb.text[s:e] != ta.text[s2:e2]:
                db.append((r, s, e, hash(_TEXT_NORM_RE.sub("", tb.text[s:e]))))
                da.append((r2, s2, e2, hash(_TEXT_NORM_RE.sub("", ta.text[s2:e2]))))
            i += 1
            j += 1
    return db, da


def _shift_row_text(text: str, shift: int, fshift) -> str:
    """The before row's text as it would read after the insert: row/cell refs moved, formulas rewritten."""
    if shift:
        text = _TEXT_REF_RE.sub(lambda m: m.group(1) + m.group(2) + str(int(m.group(3)) + shift) + m.group(4), text)
    if fshift is not None and "<f" in text:
        from xml.sax.saxutils import escape, unescape

        text = _TEXT_F_RE.sub(lambda m: m.group(1) + escape(fshift(unescape(m.group(2)))) + m.group(3), text)
    return text


def _compare_text_empty() -> dict:
    return {"values_changed": 0, "formulas_changed": 0, "styles_changed": 0, "added": 0, "removed": 0,
            "rows_removed": 0, "cells_changed": 0, "complete": True}


def _compare_text(tb, ta, db, da, sst_b, sst_a, st_b: dict, st_a: dict, blocks=(), fshift=None) -> dict:
    """Same result shape as _compare_sheets, from the differing rows only."""
    from collections import Counter

    out = {"values_changed": 0, "formulas_changed": 0, "styles_changed": 0, "added": 0, "removed": 0,
           "rows_removed": 0, "cells_changed": 0, "complete": True}
    if blocks:
        out["rows_inserted"] = {"count": sum(k for _p, k, _a in blocks),
                                "blocks": [{"row": pa_, "count": k} for _p, k, pa_ in blocks[:_MAX_LISTED]]}
    if fshift:
        out["formulas_shifted"] = 0
    moved = (Counter(), Counter())
    moved_n = [0, 0]

    def note(which: int, col: int, cell) -> None:
        if cell is not None and cell[0] is not None and moved_n[which] < _MAX_MOVED_KEYS:
            moved[which][(col, cell[0])] += 1
            moved_n[which] += 1

    _text_stats(tb, st_b)
    _text_stats(ta, st_a)
    bi, add, i, j, seen = 0, 0, 0, 0, 0
    while i < len(db) or j < len(da):
        mb = None
        if i < len(db):
            while bi < len(blocks) and db[i][0] >= blocks[bi][0]:
                add += blocks[bi][1]
                bi += 1
            mb = db[i][0] + add
        if j >= len(da) or (mb is not None and mb < da[j][0]):
            _r, cb = _parse_text_row(tb, tb.text[db[i][1]:db[i][2]], sst_b, st_b)
            out["rows_removed"] += 1
            out["removed"] += len(cb)
            out["cells_changed"] += len(cb)
            for col, cell in cb.items():
                note(0, col, cell)
            i += 1
        elif mb is None or da[j][0] < mb:
            _r, ca = _parse_text_row(ta, ta.text[da[j][1]:da[j][2]], sst_a, st_a)
            out["added"] += len(ca)
            out["cells_changed"] += len(ca)
            for col, cell in ca.items():
                note(1, col, cell)
            j += 1
        else:
            text_b, text_a = tb.text[db[i][1]:db[i][2]], ta.text[da[j][1]:da[j][2]]
            if blocks and _shift_row_text(text_b, add, fshift) == text_a:
                out["formulas_shifted"] += text_b.count("<f")  # the row only moved: nothing to count
                i += 1
                j += 1
                continue
            _r, cb = _parse_text_row(tb, text_b, sst_b, st_b)
            _r, ca = _parse_text_row(ta, text_a, sst_a, st_a)
            i += 1
            j += 1
            seen += len(cb) + len(ca)
            for col in cb.keys() | ca.keys():
                x, y = cb.get(col), ca.get(col)
                if x is None:
                    out["added"] += 1
                    out["cells_changed"] += 1
                    note(1, col, y)
                elif y is None:
                    out["removed"] += 1
                    out["cells_changed"] += 1
                    note(0, col, x)
                else:
                    diff = False
                    if x[0] != y[0]:
                        out["values_changed"] += 1
                        diff = True
                        note(0, col, x)
                        note(1, col, y)
                    if x[1] != y[1]:
                        if fshift and x[1] is not None and y[1] is not None and (fshift(x[1][0]),) + x[1][1:] == y[1]:
                            out["formulas_shifted"] += 1
                        else:
                            out["formulas_changed"] += 1
                            diff = True
                    if x[2] != y[2]:
                        out["styles_changed"] += 1
                        diff = True
                    out["cells_changed"] += diff
        if seen > MAX_CELLS_COMPARED:
            out["complete"] = False
            break
    if not blocks:
        overlap = sum((moved[0] & moved[1]).values())
        if overlap >= 2 and overlap * 2 >= min(moved_n):
            out["shift_detected"] = False
            out["positional"] = True
    return out


def _external_targets(zf: zipfile.ZipFile) -> set[str]:
    found: set[str] = set()
    for name in zf.namelist():
        if name.endswith(".rels"):
            for tag in _RE_REL_TAG.findall(zf.read(name)):
                if _RE_EXTERNAL.search(tag):
                    m = _RE_TARGET.search(tag)
                    if m:
                        found.add(m.group(1).decode("utf-8", errors="replace"))
    return found


def _analyse_changes(before: str | Path, after: str | Path, fp_before: dict, fp_after: dict) -> tuple[list[dict], dict, list[str]]:
    """-> (added_risk, cell_changes, problems). Never raises on malformed XML: ``problems`` says what was skipped."""
    risks: list[dict] = []
    changes: dict[str, dict] = {}
    problems: list[str] = []

    def risk(name: str, detail: str) -> None:
        risks.append({"risk": name, "detail": detail})

    fb, fa = fp_before["features"], fp_after["features"]
    if fa["vba"] and not fb["vba"]:
        risk("vba_project_added", "the edited file contains a VBA project (macros) the original did not have")
    if fa["external_links"] > fb["external_links"]:
        risk("external_link_added", f"external link parts: {fb['external_links']} -> {fa['external_links']}")

    with zipfile.ZipFile(before) as zb, zipfile.ZipFile(after) as za:
        new_targets = sorted(_external_targets(za) - _external_targets(zb))
        if new_targets:
            shown = ", ".join(new_targets[:_MAX_LISTED]) + (", ..." if len(new_targets) > _MAX_LISTED else "")
            risk("external_target_added", f"{len(new_targets)} new external relationship target(s): {shown}")
        try:
            sheets_b, calc_b = _workbook_info(zb)
            sheets_a, calc_a = _workbook_info(za)
        except (ET.ParseError, KeyError) as exc:
            problems.append(f"workbook.xml could not be compared ({type(exc).__name__})")
            return risks, changes, problems

        diffs = [f"{k}: {calc_b.get(k, 'absent')} -> {calc_a.get(k, 'absent')}"
                 for k in sorted((calc_b.keys() | calc_a.keys()) - {"fullCalcOnLoad", "calcId"})
                 if calc_b.get(k) != calc_a.get(k)]
        if diffs:
            risk("calculation_settings_changed", "calcPr " + "; ".join(diffs))

        for sid, info in sheets_a.items():
            was = sheets_b.get(sid)
            if info["state"] in ("hidden", "veryHidden") and (was is None or was["state"] != info["state"]):
                risk("sheet_hidden", f"sheet {info['name']!r} state is now {info['state']}")

        try:
            sst_b, ok_b = _shared_strings(zb)
            sst_a, ok_a = _shared_strings(za)
        except ET.ParseError as exc:
            problems.append(f"shared strings could not be read ({type(exc).__name__})")
            sst_b = sst_a = None
            ok_b = ok_a = False
        tot_b = _new_stats()
        tot_a = _new_stats()
        protected_added: list[str] = []
        hidden_added: list[str] = []
        names = zb.namelist(), za.namelist()
        order = list(sheets_a) + [k for k in sheets_b if k not in sheets_a]

        def parts_of(sid: str):
            ib, ia = sheets_b.get(sid), sheets_a.get(sid)
            pb = ib["part"] if ib and ib["part"] in names[0] else None
            pa = ia["part"] if ia and ia["part"] in names[1] else None
            return ib, ia, pb, pa

        # Phase 1. A sheet pair is scanned at text level when possible (identical rows dropped without
        # parsing); otherwise it is compared element by element. Sheets that differ get a row-alignment
        # attempt so a whole-row insert is not reported as thousands of changed cells.
        results: dict[str, tuple] = {}
        inserts: dict[str, list] = {}
        texts: dict[str, tuple] = {}
        try:
            same_sst = ("xl/sharedStrings.xml" in names[0]) == ("xl/sharedStrings.xml" in names[1]) and (
                "xl/sharedStrings.xml" not in names[0]
                or zb.read("xl/sharedStrings.xml") == za.read("xl/sharedStrings.xml")
            )
        except (OSError, zipfile.BadZipFile, RuntimeError):
            same_sst = False
        for sid in order:
            ib, ia, pb, pa = parts_of(sid)
            label = (ia or ib)["name"]
            st_b, st_a = _new_stats(), _new_stats()
            try:
                if pb and pa and same_sst and ok_b and ok_a:
                    if fp_before["parts"].get(pb) == fp_after["parts"].get(pa):
                        results[sid] = (label, _compare_text_empty(), st_b, st_a)
                        continue
                    tb, ta = _text_sheet(zb, pb), _text_sheet(za, pa)
                    if tb is not None and ta is not None:
                        db, da = _text_scan(tb, ta)
                        if {x[0] for x in db} != {x[0] for x in da}:
                            found = _detect_row_inserts([(x[0], x[3]) for x in db], [(x[0], x[3]) for x in da])
                            if found:
                                inserts[sid] = found
                        texts[sid] = (tb, ta, db, da)
                        continue
                res = _compare_sheets(zb, pb, sst_b, za, pa, sst_a, st_b, st_a)
                rows = res.pop("_rows", None)  # row signatures collected during the same pass: no extra scan
                if rows and pb and pa and res["complete"] and ok_b and ok_a:
                    found = _detect_row_inserts(*rows)
                    if found:
                        inserts[sid] = found
            except (ET.ParseError, ValueError, KeyError) as exc:
                problems.append(f"sheet {label!r} could not be compared ({type(exc).__name__})")
                continue
            results[sid] = (label, res, st_b, st_a)

        # Phase 2: text-scanned sheets are counted now; when rows were inserted anywhere, affected sheets are
        # compared with the shift undone and every formula is rewritten for every detected insert.
        from .refs import shift_rows_in_formula

        plain_ref = re.compile(r"(?<![A-Za-z0-9_.$!])(\$?[A-Za-z]{1,3}\$?)(\d+)(?![A-Za-z0-9_(.])")

        def make_fshift(owner: str):
            if not inserts:
                return None
            own_blocks = [b for tsid, blks in inserts.items() if sheets_b[tsid]["name"].lower() == owner.lower() for b in blks]

            def fshift(text: str) -> str:
                # Fast path, one regex pass: a formula with no sheet qualifier, quote or bracket only holds
                # references to its own sheet, so it is shifted by that sheet's blocks and nothing else.
                if "!" not in text and "'" not in text and '"' not in text and "[" not in text:
                    for b in own_blocks:
                        p, k = b[2], b[1]
                        text = plain_ref.sub(lambda m: m.group(1) + str(int(m.group(2)) + k) if int(m.group(2)) >= p else m.group(0), text)
                    return text
                for tsid, blks in inserts.items():
                    target = sheets_b[tsid]["name"]
                    for _p, k, pa_ in blks:
                        try:
                            text = shift_rows_in_formula(text, owner, target, pa_, k)
                        except Exception:  # noqa: BLE001 - leave an unparseable formula as it was
                            pass
                return text
            return fshift

        for sid, (tb, ta, db, da) in texts.items():
            ib, ia, pb, pa = parts_of(sid)
            label = (ia or ib)["name"]
            st_b, st_a = _new_stats(), _new_stats()
            try:
                res2 = _compare_text(tb, ta, db, da, sst_b, sst_a, st_b, st_a, inserts.get(sid, ()), make_fshift((ib or ia)["name"]))
            except (ET.ParseError, ValueError, KeyError) as exc:
                problems.append(f"sheet {label!r} could not be compared ({type(exc).__name__})")
                continue
            results[sid] = (label, res2, st_b, st_a)
        if inserts:
            for sid in list(results):
                label, res, _sb, _sa = results[sid]
                if sid in texts or not res["cells_changed"]:
                    continue
                ib, ia, pb, pa = parts_of(sid)
                st_b, st_a = _new_stats(), _new_stats()
                try:
                    res2 = _compare_sheets(zb, pb, sst_b, za, pa, sst_a, st_b, st_a, inserts.get(sid, ()),
                                           make_fshift((ib or ia)["name"]))
                except (ET.ParseError, ValueError, KeyError) as exc:
                    problems.append(f"sheet {label!r} could not be compared ({type(exc).__name__})")
                    continue
                results[sid] = (label, res2, st_b, st_a)


        for sid in order:
            if sid not in results:
                continue
            label, res, st_b, st_a = results[sid]
            if not (ok_b and ok_a):
                res["complete"] = False
            if any(v for k, v in res.items() if k != "complete") or not res["complete"]:
                changes[label] = res
            for tot, st in ((tot_b, st_b), (tot_a, st_a)):
                for name, n in st["risky"].items():
                    tot["risky"][name] = tot["risky"].get(name, 0) + n
            if st_a["protected"] and not st_b["protected"]:
                protected_added.append(label)
            more = (st_a["hidden_cols"] - st_b["hidden_cols"]) + (st_a["hidden_rows"] - st_b["hidden_rows"])
            if more > 0:
                hidden_added.append(f"{label} (+{more})")
        if protected_added:
            risk("sheet_protection_added", "sheet protection added on " + ", ".join(protected_added))
        if hidden_added:
            risk("hidden_rows_or_columns_added", "more hidden rows/columns on " + ", ".join(hidden_added))
        for name in _RISKY_FORMULAS:
            n_b, n_a = tot_b["risky"].get(name, 0), tot_a["risky"].get(name, 0)
            if n_a > n_b:
                risk(f"formula_{name.lower()}_added", f"formulas calling {name}: {n_b} -> {n_a}")
    return risks, changes, problems


def _plural(n: int, word: str) -> str:
    return f"{n} {word}" + ("" if n == 1 else "s")


def _summary(receipt: dict[str, Any]) -> str:
    verdict = receipt["verdict"]
    if verdict == "invalid":
        return "Not checked: the original or the edited file is not a valid workbook package."
    lost = receipt["lost_features"]
    if verdict == "damaged":
        names = ", ".join(x["feature"] for x in lost)
        lead = f"{_plural(len(lost), 'tracked feature')} lost ({names})"
    elif verdict == "rewritten":
        lead = "Another tool rewrote the file, so untouched parts are not byte-identical and nothing can be certified; no tracked feature was lost"
    else:
        lead = "No tracked feature was lost and untouched parts are byte-identical"
    cc = receipt["cell_changes"]
    bits = []
    for sheet, c in cc.items():
        if c["cells_changed"]:
            bits.append(f"{_plural(c['cells_changed'], 'cell')} changed on {sheet}")
        if c["rows_removed"]:
            bits.append(f"{_plural(c['rows_removed'], 'row')} removed on {sheet}")
        if c.get("rows_inserted"):
            bits.append(f"{_plural(c['rows_inserted']['count'], 'row')} inserted on {sheet}")
        if c.get("positional"):
            bits.append(f"cells moved within {sheet}; counts for that sheet are positional")
    cells = "; ".join(bits) if bits else "no cell changes"
    if any(not c["complete"] for c in cc.values()):
        cells += " (comparison capped or unordered: counts are lower bounds)"
    risk = receipt["added_risk"]
    risky = "nothing risky was added" if not risk else "risky additions: " + ", ".join(r["risk"] for r in risk)
    return f"{lead}; {cells}; {risky}."


def verify(before: str | Path, after: str | Path) -> dict[str, Any]:
    """Receipt for the edit that turned ``before`` into ``after``."""
    receipt: dict[str, Any] = {
        "verdict": "invalid",
        "valid": False,
        "lost_features": [],
        "rebuilt": False,
        "parts": {"unchanged": 0, "changed": [], "added": [], "removed": []},
        "before_sha256": _file_sha256(before),
        "after_sha256": _file_sha256(after),
        "reasons": [],
        "added_risk": [],
        "cell_changes": {},
        "summary": "",
    }
    reasons: list[str] = receipt["reasons"]
    fp_before, fp_after = fingerprint(before), fingerprint(after)
    for label, fp in (("original", fp_before), ("edited file", fp_after)):
        if not fp["valid_zip"]:
            reasons.append(f"{label} is not a readable workbook package: {fp['error']}")
    if reasons:
        receipt["summary"] = _summary(receipt)
        return receipt

    problems = _validity_problems(after)
    receipt["valid"] = not problems
    reasons.extend(f"edited file is not a valid package: {p}" for p in problems)

    pb, pa = fp_before["parts"], fp_after["parts"]
    common = pb.keys() & pa.keys()
    receipt["parts"] = {
        "unchanged": sum(1 for n in common if pb[n] == pa[n]),
        "changed": sorted(n for n in common if pb[n] != pa[n]),
        "added": sorted(pa.keys() - pb.keys()),
        "removed": sorted(pb.keys() - pa.keys()),
    }

    lost = lost_features(fp_before, fp_after)
    receipt["lost_features"] = lost
    for x in lost:
        more = "; ".join(x.get("detail", [])[:3])
        reasons.append(f"lost {x['feature']}: {x['before']} -> {x['after']}" + (f" ({more})" if more else ""))

    rewritten = untouchable_changed(pb, pa)
    receipt["rebuilt"] = bool(rewritten)
    if rewritten:
        shown = ", ".join(rewritten[:6]) + (", ..." if len(rewritten) > 6 else "")
        reasons.append(f"rebuilt by a non-SafeXL writer: {shown} rewritten")
        if not lost and receipt["valid"]:
            reasons.append("The file was rewritten by another tool, so its untouched parts are no longer "
                           "byte-identical to the original and nothing can be certified; no tracked feature was lost.")

    if not receipt["valid"]:
        receipt["verdict"] = "invalid"
    elif lost:
        receipt["verdict"] = "damaged"
    elif rewritten:
        receipt["verdict"] = "rewritten"
    else:
        receipt["verdict"] = "intact"

    try:
        risks, changes, problems = _analyse_changes(before, after, fp_before, fp_after)
    except (zipfile.BadZipFile, OSError, ET.ParseError, KeyError, ValueError, RuntimeError) as exc:
        risks, changes, problems = [], {}, [f"added-risk and cell comparison skipped ({type(exc).__name__}: {exc})"]
    receipt["added_risk"], receipt["cell_changes"] = risks, changes
    reasons.extend(f"not compared: {p}" for p in problems)
    receipt["summary"] = _summary(receipt)
    return receipt
