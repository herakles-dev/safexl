"""SafeXL structural edits: insert_rows, rename_sheet, add_sheet.

Same XML-injection technique as ``safexl_mcp.engine`` (module docstring
there has the full rationale): every zip member is copied byte-for-byte
except the specific parts a given operation must touch, patched with
regex/string surgery rather than a parse/rebuild round trip, so charts,
pivot tables, conditional formatting, data validation, and VBA are never
touched unless a reference inside them has to follow the edit. Every
function here follows the same rollback promise as ``engine.edit()``: it
writes to a private temp file, runs ``safexl_mcp.validate.validate_workbook``
against it, and only moves it into place at ``dst`` if the edit introduced no
new CRITICAL finding (findings the source already had are not held against
the edit); on any failure ``dst`` is left completely untouched and a
``SafeXLError`` is raised.

Reference text is rewritten by ``safexl_mcp.refs`` (a formula tokenizer that
skips string literals, structured references and external-workbook refs).

insert_rows
-----------
Inserts ``count`` blank rows at ``row`` in ``sheet_name``; the old row
``row`` and everything below moves down by ``count``. A blank row is
represented the way Excel does: there is simply no ``<row>`` element there.
Every endpoint at or below ``row`` moves; a range that straddles ``row``
grows (Excel's rule), one that ends just above it does not.

Moved or rewritten, in the target sheet: ``<row>``/``<c>`` numbers,
``<dimension>``, merged cells, ``sqref`` of conditional formatting / data
validation / ignoredErrors / protectedRanges, hyperlink ``ref``, autoFilter
and sortState, array and shared-formula ``ref=``, x14 ``<xm:sqref>``
(sparkline locations, x14 CF/DV), comments (``<comment ref>``, threaded
comments, VML note anchors), drawing anchors (charts/images move with their
cells unless ``editAs="absolute"``), tables (``ref``, autoFilter, sortState),
pivot-table ``<location>``.

Rewritten wherever they appear, when they point into the target sheet:
``<f>`` / CF ``<formula>`` / DV ``<formula1|2>`` / x14 ``<xm:f>`` (so
sparkline data ranges and cross-sheet formulas follow), hyperlink
``location``, defined names (including Print_Area / Print_Titles / filter
names), chart series (``<c:f>``), pivot cache ``worksheetSource ref``.

Shared formulas: a shared group is expanded to explicit per-cell formulas
when the insert would break it (the insert splits the group, or the shift
changes any member's references), because a shared child is computed from
the master's position. Groups the insert does not affect stay shared and
only get their ``ref=`` shifted.

Because rows physically move, ``xl/calcChain.xml`` is dropped unconditionally
(same as ``engine.edit()``) and ``xl/workbook.xml``'s ``<calcPr>`` gets
``fullCalcOnLoad="1"`` forced, so the next open recomputes everything.

Not handled (left as they were, nothing is corrupted): ``<cfvo type="formula"
val=...>`` attribute refs, row page breaks, sheet-view selection/pane cells,
3-D refs (``Sheet1:Sheet3!A1``), refs inside connection/query definitions.

rename_sheet
------------
Renames ``xl/workbook.xml``'s ``<sheet name="...">`` and rewrites every
reference to the old name, quoted or not, matched case-insensitively:
formulas on every sheet (cells, CF, DV, x14 sparklines/CF/DV), defined names,
chart series, table formulas, hyperlink ``location``, drawing ``textlink``,
and pivot cache ``worksheetSource sheet=``. The sheet's zip path is never
touched. Formula *values* are unaffected, so ``calcChain.xml`` is kept. VBA
source in ``vbaProject.bin`` is binary and is never edited: macro code that
addresses the sheet by its old *name* (``Sheets("Old")``) still needs a
human; code that uses the sheet's CodeName keeps working.

add_sheet
---------
Adds a new, empty worksheet: a minimal ``<sheetData/>`` part at the next
available ``xl/worksheets/sheetN.xml`` path, a ``<sheet>`` entry in
``xl/workbook.xml`` with a fresh ``sheetId``/``r:id``, a matching
``<Relationship>`` in ``xl/_rels/workbook.xml.rels``, and a matching
``<Override>`` in ``[Content_Types].xml``. Nothing else in the workbook is
modified.
"""
from __future__ import annotations

import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Sequence
from xml.etree import ElementTree as ET

from .engine import (
    Change,
    SafeXLError,
    _DIMENSION_RE,
    _dec,
    _enc,
    _refuse_prefixed_worksheet,
    _iter_cells,
    _local,
    _group_changes,
    _REL_NS,
    _col_to_num,
    _inject_cells,
    _num_to_col,
    _rel_prefix_alt,
    _force_full_calc_on_load,
    _normalize_self_closing,
    _resolve_sheet_zip_path,
    _split_ref,
    _strip_calc_chain_override,
    _strip_calc_chain_relationship,
)
from .refs import (
    MAX_COL,
    MAX_ROW,
    TornReference,
    escape_attr,
    quote_sheet_name_if_needed,
    rename_sheet_in_formula,
    shift_cells_in_formula,
    shift_ref_cols,
    shift_rows_in_formula,
    shift_sqref,
    shift_sqref_cols,
    translate_formula,
    xml_escape,
    xml_unescape,
)
from .validate import validate_workbook

__all__ = ["insert_rows", "rename_sheet", "add_sheet", "resize_table", "plan_table_extensions", "insert_cells"]

_MAX_ROW = MAX_ROW


# ---------------------------------------------------------------------------
# Shared: write-to-temp / validate / move (same rollback promise as engine.edit)
# ---------------------------------------------------------------------------


def _atomic_write(
    src: Path,
    dst: Path,
    build: Callable[[zipfile.ZipFile, zipfile.ZipFile], None],
    op_name: str,
    check: Callable[[zipfile.ZipFile, zipfile.ZipFile], None] | None = None,
) -> None:
    tmp_dir = dst.parent if dst.parent.exists() else Path(tempfile.gettempdir())
    tmp_fd, tmp_path_str = tempfile.mkstemp(suffix=dst.suffix or ".xlsx", dir=tmp_dir)
    os.close(tmp_fd)
    tmp_path = Path(tmp_path_str)
    try:
        with zipfile.ZipFile(src, "r") as zs, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            build(zs, zout)

        if check is not None:  # read-back: re-read the output with a real parser; raises SafeXLError
            with zipfile.ZipFile(src, "r") as zbefore, zipfile.ZipFile(tmp_path, "r") as zafter:
                try:
                    check(zbefore, zafter)
                except ET.ParseError as exc:
                    raise SafeXLError(f"{op_name} read-back failed: the written part is not parseable ({exc})") from exc

        report = validate_workbook(tmp_path)
        if report.critical_count:
            # A finding the source already had is not this edit's doing (e.g. a
            # stale calcChain in a workbook Excel itself wrote); only new ones block.
            try:
                inherited = set(validate_workbook(src).critical_details())
            except Exception:  # noqa: BLE001 - baseline is best-effort
                inherited = set()
            new_details = [d for d in report.critical_details() if d not in inherited]
            if new_details:
                raise SafeXLError(
                    f"{op_name} produced an invalid workbook, dst NOT written "
                    f"({len(new_details)} new critical structural finding(s)): {'; '.join(new_details)}"
                )
        shutil.move(str(tmp_path), str(dst))
    except SafeXLError:
        raise
    except Exception as exc:  # noqa: BLE001 - re-raise everything as SafeXLError
        raise SafeXLError(f"{op_name} failed to write workbook: {exc}") from exc
    finally:
        tmp_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Read-back checks for the structural ops (real XML parser, bounded)
# ---------------------------------------------------------------------------
#
# CAPS: row-shifting ops compare a SAMPLE of rows, not every row: the rows around
# the change point (ROW_WINDOW on each side) plus every row whose number is a
# multiple of ROW_STRIDE (<= 10,800 rows, about 1% of the sheet limit), and stop
# collecting after READBACK_MAX_CELLS cells. Streaming keeps the cost linear in
# the sheet size and the memory bounded by those caps.

READBACK_ROW_STRIDE = 97
READBACK_ROW_WINDOW = 5
READBACK_MAX_CELLS = 1_000_000


def _workbook_sheet_names(zf: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    return [el.get("name") or "" for el in root.iter() if _local(el.tag) == "sheet"]


def _sample_row(r: int, row: int, count: int) -> bool:
    # Both edges of the insert are sampled, not every row between them: the cost must not grow with ``count``.
    w = READBACK_ROW_WINDOW
    return r % READBACK_ROW_STRIDE == 0 or abs(r - row) <= w or abs(r - (row + count)) <= w


def _row_cells(zf: zipfile.ZipFile, part: str, wanted: Callable[[int], bool]) -> dict[int, dict[int, tuple]]:
    """{row: {col: (t, has_formula, v, inline)}} for rows where ``wanted(row)``.
    Rows are keyed by the cell's own ref (hand-built workbooks do not always nest cells in the matching <row>)."""
    out: dict[int, dict[int, tuple]] = {}
    total = 0
    for row_r, ref, t, f, v, inline, in_row in _iter_cells(zf, part, wanted):
        m = re.match(r"([A-Za-z]+)(\d+)$", ref or "")
        if not m:
            raise SafeXLError(f"read-back failed on {part}: cell has no valid ref ({ref!r})")
        r = int(m.group(2))
        if not wanted(r):
            continue
        total += 1
        if total > READBACK_MAX_CELLS:
            break
        out.setdefault(r, {})[_col_to_num(m.group(1).upper())] = (t, f is not None, v, inline)
    return out


def _content(cell: tuple) -> bool:
    return cell[1] or cell[2] not in (None, "") or cell[3] not in (None, "")


def _readback_row_shift(
    zbefore: zipfile.ZipFile,
    zafter: zipfile.ZipFile,
    parts: dict[str, str],
    sheet_name: str,
    row: int,
    count: int,
    c1: int = 1,
    c2: int = MAX_COL,
    op: str = "insert_rows",
) -> None:
    """Cells of columns c1..c2 at rows >= ``row`` must now sit ``count`` rows lower
    (same type/value; formula text may legitimately be shifted so only its presence
    is compared), the ``count`` rows opened at ``row`` must be blank in those
    columns, and everything else must be where it was."""
    part = next((p for p, n in parts.items() if n.lower() == sheet_name.lower()), None)
    if part is None:
        raise SafeXLError(f"{op} read-back failed: sheet {sheet_name!r} missing from the output")
    keep = lambda r: _sample_row(r, row, count)  # noqa: E731
    before = _row_cells(zbefore, part, keep)
    # the opened rows are wanted too (they should be absent, so this costs nothing), to prove they are blank
    after = _row_cells(zafter, part, lambda r: keep(r) or keep(r - count) or row <= r < row + count)
    for r, cells in before.items():
        for col, cell in cells.items():
            dest = r + count if (r >= row and c1 <= col <= c2) else r
            got = after.get(dest, {}).get(col)
            if got is None or got[:3] != cell[:3] or (got[3] != cell[3]):
                raise SafeXLError(
                    f"{op} read-back failed on {sheet_name}: the cell at {_num_to_col(col)}{r} "
                    f"is not at {_num_to_col(col)}{dest} in the output"
                )
    for r in range(row, row + count):
        for col, cell in after.get(r, {}).items():
            if c1 <= col <= c2 and _content(cell):
                raise SafeXLError(
                    f"{op} read-back failed on {sheet_name}: {_num_to_col(col)}{r} should be blank "
                    "at the insert point but holds content"
                )


# ---------------------------------------------------------------------------
# Shared: sheet-name <-> zip-path resolution for every sheet at once
# ---------------------------------------------------------------------------


def _all_worksheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    """{worksheet zip path: sheet name} for every <sheet> in xl/workbook.xml."""
    wb_xml = _dec(zf.read("xl/workbook.xml"))
    rels_xml = _dec(zf.read("xl/_rels/workbook.xml.rels"))

    rid_to_target: dict[str, str] = {}
    for rel_m in re.finditer(r"<Relationship\b[^>]*/>", rels_xml):
        tag = rel_m.group(0)
        id_m = re.search(r'\bId="([^"]+)"', tag)
        target_m = re.search(r'\bTarget="([^"]+)"', tag)
        if not id_m or not target_m:
            continue
        target = target_m.group(1).lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        rid_to_target[id_m.group(1)] = target

    out: dict[str, str] = {}
    for sheet_m in re.finditer(r"<sheet\b[^>]*/>", wb_xml):
        tag = sheet_m.group(0)
        name_m = re.search(r'\bname="([^"]*)"', tag)
        rid_m = re.search(rf'\b(?:{_rel_prefix_alt(wb_xml)}):id="([^"]+)"', tag)
        if not name_m or not rid_m:
            continue
        target = rid_to_target.get(rid_m.group(1))
        if target:
            out[target] = xml_unescape(name_m.group(1))
    return out


# Parts a worksheet points at through its own .rels, by relationship-type suffix.
_RELATED_KINDS = {
    "table": "table",
    "comments": "comments",
    "threadedComment": "threadedComment",
    "vmlDrawing": "vml",
    "drawing": "drawing",
    "pivotTable": "pivotTable",
}


def _related_parts(zf: zipfile.ZipFile, sheet_of_part: dict[str, str]) -> dict[str, tuple[str, str]]:
    """{part path: (kind, owning sheet name)} for tables, comments, VML, drawings
    and pivot tables reachable from a worksheet's .rels."""
    names = set(zf.namelist())
    out: dict[str, tuple[str, str]] = {}
    for sheet_path, sheet_name in sheet_of_part.items():
        rels_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
        if rels_path not in names:
            continue
        rels_xml = _dec(zf.read(rels_path))
        for m in re.finditer(r"<Relationship\b[^>]*/>", rels_xml):
            tag = m.group(0)
            if 'TargetMode="External"' in tag:
                continue
            type_m = re.search(r'\bType="[^"]*/([A-Za-z]+)"', tag)
            target_m = re.search(r'\bTarget="([^"]+)"', tag)
            if not type_m or not target_m or type_m.group(1) not in _RELATED_KINDS:
                continue
            target = target_m.group(1)
            path = (
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target))
            )
            if path in names:
                out[path] = (_RELATED_KINDS[type_m.group(1)], sheet_name)
    return out


# ---------------------------------------------------------------------------
# Shared: regex plumbing over reference-bearing XML
# ---------------------------------------------------------------------------

# Elements whose text is a formula: cell <f>, CF <formula>, DV <formula1/2>,
# x14 <xm:f>, chart <c:f>, table calculated-column / totals formulas.
_FORMULA_ELEM_RE = re.compile(
    r"<(?P<tag>(?:[A-Za-z0-9]+:)?(?:f|formula|formula1|formula2|calculatedColumnFormula|totalsRowFormula))"
    r"(?P<attrs>\s[^>]*)?>(?P<body>[^<]*)</(?P=tag)>"
)
_DEFINED_NAME_RE = re.compile(r"(<definedName\b[^>]*>)([^<]*)(</definedName>)")
_HYPERLINK_TAG_RE = re.compile(r"<hyperlink\b[^>]*>")
_LOCATION_ATTR_RE = re.compile(r'(\blocation=")([^"]*)(")')
_TEXTLINK_ATTR_RE = re.compile(r'(\btextlink=")([^"]*)(")')


def _map_formula_elements(xml: str, fn: Callable[[str], str]) -> str:
    def repl(m: "re.Match[str]") -> str:
        body = m.group("body")
        if not body:
            return m.group(0)
        old = xml_unescape(body)
        new = fn(old)
        if new == old:
            return m.group(0)
        return m.group(0)[: m.start("body") - m.start()] + xml_escape(new) + m.group(0)[m.end("body") - m.start() :]

    return _FORMULA_ELEM_RE.sub(repl, xml)


def _map_defined_names(xml: str, fn: Callable[[str], str]) -> str:
    def repl(m: "re.Match[str]") -> str:
        old = xml_unescape(m.group(2))
        new = fn(old)
        return m.group(0) if new == old else m.group(1) + xml_escape(new) + m.group(3)

    return _DEFINED_NAME_RE.sub(repl, xml)


def _map_attr_in_tags(xml: str, tag_re: "re.Pattern[str]", attr_re: "re.Pattern[str]", fn: Callable[[str], str]) -> str:
    def attr_repl(m: "re.Match[str]") -> str:
        old = xml_unescape(m.group(2))
        new = fn(old)
        return m.group(0) if new == old else m.group(1) + escape_attr(new) + m.group(3)

    return tag_re.sub(lambda tm: attr_re.sub(attr_repl, tm.group(0)), xml)


def _map_hyperlink_locations(xml: str, fn: Callable[[str], str]) -> str:
    return _map_attr_in_tags(xml, _HYPERLINK_TAG_RE, _LOCATION_ATTR_RE, fn)


def _map_drawing_textlinks(xml: str, fn: Callable[[str], str]) -> str:
    return _map_attr_in_tags(xml, re.compile(r"<(?:\w+:)?(?:sp|cxnSp)\b[^>]*>"), _TEXTLINK_ATTR_RE, fn)


# ---------------------------------------------------------------------------
# insert_rows -- physical row/cell/dimension shifting (target sheet only)
# ---------------------------------------------------------------------------

_ROW_FULL_RE = re.compile(
    r'<row\b(?P<pre>[^>]*)\br="(?P<r>\d+)"(?P<post>[^>]*?)(?:(?P<selfclose>/)>|>(?P<inner>.*?)</row>)',
    re.DOTALL,
)
_CELL_REF_ATTR_RE = re.compile(r'(?<=\s)r="([A-Za-z]{1,3})(\d+)"')


def _shift_rows_down(sheet_xml: str, row: int, count: int) -> str:
    """Renumber every <row r="N"> (and its cells' r="..N") with N >= row to
    N + count. Rows above the insertion point are left byte-identical."""
    sheet_xml = _normalize_self_closing(sheet_xml)

    def shift_inner_refs(inner: str, old_row: int, new_row: int) -> str:
        def cell_repl(cm: "re.Match[str]") -> str:
            col, r = cm.group(1), int(cm.group(2))
            if r == old_row:
                return f'r="{col}{new_row}"'
            return cm.group(0)

        return _CELL_REF_ATTR_RE.sub(cell_repl, inner)

    def row_repl(m: "re.Match[str]") -> str:
        old_r = int(m.group("r"))
        if old_r < row:
            return m.group(0)
        new_r = old_r + count
        if new_r > _MAX_ROW:
            raise SafeXLError(
                f"insert would push row {old_r} to row {new_r}, past Excel's row limit ({_MAX_ROW})"
            )
        pre, post = m.group("pre"), m.group("post")
        if m.group("selfclose"):
            return f'<row{pre}r="{new_r}"{post}/>'
        inner = shift_inner_refs(m.group("inner") or "", old_r, new_r)
        return f'<row{pre}r="{new_r}"{post}>{inner}</row>'

    return _ROW_FULL_RE.sub(row_repl, sheet_xml)


def _shift_dimension_for_insert(sheet_xml: str, row: int, count: int) -> str:
    m = _DIMENSION_RE.search(sheet_xml)
    if not m:
        return sheet_xml
    existing_ref = m.group(1)
    parts = existing_ref.split(":")
    try:
        if len(parts) == 1:
            sc, sr = _split_ref(parts[0])
            ec, er = sc, sr
        elif len(parts) == 2:
            sc, sr = _split_ref(parts[0])
            ec, er = _split_ref(parts[1])
        else:
            raise SafeXLError("unparseable")
    except SafeXLError:
        return sheet_xml  # existing dimension isn't parseable -- don't guess

    if sr >= row:
        sr += count
    if er >= row:
        er += count

    new_ref = f"{sc}{sr}" if (sc == ec and sr == er) else f"{sc}{sr}:{ec}{er}"
    if new_ref == existing_ref:
        return sheet_xml
    return sheet_xml[: m.start()] + f'<dimension ref="{new_ref}"/>' + sheet_xml[m.end() :]


# -- range-valued attributes and elements in the target sheet -----------------

# sqref="..." on any element except <selection> (a view detail, left alone)
_SQREF_ATTR_RE = re.compile(r'(<(?P<tag>[A-Za-z0-9:]+)\b[^>]*?\ssqref=")(?P<v>[^"]*)(")')
# ref="..." on the elements that carry a cell range in a worksheet
_REF_ATTR_RE = re.compile(r'(<(?:mergeCell|hyperlink|autoFilter|sortState|sortCondition|f)\b[^>]*?\sref=")([^"]*)(")')
_XM_SQREF_RE = re.compile(r"(<(?:\w+:)?sqref>)([^<]*)(</(?:\w+:)?sqref>)")


def _shift_target_sheet_ranges(sheet_xml: str, row: int, count: int) -> str:
    def sqref_repl(m: "re.Match[str]") -> str:
        if m.group("tag") == "selection":
            return m.group(0)
        return m.group(1) + shift_sqref(m.group("v"), row, count) + m.group(4)

    sheet_xml = _SQREF_ATTR_RE.sub(sqref_repl, sheet_xml)
    sheet_xml = _REF_ATTR_RE.sub(lambda m: m.group(1) + shift_sqref(m.group(2), row, count) + m.group(3), sheet_xml)
    sheet_xml = _XM_SQREF_RE.sub(lambda m: m.group(1) + shift_sqref(m.group(2), row, count) + m.group(3), sheet_xml)
    return sheet_xml


# -- shared formulas ----------------------------------------------------------

_CELL_RE = re.compile(r"<c\b(?P<attrs>[^>]*?)(?:/>|>(?P<inner>.*?)</c>)", re.DOTALL)
_F_ELEM_RE = re.compile(r"<f\b(?P<attrs>[^>]*?)(?:/>|>(?P<text>[^<]*)</f>)")
_CELL_ADDR_RE = re.compile(r'\sr="([A-Za-z]+)(\d+)"')
_SI_RE = re.compile(r'\ssi="(\d+)"')
_SHARED_ATTRS_RE = re.compile(r'\s(?:t|ref|si)="[^"]*"')


def _unshare_affected_formulas(
    sheet_xml: str,
    sheet_name: str,
    target: str,
    row: int,
    count: int,
    *,
    shift: Callable[[str], str] | None = None,
    force: Callable[[list[tuple[int, int, int]]], bool] | None = None,
) -> str:
    """Expand shared-formula groups that the insert would otherwise corrupt.

    A shared child has no formula text of its own; Excel derives it from the
    master by cell offset. That stops being true when (a) the insert moves
    only part of the group, or (b) shifting changes a member's references
    (a child may point above the insertion row while the master points below,
    or vice versa). Such groups are rewritten as explicit per-cell formulas,
    *un-shifted*: the ordinary formula pass then shifts them like any other
    formula. Groups the insert doesn't affect are left byte-identical.
    """
    if 't="shared"' not in sheet_xml:
        return sheet_xml

    members: dict[str, list[tuple[int, int, int]]] = {}  # si -> [(cell start, col, row)]
    masters: dict[str, tuple[int, int, str]] = {}  # si -> (col, row, formula)
    for cm in _CELL_RE.finditer(sheet_xml):
        inner = cm.group("inner")
        if not inner or "<f" not in inner:
            continue
        fm = _F_ELEM_RE.search(inner)
        if not fm or 't="shared"' not in fm.group("attrs"):
            continue
        addr = _CELL_ADDR_RE.search(cm.group("attrs"))
        si = _SI_RE.search(fm.group("attrs"))
        if not addr or not si:
            continue
        col_n = 0
        for ch in addr.group(1).upper():
            col_n = col_n * 26 + ord(ch) - 64
        members.setdefault(si.group(1), []).append((cm.start(), col_n, int(addr.group(2))))
        if fm.group("text"):
            masters[si.group(1)] = (col_n, int(addr.group(2)), xml_unescape(fm.group("text")))

    expanded: dict[int, str] = {}  # cell start -> explicit formula
    is_target = sheet_name.lower() == target.lower()
    for si, cells in members.items():
        if si not in masters:
            continue
        mcol, mrow, mtext = masters[si]
        formulas = {start: translate_formula(mtext, r - mrow, c - mcol) for start, c, r in cells}
        if force is None:
            straddles = is_target and mrow < row and any(r >= row for _, _, r in cells)
        else:
            straddles = force(cells)
        shift_one = shift or (lambda f: shift_rows_in_formula(f, sheet_name, target, row, count))
        changed = any(shift_one(f) != f for f in formulas.values())
        if straddles or changed:
            expanded.update(formulas)

    if not expanded:
        return sheet_xml

    def cell_repl(cm: "re.Match[str]") -> str:
        formula = expanded.get(cm.start())
        if formula is None:
            return cm.group(0)
        whole = cm.group(0)
        inner_start = cm.start("inner") - cm.start()
        inner = cm.group("inner")
        fm = _F_ELEM_RE.search(inner)
        attrs = _SHARED_ATTRS_RE.sub("", fm.group("attrs"))
        new_f = f"<f{attrs}>{xml_escape(formula)}</f>"
        return whole[:inner_start] + inner[: fm.start()] + new_f + inner[fm.end() :] + whole[inner_start + len(inner) :]

    return _CELL_RE.sub(cell_repl, sheet_xml)


# -- per-part transforms ------------------------------------------------------

_ANCHOR_RE = re.compile(
    r"<(?P<p>(?:[A-Za-z0-9]+:)?)(?P<kind>twoCellAnchor|oneCellAnchor)\b(?P<attrs>[^>]*)>(?P<body>.*?)</(?P=p)(?P=kind)>",
    re.DOTALL,
)
_ANCHOR_ROW_RE = re.compile(r"(<(?:[A-Za-z0-9]+:)?row>)(\d+)(</(?:[A-Za-z0-9]+:)?row>)")


def _shift_drawing_anchors(xml: str, row: int, count: int) -> str:
    def anchor_repl(m: "re.Match[str]") -> str:
        if 'editAs="absolute"' in m.group("attrs"):
            return m.group(0)

        def row_repl(rm: "re.Match[str]") -> str:
            r0 = int(rm.group(2))  # 0-based
            return f"{rm.group(1)}{r0 + count if r0 + 1 >= row else r0}{rm.group(3)}"

        body = _ANCHOR_ROW_RE.sub(row_repl, m.group("body"))
        return m.group(0)[: m.start("body") - m.start()] + body + m.group(0)[m.end("body") - m.start() :]

    return _ANCHOR_RE.sub(anchor_repl, xml)


_COMMENT_REF_RE = re.compile(r'(<(?:comment|threadedComment)\b[^>]*?\sref=")([^"]*)(")')
_VML_ROW_RE = re.compile(r"(<x:Row>\s*)(\d+)(\s*</x:Row>)")
_VML_ANCHOR_RE = re.compile(r"(<x:Anchor>)([^<]*)(</x:Anchor>)")


def _shift_vml(xml: str, row: int, count: int) -> str:
    def anchor_repl(m: "re.Match[str]") -> str:
        parts = [p.strip() for p in m.group(2).split(",")]
        if len(parts) != 8 or not all(p.lstrip("-").isdigit() for p in parts):
            return m.group(0)
        for idx in (2, 6):  # top row / bottom row of the note box, 0-based
            r0 = int(parts[idx])
            if r0 + 1 >= row:
                parts[idx] = str(r0 + count)
        return m.group(1) + ", ".join(parts) + m.group(3)

    xml = _VML_ROW_RE.sub(
        lambda m: f"{m.group(1)}{int(m.group(2)) + count if int(m.group(2)) + 1 >= row else m.group(2)}{m.group(3)}", xml
    )
    return _VML_ANCHOR_RE.sub(anchor_repl, xml)


_TABLE_REF_RE = re.compile(r'(<(?:table|autoFilter|sortState|sortCondition)\b[^>]*?\sref=")([^"]*)(")')
_PIVOT_LOCATION_RE = re.compile(r'(<location\b[^>]*?\sref=")([^"]*)(")')
_WORKSHEET_SOURCE_RE = re.compile(r"<worksheetSource\b[^>]*>")


def _shift_pivot_cache_source(xml: str, target: str, row: int, count: int) -> str:
    def tag_repl(m: "re.Match[str]") -> str:
        tag = m.group(0)
        sheet_m = re.search(r'\ssheet="([^"]*)"', tag)
        if not sheet_m or xml_unescape(sheet_m.group(1)).lower() != target.lower():
            return tag
        return re.sub(
            r'(\sref=")([^"]*)(")', lambda rm: rm.group(1) + shift_sqref(rm.group(2), row, count) + rm.group(3), tag
        )

    return _WORKSHEET_SOURCE_RE.sub(tag_repl, xml)


def _insert_in_sheet(text: str, sheet_name: str, target: str, row: int, count: int) -> str:
    is_target = sheet_name.lower() == target.lower()
    if is_target:
        _refuse_prefixed_worksheet(text, sheet_name)
    text = _unshare_affected_formulas(text, sheet_name, target, row, count)
    if is_target:
        text = _shift_rows_down(text, row, count)
        text = _shift_dimension_for_insert(text, row, count)
        text = _shift_target_sheet_ranges(text, row, count)
    text = _map_formula_elements(text, lambda f: shift_rows_in_formula(f, sheet_name, target, row, count))
    text = _map_hyperlink_locations(text, lambda f: shift_rows_in_formula(f, sheet_name, target, row, count))
    return text


def insert_rows(src: str | Path, dst: str | Path, sheet_name: str, row: int, count: int = 1) -> None:
    """Insert ``count`` blank rows at ``row`` (1-based) in ``sheet_name``.
    The old row ``row`` and everything below moves down; every reference that
    points at/below ``row`` follows (see the module docstring for the full
    list of parts covered).

    Raises ``SafeXLError`` on an invalid ``row``/``count``, an unresolvable
    sheet, or a post-write structural validation failure. ``dst`` is left
    untouched in every failure case -- see the module docstring.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    if not sheet_name or not isinstance(sheet_name, str):
        raise SafeXLError("sheet_name must be a non-empty string")
    if not isinstance(row, int) or isinstance(row, bool) or row < 1:
        raise SafeXLError(f"row must be a positive integer, got {row!r}")
    if not isinstance(count, int) or isinstance(count, bool) or count < 1:
        raise SafeXLError(f"count must be a positive integer, got {count!r}")
    if row > _MAX_ROW:
        raise SafeXLError(f"row {row} exceeds Excel's row limit ({_MAX_ROW})")

    with zipfile.ZipFile(src, "r") as zs:
        _resolve_sheet_zip_path(zs, sheet_name)  # raises if sheet_name not found
        sheet_of_part = _all_worksheet_paths(zs)
        related = _related_parts(zs, sheet_of_part)

    def build(zs: zipfile.ZipFile, zout: zipfile.ZipFile) -> None:
        for item in zs.infolist():
            name = item.filename
            if name == "xl/calcChain.xml":
                continue  # dropped unconditionally -- rows moved, any chain is potentially stale

            data = zs.read(name)
            text = None
            new = None

            if name in sheet_of_part:
                text = _dec(data)
                new = _insert_in_sheet(text, sheet_of_part[name], sheet_name, row, count)
            elif name == "xl/workbook.xml":
                text = _dec(data)
                new = _map_defined_names(text, lambda f: shift_rows_in_formula(f, None, sheet_name, row, count))
                new = _force_full_calc_on_load(new)
            elif name == "[Content_Types].xml":
                text = _dec(data)
                new = _strip_calc_chain_override(text)
            elif name == "xl/_rels/workbook.xml.rels":
                text = _dec(data)
                new = _strip_calc_chain_relationship(text)
            elif re.match(r"xl/charts/[^/]+\.xml$", name):
                text = _dec(data)
                new = _map_formula_elements(text, lambda f: shift_rows_in_formula(f, None, sheet_name, row, count))
            elif re.match(r"xl/pivotCache/pivotCacheDefinition[^/]*\.xml$", name):
                text = _dec(data)
                new = _shift_pivot_cache_source(text, sheet_name, row, count)
            elif name in related:
                kind, owner = related[name]
                text = _dec(data)
                owned = owner.lower() == sheet_name.lower()
                new = text
                if kind == "table":
                    if owned:
                        new = _TABLE_REF_RE.sub(lambda m: m.group(1) + shift_sqref(m.group(2), row, count) + m.group(3), new)
                    new = _map_formula_elements(new, lambda f: shift_rows_in_formula(f, owner, sheet_name, row, count))
                elif kind in ("comments", "threadedComment") and owned:
                    new = _COMMENT_REF_RE.sub(lambda m: m.group(1) + shift_sqref(m.group(2), row, count) + m.group(3), new)
                elif kind == "vml" and owned:
                    new = _shift_vml(new, row, count)
                elif kind == "drawing":
                    if owned:
                        new = _shift_drawing_anchors(new, row, count)
                    new = _map_drawing_textlinks(new, lambda f: shift_rows_in_formula(f, owner, sheet_name, row, count))
                elif kind == "pivotTable" and owned:
                    new = _PIVOT_LOCATION_RE.sub(lambda m: m.group(1) + shift_sqref(m.group(2), row, count) + m.group(3), new)

            if new is not None and new != text:
                data = _enc(new)
            zout.writestr(item, data)

    def check(zb: zipfile.ZipFile, za: zipfile.ZipFile) -> None:
        _readback_row_shift(zb, za, _all_worksheet_paths(zb), sheet_name, row, count)

    _atomic_write(src, dst, build, "insert_rows", check)


# ---------------------------------------------------------------------------
# rename_sheet
# ---------------------------------------------------------------------------

_SHEET_NAME_INVALID_RE = re.compile(r"[\[\]\*\?/\\:]")


def _validate_sheet_name(name: str, *, field: str) -> str:
    if not name or not isinstance(name, str):
        raise SafeXLError(f"{field} must be a non-empty string")
    if len(name) > 31:
        raise SafeXLError(f"{field} {name!r} exceeds Excel's 31-character sheet name limit")
    if _SHEET_NAME_INVALID_RE.search(name):
        raise SafeXLError(f"{field} {name!r} contains a character Excel forbids in sheet names ([ ] * ? / \\ :)")
    if name.startswith("'") or name.endswith("'"):
        raise SafeXLError(f"{field} {name!r} must not start or end with a single quote")
    return name


def _quote_sheet_name_if_needed(name: str) -> str:
    return quote_sheet_name_if_needed(name)


def _rename_sheet_in_workbook_xml(wb_xml: str, old_name: str, new_name: str) -> str:
    done = False

    def repl(m: "re.Match[str]") -> str:
        nonlocal done
        name_m = re.search(r'\bname="([^"]*)"', m.group(0))
        if done or not name_m or xml_unescape(name_m.group(1)) != old_name:
            return m.group(0)
        done = True
        tag = m.group(0)
        return tag[: name_m.start(1)] + escape_attr(new_name) + tag[name_m.end(1) :]

    new_text = re.sub(r"<sheet\b[^>]*/>", repl, wb_xml)
    if not done:
        raise SafeXLError(f"sheet {old_name!r} not found in xl/workbook.xml name attribute")
    return new_text


def _rename_sheet_in_pivot_cache(xml: str, old_name: str, new_name: str) -> str:
    def tag_repl(m: "re.Match[str]") -> str:
        tag = m.group(0)
        sheet_m = re.search(r'\ssheet="([^"]*)"', tag)
        if not sheet_m or xml_unescape(sheet_m.group(1)).lower() != old_name.lower():
            return tag
        return tag[: sheet_m.start(1)] + escape_attr(new_name) + tag[sheet_m.end(1) :]

    return _WORKSHEET_SOURCE_RE.sub(tag_repl, xml)


def rename_sheet(src: str | Path, dst: str | Path, old_name: str, new_name: str) -> None:
    """Rename worksheet ``old_name`` to ``new_name`` and rewrite every
    reference to it (see the module docstring for the list). The sheet's zip
    path is never changed.

    Raises ``SafeXLError`` if ``old_name`` doesn't exist, ``new_name`` is
    invalid or already used by another sheet, or post-write validation
    fails. ``dst`` is left untouched in every failure case.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    _validate_sheet_name(old_name, field="old_name")
    _validate_sheet_name(new_name, field="new_name")
    if old_name == new_name:
        raise SafeXLError("new_name must be different from old_name")

    with zipfile.ZipFile(src, "r") as zs:
        _resolve_sheet_zip_path(zs, old_name)  # raises SafeXLError if not found
        existing = _all_worksheet_paths(zs)
    if any(n.lower() == new_name.lower() for n in existing.values()):
        raise SafeXLError(f"a sheet named {new_name!r} already exists (sheet names are case-insensitive)")

    def rn(f: str) -> str:
        return rename_sheet_in_formula(f, old_name, new_name)

    def build(zs: zipfile.ZipFile, zout: zipfile.ZipFile) -> None:
        sheet_of_part = _all_worksheet_paths(zs)
        related = _related_parts(zs, sheet_of_part)
        for item in zs.infolist():
            name = item.filename
            data = zs.read(name)
            text = None
            new = None
            if name == "xl/workbook.xml":
                text = _dec(data)
                new = _rename_sheet_in_workbook_xml(text, old_name, new_name)
                new = _map_defined_names(new, rn)
            elif name in sheet_of_part:
                text = _dec(data)
                new = _map_formula_elements(text, rn)
                new = _map_hyperlink_locations(new, rn)
            elif re.match(r"xl/charts/[^/]+\.xml$", name):
                text = _dec(data)
                new = _map_formula_elements(text, rn)
            elif re.match(r"xl/pivotCache/pivotCacheDefinition[^/]*\.xml$", name):
                text = _dec(data)
                new = _rename_sheet_in_pivot_cache(text, old_name, new_name)
            elif name in related and related[name][0] == "table":
                text = _dec(data)
                new = _map_formula_elements(text, rn)
            elif name in related and related[name][0] == "drawing":
                text = _dec(data)
                new = _map_drawing_textlinks(text, rn)
            if new is not None and new != text:
                data = _enc(new)
            zout.writestr(item, data)

    def check(zb: zipfile.ZipFile, za: zipfile.ZipFile) -> None:
        names = _workbook_sheet_names(za)
        if new_name not in names or old_name in names:
            raise SafeXLError(f"rename_sheet read-back failed: the output's sheet list is {names!r}, expected {new_name!r} in place of {old_name!r}")
        if len(names) != len(_workbook_sheet_names(zb)):
            raise SafeXLError("rename_sheet read-back failed: the number of sheets changed")

    _atomic_write(src, dst, build, "rename_sheet", check)


# ---------------------------------------------------------------------------
# add_sheet
# ---------------------------------------------------------------------------

_SHEET_FILE_RE = re.compile(r"^xl/worksheets/sheet(\d+)\.xml$")
_SHEET_ID_RE = re.compile(r'<sheet\b[^>]*\bsheetId="(\d+)"')
_REL_ID_RE = re.compile(r'\bId="rId(\d+)"')

_MINIMAL_SHEET_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    "<sheetData/></worksheet>"
)


def _next_sheet_zip_path(names: list[str]) -> str:
    max_n = 0
    for n in names:
        m = _SHEET_FILE_RE.match(n)
        if m:
            max_n = max(max_n, int(m.group(1)))
    return f"xl/worksheets/sheet{max_n + 1}.xml"


def _next_sheet_id(wb_xml: str) -> int:
    ids = [int(x) for x in _SHEET_ID_RE.findall(wb_xml)]
    return (max(ids) + 1) if ids else 1


def _next_rel_id(rels_xml: str) -> str:
    ids = [int(x) for x in _REL_ID_RE.findall(rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def _add_sheet_to_workbook_xml(wb_xml: str, sheet_name: str, sheet_id: int, rid: str) -> str:
    if "</sheets>" not in wb_xml:
        raise SafeXLError("xl/workbook.xml has no <sheets> element to append to")
    escaped_name = escape_attr(sheet_name)
    root = re.search(r"<workbook\b[^>]*>", wb_xml)
    root_prefix = re.search(rf'xmlns:([A-Za-z_][\w.-]*)="{re.escape(_REL_NS)}"', root.group(0)) if root else None
    if root_prefix:
        prefix, decl = root_prefix.group(1), ""
    else:
        # Declared per <sheet> (openpyxl) or not at all: declare it on the new element.
        prefix, decl = "r", f' xmlns:r="{_REL_NS}"'
    new_sheet = f'<sheet name="{escaped_name}" sheetId="{sheet_id}"{decl} {prefix}:id="{rid}"/>'
    return wb_xml.replace("</sheets>", new_sheet + "</sheets>", 1)


def _add_relationship(rels_xml: str, rid: str, sheet_filename: str) -> str:
    if "</Relationships>" not in rels_xml:
        raise SafeXLError("xl/_rels/workbook.xml.rels malformed: no </Relationships>")
    rel = (
        f'<Relationship Id="{rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/{sheet_filename}"/>'
    )
    return rels_xml.replace("</Relationships>", rel + "</Relationships>", 1)


def _add_content_type_override(ct_xml: str, part_name: str) -> str:
    if "</Types>" not in ct_xml:
        raise SafeXLError("[Content_Types].xml malformed: no </Types>")
    override = (
        f'<Override PartName="{part_name}" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    )
    return ct_xml.replace("</Types>", override + "</Types>", 1)


def add_sheet(src: str | Path, dst: str | Path, sheet_name: str) -> None:
    """Add a new, empty worksheet named ``sheet_name`` to the workbook, at
    the next available ``xl/worksheets/sheetN.xml`` path, registered in
    ``xl/workbook.xml``, ``xl/_rels/workbook.xml.rels``, and
    ``[Content_Types].xml``.

    Raises ``SafeXLError`` if ``sheet_name`` is invalid, already used by
    another sheet, or post-write validation fails. ``dst`` is left
    untouched in every failure case.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    _validate_sheet_name(sheet_name, field="sheet_name")

    with zipfile.ZipFile(src, "r") as zs:
        names = zs.namelist()
        wb_xml = _dec(zs.read("xl/workbook.xml"))
        rels_xml = _dec(zs.read("xl/_rels/workbook.xml.rels"))
        ct_xml = _dec(zs.read("[Content_Types].xml"))
        existing = _all_worksheet_paths(zs)

    if any(n.lower() == sheet_name.lower() for n in existing.values()):
        raise SafeXLError(f"a sheet named {sheet_name!r} already exists (sheet names are case-insensitive)")

    new_sheet_path = _next_sheet_zip_path(names)
    new_sheet_filename = new_sheet_path.rsplit("/", 1)[-1]
    sheet_id = _next_sheet_id(wb_xml)
    rid = _next_rel_id(rels_xml)

    new_wb_xml = _add_sheet_to_workbook_xml(wb_xml, sheet_name, sheet_id, rid)
    new_rels_xml = _add_relationship(rels_xml, rid, new_sheet_filename)
    new_ct_xml = _add_content_type_override(ct_xml, f"/{new_sheet_path}")

    def build(zs: zipfile.ZipFile, zout: zipfile.ZipFile) -> None:
        for item in zs.infolist():
            data = zs.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = _enc(new_wb_xml)
            elif item.filename == "xl/_rels/workbook.xml.rels":
                data = _enc(new_rels_xml)
            elif item.filename == "[Content_Types].xml":
                data = _enc(new_ct_xml)
            zout.writestr(item, data)
        zout.writestr(new_sheet_path, _MINIMAL_SHEET_XML)

    def check(zb: zipfile.ZipFile, za: zipfile.ZipFile) -> None:
        names = _workbook_sheet_names(za)
        if sheet_name not in names or len(names) != len(_workbook_sheet_names(zb)) + 1:
            raise SafeXLError(f"add_sheet read-back failed: the output's sheet list is {names!r}, expected {sheet_name!r} added")
        if new_sheet_path not in za.namelist():
            raise SafeXLError(f"add_sheet read-back failed: part {new_sheet_path} is missing")

    _atomic_write(src, dst, build, "add_sheet", check)


# ---------------------------------------------------------------------------
# Tables: resize_table, and the table growth edit(extend_tables=True) uses
# ---------------------------------------------------------------------------

_TABLE_ROOT_RE = re.compile(r"<table\b[^>]*>")
_TABLE_COLUMNS_RE = re.compile(r"<tableColumns\b[^>]*>")
_TABLE_COLUMN_RE = re.compile(r"<tableColumn\b[^>]*?(?:/>|>.*?</tableColumn>)", re.DOTALL)
_MERGE_REF_RE = re.compile(r'<mergeCell\b[^>]*\bref="([^"]+)"')
_SST_SI_RE = re.compile(r"<si\b[^>]*?(?:/>|>(.*?)</si>)", re.DOTALL)
_T_RE = re.compile(r"<t\b[^>]*>(.*?)</t>", re.DOTALL)
_RPH_RE = re.compile(r"<rPh\b.*?</rPh>", re.DOTALL)

Bounds = tuple[int, int, int, int]  # (first col, first row, last col, last row)


def _tag_attr(tag: str, name: str) -> str | None:
    m = re.search(rf'\s{name}="([^"]*)"', tag)
    return xml_unescape(m.group(1)) if m else None


def _set_tag_ref(tag: str, new_ref: str) -> str:
    return re.sub(r'(\sref=")[^"]*(")', lambda m: m.group(1) + new_ref + m.group(2), tag, count=1)


def _parse_range(ref: str) -> Bounds:
    text = ref.replace("$", "").strip().upper()
    parts = text.split(":")
    if len(parts) > 2:
        raise SafeXLError(f"invalid range: {ref!r}")
    try:
        c1, r1 = _split_ref(parts[0])
        c2, r2 = _split_ref(parts[-1])
    except SafeXLError:
        raise SafeXLError(f"invalid range: {ref!r}") from None
    a, b, c, d = _col_to_num(c1), r1, _col_to_num(c2), r2
    return (min(a, c), min(b, d), max(a, c), max(b, d))


def _fmt_range(b: Bounds) -> str:
    return f"{_num_to_col(b[0])}{b[1]}:{_num_to_col(b[2])}{b[3]}"


def _intersects(a: Bounds, b: Bounds) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def _contains(outer: Bounds, inner: Bounds) -> bool:
    return outer[0] <= inner[0] and outer[1] <= inner[1] and inner[2] <= outer[2] and inner[3] <= outer[3]


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    xml = zf.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
    out = []
    for m in _SST_SI_RE.finditer(xml):
        body = _RPH_RE.sub("", m.group(1) or "")
        out.append(xml_unescape("".join(_T_RE.findall(body))))
    return out


def _cell_text(sheet_xml: str, ref: str, shared: list[str]) -> str | None:
    """Text of cell ``ref`` (shared / inline / cached string, or the raw number), None if empty."""
    m = re.search(rf'(<c\b[^>]*?\br="{ref}"[^>]*?)(?:/>|>(.*?)</c>)', sheet_xml, re.DOTALL)
    if not m or m.group(2) is None:
        return None
    t_m = re.search(r'\bt="(\w+)"', m.group(1))
    kind = t_m.group(1) if t_m else "n"
    body = m.group(2)
    if kind == "inlineStr":
        text = xml_unescape("".join(_T_RE.findall(_RPH_RE.sub("", body))))
    else:
        v_m = re.search(r"<v>(.*?)</v>", body, re.DOTALL)
        if not v_m:
            return None
        text = xml_unescape(v_m.group(1))
        if kind == "s":
            try:
                text = shared[int(text)]
            except (ValueError, IndexError):
                return None
    return text if text != "" else None


def _has_query_table_rel(zf: zipfile.ZipFile, part: str) -> bool:
    rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rels not in zf.namelist():
        return False
    return bool(re.search(r'Type="[^"]*/queryTable"', _dec(zf.read(rels))))


def _load_tables(zf: zipfile.ZipFile) -> list[dict]:
    sheet_of_part = _all_worksheet_paths(zf)
    related = _related_parts(zf, sheet_of_part)
    path_of_sheet = {n: p for p, n in sheet_of_part.items()}
    out = []
    for part, (kind, owner) in related.items():
        if kind != "table":
            continue
        xml = _dec(zf.read(part))
        root = _TABLE_ROOT_RE.search(xml)
        ref = _tag_attr(root.group(0), "ref") if root else None
        if not root or not ref:
            continue
        out.append(
            {
                "part": part,
                "sheet": owner,
                "sheet_path": path_of_sheet[owner],
                "name": _tag_attr(root.group(0), "name") or "",
                "display": _tag_attr(root.group(0), "displayName") or "",
                "bounds": _parse_range(ref),
                "header_rows": int(_tag_attr(root.group(0), "headerRowCount") or 1),
                "totals_rows": int(_tag_attr(root.group(0), "totalsRowCount") or 0),
                "query": _tag_attr(root.group(0), "tableType") == "queryTable"
                or _has_query_table_rel(zf, part),
                "xml": xml,
            }
        )
    return out


def _check_added_area(info: dict, nb: Bounds, others: list[dict], merges: list[Bounds]) -> None:
    for other in others:
        if other is not info and other["sheet"] == info["sheet"] and _intersects(nb, other["bounds"]):
            raise SafeXLError(
                f"table {info['name']!r} cannot grow to {_fmt_range(nb)}: overlaps table "
                f"{other['name']!r} ({_fmt_range(other['bounds'])})"
            )
    for mb in merges:
        if _intersects(nb, mb) and not _contains(info["bounds"], mb):
            raise SafeXLError(
                f"table {info['name']!r} cannot grow to {_fmt_range(nb)}: overlaps merged range {_fmt_range(mb)}"
            )


def _bracket_spans(f: str, start: int) -> str:
    """The balanced ``[...]`` text beginning at ``f[start]`` ('[')."""
    depth = 0
    for i in range(start, len(f)):
        if f[i] == "[":
            depth += 1
        elif f[i] == "]":
            depth -= 1
            if depth == 0:
                return f[start : i + 1]
    return f[start:]


def _formula_uses_column(formula: str, table_name: str, column: str) -> bool:
    col = column.lower()
    for m in re.finditer(rf"(?<![\w.]){re.escape(table_name)}\[", formula, re.IGNORECASE):
        spec = _bracket_spans(formula, m.end() - 1).lower()
        if f"[{col}]" in spec or f"[@{col}]" in spec:
            return True
    return False


def _column_users(zf: zipfile.ZipFile, info: dict, table_xml: str, column: str) -> list[str]:
    """Where a structured reference to ``info``'s ``column`` is still used."""
    tname = info["name"]
    users: list[str] = []
    sheet_of_part = _all_worksheet_paths(zf)

    def scan(xml: str, label: Callable[[str, int], str]) -> None:
        for m in _FORMULA_ELEM_RE.finditer(xml):
            body = xml_unescape(m.group("body"))
            if body and _formula_uses_column(body, tname, column):
                users.append(label(xml, m.start()))

    for part, sname in sheet_of_part.items():
        def cell_label(xml: str, pos: int, sname: str = sname) -> str:
            cm = re.search(r'r="([A-Z]+\d+)"', xml[xml.rfind("<c ", 0, pos) :]) if xml.rfind("<c ", 0, pos) >= 0 else None
            return f"formula in {sname}!{cm.group(1)}" if cm else f"formula on sheet {sname}"

        scan(_dec(zf.read(part)), cell_label)

    wb = _dec(zf.read("xl/workbook.xml"))
    for m in re.finditer(r"<definedName\b([^>]*)>([^<]*)</definedName>", wb):
        if _formula_uses_column(xml_unescape(m.group(2)), tname, column):
            users.append(f"defined name {_tag_attr('<x' + m.group(1) + '>', 'name')!r}")

    def table_label_for(t: dict, xml: str):
        def label(_x: str, pos: int) -> str:
            before = xml[:pos]
            idx = before.rfind("<tableColumn")
            cname = _tag_attr(xml[idx : xml.find(">", idx) + 1], "name") if idx >= 0 else None
            return f"calculated column {cname!r} of table {t['name']!r}"

        return label

    for t in _load_tables(zf):
        xml = table_xml if t["part"] == info["part"] else t["xml"]
        scan(xml, table_label_for(t, xml))

    for name in zf.namelist():
        if re.match(r"xl/charts/[^/]+\.xml$", name):
            scan(_dec(zf.read(name)), lambda _x, _p, n=name: f"chart part {n}")
    return users


def _build_resized_table(
    info: dict,
    nb: Bounds,
    header_text: Callable[[int, int], str | None],
    users_of: Callable[[str, str], list[str]],
    *,
    auto_header: bool = False,
    new_columns: list[dict] | None = None,
) -> str:
    """The table part's new XML for new bounds ``nb`` (right/bottom edge moves only).

    ``auto_header=True`` is Excel's behaviour for a column typed next to a
    table: an empty header cell gets the first unused ``Column{n}`` name. Every
    added column is appended to ``new_columns`` as {header_cell, name, auto_named}.
    """
    c1, r1, c2, r2 = info["bounds"]
    if (nb[0], nb[1]) != (c1, r1):
        raise SafeXLError(
            f"table {info['name']!r}: ref must keep the top-left cell {_num_to_col(c1)}{r1}; "
            "only the right and bottom edges can move"
        )
    if nb[3] - nb[1] + 1 < info["header_rows"] + info["totals_rows"] + 1:
        raise SafeXLError(f"table {info['name']!r}: {_fmt_range(nb)} leaves no data row")
    xml = info["xml"]
    if nb == info["bounds"]:
        return xml

    cols_open = _TABLE_COLUMNS_RE.search(xml)
    if cols_open:
        cols = list(_TABLE_COLUMN_RE.finditer(xml))
        old_n, new_n = c2 - c1 + 1, nb[2] - nb[0] + 1
        if len(cols) != old_n:
            raise SafeXLError(f"table {info['name']!r}: {len(cols)} tableColumn entries but ref spans {old_n} columns")
        if new_n < old_n:
            removed = cols[new_n:]
            xml = xml[: removed[0].start()] + xml[removed[-1].end() :]
            for m in removed:
                cname = _tag_attr(m.group(0), "name") or ""
                users = users_of(cname, xml)
                if users:
                    raise SafeXLError(
                        f"table {info['name']!r}: cannot remove column {cname!r}, still used by {'; '.join(users)}"
                    )
        elif new_n > old_n:
            taken = {(_tag_attr(m.group(0), "name") or "").lower() for m in cols}
            next_id = max([int(_tag_attr(m.group(0), "id") or 0) for m in cols] + [0]) + 1
            added = ""
            for k in range(old_n, new_n):
                colnum = c1 + k
                auto = False
                cell_ref = f"{_num_to_col(colnum)}{r1}" if info["header_rows"] else None
                text = header_text(colnum, r1) if info["header_rows"] else None
                if not text:
                    if info["header_rows"] and not auto_header:
                        raise SafeXLError(
                            f"table {info['name']!r}: new column {_num_to_col(colnum)} needs a header, "
                            f"but header cell {cell_ref} is empty"
                        )
                    n = 1
                    while f"column{n}" in taken:
                        n += 1
                    text, auto = f"Column{n}", True
                if text.lower() in taken:
                    raise SafeXLError(
                        f"table {info['name']!r}: header {text!r} in {_num_to_col(colnum)}{r1} "
                        "duplicates an existing column name"
                    )
                taken.add(text.lower())
                if new_columns is not None:
                    new_columns.append({"header_cell": cell_ref, "name": text, "auto_named": auto})
                added += f'<tableColumn id="{next_id}" name="{escape_attr(text)}"/>'
                next_id += 1
            end = xml.index("</tableColumns>")
            xml = xml[:end] + added + xml[end:]
        if new_n != old_n:
            xml = re.sub(
                r"(<tableColumns\b[^>]*?\bcount=\")\d+(\")",
                lambda m: m.group(1) + str(new_n) + m.group(2),
                xml,
                count=1,
            )

    root = _TABLE_ROOT_RE.search(xml)
    xml = xml[: root.start()] + _set_tag_ref(root.group(0), _fmt_range(nb)) + xml[root.end() :]
    return _adjust_filters(xml, info, nb)


def _adjust_filters(xml: str, info: dict, nb: Bounds) -> str:
    """Make the table's autoFilter and sortState follow the new range.

    A ref that covered the full old range (autoFilter: header + data; sortState:
    data rows; a sortCondition: its column over the data rows) moves to the new
    range. One that covered only part of it is left alone, except when the table
    shrank and the ref no longer fits: a stray autoFilter becomes the full new
    range, and a sortState that no longer fits is dropped (Excel only forgets the
    last sort).
    """
    c1, r1, c2, r2 = info["bounds"]
    hdr, tot = info["header_rows"], info["totals_rows"]
    old_af, new_af = (c1, r1, c2, r2 - tot), (nb[0], nb[1], nb[2], nb[3] - tot)
    old_data, new_data = (c1, r1 + hdr, c2, r2 - tot), (nb[0], nb[1] + hdr, nb[2], nb[3] - tot)

    af = re.search(r"<autoFilter\b[^>]*>", xml)
    if af:
        b = _parse_range(_tag_attr(af.group(0), "ref") or _fmt_range(old_af))
        if b == old_af or not _contains(new_af, b):
            xml = xml[: af.start()] + _set_tag_ref(af.group(0), _fmt_range(new_af)) + xml[af.end() :]

    ss = re.search(r"<sortState\b[^>]*?(?:/>|>.*?</sortState>)", xml, re.DOTALL)
    if ss:
        block = ss.group(0)
        root = re.match(r"<sortState\b[^>]*>", block).group(0)
        sb = _parse_range(_tag_attr(root, "ref") or _fmt_range(old_data))
        if sb == old_data:
            block = block.replace(root, _set_tag_ref(root, _fmt_range(new_data)), 1)

            def cond(m: "re.Match[str]") -> str:
                cb = _parse_range(_tag_attr(m.group(0), "ref") or "A1")
                if cb[1] == old_data[1] and cb[3] == old_data[3]:
                    return _set_tag_ref(m.group(0), _fmt_range((cb[0], new_data[1], cb[2], new_data[3])))
                return m.group(0)

            block = re.sub(r"<sortCondition\b[^>]*>", cond, block)
        shrunk = nb[2] < c2 or nb[3] < r2
        fits = all(_contains(new_data, _parse_range(r)) for r in re.findall(r'\sref="([^"]*)"', block))
        xml = xml[: ss.start()] + ("" if shrunk and not fits else block) + xml[ss.end() :]
    return xml


def _merges_of(sheet_xml: str) -> list[Bounds]:
    return [_parse_range(m) for m in _MERGE_REF_RE.findall(sheet_xml)]


def resize_table(
    src: str | Path,
    dst: str | Path,
    *,
    table: str,
    ref: str,
    sheet: str | None = None,
) -> None:
    """Set the range of an Excel table to ``ref`` (e.g. ``"A1:D12"``).

    ``table`` is the table's name or displayName; ``sheet`` optionally limits
    the lookup to one sheet. Only the table part changes (``ref``, its
    ``autoFilter`` ref, and ``tableColumns`` count and entries); every other
    part stays byte-identical. The top-left cell must stay put; the right and
    bottom edges move. A new column takes its name from the sheet's header
    cell (it must be non-empty and unique within the table). Removing a
    column still used by a structured reference, a calculated-column formula
    or a defined name is refused with the user named. Growing into another
    table or a merged range is refused.

    Raises ``SafeXLError`` on any refusal; ``dst`` is left untouched.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    nb = _parse_range(ref)

    with zipfile.ZipFile(src, "r") as zs:
        tables = _load_tables(zs)
        key = table.lower()
        hits = [
            t
            for t in tables
            if key in (t["name"].lower(), t["display"].lower()) and (sheet is None or t["sheet"].lower() == sheet.lower())
        ]
        if not hits:
            raise SafeXLError(f"table {table!r} not found" + (f" on sheet {sheet!r}" if sheet else ""))
        info = hits[0]
        sheet_xml = _dec(zs.read(info["sheet_path"]))
        _refuse_prefixed_worksheet(sheet_xml, info["sheet"])
        shared = _load_shared_strings(zs)
        _check_added_area(info, nb, tables, _merges_of(sheet_xml))
        new_xml = _build_resized_table(
            info,
            nb,
            lambda col, row: _cell_text(sheet_xml, f"{_num_to_col(col)}{row}", shared),
            lambda column, remaining: _column_users(zs, info, remaining, column),
        )

    def build(zs: zipfile.ZipFile, zout: zipfile.ZipFile) -> None:
        for item in zs.infolist():
            data = zs.read(item.filename)
            if item.filename == info["part"] and new_xml != info["xml"]:
                data = _enc(new_xml)
            zout.writestr(item, data)

    def check(zb: zipfile.ZipFile, za: zipfile.ZipFile) -> None:
        root = ET.fromstring(za.read(info["part"]))
        got = (root.get("ref") or "").replace("$", "").upper()
        if _parse_range(got or "A1") != nb or not got:
            raise SafeXLError(f"resize_table read-back failed: table {table!r} ref reads back as {got!r}, expected {_fmt_range(nb)}")

    _atomic_write(src, dst, build, "resize_table", check)


def _plan_table_extensions(
    src: Path,
    grouped: dict[str, dict[str, dict]],
    sheet_targets: dict[str, str],
) -> list[dict]:
    """The tables ``edit(extend_tables=True)`` will actually change (query-table
    skips are left out; ``plan_table_extensions`` reports them)."""
    return [p for p in _plan_all_table_extensions(src, grouped, sheet_targets) if "skipped_reason" not in p]


def _plan_all_table_extensions(
    src: Path,
    grouped: dict[str, dict[str, dict]],
    sheet_targets: dict[str, str],
) -> list[dict]:
    """For ``edit(extend_tables=True)``: one entry per table a written cell sits
    directly next to (right column within its rows, or the row below within its
    columns; no row growth when a totals row is present). Each entry has
    ``part``/``xml`` (the new table part), ``grouped_sheet`` (the key in
    ``grouped``), and the public fields of ``plan_table_extensions``. A new
    column with an empty header cell is auto-named like Excel does."""
    out: list[dict] = []
    with zipfile.ZipFile(src, "r") as zs:
        tables = _load_tables(zs)
        shared = _load_shared_strings(zs)
        post: dict[str, str] = {}
        for sname, target in sheet_targets.items():
            text = _dec(zs.read(target))
            post[target] = _inject_cells(text, grouped[sname])[0]
        written: dict[str, list[tuple[int, int]]] = {}
        key_of: dict[str, str] = {}
        for sname, cells in grouped.items():
            written[sheet_targets[sname]] = [(_col_to_num(_split_ref(r)[0]), _split_ref(r)[1]) for r in cells]
            key_of[sheet_targets[sname]] = sname

        for info in tables:
            cells = written.get(info["sheet_path"])
            if not cells:
                continue
            c1, r1, c2, r2 = info["bounds"]
            grew = True
            while grew:
                grew = False
                for col, row in cells:
                    if col == c2 + 1 and r1 <= row <= r2:
                        c2, grew = c2 + 1, True
                    elif row == r2 + 1 and not info["totals_rows"] and c1 <= col <= c2:
                        r2, grew = r2 + 1, True
            nb = (c1, r1, c2, r2)
            if nb == info["bounds"]:
                continue
            if info["query"]:
                out.append(
                    {
                        "table": info["name"],
                        "sheet": info["sheet"],
                        "old_ref": _fmt_range(info["bounds"]),
                        "new_ref": None,
                        "new_columns": [],
                        "skipped_reason": (
                            "backed by a query table (external data); a column added beside it is not part of "
                            "the query, so the table was left unchanged"
                        ),
                    }
                )
                continue
            sheet_xml = post[info["sheet_path"]]
            _check_added_area(info, nb, tables, _merges_of(sheet_xml))
            new_columns: list[dict] = []
            xml = _build_resized_table(
                info,
                nb,
                lambda col, row, sx=sheet_xml: _cell_text(sx, f"{_num_to_col(col)}{row}", shared),
                lambda column, remaining: [],
                auto_header=True,
                new_columns=new_columns,
            )
            out.append(
                {
                    "table": info["name"],
                    "sheet": info["sheet"],
                    "old_ref": _fmt_range(info["bounds"]),
                    "new_ref": _fmt_range(nb),
                    "new_columns": new_columns,
                    "part": info["part"],
                    "xml": xml,
                    "grouped_sheet": key_of[info["sheet_path"]],
                }
            )
    return out


def plan_table_extensions(src: str | Path, changes: Sequence[Change]) -> list[dict]:
    """What ``edit(src, dst, changes, extend_tables=True)`` would do to tables.

    Read-only: nothing is written. Returns one dict per table that would grow:
    ``{"table", "sheet", "old_ref", "new_ref", "new_columns": [{"header_cell",
    "name", "auto_named"}]}``. A table backed by a query table is never extended:
    its entry has ``new_ref`` None and a ``skipped_reason``. ``auto_named`` is True when the header cell was
    empty and the column gets Excel's ``Column{n}`` name (``header_cell`` is None
    for a table with no header row). Raises ``SafeXLError`` for the same
    refusals ``edit`` would hit.
    """
    src = Path(src)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    grouped = _group_changes(changes)
    with zipfile.ZipFile(src, "r") as zs:
        targets = {name: _resolve_sheet_zip_path(zs, name) for name in grouped}
    plans = _plan_all_table_extensions(src, grouped, targets)
    return [
        {k: p[k] for k in ("table", "sheet", "old_ref", "new_ref", "new_columns", "skipped_reason") if k in p}
        for p in plans
    ]


# ---------------------------------------------------------------------------
# insert_cells -- Excel's "Insert cells, shift cells down" within a column span
# ---------------------------------------------------------------------------
#
# insert_cells("B5:C6") makes two blank rows in columns B:C only: cells of B:C
# at row >= 5 move down 2; every other column is left where it is. References
# follow the cells (see ``refs.shift_cells_in_formula``):
#   * wholly inside B:C: a ref at/below row 5 moves; a range that straddles
#     row 5 grows; one ending above row 5 does not;
#   * wholly outside B:C, or ending above row 5: untouched;
#   * a range that straddles the edge of B:C with rows at/below row 5 is torn
#     (some cells move, some stay). Excel's formula refs would stay unchanged
#     and silently change meaning, so this refuses with the reference named.
#     The exception is conditional-format / data-validation sqrefs, which belong
#     to cells and are split. Merges, array formulas, hyperlinks, comments and
#     pivot locations that contain row 5 inside the span cannot grow and are
#     refused; tables, autoFilters, sortStates and pivot sources grow.
# Not followed (left as they were): drawing/chart anchors, sheet-view selection.

_SQREF_ATTR_NOSEL_RE = _SQREF_ATTR_RE
_GROW_TAG_RE = re.compile(r'(<(?P<tag>mergeCell|hyperlink|autoFilter|sortState|sortCondition)\b[^>]*?\sref=")(?P<v>[^"]*)(")')
_F_TAG_RE = re.compile(r"<f\b[^>]*>")
_VML_SHAPE_RE = re.compile(r"<v:shape\b.*?</v:shape>", re.DOTALL)
_VML_COL_RE = re.compile(r"<x:Column>\s*(\d+)\s*</x:Column>")


def _move_cells_down(sheet_xml: str, c1: int, c2: int, row: int, count: int) -> tuple[str, int]:
    """Move cells in columns c1..c2 at row >= ``row`` down by ``count``.
    Returns (xml, highest row now holding a moved cell, 0 if nothing moved).
    Rows that don't change stay byte-identical."""
    sheet_xml = _normalize_self_closing(sheet_xml)
    sd = re.search(r"<sheetData>(.*)</sheetData>", sheet_xml, re.DOTALL)
    if not sd:
        return sheet_xml, 0

    rows: dict[int, dict] = {}
    for m in _ROW_FULL_RE.finditer(sd.group(1)):
        inner = m.group("inner") or ""
        cells = [(_col_to_num(_CELL_ADDR_RE.search(cm.group("attrs")).group(1).upper()), cm.group(0)) for cm in _CELL_RE.finditer(inner)]  # type: ignore[union-attr]
        if "".join(c for _, c in cells) != inner:
            raise SafeXLError(f"row {m.group('r')} holds content other than cells; insert_cells cannot move it safely")
        rows[int(m.group("r"))] = {"m": m, "cells": cells, "dirty": False}

    top = 0
    for rnum in sorted(rows, reverse=True):
        if rnum < row:
            continue
        info = rows[rnum]
        keep = []
        for col, xml in info["cells"]:
            if c1 <= col <= c2:
                new_r = rnum + count
                if new_r > _MAX_ROW:
                    raise SafeXLError(f"insert would push row {rnum} to row {new_r}, past Excel's row limit ({_MAX_ROW})")
                xml = re.sub(r'(\sr=")([A-Za-z]+)\d+(")', lambda m, n=new_r: f"{m.group(1)}{m.group(2)}{n}{m.group(3)}", xml, count=1)
                tgt = rows.setdefault(new_r, {"m": None, "cells": [], "dirty": True})
                tgt["cells"].append((col, xml))
                tgt["dirty"] = True
                top = max(top, new_r)
                info["dirty"] = True
            else:
                keep.append((col, xml))
        info["cells"] = keep
    if not top:
        return sheet_xml, 0

    out = []
    for rnum in sorted(rows):
        info = rows[rnum]
        if not info["dirty"]:
            out.append(info["m"].group(0))
            continue
        cells = "".join(x for _, x in sorted(info["cells"], key=lambda t: t[0]))
        m = info["m"]
        if m is None:
            out.append(f'<row r="{rnum}">{cells}</row>')
        else:
            pre, post = re.sub(r'\sspans="[^"]*"', "", m.group("pre")), re.sub(r'\sspans="[^"]*"', "", m.group("post"))
            out.append(f'<row{pre}r="{rnum}"{post}>{cells}</row>')
    body = "".join(out)
    return sheet_xml[: sd.start(1)] + body + sheet_xml[sd.end(1) :], top


def _widen_dimension(sheet_xml: str, bottom: int) -> str:
    m = _DIMENSION_RE.search(sheet_xml)
    if not m or not bottom:
        return sheet_xml
    parts = m.group(1).split(":")
    try:
        sc, sr = _split_ref(parts[0])
        ec, er = _split_ref(parts[-1])
    except SafeXLError:
        return sheet_xml
    if bottom <= er:
        return sheet_xml
    return sheet_xml[: m.start()] + f'<dimension ref="{sc}{sr}:{ec}{bottom}"/>' + sheet_xml[m.end() :]


def _shift_vml_cols(xml: str, c1: int, c2: int, row: int, count: int) -> str:
    def shape(m: "re.Match[str]") -> str:
        blk = m.group(0)
        rm, cm = _VML_ROW_RE.search(blk), _VML_COL_RE.search(blk)
        if not rm or not cm:
            return blk
        r0, col = int(rm.group(2)), int(cm.group(1)) + 1
        if not (c1 <= col <= c2 and r0 + 1 >= row):
            return blk
        blk = _VML_ROW_RE.sub(lambda x: f"{x.group(1)}{r0 + count}{x.group(3)}", blk)

        def anchor(a: "re.Match[str]") -> str:
            parts = [p.strip() for p in a.group(2).split(",")]
            if len(parts) != 8 or not all(p.lstrip("-").isdigit() for p in parts):
                return a.group(0)
            for idx in (2, 6):
                parts[idx] = str(int(parts[idx]) + count)
            return a.group(1) + ", ".join(parts) + a.group(3)

        return _VML_ANCHOR_RE.sub(anchor, blk)

    return _VML_SHAPE_RE.sub(shape, xml)


def _cells_in_sheet(text: str, sheet_name: str, target: str, c1: int, c2: int, row: int, count: int) -> str:
    def fshift(f: str) -> str:
        return shift_cells_in_formula(f, sheet_name, target, c1, c2, row, count)

    is_target = sheet_name.lower() == target.lower()
    if is_target:
        _refuse_prefixed_worksheet(text, sheet_name)
    text = _unshare_affected_formulas(
        text,
        sheet_name,
        target,
        row,
        count,
        shift=fshift,
        force=lambda cells: is_target and any(c1 <= c <= c2 and r >= row for _, c, r in cells),
    )
    if is_target:
        text, bottom = _move_cells_down(text, c1, c2, row, count)
        text = _widen_dimension(text, bottom)

        def sq(m: "re.Match[str]") -> str:
            if m.group("tag") == "selection":
                return m.group(0)
            return m.group(1) + shift_sqref_cols(m.group("v"), c1, c2, row, count) + m.group(4)

        text = _SQREF_ATTR_RE.sub(sq, text)
        text = _XM_SQREF_RE.sub(lambda m: m.group(1) + shift_sqref_cols(m.group(2), c1, c2, row, count) + m.group(3), text)

        def grow(m: "re.Match[str]") -> str:
            allow = m.group("tag") in ("autoFilter", "sortState", "sortCondition")
            return m.group(1) + shift_ref_cols(m.group("v"), c1, c2, row, count, allow_grow=allow) + m.group(4)

        text = _GROW_TAG_RE.sub(grow, text)

        def array_f(m: "re.Match[str]") -> str:
            tag = m.group(0)
            if 't="array"' not in tag:
                return tag
            return re.sub(
                r'(\sref=")([^"]*)(")',
                lambda r: r.group(1) + shift_ref_cols(r.group(2), c1, c2, row, count, allow_grow=False) + r.group(3),
                tag,
            )

        text = _F_TAG_RE.sub(array_f, text)
    text = _map_formula_elements(text, fshift)
    text = _map_hyperlink_locations(text, fshift)
    return text


def insert_cells(src: str | Path, dst: str | Path, *, sheet: str, cell_range: str, shift: str = "down") -> None:
    """Insert blank cells over ``cell_range`` (e.g. ``"B5:C6"``) in ``sheet`` and
    shift the cells of those columns at/below the range's first row down by the
    range's height, like Excel's "Insert cells, shift cells down". Columns
    outside the range do not move. Every reference follows as described above
    (formulas, defined names, DV/CF, charts, tables, merges, hyperlinks,
    comments, pivots).

    Only ``shift="down"`` is supported. Raises ``SafeXLError`` for a bad
    range/sheet, ``shift="right"``, or a reference/merge/array formula that
    would be torn (the message names it); ``dst`` is left untouched.
    """
    src = Path(src)
    dst = Path(dst)
    if shift == "right":
        raise SafeXLError("insert_cells: shift='right' is not supported (only shift='down')")
    if shift != "down":
        raise SafeXLError(f"insert_cells: shift must be 'down', got {shift!r}")
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    if not sheet or not isinstance(sheet, str):
        raise SafeXLError("sheet must be a non-empty string")
    try:
        if not isinstance(cell_range, str) or not cell_range.strip():
            raise SafeXLError("empty")
        bounds = _parse_range(cell_range)
    except SafeXLError:
        raise SafeXLError(f"invalid cell_range: {cell_range!r}") from None
    c1, row, c2, last = bounds
    count = last - row + 1
    if c2 > MAX_COL or last > _MAX_ROW:
        raise SafeXLError(f"cell_range {cell_range!r} is outside the worksheet")

    with zipfile.ZipFile(src, "r") as zs:
        _resolve_sheet_zip_path(zs, sheet)  # raises if the sheet is missing
        sheet_of_part = _all_worksheet_paths(zs)
        related = _related_parts(zs, sheet_of_part)

    def refs_to(owner: str | None):
        return lambda f: shift_cells_in_formula(f, owner, sheet, c1, c2, row, count)

    def build(zs: zipfile.ZipFile, zout: zipfile.ZipFile) -> None:
        for item in zs.infolist():
            name = item.filename
            if name == "xl/calcChain.xml":
                continue  # cells moved: any chain is potentially stale
            data = zs.read(name)
            text = None
            new = None
            try:
                if name in sheet_of_part:
                    text = _dec(data)
                    new = _cells_in_sheet(text, sheet_of_part[name], sheet, c1, c2, row, count)
                elif name == "xl/workbook.xml":
                    text = _dec(data)
                    new = _force_full_calc_on_load(_map_defined_names(text, refs_to(None)))
                elif name == "[Content_Types].xml":
                    text = _dec(data)
                    new = _strip_calc_chain_override(text)
                elif name == "xl/_rels/workbook.xml.rels":
                    text = _dec(data)
                    new = _strip_calc_chain_relationship(text)
                elif re.match(r"xl/charts/[^/]+\.xml$", name):
                    text = _dec(data)
                    new = _map_formula_elements(text, refs_to(None))
                elif re.match(r"xl/pivotCache/pivotCacheDefinition[^/]*\.xml$", name):
                    text = _dec(data)

                    def src_tag(m: "re.Match[str]") -> str:
                        tag = m.group(0)
                        sm = re.search(r'\ssheet="([^"]*)"', tag)
                        if not sm or xml_unescape(sm.group(1)).lower() != sheet.lower():
                            return tag
                        return re.sub(
                            r'(\sref=")([^"]*)(")',
                            lambda r: r.group(1) + shift_ref_cols(r.group(2), c1, c2, row, count, allow_grow=True) + r.group(3),
                            tag,
                        )

                    new = _WORKSHEET_SOURCE_RE.sub(src_tag, text)
                elif name in related:
                    kind, owner = related[name]
                    text = _dec(data)
                    owned = owner.lower() == sheet.lower()
                    new = text

                    def strict(m: "re.Match[str]") -> str:
                        return m.group(1) + shift_ref_cols(m.group(2), c1, c2, row, count, allow_grow=False) + m.group(3)

                    if kind == "table":
                        if owned:
                            new = _TABLE_REF_RE.sub(
                                lambda m: m.group(1) + shift_ref_cols(m.group(2), c1, c2, row, count, allow_grow=True) + m.group(3),
                                new,
                            )
                        new = _map_formula_elements(new, refs_to(owner))
                    elif kind in ("comments", "threadedComment") and owned:
                        new = _COMMENT_REF_RE.sub(strict, new)
                    elif kind == "vml" and owned:
                        new = _shift_vml_cols(new, c1, c2, row, count)
                    elif kind == "drawing":
                        new = _map_drawing_textlinks(new, refs_to(owner))
                    elif kind == "pivotTable" and owned:
                        new = _PIVOT_LOCATION_RE.sub(strict, new)
            except TornReference as exc:
                raise SafeXLError(f"insert_cells refused: {exc}; in part {name}") from None
            if new is not None and new != text:
                data = _enc(new)
            zout.writestr(item, data)

    def check(zb: zipfile.ZipFile, za: zipfile.ZipFile) -> None:
        _readback_row_shift(zb, za, _all_worksheet_paths(zb), sheet, row, count, c1, c2, "insert_cells")

    _atomic_write(src, dst, build, "insert_cells", check)
