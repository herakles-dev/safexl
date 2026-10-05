"""SafeXL engine core: edit .xlsx/.xlsm workbooks in headless environments
by changing only the XML an edit needs, so Excel (or LibreOffice) can open them.

TECHNIQUE
---------
Direct XML injection / "zip surgery": every zip member of the source OOXML
package is copied BYTE-FOR-BYTE into the output, except the worksheet part(s)
actually being edited (and, incidentally, ``xl/workbook.xml``,
``[Content_Types].xml`` and ``xl/calcChain.xml`` -- see "What always changes"
below). Cell writes are surgical regex insert/replace operations against
``<sheetData>``, not a parse/rebuild round trip.

This is why SafeXL preserves OOXML parts that a load/rebuild/save engine
(openpyxl, most Excel-editing MCP servers) cannot: charts, drawings, media,
pivot tables/caches, data validations, conditional formatting (including
extLst/x14 blocks), and VBA (``xl/vbaProject.bin``) are never touched, never
re-serialized, and are therefore byte-identical to the source. SafeXL NEVER
executes or rewrites VBA -- macro-enabled workbooks (.xlsm) are edited the
same way as .xlsx, with ``xl/vbaProject.bin`` copied byte-for-byte untouched.

This module is the generalization of a prototype that was validated on
real workbooks and then audited adversarially: 7/7 real-world OOXML fixtures
survived with every feature category preserved, all 7 outputs opened cleanly
in LibreOffice headless with the injected formula computing correctly, and a
from-scratch structural audit found zero CRITICAL findings. That audit's one
required fix -- stale ``<dimension ref>`` after injecting a cell beyond the
sheet's original range -- is implemented here (see ``_update_dimension``).

FORMULA CACHING CAVEAT (read this before relying on a formula result)
-----------------------------------------------------------------------
When a change writes a formula, SafeXL emits a bare ``<f>formula</f>`` element
with **no cached ``<v>``** -- it does not evaluate the formula itself (it has
no spreadsheet engine; doing so correctly, including cross-sheet references,
array formulas, and volatile functions, is out of scope for a zip-surgery
writer). Instead it sets ``<calcPr fullCalcOnLoad="1"/>`` on
``xl/workbook.xml``, which tells Excel/LibreOffice/any other calculating
application to recompute every formula the moment the file is opened.

Practical implication, confirmed in the adversarial verification: **a
downstream consumer that reads a SafeXL-edited file with a library (openpyxl,
pandas, etc.) and never opens it in a calculating application first will see
the formula cell's value as blank/None**, not the computed result. If a
caller needs the cached value immediately without a human or agent opening
the file first, it must shell out to a recalculating application itself (e.g.
``soffice --headless --convert-to xlsx``) as a post-write step -- SafeXL does
not do this automatically, to keep the core engine dependency-free and fast.

What always changes (and why byte-identity still mostly holds)
----------------------------------------------------------------
- The edited worksheet part(s) -- obviously; that is the edit.
- ``xl/workbook.xml`` -- only if ``fullCalcOnLoad="1"`` was not already set
  (if it was, the bytes read from the source are written back unchanged and
  the part comes out byte-identical automatically).
- ``[Content_Types].xml`` -- only if it declared an Override for
  ``xl/calcChain.xml`` that must now be stripped (same "no-op if not present"
  property as above).
- ``xl/_rels/workbook.xml.rels`` -- only if it has a Relationship to
  ``calcChain.xml``, which is removed along with the part (a relationship to a
  missing part is itself a repair trigger).
- ``xl/calcChain.xml`` -- always dropped if present. A calc chain that is
  stale after *any* edit (not just a formula edit -- e.g. editing a plain
  value that an untouched formula elsewhere depends on) is a documented
  repair trigger; dropping the whole part and relying on
  ``fullCalcOnLoad="1"`` to force a full rebuild is the safe route explicitly
  validated by an adversarial verification (WARNING-only,
  zero CRITICAL findings across 7/7 fixtures). This is a superset of "remove
  just the edited formula cells' calc-chain entries" and was chosen because
  it is simpler and already proven safe, at the cost of calcChain.xml never
  being byte-identical across any edit that touches at least one cell.

Every other part -- xl/charts/*, xl/drawings/*, xl/media/*, xl/pivotTables/*,
xl/pivotCache/*, xl/vbaProject.bin, dataValidations, conditionalFormatting,
extLst blocks, printerSettings, ctrlProps, custom XML, sharedStrings, styles,
theme, etc. -- is untouched and byte-identical to the source.

ROLLBACK PROMISE
-----------------
``edit()`` never writes directly to ``dst``. It writes to a private temp file,
runs the structural validator (``safexl_mcp.validate.validate_workbook``)
against that temp file, and only moves it into place at ``dst`` if the
validator reports zero CRITICAL findings. On any failure -- a bad change
spec, an unresolvable sheet/cell, a write error, or a validator CRITICAL
finding -- ``dst`` is left completely untouched (never created, never
partially written) and a ``SafeXLError`` is raised describing exactly what
went wrong.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Mapping, Sequence
from xml.sax.saxutils import escape as _xml_escape

from .validate import validate_workbook

__all__ = ["edit", "SafeXLError", "Change"]

Change = Mapping[str, Any]

_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_ROW_TAG_RE = re.compile(r'<row\b[^>]*\br="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.DOTALL)
_SELF_CLOSING_ROW_RE = re.compile(r"<row((?:\s+[^>]*)?)/>")
_DIMENSION_RE = re.compile(r'<dimension\s+ref="([^"]*)"\s*/>')


class SafeXLError(Exception):
    """Raised when a change set cannot be safely applied.

    Whenever this is raised by ``edit()``, ``dst`` has NOT been written or
    modified -- the source file and any pre-existing dst are exactly as they
    were before the call.
    """


# ---------------------------------------------------------------------------
# Decoy-proof text handling (S3)
# ---------------------------------------------------------------------------
#
# Every locator in this package finds cells, rows, <sheetData>, merges and
# tables by regex on raw XML text. A comment, CDATA section or processing
# instruction can contain text that looks like any of those. MECHANISM: every
# part is decoded through ``_dec``, which replaces each comment / CDATA / PI
# span with an opaque placeholder ``\x00<hex of the span's UTF-8 bytes>\x00``
# (NUL is illegal in XML, so it never occurs in a real part, and the
# placeholder holds no ``<``, ``>``, quote or letter a locator could match).
# Regex edits then run on the masked text, and ``_enc`` turns placeholders back
# into the original bytes, so spans outside an overwritten cell are preserved
# byte for byte. The placeholder carries its own content, so no state is shared.

_PROTECTED_SPAN_RE = re.compile(r"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>", re.DOTALL)
_PLACEHOLDER_RE = re.compile("\x00([0-9a-f]*)\x00")


def _has_spans(xml: str) -> bool:
    """Cheap C-speed pre-check, so a part with no comment, CDATA or PI (the usual case) is one
    substring scan instead of a regex attempt at every '<'. The XML declaration is not a span."""
    start = xml.find("?>") + 2 if xml.startswith("<?xml") else 0
    return "<!--" in xml or "<![CDATA[" in xml or xml.find("<?", start) != -1


def mask_spans_same_length(xml: str) -> str:
    """Comments, CDATA and PIs overwritten with spaces, so offsets in the result are offsets in ``xml``."""
    if not _has_spans(xml):
        return xml
    return _PROTECTED_SPAN_RE.sub(lambda m: " " * len(m.group(0)), xml)


def protect_spans(xml: str) -> str:
    if not _has_spans(xml):
        return xml
    return _PROTECTED_SPAN_RE.sub(lambda m: "\x00" + m.group(0).encode("utf-8").hex() + "\x00", xml)


def restore_spans(xml: str) -> str:
    if "\x00" not in xml:
        return xml
    return _PLACEHOLDER_RE.sub(lambda m: bytes.fromhex(m.group(1)).decode("utf-8"), xml)


def _dec(data: bytes) -> str:
    """Decode a part and mask its comments, CDATA sections and PIs."""
    return protect_spans(data.decode("utf-8", errors="replace"))


def _enc(text: str) -> bytes:
    """Inverse of ``_dec``: restore the masked spans and encode."""
    return restore_spans(text).encode("utf-8")


_PREFIXED_SHEET_RE = re.compile(r"<[A-Za-z_][\w.-]*:(?:worksheet|sheetData)\b")
_PLAIN_SHEETDATA_RE = re.compile(r"<sheetData\b")


def _refuse_prefixed_worksheet(sheet_xml: str, where: str = "") -> None:
    """The regex locators assume unprefixed elements. A worksheet written with a
    namespace prefix on every element (the .NET OpenXML SDK does this) would
    otherwise fail deep inside, or silently match nothing."""
    if _PREFIXED_SHEET_RE.search(sheet_xml) and not _PLAIN_SHEETDATA_RE.search(sheet_xml):
        raise SafeXLError(
            "worksheet uses a namespace prefix (written by the OpenXML SDK); not supported yet"
            + (f" [{where}]" if where else "")
        )


# ---------------------------------------------------------------------------
# Cell-reference helpers
# ---------------------------------------------------------------------------


def _split_ref(ref: str) -> tuple[str, int]:
    m = _CELL_REF_RE.match(ref.strip())
    if not m:
        raise SafeXLError(f"invalid cell reference: {ref!r}")
    return m.group(1).upper(), int(m.group(2))


def _col_to_num(col_letters: str) -> int:
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col(n: int) -> str:
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _col_sort_key(ref: str) -> int:
    letters, _ = _split_ref(ref)
    return _col_to_num(letters)


def _row_num(ref: str) -> int:
    _, row = _split_ref(ref)
    return row


def _normalize_ref(ref: str) -> str:
    letters, row = _split_ref(ref)
    return f"{letters}{row}"


def _normalize_formula(formula: str) -> str:
    """Strip a single leading '=' if present. The <f> element must never
    contain a leading '=' -- Excel prepends that only for display in the
    formula bar; storing it is what causes negokaz/excel-mcp-server's
    documented formula-corruption bug."""
    f = formula.strip()
    if f.startswith("="):
        f = f[1:]
    return f.strip()


# ---------------------------------------------------------------------------
# Change validation
# ---------------------------------------------------------------------------


def _validate_change(change: Change, index: int) -> tuple[str, str, Change]:
    if not isinstance(change, Mapping):
        raise SafeXLError(f"changes[{index}]: expected a dict/mapping, got {type(change).__name__}")
    sheet = change.get("sheet")
    cell = change.get("cell")
    if not sheet or not isinstance(sheet, str):
        raise SafeXLError(f"changes[{index}]: missing/invalid 'sheet'")
    if not cell or not isinstance(cell, str):
        raise SafeXLError(f"changes[{index}]: missing/invalid 'cell'")
    _split_ref(cell)  # raises SafeXLError if malformed

    has_value = "value" in change
    has_formula = "formula" in change
    if has_value and has_formula:
        raise SafeXLError(f"changes[{index}] ({sheet}!{cell}): specify 'value' or 'formula', not both")
    if not has_value and not has_formula:
        raise SafeXLError(f"changes[{index}] ({sheet}!{cell}): must specify 'value' or 'formula'")

    if has_value:
        value = change["value"]
        if not isinstance(value, (bool, int, float, str)):
            raise SafeXLError(
                f"changes[{index}] ({sheet}!{cell}): unsupported value type "
                f"{type(value).__name__} (supported: bool, int, float, str)"
            )
    else:
        formula = change["formula"]
        if not isinstance(formula, str) or not formula.strip():
            raise SafeXLError(f"changes[{index}] ({sheet}!{cell}): 'formula' must be a non-empty string")

    return sheet, _normalize_ref(cell), change


def _group_changes(changes: Sequence[Change]) -> dict[str, dict[str, Change]]:
    """Validate and group changes by sheet -> {normalized cell ref: change}.
    A later change for the same (sheet, cell) overwrites an earlier one,
    same as assigning into a dict twice -- deterministic, last write wins."""
    grouped: dict[str, dict[str, Change]] = {}
    for i, change in enumerate(changes):
        sheet, ref, validated = _validate_change(change, i)
        grouped.setdefault(sheet, {})[ref] = validated
    return grouped


# ---------------------------------------------------------------------------
# Cell XML construction
# ---------------------------------------------------------------------------


def _make_cell_xml(ref: str, change: Change, style: str = "") -> str:
    s_attr = f' s="{style}"' if style else ""
    if "formula" in change:
        formula = _xml_escape(_normalize_formula(change["formula"]))
        return f'<c r="{ref}"{s_attr}><f>{formula}</f></c>'

    value = change["value"]
    if isinstance(value, bool):
        # Must be checked before (int, float): bool is an int subclass.
        return f'<c r="{ref}"{s_attr} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{value}</v></c>'
    escaped = _xml_escape(str(value))
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{escaped}</t></is></c>'


# ---------------------------------------------------------------------------
# Sheet resolution
# ---------------------------------------------------------------------------


_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _rel_prefix_alt(xml: str) -> str:
    """Regex alternation of every prefix bound to the relationships namespace
    anywhere in ``xml`` (openpyxl declares it on each <sheet>, others on the
    root, a few under a prefix other than ``r``). Falls back to ``r``."""
    prefixes = re.findall(rf'xmlns:([A-Za-z_][\w.-]*)="{re.escape(_REL_NS)}"', xml)
    return "|".join(re.escape(p) for p in dict.fromkeys(prefixes or ["r"]))


def _resolve_sheet_zip_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb_xml = _dec(zf.read("xl/workbook.xml"))
    rels_xml = _dec(zf.read("xl/_rels/workbook.xml.rels"))
    m = re.search(rf'<sheet\b[^>]*\bname="{re.escape(sheet_name)}"[^>]*/>', wb_xml)
    if not m:
        raise SafeXLError(f"sheet {sheet_name!r} not found in workbook")
    tag = m.group(0)
    rid_m = re.search(rf'\b(?:{_rel_prefix_alt(wb_xml)}):id="([^"]+)"', tag)
    if not rid_m:
        raise SafeXLError(f"sheet {sheet_name!r}: no r:id attribute in xl/workbook.xml")
    rid = rid_m.group(1)
    rel_m = re.search(rf'<Relationship\b[^>]*\bId="{rid}"[^>]*/>', rels_xml)
    if not rel_m:
        raise SafeXLError(f"sheet {sheet_name!r}: no relationship for {rid} in workbook.xml.rels")
    target_m = re.search(r'Target="([^"]+)"', rel_m.group(0))
    target = target_m.group(1).lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


# ---------------------------------------------------------------------------
# Row/cell injection (order-preserving)
# ---------------------------------------------------------------------------


def _normalize_self_closing(sheet_xml: str) -> str:
    """A sheet with no rows serializes <sheetData/> self-closed, and a row
    with no cells can serialize <row .../> self-closed. Normalize both so the
    "find/insert inside" logic below always has an open/close pair to target."""
    if "<sheetData/>" in sheet_xml:
        sheet_xml = sheet_xml.replace("<sheetData/>", "<sheetData></sheetData>")
    sheet_xml = _SELF_CLOSING_ROW_RE.sub(lambda m: f"<row{m.group(1)}></row>", sheet_xml)
    return sheet_xml


def _insert_row_in_order(sheet_xml: str, row_num: int, new_row_xml: str) -> str:
    """Insert a brand-new <row> in correct ascending-r order among existing
    rows (not always at the end -- a missing row can be anywhere)."""
    insert_pos = None
    for m in _ROW_TAG_RE.finditer(sheet_xml):
        if int(m.group(1)) > row_num:
            insert_pos = m.start()
            break
    if insert_pos is None:
        idx = sheet_xml.index("</sheetData>")
        return sheet_xml[:idx] + new_row_xml + sheet_xml[idx:]
    return sheet_xml[:insert_pos] + new_row_xml + sheet_xml[insert_pos:]


def _inject_cells(sheet_xml: str, cell_updates: dict[str, Change]) -> tuple[str, bool]:
    """Update or insert every (ref -> change) into sheetData, in ref order
    for new rows/cells. Returns (new_sheet_xml, formula_touched) where
    formula_touched is True if any injected cell writes a formula OR
    overwrites a cell that used to contain one (both make an existing
    calcChain.xml potentially stale)."""
    _refuse_prefixed_worksheet(sheet_xml)
    sheet_xml = _normalize_self_closing(sheet_xml)
    formula_touched = False

    for ref, change in cell_updates.items():
        if "formula" in change:
            formula_touched = True

        cell_pat = re.compile(rf'<c r="{ref}"(?:/>|[^>]*(?:/>|>.*?</c>))', re.DOTALL)
        m = cell_pat.search(sheet_xml)
        if m:
            old = m.group(0)
            if "<f" in old:
                formula_touched = True
            s_m = re.search(r's="(\d+)"', old)
            style = s_m.group(1) if s_m else ""
            new_cell = _make_cell_xml(ref, change, style)
            sheet_xml = sheet_xml[: m.start()] + new_cell + sheet_xml[m.end():]
            continue

        row = _row_num(ref)
        row_pat = re.compile(rf'(<row\b[^>]*\br="{row}"[^>]*>)(.*?)(</row>)', re.DOTALL)
        row_m = row_pat.search(sheet_xml)
        new_cell = _make_cell_xml(ref, change)
        if row_m:
            new_col_key = _col_sort_key(ref)
            row_body = sheet_xml[row_m.start(2): row_m.end(2)]
            insert_offset = len(row_body)
            for cm in re.finditer(r'<c r="([A-Za-z]+\d+)"', row_body):
                if _col_sort_key(cm.group(1)) > new_col_key:
                    insert_offset = cm.start()
                    break
            abs_insert = row_m.start(2) + insert_offset
            sheet_xml = sheet_xml[:abs_insert] + new_cell + sheet_xml[abs_insert:]
        else:
            new_row = f'<row r="{row}">{new_cell}</row>'
            sheet_xml = _insert_row_in_order(sheet_xml, row, new_row)

    return sheet_xml, formula_touched


# ---------------------------------------------------------------------------
# Read-back confirmation (real XML parser, streamed)
# ---------------------------------------------------------------------------
#
# After a write, the output is re-read with ElementTree (comments, CDATA and PIs
# are handled by the parser, so a regex that landed in a decoy cannot fool it).
# Cost is one streamed pass per checked part; elements are released as each row
# closes, so memory is bounded by one row plus the checked cells.


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


_ROW_OPEN_RE = re.compile(r"<row\b[^>]*>")
_ROW_R_RE = re.compile(r'\br="(\d+)"')
_ROOT_OPEN_RE = re.compile(r"<worksheet\b[^>]*>")


def _cell_tuple(el, row_r):
    f_el = next((c for c in el if _local(c.tag) == "f"), None)
    v_el = next((c for c in el if _local(c.tag) == "v"), None)
    is_el = next((c for c in el if _local(c.tag) == "is"), None)
    inline = None
    if is_el is not None:
        inline = "".join(t.text or "" for t in is_el.iter() if _local(t.tag) == "t")
    return (row_r, el.get("r"), el.get("t"), None if f_el is None else (f_el.text or ""),
            None if v_el is None else (v_el.text or ""), inline, True)


def _iter_wanted_rows(zf: zipfile.ZipFile, part: str, wanted):
    """Cells of the rows for which ``wanted(row_number)`` is true. The rows are LOCATED with one C-speed
    regex pass over the part with comments/CDATA/PIs blanked (offsets preserved), then only those
    rows are cut out and handed to a real XML parser, so the cost is one scan plus a parse of the
    wanted rows, not of the whole sheet. Memory is the part text plus the wanted rows."""
    from xml.etree import ElementTree as ET

    text = zf.read(part).decode("utf-8", errors="replace")
    masked = mask_spans_same_length(text)
    root = _ROOT_OPEN_RE.search(masked)
    if not root:
        raise SafeXLError(f"read-back failed on {part}: no <worksheet> root found")
    slices = []
    for m in _ROW_OPEN_RE.finditer(masked):
        tag = m.group(0)
        rm = _ROW_R_RE.search(tag)
        if not rm or not wanted(int(rm.group(1))):
            continue
        if tag.endswith("/>"):
            slices.append(text[m.start():m.end()])
        else:
            end = masked.find("</row>", m.end())
            if end < 0:
                raise SafeXLError(f"read-back failed on {part}: row {rm.group(1)} is not closed")
            slices.append(text[m.start():end + 6])
    if not slices:
        return
    doc = text[root.start():root.end()] + "".join(slices) + "</worksheet>"
    for row in ET.fromstring(doc):
        r = row.get("r")
        for c in row:
            if _local(c.tag) == "c":
                yield _cell_tuple(c, r)


def _iter_cells(zf: zipfile.ZipFile, part: str, wanted=None):
    """Yield ``(row_r, cell_r, t, formula, v, inline_text, in_sheetdata_row)`` for every
    <c> of a worksheet part (or only of rows where ``wanted(row)``), from a real XML parse."""
    if wanted is not None:
        yield from _iter_wanted_rows(zf, part, wanted)
        return
    from xml.etree import ElementTree as ET

    row_r: str | None = None
    depth_row = False
    with zf.open(part) as fh:
        for event, el in ET.iterparse(fh, events=("start", "end")):
            name = _local(el.tag)
            if event == "start":
                if name == "row":
                    row_r, depth_row = el.get("r"), True
                continue
            if name == "c":
                yield _cell_tuple(el, row_r)[:6] + (depth_row,)
            elif name == "row":
                el.clear()
                row_r, depth_row = None, False


def _expected_matches(change: Change, t, f, v, inline) -> str | None:
    """None when the parsed cell holds what ``change`` asked for, else a reason."""
    if "formula" in change:
        want = _normalize_formula(change["formula"])
        return None if f is not None and f.strip() == want else f"formula reads back as {f!r}, expected {want!r}"
    value = change["value"]
    if isinstance(value, bool):
        return None if t == "b" and v == ("1" if value else "0") else f"boolean reads back as t={t!r} v={v!r}"
    if isinstance(value, (int, float)):
        try:
            got = float(v)
        except (TypeError, ValueError):
            return f"number reads back as {v!r}"
        want = float(value)
        same = got == want or (got != got and want != want)
        return None if same and t in (None, "n") and f is None else f"number reads back as t={t!r} v={v!r}, expected {value!r}"
    return None if t == "inlineStr" and inline == str(value) else f"text reads back as {inline if inline is not None else v!r}, expected {str(value)!r}"


def _readback_cells(zf: zipfile.ZipFile, part: str, sheet: str, expected: Mapping[str, Change]) -> None:
    """Confirm every cell in ``expected`` ({normalized ref: change}) reads back
    from ``part`` as requested, inside the <row> with the matching number.
    Raises SafeXLError naming the first cell that does not."""
    from xml.etree import ElementTree as ET

    pending = dict(expected)
    try:
        rows = {_row_num(k) for k in pending}
        for row_r, ref, t, f, v, inline, in_row in _iter_cells(zf, part, rows.__contains__):
            key = (ref or "").upper()
            if key not in pending:
                continue
            change = pending.pop(key)
            if not in_row or row_r is None or str(row_r) != str(_row_num(key)):
                raise SafeXLError(f"read-back failed at {sheet}!{key}: the cell is not inside row {_row_num(key)}")
            bad = _expected_matches(change, t, f, v, inline)
            if bad:
                raise SafeXLError(f"read-back failed at {sheet}!{key}: {bad}")
            if not pending:
                break
    except ET.ParseError as exc:
        raise SafeXLError(f"read-back failed on {sheet}: the written sheet is not parseable ({exc})") from exc
    if pending:
        missing = sorted(pending, key=lambda r: (_row_num(r), _col_sort_key(r)))[0]
        raise SafeXLError(f"read-back failed at {sheet}!{missing}: the cell is not present in the written sheet")


# ---------------------------------------------------------------------------
# <dimension> update (a fix required by the adversarial verification)
# ---------------------------------------------------------------------------


def _update_dimension(sheet_xml: str, touched_refs: Sequence[str]) -> str:
    """Recompute <dimension ref="..."> as the union of the sheet's existing
    declared dimension and every cell ref touched by this edit. Cheap,
    surgical fix for the stale-dimension defect confirmed in 7/7 outputs by
    the adversarial verification. Excel/LibreOffice compute the real used
    range from <sheetData> regardless (dimension is a non-authoritative hint),
    so this is a best-effort widen, not a full sheetData rescan."""
    m = _DIMENSION_RE.search(sheet_xml)
    if not m:
        return sheet_xml  # no <dimension> element present -- nothing to widen

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
        min_col = min(_col_to_num(sc), _col_to_num(ec))
        max_col = max(_col_to_num(sc), _col_to_num(ec))
        min_row = min(sr, er)
        max_row = max(sr, er)
    except SafeXLError:
        # Existing dimension isn't parseable -- don't guess, leave as-is.
        return sheet_xml

    for ref in touched_refs:
        letters, row = _split_ref(ref)
        col = _col_to_num(letters)
        min_col, max_col = min(min_col, col), max(max_col, col)
        min_row, max_row = min(min_row, row), max(max_row, row)

    if min_col == max_col and min_row == max_row:
        new_ref = f"{_num_to_col(min_col)}{min_row}"
    else:
        new_ref = f"{_num_to_col(min_col)}{min_row}:{_num_to_col(max_col)}{max_row}"

    if new_ref == existing_ref:
        return sheet_xml  # no-op -- keeps the part byte-identical if nothing grew
    return sheet_xml[: m.start()] + f'<dimension ref="{new_ref}"/>' + sheet_xml[m.end():]


# ---------------------------------------------------------------------------
# calcChain.xml / workbook.xml / Content_Types plumbing
# ---------------------------------------------------------------------------


def _force_full_calc_on_load(workbook_xml_text: str) -> str:
    if "fullCalcOnLoad" in workbook_xml_text:
        return workbook_xml_text
    if "<calcPr " in workbook_xml_text:
        return re.sub(r"<calcPr ", '<calcPr fullCalcOnLoad="1" ', workbook_xml_text, count=1)
    if "<calcPr/>" in workbook_xml_text:
        return workbook_xml_text.replace("<calcPr/>", '<calcPr fullCalcOnLoad="1"/>')
    return workbook_xml_text.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def _strip_calc_chain_override(content_types_text: str) -> str:
    return re.sub(
        r'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>\s*',
        "",
        content_types_text,
    )


def _strip_calc_chain_relationship(workbook_rels_text: str) -> str:
    """Drop workbook.xml.rels' Relationship to calcChain.xml. Dropping the part
    but keeping this relationship leaves a dangling reference, which is an
    Excel repair trigger (caught on a recorded benchmark run, 2026-09-26)."""
    return re.sub(
        r'<Relationship\b[^>]*Target="(?:/xl/)?calcChain\.xml"[^>]*/>\s*',
        "",
        workbook_rels_text,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def edit(
    src: str | Path,
    dst: str | Path,
    changes: Sequence[Change],
    *,
    extend_tables: bool = False,
) -> None:
    """Apply ``changes`` to the workbook at ``src`` and write the result to
    ``dst``, preserving every untouched OOXML part byte-for-byte.

    Parameters
    ----------
    src : path to the source .xlsx/.xlsm to read (never modified).
    dst : path to write the edited workbook to. Only written on success --
        see the module docstring's "ROLLBACK PROMISE". May equal ``src``
        (edit-in-place) since the write goes through a temp file first.
    changes : a sequence of dicts, each either
        ``{"sheet": str, "cell": "A1", "value": <bool|int|float|str>}`` or
        ``{"sheet": str, "cell": "A1", "formula": "A1*2"}`` (leading "=" on
        the formula is accepted and stripped; it is never stored inside the
        written ``<f>`` element -- see ``_normalize_formula``). A later
        change for the same (sheet, cell) overwrites an earlier one.
    extend_tables : opt-in. When True, a written cell directly adjacent to an
        Excel table (the next column within its row span, or the next row
        within its column span) grows that table to include it, in the same
        write, the way Excel does when you type next to a table. A new column
        is named from its header cell; an empty header cell gets Excel's
        ``Column{n}`` name, written into the cell too. ``structural.plan_table_extensions``
        previews this without writing.
        Default False leaves every table part byte-identical.

    Raises
    ------
    SafeXLError on any invalid change, unresolvable sheet/cell, or a
    post-write structural validation failure. ``dst`` is left untouched in
    every failure case.
    """
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise SafeXLError(f"source file not found: {src}")
    if not changes:
        raise SafeXLError("changes must be a non-empty sequence")

    grouped = _group_changes(changes)  # validates everything up front

    with zipfile.ZipFile(src, "r") as zs:
        sheet_targets: dict[str, str] = {}
        for sheet_name in grouped:
            sheet_targets[sheet_name] = _resolve_sheet_zip_path(zs, sheet_name)

    target_to_sheet = {v: k for k, v in sheet_targets.items()}

    table_patches: dict[str, str] = {}
    if extend_tables:
        from .structural import _plan_table_extensions  # lazy: structural imports this module

        for plan in _plan_table_extensions(src, grouped, sheet_targets):
            table_patches[plan["part"]] = plan["xml"]
            for col in plan["new_columns"]:
                if col["auto_named"] and col["header_cell"]:
                    # Excel writes the generated name into the header cell too.
                    grouped[plan["grouped_sheet"]][col["header_cell"]] = {
                        "sheet": plan["grouped_sheet"],
                        "cell": col["header_cell"],
                        "value": col["name"],
                    }

    tmp_dir = dst.parent if dst.parent.exists() else Path(tempfile.gettempdir())
    tmp_fd, tmp_path_str = tempfile.mkstemp(suffix=dst.suffix or ".xlsx", dir=tmp_dir)
    os.close(tmp_fd)
    tmp_path = Path(tmp_path_str)

    try:
        with zipfile.ZipFile(src, "r") as zs, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zs.infolist():
                if item.filename == "xl/calcChain.xml":
                    continue  # dropped unconditionally; see module docstring

                data = zs.read(item.filename)

                if item.filename in target_to_sheet:
                    text = _dec(data)
                    cell_updates = grouped[target_to_sheet[item.filename]]
                    text, _formula_touched = _inject_cells(text, cell_updates)
                    text = _update_dimension(text, list(cell_updates.keys()))
                    data = _enc(text)

                elif item.filename in table_patches:
                    data = _enc(table_patches[item.filename])

                elif item.filename == "xl/workbook.xml":
                    text = _dec(data)
                    text = _force_full_calc_on_load(text)
                    data = _enc(text)

                elif item.filename == "[Content_Types].xml":
                    text = _dec(data)
                    text = _strip_calc_chain_override(text)
                    data = _enc(text)

                elif item.filename == "xl/_rels/workbook.xml.rels":
                    text = _dec(data)
                    stripped = _strip_calc_chain_relationship(text)
                    if stripped != text:  # untouched (byte-identical) when there's no calc chain
                        data = _enc(stripped)

                zout.writestr(item, data)

        with zipfile.ZipFile(tmp_path, "r") as zchk:
            for sheet_name, part in sheet_targets.items():
                _readback_cells(zchk, part, sheet_name, grouped[sheet_name])

        report = validate_workbook(tmp_path)
        if report.critical_count:
            details = "; ".join(report.critical_details())
            raise SafeXLError(
                f"edit produced an invalid workbook, dst NOT written "
                f"({report.critical_count} critical structural finding(s)): {details}"
            )

        shutil.move(str(tmp_path), str(dst))
    except SafeXLError:
        raise
    except Exception as exc:  # noqa: BLE001 - re-raise everything as SafeXLError
        raise SafeXLError(f"failed to write edited workbook: {exc}") from exc
    finally:
        tmp_path.unlink(missing_ok=True)
