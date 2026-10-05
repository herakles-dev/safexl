"""Reference-rewriting helpers shared by the structural edits.

Pure string functions, stdlib only, no zip or file I/O. ``structural.py``
decides *which* OOXML parts to touch; this module knows how to rewrite the
reference text inside them:

* ``map_refs``          -- scan a formula and rewrite each cell/range token
* ``shift_rows_in_formula`` / ``shift_sqref`` -- row-insert semantics
* ``translate_formula`` -- move relative refs by (dr, dc); used to expand
  shared formulas into explicit ones
* ``rename_sheet_in_formula`` -- rewrite ``Old!`` / ``'Old'!`` prefixes

The formula scanner is a tokenizer, not a full parser. It skips string
literals (``"..."``), structured references (``Table1[[#This Row],[Q1]]``)
and external-workbook references (``[1]Sheet1!A1``), and it leaves anything
that is not a clean A1 / row-range / column-range token alone (function
names such as ``LOG10(``, defined names, ``#REF!``).

Row-insert rule (same as Excel): every endpoint whose row is ``>= row`` moves
down by ``count``. A range that straddles the insertion point therefore
grows; a range that ends just above it does not.
"""
from __future__ import annotations

import re
from typing import Callable
from xml.sax.saxutils import escape as _sax_escape
from xml.sax.saxutils import unescape as _sax_unescape

MAX_ROW = 1_048_576  # Excel's row limit
MAX_COL = 16_384  # XFD

# ---------------------------------------------------------------------------
# XML text escaping that round-trips numeric character references
# ---------------------------------------------------------------------------

_NUMERIC_ENTITY_RE = re.compile(r"&#(?:x([0-9A-Fa-f]+)|([0-9]+));")
_NAMED = {"&apos;": "'", "&quot;": '"'}


def xml_unescape(text: str) -> str:
    text = _NUMERIC_ENTITY_RE.sub(lambda m: chr(int(m.group(1), 16) if m.group(1) else int(m.group(2))), text)
    return _sax_unescape(text, _NAMED)


def xml_escape(text: str) -> str:
    return _sax_escape(text).replace("\r", "&#13;")


def escape_attr(text: str) -> str:
    return _sax_escape(text, {'"': "&quot;"}).replace("\r", "&#13;").replace("\n", "&#10;")


# ---------------------------------------------------------------------------
# Column / cell helpers
# ---------------------------------------------------------------------------


def col_to_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


_ENDPOINT_RE = re.compile(r"^(\$?)([A-Za-z]{0,3})(\$?)([0-9]+)$")


def _shift_endpoint(text: str, row: int, count: int) -> str:
    """Shift one range endpoint ('A5', '$A$5', or a bare row '5' / '$5').
    Column-only endpoints ('A', '$A') are returned unchanged."""
    m = _ENDPOINT_RE.match(text)
    if not m:
        return text
    cdollar, col, rdollar, rnum = m.groups()
    n = int(rnum)
    if n >= row:
        n = min(n + count, MAX_ROW)
    return f"{cdollar}{col}{rdollar}{n}"


def shift_sqref(sqref: str, row: int, count: int) -> str:
    """Shift a whitespace-separated list of ranges ('A1:B5 D7 9:12')."""
    out = []
    for tok in sqref.split():
        out.append(":".join(_shift_endpoint(p, row, count) for p in tok.split(":")))
    return " ".join(out)


# ---------------------------------------------------------------------------
# Formula tokenizer
# ---------------------------------------------------------------------------

_IDENT = r"[^\W\d][\w.]*"
_PREFIX = rf"(?:'(?:[^']|'')+'|{_IDENT})!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
_ROWNUM = r"\$?[1-9][0-9]{0,6}"
_COLLET = r"\$?[A-Za-z]{1,3}"

_REF_TOKEN = re.compile(
    rf"(?P<prefix>{_PREFIX})?"
    rf"(?:(?P<c1>{_CELL})(?::(?P<c2>{_CELL}))?"
    rf"|(?P<r1>{_ROWNUM}):(?P<r2>{_ROWNUM})"
    rf"|(?P<k1>{_COLLET}):(?P<k2>{_COLLET}))"
    r"(?![\w.(\[!])"
)
_PREFIX_ONLY = re.compile(_PREFIX)
_EXTERNAL_PREFIX = re.compile(rf"\[[0-9]+\](?:'(?:[^']|'')+'|{_IDENT})!")

_BOUNDARY_BLOCKERS = "_.$#'\\]"


def prefix_sheet_name(prefix: str) -> str:
    """'Sheet 1'! -> Sheet 1 ; Data! -> Data"""
    name = prefix[:-1]
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return name


def _cols_valid(m: "re.Match[str]") -> bool:
    for key in ("c1", "c2", "k1", "k2"):
        tok = m.group(key)
        if tok:
            letters = re.sub(r"[^A-Za-z]", "", tok)
            if col_to_num(letters) > MAX_COL:
                return False
    return True


def _skip_span(formula: str, i: int) -> int:
    """If a string literal, structured reference or external-workbook
    reference starts at ``i``, return the index just past it (the caller
    copies it verbatim); otherwise return ``i``."""
    n = len(formula)
    ch = formula[i]
    if ch == '"':
        j = i + 1
        while j < n:
            if formula[j] == '"':
                if j + 1 < n and formula[j + 1] == '"':
                    j += 2
                    continue
                return j + 1
            j += 1
        return n
    if ch == "[":
        ext = _EXTERNAL_PREFIX.match(formula, i)
        if ext:
            j = ext.end()
            ref = _REF_TOKEN.match(formula, j)
            if ref and ref.group("prefix") is None:
                return ref.end()
            return j
        depth = 0
        j = i
        while j < n:
            c = formula[j]
            if c == "'" and j + 1 < n and formula[j + 1] in "[]#'":  # Excel's escape set
                j += 2
                continue
            if c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
                if depth == 0:
                    return j + 1
            j += 1
        return n
    return i


def _at_boundary(formula: str, i: int) -> bool:
    if i == 0:
        return True
    prev = formula[i - 1]
    return not (prev.isalnum() or prev in _BOUNDARY_BLOCKERS)


def map_refs(formula: str, on_ref: Callable[["re.Match[str]"], str]) -> str:
    """Rewrite every cell / row-range / column-range token in ``formula``.
    ``on_ref`` gets the regex match (groups: prefix, c1, c2, r1, r2, k1, k2)
    and returns the full replacement text including any prefix."""
    out: list[str] = []
    i, n = 0, len(formula)
    while i < n:
        j = _skip_span(formula, i)
        if j != i:
            out.append(formula[i:j])
            i = j
            continue
        if _at_boundary(formula, i):
            m = _REF_TOKEN.match(formula, i)
            if m and _cols_valid(m):
                out.append(on_ref(m))
                i = m.end()
                continue
        out.append(formula[i])
        i += 1
    return "".join(out)


# ---------------------------------------------------------------------------
# Row insert
# ---------------------------------------------------------------------------


def shift_rows_in_formula(
    formula: str, current_sheet: str | None, target_sheet: str, row: int, count: int
) -> str:
    """Shift references into ``target_sheet`` for an insert of ``count`` rows
    at ``row``. Unqualified refs belong to ``current_sheet`` (None: the
    formula has no home sheet, e.g. a defined name, so only qualified refs
    are touched). Sheet names compare case-insensitively, like Excel."""
    tl = target_sheet.lower()

    def on_ref(m: "re.Match[str]") -> str:
        prefix = m.group("prefix")
        sheet = current_sheet if prefix is None else prefix_sheet_name(prefix)
        if sheet is None or sheet.lower() != tl:
            return m.group(0)
        if m.group("c1"):
            body = _shift_endpoint(m.group("c1"), row, count)
            if m.group("c2"):
                body += ":" + _shift_endpoint(m.group("c2"), row, count)
        elif m.group("r1"):
            body = _shift_endpoint(m.group("r1"), row, count) + ":" + _shift_endpoint(m.group("r2"), row, count)
        else:
            return m.group(0)  # column range: rows don't matter
        return (prefix or "") + body

    return map_refs(formula, on_ref)


# ---------------------------------------------------------------------------
# Relative-reference translation (shared-formula expansion)
# ---------------------------------------------------------------------------

_CELL_PARTS_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")


def _translate_cell(tok: str, dr: int, dc: int) -> str | None:
    cdollar, col, rdollar, rnum = _CELL_PARTS_RE.match(tok).groups()  # type: ignore[union-attr]
    c, r = col_to_num(col), int(rnum)
    if not cdollar:
        c += dc
    if not rdollar:
        r += dr
    if not (1 <= c <= MAX_COL and 1 <= r <= MAX_ROW):
        return None
    return f"{cdollar}{num_to_col(c)}{rdollar}{r}"


def _translate_row(tok: str, dr: int) -> str | None:
    if tok.startswith("$"):
        return tok
    r = int(tok) + dr
    return str(r) if 1 <= r <= MAX_ROW else None


def _translate_col(tok: str, dc: int) -> str | None:
    if tok.startswith("$"):
        return tok
    c = col_to_num(tok) + dc
    return num_to_col(c) if 1 <= c <= MAX_COL else None


def translate_formula(formula: str, dr: int, dc: int) -> str:
    """Move every relative reference by (dr rows, dc columns), leaving
    ``$``-absolute parts alone -- what Excel does when it fills a formula."""

    def on_ref(m: "re.Match[str]") -> str:
        prefix = m.group("prefix") or ""
        if m.group("c1"):
            parts = [_translate_cell(m.group("c1"), dr, dc)]
            if m.group("c2"):
                parts.append(_translate_cell(m.group("c2"), dr, dc))
        elif m.group("r1"):
            parts = [_translate_row(m.group("r1"), dr), _translate_row(m.group("r2"), dr)]
        else:
            parts = [_translate_col(m.group("k1"), dc), _translate_col(m.group("k2"), dc)]
        if any(p is None for p in parts):
            return "#REF!"
        return prefix + ":".join(parts)  # type: ignore[arg-type]

    return map_refs(formula, on_ref)


# ---------------------------------------------------------------------------
# Sheet rename
# ---------------------------------------------------------------------------

_UNQUOTED_OK_RE = re.compile(r"^[^\W\d][\w]*$")
_CELL_LIKE_RE = re.compile(r"^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*[Cc][0-9]*|[Rr]|[Cc])$")


def quote_sheet_name_if_needed(name: str) -> str:
    if _UNQUOTED_OK_RE.match(name) and not _CELL_LIKE_RE.match(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def rename_sheet_in_formula(formula: str, old_name: str, new_name: str) -> str:
    """Rewrite ``Old!`` / ``'Old'!`` prefixes (case-insensitive match) to the
    new name, quoted only when Excel would need quotes."""
    ol = old_name.lower()
    new_prefix = quote_sheet_name_if_needed(new_name) + "!"
    out: list[str] = []
    i, n = 0, len(formula)
    while i < n:
        j = _skip_span(formula, i)
        if j != i:
            out.append(formula[i:j])
            i = j
            continue
        if _at_boundary(formula, i):
            m = _PREFIX_ONLY.match(formula, i)
            if m:
                out.append(new_prefix if prefix_sheet_name(m.group(0)).lower() == ol else m.group(0))
                i = m.end()
                continue
        out.append(formula[i])
        i += 1
    return "".join(out)


# ---------------------------------------------------------------------------
# Column-scoped insert ("insert cells, shift down" within columns c1..c2)
# ---------------------------------------------------------------------------


class TornReference(ValueError):
    """A reference Excel would have to tear apart (see ``shift_cells_in_formula``)."""


def _torn(text: str, c1: int, c2: int, why: str) -> TornReference:
    return TornReference(
        f"{text!r} {why} (columns {num_to_col(c1)}:{num_to_col(c2)} are shifted)"
    )


def _classify(
    a: int, b: int, ra: int, rb: int, s1: int, s2: int, row: int, *, allow_grow: bool, text: str
) -> str:
    """'none' (untouched), 'move' (apply the endpoint rule) -- or raise TornReference.

    A range wholly inside the shifted column span follows the row-insert rule
    (moves, or grows when it straddles ``row`` and ``allow_grow``). One wholly
    outside the span, or ending above ``row``, is untouched. A range that
    straddles the span's edge with rows at/below ``row`` has some cells that
    move and some that stay, so it is torn."""
    if b < s1 or a > s2 or rb < row:
        return "none"
    if s1 <= a and b <= s2:
        if ra < row <= rb and not allow_grow:
            raise _torn(text, s1, s2, "straddles the insertion point and cannot grow")
        return "move"
    raise _torn(text, s1, s2, "partly overlaps the shifted columns")


def shift_cells_in_formula(
    formula: str, current_sheet: str | None, target_sheet: str, c1: int, c2: int, row: int, count: int
) -> str:
    """Reference following for an insert of ``count`` cells shifting down at
    ``row`` within columns ``c1..c2`` of ``target_sheet``. Same sheet rules as
    ``shift_rows_in_formula``. Raises ``TornReference`` for a range that
    partly overlaps the span (or a row range such as ``5:9``) with rows at or
    below ``row``; whole-column ranges are never affected."""
    tl = target_sheet.lower()

    def on_ref(m: "re.Match[str]") -> str:
        prefix = m.group("prefix")
        sheet = current_sheet if prefix is None else prefix_sheet_name(prefix)
        if sheet is None or sheet.lower() != tl:
            return m.group(0)
        if m.group("c1"):
            parts = [m.group("c1")] + ([m.group("c2")] if m.group("c2") else [])
            pts = [_CELL_PARTS_RE.match(p).groups() for p in parts]  # type: ignore[union-attr]
            cols = [col_to_num(p[1]) for p in pts]
            rows = [int(p[3]) for p in pts]
            if _classify(min(cols), max(cols), min(rows), max(rows), c1, c2, row, allow_grow=True, text=m.group(0)) == "none":
                return m.group(0)
            return (prefix or "") + ":".join(_shift_endpoint(p, row, count) for p in parts)
        if m.group("r1"):
            lo, hi = sorted(int(re.sub(r"\$", "", m.group(k))) for k in ("r1", "r2"))
            if hi >= row and not (c1 <= 1 and c2 >= MAX_COL):
                raise _torn(m.group(0), c1, c2, "is a row range that spans the shifted columns partly")
            if hi >= row:
                return (prefix or "") + ":".join(_shift_endpoint(m.group(k), row, count) for k in ("r1", "r2"))
            return m.group(0)
        return m.group(0)  # whole-column range: unaffected

    return map_refs(formula, on_ref)


def _range_bounds(token: str) -> tuple[int, int, int, int] | None:
    parts = token.replace("$", "").split(":")
    pts = []
    for p in parts[:2]:
        m = re.match(r"^([A-Za-z]{1,3})([0-9]+)$", p)
        if not m:
            return None
        pts.append((col_to_num(m.group(1)), int(m.group(2))))
    if len(pts) == 1:
        pts.append(pts[0])
    return min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts)


def shift_ref_cols(ref: str, c1: int, c2: int, row: int, count: int, *, allow_grow: bool) -> str:
    """Column-scoped shift of a whitespace-separated ref list (hyperlink, merge,
    table, autoFilter, array-formula refs). Raises ``TornReference`` as above;
    ``allow_grow=False`` also refuses a range that contains the insertion point."""
    out = []
    for tok in ref.split():
        b = _range_bounds(tok)
        if b is None:
            out.append(tok)  # whole-column / unparseable: leave alone
            continue
        if _classify(b[0], b[2], b[1], b[3], c1, c2, row, allow_grow=allow_grow, text=tok) == "none":
            out.append(tok)
        else:
            out.append(":".join(_shift_endpoint(p, row, count) for p in tok.split(":")))
    return " ".join(out)


def shift_sqref_cols(sqref: str, c1: int, c2: int, row: int, count: int) -> str:
    """Column-scoped shift of a conditional-format / data-validation ``sqref``.
    Unlike the strict refs above, a range that straddles the span's edge is
    split: the part inside the span follows the row-insert rule, the rest stays,
    because a CF/DV rule belongs to cells and the cells that moved take it along."""
    out = []
    for tok in sqref.split():
        b = _range_bounds(tok)
        if b is None:
            if re.fullmatch(r"\$?\d+:\$?\d+", tok) and max(int(x.replace("$", "")) for x in tok.split(":")) >= row:
                raise _torn(tok, c1, c2, "is a row range that spans the shifted columns partly")
            out.append(tok)
            continue
        a, ra, bb, rb = b
        if bb < c1 or a > c2 or rb < row:
            out.append(tok)
            continue
        lo, hi = max(a, c1), min(bb, c2)
        mid_ra = ra + count if ra >= row else ra
        mid = f"{num_to_col(lo)}{mid_ra}:{num_to_col(hi)}{rb + count}"
        if lo == hi and mid_ra == rb + count:
            mid = f"{num_to_col(lo)}{mid_ra}"
        pieces = []
        if a < c1:
            pieces.append(f"{num_to_col(a)}{ra}:{num_to_col(c1 - 1)}{rb}")
        pieces.append(mid)
        if bb > c2:
            pieces.append(f"{num_to_col(c2 + 1)}{ra}:{num_to_col(bb)}{rb}")
        out.append(" ".join(p if ":" in p and p.split(":")[0] != p.split(":")[1] else p.split(":")[0] for p in pieces))
    return " ".join(out)
