"""safexl_mcp.features -- read-only feature inventory + per-part hashes of an OOXML workbook.

``fingerprint(path)`` answers "what is in this package": counts of the constructs a lossy save
(openpyxl / pandas round trip) destroys, plus the sha256 of every part. It reads the zip only: it
never executes macros (the VBA project is a part that is counted, not opened) and never fetches
anything (external links are counted, not followed).

Resource use is bounded: a package over MAX_PARTS parts, MAX_PART_BYTES for one part or
MAX_TOTAL_BYTES decompressed in total is refused (``PackageTooLarge``) before any part is read;
the byte counts are enforced while reading, so a zip that lies about its sizes is cut off too.
"""
from __future__ import annotations

import fnmatch
import hashlib
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Iterator

MAX_PARTS = 5000
MAX_PART_BYTES = 128 * 1024 * 1024
MAX_TOTAL_BYTES = 512 * 1024 * 1024

FEATURE_KEYS: tuple[str, ...] = (
    "sheets", "defined_names", "charts", "drawings", "images", "pivot_tables", "pivot_caches",
    "tables", "vba", "external_links", "sparkline_groups", "data_validations",
    "conditional_formats", "comments", "merged_cells", "hyperlinks", "protected_sheets",
)

# Parts a SafeXL edit / insert-rows / rename-sheet / add-sheet never rewrites (byte-identical
# passthrough) but every openpyxl / pandas save regenerates. Measured in a local-model
# test battery: 0 false positives on 1026 SafeXL ops, 24/24 openpyxl round trips caught.
UNTOUCHABLE: tuple[str, ...] = (
    "xl/styles.xml", "docProps/*", "_rels/.rels", "xl/theme/*", "xl/media/*", "xl/vbaProject.bin",
    "xl/printerSettings/*", "customXml/*",
)

# Features are identified by content type (from [Content_Types].xml), never by part path or file
# name: another writer may keep a comment, chart or pivot under a different path.
_CT = "application/vnd.openxmlformats-officedocument."
_PART_TYPES: dict[str, tuple[str, ...]] = {
    "charts": (_CT + "drawingml.chart+xml",),
    "drawings": (_CT + "drawing+xml",),
    "pivot_tables": (_CT + "spreadsheetml.pivotTable+xml",),
    "pivot_caches": (_CT + "spreadsheetml.pivotCacheDefinition+xml",),
    "tables": (_CT + "spreadsheetml.table+xml",),
    "external_links": (_CT + "spreadsheetml.externalLink+xml",),
    "vba": ("application/vnd.ms-office.vbaProject",),
}
_SHEET_TYPES = frozenset({
    _CT + "spreadsheetml.worksheet+xml", _CT + "spreadsheetml.chartsheet+xml",
    _CT + "spreadsheetml.dialogsheet+xml", "application/vnd.ms-excel.macrosheet+xml",
})
_COMMENT_TYPES = frozenset({_CT + "spreadsheetml.comments+xml", "application/vnd.ms-excel.threadedcomments+xml"})
_WORKBOOK_TYPES = frozenset({
    _CT + "spreadsheetml.sheet.main+xml", _CT + "spreadsheetml.template.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml", "application/vnd.ms-excel.template.macroEnabled.main+xml",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
})
_RE_SPARK = re.compile(rb"<(?:\w+:)?sparklineGroup[\s>/]")
_RE_DV = re.compile(rb"<(?:\w+:)?dataValidation[\s>/]")
_RE_CF = re.compile(rb"<(?:\w+:)?conditionalFormatting[\s>/]")  # also matches x14:conditionalFormatting
_RE_COMMENT = re.compile(rb"<(?:\w+:)?(?:comment|threadedComment)[\s>/]")
_RE_CHART_F = re.compile(rb"<(?:\w+:)?f>([^<]*)</(?:\w+:)?f>")
_RE_SER = re.compile(rb"<(?:\w+:)?ser[\s>]")
_RE_SHEET = re.compile(rb"<(?:\w+:)?sheet\s")
_RE_NAME = re.compile(rb"<(?:\w+:)?definedName[\s>]")


def _content_types(data: bytes | None) -> tuple[dict[str, str], dict[str, str]]:
    """({extension: type} defaults, {/part: type} overrides) from [Content_Types].xml; empty if absent/unparseable."""
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    if data is None:
        return defaults, overrides
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return defaults, overrides
    for child in root:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "Default":
            defaults[child.get("Extension", "").lower()] = child.get("ContentType", "")
        elif tag == "Override":
            overrides[child.get("PartName", "")] = child.get("ContentType", "")
    return defaults, overrides


class PackageTooLarge(ValueError):
    """The package is over the part-count or decompressed-size caps."""


def iter_parts(path: str | Path) -> Iterator[tuple[str, bytes]]:
    """Yield (name, bytes) for every non-directory member, enforcing the caps.

    Raises zipfile.BadZipFile / OSError for an unreadable file and PackageTooLarge over the caps."""
    with zipfile.ZipFile(path) as z:
        infos = [i for i in z.infolist() if not i.filename.endswith("/")]
        if len(infos) > MAX_PARTS:
            raise PackageTooLarge(f"{len(infos)} parts exceeds the cap of {MAX_PARTS}")
        if sum(i.file_size for i in infos) > MAX_TOTAL_BYTES:
            raise PackageTooLarge(f"declared size exceeds the cap of {MAX_TOTAL_BYTES} bytes")
        total = 0
        for info in infos:
            if info.file_size > MAX_PART_BYTES:
                raise PackageTooLarge(f"{info.filename} exceeds the per-part cap of {MAX_PART_BYTES} bytes")
            with z.open(info) as fh:
                data = fh.read(MAX_PART_BYTES + 1)
            total += len(data)
            if len(data) > MAX_PART_BYTES or total > MAX_TOTAL_BYTES:
                raise PackageTooLarge("decompressed size exceeds the caps")
            yield info.filename, data


# --------------------------------------------------------------------------------------------
# Per-item detail: what a lossy save silently drops or alters (same categories the benchmark grader
# checks that do not depend on a task's `expect` block). Items are compared with rules that let a
# legitimate structural edit through:
#   * sheets are paired by NAME first, then by worksheet part (a renamed sheet keeps its part), so a
#     rename is never a loss and sheet names are never part of an item's signature;
#   * a range-bearing item (data validation, conditional-format block, merged cell, sparkline group)
#     is matched on its range-free signature, then on its ranges: an original rectangle is kept when
#     some rectangle in the new item has the same-or-wider columns, a last row at or below the
#     original's, and a height at least the original's. That is "shifted down or grown" (insert-rows,
#     a table growing); a rectangle that moved up, shrank, or vanished is a removal;
#   * formulas are not compared (insert-rows / rename-sheet legitimately rewrite them), only literal
#     list validations, hyperlink targets, and VBA / image bytes.
# --------------------------------------------------------------------------------------------

_NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_REL_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_RE_RECT = re.compile(r"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")


def _local(tag: str) -> tuple[str, str]:
    if tag.startswith("{"):
        ns, _, name = tag[1:].partition("}")
        return ns, name
    return "", tag


def _colnum(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _rects(sqref: str | None) -> list[list[int]]:
    """[[c1, r1, c2, r2], ...] for an A1-style sqref; [] when it has any token that is not a plain range."""
    out = []
    for tok in (sqref or "").split():
        m = _RE_RECT.match(tok)
        if not m:
            return []
        c1, r1 = _colnum(m.group(1)), int(m.group(2))
        c2, r2 = (_colnum(m.group(3)), int(m.group(4))) if m.group(3) else (c1, r1)
        out.append([min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)])
    return out


def _resolve(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _rels_of(data_of: dict[str, bytes], part: str) -> dict[str, tuple[str, str, bool]]:
    """{rId: (type, target, external)} from the .rels part beside `part`."""
    rels_name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    out: dict[str, tuple[str, str, bool]] = {}
    raw = data_of.get(rels_name)
    if raw is None:
        return out
    try:
        root = ET.fromstring(raw)
    except ET.ParseError:
        return out
    for rel in root:
        if _local(rel.tag)[1] == "Relationship":
            out[rel.get("Id", "")] = (rel.get("Type", ""), rel.get("Target", ""), rel.get("TargetMode") == "External")
    return out


def _parse(data: bytes):
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _without_sheet_rows(raw: bytes) -> bytes:
    """The worksheet with the content of <sheetData> removed. _sheet_detail reads only elements outside
    the cell grid (protection, merges, links, validations, formats, sparklines), so parsing millions of
    cells to build a tree it never looks at is pure cost. The span is located on a copy with comments,
    CDATA and PIs blanked (offsets preserved), so a decoy '<sheetData>' in one cannot move it. A sheet
    that does not carry an unprefixed <sheetData>...</sheetData> pair is returned unchanged."""
    from .engine import mask_spans_same_length

    text = raw.decode("utf-8", errors="replace")
    if len(text) != len(raw) and not raw.isascii():
        return raw  # offsets would not line up with the bytes; keep the full parse
    masked = mask_spans_same_length(text)
    start = masked.find("<sheetData>")
    end = masked.rfind("</sheetData>")
    if start < 0 or end < start:
        return raw
    return (text[: start + len("<sheetData>")] + text[end:]).encode("utf-8")


def _sheet_detail(root, rels: dict[str, tuple[str, str, bool]]) -> dict:
    d: dict = {"dv": [], "cf": [], "merged": [], "spark": [], "links": [], "protected": False}
    for el in root.iter():
        ns, name = _local(el.tag)
        if name == "sheetProtection" and ns != _NS_X14:
            d["protected"] = True
        elif name == "mergeCell":
            d["merged"].append({"sig": [], "ref": el.get("ref", ""), "rects": _rects(el.get("ref"))})
        elif name == "hyperlink":
            rid = el.get(f"{{{_REL_OFFICE}}}id")
            target = rels.get(rid or "", (None, None, False))[1]
            d["links"].append(["ext", target] if target else ["int"])
        elif name == "dataValidation":
            kind = "x14" if ns == _NS_X14 else "classic"
            f1 = ""
            sqref = el.get("sqref", "")
            for ch in el.iter():
                cn = _local(ch.tag)[1]
                if cn == "formula1" and not f1:
                    f1 = "".join(ch.itertext()).strip()
                elif cn == "sqref" and not sqref:
                    sqref = (ch.text or "").strip()
            lit = f1 if f1.startswith('"') else ""
            d["dv"].append({"sig": [kind, el.get("type", "none"), lit], "ref": sqref, "rects": _rects(sqref)})
        elif name == "conditionalFormatting":
            kind = "x14" if ns == _NS_X14 else "classic"
            types = sorted(c.get("type", "") for c in el.iter() if _local(c.tag)[1] == "cfRule")
            sqref = el.get("sqref", "")
            if not sqref:
                for ch in el.iter():
                    if _local(ch.tag)[1] == "sqref":
                        sqref = (ch.text or "").strip()
            d["cf"].append({"sig": [kind] + types, "ref": sqref, "rects": _rects(sqref)})
        elif name == "sparklineGroup":
            n = sum(1 for c in el.iter() if _local(c.tag)[1] == "sparkline")
            d["spark"].append({"sig": [el.get("type", "line"), n], "ref": f"{n} sparkline(s)", "rects": []})
    return d


def _build_detail(data_of: dict[str, bytes], parts: dict[str, str], ctype_of: dict[str, str]) -> dict:
    detail: dict = {"sheets": [], "vba": [], "images": [], "workbook_protected": False, "defined_names": [],
                    "per_sheet": {}, "tables": [], "pivot_tables": [], "pivot_caches": [],
                    "external_links": [], "charts": [], "chart_refs": []}
    wb_part = next((n for n, c in ctype_of.items() if c in _WORKBOOK_TYPES), None)
    sheet_names_by_index: list[str] = []
    if wb_part and (root := _parse(data_of[wb_part])) is not None:
        rels = _rels_of(data_of, wb_part)
        for el in root.iter():
            _, nm = _local(el.tag)
            if nm == "sheet":
                rid = el.get(f"{{{_REL_OFFICE}}}id") or ""
                target = rels.get(rid, ("", "", False))[1]
                detail["sheets"].append({"name": el.get("name", ""), "part": _resolve(wb_part, target) if target else ""})
            elif nm == "workbookProtection":
                detail["workbook_protected"] = True
        sheet_names_by_index = [x["name"] for x in detail["sheets"]]
        for el in root.iter():
            if _local(el.tag)[1] == "definedName":
                lsid = el.get("localSheetId")
                scope = None
                if lsid is not None and lsid.isdigit() and int(lsid) < len(sheet_names_by_index):
                    scope = sheet_names_by_index[int(lsid)]
                detail["defined_names"].append([el.get("name", ""), scope, "".join(el.itertext()).strip()])
    for sh in detail["sheets"]:
        raw = data_of.get(sh["part"])
        root = _parse(_without_sheet_rows(raw)) if raw is not None and ctype_of.get(sh["part"]) in _SHEET_TYPES else None
        if root is not None:
            detail["per_sheet"][sh["name"]] = _sheet_detail(root, _rels_of(data_of, sh["part"]))
    for name, data in data_of.items():
        ct = ctype_of.get(name, "")
        if ct in _PART_TYPES["vba"]:
            detail["vba"].append(parts[name])
        elif ct.startswith("image/"):
            detail["images"].append(parts[name])
        elif ct in _PART_TYPES["tables"]:
            r = _parse(data)
            if r is not None:
                detail["tables"].append(r.get("name") or r.get("displayName") or "")
        elif ct in _PART_TYPES["pivot_tables"]:
            r = _parse(data)
            if r is not None:
                detail["pivot_tables"].append(r.get("name", ""))
        elif ct in _PART_TYPES["pivot_caches"]:
            r = _parse(data)
            if r is not None:
                for el in r.iter():
                    if _local(el.tag)[1] == "worksheetSource":
                        detail["pivot_caches"].append(["worksheet", el.get("name", "")])
                        break
                else:
                    detail["pivot_caches"].append(["other", ""])
        elif ct in _PART_TYPES["external_links"]:
            targets = sorted(t for _, t, ext in _rels_of(data_of, name).values() if ext) or [""]
            detail["external_links"].append(targets[0])
        elif ct in _PART_TYPES["charts"]:
            detail["charts"].append(len(_RE_SER.findall(data)))
            detail["chart_refs"].extend(m.decode("utf-8", "replace").strip() for m in _RE_CHART_F.findall(data))
    for k in ("vba", "images", "tables", "pivot_tables", "external_links", "charts"):
        detail[k].sort()
    return detail


def fingerprint(path: str | Path) -> dict:
    """Feature inventory + per-part sha256 of one workbook.

    Returns {"valid_zip": bool, "error": str | None, "features": {key: int}, "parts": {name: sha256},
    "detail": per-item inventory used by lost_features}.
    A non-zip, unreadable or over-cap file gives valid_zip False, an error string, and empty
    features / parts; it never raises."""
    features = {k: 0 for k in FEATURE_KEYS}
    parts: dict[str, str] = {}
    try:
        data_of: dict[str, bytes] = {}
        for name, data in iter_parts(path):
            parts[name] = hashlib.sha256(data).hexdigest()
            data_of[name] = data
        defaults, overrides = _content_types(data_of.get("[Content_Types].xml"))
        ctype_of = {n: overrides.get("/" + n) or defaults.get(n.rsplit(".", 1)[-1].lower() if "." in n else "", "")
                    for n in data_of}
        for name, data in data_of.items():
            ctype = ctype_of[name]
            for key, types in _PART_TYPES.items():
                if ctype in types:
                    features[key] = 1 if key == "vba" else features[key] + 1
            if ctype.startswith("image/"):
                features["images"] += 1
            if ctype in _WORKBOOK_TYPES:
                features["sheets"] += len(_RE_SHEET.findall(data))
                features["defined_names"] += len(_RE_NAME.findall(data))
            elif ctype in _SHEET_TYPES:
                # substring pre-checks (memmem speed): the regexes try every '<' of a multi-million-cell sheet
                if b"sparklineGroup" in data:
                    features["sparkline_groups"] += len(_RE_SPARK.findall(data))
                if b"dataValidation" in data:
                    features["data_validations"] += len(_RE_DV.findall(data))
                if b"conditionalFormatting" in data:
                    features["conditional_formats"] += len(_RE_CF.findall(data))
            elif ctype in _COMMENT_TYPES:
                features["comments"] += len(_RE_COMMENT.findall(data))
    except (zipfile.BadZipFile, PackageTooLarge, OSError, EOFError, NotImplementedError, RuntimeError, ValueError) as exc:
        return {"valid_zip": False, "error": f"{type(exc).__name__}: {exc}",
                "features": {k: 0 for k in FEATURE_KEYS}, "parts": {}, "detail": {}}
    detail = _build_detail(data_of, parts, ctype_of)
    features["merged_cells"] = sum(len(x["merged"]) for x in detail["per_sheet"].values())
    features["hyperlinks"] = sum(len(x["links"]) for x in detail["per_sheet"].values())
    features["protected_sheets"] = sum(1 for x in detail["per_sheet"].values() if x["protected"])
    return {"valid_zip": True, "error": None, "features": features, "parts": parts, "detail": detail}


def _covers(a: list[int], r: list[int]) -> bool:
    """True when rectangle `a` is `r` shifted down or grown: same-or-wider columns, last row not above
    r's, height not smaller."""
    return a[0] <= r[0] and a[2] >= r[2] and a[3] >= r[3] and (a[3] - a[1]) >= (r[3] - r[1])


def _unmatched_ranged(orig: list[dict], new: list[dict]) -> list[dict]:
    """Original items with no counterpart in `new`. Exact (same signature and ranges) pairs are taken
    first; the rest are paired with the nearest candidate that is the item shifted down or grown."""
    pool = list(new)
    rest = []
    for item in orig:
        hit = next((c for c in pool if c["sig"] == item["sig"] and c["rects"] == item["rects"]), None)
        if hit is None:
            rest.append(item)
        else:
            pool.remove(hit)
    lost = []
    for item in rest:
        best, best_d = None, None
        for cand in pool:
            if cand["sig"] != item["sig"]:
                continue
            if not item["rects"]:
                ok, d = not cand["rects"], 0
            else:
                ok = bool(cand["rects"]) and all(any(_covers(a, r) for a in cand["rects"]) for r in item["rects"])
                d = sum(abs(a[1] - r[1]) for r, a in zip(item["rects"], cand["rects"])) if ok else 0
            if ok and (best is None or d < best_d):
                best, best_d = cand, d
        if best is None:
            lost.append(item)
        else:
            pool.remove(best)
    return lost


def _unmatched_multiset(orig: list, new: list) -> list:
    pool = list(new)
    lost = []
    for item in orig:
        if item in pool:
            pool.remove(item)
        else:
            lost.append(item)
    return lost


def _sheet_pairs(before: dict, after: dict) -> dict[str, str]:
    """Original sheet name -> new sheet name: same name first, then same worksheet part (a rename)."""
    new_by_name = {s["name"]: s for s in after["sheets"]}
    new_by_part = {s["part"]: s["name"] for s in after["sheets"] if s["part"]}
    out = {}
    for s in before["sheets"]:
        if s["name"] in new_by_name:
            out[s["name"]] = s["name"]
        elif s["part"] in new_by_part:
            out[s["name"]] = new_by_part[s["part"]]
    return out


_RE_BOOL_CALL = re.compile(r"\b(TRUE|FALSE)\(\)", re.I)
_RE_SHEET_PREFIX = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!")
_RE_ROWNUM = re.compile(r"(\$?[A-Z]{1,3})\$?\d+")


def _norm_ref(text: str, pairs: dict[str, str]) -> str:
    """Normalise a formula for comparison across a legitimate edit: sheet names are mapped through the
    rename pairs, quoting is dropped, and row numbers are removed (insert-rows shifts them)."""
    def sheet(m: re.Match) -> str:
        name = (m.group(1) or "").replace("''", "'") or m.group(2)
        return pairs.get(name, name) + "!"
    text = _RE_BOOL_CALL.sub(lambda m: m.group(1).upper(), text.strip())  # TRUE() == TRUE
    return _RE_ROWNUM.sub(r"\1", _RE_SHEET_PREFIX.sub(sheet, text))


def _entry(feature: str, total: int, lost: list[str]) -> dict:
    return {"feature": feature, "before": total, "after": max(total - len(lost), 0), "detail": lost[:20]}


def lost_features(before: dict, after: dict) -> list[dict]:
    """What was lost going from before to after: [{feature, before, after, detail?}] (additions are fine).

    Count-only categories (sheets, drawings, comments) compare counts. Every other category compares
    items (see the rules in the section comment above); `before` is the original's item count, `after`
    how many of them survive, `detail` names what is gone, with sheet and range where one exists."""
    out: list[dict] = []
    for k in ("sheets", "drawings", "comments"):
        a, b = int(before["features"].get(k, 0)), int(after["features"].get(k, 0))
        if b < a:
            out.append({"feature": k, "before": a, "after": b})
    bd, ad = before.get("detail") or {}, after.get("detail") or {}
    if not bd or not ad:
        return out
    pairs = _sheet_pairs(bd, ad)

    def add(feature: str, total: int, lost: list[str]) -> None:
        if lost:
            out.append(_entry(feature, total, lost))

    # whole-file byte categories
    add("vba", len(bd["vba"]), [f"vbaProject.bin changed or missing (sha256 {h[:12]} not in the output)"
                                for h in _unmatched_multiset(bd["vba"], ad["vba"])])
    add("images", len(bd["images"]), [f"image {h[:12]} gone" for h in _unmatched_multiset(bd["images"], ad["images"])])
    add("tables", len(bd["tables"]), [f"table {n} missing" for n in _unmatched_multiset(bd["tables"], ad["tables"])])
    add("pivot_tables", len(bd["pivot_tables"]),
        [f"pivot table {n} missing" for n in _unmatched_multiset(bd["pivot_tables"], ad["pivot_tables"])])
    add("pivot_caches", len(bd["pivot_caches"]),
        [f"pivot cache on {c[1] or c[0]} missing" for c in _unmatched_multiset(bd["pivot_caches"], ad["pivot_caches"])])
    add("external_links", len(bd["external_links"]),
        [f"external link {t} missing" for t in _unmatched_multiset(bd["external_links"], ad["external_links"])])
    pool = sorted(ad["charts"])
    lost_charts = []
    for n in sorted(bd["charts"]):
        hit = next((c for c in pool if c >= n), None)
        if hit is None:
            lost_charts.append(f"chart with {n} series missing")
        else:
            pool.remove(hit)
    add("charts", len(bd["charts"]), lost_charts)

    # protection
    prot_lost = []
    if bd["workbook_protected"] and not ad["workbook_protected"]:
        prot_lost.append("workbook protection removed")
    n_prot = int(bool(bd["workbook_protected"]))
    for name, sd in bd["per_sheet"].items():
        if sd["protected"]:
            n_prot += 1
            new = pairs.get(name)
            if new is not None and not (ad["per_sheet"].get(new) or {}).get("protected"):
                prot_lost.append(f"sheet protection removed on {name}")
    add("protection", n_prot, prot_lost)

    # defined names: (name, scope) must exist and its formula must survive after sheet renames / row shifts
    new_names = [list(x) for x in ad["defined_names"]]
    lost_names = []
    for n, sc, f in bd["defined_names"]:
        scope = pairs.get(sc, sc) if sc is not None else None
        cand = next((x for x in new_names if x[0] == n and x[1] == scope), None)
        if cand is None:
            lost_names.append(f"defined name {n}" + (f" (scope {sc})" if sc else "") + " missing")
            continue
        new_names.remove(cand)
        if _norm_ref(f, pairs) != _norm_ref(cand[2], {}):
            lost_names.append(f"defined name {n} changed ({f} -> {cand[2]})")
    add("defined_names", len(bd["defined_names"]), lost_names)

    # chart series / label references: sheet and columns must follow (rows may shift)
    pool = [_norm_ref(r, {}) for r in ad["chart_refs"]]
    lost_refs = []
    for r in bd["chart_refs"]:
        want = _norm_ref(r, pairs)
        if want in pool:
            pool.remove(want)
        else:
            lost_refs.append(f"chart reference {r} changed or gone")
    add("chart_references", len(bd["chart_refs"]), lost_refs)

    # per-sheet range items and link multisets
    labels = {"dv": ("data_validations", "data validation"), "cf": ("conditional_formats", "conditional format"),
              "merged": ("merged_cells", "merged range"), "spark": ("sparkline_groups", "sparkline group")}
    for key, (feature, label) in labels.items():
        total, lost = 0, []
        for name, sd in bd["per_sheet"].items():
            total += len(sd[key])
            new = pairs.get(name)
            if new is None or new not in ad["per_sheet"]:
                continue
            for it in _unmatched_ranged(sd[key], ad["per_sheet"][new][key]):
                lost.append(f"{name}: {label} {'/'.join(map(str, it['sig'])) or ''} on {it['ref']} removed".replace("  ", " "))
        add(feature, total, lost)
    total, lost = 0, []
    for name, sd in bd["per_sheet"].items():
        total += len(sd["links"])
        new = pairs.get(name)
        if new is None or new not in ad["per_sheet"]:
            continue
        for lk in _unmatched_multiset(sd["links"], ad["per_sheet"][new]["links"]):
            lost.append(f"{name}: hyperlink {lk[1] if len(lk) > 1 else '(internal)'} removed")
    add("hyperlinks", total, lost)
    return out


def untouchable_changed(before_parts: dict[str, str], after_parts: dict[str, str]) -> list[str]:
    """Never-rewritten parts that differ between the two part maps (changed or removed)."""
    out = []
    for name, digest in sorted(before_parts.items()):
        if not any(fnmatch.fnmatchcase(name, p) for p in UNTOUCHABLE):
            continue
        if name not in after_parts or after_parts[name] != digest:
            out.append(name)
    return out
