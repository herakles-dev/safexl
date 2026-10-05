#!/usr/bin/env python3
"""inventory.py — feature inventory / preservation checker for OOXML (.xlsx/.xlsm) workbooks.

Reads the zip container directly (zipfile + xml.etree, stdlib only — no openpyxl
or any other object-model library) and reports, per file:

  - VBA project presence + sha256 of xl/vbaProject.bin
  - Charts: count + each series' category/value/name range refs
  - Pivot tables + pivot caches: count + cache source ranges
  - Excel Tables (ListObjects): name + ref
  - Conditional formatting: classic + x14 (extLst) blocks, counts + sqrefs
  - Data validation: classic + x14 blocks, counts + sqrefs
  - Defined names: name -> formula
  - Images/media: count + sha256 of each part
  - Comments (legacy + threaded): count
  - Sparkline groups (x14): count + per-group sqrefs/refs
  - Slicers: count
  - External links: count + targets
  - Protection: workbook-level + per-sheet flags
  - Hyperlinks: count + per-link ref/target
  - Merged cell ranges: per sheet
  - Frozen panes: per sheet
  - Sheet names, in workbook-declared order
  - Per-part sha256 (every member of the zip)

This module is deliberately conservative: a missing/malformed part yields an
empty/False result for that feature rather than raising, so a corrupted
"after edit" file still produces a diffable inventory instead of a crash.
The one exception is that the zip itself must open — a workbook that isn't a
valid zip is a hard failure the grader wants to see.

CLI:
    python3 inventory.py FILE.xlsx [--json]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


# --------------------------------------------------------------------------
# small XML helpers (namespace-wildcard so we don't care which prefix a
# producer used for x14/xm/etc — we match on local tag name only)
# --------------------------------------------------------------------------

def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _findall(elem, local_tag):
    """All descendants (any namespace) whose local name == local_tag."""
    return elem.findall(f".//{{*}}{local_tag}")


def _find(elem, local_tag):
    return elem.find(f".//{{*}}{local_tag}")


def _children(elem, local_tag):
    """Direct children (any namespace) whose local name == local_tag."""
    return [c for c in list(elem) if _local(c.tag) == local_tag]


def _parse_xml(zf: zipfile.ZipFile, part: str):
    try:
        data = zf.read(part)
    except KeyError:
        return None
    try:
        return ET.fromstring(data)
    except ET.ParseError:
        return None


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# --------------------------------------------------------------------------
# relationship helpers
# --------------------------------------------------------------------------

def _rels_for(zf: zipfile.ZipFile, part: str) -> dict:
    """Return {rId: target_path} for the .rels file belonging to `part`."""
    if "/" in part:
        directory, name = part.rsplit("/", 1)
        rels_part = f"{directory}/_rels/{name}.rels"
    else:
        rels_part = f"_rels/{part}.rels"
    root = _parse_xml(zf, rels_part)
    out = {}
    if root is None:
        return out
    base_dir = part.rsplit("/", 1)[0] if "/" in part else ""
    for rel in _children(root, "Relationship"):
        rid = rel.get("Id")
        target = rel.get("Target", "")
        mode = rel.get("TargetMode", "Internal")
        if mode == "External":
            out[rid] = target
        else:
            # resolve relative to the part's directory
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                parts = (base_dir.split("/") if base_dir else []) + target.split("/")
                stack = []
                for p in parts:
                    if p == "..":
                        if stack:
                            stack.pop()
                    elif p and p != ".":
                        stack.append(p)
                resolved = "/".join(stack)
            out[rid] = resolved
    return out


# --------------------------------------------------------------------------
# sheet name / order + sheetN.xml part mapping
# --------------------------------------------------------------------------

def _sheet_map(zf: zipfile.ZipFile):
    """Return (ordered list of sheet names, {sheet_name: part_path}, {sheet_name: sheetId})."""
    wb = _parse_xml(zf, "xl/workbook.xml")
    if wb is None:
        return [], {}, {}
    rels = _rels_for(zf, "xl/workbook.xml")
    names = []
    part_of = {}
    sheet_id_of = {}
    sheets_el = _find(wb, "sheets")
    if sheets_el is None:
        return [], {}, {}
    for sheet in _children(sheets_el, "sheet"):
        name = sheet.get("name", "")
        rid = sheet.get(f"{{{REL_NS}}}id")
        names.append(name)
        sheet_id_of[name] = sheet.get("sheetId")
        target = rels.get(rid)
        if target:
            part_of[name] = target
    return names, part_of, sheet_id_of


# --------------------------------------------------------------------------
# feature extractors
# --------------------------------------------------------------------------

def _vba(zf, names):
    part = "xl/vbaProject.bin"
    present = part in names
    sha = _sha256_bytes(zf.read(part)) if present else None
    return {"present": present, "sha256": sha}


def _charts(zf, names):
    chart_parts = sorted(n for n in names if n.startswith("xl/charts/chart") and n.endswith(".xml"))
    charts = []
    for part in chart_parts:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        series_list = []
        for ser in _findall(root, "ser"):
            def ref_of(container_tag):
                container = None
                for c in list(ser):
                    if _local(c.tag) == container_tag:
                        container = c
                        break
                if container is None:
                    return None
                f = _find(container, "f")
                return f.text if f is not None else None

            series_list.append(
                {
                    "name_ref": ref_of("tx"),
                    "cat_ref": ref_of("cat"),
                    "val_ref": ref_of("val"),
                }
            )
        charts.append({"part": part, "n_series": len(series_list), "series": series_list})
    return {"count": len(charts), "charts": charts}


def _pivot_tables(zf, names):
    pt_parts = sorted(n for n in names if n.startswith("xl/pivotTables/") and n.endswith(".xml"))
    tables = []
    for part in pt_parts:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        loc = _find(root, "location")
        tables.append(
            {
                "part": part,
                "name": root.get("name"),
                "cache_id": root.get("cacheId"),
                "location_ref": loc.get("ref") if loc is not None else None,
            }
        )
    return {"count": len(tables), "tables": tables}


def _pivot_caches(zf, names):
    pc_parts = sorted(
        n for n in names if n.startswith("xl/pivotCache/pivotCacheDefinition") and n.endswith(".xml")
    )
    caches = []
    for part in pc_parts:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        src = _find(root, "cacheSource")
        entry = {"part": part, "source_type": src.get("type") if src is not None else None}
        if src is not None:
            ws = _find(src, "worksheetSource")
            if ws is not None:
                entry["source_sheet"] = ws.get("sheet")
                entry["source_ref"] = ws.get("ref")
                entry["source_name"] = ws.get("name")
        caches.append(entry)
    return {"count": len(caches), "caches": caches}


def _tables(zf, names, sheet_of_part):
    t_parts = sorted(n for n in names if n.startswith("xl/tables/") and n.endswith(".xml"))
    tables = []
    for part in t_parts:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        tables.append(
            {
                "part": part,
                "name": root.get("name"),
                "ref": root.get("ref"),
                "sheet": sheet_of_part.get(_owning_sheet_part(zf, part)),
            }
        )
    return {"count": len(tables), "tables": tables}


def _owning_sheet_part(zf, table_part):
    """Given xl/tables/tableN.xml, find the worksheet part that references it."""
    # search every worksheet's rels for a relationship whose target == table_part
    for name in zf.namelist():
        if name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
            sheet_part = "xl/worksheets/" + name.rsplit("/", 1)[-1][: -len(".rels")]
            rels = _rels_for(zf, sheet_part)
            if table_part in rels.values():
                return sheet_part
    return None


def _cf_and_dv(zf, sheet_names, part_of):
    """Classic + x14 conditional formatting and data validation, per sheet."""
    cf_classic, cf_x14 = [], []
    dv_classic, dv_x14 = [], []
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        # classic CF: direct <conditionalFormatting> elements (not inside extLst)
        for cf in _children(root, "conditionalFormatting"):
            types = sorted({r.get("type") for r in _children(cf, "cfRule") if r.get("type")})
            cf_classic.append({"sheet": sheet, "sqref": cf.get("sqref"), "rule_types": types})
        # classic DV
        dvs_el = None
        for c in list(root):
            if _local(c.tag) == "dataValidations":
                dvs_el = c
                break
        if dvs_el is not None:
            for dv in _children(dvs_el, "dataValidation"):
                dv_classic.append(
                    {"sheet": sheet, "sqref": dv.get("sqref"), "type": dv.get("type")}
                )
        # x14 (extLst) CF + DV
        ext_lst = None
        for c in list(root):
            if _local(c.tag) == "extLst":
                ext_lst = c
                break
        if ext_lst is not None:
            for x14cf in _findall(ext_lst, "conditionalFormatting"):
                sqref_el = _find(x14cf, "sqref")
                types = sorted({r.get("type") for r in _findall(x14cf, "cfRule") if r.get("type")})
                cf_x14.append(
                    {
                        "sheet": sheet,
                        "sqref": sqref_el.text if sqref_el is not None else None,
                        "rule_types": types,
                    }
                )
            for x14dv in _findall(ext_lst, "dataValidation"):
                sqref_el = _find(x14dv, "sqref")
                dv_x14.append(
                    {
                        "sheet": sheet,
                        "sqref": sqref_el.text if sqref_el is not None else None,
                        "type": x14dv.get("type"),
                    }
                )
    return (
        {"count": len(cf_classic), "blocks": cf_classic},
        {"count": len(cf_x14), "blocks": cf_x14},
        {"count": len(dv_classic), "blocks": dv_classic},
        {"count": len(dv_x14), "blocks": dv_x14},
    )


def _defined_names(zf):
    wb = _parse_xml(zf, "xl/workbook.xml")
    out = {}
    if wb is None:
        return out
    dns = None
    for c in list(wb):
        if _local(c.tag) == "definedNames":
            dns = c
            break
    if dns is None:
        return out
    for dn in _children(dns, "definedName"):
        name = dn.get("name", "")
        if dn.get("localSheetId") is not None:
            name = f"{name}#localSheetId={dn.get('localSheetId')}"
        out[name] = (dn.text or "").strip()
    return out


def _images(zf, names):
    media = sorted(n for n in names if n.startswith("xl/media/"))
    out = [{"part": n, "sha256": _sha256_bytes(zf.read(n))} for n in media]
    return {"count": len(out), "media": out}


def _comments(zf, names):
    legacy = sorted(n for n in names if n.startswith("xl/comments") and n.endswith(".xml"))
    threaded = sorted(n for n in names if "threadedComment" in n and n.endswith(".xml"))
    count = 0
    for part in legacy:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        count += len(_findall(root, "comment"))
    threaded_count = 0
    for part in threaded:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        threaded_count += len(_findall(root, "threadedComment"))
    return {
        "count": count + threaded_count,
        "legacy_count": count,
        "threaded_count": threaded_count,
        "parts": legacy + threaded,
    }


def _sparklines(zf, sheet_names, part_of):
    groups = []
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        for grp in _findall(root, "sparklineGroup"):
            refs = []
            for spk in _findall(grp, "sparkline"):
                f_el = _find(spk, "f")
                sqref_el = _find(spk, "sqref")
                refs.append(
                    {
                        "f": f_el.text if f_el is not None else None,
                        "sqref": sqref_el.text if sqref_el is not None else None,
                    }
                )
            groups.append({"sheet": sheet, "type": grp.get("type"), "sparklines": refs})
    return {"count": len(groups), "groups": groups}


def _slicers(zf, names):
    slicer_parts = sorted(n for n in names if n.startswith("xl/slicers/") and n.endswith(".xml"))
    cache_parts = sorted(n for n in names if n.startswith("xl/slicerCaches/") and n.endswith(".xml"))
    total = 0
    detail = []
    for part in slicer_parts:
        root = _parse_xml(zf, part)
        if root is None:
            continue
        slicer_els = _findall(root, "slicer")
        total += len(slicer_els)
        for s in slicer_els:
            detail.append({"part": part, "name": s.get("name"), "cache": s.get("cache")})
    return {"count": total, "slicer_parts": slicer_parts, "slicer_cache_parts": cache_parts, "detail": detail}


def _external_links(zf, names):
    ext_parts = sorted(
        n for n in names if n.startswith("xl/externalLinks/externalLink") and n.endswith(".xml")
    )
    out = []
    for part in ext_parts:
        rels = _rels_for(zf, part)
        target = next(iter(rels.values()), None)
        root = _parse_xml(zf, part)
        book = _find(root, "externalBook") if root is not None else None
        out.append({"part": part, "target": target, "has_externalBook": book is not None})
    return {"count": len(out), "links": out}


def _protection(zf, sheet_names, part_of):
    wb = _parse_xml(zf, "xl/workbook.xml")
    wb_prot = None
    if wb is not None:
        el = None
        for c in list(wb):
            if _local(c.tag) == "workbookProtection":
                el = c
                break
        if el is not None:
            wb_prot = dict(el.attrib)
    sheets = {}
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        el = None
        for c in list(root):
            if _local(c.tag) == "sheetProtection":
                el = c
                break
        if el is not None:
            sheets[sheet] = dict(el.attrib)
    return {"workbook": wb_prot, "sheets": sheets}


def _hyperlinks(zf, sheet_names, part_of):
    out = []
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        rels = _rels_for(zf, part)
        hls_el = None
        for c in list(root):
            if _local(c.tag) == "hyperlinks":
                hls_el = c
                break
        if hls_el is None:
            continue
        for hl in _children(hls_el, "hyperlink"):
            rid = hl.get(f"{{{REL_NS}}}id")
            out.append(
                {
                    "sheet": sheet,
                    "ref": hl.get("ref"),
                    "target": rels.get(rid) if rid else hl.get("location"),
                    "internal": rid is None,
                }
            )
    return {"count": len(out), "links": out}


def _merged_cells(zf, sheet_names, part_of):
    by_sheet = {}
    total = 0
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        mcs_el = None
        for c in list(root):
            if _local(c.tag) == "mergeCells":
                mcs_el = c
                break
        if mcs_el is None:
            continue
        ranges = [mc.get("ref") for mc in _children(mcs_el, "mergeCell")]
        if ranges:
            by_sheet[sheet] = ranges
            total += len(ranges)
    return {"count": total, "by_sheet": by_sheet}


def _frozen_panes(zf, sheet_names, part_of):
    sheets = []
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        for pane in _findall(root, "pane"):
            state = pane.get("state")
            if state in ("frozen", "frozenSplit"):
                sheets.append(
                    {
                        "sheet": sheet,
                        "state": state,
                        "topLeftCell": pane.get("topLeftCell"),
                        "xSplit": pane.get("xSplit"),
                        "ySplit": pane.get("ySplit"),
                    }
                )
    return {"count": len(sheets), "sheets": sheets}


_CROSS_SHEET_RE = None


def _cross_sheet_formula_refs(zf, sheet_names, part_of):
    """Formulas whose text contains a sheet-qualified reference (Sheet1!A1,
    'My Sheet'!$B$2, or an external-book form like '[1]Sheet'!A1).

    This is a lexical scan of <f> formula text, not a real formula parser —
    it is good enough to flag "this workbook has cross-sheet wiring that a
    naive row-insert could silently break", which is what the benchmark
    corpus needs it for.
    """
    import re

    global _CROSS_SHEET_RE
    if _CROSS_SHEET_RE is None:
        _CROSS_SHEET_RE = re.compile(
            r"(?P<ref>(?:'[^']+'|\[\d+\][A-Za-z_][A-Za-z0-9_. ]*|[A-Za-z_][A-Za-z0-9_.]*))!"
            r"\$?[A-Z]{1,3}\$?[0-9]+"
        )
    out = []
    for sheet in sheet_names:
        part = part_of.get(sheet)
        if not part:
            continue
        root = _parse_xml(zf, part)
        if root is None:
            continue
        for f_el in _findall(root, "f"):
            text = f_el.text or ""
            if _CROSS_SHEET_RE.search(text):
                out.append({"sheet": sheet, "formula": text})
    return {"count": len(out), "formulas": out}


def _drawings_and_media_refs(zf, names):
    drawing_parts = sorted(
        n for n in names if n.startswith("xl/drawings/drawing") and n.endswith(".xml")
    )
    return {"count": len(drawing_parts), "parts": drawing_parts}


def _part_sha256(zf):
    out = {}
    for info in zf.infolist():
        if info.is_dir():
            continue
        out[info.filename] = _sha256_bytes(zf.read(info.filename))
    return out


# --------------------------------------------------------------------------
# top-level entry point
# --------------------------------------------------------------------------

def inventory(path: str) -> dict:
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        sheet_names, part_of, sheet_id_of = _sheet_map(zf)
        sheet_of_part = {v: k for k, v in part_of.items()}

        cf_classic, cf_x14, dv_classic, dv_x14 = _cf_and_dv(zf, sheet_names, part_of)

        result = {
            "file": path,
            "size_bytes": __import__("os").path.getsize(path),
            "sheet_names": sheet_names,
            "vba": _vba(zf, names),
            "charts": _charts(zf, names),
            "pivot_tables": _pivot_tables(zf, names),
            "pivot_caches": _pivot_caches(zf, names),
            "tables": _tables(zf, names, sheet_of_part),
            "conditional_formatting": {"classic": cf_classic, "x14": cf_x14},
            "data_validations": {"classic": dv_classic, "x14": dv_x14},
            "defined_names": _defined_names(zf),
            "images": _images(zf, names),
            "comments": _comments(zf, names),
            "sparklines": _sparklines(zf, sheet_names, part_of),
            "slicers": _slicers(zf, names),
            "external_links": _external_links(zf, names),
            "protection": _protection(zf, sheet_names, part_of),
            "hyperlinks": _hyperlinks(zf, sheet_names, part_of),
            "merged_cells": _merged_cells(zf, sheet_names, part_of),
            "frozen_panes": _frozen_panes(zf, sheet_names, part_of),
            "cross_sheet_formulas": _cross_sheet_formula_refs(zf, sheet_names, part_of),
            "drawings": _drawings_and_media_refs(zf, names),
            "part_sha256": _part_sha256(zf),
        }
        return result


def main(argv=None):
    parser = argparse.ArgumentParser(description="Feature inventory for .xlsx/.xlsm workbooks")
    parser.add_argument("file", help="path to .xlsx/.xlsm file")
    parser.add_argument("--json", action="store_true", help="print full JSON (default: summary)")
    args = parser.parse_args(argv)

    result = inventory(args.file)

    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0

    print(f"file: {result['file']} ({result['size_bytes']} bytes)")
    print(f"sheets ({len(result['sheet_names'])}): {result['sheet_names']}")
    print(f"vba present: {result['vba']['present']}" + (f" sha256={result['vba']['sha256']}" if result['vba']['present'] else ""))
    print(f"charts: {result['charts']['count']}")
    print(f"pivot tables: {result['pivot_tables']['count']}  pivot caches: {result['pivot_caches']['count']}")
    print(f"tables (ListObjects): {result['tables']['count']}")
    print(f"conditional formatting: classic={result['conditional_formatting']['classic']['count']} x14={result['conditional_formatting']['x14']['count']}")
    print(f"data validations: classic={result['data_validations']['classic']['count']} x14={result['data_validations']['x14']['count']}")
    print(f"defined names: {len(result['defined_names'])}")
    print(f"images/media: {result['images']['count']}")
    print(f"comments: {result['comments']['count']}")
    print(f"sparkline groups: {result['sparklines']['count']}")
    print(f"slicers: {result['slicers']['count']}")
    print(f"external links: {result['external_links']['count']}")
    print(f"hyperlinks: {result['hyperlinks']['count']}")
    print(f"merged cells: {result['merged_cells']['count']}")
    print(f"frozen panes: {result['frozen_panes']['count']}")
    print(f"cross-sheet formulas: {result['cross_sheet_formulas']['count']}")
    print(f"drawings: {result['drawings']['count']}")
    print(f"workbook protection: {result['protection']['workbook'] is not None}")
    print(f"sheet protection on: {list(result['protection']['sheets'].keys())}")
    print(f"parts: {len(result['part_sha256'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
