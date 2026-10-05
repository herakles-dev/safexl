#!/usr/bin/env python3
"""grade.py — deterministic grader for xlsx-bench.

Grades one run (an edited workbook produced by either arm A — Anthropic's
xlsx skill — or arm B — SafeXL) against a task definition from tasks.json
(schema: TASK_SCHEMA.md). Never looks at which arm produced a file while
grading: `arm`/`rep` are attached to the result dict only by the CLI, after
`grade()` returns.

Three independent booleans, ANDed into `pass`:
  - correct:   every `expect` check in the task passes (cell/table/chart/
               CF/DV/defined-name/sheet-name assertions, checked against a
               LibreOffice-headless recalculation of the OUTPUT file).
  - preserved: every feature class named in the task's `preserve` list
               still has the same identity in the output as in the
               original (per inventory.py), except for the precise deltas
               the task's own `expect` asserts (e.g. a table ref that
               legitimately moved because of an inserted row).
  - valid:     the output is a well-formed OOXML package (unzips, every
               XML part parses, [Content_Types].xml covers every part),
               LibreOffice can open/convert it, and no NEW #REF!/#NAME?/
               #VALUE!/#DIV/0!/#N/A/#NULL!/#NUM! cell appeared that wasn't
               already present (same cell) in the original's recalc.

CLI:
    python3 grade.py --tasks tasks.json --runs runs/ --out results.jsonl [--summary]
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

# --------------------------------------------------------------------------
# load inventory.py by path (never a second parser — reuse it exactly)
# --------------------------------------------------------------------------

_HERE = Path(__file__).resolve().parent
# research/xlsx-bench -> research -> repo root. Matches runner.py's
# REPO_ROOT exactly: task["workbook"] strings are written relative to the
# repo root (e.g. "research/xlsx-bench/corpus/foo.xlsx" or a fixture path
# like "research/safexl-moat/fixtures/foo.xlsm"), not to this script.
REPO_ROOT = _HERE.parent.parent
_inv_spec = importlib.util.spec_from_file_location("xlsx_bench_inventory", _HERE / "inventory.py")
_inv_mod = importlib.util.module_from_spec(_inv_spec)
sys.modules.setdefault("xlsx_bench_inventory", _inv_mod)
_inv_spec.loader.exec_module(_inv_mod)
inventory = _inv_mod.inventory

try:
    import openpyxl
except ImportError as e:  # pragma: no cover
    raise SystemExit("grade.py requires openpyxl (pip install openpyxl)") from e


ERROR_STRINGS = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NULL!", "#NUM!"}

# Feature classes recognized in a task's `preserve` list (== inventory.py's
# top-level keys that carry identity, minus the always-checked structural
# ones like sheet_names/part_sha256/frozen_panes/slicers/drawings which are
# not part of the TASK_SCHEMA preserve vocabulary).
KNOWN_PRESERVE_CLASSES = {
    "vba", "images", "charts", "pivot_tables", "pivot_caches", "comments",
    "hyperlinks", "protection", "sparklines", "external_links",
    "merged_cells", "conditional_formatting", "data_validations",
    "defined_names", "tables",
}

# A fresh, unique LibreOffice user profile per recalc call (per point 2 of
# the task spec) means parallel calls never collide on lock files or config.
_RECALC_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""
# OOXMLRecalcMode/ODFRecalcMode default to 1 ("prompt user"), which under
# --headless silently means "trust the cached value, don't recalculate".
# Setting both to 0 ("always recalculate") is the only way to get real
# recalculated values out of `soffice --convert-to xlsx` instead of the
# stale cached <v> the editor happened to leave behind. Verified empirically
# against corpus/simple-monthly-budget.xlsx before relying on it here.

_RECALC_CACHE: dict[str, dict] = {}


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _run_libreoffice_recalc(path: Path, timeout: int = 60) -> dict:
    """Convert a COPY of `path` to xlsx in an isolated temp dir/profile,
    forcing a full recalculation, then read back computed values via
    openpyxl(data_only=True). Never touches `path` itself or its directory
    (no lock files, no mutation) — safe to call on corpus files or run
    outputs alike.
    """
    with tempfile.TemporaryDirectory(prefix="xlsxbench-recalc-") as tmp_str:
        tmp = Path(tmp_str)
        work_copy = tmp / path.name
        shutil.copy2(path, work_copy)

        profile_dir = tmp / "profile"
        user_dir = profile_dir / "user"
        user_dir.mkdir(parents=True, exist_ok=True)
        (user_dir / "registrymodifications.xcu").write_text(_RECALC_XCU, encoding="utf-8")

        outdir = tmp / "out"
        outdir.mkdir()

        cmd = [
            "soffice",
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--norestore",
            "--nologo",
            "--nolockcheck",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to", "xlsx",
            "--outdir", str(outdir),
            str(work_copy),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout, text=True)
        except subprocess.TimeoutExpired:
            return {
                "ok": False,
                "reason": f"soffice timed out after {timeout}s",
                "values": {},
                "error_cells": set(),
            }
        except FileNotFoundError as e:
            return {
                "ok": False,
                "reason": f"soffice not found: {e}",
                "values": {},
                "error_cells": set(),
            }

        converted = outdir / (work_copy.stem + ".xlsx")
        if not converted.exists():
            stderr = (proc.stderr or "").strip()[:500]
            return {
                "ok": False,
                "reason": f"soffice conversion failed (rc={proc.returncode}): {stderr}",
                "values": {},
                "error_cells": set(),
            }

        try:
            wb = openpyxl.load_workbook(str(converted), data_only=True, read_only=True)
        except Exception as e:
            return {
                "ok": False,
                "reason": f"openpyxl couldn't read the recalculated file: {e}",
                "values": {},
                "error_cells": set(),
            }

        values: dict[tuple[str, str], object] = {}
        error_cells: set[tuple[str, str]] = set()
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        key = (ws.title, cell.coordinate)
                        values[key] = cell.value
                        if isinstance(cell.value, str) and cell.value.strip() in ERROR_STRINGS:
                            error_cells.add(key)
        finally:
            wb.close()

        return {"ok": True, "reason": None, "values": values, "error_cells": error_cells}


def _recalc_workbook(path: Path, timeout: int = 60) -> dict:
    path = Path(path)
    try:
        key = _sha256_file(path)
    except OSError as e:
        return {"ok": False, "reason": f"couldn't read file: {e}", "values": {}, "error_cells": set()}
    cached = _RECALC_CACHE.get(key)
    if cached is not None:
        return cached
    result = _run_libreoffice_recalc(path, timeout=timeout)
    _RECALC_CACHE[key] = result
    return result


# --------------------------------------------------------------------------
# validity: structural checks on the OOXML package itself
# --------------------------------------------------------------------------

def _check_content_types(zf: zipfile.ZipFile, names: list[str]) -> tuple[bool, str]:
    CT_NS_PREFIX = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    try:
        root = ET.fromstring(zf.read("[Content_Types].xml"))
    except KeyError:
        return False, "[Content_Types].xml missing"
    except ET.ParseError as e:
        return False, f"[Content_Types].xml malformed: {e}"

    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for el in root:
        tag = el.tag[len(CT_NS_PREFIX):] if el.tag.startswith(CT_NS_PREFIX) else el.tag.rsplit("}", 1)[-1]
        if tag == "Default":
            defaults[(el.get("Extension") or "").lower()] = el.get("ContentType")
        elif tag == "Override":
            overrides[el.get("PartName") or ""] = el.get("ContentType")

    missing = []
    for n in names:
        if n == "[Content_Types].xml" or n.endswith("/"):
            continue
        partname = "/" + n
        if partname in overrides:
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n.rsplit("/", 1)[-1] else ""
        if ext in defaults:
            continue
        missing.append(n)

    if missing:
        return False, f"{len(missing)} part(s) not covered by [Content_Types].xml: {missing}"
    return True, f"{len(overrides)} override(s) + {len(defaults)} default extension(s) cover all {len(names)} part(s)"


def _check_validity_structure(output: Path) -> tuple[list[dict], bool]:
    checks = []
    try:
        zf = zipfile.ZipFile(output)
    except (zipfile.BadZipFile, FileNotFoundError, OSError) as e:
        checks.append({"name": "unzip", "ok": False, "detail": str(e)})
        return checks, False

    bad_member = zf.testzip()
    if bad_member is not None:
        checks.append({"name": "unzip", "ok": False, "detail": f"corrupt member: {bad_member}"})
        return checks, False
    names = zf.namelist()
    checks.append({"name": "unzip", "ok": True, "detail": f"{len(names)} part(s)"})

    bad_xml = []
    for n in names:
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                ET.fromstring(zf.read(n))
            except ET.ParseError as e:
                bad_xml.append(f"{n}: {e}")
    xml_ok = not bad_xml
    checks.append({
        "name": "xml_parts_parse",
        "ok": xml_ok,
        "detail": "all XML parts parsed" if xml_ok else "; ".join(bad_xml),
    })

    ct_ok, ct_detail = _check_content_types(zf, names)
    checks.append({"name": "content_types_coverage", "ok": ct_ok, "detail": ct_detail})

    rel_ok, rel_detail = _check_relationship_targets(zf, names)
    checks.append({"name": "relationship_targets", "ok": rel_ok, "detail": rel_detail})

    return checks, xml_ok and ct_ok and rel_ok


_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _check_relationship_targets(zf: zipfile.ZipFile, names: list[str]) -> tuple[bool, str]:
    """Every internal Relationship must point at a part that exists. A dangling
    one (e.g. a calcChain.xml dropped but still referenced) is an Excel repair
    trigger that LibreOffice silently tolerates, so opening in LibreOffice
    alone can't catch it. TargetMode="External" targets are skipped."""
    present = set(names)
    dangling = []
    for rels in (n for n in names if n.endswith(".rels")):
        try:
            root = ET.fromstring(zf.read(rels))
        except ET.ParseError:
            continue  # already reported by xml_parts_parse
        base = posixpath.dirname(posixpath.dirname(rels))
        for rel in root.iter(f"{_RELS_NS}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if part not in present:
                dangling.append(f"{rels} {rel.get('Id')} -> {part}")
    if dangling:
        return False, f"{len(dangling)} dangling relationship(s): {dangling}"
    return True, "every internal relationship target exists"


# --------------------------------------------------------------------------
# correctness: `expect` checks, against a recalculated OUTPUT
# --------------------------------------------------------------------------

def _col_of(ref: str) -> str:
    return re.match(r"[A-Z]+", ref).group(0)


def _check_column_values(specs: list[dict], out_recalc: dict, orig_recalc: dict | None) -> list[dict]:
    """Position-independent check for a NEW column in a plain range: some column
    of `sheet` must hold `rows` ({row: value}) within `tol` after recalc, and it
    must not already have held those values in the original. Where the column
    goes is the editor's call (e.g. Claude put one at EA to keep a macro's fixed
    column reads intact); only the values are the task."""
    results = []
    orig_values = (orig_recalc or {}).get("values", {})
    for spec in specs:
        sheet, tol, use_abs = spec["sheet"], spec.get("tol", 0.0), spec.get("abs", False)
        # `any_of`: the prompt allows more than one right answer (e.g. a row holding
        # four dates asked for "the days between the two dates"); any one passes.
        alternatives = [{int(r): float(v) for r, v in alt.items()} for alt in spec.get("any_of", [spec.get("rows", {})])]
        cols = {_col_of(ref) for (s, ref) in out_recalc["values"] if s == sheet}

        def matches(values: dict, col: str, rows: dict) -> bool:
            for r, want in rows.items():
                got = values.get((sheet, f"{col}{r}"))
                try:
                    got = abs(float(got)) if use_abs else float(got)
                    if abs(got - want) > tol:
                        return False
                except (TypeError, ValueError):
                    return False
            return True

        found, rows = [], alternatives[0]
        for alt in alternatives:
            found = sorted((c for c in cols if matches(out_recalc["values"], c, alt) and not matches(orig_values, c, alt)),
                           key=lambda c: (len(c), c))
            if found:
                rows = alt
                break
        if not found and len(alternatives) > 1:
            rows = {"any_of": alternatives}
        label = spec.get("label", f"{sheet} new column")
        detail = (f"{label}: found in column {found[0]} ({', '.join(f'row {r}={v}' for r, v in rows.items())})"
                  if found else f"{label}: no new column holds {rows} (tol {tol})")
        results.append({"name": f"column_values:{sheet}", "ok": bool(found), "detail": detail})
    return results


_LOCAL_REF = re.compile(r"(?<![!A-Za-z0-9_$.'])\$?([A-Z]{1,3})\$?([0-9]+)(?![0-9(A-Za-z_])")


def _formula_reaches(ref: str, needle: str, formulas: dict[str, str], depth: int = 3) -> bool:
    """True when the cell's formula contains `needle`, or reaches it through
    same-sheet cells it references (up to `depth` hops). =B6/B5 with
    B5 = Summary_report!C33 is as live a link to the source as the one-line form."""
    f = formulas.get(ref)
    if not f:
        return False
    if needle.upper() in f.upper():
        return True
    if depth == 0:
        return False
    return any(_formula_reaches(f"{col}{row}", needle, formulas, depth - 1)
               for col, row in _LOCAL_REF.findall(f) if f"{col}{row}" != ref)


def _check_sheet_contains(specs: list[dict], out_recalc: dict, output: Path) -> list[dict]:
    """Position-independent check for a NEW summary-style sheet: each expected
    value must appear in SOME cell of `sheet` after recalc, and, when
    `formula_contains` is given, that cell must be a live formula containing it.
    Where the numbers go (B1 vs B5 under a header row) is the editor's call."""
    results = []
    try:
        wb = openpyxl.load_workbook(str(output), data_only=False, read_only=True)
    except Exception as e:
        return [{"name": "sheet_contains", "ok": False, "detail": f"couldn't open output to read formulas: {e}"}]
    try:
        for spec in specs:
            sheet = spec["sheet"]
            formulas = {}
            if sheet in wb.sheetnames:
                for row in wb[sheet].iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("="):
                            formulas[c.coordinate] = c.value
            cells = {ref: v for (s, ref), v in out_recalc["values"].items() if s == sheet}
            for item in spec["values"]:
                want, tol, needle = float(item["value"]), item.get("tol", 0.0), item.get("formula_contains")
                hits = []
                for ref, v in cells.items():
                    try:
                        if abs(float(v) - want) > tol:
                            continue
                    except (TypeError, ValueError):
                        continue
                    # formula_contains may list several acceptable sources (any one passes)
                    needles = needle if isinstance(needle, list) else [needle]
                    if needle is None or any(_formula_reaches(ref, n, formulas) for n in needles):
                        hits.append(ref)
                name = f"sheet_contains:{sheet}:{want:g}"
                if sheet not in wb.sheetnames:
                    results.append({"name": name, "ok": False, "detail": f"sheet {sheet!r} not found"})
                elif hits:
                    results.append({"name": name, "ok": True, "detail": f"{want:g} found at {sheet}!{hits[0]}" + (f" as a formula containing {needle!r}" if needle else "")})
                else:
                    results.append({"name": name, "ok": False, "detail": f"no cell on {sheet!r} holds {want:g} (tol {tol})" + (f" as a formula containing {needle!r}" if needle else "")})
    finally:
        wb.close()
    return results


def _check_cells(expect_cells: list[dict], out_recalc: dict, orig_recalc: dict | None = None) -> list[dict]:
    results = []
    orig_values = (orig_recalc or {}).get("values", {})
    for c in expect_cells:
        sheet, ref = c["sheet"], c["ref"]
        before = orig_values.get((sheet, ref))
        if isinstance(before, str) and before.strip() in ERROR_STRINGS:
            # LibreOffice can't compute this cell even in the untouched original (e.g. Excel-only
            # structured refs like Table[#All]), so it can't grade anyone's result here.
            results.append({"name": f"cell:{sheet}!{ref}", "ok": True, "unverifiable": True,
                            "detail": f"UNVERIFIABLE: {sheet}!{ref} is {before.strip()} in LibreOffice's recalc of the original too"})
            continue
        expected = c["value"]
        tol = c.get("tol", 0.0)
        actual = out_recalc["values"].get((sheet, ref))
        ok = False
        if actual is None:
            detail = f"{sheet}!{ref} is empty (expected {expected} +/- {tol})"
        elif isinstance(actual, str) and actual.strip() in ERROR_STRINGS:
            detail = f"{sheet}!{ref} holds error {actual} (expected {expected} +/- {tol})"
        else:
            try:
                actual_num = float(actual)
            except (TypeError, ValueError):
                detail = f"{sheet}!{ref} = {actual!r} is not numeric (expected {expected} +/- {tol})"
            else:
                ok = abs(actual_num - float(expected)) <= tol
                detail = f"{sheet}!{ref} = {actual_num} (expected {expected} +/- {tol})"
        results.append({"name": f"cell:{sheet}!{ref}", "ok": ok, "detail": detail})
    return results


def _check_text_cells(expect_text_cells: list[dict], out_recalc: dict) -> list[dict]:
    results = []
    for c in expect_text_cells:
        sheet, ref = c["sheet"], c["ref"]
        expected = c["value"]
        actual = out_recalc["values"].get((sheet, ref))
        actual_str = "" if actual is None else str(actual)
        ok = actual_str == expected
        results.append({
            "name": f"text_cell:{sheet}!{ref}",
            "ok": ok,
            "detail": f"{sheet}!{ref} = {actual_str!r} (expected {expected!r})",
        })
    return results


def _check_formulas(expect_formulas: list[dict], output: Path) -> list[dict]:
    try:
        wb = openpyxl.load_workbook(str(output), data_only=False)
    except Exception as e:
        return [{"name": "formulas", "ok": False, "detail": f"couldn't open output to read formulas: {e}"}]

    results = []
    for f in expect_formulas:
        sheet, ref = f["sheet"], f["ref"]
        # contains_any: accept alternatives, e.g. M3 or the structured ref [total1] inside an Excel Table.
        options = f.get("contains_any") or [f["contains"]]
        contains = " | ".join(options)
        if sheet not in wb.sheetnames:
            results.append({"name": f"formula:{sheet}!{ref}", "ok": False, "detail": f"sheet {sheet!r} not found"})
            continue
        ws = wb[sheet]
        val = ws[ref].value
        if val is None:
            detail = f"{sheet}!{ref} has no formula (empty or value-only cell)"
            ok = False
        else:
            formula_text = str(val)
            ok = any(o.upper() in formula_text.upper() for o in options)
            detail = f"{sheet}!{ref} formula = {formula_text!r} (expected to contain {contains!r})"
        results.append({"name": f"formula:{sheet}!{ref}", "ok": ok, "detail": detail})
    wb.close()
    return results


def _check_tables(expect_tables: list[dict], out_inv: dict) -> list[dict]:
    by_name = {t["name"]: t for t in out_inv["tables"]["tables"]}
    results = []
    for t in expect_tables:
        name, ref = t["name"], t["ref"]
        actual = by_name.get(name)
        if actual is None:
            results.append({"name": f"table:{name}", "ok": False, "detail": f"table {name!r} not found in output"})
        else:
            ok = actual["ref"] == ref
            results.append({
                "name": f"table:{name}",
                "ok": ok,
                "detail": f"{name} ref = {actual['ref']!r} (expected {ref!r})",
            })
    return results


def _check_chart_series(expect_chart_series: list[dict], out_inv: dict) -> list[dict]:
    """Each entry names a chart part and one of val_ref / cat_ref / name_ref that
    some series in that chart must carry."""
    by_part = {c["part"]: c for c in out_inv["charts"]["charts"]}
    results = []
    for cs in expect_chart_series:
        part = cs["part"]
        key = next((k for k in ("val_ref", "cat_ref", "name_ref") if k in cs), None)
        if key is None:
            results.append({"name": f"chart_series:{part}", "ok": False,
                            "detail": f"task entry names no val_ref/cat_ref/name_ref: {cs}"})
            continue
        want = cs[key]
        chart = by_part.get(part)
        if chart is None:
            results.append({"name": f"chart_series:{part}", "ok": False,
                            "detail": f"chart part {part!r} not found in output"})
            continue
        refs = [s.get(key) for s in chart["series"]]
        results.append({"name": f"chart_series:{part}:{key}", "ok": want in refs,
                        "detail": f"series {key}s = {refs} (expected {want!r} present)"})
    return results


def _sqref_tokens_match(actual_sqref: str | None, expected_sqref: str) -> bool:
    if not actual_sqref:
        return False
    return set(expected_sqref.split()) <= set(actual_sqref.split())


def _check_sqref_class(expect_list: list[dict], blocks: list[dict], label: str) -> list[dict]:
    results = []
    for e in expect_list:
        sheet, sqref = e["sheet"], e["sqref"]
        matches = [b for b in blocks if b["sheet"] == sheet and _sqref_tokens_match(b.get("sqref"), sqref)]
        ok = bool(matches)
        results.append({
            "name": f"{label}:{sheet}!{sqref}",
            "ok": ok,
            "detail": (f"found {len(matches)} matching block(s)" if ok
                       else f"no {label} block on sheet {sheet!r} covering {sqref!r}"),
        })
    return results


_NEEDLESS_SHEET_QUOTES = re.compile(r"'([A-Za-z_][A-Za-z0-9_.]*)'!")
_BOOL_CALL = re.compile(r"\b(TRUE|FALSE)\(\)", re.IGNORECASE)


def _same_formula(a: str | None, b: str | None) -> bool:
    """Equal up to spellings Excel treats as identical: quotes around a sheet
    name that doesn't need them ('KYC'!$1:$1 == KYC!$1:$1, Haiku round, kyc-t4)
    and TRUE()/FALSE() for TRUE/FALSE."""
    if a is None or b is None:
        return a is b
    norm = lambda f: _BOOL_CALL.sub(lambda m: m.group(1).upper(), _NEEDLESS_SHEET_QUOTES.sub(r"\1!", str(f).strip()))
    return norm(a) == norm(b)


def _check_defined_names(expect_defined_names: dict, out_inv: dict) -> list[dict]:
    actual = out_inv["defined_names"]
    results = []
    for name, formula in expect_defined_names.items():
        a = actual.get(name)
        ok = a is not None and _same_formula(a, str(formula))
        results.append({
            "name": f"defined_name:{name}",
            "ok": ok,
            "detail": f"= {a!r} (expected {formula!r})",
        })
    return results


def _check_sheet_names(expect_sheet_names: list[str], out_inv: dict, orig_sheet_names: list[str] | None = None) -> list[dict]:
    """Exact names, and the existing sheets in the expected order. A sheet the task
    ADDS (a name the original doesn't have, when the expected list is longer) may go
    anywhere: no prompt says where, and Claude reasonably put 'Pilot KPIs' before
    three empty leftover sheets instead of after them."""
    actual = out_inv["sheet_names"]
    ok = actual == expect_sheet_names
    note = ""
    if not ok and orig_sheet_names is not None and len(expect_sheet_names) > len(orig_sheet_names):
        added = set(expect_sheet_names) - set(orig_sheet_names)
        ok = (sorted(actual) == sorted(expect_sheet_names)
              and [n for n in actual if n not in added] == [n for n in expect_sheet_names if n not in added])
        if ok:
            note = f"; added sheet(s) {sorted(added)} placed elsewhere, existing order kept"
    return [{"name": "sheet_names", "ok": ok, "detail": f"= {actual} (expected {expect_sheet_names}){note}"}]


def _check_expect(expect: dict, output: Path, out_inv: dict | None, out_recalc: dict,
                  orig_recalc: dict | None = None, orig_sheet_names: list[str] | None = None) -> list[dict]:
    checks: list[dict] = []
    if "cells" in expect:
        checks += _check_cells(expect["cells"], out_recalc, orig_recalc)
    if "column_values" in expect:
        checks += _check_column_values(expect["column_values"], out_recalc, orig_recalc)
    if "sheet_contains" in expect:
        checks += _check_sheet_contains(expect["sheet_contains"], out_recalc, output)
    if "text_cells" in expect:
        checks += _check_text_cells(expect["text_cells"], out_recalc)
    if "formulas" in expect:
        checks += _check_formulas(expect["formulas"], output)
    if out_inv is not None:
        if "tables" in expect:
            checks += _check_tables(expect["tables"], out_inv)
        if "chart_series" in expect:
            checks += _check_chart_series(expect["chart_series"], out_inv)
        if "cf_sqrefs" in expect:
            cf_blocks = out_inv["conditional_formatting"]["classic"]["blocks"] + out_inv["conditional_formatting"]["x14"]["blocks"]
            checks += _check_sqref_class(expect["cf_sqrefs"], cf_blocks, "cf_sqref")
        if "dv_sqrefs" in expect:
            dv_blocks = out_inv["data_validations"]["classic"]["blocks"] + out_inv["data_validations"]["x14"]["blocks"]
            checks += _check_sqref_class(expect["dv_sqrefs"], dv_blocks, "dv_sqref")
        if "defined_names" in expect:
            checks += _check_defined_names(expect["defined_names"], out_inv)
        if "sheet_names" in expect:
            checks += _check_sheet_names(expect["sheet_names"], out_inv, orig_sheet_names)
    else:
        for key in ("tables", "chart_series", "cf_sqrefs", "dv_sqrefs", "defined_names", "sheet_names"):
            if key in expect:
                checks.append({"name": key, "ok": False, "detail": "output inventory unavailable"})
    return checks


# --------------------------------------------------------------------------
# preservation: inventory(original) vs inventory(output), per preserve class
# --------------------------------------------------------------------------

def _diff_vba(orig: dict, out: dict) -> list[str]:
    ov, nv = orig["vba"], out["vba"]
    if ov["present"] and not nv["present"]:
        return ["vba: vbaProject.bin missing"]
    if ov["present"] and nv["present"] and ov["sha256"] != nv["sha256"]:
        return ["vba: vbaProject.bin changed (sha256 mismatch)"]
    return []


def _diff_images(orig: dict, out: dict) -> list[str]:
    orig_media = orig["images"]["media"]
    orig_shas = Counter(m["sha256"] for m in orig_media)
    out_shas = Counter(m["sha256"] for m in out["images"]["media"])
    missing_count = 0
    missing_parts: list[str] = []
    for sha, cnt in orig_shas.items():
        deficit = cnt - out_shas.get(sha, 0)
        if deficit > 0:
            missing_count += deficit
            parts = [m["part"] for m in orig_media if m["sha256"] == sha][:deficit]
            missing_parts.extend(parts)
    if missing_count:
        total = len(orig_media)
        return [f"images: {missing_count} of {total} media part(s) gone ({', '.join(missing_parts)})"]
    return []


def _diff_charts(orig: dict, out: dict, chart_series_expect: list[dict]) -> list[str]:
    losses: list[str] = []
    orig_charts = {c["part"]: c for c in orig["charts"]["charts"]}
    out_charts = {c["part"]: c for c in out["charts"]["charts"]}
    if orig["charts"]["count"] != out["charts"]["count"]:
        losses.append(f"charts: count changed ({orig['charts']['count']} -> {out['charts']['count']})")

    # A series ref may change only where expect.chart_series asserts the new value
    # (val_ref, cat_ref or name_ref), e.g. a row insert growing a range.
    allowed = {(e["part"], k, e[k]) for e in chart_series_expect for k in ("val_ref", "cat_ref", "name_ref") if k in e}
    for part, oc in orig_charts.items():
        nc = out_charts.get(part)
        if nc is None:
            losses.append(f"charts: {part} missing")
            continue
        if oc["n_series"] != nc["n_series"]:
            losses.append(f"charts: {part} series count changed ({oc['n_series']} -> {nc['n_series']})")
            continue
        for i, (os_, ns_) in enumerate(zip(oc["series"], nc["series"])):
            for key in ("name_ref", "cat_ref", "val_ref"):
                if os_.get(key) != ns_.get(key) and (part, key, ns_.get(key)) not in allowed:
                    losses.append(f"charts: {part} series[{i}].{key} changed "
                                  f"({os_.get(key)!r} -> {ns_.get(key)!r}, not asserted in expect.chart_series)")
    return losses


def _diff_pivot_tables(orig: dict, out: dict) -> list[str]:
    orig_names = Counter(t["name"] for t in orig["pivot_tables"]["tables"])
    out_names = Counter(t["name"] for t in out["pivot_tables"]["tables"])
    missing = orig_names - out_names
    if missing:
        total = orig["pivot_tables"]["count"]
        names = sorted(missing.elements())
        return [f"pivot_tables: {sum(missing.values())} of {total} missing ({', '.join(names)})"]
    return []


def _diff_pivot_caches(orig: dict, out: dict) -> list[str]:
    # Identity is (type, sheet, named source), not the range: inserting a data row
    # correctly grows a cache's source (A1:D13 -> A1:D14), and both arms did exactly
    # that on pivotrpt-t4. A cache whose range SHRINKS still counts as damaged.
    def sig(c):
        return (c.get("source_type"), c.get("source_sheet"), c.get("source_name"))
    orig_caches, out_caches = orig["pivot_caches"]["caches"], out["pivot_caches"]["caches"]
    missing = Counter(sig(c) for c in orig_caches) - Counter(sig(c) for c in out_caches)
    losses = []
    if missing:
        total = orig["pivot_caches"]["count"]
        losses.append(f"pivot_caches: {sum(missing.values())} of {total} missing (sources: {list(missing.elements())})")
    for a, b in zip(sorted(orig_caches, key=sig), sorted(out_caches, key=sig)):
        ra, rb = _ref_rows(a.get("source_ref")), _ref_rows(b.get("source_ref"))
        if sig(a) == sig(b) and ra and rb and rb[1] < ra[1]:
            losses.append(f"pivot_caches: source {a.get('source_ref')} shrank to {b.get('source_ref')}")
    return losses


def _ref_rows(ref):
    """(first row, last row) of an A1-style range, or None."""
    m = re.fullmatch(r"\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?", ref or "")
    return (int(m.group(1)), int(m.group(2) or m.group(1))) if m else None


def _diff_comments(orig: dict, out: dict) -> list[str]:
    if orig["comments"]["count"] > out["comments"]["count"]:
        gone = orig["comments"]["count"] - out["comments"]["count"]
        return [f"comments: {gone} of {orig['comments']['count']} missing"]
    return []


def _diff_hyperlinks(orig: dict, out: dict) -> list[str]:
    def sig(l):
        return (l["sheet"], l["ref"], l["target"])
    orig_sigs = Counter(sig(l) for l in orig["hyperlinks"]["links"])
    out_sigs = Counter(sig(l) for l in out["hyperlinks"]["links"])
    missing = orig_sigs - out_sigs
    if missing:
        total = orig["hyperlinks"]["count"]
        sample = list(missing.elements())[:5]
        return [f"hyperlinks: {sum(missing.values())} of {total} missing (e.g. {sample})"]
    return []


def _diff_protection(orig: dict, out: dict) -> list[str]:
    losses = []
    if orig["protection"]["workbook"] is not None and out["protection"]["workbook"] is None:
        losses.append("protection: workbook-level protection removed")
    orig_sheets = set(orig["protection"]["sheets"].keys())
    out_sheets = set(out["protection"]["sheets"].keys())
    missing = orig_sheets - out_sheets
    if missing:
        losses.append(f"protection: sheet protection removed on {sorted(missing)}")
    return losses


def _diff_sparklines(orig: dict, out: dict) -> list[str]:
    def group_sig(g):
        return (g["sheet"], g["type"], tuple(sorted((s["f"], s["sqref"]) for s in g["sparklines"])))
    orig_sigs = Counter(group_sig(g) for g in orig["sparklines"]["groups"])
    out_sigs = Counter(group_sig(g) for g in out["sparklines"]["groups"])
    missing = orig_sigs - out_sigs
    if missing:
        total = orig["sparklines"]["count"]
        return [f"sparklines: {sum(missing.values())} of {total} group(s) missing or changed"]
    return []


def _diff_external_links(orig: dict, out: dict) -> list[str]:
    orig_targets = Counter(l["target"] for l in orig["external_links"]["links"])
    out_targets = Counter(l["target"] for l in out["external_links"]["links"])
    missing = orig_targets - out_targets
    if missing:
        total = orig["external_links"]["count"]
        return [f"external_links: {sum(missing.values())} of {total} missing ({list(missing.elements())})"]
    return []


def _diff_merged_cells(orig: dict, out: dict) -> list[str]:
    losses = []
    out_by_sheet = out["merged_cells"]["by_sheet"]
    for sheet, ranges in orig["merged_cells"]["by_sheet"].items():
        orig_c = Counter(ranges)
        out_c = Counter(out_by_sheet.get(sheet, []))
        missing = orig_c - out_c
        if missing:
            losses.append(
                f"merged_cells: {sheet}: {sum(missing.values())} of {len(ranges)} merge(s) missing "
                f"({list(missing.elements())})"
            )
    return losses


def _diff_sqref_class(orig: dict, out: dict, expect_list: list[dict], label: str, blocks_of) -> list[str]:
    """Generic diff for conditional_formatting / data_validations: a block
    that disappears from its original (sheet, sqref) is only NOT a loss if
    the task's own `expect` asserts a block on that sheet at a new sqref
    that is actually present in the output (a legitimate row-insert move).
    """
    orig_blocks = blocks_of(orig)
    out_blocks = blocks_of(out)
    orig_sig = Counter(orig_blocks)
    out_sig = Counter(out_blocks)
    missing = orig_sig - out_sig

    exemptions = Counter()
    out_sqref_set = {(b[0], b[1]) for b in out_blocks}
    for e in expect_list:
        sheet, sqref = e["sheet"], e["sqref"]
        if (sheet, sqref) in out_sqref_set:
            exemptions[sheet] += 1

    # A block extended over newly added cells (A2:G2 -> A2:H2 when a column joins the
    # table) is preserved, not lost: match it to an extra output block with the same
    # sheet/rules/kind whose sqref covers the original's.
    extra = list((out_sig - orig_sig).elements())
    losses = []
    for block_key, cnt in missing.items():
        sheet, sqref, rule_types, kind = block_key
        remaining = cnt
        for cand in list(extra):
            if remaining and cand[0] == sheet and cand[2] == rule_types and cand[3] == kind and _sqref_covers(cand[1], sqref):
                extra.remove(cand)
                remaining -= 1
        if exemptions[sheet] > 0:
            take = min(exemptions[sheet], remaining)
            exemptions[sheet] -= take
            remaining -= take
        for _ in range(remaining):
            losses.append(f"{label}: {kind} block on {sqref} removed")
    return losses


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _sqref_rects(sqref: str) -> list[tuple[int, int, int, int]]:
    rects = []
    for tok in (sqref or "").replace("$", "").split():
        m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", tok)
        if not m:
            return []
        c1, r1 = _col_num(m.group(1)), int(m.group(2))
        c2, r2 = (_col_num(m.group(3)), int(m.group(4))) if m.group(3) else (c1, r1)
        rects.append((min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)))
    return rects


def _sqref_covers(big: str, small: str) -> bool:
    """True if every range of `small` sits inside some range of `big`."""
    b, s_ = _sqref_rects(big), _sqref_rects(small)
    return bool(b) and bool(s_) and all(any(bx1 <= x1 and by1 <= y1 and x2 <= bx2 and y2 <= by2
                                             for bx1, by1, bx2, by2 in b) for x1, y1, x2, y2 in s_)


def _cf_blocks_of(inv: dict):
    out = []
    for b in inv["conditional_formatting"]["classic"]["blocks"]:
        out.append((b["sheet"], b["sqref"], tuple(b["rule_types"]), "classic"))
    for b in inv["conditional_formatting"]["x14"]["blocks"]:
        out.append((b["sheet"], b["sqref"], tuple(b["rule_types"]), "x14"))
    return out


def _dv_blocks_of(inv: dict):
    out = []
    for b in inv["data_validations"]["classic"]["blocks"]:
        out.append((b["sheet"], b["sqref"], b["type"], "classic"))
    for b in inv["data_validations"]["x14"]["blocks"]:
        out.append((b["sheet"], b["sqref"], b["type"], "x14"))
    return out


def _diff_defined_names(orig: dict, out: dict, expect_defined_names: dict) -> list[str]:
    losses = []
    allowed = expect_defined_names or {}
    out_names = out["defined_names"]
    for name, formula in orig["defined_names"].items():
        if name not in out_names:
            losses.append(f"defined_names: {name} missing")
            continue
        if not _same_formula(out_names[name], formula):
            if name in allowed and _same_formula(out_names[name], str(allowed[name])):
                continue
            losses.append(f"defined_names: {name} changed ({formula!r} -> {out_names[name]!r})")
    return losses


def _diff_tables(orig: dict, out: dict, expect_tables: list[dict]) -> list[str]:
    losses = []
    allowed = {t["name"]: t["ref"] for t in expect_tables}
    out_tables = {t["name"]: t for t in out["tables"]["tables"]}
    for t in orig["tables"]["tables"]:
        name = t["name"]
        nt = out_tables.get(name)
        if nt is None:
            losses.append(f"tables: {name} missing")
            continue
        if nt["ref"] != t["ref"]:
            if name in allowed and nt["ref"] == allowed[name]:
                continue
            losses.append(f"tables: {name} ref changed ({t['ref']!r} -> {nt['ref']!r}, not asserted in expect.tables)")
    return losses


_PRESERVE_DIFFERS = {
    "vba": lambda task, o, n: _diff_vba(o, n),
    "images": lambda task, o, n: _diff_images(o, n),
    "charts": lambda task, o, n: _diff_charts(o, n, task.get("expect", {}).get("chart_series", [])),
    "pivot_tables": lambda task, o, n: _diff_pivot_tables(o, n),
    "pivot_caches": lambda task, o, n: _diff_pivot_caches(o, n),
    "comments": lambda task, o, n: _diff_comments(o, n),
    "hyperlinks": lambda task, o, n: _diff_hyperlinks(o, n),
    "protection": lambda task, o, n: _diff_protection(o, n),
    "sparklines": lambda task, o, n: _diff_sparklines(o, n),
    "external_links": lambda task, o, n: _diff_external_links(o, n),
    "merged_cells": lambda task, o, n: _diff_merged_cells(o, n),
    "conditional_formatting": lambda task, o, n: _diff_sqref_class(
        o, n, task.get("expect", {}).get("cf_sqrefs", []), "conditional_formatting", _cf_blocks_of
    ),
    "data_validations": lambda task, o, n: _diff_sqref_class(
        o, n, task.get("expect", {}).get("dv_sqrefs", []), "data_validations", _dv_blocks_of
    ),
    "defined_names": lambda task, o, n: _diff_defined_names(o, n, task.get("expect", {}).get("defined_names", {})),
    "tables": lambda task, o, n: _diff_tables(o, n, task.get("expect", {}).get("tables", [])),
}


_REL_NS_OFFICE = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _sheet_parts(path: Path) -> dict[str, str]:
    """{sheet name: worksheet part path}. The part path survives a rename."""
    try:
        with zipfile.ZipFile(path) as zf:
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, OSError, zipfile.BadZipFile, ET.ParseError):
        return {}
    targets = {r.get("Id"): r.get("Target", "") for r in rels.iter(f"{_RELS_NS}Relationship")}
    out = {}
    for sh in wb.iter(f"{_MAIN_NS}sheet"):
        t = targets.get(sh.get(f"{_REL_NS_OFFICE}id"), "")
        out[sh.get("name")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
    return out


def sheet_rename_map(original: Path, output: Path) -> dict[str, str]:
    """Old name -> new name for sheets whose worksheet part is the same file."""
    before, after = _sheet_parts(original), _sheet_parts(output)
    by_part = {part: name for name, part in after.items()}
    return {old: by_part[part] for old, part in before.items() if part in by_part and by_part[part] != old}


def _quote_sheet(name: str) -> str:
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) else "'" + name.replace("'", "''") + "'"


def apply_sheet_renames(value, renames: dict[str, str]):
    """Rewrite an inventory as it should look after a CORRECT rename: every
    `sheet` field and every sheet-qualified reference ('Old'!A1 / Old!A1) moves
    to the new name. Whatever still differs from the output is a real loss."""
    if not renames:
        return value
    if isinstance(value, dict):
        # Keys can be sheet names too (merged_cells.by_sheet, protection.sheets).
        return {renames.get(k, k): (renames.get(v, v) if k == "sheet" and isinstance(v, str) else apply_sheet_renames(v, renames))
                for k, v in value.items()}
    if isinstance(value, list):
        return [renames.get(v, v) if isinstance(v, str) and v in renames else apply_sheet_renames(v, renames)
                for v in value]
    if isinstance(value, str):
        for old, new in renames.items():
            value = value.replace("'" + old.replace("'", "''") + "'!", _quote_sheet(new) + "!")
            if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", old):
                value = re.sub(rf"(?<![A-Za-z0-9_.']){re.escape(old)}!", _quote_sheet(new) + "!", value)
        return value
    return value


def _diff_preserved(task: dict, orig_inv: dict, out_inv: dict) -> list[str]:
    losses: list[str] = []
    for cls in task.get("preserve", []):
        differ = _PRESERVE_DIFFERS.get(cls)
        if differ is None:
            continue  # unrecognized class name; not this grader's job to validate tasks.json
        losses.extend(differ(task, orig_inv, out_inv))
    return losses


# --------------------------------------------------------------------------
# top-level grade()
# --------------------------------------------------------------------------

def grade(task: dict, original: Path, output: Path | None, recalc_timeout: int = 60) -> dict:
    """Grade one run of `task` against its `original` workbook and the
    edited `output` workbook. Never receives or looks at which arm produced
    `output` — arm/rep are attached by the caller after this returns.
    """
    task_id = task.get("id", "<unknown>")
    original = Path(original)

    if output is None or not Path(output).exists():
        return {
            "task_id": task_id,
            "correct": False,
            "preserved": False,
            "valid": False,
            "pass": False,
            "checks": [{
                "name": "output_exists",
                "ok": False,
                "detail": "no output file found for this run",
            }],
            "losses": [],
            "new_errors": [],
        }
    output = Path(output)

    checks: list[dict] = []
    new_errors: list[dict] = []

    struct_checks, structurally_valid = _check_validity_structure(output)
    checks.extend(struct_checks)

    # Stop early only when the package can't be read at all. A dangling
    # relationship or a Content_Types gap still leaves a readable file, and the
    # report should show both "invalid" and what was lost (e.g. the macros),
    # not hide the losses behind the validity failure.
    readable = all(c["ok"] for c in struct_checks if c["name"] in ("unzip", "xml_parts_parse"))
    if not readable:
        return {
            "task_id": task_id,
            "correct": False,
            "preserved": False,
            "valid": False,
            "pass": False,
            "checks": checks,
            "losses": [],
            "new_errors": [],
        }

    try:
        orig_inv = inventory(str(original))
    except Exception as e:
        checks.append({"name": "original_inventory", "ok": False, "detail": f"failed to inventory original: {e}"})
        orig_inv = None

    try:
        out_inv = inventory(str(output))
        checks.append({"name": "output_inventory", "ok": True, "detail": f"{len(out_inv['part_sha256'])} part(s)"})
    except Exception as e:
        checks.append({"name": "output_inventory", "ok": False, "detail": f"failed to inventory output: {e}"})
        out_inv = None

    orig_recalc = _recalc_workbook(original, timeout=recalc_timeout)
    out_recalc = _recalc_workbook(output, timeout=recalc_timeout)
    checks.append({
        "name": "libreoffice_opens",
        "ok": out_recalc["ok"],
        "detail": out_recalc["reason"] or "opened and recalculated cleanly",
    })

    no_new_errors_ok: bool
    if out_recalc["ok"] and orig_recalc["ok"]:
        new_error_keys = sorted(out_recalc["error_cells"] - orig_recalc["error_cells"])
        no_new_errors_ok = not new_error_keys
        for sheet, ref in new_error_keys:
            new_errors.append({"sheet": sheet, "ref": ref, "value": out_recalc["values"].get((sheet, ref))})
    else:
        no_new_errors_ok = False
        if not orig_recalc["ok"]:
            new_errors.append({"sheet": None, "ref": None, "value": f"original recalc failed: {orig_recalc['reason']}"})

    if no_new_errors_ok:
        no_new_errors_detail = "no new error cells vs. the original's recalc"
    else:
        detail_items = [f"{e['sheet']}!{e['ref']}={e['value']}" for e in new_errors if e["sheet"]]
        if detail_items:
            no_new_errors_detail = f"{len(detail_items)} new error cell(s): {detail_items}"
        else:
            no_new_errors_detail = "could not establish a recalc baseline"
    checks.append({"name": "no_new_errors", "ok": no_new_errors_ok, "detail": no_new_errors_detail})

    expect = task.get("expect", {})
    expect_checks = _check_expect(expect, output, out_inv, out_recalc, orig_recalc,
                                  orig_inv["sheet_names"] if orig_inv is not None else None)
    checks.extend(expect_checks)
    if "no_new_errors" in expect:
        checks.append({"name": "expect.no_new_errors", "ok": no_new_errors_ok, "detail": "see no_new_errors check above"})

    correct = all(c["ok"] for c in expect_checks) and no_new_errors_ok

    if orig_inv is not None and out_inv is not None:
        renames = sheet_rename_map(Path(original), Path(output))
        if renames:
            checks.append({"name": "sheet_renames", "ok": True, "detail": f"compared across renames {renames}"})
        losses = _diff_preserved(task, apply_sheet_renames(orig_inv, renames), out_inv)
    else:
        losses = ["preservation: could not compare (inventory of original or output failed)"]
    preserved = not losses
    checks.append({
        "name": "preserved",
        "ok": preserved,
        "detail": "no losses" if preserved else f"{len(losses)} loss(es)",
    })

    valid = structurally_valid and out_recalc["ok"] and no_new_errors_ok

    return {
        "task_id": task_id,
        "correct": correct,
        "preserved": preserved,
        "valid": valid,
        "pass": correct and preserved and valid,
        "checks": checks,
        "losses": losses,
        "new_errors": new_errors,
    }


# --------------------------------------------------------------------------
# CLI: walk runs/<task_id>/<arm>/<rep>/, grade each, write results.jsonl
# --------------------------------------------------------------------------

def _load_run_meta(rep_dir: Path, work_dir: Path) -> dict:
    for candidate in (rep_dir / "run.json", work_dir / "run.json"):
        if candidate.exists():
            try:
                return json.loads(candidate.read_text())
            except (json.JSONDecodeError, OSError):
                continue
    return {}


_BACKUP_HINT = re.compile(r"(backup|\.bak|\borig(inal)?\b|copy)", re.IGNORECASE)


def find_output_file(
    work_dir: Path, original_name: str, rep_dir: Path, original_path: Path | None = None
) -> Path | None:
    """The output is the work copy of the original if it was edited in place
    (its sha256 differs from the pristine original): prompts ask to edit the
    workbook and rarely say "save as", and a session that builds intermediates
    and then moves its final file onto the original name has delivered that
    file. Otherwise it is whatever run.json lists (a harness file listing, not
    a statement by the model), else a new workbook in work_dir. Among several
    new files, backups (backup/.bak/orig/copy in the name) lose to the rest,
    then newest wins.
    """
    in_place = work_dir / original_name
    if in_place.is_file() and original_path is not None:
        if _sha256_file(in_place) != _sha256_file(Path(original_path)):
            return in_place

    meta = _load_run_meta(rep_dir, work_dir)
    for key in ("output_files", "output_file", "outputs", "output"):
        if key not in meta:
            continue
        val = meta[key]
        candidates = [val] if isinstance(val, str) else (val if isinstance(val, list) else [])
        for c in candidates:
            p = work_dir / c
            # run.json lists every file Claude touched, e.g. the XML parts of a workbook it
            # unzipped into x/ for inspection; only a spreadsheet can be the output.
            if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm") and zipfile.is_zipfile(p):
                return p

    new_files = [p for p in work_dir.glob("*.xls*") if p.is_file() and p.name != original_name]
    if new_files:
        new_files.sort(key=lambda p: (bool(_BACKUP_HINT.search(p.name)), -p.stat().st_mtime))
        return new_files[0]

    # Only reached with original_path=None (legacy callers): no pristine copy to compare against.
    in_place = work_dir / original_name
    if in_place.is_file():
        if original_path is None or _sha256_file(in_place) != _sha256_file(Path(original_path)):
            return in_place
    return None


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def resolve_workbook(workbook: str, corpus_dir: Path) -> Path:
    """Resolve a task's `workbook` string to a real file, agreeing with
    runner.py's own resolve_workbook() so the grader always checks the
    exact same original the runner copied into each run's work/ dir.
    Falls back to corpus_dir/<basename> for the bare-filename form shown
    in TASK_SCHEMA.md's example, in case a task doesn't use a repo-root-
    relative path.
    """
    p = Path(workbook)
    if p.is_absolute():
        return p
    repo_relative = REPO_ROOT / p
    if repo_relative.exists():
        return repo_relative
    return corpus_dir / p.name


def _iter_runs(runs_dir: Path):
    for task_dir in sorted(p for p in runs_dir.iterdir() if p.is_dir()):
        for arm_dir in sorted(p for p in task_dir.iterdir() if p.is_dir()):
            for rep_dir in sorted(p for p in arm_dir.iterdir() if p.is_dir()):
                # Skip runs still in progress (no run.json yet) and runs set aside
                # by hand (e.g. '1.cut-by-usage-limit'): neither is a finished result.
                if not rep_dir.name.isdigit() or not (rep_dir / "run.json").exists():
                    continue
                yield task_dir.name, arm_dir.name, rep_dir.name, rep_dir


def _run_contaminated(rep_dir: Path) -> bool:
    """run.json's contamination.contaminated flag (False when absent)."""
    try:
        data = json.loads((rep_dir / "run.json").read_text())
    except (OSError, ValueError):
        return False
    return bool((data.get("contamination") or {}).get("contaminated"))


def print_summary(tasks_by_id: dict, results: list[dict]) -> None:
    by_arm = defaultdict(lambda: {"n": 0, "correct": 0, "preserved": 0, "valid": 0, "pass": 0})
    by_arm_tier = defaultdict(lambda: {"n": 0, "correct": 0, "preserved": 0, "valid": 0, "pass": 0})

    for r in results:
        arm = r.get("arm", "?")
        task = tasks_by_id.get(r["task_id"], {})
        tier = task.get("tier", "?")
        for bucket, key in ((by_arm, arm), (by_arm_tier, (arm, tier))):
            b = bucket[key]
            b["n"] += 1
            for field in ("correct", "preserved", "valid", "pass"):
                if r.get(field):
                    b[field] += 1

    def pct(b, field):
        return 100.0 * b[field] / b["n"] if b["n"] else 0.0

    print("\n=== Per-arm summary ===")
    header = f"{'arm':<14}{'n':>5}{'correct%':>10}{'preserved%':>12}{'valid%':>9}{'pass%':>8}"
    print(header)
    for arm, b in sorted(by_arm.items()):
        print(f"{arm:<14}{b['n']:>5}{pct(b,'correct'):>9.1f}%{pct(b,'preserved'):>11.1f}%{pct(b,'valid'):>8.1f}%{pct(b,'pass'):>7.1f}%")

    print("\n=== Per-arm-per-tier summary ===")
    header = f"{'arm':<14}{'tier':>5}{'n':>5}{'correct%':>10}{'preserved%':>12}{'valid%':>9}{'pass%':>8}"
    print(header)
    for (arm, tier), b in sorted(by_arm_tier.items(), key=lambda kv: (kv[0][0], str(kv[0][1]))):
        print(f"{arm:<14}{str(tier):>5}{b['n']:>5}{pct(b,'correct'):>9.1f}%{pct(b,'preserved'):>11.1f}%{pct(b,'valid'):>8.1f}%{pct(b,'pass'):>7.1f}%")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Deterministic grader for xlsx-bench runs")
    ap.add_argument("--tasks", required=True, help="path to tasks.json")
    ap.add_argument("--runs", required=True, help="path to runs/ directory (runs/<task>/<arm>/<rep>/)")
    ap.add_argument("--out", default=None, help="write one JSON line per run here (default: stdout)")
    ap.add_argument("--summary", action="store_true", help="print pass/correct/preserved/valid rates per arm and tier")
    ap.add_argument("--exclude-contaminated", action="store_true",
                    help="summary only: list runs whose run.json says contamination.contaminated "
                         "(session touched the bench tree or another run) but leave them out of the rates")
    ap.add_argument("--corpus-dir", default=None, help="override corpus dir (default: sibling corpus/ of this script)")
    ap.add_argument("--recalc-timeout", type=int, default=60, help="soffice recalc timeout in seconds (default 60)")
    args = ap.parse_args(argv)

    tasks = json.loads(Path(args.tasks).read_text())
    tasks_by_id = {t["id"]: t for t in tasks}

    corpus_dir = Path(args.corpus_dir) if args.corpus_dir else (_HERE / "corpus")
    runs_dir = Path(args.runs)

    results = []
    for task_id, arm, rep, rep_dir in _iter_runs(runs_dir):
        task = tasks_by_id.get(task_id)
        if task is None:
            results.append({
                "task_id": task_id,
                "correct": False,
                "preserved": False,
                "valid": False,
                "pass": False,
                "checks": [{"name": "task_lookup", "ok": False, "detail": f"task {task_id!r} not found in {args.tasks}"}],
                "losses": [],
                "new_errors": [],
                "arm": arm,
                "rep": rep,
            })
            continue

        original = resolve_workbook(task["workbook"], corpus_dir)
        work_dir = rep_dir / "work"
        if not work_dir.is_dir():
            work_dir = rep_dir
        output = find_output_file(work_dir, Path(task["workbook"]).name, rep_dir, original_path=original)
        result_output = str(output.relative_to(rep_dir)) if output else None

        try:
            result = grade(task, original, output, recalc_timeout=args.recalc_timeout)
        except Exception as e:  # one bad task or file must not stop the whole grading run
            result = {"task_id": task_id, "correct": False, "preserved": False, "valid": False,
                      "pass": False, "losses": [], "new_errors": [],
                      "checks": [{"name": "grader_error", "ok": False, "detail": f"{type(e).__name__}: {e}"}]}
        result["arm"] = arm
        result["rep"] = rep
        result["output_graded"] = result_output
        result["contaminated"] = _run_contaminated(rep_dir)
        results.append(result)

    if args.out:
        with open(args.out, "w") as f:
            for r in results:
                f.write(json.dumps(r) + "\n")
    else:
        for r in results:
            print(json.dumps(r))

    if args.summary:
        if args.exclude_contaminated:
            dirty = [r for r in results if r.get("contaminated")]
            print(f"\n=== Excluded as contaminated: {len(dirty)} ===")
            for r in dirty:
                print(f"  {r['task_id']}/{r.get('arm')}/{r.get('rep')}")
            results = [r for r in results if not r.get("contaminated")]
        print_summary(tasks_by_id, results)

    return 0


if __name__ == "__main__":
    sys.exit(main())
