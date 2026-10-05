#!/usr/bin/env python3
"""Self-check for research/xlsx-bench/tasks.json against TASK_SCHEMA.md.

Validates, without needing grade.py or a live recalc:
  - tasks.json is a list of task dicts with all required keys present and well-typed
  - tier is one of 1-5; id is unique across the whole suite
  - workbook exists in research/xlsx-bench/corpus/
  - every `preserve` class is a real feature class that exists in that workbook's
    inventory.py output (i.e. we're not claiming to preserve something the workbook
    doesn't have)
  - every `expect.tables` / `expect.chart_series` / `expect.cf_sqrefs` / `expect.dv_sqrefs`
    entry either (a) names something that exists in the ORIGINAL workbook's inventory
    (the pre-change identity), or (b) is explicitly flagged as an expected POST-CHANGE
    state in that task's `notes` field (a lightweight heuristic: the notes text must
    mention the new name/ref, since we can't recompute the post-edit workbook here)

Exits 0 iff there are zero errors. Prints a bullet list of errors (and a summary of
warnings, which don't affect the exit code) otherwise.

Usage: python3 research/xlsx-bench/check_tasks.py [tasks.json]
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
CORPUS_DIR = HERE / "corpus"

sys.path.insert(0, str(HERE))
import inventory as inventory_mod  # noqa: E402

REQUIRED_KEYS = {"id", "workbook", "tier", "kind", "prompt", "expect", "preserve", "notes"}
VALID_TIERS = {1, 2, 3, 4, 5}
VALID_KINDS = {"value", "formula_column", "bulk", "insert_row", "sheet_op"}
# Feature classes inventory.py can report, mapped to the `preserve` vocabulary used in
# TASK_SCHEMA.md's example (`vba, images, charts, pivot_tables, pivot_caches, comments,
# hyperlinks, protection, sparklines, external_links, merged_cells,
# conditional_formatting, data_validations, defined_names, tables`), plus two real
# feature classes inventory.py also reports that TASK_SCHEMA.md's example list doesn't
# spell out (frozen_panes -- used because benefits-contribution-calculator.xlsm has real
# frozen panes untouched by any of its 5 tasks).
PRESERVE_CLASSES = {
    "vba", "images", "charts", "pivot_tables", "pivot_caches", "comments",
    "hyperlinks", "protection", "sparklines", "external_links", "merged_cells",
    "conditional_formatting", "data_validations", "defined_names", "tables",
    "frozen_panes",
}

EXPECT_KEYS = {
    "cells", "text_cells", "column_values", "sheet_contains", "formulas", "tables", "chart_series", "cf_sqrefs",
    "dv_sqrefs", "defined_names", "sheet_names", "no_new_errors",
}


def has_feature(inv: dict, cls: str) -> bool:
    """Does this workbook's inventory actually have a nonzero instance of `cls`?"""
    if cls == "vba":
        return bool(inv["vba"]["present"])
    if cls == "images":
        return inv["images"]["count"] > 0
    if cls == "charts":
        return inv["charts"]["count"] > 0
    if cls == "pivot_tables":
        return inv["pivot_tables"]["count"] > 0
    if cls == "pivot_caches":
        return inv["pivot_caches"]["count"] > 0
    if cls == "comments":
        return inv["comments"]["count"] > 0
    if cls == "hyperlinks":
        return inv["hyperlinks"]["count"] > 0
    if cls == "protection":
        return inv["protection"]["workbook"] is not None or bool(inv["protection"]["sheets"])
    if cls == "sparklines":
        return inv["sparklines"]["count"] > 0
    if cls == "external_links":
        return inv["external_links"]["count"] > 0
    if cls == "merged_cells":
        return inv["merged_cells"]["count"] > 0
    if cls == "conditional_formatting":
        cf = inv["conditional_formatting"]
        return cf["classic"]["count"] > 0 or cf["x14"]["count"] > 0
    if cls == "data_validations":
        dv = inv["data_validations"]
        return dv["classic"]["count"] > 0 or dv["x14"]["count"] > 0
    if cls == "defined_names":
        return len(inv["defined_names"]) > 0
    if cls == "tables":
        return inv["tables"]["count"] > 0
    if cls == "frozen_panes":
        return inv["frozen_panes"]["count"] > 0
    return False


def workbook_table_names(inv: dict) -> set[str]:
    return {t["name"] for t in inv["tables"]["tables"]}


def workbook_chart_refs(inv: dict) -> set[str]:
    refs = set()
    for chart in inv["charts"]["charts"]:
        for ser in chart["series"]:
            for key in ("cat_ref", "name_ref", "val_ref"):
                if ser.get(key):
                    refs.add(ser[key])
    return refs


def workbook_cf_sqrefs(inv: dict) -> set[tuple[str, str]]:
    out = set()
    for block in inv["conditional_formatting"]["classic"]["blocks"] + inv["conditional_formatting"]["x14"]["blocks"]:
        if block.get("sqref"):
            out.add((block["sheet"], block["sqref"]))
    return out


def workbook_dv_sqrefs(inv: dict) -> set[tuple[str, str]]:
    out = set()
    for block in inv["data_validations"]["classic"]["blocks"] + inv["data_validations"]["x14"]["blocks"]:
        if block.get("sqref"):
            out.add((block["sheet"], block["sqref"]))
    return out


def main(argv: list[str]) -> int:
    tasks_path = Path(argv[1]) if len(argv) > 1 else HERE / "tasks.json"
    errors: list[str] = []
    warnings: list[str] = []

    try:
        tasks = json.loads(tasks_path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        print(f"FATAL: could not read/parse {tasks_path}: {exc}")
        return 1

    if not isinstance(tasks, list):
        print("FATAL: tasks.json must be a JSON list")
        return 1

    seen_ids: set[str] = set()
    inv_cache: dict[str, dict] = {}

    for idx, t in enumerate(tasks):
        loc = f"task[{idx}] (id={t.get('id', '<missing>')!r})"

        if not isinstance(t, dict):
            errors.append(f"{loc}: not a JSON object")
            continue

        missing = REQUIRED_KEYS - t.keys()
        if missing:
            errors.append(f"{loc}: missing required keys: {sorted(missing)}")
            continue

        tid = t["id"]
        if not isinstance(tid, str) or not tid:
            errors.append(f"{loc}: id must be a non-empty string")
        elif tid in seen_ids:
            errors.append(f"{loc}: duplicate id {tid!r}")
        else:
            seen_ids.add(tid)

        if t["tier"] not in VALID_TIERS:
            errors.append(f"{loc}: tier {t['tier']!r} not in 1-5")

        if t["kind"] not in VALID_KINDS:
            errors.append(f"{loc}: kind {t['kind']!r} not one of {sorted(VALID_KINDS)}")

        if not isinstance(t["prompt"], str) or not t["prompt"].strip():
            errors.append(f"{loc}: prompt must be a non-empty string")

        if not isinstance(t["notes"], str) or not t["notes"].strip():
            errors.append(f"{loc}: notes must be a non-empty string")

        # workbook must exist
        wb_name = t["workbook"]
        wb_path = CORPUS_DIR / wb_name
        if not wb_path.exists():
            errors.append(f"{loc}: workbook {wb_name!r} not found in corpus/")
            continue

        # load (and cache) inventory of the ORIGINAL workbook
        if wb_name not in inv_cache:
            try:
                inv_cache[wb_name] = inventory_mod.inventory(str(wb_path))
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{loc}: inventory.py failed on {wb_name}: {exc}")
                continue
        inv = inv_cache[wb_name]

        # preserve: list of known classes, each must be a real feature of this workbook
        preserve = t["preserve"]
        if not isinstance(preserve, list) or not preserve:
            errors.append(f"{loc}: preserve must be a non-empty list")
        else:
            for cls in preserve:
                if cls not in PRESERVE_CLASSES:
                    errors.append(f"{loc}: preserve class {cls!r} is not a recognized feature class")
                elif not has_feature(inv, cls):
                    errors.append(
                        f"{loc}: preserve class {cls!r} claimed but workbook {wb_name!r} "
                        "has none (per inventory.py)"
                    )

        # expect: keys must be from the known set, cells/text_cells must reference the
        # right sheet+ref shape, tables/chart_series/cf_sqrefs/dv_sqrefs must name
        # something that exists pre-change OR be flagged as an expected change in notes
        expect = t["expect"]
        if not isinstance(expect, dict):
            errors.append(f"{loc}: expect must be an object")
            continue

        unknown_expect_keys = set(expect.keys()) - EXPECT_KEYS
        if unknown_expect_keys:
            errors.append(f"{loc}: unknown expect keys {sorted(unknown_expect_keys)}")

        for key in ("cells", "text_cells"):
            for entry in expect.get(key, []):
                for req in ("sheet", "ref", "value"):
                    if req not in entry:
                        errors.append(f"{loc}: expect.{key} entry missing {req!r}: {entry}")
                if key == "cells" and "value" in entry and not isinstance(entry["value"], (int, float)):
                    errors.append(f"{loc}: expect.cells value must be numeric: {entry}")
                if key == "text_cells" and "value" in entry and not isinstance(entry["value"], str):
                    errors.append(f"{loc}: expect.text_cells value must be a string: {entry}")

        for entry in expect.get("formulas", []):
            for req in ("sheet", "ref", "contains"):
                if req not in entry:
                    errors.append(f"{loc}: expect.formulas entry missing {req!r}: {entry}")

        notes_text = t["notes"]

        # tables
        orig_tables = workbook_table_names(inv)
        for entry in expect.get("tables", []):
            name = entry.get("name")
            ref = entry.get("ref")
            if name is None or ref is None:
                errors.append(f"{loc}: expect.tables entry missing name/ref: {entry}")
                continue
            if name not in orig_tables:
                errors.append(
                    f"{loc}: expect.tables names {name!r}, not a table in the original "
                    f"{wb_name!r} (tables present: {sorted(orig_tables)})"
                )
                continue
            orig_ref = next(x["ref"] for x in inv["tables"]["tables"] if x["name"] == name)
            if ref != orig_ref and ref not in notes_text:
                errors.append(
                    f"{loc}: expect.tables {name!r} ref {ref!r} differs from original "
                    f"{orig_ref!r} and is not mentioned in notes as the expected new ref"
                )

        # chart_series
        orig_chart_refs = workbook_chart_refs(inv)
        for entry in expect.get("chart_series", []):
            part = entry.get("part")
            if part is None:
                errors.append(f"{loc}: expect.chart_series entry missing part: {entry}")
                continue
            parts_present = {c["part"] for c in inv["charts"]["charts"]}
            if part not in parts_present:
                errors.append(
                    f"{loc}: expect.chart_series part {part!r} not a chart part in the "
                    f"original {wb_name!r} (parts present: {sorted(parts_present)})"
                )
                continue
            for key in ("cat_ref", "val_ref", "name_ref"):
                if key not in entry:
                    continue
                ref = entry[key]
                if ref in orig_chart_refs:
                    continue
                if ref in notes_text:
                    continue
                errors.append(
                    f"{loc}: expect.chart_series {key}={ref!r} is neither an existing "
                    f"chart ref in {wb_name!r} nor mentioned in notes as the expected "
                    "post-change ref"
                )

        # cf_sqrefs / dv_sqrefs
        orig_cf = workbook_cf_sqrefs(inv)
        for entry in expect.get("cf_sqrefs", []):
            sheet, sqref = entry.get("sheet"), entry.get("sqref")
            if sheet is None or sqref is None:
                errors.append(f"{loc}: expect.cf_sqrefs entry missing sheet/sqref: {entry}")
                continue
            if (sheet, sqref) in orig_cf:
                continue
            # sheet may have been renamed (sqref unchanged) -- accept if the sqref alone
            # matches an original CF block AND the new sheet name is mentioned in notes
            sqref_matches_any_sheet = any(s == sqref for _, s in orig_cf)
            if sqref_matches_any_sheet and sheet in notes_text:
                continue
            if (sheet, sqref) not in orig_cf and sqref in notes_text:
                continue
            errors.append(
                f"{loc}: expect.cf_sqrefs {sheet!r}!{sqref!r} not found in original "
                f"{wb_name!r} CF blocks and not clearly explained as a post-change state "
                "in notes"
            )

        orig_dv = workbook_dv_sqrefs(inv)
        for entry in expect.get("dv_sqrefs", []):
            sheet, sqref = entry.get("sheet"), entry.get("sqref")
            if sheet is None or sqref is None:
                errors.append(f"{loc}: expect.dv_sqrefs entry missing sheet/sqref: {entry}")
                continue
            if (sheet, sqref) in orig_dv:
                continue
            sqref_matches_any_sheet = any(s == sqref for _, s in orig_dv)
            if sqref_matches_any_sheet and sheet in notes_text:
                continue
            errors.append(
                f"{loc}: expect.dv_sqrefs {sheet!r}!{sqref!r} not found in original "
                f"{wb_name!r} DV blocks and not clearly explained as a post-change state "
                "in notes"
            )

        # sheet_names sanity: every name should either exist in the original sheet list
        # or be explained (new/renamed) in notes
        orig_sheets = set(inv["sheet_names"])
        for name in expect.get("sheet_names", []):
            if name in orig_sheets:
                continue
            if name in notes_text or name in t["prompt"]:
                continue
            warnings.append(
                f"{loc}: expect.sheet_names has {name!r}, not an original sheet name and "
                "not obviously explained in notes/prompt (soft warning)"
            )

    # tier/workbook coverage summary (not an error condition -- substitutions are allowed
    # and documented per-task in `notes`)
    by_tier: dict[int, int] = {}
    by_workbook: dict[str, list[int]] = {}
    for t in tasks:
        if "tier" in t:
            by_tier[t["tier"]] = by_tier.get(t["tier"], 0) + 1
        by_workbook.setdefault(t.get("workbook", "?"), []).append(t.get("tier"))

    print(f"Loaded {len(tasks)} tasks from {tasks_path}")
    print(f"by_tier: {dict(sorted(by_tier.items()))}")
    for wb, tiers in sorted(by_workbook.items()):
        print(f"  {wb}: tiers {sorted(tiers)}")

    if warnings:
        print(f"\n{len(warnings)} warning(s):")
        for w in warnings:
            print(f"  - {w}")

    if errors:
        print(f"\n{len(errors)} error(s):")
        for e in errors:
            print(f"  - {e}")
        return 1

    print("\nOK: 0 errors")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
