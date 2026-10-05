"""safexl: command-line interface for the SafeXL engine core.

    safexl edit SRC DST --set 'Sheet1!A1=42' --formula 'Sheet1!B1=A1*2'
    safexl insert-rows SRC DST --sheet Sheet1 --row 5 [--count 1]
    safexl rename-sheet SRC DST OLD_NAME NEW_NAME
    safexl add-sheet SRC DST NAME
    safexl insert-cells SRC DST --sheet S --range B5:C6
    safexl inspect FILE [--json]
    safexl diff A B [--json]
    safexl verify BEFORE AFTER [--json] [--strict]
    safexl view PATH [--sheet S] [--range A1:D20] [--find TEXT] [--max-cells N] [--json]
    safexl salvage ORIGINAL DAMAGED -o OUTPUT [--force] [--json]
    safexl guard snapshot --dir SNAP PATH [PATH...]
    safexl guard check --dir SNAP [--restore] [--lenient] [--json]
    safexl resize-table SRC DST --table NAME --ref A1:D20 [--sheet S]

--set, --text and --formula are repeatable; --text never coerces. --set values are coerced: TRUE/FALSE
(case-insensitive) become booleans, values parseable as int/float become
numbers, everything else is written as a string. --formula accepts the
formula with or without a leading '=' (see safexl_mcp.engine for why the
stored <f> element never contains one).

``insert-rows``, ``rename-sheet`` and ``add-sheet`` are the structural edits from
safexl_mcp.structural (same SRC/DST rules as ``edit``: SRC is never modified,
DST is only written on success). ``inspect`` and ``diff`` are read-only and never execute macros -- see
safexl_mcp.inspect / safexl_mcp.diff for details.

``edit`` grows an Excel table when a written cell sits next to it (what Excel does when you type
beside a table); ``--no-extend-tables`` turns that off. Errors keep one convention everywhere:
1 = SafeXLError (missing file, bad input), 2 = usage. Result codes sit outside that range: ``verify``
exits 0 intact, 10 rewritten, 11 damaged, 12 invalid (``--strict`` adds 13: intact, but something risky
was added; 14: intact, nothing risky added, but a defined name or table range changed; 13 wins when both apply);``salvage`` exits 0 salvaged/nothing_to_do,
10 partial, 11 unsupported; ``guard check`` exits 0 when nothing was wrong, 10 when a file was flagged.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

from .diff import WorkbookDiff, diff_workbooks
from .engine import SafeXLError, edit
from .guard import check as guard_check, check_and_restore, snapshot as guard_snapshot
from .inspect import WorkbookInspection, inspect_workbook
from .salvage import salvage
from .refs import TornReference
from .structural import add_sheet, insert_cells, insert_rows, plan_table_extensions, rename_sheet, resize_table
from .verify import verify
from .view import render_text, view


def _coerce_value(raw: str) -> Any:
    if raw.upper() == "TRUE":
        return True
    if raw.upper() == "FALSE":
        return False
    try:
        return int(raw)
    except ValueError:
        pass
    try:
        return float(raw)
    except ValueError:
        pass
    return raw


def _split_sheet_cell_expr(arg: str, flag: str, default_sheet: str | None = None) -> tuple[str, str, str]:
    # With --sheet, a reference has a sheet prefix only if a '!' comes before the first '='
    # (so a value like 'hello!' is not mistaken for one).
    if default_sheet is not None and "!" not in arg.split("=", 1)[0]:
        arg = f"{default_sheet}!{arg}"
    if "!" not in arg:
        raise SystemExit(f"{flag}: expected 'Sheet!Cell=value', got: {arg!r} (missing '!')")
    sheet, rest = arg.split("!", 1)
    if "=" not in rest:
        raise SystemExit(f"{flag}: expected 'Sheet!Cell=value', got: {arg!r} (missing '=')")
    cell, payload = rest.split("=", 1)
    sheet, cell = sheet.strip(), cell.strip()
    if not sheet or not cell:
        raise SystemExit(f"{flag}: expected 'Sheet!Cell=value', got: {arg!r}")
    return sheet, cell, payload


def _parse_set(arg: str, default_sheet: str | None = None) -> dict:
    sheet, cell, raw_value = _split_sheet_cell_expr(arg, "--set", default_sheet)
    return {"sheet": sheet, "cell": cell, "value": _coerce_value(raw_value)}


def _parse_text(arg: str, default_sheet: str | None = None) -> dict:
    sheet, cell, raw_value = _split_sheet_cell_expr(arg, "--text", default_sheet)
    return {"sheet": sheet, "cell": cell, "value": raw_value}


def _parse_formula(arg: str, default_sheet: str | None = None) -> dict:
    sheet, cell, formula = _split_sheet_cell_expr(arg, "--formula", default_sheet)
    if not formula.strip():
        raise SystemExit(f"--formula: empty formula in {arg!r}")
    return {"sheet": sheet, "cell": cell, "formula": formula}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="safexl",
        description="Edit existing Excel workbooks in headless environments by changing only the XML an edit needs, and verify what changed.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    edit_p = sub.add_parser(
        "edit",
        help="Write cell values and/or formulas into a workbook, preserving every other OOXML part",
    )
    edit_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    edit_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    edit_p.add_argument(
        "--set",
        action="append",
        default=[],
        dest="sets",
        metavar="Sheet!Cell=value",
        help="write a literal value (repeatable); TRUE/FALSE/int/float are coerced, else string",
    )
    edit_p.add_argument(
        "--text",
        action="append",
        default=[],
        dest="texts",
        metavar="Sheet!Cell=text",
        help="write a string verbatim, never coerced (repeatable); use for codes like 007 or 1E5 "
        "that --set would turn into numbers",
    )
    edit_p.add_argument(
        "--formula",
        action="append",
        default=[],
        dest="formulas",
        metavar="Sheet!Cell=formula",
        help="write a formula (repeatable); leading '=' optional. No cached value is written -- "
        "the workbook is marked fullCalcOnLoad so the next calculating application computes it",
    )
    edit_p.add_argument(
        "--sheet",
        help="default sheet for references written without a 'Sheet!' prefix "
        "(`--sheet Data --set B2=5` is `--set Data!B2=5`); an explicit prefix wins",
    )
    edit_p.add_argument(
        "--no-extend-tables",
        action="store_false",
        dest="extend_tables",
        help="do not grow an Excel table when a written cell is next to it (default: grow it)",
    )

    ins_p = sub.add_parser(
        "insert-rows",
        help="Insert blank rows, shifting rows below down and rewriting every reference "
        "(formulas, names, tables, charts, CF/DV, merges, ...)",
    )
    ins_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    ins_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    ins_p.add_argument("--sheet", required=True, help="sheet to insert into")
    ins_p.add_argument("--row", required=True, type=int, help="1-based row to insert at; the old row --row moves down")
    ins_p.add_argument("--count", type=int, default=1, help="number of rows to insert (default 1)")

    ren_p = sub.add_parser(
        "rename-sheet",
        help="Rename a sheet and every reference to it (formulas, names, charts, pivot sources, hyperlinks)",
    )
    ren_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    ren_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    ren_p.add_argument("old_name", help="current sheet name")
    ren_p.add_argument("new_name", help="new sheet name")

    add_p = sub.add_parser(
        "add-sheet",
        help="Append a new empty sheet (fill it afterwards with `safexl edit`)",
    )
    add_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    add_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    add_p.add_argument("name", help="name of the new sheet")

    cells_help = (
        "Insert blank cells in a column span only, shifting those columns down and rewriting every "
        "reference. Use it instead of insert-rows when the sheet has other lists or tables side by side "
        "that must not move"
    )
    cells_p = sub.add_parser("insert-cells", help=cells_help, description=cells_help)
    cells_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    cells_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    cells_p.add_argument("--sheet", required=True, help="sheet to insert into")
    cells_p.add_argument(
        "--range",
        required=True,
        dest="cell_range",
        metavar="B5:C6",
        help="cells to insert; only these columns shift down, by the range's height, from its first row",
    )

    inspect_p = sub.add_parser(
        "inspect",
        help="Read-only structural + feature report for a workbook (never executes macros)",
    )
    inspect_p.add_argument("file", help=".xlsx/.xlsm to inspect")
    inspect_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")

    diff_p = sub.add_parser(
        "diff",
        help="Per-sheet cell diff plus OOXML-part diff (by sha256) between two workbooks",
    )
    diff_p.add_argument("a", help="'before' .xlsx/.xlsm")
    diff_p.add_argument("b", help="'after' .xlsx/.xlsm")
    diff_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")

    ver_p = sub.add_parser(
        "verify",
        help="Check an edited workbook against its original: intact, rewritten, damaged or invalid "
        "(exit 0/10/11/12)",
    )
    ver_p.add_argument("before", help="the original .xlsx/.xlsm")
    ver_p.add_argument("after", help="the edited .xlsx/.xlsm")
    ver_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")
    ver_p.add_argument(
        "--strict",
        action="store_true",
        help="exit 13 when the file is intact but something risky was added (external link, hidden sheet, ...); exit 14 when it is intact, nothing risky was added, but a defined name or table range changed (13 wins if both)",
    )

    view_p = sub.add_parser(
        "view",
        help="Read-only compact cell listing, one line per cell (use instead of reading sheet XML)",
    )
    view_p.add_argument("path", help=".xlsx/.xlsm to read")
    view_p.add_argument("--sheet", help="sheet to list (default: the first; with --find, all sheets)")
    view_p.add_argument("--range", dest="cell_range", metavar="A1:D20", help="only cells inside this range")
    view_p.add_argument("--find", help="list cells whose value or formula contains this text (case-insensitive)")
    view_p.add_argument("--max-cells", type=int, default=400, help="cells per call (default 400)")
    view_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")

    sal_p = sub.add_parser(
        "salvage",
        help="Re-apply the cell edits found in a damaged copy onto the original (exit 0 salvaged, "
        "10 partial, 11 unsupported)",
    )
    sal_p.add_argument("original", help="the original .xlsx/.xlsm")
    sal_p.add_argument("damaged", help="the copy another tool re-saved")
    sal_p.add_argument(
        "--force",
        action="store_true",
        help="also rebuild a file that was re-saved without losing anything, so untouched parts are "
        "byte-identical again",
    )
    sal_p.add_argument("-o", "--output", required=True, help="path to write the salvaged workbook to")
    sal_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")

    guard_p = sub.add_parser(
        "guard",
        help="Snapshot workbooks, then detect and undo damage done by another writer",
    )
    guard_sub = guard_p.add_subparsers(dest="guard_command", required=True)
    gs_p = guard_sub.add_parser("snapshot", help="Copy each workbook into the snapshot directory")
    gs_p.add_argument("--dir", required=True, help="snapshot directory (created if missing)")
    gs_p.add_argument("paths", nargs="+", metavar="PATH", help="workbooks to track")
    gc_p = guard_sub.add_parser(
        "check",
        help="Report tracked workbooks that were damaged, rewritten or removed (exit 10 if any)",
    )
    gc_p.add_argument("--dir", required=True, help="snapshot directory from `guard snapshot`")
    gc_p.add_argument("--restore", action="store_true", help="restore flagged workbooks from the snapshot")
    gc_p.add_argument(
        "--lenient",
        action="store_true",
        help="with --restore, leave a merely rewritten file in place (restore only damaged/invalid/missing)",
    )
    gc_p.add_argument("--json", action="store_true", help="emit machine-readable JSON instead of text")

    rt_p = sub.add_parser(
        "resize-table",
        help="Set the range of an Excel table (grow or shrink columns/rows); only the table part changes",
    )
    rt_p.add_argument("src", help="source .xlsx/.xlsm to read (never modified)")
    rt_p.add_argument("dst", help="path to write the edited workbook to (only written on success)")
    rt_p.add_argument("--table", required=True, help="table name or displayName")
    rt_p.add_argument("--ref", required=True, metavar="A1:D20", help="new range; the top-left cell must stay")
    rt_p.add_argument("--sheet", help="only look for the table on this sheet")

    return parser


_VERIFY_EXIT = {"intact": 0, "rewritten": 10, "damaged": 11, "invalid": 12}
_SALVAGE_EXIT = {"salvaged": 0, "nothing_to_do": 0, "partial": 10, "unsupported": 11}


def _cell_change_line(c: dict) -> str:
    """One compact line for a sheet's cell_changes entry, or '' when the sheet has no changes."""
    n = c.get("cells_changed") or 0
    parts = []
    if n:
        parts.append(f"{n} cell{'s' if n != 1 else ''} changed")
    for key, label in (("rows_inserted", "inserted"), ("rows_removed", "removed")):
        k = c.get(key) or 0
        if k:
            parts.append(f"{k} row{'s' if k != 1 else ''} {label}")
    for key in ("added", "removed"):
        k = c.get(key) or 0
        if k and not n:
            parts.append(f"{k} cell{'s' if k != 1 else ''} {key}")
    if not parts:
        return ""
    line = ", ".join(parts)
    if c.get("positional"):
        line += " (positional)"
    return line


def _print_verify(r: dict) -> None:
    print(f"safexl verify: {r['verdict']}")
    if r.get("summary"):
        print(r["summary"])
    for risk in r.get("added_risk", []):
        print(f"  added: {risk['risk']}: {risk['detail']}")
    for ref in r.get("reference_changes") or []:
        print(f"  reference: {ref['kind']} {ref['name']} ({ref['scope']}): {ref['before']} -> {ref['after']}")
    for sheet, c in (r.get("cell_changes") or {}).items():
        line = _cell_change_line(c)
        if line:
            print(f"  {sheet}: {line}")
    for x in r["lost_features"]:
        print(f"  lost {x['feature']}: {x['before']} -> {x['after']}")
        for d in x.get("detail", []):
            print(f"    - {d}")
    for reason in r["reasons"]:
        print(f"  {reason}")
    p = r["parts"]
    print(f"  parts: {p['unchanged']} unchanged, {len(p['changed'])} changed, "
          f"{len(p['added'])} added, {len(p['removed'])} removed")
    for label in ("changed", "added", "removed"):
        for name in p[label]:
            print(f"    {label}: {name}")


def _print_salvage(r: dict, output: str) -> None:
    print(f"safexl salvage: {r['status']}")
    print(f"  applied: {len(r['applied'])}, skipped: {len(r['skipped'])}")
    for a in r["applied"]:
        print(f"    applied {a['sheet']}!{a['ref']} ({a['kind']})")
    for k in r["skipped"]:
        print(f"    skipped {k['sheet']}!{k['ref']}: {k['reason']}")
    for d in r["structural"]["detected"]:
        detail = {k: v for k, v in d.items() if k not in ("kind", "handled")}
        state = "handled" if d["handled"] else "NOT handled"
        print(f"    structural: {d['kind']} {json.dumps(detail)} ({state})")
    if r["status"] in ("salvaged", "partial"):
        print(f"  wrote {output}")
    elif r["status"] == "unsupported":
        print("  no output written")


def _print_table_extensions(plans: list[dict]) -> None:
    for t in plans:
        cols = t["new_columns"]
        bits = ""
        if cols:
            names = ", ".join(f'"{c["name"]}"' for c in cols)
            bits = f' (new column{"s" if len(cols) > 1 else ""} {names})'
        print(f"table {t['table']} extended {t['old_ref']} -> {t['new_ref']}{bits}")
        for c in cols:
            if c["auto_named"]:
                where = f"header cell {c['header_cell']} is empty" if c["header_cell"] else "the table has no header row"
                print(f'  note: column "{c["name"]}" was auto-named ({where}); give it a real header')


def _print_inspect(report: WorkbookInspection) -> None:
    print(f"safexl inspect: {report.path}")
    print(f"  sheets: {len(report.sheets)}")
    for s in report.sheets:
        mismatch = "" if s.dimension_matches else "  <-- MISMATCH"
        print(f"    - {s.name!r} (sheetId={s.sheet_id}): declared={s.declared_dimension!r} "
              f"actual={s.actual_used_range!r}{mismatch}")
        print(f"        formulas={s.formula_count} "
              f"dataValidation(native={s.data_validation_native}, extLst={s.data_validation_extlst}) "
              f"conditionalFormatting(native={s.conditional_formatting_native}, "
              f"extLst={s.conditional_formatting_extlst})")

    print(f"  defined names: {len(report.defined_names)}")
    for dn in report.defined_names:
        scope = f" (sheet-local id={dn['local_sheet_id']})" if dn.get("local_sheet_id") is not None else ""
        print(f"    - {dn['name']}{scope} = {dn['value']}")

    f = report.features
    if f is None:
        return
    print("  features:")
    print(f"    charts={f.charts} drawings={f.drawings} media_files={f.media_files}")
    print(f"    pivot_tables={f.pivot_tables} pivot_caches={f.pivot_caches}")
    print(f"    dataValidation total={f.data_validation_native + f.data_validation_extlst} "
          f"(native={f.data_validation_native}, extLst={f.data_validation_extlst})")
    print(f"    conditionalFormatting total={f.conditional_formatting_native + f.conditional_formatting_extlst} "
          f"(native={f.conditional_formatting_native}, extLst={f.conditional_formatting_extlst})")
    vba = f.vba
    if vba.present:
        print(f"    vba: present, {vba.size_bytes} bytes (xl/vbaProject.bin) -- macros NEVER executed")
        scan = vba.macro_scan
        note = f" -- {scan['note']}" if scan.get("note") else ""
        print(f"      macro scan: {scan['mode']}{note}")
        for hit in scan.get("hits", []):
            print(f"        [{hit['type']}] {hit['keyword']}: {hit['description']}")
    else:
        print("    vba: not present")


def _print_diff(d: WorkbookDiff) -> None:
    print(f"safexl diff: {d.path_before} -> {d.path_after}")
    if d.sheets_added:
        print(f"  sheets added: {d.sheets_added}")
    if d.sheets_removed:
        print(f"  sheets removed: {d.sheets_removed}")
    if not d.sheets:
        print("  no cell changes in any common sheet")
    for s in d.sheets:
        print(f"  sheet {s.sheet!r}:")
        for c in s.changed_cells:
            print(f"    changed {c['cell']}: {c['before']} -> {c['after']}")
        for c in s.added_cells:
            print(f"    added   {c['cell']}: {c['after']}")
        for c in s.removed_cells:
            print(f"    removed {c['cell']}: {c['before']}")

    p = d.parts
    print(f"  OOXML parts: {len(p['changed'])} changed, {len(p['added'])} added, "
          f"{len(p['removed'])} removed, {p['unchanged_count']} unchanged")
    for c in p["changed"]:
        print(f"    changed: {c['path']}")
    for a in p["added"]:
        print(f"    added:   {a}")
    for r in p["removed"]:
        print(f"    removed: {r}")


class _QuietStdout:
    """stdout that swallows a closed pipe (``safexl view ... | head``): the command keeps its real exit code."""

    def __init__(self, real):
        self._real = real

    def __getattr__(self, name):
        return getattr(self._real, name)

    def _closed(self) -> None:
        try:  # later writes, and the flush at interpreter exit, go nowhere
            fd = self._real.fileno()
            os.dup2(os.open(os.devnull, os.O_WRONLY), fd)
        except (OSError, ValueError, AttributeError):
            pass

    def write(self, text):
        try:
            return self._real.write(text)
        except BrokenPipeError:
            self._closed()
            return len(text)

    def flush(self):
        try:
            self._real.flush()
        except BrokenPipeError:
            self._closed()


def main(argv: list[str] | None = None) -> int:
    real = sys.stdout
    sys.stdout = _QuietStdout(real)
    try:
        return _main(argv)
    finally:
        sys.stdout.flush()
        sys.stdout = real


def _main(argv: list[str] | None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "edit":
        changes = (
            [_parse_set(s, args.sheet) for s in args.sets]
            + [_parse_text(t, args.sheet) for t in args.texts]
            + [_parse_formula(f, args.sheet) for f in args.formulas]
        )
        if not changes:
            print("safexl edit: at least one --set, --text or --formula is required", file=sys.stderr)
            return 2
        try:
            plans = plan_table_extensions(args.src, changes) if args.extend_tables else []
            edit(args.src, args.dst, changes, extend_tables=args.extend_tables)
        except SafeXLError as exc:
            print(f"safexl edit: {exc}", file=sys.stderr)
            return 1
        print(f"safexl edit: wrote {args.dst} ({len(changes)} change(s))")
        _print_table_extensions(plans)
        return 0

    if args.command == "insert-rows":
        try:
            insert_rows(args.src, args.dst, args.sheet, args.row, args.count)
        except SafeXLError as exc:
            print(f"safexl insert-rows: {exc}", file=sys.stderr)
            return 1
        print(f"safexl insert-rows: wrote {args.dst} ({args.count} row(s) inserted at {args.sheet}!{args.row})")
        return 0

    if args.command == "insert-cells":
        try:
            insert_cells(args.src, args.dst, sheet=args.sheet, cell_range=args.cell_range)
        except (SafeXLError, TornReference) as exc:
            print(f"safexl insert-cells: {exc}", file=sys.stderr)
            return 1
        print(f"safexl insert-cells: wrote {args.dst} (blank cells at {args.sheet}!{args.cell_range}, "
              "those columns shifted down)")
        return 0

    if args.command == "rename-sheet":
        try:
            rename_sheet(args.src, args.dst, args.old_name, args.new_name)
        except SafeXLError as exc:
            print(f"safexl rename-sheet: {exc}", file=sys.stderr)
            return 1
        print(f"safexl rename-sheet: wrote {args.dst} ({args.old_name!r} -> {args.new_name!r})")
        return 0

    if args.command == "add-sheet":
        try:
            add_sheet(args.src, args.dst, args.name)
        except SafeXLError as exc:
            print(f"safexl add-sheet: {exc}", file=sys.stderr)
            return 1
        print(f"safexl add-sheet: wrote {args.dst} (sheet {args.name!r})")
        return 0

    if args.command == "inspect":
        try:
            report = inspect_workbook(args.file)
        except SafeXLError as exc:
            print(f"safexl inspect: {exc}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(report.to_dict(), indent=2))
        else:
            _print_inspect(report)
        return 0

    if args.command == "diff":
        try:
            result = diff_workbooks(args.a, args.b)
        except SafeXLError as exc:
            print(f"safexl diff: {exc}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(result.to_dict(), indent=2))
        else:
            _print_diff(result)
        return 0

    if args.command == "verify":
        missing = [p for p in (args.before, args.after) if not Path(p).is_file()]
        if missing:
            print(f"safexl verify: file not found: {missing[0]}", file=sys.stderr)
            return 1
        try:
            receipt = verify(args.before, args.after)
        except SafeXLError as exc:
            print(f"safexl verify: {exc}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(receipt, indent=2))
        else:
            _print_verify(receipt)
        if args.strict and receipt["verdict"] == "intact":
            if receipt.get("added_risk"):
                return 13
            if receipt.get("reference_changes"):
                return 14
        return _VERIFY_EXIT[receipt["verdict"]]

    if args.command == "view":
        try:
            result = view(args.path, sheet=args.sheet, cell_range=args.cell_range, find=args.find,
                          max_cells=args.max_cells)
        except SafeXLError as exc:
            print(f"safexl view: {exc}", file=sys.stderr)
            return 1
        print(json.dumps(result, indent=2) if args.json else render_text(result))
        return 0

    if args.command == "salvage":
        try:
            result = salvage(args.original, args.damaged, args.output, force=args.force)
        except SafeXLError as exc:
            print(f"safexl salvage: {exc}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            _print_salvage(result, args.output)
        return _SALVAGE_EXIT[result["status"]]

    if args.command == "guard":
        if args.guard_command == "snapshot":
            try:
                tracked = guard_snapshot(args.paths, args.dir)
            except (SafeXLError, OSError) as exc:
                print(f"safexl guard: {exc}", file=sys.stderr)
                return 1
            missing = [p for p in args.paths if str(Path(p).absolute()) not in tracked]
            for p in missing:
                print(f"safexl guard: not found, not tracked: {p}", file=sys.stderr)
            print(f"safexl guard: snapshot of {len(tracked)} file(s) in {args.dir}")
            return 1 if missing else 0
        try:
            if args.restore:
                events = check_and_restore(args.dir, strict=not args.lenient)
            else:
                events = guard_check(args.dir)
        except (SafeXLError, OSError) as exc:
            print(f"safexl guard: {exc}", file=sys.stderr)
            return 1
        if args.json:
            print(json.dumps(events, indent=2))
        else:
            for ev in events:
                print(ev["message"])
            if not events:
                print("safexl guard: all tracked workbooks are fine")
        return 10 if events else 0

    if args.command == "resize-table":
        try:
            resize_table(args.src, args.dst, table=args.table, ref=args.ref, sheet=args.sheet)
        except SafeXLError as exc:
            print(f"safexl resize-table: {exc}", file=sys.stderr)
            return 1
        print(f"safexl resize-table: wrote {args.dst} (table {args.table!r} -> {args.ref})")
        return 0

    parser.print_help()
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
