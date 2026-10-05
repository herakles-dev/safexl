"""Pytest checks for research/xlsx-bench/grade.py.

Every fixture here is built via direct zip-level edits on a COPY of a
corpus file (never openpyxl save+re-save), so each test isolates exactly
one feature loss with everything else byte-identical to the source
workbook. The one deliberate exception is
`test_openpyxl_roundtrip_is_not_preserved`, which does the real-world
failure the grader exists to catch: a naive load+save round-trip that
silently drops the x14 conditional-formatting block and the embedded
image. No Claude or SafeXL runs are needed — grade() is exercised
directly against synthetic (task, original, output) triples.
"""
from __future__ import annotations

import importlib.util
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

from _corpus import require

HERE = Path(__file__).resolve().parent
BENCH_DIR = HERE.parent
CORPUS_DIR = BENCH_DIR / "corpus"

_spec = importlib.util.spec_from_file_location("xlsx_bench_grade", BENCH_DIR / "grade.py")
grade_module = importlib.util.module_from_spec(_spec)
sys.modules["xlsx_bench_grade"] = grade_module
_spec.loader.exec_module(grade_module)
grade = grade_module.grade
inventory = grade_module.inventory


def corpus_path(name: str) -> Path:
    return require(CORPUS_DIR / name)


def _edit_zip(src: Path, dst: Path, replace: dict | None = None, remove: set | None = None) -> None:
    """Copy `src` to `dst` as a zip, replacing/removing specific members by
    filename. Every OOXML edit in this file goes through this helper —
    never openpyxl.save() — producing precise, minimal, single-feature
    diffs from the corpus file.
    """
    replace = replace or {}
    remove = remove or set()
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in remove:
                continue
            data = zin.read(item.filename)
            if item.filename in replace:
                data = replace[item.filename]
            zout.writestr(item, data)


def _read(src: Path, member: str) -> bytes:
    with zipfile.ZipFile(src) as zf:
        return zf.read(member)


def _task(id_, workbook, preserve=(), expect=None, tier=1):
    return {
        "id": id_,
        "workbook": workbook,
        "tier": tier,
        "kind": "value",
        "prompt": "test fixture",
        "expect": expect or {},
        "preserve": list(preserve),
        "notes": "synthetic task for grade.py tests",
    }


ALL_PRESERVE_CLASSES = [
    "vba", "images", "charts", "pivot_tables", "pivot_caches", "comments",
    "hyperlinks", "protection", "sparklines", "external_links",
    "merged_cells", "conditional_formatting", "data_validations",
    "defined_names", "tables",
]


# --------------------------------------------------------------------------
# (a) identical copy -> preserved + valid True
# --------------------------------------------------------------------------

def test_identical_copy_is_preserved_and_valid(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "identical.xlsx"
    shutil.copy2(src, dst)

    task = _task("identical", "simple-monthly-budget.xlsx", preserve=ALL_PRESERVE_CLASSES)
    result = grade(task, src, dst)

    assert result["valid"] is True, result["checks"]
    assert result["preserved"] is True, result["losses"]
    assert result["losses"] == []
    assert result["correct"] is True  # empty expect -> vacuously true
    assert result["pass"] is True


# --------------------------------------------------------------------------
# (b) xl/vbaProject.bin removed -> preserved False, vba loss named
# --------------------------------------------------------------------------

def test_vba_removed_is_not_preserved(tmp_path):
    src = corpus_path("benefits-contribution-calculator.xlsm")
    dst = tmp_path / "no-vba.xlsm"
    _edit_zip(src, dst, remove={"xl/vbaProject.bin"})

    task = _task("vba-removed", "benefits-contribution-calculator.xlsm", preserve=["vba"])
    result = grade(task, src, dst)

    assert result["preserved"] is False
    assert result["losses"] == ["vba: vbaProject.bin missing"], result["losses"]


# --------------------------------------------------------------------------
# (c) one media part removed -> images loss, part named
# --------------------------------------------------------------------------

def test_media_part_removed_is_not_preserved(tmp_path):
    src = corpus_path("parking-study-report.xlsm")  # 2 embedded images
    dst = tmp_path / "missing-image.xlsm"

    orig_inv = inventory(str(src))
    assert orig_inv["images"]["count"] == 2
    removed_part = orig_inv["images"]["media"][0]["part"]

    _edit_zip(src, dst, remove={removed_part})

    task = _task("image-removed", "parking-study-report.xlsm", preserve=["images"])
    result = grade(task, src, dst)

    assert result["preserved"] is False
    assert len(result["losses"]) == 1
    loss = result["losses"][0]
    assert loss.startswith("images:")
    assert "1 of 2" in loss
    assert removed_part in loss


# --------------------------------------------------------------------------
# (d) x14 CF block stripped -> conditional_formatting loss
# --------------------------------------------------------------------------

def test_x14_cf_block_stripped_is_not_preserved(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "no-x14-cf.xlsx"

    sheet_xml = _read(src, "xl/worksheets/sheet1.xml").decode("utf-8")
    marker_start = sheet_xml.index('<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"')
    marker_end = sheet_xml.index("</extLst>", marker_start) + len("</extLst>")
    assert "x14:conditionalFormattings" in sheet_xml[marker_start:marker_end]
    stripped = sheet_xml[:marker_start] + sheet_xml[marker_end:]
    assert "x14:conditionalFormattings" not in stripped

    _edit_zip(src, dst, replace={"xl/worksheets/sheet1.xml": stripped.encode("utf-8")})

    # sanity: exactly the x14 block is gone; classic CF on the same sheet survives
    orig_inv = inventory(str(src))
    out_inv = inventory(str(dst))
    assert orig_inv["conditional_formatting"]["x14"]["count"] == 1
    assert out_inv["conditional_formatting"]["x14"]["count"] == 0
    assert out_inv["conditional_formatting"]["classic"]["count"] == 1

    task = _task("cf-x14-stripped", "simple-monthly-budget.xlsx", preserve=["conditional_formatting"])
    result = grade(task, src, dst)

    assert result["preserved"] is False
    assert result["losses"] == ["conditional_formatting: x14 block on E4 removed"], result["losses"]


# --------------------------------------------------------------------------
# (e) a Content_Types entry removed -> valid False
# --------------------------------------------------------------------------

def test_content_types_gap_makes_output_invalid(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "bad-content-types.xlsx"

    ct_xml = _read(src, "[Content_Types].xml").decode("utf-8")
    # customXml/item1.xml has no explicit <Override>; it is covered only by
    # the <Default Extension="xml"> fallback. Removing that Default leaves
    # exactly that one part uncovered, without touching anything else.
    default_xml_entry = '<Default Extension="xml" ContentType="application/xml"/>'
    assert default_xml_entry in ct_xml
    stripped = ct_xml.replace(default_xml_entry, "", 1)
    assert stripped != ct_xml

    _edit_zip(src, dst, replace={"[Content_Types].xml": stripped.encode("utf-8")})

    task = _task("content-types-gap", "simple-monthly-budget.xlsx")
    result = grade(task, src, dst)

    assert result["valid"] is False
    assert result["pass"] is False
    ct_checks = [c for c in result["checks"] if c["name"] == "content_types_coverage"]
    assert ct_checks and ct_checks[0]["ok"] is False
    assert "customXml/item1.xml" in ct_checks[0]["detail"]


# --------------------------------------------------------------------------
# (f) a correct cell value patched -> expect.cells passes after recalc;
#     a wrong expected value -> fails
# --------------------------------------------------------------------------

def test_correct_cell_value_passes_after_recalc(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "income-raised.xlsx"

    sheet_xml = _read(src, "xl/worksheets/sheet1.xml").decode("utf-8")
    old_cell = '<c r="C5" s="1"><v>2500</v></c>'
    new_cell = '<c r="C5" s="1"><v>5000</v></c>'
    assert old_cell in sheet_xml
    patched = sheet_xml.replace(old_cell, new_cell, 1)

    _edit_zip(src, dst, replace={"xl/worksheets/sheet1.xml": patched.encode("utf-8")})

    # TotalMonthlyIncome = SUM(tblIncome[Amount]) = C5+C6+C7 = 5000+1000+250 = 6250
    task_pass = _task(
        "income-raise-correct",
        "simple-monthly-budget.xlsx",
        expect={"cells": [{"sheet": "Simple Monthly Budget", "ref": "E9", "value": 6250.0, "tol": 0.005}]},
    )
    result_pass = grade(task_pass, src, dst)
    assert result_pass["correct"] is True, result_pass["checks"]

    task_fail = _task(
        "income-raise-wrong",
        "simple-monthly-budget.xlsx",
        expect={"cells": [{"sheet": "Simple Monthly Budget", "ref": "E9", "value": 9999.0, "tol": 0.005}]},
    )
    result_fail = grade(task_fail, src, dst)
    assert result_fail["correct"] is False
    cell_checks = [c for c in result_fail["checks"] if c["name"].startswith("cell:")]
    assert cell_checks and cell_checks[0]["ok"] is False


# --------------------------------------------------------------------------
# real-world failure mode: an openpyxl load+save round-trip of
# simple-monthly-budget.xlsx must be caught as NOT preserved
# --------------------------------------------------------------------------

def test_openpyxl_roundtrip_is_not_preserved(tmp_path):
    import openpyxl

    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "openpyxl-roundtrip.xlsx"

    wb = openpyxl.load_workbook(str(src))
    wb.save(str(dst))
    wb.close()

    task = _task(
        "openpyxl-roundtrip",
        "simple-monthly-budget.xlsx",
        preserve=["images", "conditional_formatting", "tables", "defined_names", "merged_cells", "charts"],
    )
    result = grade(task, src, dst)

    assert result["preserved"] is False, "openpyxl round-trip must be caught, not silently accepted"
    assert any(l.startswith("images:") for l in result["losses"]), result["losses"]
    assert any(l.startswith("conditional_formatting:") for l in result["losses"]), result["losses"]


# --------------------------------------------------------------------------
# extra rigor: the preserve-vs-expect exemption for a legitimately moved
# table ref (same pattern applies to cf_sqrefs/dv_sqrefs/defined_names)
# --------------------------------------------------------------------------

def test_table_ref_move_without_expect_is_a_loss(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "table-moved.xlsx"
    table_xml = _read(src, "xl/tables/table1.xml").decode("utf-8")
    assert 'ref="B4:C7"' in table_xml
    patched = table_xml.replace('ref="B4:C7"', 'ref="B4:C8"', 1)
    _edit_zip(src, dst, replace={"xl/tables/table1.xml": patched.encode("utf-8")})

    task = _task("table-moved-unasserted", "simple-monthly-budget.xlsx", preserve=["tables"])
    result = grade(task, src, dst)

    assert result["preserved"] is False
    assert result["losses"] == [
        "tables: tblIncome ref changed ('B4:C7' -> 'B4:C8', not asserted in expect.tables)"
    ], result["losses"]


def test_table_ref_move_asserted_in_expect_is_exempted(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "table-moved-asserted.xlsx"
    table_xml = _read(src, "xl/tables/table1.xml").decode("utf-8")
    patched = table_xml.replace('ref="B4:C7"', 'ref="B4:C8"', 1)
    _edit_zip(src, dst, replace={"xl/tables/table1.xml": patched.encode("utf-8")})

    task = _task(
        "table-moved-asserted",
        "simple-monthly-budget.xlsx",
        preserve=["tables"],
        expect={"tables": [{"name": "tblIncome", "ref": "B4:C8"}]},
    )
    result = grade(task, src, dst)

    assert result["preserved"] is True, result["losses"]
    table_checks = [c for c in result["checks"] if c["name"] == "table:tblIncome"]
    assert table_checks and table_checks[0]["ok"] is True


# --------------------------------------------------------------------------
# robustness: missing/corrupt outputs never crash the grader
# --------------------------------------------------------------------------

def test_missing_output_path_fails_everything(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    missing = tmp_path / "does-not-exist.xlsx"
    task = _task("missing-output", "simple-monthly-budget.xlsx")
    result = grade(task, src, missing)
    assert (result["correct"], result["preserved"], result["valid"], result["pass"]) == (False, False, False, False)
    assert result["checks"] and result["checks"][0]["ok"] is False


def test_none_output_fails_everything(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    task = _task("missing-output-none", "simple-monthly-budget.xlsx")
    result = grade(task, src, None)
    assert result["pass"] is False


def test_corrupt_zip_is_invalid(tmp_path):
    src = corpus_path("simple-monthly-budget.xlsx")
    dst = tmp_path / "corrupt.xlsx"
    dst.write_bytes(b"not a zip file at all")
    task = _task("corrupt", "simple-monthly-budget.xlsx")
    result = grade(task, src, dst)
    assert result["valid"] is False
    assert result["pass"] is False
    assert result["checks"][0]["name"] == "unzip"
    assert result["checks"][0]["ok"] is False


def test_find_output_file_in_place_edit(tmp_path):
    """Prompts rarely say "save as": an edited-in-place work copy is the output,
    an untouched one is not, and a new workbook beats a backup copy."""
    import shutil as _shutil
    find_output_file = grade_module.find_output_file
    src = require(Path(__file__).resolve().parents[1] / "corpus" / "simple-monthly-budget.xlsx")
    work = tmp_path / "work"
    work.mkdir()
    _shutil.copy(src, work / src.name)
    assert find_output_file(work, src.name, tmp_path, original_path=src) is None
    _shutil.copy(src, work / "budget-backup.xlsx")
    _shutil.copy(src, work / "budget-updated.xlsx")
    # original untouched: the new workbook is the output, never the backup
    assert find_output_file(work, src.name, tmp_path, original_path=src).name == "budget-updated.xlsx"
    with open(work / src.name, "ab") as fh:
        fh.write(b"\0")  # any byte change counts as "edited"
    # original edited in place: it is the output, even with new workbooks around (#17)
    assert find_output_file(work, src.name, tmp_path, original_path=src) == work / src.name


def test_in_place_final_beats_intermediates_listed_in_run_json(tmp_path):
    """#17: a session writes output.xlsx (intermediate), then final.xlsx, then
    moves final onto the original name. run.json lists output.xlsx first; the
    edited original is still the deliverable."""
    import json as _json, shutil as _shutil
    src = require(Path(__file__).resolve().parents[1] / "corpus" / "simple-monthly-budget.xlsx")
    work = tmp_path / "work"; work.mkdir()
    _shutil.copy(src, work / "output.xlsx")
    _shutil.copy(src, work / src.name)
    with open(work / src.name, "ab") as fh:
        fh.write(b"\0")
    (tmp_path / "run.json").write_text(_json.dumps({"output_files": ["output.xlsx", src.name]}))
    assert grade_module.find_output_file(work, src.name, tmp_path, original_path=src) == work / src.name


def test_dangling_relationship_fails_validity(tmp_path):
    """calcChain.xml dropped but still referenced from workbook.xml.rels: LibreOffice
    opens it fine, Excel may not, so validity must fail and name the target."""
    src = require(Path(__file__).resolve().parents[1] / "corpus" / "simple-monthly-budget.xlsx")
    bad = tmp_path / "bad.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename != "xl/calcChain.xml":
                zout.writestr(item, zin.read(item.filename))
    checks, ok = grade_module._check_validity_structure(bad)
    assert ok is False
    rel = next(c for c in checks if c["name"] == "relationship_targets")
    assert rel["ok"] is False and "calcChain.xml" in rel["detail"]
    checks_ok, ok_src = grade_module._check_validity_structure(src)
    assert ok_src is True


def test_rename_is_compared_across_the_rename():
    """A correct sheet rename moves every sheet-scoped feature to the new name;
    that must not read as 'removed'."""
    orig = {"conditional_formatting": {"classic": {"blocks": [{"sheet": "Simple Monthly Budget", "sqref": "E4"}]}},
            "charts": {"charts": [{"series": [{"val_ref": "'Simple Monthly Budget'!$E$9:$F$9"}]}]},
            "defined_names": {"X": "Data!$A$1", "Y": "'Simple Monthly Budget'!$B$2"},
            "sheet_names": ["Simple Monthly Budget", "Data"]}
    moved = grade_module.apply_sheet_renames(orig, {"Simple Monthly Budget": "October Budget", "Data": "Raw"})
    assert moved["conditional_formatting"]["classic"]["blocks"][0]["sheet"] == "October Budget"
    assert moved["charts"]["charts"][0]["series"][0]["val_ref"] == "'October Budget'!$E$9:$F$9"
    assert moved["defined_names"] == {"X": "Raw!$A$1", "Y": "'October Budget'!$B$2"}
    assert moved["sheet_names"] == ["October Budget", "Raw"]


def test_sheet_rename_map_uses_worksheet_parts(tmp_path):
    src = require(Path(__file__).resolve().parents[1] / "corpus" / "simple-monthly-budget.xlsx")
    renamed = tmp_path / "renamed.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(renamed, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.replace(b'name="Simple Monthly Budget"', b'name="October Budget"')
            zout.writestr(item, data)
    assert grade_module.sheet_rename_map(src, renamed) == {"Simple Monthly Budget": "October Budget"}
    assert grade_module.sheet_rename_map(src, src) == {}


def test_rename_moves_sheet_keyed_dicts():
    orig = {"merged_cells": {"by_sheet": {"Simple Monthly Budget": ["H4:H5"]}, "count": 1}}
    moved = grade_module.apply_sheet_renames(orig, {"Simple Monthly Budget": "October Budget"})
    assert moved == {"merged_cells": {"by_sheet": {"October Budget": ["H4:H5"]}, "count": 1}}


def test_column_values_finds_the_new_column_anywhere():
    spec = [{"sheet": "S", "rows": {"5": 0.13, "88": 0.5}, "tol": 0.001}]
    orig = {"values": {("S", "B5"): 163810, ("S", "M5"): 0}}
    out_ea = {"values": {("S", "B5"): 163810, ("S", "M5"): 0, ("S", "EA5"): 0.13, ("S", "EA88"): 0.5}}
    r = grade_module._check_column_values(spec, out_ea, orig)[0]
    assert r["ok"] and "EA" in r["detail"]
    out_wrong = {"values": {("S", "EA5"): 0.2, ("S", "EA88"): 0.5}}
    assert not grade_module._check_column_values(spec, out_wrong, orig)[0]["ok"]
    # a column that already held the values in the original doesn't count as new
    same = {"values": {("S", "C5"): 0.13, ("S", "C88"): 0.5}}
    assert not grade_module._check_column_values(spec, same, same)[0]["ok"]


def test_sheet_contains_finds_values_anywhere_as_formulas(tmp_path):
    import openpyxl as _ox
    wb = _ox.Workbook(); ws = wb.active; ws.title = "Summary"
    ws["A4"], ws["B4"] = "Metric", "Value"
    ws["B5"] = "='Enrollment Restrictions'!C29"
    ws["B6"] = 1262503  # a typed-in value, not a formula
    out = tmp_path / "s.xlsx"; wb.save(out)  # test fixture only; never a corpus file
    recalc = {"values": {("Summary", "B5"): 2303558, ("Summary", "B6"): 1262503}}
    spec = [{"sheet": "Summary", "values": [{"value": 2303558, "tol": 0.5, "formula_contains": "Enrollment Restrictions"},
                                            {"value": 1262503, "tol": 0.5, "formula_contains": "SUM("}]}]
    r = grade_module._check_sheet_contains(spec, recalc, out)
    assert r[0]["ok"] and "B5" in r[0]["detail"]
    assert not r[1]["ok"]  # right number, but typed in rather than a SUM formula


def test_pivot_cache_range_growth_is_not_a_loss_but_shrink_is():
    def inv(ref):
        return {"pivot_caches": {"count": 1, "caches": [{"source_type": "worksheet", "source_sheet": "Sheet1", "source_ref": ref, "source_name": None}]}}
    assert grade_module._diff_pivot_caches(inv("A1:D13"), inv("A1:D14")) == []
    assert grade_module._diff_pivot_caches(inv("A1:D13"), inv("A1:D12"))
    assert grade_module._diff_pivot_caches(inv("A1:D13"), {"pivot_caches": {"count": 0, "caches": []}})


def test_find_output_ignores_unzipped_parts_listed_in_run_json(tmp_path):
    import json as _json, shutil as _shutil
    src = require(Path(__file__).resolve().parents[1] / "corpus" / "simple-monthly-budget.xlsx")
    work = tmp_path / "work"; (work / "x").mkdir(parents=True)
    _shutil.copy(src, work / src.name)
    with open(work / src.name, "ab") as fh:
        fh.write(b"\0")
    (work / "x" / "[Content_Types].xml").write_text("<Types/>")
    (tmp_path / "run.json").write_text(_json.dumps({"output_files": ["x/[Content_Types].xml"]}))
    assert grade_module.find_output_file(work, src.name, tmp_path, original_path=src) == work / src.name


def test_cf_block_extended_over_new_column_is_preserved():
    def inv(sqref):
        return {"conditional_formatting": {"classic": {"blocks": [{"sheet": "S", "sqref": sqref, "rule_types": ["expression"]}]},
                                           "x14": {"blocks": []}}}
    diff = lambda o, n: grade_module._diff_sqref_class(o, n, [], "conditional_formatting", grade_module._cf_blocks_of)
    assert diff(inv("A2:G2"), inv("A2:H2")) == []          # extended over the new column
    assert diff(inv("A2:G2"), inv("A2:F2"))                 # shrank: a real loss
    assert grade_module._sqref_covers("A2:H2 J5", "A2:G2 J5")


def test_sheet_contains_follows_same_sheet_helper_cells(tmp_path):
    """parking-t1sub, arm A: B7 = B6/B5 with B5/B6 pulling from Summary_report
    is as live a link as the one-line form. A typed-in chain still fails."""
    import openpyxl as _ox
    wb = _ox.Workbook(); ws = wb.active; ws.title = "Pilot KPIs"
    ws["B5"], ws["B6"] = "=Summary_report!C33", "=Summary_report!D33"
    ws["B7"] = '=IF(B5=0,"n/a",B6/B5)'
    ws["C5"], ws["C6"], ws["C7"] = 40, 40, "=C6/C5"  # right answer, no link to the source
    out = tmp_path / "k.xlsx"; wb.save(out)  # test fixture only
    spec = [{"sheet": "Pilot KPIs", "values": [{"value": 1.0, "tol": 0.001, "formula_contains": "Summary_report"}]}]
    r = grade_module._check_sheet_contains(spec, {"values": {("Pilot KPIs", "B7"): 1.0}}, out)
    assert r[0]["ok"] and "B7" in r[0]["detail"]
    r = grade_module._check_sheet_contains(spec, {"values": {("Pilot KPIs", "C7"): 1.0}}, out)
    assert not r[0]["ok"]


def test_formula_reach_ignores_other_sheets_refs_and_function_names():
    formulas = {"A1": "=Other!B2+LOG10(3)", "B2": "=Summary_report!A1"}
    assert not grade_module._formula_reaches("A1", "Summary_report", formulas)  # Other!B2 isn't this sheet's B2
    assert grade_module._formula_reaches("B2", "Summary_report", formulas)


def test_added_sheet_may_go_anywhere_but_existing_order_is_kept():
    orig = ["General", "Summary_report", "Tabelle1", "Tabelle2"]
    want = orig + ["Pilot KPIs"]
    ok = lambda actual: grade_module._check_sheet_names(want, {"sheet_names": actual}, orig)[0]["ok"]
    assert ok(["General", "Summary_report", "Tabelle1", "Tabelle2", "Pilot KPIs"])
    assert ok(["General", "Summary_report", "Pilot KPIs", "Tabelle1", "Tabelle2"])
    assert not ok(["Summary_report", "General", "Pilot KPIs", "Tabelle1", "Tabelle2"])  # moved an existing sheet
    assert not ok(["General", "Summary_report", "Tabelle1", "Tabelle2"])  # sheet not added
    # a rename (same count) still needs the exact order
    renamed = ["General", "Summary", "Tabelle1", "Tabelle2"]
    assert not grade_module._check_sheet_names(renamed, {"sheet_names": ["Summary", "General", "Tabelle1", "Tabelle2"]}, orig)[0]["ok"]


def test_column_values_any_of_accepts_each_reading_and_abs():
    recalc = lambda v: {"values": {("Sheet2", "G1"): 1, ("Sheet2", "EU1"): v}}
    spec = [{"sheet": "Sheet2", "any_of": [{"1": 50.0}, {"1": 364.0}], "abs": True, "tol": 0.5}]
    assert grade_module._check_column_values(spec, recalc(364), {"values": {}})[0]["ok"]
    assert grade_module._check_column_values(spec, recalc(-50), {"values": {}})[0]["ok"]
    assert not grade_module._check_column_values(spec, recalc(7), {"values": {}})[0]["ok"]


def test_sheet_contains_formula_contains_accepts_a_list(tmp_path):
    import openpyxl as _ox
    wb = _ox.Workbook(); ws = wb.active; ws.title = "QA Notes"
    ws["B4"] = "=Summary_report!E33"
    out = tmp_path / "q.xlsx"; wb.save(out)  # test fixture only
    recalc = {"values": {("QA Notes", "B4"): 0}}
    ok = lambda needle: grade_module._check_sheet_contains(
        [{"sheet": "QA Notes", "values": [{"value": 0, "tol": 0.5, "formula_contains": needle}]}], recalc, out)[0]["ok"]
    assert ok(["Detailed_report", "Summary_report"])
    assert not ok("Detailed_report")


def test_defined_names_equal_up_to_needless_quotes_and_bool_calls():
    same = grade_module._same_formula
    assert same("KYC!$1:$1", "'KYC'!$1:$1")
    assert same("TRUE", "TRUE()")
    assert not same("'Detail Report'!$Z$6", "Detailed_report!$Z$6")
    assert same("'Detail Report'!$Z$6", "'Detail Report'!$Z$6")  # quotes that ARE needed stay
    assert not same("KYC!$1:$1", "KYC!$1:$2")
