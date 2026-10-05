"""safexl_mcp.diff -- read-only before/after workbook diff.

Exact changes after an edit (both the per-cell content and the OOXML
part-level diff); identical files -> no changes at all.
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError, diff_workbooks, edit

from _util import (
    FIXTURE_NAMES,
    build_minimal_workbook,
    fixture_path,
    first_sheet_name,
)


# ---------------------------------------------------------------------------
# Identical files -> no changes
# ---------------------------------------------------------------------------


def test_diff_same_path_against_itself_shows_no_changes():
    path = fixture_path("WithChart.xlsx")
    d = diff_workbooks(path, path).to_dict()
    assert d["sheets"] == []
    assert d["sheets_added"] == []
    assert d["sheets_removed"] == []
    assert d["parts"]["changed"] == []
    assert d["parts"]["added"] == []
    assert d["parts"]["removed"] == []


def test_diff_byte_identical_copy_shows_no_changes(tmp_path):
    src = fixture_path("WithChart.xlsx")
    copy = tmp_path / "copy.xlsx"
    copy.write_bytes(src.read_bytes())
    d = diff_workbooks(src, copy).to_dict()
    assert d["sheets"] == []
    assert d["parts"]["changed"] == []
    assert d["parts"]["unchanged_count"] > 0


def test_diff_identical_synthetic_workbook_shows_no_changes(tmp_path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    cells = {"A1": '<c r="A1"><v>1</v></c>', "B2": '<c r="B2" t="inlineStr"><is><t xml:space="preserve">x</t></is></c>'}
    build_minimal_workbook(a, cells=cells)
    build_minimal_workbook(b, cells=cells)
    d = diff_workbooks(a, b).to_dict()
    assert d["sheets"] == []
    assert d["parts"]["changed"] == []


# ---------------------------------------------------------------------------
# Exact per-cell changes after an edit (synthetic, deterministic)
# ---------------------------------------------------------------------------


def test_diff_reports_exact_changed_added_cells(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src, cells={"A1": '<c r="A1"><v>1</v></c>'})
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "A1", "value": 99},  # changed
            {"sheet": "Sheet1", "cell": "B1", "value": "new"},  # added
            {"sheet": "Sheet1", "cell": "C1", "formula": "A1+1"},  # added formula
        ],
    )
    d = diff_workbooks(src, dst).to_dict()
    assert len(d["sheets"]) == 1
    sheet = d["sheets"][0]
    assert sheet["sheet"] == "Sheet1"

    changed = {c["cell"]: c for c in sheet["changed_cells"]}
    assert changed["A1"]["before"] == {"value": 1}
    assert changed["A1"]["after"] == {"value": 99}

    added = {c["cell"]: c for c in sheet["added_cells"]}
    assert added["B1"]["after"] == {"value": "new"}
    assert added["C1"]["after"] == {"formula": "A1+1"}

    assert sheet["removed_cells"] == []
    assert d["sheets_added"] == []
    assert d["sheets_removed"] == []


def test_diff_reports_bool_and_formula_before_after_exactly(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src, cells={"A1": '<c r="A1" t="b"><v>0</v></c>', "B1": '<c r="B1"><f>A1</f></c>'})
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "A1", "value": True},
            {"sheet": "Sheet1", "cell": "B1", "formula": "A1*2"},
        ],
    )
    d = diff_workbooks(src, dst).to_dict()
    sheet = d["sheets"][0]
    changed = {c["cell"]: c for c in sheet["changed_cells"]}
    assert changed["A1"]["before"] == {"value": False}
    assert changed["A1"]["after"] == {"value": True}
    assert changed["B1"]["before"] == {"formula": "A1"}
    assert changed["B1"]["after"] == {"formula": "A1*2"}


def test_diff_removed_cell_reported_when_after_workbook_lacks_a_cell(tmp_path):
    """Exercises diff's own removed_cells logic directly (edit() never
    removes cells, so this compares two independently-built workbooks)."""
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    build_minimal_workbook(
        a,
        cells={
            "A1": '<c r="A1"><v>1</v></c>',
            "B1": '<c r="B1" t="inlineStr"><is><t xml:space="preserve">gone</t></is></c>',
        },
    )
    build_minimal_workbook(b, cells={"A1": '<c r="A1"><v>1</v></c>'})

    d = diff_workbooks(a, b).to_dict()
    assert len(d["sheets"]) == 1
    sheet = d["sheets"][0]
    assert sheet["changed_cells"] == []
    assert sheet["added_cells"] == []
    removed = {c["cell"]: c for c in sheet["removed_cells"]}
    assert removed["B1"]["before"] == {"value": "gone"}


def test_diff_parts_reports_changed_worksheet_and_sha256(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src)
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    d = diff_workbooks(src, dst).to_dict()
    changed_paths = {c["path"] for c in d["parts"]["changed"]}
    assert "xl/worksheets/sheet1.xml" in changed_paths
    for c in d["parts"]["changed"]:
        assert c["sha256_before"] != c["sha256_after"]
        assert len(c["sha256_before"]) == 64
        assert len(c["sha256_after"]) == 64


def test_diff_calc_chain_removal_reported_as_removed_part(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src, include_calc_chain=True)
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    d = diff_workbooks(src, dst).to_dict()
    assert "xl/calcChain.xml" in d["parts"]["removed"]
    assert d["parts"]["added"] == []


# ---------------------------------------------------------------------------
# Exact changes across all 7 real fixtures, after the standard smoke-test
# edit shape (1 value + 1 formula at ZZ9000/ZZ9001) -- pytest-native
# equivalent of scripts/smoke_safexl_tools.sh Part (b), with exact cell
# content assertions instead of ref-only checks.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_diff_after_standard_edit_shows_exactly_the_two_injected_cells(tmp_path, name):
    src = fixture_path(name)
    dst = tmp_path / f"out-{name}"
    sheet = first_sheet_name(src)
    edit(
        src,
        dst,
        [
            {"sheet": sheet, "cell": "ZZ9000", "value": "SAFEXL_SMOKE"},
            {"sheet": sheet, "cell": "ZZ9001", "formula": "1+1"},
        ],
    )

    d = diff_workbooks(src, dst).to_dict()
    sheets_by_name = {s["sheet"]: s for s in d["sheets"]}
    assert list(sheets_by_name) == [sheet], f"expected exactly 1 changed sheet, got {list(sheets_by_name)}"

    s = sheets_by_name[sheet]
    assert s["removed_cells"] == []
    all_cells = {c["cell"]: c for c in s["changed_cells"] + s["added_cells"]}
    assert set(all_cells) == {"ZZ9000", "ZZ9001"}
    assert all_cells["ZZ9000"]["after"] == {"value": "SAFEXL_SMOKE"}
    assert all_cells["ZZ9001"]["after"] == {"formula": "1+1"}

    assert d["sheets_added"] == []
    assert d["sheets_removed"] == []

    allowed_extra_changed = {"xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
    worksheet_changed = [p for p in d["parts"]["changed"] if p["path"] not in allowed_extra_changed]
    assert len(worksheet_changed) == 1, worksheet_changed
    assert set(d["parts"]["removed"]) <= {"xl/calcChain.xml"}
    assert d["parts"]["added"] == []


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


def test_diff_missing_file_a_raises(tmp_path):
    b = fixture_path("WithChart.xlsx")
    with pytest.raises(SafeXLError, match="file not found"):
        diff_workbooks(tmp_path / "does-not-exist.xlsx", b)


def test_diff_missing_file_b_raises(tmp_path):
    a = fixture_path("WithChart.xlsx")
    with pytest.raises(SafeXLError, match="file not found"):
        diff_workbooks(a, tmp_path / "does-not-exist.xlsx")


def test_diff_never_writes_to_either_input(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src)
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    src_before = src.read_bytes()
    dst_before = dst.read_bytes()
    diff_workbooks(src, dst)
    assert src.read_bytes() == src_before
    assert dst.read_bytes() == dst_before
