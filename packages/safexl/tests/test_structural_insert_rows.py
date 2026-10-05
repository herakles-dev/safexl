"""safexl_mcp.structural -- insert_rows: shifts existing rows down, rewrites
formula references (own-sheet and cross-sheet) and merged cell ranges that
point at or below the insertion point, and widens <dimension>. See
structural.py's module docstring for exactly what is/isn't recognized in
formula text.
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError
from safexl_mcp.structural import insert_rows

from _util import build_minimal_workbook, build_two_sheet_workbook, read_zip_text, row_order, sheet1_xml


# ---------------------------------------------------------------------------
# Physical row shifting
# ---------------------------------------------------------------------------


def test_insert_rows_shifts_existing_rows_down(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={
            "A1": '<c r="A1"><v>1</v></c>',
            "A2": '<c r="A2"><v>2</v></c>',
            "A3": '<c r="A3"><v>3</v></c>',
        },
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=2, count=1)
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 3, 4]
    assert '<c r="A1"><v>1</v></c>' in xml  # below insertion point: untouched
    assert '<c r="A3"><v>2</v></c>' in xml  # old row 2 -> row 3
    assert '<c r="A4"><v>3</v></c>' in xml  # old row 3 -> row 4


def test_insert_rows_before_first_row_shifts_everything(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "A2": '<c r="A2"><v>2</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=1, count=1)
    xml = sheet1_xml(dst)
    assert row_order(xml) == [2, 3]


def test_insert_rows_multiple_count_shifts_by_count(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"A5": '<c r="A5"><v>5</v></c>'})
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=2, count=3)
    xml = sheet1_xml(dst)
    assert row_order(xml) == [8]


def test_insert_rows_row_below_insertion_point_is_untouched(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "A9": '<c r="A9"><v>9</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=2)
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 11]


# ---------------------------------------------------------------------------
# <dimension>
# ---------------------------------------------------------------------------


def test_insert_rows_updates_dimension(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "C10": '<c r="C10"><v>10</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=2)
    xml = sheet1_xml(dst)
    assert '<dimension ref="A1:C12"/>' in xml


# ---------------------------------------------------------------------------
# Formula references
# ---------------------------------------------------------------------------


def test_insert_rows_updates_relative_formula_reference_at_or_below_insertion(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "B10": '<c r="B10"><f>A5+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=3, count=1)
    xml = sheet1_xml(dst)
    assert "<f>A6+1</f>" in xml


def test_insert_rows_leaves_formula_reference_above_insertion_unchanged(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "B20": '<c r="B20"><f>A2+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=10, count=1)
    xml = sheet1_xml(dst)
    assert "<f>A2+1</f>" in xml


def test_insert_rows_updates_absolute_formula_reference(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "B10": '<c r="B10"><f>$A$5*2</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=1)
    xml = sheet1_xml(dst)
    assert "<f>$A$6*2</f>" in xml


def test_insert_rows_updates_range_formula_reference(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "B20": '<c r="B20"><f>SUM(A3:A10)</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=1)
    xml = sheet1_xml(dst)
    # start (row 3) is above the insertion point and stays; end (row 10) shifts to 11.
    assert "<f>SUM(A3:A11)</f>" in xml


def test_insert_rows_updates_cross_sheet_formula_reference(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet1_cells={"A5": '<c r="A5"><v>5</v></c>'},
        sheet2_cells={"B2": '<c r="B2"><f>Sheet1!A5+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=3, count=1)
    sheet2_out = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<f>Sheet1!A6+1</f>" in sheet2_out


def test_insert_rows_leaves_other_sheets_own_unqualified_refs_unchanged(tmp_path):
    """A bare 'A5' formula living on Sheet2 refers to Sheet2!A5, not
    Sheet1!A5 -- inserting rows into Sheet1 must not touch it."""
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet2_cells={"B2": '<c r="B2"><f>A5+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=1, count=5)
    sheet2_out = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<f>A5+1</f>" in sheet2_out


# ---------------------------------------------------------------------------
# Merged cells
# ---------------------------------------------------------------------------


def test_insert_rows_expands_merge_cell_spanning_insertion(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>'},
        extra_sheet_xml='<mergeCells count="1"><mergeCell ref="B3:C6"/></mergeCells>',
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=1)
    xml = sheet1_xml(dst)
    assert '<mergeCell ref="B3:C7"/>' in xml


def test_insert_rows_shifts_merge_cell_entirely_below_insertion(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>'},
        extra_sheet_xml='<mergeCells count="1"><mergeCell ref="B10:C12"/></mergeCells>',
    )
    dst = tmp_path / "out.xlsx"
    insert_rows(src, dst, "Sheet1", row=5, count=2)
    xml = sheet1_xml(dst)
    assert '<mergeCell ref="B12:C14"/>' in xml


# ---------------------------------------------------------------------------
# Errors / rollback
# ---------------------------------------------------------------------------


def test_insert_rows_sheet_not_found_raises_and_leaves_dst_untouched(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        insert_rows(src, dst, "NoSuchSheet", row=1, count=1)
    assert not dst.exists()


def test_insert_rows_invalid_row_raises(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        insert_rows(src, dst, "Sheet1", row=0, count=1)
    assert not dst.exists()


def test_insert_rows_invalid_count_raises(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        insert_rows(src, dst, "Sheet1", row=1, count=0)
    assert not dst.exists()
