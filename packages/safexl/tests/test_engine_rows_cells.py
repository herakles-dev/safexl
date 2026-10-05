"""safexl_mcp.engine -- row/cell insertion order and <dimension> updates.

New rows and cells must always be inserted in ascending ref order (ascending
row, then ascending column within a row), regardless of the order changes
are supplied in or where in the existing sheet a gap falls -- an
"out-of-order" insert (e.g. a new row 2 between existing rows 1 and 3) must
still land in the right place, not get appended at the end.
See engine._insert_row_in_order / engine._inject_cells / engine._update_dimension.
"""
from __future__ import annotations

from safexl_mcp import edit

from _util import build_minimal_workbook, cell_refs_in_row, row_order, sheet1_xml


def _dimension_ref(xml: str) -> str | None:
    import re

    m = re.search(r'<dimension ref="([^"]*)"/>', xml)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# Row ordering
# ---------------------------------------------------------------------------


def test_new_row_inserted_between_existing_rows_stays_ascending(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "A3": '<c r="A3"><v>3</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 2}])
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 2, 3]


def test_new_row_inserted_before_first_row(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A5": '<c r="A5"><v>5</v></c>', "A10": '<c r="A10"><v>10</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 1}])
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 5, 10]


def test_new_row_inserted_after_last_row(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "A2": '<c r="A2"><v>2</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "A100", "value": 100}])
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 2, 100]


def test_multiple_out_of_order_new_rows_end_up_fully_sorted(tmp_path):
    """Changes for rows 7, 3, 5 are supplied in that scrambled order (a dict,
    which the engine iterates in insertion order) against a sheet that only
    has rows 1 and 10 -- the output must still be strictly ascending."""
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "A10": '<c r="A10"><v>10</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "B7", "value": 7},
            {"sheet": "Sheet1", "cell": "B3", "value": 3},
            {"sheet": "Sheet1", "cell": "B5", "value": 5},
        ],
    )
    xml = sheet1_xml(dst)
    assert row_order(xml) == [1, 3, 5, 7, 10]


def test_row_inserted_into_originally_empty_sheet(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={})
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "C5", "value": "x"}])
    xml = sheet1_xml(dst)
    assert row_order(xml) == [5]
    assert cell_refs_in_row(xml, 5) == ["C5"]


# ---------------------------------------------------------------------------
# Cell ordering within a row
# ---------------------------------------------------------------------------


def test_new_cell_inserted_between_existing_cells_in_a_row(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "C1": '<c r="C1"><v>3</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B1", "value": 2}])
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["A1", "B1", "C1"]


def test_new_cell_inserted_before_first_cell_in_a_row(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"C1": '<c r="C1"><v>3</v></c>', "E1": '<c r="E1"><v>5</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 1}])
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["A1", "C1", "E1"]


def test_new_cell_inserted_after_last_cell_in_a_row(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "C1": '<c r="C1"><v>3</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "E1", "value": 5}])
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["A1", "C1", "E1"]


def test_multiple_out_of_order_cells_in_same_row_end_up_sorted(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "J1": '<c r="J1"><v>10</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "F1", "value": "f"},
            {"sheet": "Sheet1", "cell": "C1", "value": "c"},
            {"sheet": "Sheet1", "cell": "H1", "value": "h"},
        ],
    )
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["A1", "C1", "F1", "H1", "J1"]


def test_ten_letter_wide_column_ref_sorts_correctly(tmp_path):
    """AA/AB etc. must sort by numeric column value, not lexically after
    single-letter columns already covers this (A < Z < AA), but exercising a
    double-letter column landing between two others catches an off-by-one in
    _col_to_num."""
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"Z1": '<c r="Z1"><v>26</v></c>', "AB1": '<c r="AB1"><v>28</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "AA1", "value": 27}])
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["Z1", "AA1", "AB1"]


def test_updating_existing_cell_does_not_change_row_or_cell_order(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={
            "A1": '<c r="A1"><v>1</v></c>',
            "B1": '<c r="B1"><v>2</v></c>',
            "C1": '<c r="C1"><v>3</v></c>',
        },
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B1", "value": 99}])
    xml = sheet1_xml(dst)
    assert cell_refs_in_row(xml, 1) == ["A1", "B1", "C1"]
    assert '<c r="B1"><v>99</v></c>' in xml


# ---------------------------------------------------------------------------
# <dimension> updates
# ---------------------------------------------------------------------------


def test_dimension_expands_to_cover_a_far_away_new_cell(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"A1": '<c r="A1"><v>1</v></c>'})
    assert _dimension_ref(sheet1_xml(src)) == "A1"

    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "Z50", "value": 1}])
    assert _dimension_ref(sheet1_xml(dst)) == "A1:Z50"


def test_dimension_expands_to_cover_every_touched_ref_in_one_edit(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"A1": '<c r="A1"><v>1</v></c>'})
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "B5", "value": 1},
            {"sheet": "Sheet1", "cell": "D2", "value": 2},
        ],
    )
    assert _dimension_ref(sheet1_xml(dst)) == "A1:D5"


def test_dimension_unchanged_text_when_edit_stays_within_existing_bounds(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>', "C3": '<c r="C3"><v>3</v></c>'},
    )
    assert _dimension_ref(sheet1_xml(src)) == "A1:C3"

    dst = tmp_path / "out.xlsx"
    # B2 is a brand-new cell, but already inside the A1:C3 bounding box.
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 2}])
    assert _dimension_ref(sheet1_xml(dst)) == "A1:C3"


def test_dimension_single_cell_ref_format_when_touched_ref_equals_existing(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"A1": '<c r="A1"><v>1</v></c>'})
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 99}])
    # min == max on both axes -> single-ref form, not "A1:A1".
    assert _dimension_ref(sheet1_xml(dst)) == "A1"


def test_no_dimension_element_stays_absent(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx",
        cells={"A1": '<c r="A1"><v>1</v></c>'},
        include_dimension=False,
    )
    assert "<dimension" not in sheet1_xml(src)

    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "Z99", "value": 1}])
    assert "<dimension" not in sheet1_xml(dst)
