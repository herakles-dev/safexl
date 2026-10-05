"""safexl_mcp.engine -- calc-chain handling and fullCalcOnLoad plumbing.

See engine._force_full_calc_on_load / engine._strip_calc_chain_override and
the "What always changes" section of engine.py's module docstring: any edit
unconditionally drops xl/calcChain.xml (if present) and its Content_Types
Override, and forces <calcPr fullCalcOnLoad="1"/> on xl/workbook.xml.
"""
from __future__ import annotations

import zipfile

from safexl_mcp import edit

from _util import build_minimal_workbook, fixture_path, first_sheet_name


# ---------------------------------------------------------------------------
# fullCalcOnLoad -- all four branches of _force_full_calc_on_load
# ---------------------------------------------------------------------------


def test_fullcalconload_inserted_when_no_calcpr_element_exists(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")  # no <calcPr> at all
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])
    with zipfile.ZipFile(dst) as z:
        wb = z.read("xl/workbook.xml").decode()
    assert "<calcPr fullCalcOnLoad=\"1\"/></workbook>" in wb


def test_fullcalconload_added_to_self_closing_empty_calcpr(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", extra_workbook_xml="<calcPr/>")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])
    with zipfile.ZipFile(dst) as z:
        wb = z.read("xl/workbook.xml").decode()
    assert '<calcPr fullCalcOnLoad="1"/>' in wb


def test_fullcalconload_added_to_calcpr_with_existing_attributes(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", extra_workbook_xml='<calcPr calcId="999"/>')
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])
    with zipfile.ZipFile(dst) as z:
        wb = z.read("xl/workbook.xml").decode()
    assert 'fullCalcOnLoad="1"' in wb
    assert 'calcId="999"' in wb  # existing attribute preserved, not dropped


def test_workbook_xml_byte_identical_when_fullcalconload_already_set(tmp_path):
    """If the source already declares fullCalcOnLoad="1", the part must come
    out byte-identical -- _force_full_calc_on_load returns the text
    unchanged, so nothing is reformatted or rewritten."""
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src, extra_workbook_xml='<calcPr fullCalcOnLoad="1"/>')
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    with zipfile.ZipFile(src) as zs, zipfile.ZipFile(dst) as zd:
        assert zs.read("xl/workbook.xml") == zd.read("xl/workbook.xml")


# ---------------------------------------------------------------------------
# calcChain.xml -- always dropped if present, on a controlled synthetic file
# ---------------------------------------------------------------------------


def test_calc_chain_and_its_content_types_override_dropped_synthetic(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(src, include_calc_chain=True)
    with zipfile.ZipFile(src) as z:
        assert "xl/calcChain.xml" in z.namelist()
        assert "calcChain" in z.read("[Content_Types].xml").decode()

    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])
    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        assert "xl/calcChain.xml" not in names
        assert "calcChain" not in z.read("[Content_Types].xml").decode()


def test_calc_chain_dropped_even_for_an_edit_unrelated_to_any_formula(tmp_path):
    """Editing a plain value (not a formula cell) still drops calcChain.xml:
    an untouched formula elsewhere could depend on the edited value, so any
    pre-existing calc chain is conservatively treated as stale by any edit."""
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(
        src,
        cells={"A1": '<c r="A1"><v>1</v></c>', "B1": '<c r="B1"><v>2</v></c>'},
        include_calc_chain=True,
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 5}])  # plain value edit
    with zipfile.ZipFile(dst) as z:
        assert "xl/calcChain.xml" not in z.namelist()


# ---------------------------------------------------------------------------
# calcChain.xml -- real fixture that ships with one already
# (NewStyleConditionalFormattings.xlsx)
# ---------------------------------------------------------------------------


def test_calc_chain_dropped_on_real_fixture_that_ships_with_one(tmp_path):
    src = fixture_path("NewStyleConditionalFormattings.xlsx")
    with zipfile.ZipFile(src) as z:
        assert "xl/calcChain.xml" in z.namelist()
        assert "calcChain" in z.read("[Content_Types].xml").decode()
        assert 'fullCalcOnLoad="1"' not in z.read("xl/workbook.xml").decode()

    dst = tmp_path / "out.xlsx"
    sheet = first_sheet_name(src)
    edit(src, dst, [{"sheet": sheet, "cell": "ZZ9000", "value": 1}])

    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        assert "xl/calcChain.xml" not in names
        assert "calcChain" not in z.read("[Content_Types].xml").decode()
        assert 'fullCalcOnLoad="1"' in z.read("xl/workbook.xml").decode()


def test_no_calc_chain_present_leaves_content_types_unaffected(tmp_path):
    """A fixture with no calcChain.xml to begin with must not gain a dangling
    Override or any other calcChain-shaped artifact after an edit."""
    src = fixture_path("WithChart.xlsx")
    with zipfile.ZipFile(src) as z:
        assert "xl/calcChain.xml" not in z.namelist()

    dst = tmp_path / "out.xlsx"
    sheet = first_sheet_name(src)
    edit(src, dst, [{"sheet": sheet, "cell": "ZZ9000", "value": 1}])
    with zipfile.ZipFile(dst) as z:
        assert "xl/calcChain.xml" not in z.namelist()
        assert "calcChain" not in z.read("[Content_Types].xml").decode()
