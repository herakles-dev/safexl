"""safexl_mcp.engine -- formula writing: leading '=' accepted and stripped,
never stored inside <f>, XML-escaped. See engine._normalize_formula /
engine._make_cell_xml.
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError, edit

from _util import build_minimal_workbook, sheet1_xml


def test_formula_without_leading_equals_stored_bare(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "A1*2"}])
    xml = sheet1_xml(dst)
    assert "<c r=\"B2\"><f>A1*2</f></c>" in xml


def test_formula_with_leading_equals_is_stripped(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "=A1*2"}])
    xml = sheet1_xml(dst)
    assert "<c r=\"B2\"><f>A1*2</f></c>" in xml
    f_body = xml.split("<f>", 1)[1].split("</f>", 1)[0]
    assert "=" not in f_body


def test_formula_leading_equals_never_stored_across_the_whole_output(tmp_path):
    """The <f> element must NEVER contain a leading '=' -- storing it is the
    documented negokaz/excel-mcp-server-style formula-corruption bug this
    engine exists to avoid (see engine.py module docstring)."""
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "  =SUM(A1:A2)  "}])
    xml = sheet1_xml(dst)
    assert "<f>SUM(A1:A2)</f>" in xml
    assert "<f>=SUM(A1:A2)</f>" not in xml
    assert "<f> =SUM" not in xml


def test_formula_only_a_single_leading_equals_is_stripped(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "==A1"}])
    xml = sheet1_xml(dst)
    assert "<f>=A1</f>" in xml


def test_formula_special_characters_are_xml_escaped(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": '=IF(A1<2,"x","y")'}])
    xml = sheet1_xml(dst)
    assert "&lt;" in xml
    assert "A1<2" not in xml


def test_formula_cell_has_no_style_attribute_when_new(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "A1"}])
    xml = sheet1_xml(dst)
    assert '<c r="B2"><f>A1</f></c>' in xml  # no s="..." on a brand-new cell


def test_formula_overwrite_preserves_existing_style_attribute(tmp_path):
    src = build_minimal_workbook(
        tmp_path / "src.xlsx", cells={"B2": '<c r="B2" s="3"><v>1</v></c>'}
    )
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "A1"}])
    xml = sheet1_xml(dst)
    assert '<c r="B2" s="3"><f>A1</f></c>' in xml


def test_empty_formula_string_rejected(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "   "}])
    assert not dst.exists()


def test_formula_non_string_rejected(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": 123}])
    assert not dst.exists()


def test_fullcalconload_set_when_a_formula_is_written(tmp_path):
    import zipfile

    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "formula": "A1*2"}])
    with zipfile.ZipFile(dst) as z:
        wb = z.read("xl/workbook.xml").decode()
    assert 'fullCalcOnLoad="1"' in wb


def test_overwriting_a_formula_cell_with_a_value_also_sets_fullcalconload(tmp_path):
    """Overwriting a formula cell (even with a plain value) makes any
    pre-existing calcChain potentially stale -- engine._inject_cells flags
    formula_touched=True in this case too, though the engine currently drops
    calcChain.xml unconditionally on every edit regardless."""
    import zipfile

    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"B2": '<c r="B2"><f>A1</f></c>'})
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 5}])
    xml = sheet1_xml(dst)
    assert "<f>" not in xml  # the formula is gone, replaced by a plain value
    assert '<c r="B2"><v>5</v></c>' in xml
    with zipfile.ZipFile(dst) as z:
        wb = z.read("xl/workbook.xml").decode()
    assert 'fullCalcOnLoad="1"' in wb
