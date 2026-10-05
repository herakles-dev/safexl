"""safexl_mcp.structural -- rename_sheet: updates xl/workbook.xml's <sheet
name=...>, and every cross-sheet formula reference to the renamed sheet
(both quoted and unquoted forms), without touching the sheet's own zip path.
"""
from __future__ import annotations

import zipfile

import pytest

from safexl_mcp import SafeXLError
from safexl_mcp.structural import rename_sheet

from _util import build_two_sheet_workbook, read_zip_text


def test_rename_sheet_updates_workbook_xml_name(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx", sheet1_name="Sheet1", sheet2_name="Sheet2")
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Sheet1", "Renamed")
    wb = read_zip_text(dst, "xl/workbook.xml")
    assert 'name="Renamed"' in wb
    assert 'name="Sheet1"' not in wb
    assert 'name="Sheet2"' in wb  # untouched


def test_rename_sheet_keeps_zip_path_unchanged(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx", sheet1_name="Sheet1", sheet2_name="Sheet2")
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Sheet1", "Renamed")
    with zipfile.ZipFile(dst) as z:
        names = set(z.namelist())
    assert "xl/worksheets/sheet1.xml" in names
    assert "xl/worksheets/sheet2.xml" in names


def test_rename_sheet_updates_unquoted_cross_sheet_formula_reference(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet2_cells={"B2": '<c r="B2"><f>Sheet1!A1+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Sheet1", "Renamed")
    sheet2_xml = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<f>Renamed!A1+1</f>" in sheet2_xml


def test_rename_sheet_updates_quoted_cross_sheet_formula_reference(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Old Sheet",
        sheet2_name="Sheet2",
        sheet2_cells={"B2": "<c r=\"B2\"><f>'Old Sheet'!A1+1</f></c>"},
    )
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Old Sheet", "Renamed")
    sheet2_xml = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<f>Renamed!A1+1</f>" in sheet2_xml


def test_rename_sheet_to_name_with_space_gets_quoted_in_formula(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet2_cells={"B2": '<c r="B2"><f>Sheet1!A1+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Sheet1", "New Name")
    sheet2_xml = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<f>'New Name'!A1+1</f>" in sheet2_xml


def test_rename_sheet_does_not_touch_reference_to_other_sheet(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet1_cells={"B2": '<c r="B2"><f>Sheet2!A1+1</f></c>'},
    )
    dst = tmp_path / "out.xlsx"
    rename_sheet(src, dst, "Sheet1", "Renamed")
    sheet1_xml = read_zip_text(dst, "xl/worksheets/sheet1.xml")
    assert "<f>Sheet2!A1+1</f>" in sheet1_xml


def test_rename_sheet_not_found_raises_and_leaves_dst_untouched(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        rename_sheet(src, dst, "NoSuchSheet", "Renamed")
    assert not dst.exists()


def test_rename_sheet_duplicate_name_raises(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx", sheet1_name="Sheet1", sheet2_name="Sheet2")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        rename_sheet(src, dst, "Sheet1", "Sheet2")
    assert not dst.exists()


def test_rename_sheet_invalid_name_raises(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx", sheet1_name="Sheet1", sheet2_name="Sheet2")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        rename_sheet(src, dst, "Sheet1", "Bad/Name")
    assert not dst.exists()
