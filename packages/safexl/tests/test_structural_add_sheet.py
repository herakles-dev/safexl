"""safexl_mcp.structural -- add_sheet: creates a new empty worksheet part and
registers it in xl/workbook.xml, xl/_rels/workbook.xml.rels, and
[Content_Types].xml.
"""
from __future__ import annotations

import re
import zipfile

import pytest

from safexl_mcp import SafeXLError, edit
from safexl_mcp.structural import add_sheet

from _util import build_minimal_workbook, build_two_sheet_workbook, read_zip_text


def test_add_sheet_creates_new_worksheet_part(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "NewSheet")
    with zipfile.ZipFile(dst) as z:
        assert "xl/worksheets/sheet2.xml" in z.namelist()
    sheet_xml = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert "<sheetData/>" in sheet_xml


def test_add_sheet_registers_in_workbook_xml(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "NewSheet")
    wb = read_zip_text(dst, "xl/workbook.xml")
    m = re.search(r'<sheet\b[^>]*\bname="NewSheet"[^>]*/>', wb)
    assert m is not None
    assert 'sheetId="2"' in m.group(0)


def test_add_sheet_registers_relationship(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "NewSheet")
    wb = read_zip_text(dst, "xl/workbook.xml")
    rels = read_zip_text(dst, "xl/_rels/workbook.xml.rels")
    m = re.search(r'<sheet\b[^>]*\bname="NewSheet"[^>]*/>', wb)
    rid = re.search(r'r:id="([^"]+)"', m.group(0)).group(1)
    assert f'Id="{rid}"' in rels
    assert 'Target="worksheets/sheet2.xml"' in rels


def test_add_sheet_registers_content_type_override(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "NewSheet")
    ct = read_zip_text(dst, "[Content_Types].xml")
    assert '<Override PartName="/xl/worksheets/sheet2.xml"' in ct


def test_add_sheet_picks_next_available_sheet_number(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "src.xlsx", sheet1_name="Sheet1", sheet2_name="Sheet2")
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "Sheet3")
    with zipfile.ZipFile(dst) as z:
        assert "xl/worksheets/sheet3.xml" in z.namelist()
    wb = read_zip_text(dst, "xl/workbook.xml")
    m = re.search(r'<sheet\b[^>]*\bname="Sheet3"[^>]*/>', wb)
    assert m is not None
    assert 'sheetId="3"' in m.group(0)


def test_add_sheet_existing_sheets_untouched(tmp_path):
    src = build_two_sheet_workbook(
        tmp_path / "src.xlsx",
        sheet1_name="Sheet1",
        sheet2_name="Sheet2",
        sheet1_cells={"A1": '<c r="A1"><v>1</v></c>'},
    )
    dst = tmp_path / "out.xlsx"
    add_sheet(src, dst, "Sheet3")
    sheet1_xml = read_zip_text(dst, "xl/worksheets/sheet1.xml")
    assert '<c r="A1"><v>1</v></c>' in sheet1_xml


def test_add_sheet_duplicate_name_raises(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", sheet_name="Sheet1")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        add_sheet(src, dst, "Sheet1")
    assert not dst.exists()


def test_add_sheet_duplicate_name_case_insensitive_raises(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx", sheet_name="Sheet1")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        add_sheet(src, dst, "sheet1")
    assert not dst.exists()


def test_add_sheet_invalid_name_raises(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError):
        add_sheet(src, dst, "Bad*Name")
    assert not dst.exists()


def test_add_sheet_new_sheet_is_editable_with_edit(tmp_path):
    """Integration check: add_sheet's output is a real, valid workbook that
    the already-tested engine.edit() can write into by the new sheet's name."""
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    mid = tmp_path / "mid.xlsx"
    add_sheet(src, mid, "NewSheet")
    dst = tmp_path / "out.xlsx"
    edit(mid, dst, [{"sheet": "NewSheet", "cell": "A1", "value": 42}])
    sheet_xml = read_zip_text(dst, "xl/worksheets/sheet2.xml")
    assert '<c r="A1"><v>42</v></c>' in sheet_xml
