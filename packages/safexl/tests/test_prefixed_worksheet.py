"""A worksheet written with a namespace prefix on every element (the .NET OpenXML SDK) is refused
with a typed SafeXLError, never a bare ValueError and never a silent mis-edit (I07)."""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError, edit, insert_rows
from safexl_mcp.verify import verify
from _pkg import NS_MAIN, make_package

PREFIXED = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<x:worksheet xmlns:x="{NS_MAIN}"><x:sheetData>'
    '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>'
)


def test_edit_refuses_a_prefixed_worksheet_with_a_typed_error(tmp_path):
    src = make_package(tmp_path / "in.xlsx", [("Sheet1", PREFIXED)])
    out = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="namespace prefix"):
        edit(src, out, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])
    assert not out.exists()


def test_insert_rows_refuses_a_prefixed_worksheet(tmp_path):
    src = make_package(tmp_path / "in.xlsx", [("Sheet1", PREFIXED)])
    out = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="not supported yet"):
        insert_rows(src, out, "Sheet1", 1, 1)
    assert not out.exists()


def test_verify_still_reads_a_prefixed_worksheet(tmp_path):
    src = make_package(tmp_path / "in.xlsx", [("Sheet1", PREFIXED)])
    r = verify(src, src)
    assert r["verdict"] == "intact" and r["cell_changes"] == {}
