"""Comments, CDATA and processing instructions that look like cells/rows must never be edited
(S3 of the red-team): the edit goes to the REAL cell and the decoy bytes are preserved."""
from __future__ import annotations

import zipfile
from xml.etree import ElementTree as ET

import pytest

from safexl_mcp import edit, insert_rows
from _pkg import NS_MAIN, make_package, sheet_xml

DECOY_COMMENT = '<!-- <row r="2"><c r="C2"><v>1</v></c></row> -->'
DECOY_CDATA = '<![CDATA[<c r="C2"><v>1</v></c></row></sheetData>]]>'
DECOY_PI = '<?decoy <c r="C2"><v>1</v></c> ?>'


def _sheet_bytes(path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _real_cells(path) -> dict[str, str | None]:
    root = ET.fromstring(_sheet_bytes(path))
    out = {}
    for c in root.iter(f"{{{NS_MAIN}}}c"):
        v = c.find(f"{{{NS_MAIN}}}v")
        t = c.find(f"{{{NS_MAIN}}}is/{{{NS_MAIN}}}t")
        out[c.get("r")] = v.text if v is not None else (t.text if t is not None else None)
    return out


def _pkg(tmp_path, body):
    return make_package(tmp_path / "in.xlsx", [("Sheet1", sheet_xml(body))])


def test_comment_decoy_before_data_edits_the_real_cell(tmp_path):
    src = _pkg(tmp_path, DECOY_COMMENT + '<row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="C2"><v>500</v></c></row>')
    out = tmp_path / "out.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "C2", "value": 9999}])
    assert _real_cells(out)["C2"] == "9999"
    assert DECOY_COMMENT in _sheet_bytes(out)


def test_cdata_decoy_does_not_swallow_the_edit(tmp_path):
    body = (
        '<row r="1"><c r="A1" t="inlineStr"><is><t>' + DECOY_CDATA + "</t></is></c></row>"
        '<row r="2"><c r="A2"><v>3</v></c></row>'
    )
    src = _pkg(tmp_path, body)
    out = tmp_path / "out.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "C2", "value": 9999}])
    cells = _real_cells(out)
    assert cells["C2"] == "9999"
    assert DECOY_CDATA in _sheet_bytes(out)
    assert _sheet_bytes(out).count('r="C2"') == 2  # one inside the CDATA, one real


def test_processing_instruction_decoy_edits_the_real_cell(tmp_path):
    src = _pkg(tmp_path, DECOY_PI + '<row r="2"><c r="C2"><v>500</v></c></row>')
    out = tmp_path / "out.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "C2", "value": 9999}])
    assert _real_cells(out)["C2"] == "9999"
    assert DECOY_PI in _sheet_bytes(out)


def test_comment_decoy_row_does_not_capture_a_new_cell_in_row_order(tmp_path):
    src = _pkg(tmp_path, '<row r="1"><c r="A1"><v>1</v></c></row><!-- <row r="3"></row> --><row r="5"><c r="A5"><v>5</v></c></row>')
    out = tmp_path / "out.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "B4", "value": 4}])
    xml = _sheet_bytes(out)
    assert xml.index('<row r="4"') < xml.index('<row r="5"')
    assert xml.index("<!-- <row") < xml.index('<row r="4"')
    assert _real_cells(out)["B4"] == "4"


def test_insert_rows_ignores_a_decoy_row_inside_a_comment(tmp_path):
    src = _pkg(tmp_path, '<!-- <row r="2"><c r="A2"><v>9</v></c></row> --><row r="2"><c r="A2"><v>2</v></c></row>')
    out = tmp_path / "out.xlsx"
    insert_rows(src, out, "Sheet1", 2, 1)
    cells = _real_cells(out)
    assert cells == {"A3": "2"}
    assert '<!-- <row r="2"><c r="A2"><v>9</v></c></row> -->' in _sheet_bytes(out)


# --- read-back confirmation ----------------------------------------------------


def test_readback_catches_a_broken_write_and_leaves_no_output(tmp_path, monkeypatch):
    from safexl_mcp import SafeXLError, engine

    src = _pkg(tmp_path, '<row r="1"><c r="A1"><v>1</v></c></row>')
    out = tmp_path / "out.xlsx"
    monkeypatch.setattr(engine, "_make_cell_xml", lambda ref, change, style="": f'<c r="{ref}"><v>0</v></c>')
    with pytest.raises(SafeXLError, match=r"read-back failed at Sheet1!A1"):
        edit(src, out, [{"sheet": "Sheet1", "cell": "A1", "value": 5}])
    assert not out.exists()
    assert not list(tmp_path.glob("tmp*"))


def test_readback_catches_a_write_that_lands_outside_its_row(tmp_path, monkeypatch):
    from safexl_mcp import SafeXLError, engine

    src = _pkg(tmp_path, '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>')
    out = tmp_path / "out.xlsx"
    monkeypatch.setattr(engine, "_inject_cells", lambda xml, u: (xml.replace("</sheetData>", '<row r="1"><c r="B3"><v>8</v></c></row></sheetData>'), False))
    with pytest.raises(SafeXLError, match="read-back failed"):
        edit(src, out, [{"sheet": "Sheet1", "cell": "B3", "value": 8}])
    assert not out.exists()


def test_structural_readback_catches_a_wrong_shift(tmp_path, monkeypatch):
    from safexl_mcp import SafeXLError, structural

    src = _pkg(tmp_path, '<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>')
    out = tmp_path / "out.xlsx"
    monkeypatch.setattr(structural, "_shift_rows_down", lambda text, row, count: text)  # writer forgets to move the rows
    with pytest.raises(SafeXLError, match="read-back failed"):
        insert_rows(src, out, "Sheet1", 2, 1)
    assert not out.exists()
