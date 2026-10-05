"""safexl_mcp.server v0.2 tools: verify, view, salvage, resize_table, and edit's extend_tables.

Same approach as test_server.py: call the handler functions directly (they return
``CallToolResult``) in both transports."""
from __future__ import annotations

import asyncio
import base64
import io
import tempfile
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("mcp.server.mcpserver", reason="safexl_mcp.server needs mcp>=2.2.0")

from safexl_mcp import edit
from safexl_mcp import server as S
from safexl_mcp.server import (
    CellChange,
    build_server,
    http_edit,
    http_insert_cells,
    http_resize_table,
    http_salvage,
    http_verify,
    http_view,
    stdio_edit,
    stdio_insert_cells,
    stdio_resize_table,
    stdio_salvage,
    stdio_verify,
    stdio_view,
)

from _pkg import cell, make_package, row, sheet_xml
from _util import build_minimal_workbook

pytestmark = pytest.mark.filterwarnings("ignore::DeprecationWarning")

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
EXPECTED_TOOLS = [
    "safexl_add_sheet",
    "safexl_diff",
    "safexl_edit",
    "safexl_insert_cells",
    "safexl_insert_rows",
    "safexl_inspect",
    "safexl_rename_sheet",
    "safexl_resize_table",
    "safexl_salvage",
    "safexl_verify",
    "safexl_view",
]


def _b64(path: Path) -> str:
    return base64.b64encode(Path(path).read_bytes()).decode("ascii")


def _istr(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def _table_pkg(tmp_path: Path) -> Path:
    body = (
        row(1, _istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma"))
        + row(2, '<c r="A2"><v>1</v></c>', '<c r="B2"><v>2</v></c>', '<c r="C2"><v>3</v></c>')
        + row(3, '<c r="A3"><v>4</v></c>', '<c r="B3"><v>5</v></c>', '<c r="C3"><v>6</v></c>')
    )
    table = (
        f'<table xmlns="{NS}" id="1" name="Sales" displayName="Sales" ref="A1:C3">'
        '<autoFilter ref="A1:C3"/><tableColumns count="3">'
        '<tableColumn id="1" name="Alpha"/><tableColumn id="2" name="Beta"/><tableColumn id="3" name="Gamma"/>'
        "</tableColumns></table>"
    )
    return make_package(
        tmp_path / "tbl.xlsx",
        [("Data", sheet_xml(body))],
        extra_parts={"xl/tables/table1.xml": table},
        sheet_rels={1: [("table", "../tables/table1.xml")]},
    )


def _zip_names(raw: bytes) -> set[str]:
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        return set(z.namelist())


def _table_xml(raw: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        return z.read("xl/tables/table1.xml").decode()


def _no_leftover_temp(before: set[Path]) -> None:
    after = set(Path(tempfile.gettempdir()).glob(f"{S._TMP_PREFIX}*"))
    assert after <= before


def _tmp_snapshot() -> set[Path]:
    return set(Path(tempfile.gettempdir()).glob(f"{S._TMP_PREFIX}*"))


def _err_code(result) -> str:
    assert result.is_error is True
    return result.structured_content["error"]["code"]


@pytest.mark.parametrize("mode", ["stdio", "http"])
def test_tool_list_matches_documented_set(mode):
    tools = asyncio.run(build_server(mode).list_tools())
    assert sorted(t.name for t in tools) == EXPECTED_TOOLS


@pytest.mark.parametrize("mode", ["stdio", "http"])
def test_new_tool_descriptions_lead_with_when_to_use(mode):
    tools = {t.name: t.description for t in asyncio.run(build_server(mode).list_tools())}
    assert "check that an edited workbook lost nothing" in tools["safexl_verify"].split(". ")[0]
    first = tools["safexl_salvage"].split(": ")[0]
    assert "re-saved by openpyxl, pandas or another library and lost features" in first
    assert tools["safexl_view"].startswith("Use this")
    assert tools["safexl_resize_table"].startswith("Use this")


def test_version_is_0_2_0():
    import safexl_mcp

    assert safexl_mcp.__version__ == "0.2.1"


# ---------------------------------------------------------------- verify


def test_stdio_verify_intact(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    out = tmp_path / "b.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "A1", "value": 5}])
    body = stdio_verify(str(src), str(out)).structured_content
    assert body["ok"] is True
    assert body["verdict"] == "intact"


def test_stdio_verify_refusals(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    assert _err_code(stdio_verify(str(src), str(tmp_path / "nope.xlsx"))) == "file_not_found"
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"not a zip at all")
    assert _err_code(stdio_verify(str(src), str(bad))) == "not_a_workbook"


def test_http_verify_intact_and_no_paths(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    out = tmp_path / "b.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "A1", "value": 5}])
    snap = _tmp_snapshot()
    result = http_verify(_b64(src), _b64(out))
    assert result.structured_content["verdict"] == "intact"
    assert "safexl-mcp-" not in result.content[0].text
    _no_leftover_temp(snap)


def test_http_verify_refusals(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    good = _b64(src)
    assert _err_code(http_verify("/etc/passwd", good)) == "path_not_allowed"
    assert _err_code(http_verify(good, "https://example.com/x.xlsx")) == "path_not_allowed"
    assert _err_code(http_verify(good, base64.b64encode(b"plain text").decode())) == "not_a_workbook"
    assert _err_code(http_verify(good, "A" * (S.MAX_UPLOAD_B64_CHARS + 4))) == "file_too_large"


# ---------------------------------------------------------------- view


def test_stdio_view_happy_and_unknown_sheet(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    body = stdio_view(str(src)).structured_content
    assert body["ok"] is True and body["sheet"] == "Sheet1" and "cells" in body
    result = stdio_view(str(src), sheet="Nope")
    assert _err_code(result) == "sheet_not_found"
    assert "Sheet1" in result.structured_content["error"]["details"]["available_sheets"]
    assert _err_code(stdio_view(str(src), cell_range="not-a-range")) == "invalid_range"
    assert _err_code(stdio_view(str(tmp_path / "x.xlsx"))) == "file_not_found"


def test_stdio_view_find(tmp_path):
    src = _table_pkg(tmp_path)
    body = stdio_view(str(src), find="beta").structured_content
    assert any(c["ref"] == "B1" for c in body["cells"])


def test_http_view_happy_and_refusals(tmp_path):
    src = _table_pkg(tmp_path)
    snap = _tmp_snapshot()
    body = http_view(_b64(src), sheet="Data", max_cells=2).structured_content
    assert body["ok"] is True and len(body["cells"]) == 2 and body["truncated"] is True
    assert _err_code(http_view(_b64(src), sheet="Nope")) == "sheet_not_found"
    assert _err_code(http_view("/home/x/a.xlsx")) == "path_not_allowed"
    assert _err_code(http_view(base64.b64encode(b"hello").decode())) == "not_a_workbook"
    assert _err_code(http_view("A" * (S.MAX_UPLOAD_B64_CHARS + 4))) == "file_too_large"
    _no_leftover_temp(snap)


def test_view_max_cells_has_hard_ceiling():
    tool = {t.name: t for t in asyncio.run(build_server("http").list_tools())}["safexl_view"]
    prop = tool.input_schema["properties"]["max_cells"]
    assert prop["maximum"] == S.MAX_VIEW_CELLS


# ---------------------------------------------------------------- resize_table


def test_stdio_resize_table_happy_and_refusals(tmp_path):
    src = _table_pkg(tmp_path)
    out = tmp_path / "out.xlsx"
    body = stdio_resize_table(str(src), "Sales", "A1:C4", output_path=str(out)).structured_content
    assert body["ok"] is True and body["output_path"] == str(out)
    assert 'ref="A1:C4"' in _table_xml(out.read_bytes())
    # output exists, no overwrite
    assert _err_code(stdio_resize_table(str(src), "Sales", "A1:C4", output_path=str(out))) == "output_exists"
    result = stdio_resize_table(str(src), "Nope", "A1:C4", output_path=str(tmp_path / "o2.xlsx"))
    assert _err_code(result) == "table_not_found"
    assert _err_code(stdio_resize_table(str(tmp_path / "x.xlsx"), "Sales", "A1:C4")) == "file_not_found"


def test_http_resize_table_happy_and_refusals(tmp_path):
    src = _table_pkg(tmp_path)
    snap = _tmp_snapshot()
    result = http_resize_table(_b64(src), "Sales", "A1:C4")
    body = result.structured_content
    assert body["ok"] is True
    assert 'ref="A1:C4"' in _table_xml(base64.b64decode(body["file_base64"]))
    assert "file_base64" in body and "base64 chars" in result.content[0].text
    assert _err_code(http_resize_table(_b64(src), "Nope", "A1:C4")) == "table_not_found"
    assert _err_code(http_resize_table("~/a.xlsx", "Sales", "A1:C4")) == "path_not_allowed"
    assert _err_code(http_resize_table(base64.b64encode(b"x" * 20).decode(), "Sales", "A1:C4")) == "not_a_workbook"
    assert _err_code(http_resize_table("A" * (S.MAX_UPLOAD_B64_CHARS + 4), "Sales", "A1:C4")) == "file_too_large"
    _no_leftover_temp(snap)


# ---------------------------------------------------------------- salvage


def _chart_pair(tmp_path, *, lose_chart=True):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.chart import BarChart, Reference

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["name", "qty"])
    for i in range(2, 7):
        ws.append([f"item{i}", i * 10])
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
    ws.add_chart(chart, "F2")
    src = tmp_path / "orig.xlsx"
    wb.save(src)
    # Excel declares xmlns:r on the root; openpyxl repeats it on each <sheet>
    ns = ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    tmp = tmp_path / "hoist.tmp"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.decode().replace(ns, "").replace("<workbook ", "<workbook" + ns + " ", 1).encode()
            zo.writestr(item, data)
    tmp.replace(src)
    wb2 = openpyxl.load_workbook(src)
    wb2["Data"]["B3"] = 5
    if lose_chart:
        wb2["Data"]._charts = []
    dmg = tmp_path / "dmg.xlsx"
    wb2.save(dmg)
    return src, dmg


def test_stdio_salvage_lossy_resave_is_salvaged_and_verifies_intact(tmp_path):
    src, dmg = _chart_pair(tmp_path)
    body = stdio_salvage(str(src), str(dmg)).structured_content
    assert body["ok"] is True and body["status"] == "salvaged" and body["output_written"] is True
    assert Path(body["output_path"]).is_file()
    assert "xl/charts/chart1.xml" in _zip_names(Path(body["output_path"]).read_bytes())
    assert stdio_verify(str(src), body["output_path"]).structured_content["verdict"] == "intact"


def test_stdio_salvage_nothing_to_do_is_unambiguous_and_force_salvages(tmp_path):
    src, dmg = _chart_pair(tmp_path, lose_chart=False)
    body = stdio_salvage(str(src), str(dmg)).structured_content
    assert body["status"] == "nothing_to_do" and body["output_written"] is False
    assert "output_path" not in body and "file_base64" not in body
    assert "intact" in body["message"] and "keep" in body["message"]
    assert not (tmp_path / "orig.salvaged.xlsx").exists()
    forced = stdio_salvage(str(src), str(dmg), force=True).structured_content
    assert forced["status"] == "salvaged" and forced["output_written"] is True
    assert Path(forced["output_path"]).is_file()
    assert "message" not in forced


def test_stdio_salvage_refusals(tmp_path):
    src, dmg = _chart_pair(tmp_path)
    assert _err_code(stdio_salvage(str(src), str(tmp_path / "no.xlsx"))) == "file_not_found"
    clash = tmp_path / "clash.xlsx"
    clash.write_bytes(b"x")
    assert _err_code(stdio_salvage(str(src), str(dmg), output_path=str(clash))) == "output_exists"
    assert _err_code(stdio_salvage(str(src), str(dmg), output_path=str(tmp_path / "o.xlsm"))) == "extension_mismatch"


def test_http_salvage_statuses_and_refusals(tmp_path):
    src, dmg = _chart_pair(tmp_path)
    snap = _tmp_snapshot()
    result = http_salvage(_b64(src), _b64(dmg))
    body = result.structured_content
    assert body["status"] == "salvaged" and body["output_written"] is True
    assert "xl/charts/chart1.xml" in _zip_names(base64.b64decode(body["file_base64"]))
    assert "safexl-mcp-" not in result.content[0].text

    src2, dmg2 = _chart_pair(tmp_path, lose_chart=False)
    body = http_salvage(_b64(src2), _b64(dmg2)).structured_content
    assert body["status"] == "nothing_to_do" and body["output_written"] is False
    assert "file_base64" not in body and "output_path" not in body and "message" in body
    body = http_salvage(_b64(src2), _b64(dmg2), force=True).structured_content
    assert body["status"] == "salvaged" and "file_base64" in body

    assert _err_code(http_salvage("/tmp/a.xlsx", _b64(dmg))) == "path_not_allowed"
    assert _err_code(http_salvage(_b64(src), "http://x/y.xlsx")) == "path_not_allowed"
    assert _err_code(http_salvage(_b64(src), base64.b64encode(b"junk junk").decode())) == "not_a_workbook"
    assert _err_code(http_salvage("A" * (S.MAX_UPLOAD_B64_CHARS + 4), _b64(dmg))) == "file_too_large"
    _no_leftover_temp(snap)


# ---------------------------------------------------------------- edit extend_tables


def _change(cell, value):
    return [CellChange(sheet="Data", cell=cell, value=value)]


def test_stdio_edit_extends_table_by_default(tmp_path):
    src = _table_pkg(tmp_path)
    out = tmp_path / "e.xlsx"
    body = stdio_edit(str(src), _change("D1", "Delta") + _change("D2", 1), output_path=str(out)).structured_content
    assert body["ok"] is True
    assert [t["table"] for t in body["tables_extended"]] == ["Sales"]
    assert body["tables_extended"][0]["new_ref"].startswith("A1:D")
    assert 'ref="A1:D3"' in _table_xml(out.read_bytes())
    assert body["cells_written"] == 2


def test_stdio_edit_extend_tables_false(tmp_path):
    src = _table_pkg(tmp_path)
    out = tmp_path / "e.xlsx"
    body = stdio_edit(
        str(src), _change("D1", "Delta"), output_path=str(out), extend_tables=False
    ).structured_content
    assert body["ok"] is True and body["tables_extended"] == []
    assert 'ref="A1:C3"' in _table_xml(out.read_bytes())


def test_http_edit_extends_and_honours_false(tmp_path):
    src = _table_pkg(tmp_path)
    snap = _tmp_snapshot()
    body = http_edit(_b64(src), _change("D1", "Delta")).structured_content
    assert [t["table"] for t in body["tables_extended"]] == ["Sales"]
    assert 'ref="A1:D3"' in _table_xml(base64.b64decode(body["file_base64"]))
    body = http_edit(_b64(src), _change("D1", "Delta"), extend_tables=False).structured_content
    assert body["tables_extended"] == []
    assert 'ref="A1:C3"' in _table_xml(base64.b64decode(body["file_base64"]))
    _no_leftover_temp(snap)


def test_edit_required_arguments_unchanged():
    for mode in ("stdio", "http"):
        tool = {t.name: t for t in asyncio.run(build_server(mode).list_tools())}["safexl_edit"]
        required = set(tool.input_schema.get("required", []))
        assert required == ({"path", "changes"} if mode == "stdio" else {"file_base64", "changes"})


# ---------------------------------------------------------------- insert_cells


def _cells_pkg(tmp_path: Path, *, torn: bool = False) -> Path:
    extra = (cell("G1", f="SUM(A1:C6)"),) if torn else ()
    body = "".join(
        row(n, cell(f"A{n}", v=str(n * 10)), cell(f"C{n}", v=str(n * 100)), *(extra if n == 1 else ()))
        for n in range(1, 7)
    )
    return make_package(tmp_path / "cells.xlsx", [("Data", sheet_xml(body))])


def _sheet_cells(raw: bytes) -> set[str]:
    import re

    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        return set(re.findall(r'<c r="([A-Z]+\d+)"', z.read("xl/worksheets/sheet1.xml").decode()))


def test_stdio_insert_cells_happy_and_refusals(tmp_path):
    src = _cells_pkg(tmp_path)
    out = tmp_path / "o.xlsx"
    body = stdio_insert_cells(str(src), "Data", "A3", output_path=str(out)).structured_content
    assert body["ok"] is True and body["output_path"] == str(out) and body["shift"] == "down"
    cells = _sheet_cells(out.read_bytes())
    assert "A7" in cells and "A3" not in cells and "C6" in cells and "C7" not in cells
    assert _err_code(stdio_insert_cells(str(src), "Data", "A3", output_path=str(out))) == "output_exists"
    result = stdio_insert_cells(str(src), "Nope", "A3", output_path=str(tmp_path / "o2.xlsx"))
    assert _err_code(result) == "sheet_not_found"
    assert _err_code(stdio_insert_cells(str(src), "Data", "zzz", output_path=str(tmp_path / "o3.xlsx"))) == "invalid_range"
    assert _err_code(stdio_insert_cells(str(tmp_path / "x.xlsx"), "Data", "A3")) == "file_not_found"


def test_stdio_insert_cells_torn_reference_is_typed(tmp_path):
    src = _cells_pkg(tmp_path, torn=True)
    out = tmp_path / "o.xlsx"
    result = stdio_insert_cells(str(src), "Data", "A3", output_path=str(out))
    assert _err_code(result) == "reference_would_tear"
    assert "A1:C6" in result.structured_content["error"]["message"]
    assert not out.exists()


def test_http_insert_cells_happy_and_refusals(tmp_path):
    src = _cells_pkg(tmp_path)
    snap = _tmp_snapshot()
    result = http_insert_cells(_b64(src), "Data", "A3")
    body = result.structured_content
    assert body["ok"] is True and "base64 chars" in result.content[0].text
    assert "A7" in _sheet_cells(base64.b64decode(body["file_base64"]))
    assert _err_code(http_insert_cells(_b64(src), "Nope", "A3")) == "sheet_not_found"
    assert _err_code(http_insert_cells(_b64(_cells_pkg(tmp_path, torn=True)), "Data", "A3")) == "reference_would_tear"
    assert _err_code(http_insert_cells("/etc/x.xlsx", "Data", "A3")) == "path_not_allowed"
    assert _err_code(http_insert_cells(base64.b64encode(b"x" * 20).decode(), "Data", "A3")) == "not_a_workbook"
    assert _err_code(http_insert_cells("A" * (S.MAX_UPLOAD_B64_CHARS + 4), "Data", "A3")) == "file_too_large"
    _no_leftover_temp(snap)


def test_insert_cells_description_leads_with_when_to_use():
    tools = {t.name: t.description for t in asyncio.run(build_server("http").list_tools())}
    assert tools["safexl_insert_cells"].startswith("Use this instead of safexl_insert_rows when the sheet has other")


# ---------------------------------------------------------------- receipt surface


_PREFIXED_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>'
    '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>'
)


@pytest.mark.parametrize("mode", ["stdio", "http"])
def test_verify_description_says_what_intact_covers(mode):
    tools = {t.name: t.description for t in asyncio.run(build_server(mode).list_tools())}
    d = tools["safexl_verify"]
    assert "added_risk" in d and "cell_changes" in d and "tracked features" in d and "summary" in d


def test_verify_receipt_has_summary_risk_and_cell_changes(tmp_path):
    src = build_minimal_workbook(tmp_path / "a.xlsx")
    out = tmp_path / "b.xlsx"
    edit(src, out, [{"sheet": "Sheet1", "cell": "A1", "value": 5}])
    for body in (stdio_verify(str(src), str(out)).structured_content,
                 http_verify(_b64(src), _b64(out)).structured_content):
        assert isinstance(body["summary"], str) and body["summary"]
        assert body["added_risk"] == []
        assert "cell_changes" in body and "cell_changes_truncated" not in body


def test_verify_receipt_caps_sheets_and_says_so():
    receipt = {"verdict": "intact", "cell_changes": {f"S{i}": {"cells_changed": 1} for i in range(80)}}
    out = S._bounded_receipt(receipt)
    assert len(out["cell_changes"]) == S.MAX_RECEIPT_SHEETS == 50
    assert out["cell_changes_truncated"] == {"shown": 50, "total": 80}
    assert len(receipt["cell_changes"]) == 80
    small = {"cell_changes": {"S": {}}}
    assert S._bounded_receipt(small) is small


def test_stdio_edit_and_structural_namespace_prefix_is_unsupported_workbook(tmp_path):
    src = make_package(tmp_path / "p.xlsx", [("Sheet1", _PREFIXED_SHEET)])
    out = tmp_path / "o.xlsx"
    r = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="B2", value=1)], output_path=str(out))
    assert _err_code(r) == "unsupported_workbook" and not out.exists()
    r = S.stdio_insert_rows(str(src), "Sheet1", 1, output_path=str(out))
    assert _err_code(r) == "unsupported_workbook" and not out.exists()


def test_stdio_edit_read_back_failure_is_write_not_confirmed(tmp_path, monkeypatch):
    from safexl_mcp.engine import SafeXLError

    src = build_minimal_workbook(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"

    def boom(*a, **k):
        raise SafeXLError("read-back failed at Sheet1!A1: value differs")

    monkeypatch.setattr(S, "edit", boom)
    r = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="A1", value=1)], output_path=str(out))
    assert _err_code(r) == "write_not_confirmed" and not out.exists()
    assert S._resize_error(SafeXLError("resize_table read-back failed: x"), src).code == "write_not_confirmed"
