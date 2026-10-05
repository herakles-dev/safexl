"""safexl_mcp.server -- the MCP server's tool handlers.

Covers the security-relevant surface the module docstring calls out: the
stdio transport's local-path tools, the HTTP transport's base64-only intake
(paths/URLs rejected, size cap enforced), structured ok/error results for
every tool, and that concurrent HTTP requests don't corrupt each other's
output. Calls the tool handler functions (``stdio_inspect``, ``http_edit``,
etc.) directly -- they are plain functions that return ``CallToolResult``,
the same objects the MCP SDK would get back through the wire protocol -- so
no transport/session plumbing is needed to exercise them.

Engine-level rollback behavior (shutil.move / raw OSError during the write)
is covered separately in test_engine_rollback_io.py; this file only tests
the server's own request handling (validation, intake, shaping, transport).

ENVIRONMENT NOTE
-----------------
``safexl_mcp/server.py`` does ``from mcp.server.mcpserver import MCPServer``,
which only exists in ``mcp>=2.0`` (``FastMCP`` was renamed to ``MCPServer``
in the 1.x -> 2.x jump; see pyproject.toml's dependency comment, which pins
``mcp>=2.2.0,<3``). If the environment running this file has an older
``mcp`` installed, the import fails and every test below is skipped with a
loud reason rather than erroring out -- see bugs_found in this task's
report for why that matters and what environment actually has the pin
satisfied.
"""
from __future__ import annotations

import asyncio
import base64
import io
import threading
import zipfile
from pathlib import Path

import pytest

pytest.importorskip(
    "mcp.server.mcpserver",
    reason=(
        "BUG (environment, not code): safexl_mcp/server.py requires "
        "mcp>=2.2.0 (mcp.server.mcpserver.MCPServer only exists in mcp 2.x; "
        "FastMCP was renamed). This interpreter's installed mcp is older "
        "and lacks that module, so safexl_mcp.server cannot be imported "
        "here at all -- these tests are skipped, not passing vacuously. "
        "pyproject.toml already pins mcp>=2.2.0,<3; install that version "
        "to run this file for real."
    ),
)

from safexl_mcp import server as S
from safexl_mcp.server import (
    CellChange,
    build_server,
    http_add_sheet,
    http_diff,
    http_edit,
    http_insert_rows,
    http_inspect,
    http_rename_sheet,
    stdio_add_sheet,
    stdio_diff,
    stdio_edit,
    stdio_insert_rows,
    stdio_inspect,
    stdio_rename_sheet,
)

from _util import build_minimal_workbook, build_two_sheet_workbook

pytestmark = pytest.mark.filterwarnings("ignore::DeprecationWarning")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _b64_bytes(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _b64_file(path: Path) -> str:
    return _b64_bytes(Path(path).read_bytes())


def _sheet1_xml_from_bytes(raw: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


def _tiny_zip_without_workbook_xml() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("hello.txt", "not a workbook")
    return buf.getvalue()


# ===========================================================================
# stdio_inspect
# ===========================================================================


def test_stdio_inspect_happy_path(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_inspect(str(src))
    body = result.structured_content
    assert result.is_error in (False, None)
    assert body["ok"] is True
    assert body["path"] == str(src)
    assert body["sheets"][0]["name"] == "Sheet1"
    assert body["defined_names_total"] == 0


def test_stdio_inspect_missing_file_returns_file_not_found(tmp_path):
    result = stdio_inspect(str(tmp_path / "does-not-exist.xlsx"))
    body = result.structured_content
    assert result.is_error is True
    assert body["ok"] is False
    assert body["error"]["code"] == "file_not_found"


def test_stdio_inspect_wrong_extension_returns_unsupported_file_type(tmp_path):
    bad = tmp_path / "book.docx"
    bad.write_bytes(b"whatever")
    result = stdio_inspect(str(bad))
    assert result.structured_content["error"]["code"] == "unsupported_file_type"


def test_stdio_inspect_non_zip_content_returns_not_a_workbook(tmp_path):
    bad = tmp_path / "book.xlsx"
    bad.write_bytes(b"this is not a zip file at all")
    result = stdio_inspect(str(bad))
    assert result.structured_content["error"]["code"] == "not_a_workbook"


def test_stdio_inspect_legacy_ole_magic_returns_unsupported_file_type(tmp_path):
    bad = tmp_path / "book.xlsx"
    bad.write_bytes(S._OLE_MAGIC + b"\x00" * 32)
    result = stdio_inspect(str(bad))
    assert result.structured_content["error"]["code"] == "unsupported_file_type"
    assert "encrypted" in result.structured_content["error"]["message"] or "legacy" in result.structured_content["error"]["message"]


# ===========================================================================
# stdio_edit
# ===========================================================================


def test_stdio_edit_happy_path_writes_output(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="B2", value=42)])
    body = result.structured_content
    assert body["ok"] is True
    assert body["cells_written"] == 1
    assert body["formula_cells"] == 0
    assert body["sheets_edited"] == ["Sheet1"]
    out = Path(body["output_path"])
    assert out.exists()
    assert out.name == "book.edited.xlsx"
    assert ">42<" in _sheet1_xml_from_bytes(out.read_bytes())


def test_stdio_edit_formula_change_reports_formula_cell_and_caveat_note(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="B2", formula="=A1*2")])
    body = result.structured_content
    assert body["ok"] is True
    assert body["formula_cells"] == 1
    assert any("recalculate" in note for note in body["notes"])


def test_stdio_edit_default_output_path_uses_edited_suffix(tmp_path):
    src = build_minimal_workbook(tmp_path / "widgets.xlsx")
    result = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="A1", value=1)], output_path=None)
    assert Path(result.structured_content["output_path"]).name == "widgets.edited.xlsx"


def test_stdio_edit_bad_sheet_returns_sheet_not_found_with_available_sheets(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_edit(str(src), [CellChange(sheet="NoSuchSheet", cell="A1", value=1)])
    error = result.structured_content["error"]
    assert error["code"] == "sheet_not_found"
    assert error["details"]["available_sheets"] == ["Sheet1"]


def test_stdio_edit_invalid_cell_ref_returns_invalid_change(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="ZZZZ1", value=1)])
    assert result.structured_content["error"]["code"] == "invalid_change"


def test_stdio_edit_too_many_changes_returns_too_many_changes(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    changes = [CellChange(sheet="Sheet1", cell=f"A{i + 1}", value=i) for i in range(S.MAX_CHANGES + 1)]
    result = stdio_edit(str(src), changes)
    error = result.structured_content["error"]
    assert error["code"] == "too_many_changes"
    assert error["details"]["limit"] == S.MAX_CHANGES


def test_stdio_edit_output_exists_without_overwrite_returns_output_exists(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    dst = tmp_path / "out.xlsx"
    dst.write_bytes(b"pre-existing")
    result = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="A1", value=1)], output_path=str(dst))
    assert result.structured_content["error"]["code"] == "output_exists"
    assert dst.read_bytes() == b"pre-existing"


def test_stdio_edit_overwrite_true_replaces_existing_output(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    dst = tmp_path / "out.xlsx"
    dst.write_bytes(b"pre-existing")
    result = stdio_edit(
        str(src), [CellChange(sheet="Sheet1", cell="A1", value=7)], output_path=str(dst), overwrite=True
    )
    assert result.structured_content["ok"] is True
    assert dst.read_bytes() != b"pre-existing"


def test_stdio_edit_extension_mismatch_returns_error(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_edit(
        str(src), [CellChange(sheet="Sheet1", cell="A1", value=1)], output_path=str(tmp_path / "out.xlsm")
    )
    assert result.structured_content["error"]["code"] == "extension_mismatch"


# ===========================================================================
# stdio_diff
# ===========================================================================


def test_stdio_diff_happy_path_reports_summary(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    edited = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="B2", value=99)])
    after = edited.structured_content["output_path"]
    result = stdio_diff(str(src), after)
    body = result.structured_content
    assert body["ok"] is True
    assert body["summary"]["cells_changed"] + body["summary"]["cells_added"] >= 1


def test_stdio_diff_missing_before_path_returns_file_not_found(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = stdio_diff(str(tmp_path / "nope.xlsx"), str(src))
    assert result.structured_content["error"]["code"] == "file_not_found"


# ===========================================================================
# http_inspect / http intake rules (base64 only: no paths, no URLs)
# ===========================================================================


def test_http_inspect_happy_path_hides_server_path(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = http_inspect(_b64_file(src))
    body = result.structured_content
    assert body["ok"] is True
    assert "path" not in body
    assert body["sheets"][0]["name"] == "Sheet1"


@pytest.mark.parametrize(
    "bad_input",
    [
        "/etc/passwd",
        "~/secret.xlsx",
        "C:\\Users\\me\\book.xlsx",
        "file:///etc/passwd",
        "https://example.com/book.xlsx",
        "../../etc/passwd",
        "somefile.xlsx",  # not path-shaped, but ends with a workbook suffix
    ],
)
def test_http_transports_reject_path_and_url_inputs(bad_input):
    for fn, kwargs in (
        (http_inspect, {"file_base64": bad_input}),
        (http_edit, {"file_base64": bad_input, "changes": [CellChange(sheet="Sheet1", cell="A1", value=1)]}),
        (http_diff, {"before_base64": bad_input, "after_base64": bad_input}),
    ):
        result = fn(**kwargs)
        error = result.structured_content["error"]
        assert error["code"] == "path_not_allowed", (fn.__name__, bad_input, error)
        assert "/etc/passwd" not in error["message"]  # never echoes the raw dangerous string as a path either


def test_http_inspect_invalid_base64_returns_error():
    result = http_inspect("not-valid-base64-!!!@@@")
    assert result.structured_content["error"]["code"] == "invalid_base64"


def test_http_inspect_empty_string_returns_invalid_base64():
    result = http_inspect("   ")
    assert result.structured_content["error"]["code"] == "invalid_base64"


def test_http_inspect_non_zip_bytes_returns_not_a_workbook():
    result = http_inspect(_b64_bytes(b"just some random bytes, not a zip"))
    assert result.structured_content["error"]["code"] == "not_a_workbook"


def test_http_inspect_legacy_ole_magic_returns_unsupported_file_type():
    result = http_inspect(_b64_bytes(S._OLE_MAGIC + b"\x00" * 32))
    assert result.structured_content["error"]["code"] == "unsupported_file_type"


def test_http_inspect_zip_without_workbook_xml_returns_not_a_workbook():
    result = http_inspect(_b64_bytes(_tiny_zip_without_workbook_xml()))
    error = result.structured_content["error"]
    assert error["code"] == "not_a_workbook"
    assert "workbook.xml" in error["message"]


def test_http_upload_size_cap_is_enforced():
    # Deliberately longer than the base64-char cap so the length check fires
    # before any base64 decoding is attempted (cheap: no ~20MB decode).
    oversized = "A" * (S.MAX_UPLOAD_B64_CHARS + 1)
    result = http_inspect(oversized)
    error = result.structured_content["error"]
    assert error["code"] == "file_too_large"
    assert error["details"]["limit_bytes"] == S.MAX_UPLOAD_BYTES


def test_http_zip_bomb_entry_count_is_rejected():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
        for i in range(S.MAX_ZIP_ENTRIES + 1):
            z.writestr(f"junk/{i}.txt", "")
    result = http_inspect(_b64_bytes(buf.getvalue()))
    assert result.structured_content["error"]["code"] == "zip_rejected"


# ===========================================================================
# http_edit
# ===========================================================================


def test_http_edit_happy_path_returns_base64_and_truncated_text_placeholder(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = http_edit(_b64_file(src), [CellChange(sheet="Sheet1", cell="B2", value=42)])
    body = result.structured_content
    assert body["ok"] is True
    assert body["file_extension"] == ".xlsx"
    raw = base64.b64decode(body["file_base64"])
    assert body["size_bytes"] == len(raw)
    assert ">42<" in _sheet1_xml_from_bytes(raw)
    # The text content block must not repeat the (potentially huge) blob --
    # only structuredContent carries the full base64 string.
    text = result.content[0].text
    assert "base64 chars; read structuredContent.file_base64" in text
    assert body["file_base64"] not in text


def test_http_edit_bad_sheet_returns_sheet_not_found(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = http_edit(_b64_file(src), [CellChange(sheet="Nope", cell="A1", value=1)])
    error = result.structured_content["error"]
    assert error["code"] == "sheet_not_found"
    assert error["details"]["available_sheets"] == ["Sheet1"]


def test_http_edit_empty_changes_returns_invalid_change(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = http_edit(_b64_file(src), [])
    assert result.structured_content["error"]["code"] == "invalid_change"


# ===========================================================================
# http_diff
# ===========================================================================


def test_http_diff_happy_path_hides_server_paths(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    edited = stdio_edit(str(src), [CellChange(sheet="Sheet1", cell="B2", value=7)])
    after_path = Path(edited.structured_content["output_path"])
    result = http_diff(_b64_file(src), _b64_file(after_path))
    body = result.structured_content
    assert body["ok"] is True
    assert "path_before" not in body
    assert "path_after" not in body
    assert body["summary"]["cells_changed"] + body["summary"]["cells_added"] >= 1


def test_http_diff_invalid_base64_on_one_side_returns_error(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    result = http_diff(_b64_file(src), "!!!not-base64!!!")
    assert result.structured_content["error"]["code"] == "invalid_base64"


# ===========================================================================
# Concurrency: parallel HTTP requests must not corrupt each other
# ===========================================================================


def test_http_edit_concurrent_requests_do_not_corrupt_each_other(tmp_path):
    n = 8
    payloads = []
    for i in range(n):
        src = build_minimal_workbook(tmp_path / f"src{i}.xlsx")
        payloads.append(_b64_file(src))

    results: list = [None] * n
    errors: list = [None] * n

    def worker(i: int) -> None:
        try:
            results[i] = http_edit(payloads[i], [CellChange(sheet="Sheet1", cell="B2", value=1000 + i)])
        except Exception as exc:  # noqa: BLE001 - captured for the assertion below
            errors[i] = exc

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(n)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=30)

    assert errors == [None] * n
    assert all(r is not None and r.structured_content["ok"] for r in results)

    for i, r in enumerate(results):
        raw = base64.b64decode(r.structured_content["file_base64"])
        xml = _sheet1_xml_from_bytes(raw)
        assert f">{1000 + i}<" in xml
        for j in range(n):
            if j != i:
                assert f">{1000 + j}<" not in xml, f"thread {i}'s output contains thread {j}'s value"


def test_http_server_busy_when_concurrency_slots_are_exhausted(monkeypatch):
    acquired = [S._http_slots.acquire(timeout=5) for _ in range(S.HTTP_MAX_CONCURRENT_JOBS)]
    assert all(acquired), "test setup could not acquire all concurrency slots"
    monkeypatch.setattr(S, "HTTP_BUSY_TIMEOUT_S", 0.05)
    try:
        result = S._run_http(lambda: {"unreachable": True})
        assert result.structured_content["error"]["code"] == "server_busy"
        assert result.is_error is True
    finally:
        for _ in acquired:
            S._http_slots.release()


# ===========================================================================
# build_server: tool wiring sanity (both transports)
# ===========================================================================


def test_build_server_rejects_unknown_mode():
    with pytest.raises(ValueError, match="unknown mode"):
        build_server("carrier-pigeon")  # type: ignore[arg-type]


@pytest.mark.parametrize("mode", ["stdio", "http"])
def test_build_server_registers_exactly_the_documented_tools(mode):
    server = build_server(mode)
    tools = asyncio.run(server.list_tools())
    names = sorted(t.name for t in tools)
    assert names == [
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
    for tool in tools:
        assert "never execute" not in tool.description or "macro" in tool.description.lower()


# ===========================================================================
# Structural tools: insert_rows / rename_sheet / add_sheet (both transports)
# ===========================================================================


def _part_from_bytes(raw: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        return z.read(name).decode("utf-8")


def test_stdio_insert_rows_shifts_cells_down(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx", cells={"A2": '<c r="A2"><v>7</v></c>'})
    result = stdio_insert_rows(str(src), "Sheet1", 2, 2)
    body = result.structured_content
    assert body["ok"] is True
    xml = _sheet1_xml_from_bytes(Path(body["output_path"]).read_bytes())
    assert 'r="A4"' in xml
    assert 'r="A2"' not in xml


def test_stdio_insert_rows_bad_sheet_returns_sheet_not_found(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    error = stdio_insert_rows(str(src), "Nope", 1).structured_content["error"]
    assert error["code"] == "sheet_not_found"
    assert error["details"]["available_sheets"] == ["Sheet1"]


def test_stdio_rename_sheet_updates_workbook_xml(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    body = stdio_rename_sheet(str(src), "Sheet1", "Budget").structured_content
    assert body["ok"] is True
    wb = _part_from_bytes(Path(body["output_path"]).read_bytes(), "xl/workbook.xml")
    assert 'name="Budget"' in wb
    assert 'name="Sheet1"' not in wb


def test_stdio_rename_sheet_to_existing_name_returns_conflict(tmp_path):
    src = build_two_sheet_workbook(tmp_path / "book.xlsx")
    error = stdio_rename_sheet(str(src), "Sheet1", "Sheet2").structured_content["error"]
    assert error["code"] == "sheet_name_conflict"


def test_stdio_add_sheet_registers_new_sheet(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    body = stdio_add_sheet(str(src), "Notes").structured_content
    assert body["ok"] is True
    out = Path(body["output_path"])
    assert S._sheet_names(out) == ["Sheet1", "Notes"]


def test_stdio_add_sheet_duplicate_name_returns_conflict(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    error = stdio_add_sheet(str(src), "Sheet1").structured_content["error"]
    assert error["code"] == "sheet_name_conflict"


def test_http_insert_rows_returns_shifted_workbook(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx", cells={"A2": '<c r="A2"><v>7</v></c>'})
    body = http_insert_rows(_b64_file(src), "Sheet1", 1).structured_content
    assert body["ok"] is True
    assert 'r="A3"' in _sheet1_xml_from_bytes(base64.b64decode(body["file_base64"]))
    assert "output_path" not in body


def test_http_rename_sheet_returns_renamed_workbook(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    body = http_rename_sheet(_b64_file(src), "Sheet1", "Q3").structured_content
    assert body["ok"] is True
    assert 'name="Q3"' in _part_from_bytes(base64.b64decode(body["file_base64"]), "xl/workbook.xml")


def test_http_add_sheet_returns_workbook_with_new_sheet(tmp_path):
    src = build_minimal_workbook(tmp_path / "book.xlsx")
    body = http_add_sheet(_b64_file(src), "Notes").structured_content
    assert body["ok"] is True
    assert 'name="Notes"' in _part_from_bytes(base64.b64decode(body["file_base64"]), "xl/workbook.xml")
