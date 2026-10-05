"""SafeXL MCP server: lets any MCP-capable agent inspect, edit, and diff
existing Excel workbooks (.xlsx/.xlsm) without breaking them.

    safexl-mcp                          # stdio (default) -- local file paths
    safexl-mcp --http --port 8000       # streamable HTTP -- base64 file content only

Thin wrappers over the library in this package:

    safexl_inspect  -> safexl_mcp.inspect.inspect_workbook   (read-only)
    safexl_edit     -> safexl_mcp.engine.edit                 (zip surgery + validate-then-move)
    safexl_diff     -> safexl_mcp.diff.diff_workbooks         (read-only)
    safexl_verify   -> safexl_mcp.verify.verify               (read-only receipt)
    safexl_view     -> safexl_mcp.view.view                   (read-only, bounded)
    safexl_salvage  -> safexl_mcp.salvage.salvage             (rebuild from original + intended edits)
    safexl_resize_table -> safexl_mcp.structural.resize_table
    plus insert_rows / insert_cells / rename_sheet / add_sheet structural edits

FILE INTAKE (the security boundary)
-----------------------------------
stdio: the server runs on the user's own machine as the user, so tools take
local file paths -- the same trust level as the ``safexl`` CLI.

HTTP: the server may be reachable by other people's agents, so tools accept
base64 file content ONLY. There is no path parameter, strings that look like a
path or URL are rejected before decoding, nothing is ever fetched from a URL,
and uploads are capped at 20 MB (decoded) per file. Each call works inside a
private temp directory that is deleted when the call returns, and temp paths
are never echoed back. The edited workbook comes back as base64. Uploads are
also checked for zip-bomb shape (entry count, declared uncompressed size)
before the engine touches them, and at most ``HTTP_MAX_CONCURRENT_JOBS`` file
jobs run at once.

ERRORS
------
Every tool returns a JSON object. Success: ``{"ok": true, ...}``. Failure:
``{"ok": false, "error": {"code": ..., "message": ..., "details": {...}}}``
with the MCP ``isError`` flag set, so an agent can branch on ``error.code``
instead of parsing prose. Argument *shape* errors (wrong JSON type, unknown
field) are rejected by the SDK's schema validation before a tool runs and
also come back with ``isError`` set.

Macros are never executed by any tool here: the engine copies
``xl/vbaProject.bin`` byte-for-byte, and ``inspect`` only reports presence,
size, and (with the optional ``scan`` extra) a static keyword scan.
"""

import argparse
import base64
import binascii
import hashlib
import io
import json
import math
import re
import sys
import tempfile
import threading
import zipfile
from pathlib import Path
from typing import Annotated, Any, Callable, Literal
from xml.etree import ElementTree as ET

from mcp.server.mcpserver import MCPServer
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field

from . import __version__
from .diff import diff_workbooks
from .engine import SafeXLError, edit
from .inspect import inspect_workbook
from .salvage import salvage
from .refs import TornReference
from .structural import add_sheet, insert_cells, insert_rows, plan_table_extensions, rename_sheet, resize_table
from .verify import verify
from .view import ViewError, view

__all__ = ["build_server", "main", "CellChange"]

# ---------------------------------------------------------------------------
# Limits
# ---------------------------------------------------------------------------

MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB decoded, per file (HTTP mode)
MAX_UPLOAD_B64_CHARS = 4 * math.ceil(MAX_UPLOAD_BYTES / 3)
# A legitimate 20 MB workbook rarely expands past ~10x; 256 MB leaves headroom
# while keeping a zip bomb from exhausting memory in the XML parse/validate step.
MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024
MAX_ZIP_ENTRIES = 10_000
# safexl_diff takes two uploads, so the transport must admit two capped files.
HTTP_MAX_REQUEST_BODY = 2 * MAX_UPLOAD_B64_CHARS + 1024 * 1024
HTTP_MAX_CONCURRENT_JOBS = 4
HTTP_BUSY_TIMEOUT_S = 120

MAX_CHANGES = 1000
MAX_TEXT_LEN = 32_767  # Excel's per-cell character limit
MAX_FORMULA_LEN = 8_192  # Excel's formula length limit
MAX_COL = 16_384  # XFD
MAX_ROW = 1_048_576
MAX_DEFINED_NAMES = 200  # inspect output cap (token budget)
DEFAULT_DIFF_CELLS = 200
DEFAULT_VIEW_CELLS = 400
MAX_VIEW_CELLS = 2000  # hard server-side ceiling for safexl_view output (token budget)

_ALLOWED_SUFFIXES = (".xlsx", ".xlsm")
_ZIP_MAGIC = b"PK\x03\x04"
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy .xls, or an encrypted OOXML file
_CELL_RE = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$")
_URL_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*://")
_PATHLIKE_RE = re.compile(r"""^(?:/|~|\.{1,2}[/\\]|[A-Za-z]:[\\/]|\\\\|file:|[A-Za-z][A-Za-z0-9+.-]*://)""")
_DATA_URI_RE = re.compile(r"^data:[^,]*;base64,", re.IGNORECASE)
_TMP_PREFIX = "safexl-mcp-"
# Matches "<anything>/safexl-mcp-XXXX/" so HTTP error messages never reveal server paths.
_TMP_PATH_RE = re.compile(r"[^\s'\"]*" + _TMP_PREFIX + r"[^/\\\s'\"]+[/\\]")

_FORMULA_NOTE = (
    "Formula cells are stored without a cached result and compute when the workbook is next opened "
    "in Excel or LibreOffice (the workbook is flagged to recalculate on open). A library that reads "
    "the file without opening it first (openpyxl, pandas) sees those cells as empty."
)

_http_slots = threading.BoundedSemaphore(HTTP_MAX_CONCURRENT_JOBS)


# ---------------------------------------------------------------------------
# Structured results
# ---------------------------------------------------------------------------


class ToolInputError(Exception):
    """A caller-fixable problem. Becomes ``{"ok": false, "error": {...}}``."""

    def __init__(self, code: str, message: str, **details: Any) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details


def _ok(payload: dict[str, Any], *, blob_field: str | None = None) -> CallToolResult:
    body = {"ok": True, **payload}
    text_body = body
    if blob_field and blob_field in body:
        # Don't send a 27 MB base64 string twice: structuredContent carries it,
        # the text block carries everything else plus a pointer to it.
        text_body = dict(body)
        text_body[blob_field] = f"<{len(body[blob_field])} base64 chars; read structuredContent.{blob_field}>"
    return CallToolResult(
        content=[TextContent(type="text", text=json.dumps(text_body, ensure_ascii=False))],
        structured_content=body,
    )


def _err(code: str, message: str, details: dict[str, Any] | None = None) -> CallToolResult:
    error: dict[str, Any] = {"code": code, "message": message}
    if details:
        error["details"] = details
    body = {"ok": False, "error": error}
    return CallToolResult(
        content=[TextContent(type="text", text=json.dumps(body, ensure_ascii=False))],
        structured_content=body,
        is_error=True,
    )


def _run(job: Callable[[], dict[str, Any]], *, blob_field: str | None = None, remote: bool = False) -> CallToolResult:
    """Run a tool body and turn every outcome into a structured result.

    ``remote`` (HTTP mode) strips server temp-dir paths from error messages and
    hides unexpected-exception text, so nothing about the host leaks to callers."""

    def clean(msg: str) -> str:
        return _TMP_PATH_RE.sub("", msg) if remote else msg

    try:
        return _ok(job(), blob_field=blob_field)
    except ToolInputError as exc:
        return _err(exc.code, clean(exc.message), exc.details)
    except SafeXLError as exc:
        return _err("invalid_workbook", clean(str(exc)))
    except Exception as exc:  # noqa: BLE001 - never crash the session
        detail = "" if remote else f": {exc}"
        return _err("internal_error", f"unexpected {type(exc).__name__} while processing the workbook{detail}")


def _run_http(job: Callable[[], dict[str, Any]], *, blob_field: str | None = None) -> CallToolResult:
    if not _http_slots.acquire(timeout=HTTP_BUSY_TIMEOUT_S):
        return _err("server_busy", "too many workbooks in flight on this server; retry shortly")
    try:
        return _run(job, blob_field=blob_field, remote=True)
    finally:
        _http_slots.release()


# ---------------------------------------------------------------------------
# Change validation (tool-boundary checks; the engine re-validates too)
# ---------------------------------------------------------------------------


class CellChange(BaseModel):
    """One cell write. Exactly one of ``value`` / ``formula``."""

    model_config = ConfigDict(extra="forbid")

    sheet: str = Field(
        description="Worksheet name exactly as it appears on the tab (safexl_inspect lists them), e.g. 'Sheet1'."
    )
    cell: str = Field(description="Single A1-style cell reference, e.g. 'B7'. No sheet prefix, no ranges.")
    value: bool | int | float | str | None = Field(
        default=None,
        description=(
            "Literal value to store. JSON numbers and booleans are stored as numbers and booleans; "
            "strings are stored as text (send 42, not \"42\", for a number). Give value OR formula."
        ),
    )
    formula: str | None = Field(
        default=None,
        description=(
            "Formula to store, with or without a leading '=', e.g. '=SUM(B2:B9)'. It computes when the "
            "workbook is next opened in Excel/LibreOffice. Give value OR formula."
        ),
    )


def _col_to_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _engine_changes(changes: list[CellChange]) -> list[dict[str, Any]]:
    if not changes:
        raise ToolInputError("invalid_change", "changes must contain at least one change")
    if len(changes) > MAX_CHANGES:
        raise ToolInputError(
            "too_many_changes",
            f"{len(changes)} changes in one call; the limit is {MAX_CHANGES}. Split the edit into several calls.",
            limit=MAX_CHANGES,
        )

    out: list[dict[str, Any]] = []
    for i, ch in enumerate(changes):
        where = f"changes[{i}]"
        if not ch.sheet:
            raise ToolInputError("invalid_change", f"{where}: 'sheet' must be a non-empty sheet name", index=i)

        cell = ch.cell.strip().replace("$", "")
        if "!" in cell:
            raise ToolInputError(
                "invalid_change",
                f"{where}: 'cell' must not include a sheet prefix ({ch.cell!r}); put the sheet name in 'sheet'",
                index=i,
            )
        m = _CELL_RE.match(cell)
        if not m or _col_to_num(m.group(1)) > MAX_COL or int(m.group(2)) > MAX_ROW:
            raise ToolInputError(
                "invalid_change",
                f"{where}: {ch.cell!r} is not a single cell reference within A1:XFD1048576",
                index=i,
            )
        cell = cell.upper()

        has_value = ch.value is not None
        has_formula = ch.formula is not None
        if has_value and has_formula:
            raise ToolInputError(
                "invalid_change", f"{where} ({ch.sheet}!{cell}): give 'value' or 'formula', not both", index=i
            )
        if not has_value and not has_formula:
            raise ToolInputError(
                "invalid_change",
                f"{where} ({ch.sheet}!{cell}): give a non-null 'value' or a 'formula' (clearing cells is not supported)",
                index=i,
            )

        if has_formula:
            formula = ch.formula.strip()  # type: ignore[union-attr]
            if not formula.lstrip("=").strip():
                raise ToolInputError("invalid_change", f"{where} ({ch.sheet}!{cell}): formula is empty", index=i)
            if len(formula) > MAX_FORMULA_LEN:
                raise ToolInputError(
                    "invalid_change",
                    f"{where} ({ch.sheet}!{cell}): formula is {len(formula)} chars; Excel's limit is {MAX_FORMULA_LEN}",
                    index=i,
                )
            out.append({"sheet": ch.sheet, "cell": cell, "formula": formula})
            continue

        value = ch.value
        if isinstance(value, float) and not math.isfinite(value):
            raise ToolInputError(
                "invalid_change", f"{where} ({ch.sheet}!{cell}): NaN/Infinity cannot be stored in a cell", index=i
            )
        if isinstance(value, str) and len(value) > MAX_TEXT_LEN:
            raise ToolInputError(
                "invalid_change",
                f"{where} ({ch.sheet}!{cell}): text is {len(value)} chars; Excel's per-cell limit is {MAX_TEXT_LEN}",
                index=i,
            )
        out.append({"sheet": ch.sheet, "cell": cell, "value": value})
    return out


def _sheet_names(path: Path) -> list[str]:
    try:
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("xl/workbook.xml"))
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        return [el.get("name", "") for el in root.iter(f"{ns}sheet")]
    except Exception:  # noqa: BLE001 - best-effort hint only
        return []


def _write_refusal(msg: str) -> ToolInputError | None:
    """Typed codes for the two write-side refusals every editing tool can hit; None for anything else."""
    if "read-back failed" in msg:
        return ToolInputError("write_not_confirmed", msg)
    if "namespace prefix" in msg:
        return ToolInputError("unsupported_workbook", msg)
    return None


def _edit_error(exc: SafeXLError, src: Path) -> ToolInputError:
    msg = str(exc)
    refusal = _write_refusal(msg)
    if refusal:
        return refusal
    if re.match(r"sheet .* not found in workbook", msg):
        return ToolInputError("sheet_not_found", msg, available_sheets=_sheet_names(src))
    return ToolInputError("edit_failed", msg)


def _structural_error(exc: SafeXLError, src: Path) -> ToolInputError:
    """Same shape as _edit_error, plus the sheet-name-conflict case that
    insert_rows/rename_sheet/add_sheet can raise that safexl_edit cannot."""
    msg = str(exc)
    refusal = _write_refusal(msg)
    if refusal:
        return refusal
    if re.match(r"sheet .* not found in workbook", msg):
        return ToolInputError("sheet_not_found", msg, available_sheets=_sheet_names(src))
    if "already exists" in msg:
        return ToolInputError("sheet_name_conflict", msg, available_sheets=_sheet_names(src))
    return ToolInputError("structural_edit_failed", msg)


MAX_RECEIPT_SHEETS = 50


def _bounded_receipt(receipt: dict[str, Any]) -> dict[str, Any]:
    """Cap cell_changes at MAX_RECEIPT_SHEETS sheets; say so when it was cut."""
    changes = receipt.get("cell_changes")
    if isinstance(changes, dict) and len(changes) > MAX_RECEIPT_SHEETS:
        receipt = dict(receipt)
        receipt["cell_changes"] = dict(list(changes.items())[:MAX_RECEIPT_SHEETS])
        receipt["cell_changes_truncated"] = {"shown": MAX_RECEIPT_SHEETS, "total": len(changes)}
    return receipt


def _part_hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as zf:
        return {n: hashlib.sha256(zf.read(n)).hexdigest() for n in zf.namelist() if not n.endswith("/")}


def _parts_summary(before: dict[str, str], after: dict[str, str]) -> dict[str, Any]:
    common = before.keys() & after.keys()
    return {
        "changed": sorted(n for n in common if before[n] != after[n]),
        "added": sorted(after.keys() - before.keys()),
        "removed": sorted(before.keys() - after.keys()),
        "unchanged_count": sum(1 for n in common if before[n] == after[n]),
    }


def _edit_result(changes: list[dict[str, Any]], parts: dict[str, Any]) -> dict[str, Any]:
    cells = {(c["sheet"], c["cell"]) for c in changes}
    formulas = sum(1 for c in changes if "formula" in c)
    notes = [
        f"{parts['unchanged_count']} internal parts are byte-identical to the source (charts, pivot tables, "
        "validation, conditional formatting, and VBA live there); parts.changed lists the rest. "
        "Call safexl_diff for a cell-level before/after."
    ]
    if formulas:
        notes.insert(0, _FORMULA_NOTE)
    as_text = [f"{c['sheet']}!{c['cell']}" for c in changes if isinstance(c.get("value"), str) and c["value"].startswith("=")]
    if as_text:
        more = " ..." if len(as_text) > 10 else ""
        notes.append(
            f"Stored as TEXT because they were sent as 'value': {', '.join(as_text[:10])}{more}. "
            "Send them as 'formula' to store a formula."
        )
    return {
        "cells_written": len(cells),
        "formula_cells": formulas,
        "sheets_edited": sorted({c["sheet"] for c in changes}),
        "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        "parts": parts,
        "notes": notes,
    }


# ---------------------------------------------------------------------------
# Output shaping (keep agent context small)
# ---------------------------------------------------------------------------


def _shape_inspection(report: dict[str, Any]) -> dict[str, Any]:
    names = report.get("defined_names") or []
    if len(names) > MAX_DEFINED_NAMES:
        report["defined_names"] = names[:MAX_DEFINED_NAMES]
        report["defined_names_truncated"] = True
    report["defined_names_total"] = len(names)
    return report


def _shape_diff(result: dict[str, Any], max_cells: int) -> dict[str, Any]:
    budget = max_cells
    totals = {"changed": 0, "added": 0, "removed": 0}
    truncated = False
    for sheet in result["sheets"]:
        for kind in ("changed", "added", "removed"):
            key = f"{kind}_cells"
            cells = sheet[key]
            sheet[f"{kind}_count"] = len(cells)
            totals[kind] += len(cells)
            if len(cells) > budget:
                sheet[key] = cells[:budget]
                truncated = True
            budget -= len(sheet[key])
    parts = result["parts"]
    result["summary"] = {
        "sheets_with_cell_changes": len(result["sheets"]),
        "cells_changed": totals["changed"],
        "cells_added": totals["added"],
        "cells_removed": totals["removed"],
        "parts_changed": len(parts["changed"]),
        "parts_added": len(parts["added"]),
        "parts_removed": len(parts["removed"]),
        "parts_unchanged": parts["unchanged_count"],
    }
    result["cells_truncated"] = truncated
    return result


# ---------------------------------------------------------------------------
# stdio mode: local paths
# ---------------------------------------------------------------------------


def _local_path(raw: str, field: str) -> Path:
    s = (raw or "").strip()
    if not s:
        raise ToolInputError("invalid_path", f"'{field}' must be a local file path")
    if _URL_RE.match(s):
        raise ToolInputError("invalid_path", f"'{field}': URLs are not accepted; pass a local file path")
    p = Path(s).expanduser()
    if not p.is_absolute():
        p = Path.cwd() / p
    p = p.resolve()
    if p.suffix.lower() not in _ALLOWED_SUFFIXES:
        raise ToolInputError(
            "unsupported_file_type",
            f"'{field}' must be an .xlsx or .xlsm file (got {p.name!r}); legacy .xls, .xlsb, .csv and .ods "
            "are not supported",
            path=str(p),
        )
    return p


def _existing_workbook(raw: str, field: str) -> Path:
    p = _local_path(raw, field)
    if not p.is_file():
        raise ToolInputError("file_not_found", f"'{field}': no such file: {p}", path=str(p))
    with p.open("rb") as fh:
        head = fh.read(8)
    _check_magic(head, field)
    return p


def _check_magic(head: bytes, field: str) -> None:
    if head.startswith(_ZIP_MAGIC):
        return
    if head.startswith(_OLE_MAGIC):
        raise ToolInputError(
            "unsupported_file_type",
            f"'{field}' is a legacy .xls file or a password-protected (encrypted) workbook; SafeXL edits "
            "unencrypted .xlsx/.xlsm only",
        )
    raise ToolInputError("not_a_workbook", f"'{field}' is not an .xlsx/.xlsm (OOXML zip) file")


PathArg = Annotated[str, Field(description="Path to an existing .xlsx or .xlsm file on this machine (absolute path recommended).")]


def _resolve_output_path(src: Path, output_path: str | None, overwrite: bool, *, suffix_tag: str) -> Path:
    """Same output_path/overwrite handling as stdio_edit's inline logic, shared
    by the structural-edit tools below (kept separate from stdio_edit itself
    so its existing, already-tested behavior is untouched)."""
    if output_path is None or not output_path.strip():
        dst = src.with_name(f"{src.stem}.{suffix_tag}{src.suffix}")
    else:
        dst = _local_path(output_path, "output_path")
    if dst.suffix.lower() != src.suffix.lower():
        raise ToolInputError(
            "extension_mismatch",
            f"output_path must keep the source extension {src.suffix!r} (got {dst.suffix!r}); Excel refuses "
            "to open a workbook whose extension doesn't match its content type",
        )
    if not dst.parent.is_dir():
        raise ToolInputError("invalid_path", f"output directory does not exist: {dst.parent}")
    if dst.exists() and not overwrite:
        raise ToolInputError(
            "output_exists",
            f"{dst} already exists; pass overwrite=true to replace it, or choose another output_path",
            output_path=str(dst),
        )
    return dst


def stdio_inspect(path: PathArg) -> CallToolResult:
    def job() -> dict[str, Any]:
        p = _existing_workbook(path, "path")
        return _shape_inspection(inspect_workbook(p).to_dict())

    return _run(job)


ForceArg = Annotated[
    bool,
    Field(
        description=(
            "Rebuild from the original even when the damaged file looks intact or was only re-serialised without "
            "losing a feature (default false: that case returns nothing_to_do and writes nothing)."
        )
    ),
]
ExtendTablesArg = Annotated[
    bool,
    Field(
        description=(
            "Grow an Excel table when a written cell sits directly next to it (a new column to its right or a new "
            "row below), so the new cells join the table. Default true; the result's tables_extended lists each "
            "table that grew. Pass false to write the cells without touching any table."
        )
    ),
]


def stdio_edit(
    path: PathArg,
    changes: Annotated[
        list[CellChange],
        Field(description=f"Cell writes to apply (1-{MAX_CHANGES}). A later write to the same cell wins."),
    ],
    output_path: Annotated[
        str | None,
        Field(
            description=(
                "Where to write the edited workbook; must keep the source's extension (.xlsx stays .xlsx, "
                ".xlsm stays .xlsm). Default: '<name>.edited.<ext>' next to the source. The source file is "
                "never modified unless output_path equals path and overwrite is true."
            )
        ),
    ] = None,
    overwrite: Annotated[
        bool, Field(description="Allow replacing an existing file at output_path (including the source itself).")
    ] = False,
    extend_tables: ExtendTablesArg = True,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        engine_changes = _engine_changes(changes)
        if output_path is None or not output_path.strip():
            dst = src.with_name(f"{src.stem}.edited{src.suffix}")
        else:
            dst = _local_path(output_path, "output_path")
        if dst.suffix.lower() != src.suffix.lower():
            raise ToolInputError(
                "extension_mismatch",
                f"output_path must keep the source extension {src.suffix!r} (got {dst.suffix!r}); Excel refuses "
                "to open a workbook whose extension doesn't match its content type",
            )
        if not dst.parent.is_dir():
            raise ToolInputError("invalid_path", f"output directory does not exist: {dst.parent}")
        if dst.exists() and not overwrite:
            raise ToolInputError(
                "output_exists",
                f"{dst} already exists; pass overwrite=true to replace it, or choose another output_path",
                output_path=str(dst),
            )

        before = _part_hashes(src)
        try:
            extended = plan_table_extensions(src, engine_changes) if extend_tables else []
            edit(src, dst, engine_changes, extend_tables=extend_tables)
        except SafeXLError as exc:
            raise _edit_error(exc, src) from exc
        result = _edit_result(engine_changes, _parts_summary(before, _part_hashes(dst)))
        return {"source_path": str(src), "output_path": str(dst), "tables_extended": extended, **result}

    return _run(job)


def stdio_diff(
    before_path: Annotated[str, Field(description="The original .xlsx/.xlsm (e.g. the file you passed to safexl_edit).")],
    after_path: Annotated[str, Field(description="The changed .xlsx/.xlsm (e.g. safexl_edit's output_path).")],
    max_cells: Annotated[
        int, Field(ge=1, le=5000, description="Maximum cell entries to list across all sheets; counts are always exact.")
    ] = DEFAULT_DIFF_CELLS,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        a = _existing_workbook(before_path, "before_path")
        b = _existing_workbook(after_path, "after_path")
        return _shape_diff(diff_workbooks(a, b).to_dict(), max_cells)

    return _run(job)


OutputPathArg = Annotated[
    str | None,
    Field(
        description=(
            "Where to write the edited workbook; must keep the source's extension (.xlsx stays .xlsx, "
            ".xlsm stays .xlsm). Default: '<name>.<op>.<ext>' next to the source. The source file is "
            "never modified unless output_path equals path and overwrite is true."
        )
    ),
]
OverwriteArg = Annotated[
    bool, Field(description="Allow replacing an existing file at output_path (including the source itself).")
]
SheetArg = Annotated[
    str | None, Field(description="Worksheet name to read. Default: the first sheet (all sheets when 'find' is set).")
]
CellRangeArg = Annotated[
    str | None, Field(description="A1-style window to read, e.g. 'A1:F40'. Default: from the top of the sheet.")
]
FindArg = Annotated[
    str | None,
    Field(description="Case-insensitive text to search for in cell values and formulas; returns the matching cells."),
]
MaxViewCellsArg = Annotated[
    int,
    Field(
        ge=1,
        le=MAX_VIEW_CELLS,
        description=f"Maximum cells to return (hard ceiling {MAX_VIEW_CELLS}); 'truncated' and 'next_range' say how to continue.",
    ),
]
FormulasArg = Annotated[bool, Field(description="Include formula text for formula cells (default true).")]
CellsRangeArg = Annotated[
    str,
    Field(
        description=(
            "A1-style range of blank cells to insert, e.g. 'B5:C6'. Cells in those columns at or below the range's "
            "first row shift down by its height; columns outside the range do not move."
        )
    ),
]
InsertShiftArg = Annotated[
    Literal["down"], Field(description="Only 'down' is supported (cells shift down within the range's columns).")
]
TableArg = Annotated[str, Field(description="The table's name or displayName (safexl_inspect lists tables).")]
TableRefArg = Annotated[
    str,
    Field(description="New range for the table, e.g. 'A1:D12'. The top-left cell must stay put; the right and bottom edges move."),
]
TableSheetArg = Annotated[str | None, Field(description="Limit the table lookup to this worksheet (optional).")]


def stdio_insert_rows(
    path: PathArg,
    sheet_name: Annotated[str, Field(description="Worksheet name to insert rows into, exactly as it appears on the tab.")],
    row: Annotated[
        int,
        Field(ge=1, description="1-based row number where the blank row(s) go; existing rows at/below this shift down."),
    ],
    count: Annotated[int, Field(ge=1, description="Number of blank rows to insert.")] = 1,
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="rows-inserted")
        before = _part_hashes(src)
        try:
            insert_rows(src, dst, sheet_name, row, count)
        except SafeXLError as exc:
            raise _structural_error(exc, src) from exc
        return {
            "source_path": str(src),
            "output_path": str(dst),
            "sheet": sheet_name,
            "row": row,
            "count": count,
            "parts": _parts_summary(before, _part_hashes(dst)),
            "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        }

    return _run(job)


def stdio_insert_cells(
    path: PathArg,
    sheet_name: Annotated[str, Field(description="Worksheet name to insert cells into, exactly as it appears on the tab.")],
    cell_range: CellsRangeArg,
    shift: InsertShiftArg = "down",
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="cells-inserted")
        before = _part_hashes(src)
        try:
            insert_cells(src, dst, sheet=sheet_name, cell_range=cell_range, shift=shift)
        except (SafeXLError, TornReference) as exc:
            raise _insert_cells_error(exc, src) from exc
        return {
            "source_path": str(src),
            "output_path": str(dst),
            "sheet": sheet_name,
            "cell_range": cell_range,
            "shift": shift,
            "parts": _parts_summary(before, _part_hashes(dst)),
            "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        }

    return _run(job)


def stdio_rename_sheet(
    path: PathArg,
    old_name: Annotated[str, Field(description="Worksheet name to rename, exactly as it appears on the tab.")],
    new_name: Annotated[
        str, Field(description="New worksheet name (max 31 chars; may not contain [ ] * ? / \\ : or start/end with ').")
    ],
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="renamed")
        before = _part_hashes(src)
        try:
            rename_sheet(src, dst, old_name, new_name)
        except SafeXLError as exc:
            raise _structural_error(exc, src) from exc
        return {
            "source_path": str(src),
            "output_path": str(dst),
            "old_name": old_name,
            "new_name": new_name,
            "parts": _parts_summary(before, _part_hashes(dst)),
            "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        }

    return _run(job)


def stdio_add_sheet(
    path: PathArg,
    sheet_name: Annotated[
        str, Field(description="Name for the new, empty worksheet (max 31 chars; may not contain [ ] * ? / \\ :).")
    ],
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="sheet-added")
        before = _part_hashes(src)
        try:
            add_sheet(src, dst, sheet_name)
        except SafeXLError as exc:
            raise _structural_error(exc, src) from exc
        return {
            "source_path": str(src),
            "output_path": str(dst),
            "sheet": sheet_name,
            "parts": _parts_summary(before, _part_hashes(dst)),
            "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        }

    return _run(job)


def _view_error(exc: SafeXLError, src: Path) -> ToolInputError:
    msg = str(exc)
    if isinstance(exc, ViewError):
        if msg.startswith("no sheet"):
            return ToolInputError("sheet_not_found", msg, available_sheets=_sheet_names(src))
        return ToolInputError("invalid_range", msg)
    return ToolInputError("view_failed", msg)


def _resize_error(exc: SafeXLError, src: Path) -> ToolInputError:
    msg = str(exc)
    refusal = _write_refusal(msg)
    if refusal:
        return refusal
    if re.match(r"table .* not found", msg):
        return ToolInputError("table_not_found", msg, available_sheets=_sheet_names(src))
    return ToolInputError("resize_failed", msg)


def _insert_cells_error(exc: Exception, src: Path) -> ToolInputError:
    msg = str(exc)
    refusal = _write_refusal(msg)
    if refusal:
        return refusal
    if isinstance(exc, TornReference) or "insert_cells refused" in msg:
        return ToolInputError("reference_would_tear", msg)
    if re.match(r"sheet .* not found in workbook", msg):
        return ToolInputError("sheet_not_found", msg, available_sheets=_sheet_names(src))
    if msg.startswith("invalid cell_range") or "outside the worksheet" in msg:
        return ToolInputError("invalid_range", msg)
    return ToolInputError("structural_edit_failed", msg)


def _salvage_result(result: dict[str, Any]) -> dict[str, Any]:
    result = dict(result)
    result["output_written"] = result.get("status") in ("salvaged", "partial")
    if result.get("status") == "nothing_to_do":
        result["message"] = (
            "The damaged file is already intact (nothing was lost); keep it as is. No output file was written. "
            "Pass force=true to rebuild from the original anyway."
        )
    return result


def stdio_verify(
    before_path: Annotated[str, Field(description="The ORIGINAL .xlsx/.xlsm, before the edit.")],
    after_path: Annotated[str, Field(description="The edited .xlsx/.xlsm to check against the original.")],
) -> CallToolResult:
    def job() -> dict[str, Any]:
        a = _existing_workbook(before_path, "before_path")
        b = _existing_workbook(after_path, "after_path")
        return _bounded_receipt(verify(a, b))

    return _run(job)


def stdio_view(
    path: PathArg,
    sheet: SheetArg = None,
    cell_range: CellRangeArg = None,
    find: FindArg = None,
    max_cells: MaxViewCellsArg = DEFAULT_VIEW_CELLS,
    formulas: FormulasArg = True,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        try:
            return view(src, sheet, cell_range, find, max_cells, formulas)
        except SafeXLError as exc:
            raise _view_error(exc, src) from exc

    return _run(job)


def stdio_salvage(
    original_path: Annotated[str, Field(description="The ORIGINAL .xlsx/.xlsm from before the damaging re-save.")],
    damaged_path: Annotated[str, Field(description="The workbook that was re-saved by another library and lost features.")],
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
    force: ForceArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(original_path, "original_path")
        dmg = _existing_workbook(damaged_path, "damaged_path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="salvaged")
        try:
            result = salvage(src, dmg, dst, force=force)
        except SafeXLError as exc:
            raise ToolInputError("salvage_failed", str(exc)) from exc
        result = _salvage_result(result)
        if result["output_written"]:
            result["output_path"] = str(dst)
        return {"source_path": str(src), **result}

    return _run(job)


def stdio_resize_table(
    path: PathArg,
    table: TableArg,
    ref: TableRefArg,
    sheet_name: TableSheetArg = None,
    output_path: OutputPathArg = None,
    overwrite: OverwriteArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        src = _existing_workbook(path, "path")
        dst = _resolve_output_path(src, output_path, overwrite, suffix_tag="table-resized")
        before = _part_hashes(src)
        try:
            resize_table(src, dst, table=table, ref=ref, sheet=sheet_name)
        except SafeXLError as exc:
            raise _resize_error(exc, src) from exc
        return {
            "source_path": str(src),
            "output_path": str(dst),
            "table": table,
            "ref": ref,
            "parts": _parts_summary(before, _part_hashes(dst)),
            "validation": "passed: 0 CRITICAL structural findings (the file is only written if this holds)",
        }

    return _run(job)


# ---------------------------------------------------------------------------
# HTTP mode: base64 content only
# ---------------------------------------------------------------------------


def _decode_upload(data: str, field: str) -> bytes:
    s = (data or "").strip()
    if not s:
        raise ToolInputError("invalid_base64", f"'{field}' must be the workbook's bytes, base64-encoded")
    s = _DATA_URI_RE.sub("", s, count=1)
    # An .xlsx/.xlsm is a zip, so its base64 always starts "UEsDB"; anything that
    # looks like a path or URL is refused before any decoding is attempted.
    if _PATHLIKE_RE.match(s) or s.lower().endswith((".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")):
        raise ToolInputError(
            "path_not_allowed",
            f"'{field}' looks like a file path or URL. This HTTP server never reads server-side paths or fetches "
            "URLs; send the workbook's bytes base64-encoded.",
        )
    s = re.sub(r"\s+", "", s)
    if len(s) > MAX_UPLOAD_B64_CHARS:
        raise ToolInputError(
            "file_too_large",
            f"'{field}' exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB upload limit",
            limit_bytes=MAX_UPLOAD_BYTES,
        )
    try:
        raw = base64.b64decode(s, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ToolInputError(
            "invalid_base64",
            f"'{field}' is not valid base64 ({exc}); send the workbook's bytes base64-encoded -- paths and URLs "
            "are not accepted",
        ) from exc
    if len(raw) > MAX_UPLOAD_BYTES:
        raise ToolInputError(
            "file_too_large",
            f"'{field}' is {len(raw)} bytes; the limit is {MAX_UPLOAD_BYTES}",
            limit_bytes=MAX_UPLOAD_BYTES,
        )
    _check_magic(raw[:8], field)
    _check_zip_shape(raw, field)
    return raw


def _check_zip_shape(raw: bytes, field: str) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            infos = zf.infolist()
            names = {i.filename for i in infos}
    except zipfile.BadZipFile as exc:
        raise ToolInputError("not_a_workbook", f"'{field}' is not a readable zip/OOXML package: {exc}") from exc
    if len(infos) > MAX_ZIP_ENTRIES:
        raise ToolInputError("zip_rejected", f"'{field}' has {len(infos)} zip entries; the limit is {MAX_ZIP_ENTRIES}")
    expanded = sum(i.file_size for i in infos)
    if expanded > MAX_UNCOMPRESSED_BYTES:
        raise ToolInputError(
            "zip_rejected",
            f"'{field}' expands to {expanded // (1024 * 1024)} MB uncompressed; the limit is "
            f"{MAX_UNCOMPRESSED_BYTES // (1024 * 1024)} MB",
        )
    if "xl/workbook.xml" not in names:
        raise ToolInputError("not_a_workbook", f"'{field}' is a zip but not an Excel workbook (no xl/workbook.xml)")


def _workbook_ext(raw: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        try:
            ct = zf.read("[Content_Types].xml")
        except KeyError:
            return ".xlsx"
    return ".xlsm" if b"macroEnabled" in ct else ".xlsx"


def _stage(tmp: Path, raw: bytes, stem: str) -> Path:
    p = tmp / f"{stem}{_workbook_ext(raw)}"
    p.write_bytes(raw)
    return p


B64Arg = Annotated[
    str,
    Field(
        description=(
            "The workbook's bytes (.xlsx or .xlsm), base64-encoded; max 20 MB decoded. File paths and URLs are "
            "rejected: this server never reads its own disk or fetches anything."
        )
    ),
]


def http_inspect(file_base64: B64Arg) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            report = inspect_workbook(_stage(Path(tmp), raw, "upload")).to_dict()
        report.pop("path", None)
        return _shape_inspection(report)

    return _run_http(job)


def http_edit(
    file_base64: B64Arg,
    changes: Annotated[
        list[CellChange],
        Field(description=f"Cell writes to apply (1-{MAX_CHANGES}). A later write to the same cell wins."),
    ],
    extend_tables: ExtendTablesArg = True,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        engine_changes = _engine_changes(changes)
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                extended = plan_table_extensions(src, engine_changes) if extend_tables else []
                edit(src, dst, engine_changes, extend_tables=extend_tables)
            except SafeXLError as exc:
                raise _edit_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "tables_extended": extended,
            **_edit_result(engine_changes, parts),
        }

    return _run_http(job, blob_field="file_base64")


def http_insert_rows(
    file_base64: B64Arg,
    sheet_name: Annotated[str, Field(description="Worksheet name to insert rows into, exactly as it appears on the tab.")],
    row: Annotated[
        int,
        Field(ge=1, description="1-based row number where the blank row(s) go; existing rows at/below this shift down."),
    ],
    count: Annotated[int, Field(ge=1, description="Number of blank rows to insert.")] = 1,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                insert_rows(src, dst, sheet_name, row, count)
            except SafeXLError as exc:
                raise _structural_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "sheet": sheet_name,
            "row": row,
            "count": count,
            "parts": parts,
        }

    return _run_http(job, blob_field="file_base64")


def http_insert_cells(
    file_base64: B64Arg,
    sheet_name: Annotated[str, Field(description="Worksheet name to insert cells into, exactly as it appears on the tab.")],
    cell_range: CellsRangeArg,
    shift: InsertShiftArg = "down",
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                insert_cells(src, dst, sheet=sheet_name, cell_range=cell_range, shift=shift)
            except (SafeXLError, TornReference) as exc:
                raise _insert_cells_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "sheet": sheet_name,
            "cell_range": cell_range,
            "shift": shift,
            "parts": parts,
        }

    return _run_http(job, blob_field="file_base64")


def http_rename_sheet(
    file_base64: B64Arg,
    old_name: Annotated[str, Field(description="Worksheet name to rename, exactly as it appears on the tab.")],
    new_name: Annotated[
        str, Field(description="New worksheet name (max 31 chars; may not contain [ ] * ? / \\ : or start/end with ').")
    ],
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                rename_sheet(src, dst, old_name, new_name)
            except SafeXLError as exc:
                raise _structural_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "old_name": old_name,
            "new_name": new_name,
            "parts": parts,
        }

    return _run_http(job, blob_field="file_base64")


def http_add_sheet(
    file_base64: B64Arg,
    sheet_name: Annotated[
        str, Field(description="Name for the new, empty worksheet (max 31 chars; may not contain [ ] * ? / \\ :).")
    ],
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                add_sheet(src, dst, sheet_name)
            except SafeXLError as exc:
                raise _structural_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "sheet": sheet_name,
            "parts": parts,
        }

    return _run_http(job, blob_field="file_base64")


def http_diff(
    before_base64: Annotated[str, Field(description="The original workbook, base64-encoded (max 20 MB decoded).")],
    after_base64: Annotated[
        str, Field(description="The changed workbook (e.g. safexl_edit's file_base64), base64-encoded (max 20 MB).")
    ],
    max_cells: Annotated[
        int, Field(ge=1, le=5000, description="Maximum cell entries to list across all sheets; counts are always exact.")
    ] = DEFAULT_DIFF_CELLS,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw_a = _decode_upload(before_base64, "before_base64")
        raw_b = _decode_upload(after_base64, "after_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            result = diff_workbooks(_stage(Path(tmp), raw_a, "before"), _stage(Path(tmp), raw_b, "after")).to_dict()
        result.pop("path_before", None)
        result.pop("path_after", None)
        return _shape_diff(result, max_cells)

    return _run_http(job)


def http_verify(
    before_base64: Annotated[str, Field(description="The original workbook, base64-encoded (max 20 MB decoded).")],
    after_base64: Annotated[str, Field(description="The edited workbook, base64-encoded (max 20 MB decoded).")],
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw_a = _decode_upload(before_base64, "before_base64")
        raw_b = _decode_upload(after_base64, "after_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            return _bounded_receipt(verify(_stage(Path(tmp), raw_a, "before"), _stage(Path(tmp), raw_b, "after")))

    return _run_http(job)


def http_view(
    file_base64: B64Arg,
    sheet: SheetArg = None,
    cell_range: CellRangeArg = None,
    find: FindArg = None,
    max_cells: MaxViewCellsArg = DEFAULT_VIEW_CELLS,
    formulas: FormulasArg = True,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "upload")
            try:
                return view(src, sheet, cell_range, find, max_cells, formulas)
            except SafeXLError as exc:
                raise _view_error(exc, src) from exc

    return _run_http(job)


def http_salvage(
    original_base64: Annotated[
        str, Field(description="The ORIGINAL workbook from before the damaging re-save, base64-encoded (max 20 MB).")
    ],
    damaged_base64: Annotated[
        str, Field(description="The workbook another library re-saved and damaged, base64-encoded (max 20 MB).")
    ],
    force: ForceArg = False,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw_a = _decode_upload(original_base64, "original_base64")
        raw_b = _decode_upload(damaged_base64, "damaged_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw_a, "original")
            dmg = _stage(Path(tmp), raw_b, "damaged")
            dst = src.with_name(f"salvaged{src.suffix}")
            try:
                result = _salvage_result(salvage(src, dmg, dst, force=force))
            except SafeXLError as exc:
                raise ToolInputError("salvage_failed", str(exc)) from exc
            if result["output_written"]:
                out = dst.read_bytes()
                result["file_base64"] = base64.b64encode(out).decode("ascii")
                result["file_extension"] = src.suffix
                result["size_bytes"] = len(out)
        return result

    return _run_http(job, blob_field="file_base64")


def http_resize_table(
    file_base64: B64Arg,
    table: TableArg,
    ref: TableRefArg,
    sheet_name: TableSheetArg = None,
) -> CallToolResult:
    def job() -> dict[str, Any]:
        raw = _decode_upload(file_base64, "file_base64")
        with tempfile.TemporaryDirectory(prefix=_TMP_PREFIX) as tmp:
            src = _stage(Path(tmp), raw, "source")
            dst = src.with_name(f"edited{src.suffix}")
            try:
                resize_table(src, dst, table=table, ref=ref, sheet=sheet_name)
            except SafeXLError as exc:
                raise _resize_error(exc, src) from exc
            out = dst.read_bytes()
            parts = _parts_summary(_part_hashes(src), _part_hashes(dst))
        return {
            "file_base64": base64.b64encode(out).decode("ascii"),
            "file_extension": src.suffix,
            "size_bytes": len(out),
            "table": table,
            "ref": ref,
            "parts": parts,
        }

    return _run_http(job, blob_field="file_base64")


# ---------------------------------------------------------------------------
# Tool descriptions (agents choose tools by these -- keep them concrete)
# ---------------------------------------------------------------------------

_WHEN = (
    "Use SafeXL for EXISTING .xlsx/.xlsm workbooks whose formulas, charts, pivot tables, conditional "
    "formatting, data validation, images, or macros must survive the edit -- instead of openpyxl, pandas, or "
    "Excel tools that load and re-save the whole file and silently drop what they don't model."
)

_DESCRIPTIONS = {
    "inspect": (
        "Read-only report on an existing Excel workbook (.xlsx/.xlsm); call it before editing. Returns each "
        "sheet's name, declared vs actual used range and formula count; defined names; and a feature inventory "
        "(charts, drawings/images, pivot tables, data validation, conditional formatting, and whether a VBA macro "
        "project is present, with a static keyword flag for AutoOpen/Shell/CreateObject-style patterns when the "
        "optional oletools extra is installed -- a heads-up, not a security audit). Never modifies the file. "
        "Macros are never executed."
    ),
    "edit": (
        "Write cell values and/or formulas into an EXISTING Excel workbook (.xlsx/.xlsm) without breaking it. "
        + _WHEN
        + " SafeXL rewrites only the edited sheets' XML plus a recalculate-on-open flag and copies every other "
        "part of the file byte-for-byte; the result's 'parts' field shows exactly which parts changed. Macros "
        "(VBA) are preserved and never executed. Formula values are not computed by SafeXL: they compute when "
        "the file is next opened in Excel or LibreOffice, so a library reading the file before that sees formula "
        "cells as empty. The output is structurally validated before it is written; on any failure nothing is "
        "written and a structured error is returned. Not for creating new workbooks, adding sheets, styling, or "
        ".xls/.xlsb/.csv files."
    ),
    "diff": (
        "Compare two versions of an Excel workbook (.xlsx/.xlsm) -- typically the original and safexl_edit's "
        "output -- and report exactly what changed: per-sheet cell changes (value/formula before -> after, "
        "added, removed) plus every internal file part added, removed, or changed by SHA-256. Use it after an "
        "edit to confirm only the intended cells changed and that charts, pivot tables, and macros are "
        "byte-identical. Read-only. Macros are never executed."
    ),
    "insert_rows": (
        "Insert blank rows into an EXISTING Excel worksheet (.xlsx/.xlsm) without breaking it: existing rows "
        "at or below the insertion point shift down, and formula references that point at or below the "
        "insertion point are rewritten to follow -- both within the sheet and cross-sheet (e.g. 'Sheet1!A5'), "
        "and merged cell ranges that span the insertion point are expanded. "
        + _WHEN
        + " Not for inserting or removing columns. The output is structurally validated before it is written; "
        "on any failure nothing is written and a structured error is returned."
    ),
    "insert_cells": (
        "Use this instead of safexl_insert_rows when the sheet has other lists or tables side by side that must "
        "not move: it inserts blank cells over a range (e.g. 'B5:C6') in an EXISTING workbook (.xlsx/.xlsm) and "
        "shifts only those columns down, like Excel's 'Insert cells, shift cells down'. Formulas, defined names, "
        "validation, conditional formatting, charts, tables, merges, hyperlinks, comments and pivots follow. If a "
        "reference, merge, table or array formula would be torn apart, nothing is written and error "
        "reference_would_tear names it. Only shift='down' is supported."
    ),
    "rename_sheet": (
        "Rename a worksheet in an EXISTING Excel workbook (.xlsx/.xlsm) and update every formula reference to "
        "it (both quoted \"'Old Name'!A1\" and unquoted 'OldName!A1' forms) across every sheet, so formulas "
        "keep pointing at the right sheet. The worksheet's underlying file inside the workbook is unchanged -- "
        "only its displayed name and references to it are updated. Fails if new_name is already used by "
        "another sheet."
    ),
    "verify": (
        "Use this to check that an edited workbook lost nothing: compares the edited .xlsx/.xlsm against its "
        "original and returns a receipt with a verdict -- intact (only intended parts changed), rewritten (the "
        "file was rebuilt by another writer), damaged (features such as charts, pivots, validation or macros "
        "went missing; lost_features lists them) or invalid (not a readable workbook) -- plus which internal "
        "parts changed. The receipt also has a one-sentence summary, added_risk (risky things the edit added, "
        "such as external links, hidden sheets or a VBA project) and cell_changes (changed-cell and inserted-row "
        "counts per sheet, capped at 50 sheets). Intact covers tracked features and untouched parts only: it "
        "does not mean no cell changed, so read summary and added_risk too. Read-only. Macros are never executed."
    ),
    "view": (
        "Use this to read cells from an EXISTING Excel workbook (.xlsx/.xlsm) in a compact, bounded form instead "
        "of dumping the whole sheet: a window of cells (cell_range) or the cells matching a text search (find), "
        "each with its type, value and formula. Output is capped by max_cells (hard ceiling "
        f"{MAX_VIEW_CELLS}); when truncated, next_range tells you the next window to ask for. Shared and array "
        "formulas are marked, never expanded or evaluated. Read-only."
    ),
    "salvage": (
        "Use this when a workbook was re-saved by openpyxl, pandas or another library and lost features: "
        "rebuild it from the original plus the intended edits. SafeXL finds the cell edits that were made in "
        "the damaged file and re-applies them to the original, so charts, pivot tables, validation and macros "
        "survive. status is salvaged (every difference explained), partial (some edits skipped; see skipped), "
        "unsupported (an unhandled structural change; nothing written) or nothing_to_do (nothing written). "
        "The output is only produced for salvaged and partial. nothing_to_do means the file lost nothing: keep it. "
        "Set force=true only to rebuild from the original even though the file looks intact."
    ),
    "resize_table": (
        "Use this to change the range of an Excel table (ListObject) in an EXISTING workbook (.xlsx/.xlsm), "
        "e.g. to grow it to 'A1:D12' after adding a column or rows next to it. Only the table definition "
        "changes; every other part of the file stays byte-identical. The top-left cell must stay put. Fails "
        "with table_not_found if no table has that name. Not for creating or deleting tables."
    ),
    "add_sheet": (
        "Add a new, empty worksheet to an EXISTING Excel workbook (.xlsx/.xlsm): registers it in the "
        "workbook's sheet list, relationships, and content types so it shows up as a normal tab. The new sheet "
        "has no cells -- call safexl_edit afterward to write values or formulas into it."
    ),
}

_INSTRUCTIONS = (
    "SafeXL edits existing Excel workbooks (.xlsx/.xlsm) without breaking them. Flow: safexl_inspect (sheets, "
    "formulas, charts, pivots, macros) -> safexl_edit (write values/formulas) -> safexl_diff (confirm only the "
    "intended cells changed; safexl_verify checks nothing was lost). If a workbook was already damaged by a "
    "re-save, use safexl_salvage. " + _WHEN + " Macros are never executed. Formulas written by SafeXL compute when "
    "the file is next opened in Excel or LibreOffice."
)

_HTTP_NOTE = (
    " This server is in HTTP mode: pass workbooks as base64 content (max 20 MB each); file paths and URLs are "
    "rejected, and safexl_edit returns the edited workbook as base64 in file_base64."
)


def build_server(mode: Literal["stdio", "http"] = "stdio") -> MCPServer:
    """Build the SafeXL MCP server. ``mode`` picks the file-intake contract:
    ``"stdio"`` tools take local paths; ``"http"`` tools take base64 content only."""
    if mode not in ("stdio", "http"):
        raise ValueError(f"unknown mode: {mode!r}")
    http = mode == "http"
    server = MCPServer(
        "safexl",
        title="SafeXL",
        instructions=_INSTRUCTIONS + (_HTTP_NOTE if http else ""),
        website_url="https://github.com/herakles-dev/safexl",
        version=__version__,
    )
    read_only = ToolAnnotations(read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=False)
    if http:
        # Nothing on the server changes: the edited workbook is returned, not stored.
        edit_hints = read_only
    else:
        edit_hints = ToolAnnotations(
            read_only_hint=False, destructive_hint=True, idempotent_hint=True, open_world_hint=False
        )

    tools = (
        ("safexl_inspect", "Inspect workbook", "inspect", http_inspect if http else stdio_inspect, read_only),
        ("safexl_edit", "Edit workbook safely", "edit", http_edit if http else stdio_edit, edit_hints),
        ("safexl_diff", "Diff two workbooks", "diff", http_diff if http else stdio_diff, read_only),
        (
            "safexl_insert_rows",
            "Insert rows",
            "insert_rows",
            http_insert_rows if http else stdio_insert_rows,
            edit_hints,
        ),
        (
            "safexl_insert_cells",
            "Insert cells",
            "insert_cells",
            http_insert_cells if http else stdio_insert_cells,
            edit_hints,
        ),
        (
            "safexl_rename_sheet",
            "Rename sheet",
            "rename_sheet",
            http_rename_sheet if http else stdio_rename_sheet,
            edit_hints,
        ),
        ("safexl_add_sheet", "Add sheet", "add_sheet", http_add_sheet if http else stdio_add_sheet, edit_hints),
        ("safexl_verify", "Verify an edit", "verify", http_verify if http else stdio_verify, read_only),
        ("safexl_view", "View cells", "view", http_view if http else stdio_view, read_only),
        ("safexl_salvage", "Salvage a damaged workbook", "salvage", http_salvage if http else stdio_salvage, edit_hints),
        (
            "safexl_resize_table",
            "Resize table",
            "resize_table",
            http_resize_table if http else stdio_resize_table,
            edit_hints,
        ),
    )
    for name, title, key, fn, hints in tools:
        server.add_tool(fn, name=name, title=title, description=_DESCRIPTIONS[key], annotations=hints)
    return server


# ---------------------------------------------------------------------------
# Console entry point
# ---------------------------------------------------------------------------


def _port(raw: str) -> int:
    try:
        n = int(raw)
    except ValueError:
        raise argparse.ArgumentTypeError(f"not a port number: {raw!r}") from None
    if not 1 <= n <= 65535:
        raise argparse.ArgumentTypeError(f"port out of range: {n}")
    return n


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="safexl-mcp",
        description="SafeXL MCP server: inspect, edit, and diff existing Excel workbooks without breaking them.",
    )
    parser.add_argument(
        "--http",
        action="store_true",
        help="serve streamable HTTP instead of stdio (tools accept base64 file content only; no server-side paths)",
    )
    parser.add_argument("--host", default=None, help="HTTP bind address (default 127.0.0.1)")
    parser.add_argument("--port", type=_port, default=None, help="HTTP port (default 8000)")
    parser.add_argument("--version", action="version", version=f"safexl-mcp {__version__}")
    args = parser.parse_args(argv)

    if not args.http:
        if args.host is not None or args.port is not None:
            parser.error("--host/--port only apply with --http")
        build_server("stdio").run()
        return 0

    host = args.host or "127.0.0.1"
    port = args.port or 8000
    if host not in ("127.0.0.1", "localhost", "::1"):
        print(
            f"safexl-mcp: WARNING binding to {host}: this server has no authentication. Put it behind an "
            "authenticating reverse proxy before exposing it beyond this machine.",
            file=sys.stderr,
        )
    build_server("http").run(
        "streamable-http",
        host=host,
        port=port,
        stateless_http=True,
        json_response=True,
        max_request_body_size=HTTP_MAX_REQUEST_BODY,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
