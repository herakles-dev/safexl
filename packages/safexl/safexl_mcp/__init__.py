"""SafeXL — edit Excel workbooks in headless environments without breaking them.

Public API:

    from safexl_mcp import edit, SafeXLError

    edit(
        "in.xlsx",
        "out.xlsx",
        [
            {"sheet": "Sheet1", "cell": "A1", "value": 42},
            {"sheet": "Sheet1", "cell": "B1", "formula": "A1*2"},
        ],
    )

See ``safexl_mcp.engine`` for the full docstring covering the direct-XML-injection
technique, the formula-caching caveat, and the rollback promise. See
``safexl_mcp.inspect`` and ``safexl_mcp.diff`` for the read-only, never-executes-
macros inspection and diff tools.
"""
from .diff import SheetDiff, WorkbookDiff, diff_workbooks
from .engine import Change, SafeXLError, edit
from .inspect import FeatureInventory, SheetInspection, VbaInfo, WorkbookInspection, inspect_workbook
from .structural import add_sheet, insert_rows, rename_sheet
from .validate import FileReport, Finding, validate_bytes_ok, validate_workbook

__version__ = "0.2.0"

__all__ = [
    "edit",
    "SafeXLError",
    "Change",
    "insert_rows",
    "rename_sheet",
    "add_sheet",
    "validate_workbook",
    "validate_bytes_ok",
    "FileReport",
    "Finding",
    "inspect_workbook",
    "WorkbookInspection",
    "SheetInspection",
    "FeatureInventory",
    "VbaInfo",
    "diff_workbooks",
    "WorkbookDiff",
    "SheetDiff",
    "__version__",
]
