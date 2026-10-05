"""SafeXL read-only workbook inspector.

Reports the structure of a .xlsx/.xlsm without ever writing to it and without
ever executing a macro. Two things this module answers:

1. Per-sheet: declared ``<dimension ref="...">`` vs the actual used range
   computed by scanning ``<sheetData>`` (the same "dimension can lie, Excel
   trusts the real cells" fact ``safexl_mcp.engine._update_dimension`` and
   ``safexl_mcp.validate.check_dimension`` are built around), formula count,
   and per-sheet data-validation / conditional-formatting counts (native and
   extLst/x14).
2. Workbook-wide: a feature inventory (charts, drawings/images, pivot
   tables+caches, data validation, conditional formatting, VBA) built by
   walking the OOXML zip parts
   directly, not by loading the workbook into any Excel-editing library.

VBA handling
------------
``xl/vbaProject.bin`` (if present) is reported by presence and byte size
only in stdlib-only mode. The VBA source stored inside that OLE compound
file is MS-OVBA-compressed, so nothing about its *content* -- module names,
AutoOpen hooks, Shell calls -- can be determined without a real OLE/VBA
parser; stdlib's ``zipfile``/``xml.etree`` cannot read compound-file storage.

If the optional ``oletools`` package (the ``scan`` extra) is
importable, its ``olevba`` static analyzer scans the extracted VBA project
for known-suspicious keywords (AutoOpen/Auto_Open/Workbook_Open, Shell,
CreateObject, URLDownloadToFile, WScript, PowerShell) and reports hits.
olevba performs **static** analysis only (it can decode obfuscated string
literals as text) -- it never executes a single line of VBA, matching the
same never-execute guarantee ``safexl_mcp.engine`` makes for edits.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .engine import SafeXLError

__all__ = [
    "SheetInspection",
    "VbaInfo",
    "FeatureInventory",
    "WorkbookInspection",
    "inspect_workbook",
]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
_CHART_RE = re.compile(r"^xl/charts/chart\d+\.xml$")
_DRAWING_RE = re.compile(r"^xl/drawings/drawing\d+\.xml$")
_PIVOT_TABLE_RE = re.compile(r"^xl/pivotTables/pivotTable\d+\.xml$")
_PIVOT_CACHE_RE = re.compile(r"^xl/pivotCache/pivotCacheDefinition\d+\.xml$")
_MEDIA_PREFIX = "xl/media/"
_VBA_PART = "xl/vbaProject.bin"

_SUSPICIOUS_KEYWORDS = (
    "autoopen",
    "auto_open",
    "workbook_open",
    "shell",
    "createobject",
    "urldownloadtofile",
    "wscript",
    "powershell",
)


# ---------------------------------------------------------------------------
# Cell-reference helpers (local copies -- see engine.py / validate.py for why
# each module keeps its own: no cross-module coupling for a two-line helper)
# ---------------------------------------------------------------------------


def _split_ref(ref: str) -> tuple[str, int]:
    m = _CELL_REF_RE.match(ref)
    if not m:
        raise ValueError(f"not a cell ref: {ref!r}")
    return m.group(1).upper(), int(m.group(2))


def _col_to_num(col_letters: str) -> int:
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col(n: int) -> str:
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class SheetInspection:
    name: str
    sheet_id: str
    declared_dimension: str | None
    actual_used_range: str | None
    dimension_matches: bool
    formula_count: int
    data_validation_native: int
    data_validation_extlst: int
    conditional_formatting_native: int
    conditional_formatting_extlst: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "sheet_id": self.sheet_id,
            "declared_dimension": self.declared_dimension,
            "actual_used_range": self.actual_used_range,
            "dimension_matches": self.dimension_matches,
            "formula_count": self.formula_count,
            "data_validation_native": self.data_validation_native,
            "data_validation_extlst": self.data_validation_extlst,
            "conditional_formatting_native": self.conditional_formatting_native,
            "conditional_formatting_extlst": self.conditional_formatting_extlst,
        }


@dataclass
class VbaInfo:
    present: bool
    size_bytes: int | None
    macro_scan: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return {
            "present": self.present,
            "size_bytes": self.size_bytes,
            "macro_scan": self.macro_scan,
        }


@dataclass
class FeatureInventory:
    charts: int
    drawings: int
    media_files: int
    pivot_tables: int
    pivot_caches: int
    data_validation_native: int
    data_validation_extlst: int
    conditional_formatting_native: int
    conditional_formatting_extlst: int
    vba: VbaInfo

    def to_dict(self) -> dict[str, Any]:
        return {
            "charts": self.charts,
            "drawings": self.drawings,
            "media_files": self.media_files,
            "pivot_tables": self.pivot_tables,
            "pivot_caches": self.pivot_caches,
            "data_validation_native": self.data_validation_native,
            "data_validation_extlst": self.data_validation_extlst,
            "conditional_formatting_native": self.conditional_formatting_native,
            "conditional_formatting_extlst": self.conditional_formatting_extlst,
            "vba": self.vba.to_dict(),
        }


@dataclass
class WorkbookInspection:
    path: str
    sheets: list[SheetInspection] = field(default_factory=list)
    defined_names: list[dict[str, Any]] = field(default_factory=list)
    features: FeatureInventory | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "sheets": [s.to_dict() for s in self.sheets],
            "defined_names": self.defined_names,
            "features": self.features.to_dict() if self.features is not None else None,
        }


# ---------------------------------------------------------------------------
# workbook.xml / rels parsing
# ---------------------------------------------------------------------------


def _sheet_list(zf: zipfile.ZipFile, wb_root: ET.Element) -> list[dict[str, str]]:
    """[{"name":, "sheet_id":, "rid":, "target": zip path of the sheet XML}], workbook order."""
    try:
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    except KeyError:
        rels_root = None

    rid_to_target: dict[str, str] = {}
    if rels_root is not None:
        for rel in rels_root:
            rid = rel.get("Id")
            target = rel.get("Target", "").lstrip("/")
            if target and not target.startswith("xl/"):
                target = "xl/" + target
            if rid:
                rid_to_target[rid] = target

    sheets: list[dict[str, str]] = []
    for sheet_el in wb_root.iter(f"{{{NS_MAIN}}}sheet"):
        rid = sheet_el.get(f"{{{R_NS}}}id")
        sheets.append(
            {
                "name": sheet_el.get("name", ""),
                "sheet_id": sheet_el.get("sheetId", ""),
                "rid": rid or "",
                "target": rid_to_target.get(rid, ""),
            }
        )
    return sheets


def _defined_names(wb_root: ET.Element) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    container = wb_root.find(f"{{{NS_MAIN}}}definedNames")
    if container is None:
        return out
    for dn in container.findall(f"{{{NS_MAIN}}}definedName"):
        out.append(
            {
                "name": dn.get("name", ""),
                "value": dn.text or "",
                "local_sheet_id": dn.get("localSheetId"),
            }
        )
    return out


# ---------------------------------------------------------------------------
# Per-sheet structural facts
# ---------------------------------------------------------------------------


def _declared_dimension(sheet_root: ET.Element) -> str | None:
    dim = sheet_root.find(f"{{{NS_MAIN}}}dimension")
    return dim.get("ref") if dim is not None else None


def _actual_used_range(sheet_root: ET.Element) -> str | None:
    sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        return None

    min_col = max_col = min_row = max_row = None
    for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
        for c in row.findall(f"{{{NS_MAIN}}}c"):
            ref = c.get("r")
            if not ref:
                continue
            try:
                letters, num = _split_ref(ref)
            except ValueError:
                continue
            col = _col_to_num(letters)
            min_col = col if min_col is None else min(min_col, col)
            max_col = col if max_col is None else max(max_col, col)
            min_row = num if min_row is None else min(min_row, num)
            max_row = num if max_row is None else max(max_row, num)

    if min_col is None:
        return None
    if min_col == max_col and min_row == max_row:
        return f"{_num_to_col(min_col)}{min_row}"
    return f"{_num_to_col(min_col)}{min_row}:{_num_to_col(max_col)}{max_row}"


def _formula_count(sheet_root: ET.Element) -> int:
    return sum(1 for c in sheet_root.iter(f"{{{NS_MAIN}}}c") if c.find(f"{{{NS_MAIN}}}f") is not None)


def _data_validation_counts(sheet_root: ET.Element) -> dict[str, int]:
    native = sum(1 for _ in sheet_root.iter(f"{{{NS_MAIN}}}dataValidation"))
    extlst = sum(1 for _ in sheet_root.iter(f"{{{X14_NS}}}dataValidation"))
    return {"native": native, "extlst": extlst}


def _conditional_formatting_counts(sheet_root: ET.Element) -> dict[str, int]:
    native = sum(1 for _ in sheet_root.iter(f"{{{NS_MAIN}}}conditionalFormatting"))
    extlst = sum(1 for _ in sheet_root.iter(f"{{{X14_NS}}}conditionalFormatting"))
    return {"native": native, "extlst": extlst}


# ---------------------------------------------------------------------------
# Workbook-wide feature counts (by OOXML zip part, walking the zip directly)
# ---------------------------------------------------------------------------


def _count_parts(names: list[str], pattern: re.Pattern[str]) -> int:
    return sum(1 for n in names if pattern.match(n))


def _macro_scan(vba_bytes: bytes) -> dict[str, Any]:
    """Static keyword scan of vbaProject.bin via oletools' olevba, if installed.

    NEVER executes VBA: olevba's analyze_macros() is purely static (it can
    decode obfuscated string literals -- Hex/Base64/Dridex-style -- as text,
    but does not run any macro code). Falls back to a presence/size-only
    report if oletools is not importable, per the module docstring.
    """
    try:
        from oletools.olevba import VBA_Parser  # type: ignore[import-not-found]
    except ImportError:
        return {
            "mode": "stdlib-only",
            "note": "macro scan unavailable (install the scan extra: oletools)",
            "hits": [],
        }

    hits: list[dict[str, str]] = []
    try:
        parser = VBA_Parser("vbaProject.bin", data=vba_bytes)
        try:
            for entry in parser.analyze_macros():
                # olevba's analyze_macros() yields (kw_type, keyword, description).
                kw_type, keyword, description = entry[0], entry[1], entry[2]
                if any(s in keyword.lower() for s in _SUSPICIOUS_KEYWORDS):
                    hits.append({"type": kw_type, "keyword": keyword, "description": description})
        finally:
            parser.close()
    except Exception as exc:  # noqa: BLE001 - a scan failure must not break inspect()
        return {"mode": "oletools", "note": f"macro scan failed: {exc}", "hits": []}

    return {"mode": "oletools", "note": "", "hits": hits}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def inspect_workbook(path: str | Path) -> WorkbookInspection:
    """Read-only structural + feature report for ``path``. Never writes to
    ``path``, never executes any macro it contains.

    Raises
    ------
    SafeXLError if ``path`` doesn't exist or isn't a well-formed OOXML
    package.
    """
    path = Path(path)
    if not path.is_file():
        raise SafeXLError(f"file not found: {path}")

    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = zf.namelist()

            try:
                wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            except KeyError as exc:
                raise SafeXLError(f"{path}: not a valid OOXML workbook (missing xl/workbook.xml)") from exc

            sheet_meta = _sheet_list(zf, wb_root)
            defined_names = _defined_names(wb_root)

            sheets: list[SheetInspection] = []
            for meta in sheet_meta:
                target = meta["target"]
                if not target or target not in names:
                    sheets.append(
                        SheetInspection(
                            name=meta["name"],
                            sheet_id=meta["sheet_id"],
                            declared_dimension=None,
                            actual_used_range=None,
                            dimension_matches=True,
                            formula_count=0,
                            data_validation_native=0,
                            data_validation_extlst=0,
                            conditional_formatting_native=0,
                            conditional_formatting_extlst=0,
                        )
                    )
                    continue

                sheet_root = ET.fromstring(zf.read(target))
                declared = _declared_dimension(sheet_root)
                actual = _actual_used_range(sheet_root)
                dv = _data_validation_counts(sheet_root)
                cf = _conditional_formatting_counts(sheet_root)
                sheets.append(
                    SheetInspection(
                        name=meta["name"],
                        sheet_id=meta["sheet_id"],
                        declared_dimension=declared,
                        actual_used_range=actual,
                        dimension_matches=True if actual is None else declared == actual,
                        formula_count=_formula_count(sheet_root),
                        data_validation_native=dv["native"],
                        data_validation_extlst=dv["extlst"],
                        conditional_formatting_native=cf["native"],
                        conditional_formatting_extlst=cf["extlst"],
                    )
                )

            charts = _count_parts(names, _CHART_RE)
            drawings = _count_parts(names, _DRAWING_RE)
            media_files = sum(1 for n in names if n.startswith(_MEDIA_PREFIX) and not n.endswith("/"))
            pivot_tables = _count_parts(names, _PIVOT_TABLE_RE)
            pivot_caches = _count_parts(names, _PIVOT_CACHE_RE)

            vba_present = _VBA_PART in names
            if vba_present:
                vba_bytes = zf.read(_VBA_PART)
                vba_size: int | None = len(vba_bytes)
                macro_scan = _macro_scan(vba_bytes)
            else:
                vba_size = None
                macro_scan = {"mode": "n/a", "note": "no VBA project present", "hits": []}

            features = FeatureInventory(
                charts=charts,
                drawings=drawings,
                media_files=media_files,
                pivot_tables=pivot_tables,
                pivot_caches=pivot_caches,
                data_validation_native=sum(s.data_validation_native for s in sheets),
                data_validation_extlst=sum(s.data_validation_extlst for s in sheets),
                conditional_formatting_native=sum(s.conditional_formatting_native for s in sheets),
                conditional_formatting_extlst=sum(s.conditional_formatting_extlst for s in sheets),
                vba=VbaInfo(present=vba_present, size_bytes=vba_size, macro_scan=macro_scan),
            )
    except SafeXLError:
        raise
    except zipfile.BadZipFile as exc:
        raise SafeXLError(f"{path}: not a valid zip/OOXML package: {exc}") from exc
    except ET.ParseError as exc:
        raise SafeXLError(f"{path}: malformed XML in workbook: {exc}") from exc

    return WorkbookInspection(path=str(path), sheets=sheets, defined_names=defined_names, features=features)
