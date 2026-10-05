"""SafeXL read-only before/after diff for two .xlsx/.xlsm workbooks.

Two questions this module answers, both driven by the "nothing else changed"
promise documented in ``safexl_mcp.engine``'s module docstring:

1. **Per sheet, what cell content changed?** Value/formula before -> after,
   plus cells added (present only in B) and removed (present only in A).
   Comparison is on resolved semantic content -- a shared-string index and an
   inline string that display the same text are treated as equal -- not on
   raw ``<c>`` XML, so the diff reflects what a person or downstream reader
   actually sees.
2. **What OOXML zip parts changed, by hash?** This is the machine-checkable
   half of "every part I didn't touch is byte-identical to the source"
   (``safexl_mcp.engine``'s core claim): added/removed/changed parts, with a
   sha256 before/after for every changed part.

Never writes to either input file. Never executes any macro either file
might contain (this module doesn't even look at ``xl/vbaProject.bin``'s
content -- see ``safexl_mcp.inspect`` for that).
"""
from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .engine import SafeXLError

__all__ = ["SheetDiff", "WorkbookDiff", "diff_workbooks"]

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


# ---------------------------------------------------------------------------
# Cell-reference helpers (local copy -- see inspect.py/engine.py/validate.py
# for why each module keeps its own rather than sharing a private import)
# ---------------------------------------------------------------------------


def _split_ref(ref: str) -> tuple[str, int]:
    m = _CELL_REF_RE.match(ref)
    if not m:
        return "", 0
    return m.group(1).upper(), int(m.group(2))


def _col_to_num(col_letters: str) -> int:
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _ref_sort_key(ref: str) -> tuple[int, int]:
    letters, row = _split_ref(ref)
    return (row, _col_to_num(letters))


# ---------------------------------------------------------------------------
# workbook.xml / rels / sharedStrings parsing
# ---------------------------------------------------------------------------


def _sheet_target_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """sheet name -> zip path of its worksheet XML, workbook order preserved."""
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
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

    out: dict[str, str] = {}
    for sheet_el in wb_root.iter(f"{{{NS_MAIN}}}sheet"):
        name = sheet_el.get("name", "")
        rid = sheet_el.get(f"{{{R_NS}}}id")
        out[name] = rid_to_target.get(rid, "")
    return out


def _extract_text(el: ET.Element) -> str:
    """Text of an <is> or <si> element: a direct <t>, or concatenated <r><t> runs."""
    t = el.find(f"{{{NS_MAIN}}}t")
    if t is not None:
        return t.text or ""
    parts = []
    for r in el.findall(f"{{{NS_MAIN}}}r"):
        rt = r.find(f"{{{NS_MAIN}}}t")
        if rt is not None and rt.text:
            parts.append(rt.text)
    return "".join(parts)


def _read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    return [_extract_text(si) for si in root.findall(f"{{{NS_MAIN}}}si")]


def _coerce_number(raw: str) -> int | float | str:
    try:
        return int(raw)
    except ValueError:
        pass
    try:
        return float(raw)
    except ValueError:
        return raw


# ---------------------------------------------------------------------------
# Per-cell semantic content
# ---------------------------------------------------------------------------


def _cell_repr(c_el: ET.Element, shared_strings: list[str]) -> dict[str, Any]:
    f_el = c_el.find(f"{{{NS_MAIN}}}f")
    if f_el is not None:
        if f_el.text:
            return {"kind": "formula", "formula": f_el.text}
        si = f_el.get("si")
        return {"kind": "formula", "formula": f"<shared formula, master si={si}>" if si else ""}

    t = c_el.get("t")
    v_el = c_el.find(f"{{{NS_MAIN}}}v")

    if t == "inlineStr":
        is_el = c_el.find(f"{{{NS_MAIN}}}is")
        return {"kind": "value", "value": _extract_text(is_el) if is_el is not None else ""}
    if t == "s":
        if v_el is None or v_el.text is None:
            return {"kind": "empty"}
        idx = int(v_el.text)
        value = shared_strings[idx] if 0 <= idx < len(shared_strings) else f"<sharedString#{idx} out of range>"
        return {"kind": "value", "value": value}
    if t == "b":
        return {"kind": "value", "value": v_el is not None and v_el.text == "1"}
    if t == "str":
        return {"kind": "value", "value": v_el.text if v_el is not None else ""}
    if t == "e":
        return {"kind": "value", "value": v_el.text if v_el is not None else "", "error": True}

    if v_el is None or v_el.text is None:
        return {"kind": "empty"}
    return {"kind": "value", "value": _coerce_number(v_el.text)}


def _export_cell(rep: dict[str, Any]) -> dict[str, Any]:
    if rep["kind"] == "formula":
        return {"formula": rep["formula"]}
    if rep["kind"] == "empty":
        return {"empty": True}
    out: dict[str, Any] = {"value": rep["value"]}
    if rep.get("error"):
        out["error"] = True
    return out


def _sheet_cells(zf: zipfile.ZipFile, target: str, shared_strings: list[str]) -> dict[str, dict[str, Any]]:
    if not target:
        return {}
    try:
        root = ET.fromstring(zf.read(target))
    except KeyError:
        return {}
    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        return {}

    out: dict[str, dict[str, Any]] = {}
    for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
        for c in row.findall(f"{{{NS_MAIN}}}c"):
            ref = c.get("r")
            if not ref:
                continue
            out[ref] = _cell_repr(c, shared_strings)
    return out


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class SheetDiff:
    sheet: str
    changed_cells: list[dict[str, Any]] = field(default_factory=list)
    added_cells: list[dict[str, Any]] = field(default_factory=list)
    removed_cells: list[dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet": self.sheet,
            "changed_cells": self.changed_cells,
            "added_cells": self.added_cells,
            "removed_cells": self.removed_cells,
        }


@dataclass
class WorkbookDiff:
    path_before: str
    path_after: str
    sheets: list[SheetDiff] = field(default_factory=list)
    sheets_added: list[str] = field(default_factory=list)
    sheets_removed: list[str] = field(default_factory=list)
    parts: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "path_before": self.path_before,
            "path_after": self.path_after,
            "sheets": [s.to_dict() for s in self.sheets],
            "sheets_added": self.sheets_added,
            "sheets_removed": self.sheets_removed,
            "parts": self.parts,
        }


# ---------------------------------------------------------------------------
# Diff assembly
# ---------------------------------------------------------------------------


def _diff_sheet(name: str, before: dict[str, dict[str, Any]], after: dict[str, dict[str, Any]]) -> SheetDiff | None:
    changed: list[dict[str, Any]] = []
    added: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []

    for ref in sorted(set(before) | set(after), key=_ref_sort_key):
        b = before.get(ref)
        a = after.get(ref)
        if b is None and a is not None:
            added.append({"cell": ref, "after": _export_cell(a)})
        elif b is not None and a is None:
            removed.append({"cell": ref, "before": _export_cell(b)})
        elif b is not None and a is not None and b != a:
            changed.append({"cell": ref, "before": _export_cell(b), "after": _export_cell(a)})

    if not changed and not added and not removed:
        return None
    return SheetDiff(sheet=name, changed_cells=changed, added_cells=added, removed_cells=removed)


def _parts_diff(zf_a: zipfile.ZipFile, zf_b: zipfile.ZipFile) -> dict[str, Any]:
    names_a = {n for n in zf_a.namelist() if not n.endswith("/")}
    names_b = {n for n in zf_b.namelist() if not n.endswith("/")}
    added = sorted(names_b - names_a)
    removed = sorted(names_a - names_b)
    common = sorted(names_a & names_b)

    changed: list[dict[str, str]] = []
    unchanged = 0
    for name in common:
        ha = hashlib.sha256(zf_a.read(name)).hexdigest()
        hb = hashlib.sha256(zf_b.read(name)).hexdigest()
        if ha != hb:
            changed.append({"path": name, "sha256_before": ha, "sha256_after": hb})
        else:
            unchanged += 1

    return {"added": added, "removed": removed, "changed": changed, "unchanged_count": unchanged}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def diff_workbooks(path_a: str | Path, path_b: str | Path) -> WorkbookDiff:
    """Diff two workbooks: per-sheet cell changes (by resolved semantic
    content, not raw XML) plus an OOXML-part-level diff by sha256. Read-only
    -- never writes to either file, never executes any macro.

    Raises
    ------
    SafeXLError if either path doesn't exist or isn't a well-formed OOXML
    package.
    """
    path_a = Path(path_a)
    path_b = Path(path_b)
    if not path_a.is_file():
        raise SafeXLError(f"file not found: {path_a}")
    if not path_b.is_file():
        raise SafeXLError(f"file not found: {path_b}")

    try:
        with zipfile.ZipFile(path_a, "r") as za, zipfile.ZipFile(path_b, "r") as zb:
            targets_a = _sheet_target_map(za)
            targets_b = _sheet_target_map(zb)
            shared_a = _read_shared_strings(za)
            shared_b = _read_shared_strings(zb)

            names_common = [n for n in targets_a if n in targets_b]
            sheets_added = [n for n in targets_b if n not in targets_a]
            sheets_removed = [n for n in targets_a if n not in targets_b]

            sheets: list[SheetDiff] = []
            for name in names_common:
                before = _sheet_cells(za, targets_a[name], shared_a)
                after = _sheet_cells(zb, targets_b[name], shared_b)
                sd = _diff_sheet(name, before, after)
                if sd is not None:
                    sheets.append(sd)

            parts = _parts_diff(za, zb)
    except SafeXLError:
        raise
    except zipfile.BadZipFile as exc:
        raise SafeXLError(f"not a valid zip/OOXML package: {exc}") from exc
    except ET.ParseError as exc:
        raise SafeXLError(f"malformed XML in workbook: {exc}") from exc

    return WorkbookDiff(
        path_before=str(path_a),
        path_after=str(path_b),
        sheets=sheets,
        sheets_added=sheets_added,
        sheets_removed=sheets_removed,
        parts=parts,
    )
