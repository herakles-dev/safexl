"""Structural repair-trigger validator for OOXML (.xlsx/.xlsm) packages.

Ported from an adversarial checker script into a reusable
library module. That check measured feature-category *preservation* via
zip-part/hash diffs; this module checks whether an OOXML package is actually
well-formed, internally consistent XML -- the class of defect that classically
triggers Excel's "We found a problem with some content... repair?" dialog.

``safexl_mcp.engine`` runs this validator against every edited output
*before* it is moved into place (see the module docstring there for the
write-to-temp / validate / move rollback promise). It can also be called
directly, e.g. from a CLI or test suite, against any .xlsx/.xlsm file.

Checks performed per workbook (independent of any "before" file -- these are
all properties a single OOXML package must satisfy on its own):

  1. well_formed_xml   -- every .xml/.rels part in the zip parses with
                           xml.etree.ElementTree.
  2. content_types      -- every Override PartName in [Content_Types].xml
                           actually exists in the zip (a dangling Override for
                           a part that was deleted, e.g. calcChain.xml after a
                           formula edit, is a classic self-inflicted repair
                           trigger); every part not covered by an Override has
                           a Default for its extension.
  3. calc_chain         -- if xl/calcChain.xml is present, every entry's cell
                           ref must correspond to a cell that actually
                           contains <f> on the sheet calcChain claims it's on.
                           If calcChain.xml is absent but the workbook has
                           formula cells, calcPr should carry
                           fullCalcOnLoad="1" so the opening application
                           rebuilds it instead of trusting nothing.
  4. shared_strings     -- xl/sharedStrings.xml's uniqueCount must equal its
                           actual <si> child count; every t="s" cell's string
                           index across every sheet must be within
                           [0, uniqueCount).
  5. row_cell_order     -- within each worksheet, <row r="N"> elements must
                           appear in strictly ascending, non-duplicate N
                           order, and within each row, <c r="..."> cells must
                           appear in strictly ascending, non-duplicate column
                           order.
  6. dimension          -- each worksheet's <dimension ref="..."/> must be
                           parseable as a cell or range ref, and must bound
                           every actual cell present in that sheet's
                           <sheetData>.

CRITICAL = well_formed_xml failure, dangling Content_Types Override, or a
stale calcChain entry (all three are documented repair triggers). WARNING =
everything else (dimension undershoot, missing fullCalcOnLoad fallback,
sharedStrings count drift) -- structural defects worth fixing but not
independently confirmed to produce a repair prompt on their own.
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

__all__ = [
    "Finding",
    "FileReport",
    "validate_workbook",
    "validate_bytes_ok",
]


@dataclass
class Finding:
    check: str
    severity: str  # "CRITICAL" | "WARNING"
    detail: str


@dataclass
class FileReport:
    path: str
    findings: list = field(default_factory=list)

    def add(self, check: str, severity: str, detail: str) -> None:
        self.findings.append(Finding(check, severity, detail))

    @property
    def critical_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == "CRITICAL")

    @property
    def warning_count(self) -> int:
        return sum(1 for f in self.findings if f.severity == "WARNING")

    @property
    def clean(self) -> bool:
        return not self.findings

    def critical_details(self) -> list[str]:
        return [f"[{f.check}] {f.detail}" for f in self.findings if f.severity == "CRITICAL"]


def _col_to_num(col_letters: str) -> int:
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"not a cell ref: {ref!r}")
    return m.group(1), int(m.group(2))


# ---------------------------------------------------------------------------
# 1. well-formed XML
# ---------------------------------------------------------------------------


def check_well_formed_xml(z: zipfile.ZipFile, report: FileReport) -> dict[str, ET.Element]:
    """Parse every .xml/.rels part; return {name: root} for parts that succeeded."""
    parsed: dict[str, ET.Element] = {}
    for name in z.namelist():
        if not name.endswith(".xml") and not name.endswith(".rels"):
            continue
        data = z.read(name)
        try:
            parsed[name] = ET.fromstring(data)
        except ET.ParseError as e:
            report.add("well_formed_xml", "CRITICAL", f"{name}: XML parse error: {e}")
    return parsed


# ---------------------------------------------------------------------------
# 2. [Content_Types].xml consistency
# ---------------------------------------------------------------------------


def check_content_types(z: zipfile.ZipFile, parsed: dict[str, ET.Element], report: FileReport) -> None:
    if "[Content_Types].xml" not in parsed:
        report.add("content_types", "CRITICAL", "[Content_Types].xml missing or unparseable")
        return
    root = parsed["[Content_Types].xml"]
    names = set(z.namelist())

    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "Default":
            defaults[child.get("Extension", "").lower()] = child.get("ContentType", "")
        elif tag == "Override":
            overrides[child.get("PartName", "")] = child.get("ContentType", "")

    for part_name in overrides:
        zip_path = part_name.lstrip("/")
        if zip_path not in names:
            report.add(
                "content_types",
                "CRITICAL",
                f"dangling Override for {part_name!r}: no such part in the zip "
                f"(classic pattern: a part was deleted but Content_Types still "
                f"declares it -- e.g. calcChain.xml)",
            )

    for name in sorted(names):
        if name.endswith("/"):
            continue
        part_name = "/" + name
        if part_name in overrides:
            continue
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext in defaults:
            continue
        report.add(
            "content_types",
            "WARNING",
            f"{name}: not covered by any Override, and extension {ext!r} has no Default",
        )


# ---------------------------------------------------------------------------
# 3. calcChain.xml handling
# ---------------------------------------------------------------------------


def _sheet_targets(parsed: dict[str, ET.Element]) -> dict[str, str]:
    rels_name = "xl/_rels/workbook.xml.rels"
    if rels_name not in parsed:
        return {}
    rid_to_target: dict[str, str] = {}
    for rel in parsed[rels_name]:
        rid = rel.get("Id")
        target = rel.get("Target", "").lstrip("/")
        if target and not target.startswith("xl/"):
            target = "xl/" + target
        if rid:
            rid_to_target[rid] = target
    return rid_to_target


def _sheet_id_to_target(parsed: dict[str, ET.Element]) -> dict[str, str]:
    wb_name = "xl/workbook.xml"
    if wb_name not in parsed:
        return {}
    rid_to_target = _sheet_targets(parsed)
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    out: dict[str, str] = {}
    for sheet_el in parsed[wb_name].iter():
        tag = sheet_el.tag.rsplit("}", 1)[-1]
        if tag != "sheet":
            continue
        sheet_id = sheet_el.get("sheetId")
        rid = sheet_el.get(f"{r_ns}id")
        target = rid_to_target.get(rid)
        if sheet_id and target:
            out[sheet_id] = target
    return out


def _formula_cell_refs(sheet_root: ET.Element) -> set[str]:
    refs: set[str] = set()
    for c in sheet_root.iter(f"{{{NS['main']}}}c"):
        if c.find(f"{{{NS['main']}}}f") is not None:
            ref = c.get("r")
            if ref:
                refs.add(ref)
    return refs


def check_calc_chain(z: zipfile.ZipFile, parsed: dict[str, ET.Element], report: FileReport) -> None:
    names = set(z.namelist())
    has_calc_chain = "xl/calcChain.xml" in names

    sheet_id_to_target = _sheet_id_to_target(parsed)
    formula_refs_by_target: dict[str, set[str]] = {}
    any_formulas = False
    for target in sheet_id_to_target.values():
        if target in parsed:
            refs = _formula_cell_refs(parsed[target])
            formula_refs_by_target[target] = refs
            if refs:
                any_formulas = True

    if has_calc_chain and "xl/calcChain.xml" in parsed:
        stale = []
        for c in parsed["xl/calcChain.xml"]:
            ref = c.get("r")
            sheet_id = c.get("i")
            target = sheet_id_to_target.get(sheet_id)
            if target is None:
                stale.append(f"{ref} (unknown sheetId {sheet_id})")
                continue
            if ref not in formula_refs_by_target.get(target, set()):
                stale.append(f"{ref} on {target} (no <f> at that ref)")
        if stale:
            report.add(
                "calc_chain",
                "CRITICAL",
                f"xl/calcChain.xml has {len(stale)} stale entr{'y' if len(stale) == 1 else 'ies'} "
                f"not matching an actual formula cell: {stale[:5]}" + (" ..." if len(stale) > 5 else ""),
            )
    elif any_formulas and not has_calc_chain:
        wb_root = parsed.get("xl/workbook.xml")
        full_calc = False
        if wb_root is not None:
            for calc_pr in wb_root.iter(f"{{{NS['main']}}}calcPr"):
                full_calc = calc_pr.get("fullCalcOnLoad") == "1"
        if not full_calc:
            report.add(
                "calc_chain",
                "WARNING",
                "formula cells present, xl/calcChain.xml absent, but "
                "<calcPr fullCalcOnLoad=\"1\"> not set -- the opening application "
                "may not know to recompute rather than trust nothing",
            )


# ---------------------------------------------------------------------------
# 4. sharedStrings.xml consistency
# ---------------------------------------------------------------------------


def check_shared_strings(parsed: dict[str, ET.Element], report: FileReport) -> None:
    ss_name = "xl/sharedStrings.xml"
    if ss_name not in parsed:
        return
    root = parsed[ss_name]
    si_elements = [c for c in root if c.tag == f"{{{NS['main']}}}si"]
    unique_count_attr = root.get("uniqueCount")
    count_attr = root.get("count")

    if unique_count_attr is not None and int(unique_count_attr) != len(si_elements):
        report.add(
            "shared_strings",
            "WARNING",
            f"sharedStrings uniqueCount={unique_count_attr} but actual <si> count={len(si_elements)}",
        )

    n_unique = len(si_elements)
    total_refs = 0
    out_of_range = []
    for name, root_el in parsed.items():
        if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
            continue
        for c in root_el.iter(f"{{{NS['main']}}}c"):
            if c.get("t") == "s":
                v_el = c.find(f"{{{NS['main']}}}v")
                if v_el is None or v_el.text is None:
                    continue
                total_refs += 1
                try:
                    idx = int(v_el.text)
                except ValueError:
                    continue
                if idx < 0 or idx >= n_unique:
                    out_of_range.append((name, c.get("r"), idx))

    if out_of_range:
        report.add(
            "shared_strings",
            "CRITICAL",
            f"{len(out_of_range)} cell(s) reference a sharedStrings index out of range "
            f"[0,{n_unique}): {out_of_range[:5]}" + (" ..." if len(out_of_range) > 5 else ""),
        )

    if count_attr is not None:
        try:
            declared_count = int(count_attr)
        except ValueError:
            declared_count = None
        if declared_count is not None and declared_count != total_refs:
            report.add(
                "shared_strings",
                "WARNING",
                f"sharedStrings count={declared_count} but actual t=\"s\" cell references "
                f"across all sheets={total_refs}",
            )


# ---------------------------------------------------------------------------
# 5. row/cell ascending order, no duplicates
# ---------------------------------------------------------------------------


def check_row_cell_order(parsed: dict[str, ET.Element], report: FileReport) -> None:
    for name, root in parsed.items():
        if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
            continue
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")
        if sheet_data is None:
            continue
        prev_row = -1
        for row in sheet_data:
            if row.tag != f"{{{NS['main']}}}row":
                continue
            r_attr = row.get("r")
            if r_attr is None:
                report.add("row_cell_order", "WARNING", f"{name}: <row> with no r attribute")
                continue
            row_num = int(r_attr)
            if row_num <= prev_row:
                report.add(
                    "row_cell_order",
                    "CRITICAL",
                    f"{name}: row r={row_num} out of order/duplicate (previous row r={prev_row})",
                )
            prev_row = max(prev_row, row_num)

            prev_col = -1
            for c in row:
                if c.tag != f"{{{NS['main']}}}c":
                    continue
                ref = c.get("r")
                if not ref:
                    report.add("row_cell_order", "WARNING", f"{name}: <c> with no r attribute in row {row_num}")
                    continue
                try:
                    col_letters, cell_row = _split_ref(ref)
                except ValueError:
                    report.add("row_cell_order", "WARNING", f"{name}: malformed cell ref {ref!r}")
                    continue
                if cell_row != row_num:
                    report.add(
                        "row_cell_order",
                        "CRITICAL",
                        f"{name}: cell {ref} inside <row r={row_num}> but ref implies row {cell_row}",
                    )
                col_num = _col_to_num(col_letters)
                if col_num <= prev_col:
                    report.add(
                        "row_cell_order",
                        "CRITICAL",
                        f"{name}: row {row_num} cell {ref} out of order/duplicate column "
                        f"(previous column index {prev_col})",
                    )
                prev_col = max(prev_col, col_num)


# ---------------------------------------------------------------------------
# 6. <dimension> validity
# ---------------------------------------------------------------------------


def check_dimension(parsed: dict[str, ET.Element], report: FileReport) -> None:
    for name, root in parsed.items():
        if not (name.startswith("xl/worksheets/sheet") and name.endswith(".xml")):
            continue
        dim_el = root.find(f"{{{NS['main']}}}dimension")
        sheet_data = root.find(f"{{{NS['main']}}}sheetData")
        if sheet_data is None:
            continue

        actual_cells = []
        for row in sheet_data:
            if row.tag != f"{{{NS['main']}}}row":
                continue
            for c in row:
                if c.tag != f"{{{NS['main']}}}c":
                    continue
                ref = c.get("r")
                if ref:
                    try:
                        actual_cells.append(_split_ref(ref))
                    except ValueError:
                        pass
        if not actual_cells:
            continue

        if dim_el is None:
            report.add("dimension", "WARNING", f"{name}: no <dimension> element but sheet has data")
            continue

        ref = dim_el.get("ref", "")
        parts = ref.split(":")
        try:
            if len(parts) == 1:
                start_col, start_row = _split_ref(parts[0])
                end_col, end_row = start_col, start_row
            elif len(parts) == 2:
                start_col, start_row = _split_ref(parts[0])
                end_col, end_row = _split_ref(parts[1])
            else:
                raise ValueError
        except ValueError:
            report.add("dimension", "WARNING", f"{name}: <dimension ref={ref!r}> is not a parseable cell/range")
            continue

        min_col_n = min(_col_to_num(start_col), _col_to_num(end_col))
        max_col_n = max(_col_to_num(start_col), _col_to_num(end_col))
        min_row_n = min(start_row, end_row)
        max_row_n = max(start_row, end_row)

        actual_max_row = max(r for _, r in actual_cells)
        actual_max_col = max(_col_to_num(c) for c, _ in actual_cells)
        actual_min_row = min(r for _, r in actual_cells)
        actual_min_col = min(_col_to_num(c) for c, _ in actual_cells)

        if actual_max_row > max_row_n or actual_max_col > max_col_n:
            report.add(
                "dimension",
                "WARNING",
                f"{name}: <dimension ref={ref!r}> undershoots actual data -- "
                f"data extends to row {actual_max_row}, col {actual_max_col} "
                f"(dimension bounds: row<={max_row_n}, col<={max_col_n})",
            )
        if actual_min_row < min_row_n or actual_min_col < min_col_n:
            report.add(
                "dimension",
                "WARNING",
                f"{name}: <dimension ref={ref!r}> starts after actual data "
                f"(data starts at row {actual_min_row}, col {actual_min_col})",
            )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


_RELS_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def check_relationships(z: zipfile.ZipFile, parsed: dict[str, ET.Element], report: FileReport) -> None:
    """CRITICAL for every internal relationship whose target part is missing.

    A ``*.rels`` part at ``<dir>/_rels/<name>.rels`` resolves relative targets
    against ``<dir>``; a leading ``/`` makes a target package-absolute.
    ``TargetMode="External"`` targets (URLs, other workbooks) are skipped.
    """
    names = set(z.namelist())
    for rels_name, root in parsed.items():
        if not rels_name.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(rels_name))
        for rel in root.iter(f"{_RELS_NS}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if part not in names:
                report.add(
                    "relationships",
                    "CRITICAL",
                    f"{rels_name}: {rel.get('Id')} points at missing part {part!r}",
                )


def validate_workbook(path: str | Path) -> FileReport:
    """Run all structural checks against a single .xlsx/.xlsm file.

    Returns a FileReport. ``report.clean`` is True iff there are no findings
    at all; ``report.critical_count == 0`` is the weaker "safe to ship"
    bar used by the rollback gate in ``safexl_mcp.engine`` (WARNING findings,
    e.g. a still-imperfect <dimension>, do not block a write).
    """
    path = Path(path)
    report = FileReport(path=str(path))
    try:
        with zipfile.ZipFile(path, "r") as z:
            parsed = check_well_formed_xml(z, report)
            check_content_types(z, parsed, report)
            check_relationships(z, parsed, report)
            check_calc_chain(z, parsed, report)
            check_shared_strings(parsed, report)
            check_row_cell_order(parsed, report)
            check_dimension(parsed, report)
    except zipfile.BadZipFile as e:
        report.add("zip", "CRITICAL", f"not a valid zip/OOXML package: {e}")
    return report


def validate_bytes_ok(path: str | Path) -> tuple[bool, FileReport]:
    """Convenience wrapper: (no CRITICAL findings, full report)."""
    report = validate_workbook(path)
    return report.critical_count == 0, report
