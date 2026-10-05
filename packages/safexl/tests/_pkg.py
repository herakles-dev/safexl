"""Tiny OOXML package builder for the structural-edit reference tests.

Unlike ``_util.build_minimal_workbook`` this lets a test hand in whole sheet
bodies plus arbitrary extra parts (tables, charts, comments, drawings, pivot
caches) and per-sheet relationships, so each reference-bearing part the
structural edits must follow can be asserted on directly.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def sheet_xml(body: str, *, after_data: str = "", dimension: str = "") -> str:
    """A worksheet: ``body`` is the <row> elements, ``after_data`` is everything
    that follows </sheetData> (mergeCells, CF, DV, hyperlinks, extLst...)."""
    dim = f'<dimension ref="{dimension}"/>' if dimension else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}" '
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
        'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        f"{dim}<sheetData>{body}</sheetData>{after_data}</worksheet>"
    )


def make_package(
    dst: Path,
    sheets: list[tuple[str, str]],
    *,
    defined_names: str = "",
    extra_parts: dict[str, str] | None = None,
    sheet_rels: dict[int, list[tuple[str, str]]] | None = None,
    calc_chain: bool = False,
) -> Path:
    """``sheets`` is [(name, worksheet xml)], part paths sheet1.xml, sheet2.xml...
    ``sheet_rels`` maps a 1-based sheet index to [(relationship type suffix, target)]."""
    n = len(sheets)
    ct = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]
    for i in range(1, n + 1):
        ct.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    if calc_chain:
        ct.append(
            '<Override PartName="/xl/calcChain.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
        )
    ct.append("</Types>")

    sheet_tags = "".join(
        f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _) in enumerate(sheets, start=1)
    )
    wb = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheets>{sheet_tags}</sheets>'
        f"{defined_names}</workbook>"
    )
    wb_rels = "".join(
        f'<Relationship Id="rId{i}" Type="{REL}worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, n + 1)
    )
    if calc_chain:
        wb_rels += f'<Relationship Id="rId{n + 1}" Type="{REL}calcChain" Target="calcChain.xml"/>'

    dst = Path(dst)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}officeDocument" Target="xl/workbook.xml"/></Relationships>',
        )
        z.writestr("xl/workbook.xml", wb)
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{wb_rels}</Relationships>',
        )
        for i, (_, xml) in enumerate(sheets, start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
        for idx, rels in (sheet_rels or {}).items():
            body = "".join(
                f'<Relationship Id="rId{k}" Type="{REL}{t}" Target="{tgt}"/>' for k, (t, tgt) in enumerate(rels, start=1)
            )
            z.writestr(
                f"xl/worksheets/_rels/sheet{idx}.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>',
            )
        if calc_chain:
            z.writestr(
                "xl/calcChain.xml",
                f'<calcChain xmlns="{NS_MAIN}"><c r="A1" i="1"/></calcChain>',
            )
        for path, text in (extra_parts or {}).items():
            z.writestr(path, text)
    return dst


def read_part(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def cell(ref: str, *, v: str | None = None, f: str | None = None, fattrs: str = "", t: str | None = None) -> str:
    tt = f' t="{t}"' if t else ""
    inner = (f"<f{fattrs}>{f}</f>" if f is not None else "") + (f"<v>{v}</v>" if v is not None else "")
    return f'<c r="{ref}"{tt}>{inner}</c>'


def row(n: int, *cells: str) -> str:
    return f'<row r="{n}">{"".join(cells)}</row>'
