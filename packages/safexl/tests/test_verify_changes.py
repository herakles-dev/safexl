"""verify() keeps its loss-based verdicts and adds added_risk, cell_changes and summary (S2)."""
from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from safexl_mcp import edit
from safexl_mcp.verify import verify
from _pkg import make_package, sheet_xml

BODY = (
    '<row r="1"><c r="A1" t="inlineStr"><is><t>name</t></is></c><c r="B1"><v>5</v></c><c r="C1" s="0"><v>7</v></c></row>'
    '<row r="2"><c r="A2" t="inlineStr"><is><t>bob</t></is></c><c r="B2"><v>6</v></c></row>'
    '<row r="3"><c r="A3"><v>1</v></c></row>'
)


def _base(tmp_path) -> Path:
    return make_package(tmp_path / "in.xlsx", [("Sheet1", sheet_xml(BODY))])


def _tamper(src: Path, dst: Path, edits: dict[str, tuple[str, str]] | None = None, add: dict[str, str] | None = None) -> Path:
    """Copy ``src`` to ``dst`` replacing text in parts: {part: (old, new)}; ``add`` adds parts."""
    edits = edits or {}
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename in edits:
                old, new = edits[item.filename]
                text = data.decode("utf-8")
                assert old in text, (item.filename, old)
                data = text.replace(old, new, 1).encode("utf-8")
            zo.writestr(item, data)
        for name, content in (add or {}).items():
            zo.writestr(name, content)
    return dst


SHEET = "xl/worksheets/sheet1.xml"


def _risks(r):
    return {x["risk"] for x in r["added_risk"]}


def test_identical_files_report_nothing(tmp_path):
    src = _base(tmp_path)
    r = verify(src, src)
    assert r["verdict"] == "intact" and r["added_risk"] == [] and r["cell_changes"] == {}
    assert r["summary"] == "No tracked feature was lost and untouched parts are byte-identical; no cell changes; nothing risky was added."


def test_normal_edit_of_three_cells_reports_exactly_three(tmp_path):
    src, out = _base(tmp_path), tmp_path / "out.xlsx"
    edit(src, out, [
        {"sheet": "Sheet1", "cell": "B2", "value": 60},          # overwrite a number
        {"sheet": "Sheet1", "cell": "D1", "value": "new"},        # new cell in an existing row
        {"sheet": "Sheet1", "cell": "A9", "formula": "B1+B2"},    # new cell in a new row
    ])
    r = verify(src, out)
    assert r["verdict"] == "intact"
    assert r["added_risk"] == []
    c = r["cell_changes"]["Sheet1"]
    assert c["cells_changed"] == 3 and c["added"] == 2 and c["values_changed"] == 1 and c["removed"] == 0
    assert "3 cells changed on Sheet1" in r["summary"] and "nothing risky was added" in r["summary"]


def test_numeric_and_text_changes_are_counted(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ('<c r="B1"><v>5</v></c>', '<c r="B1"><v>5000</v></c>')})
    out2 = _tamper(out, tmp_path / "o2.xlsx", {SHEET: ("<t>bob</t>", "<t>mallory</t>")})
    r = verify(src, out2)
    assert r["verdict"] == "intact"  # loss-based verdict is unchanged
    assert r["cell_changes"]["Sheet1"]["values_changed"] == 2
    assert r["cell_changes"]["Sheet1"]["cells_changed"] == 2


def test_formula_injected_into_a_value_cell(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ('<c r="B2"><v>6</v></c>', '<c r="B2"><f>WEBSERVICE("http://evil/x")</f></c>')})
    r = verify(src, out)
    assert r["verdict"] == "intact"
    assert r["cell_changes"]["Sheet1"]["formulas_changed"] == 1
    assert "formula_webservice_added" in _risks(r)


def test_hyperlink_and_dde_formulas_are_named(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ('<c r="B2"><v>6</v></c>', '<c r="B2"><f>HYPERLINK("http://e","x")</f></c><c r="C2"><f>cmd|\'/c calc\'!A1</f></c>')})
    assert {"formula_hyperlink_added", "formula_dde_added"} <= _risks(verify(src, out))


def test_deleted_cell_and_deleted_row(tmp_path):
    src = _base(tmp_path)
    cell_gone = _tamper(src, tmp_path / "o1.xlsx", {SHEET: ('<c r="C1" s="0"><v>7</v></c>', "")})
    r = verify(src, cell_gone)
    assert r["cell_changes"]["Sheet1"]["removed"] == 1 and r["cell_changes"]["Sheet1"]["rows_removed"] == 0
    row_gone = _tamper(src, tmp_path / "o2.xlsx", {SHEET: ('<row r="3"><c r="A3"><v>1</v></c></row>', "")})
    r = verify(src, row_gone)
    assert r["cell_changes"]["Sheet1"]["rows_removed"] == 1 and r["cell_changes"]["Sheet1"]["removed"] == 1
    assert "1 row removed on Sheet1" in r["summary"]


def test_style_index_change_is_counted(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ('<c r="C1" s="0">', '<c r="C1" s="3">')})
    assert verify(src, out)["cell_changes"]["Sheet1"]["styles_changed"] == 1


def test_hidden_column_and_sheet_protection(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ("<sheetData>", '<cols><col min="2" max="2" width="9" hidden="1"/></cols><sheetData>')})
    assert "hidden_rows_or_columns_added" in _risks(verify(src, out))
    out = _tamper(src, tmp_path / "o2.xlsx", {SHEET: ("</sheetData>", '</sheetData><sheetProtection sheet="1"/>')})
    assert "sheet_protection_added" in _risks(verify(src, out))


def test_veryhidden_sheet_and_calc_settings(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {"xl/workbook.xml": ('<sheet name="Sheet1"', '<sheet state="veryHidden" name="Sheet1"')})
    assert "sheet_hidden" in _risks(verify(src, out))
    out = _tamper(src, tmp_path / "o2.xlsx", {"xl/workbook.xml": ("</workbook>", '<calcPr iterate="1"/></workbook>')})
    r = verify(src, out)
    assert "calculation_settings_changed" in _risks(r) and "iterate" in r["added_risk"][0]["detail"]


def test_fullcalconload_alone_is_not_a_calc_change(tmp_path):
    src = _base(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {"xl/workbook.xml": ("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')})
    assert "calculation_settings_changed" not in _risks(verify(src, out))


def test_macro_part_and_external_link_added(tmp_path):
    src = _base(tmp_path)
    ct = ("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    out = _tamper(src, tmp_path / "o.xlsx", {"[Content_Types].xml": ct}, {"xl/vbaProject.bin": "MACRO"})
    assert "vba_project_added" in _risks(verify(src, out))

    ct = ("</Types>", '<Override PartName="/xl/externalLinks/externalLink1.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>')
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://x/externalLinkPath" Target="http://evil.example/x.xlsx" TargetMode="External"/></Relationships>')
    out = _tamper(src, tmp_path / "o2.xlsx", {"[Content_Types].xml": ct}, {
        "xl/externalLinks/externalLink1.xml": '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        "xl/externalLinks/_rels/externalLink1.xml.rels": rels,
    })
    r = verify(src, out)
    assert {"external_link_added", "external_target_added"} <= _risks(r)
    assert "http://evil.example/x.xlsx" in " ".join(x["detail"] for x in r["added_risk"])


def test_invalid_file_has_the_new_fields(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"nope")
    r = verify(bad, bad)
    assert r["verdict"] == "invalid" and r["added_risk"] == [] and r["cell_changes"] == {} and r["summary"]


# --- real recorded runs ---------------------------------------------------------

_REPO = Path(__file__).resolve().parents[3]
_GUARD = _REPO / "research/xlsx-bench/recordings/haiku-guard-2026-10-04"


def test_all_arm_b_guard_outputs_still_verify_intact():
    if not _GUARD.is_dir():
        pytest.skip("recorded runs not present")
    checked = 0
    for run in sorted(_GUARD.glob("*/B/*/run.json")):
        meta = json.loads(run.read_text())
        original = _REPO / meta["workbook"]
        outs = sorted((run.parent / "work").glob("*.xls*"))
        if not original.is_file() or not outs:
            continue
        for out in outs:
            r = verify(original, out)
            assert r["verdict"] == "intact", (run.parent, r["reasons"])
            assert isinstance(r["added_risk"], list) and r["summary"]
            checked += 1
    if not checked:
        pytest.skip("no arm-B outputs found")
    assert checked >= 1


# --- row inserts are not reported as thousands of changed cells ---------------------------------

from safexl_mcp import insert_rows  # noqa: E402
from safexl_mcp.structural import insert_cells  # noqa: E402


def _tall(tmp_path):
    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><v>{r * 10}</v></c><c r="B{r}"><f>A{r}*2</f><v>{r * 20}</v></c></row>'
        for r in range(1, 31)
    )
    rows += '<row r="31"><c r="A31"><f>SUM(A1:A30)</f><v>4650</v></c></row>'
    return make_package(tmp_path / "tall.xlsx", [("Sheet1", sheet_xml(rows))])


def test_insert_rows_reports_the_insert_not_the_shift(tmp_path):
    src, out = _tall(tmp_path), tmp_path / "o.xlsx"
    insert_rows(src, out, "Sheet1", 5, 2)
    r = verify(src, out)
    c = r["cell_changes"]["Sheet1"]
    assert r["verdict"] == "intact"
    assert c["rows_inserted"] == {"count": 2, "blocks": [{"row": 5, "count": 2}]}
    assert c["values_changed"] == 0 and c["cells_changed"] == 0 and c["rows_removed"] == 0
    assert c["formulas_changed"] == 0 and c["formulas_shifted"] > 0
    assert r["summary"].count(";") == 2 and "2 rows inserted on Sheet1" in r["summary"] and "changed" not in r["summary"].split(";")[1]


def test_insert_then_edit_reports_one_changed_cell(tmp_path):
    src, mid, out = _tall(tmp_path), tmp_path / "m.xlsx", tmp_path / "o.xlsx"
    insert_rows(src, mid, "Sheet1", 5, 2)
    edit(mid, out, [{"sheet": "Sheet1", "cell": "A5", "value": "new"}])
    c = verify(src, out)["cell_changes"]["Sheet1"]
    assert c["rows_inserted"]["count"] == 2 and c["cells_changed"] == 1 and c["added"] == 1


def test_a_real_row_deletion_is_still_reported_next_to_an_insert(tmp_path):
    src, mid = _tall(tmp_path), tmp_path / "m.xlsx"
    insert_rows(src, mid, "Sheet1", 5, 2)
    out = _tamper(mid, tmp_path / "o.xlsx", {SHEET: ('<row r="22"><c r="A22"><v>200</v></c><c r="B22"><f>A22*2</f><v>400</v></c></row>', "")})
    c = verify(src, out)["cell_changes"]["Sheet1"]
    assert c["rows_inserted"]["count"] == 2 and c["rows_removed"] == 1 and c["removed"] == 2


def test_a_plain_deletion_is_not_mistaken_for_an_insert(tmp_path):
    src = _tall(tmp_path)
    out = _tamper(src, tmp_path / "o.xlsx", {SHEET: ('<row r="10"><c r="A10"><v>100</v></c><c r="B10"><f>A10*2</f><v>200</v></c></row>', "")})
    c = verify(src, out)["cell_changes"]["Sheet1"]
    assert "rows_inserted" not in c and c["rows_removed"] == 1


def test_insert_cells_is_aligned_or_flagged_positional(tmp_path):
    src, out = _tall(tmp_path), tmp_path / "o.xlsx"
    insert_cells(src, out, sheet="Sheet1", cell_range="A5:A6")
    r = verify(src, out)
    c = r["cell_changes"]["Sheet1"]
    assert c.get("rows_inserted") or (c.get("positional") is True and c.get("shift_detected") is False)
    assert r["summary"].endswith("nothing risky was added.")
    if c.get("positional"):
        assert "counts for that sheet are positional" in r["summary"]
