"""safexl_mcp.view -- compact read-only cell viewer + terse text renderer."""
from __future__ import annotations

import hashlib

import pytest

from safexl_mcp import SafeXLError
from safexl_mcp.view import ViewError, render_text, view

from _pkg import make_package, sheet_xml
from _util import fixture_path, FIXTURE_NAMES

SST = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">'
    '<si><t>Groceries</t></si>'
    '<si><r><t>Rent </t></r><r><t>due</t></r></si>'
    '<si><t>1250</t></si>'
    "</sst>"
)


def _wb(tmp_path):
    s1 = sheet_xml(
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>420</v></c><c r="C1" t="s"><v>2</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>1250.5</v></c>'
        '<c r="C2"><f>SUM(B1:B2)</f><v>1670.5</v></c></row>'
        '<row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3" t="e"><v>#DIV/0!</v></c>'
        '<c r="C3" t="inlineStr"><is><t>inline</t></is></c><c r="D3" s="4"/></row>'
        '<row r="4"><c r="A4"><f t="shared" ref="A4:A5" si="0">B1*2</f><v>840</v></c></row>'
        '<row r="5"><c r="A5"><f t="shared" si="0"/><v>2500</v></c>'
        '<c r="B5"><f t="array" ref="B5:B5">SUM(B1:B2)</f><v>1670.5</v></c>'
        '<c r="C5" t="str"><f>"x"&amp;"y"</f><v>xy</v></c></row>'
    )
    s2 = sheet_xml('<row r="10"><c r="B10" t="inlineStr"><is><t>Groceries tax</t></is></c></row>')
    return make_package(
        tmp_path / "w.xlsx", [("Budget", s1), ("Other", s2)], extra_parts={"xl/sharedStrings.xml": SST}
    )


def _by_ref(res):
    return {c["ref"]: c for c in res["cells"]}


def test_values_types_and_shared_strings(tmp_path):
    res = view(_wb(tmp_path))
    c = _by_ref(res)
    assert res["sheets"] == ["Budget", "Other"]
    assert res["sheet"] == "Budget"
    assert res["dims"] == "A1:C5"
    assert c["A1"]["value"] == "Groceries" and c["A1"]["type"] == "s"
    assert c["A2"]["value"] == "Rent due"
    assert c["B1"]["value"] == 420 and c["B1"]["type"] == "n"
    assert c["B2"]["value"] == 1250.5
    assert c["C1"]["value"] == "1250" and c["C1"]["type"] == "s"  # number stored as text stays a string
    assert c["A3"] == {"ref": "A3", "type": "b", "value": True}
    assert c["B3"]["type"] == "e" and c["B3"]["value"] == "#DIV/0!"
    assert c["C3"]["value"] == "inline"
    assert "D3" not in c  # styled-but-empty cell skipped
    assert res["truncated"] is False


def test_formula_text_cached_value_and_markers(tmp_path):
    c = _by_ref(view(_wb(tmp_path)))
    assert c["C2"]["formula"] == "=SUM(B1:B2)" and c["C2"]["value"] == 1670.5
    assert c["A4"]["formula"] == "=B1*2" and c["A4"]["shared"] == "0"
    assert c["A5"]["formula"] == "=B2*2" and c["A5"]["shared"] == "0" and c["A5"]["value"] == 2500
    assert c["B5"]["array"] == "B5:B5" and c["B5"]["formula"] == "=SUM(B1:B2)"
    assert c["C5"]["type"] == "s" and c["C5"]["value"] == "xy"
    assert "formula" not in _by_ref(view(_wb(tmp_path), formulas=False))["C2"]


def test_render_text_lines(tmp_path):
    text = render_text(view(_wb(tmp_path)))
    lines = text.splitlines()
    assert lines[0].startswith("# Budget")
    assert "C2  n  1670.5  =SUM(B1:B2)" in lines
    assert 'A1  s  "Groceries"' in lines
    assert "A3  b  TRUE" in lines
    assert "A5  n  2500  =B2*2" in lines
    assert "A4  n  840  =B1*2 [shared 0]" in lines
    assert "B5  n  1670.5  =SUM(B1:B2) [array]" in lines


def test_find_hits_value_and_formula_across_sheets(tmp_path):
    res = view(_wb(tmp_path), find="groceries")
    assert [(c["sheet"], c["ref"]) for c in res["cells"]] == [("Budget", "A1"), ("Other", "B10")]
    res = view(_wb(tmp_path), find="sum(b1")
    assert {c["ref"] for c in res["cells"]} == {"C2", "B5"}
    assert "Budget!C2  n  1670.5  =SUM(B1:B2)" in render_text(res)
    res = view(_wb(tmp_path), sheet="Other", find="groceries")
    assert [c["ref"] for c in res["cells"]] == ["B10"]
    assert view(_wb(tmp_path), find="nomatch-zzz")["cells"] == []


def test_find_truncates_with_hint(tmp_path):
    res = view(_wb(tmp_path), find="e", max_cells=1)
    assert len(res["cells"]) == 1 and res["truncated"] is True and "hint" in res
    assert "truncated" in render_text(res)


def test_range_window_and_truncation_next_range(tmp_path):
    p = _wb(tmp_path)
    res = view(p, sheet="Budget", cell_range="B1:C2")
    assert [c["ref"] for c in res["cells"]] == ["B1", "C1", "B2", "C2"]
    res = view(p, cell_range="B2")
    assert [c["ref"] for c in res["cells"]] == ["B2"]

    first = view(p, max_cells=4)
    # the cut fell mid-row 2 (after A2), so the window ends at row 1 and the next one starts at row 2
    assert [c["ref"] for c in first["cells"]] == ["A1", "B1", "C1"] and first["truncated"] is True
    assert first["next_range"] == "A2:C5"
    seen = [c["ref"] for c in first["cells"]]
    while True:
        nxt = view(p, cell_range=first["next_range"], max_cells=4)
        seen += [c["ref"] for c in nxt["cells"] if c["ref"] not in seen]
        if not nxt["truncated"]:
            break
        first = nxt
    assert set(seen) == {c["ref"] for c in view(p)["cells"]}
    assert "cell_range='" in render_text(view(p, max_cells=4))


def test_clean_errors(tmp_path):
    p = _wb(tmp_path)
    with pytest.raises(ValueError, match="no sheet 'Nope'"):
        view(p, sheet="Nope")
    with pytest.raises(ViewError):
        view(p, sheet="Nope", find="x")
    for bad in ("ZZ", "A0:B2", "1:2", "A1:"):
        with pytest.raises(ValueError, match="bad range"):
            view(p, cell_range=bad)
    with pytest.raises(ValueError):
        view(p, max_cells=0)
    with pytest.raises(SafeXLError):
        view(tmp_path / "missing.xlsx")
    junk = tmp_path / "junk.xlsx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(SafeXLError):
        view(junk)


def test_macro_workbook_read_without_executing_or_writing():
    p = fixture_path("SimpleMacro.xlsm")
    before = hashlib.sha256(p.read_bytes()).hexdigest()
    res = view(p)
    assert res["sheets"] and isinstance(res["cells"], list)
    render_text(res)
    assert hashlib.sha256(p.read_bytes()).hexdigest() == before


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_every_fixture_views_and_renders(name):
    res = view(fixture_path(name), max_cells=50)
    assert len(res["cells"]) <= 50
    render_text(res)


def test_formulas_fixture_has_formulas():
    res = view(fixture_path("formulas_definednames.xlsx"), find="=")
    assert res["cells"] and all("formula" in c for c in res["cells"])


def test_next_range_never_repeats_cells(tmp_path):
    p = _wb(tmp_path)
    full = [c["ref"] for c in view(p)["cells"]]
    for budget in (3, 4, 5, 6):
        seen, rng = [], None
        for _ in range(20):
            res = view(p, cell_range=rng, max_cells=budget)
            seen += [c["ref"] for c in res["cells"]]
            if not res["truncated"]:
                break
            rng = res["next_range"]
        assert seen == full, (budget, seen)


def test_row_wider_than_budget_is_cut_and_paging_goes_on_from_next_row(tmp_path):
    p = _wb(tmp_path)
    res = view(p, max_cells=2)
    assert [c["ref"] for c in res["cells"]] == ["A1", "B1"]
    assert res["next_range"] == "A2:C5" and "rest starts at C1" in res["hint"]
    assert "rest starts at C1" in render_text(res)


def _shared_wb(tmp_path):
    s1 = sheet_xml(
        '<row r="2"><c r="C2"><f t="shared" ref="C2:C4" si="3">A2*$B$1+Other!D2+SUM(A2:B$2)+"A2"</f><v>1</v></c></row>'
        '<row r="3"><c r="C3"><f t="shared" si="3"/><v>2</v></c></row>'
        '<row r="4"><c r="C4"><f t="shared" si="3"/><v>3</v></c><c r="D4"><f t="shared" ref="D4" si="4">LOG10(C4)</f><v>0</v></c></row>'
        '<row r="5"><c r="D5"><f t="shared" si="4"/><v>0</v></c></row>'
    )
    return make_package(tmp_path / "s.xlsx", [("S", s1), ("Other", sheet_xml(""))])


def test_shared_child_shows_shifted_formula_text(tmp_path):
    p = _shared_wb(tmp_path)
    c = _by_ref(view(p))
    assert c["C2"]["formula"] == '=A2*$B$1+Other!D2+SUM(A2:B$2)+"A2"'
    # relative refs shift by the row offset, $ refs and string literals do not
    assert c["C3"]["formula"] == '=A3*$B$1+Other!D3+SUM(A3:B$2)+"A2"'
    assert c["C4"]["formula"] == '=A4*$B$1+Other!D4+SUM(A4:B$2)+"A2"'
    assert c["C4"]["shared"] == "3" and c["C4"]["shared_from"] == "3"
    # a function name that looks like a ref is not shifted
    assert c["D5"]["formula"] == "=LOG10(C5)"
    lines = render_text(view(p)).splitlines()
    assert any(l.startswith("C3  n  2  =A3*$B$1") and "[shared" not in l for l in lines)
    assert any(l.startswith("C2  n  1  =A2*") and "[shared 3]" in l for l in lines)


def test_shared_child_column_shift_and_find(tmp_path):
    from safexl_mcp.view import _shift_formula

    assert _shift_formula("A1+$A1+A$1+Sheet1!B2", 2, 1) == "C2+$A2+C$1+Sheet1!D3"
    assert _shift_formula("SUM(A:B)+SUM(1:2)", 1, 1) == "SUM(B:C)+SUM(2:3)"
    assert _shift_formula("A1", -1, 0) == "#REF!"
    assert _shift_formula("'My Sheet'!A1+1E5", 0, 1) == "'My Sheet'!A2+1E5"
    hits = view(_shared_wb(tmp_path), find="a4*")["cells"]
    assert [h["ref"] for h in hits] == ["C4"]
