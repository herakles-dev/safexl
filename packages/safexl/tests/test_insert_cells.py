"""insert_cells: Excel's 'insert cells, shift down' inside a column span.

Layout (built by hand): "Data" has three side-by-side lists in columns A, C and E,
rows 1..6, cell value = row * 10 / 100 / 1000. Inserting into A must leave C and E
exactly where they are, and every reference into A below the insert point follows.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.dom import minidom

import pytest

from safexl_mcp import SafeXLError
from safexl_mcp.structural import insert_cells, insert_rows
from safexl_mcp.validate import validate_workbook

from _pkg import cell, make_package, read_part, row, sheet_xml

DATA = "xl/worksheets/sheet1.xml"
OTHER = "xl/worksheets/sheet2.xml"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _rows(extra=None):
    extra = extra or {}
    return "".join(
        row(n, cell(f"A{n}", v=str(n * 10)), cell(f"C{n}", v=str(n * 100)), cell(f"E{n}", v=str(n * 1000)), *extra.get(n, ()))
        for n in range(1, 7)
    )


def _pkg(tmp_path, *, rows=None, after="", other=None, **kw):
    sheets = [("Data", sheet_xml(rows if rows is not None else _rows(), after_data=after))]
    if other is not None:
        sheets.append(("Other", other))
    return make_package(tmp_path / "src.xlsx", sheets, **kw)


def _ins(tmp_path, pkg, rng, **kw):
    out = tmp_path / "out.xlsx"
    insert_cells(pkg, out, sheet="Data", cell_range=rng, **kw)
    return out


def _cells(path, part=DATA):
    xml = read_part(path, part)
    return {m.group(1): m.group(2) for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(?:<v>([^<]*)</v>)?', xml)}


def test_only_the_column_block_moves_down(tmp_path):
    out = _ins(tmp_path, _pkg(tmp_path), "A3")
    c = _cells(out)
    assert "A3" not in c and c["A4"] == "30" and c["A7"] == "60" and c["A2"] == "20"
    assert all(c[f"C{n}"] == str(n * 100) and c[f"E{n}"] == str(n * 1000) for n in range(1, 7))
    assert validate_workbook(out).critical_count == 0


def test_multi_row_multi_column_span(tmp_path):
    out = _ins(tmp_path, _pkg(tmp_path), "C3:E4")  # 2 rows in C:E
    c = _cells(out)
    assert c["C5"] == "300" and c["E6"] == "4000" and c["E8"] == "6000" and "C3" not in c and "E4" not in c
    assert c["A3"] == "30" and c["A6"] == "60"


def test_inserting_past_the_data_creates_no_phantom_rows(tmp_path):
    out = _ins(tmp_path, _pkg(tmp_path), "A9")
    assert _cells(out) == _cells(_pkg(tmp_path))


def test_shift_right_is_refused_clearly(tmp_path):
    with pytest.raises(SafeXLError, match="shift='right' is not supported"):
        _ins(tmp_path, _pkg(tmp_path), "A3", shift="right")
    assert not (tmp_path / "out.xlsx").exists()


def test_bad_range_and_sheet(tmp_path):
    with pytest.raises(SafeXLError, match="invalid cell_range"):
        _ins(tmp_path, _pkg(tmp_path), "nope")
    with pytest.raises(SafeXLError):
        insert_cells(_pkg(tmp_path), tmp_path / "o.xlsx", sheet="Missing", cell_range="A1")


# --- formulas -----------------------------------------------------------------


def test_formula_refs_follow_only_inside_the_span(tmp_path):
    f = {
        1: (
            cell("G1", f="SUM(A1:A6)"),  # straddles A3 -> grows
            cell("G2", f="A4+A2"),  # A4 moves, A2 does not
            cell("G3", f="C4+E4"),  # other columns untouched
            cell("G4", f="SUM(A1:A2)"),  # ends above -> untouched
            cell("G5", f="$A$5:$A$6"),  # wholly below -> moves
            cell("G6", f="C:C+A:A"),  # whole-column refs never change
        )
    }
    out = _ins(tmp_path, _pkg(tmp_path, rows=_rows(f)), "A3")
    xml = read_part(out, DATA)
    for ref, formula in [
        ("G1", "SUM(A1:A7)"), ("G2", "A5+A2"), ("G3", "C4+E4"), ("G4", "SUM(A1:A2)"), ("G5", "$A$6:$A$7"), ("G6", "C:C+A:A")
    ]:
        assert f'<c r="{ref}"><f>{formula}</f></c>' in xml, ref


def test_cross_sheet_refs_follow_and_other_sheet_own_cells_do_not(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="Data!A4"), cell("B1", f="Data!C4"), cell("A4", v="9")))
    out = _ins(tmp_path, _pkg(tmp_path, other=other), "A3")
    xml = read_part(out, OTHER)
    assert "<f>Data!A5</f>" in xml and "<f>Data!C4</f>" in xml and '<c r="A4"><v>9</v></c>' in xml


def test_defined_name_grows_and_other_list_name_is_untouched(tmp_path):
    names = (
        "<definedNames>"
        '<definedName name="listA">Data!$A$2:$A$6</definedName>'
        '<definedName name="listC">Data!$C$2:$C$6</definedName>'
        '<definedName name="tail">Data!$A$5</definedName>'
        "</definedNames>"
    )
    out = _ins(tmp_path, _pkg(tmp_path, defined_names=names), "A3")
    wb = read_part(out, "xl/workbook.xml")
    assert ">Data!$A$2:$A$7<" in wb and ">Data!$C$2:$C$6<" in wb and ">Data!$A$6<" in wb


def test_dropdown_list_range_grows_to_include_the_inserted_cell(tmp_path):
    names = '<definedNames><definedName name="listA">Data!$A$2:$A$6</definedName></definedNames>'
    after = (
        '<dataValidations count="2">'
        '<dataValidation type="list" sqref="G1:G5"><formula1>listA</formula1></dataValidation>'
        '<dataValidation type="list" sqref="H1"><formula1>Data!$A$2:$A$6</formula1></dataValidation>'
        "</dataValidations>"
    )
    out = _ins(tmp_path, _pkg(tmp_path, defined_names=names, after=after), "A4")
    xml = read_part(out, DATA)
    assert "<formula1>Data!$A$2:$A$7</formula1>" in xml and "<formula1>listA</formula1>" in xml
    assert ">Data!$A$2:$A$7<" in read_part(out, "xl/workbook.xml")
    assert 'sqref="G1:G5"' in xml  # a DV on other columns is not moved


def test_dv_and_cf_sqrefs_move_with_their_cells_and_split_at_the_span_edge(tmp_path):
    after = (
        '<conditionalFormatting sqref="A4:A6"><cfRule type="expression" priority="1"><formula>A4&gt;1</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="A1:C6"><cfRule type="expression" priority="2"><formula>C1&gt;1</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="E1:E6"><cfRule type="expression" priority="3"><formula>E1&gt;1</formula></cfRule></conditionalFormatting>'
        '<dataValidations count="1"><dataValidation type="whole" sqref="A2:A5"><formula1>0</formula1></dataValidation></dataValidations>'
    )
    xml = read_part(_ins(tmp_path, _pkg(tmp_path, after=after), "A3"), DATA)
    assert 'sqref="A5:A7"' in xml and "<formula>A5&gt;1</formula>" in xml
    assert 'sqref="B1:C6 A1:A7"' in xml or 'sqref="A1:A7 B1:C6"' in xml
    assert 'sqref="E1:E6"' in xml and 'sqref="A2:A6"' in xml


def test_shared_formula_group_in_the_span_is_expanded_and_correct(tmp_path):
    child = '<c r="B{n}"><f t="shared" si="0"/></c>'
    rows = (
        row(1, cell("A1", v="1"), cell("B1", v="1"))
        + row(2, cell("A2", v="1"), cell("B2", f="A2*2", fattrs=' t="shared" ref="B2:B4" si="0"'))
        + row(3, cell("A3", v="1"), child.format(n=3))
        + row(4, cell("A4", v="1"), child.format(n=4))
    )
    out = _ins(tmp_path, _pkg(tmp_path, rows=rows), "B3")  # moves B3:B4 only, not the master
    xml = read_part(out, DATA)
    assert 't="shared"' not in xml
    assert '<c r="B2"><f>A2*2</f></c>' in xml
    assert '<c r="B4"><f>A3*2</f></c>' in xml and '<c r="B5"><f>A4*2</f></c>' in xml


# --- torn references and structures ---------------------------------------------


def test_torn_formula_range_is_refused_with_the_reference_named(tmp_path):
    rows = _rows({1: (cell("G1", f="SUM(A1:C6)"),)})
    with pytest.raises(SafeXLError, match=r"A1:C6.*partly overlaps the shifted columns.*A:A"):
        _ins(tmp_path, _pkg(tmp_path, rows=rows), "A3")
    assert not (tmp_path / "out.xlsx").exists()


def test_torn_range_above_the_insert_point_is_fine(tmp_path):
    rows = _rows({1: (cell("G1", f="SUM(A1:C2)"),)})
    assert "<f>SUM(A1:C2)</f>" in read_part(_ins(tmp_path, _pkg(tmp_path, rows=rows), "A3"), DATA)


def test_torn_defined_name_is_refused(tmp_path):
    names = '<definedNames><definedName name="wide">Data!$A$2:$C$6</definedName></definedNames>'
    with pytest.raises(SafeXLError, match="wide|A\\$2:\\$C\\$6"):
        _ins(tmp_path, _pkg(tmp_path, defined_names=names), "A3")


def test_merge_moves_but_straddling_or_partial_merges_are_refused(tmp_path):
    ok = read_part(_ins(tmp_path, _pkg(tmp_path, after='<mergeCells count="2"><mergeCell ref="A5:A6"/><mergeCell ref="C5:D6"/></mergeCells>'), "A3"), DATA)
    assert 'ref="A6:A7"' in ok and 'ref="C5:D6"' in ok
    for ref in ("A2:A4", "A1:B6"):
        with pytest.raises(SafeXLError, match="insert_cells refused"):
            _ins(tmp_path, _pkg(tmp_path, after=f'<mergeCells count="1"><mergeCell ref="{ref}"/></mergeCells>'), "A3")


def _table(ref, name="T", n=1, cols=("a", "b")):
    entries = "".join(f'<tableColumn id="{i}" name="{c}"/>' for i, c in enumerate(cols, 1))
    return (
        f'<table xmlns="{NS}" id="{n}" name="{name}" displayName="{name}" ref="{ref}"><autoFilter ref="{ref}"/>'
        f'<tableColumns count="{len(cols)}">{entries}</tableColumns></table>'
    )


def test_table_inside_the_span_grows_moves_and_a_torn_table_is_refused(tmp_path):
    kw = lambda ref: dict(extra_parts={"xl/tables/table1.xml": _table(ref)}, sheet_rels={1: [("table", "../tables/table1.xml")]})
    grown = read_part(_ins(tmp_path, _pkg(tmp_path, **kw("A1:B6")), "A3:B3"), "xl/tables/table1.xml")
    assert 'ref="A1:B7"' in grown and '<autoFilter ref="A1:B7"/>' in grown
    moved = read_part(_ins(tmp_path, _pkg(tmp_path, **kw("A4:B6")), "A3:B3"), "xl/tables/table1.xml")
    assert 'ref="A5:B7"' in moved
    with pytest.raises(SafeXLError, match=r"A1:B6.*partly overlaps"):
        _ins(tmp_path, _pkg(tmp_path, **kw("A1:B6")), "A3")  # only column A shifts: table is torn
    assert 'ref="A1:B6"' in read_part(_ins(tmp_path, _pkg(tmp_path, **kw("A1:B6")), "E3"), "xl/tables/table1.xml")


def test_array_formula_that_would_be_split_is_refused(tmp_path):
    rows = _rows({4: (cell("G4", f="A4:A5", fattrs=' t="array" ref="G4:G5"'),)})
    # G is outside the span: untouched
    assert 'ref="G4:G5"' in read_part(_ins(tmp_path, _pkg(tmp_path, rows=rows), "A3"), DATA)
    with pytest.raises(SafeXLError, match="insert_cells refused"):
        _ins(tmp_path, _pkg(tmp_path, rows=rows), "G5")


# --- other reference-bearing parts ------------------------------------------------


def test_hyperlinks_comments_vml_chart_and_pivot_follow(tmp_path):
    comments = (
        f'<comments xmlns="{NS}"><authors><author>a</author></authors><commentList>'
        '<comment ref="A5" authorId="0"><text><t>x</t></text></comment>'
        '<comment ref="C5" authorId="0"><text><t>y</t></text></comment></commentList></comments>'
    )
    vml = (
        '<xml><v:shape><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 3, 2, 4, 15, 7, 4</x:Anchor>'
        "<x:Row>4</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>"
        '<v:shape><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 3, 2, 4, 15, 7, 4</x:Anchor>'
        "<x:Row>4</x:Row><x:Column>2</x:Column></x:ClientData></v:shape></xml>"
    )
    chart = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart>'
        "<c:ser><c:cat><c:strRef><c:f>Data!$A$2:$A$6</c:f></c:strRef></c:cat>"
        "<c:val><c:numRef><c:f>Data!$C$2:$C$6</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
    )
    cache = f'<pivotCacheDefinition xmlns="{NS}"><cacheSource type="worksheet"><worksheetSource ref="A1:A6" sheet="Data"/></cacheSource></pivotCacheDefinition>'
    pt = f'<pivotTableDefinition xmlns="{NS}" name="P"><location ref="G5:H8" firstHeaderRow="1"/></pivotTableDefinition>'
    after = '<hyperlinks><hyperlink ref="A5" location="Data!A5"/><hyperlink ref="C5" location="Data!C5"/></hyperlinks>'
    pkg = _pkg(
        tmp_path,
        after=after,
        extra_parts={
            "xl/comments1.xml": comments,
            "xl/drawings/vmlDrawing1.vml": vml,
            "xl/charts/chart1.xml": chart,
            "xl/pivotCache/pivotCacheDefinition1.xml": cache,
            "xl/pivotTables/pivotTable1.xml": pt,
        },
        sheet_rels={
            1: [
                ("comments", "../comments1.xml"),
                ("vmlDrawing", "../drawings/vmlDrawing1.vml"),
                ("pivotTable", "../pivotTables/pivotTable1.xml"),
            ]
        },
    )
    out = _ins(tmp_path, pkg, "A3")
    assert '<hyperlink ref="A6" location="Data!A6"/>' in read_part(out, DATA)
    assert '<hyperlink ref="C5" location="Data!C5"/>' in read_part(out, DATA)
    c = read_part(out, "xl/comments1.xml")
    assert '<comment ref="A6"' in c and '<comment ref="C5"' in c
    v = read_part(out, "xl/drawings/vmlDrawing1.vml")
    assert v.count("<x:Row>5</x:Row>") == 1 and v.count("<x:Row>4</x:Row>") == 1
    ch = read_part(out, "xl/charts/chart1.xml")
    assert "<c:f>Data!$A$2:$A$7</c:f>" in ch and "<c:f>Data!$C$2:$C$6</c:f>" in ch
    assert 'ref="A1:A7" sheet="Data"' in read_part(out, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert 'location ref="G5:H8"' in read_part(out, "xl/pivotTables/pivotTable1.xml")  # G:H are outside the span


def test_unrelated_parts_stay_byte_identical(tmp_path):
    other = sheet_xml(row(1, cell("A1", v="5")))
    chart = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser>'
        "<c:val><c:numRef><c:f>Other!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"
    )
    pkg = _pkg(tmp_path, other=other, extra_parts={"xl/charts/chart1.xml": chart, "xl/sharedStrings.xml": f'<sst xmlns="{NS}"/>'})
    out = _ins(tmp_path, pkg, "A3")
    with zipfile.ZipFile(pkg) as a, zipfile.ZipFile(out) as b:
        changed = {n for n in a.namelist() if n in b.namelist() and a.read(n) != b.read(n)}
    assert changed == {DATA, "xl/workbook.xml"}  # the sheet, plus fullCalcOnLoad


def test_whole_row_insert_behaviour_is_unchanged(tmp_path):
    out = tmp_path / "rows.xlsx"
    insert_rows(_pkg(tmp_path), out, "Data", 3)
    c = _cells(out)
    assert c["A4"] == "30" and c["C4"] == "300" and c["E4"] == "3000"


# --- the real corpus workbook ------------------------------------------------------

CORPUS = Path(__file__).resolve().parents[3] / "research" / "xlsx-bench" / "corpus" / "structured-refs-lookup-tool.xlsx"


def test_corpus_two_offset_lists_stay_valid(tmp_path):
    if not CORPUS.is_file():
        pytest.skip("run research/xlsx-bench/fetch_corpus.py")
    src = tmp_path / "w.xlsx"
    shutil.copy(CORPUS, src)
    mid = tmp_path / "mid.xlsx"
    out = tmp_path / "out.xlsx"
    insert_cells(src, mid, sheet="dropdowns", cell_range="A13")  # list 'state'
    insert_cells(mid, out, sheet="dropdowns", cell_range="C14")  # list 'states', offset one row
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        wb = b.read("xl/workbook.xml").decode()
        assert re.search(r'<definedName name="state"[^>]*>dropdowns!\$A\$2:\$A\$54<', wb)
        assert re.search(r'<definedName name="states"[^>]*>dropdowns!\$C\$2:\$C\$55<', wb)
        changed = {n for n in a.namelist() if n in b.namelist() and a.read(n) != b.read(n)}
        assert "xl/worksheets/sheet4.xml" in changed
        assert not any(n.startswith(("xl/tables/", "xl/charts/", "xl/pivot")) for n in changed - {"xl/worksheets/sheet4.xml"})
        assert b.testzip() is None
        for n in b.namelist():
            if n.endswith((".xml", ".rels")):
                minidom.parseString(b.read(n))
    assert validate_workbook(out).critical_count == validate_workbook(src).critical_count
    xml = zipfile.ZipFile(out).read("xl/worksheets/sheet4.xml").decode()
    # the Category block (E:G, rows 1-13) did not move
    assert re.findall(r'<c r="E1[0-3]"', xml) == re.findall(r'<c r="E1[0-3]"', zipfile.ZipFile(src).read("xl/worksheets/sheet4.xml").decode())
