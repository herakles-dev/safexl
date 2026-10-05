"""Every reference-bearing part must follow an insert_rows / rename_sheet.

One test per part type so a failure names the thing that was left dangling.
Workbooks are built by hand (tests/_pkg.py) so each assertion is on exact XML.
Sheet names used: "Data" (the edited sheet) and "Other" (a second sheet).
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError
from safexl_mcp.structural import insert_rows, rename_sheet
from safexl_mcp.validate import validate_workbook

from _pkg import cell, make_package, read_part, row, sheet_xml

DATA = "xl/worksheets/sheet1.xml"
OTHER = "xl/worksheets/sheet2.xml"


def _ins(tmp_path, pkg, at, count=1, sheet="Data"):
    out = tmp_path / "out.xlsx"
    insert_rows(pkg, out, sheet, at, count)
    return out


def _ren(tmp_path, pkg, old, new):
    out = tmp_path / "out.xlsx"
    rename_sheet(pkg, out, old, new)
    return out


def _simple_rows(n=10):
    return "".join(row(i, cell(f"A{i}", v=str(i))) for i in range(1, n + 1))


# ===========================================================================
# insert_rows
# ===========================================================================


def test_insert_shifts_defined_names_including_print_titles(tmp_path):
    names = (
        "<definedNames>"
        '<definedName name="Rng">Data!$A$2:$A$10</definedName>'
        '<definedName name="Below">Data!$B$8</definedName>'
        '<definedName name="Above">Data!$B$1</definedName>'
        '<definedName name="OtherSheet">Other!$A$8</definedName>'
        '<definedName name="Quoted">\'Data\'!$C$5:$C$9</definedName>'
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$3</definedName>'
        '<definedName name="Lower">data!A9</definedName>'
        "</definedNames>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows())), ("Other", sheet_xml(""))],
        defined_names=names,
    )
    wb = read_part(_ins(tmp_path, pkg, 5), "xl/workbook.xml")
    assert ">Data!$A$2:$A$11</definedName>" in wb  # straddles row 5 -> grows
    assert 'name="Below">Data!$B$9<' in wb
    assert 'name="Above">Data!$B$1<' in wb
    assert 'name="OtherSheet">Other!$A$8<' in wb  # other sheet: untouched
    assert ">'Data'!$C$6:$C$10<" in wb
    assert ">Data!$1:$3<" in wb  # insert below the title rows: unchanged
    assert ">data!A10<" in wb  # sheet-name match is case-insensitive


def test_insert_grows_row_range_print_titles(tmp_path):
    names = '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$3</definedName></definedNames>'
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(_simple_rows()))], defined_names=names)
    wb = read_part(_ins(tmp_path, pkg, 2), "xl/workbook.xml")
    assert ">Data!$1:$4</definedName>" in wb


def test_insert_ignores_external_refs_structured_refs_and_strings(tmp_path):
    body = row(
        1,
        cell("A1", f='[1]Data!$B$8'),
        cell("B1", f='Table1[[#This Row],[Q8]]'),
        cell("C1", f='IF(A9="Data!A9",LOG10(A9),A9)', t="str"),
        cell("D1", f='SUM(Data!A9,\'Data\'!A9:B10)'),
    )
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(body))])
    xml = read_part(_ins(tmp_path, pkg, 5), DATA)
    assert "<f>[1]Data!$B$8</f>" in xml
    assert "<f>Table1[[#This Row],[Q8]]</f>" in xml
    assert '<f>IF(A10="Data!A9",LOG10(A10),A10)</f>' in xml
    assert "<f>SUM(Data!A10,'Data'!A10:B11)</f>" in xml


def test_insert_shifts_cross_sheet_formulas_on_other_sheet(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="Data!A8+Data!$A$3+'Data'!A5:A9"), cell("B1", f="A8")))
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(_simple_rows())), ("Other", other)])
    xml = read_part(_ins(tmp_path, pkg, 5), OTHER)
    assert "<f>Data!A9+Data!$A$3+'Data'!A6:A10</f>" in xml
    assert "<f>A8</f>" in xml  # unqualified ref on Other stays on Other


def test_insert_splits_shared_formula_group_into_explicit_formulas(tmp_path):
    # C2 = B1+1 fills down C2:C5 (each row refers one row up + 1 => B(r-1)... see below).
    body = "".join(
        [
            row(1, cell("B1", v="1")),
            row(2, cell("B2", v="1"), cell("C2", f="B1+1", fattrs=' t="shared" ref="C2:C5" si="0"', v="2")),
            row(3, cell("B3", v="1"), cell("C3", f="", fattrs=' t="shared" si="0"', v="2")),
            row(4, cell("B4", v="1"), cell("C4", f="", fattrs=' t="shared" si="0"', v="2")),
            row(5, cell("B5", v="1"), cell("C5", f="", fattrs=' t="shared" si="0"', v="2")),
        ]
    ).replace("<f t=\"shared\" si=\"0\"></f>", '<f t="shared" si="0"/>')
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(body))])
    xml = read_part(_ins(tmp_path, pkg, 4), DATA)
    # rows 2,3 stay; old row 4/5 -> 5/6. Each cell must still point at the SAME cells it did:
    assert '<c r="C2"><f>B1+1</f>' in xml
    assert '<c r="C3"><f>B2+1</f>' in xml
    assert '<c r="C5"><f>B3+1</f>' in xml  # was C4 = B3+1; B3 is above the insert
    assert '<c r="C6"><f>B5+1</f>' in xml  # was C5 = B4+1; B4 moved to B5
    assert 't="shared"' not in xml


def test_insert_leaves_unaffected_shared_group_shared_and_shifts_its_ref(tmp_path):
    # Group entirely below the insertion, every member only refers to row 1 (above it):
    # nothing about the formulas changes, so it stays shared; only ref= and the cells move.
    body = "".join(
        [
            row(10, cell("D10", f="$A$1*2", fattrs=' t="shared" ref="D10:D12" si="3"', v="2")),
            row(11, '<c r="D11"><f t="shared" si="3"/><v>2</v></c>'),
            row(12, '<c r="D12"><f t="shared" si="3"/><v>2</v></c>'),
        ]
    )
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(body))])
    xml = read_part(_ins(tmp_path, pkg, 5, count=2), DATA)
    assert '<f t="shared" ref="D12:D14" si="3">$A$1*2</f>' in xml
    assert '<c r="D13"><f t="shared" si="3"/>' in xml


def test_insert_expands_shared_group_when_its_refs_move(tmp_path):
    body = "".join(
        [
            row(10, cell("D10", f="A10*2", fattrs=' t="shared" ref="D10:D11" si="1"', v="2")),
            row(11, '<c r="D11"><f t="shared" si="1"/><v>2</v></c>'),
        ]
    )
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(body))])
    xml = read_part(_ins(tmp_path, pkg, 5), DATA)
    assert '<c r="D11"><f>A11*2</f>' in xml
    assert '<c r="D12"><f>A12*2</f>' in xml
    assert "shared" not in xml


def test_insert_shifts_array_formula_ref(tmp_path):
    body = row(8, cell("A8", f="B8:B9*2", fattrs=' t="array" ref="A8:A9"', v="1"))
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(body))])
    xml = read_part(_ins(tmp_path, pkg, 5), DATA)
    assert '<f t="array" ref="A9:A10">B9:B10*2</f>' in xml


def test_insert_shifts_cf_dv_and_x14_ranges(tmp_path):
    after = (
        '<mergeCells count="1"><mergeCell ref="A8:B8"/></mergeCells>'
        '<conditionalFormatting sqref="A2:A10 C8"><cfRule type="expression" dxfId="0" priority="1">'
        "<formula>A2&gt;A$9</formula></cfRule></conditionalFormatting>"
        '<dataValidations count="1"><dataValidation type="list" sqref="B2:B10">'
        "<formula1>$D$2:$D$10</formula1></dataValidation></dataValidations>"
        '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup>'
        "<x14:sparklines><x14:sparkline><xm:f>Data!B2:B10</xm:f><xm:sqref>E8</xm:sqref></x14:sparkline>"
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>"
        '<ext uri="{78C0D931-6437-4328-A341-0F0E8B9F4BD2}"><x14:conditionalFormattings><x14:conditionalFormatting>'
        '<x14:cfRule type="expression" priority="2"><xm:f>$A$7&gt;0</xm:f></x14:cfRule>'
        "<xm:sqref>F6:F9</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"
    )
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(_simple_rows(), after_data=after))])
    xml = read_part(_ins(tmp_path, pkg, 5), DATA)
    assert '<mergeCell ref="A9:B9"/>' in xml
    assert 'sqref="A2:A11 C9"' in xml
    assert "<formula>A2&gt;A$10</formula>" in xml
    assert '<dataValidation type="list" sqref="B2:B11">' in xml
    assert "<formula1>$D$2:$D$11</formula1>" in xml
    assert "<xm:f>Data!B2:B11</xm:f><xm:sqref>E9</xm:sqref>" in xml  # sparkline data + location
    assert "<xm:f>$A$8&gt;0</xm:f>" in xml
    assert "<xm:sqref>F7:F10</xm:sqref>" in xml


def test_insert_shifts_hyperlink_ref_and_location(tmp_path):
    after = (
        "<hyperlinks>"
        '<hyperlink ref="A8" location="Data!A9" display="own"/>'
        '<hyperlink ref="A2" location="\'Data\'!B8:B9"/>'
        '<hyperlink ref="A3" location="Other!A9"/>'
        "</hyperlinks>"
    )
    other = sheet_xml("", after_data='<hyperlinks><hyperlink ref="A1" location="Data!A9"/></hyperlinks>')
    pkg = make_package(
        tmp_path / "s.xlsx", [("Data", sheet_xml(_simple_rows(), after_data=after)), ("Other", other)]
    )
    out = _ins(tmp_path, pkg, 5)
    xml = read_part(out, DATA)
    assert '<hyperlink ref="A9" location="Data!A10" display="own"/>' in xml
    assert "<hyperlink ref=\"A2\" location=\"'Data'!B9:B10\"/>" in xml
    assert '<hyperlink ref="A3" location="Other!A9"/>' in xml
    assert '<hyperlink ref="A1" location="Data!A10"/>' in read_part(out, OTHER)


def test_insert_shifts_table_ref_autofilter_and_defined_name(tmp_path):
    table = (
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ref="A1:A10">'
        '<autoFilter ref="A1:A10"/><tableColumns count="1"><tableColumn id="1" name="c"/></tableColumns></table>'
    )
    names = '<definedNames><definedName name="Q" localSheetId="0">Data!$A$1:$A$10</definedName></definedNames>'
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows()))],
        defined_names=names,
        extra_parts={"xl/tables/table1.xml": table},
        sheet_rels={1: [("table", "../tables/table1.xml")]},
    )
    out = _ins(tmp_path, pkg, 5)
    t = read_part(out, "xl/tables/table1.xml")
    assert 'name="T" ref="A1:A11"' in t and '<autoFilter ref="A1:A11"/>' in t
    assert ">Data!$A$1:$A$11<" in read_part(out, "xl/workbook.xml")


def test_insert_above_table_moves_it_and_below_table_does_not_grow_it(tmp_path):
    table = (
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ref="A3:A6">'
        '<autoFilter ref="A3:A6"/></table>'
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows()))],
        extra_parts={"xl/tables/table1.xml": table},
        sheet_rels={1: [("table", "../tables/table1.xml")]},
    )
    assert 'ref="A5:A8"' in read_part(_ins(tmp_path, pkg, 1, count=2), "xl/tables/table1.xml")
    assert 'ref="A3:A6"' in read_part(_ins(tmp_path, pkg, 7), "xl/tables/table1.xml")


def test_insert_leaves_table_on_other_sheet_alone(tmp_path):
    table = '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ref="A1:A10"/>'
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows())), ("Other", sheet_xml(_simple_rows()))],
        extra_parts={"xl/tables/table1.xml": table},
        sheet_rels={2: [("table", "../tables/table1.xml")]},
    )
    assert 'ref="A1:A10"' in read_part(_ins(tmp_path, pkg, 5), "xl/tables/table1.xml")


def test_insert_shifts_chart_series_refs(tmp_path):
    chart = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>'
        "<c:barChart><c:ser>"
        "<c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>"
        "<c:cat><c:strRef><c:f>Data!$A$2:$A$10</c:f></c:strRef></c:cat>"
        "<c:val><c:numRef><c:f>'Data'!$B$2:$B$10</c:f><c:numCache><c:ptCount val=\"9\"/></c:numCache></c:numRef></c:val>"
        "</c:ser><c:ser><c:val><c:numRef><c:f>Other!$B$8</c:f></c:numRef></c:val></c:ser></c:barChart>"
        "</c:plotArea></c:chart></c:chartSpace>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx", [("Data", sheet_xml(_simple_rows()))], extra_parts={"xl/charts/chart1.xml": chart}
    )
    xml = read_part(_ins(tmp_path, pkg, 5), "xl/charts/chart1.xml")
    assert "<c:f>Data!$B$1</c:f>" in xml
    assert "<c:f>Data!$A$2:$A$11</c:f>" in xml
    assert "<c:f>'Data'!$B$2:$B$11</c:f>" in xml
    assert "<c:f>Other!$B$8</c:f>" in xml


def test_insert_shifts_pivot_cache_source_and_pivot_location(tmp_path):
    cache = (
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<cacheSource type="worksheet"><worksheetSource ref="A1:D13" sheet="Data"/></cacheSource>'
        "</pivotCacheDefinition>"
    )
    cache_other = cache.replace('sheet="Data"', 'sheet="Other"')
    pt = (
        '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P">'
        '<location ref="F8:H20" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/></pivotTableDefinition>'
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows())), ("Other", sheet_xml(""))],
        extra_parts={
            "xl/pivotCache/pivotCacheDefinition1.xml": cache,
            "xl/pivotCache/pivotCacheDefinition2.xml": cache_other,
            "xl/pivotTables/pivotTable1.xml": pt,
        },
        sheet_rels={1: [("pivotTable", "../pivotTables/pivotTable1.xml")]},
    )
    out = _ins(tmp_path, pkg, 8)
    assert 'ref="A1:D14" sheet="Data"' in read_part(out, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert 'ref="A1:D13" sheet="Other"' in read_part(out, "xl/pivotCache/pivotCacheDefinition2.xml")
    assert 'location ref="F9:H21"' in read_part(out, "xl/pivotTables/pivotTable1.xml")


def test_insert_shifts_comments_and_vml_note_anchor(tmp_path):
    comments = (
        '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors>'
        '<commentList><comment ref="A8" authorId="0"><text><t>x</t></text></comment>'
        '<comment ref="A2" authorId="0"><text><t>y</t></text></comment></commentList></comments>'
    )
    vml = (
        "<xml><v:shape><x:ClientData ObjectType=\"Note\"><x:Anchor>2, 15, 6, 2, 4, 15, 10, 4</x:Anchor>"
        "<x:Row>7</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>"
        "<v:shape><x:ClientData ObjectType=\"Note\"><x:Anchor>2, 15, 0, 2, 4, 15, 3, 4</x:Anchor>"
        "<x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows()))],
        extra_parts={"xl/comments1.xml": comments, "xl/drawings/vmlDrawing1.vml": vml},
        sheet_rels={1: [("comments", "../comments1.xml"), ("vmlDrawing", "../drawings/vmlDrawing1.vml")]},
    )
    out = _ins(tmp_path, pkg, 5)
    c = read_part(out, "xl/comments1.xml")
    assert '<comment ref="A9"' in c and '<comment ref="A2"' in c
    v = read_part(out, "xl/drawings/vmlDrawing1.vml")
    assert "<x:Row>8</x:Row>" in v and "<x:Row>1</x:Row>" in v
    assert "<x:Anchor>2, 15, 7, 2, 4, 15, 11, 4</x:Anchor>" in v
    assert "<x:Anchor>2, 15, 0, 2, 4, 15, 3, 4</x:Anchor>" in v


def test_insert_moves_drawing_anchors_unless_absolute(tmp_path):
    def anchor(kind, fr, to, extra=""):
        return (
            f'<xdr:twoCellAnchor{extra}><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row>'
            f"<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame/></xdr:twoCellAnchor>"
        )

    drawing = (
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
        + anchor("a", 2, 12)  # 0-based rows 2..12: straddles 1-based row 5 -> grows
        + anchor("b", 7, 14)  # below -> moves
        + anchor("c", 7, 14, ' editAs="absolute"')  # absolute -> stays
        + "</xdr:wsDr>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Data", sheet_xml(_simple_rows()))],
        extra_parts={"xl/drawings/drawing1.xml": drawing},
        sheet_rels={1: [("drawing", "../drawings/drawing1.xml")]},
    )
    xml = read_part(_ins(tmp_path, pkg, 5, count=3), "xl/drawings/drawing1.xml")
    assert "<xdr:row>2</xdr:row>" in xml and "<xdr:row>15</xdr:row>" in xml  # anchor a
    assert "<xdr:row>10</xdr:row>" in xml and "<xdr:row>17</xdr:row>" in xml  # anchor b
    assert xml.count("<xdr:row>7</xdr:row>") == 1 and xml.count("<xdr:row>14</xdr:row>") == 1  # anchor c


def test_insert_result_still_validates_and_drops_calc_chain(tmp_path):
    pkg = make_package(
        tmp_path / "s.xlsx", [("Data", sheet_xml(row(1, cell("A1", f="1+1", v="2"))))], calc_chain=True
    )
    out = _ins(tmp_path, pkg, 1)
    assert validate_workbook(out).critical_count == 0


def test_insert_past_row_limit_is_an_error_and_writes_nothing(tmp_path):
    pkg = make_package(tmp_path / "s.xlsx", [("Data", sheet_xml(row(1048576, cell("A1048576", v="1"))))])
    out = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="row limit"):
        insert_rows(pkg, out, "Data", 1, 1)
    assert not out.exists()


# ===========================================================================
# rename_sheet
# ===========================================================================


def test_rename_updates_chart_series_and_defined_names_and_formulas(tmp_path):
    chart = (
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:ser>'
        "<c:tx><c:strRef><c:f>Sheet2!$C$3</c:f></c:strRef></c:tx>"
        "<c:cat><c:strRef><c:f>Sheet2!$A$4:$B$11</c:f></c:strRef></c:cat>"
        "<c:val><c:numRef><c:f>'Keep Me'!$C$4</c:f></c:numRef></c:val></c:ser></c:chartSpace>"
    )
    names = (
        "<definedNames>"
        '<definedName name="N1">Sheet2!$Z$6</definedName>'
        '<definedName name="N2">SUM(Sheet2!$A$1:$A$3,Other!A1)</definedName>'
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet2!$1:$3</definedName>'
        "</definedNames>"
    )
    other = sheet_xml(
        row(1, cell("A1", f="Sheet2!A1+SHEET2!B1+'Sheet2'!C1"), cell("B1", f='"Sheet2!A1"&Sheet2!D1', t="str"))
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Sheet2", sheet_xml(_simple_rows())), ("Other", other)],
        defined_names=names,
        extra_parts={"xl/charts/chart1.xml": chart},
    )
    out = _ren(tmp_path, pkg, "Sheet2", "Yearly Totals")
    ch = read_part(out, "xl/charts/chart1.xml")
    assert "<c:f>'Yearly Totals'!$C$3</c:f>" in ch
    assert "<c:f>'Yearly Totals'!$A$4:$B$11</c:f>" in ch
    assert "<c:f>'Keep Me'!$C$4</c:f>" in ch
    wb = read_part(out, "xl/workbook.xml")
    assert '<sheet name="Yearly Totals"' in wb
    assert ">'Yearly Totals'!$Z$6<" in wb
    assert ">SUM('Yearly Totals'!$A$1:$A$3,Other!A1)<" in wb
    assert ">'Yearly Totals'!$1:$3<" in wb
    xml = read_part(out, OTHER)
    assert "<f>'Yearly Totals'!A1+'Yearly Totals'!B1+'Yearly Totals'!C1</f>" in xml
    assert "<f>\"Sheet2!A1\"&amp;'Yearly Totals'!D1</f>" in xml or "<f>&quot;Sheet2!A1&quot;&amp;'Yearly Totals'!D1</f>" in xml


def test_rename_from_quoted_to_plain_and_back(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="'Old Name'!A1+'Old Name'!B2")))
    pkg = make_package(tmp_path / "s.xlsx", [("Old Name", sheet_xml(_simple_rows())), ("Other", other)])
    out = _ren(tmp_path, pkg, "Old Name", "Plain")
    assert "<f>Plain!A1+Plain!B2</f>" in read_part(out, OTHER)


def test_rename_quotes_names_that_look_like_cells(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="Sheet2!A1")))
    pkg = make_package(tmp_path / "s.xlsx", [("Sheet2", sheet_xml(_simple_rows())), ("Other", other)])
    assert "<f>'Q1'!A1</f>" in read_part(_ren(tmp_path, pkg, "Sheet2", "Q1"), OTHER)


def test_rename_updates_cf_dv_and_sparkline_formulas(tmp_path):
    after = (
        '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>Sheet2!A1&gt;0</formula></cfRule>'
        "</conditionalFormatting>"
        '<dataValidations count="1"><dataValidation type="list" sqref="B1"><formula1>Sheet2!$A$1:$A$3</formula1>'
        "</dataValidation></dataValidations>"
        '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup>'
        "<x14:sparklines><x14:sparkline><xm:f>Sheet2!B2:B10</xm:f><xm:sqref>E8</xm:sqref></x14:sparkline>"
        "</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx", [("Sheet2", sheet_xml(_simple_rows())), ("Other", sheet_xml("", after_data=after))]
    )
    xml = read_part(_ren(tmp_path, pkg, "Sheet2", "Yearly"), OTHER)
    assert "<formula>Yearly!A1&gt;0</formula>" in xml
    assert "<formula1>Yearly!$A$1:$A$3</formula1>" in xml
    assert "<xm:f>Yearly!B2:B10</xm:f>" in xml


def test_rename_updates_hyperlink_location(tmp_path):
    after = (
        '<hyperlinks><hyperlink ref="A1" location="Sheet2!A1" display="d"/>'
        '<hyperlink ref="A2" location="\'Sheet2\'!B2:B3"/><hyperlink ref="A3" location="Other!A1"/></hyperlinks>'
    )
    pkg = make_package(
        tmp_path / "s.xlsx", [("Sheet2", sheet_xml(_simple_rows())), ("Other", sheet_xml("", after_data=after))]
    )
    xml = read_part(_ren(tmp_path, pkg, "Sheet2", "Yearly Totals"), OTHER)
    assert 'location="\'Yearly Totals\'!A1" display="d"' in xml
    assert "location=\"'Yearly Totals'!B2:B3\"" in xml
    assert 'location="Other!A1"' in xml


def test_rename_updates_pivot_cache_source_sheet(tmp_path):
    cache = (
        '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<cacheSource type="worksheet"><worksheetSource ref="A1:D13" sheet="Sheet2"/></cacheSource>'
        "</pivotCacheDefinition>"
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Sheet2", sheet_xml(_simple_rows()))],
        extra_parts={"xl/pivotCache/pivotCacheDefinition1.xml": cache},
    )
    out = _ren(tmp_path, pkg, "Sheet2", "R&D Data")
    assert 'ref="A1:D13" sheet="R&amp;D Data"' in read_part(out, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert '<sheet name="R&amp;D Data"' in read_part(out, "xl/workbook.xml")


def test_rename_updates_table_formula_and_textlink(tmp_path):
    table = (
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ref="A1:B3">'
        "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"c\"><calculatedColumnFormula>Sheet2!A1*2"
        "</calculatedColumnFormula></tableColumn></tableColumns></table>"
    )
    drawing = (
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
        '<xdr:twoCellAnchor><xdr:sp macro="" textlink="Sheet2!$A$1"/></xdr:twoCellAnchor></xdr:wsDr>'
    )
    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Sheet2", sheet_xml(_simple_rows())), ("Other", sheet_xml(""))],
        extra_parts={"xl/tables/table1.xml": table, "xl/drawings/drawing1.xml": drawing},
        sheet_rels={2: [("table", "../tables/table1.xml"), ("drawing", "../drawings/drawing1.xml")]},
    )
    out = _ren(tmp_path, pkg, "Sheet2", "Yearly")
    assert "<calculatedColumnFormula>Yearly!A1*2<" in read_part(out, "xl/tables/table1.xml")
    assert 'textlink="Yearly!$A$1"' in read_part(out, "xl/drawings/drawing1.xml")


def test_rename_leaves_external_and_structured_refs_alone(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="[1]Sheet2!A1+Sheet2!A1+Table1[Sheet2!x]")))
    pkg = make_package(tmp_path / "s.xlsx", [("Sheet2", sheet_xml(_simple_rows())), ("Other", other)])
    xml = read_part(_ren(tmp_path, pkg, "Sheet2", "Yearly"), OTHER)
    assert "<f>[1]Sheet2!A1+Yearly!A1+Table1[Sheet2!x]</f>" in xml


def test_rename_unicode_sheet_name_unquoted_refs(tmp_path):
    other = sheet_xml(row(1, cell("A1", f="Übersicht!A1")))
    pkg = make_package(tmp_path / "s.xlsx", [("Übersicht", sheet_xml(_simple_rows())), ("Other", other)])
    assert "<f>'Overview 1'!A1</f>" in read_part(_ren(tmp_path, pkg, "Übersicht", "Overview 1"), OTHER)


def test_rename_not_blocked_by_validation_finding_the_source_already_had(tmp_path):
    # calcChain has an entry for A1 but A1 has no formula: a CRITICAL the source already carries.
    pkg = make_package(tmp_path / "s.xlsx", [("Sheet1", sheet_xml(row(1, cell("A1", v="1"))))], calc_chain=True)
    assert validate_workbook(pkg).critical_count > 0
    out = _ren(tmp_path, pkg, "Sheet1", "Renamed")
    assert '<sheet name="Renamed"' in read_part(out, "xl/workbook.xml")


def test_rename_leaves_untouched_parts_byte_identical(tmp_path):
    import hashlib
    import zipfile

    pkg = make_package(
        tmp_path / "s.xlsx",
        [("Sheet2", sheet_xml(_simple_rows())), ("Other", sheet_xml(row(1, cell("A1", f="B1"))))],
        extra_parts={"xl/media/x.png": "png-bytes"},
    )
    out = _ren(tmp_path, pkg, "Sheet2", "Yearly")
    with zipfile.ZipFile(pkg) as a, zipfile.ZipFile(out) as b:
        for n in ("xl/media/x.png", OTHER, DATA):
            assert hashlib.sha256(a.read(n)).digest() == hashlib.sha256(b.read(n)).digest()
