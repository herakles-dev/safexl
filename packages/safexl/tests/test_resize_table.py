"""resize_table and edit(extend_tables=True): growing an Excel table next to an added column or row."""
from __future__ import annotations

import posixpath
import re
import shutil
import zipfile
from pathlib import Path
from xml.dom import minidom

import pytest

from safexl_mcp import SafeXLError, edit
from safexl_mcp.structural import plan_table_extensions, resize_table

from _pkg import make_package, read_part, row, sheet_xml

TBL = "xl/tables/table1.xml"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _istr(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def _table(ref="A1:C4", cols=("Alpha", "Beta", "Gamma"), extra="", af=None):
    entries = "".join(f'<tableColumn id="{i}" name="{c}"/>' for i, c in enumerate(cols, start=1))
    return (
        f'<table xmlns="{NS}" id="1" name="Sales" displayName="Sales" ref="{ref}">'
        f'<autoFilter ref="{af or ref}"/><tableColumns count="{len(cols)}">{entries}{extra}</tableColumns></table>'
    )


def _pkg(tmp_path, *, body=None, after="", table=None, defined="", extra_parts=None, second_table=None):
    if body is None:
        body = (
            row(1, _istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma"), _istr("D1", "Delta"))
            + row(2, '<c r="A2"><v>1</v></c>', '<c r="B2"><v>2</v></c>', '<c r="C2"><v>3</v></c>')
            + row(3, '<c r="A3"><v>4</v></c>', '<c r="B3"><v>5</v></c>', '<c r="C3"><v>6</v></c>')
            + row(4, '<c r="A4"><v>7</v></c>', '<c r="B4"><v>8</v></c>', '<c r="C4"><v>9</v></c>')
        )
    parts = {TBL: table or _table()}
    rels = [("table", "../tables/table1.xml")]
    if second_table:
        parts["xl/tables/table2.xml"] = second_table
        rels.append(("table", "../tables/table2.xml"))
    parts.update(extra_parts or {})
    return make_package(
        tmp_path / "src.xlsx",
        [("Data", sheet_xml(body, after_data=after))],
        defined_names=defined,
        extra_parts=parts,
        sheet_rels={1: rels},
    )


def _resize(tmp_path, pkg, **kw):
    out = tmp_path / "out.xlsx"
    resize_table(pkg, out, **kw)
    return out


def test_grow_one_column_takes_header_from_sheet(tmp_path):
    out = _resize(tmp_path, _pkg(tmp_path), table="Sales", ref="A1:D4")
    t = read_part(out, TBL)
    assert re.search(r'<table\b[^>]*\bref="A1:D4"', t)
    assert '<autoFilter ref="A1:D4"/>' in t
    assert '<tableColumns count="4">' in t
    assert '<tableColumn id="4" name="Delta"/>' in t
    assert t.count("<tableColumn ") == 4


def test_only_table_part_changes(tmp_path):
    pkg = _pkg(tmp_path)
    out = _resize(tmp_path, pkg, table="Sales", ref="A1:D4")
    with zipfile.ZipFile(pkg) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
    assert changed == [TBL]


def test_lookup_by_display_name_and_sheet(tmp_path):
    out = _resize(tmp_path, _pkg(tmp_path), table="sales", ref="A1:D4", sheet="Data")
    assert 'ref="A1:D4"' in read_part(out, TBL)


def test_grow_and_shrink_rows(tmp_path):
    pkg = _pkg(tmp_path)
    grown = _resize(tmp_path, pkg, table="Sales", ref="A1:C9")
    t = read_part(grown, TBL)
    assert 'ref="A1:C9"' in t and '<autoFilter ref="A1:C9"/>' in t and '<tableColumns count="3">' in t
    shrunk = tmp_path / "shrunk.xlsx"
    resize_table(pkg, shrunk, table="Sales", ref="A1:C2")
    assert 'ref="A1:C2"' in read_part(shrunk, TBL)


def test_header_from_shared_string(tmp_path):
    body = (
        row(1, _istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma"), '<c r="D1" t="s"><v>0</v></c>')
        + row(2, '<c r="A2"><v>1</v></c>')
        + row(3, '<c r="A3"><v>1</v></c>')
        + row(4, '<c r="A4"><v>1</v></c>')
    )
    pkg = _pkg(
        tmp_path,
        body=body,
        extra_parts={"xl/sharedStrings.xml": f'<sst xmlns="{NS}" count="1" uniqueCount="1"><si><t>Shared &amp; Co</t></si></sst>'},
    )
    out = _resize(tmp_path, pkg, table="Sales", ref="A1:D4")
    assert 'name="Shared &amp; Co"' in read_part(out, TBL)


def test_refuse_empty_header(tmp_path):
    pkg = _pkg(tmp_path)
    with pytest.raises(SafeXLError, match="header cell E1 is empty"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:E4")
    assert not (tmp_path / "o.xlsx").exists()


def test_refuse_duplicate_header(tmp_path):
    body = row(1, _istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma"), _istr("D1", "beta")) + row(2) + row(3) + row(4)
    pkg = _pkg(tmp_path, body=body)
    with pytest.raises(SafeXLError, match="duplicates an existing column name"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:D4")


def test_refuse_removing_column_used_by_structured_reference(tmp_path):
    body = (
        row(1, _istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma"))
        + row(2, '<c r="A2"><v>1</v></c>')
        + row(3)
        + row(4)
        + row(6, '<c r="F6"><f>SUM(Sales[Gamma])</f></c>')
    )
    pkg = _pkg(tmp_path, body=body)
    with pytest.raises(SafeXLError, match=r"Gamma.*formula in Data!F6"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:B4")


def test_refuse_removing_column_used_by_defined_name(tmp_path):
    pkg = _pkg(tmp_path, defined='<definedNames><definedName name="Totals">SUM(Sales[Gamma])</definedName></definedNames>')
    with pytest.raises(SafeXLError, match="defined name 'Totals'"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:B4")


def test_refuse_removing_column_used_by_calculated_column(tmp_path):
    calc = '<tableColumn id="4" name="Calc"><calculatedColumnFormula>Sales[[#This Row],[Gamma]]*2</calculatedColumnFormula></tableColumn>'
    # Gamma is the 3rd column; shrink the 4-column table to A:B so Gamma and Calc both go,
    # but Calc is removed along with it, so the user is only the removed column itself: allowed.
    table = _table(ref="A1:D4", cols=("Alpha", "Beta", "Gamma"), extra=calc)
    pkg = _pkg(tmp_path, table=table)
    out = _resize(tmp_path, pkg, table="Sales", ref="A1:B4")
    assert "Calc" not in read_part(out, TBL)
    # Dropping only Calc is fine too; dropping Gamma while Calc stays is impossible by position.
    other = (
        f'<table xmlns="{NS}" id="2" name="Other" displayName="Other" ref="H1:H3"><tableColumns count="1">'
        '<tableColumn id="1" name="X"><calculatedColumnFormula>SUM(Sales[Gamma])</calculatedColumnFormula></tableColumn>'
        "</tableColumns></table>"
    )
    pkg2 = _pkg(tmp_path, second_table=other)
    with pytest.raises(SafeXLError, match="calculated column 'X' of table 'Other'"):
        resize_table(pkg2, tmp_path / "o2.xlsx", table="Sales", ref="A1:B4")


def test_removing_unreferenced_column_is_allowed(tmp_path):
    out = _resize(tmp_path, _pkg(tmp_path), table="Sales", ref="A1:B4")
    t = read_part(out, TBL)
    assert 'ref="A1:B4"' in t and '<tableColumns count="2">' in t and "Gamma" not in t


def test_refuse_overlap_with_other_table(tmp_path):
    other = f'<table xmlns="{NS}" id="2" name="Other" displayName="Other" ref="D1:E3"/>'
    pkg = _pkg(tmp_path, second_table=other)
    with pytest.raises(SafeXLError, match="overlaps table 'Other'"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:D4")


def test_refuse_overlap_with_merged_range(tmp_path):
    pkg = _pkg(tmp_path, after='<mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>')
    with pytest.raises(SafeXLError, match="merged range D1:E1"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="A1:D4")


def test_refuse_moving_top_left_and_unknown_table(tmp_path):
    pkg = _pkg(tmp_path)
    with pytest.raises(SafeXLError, match="top-left"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Sales", ref="B1:C4")
    with pytest.raises(SafeXLError, match="not found"):
        resize_table(pkg, tmp_path / "o.xlsx", table="Nope", ref="A1:C4")


def test_totals_row_autofilter_excludes_totals(tmp_path):
    table = _table(ref="A1:C5", af="A1:C4").replace('ref="A1:C5">', 'ref="A1:C5" totalsRowCount="1">', 1)
    out = _resize(tmp_path, _pkg(tmp_path, table=table), table="Sales", ref="A1:C7")
    t = read_part(out, TBL)
    assert 'ref="A1:C7"' in t and '<autoFilter ref="A1:C6"/>' in t


# --- edit(extend_tables=...) -------------------------------------------------


def _edit(tmp_path, pkg, changes, **kw):
    out = tmp_path / "edited.xlsx"
    edit(pkg, out, changes, **kw)
    return out


def test_edit_extend_tables_grows_table_for_adjacent_column(tmp_path):
    pkg = _pkg(tmp_path)
    out = _edit(
        tmp_path,
        pkg,
        [
            {"sheet": "Data", "cell": "D1", "value": "Total"},
            {"sheet": "Data", "cell": "D2", "formula": "A2+B2"},
            {"sheet": "Data", "cell": "D3", "formula": "A3+B3"},
        ],
        extend_tables=True,
    )
    t = read_part(out, TBL)
    assert re.search(r'<table\b[^>]*\bref="A1:D4"', t)
    assert '<autoFilter ref="A1:D4"/>' in t
    assert '<tableColumns count="4">' in t
    assert '<tableColumn id="4" name="Total"/>' in t


def test_edit_extend_tables_grows_table_for_row_below(tmp_path):
    out = _edit(tmp_path, _pkg(tmp_path), [{"sheet": "Data", "cell": "B5", "value": 10}], extend_tables=True)
    t = read_part(out, TBL)
    assert 'ref="A1:C5"' in t and '<tableColumns count="3">' in t


def test_edit_extend_tables_ignores_non_adjacent_writes(tmp_path):
    pkg = _pkg(tmp_path)
    out = _edit(
        tmp_path,
        pkg,
        [{"sheet": "Data", "cell": "E2", "value": 1}, {"sheet": "Data", "cell": "A6", "value": 1}, {"sheet": "Data", "cell": "D5", "value": 1}],
        extend_tables=True,
    )
    assert read_part(out, TBL) == read_part(pkg, TBL)


def _headerless_pkg(tmp_path, cols=("Alpha", "Beta", "Gamma")):
    heads = [_istr(f"{c}1", n) for c, n in zip("ABC", cols)]
    return _pkg(tmp_path, body=row(1, *heads) + row(2) + row(3) + row(4), table=_table(cols=cols))


def test_edit_extend_tables_auto_names_empty_header_like_excel(tmp_path):
    pkg = _headerless_pkg(tmp_path)
    out = _edit(
        tmp_path,
        pkg,
        [{"sheet": "Data", "cell": "D2", "formula": "A2+B2"}, {"sheet": "Data", "cell": "E3", "value": 1}],
        extend_tables=True,
    )
    t = read_part(out, TBL)
    assert 'ref="A1:E4"' in t and '<tableColumns count="5">' in t
    assert '<tableColumn id="4" name="Column1"/>' in t and '<tableColumn id="5" name="Column2"/>' in t
    sheet = read_part(out, "xl/worksheets/sheet1.xml")
    assert re.search(r'<c r="D1"[^>]*t="inlineStr"><is><t[^>]*>Column1</t>', sheet)
    assert re.search(r'<c r="E1"[^>]*t="inlineStr"><is><t[^>]*>Column2</t>', sheet)


def test_auto_name_skips_names_already_in_use(tmp_path):
    pkg = _headerless_pkg(tmp_path, cols=("Column1", "Beta", "Column2"))
    out = _edit(tmp_path, pkg, [{"sheet": "Data", "cell": "D2", "value": 1}], extend_tables=True)
    assert 'name="Column3"' in read_part(out, TBL)


def test_edit_extend_tables_still_refuses_duplicate_header(tmp_path):
    pkg = _pkg(tmp_path)
    with pytest.raises(SafeXLError, match="duplicates"):
        edit(pkg, tmp_path / "x.xlsx", [{"sheet": "Data", "cell": "D1", "value": "alpha"}, {"sheet": "Data", "cell": "D2", "value": 1}], extend_tables=True)
    assert not (tmp_path / "x.xlsx").exists()


def test_resize_table_stays_strict_about_empty_header(tmp_path):
    with pytest.raises(SafeXLError, match="header cell D1 is empty"):
        resize_table(_headerless_pkg(tmp_path), tmp_path / "o.xlsx", table="Sales", ref="A1:D4")


def test_plan_table_extensions_is_read_only_and_reports(tmp_path):
    pkg = _headerless_pkg(tmp_path)
    before = pkg.read_bytes()
    plan = plan_table_extensions(
        pkg, [{"sheet": "Data", "cell": "D1", "value": "Total"}, {"sheet": "Data", "cell": "D2", "value": 1}, {"sheet": "Data", "cell": "E2", "value": 1}]
    )
    assert plan == [
        {
            "table": "Sales",
            "sheet": "Data",
            "old_ref": "A1:C4",
            "new_ref": "A1:E4",
            "new_columns": [
                {"header_cell": "D1", "name": "Total", "auto_named": False},
                {"header_cell": "E1", "name": "Column1", "auto_named": True},
            ],
        }
    ]
    assert pkg.read_bytes() == before
    assert plan_table_extensions(pkg, [{"sheet": "Data", "cell": "H9", "value": 1}]) == []


def _sorted_table(sort_ref, cond_ref):
    return (
        f'<table xmlns="{NS}" id="1" name="Sales" displayName="Sales" ref="A1:C4"><autoFilter ref="A1:C4"/>'
        f'<sortState ref="{sort_ref}"><sortCondition ref="{cond_ref}"/></sortState>'
        '<tableColumns count="3"><tableColumn id="1" name="Alpha"/><tableColumn id="2" name="Beta"/>'
        '<tableColumn id="3" name="Gamma"/></tableColumns></table>'
    )


def test_grow_moves_full_cover_sortstate_and_leaves_partial_alone(tmp_path):
    full = _resize(tmp_path, _pkg(tmp_path, table=_sorted_table("A2:C4", "B2:B4")), table="Sales", ref="A1:D9")
    t = read_part(full, TBL)
    assert '<sortState ref="A2:D9">' in t and '<sortCondition ref="B2:B9"/>' in t and '<autoFilter ref="A1:D9"/>' in t
    part_dir = tmp_path / "p"
    part_dir.mkdir()
    partial = _resize(part_dir, _pkg(part_dir, table=_sorted_table("A2:C3", "B2:B3")), table="Sales", ref="A1:D9")
    t = read_part(partial, TBL)
    assert '<sortState ref="A2:C3">' in t and '<sortCondition ref="B2:B3"/>' in t


def test_shrink_drops_sortstate_that_no_longer_fits(tmp_path):
    out = _resize(tmp_path, _pkg(tmp_path, table=_sorted_table("A2:C4", "C2:C4")), table="Sales", ref="A1:B4")
    assert "sortState" not in read_part(out, TBL)


def test_edit_default_leaves_table_byte_identical(tmp_path):
    pkg = _pkg(tmp_path)
    changes = [{"sheet": "Data", "cell": "D2", "formula": "A2+B2"}, {"sheet": "Data", "cell": "A5", "value": 3}]
    plain = _edit(tmp_path, pkg, changes)
    with zipfile.ZipFile(pkg) as a, zipfile.ZipFile(plain) as b:
        assert a.read(TBL) == b.read(TBL)
    explicit = tmp_path / "e.xlsx"
    edit(pkg, explicit, changes, extend_tables=False)
    with zipfile.ZipFile(pkg) as a, zipfile.ZipFile(explicit) as b:
        assert a.read(TBL) == b.read(TBL)


# --- real workbooks from the benchmark corpus (copied; corpus is never written) ---

CORPUS = Path(__file__).resolve().parents[3] / "research" / "xlsx-bench" / "corpus"

# (corpus file, a real plain table in it)
CASES = [
    ("pension-hedging-stress-model.xlsx", "Portfolios"),
]


def _copy(tmp_path, name):
    src = CORPUS / name
    if not src.is_file():
        pytest.skip("run research/xlsx-bench/fetch_corpus.py")
    dst = tmp_path / name
    shutil.copy(src, dst)
    return dst


def _table_location(path, table):
    """(table part, sheet name, sheet part, ref) for ``table``."""
    with zipfile.ZipFile(path) as z:
        wb = z.read("xl/workbook.xml").decode()
        rels = z.read("xl/_rels/workbook.xml.rels").decode()
        for part in (n for n in z.namelist() if re.match(r"xl/tables/table\d+\.xml$", n)):
            xml = z.read(part).decode()
            if f'name="{table}"' not in xml:
                continue
            ref = re.search(r'<table\b[^>]*\bref="([^"]+)"', xml).group(1)
            for sheet_rels in (n for n in z.namelist() if n.startswith("xl/worksheets/_rels/")):
                if f"/{posixpath.basename(part)}" in z.read(sheet_rels).decode():
                    sheet_part = "xl/worksheets/" + posixpath.basename(sheet_rels)[: -len(".rels")]
                    rid = re.search(
                        rf'<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="[^"]*{posixpath.basename(sheet_part)}"', rels
                    ) or re.search(rf'<Relationship\b[^>]*Target="[^"]*{posixpath.basename(sheet_part)}"[^>]*Id="([^"]+)"', rels)
                    name = re.search(rf'<sheet\b[^>]*name="([^"]*)"[^>]*r:id="{rid.group(1)}"', wb).group(1)
                    return part, name, sheet_part, ref
    raise AssertionError(f"table {table} not found")


def _split(ref):
    a, b = ref.split(":")
    m1, m2 = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", b)
    return m1.group(1), int(m1.group(2)), m2.group(1), int(m2.group(2))


def _next_col(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    n += 1
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _assert_valid_package(path):
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        names = set(z.namelist())
        for n in names:
            if n.endswith((".xml", ".rels", ".vml")) and n != "xl/vbaProject.bin":
                minidom.parseString(z.read(n))
        for rels in (n for n in names if n.endswith(".rels")):
            base = posixpath.dirname(posixpath.dirname(rels))
            for m in re.finditer(r"<Relationship\b[^>]*/>", z.read(rels).decode()):
                tag = m.group(0)
                if 'TargetMode="External"' in tag:
                    continue
                target = re.search(r'Target="([^"]*)"', tag).group(1)
                full = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
                assert full in names, f"{rels}: dangling target {target}"
        ct = z.read("[Content_Types].xml").decode()
        defaults = set(re.findall(r'<Default\b[^>]*Extension="([^"]+)"', ct))
        overrides = set(re.findall(r'<Override\b[^>]*PartName="([^"]+)"', ct))
        for n in names:
            if n == "[Content_Types].xml":
                continue
            assert f"/{n}" in overrides or n.rsplit(".", 1)[-1] in defaults, f"no content type for {n}"


@pytest.mark.parametrize("name,table", CASES)
@pytest.mark.parametrize("with_header", [True, False])
def test_extend_real_table(tmp_path, name, table, with_header):
    src = _copy(tmp_path, name)
    part, sheet, sheet_part, ref = _table_location(src, table)
    c1, r1, c2, r2 = _split(ref)
    new_col = _next_col(c2)
    changes = [{"sheet": sheet, "cell": f"{new_col}{r1 + 1}", "formula": "1+1"}]
    if with_header:
        changes.insert(0, {"sheet": sheet, "cell": f"{new_col}{r1}", "value": "ZedNew"})

    plan = plan_table_extensions(src, changes)
    assert len(plan) == 1 and plan[0]["table"] == table and plan[0]["old_ref"] == ref
    assert plan[0]["new_columns"][0]["auto_named"] is (not with_header)

    out = tmp_path / "out.xlsx"
    edit(src, out, changes, extend_tables=True)

    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        old_t, new_t = a.read(part).decode(), b.read(part).decode()
        n_old = len(re.findall(r"<tableColumn\b", old_t))
        assert re.search(rf'<table\b[^>]*\bref="{c1}{r1}:{new_col}{r2}"', new_t)
        assert len(re.findall(r"<tableColumn\b", new_t)) == n_old + 1
        assert f'<tableColumns count="{n_old + 1}"' in new_t
        added = re.findall(r'<tableColumn\b[^>]*\bname="([^"]*)"', new_t)[-1]
        assert added == "ZedNew" if with_header else re.fullmatch(r"Column\d+", added)
        changed = {n for n in a.namelist() if n in b.namelist() and a.read(n) != b.read(n)}
        assert {sheet_part, part} <= changed
        # Besides the sheet, the table part and sharedStrings, only edit()'s standing
        # calc-chain bookkeeping (fullCalcOnLoad, dropped calcChain) may differ.
        bookkeeping = {"xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
        assert changed <= {sheet_part, part, "xl/sharedStrings.xml"} | bookkeeping
        assert set(a.namelist()) - set(b.namelist()) <= {"xl/calcChain.xml"}

    _assert_valid_package(out)


def test_query_table_is_never_extended_and_reported(tmp_path):
    src = _copy(tmp_path, "structured-refs-lookup-tool.xlsx")
    part, sheet, sheet_part, ref = _table_location(src, "Table_Query_from_RDS2")
    c1, r1, c2, r2 = _split(ref)
    changes = [
        {"sheet": sheet, "cell": f"{_next_col(c2)}{r1}", "value": "ZedNew"},
        {"sheet": sheet, "cell": f"{_next_col(c2)}{r1 + 1}", "value": 1},
    ]
    plan = plan_table_extensions(src, changes)
    assert len(plan) == 1 and plan[0]["table"] == "Table_Query_from_RDS2"
    assert plan[0]["new_ref"] is None and plan[0]["new_columns"] == []
    assert "query table" in plan[0]["skipped_reason"]

    out = tmp_path / "out.xlsx"
    edit(src, out, changes, extend_tables=True)
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(out) as b:
        assert a.read(part) == b.read(part)
        assert a.read(sheet_part) != b.read(sheet_part)  # the cells themselves are still written
    _assert_valid_package(out)


def test_query_table_detected_by_relationship_alone(tmp_path):
    pkg = _pkg(
        tmp_path,
        extra_parts={
            "xl/tables/_rels/table1.xml.rels": (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" '
                'Target="../queryTables/queryTable1.xml"/></Relationships>'
            )
        },
    )
    plan = plan_table_extensions(pkg, [{"sheet": "Data", "cell": "D1", "value": "X"}, {"sheet": "Data", "cell": "D2", "value": 1}])
    assert plan[0]["new_ref"] is None and "skipped_reason" in plan[0]
