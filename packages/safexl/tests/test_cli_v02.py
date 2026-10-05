"""safexl CLI v0.2: verify, view, salvage, guard, resize-table, and edit's default table extension.

Every test goes through ``main()``. Exit codes: verify 0 intact / 10 rewritten / 11 damaged / 12 invalid;
salvage 0 salvaged or nothing_to_do / 10 partial / 11 unsupported; guard check 0 clean / 10 flagged.
Errors stay 1 (SafeXLError, missing input) and 2 (usage).
"""
from __future__ import annotations

import json
import re
import zipfile
import shutil
import subprocess
import sys

import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl.chart import BarChart, Reference  # noqa: E402

from safexl_mcp.cli import main

from _pkg import make_package, read_part, row, sheet_xml
from _util import PACKAGE_DIR, fixture_path, first_sheet_name, sha256_of

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
TBL = "xl/tables/table1.xml"


def _istr(ref: str, text: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>'


def _table_pkg(tmp_path, *, d1_header=True):
    """Sheet "Data" with table Sales over A1:C4 and a header cell next to it (D1 = Delta, or empty)."""
    entries = "".join(f'<tableColumn id="{i}" name="{c}"/>' for i, c in enumerate(("Alpha", "Beta", "Gamma"), 1))
    table = (
        f'<table xmlns="{NS}" id="1" name="Sales" displayName="Sales" ref="A1:C4">'
        f'<autoFilter ref="A1:C4"/><tableColumns count="3">{entries}</tableColumns></table>'
    )
    head = [_istr("A1", "Alpha"), _istr("B1", "Beta"), _istr("C1", "Gamma")]
    if d1_header:
        head.append(_istr("D1", "Delta"))
    body = row(1, *head) + "".join(
        row(n, *(f'<c r="{c}{n}"><v>{n}</v></c>' for c in "ABC")) for n in (2, 3, 4)
    )
    return make_package(
        tmp_path / "src.xlsx",
        [("Data", sheet_xml(body))],
        extra_parts={TBL: table},
        sheet_rels={1: [("table", "../tables/table1.xml")]},
    )


def _simple_pkg(tmp_path):
    body = row(1, '<c r="A1"><v>1</v></c>') + row(2, '<c r="A2"><v>2</v></c>')
    return make_package(tmp_path / "simple.xlsx", [("Sheet1", sheet_xml(body))])


def _original(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["name", "qty"])
    for i in range(2, 6):
        ws.append([f"item{i}", i * 10])
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    ws.add_chart(chart, "F2")
    wb.save(path)
    return path


def _strip_chart(path):
    """Make a re-save lose the chart, the way a lossy writer would."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            if item.filename.startswith(("xl/charts/", "xl/drawings/")):
                continue
            data = zi.read(item.filename)
            if item.filename.endswith(".rels") and b"drawing" in data:
                data = b"".join(l for l in data.replace(b"/>", b"/>\n").splitlines(True) if b"drawing" not in l)
            if item.filename == "[Content_Types].xml":
                data = b"".join(
                    l for l in data.replace(b"/>", b"/>\n").splitlines(True) if b"chart" not in l and b"drawing" not in l
                )
            if item.filename.startswith("xl/worksheets/sheet") and b"<drawing" in data:
                data = re.sub(rb"<drawing[^>]*/>", b"", data)
            zo.writestr(item, data)
    tmp.replace(path)


def _lossy(src, dst, mutate=None):
    """A re-save that loses the chart (so verify calls it damaged)."""
    _resave(src, dst, mutate)
    _strip_chart(dst)
    return dst


def _resave(src, dst, mutate=None):
    wb = openpyxl.load_workbook(src)
    if mutate:
        mutate(wb)
    wb.save(dst)
    return dst


def _run(argv, capsys):
    rc = main(argv)
    cap = capsys.readouterr()
    return rc, cap.out, cap.err


# --- help --------------------------------------------------------------------------------------


def test_help_lists_new_subcommands():
    out = subprocess.run(
        [sys.executable, "-m", "safexl_mcp.cli", "--help"], cwd=PACKAGE_DIR, capture_output=True, text=True
    )
    assert out.returncode == 0
    for name in ("verify", "view", "salvage", "guard", "resize-table"):
        assert name in out.stdout


# --- verify ------------------------------------------------------------------------------------


def test_verify_intact_exit_0(tmp_path, capsys):
    src = fixture_path("WithChart.xlsx")
    out = tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(out), "--set", f"{first_sheet_name(src)}!Z90=1"]) == 0
    capsys.readouterr()
    rc, text, _ = _run(["verify", str(src), str(out)], capsys)
    assert rc == 0
    assert text.startswith("safexl verify: intact") and "parts:" in text


def test_verify_rewritten_exit_1(tmp_path, capsys):
    src = fixture_path("WithChart.xlsx")
    out = tmp_path / "rt.xlsx"
    _resave(src, out)
    rc, text, _ = _run(["verify", str(src), str(out)], capsys)
    assert rc == 10
    assert "safexl verify: rewritten" in text and "rewritten by another tool" in text


def test_verify_damaged_exit_2_names_lost_feature(tmp_path, capsys):
    src = fixture_path("SimpleMacro.xlsm")
    out = tmp_path / "rt.xlsm"
    _resave(src, out)  # no keep_vba: the macro project is dropped
    rc, text, _ = _run(["verify", str(src), str(out)], capsys)
    assert rc == 11
    assert "safexl verify: damaged" in text and "lost vba: 1 -> 0" in text


def test_verify_invalid_exit_3(tmp_path, capsys):
    src = fixture_path("WithChart.xlsx")
    bad = tmp_path / "bad.xlsx"
    bad.write_text("not a workbook")
    rc, text, _ = _run(["verify", str(src), str(bad)], capsys)
    assert rc == 12 and "safexl verify: invalid" in text


def test_verify_json_parses_and_keeps_exit_code(tmp_path, capsys):
    src = fixture_path("SimpleMacro.xlsm")
    out = tmp_path / "rt.xlsm"
    _resave(src, out)
    rc, text, _ = _run(["verify", str(src), str(out), "--json"], capsys)
    data = json.loads(text)
    assert rc == 11 and data["verdict"] == "damaged"
    assert {"lost_features", "parts", "reasons", "before_sha256", "after_sha256"} <= set(data)


def test_verify_missing_input_is_error_1_not_a_result(tmp_path, capsys):
    src = fixture_path("WithChart.xlsx")
    rc, _, err = _run(["verify", str(src), str(tmp_path / "gone.xlsx")], capsys)
    assert rc == 1 and "safexl verify:" in err
    rc, _, err = _run(["verify", str(tmp_path / "gone.xlsx"), str(src)], capsys)
    assert rc == 1 and "safexl verify:" in err


# --- view --------------------------------------------------------------------------------------


def test_view_text_default_and_range(tmp_path, capsys):
    p = _simple_pkg(tmp_path)
    rc, text, _ = _run(["view", str(p)], capsys)
    assert rc == 0 and "A1  n  1" in text and "A2  n  2" in text and text.startswith("# Sheet1")
    rc, text, _ = _run(["view", str(p), "--range", "A2"], capsys)
    assert "A2  n  2" in text and "A1  n" not in text


def test_view_find_and_max_cells_and_json(tmp_path, capsys):
    p = _simple_pkg(tmp_path)
    rc, text, _ = _run(["view", str(p), "--find", "2", "--sheet", "Sheet1"], capsys)
    assert rc == 0 and "Sheet1!A2" in text
    rc, text, _ = _run(["view", str(p), "--max-cells", "1", "--json"], capsys)
    data = json.loads(text)
    assert rc == 0 and len(data["cells"]) == 1 and data["truncated"] is True and data["next_range"] == "A2:A2"


def test_view_errors_exit_1(tmp_path, capsys):
    p = _simple_pkg(tmp_path)
    rc, _, err = _run(["view", str(p), "--sheet", "Nope"], capsys)
    assert rc == 1 and "safexl view:" in err
    rc, _, err = _run(["view", str(tmp_path / "missing.xlsx")], capsys)
    assert rc == 1 and "safexl view:" in err


# --- salvage -----------------------------------------------------------------------------------


def test_salvage_salvaged_exit_0(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")
    dam = _lossy(orig, tmp_path / "dam.xlsx", lambda wb: wb["Data"].__setitem__("B3", 999))
    out = tmp_path / "out.xlsx"
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out)], capsys)
    assert rc == 0 and out.exists()
    assert "safexl salvage: salvaged" in text and "applied: 1, skipped: 0" in text and "Data!B3" in text
    assert openpyxl.load_workbook(out)["Data"]["B3"].value == 999


def test_salvage_nothing_to_do_exit_0_and_no_output(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")
    dam = _resave(orig, tmp_path / "dam.xlsx")
    out = tmp_path / "out.xlsx"
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out)], capsys)
    assert rc == 0 and not out.exists() and "nothing_to_do" in text


def test_salvage_partial_exit_1_lists_reason(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")

    def mutate(wb):
        wb["Data"]["B3"] = 999
        wb["Data"]["B4"] = None  # a cleared cell cannot be re-applied

    dam = _lossy(orig, tmp_path / "dam.xlsx", mutate)
    out = tmp_path / "out.xlsx"
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out)], capsys)
    assert rc == 10 and out.exists()
    assert "safexl salvage: partial" in text and "applied: 1, skipped: 1" in text
    assert "skipped Data!B4: cell cleared" in text


def test_salvage_unsupported_exit_2_no_output_and_structural(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")
    dam = _lossy(orig, tmp_path / "dam.xlsx", lambda wb: wb["Data"].insert_cols(2))
    out = tmp_path / "out.xlsx"
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out)], capsys)
    assert rc == 11 and not out.exists()
    assert "safexl salvage: unsupported" in text and "structural: columns_shifted" in text and "NOT handled" in text


def test_salvage_json_and_missing_file(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")
    dam = _lossy(orig, tmp_path / "dam.xlsx", lambda wb: wb["Data"].__setitem__("B3", 5))
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(tmp_path / "o.xlsx"), "--json"], capsys)
    data = json.loads(text)
    assert rc == 0 and data["status"] == "salvaged" and data["applied"][0]["ref"] == "B3"
    rc, _, err = _run(["salvage", str(orig), str(tmp_path / "gone.xlsx"), "-o", str(tmp_path / "o2.xlsx")], capsys)
    assert rc == 1 and "safexl salvage:" in err


def test_salvage_force_rebuilds_a_rewritten_file_and_verify_calls_it_intact(tmp_path, capsys):
    orig = _original(tmp_path / "orig.xlsx")
    dam = _resave(orig, tmp_path / "dam.xlsx", lambda wb: wb["Data"].__setitem__("B3", 5))  # nothing lost
    out = tmp_path / "out.xlsx"
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out)], capsys)
    assert rc == 0 and "nothing_to_do" in text and not out.exists()
    rc, text, _ = _run(["salvage", str(orig), str(dam), "-o", str(out), "--force"], capsys)
    assert rc == 0 and "salvaged" in text and out.exists()
    rc, text, _ = _run(["verify", str(orig), str(out)], capsys)
    assert rc == 0 and "intact" in text


def test_salvage_requires_output():
    with pytest.raises(SystemExit) as e:
        main(["salvage", "a.xlsx", "b.xlsx"])
    assert e.value.code == 2


# --- guard -------------------------------------------------------------------------------------


def _guard_setup(tmp_path, capsys, name="WithChart.xlsx"):
    wbk = tmp_path / name
    shutil.copy(fixture_path(name), wbk)
    snap = tmp_path / "snap"
    rc, text, _ = _run(["guard", "snapshot", "--dir", str(snap), str(wbk)], capsys)
    assert rc == 0 and "snapshot of 1 file(s)" in text
    return wbk, snap, sha256_of(wbk.read_bytes())


def test_guard_clean_check_exit_0(tmp_path, capsys):
    _, snap, _ = _guard_setup(tmp_path, capsys)
    rc, text, _ = _run(["guard", "check", "--dir", str(snap)], capsys)
    assert rc == 0 and "fine" in text
    rc, _, _ = _run(["guard", "check", "--dir", str(snap), "--restore"], capsys)
    assert rc == 0


def test_guard_check_flags_without_restoring_exit_1(tmp_path, capsys):
    wbk, snap, orig = _guard_setup(tmp_path, capsys)
    _resave(wbk, wbk)
    rc, text, _ = _run(["guard", "check", "--dir", str(snap)], capsys)
    assert rc == 10
    assert text.count("\n") == 1 and "NOT restored" in text  # one event, one line, printed as is
    assert sha256_of(wbk.read_bytes()) != orig


def test_guard_restore_exit_1_and_file_restored(tmp_path, capsys):
    wbk, snap, orig = _guard_setup(tmp_path, capsys)
    _resave(wbk, wbk)
    rc, text, _ = _run(["guard", "check", "--dir", str(snap), "--restore"], capsys)
    assert rc == 10 and "restored" in text and "NOT restored" not in text
    assert sha256_of(wbk.read_bytes()) == orig
    rc, _, _ = _run(["guard", "check", "--dir", str(snap)], capsys)
    assert rc == 0


def test_guard_lenient_leaves_rewritten_file_but_restores_damaged(tmp_path, capsys):
    wbk, snap, orig = _guard_setup(tmp_path, capsys)
    _resave(wbk, wbk)  # rewritten, nothing lost
    rc, text, _ = _run(["guard", "check", "--dir", str(snap), "--restore", "--lenient"], capsys)
    assert rc == 10 and "NOT restored" in text and sha256_of(wbk.read_bytes()) != orig

    (tmp_path / "m").mkdir()
    macro, snap2, orig2 = _guard_setup(tmp_path / "m", capsys, "SimpleMacro.xlsm")
    _resave(macro, macro)  # no keep_vba: the macro project is lost
    rc, text, _ = _run(["guard", "check", "--dir", str(snap2), "--restore", "--lenient"], capsys)
    assert rc == 10 and "lost vba" in text and sha256_of(macro.read_bytes()) == orig2


def test_guard_json_events_parse(tmp_path, capsys):
    wbk, snap, _ = _guard_setup(tmp_path, capsys)
    _resave(wbk, wbk)
    rc, text, _ = _run(["guard", "check", "--dir", str(snap), "--json"], capsys)
    events = json.loads(text)
    assert rc == 10 and events[0]["file"] == wbk.name and events[0]["kind"] == "rewritten"
    shutil.copy(fixture_path("WithChart.xlsx"), wbk)
    rc, text, _ = _run(["guard", "check", "--dir", str(snap), "--json"], capsys)
    assert rc == 0 and json.loads(text) == []


def test_guard_snapshot_missing_path_exit_1(tmp_path, capsys):
    rc, _, err = _run(["guard", "snapshot", "--dir", str(tmp_path / "s"), str(tmp_path / "nope.xlsx")], capsys)
    assert rc == 1 and "not found" in err


def test_guard_requires_subcommand_and_dir():
    for argv in (["guard"], ["guard", "check"], ["guard", "snapshot", "a.xlsx"]):
        with pytest.raises(SystemExit) as e:
            main(argv)
        assert e.value.code == 2


# --- resize-table ------------------------------------------------------------------------------


def test_resize_table_success(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, text, _ = _run(["resize-table", str(src), str(dst), "--table", "Sales", "--ref", "A1:D4", "--sheet", "Data"], capsys)
    assert rc == 0 and "resize-table" in text and str(dst) in text
    t = read_part(dst, TBL)
    assert re.search(r'<table\b[^>]*\bref="A1:D4"', t) and '<tableColumn id="4" name="Delta"/>' in t


def test_resize_table_unknown_table_exit_1_and_no_dst(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, _, err = _run(["resize-table", str(src), str(dst), "--table", "Nope", "--ref", "A1:D4"], capsys)
    assert rc == 1 and "safexl resize-table:" in err and not dst.exists()


def test_resize_table_requires_table_and_ref():
    with pytest.raises(SystemExit) as e:
        main(["resize-table", "a.xlsx", "b.xlsx", "--table", "Sales"])
    assert e.value.code == 2


# --- edit: table extension ---------------------------------------------------------------------


def test_edit_extends_table_by_default_and_reports_it(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, text, _ = _run(["edit", str(src), str(dst), "--set", "Data!D2=10", "--set", "Data!D3=20", "--set", "Data!D4=30"], capsys)
    assert rc == 0
    assert 'table Sales extended A1:C4 -> A1:D4 (new column "Delta")' in text
    assert "auto-named" not in text
    t = read_part(dst, TBL)
    assert re.search(r'<table\b[^>]*\bref="A1:D4"', t) and 'name="Delta"' in t


def test_edit_no_extend_tables_leaves_table_alone(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, text, _ = _run(["edit", str(src), str(dst), "--set", "Data!D2=10", "--no-extend-tables"], capsys)
    assert rc == 0 and "extended" not in text
    assert read_part(dst, TBL) == read_part(src, TBL)


def test_edit_flags_auto_named_column(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path, d1_header=False), tmp_path / "out.xlsx"
    rc, text, _ = _run(["edit", str(src), str(dst), "--set", "Data!D2=10"], capsys)
    assert rc == 0
    assert 'table Sales extended A1:C4 -> A1:D4 (new column "Column1")' in text
    assert "auto-named" in text and "D1" in text


def test_edit_without_adjacent_table_prints_no_table_line(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, text, _ = _run(["edit", str(src), str(dst), "--set", "Data!H20=1"], capsys)
    assert rc == 0 and "table " not in text


def test_edit_error_still_exit_1(tmp_path, capsys):
    src, dst = _table_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, _, err = _run(["edit", str(src), str(dst), "--set", "Nope!A1=1"], capsys)
    assert rc == 1 and "safexl edit:" in err and not dst.exists()


# --- insert-cells ------------------------------------------------------------------------------


def _side_by_side_pkg(tmp_path):
    """Two lists side by side: A2:A3 and C2:C3."""
    body = (
        row(1, '<c r="A1"><v>1</v></c>', '<c r="C1"><v>10</v></c>')
        + row(2, '<c r="A2"><v>2</v></c>', '<c r="C2"><v>20</v></c>')
        + row(3, '<c r="A3"><v>3</v></c>', '<c r="C3"><v>30</v></c>')
    )
    return make_package(tmp_path / "side.xlsx", [("Sheet1", sheet_xml(body))])


def test_insert_cells_shifts_only_the_column_span(tmp_path, capsys):
    src, dst = _side_by_side_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, text, _ = _run(["insert-cells", str(src), str(dst), "--sheet", "Sheet1", "--range", "A2:A2"], capsys)
    assert rc == 0 and "insert-cells" in text and str(dst) in text
    rc, text, _ = _run(["view", str(dst), "--json"], capsys)
    cells = {c["ref"]: c["value"] for c in json.loads(text)["cells"]}
    assert cells["A1"] == 1 and cells["A3"] == 2 and cells["A4"] == 3  # column A shifted down
    assert cells["C2"] == 20 and cells["C3"] == 30 and "C4" not in cells  # column C stayed


def test_insert_cells_errors_exit_1_with_reason_and_no_dst(tmp_path, capsys):
    src, dst = _side_by_side_pkg(tmp_path), tmp_path / "out.xlsx"
    rc, _, err = _run(["insert-cells", str(src), str(dst), "--sheet", "Nope", "--range", "A2:A2"], capsys)
    assert rc == 1 and "safexl insert-cells:" in err and not dst.exists()
    rc, _, err = _run(["insert-cells", str(src), str(dst), "--sheet", "Sheet1", "--range", "bogus"], capsys)
    assert rc == 1 and "invalid cell_range" in err and not dst.exists()


def test_insert_cells_help_says_when_to_use_it(capsys):
    with pytest.raises(SystemExit) as e:
        main(["insert-cells", "--help"])
    assert e.value.code == 0
    out = " ".join(capsys.readouterr().out.split())
    assert "side by side" in out and "insert-rows" in out


def test_insert_cells_requires_sheet_and_range():
    with pytest.raises(SystemExit) as e:
        main(["insert-cells", "a.xlsx", "b.xlsx", "--sheet", "S"])
    assert e.value.code == 2


# --- closed stdout -----------------------------------------------------------------------------


def test_closed_pipe_exits_quietly_with_real_code(tmp_path):
    body = "".join(row(n, f'<c r="A{n}"><v>{n}</v></c>') for n in range(1, 6001))
    p = make_package(tmp_path / "big.xlsx", [("Sheet1", sheet_xml(body))])
    proc = subprocess.Popen(
        [sys.executable, "-m", "safexl_mcp.cli", "view", str(p), "--max-cells", "6000"],
        cwd=PACKAGE_DIR, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    )
    first = proc.stdout.readline()
    proc.stdout.close()  # what `| head -1` does
    err = proc.stderr.read()
    rc = proc.wait(timeout=30)
    assert first.startswith(b"# Sheet1")
    assert rc == 0 and err == b""


# --- edit --sheet ------------------------------------------------------------------------------


def test_edit_sheet_default_equals_prefixed_form(tmp_path, capsys):
    src = _simple_pkg(tmp_path)
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    assert main(["edit", str(src), str(a), "--sheet", "Sheet1", "--set", "B2=5", "--text", "C2=007",
                 "--formula", "D2=B2*2"]) == 0
    assert main(["edit", str(src), str(b), "--set", "Sheet1!B2=5", "--text", "Sheet1!C2=007",
                 "--formula", "Sheet1!D2=B2*2"]) == 0
    capsys.readouterr()
    assert read_part(a, "xl/worksheets/sheet1.xml") == read_part(b, "xl/worksheets/sheet1.xml")


def test_edit_explicit_prefix_wins_over_sheet(tmp_path, capsys):
    body = row(1, '<c r="A1"><v>1</v></c>')
    src = make_package(tmp_path / "two.xlsx", [("One", sheet_xml(body)), ("Two", sheet_xml(body))])
    dst = tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(dst), "--sheet", "One", "--set", "A1=5", "--set", "Two!A1=7"]) == 0
    capsys.readouterr()
    rc, text, _ = _run(["view", str(dst), "--sheet", "One", "--json"], capsys)
    assert json.loads(text)["cells"][0]["value"] == 5
    rc, text, _ = _run(["view", str(dst), "--sheet", "Two", "--json"], capsys)
    assert json.loads(text)["cells"][0]["value"] == 7


def test_edit_sheet_default_keeps_bang_in_values(tmp_path, capsys):
    src, dst = _simple_pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(dst), "--sheet", "Sheet1", "--set", "B2=hello!"]) == 0
    capsys.readouterr()
    rc, text, _ = _run(["view", str(dst), "--find", "hello", "--json"], capsys)
    assert json.loads(text)["cells"][0]["value"] == "hello!"


def test_edit_without_sheet_and_without_prefix_still_errors(tmp_path):
    src, dst = _simple_pkg(tmp_path), tmp_path / "out.xlsx"
    with pytest.raises(SystemExit) as e:
        main(["edit", str(src), str(dst), "--set", "B2=5"])
    assert "missing '!'" in str(e.value)
    assert not dst.exists()


# --- receipt surface: summary, added risk, --strict, refusals -----------------------------------

_PREFIXED_SHEET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<x:worksheet xmlns:x="{NS}"><x:sheetData>'
    '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>'
)


def _fake_receipt(**over):
    r = {
        "verdict": "intact", "summary": "No tracked feature was lost; 1 external link was added.",
        "added_risk": [{"risk": "external_link_added", "detail": "xl/externalLinks/externalLink1.xml"}],
        "cell_changes": {"Sheet1": {"cells_changed": 3, "rows_inserted": 1, "positional": True},
                         "Quiet": {"cells_changed": 0, "rows_inserted": 0}},
        "lost_features": [], "reasons": [], "parts": {"unchanged": 4, "changed": [], "added": [], "removed": []},
    }
    r.update(over)
    return r


def test_verify_text_shows_summary_added_risk_and_sheet_changes(tmp_path, capsys):
    src = _simple_pkg(tmp_path)
    out = tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(out), "--set", "Sheet1!A1=9"]) == 0
    capsys.readouterr()
    rc, text, _ = _run(["verify", str(src), str(out)], capsys)
    lines = text.splitlines()
    assert rc == 0 and lines[0] == "safexl verify: intact"
    assert "1 cell changed" in lines[1]
    assert "  Sheet1: 1 cell changed" in text


def test_verify_text_and_strict_exit_13_on_added_risk(tmp_path, capsys, monkeypatch):
    src = _simple_pkg(tmp_path)
    monkeypatch.setattr("safexl_mcp.cli.verify", lambda a, b: _fake_receipt())
    rc, text, _ = _run(["verify", str(src), str(src)], capsys)
    assert rc == 0  # an added risk does not change the default verdict or exit code
    assert "  added: external_link_added: xl/externalLinks/externalLink1.xml" in text
    assert "  Sheet1: 3 cells changed, 1 row inserted (positional)" in text
    assert "Quiet" not in text
    rc, _, _ = _run(["verify", str(src), str(src), "--strict"], capsys)
    assert rc == 13


def test_verify_strict_keeps_other_verdicts_and_clean_intact(tmp_path, capsys, monkeypatch):
    src = _simple_pkg(tmp_path)
    monkeypatch.setattr("safexl_mcp.cli.verify", lambda a, b: _fake_receipt(verdict="damaged"))
    assert _run(["verify", str(src), str(src), "--strict"], capsys)[0] == 11
    monkeypatch.setattr("safexl_mcp.cli.verify", lambda a, b: _fake_receipt(added_risk=[]))
    assert _run(["verify", str(src), str(src), "--strict"], capsys)[0] == 0


def test_edit_namespace_prefix_refusal_exit_1_no_output(tmp_path, capsys):
    src = make_package(tmp_path / "p.xlsx", [("Sheet1", _PREFIXED_SHEET)])
    out = tmp_path / "out.xlsx"
    rc, _, err = _run(["edit", str(src), str(out), "--set", "Sheet1!B2=1"], capsys)
    assert rc == 1 and "namespace prefix" in err and not out.exists()
    rc, _, err = _run(["insert-rows", str(src), str(out), "--sheet", "Sheet1", "--row", "1"], capsys)
    assert rc == 1 and "namespace prefix" in err and not out.exists()


def test_edit_read_back_failure_exit_1_no_output(tmp_path, capsys, monkeypatch):
    from safexl_mcp.engine import SafeXLError

    src = _simple_pkg(tmp_path)
    out = tmp_path / "out.xlsx"

    def boom(*a, **k):
        raise SafeXLError("read-back failed at Sheet1!A1: value differs")

    monkeypatch.setattr("safexl_mcp.cli.edit", boom)
    rc, _, err = _run(["edit", str(src), str(out), "--set", "Sheet1!A1=9"], capsys)
    assert rc == 1 and "read-back failed at Sheet1!A1" in err and not out.exists()
