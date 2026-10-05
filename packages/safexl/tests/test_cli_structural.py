"""safexl CLI: insert-rows, rename-sheet, add-sheet, and edit --text.

Same conventions as `edit`: SRC is never modified, DST is only written on
success, exit 0 on success, 1 on a SafeXLError (message on stderr, prefixed
with the subcommand), 2 on a usage error.
"""
from __future__ import annotations

import subprocess
import sys
import zipfile

import pytest

from safexl_mcp.cli import main

from _pkg import cell, make_package, read_part, row, sheet_xml
from _util import PACKAGE_DIR


def _pkg(tmp_path):
    body = row(1, cell("A1", v="1")) + row(2, cell("A2", v="2")) + row(3, cell("A3", f="A1+A2", v="3"))
    return make_package(tmp_path / "src.xlsx", [("Sheet1", sheet_xml(body)), ("Other", sheet_xml(""))])


def test_insert_rows_success(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    rc = main(["insert-rows", str(src), str(dst), "--sheet", "Sheet1", "--row", "2", "--count", "2"])
    assert rc == 0
    xml = read_part(dst, "xl/worksheets/sheet1.xml")
    assert '<row r="4">' in xml and "<f>A1+A4</f>" in xml
    out = capsys.readouterr().out
    assert "insert-rows" in out and str(dst) in out


def test_insert_rows_count_defaults_to_one(tmp_path):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["insert-rows", str(src), str(dst), "--sheet", "Sheet1", "--row", "1"]) == 0
    assert '<row r="2">' in read_part(dst, "xl/worksheets/sheet1.xml")


def test_insert_rows_unknown_sheet_exit_one_and_dst_not_created(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["insert-rows", str(src), str(dst), "--sheet", "Nope", "--row", "2"]) == 1
    assert not dst.exists()
    assert "safexl insert-rows:" in capsys.readouterr().err


def test_insert_rows_bad_row_exit_one(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["insert-rows", str(src), str(dst), "--sheet", "Sheet1", "--row", "0"]) == 1
    assert not dst.exists()


def test_insert_rows_requires_sheet_and_row():
    with pytest.raises(SystemExit) as e:
        main(["insert-rows", "a.xlsx", "b.xlsx", "--row", "2"])
    assert e.value.code == 2


def test_insert_rows_never_modifies_src(tmp_path):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    before = src.read_bytes()
    main(["insert-rows", str(src), str(dst), "--sheet", "Sheet1", "--row", "2"])
    assert src.read_bytes() == before


def test_rename_sheet_success(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["rename-sheet", str(src), str(dst), "Sheet1", "Budget 2026"]) == 0
    assert '<sheet name="Budget 2026"' in read_part(dst, "xl/workbook.xml")
    assert "rename-sheet" in capsys.readouterr().out


def test_rename_sheet_to_existing_name_exit_one(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["rename-sheet", str(src), str(dst), "Sheet1", "Other"]) == 1
    assert not dst.exists()
    assert "safexl rename-sheet:" in capsys.readouterr().err


def test_add_sheet_success_then_edit_into_it(tmp_path, capsys):
    src, mid, dst = _pkg(tmp_path), tmp_path / "mid.xlsx", tmp_path / "out.xlsx"
    assert main(["add-sheet", str(src), str(mid), "Lookup Check"]) == 0
    assert '<sheet name="Lookup Check"' in read_part(mid, "xl/workbook.xml")
    assert main(["edit", str(mid), str(dst), "--set", "Lookup Check!A1=hello"]) == 0
    with zipfile.ZipFile(dst) as z:
        assert any("hello" in z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith(".xml"))


def test_add_sheet_duplicate_name_exit_one(tmp_path, capsys):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    assert main(["add-sheet", str(src), str(dst), "sheet1"]) == 1
    assert not dst.exists()
    assert "safexl add-sheet:" in capsys.readouterr().err


def test_edit_text_keeps_leading_zeros_and_set_still_coerces(tmp_path):
    src, dst = _pkg(tmp_path), tmp_path / "out.xlsx"
    rc = main(["edit", str(src), str(dst), "--text", "Sheet1!B1=038", "--set", "Sheet1!B2=038"])
    assert rc == 0
    xml = read_part(dst, "xl/worksheets/sheet1.xml")
    with zipfile.ZipFile(dst) as z:
        shared = z.read("xl/sharedStrings.xml").decode() if "xl/sharedStrings.xml" in z.namelist() else ""
    assert ">038<" in xml + shared  # the text cell keeps its zero
    assert '<c r="B2"><v>38</v></c>' in xml or '<c r="B2" t="n"><v>38</v></c>' in xml  # --set coerced to a number


def test_edit_text_rejects_malformed_expression():
    with pytest.raises(SystemExit):
        main(["edit", "a.xlsx", "b.xlsx", "--text", "no-bang"])


def test_console_script_lists_new_subcommands():
    r = subprocess.run([sys.executable, "-m", "safexl_mcp.cli", "--help"], capture_output=True, text=True, cwd=PACKAGE_DIR)
    assert r.returncode == 0
    for name in ("insert-rows", "rename-sheet", "add-sheet"):
        assert name in r.stdout
