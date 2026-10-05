"""safexl_mcp.cli -- exit codes and valid --json output for edit/inspect/diff.

Most tests call cli.main(argv) -> int in-process (this is exactly the value
`if __name__ == "__main__": raise SystemExit(main())` turns into the process
exit code, so asserting on the return value IS asserting the exit code). One
end-to-end subprocess test exercises the real `python -m safexl_mcp.cli`
entry point for belt-and-suspenders coverage of the packaging/entry point
itself.
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile

import pytest

from safexl_mcp.cli import main

from _util import PACKAGE_DIR, build_minimal_workbook, fixture_path


# ---------------------------------------------------------------------------
# edit
# ---------------------------------------------------------------------------


def test_edit_success_exit_code_zero(tmp_path, capsys):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    rc = main(["edit", str(src), str(dst), "--set", "Sheet1!B2=42"])
    assert rc == 0
    assert dst.is_file()
    out = capsys.readouterr().out
    assert "wrote" in out
    assert str(dst) in out


def test_edit_no_set_or_formula_exit_code_two(tmp_path, capsys):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    rc = main(["edit", str(src), str(dst)])
    assert rc == 2
    assert not dst.exists()
    err = capsys.readouterr().err
    assert "at least one --set, --text or --formula" in err


def test_edit_safexlerror_exit_code_one_and_dst_not_created(tmp_path, capsys):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    rc = main(["edit", str(src), str(dst), "--set", "NoSuchSheet!A1=1"])
    assert rc == 1
    assert not dst.exists()
    err = capsys.readouterr().err
    assert "safexl edit:" in err


def test_edit_set_expr_missing_bang_raises_systemexit_two(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SystemExit):
        main(["edit", str(src), str(dst), "--set", "A1=1"])  # missing '!'


def test_edit_set_expr_missing_equals_raises_systemexit(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SystemExit):
        main(["edit", str(src), str(dst), "--set", "Sheet1!A1"])  # missing '='


def test_edit_set_expr_empty_sheet_after_strip_raises_systemexit(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SystemExit):
        main(["edit", str(src), str(dst), "--set", "!A1=1"])  # empty sheet name


def test_edit_formula_expr_empty_after_strip_raises_systemexit(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SystemExit):
        main(["edit", str(src), str(dst), "--formula", "Sheet1!A1=   "])  # blank formula


def test_edit_formula_flag_and_value_coercion(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    rc = main(
        [
            "edit",
            str(src),
            str(dst),
            "--set",
            "Sheet1!A1=TRUE",
            "--set",
            "Sheet1!A2=false",
            "--set",
            "Sheet1!A3=42",
            "--set",
            "Sheet1!A4=3.5",
            "--set",
            "Sheet1!A5=hello",
            "--formula",
            "Sheet1!A6==A1",
        ]
    )
    assert rc == 0
    with zipfile.ZipFile(dst) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<c r="A1" t="b"><v>1</v></c>' in xml
    assert '<c r="A2" t="b"><v>0</v></c>' in xml
    assert '<c r="A3"><v>42</v></c>' in xml
    assert '<c r="A4"><v>3.5</v></c>' in xml
    assert '<c r="A5" t="inlineStr"><is><t xml:space="preserve">hello</t></is></c>' in xml
    # "Sheet1!A6==A1": the CLI's cell/payload split on the first '=' consumes
    # one, leaving formula text "=A1"; _normalize_formula then strips its own
    # single leading '=' -- so both are gone and "A1" is what's stored.
    assert '<c r="A6"><f>A1</f></c>' in xml


# ---------------------------------------------------------------------------
# inspect
# ---------------------------------------------------------------------------


def test_inspect_json_exit_code_zero_and_valid_json():
    path = fixture_path("WithChart.xlsx")
    rc = main(["inspect", str(path), "--json"])
    assert rc == 0


def test_inspect_json_output_is_valid_and_matches_library_call(capsys):
    from safexl_mcp import inspect_workbook

    path = fixture_path("WithChart.xlsx")
    rc = main(["inspect", str(path), "--json"])
    assert rc == 0
    out = capsys.readouterr().out
    data = json.loads(out)  # must parse as valid JSON
    assert data == inspect_workbook(path).to_dict()
    assert data["features"]["charts"] >= 1


def test_inspect_missing_file_exit_code_one(tmp_path, capsys):
    rc = main(["inspect", str(tmp_path / "does-not-exist.xlsx")])
    assert rc == 1
    err = capsys.readouterr().err
    assert "safexl inspect:" in err


def test_inspect_text_mode_is_not_valid_json(capsys):
    path = fixture_path("WithChart.xlsx")
    rc = main(["inspect", str(path)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "safexl inspect:" in out
    with pytest.raises(json.JSONDecodeError):
        json.loads(out)


def test_inspect_text_mode_with_vba_present_prints_macro_scan_line(capsys):
    path = fixture_path("SimpleMacro.xlsm")
    rc = main(["inspect", str(path)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "vba: present" in out
    assert "macro scan:" in out


def test_inspect_text_mode_with_defined_names_prints_them(capsys):
    path = fixture_path("formulas_definednames.xlsx")
    rc = main(["inspect", str(path)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "defined names:" in out


# ---------------------------------------------------------------------------
# diff
# ---------------------------------------------------------------------------


def test_diff_json_exit_code_zero_and_valid_json(tmp_path, capsys):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(dst), "--set", "Sheet1!B2=1"]) == 0
    capsys.readouterr()  # discard edit's own stdout

    rc = main(["diff", str(src), str(dst), "--json"])
    assert rc == 0
    out = capsys.readouterr().out
    data = json.loads(out)
    assert data["path_before"] == str(src)
    assert data["path_after"] == str(dst)
    assert len(data["sheets"]) == 1
    assert data["sheets"][0]["added_cells"][0]["cell"] == "B2"


def test_diff_missing_file_exit_code_one(tmp_path, capsys):
    rc = main(["diff", str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")])
    assert rc == 1
    err = capsys.readouterr().err
    assert "safexl diff:" in err


def test_diff_identical_files_json_shows_no_changes(capsys):
    path = fixture_path("WithChart.xlsx")
    rc = main(["diff", str(path), str(path), "--json"])
    assert rc == 0
    out = capsys.readouterr().out
    data = json.loads(out)
    assert data["sheets"] == []
    assert data["parts"]["changed"] == []


def test_diff_text_mode_with_changes_prints_cell_and_part_lines(tmp_path, capsys):
    src = build_minimal_workbook(tmp_path / "src.xlsx", cells={"A1": '<c r="A1"><v>1</v></c>'})
    dst = tmp_path / "out.xlsx"
    assert main(["edit", str(src), str(dst), "--set", "Sheet1!A1=99", "--set", "Sheet1!B1=new"]) == 0
    capsys.readouterr()

    rc = main(["diff", str(src), str(dst)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "safexl diff:" in out
    assert "sheet 'Sheet1'" in out
    assert "changed A1:" in out
    assert "added   B1:" in out
    assert "OOXML parts:" in out
    assert "changed: xl/worksheets/sheet1.xml" in out


def test_diff_text_mode_no_changes_prints_explicit_message(capsys):
    path = fixture_path("WithChart.xlsx")
    rc = main(["diff", str(path), str(path)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "no cell changes in any common sheet" in out


# ---------------------------------------------------------------------------
# argparse plumbing
# ---------------------------------------------------------------------------


def test_no_subcommand_exits_with_code_two():
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 2


def test_unknown_subcommand_exits_nonzero():
    with pytest.raises(SystemExit):
        main(["bogus-command"])


# ---------------------------------------------------------------------------
# One real end-to-end subprocess call through `python -m safexl_mcp.cli`
# ---------------------------------------------------------------------------


def test_cli_subprocess_end_to_end_inspect_json():
    path = fixture_path("SimpleMacro.xlsm")
    result = subprocess.run(
        [sys.executable, "-m", "safexl_mcp.cli", "inspect", str(path), "--json"],
        cwd=str(PACKAGE_DIR),
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["features"]["vba"]["present"] is True


def test_cli_subprocess_end_to_end_edit_and_exit_code(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    result = subprocess.run(
        [sys.executable, "-m", "safexl_mcp.cli", "edit", str(src), str(dst), "--set", "Sheet1!A1=1"],
        cwd=str(PACKAGE_DIR),
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stderr
    assert dst.is_file()


def test_cli_subprocess_end_to_end_error_exit_code_one(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    result = subprocess.run(
        [sys.executable, "-m", "safexl_mcp.cli", "edit", str(src), str(dst), "--set", "NoSuchSheet!A1=1"],
        cwd=str(PACKAGE_DIR),
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 1
    assert not dst.exists()
    assert "safexl edit:" in result.stderr
