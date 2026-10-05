"""safexl_mcp.inspect -- read-only structural + feature report.

Asserts the documented feature category for each of the 7 real
fixtures is reported: WithChart -> charts>=1,
WithDrawing -> drawing/media>=1, ExcelPivotTableSample -> pivot>=1,
DataValidations -> validations>=1, NewStyleConditionalFormattings ->
conditional formatting>=1, SimpleMacro.xlsm -> VBA present,
formulas_definednames -> formulas>=1 + defined names>=1.
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError, inspect_workbook

from _util import fixture_path


def test_inspect_with_chart_reports_at_least_one_chart():
    report = inspect_workbook(fixture_path("WithChart.xlsx")).to_dict()
    assert report["features"]["charts"] >= 1


def test_inspect_with_drawing_reports_drawing_or_media():
    report = inspect_workbook(fixture_path("WithDrawing.xlsx")).to_dict()
    f = report["features"]
    assert f["drawings"] + f["media_files"] >= 1


def test_inspect_pivot_sample_reports_pivot_table_and_cache():
    report = inspect_workbook(fixture_path("ExcelPivotTableSample.xlsx")).to_dict()
    f = report["features"]
    assert f["pivot_tables"] >= 1
    assert f["pivot_caches"] >= 1


def test_inspect_data_validations_fixture_reports_validations():
    report = inspect_workbook(fixture_path("DataValidations-49244.xlsx")).to_dict()
    f = report["features"]
    assert f["data_validation_native"] + f["data_validation_extlst"] >= 1


def test_inspect_new_style_conditional_formatting_reports_cf():
    report = inspect_workbook(fixture_path("NewStyleConditionalFormattings.xlsx")).to_dict()
    f = report["features"]
    assert f["conditional_formatting_native"] + f["conditional_formatting_extlst"] >= 1


def test_inspect_simple_macro_reports_vba_present_and_sized():
    report = inspect_workbook(fixture_path("SimpleMacro.xlsm")).to_dict()
    vba = report["features"]["vba"]
    assert vba["present"] is True
    assert (vba["size_bytes"] or 0) > 0
    # oletools may or may not be installed in the environment -- either way
    # inspect() must degrade gracefully rather than raising.
    assert vba["macro_scan"]["mode"] in ("stdlib-only", "oletools")
    assert isinstance(vba["macro_scan"]["hits"], list)


def test_inspect_formulas_definednames_reports_formulas_and_names():
    report = inspect_workbook(fixture_path("formulas_definednames.xlsx")).to_dict()
    total_formulas = sum(s["formula_count"] for s in report["sheets"])
    assert total_formulas >= 1
    assert len(report["defined_names"]) >= 1


# ---------------------------------------------------------------------------
# Negative-space assertions: features NOT present are reported as absent,
# not merely omitted / silently wrong.
# ---------------------------------------------------------------------------


def test_inspect_workbook_without_vba_reports_absent():
    report = inspect_workbook(fixture_path("WithChart.xlsx")).to_dict()
    vba = report["features"]["vba"]
    assert vba["present"] is False
    assert vba["size_bytes"] is None
    assert vba["macro_scan"]["mode"] == "n/a"
    assert vba["macro_scan"]["hits"] == []


def test_inspect_workbook_without_charts_reports_zero():
    report = inspect_workbook(fixture_path("DataValidations-49244.xlsx")).to_dict()
    assert report["features"]["charts"] == 0


# ---------------------------------------------------------------------------
# Per-sheet structural facts
# ---------------------------------------------------------------------------


def test_inspect_reports_every_sheet_in_workbook_order():
    report = inspect_workbook(fixture_path("ExcelPivotTableSample.xlsx")).to_dict()
    names = [s["name"] for s in report["sheets"]]
    assert names == ["Tabelle1", "Tabelle2", "Tabelle3"]


def test_inspect_dimension_matches_flag_consistent_with_declared_vs_actual():
    report = inspect_workbook(fixture_path("DataValidations-49244.xlsx")).to_dict()
    for sheet in report["sheets"]:
        if sheet["actual_used_range"] is None:
            continue
        expected = sheet["declared_dimension"] == sheet["actual_used_range"]
        assert sheet["dimension_matches"] == expected


def test_inspect_never_writes_to_the_file(tmp_path):
    src = fixture_path("WithChart.xlsx")
    original_bytes = src.read_bytes()
    inspect_workbook(src)
    assert src.read_bytes() == original_bytes


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


def test_inspect_missing_file_raises_safexlerror():
    with pytest.raises(SafeXLError, match="file not found"):
        inspect_workbook("/nonexistent/path/does-not-exist.xlsx")


def test_inspect_not_a_zip_raises_safexlerror(tmp_path):
    bad = tmp_path / "not-a-workbook.xlsx"
    bad.write_bytes(b"not a zip file at all")
    with pytest.raises(SafeXLError):
        inspect_workbook(bad)


def test_inspect_zip_without_workbook_xml_raises_safexlerror(tmp_path):
    import zipfile

    bad = tmp_path / "empty.xlsx"
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("hello.txt", "not an OOXML package")
    with pytest.raises(SafeXLError, match="missing xl/workbook.xml"):
        inspect_workbook(bad)
