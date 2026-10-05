"""safexl_mcp.validate -- structural repair-trigger validator.

Passes all 7 real fixtures clean (zero CRITICAL); catches
deliberately corrupted files across each CRITICAL check category.
"""
from __future__ import annotations

import re
import zipfile

import pytest

from safexl_mcp import validate_bytes_ok, validate_workbook

from _util import FIXTURE_NAMES, fixture_path


# ---------------------------------------------------------------------------
# Clean real fixtures
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_validate_reports_zero_critical_on_real_fixtures(name):
    report = validate_workbook(fixture_path(name))
    assert report.critical_count == 0, report.critical_details()


def test_validate_bytes_ok_true_for_a_clean_fixture():
    ok, report = validate_bytes_ok(fixture_path("SimpleMacro.xlsm"))
    assert ok is True
    assert report.critical_count == 0


def test_validate_bytes_ok_returns_full_report_object():
    ok, report = validate_bytes_ok(fixture_path("WithChart.xlsx"))
    assert ok is True
    assert report.path == str(fixture_path("WithChart.xlsx"))


# ---------------------------------------------------------------------------
# Deliberately corrupted files -- one per CRITICAL check
# ---------------------------------------------------------------------------


def test_validate_catches_a_file_that_is_not_a_zip_at_all(tmp_path):
    bad = tmp_path / "not-a-workbook.xlsx"
    bad.write_bytes(b"this is definitely not a zip file")
    report = validate_workbook(bad)
    assert report.critical_count >= 1
    assert any(f.check == "zip" for f in report.findings)


def test_validate_catches_malformed_xml_inside_a_valid_zip(tmp_path):
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<sheetData>", b"<sheetData><unclosed-tag-should-break-xml")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "well_formed_xml" for f in report.findings)


def test_validate_catches_dangling_content_types_override(tmp_path):
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/does-not-exist.xml" '
                    b'ContentType="application/xml"/></Types>',
                )
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "content_types" for f in report.findings)


def test_validate_catches_out_of_order_rows(tmp_path):
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                rows = re.findall(r'<row\b[^>]*\br="\d+"[^>]*?(?:/>|>.*?</row>)', text, re.DOTALL)
                assert len(rows) >= 2, "fixture needs >=2 rows for this corruption"
                swapped = text.replace(rows[0] + rows[1], rows[1] + rows[0], 1)
                assert swapped != text
                data = swapped.encode("utf-8")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "row_cell_order" for f in report.findings)


def test_validate_catches_duplicate_row_number(tmp_path):
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                # Force <row r="2" ...> to duplicate r="1" (a strictly
                # ascending, non-duplicate order is required).
                text = text.replace('<row r="2"', '<row r="1"', 1)
                data = text.encode("utf-8")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "row_cell_order" for f in report.findings)


def test_validate_catches_shared_string_index_out_of_range(tmp_path):
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin:
        ss_names = [n for n in zin.namelist() if n == "xl/sharedStrings.xml"]
    assert ss_names, "fixture must have xl/sharedStrings.xml for this corruption"

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                new_text = re.sub(r'(t="s"><v>)\d+(</v>)', r"\g<1>999999\g<2>", text, count=1)
                assert new_text != text, "fixture sheet1 must contain a t=\"s\" cell"
                data = new_text.encode("utf-8")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "shared_strings" for f in report.findings)


def test_validate_catches_stale_calc_chain_entry(tmp_path):
    src = fixture_path("NewStyleConditionalFormattings.xlsx")
    with zipfile.ZipFile(src) as zin:
        assert "xl/calcChain.xml" in zin.namelist()

    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/calcChain.xml":
                text = data.decode("utf-8")
                # Point one calcChain entry at a cell ref that has no <f> --
                # a classic "part deleted/edited but calcChain still claims
                # it" repair trigger.
                text = text.replace('r="U17"', 'r="ZZ9999"', 1)
                assert 'r="ZZ9999"' in text
                data = text.encode("utf-8")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert report.critical_count >= 1
    assert any(f.check == "calc_chain" for f in report.findings)


def test_validate_warns_but_not_critical_on_dimension_undershoot(tmp_path):
    """A stale <dimension> that undershoots actual data is a WARNING, not a
    CRITICAL (it doesn't trigger Excel's repair dialog on its own) -- see
    validate.py module docstring."""
    src = fixture_path("DataValidations-49244.xlsx")
    corrupted = tmp_path / "corrupted.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(corrupted, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                new_text = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:A1"/>', text, count=1)
                assert new_text != text
                data = new_text.encode("utf-8")
            zout.writestr(item, data)

    report = validate_workbook(corrupted)
    assert any(f.check == "dimension" and f.severity == "WARNING" for f in report.findings)
