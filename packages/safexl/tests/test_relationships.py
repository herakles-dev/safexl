"""Relationship integrity: an edit must never leave a Relationship pointing at a
missing part. Regression for the calcChain case caught on a recorded
benchmark run (2026-09-26): calcChain.xml was dropped, its workbook.xml.rels
entry was not, and a dangling relationship is an Excel repair trigger."""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest

from _util import SKIP_REASON, fixture_path
from safexl_mcp import engine
from safexl_mcp.validate import validate_workbook

REPO = Path(__file__).resolve().parents[3]
BUDGET = REPO / "research" / "xlsx-bench" / "corpus" / "simple-monthly-budget.xlsx"


def test_edit_drops_calc_chain_relationship_with_the_part(tmp_path):
    if not BUDGET.is_file():
        pytest.skip(SKIP_REASON)
    with zipfile.ZipFile(BUDGET) as z:
        assert "xl/calcChain.xml" in z.namelist()
        assert "calcChain" in z.read("xl/_rels/workbook.xml.rels").decode()
    out = tmp_path / "out.xlsx"
    engine.edit(BUDGET, out, [{"sheet": "Simple Monthly Budget", "cell": "C12", "value": 135}])
    with zipfile.ZipFile(out) as z:
        assert "xl/calcChain.xml" not in z.namelist()
        assert "calcChain" not in z.read("xl/_rels/workbook.xml.rels").decode()
    assert validate_workbook(out).critical_count == 0


def test_validator_flags_a_dangling_relationship(tmp_path):
    if not BUDGET.is_file():
        pytest.skip(SKIP_REASON)
    bad = tmp_path / "bad.xlsx"
    with zipfile.ZipFile(BUDGET) as src, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename != "xl/calcChain.xml":
                dst.writestr(item, src.read(item.filename))
    report = validate_workbook(bad)
    assert report.critical_count >= 1
    assert any("calcChain.xml" in d and "missing part" in d for d in report.critical_details())


def test_workbook_rels_untouched_when_there_is_no_calc_chain(tmp_path):
    src = fixture_path("WithDrawing.xlsx")
    with zipfile.ZipFile(src) as z:
        if "xl/calcChain.xml" in z.namelist():
            return
        before = z.read("xl/_rels/workbook.xml.rels")
        sheet = "Sheet1"
    out = tmp_path / "out.xlsx"
    engine.edit(src, out, [{"sheet": sheet, "cell": "A1", "value": "x"}])
    with zipfile.ZipFile(out) as z:
        assert z.read("xl/_rels/workbook.xml.rels") == before
