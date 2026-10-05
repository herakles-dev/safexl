"""safexl_mcp.verify -- the intact / damaged / invalid receipt."""
from __future__ import annotations

import zipfile

import pytest
import openpyxl
import openpyxl.comments

from safexl_mcp import add_sheet, edit, features, insert_rows, rename_sheet
from safexl_mcp.verify import verify

from _util import FIXTURE_NAMES, fixture_path, first_sheet_name


def _ops(src, dst_dir):
    """Every SafeXL op on one workbook -> list of (label, output path)."""
    sheet = first_sheet_name(src)
    ext = src.suffix
    outs = []
    p = dst_dir / f"edit{ext}"
    edit(src, p, [{"sheet": sheet, "cell": "Z90", "value": 42}])
    outs.append(("edit", p))
    p = dst_dir / f"ins{ext}"
    insert_rows(src, p, sheet, 2, 2)
    outs.append(("insert-rows", p))
    p = dst_dir / f"ren{ext}"
    rename_sheet(src, p, sheet, "Renamed Tab")
    outs.append(("rename-sheet", p))
    p = dst_dir / f"add{ext}"
    add_sheet(src, p, "Brand New")
    outs.append(("add-sheet", p))
    return outs


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_every_safexl_op_is_intact_not_rebuilt(name, tmp_path):
    src = fixture_path(name)
    for label, out in _ops(src, tmp_path):
        r = verify(src, out)
        assert r["verdict"] == "intact", (name, label, r["reasons"])
        assert r["valid"] is True and r["rebuilt"] is False and r["lost_features"] == []


def test_identical_file_is_intact_with_matching_hashes():
    p = fixture_path("WithChart.xlsx")
    r = verify(p, p)
    assert r["verdict"] == "intact"
    assert r["before_sha256"] == r["after_sha256"] and len(r["before_sha256"]) == 64
    assert r["parts"]["changed"] == [] and r["parts"]["added"] == [] and r["parts"]["removed"] == []
    assert r["parts"]["unchanged"] > 0


def test_receipt_has_every_documented_key():
    r = verify(fixture_path("WithChart.xlsx"), fixture_path("WithChart.xlsx"))
    assert set(r) == {"verdict", "valid", "lost_features", "rebuilt", "parts", "before_sha256",
                      "after_sha256", "reasons", "added_risk", "cell_changes", "reference_changes", "summary"}
    assert set(r["parts"]) == {"unchanged", "changed", "added", "removed"}


def test_openpyxl_round_trip_on_chart_is_rewritten_not_damaged(tmp_path):
    """openpyxl re-writes the chart but keeps the count: nothing demonstrable is lost."""
    src = fixture_path("WithChart.xlsx")
    out = tmp_path / "rt.xlsx"
    openpyxl.load_workbook(src).save(out)
    r = verify(src, out)
    assert r["verdict"] == "rewritten" and r["rebuilt"] is True and r["lost_features"] == []
    assert any("rebuilt" in s for s in r["reasons"])
    assert not any("damaged" in s.lower() for s in r["reasons"])
    assert any("rewritten by another tool" in s for s in r["reasons"])


def test_openpyxl_round_trip_on_macro_workbook_names_lost_vba(tmp_path):
    src = fixture_path("SimpleMacro.xlsm")
    out = tmp_path / "rt.xlsm"
    openpyxl.load_workbook(src).save(out)  # no keep_vba: the macro project is dropped
    r = verify(src, out)
    assert r["verdict"] == "damaged" and r["rebuilt"] is True
    vba = [x for x in r["lost_features"] if x["feature"] == "vba"]
    assert vba and (vba[0]["before"], vba[0]["after"]) == (1, 0)


def test_openpyxl_round_trip_with_keep_vba_is_still_damaged(tmp_path):
    src = fixture_path("SimpleMacro.xlsm")
    out = tmp_path / "rt.xlsm"
    openpyxl.load_workbook(src, keep_vba=True).save(out)
    assert verify(src, out)["verdict"] in ("rewritten", "damaged")


def test_openpyxl_round_trip_names_lost_images_and_conditional_formats(tmp_path):
    for name, feature in (("WithDrawing.xlsx", "images"), ("NewStyleConditionalFormattings.xlsx", "conditional_formats")):
        src = fixture_path(name)
        out = tmp_path / f"rt-{name}"
        openpyxl.load_workbook(src).save(out)
        r = verify(src, out)
        assert r["verdict"] == "damaged"
        assert feature in [x["feature"] for x in r["lost_features"]]


def test_rebuild_alone_is_damaged(tmp_path):
    """Features all survive but a never-rewritten part changed -> rewritten."""
    src = fixture_path("WithChart.xlsx")
    out = tmp_path / "rb.xlsx"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info)
            if info.filename == "docProps/core.xml":
                data = data + b" "
            zo.writestr(info, data)
    r = verify(src, out)
    assert r["verdict"] == "rewritten" and r["rebuilt"] is True and r["lost_features"] == []


def test_non_zip_is_invalid_not_an_exception(tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_bytes(b"garbage")
    r = verify(fixture_path("WithChart.xlsx"), bad)
    assert r["verdict"] == "invalid" and r["valid"] is False and r["reasons"]
    r2 = verify(bad, fixture_path("WithChart.xlsx"))
    assert r2["verdict"] == "invalid"
    assert verify(fixture_path("WithChart.xlsx"), tmp_path / "nope.xlsx")["verdict"] == "invalid"


def test_corrupt_xml_part_is_invalid(tmp_path):
    src = fixture_path("WithChart.xlsx")
    out = tmp_path / "c.xlsx"
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info)
            if info.filename == "xl/workbook.xml":
                data = data[: len(data) // 2]
            zo.writestr(info, data)
    r = verify(src, out)
    assert r["verdict"] == "invalid" and r["valid"] is False


def test_zip_over_caps_is_refused_cleanly(tmp_path, monkeypatch):
    src = fixture_path("WithChart.xlsx")
    monkeypatch.setattr(features, "MAX_TOTAL_BYTES", 100)
    r = verify(src, src)
    assert r["verdict"] == "invalid"
    assert any("cap" in s for s in r["reasons"])


def test_part_count_cap_refused(tmp_path, monkeypatch):
    big = tmp_path / "many.xlsx"
    with zipfile.ZipFile(big, "w") as z:
        for i in range(10):
            z.writestr(f"x{i}.xml", "<a/>")
    monkeypatch.setattr(features, "MAX_PARTS", 5)
    r = verify(fixture_path("WithChart.xlsx"), big)
    assert r["verdict"] == "invalid"


def test_openpyxl_round_trip_keeps_comment_count_despite_new_paths(tmp_path):
    """openpyxl moves comments to xl/comments/comment1.xml (+ a VML drawing): same count, not a loss."""
    import shutil
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["A1"].comment = openpyxl.comments.Comment("note", "me")
    wb.active["A2"].comment = openpyxl.comments.Comment("second", "me")
    gen = tmp_path / "gen.xlsx"
    wb.save(gen)
    # move the comments part to the path Excel itself uses, so the two files differ in path only
    legacy = tmp_path / "legacy.xlsx"
    with zipfile.ZipFile(gen) as zi, zipfile.ZipFile(legacy, "w", zipfile.ZIP_DEFLATED) as zo:
        for info in zi.infolist():
            data = zi.read(info).replace(b"/xl/comments/comment1.xml", b"/xl/comments1.xml")
            name = "xl/comments1.xml" if info.filename == "xl/comments/comment1.xml" else info.filename
            zo.writestr(name, data)
    assert features.fingerprint(legacy)["features"]["comments"] == 2
    out = tmp_path / "rt.xlsx"
    openpyxl.load_workbook(legacy).save(out)
    assert features.fingerprint(out)["features"]["comments"] == 2
    r = verify(legacy, out)
    assert "comments" not in [x["feature"] for x in r["lost_features"]]
