"""safexl_mcp.guard -- snapshot, detect damage by another writer, restore byte-identical."""
from __future__ import annotations

import shutil

import openpyxl

from safexl_mcp import edit
from safexl_mcp.guard import check, check_and_restore, snapshot

from _util import fixture_path, first_sheet_name, sha256_of


def _setup(tmp_path, name="WithChart.xlsx"):
    wbk = tmp_path / name
    shutil.copy(fixture_path(name), wbk)
    snap = tmp_path / "snap"
    snapshot([wbk], snap)
    return wbk, snap, sha256_of(wbk.read_bytes())


def _openpyxl_save(path):
    openpyxl.load_workbook(path).save(path)


def test_untouched_workbook_has_no_events(tmp_path):
    _, snap, _ = _setup(tmp_path)
    assert check(snap) == [] and check_and_restore(snap) == []


def test_damage_is_restored_to_original_sha256(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    _openpyxl_save(wbk)
    assert sha256_of(wbk.read_bytes()) != orig
    events = check_and_restore(snap)
    assert len(events) == 1
    ev = events[0]
    assert ev["file"] == wbk.name and ev["kind"] == "rewritten" and ev["restored"] is True
    assert ev["rebuilt"] is True
    assert "safexl" in ev["message"] and "\n" not in ev["message"]
    assert sha256_of(wbk.read_bytes()) == orig
    assert check_and_restore(snap) == []


def test_check_is_read_only(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    _openpyxl_save(wbk)
    damaged = sha256_of(wbk.read_bytes())
    events = check(snap)
    assert len(events) == 1 and events[0]["restored"] is False
    assert sha256_of(wbk.read_bytes()) == damaged


def test_vba_workbook_restored_and_loss_named(tmp_path):
    wbk, snap, orig = _setup(tmp_path, "SimpleMacro.xlsm")
    _openpyxl_save(wbk)
    ev = check_and_restore(snap)
    assert ev and "vba" in ev[0]["lost_features"] and sha256_of(wbk.read_bytes()) == orig


def test_safexl_edit_is_accepted_and_becomes_the_new_baseline(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    edit(wbk, wbk, [{"sheet": first_sheet_name(wbk), "cell": "Z90", "value": 1}])
    edited = sha256_of(wbk.read_bytes())
    assert edited != orig
    assert check_and_restore(snap) == []
    _openpyxl_save(wbk)
    assert check_and_restore(snap)
    assert sha256_of(wbk.read_bytes()) == edited


def test_new_workbook_is_left_alone(tmp_path):
    wbk, snap, _ = _setup(tmp_path)
    new = tmp_path / "new.xlsx"
    openpyxl.Workbook().save(new)
    before = sha256_of(new.read_bytes())
    assert check_and_restore(snap) == []
    assert sha256_of(new.read_bytes()) == before


def test_corrupt_overwrite_is_restored(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    wbk.write_bytes(b"corrupt")
    ev = check_and_restore(snap)
    assert ev and ev[0]["kind"] == "damaged"
    assert sha256_of(wbk.read_bytes()) == orig


def test_deleted_workbook_is_restored(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    wbk.unlink()
    ev = check_and_restore(snap)
    assert ev and ev[0]["kind"] == "missing" and ev[0]["restored"] is True
    assert sha256_of(wbk.read_bytes()) == orig


def test_missing_snapshot_dir_is_a_clean_no_op(tmp_path):
    assert check(tmp_path / "never-made") == []


def test_rewritten_file_restored_when_strict_with_honest_message(tmp_path):
    wbk, snap, orig = _setup(tmp_path)  # WithChart: openpyxl keeps every feature count -> rewritten
    _openpyxl_save(wbk)
    ev = check_and_restore(snap)
    assert ev[0]["kind"] == "rewritten" and ev[0]["restored"] is True and ev[0]["lost_features"] == []
    assert "rewritten by another tool; restored to keep untouched parts byte-identical" in ev[0]["message"]
    assert "lost" not in ev[0]["message"]
    assert sha256_of(wbk.read_bytes()) == orig


def test_non_strict_leaves_rewritten_but_restores_damaged(tmp_path):
    wbk, snap, orig = _setup(tmp_path)
    _openpyxl_save(wbk)
    rewritten = sha256_of(wbk.read_bytes())
    ev = check_and_restore(snap, strict=False)
    assert ev[0]["kind"] == "rewritten" and ev[0]["restored"] is False
    assert sha256_of(wbk.read_bytes()) == rewritten
    (tmp_path / "m").mkdir()
    wbk2, snap2, orig2 = _setup(tmp_path / "m", "SimpleMacro.xlsm")
    _openpyxl_save(wbk2)  # drops the macro project -> damaged
    ev2 = check_and_restore(snap2, strict=False)
    assert ev2[0]["kind"] == "damaged" and sha256_of(wbk2.read_bytes()) == orig2
