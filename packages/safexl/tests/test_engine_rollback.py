"""safexl_mcp.engine -- the ROLLBACK PROMISE (engine.py module docstring):
edit() never writes directly to dst. On any failure -- a bad change spec, an
unresolvable sheet/cell, a value+formula conflict, or a validator CRITICAL
finding on the temp output -- dst is left completely untouched: never
created if it didn't exist, never modified if it did.
"""
from __future__ import annotations

import pytest

from safexl_mcp import SafeXLError, edit

from _util import build_minimal_workbook


# ---------------------------------------------------------------------------
# dst not created (source had a valid file, dst didn't exist beforehand)
# ---------------------------------------------------------------------------


def test_bad_sheet_name_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="not found in workbook"):
        edit(src, dst, [{"sheet": "NoSuchSheet", "cell": "A1", "value": 1}])
    assert not dst.exists()


def test_bad_cell_ref_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="invalid cell reference"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "ABC", "value": 1}])  # letters, no digits
    assert not dst.exists()


def test_cell_ref_with_trailing_garbage_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="invalid cell reference"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "A1X", "value": 1}])  # digits then letter
    assert not dst.exists()


def test_value_and_formula_together_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="not both"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1, "formula": "A1"}])
    assert not dst.exists()


def test_neither_value_nor_formula_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="must specify"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2"}])
    assert not dst.exists()


def test_unsupported_value_type_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="unsupported value type"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": [1, 2, 3]}])
    assert not dst.exists()


def test_missing_sheet_key_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="missing/invalid 'sheet'"):
        edit(src, dst, [{"cell": "A1", "value": 1}])
    assert not dst.exists()


def test_missing_cell_key_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="missing/invalid 'cell'"):
        edit(src, dst, [{"sheet": "Sheet1", "value": 1}])
    assert not dst.exists()


def test_non_mapping_change_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="expected a dict/mapping"):
        edit(src, dst, ["not a dict"])
    assert not dst.exists()


def test_empty_changes_sequence_raises_and_dst_not_created(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="non-empty sequence"):
        edit(src, dst, [])
    assert not dst.exists()


def test_source_file_not_found_raises_and_dst_not_created(tmp_path):
    src = tmp_path / "does-not-exist.xlsx"
    dst = tmp_path / "out.xlsx"
    with pytest.raises(SafeXLError, match="source file not found"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 1}])
    assert not dst.exists()


# ---------------------------------------------------------------------------
# A pre-existing dst is left completely untouched on failure
# ---------------------------------------------------------------------------


def test_preexisting_dst_left_untouched_when_sheet_is_bad(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    sentinel = b"PRE-EXISTING-DST-CONTENT-DO-NOT-TOUCH"
    dst.write_bytes(sentinel)

    with pytest.raises(SafeXLError):
        edit(src, dst, [{"sheet": "NoSuchSheet", "cell": "A1", "value": 1}])

    assert dst.read_bytes() == sentinel


def test_preexisting_dst_left_untouched_when_value_formula_conflict(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    sentinel = b"PRE-EXISTING-DST-CONTENT-DO-NOT-TOUCH"
    dst.write_bytes(sentinel)

    with pytest.raises(SafeXLError):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "A1", "value": 1, "formula": "B1"}])

    assert dst.read_bytes() == sentinel


def test_preexisting_dst_left_untouched_when_validator_finds_critical_issue(tmp_path):
    """A source file with a pre-existing structural defect the engine does
    not repair (a Content_Types Override dangling for a part that was never
    there, unrelated to calcChain.xml) survives into the edited output and
    trips the validator's CRITICAL gate inside edit() -- after the temp file
    was written, but before it is moved into place at dst. This exercises
    the *validator-triggered* rollback path end-to-end via the public API,
    not just the up-front change-validation path."""
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(
        src,
        extra_content_types_xml=(
            '<Override PartName="/xl/bogus.xml" ContentType="application/xml"/>'
        ),
    )
    dst = tmp_path / "out.xlsx"
    sentinel = b"PRE-EXISTING-DST-CONTENT-DO-NOT-TOUCH"
    dst.write_bytes(sentinel)

    with pytest.raises(SafeXLError, match="critical structural finding"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    assert dst.read_bytes() == sentinel


def test_edit_rolls_back_and_creates_no_dst_when_validator_finds_critical_issue(tmp_path):
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(
        src,
        extra_content_types_xml=(
            '<Override PartName="/xl/bogus.xml" ContentType="application/xml"/>'
        ),
    )
    dst = tmp_path / "out.xlsx"  # does not exist yet

    with pytest.raises(SafeXLError, match="critical structural finding"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    assert not dst.exists()


def test_no_temp_files_left_behind_in_dst_directory_after_failed_edit(tmp_path):
    """edit()'s finally block always unlinks its private temp file, success
    or failure. Uses the validator-CRITICAL failure path specifically,
    because that is the one failure mode that occurs AFTER the temp file
    (created next to dst) has actually been written -- the up-front
    change-validation failures (bad sheet/cell/etc.) never get that far."""
    src = tmp_path / "src.xlsx"
    build_minimal_workbook(
        src,
        extra_content_types_xml=(
            '<Override PartName="/xl/bogus.xml" ContentType="application/xml"/>'
        ),
    )
    dst_dir = tmp_path / "outdir"
    dst_dir.mkdir()
    dst = dst_dir / "out.xlsx"

    with pytest.raises(SafeXLError, match="critical structural finding"):
        edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 1}])

    assert list(dst_dir.iterdir()) == []
