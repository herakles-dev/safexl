"""safexl_mcp.engine -- untouched OOXML parts are byte-identical to the
source after an edit, across all 7 real fixtures, and
xl/vbaProject.bin in particular survives with an unchanged hash (the VBA
project is never touched, never re-serialized -- see engine.py module
docstring's "What always changes" section).

This is the pytest-native equivalent of scripts/smoke_safexl.sh's check #3,
but with exact sha256 assertions per part instead of a shell script's diff.
"""
from __future__ import annotations

import hashlib
import zipfile

import pytest

from safexl_mcp import edit

from _util import FIXTURE_NAMES, fixture_path, first_sheet_name, zip_part_hashes

# workbook.xml.rels changes only when a calcChain relationship is dropped with its part.
EXPECTED_MAYBE_CHANGED = {"xl/workbook.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_untouched_parts_byte_identical_across_all_fixtures(tmp_path, name):
    src = fixture_path(name)
    dst = tmp_path / f"out-{name}"
    sheet = first_sheet_name(src)

    edit(
        src,
        dst,
        [
            {"sheet": sheet, "cell": "ZZ9000", "value": "SAFEXL_TEST"},
            {"sheet": sheet, "cell": "ZZ9001", "formula": "1+1"},
        ],
    )

    before = zip_part_hashes(src)
    after = zip_part_hashes(dst)

    missing = set(before) - set(after) - {"xl/calcChain.xml"}
    assert not missing, f"unexpected missing parts: {sorted(missing)}"

    extra = set(after) - set(before)
    assert not extra, f"unexpected new parts: {sorted(extra)}"

    changed = [n for n in sorted(set(before) & set(after)) if before[n] != after[n]]
    non_exempt_changed = [n for n in changed if n not in EXPECTED_MAYBE_CHANGED]
    assert len(non_exempt_changed) == 1, (
        f"expected exactly 1 changed non-exempt part (the edited worksheet), "
        f"got {len(non_exempt_changed)}: {non_exempt_changed}"
    )


@pytest.mark.parametrize("name", FIXTURE_NAMES)
def test_only_the_edited_worksheet_part_itself_differs_no_others(tmp_path, name):
    """Stronger form of the check above: every OTHER worksheet part (for
    fixtures with multiple sheets) must be untouched too -- only the one
    sheet actually named in the change set may differ."""
    src = fixture_path(name)
    dst = tmp_path / f"out-{name}"
    sheet = first_sheet_name(src)

    edit(src, dst, [{"sheet": sheet, "cell": "ZZ9000", "value": 1}])

    before = zip_part_hashes(src)
    after = zip_part_hashes(dst)
    other_worksheet_parts = [
        n
        for n in before
        if n.startswith("xl/worksheets/sheet")
        and n.endswith(".xml")
        and n in after
    ]
    # There must be at least the one edited worksheet; any additional
    # worksheet parts (other sheets in a multi-sheet fixture) must be
    # byte-identical.
    unchanged_others = [n for n in other_worksheet_parts if before[n] == after[n]]
    changed_worksheets = [n for n in other_worksheet_parts if before[n] != after[n]]
    assert len(changed_worksheets) == 1, changed_worksheets
    assert len(unchanged_others) == len(other_worksheet_parts) - 1


# ---------------------------------------------------------------------------
# xl/vbaProject.bin -- never touched, never re-serialized
# ---------------------------------------------------------------------------


def test_vbaproject_bin_hash_unchanged_after_editing_a_different_sheet(tmp_path):
    src = fixture_path("SimpleMacro.xlsm")
    with zipfile.ZipFile(src) as z:
        assert "xl/vbaProject.bin" in z.namelist()
        before_bytes = z.read("xl/vbaProject.bin")
    before_hash = hashlib.sha256(before_bytes).hexdigest()

    dst = tmp_path / "out.xlsm"
    sheet = first_sheet_name(src)
    edit(
        src,
        dst,
        [
            {"sheet": sheet, "cell": "ZZ9000", "value": "hello & <world>"},
            {"sheet": sheet, "cell": "ZZ9001", "formula": "1+1"},
        ],
    )

    with zipfile.ZipFile(dst) as z:
        after_bytes = z.read("xl/vbaProject.bin")
    after_hash = hashlib.sha256(after_bytes).hexdigest()

    assert after_bytes == before_bytes
    assert after_hash == before_hash


def test_vbaproject_bin_hash_unchanged_across_multiple_edits(tmp_path):
    """Two sequential edits (second edit reads the first edit's output as its
    own source) must still leave vbaProject.bin byte-identical to the
    original -- the VBA project is copied byte-for-byte on every pass, not
    just the first."""
    src = fixture_path("SimpleMacro.xlsm")
    with zipfile.ZipFile(src) as z:
        original_vba = z.read("xl/vbaProject.bin")

    sheet = first_sheet_name(src)
    mid = tmp_path / "mid.xlsm"
    edit(src, mid, [{"sheet": sheet, "cell": "A100", "value": 1}])

    final = tmp_path / "final.xlsm"
    edit(mid, final, [{"sheet": sheet, "cell": "A101", "value": 2}])

    with zipfile.ZipFile(final) as z:
        final_vba = z.read("xl/vbaProject.bin")
    assert final_vba == original_vba
