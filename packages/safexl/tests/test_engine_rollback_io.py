"""safexl_mcp.engine -- ROLLBACK PROMISE, I/O-failure paths.

test_engine_rollback.py already covers the up-front validation failures and
the validator-CRITICAL failure path. This file covers the two I/O failure
modes that were previously untested, both of which fall through to the same
``except Exception as exc: raise SafeXLError(...)`` branch in ``edit()``
(safexl_mcp/engine.py, end of the function):

1. A raw OSError while writing the private temp file (e.g. disk full,
   permission denied) -- simulated by making ``zipfile.ZipFile.writestr``
   raise, since that is the call ``edit()`` uses for every zip member it
   writes into the temp output.
2. ``shutil.move`` raising while placing the validated temp file at ``dst``
   -- simulated by making ``shutil.move`` raise directly.

Both must uphold engine.py's documented ROLLBACK PROMISE: ``dst`` is left
completely untouched (never created if it didn't exist, never modified if
it did), and the private temp file is always cleaned up (the ``finally:
tmp_path.unlink(missing_ok=True)`` in ``edit()``), success or failure.
"""
from __future__ import annotations

import shutil
import threading
import zipfile

import pytest

from safexl_mcp import SafeXLError, edit

from _util import build_minimal_workbook

_A_VALID_CHANGE = [{"sheet": "Sheet1", "cell": "B2", "value": 1}]


def _raise_oserror(*_args, **_kwargs):
    raise OSError("simulated I/O failure")


# ---------------------------------------------------------------------------
# shutil.move raising at placement
# ---------------------------------------------------------------------------


def test_move_failure_raises_safexlerror_wrapping_the_oserror(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    monkeypatch.setattr(shutil, "move", lambda *a, **k: (_ for _ in ()).throw(OSError("disk full at placement")))

    with pytest.raises(SafeXLError, match="failed to write edited workbook.*disk full at placement"):
        edit(src, dst, _A_VALID_CHANGE)


def test_move_failure_dst_not_created_when_it_did_not_exist(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"  # does not exist yet
    monkeypatch.setattr(shutil, "move", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert not dst.exists()
    # Only the source remains -- the private temp file (created next to dst)
    # was cleaned up by edit()'s finally block despite the failure.
    assert {p.name for p in tmp_path.iterdir()} == {"src.xlsx"}


def test_move_failure_preexisting_dst_left_byte_for_byte_untouched(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    sentinel = b"PRE-EXISTING-DST-CONTENT-DO-NOT-TOUCH"
    dst.write_bytes(sentinel)
    monkeypatch.setattr(shutil, "move", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert dst.read_bytes() == sentinel
    assert {p.name for p in tmp_path.iterdir()} == {"src.xlsx", "out.xlsx"}


def test_move_failure_leaves_no_temp_files_in_dst_directory(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst_dir = tmp_path / "outdir"
    dst_dir.mkdir()
    dst = dst_dir / "out.xlsx"
    monkeypatch.setattr(shutil, "move", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert list(dst_dir.iterdir()) == []


# ---------------------------------------------------------------------------
# Raw OSError while writing the temp file
# ---------------------------------------------------------------------------


def test_temp_write_oserror_raises_safexlerror_wrapping_the_oserror(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    monkeypatch.setattr(
        zipfile.ZipFile,
        "writestr",
        lambda *a, **k: (_ for _ in ()).throw(OSError("simulated disk full during temp write")),
    )

    with pytest.raises(SafeXLError, match="failed to write edited workbook.*simulated disk full during temp write"):
        edit(src, dst, _A_VALID_CHANGE)


def test_temp_write_oserror_dst_not_created_when_it_did_not_exist(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"  # does not exist yet
    monkeypatch.setattr(zipfile.ZipFile, "writestr", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert not dst.exists()
    assert {p.name for p in tmp_path.iterdir()} == {"src.xlsx"}


def test_temp_write_oserror_preexisting_dst_left_byte_for_byte_untouched(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    sentinel = b"PRE-EXISTING-DST-CONTENT-DO-NOT-TOUCH"
    dst.write_bytes(sentinel)
    monkeypatch.setattr(zipfile.ZipFile, "writestr", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert dst.read_bytes() == sentinel
    assert {p.name for p in tmp_path.iterdir()} == {"src.xlsx", "out.xlsx"}


def test_temp_write_oserror_leaves_no_temp_files_in_dst_directory(tmp_path, monkeypatch):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst_dir = tmp_path / "outdir"
    dst_dir.mkdir()
    dst = dst_dir / "out.xlsx"
    monkeypatch.setattr(zipfile.ZipFile, "writestr", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert list(dst_dir.iterdir()) == []


def test_temp_write_oserror_does_not_corrupt_the_source_file(tmp_path, monkeypatch):
    """src is opened read-only ("r") throughout edit() -- a write failure on
    the *output* side must never touch the source's bytes."""
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    original_bytes = src.read_bytes()
    dst = tmp_path / "out.xlsx"
    monkeypatch.setattr(zipfile.ZipFile, "writestr", _raise_oserror)

    with pytest.raises(SafeXLError):
        edit(src, dst, _A_VALID_CHANGE)

    assert src.read_bytes() == original_bytes


# ---------------------------------------------------------------------------
# Concurrency: two edit() calls racing to write the SAME dst
# ---------------------------------------------------------------------------
#
# edit() never writes dst directly -- it writes a private temp file (unique
# per call, from tempfile.mkstemp) in dst's own directory, validates it, and
# only then calls shutil.move(tmp_path, dst). On the same filesystem (always
# true here: the temp file is created with dir=dst.parent) shutil.move is a
# plain os.rename, which POSIX guarantees is atomic. So two concurrent
# edit() calls targeting the same dst never interleave or torn-write it --
# each rename either fully replaces dst or hasn't happened yet -- but there
# is no locking between them: whichever call's shutil.move executes last
# simply overwrites the other's result, same as any two processes racing to
# rename a file onto the same path. See the Concurrency section of README.md
# for the plain-English version of exactly this guarantee.


def test_concurrent_edits_to_the_same_dst_never_corrupt_it(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    barrier = threading.Barrier(2)
    errors: list = [None, None]

    def worker(index: int, value: int) -> None:
        try:
            barrier.wait(timeout=10)
            edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": value}])
        except Exception as exc:  # noqa: BLE001 - captured for the assertion below
            errors[index] = exc

    threads = [threading.Thread(target=worker, args=(0, 111)), threading.Thread(target=worker, args=(1, 222))]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=10)

    # Neither call fails -- there's no lock contention error, only a race
    # over which one's rename lands last.
    assert errors == [None, None]

    # dst is a single, structurally intact workbook containing exactly one
    # of the two racing values -- never both (torn write) and never neither
    # (dst always exists once both calls have returned, since shutil.move
    # only fails via the OSError paths covered above, not under a plain
    # race).
    with zipfile.ZipFile(dst) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    has_111 = ">111<" in sheet_xml
    has_222 = ">222<" in sheet_xml
    assert has_111 != has_222, "dst should contain exactly one racing writer's value, not both/neither"

    # No stray temp file from either call was left behind.
    assert {p.name for p in tmp_path.iterdir()} == {"src.xlsx", "out.xlsx"}
