"""Tests for fetch_corpus.py. No network: the opener is injected."""
import hashlib
import importlib.util
import io
import json
import sys
from pathlib import Path

BENCH_DIR = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("xlsx_bench_fetch_corpus", BENCH_DIR / "fetch_corpus.py")
fc = importlib.util.module_from_spec(_spec)
sys.modules["xlsx_bench_fetch_corpus"] = fc
_spec.loader.exec_module(fc)

GOOD = b"workbook bytes"
GOOD_HASH = hashlib.sha256(GOOD).hexdigest()


class _Resp(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def opener_for(payload, calls=None):
    def opener(url, timeout=None):
        if calls is not None:
            calls.append(url)
        if isinstance(payload, Exception):
            raise payload
        return _Resp(payload)
    return opener


def _setup(tmp_path, sha=GOOD_HASH):
    d = tmp_path / "files"
    manifest = tmp_path / "m.json"
    manifest.write_text(json.dumps({"corpus": [{"file": "a.xlsx", "source_url": "https://example.test/a.xlsx", "sha256": sha}]}))
    return [(manifest, "corpus", d)], d / "a.xlsx"


def test_downloads_missing_file_and_verifies(tmp_path):
    m, dest = _setup(tmp_path)
    out = io.StringIO()
    assert fc.run(manifests=m, opener=opener_for(GOOD), out=out) == 0
    assert dest.read_bytes() == GOOD
    assert "fetched" in out.getvalue() and "a.xlsx" in out.getvalue()


def test_existing_correct_file_is_not_downloaded(tmp_path):
    m, dest = _setup(tmp_path)
    dest.parent.mkdir()
    dest.write_bytes(GOOD)
    calls = []
    assert fc.run(manifests=m, opener=opener_for(b"other", calls), out=io.StringIO()) == 0
    assert calls == []
    assert dest.read_bytes() == GOOD


def test_hash_mismatch_writes_nothing_and_fails(tmp_path):
    m, dest = _setup(tmp_path)
    out = io.StringIO()
    assert fc.run(manifests=m, opener=opener_for(b"tampered"), out=out) == 1
    assert not dest.exists()
    assert not list(dest.parent.glob("*")) if dest.parent.exists() else True
    assert "failed" in out.getvalue()


def test_wrong_existing_file_is_replaced_only_by_a_verified_download(tmp_path):
    m, dest = _setup(tmp_path)
    dest.parent.mkdir()
    dest.write_bytes(b"old wrong")
    assert fc.run(manifests=m, opener=opener_for(b"still wrong"), out=io.StringIO()) == 1
    assert dest.read_bytes() == b"old wrong"
    assert fc.run(manifests=m, opener=opener_for(GOOD), out=io.StringIO()) == 0
    assert dest.read_bytes() == GOOD


def test_download_error_is_reported_not_raised(tmp_path):
    m, dest = _setup(tmp_path)
    assert fc.run(manifests=m, opener=opener_for(OSError("no network")), out=io.StringIO()) == 1
    assert not dest.exists()


def test_check_mode_never_downloads_and_exit_codes(tmp_path):
    m, dest = _setup(tmp_path)
    calls = []
    assert fc.run(check_only=True, manifests=m, opener=opener_for(GOOD, calls), out=io.StringIO()) == 1
    assert calls == [] and not dest.exists()
    dest.parent.mkdir()
    dest.write_bytes(b"wrong")
    out = io.StringIO()
    assert fc.run(check_only=True, manifests=m, out=out) == 1
    assert "wrong" in out.getvalue()
    dest.write_bytes(GOOD)
    assert fc.run(check_only=True, manifests=m, out=io.StringIO()) == 0


def test_real_manifests_are_well_formed():
    entries = fc.load_entries()
    assert len(entries) >= 14
    names = [dest for _, dest in entries]
    assert len(set(names)) == len(names)
    for entry, _ in entries:
        assert len(entry["sha256"]) == 64
        assert entry["source_url"].startswith("https://raw.githubusercontent.com/")
        assert "/trunk/" not in entry["source_url"] and "/develop/" not in entry["source_url"]
