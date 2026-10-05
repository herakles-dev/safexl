#!/usr/bin/env python3
"""Download the third-party workbooks the tests and the benchmark use.

The public repository holds no .xlsx/.xlsm files. Two manifests list every
workbook with its pinned upstream URL and sha256:

    research/xlsx-bench/corpus.json                 -> research/xlsx-bench/corpus/
    research/safexl-moat/fixtures/fixtures.json     -> research/safexl-moat/fixtures/

For each entry: a file that is already present with the right hash is left
alone; otherwise the upstream file is downloaded, its hash is checked, and
only then is it written. A file whose hash does not match is never left on
disk.

    python research/xlsx-bench/fetch_corpus.py           # download what is missing
    python research/xlsx-bench/fetch_corpus.py --check   # verify only, no network

Exit code 0 when every file is present and correct, 1 otherwise. Standard
library only.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
RESEARCH = HERE.parent

# (manifest path, key that holds the entry list, directory the files live in)
MANIFESTS = [
    (HERE / "corpus.json", "corpus", HERE / "corpus"),
    (RESEARCH / "safexl-moat" / "fixtures" / "fixtures.json", "fixtures", RESEARCH / "safexl-moat" / "fixtures"),
]

TIMEOUT = 60


def sha256_of(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load_entries(manifests=None):
    """[(entry dict, destination Path)] for every manifest that exists."""
    out = []
    for manifest, key, directory in manifests if manifests is not None else MANIFESTS:
        manifest = Path(manifest)
        if not manifest.is_file():
            continue
        data = json.loads(manifest.read_text())
        for entry in data.get(key, []):
            out.append((entry, Path(directory) / entry["file"]))
    return out


def _download(url: str, opener) -> bytes:
    with opener(url, timeout=TIMEOUT) as resp:
        return resp.read()


def check_one(entry: dict, dest: Path) -> str:
    """'ok', 'missing' or 'wrong' for the file on disk."""
    if not dest.is_file():
        return "missing"
    return "ok" if sha256_of(dest.read_bytes()) == entry["sha256"] else "wrong"


def fetch_one(entry: dict, dest: Path, opener=urllib.request.urlopen) -> str:
    """Make ``dest`` hold the entry's file. Returns 'ok' (already right),
    'fetched', or 'failed: <reason>'. Never leaves a file with a wrong hash
    that this call wrote; an existing wrong file is replaced only by a
    verified download."""
    if check_one(entry, dest) == "ok":
        return "ok"
    try:
        data = _download(entry["source_url"], opener)
    except Exception as exc:  # network, HTTP error, bad URL
        return f"failed: download error: {exc}"
    got = sha256_of(data)
    if got != entry["sha256"]:
        return f"failed: sha256 {got} does not match the manifest {entry['sha256']}"
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.name + ".part")
    tmp.write_bytes(data)
    os.replace(tmp, dest)
    return "fetched"


def run(check_only: bool = False, manifests=None, opener=urllib.request.urlopen, out=None) -> int:
    out = out or sys.stdout
    bad = 0
    for entry, dest in load_entries(manifests):
        if check_only:
            status = check_one(entry, dest)
            if status != "ok":
                bad += 1
        else:
            status = fetch_one(entry, dest, opener)
            if status.startswith("failed"):
                bad += 1
        print(f"{status:8} {dest.name}", file=out)
    return 1 if bad else 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--check", action="store_true", help="verify files on disk only; no download; exit 1 if any is missing or wrong")
    args = ap.parse_args(argv)
    return run(check_only=args.check)


if __name__ == "__main__":
    sys.exit(main())
