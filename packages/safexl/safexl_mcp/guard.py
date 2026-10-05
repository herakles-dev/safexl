"""safexl_mcp.guard -- snapshot tracked workbooks, then undo damage done by another writer.

    snapshot(paths, snap_dir)        copy each workbook into snap_dir and record its sha256
    check(snap_dir)                  read-only: events for tracked workbooks that are damaged
    check_and_restore(snap_dir)      same, and restore each damaged one byte-identical from the snapshot

A tracked workbook is acted on when ``verify(snapshot, current)`` says so: a feature was lost, the
package was rebuilt by a non-SafeXL writer, or the file is no longer a valid package (or is gone).
An intact change (a SafeXL edit) is accepted: ``check_and_restore`` promotes it to the new snapshot
so later damage restores to the latest good copy, not to the original. Workbooks that did not exist
at snapshot time are never touched. No hook wiring lives here.

An event is {"file", "kind" ("damaged" | "rewritten" | "missing"), "lost_features", "rebuilt", "restored",
"message"}; ``message`` is one line a model can act on.
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any, Iterable

from .verify import verify

_MANIFEST = "manifest.json"
ADVICE = ("Write workbooks only with safexl edit / insert-rows / rename-sheet / add-sheet; "
          "never save with openpyxl or pandas.")


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _write_manifest(snap_dir: Path, entries: dict[str, Any]) -> None:
    fd, tmp = tempfile.mkstemp(dir=snap_dir, prefix=".manifest-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w") as fh:
            json.dump(entries, fh, indent=2)
        os.replace(tmp, snap_dir / _MANIFEST)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _read_manifest(snap_dir: Path) -> dict[str, Any]:
    try:
        return json.loads((snap_dir / _MANIFEST).read_text())
    except (OSError, ValueError):
        return {}


def _copy_atomic(src: Path, dst: Path) -> None:
    """Copy src over dst via a same-directory temp + os.replace: never writes through a symlink at
    dst and never leaves a half-written file."""
    fd, tmp = tempfile.mkstemp(dir=dst.parent, prefix=".guard-", suffix=".tmp")
    os.close(fd)
    try:
        shutil.copyfile(src, tmp)
        shutil.copymode(src, tmp)
        os.replace(tmp, dst)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def snapshot(paths: Iterable[str | Path], snap_dir: str | Path) -> dict[str, str]:
    """Copy each existing workbook into ``snap_dir`` and record it. Returns {path: sha256}.

    Paths that do not exist are skipped. Calling it again for a tracked path refreshes its snapshot."""
    sd = Path(snap_dir)
    sd.mkdir(parents=True, exist_ok=True)
    entries = _read_manifest(sd)
    out: dict[str, str] = {}
    for p in paths:
        src = Path(p).absolute()
        if not src.is_file():
            continue
        key = str(src)
        snap_name = entries.get(key, {}).get("snapshot") or f"{hashlib.sha256(key.encode()).hexdigest()[:12]}-{src.name}"
        shutil.copyfile(src, sd / snap_name)
        digest = _sha256(sd / snap_name)
        entries[key] = {"snapshot": snap_name, "sha256": digest}
        out[key] = digest
    _write_manifest(sd, entries)
    return out


def _event(name: str, kind: str, rec: dict[str, Any] | None, restored: bool) -> dict[str, Any]:
    if kind == "rewritten":
        tail = ("rewritten by another tool; restored to keep untouched parts byte-identical." if restored
                else "rewritten by another tool; NOT restored.")
        return {"file": name, "kind": kind, "lost_features": [], "rebuilt": True, "restored": restored,
                "message": f"{name} {tail} {ADVICE}"}
    rec = rec or {}
    lost = [x["feature"] for x in rec.get("lost_features", [])]
    rebuilt = bool(rec.get("rebuilt"))
    if kind == "missing":
        why = "was moved or deleted"
    else:
        bits = []
        if lost:
            bits.append("lost " + ", ".join(lost))
        if rebuilt:
            bits.append("was rebuilt by a non-SafeXL writer")
        if not bits:
            bits.append("is no longer a valid workbook")
        why = " and ".join(bits)
    tail = "It was restored to its last good version." if restored else "It was NOT restored."
    return {"file": name, "kind": kind, "lost_features": lost, "rebuilt": rebuilt, "restored": restored,
            "message": f"{name} {why}. {tail} {ADVICE}"}


def _scan(snap_dir: str | Path, restore: bool, strict: bool) -> list[dict[str, Any]]:
    sd = Path(snap_dir)
    entries = _read_manifest(sd)
    events: list[dict[str, Any]] = []
    changed = False
    for key, ent in sorted(entries.items()):
        target, snap = Path(key), sd / ent["snapshot"]
        if not snap.is_file():
            continue
        if not target.exists():
            if restore:
                _copy_atomic(snap, target)
            events.append(_event(target.name, "missing", None, restore))
            continue
        if _sha256(target) == ent["sha256"]:
            continue
        rec = verify(snap, target)
        if rec["verdict"] == "intact":
            if restore:
                shutil.copyfile(target, snap)
                ent["sha256"] = _sha256(snap)
                changed = True
            continue
        if rec["verdict"] == "rewritten":
            do = restore and strict
            if do:
                _copy_atomic(snap, target)
            events.append(_event(target.name, "rewritten", rec, do))
            continue
        if restore:
            _copy_atomic(snap, target)
        events.append(_event(target.name, "damaged", rec, restore))
    if changed:
        _write_manifest(sd, entries)
    return events


def check(snap_dir: str | Path) -> list[dict[str, Any]]:
    """Events for tracked workbooks that are damaged, rewritten, invalid or missing. Changes nothing."""
    return _scan(snap_dir, restore=False, strict=True)


def check_and_restore(snap_dir: str | Path, strict: bool = True) -> list[dict[str, Any]]:
    """Like ``check``, and restore byte-identical from the snapshot. strict (default) restores on
    rewritten, damaged, invalid or missing; strict=False restores only on damaged, invalid or missing
    (a merely rewritten file is reported with restored False and left in place)."""
    return _scan(snap_dir, restore=True, strict=strict)
