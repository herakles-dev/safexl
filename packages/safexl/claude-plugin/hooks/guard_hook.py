#!/usr/bin/env python3
"""SafeXL write guard for Claude Code (SessionStart + PostToolUse hook). Standard library only.

  guard_hook.py session-start   snapshot the workbooks in the working directory
  guard_hook.py post-tool       after a tool call that can write files, repair any workbook another
                                writer damaged: `safexl salvage` re-applies the model's edits to the
                                snapshot original; if that fails, restore the snapshot and tell the model
                                to redo the edit with safexl. A new workbook that is a damaged copy of a
                                tracked one is repaired the same way (the original is not touched).
                                SAFEXL_GUARD_MODE=restore turns salvage and the copy handling off.

Hook I/O (Claude Code hooks reference): stdin is one JSON object with session_id, cwd, hook_event_name;
the model reads hookSpecificOutput.additionalContext, the user reads systemMessage. Exit 0 always: a
hook failure must never break the session or block a tool call. No network.

State lives outside the project: $CLAUDE_PLUGIN_DATA/sessions/<session_id>/ (temp dir fallback), one
snapshot dir per tracked workbook. Turn the guard off with SAFEXL_GUARD=off or an empty file named
.safexl-guard-off in the project directory.

Caps: workbooks up to MAX_BYTES each, at most MAX_FILES tracked, found in the project dir and up to
MAX_DEPTH levels below it (hidden dirs and dependency dirs skipped).
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

MAX_BYTES = 50 * 1024 * 1024
MAX_FILES = 20
MAX_DEPTH = 2
SUFFIXES = (".xlsx", ".xlsm")
SKIP_DIRS = {"node_modules", "venv", "env", "__pycache__", "site-packages"}
CLI_TIMEOUT = 40

ADVICE = ("Write workbooks only with safexl edit / insert-rows / rename-sheet / add-sheet; "
          "never save with openpyxl or pandas.")
NOT_INSTALLED = ("SafeXL plugin: the `safexl` command is not on PATH, so the workbook guard is idle. "
                 "Install it (pipx install \"git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl\") and restart the session.")


def _stat_key(p: Path) -> list[int]:
    st = p.stat()
    return [st.st_size, st.st_mtime_ns, st.st_ino, st.st_ctime_ns]


def _data_root() -> Path:
    base = os.environ.get("CLAUDE_PLUGIN_DATA")
    return Path(base) if base else Path(tempfile.gettempdir()) / "safexl-guard"


def _state_dir(session_id: str) -> Path:
    safe = "".join(c if c.isalnum() or c in "-_" else "-" for c in session_id)[:64] or "nosession"
    return _data_root() / "sessions" / safe


def _find_workbooks(root: Path) -> list[Path]:
    found: list[Path] = []
    stack = [(root, 0)]
    while stack and len(found) < 500:
        d, depth = stack.pop()
        try:
            entries = sorted(os.scandir(d), key=lambda e: e.name)
        except OSError:
            continue
        for e in entries:
            try:
                if e.is_file(follow_symlinks=False):
                    if e.name.lower().endswith(SUFFIXES) and not e.name.startswith("~$"):
                        found.append(Path(e.path))
                elif e.is_dir(follow_symlinks=False) and depth < MAX_DEPTH \
                        and not e.name.startswith(".") and e.name not in SKIP_DIRS:
                    stack.append((Path(e.path), depth + 1))
            except OSError:
                continue
    return sorted(found)


def _load(sd: Path) -> dict:
    try:
        st = json.loads((sd / "state.json").read_text())
        if isinstance(st, dict) and isinstance(st.get("tracked"), dict):
            return st
    except (OSError, ValueError):
        pass
    return {"tracked": {}, "warned": False}


def _save(sd: Path, st: dict) -> None:
    sd.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=sd, prefix=".state-", suffix=".tmp")
    with os.fdopen(fd, "w") as f:
        json.dump(st, f)
    os.replace(tmp, sd / "state.json")


def _snap_dir(sd: Path, path: Path) -> Path:
    return sd / "snaps" / hashlib.sha256(str(path).encode()).hexdigest()[:16]


def _run(cmd: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, timeout=CLI_TIMEOUT, stdin=subprocess.DEVNULL)


def _safexl() -> str | None:
    return os.environ.get("SAFEXL_BIN") or shutil.which("safexl")


def _mode() -> str:
    return "restore" if os.environ.get("SAFEXL_GUARD_MODE", "").lower() == "restore" else "salvage"


def _sheet_names(p: Path) -> tuple[str, ...] | None:
    try:
        with zipfile.ZipFile(p) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None
    return tuple(sorted(re.findall(r"<(?:\w+:)?sheet\b[^>]*?\bname=\"([^\"]*)\"", xml))) or None


def _snapshot_file(sd: Path, p: Path) -> Path | None:
    d = _snap_dir(sd, p)
    try:
        ent = json.loads((d / "manifest.json").read_text())[str(p)]
        f = d / ent["snapshot"]
        return f if f.is_file() else None
    except (OSError, ValueError, KeyError, TypeError):
        return None


def _reason(r: subprocess.CompletedProcess) -> str:
    """One line saying why salvage did not produce a file."""
    try:
        j = json.loads(r.stdout)
        bits = [str(j.get("status") or "")]
        for k in ("reason", "message"):
            if j.get(k):
                bits.append(str(j[k]))
        skipped = j.get("skipped")
        if skipped and not j.get("reason"):
            bits.append(f"{len(skipped)} edit(s) could not be carried over")
        line = ": ".join(b for b in bits if b)
        if line:
            return line
    except (ValueError, TypeError, AttributeError):
        pass
    for text in (r.stderr, r.stdout):
        for line in (text or "").splitlines():
            if line.strip():
                return line.strip()[:200]
    return f"salvage exited {r.returncode}"


def _replace_atomic(src: Path, dst: Path) -> None:
    fd, tmp = tempfile.mkstemp(dir=dst.parent, prefix=".guard-", suffix=".tmp")
    os.close(fd)
    try:
        shutil.copyfile(src, tmp)
        os.replace(tmp, dst)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _try_salvage(sd: Path, exe: str, original: Path, damaged: Path) -> tuple[Path | None, dict, str]:
    """Re-apply the model's edits in `damaged` onto `original`. Returns (verified output or None, json, reason)."""
    out = sd / "tmp" / f"salvage-{os.getpid()}{damaged.suffix}"
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        r = _run([exe, "salvage", str(original), str(damaged), "-o", str(out), "--force", "--json"])
    except (subprocess.TimeoutExpired, OSError):
        return None, {}, "salvage timed out"
    try:
        info = json.loads(r.stdout)
        info = info if isinstance(info, dict) else {}
    except ValueError:
        info = {}
    if r.returncode == 0 and out.is_file():
        try:
            if _run([exe, "verify", str(original), str(out)]).returncode == 0:
                return out, info, ""
        except (subprocess.TimeoutExpired, OSError):
            pass
        return None, info, "the salvaged file did not verify intact"
    return None, info, _reason(r)


def _edit_names(info: dict) -> tuple[int, str]:
    applied = info.get("applied") or []
    names = []
    for item in applied[:4]:
        if isinstance(item, dict):
            ref = next((item[k] for k in ("ref", "cell", "address", "target") if item.get(k)), None)
            item = (f"{item['sheet']}!{ref}" if item.get("sheet") else str(ref)) if ref else \
                json.dumps(item, sort_keys=True)[:40]
        names.append(str(item))
    return len(applied), ", ".join(names) + (" ..." if len(applied) > 4 else "")


def _salvaged_message(name: str, lost: str, info: dict) -> str:
    n, names = _edit_names(info)
    kept = f"its {n} edit(s) kept" + (f" ({names})" if names else "") if n else "its edits kept"
    if lost:
        head = f"{name} was saved by another writer{lost}."
        how = "It was rebuilt from the original with"
    else:
        head = f"{name} was re-saved by another library."
        how = "It was rebuilt from the original, so untouched parts are byte-identical, with"
    return (f"{head} {how} {kept}. Make any further changes only with safexl edit / insert-rows / "
            f"insert-cells / rename-sheet / add-sheet, and check with `safexl verify`.")


def _check_changed(sd: Path, st: dict, exe: str) -> list[str]:
    """Run the guard on tracked workbooks whose stat changed. Returns messages for the model."""
    messages: list[str] = []
    for key in sorted(st["tracked"]):
        p = Path(key)
        if not p.exists():
            # Moved or deleted on purpose (e.g. `mv step1.xlsx out.xlsx`): stop tracking, never resurrect.
            del st["tracked"][key]
            shutil.rmtree(_snap_dir(sd, p), ignore_errors=True)
            continue
        try:
            sk = _stat_key(p)
        except OSError:
            continue
        if sk == st["tracked"][key]:
            continue
        d = str(_snap_dir(sd, p))
        r = _run([exe, "guard", "check", "--dir", d, "--json"])  # read-only
        events = []
        if r.returncode == 10:
            try:
                events = [ev for ev in json.loads(r.stdout) if isinstance(ev, dict)]
            except (ValueError, TypeError):
                events = [{"message": f"{p.name} was damaged by another writer. {ADVICE}", "lost_features": []}]
        if not events:
            _run([exe, "guard", "check", "--dir", d, "--restore"])  # accepts an intact SafeXL edit as the new snapshot
        else:
            ev = events[0]
            snap = _snapshot_file(sd, p)
            salvaged, info, why = (None, {}, "")
            if _mode() == "salvage" and snap is not None:
                salvaged, info, why = _try_salvage(sd, exe, snap, p)
            if salvaged is not None:
                _replace_atomic(salvaged, p)
                _run([exe, "guard", "snapshot", "--dir", d, key])
                lost = ", and lost " + ", ".join(ev["lost_features"]) if ev.get("lost_features") else ""
                messages.append(_salvaged_message(p.name, lost, info))
            else:
                rr = _run([exe, "guard", "check", "--dir", d, "--restore", "--json"])
                try:
                    msgs = [str(e["message"]) for e in json.loads(rr.stdout) if e.get("message")]
                except (ValueError, TypeError, KeyError):
                    msgs = []
                msg = " ".join(msgs) or str(ev.get("message") or f"{p.name} was restored. {ADVICE}")
                if why:
                    msg += f" (Rebuild from your edits was not possible: {why}.)"
                msg += " Redo the edit now with safexl."
                messages.append(msg)
        try:
            st["tracked"][key] = _stat_key(p)
        except OSError:
            pass
    return messages


def _derivative_of(sd: Path, st: dict, exe: str, p: Path) -> Path | None:
    """The tracked workbook's snapshot that `p` is a damaged copy of, else None."""
    names = _sheet_names(p)
    if not names:
        return None
    for key in sorted(st["tracked"]):
        tp = Path(key)
        if tp == p or _sheet_names(tp) != names:
            continue
        snap = _snapshot_file(sd, tp)
        if snap is not None and _run([exe, "verify", str(snap), str(p)]).returncode in (11, 12):
            return snap
    return None


def _track_new(sd: Path, st: dict, exe: str, workbooks: list[Path], derive: bool = False) -> tuple[list[Path], list[str]]:
    """Snapshot workbooks not yet tracked (within the caps). With derive, a new copy of a tracked workbook
    that another library wrote is repaired first. Returns (newly tracked, messages for the model)."""
    added: list[Path] = []
    messages: list[str] = []
    seen = st.setdefault("derivs", {})
    for p in workbooks:
        key = str(p)
        if key in st["tracked"] or len(st["tracked"]) >= MAX_FILES:
            continue
        try:
            if p.stat().st_size > MAX_BYTES:
                continue
            sk = _stat_key(p)
        except OSError:
            continue
        if derive:
            if seen.get(key) == sk:
                continue  # already told the model about this exact file
            snap = _derivative_of(sd, st, exe, p)
            if snap is not None:
                salvaged, info, why = _try_salvage(sd, exe, snap, p)
                if salvaged is not None:
                    _replace_atomic(salvaged, p)
                    sk = _stat_key(p)
                    messages.append(_salvaged_message(p.name, "", info) +
                                    " The original workbook was not touched.")
                else:
                    seen[key] = sk
                    messages.append(
                        f"{p.name} is a copy of an existing workbook saved by another library, so it lost "
                        f"features, and it could not be rebuilt ({why}). Make it safely instead: "
                        f"`cp <original> {p.name}`, then apply the changes with safexl edit / insert-rows / "
                        f"insert-cells / rename-sheet / add-sheet.")
                    continue
        r = _run([exe, "guard", "snapshot", "--dir", str(_snap_dir(sd, p)), key])
        if r.returncode == 0:
            st["tracked"][key] = _stat_key(p)
            added.append(p)
    return added, messages


def main(argv: list[str]) -> dict | None:
    mode = argv[1] if len(argv) > 1 else ""
    try:
        payload = json.loads(sys.stdin.read() or "{}")
    except ValueError:
        payload = {}
    if not isinstance(payload, dict):
        payload = {}
    if os.environ.get("SAFEXL_GUARD", "").lower() in ("off", "0", "false"):
        return None
    cwd = Path(payload.get("cwd") or os.getcwd())
    if (cwd / ".safexl-guard-off").exists():
        return None
    workbooks = _find_workbooks(cwd)
    sd = _state_dir(str(payload.get("session_id") or ""))
    st = _load(sd)
    if not workbooks and not st["tracked"]:
        return None  # nothing to guard: silent, no state written

    exe = _safexl()
    if not exe:
        if st.get("warned"):
            return None
        st["warned"] = True
        _save(sd, st)
        return {"systemMessage": NOT_INSTALLED}

    if mode == "session-start":
        added, _ = _track_new(sd, st, exe, workbooks)
        _save(sd, st)
        if not added:
            return None
        names = ", ".join(p.name for p in added[:5]) + (" ..." if len(added) > 5 else "")
        return {"hookSpecificOutput": {"hookEventName": "SessionStart", "additionalContext":
                f"SafeXL guard is on for: {names}. To change an existing workbook use the safexl skill: "
                f"{ADVICE} A workbook saved any other way is restored automatically."}}

    if mode == "post-tool":
        before = json.dumps(st, sort_keys=True)
        messages = _check_changed(sd, st, exe)
        # created mid-session: tracked from now on (a damaged copy of a tracked workbook is rebuilt first)
        _, more = _track_new(sd, st, exe, workbooks, derive=_mode() == "salvage")
        messages += more
        if json.dumps(st, sort_keys=True) != before:
            _save(sd, st)
        if not messages:
            return None
        return {"hookSpecificOutput": {"hookEventName": "PostToolUse", "additionalContext":
                "SAFEXL GUARD: " + " ".join(messages) +
                " Then check the result with `safexl verify`."}}
    return None


if __name__ == "__main__":
    try:
        out = main(sys.argv)
        if out:
            sys.stdout.write(json.dumps(out))
    except BaseException:  # noqa: BLE001 - a hook failure must never break the session
        pass
    sys.exit(0)
