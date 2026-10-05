#!/usr/bin/env python3
"""Interactive-session recorder for the xlsx-bench two-arm benchmark.

Where runner.py drives Claude Code headlessly (``-p``, stream-json) to get a
scorable ``run.json``, this module drives the same two arms *interactively*,
inside a tmux session, under ``asciinema``, so a customer can:

  - WATCH a real session (``session.cast``, played back with ``asciinema
    play`` or embedded via asciinema-player -- no video, ~10-50KB of text),
  - READ exactly what Claude did (``session.jsonl``, the real Claude Code
    session transcript for that work folder -- rendered by ``transcript.py``),
  - VERIFY the files weren't swapped afterwards (``manifest.json``: sha256 of
    the original workbook, the output file(s), the cast, and the transcript).

It deliberately reuses runner.py's arm logic (``build_command``, ``build_env``,
``ARMS``, ``ALLOWED_TOOLS``, task loading, the ``<task>/<arm>/<rep>/work``
layout, ``is_done``/``write_run_json`` resumability) rather than forking it --
see runner.py's module docstring for why arm isolation is paranoid here.

Empirically verified against this box's `claude` 2.1.283 (see git history for
the throwaway probes; not checked in):

  - ``claude [options] PROMPT`` in interactive mode (no ``-p``) pre-fills
    *and auto-submits* PROMPT as the first user turn -- confirmed by watching
    a real reply come back with no keystroke sent.
  - This project's ``.mcp.json`` lists 7 unrelated servers (gmail, github,
    sentry, socket, hugging-face, ...). On a fresh interactive launch this
    pops a blocking "N new MCP servers found ... Enable selected" dialog --
    *before* the folder-trust dialog, and on every single launch (it's keyed
    to the project, not the work subdirectory). That dialog is not the
    folder-trust dialog and neither arm's tool use goes through MCP anyway
    (runner.py's README), so it is suppressed at the command-line with
    ``--strict-mcp-config`` -- the same "engineer the flag so the unwanted
    prompt never appears" approach ``build_env`` already uses for CLAUDE.md
    pollution, rather than teaching this tool to click through a dialog it
    was told never to auto-approve.
  - The folder-trust dialog ("Quick safety check: Is this ... a project you
    ... trust?") *is* real and *is not* suppressed by ``--permission-mode
    dontAsk`` -- confirmed by triggering it in a directory with no trusted
    git-repo ancestor. Default-selected option is "No, exit"; accepting is
    one Down arrow + Enter. It never fired against a directory nested inside
    this already-trusted repo/worktree in testing, but every recording polls
    for it anyway and accepts *only* that dialog -- any other interactive
    dialog aborts the run and is reported, never auto-approved.
  - The session transcript lives at
    ``~/.claude/projects/<cwd with / and . replaced by ->/​*.jsonl`` --
    confirmed against runner.py's own headless smoke/pilot run folders, which
    already have exactly this transcript on disk even though runner.py never
    asked for one.
  - Its ``assistant``/``user`` entries are structurally compatible with the
    stream-json events runner.py's ``summarize_events`` parses (same
    ``message.content`` block shapes for text/tool_use/tool_result), *but*:
    token ``usage`` is duplicated verbatim across every content-block entry
    that shares a ``requestId`` (one real API call can log 2-3 lines), so it
    must be deduped by ``requestId`` before summing -- see
    ``summarize_transcript``. There is no stream-json-style ``result`` event;
    a single ``cost-state`` entry near the end carries ``totalCostUSD``.

Stdlib only, except ``asciinema``/``tmux`` as external binaries (already on
this box). See README.md for the corresponding headless runner.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import secrets
import shlex
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import runner  # noqa: E402  (sibling module -- reused, not forked; see docstring)

RECORDER_VERSION = "1.0"

TMUX_COLS = 140
TMUX_ROWS = 42
DEFAULT_POLL_INTERVAL = 3.0
DEFAULT_QUIET_SECONDS = 20.0
EXIT_GRACE_SECONDS = 60
RELAUNCH_BACKOFF_SECONDS = 8.0
ASCIINEMA_BIN_DEFAULT = str(Path.home() / ".local" / "bin" / "asciinema")

# Any of these footer hints mean the pane is showing a modal dialog rather
# than the normal chat input box (whose footer reads "don't ask on ..." /
# "esc to interrupt" / "xhigh · /effort"). Built from the two dialogs
# actually observed on this box (folder-trust, MCP-servers-found); a dialog
# this doesn't recognize still matches here (both known dialogs show one of
# these hints) and is classified "other" -> aborts the run, never approved.
_DIALOG_HINT_RE = re.compile(
    r"Enter to confirm|Esc to cancel|Esc to reject|Space to select", re.IGNORECASE
)
_FOLDER_TRUST_MARKER = "trust this folder"

_ABS_PATH_RE = re.compile(r"/(?:[A-Za-z0-9_.\-]+/)*[A-Za-z0-9_.\-]+")


# ---------------------------------------------------------------------------
# Pure helpers: command construction
# ---------------------------------------------------------------------------


def interactive_command(
    claude_bin: str,
    prompt: str,
    arm: str,
    block_skill_method: str,
    plugin_dir: str = runner.SAFEXL_PLUGIN_DIR,
) -> list[str]:
    """Build the interactive (non-``-p``) claude argv for one recording.

    Starts from runner.build_command() -- the single source of truth for arm
    A/B flags -- then swaps the headless-only pieces for interactive ones:
    drops ``-p PROMPT``, ``--output-format stream-json``, ``--verbose``; adds
    ``--permission-mode dontAsk`` (deny outside the allow-list instead of
    hanging on a permission prompt) and ``--strict-mcp-config`` (see module
    docstring); re-appends PROMPT as a trailing positional argument, which
    the interactive CLI both pre-fills and auto-submits as the first turn.

    Asserts on build_command()'s exact current shape so a future change
    there fails loudly instead of silently building the wrong command.
    """
    headless = runner.build_command(claude_bin, prompt, arm, block_skill_method, plugin_dir)
    assert headless[0] == claude_bin, headless
    assert headless[1] == "-p" and headless[2] == prompt, headless
    assert headless[-3] == "--output-format" and headless[-2] == "stream-json", headless
    assert headless[-1] == "--verbose", headless
    middle = headless[3:-3]  # --allowedTools ALLOWED_TOOLS [arm B: --plugin-dir ... --disallowedTools/--settings ...]
    return [claude_bin, *middle, "--permission-mode", "dontAsk", "--strict-mcp-config", prompt]


def build_asciinema_argv(
    cast_path: Path, claude_argv: list[str], cols: int = TMUX_COLS, rows: int = TMUX_ROWS,
    asciinema_bin: str = "asciinema",
) -> list[str]:
    """argv for the asciinema wrapper. claude_argv is shlex-joined into a
    single shell-safe string for asciinema's ``-c`` (verified against
    apostrophes/backticks/`$`/unicode in real task prompts -- shlex.join's
    single-quoting keeps every one of those literal, whichever way asciinema
    ends up invoking it)."""
    return [
        asciinema_bin, "rec", "--overwrite", "-q",
        "--cols", str(cols), "--rows", str(rows),
        "-c", shlex.join(claude_argv), str(cast_path),
    ]


PANE_ENV_KEYS = ("PATH", "CLAUDE_CODE_DISABLE_CLAUDE_MDS", "CLAUDE_ENV_FILE")


def pane_env(env: dict[str, str]) -> dict[str, str]:
    """The arm-specific variables the tmux pane must receive via -e."""
    return {k: env[k] for k in PANE_ENV_KEYS if k in env}


def build_tmux_argv(
    session: str, cwd: Path, command_argv: list[str], cols: int = TMUX_COLS, rows: int = TMUX_ROWS,
    env: Optional[dict[str, str]] = None,
) -> list[str]:
    """argv for ``tmux new-session``. command_argv is passed as multiple
    tmux arguments (not one joined string) so tmux execs it directly without
    an extra ``sh -c`` layer -- see tmux(1) on new-session/shell-command.
    ``env`` entries become ``-e KEY=VALUE`` (the pane otherwise inherits the
    tmux server's environment, not the caller's)."""
    env_args = [a for k, v in (env or {}).items() for a in ("-e", f"{k}={v}")]
    return [
        "tmux", "new-session", "-d", "-s", session,
        "-x", str(cols), "-y", str(rows), "-c", str(cwd),
        *env_args,
        *command_argv,
    ]


def escape_project_dir(path: Path) -> str:
    """Claude Code's ``~/.claude/projects/<escaped>`` naming: every ``/`` and
    ``.`` in the absolute path replaced with ``-``. Verified against
    runner.py's own smoke/pilot run folders (e.g. the ``.claude`` in a
    worktree path becomes a double dash)."""
    return re.sub(r"[/.]", "-", str(path.resolve()))


# ---------------------------------------------------------------------------
# Pure helpers: dialog classification
# ---------------------------------------------------------------------------


def classify_dialog(pane_text: str) -> Optional[dict]:
    """Classify a tmux capture-pane snapshot as a dialog, or None if it's the
    ordinary chat UI. See module docstring for the two dialogs this was
    verified against."""
    if not pane_text or not _DIALOG_HINT_RE.search(pane_text):
        return None
    if _FOLDER_TRUST_MARKER in pane_text.lower():
        return {"type": "folder_trust"}
    return {"type": "other"}


def truncate_pane_text(text: str, n: int = 2000) -> str:
    text = (text or "").strip()
    if len(text) <= n:
        return text
    return text[:n].rstrip() + "…"


# ---------------------------------------------------------------------------
# Pure helpers: transcript parsing (Claude Code's ~/.claude/projects/*.jsonl)
# ---------------------------------------------------------------------------


def read_transcript_entries(path: Path) -> list[dict]:
    try:
        text = path.read_text(errors="replace")
    except OSError:
        return []
    return list(runner.iter_json_lines(text))


def transcript_is_finished(entries: list[dict]) -> bool:
    """True iff the LAST assistant-type entry has stop_reason "end_turn" and
    none of its content blocks is a pending tool_use. Uses the last entry
    specifically (not the first, not "any") because one logical turn can be
    split across several consecutive assistant-type JSONL lines that all
    share the same stop_reason (one per content block)."""
    last_assistant = None
    for e in entries:
        if e.get("type") == "assistant":
            last_assistant = e
    if last_assistant is None:
        return False
    message = last_assistant.get("message") or {}
    if message.get("stop_reason") != "end_turn":
        return False
    content = message.get("content") or []
    if any(isinstance(b, dict) and b.get("type") == "tool_use" for b in content):
        return False
    return True


def summarize_transcript(entries: list[dict]) -> dict:
    """Turn Claude Code transcript entries into run.json summary fields.

    Structurally parallel to runner.summarize_events (same tools_called /
    skills_invoked / safexl_used shape, for one shared mental model across
    both run.json flavors), but tailored to this format's differences: token
    usage lives on every assistant entry and is duplicated per requestId
    (deduped here before summing), and there's no stream-json "result" event
    (a "cost-state" entry carries totalCostUSD instead, once, near the end).
    """
    model = None
    total_cost_usd = None
    last_assistant_text = None

    usage_by_request: dict[str, dict] = {}
    request_order: list[str] = []

    tool_calls_by_id: dict[str, dict] = {}
    tools_called: list[str] = []
    tools_seen: set[str] = set()
    tool_call_count = 0

    skills_invoked: list[str] = []
    skills_seen: set[str] = set()
    skills_succeeded: set[str] = set()
    skills_denied: set[str] = set()
    skill_denial_evidence: dict[str, list[str]] = {}

    bash_commands: list[str] = []

    for entry in entries:
        etype = entry.get("type")

        if etype == "assistant":
            message = entry.get("message") or {}
            model = message.get("model", model)
            req = entry.get("requestId")
            usage = message.get("usage")
            if req and isinstance(usage, dict) and req not in usage_by_request:
                usage_by_request[req] = usage
                request_order.append(req)

            for block in message.get("content") or []:
                if not isinstance(block, dict):
                    continue
                btype = block.get("type")
                if btype == "text":
                    last_assistant_text = block.get("text", last_assistant_text)
                elif btype == "tool_use":
                    tool_call_count += 1
                    name = block.get("name")
                    tool_use_id = block.get("id")
                    tool_input = block.get("input") or {}
                    if not name:
                        continue
                    if name not in tools_seen:
                        tools_seen.add(name)
                        tools_called.append(name)
                    if tool_use_id:
                        tool_calls_by_id[tool_use_id] = {"name": name, "input": tool_input}
                    if name == "Skill":
                        skill_name = tool_input.get("skill") or tool_input.get("name")
                        if skill_name and skill_name not in skills_seen:
                            skills_seen.add(skill_name)
                            skills_invoked.append(skill_name)
                    if name == "Bash":
                        command = tool_input.get("command")
                        if isinstance(command, str):
                            bash_commands.append(command)

        elif etype == "user":
            message = entry.get("message") or {}
            content = message.get("content")
            if not isinstance(content, list):
                continue
            for block in content:
                if not isinstance(block, dict) or block.get("type") != "tool_result":
                    continue
                tool_use_id = block.get("tool_use_id")
                call = tool_calls_by_id.get(tool_use_id) if tool_use_id else None
                if not call or call["name"] != "Skill":
                    continue
                skill_name = call["input"].get("skill") or call["input"].get("name")
                if not skill_name:
                    continue
                text = runner._tool_result_text(block)
                errored = bool(block.get("is_error")) or runner._looks_like_denial(text)
                if errored:
                    skills_denied.add(skill_name)
                    skill_denial_evidence.setdefault(skill_name, []).append(text[:500])
                else:
                    skills_succeeded.add(skill_name)

        elif etype == "cost-state":
            total_cost_usd = entry.get("totalCostUSD", total_cost_usd)

    total_input = sum((usage_by_request[r].get("input_tokens") or 0) for r in request_order)
    total_output = sum((usage_by_request[r].get("output_tokens") or 0) for r in request_order)
    total_cache_read = sum(
        (usage_by_request[r].get("cache_read_input_tokens") or 0) for r in request_order
    )
    total_cache_creation = sum(
        (usage_by_request[r].get("cache_creation_input_tokens") or 0) for r in request_order
    )

    safexl_commands = [c for c in bash_commands if runner.SAFEXL_MENTION_RE.search(c)]

    return {
        "model": model,
        "num_turns": len(request_order),
        "total_cost_usd": total_cost_usd,
        "tokens": {
            "input_tokens": total_input,
            "output_tokens": total_output,
            "cache_read_input_tokens": total_cache_read,
            "cache_creation_input_tokens": total_cache_creation,
        },
        "final_text": runner._truncate(last_assistant_text, 800) if last_assistant_text else None,
        "tools_called": tools_called,
        "tool_call_count": tool_call_count,
        "skills_invoked": skills_invoked,
        "skills_succeeded": sorted(skills_succeeded),
        "skills_denied": sorted(skills_denied),
        "skill_denial_evidence": skill_denial_evidence,
        "safexl_used": bool(safexl_commands),
        "safexl_commands": safexl_commands,
        "bash_command_count": len(bash_commands),
    }


# ---------------------------------------------------------------------------
# Isolation: each session runs in a fresh temp dir outside the repo
# ---------------------------------------------------------------------------

TMP_ROOT = os.environ.get("XLSX_BENCH_TMP_ROOT", "/tmp")
TMP_PREFIX = "xlsxb-"


def make_isolated_workdir(tmp_root: str = TMP_ROOT) -> Path:
    """A fresh unique dir (``/tmp/xlsxb-XXXXXXXX``) for one session's cwd.
    Outside the repo and outside any tree shared with other runs, so ``ls`` /
    ``find`` from the cwd can only see this run's own workbook. (In the old
    layout the cwd sat in recordings/<round>/<task>/<arm>/<rep>/work, and
    sessions searching for the workbook wandered into sibling runs.)"""
    # hex-only suffix: Claude Code derives the transcript dir name from the cwd,
    # and mkdtemp's underscores are not a case escape_project_dir is verified for.
    while True:
        path = Path(tmp_root) / f"{TMP_PREFIX}{secrets.token_hex(4)}"
        try:
            path.mkdir(mode=0o700)
        except FileExistsError:
            continue
        return path.resolve()


def copy_back(tmp_work: Path, dest_work: Path) -> None:
    """Copy the session's whole work dir into the recordings layout
    (``<cell>/work``) so grade.py finds it unchanged."""
    if dest_work.exists():
        shutil.rmtree(dest_work)
    shutil.copytree(tmp_work, dest_work)


# ---------------------------------------------------------------------------
# Pure-ish helpers: hashing / manifest
# ---------------------------------------------------------------------------


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def build_manifest(
    workbook_name: str,
    workbook_sha256_before: str,
    wdir: Path,
    output_files: list[str],
    cast_path: Path,
    jsonl_path: Optional[Path],
    workbook_sha256_after: Optional[str] = None,
) -> dict:
    """workbook_sha256_before is taken right after the fresh copy, before
    Claude ever touches it. workbook_sha256_after is taken at the very end,
    at the SAME path -- for a task that edits the workbook in place (no new
    "save as" filename, the common case in this corpus), that's the only
    hash of the file actually delivered; without it a naive before/after
    comparison would misread a legitimate in-place edit as "the file got
    swapped." modified_in_place makes that explicit instead of leaving it to
    be inferred by comparing two hashes."""
    manifest: dict = {
        "generated_at": runner._now_iso(),
        "workbook": {
            "path": workbook_name,
            "sha256_before": workbook_sha256_before,
            "sha256_after": workbook_sha256_after,
            "modified_in_place": (
                None
                if workbook_sha256_before is None or workbook_sha256_after is None
                else workbook_sha256_after != workbook_sha256_before
            ),
        },
        "output_files": [],
        "session_cast": None,
        "session_jsonl": None,
    }
    for rel in output_files:
        p = wdir / rel
        if p.is_file():
            manifest["output_files"].append({"path": rel, "sha256": sha256_file(p)})
    if cast_path.exists():
        manifest["session_cast"] = {"path": cast_path.name, "sha256": sha256_file(cast_path)}
    if jsonl_path is not None and jsonl_path.exists():
        manifest["session_jsonl"] = {"path": jsonl_path.name, "sha256": sha256_file(jsonl_path)}
    return manifest


# ---------------------------------------------------------------------------
# tmux control (impure -- not unit tested, mirrors runner.py's live-execution split)
# ---------------------------------------------------------------------------


def _tmux_has_session(name: str) -> bool:
    r = subprocess.run(["tmux", "has-session", "-t", name], capture_output=True)
    return r.returncode == 0


def _tmux_disable_remain_on_exit(name: str) -> None:
    """Force THIS session's pane to disappear when its process exits,
    regardless of the tmux server's global ``remain-on-exit`` default.

    Learned the hard way: an earlier throwaway debugging session on this box
    left ``remain-on-exit`` set globally ON (a shared tmux server persists
    across unrelated work). With it on, a pane whose wrapped process has
    already died is NOT reflected by ``has-session`` -- the session object
    keeps existing, showing its last frame, forever. A real recording then
    silently polls a corpse for the full timeout instead of being caught by
    the spurious-early-death retry in record_cell. Session-scoped
    ``set-option`` overrides the server default for just this session, so
    this tool's own correctness never depends on nobody else having touched
    global tmux state. ``_tmux_session_alive`` below is a second, independent
    check (dead-pane detection) in case this override itself is ever
    ineffective."""
    subprocess.run(["tmux", "set-option", "-t", name, "remain-on-exit", "off"], capture_output=True)


def _tmux_session_alive(name: str) -> bool:
    """True iff the session exists AND its pane's process hasn't exited.
    Belt-and-suspenders alongside ``_tmux_disable_remain_on_exit``: even with
    remain-on-exit off, checking ``pane_dead`` directly means a lingering
    pane (any cause) is never mistaken for a live session."""
    if not _tmux_has_session(name):
        return False
    r = subprocess.run(
        ["tmux", "list-panes", "-t", name, "-F", "#{pane_dead}"], capture_output=True, text=True
    )
    if r.returncode != 0:
        return False
    statuses = [line.strip() for line in r.stdout.splitlines() if line.strip()]
    if not statuses:
        return False
    return all(s == "0" for s in statuses)


def _tmux_capture_pane(name: str) -> str:
    r = subprocess.run(["tmux", "capture-pane", "-p", "-t", name], capture_output=True, text=True)
    return r.stdout if r.returncode == 0 else ""


def _tmux_send_keys(name: str, keys: list[str]) -> None:
    subprocess.run(["tmux", "send-keys", "-t", name, *keys], capture_output=True)


def _tmux_kill_session(name: str) -> None:
    subprocess.run(["tmux", "kill-session", "-t", name], capture_output=True)


def verify_dontask_supported(claude_bin: str) -> bool:
    try:
        r = subprocess.run([claude_bin, "--help"], capture_output=True, text=True, timeout=30)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return "dontAsk" in (r.stdout or "")


# ---------------------------------------------------------------------------
# One recording (impure -- I/O + subprocess + tmux; not unit tested, like
# runner.run_cell)
# ---------------------------------------------------------------------------


def _launch_and_wait(
    claude_argv: list[str],
    wdir: Path,
    arm: str,
    session_name: str,
    timeout: int,
    poll_interval: float,
    quiet_seconds: float,
    asciinema_bin: str,
    cast_path: Path,
) -> dict:
    """One launch attempt: run claude_argv under asciinema in a fresh tmux
    session, poll for dialogs/completion, then exit cleanly. Returns a dict
    of everything record_cell needs -- factored out of record_cell so it can
    be retried (see record_cell's max_launch_attempts) without duplicating
    the poll loop."""
    if cast_path.exists():
        cast_path.unlink()
    asciinema_argv = build_asciinema_argv(cast_path, claude_argv, asciinema_bin=asciinema_bin)

    _tmux_kill_session(session_name)  # clean slate in case a prior attempt/cell left it lingering
    env = runner.build_env(os.environ, arm)
    # A new session on an already-running tmux server takes the SERVER's
    # environment, not this client's, so the arm's env must be passed
    # explicitly with -e (without it, CLAUDE.md loaded in arm A and arm B
    # never got SafeXL on PATH).
    tmux_argv = build_tmux_argv(session_name, wdir, asciinema_argv, env=pane_env(env))

    escaped = escape_project_dir(wdir)
    project_dir = Path.home() / ".claude" / "projects" / escaped
    existing_jsonl = {p.name for p in project_dir.glob("*.jsonl")} if project_dir.exists() else set()

    started_iso = runner._now_iso()
    started = time.monotonic()

    launch = subprocess.run(tmux_argv, env=env, capture_output=True, text=True)
    if launch.returncode != 0:
        raise RuntimeError(f"tmux new-session failed: {launch.stderr.strip()}")
    _tmux_disable_remain_on_exit(session_name)

    dialogs_seen: list[dict] = []
    transcript_path: Optional[Path] = None
    trust_accepted = False
    aborted_reason: Optional[str] = None
    timed_out = False
    session_ended_early = False

    deadline = started + timeout
    while True:
        # Check the transcript for a clean finish BEFORE checking whether the
        # tmux session/pane is still alive. A session that closes itself (or
        # is killed by something external) in the same instant it finishes
        # must not be misreported as a failure when its own transcript
        # already shows a completed, no-pending-tool-use end_turn -- the
        # transcript is the ground truth, the pane is just where we watch
        # for dialogs while waiting.
        if transcript_path is None and project_dir.exists():
            found = [p for p in project_dir.glob("*.jsonl") if p.name not in existing_jsonl]
            if found:
                found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                transcript_path = found[0]

        if transcript_path is not None and transcript_path.exists():
            entries = read_transcript_entries(transcript_path)
            if transcript_is_finished(entries):
                mtime = transcript_path.stat().st_mtime
                if time.time() - mtime >= quiet_seconds:
                    break

        if not _tmux_session_alive(session_name):
            session_ended_early = True
            break

        pane = _tmux_capture_pane(session_name)
        dialog = classify_dialog(pane)
        if dialog:
            if dialog["type"] == "folder_trust":
                if not trust_accepted:
                    dialogs_seen.append(
                        {
                            "type": "folder_trust",
                            "accepted": True,
                            "text": truncate_pane_text(pane),
                            "at": runner._now_iso(),
                        }
                    )
                    _tmux_send_keys(session_name, ["Down"])
                    time.sleep(0.4)
                    _tmux_send_keys(session_name, ["Enter"])
                    trust_accepted = True
            else:
                dialogs_seen.append(
                    {
                        "type": "other",
                        "accepted": False,
                        "text": truncate_pane_text(pane),
                        "at": runner._now_iso(),
                    }
                )
                aborted_reason = "unexpected interactive dialog (not folder-trust) -- see dialogs_seen"
                break

        if time.monotonic() > deadline:
            timed_out = True
            break

        time.sleep(poll_interval)

    if aborted_reason or timed_out or session_ended_early:
        _tmux_kill_session(session_name)
    else:
        _tmux_send_keys(session_name, ["/exit", "Enter"])
        exit_deadline = time.monotonic() + EXIT_GRACE_SECONDS
        while _tmux_session_alive(session_name) and time.monotonic() < exit_deadline:
            time.sleep(1)
        if _tmux_has_session(session_name):
            _tmux_kill_session(session_name)

    return {
        "started_iso": started_iso,
        "started_monotonic": started,
        "wall_seconds": time.monotonic() - started,
        "finished_iso": runner._now_iso(),
        "dialogs_seen": dialogs_seen,
        "transcript_path": transcript_path,
        "aborted_reason": aborted_reason,
        "timed_out": timed_out,
        "session_ended_early": session_ended_early,
    }


def record_cell(
    task: "runner.Task",
    arm: str,
    rep: int,
    cdir: Path,
    claude_bin: str,
    block_skill_method: str,
    timeout: int,
    session_name: str,
    poll_interval: float = DEFAULT_POLL_INTERVAL,
    quiet_seconds: float = DEFAULT_QUIET_SECONDS,
    asciinema_bin: str = ASCIINEMA_BIN_DEFAULT,
    max_launch_attempts: int = 4,
    relaunch_backoff_seconds: float = RELAUNCH_BACKOFF_SECONDS,
) -> dict:
    final_wdir = runner.work_dir(cdir)
    if final_wdir.exists():
        shutil.rmtree(final_wdir)
    # The session runs in a private temp dir; outputs are copied back to
    # final_wdir afterwards. wdir is the temp dir for the rest of the run.
    wdir = make_isolated_workdir()

    workbook_src = runner.resolve_workbook(task)
    if not workbook_src.exists():
        shutil.rmtree(wdir, ignore_errors=True)
        raise FileNotFoundError(f"workbook not found: {workbook_src}")
    workbook_dst = wdir / workbook_src.name
    shutil.copy2(workbook_src, workbook_dst)
    workbook_sha256 = sha256_file(workbook_dst)

    before_files = runner.snapshot_dir(wdir)

    claude_argv = interactive_command(claude_bin, task.prompt, arm, block_skill_method)
    cast_path = cdir / "session.cast"

    # Observed in practice on this box: an interactive `claude` launch
    # sometimes dies within ~3s with zero transcript ever created (no
    # evidence a turn was even queued) -- sporadically, not tied to a
    # specific arm or "only the very first launch ever": it happened once on
    # arm A's first-ever launch, then not again for many launches, then
    # twice in a row on an unrelated arm B attempt. The next launch with the
    # *identical* command always worked normally, so this reads as
    # environment flakiness (this box's tmux/asciinema/claude startup path),
    # not a defect in the command being built. Retrying that *specific*
    # signature (session gone, no transcript at all) is cheap and safe;
    # default budget is 4 attempts with an 8s backoff between them, since 2
    # rapid-fire attempts were empirically not always enough for arm B (which
    # loads a plugin via --plugin-dir -- possibly a one-time plugin-cache/
    # sync cost that a too-quick immediate retry doesn't leave room for). A
    # real timeout or a disallowed dialog is never retried.
    relaunch_events: list[dict] = []
    attempt_result = None
    used_session_name = session_name
    for attempt in range(1, max_launch_attempts + 1):
        # A distinct session name per attempt (rather than reusing one name
        # across retries) rules out any tmux-level race from killing and
        # immediately recreating the *same* named session/pty in a tight
        # loop -- cheap insurance, and it means a lingering dead session
        # from attempt N can never be confused with attempt N+1's.
        used_session_name = session_name if attempt == 1 else f"{session_name}-r{attempt}"
        attempt_result = _launch_and_wait(
            claude_argv, wdir, arm, used_session_name, timeout, poll_interval, quiet_seconds,
            asciinema_bin, cast_path,
        )
        spurious = attempt_result["session_ended_early"] and attempt_result["transcript_path"] is None
        if spurious and attempt < max_launch_attempts:
            relaunch_events.append(
                {
                    "attempt": attempt,
                    "wall_seconds": round(attempt_result["wall_seconds"], 2),
                    "reason": "tmux session ended within the first poll with no transcript ever "
                    "created -- treated as a spurious first-launch failure and retried",
                }
            )
            time.sleep(relaunch_backoff_seconds)
            continue
        break

    session_name = used_session_name
    started_iso = attempt_result["started_iso"]
    wall_seconds = attempt_result["wall_seconds"]
    finished_iso = attempt_result["finished_iso"]
    dialogs_seen = attempt_result["dialogs_seen"]
    transcript_path = attempt_result["transcript_path"]
    aborted_reason = attempt_result["aborted_reason"]
    timed_out = attempt_result["timed_out"]
    session_ended_early = attempt_result["session_ended_early"]

    entries = read_transcript_entries(transcript_path) if transcript_path else []
    summary = summarize_transcript(entries)
    rate_limit_evidence = runner.detect_rate_limit(entries, "")
    contamination = runner.detect_contamination(entries, own_tmp=str(wdir))

    after_files = runner.snapshot_dir(wdir)
    output_files = runner.new_files(before_files, after_files)
    workbook_sha256_after = sha256_file(workbook_dst) if workbook_dst.exists() else None

    # Bring the outputs home and drop the temp dir. Hash the copy for the manifest.
    copy_back(wdir, final_wdir)
    isolated_cwd = str(wdir)
    shutil.rmtree(wdir, ignore_errors=True)
    workbook_dst = final_wdir / workbook_dst.name
    wdir = final_wdir

    session_jsonl_path: Optional[Path] = None
    if transcript_path is not None and transcript_path.exists():
        session_jsonl_path = cdir / "session.jsonl"
        shutil.copy2(transcript_path, session_jsonl_path)

    result = {
        "task_id": task.id,
        "arm": arm,
        "rep": rep,
        "workbook": task.workbook,
        "prompt": task.prompt,
        "started_at": started_iso,
        "finished_at": finished_iso,
        "wall_seconds": round(wall_seconds, 2),
        "timed_out": timed_out,
        "session_ended_early": session_ended_early,
        "aborted_reason": aborted_reason,
        "dialogs_seen": dialogs_seen,
        "relaunch_events": relaunch_events,
        "output_files": output_files,
        "rate_limited": rate_limit_evidence is not None,
        "rate_limit_evidence": rate_limit_evidence,
        "contamination": contamination,
        "cast_file": "session.cast" if cast_path.exists() else None,
        "transcript_file": "session.jsonl" if session_jsonl_path else None,
        "tmux_session": session_name,
        "invocation": {
            "claude_argv": claude_argv,
            "cwd": isolated_cwd,
            "work_dir_copied_to": str(final_wdir),
            "isolated": True,
            "deny_rules": runner.bench_tree_deny_rules(),
            "block_skill_method": block_skill_method if arm == "B" else None,
            "plugin_dir": runner.SAFEXL_PLUGIN_DIR if arm == "B" else None,
            "path_prepend": [runner.SAFEXL_VENV_BIN] if arm == "B" else [],
            "anthropic_api_key_set": False,
            "claude_md_disabled": True,
            "strict_mcp_config": True,
            "permission_mode": "dontAsk",
        },
        "recorder_version": RECORDER_VERSION,
    }
    result.update(summary)

    manifest = build_manifest(
        workbook_dst.name, workbook_sha256, wdir, output_files, cast_path, session_jsonl_path,
        workbook_sha256_after,
    )
    (cdir / "manifest.json").write_text(json.dumps(manifest, indent=2, sort_keys=False) + "\n")

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="record.py",
        description="Record an interactive xlsx-bench Claude Code session (arm A vs arm B).",
    )
    parser.add_argument("--tasks", type=Path, default=SCRIPT_DIR / "tasks_run.json")
    parser.add_argument("--task", type=str, help="single task id (mutually exclusive with --all)")
    parser.add_argument("--arm", type=str, choices=runner.ARMS, help="single arm (mutually exclusive with --all)")
    parser.add_argument("--rep", type=int, default=1)
    parser.add_argument("--out", type=Path, default=SCRIPT_DIR / "recordings")
    parser.add_argument("--all", action="store_true", help="batch mode: record every (task,arm,rep) cell, resumable")
    parser.add_argument("--arms", type=str, default="A,B", help="(--all only) which arms to record")
    parser.add_argument("--reps", type=int, default=1, help="(--all only) repetitions per (task,arm)")
    parser.add_argument("--claude-bin", type=str, default=runner.CLAUDE_BIN_DEFAULT)
    parser.add_argument("--asciinema-bin", type=str, default=ASCIINEMA_BIN_DEFAULT)
    parser.add_argument(
        "--block-skill-method", choices=("disallowed", "settings"), default="disallowed",
    )
    parser.add_argument(
        "--timeout", type=int, default=runner.DEFAULT_TIMEOUT_SECONDS,
        help=f"per-run wall-clock timeout in seconds (default {runner.DEFAULT_TIMEOUT_SECONDS})",
    )
    parser.add_argument(
        "--session-name", type=str, default=None,
        help="tmux session name to use (default: xlsx-rec-<random short id>)",
    )
    parser.add_argument("--poll-interval", type=float, default=DEFAULT_POLL_INTERVAL)
    parser.add_argument("--quiet-seconds", type=float, default=DEFAULT_QUIET_SECONDS)
    parser.add_argument(
        "--max-launch-attempts", type=int, default=4,
        help="retry budget for a launch that dies within the first poll with zero transcript "
        "ever created (observed sporadic environment flakiness, not a command defect)",
    )
    parser.add_argument("--status", action="store_true", help="print done/total per arm and exit")
    return parser


def _find_task(tasks: list["runner.Task"], task_id: str) -> "runner.Task":
    for t in tasks:
        if t.id == task_id:
            return t
    raise ValueError(f"unknown task id {task_id!r}")


def main(argv: Optional[list[str]] = None) -> int:
    raw_argv = list(argv) if argv is not None else sys.argv[1:]
    parser = build_arg_parser()
    args = parser.parse_args(raw_argv)

    try:
        tasks = runner.load_tasks(args.tasks)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2

    if args.status:
        try:
            arms = runner.parse_arms(args.arms)
        except ValueError as e:
            print(f"error: {e}", file=sys.stderr)
            return 2
        print(runner.status_report(tasks, arms, args.reps, args.out))
        return 0

    if not Path(args.claude_bin).exists():
        print(f"error: claude binary not found at {args.claude_bin}", file=sys.stderr)
        return 2
    if not verify_dontask_supported(args.claude_bin):
        print(
            f"error: {args.claude_bin} --help does not mention permission mode "
            "\"dontAsk\" -- refusing to record without it (a disallowed tool call "
            "would otherwise hang on an interactive permission prompt)",
            file=sys.stderr,
        )
        return 2
    if shutil.which("tmux") is None:
        print("error: tmux is not installed", file=sys.stderr)
        return 2

    session_name = args.session_name or f"xlsx-rec-{uuid.uuid4().hex[:8]}"
    # Absolute: the tmux pane starts in the work dir, so a relative cast path
    # (or -c start dir) would resolve somewhere else and asciinema would exit
    # instantly with nothing recorded.
    out_dir: Path = Path(args.out).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.all:
        if args.task or args.arm:
            print("error: --task/--arm cannot be combined with --all", file=sys.stderr)
            return 2
        try:
            arms = runner.parse_arms(args.arms)
            cells = runner.build_matrix(tasks, arms, args.reps)
        except ValueError as e:
            print(f"error: {e}", file=sys.stderr)
            return 2

        stopped_path = out_dir / "STOPPED.json"
        if stopped_path.exists():
            print(
                f"error: {stopped_path} exists from a previously stopped batch; "
                "review it and remove it before recording again",
                file=sys.stderr,
            )
            return 2

        total = len(cells)
        done_count = 0
        for i, cell in enumerate(cells, 1):
            cdir = runner.run_dir(out_dir, cell.task.id, cell.arm, cell.rep)
            label = f"{cell.task.id}/{cell.arm}/rep{cell.rep}"
            if runner.is_done(cdir):
                done_count += 1
                print(f"[{i}/{total}] {label}: already done, skipping")
                continue

            cdir.mkdir(parents=True, exist_ok=True)
            print(f"[{i}/{total}] {label}: recording (timeout {args.timeout}s)...")
            result = record_cell(
                cell.task, cell.arm, cell.rep, cdir, args.claude_bin, args.block_skill_method,
                args.timeout, session_name, args.poll_interval, args.quiet_seconds, args.asciinema_bin,
                args.max_launch_attempts,
            )
            runner.write_run_json(cdir, result)
            print(
                f"    wall={result.get('wall_seconds')}s model={result.get('model')} "
                f"dialogs={len(result.get('dialogs_seen') or [])} "
                f"aborted={result.get('aborted_reason')!r}"
            )

            if result.get("rate_limited"):
                stopped = {
                    "stopped_at": runner._now_iso(),
                    "reason": result.get("rate_limit_evidence") or "usage/rate limit detected in output",
                    "cell": {"task": cell.task.id, "arm": cell.arm, "rep": cell.rep},
                    "run_dir": str(cdir),
                }
                stopped_path.write_text(json.dumps(stopped, indent=2) + "\n")
                print(f"RATE LIMIT DETECTED -- stopping batch. See {stopped_path}", file=sys.stderr)
                return 1

            if result.get("aborted_reason"):
                stopped = {
                    "stopped_at": runner._now_iso(),
                    "reason": result["aborted_reason"],
                    "dialogs_seen": result.get("dialogs_seen"),
                    "cell": {"task": cell.task.id, "arm": cell.arm, "rep": cell.rep},
                    "run_dir": str(cdir),
                }
                stopped_path.write_text(json.dumps(stopped, indent=2) + "\n")
                print(f"UNEXPECTED DIALOG -- stopping batch. See {stopped_path}", file=sys.stderr)
                return 1

            done_count += 1

        print(f"Done: {done_count}/{total}")
        return 0

    # single-cell mode
    if not args.task or not args.arm:
        print("error: --task and --arm are required (or use --all)", file=sys.stderr)
        return 2
    try:
        task = _find_task(tasks, args.task)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2

    cdir = runner.run_dir(out_dir, task.id, args.arm, args.rep)
    label = f"{task.id}/{args.arm}/rep{args.rep}"
    if runner.is_done(cdir):
        print(f"{label}: already done, skipping (remove {cdir / 'run.json'} to re-record)")
        return 0

    cdir.mkdir(parents=True, exist_ok=True)
    print(f"{label}: recording (timeout {args.timeout}s)...")
    result = record_cell(
        task, args.arm, args.rep, cdir, args.claude_bin, args.block_skill_method,
        args.timeout, session_name, args.poll_interval, args.quiet_seconds, args.asciinema_bin,
        args.max_launch_attempts,
    )
    runner.write_run_json(cdir, result)
    print(
        f"    wall={result.get('wall_seconds')}s model={result.get('model')} "
        f"dialogs={len(result.get('dialogs_seen') or [])} "
        f"aborted={result.get('aborted_reason')!r}"
    )
    if (
        result.get("aborted_reason")
        or result.get("timed_out")
        or result.get("rate_limited")
        or result.get("session_ended_early")
    ):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
