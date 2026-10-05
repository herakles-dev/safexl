#!/usr/bin/env python3
"""Two-arm head-to-head benchmark runner for Excel-editing agents.

Arm A: Claude Code with Anthropic's built-in `xlsx` skill (`anthropic-skills:xlsx`).
Arm B: Claude Code with the SafeXL plugin (`safexl` CLI), with the xlsx skill
       blocked so the model cannot fall back to it.

Same model (whatever the authenticated CLI defaults to), same prompts, a fresh
work folder per (task, arm, rep). Results feed a customer-facing comparison, so
this module is deliberately paranoid about two things:

1. Arm isolation — arm A must never be able to reach SafeXL, and arm B must
   never be able to *successfully* invoke the Anthropic xlsx skill. See
   ``build_env`` and ``skill_block_args``.
2. Logging — every run gets its own ``events.jsonl`` (raw stream-json),
   ``stderr.txt``, and a parsed ``run.json`` summary, so a claim like "arm B
   never used the xlsx skill" can be checked against the raw transcript, not
   just our parsing of it.

Stdlib only. See README.md for usage and the isolation write-up.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shlex
import shutil
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
# research/xlsx-bench -> research -> repo root. Used to resolve workbook paths
# in the tasks file (they're written relative to the repo root) regardless of
# which worktree/checkout this script happens to run from.
REPO_ROOT = SCRIPT_DIR.parent.parent

CLAUDE_BIN_DEFAULT = str(Path.home() / ".local" / "bin" / "claude")

# Identical for both arms: this is the whole point of the comparison. Passed
# as ONE argument (space-separated), matching the working precedent in
# research/claude-xlsx-skill/run_claude_code.sh.
ALLOWED_TOOLS = "Skill Read Write Edit Bash Glob Grep"

ANTHROPIC_XLSX_SKILL = "anthropic-skills:xlsx"

# Paths default to this checkout (resolved from this file), so a fresh clone works.
# XLSX_BENCH_PLUGIN_DIR points arm B at another copy of the plugin (e.g. a frozen
# `git archive` of one commit); XLSX_BENCH_SAFEXL_BIN points it at another
# virtualenv bin directory. Neither is ever written to by this script.
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
SAFEXL_PLUGIN_DIR = os.environ.get(
    "XLSX_BENCH_PLUGIN_DIR", str(_REPO_ROOT / "packages" / "safexl" / "claude-plugin")
)
SAFEXL_VENV_BIN = os.environ.get(
    "XLSX_BENCH_SAFEXL_BIN", str(_REPO_ROOT / "packages" / "safexl" / ".venv" / "bin")
)

DEFAULT_TIMEOUT_SECONDS = 20 * 60
TMUX_SESSION_NAME = "xlsx-bench"

ARMS = ("A", "B")
TASK_ID_RE = re.compile(r"^[A-Za-z0-9_-]+$")
SAFEXL_MENTION_RE = re.compile(r"safexl", re.IGNORECASE)
DENIAL_TEXT_RE = re.compile(
    r"permission|denied|disallowed|not allowed|not permitted|requested permission",
    re.IGNORECASE,
)

# Deliberately broad. A false positive here just stops the batch early
# (recoverable: rerun later); a false negative burns plan usage into a wall.
RATE_LIMIT_PATTERNS = [
    re.compile(p, re.IGNORECASE)
    for p in [
        r"usage limit",
        r"rate.?limit",
        # 429 only with rate-limit context: a bare \b429\b matched a 429-byte file in an
        # `unzip -l` listing and stopped a healthy run (2026-09-26).
        r"(?:http|status|error|code)\W{0,3}429\b",
        r"\b429\b\W{0,3}(?:too many|rate)",
        r"session limit",
        r"quota exceeded",
        r"exceeded your[^.\n]{0,40}limit",
        r"try again later",
        r"you've hit your",
    ]
]

RUNNER_VERSION = "1.0"


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Task:
    id: str
    workbook: str
    prompt: str


@dataclass(frozen=True)
class Cell:
    task: Task
    arm: str
    rep: int


# ---------------------------------------------------------------------------
# Tasks file parsing (pure — no I/O beyond the initial read)
# ---------------------------------------------------------------------------


def parse_tasks(data: object) -> list[Task]:
    """Validate and convert decoded tasks-file JSON into a list of Task.

    Raises ValueError with a human-readable message on any schema problem.
    """
    if not isinstance(data, list):
        raise ValueError("tasks file must contain a JSON list")
    if not data:
        raise ValueError("tasks file must contain at least one task")

    tasks: list[Task] = []
    seen_ids: set[str] = set()
    for i, entry in enumerate(data):
        if not isinstance(entry, dict):
            raise ValueError(f"task[{i}] must be a JSON object")
        missing = [k for k in ("id", "workbook", "prompt") if k not in entry]
        if missing:
            raise ValueError(f"task[{i}] missing required field(s): {', '.join(missing)}")

        task_id, workbook, prompt = entry["id"], entry["workbook"], entry["prompt"]

        if not isinstance(task_id, str) or not task_id:
            raise ValueError(f"task[{i}].id must be a non-empty string")
        if not TASK_ID_RE.match(task_id):
            raise ValueError(
                f"task[{i}].id {task_id!r} must match {TASK_ID_RE.pattern!r} "
                "(it's used as a directory name)"
            )
        if task_id in seen_ids:
            raise ValueError(f"duplicate task id: {task_id!r}")
        seen_ids.add(task_id)

        if not isinstance(workbook, str) or not workbook:
            raise ValueError(f"task[{i}].workbook must be a non-empty string")
        if not isinstance(prompt, str) or not prompt.strip():
            raise ValueError(f"task[{i}].prompt must be a non-empty string")

        tasks.append(Task(id=task_id, workbook=workbook, prompt=prompt))
    return tasks


def load_tasks(path: Path) -> list[Task]:
    try:
        raw = path.read_text()
    except OSError as e:
        raise ValueError(f"cannot read tasks file {path}: {e}") from e
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"tasks file {path} is not valid JSON: {e}") from e
    return parse_tasks(data)


def resolve_workbook(task: Task) -> Path:
    p = Path(task.workbook)
    return p if p.is_absolute() else REPO_ROOT / p


def parse_arms(spec: str) -> list[str]:
    arms = [a.strip().upper() for a in spec.split(",") if a.strip()]
    if not arms:
        raise ValueError("--arms must name at least one arm")
    bad = [a for a in arms if a not in ARMS]
    if bad:
        raise ValueError(f"unknown arm(s) {bad!r}; valid arms are {ARMS}")
    seen: list[str] = []
    for a in arms:
        if a not in seen:
            seen.append(a)
    return seen


def build_matrix(tasks: list[Task], arms: list[str], reps: int) -> list[Cell]:
    if reps < 1:
        raise ValueError("--reps must be >= 1")
    return [
        Cell(task=task, arm=arm, rep=rep)
        for task in tasks
        for arm in arms
        for rep in range(1, reps + 1)
    ]


# ---------------------------------------------------------------------------
# Run-directory layout + resumability
# ---------------------------------------------------------------------------


def run_dir(out_dir: Path, task_id: str, arm: str, rep: int) -> Path:
    return out_dir / task_id / arm / str(rep)


def work_dir(cell_dir: Path) -> Path:
    return cell_dir / "work"


def run_json_path(cell_dir: Path) -> Path:
    return cell_dir / "run.json"


def is_done(cell_dir: Path) -> bool:
    """A cell is done, and skippable on resume, iff it has a run.json that
    parses as JSON. run.json is written atomically (temp file + rename, see
    write_run_json) so a half-written file is never mistaken for a finished
    one, even if the runner was killed mid-write."""
    p = run_json_path(cell_dir)
    if not p.exists():
        return False
    try:
        json.loads(p.read_text())
        return True
    except (OSError, json.JSONDecodeError):
        return False


def pending_cells(cells: list[Cell], out_dir: Path) -> list[Cell]:
    return [
        c for c in cells if not is_done(run_dir(out_dir, c.task.id, c.arm, c.rep))
    ]


def status_report(tasks: list[Task], arms: list[str], reps: int, out_dir: Path) -> str:
    lines: list[str] = []
    grand_done = grand_total = 0
    for arm in arms:
        arm_total = arm_done = 0
        missing: list[str] = []
        for task in tasks:
            for rep in range(1, reps + 1):
                arm_total += 1
                cdir = run_dir(out_dir, task.id, arm, rep)
                if is_done(cdir):
                    arm_done += 1
                else:
                    missing.append(f"{task.id}/rep{rep}")
        grand_done += arm_done
        grand_total += arm_total
        lines.append(f"Arm {arm}: {arm_done}/{arm_total} done")
        if missing:
            lines.append(f"  pending: {', '.join(missing)}")
    lines.append(f"Total: {grand_done}/{grand_total} done")
    stopped = out_dir / "STOPPED.json"
    if stopped.exists():
        lines.append(f"STOPPED: {stopped} exists -- batch was halted, see reason inside")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Command + environment construction (pure)
# ---------------------------------------------------------------------------


# XLSX_BENCH_MODEL runs the same clean room on another model (e.g. the Haiku round);
# run.json records the model each session actually used.
BENCH_MODEL = os.environ.get("XLSX_BENCH_MODEL", "claude-opus-5-5")


def plugin_hooks_enabled(arm: Optional[str]) -> bool:
    """XLSX_BENCH_PLUGIN_HOOKS=1: arm B runs with hooks on so the plugin's own
    guard hook can fire. Off by default; arm A never gets hooks."""
    return arm == "B" and os.environ.get("XLSX_BENCH_PLUGIN_HOOKS") == "1"


def clean_room_settings(arm: Optional[str] = None) -> dict:
    """Settings overlay for both arms: no hooks, and every user-enabled plugin
    switched off (arm B adds SafeXL separately via --plugin-dir)."""
    user_settings = Path.home() / ".claude" / "settings.json"
    try:
        enabled = json.loads(user_settings.read_text()).get("enabledPlugins", {}) or {}
    except (OSError, ValueError):
        enabled = {}
    return {
        "disableAllHooks": not plugin_hooks_enabled(arm),
        "enabledPlugins": {name: False for name in enabled},
        "permissions": {"deny": bench_tree_deny_rules()},
    }


def bench_tree_deny_rules(bench_dir: Optional[Path] = None) -> list[str]:
    """Permission deny rules that keep a session out of this bench tree
    (recordings, runs, results, corpus, tasks, grader). Same rules for both
    arms. A run's cwd is a throwaway dir outside the repo (see
    record.isolated_workdir), so a session has no reason to go here; these
    rules stop it if it wanders anyway.

    Path grammar: ``//abs/path`` is a filesystem-absolute path in Claude Code
    permission rules (a single leading ``/`` means project-relative). File
    tools (Read/Edit/Write and, best effort, Grep/Glob) honour Read/Edit
    rules; Bash is matched on the command string, so ``Bash(*xlsx-bench*)``
    blocks any command that names the tree. The SafeXL plugin dir and venv
    live under packages/, outside this tree, and stay usable.
    """
    root = str((bench_dir or SCRIPT_DIR).resolve())
    return [
        f"Read(/{root}/**)",
        f"Edit(/{root}/**)",
        f"Read(/{root})",
        "Bash(*xlsx-bench*)",
    ]


def skill_block_args(method: str) -> list[str]:
    """Args that block the Anthropic xlsx skill for arm B.

    'disallowed' -> --disallowedTools "Skill(anthropic-skills:xlsx)"
        Rule-content grammar Skill(<name>) is the same one Claude Code uses
        for Skill *allow* rules (e.g. the built-in "Skill(anthropic-skills:*)"
        rule); deny rules use the same grammar. This is the first thing to
        try and, per the smoke run, is what actually blocks it -- see README.

    'settings' -> --settings '{"permissions": {"deny": ["Skill(...)"]}}'
        Fallback if disallowedTools ever stops working: same rule string,
        delivered as an inline settings JSON blob instead of a CLI list.
    """
    if method == "disallowed":
        return ["--disallowedTools", f"Skill({ANTHROPIC_XLSX_SKILL})"]
    if method == "settings":
        settings = {"permissions": {"deny": [f"Skill({ANTHROPIC_XLSX_SKILL})"]}}
        return ["--settings", json.dumps(settings)]
    raise ValueError(f"unknown block-skill method: {method!r}")


# Operator-configured round (XLSX_BENCH_OPERATOR_HINT=1): the same one-line
# instruction a team deploying an agent would add, symmetric across arms.
# Added after the Haiku round, where Haiku never used SafeXL (0/8) and loaded
# Anthropic's skill in only 4/8 runs, so the plain round measured bare Haiku.
OPERATOR_HINTS = {
    "A": "When a task involves an Excel workbook, use the xlsx skill for the edit.",
    "B": "When a task involves an Excel workbook, use SafeXL (its skill and the safexl command) for the edit.",
}


def build_command(
    claude_bin: str,
    prompt: str,
    arm: str,
    block_skill_method: str,
    plugin_dir: str = SAFEXL_PLUGIN_DIR,
) -> list[str]:
    # Clean room: load user settings (Anthropic's xlsx skill lives in the user
    # skills dir, so dropping "user" removed it from arm A) but never the
    # project's, then switch off everything that isn't a clean install: all
    # hooks and the user's other plugins. The model is pinned because it also
    # comes from user settings (without "user", runs fell back to Sonnet 5).
    cmd = [
        claude_bin, "-p", prompt, "--allowedTools", ALLOWED_TOOLS,
        "--model", BENCH_MODEL,
        # Hooks-on arm B drops the `user` source: ~/.claude/settings.json carries
        # user hooks (session spawners, mailbox checks) that must not fire, and the
        # plugin's own hooks load from --plugin-dir, not from settings.
        "--setting-sources", "local" if plugin_hooks_enabled(arm) else "user,local",
    ]
    clean = clean_room_settings(arm)
    if arm == "B" and block_skill_method == "settings":
        clean["permissions"]["deny"].append(f"Skill({ANTHROPIC_XLSX_SKILL})")
    cmd += ["--settings", json.dumps(clean)]
    if arm == "B":
        cmd += ["--plugin-dir", plugin_dir]
        if block_skill_method != "settings":
            cmd += skill_block_args(block_skill_method)
    if os.environ.get("XLSX_BENCH_OPERATOR_HINT") == "1":
        cmd += ["--append-system-prompt", OPERATOR_HINTS[arm]]
    cmd += ["--output-format", "stream-json", "--verbose"]
    return cmd


def build_env(base_env: dict[str, str], arm: str) -> dict[str, str]:
    """Per-arm subprocess environment.

    Both arms:
      - ANTHROPIC_API_KEY is stripped (never set it -- OAuth/subscription
        auth only; the runs use a Claude subscription and no API key).
      - CLAUDE_CODE_DISABLE_CLAUDE_MDS=1: the fresh work/ folder lives inside
        this git checkout, so without this, Claude Code would auto-discover
        the repo's (and the user's) CLAUDE.md and inject it into the system
        prompt -- leaking unrelated private project notes into a benchmark
        whose logs go to customers, and
        biasing arm A by telling it SafeXL exists. This does NOT disable
        skills, plugins, or hooks -- only CLAUDE.md/AGENTS.md/rules text
        (confirmed against the CLI binary's own guard checks).

    Arm A only: PATH has any SafeXL venv bin directory stripped out, so even
      if it were already on PATH, the `safexl` CLI would not resolve; and no
      CLAUDE_ENV_FILE.
    Arm B only: the SafeXL venv bin directory is prepended to PATH, AND
      CLAUDE_ENV_FILE points at a script that prepends it again. Claude
      Code's Bash tool rebuilds PATH from the user's shell startup files, so
      the process PATH alone never reaches Bash calls (`command -v safexl`
      returned NOT-FOUND without the env file, the venv path with it,
      2026-09-26). The env file mirrors a real install, where `safexl`
      is on the user's PATH.
    """
    env = dict(base_env)
    env.pop("ANTHROPIC_API_KEY", None)
    env.pop("CLAUDE_ENV_FILE", None)
    env["CLAUDE_CODE_DISABLE_CLAUDE_MDS"] = "1"

    safexl_bin = str(Path(SAFEXL_VENV_BIN))
    parts = [p for p in env.get("PATH", "").split(os.pathsep) if p]
    parts = [p for p in parts if str(Path(p)) != safexl_bin]
    if arm == "B":
        parts = [safexl_bin] + parts
        env["CLAUDE_ENV_FILE"] = str(arm_b_env_file())
    env["PATH"] = os.pathsep.join(parts)
    return env


def arm_b_env_file() -> Path:
    """The CLAUDE_ENV_FILE arm B sources before every Bash call: puts the
    SafeXL venv's bin first on PATH. Written on demand (machine-specific
    absolute path, so it isn't committed)."""
    path = SCRIPT_DIR / ".arm_b_env.sh"
    content = f'export PATH={shlex.quote(str(Path(SAFEXL_VENV_BIN)))}:"$PATH"\n'
    if not path.exists() or path.read_text() != content:
        path.write_text(content)
    return path


# ---------------------------------------------------------------------------
# stream-json event parsing (pure)
# ---------------------------------------------------------------------------


def iter_json_lines(text: str):
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            yield json.loads(line)
        except json.JSONDecodeError:
            continue  # tolerate stray non-JSON stdout noise


def _tool_result_text(block: dict) -> str:
    content = block.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for c in content:
            if isinstance(c, dict) and c.get("type") == "text":
                parts.append(c.get("text", ""))
            elif isinstance(c, str):
                parts.append(c)
        return "\n".join(parts)
    return ""


def _looks_like_denial(text: str) -> bool:
    return bool(text) and bool(DENIAL_TEXT_RE.search(text))


def _first_present(d: dict, keys: tuple[str, ...], default=None):
    for k in keys:
        if k in d and d[k] is not None:
            return d[k]
    return default


def _truncate(text: str, n: int) -> str:
    text = text.strip()
    if len(text) <= n:
        return text
    return text[:n].rstrip() + "…"


def summarize_events(events: list[dict]) -> dict:
    """Turn a list of decoded stream-json events into the run.json summary
    fields. Pure function -- no I/O, no subprocess -- so it's easy to unit
    test against small hand-built fixtures.
    """
    model = None
    total_cost_usd = None
    num_turns = None
    result_subtype = None
    is_error = None
    result_text = None
    last_assistant_text = None

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

    for event in events:
        etype = event.get("type")

        if etype == "system" and event.get("subtype") == "init":
            model = event.get("model", model)

        elif etype == "assistant":
            message = event.get("message") or {}
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
            message = event.get("message") or {}
            for block in message.get("content") or []:
                if not isinstance(block, dict):
                    continue
                if block.get("type") != "tool_result":
                    continue
                tool_use_id = block.get("tool_use_id")
                call = tool_calls_by_id.get(tool_use_id) if tool_use_id else None
                if not call or call["name"] != "Skill":
                    continue
                skill_name = call["input"].get("skill") or call["input"].get("name")
                if not skill_name:
                    continue
                text = _tool_result_text(block)
                errored = bool(block.get("is_error")) or _looks_like_denial(text)
                if errored:
                    skills_denied.add(skill_name)
                    skill_denial_evidence.setdefault(skill_name, []).append(text[:500])
                else:
                    skills_succeeded.add(skill_name)

        elif etype == "result":
            result_subtype = event.get("subtype", result_subtype)
            is_error = event.get("is_error", is_error)
            total_cost_usd = _first_present(event, ("total_cost_usd", "cost_usd"), total_cost_usd)
            num_turns = _first_present(event, ("num_turns", "turns"), num_turns)
            result_text = event.get("result", result_text)

    final_text = result_text if result_text is not None else last_assistant_text
    safexl_commands = [c for c in bash_commands if SAFEXL_MENTION_RE.search(c)]

    return {
        "model": model,
        "total_cost_usd": total_cost_usd,
        "num_turns": num_turns,
        "result_subtype": result_subtype,
        "is_error": is_error,
        "final_text": _truncate(final_text, 800) if final_text else None,
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


def _text_of_content(content) -> str:
    if isinstance(content, str):
        return content
    parts = []
    for block in content or []:
        if isinstance(block, dict) and block.get("type") == "text":
            parts.append(block.get("text", ""))
    return "\n".join(parts)


def _rate_limit_signal_texts(events: list[dict]) -> list[str]:
    """Text from the CLI's / API's own status output only. Never tool results,
    never ordinary model text: a session that reads a file mentioning a rate
    limit (e.g. another run's run.json) is not rate limited (false stop,
    2026-09-29). Real signals:

      - stream-json ``rate_limit_event`` with status "rejected"
      - stream-json ``result`` that is an error
      - a synthetic API-error assistant message (``isApiErrorMessage``, an
        ``error`` field, ``model == "<synthetic>"``, or ``quotaLimits`` status
        "rejected") -- this is how the interactive transcript records
        "You've hit your session limit"
      - ``system`` ``api_retry`` events
    """
    out: list[str] = []
    for event in events:
        etype = event.get("type")
        if etype == "rate_limit_event":
            info = event.get("rate_limit_info") or {}
            if info.get("status") == "rejected":
                out.append("rate limit rejected: " + json.dumps(info)[:300])
        elif etype == "result":
            if event.get("is_error") or str(event.get("subtype", "")).startswith("error"):
                out.append(str(event.get("result") or "") + " " + str(event.get("subtype") or ""))
        elif etype == "assistant":
            message = event.get("message") or {}
            quota = event.get("quotaLimits") or {}
            if quota.get("status") == "rejected":
                out.append("rate limit rejected (quotaLimits): " + _text_of_content(message.get("content")))
            if event.get("isApiErrorMessage") or event.get("error") or message.get("model") == "<synthetic>":
                out.append(_text_of_content(message.get("content")) + " " + str(event.get("error") or ""))
        elif etype == "system" and event.get("subtype") == "api_retry":
            out.append(json.dumps(event)[:500])
    return out


def detect_rate_limit(events: list[dict], stderr_text: str = "") -> Optional[str]:
    """Return a short evidence snippet if the session's own status/error
    output (CLI stderr, API-error messages, error results) looks like a
    usage/rate limit, else None. Tool-result text is deliberately ignored."""
    combined = "\n".join(t for t in [*_rate_limit_signal_texts(events), stderr_text or ""] if t)

    for pattern in RATE_LIMIT_PATTERNS:
        m = pattern.search(combined)
        if m:
            start = max(m.start() - 80, 0)
            end = min(m.end() + 80, len(combined))
            return combined[start:end].strip()
    return None


# ---------------------------------------------------------------------------
# Contamination detection (pure): did a session touch the bench tree or
# another run's temp dir?
# ---------------------------------------------------------------------------

BENCH_TREE_RE = re.compile(r"[^\s\"'`<>()\[\]{},;|&=]*xlsx-bench[^\s\"'`<>()\[\]{},;|&]*")
OTHER_TMP_RE = re.compile(r"/tmp/xlsxb-[A-Za-z0-9_]+")
# In the old in-repo layout (cwd = <cell>/work) a relative climb reaches sibling runs.
LEGACY_CLIMB_RE = re.compile(r"\.\./(?:\.\./)*(?:manifest\.json|run\.json|session\.|events\.|[AB]/|[0-9]+/)")
# The arm B CLAUDE_ENV_FILE lives in the bench dir and is sourced by the
# harness, not read by the model; its name showing up in `env` output is benign.
CONTAMINATION_IGNORE = (".arm_b_env.sh",)
# A grep-style output line that starts with a bench path followed by ':' (a match)
# or by a file extension and '-N-' (a context line) shows that file's content, not
# just its name. A date inside a path (haiku-2026-09-29/) is not a context line.
BENCH_CONTENT_LINE_RE = re.compile(
    r"^[ \t]*(?P<path>[^\s:]*xlsx-bench(?:[^\s:]*:|[^\s:]*?\.[A-Za-z0-9]+(?=-\d+-)))", re.MULTILINE
)


def _walk_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from _walk_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk_strings(v)


def _transcript_strings(entries: list[dict]):
    """Yield (kind, tool, text) for everything the model wrote or saw:
    tool_use inputs, tool_result text, assistant text. The opening user
    prompt and harness-injected entries are not scanned."""
    tool_names: dict[str, str] = {}
    for entry in entries:
        content = (entry.get("message") or {}).get("content")
        if not isinstance(content, list):
            continue
        for block in content:
            if not isinstance(block, dict):
                continue
            btype = block.get("type")
            if btype == "tool_use":
                tool_names[block.get("id")] = block.get("name")
                for s in _walk_strings(block.get("input")):
                    yield "tool_use", block.get("name"), s
            elif btype == "tool_result":
                yield "tool_result", tool_names.get(block.get("tool_use_id")), _tool_result_text(block)
            elif btype == "text" and entry.get("type") == "assistant":
                yield "assistant_text", None, block.get("text", "")


def _under(path: str, prefixes: list[str]) -> bool:
    return any(path == p or path.startswith(p.rstrip("/") + "/") for p in prefixes)


def detect_contamination(
    entries: list[dict],
    own_dirs: Optional[list[str]] = None,
    own_tmp: Optional[str] = None,
    max_evidence: int = 20,
) -> dict:
    """Scan a session transcript for any path under research/xlsx-bench/
    (recordings, runs, results, corpus, tasks*.json, grade.py) or another
    run's temp dir. Reading the SafeXL package, plugin/skill dirs or ~/.claude
    is fine and is not flagged.

    own_dirs: paths a session in the OLD in-repo layout may legitimately use
    (its own cell dir, which holds only its own work dir while it runs); anything else under xlsx-bench is flagged. Empty for
    the isolated layout, where no xlsx-bench path should ever appear.
    own_tmp: this run's temp dir (basename compared; every other
    /tmp/xlsxb-* is flagged).

    Only what the session asked for counts: a hit in a tool_use input is
    contamination. A hit that appears only in tool output or assistant text
    is a warning (e.g. arm B reading a SafeXL test file that names the bench
    corpus), unless the output is grep-style content from a bench file.
    Returns {"contaminated": bool, "evidence": [...]}, plus "warnings": [...]
    when there are any.
    """
    allowed = [str(Path(p)) for p in (own_dirs or [])]
    # The same dirs as the repo-relative paths tools print (research/xlsx-bench/...).
    for p in list(allowed):
        try:
            allowed.append(str(Path(p).resolve().relative_to(REPO_ROOT)))
        except ValueError:
            pass
    # Claude Code's own per-cwd state (~/.claude/projects/<escaped cwd>/...,
    # /tmp/claude-<uid>/<escaped cwd>/...) is derived from the cwd, so it names
    # this run's dir without being a sibling's.
    own_escaped = [re.sub(r"[/.]", "-", str(Path(p).resolve())) for p in (own_dirs or [])]
    own_tmp_name = Path(own_tmp).name if own_tmp else None
    evidence: list[dict] = []
    warnings: list[dict] = []
    seen: set[tuple] = set()

    def add(kind, tool, match, text, start, end, content=False):
        bucket = evidence if kind == "tool_use" or content else warnings
        key = (kind, tool, match, bucket is evidence)
        if key in seen or len(bucket) >= max_evidence:
            return
        seen.add(key)
        bucket.append({
            "kind": kind, "tool": tool, "match": match,
            "snippet": text[max(start - 60, 0):end + 60].replace("\n", " ")[:260],
        })

    for kind, tool, text in _transcript_strings(entries):
        content_starts = (
            {m.start("path") for m in BENCH_CONTENT_LINE_RE.finditer(text)} if kind == "tool_result" else set()
        )
        for m in BENCH_TREE_RE.finditer(text):
            path = m.group(0).rstrip(".:")
            if (
                any(ig in path for ig in CONTAMINATION_IGNORE)
                or _under(path, allowed)
                or any(tok in path for tok in own_escaped)
            ):
                continue
            add(kind, tool, path, text, m.start(), m.end(), content=m.start() in content_starts)
        for m in OTHER_TMP_RE.finditer(text):
            if m.group(0).rsplit("/", 1)[-1] != own_tmp_name:
                add(kind, tool, m.group(0), text, m.start(), m.end())
        if allowed and kind == "tool_use":
            for m in LEGACY_CLIMB_RE.finditer(text):
                add(kind, tool, m.group(0), text, m.start(), m.end())

    verdict = {"contaminated": bool(evidence), "evidence": evidence}
    if warnings:
        verdict["warnings"] = warnings
    return verdict


# ---------------------------------------------------------------------------
# Filesystem diffing (pure-ish; touches disk but takes no globals)
# ---------------------------------------------------------------------------


def snapshot_dir(path: Path) -> set[str]:
    if not path.exists():
        return set()
    return {str(p.relative_to(path)) for p in path.rglob("*") if p.is_file()}


def new_files(before: set[str], after: set[str]) -> list[str]:
    return sorted(after - before)


# ---------------------------------------------------------------------------
# Live execution (I/O + subprocess -- not covered by unit tests)
# ---------------------------------------------------------------------------


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _kill_process_group(proc: "subprocess.Popen") -> None:
    try:
        pgid = os.getpgid(proc.pid)
    except ProcessLookupError:
        return
    try:
        os.killpg(pgid, signal.SIGTERM)
    except ProcessLookupError:
        return
    time.sleep(2)
    try:
        os.killpg(pgid, signal.SIGKILL)
    except ProcessLookupError:
        pass


def run_cell(cell: Cell, cdir: Path, claude_bin: str, block_skill_method: str, timeout: int) -> dict:
    wdir = work_dir(cdir)
    if wdir.exists():
        shutil.rmtree(wdir)  # a fresh folder per run, even on a crashed-and-retried cell
    wdir.mkdir(parents=True, exist_ok=True)

    workbook_src = resolve_workbook(cell.task)
    if not workbook_src.exists():
        raise FileNotFoundError(f"workbook not found: {workbook_src}")
    workbook_dst = wdir / workbook_src.name
    shutil.copy2(workbook_src, workbook_dst)

    before = snapshot_dir(wdir)

    cmd = build_command(claude_bin, cell.task.prompt, cell.arm, block_skill_method)
    env = build_env(os.environ, cell.arm)

    events_path = cdir / "events.jsonl"
    stderr_path = cdir / "stderr.txt"

    started_iso = _now_iso()
    started = time.monotonic()
    exit_code: Optional[int] = None
    timed_out = False

    with open(events_path, "wb") as ev_f, open(stderr_path, "wb") as err_f:
        proc = subprocess.Popen(
            cmd, cwd=str(wdir), env=env, stdout=ev_f, stderr=err_f, preexec_fn=os.setsid
        )
        try:
            exit_code = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            _kill_process_group(proc)
            try:
                exit_code = proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                exit_code = None

    wall_seconds = time.monotonic() - started
    finished_iso = _now_iso()

    events_text = events_path.read_text(errors="replace")
    stderr_text = stderr_path.read_text(errors="replace")
    events = list(iter_json_lines(events_text))

    summary = summarize_events(events)
    after = snapshot_dir(wdir)
    output_files = new_files(before, after)
    rate_limit_evidence = detect_rate_limit(events, stderr_text)

    result = {
        "task_id": cell.task.id,
        "arm": cell.arm,
        "rep": cell.rep,
        "workbook": cell.task.workbook,
        "prompt": cell.task.prompt,
        "started_at": started_iso,
        "finished_at": finished_iso,
        "wall_seconds": round(wall_seconds, 2),
        "exit_code": exit_code,
        "timed_out": timed_out,
        "output_files": output_files,
        "rate_limited": rate_limit_evidence is not None,
        "rate_limit_evidence": rate_limit_evidence,
        "events_file": "events.jsonl",
        "stderr_file": "stderr.txt",
        "stderr_bytes": len(stderr_text.encode("utf-8", errors="replace")),
        "invocation": {
            "command": cmd,
            "cwd": str(wdir),
            "block_skill_method": block_skill_method if cell.arm == "B" else None,
            "plugin_dir": SAFEXL_PLUGIN_DIR if cell.arm == "B" else None,
            "path_prepend": [SAFEXL_VENV_BIN] if cell.arm == "B" else [],
            "anthropic_api_key_set": False,
            "claude_md_disabled": True,
        },
        "runner_version": RUNNER_VERSION,
    }
    result.update(summary)
    return result


def write_run_json(cdir: Path, result: dict) -> None:
    """Write run.json atomically (temp file + rename) so a crash mid-write
    can never leave a partial file that is_done() would mistake for done."""
    tmp = cdir / "run.json.tmp"
    tmp.write_text(json.dumps(result, indent=2, sort_keys=False) + "\n")
    tmp.replace(run_json_path(cdir))


# ---------------------------------------------------------------------------
# tmux relaunch
# ---------------------------------------------------------------------------


def relaunch_in_tmux(argv: list[str]) -> int:
    if shutil.which("tmux") is None:
        print("error: tmux is not installed", file=sys.stderr)
        return 2

    check = subprocess.run(
        ["tmux", "has-session", "-t", TMUX_SESSION_NAME], capture_output=True
    )
    if check.returncode == 0:
        print(f"tmux session '{TMUX_SESSION_NAME}' is already running.")
        print(f"Attach with: tmux attach -t {TMUX_SESSION_NAME}")
        return 0

    filtered = [a for a in argv if a != "--tmux"]
    # Start in the caller's cwd so relative --tasks/--out paths resolve the same
    # way they would without --tmux, and keep a log: a crash would otherwise
    # close the session and take the error with it.
    log = SCRIPT_DIR / "runner-tmux.log"
    inner = shlex.join([sys.executable, str(Path(__file__).resolve()), *filtered])
    cmd = [
        "tmux", "new-session", "-d", "-s", TMUX_SESSION_NAME, "-c", str(Path.cwd()),
        f"{inner} 2>&1 | tee -a {shlex.quote(str(log))}",
    ]
    subprocess.run(cmd, check=True)
    print(f"Started tmux session '{TMUX_SESSION_NAME}'. Log: {log}")
    print(f"Attach with: tmux attach -t {TMUX_SESSION_NAME}")
    return 0


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="runner.py",
        description="Two-arm benchmark: Anthropic xlsx skill (A) vs SafeXL (B).",
    )
    parser.add_argument("--tasks", type=Path, default=SCRIPT_DIR / "smoke_tasks.json")
    parser.add_argument("--arms", type=str, default="A,B")
    parser.add_argument("--reps", type=int, default=1)
    parser.add_argument("--out", type=Path, default=SCRIPT_DIR / "runs")
    parser.add_argument(
        "--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS,
        help=f"per-run wall-clock timeout in seconds (default {DEFAULT_TIMEOUT_SECONDS} = 20 min)",
    )
    parser.add_argument("--claude-bin", type=str, default=CLAUDE_BIN_DEFAULT)
    parser.add_argument(
        "--block-skill-method", choices=("disallowed", "settings"), default="disallowed",
        help="how arm B blocks the Anthropic xlsx skill (see README.md)",
    )
    parser.add_argument("--tmux", action="store_true", help="relaunch in a detached tmux session")
    parser.add_argument("--status", action="store_true", help="print done/total per arm and exit")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    raw_argv = list(argv) if argv is not None else sys.argv[1:]
    parser = build_arg_parser()
    args = parser.parse_args(raw_argv)

    if args.tmux:
        return relaunch_in_tmux(raw_argv)

    try:
        tasks = load_tasks(args.tasks)
        arms = parse_arms(args.arms)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2

    if args.status:
        print(status_report(tasks, arms, args.reps, args.out))
        return 0

    if not Path(args.claude_bin).exists():
        print(f"error: claude binary not found at {args.claude_bin}", file=sys.stderr)
        return 2

    out_dir = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    stopped_path = out_dir / "STOPPED.json"
    if stopped_path.exists():
        print(
            f"error: {stopped_path} exists from a previously stopped batch; "
            "review it and remove it before running again",
            file=sys.stderr,
        )
        return 2

    try:
        cells = build_matrix(tasks, arms, args.reps)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2

    total = len(cells)
    done_count = 0
    for i, cell in enumerate(cells, 1):
        cdir = run_dir(out_dir, cell.task.id, cell.arm, cell.rep)
        label = f"{cell.task.id}/{cell.arm}/rep{cell.rep}"
        if is_done(cdir):
            done_count += 1
            print(f"[{i}/{total}] {label}: already done, skipping")
            continue

        cdir.mkdir(parents=True, exist_ok=True)
        print(f"[{i}/{total}] {label}: running (timeout {args.timeout}s)...")
        try:
            result = run_cell(cell, cdir, args.claude_bin, args.block_skill_method, args.timeout)
        except FileNotFoundError as e:
            print(f"error: {e}", file=sys.stderr)
            return 2

        write_run_json(cdir, result)
        cost = result.get("total_cost_usd")
        cost_s = f"${cost:.4f}" if isinstance(cost, (int, float)) else "?"
        print(
            f"    exit={result.get('exit_code')} cost={cost_s} "
            f"wall={result.get('wall_seconds')}s model={result.get('model')}"
        )

        if result.get("rate_limited"):
            stopped = {
                "stopped_at": _now_iso(),
                "reason": result.get("rate_limit_evidence") or "usage/rate limit detected in output",
                "cell": {"task": cell.task.id, "arm": cell.arm, "rep": cell.rep},
                "run_dir": str(cdir),
            }
            stopped_path.write_text(json.dumps(stopped, indent=2) + "\n")
            print(f"RATE LIMIT DETECTED -- stopping batch. See {stopped_path}", file=sys.stderr)
            return 1

        done_count += 1

    print(f"Done: {done_count}/{total}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
