"""Unit tests for research/xlsx-bench/record.py.

Covers the pure logic only: command construction (interactive_command /
build_asciinema_argv / build_tmux_argv), dialog classification, transcript
parsing (transcript_is_finished / summarize_transcript), hashing/manifest,
and the CLI's non-live paths (--status, resumability skip, argument
validation). No live `claude`/`tmux`/`asciinema` invocation anywhere in this
file -- that's the job of the two smoke recordings, not pytest.

record.py (and the runner.py it imports) is loaded directly from disk by
path, same pattern as tests/test_runner.py.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
from pathlib import Path

import pytest

_RUNNER_PATH = Path(__file__).resolve().parent.parent / "runner.py"
_spec_runner = importlib.util.spec_from_file_location("xlsx_bench_runner_for_record", _RUNNER_PATH)
_runner_mod = importlib.util.module_from_spec(_spec_runner)
assert _spec_runner.loader is not None
sys.modules["xlsx_bench_runner_for_record"] = _runner_mod
_spec_runner.loader.exec_module(_runner_mod)

_RECORD_PATH = Path(__file__).resolve().parent.parent / "record.py"
_spec_record = importlib.util.spec_from_file_location("xlsx_bench_record", _RECORD_PATH)
rec = importlib.util.module_from_spec(_spec_record)
assert _spec_record.loader is not None
sys.modules["xlsx_bench_record"] = rec
_spec_record.loader.exec_module(rec)

runner = rec.runner  # the same runner module record.py imported for itself


# ---------------------------------------------------------------------------
# escape_project_dir
# ---------------------------------------------------------------------------


def test_escape_project_dir_replaces_slashes():
    p = Path("/home/alice/project/research/xlsx-bench/runs/pilot/t1/A/1/work")
    assert rec.escape_project_dir(p) == (
        "-home-alice-project-research-xlsx-bench-runs-pilot-t1-A-1-work"
    )


def test_escape_project_dir_replaces_dots_too():
    # a ".claude" path segment becomes a double dash (the "/" and the "."
    # each become their own "-") -- verified against a real worktree's
    # ~/.claude/projects/ directory name.
    p = Path("/home/alice/project/.claude/worktrees/agent-xyz/work")
    escaped = rec.escape_project_dir(p)
    assert "--claude-worktrees-agent-xyz-work" in escaped
    assert escaped.startswith("-home-alice-project")


# ---------------------------------------------------------------------------
# classify_dialog / truncate_pane_text
# ---------------------------------------------------------------------------

_IDLE_PANE = """ ▐▛███▛█   Claude Code v2.1.283
▝▜██████▀  Opus 5.5 (1M context) with xhigh effort · Claude Pro
 ▝▝   ▝▝   ~/…/.scratch/autosubmit-test-1


❯ Reply with exactly the single word PONG and nothing else, then
  stop.

● PONG

✻ Sauteed for 2s · done 2:15 AM

───────────────────────────────────────────────────────────────────
❯
───────────────────────────────────────────────────────────────────
  ⏵⏵ don't ask on (shift+tab to cycle) · ← 2 agents
                                                ◉ xhigh · /effort
"""

_FOLDER_TRUST_PANE = """
───────────────────────────────────────────────────────────────────
 Accessing workspace:

 /tmp/claude-1000/never-trusted-probe-dir

 Quick safety check: Is this a project you created or one you
 trust? (Like your own code, a well-known open source project, or
 work from your team). If not, take a moment to review what's in
 this folder first.

 Claude Code'll be able to read, edit, and execute files here.

 Security guide

 ❯ No, exit
   Yes, I trust this folder

 Enter to confirm · Esc to cancel
"""

_MCP_DIALOG_PANE = """
───────────────────────────────────────────────────────────────────
  7 new MCP servers found in this project
  Select any you wish to enable.

  MCP servers may execute code or access system resources. All
  tool calls require approval. Learn more in the MCP
  documentation.

  ❯ [✔] gmail
    [✔] github
    [✔] sentry
    [✔] socket
  ↓ [✔] hugging-face
       Enable selected
 Space to select · Esc to reject all
"""


def test_classify_dialog_none_for_idle_ui():
    assert rec.classify_dialog(_IDLE_PANE) is None


def test_classify_dialog_none_for_empty():
    assert rec.classify_dialog("") is None
    assert rec.classify_dialog(None) is None


def test_classify_dialog_folder_trust():
    result = rec.classify_dialog(_FOLDER_TRUST_PANE)
    assert result == {"type": "folder_trust"}


def test_classify_dialog_other_for_mcp_servers_dialog():
    result = rec.classify_dialog(_MCP_DIALOG_PANE)
    assert result == {"type": "other"}


def test_truncate_pane_text_short_unchanged():
    assert rec.truncate_pane_text("hello") == "hello"


def test_truncate_pane_text_long_truncates():
    long_text = "x" * 3000
    out = rec.truncate_pane_text(long_text, n=100)
    assert len(out) <= 101
    assert out.endswith("…")


# ---------------------------------------------------------------------------
# interactive_command
# ---------------------------------------------------------------------------


def test_interactive_command_arm_a_shape():
    cmd = rec.interactive_command("claude", "add a Reviewed column", "A", "disallowed")
    assert cmd[0] == "claude"
    assert cmd[-1] == "add a Reviewed column"  # prompt trails, positional
    assert "-p" not in cmd
    assert "--output-format" not in cmd
    assert "--verbose" not in cmd
    assert "--allowedTools" in cmd
    idx = cmd.index("--allowedTools")
    assert cmd[idx + 1] == runner.ALLOWED_TOOLS
    assert "--permission-mode" in cmd
    pm_idx = cmd.index("--permission-mode")
    assert cmd[pm_idx + 1] == "dontAsk"
    assert "--strict-mcp-config" in cmd
    assert "--plugin-dir" not in cmd
    assert "--disallowedTools" not in cmd


def test_interactive_command_arm_b_shape():
    cmd = rec.interactive_command("claude", "add a Reviewed column", "B", "disallowed")
    assert "--plugin-dir" in cmd
    plugin_idx = cmd.index("--plugin-dir")
    assert cmd[plugin_idx + 1] == runner.SAFEXL_PLUGIN_DIR
    assert "--disallowedTools" in cmd
    deny_idx = cmd.index("--disallowedTools")
    assert cmd[deny_idx + 1] == f"Skill({runner.ANTHROPIC_XLSX_SKILL})"
    assert cmd[-1] == "add a Reviewed column"
    assert "--permission-mode" in cmd and "dontAsk" in cmd
    assert "--strict-mcp-config" in cmd


def test_interactive_command_arm_b_settings_method():
    cmd = rec.interactive_command("claude", "p", "B", "settings")
    assert "--settings" in cmd
    assert "--disallowedTools" not in cmd
    assert cmd[-1] == "p"


def test_interactive_command_never_leaks_prompt_into_flags():
    prompt = "Reply with exactly the single word PONG"
    cmd = rec.interactive_command("claude", prompt, "A", "disallowed")
    assert cmd.count(prompt) == 1
    assert cmd[-1] == prompt


# ---------------------------------------------------------------------------
# build_asciinema_argv / build_tmux_argv
# ---------------------------------------------------------------------------


def test_build_asciinema_argv_shape(tmp_path):
    cast = tmp_path / "session.cast"
    claude_argv = ["claude", "--allowedTools", "Skill Read", "hello world"]
    argv = rec.build_asciinema_argv(cast, claude_argv, cols=140, rows=42, asciinema_bin="asciinema")
    assert argv[0] == "asciinema"
    assert argv[1] == "rec"
    assert "--overwrite" in argv
    assert "-q" in argv
    cols_idx = argv.index("--cols")
    assert argv[cols_idx + 1] == "140"
    rows_idx = argv.index("--rows")
    assert argv[rows_idx + 1] == "42"
    assert argv[-1] == str(cast)
    c_idx = argv.index("-c")
    assert argv[c_idx + 1] != claude_argv  # joined into one string, not the list


def test_build_asciinema_argv_roundtrips_tricky_prompt(tmp_path):
    import shlex

    cast = tmp_path / "x.cast"
    tricky_prompt = "It's a test — with $DOLLAR `backtick` and \"quotes\"."
    claude_argv = ["claude", "--allowedTools", "Skill Read Write Edit Bash Glob Grep", tricky_prompt]
    argv = rec.build_asciinema_argv(cast, claude_argv)
    c_idx = argv.index("-c")
    joined = argv[c_idx + 1]
    assert shlex.split(joined) == claude_argv


def test_build_tmux_argv_shape(tmp_path):
    command_argv = ["asciinema", "rec", "-c", "claude foo", "cast.cast"]
    argv = rec.build_tmux_argv("xlsx-rec-abcd", tmp_path, command_argv, cols=140, rows=42)
    assert argv[:5] == ["tmux", "new-session", "-d", "-s", "xlsx-rec-abcd"]
    assert "-x" in argv and argv[argv.index("-x") + 1] == "140"
    assert "-y" in argv and argv[argv.index("-y") + 1] == "42"
    assert "-c" in argv and argv[argv.index("-c") + 1] == str(tmp_path)
    # command_argv appended as separate argv elements, not shell-joined
    assert argv[-len(command_argv):] == command_argv


# ---------------------------------------------------------------------------
# transcript_is_finished
# ---------------------------------------------------------------------------


def _assistant(stop_reason, content_types, request_id="req1"):
    content = []
    for i, t in enumerate(content_types):
        if t == "text":
            content.append({"type": "text", "text": f"text-{i}"})
        elif t == "tool_use":
            content.append({"type": "tool_use", "id": f"tu-{i}", "name": "Bash", "input": {}})
        elif t == "thinking":
            content.append({"type": "thinking", "thinking": "..."})
    return {
        "type": "assistant",
        "requestId": request_id,
        "message": {"stop_reason": stop_reason, "content": content},
    }


def test_transcript_is_finished_empty():
    assert rec.transcript_is_finished([]) is False


def test_transcript_is_finished_false_when_tool_use_pending():
    entries = [_assistant("tool_use", ["tool_use"])]
    assert rec.transcript_is_finished(entries) is False


def test_transcript_is_finished_true_on_end_turn_text_only():
    entries = [_assistant("tool_use", ["tool_use"]), _assistant("end_turn", ["text"])]
    assert rec.transcript_is_finished(entries) is True


def test_transcript_is_finished_false_if_end_turn_entry_still_has_tool_use_block():
    # defensive: shouldn't happen in practice, but a stray tool_use block
    # alongside end_turn must not be treated as finished
    entries = [_assistant("end_turn", ["text", "tool_use"])]
    assert rec.transcript_is_finished(entries) is False


def test_transcript_is_finished_uses_last_assistant_not_first():
    entries = [
        _assistant("end_turn", ["text"], request_id="req1"),
        _assistant("tool_use", ["tool_use"], request_id="req2"),
    ]
    assert rec.transcript_is_finished(entries) is False


def test_transcript_is_finished_ignores_non_assistant_lines():
    entries = [
        {"type": "queue-operation", "operation": "enqueue"},
        {"type": "attachment", "attachment": {}},
        _assistant("end_turn", ["text"]),
        {"type": "cost-state", "totalCostUSD": 0.01},
    ]
    assert rec.transcript_is_finished(entries) is True


# ---------------------------------------------------------------------------
# summarize_transcript
# ---------------------------------------------------------------------------


def _assistant_tool_use(request_id, tool_use_id, name, tool_input, usage=None, model="claude-opus-5-5"):
    return {
        "type": "assistant",
        "requestId": request_id,
        "message": {
            "model": model,
            "stop_reason": "tool_use",
            "content": [{"type": "tool_use", "id": tool_use_id, "name": name, "input": tool_input}],
            "usage": usage or {},
        },
    }


def _user_tool_result(tool_use_id, text, is_error=False):
    return {
        "type": "user",
        "message": {
            "content": [
                {"type": "tool_result", "tool_use_id": tool_use_id, "content": text, "is_error": is_error}
            ]
        },
    }


def test_summarize_transcript_model_and_final_text():
    entries = [
        _assistant_tool_use("r1", "tu1", "Read", {"file_path": "x"}),
        {
            "type": "assistant",
            "requestId": "r2",
            "message": {
                "model": "claude-opus-5-5",
                "stop_reason": "end_turn",
                "content": [{"type": "text", "text": "all done"}],
                "usage": {"input_tokens": 1, "output_tokens": 2},
            },
        },
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["model"] == "claude-opus-5-5"
    assert summary["final_text"] == "all done"


def test_summarize_transcript_dedupes_usage_by_request_id():
    # same requestId logged twice (one per content block), as real Claude
    # Code transcripts do -- must sum once, not twice.
    usage = {
        "input_tokens": 2,
        "output_tokens": 373,
        "cache_read_input_tokens": 46737,
        "cache_creation_input_tokens": 878,
    }
    entries = [
        {
            "type": "assistant",
            "requestId": "req-A",
            "message": {"model": "m", "stop_reason": "end_turn",
                        "content": [{"type": "thinking", "thinking": "..."}], "usage": usage},
        },
        {
            "type": "assistant",
            "requestId": "req-A",
            "message": {"model": "m", "stop_reason": "end_turn",
                        "content": [{"type": "text", "text": "done"}], "usage": usage},
        },
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["num_turns"] == 1
    assert summary["tokens"]["input_tokens"] == 2
    assert summary["tokens"]["output_tokens"] == 373
    assert summary["tokens"]["cache_read_input_tokens"] == 46737
    assert summary["tokens"]["cache_creation_input_tokens"] == 878


def test_summarize_transcript_sums_across_distinct_requests():
    u = {"input_tokens": 10, "output_tokens": 20, "cache_read_input_tokens": 0, "cache_creation_input_tokens": 0}
    entries = [
        {"type": "assistant", "requestId": "r1", "message": {"stop_reason": "tool_use", "content": [], "usage": u}},
        {"type": "assistant", "requestId": "r2", "message": {"stop_reason": "end_turn", "content": [], "usage": u}},
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["num_turns"] == 2
    assert summary["tokens"]["input_tokens"] == 20
    assert summary["tokens"]["output_tokens"] == 40


def test_summarize_transcript_skill_invoked_and_succeeded():
    entries = [
        _assistant_tool_use("r1", "tu1", "Skill", {"skill": "safexl"}),
        _user_tool_result("tu1", "ran safexl edit ok", is_error=False),
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["skills_invoked"] == ["safexl"]
    assert summary["skills_succeeded"] == ["safexl"]
    assert summary["skills_denied"] == []


def test_summarize_transcript_skill_denied():
    entries = [
        _assistant_tool_use("r1", "tu1", "Skill", {"skill": "anthropic-skills:xlsx"}),
        _user_tool_result("tu1", "Permission denied: this skill is disallowed for this session", is_error=True),
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["skills_denied"] == ["anthropic-skills:xlsx"]
    assert summary["skills_succeeded"] == []
    assert "anthropic-skills:xlsx" in summary["skill_denial_evidence"]


def test_summarize_transcript_safexl_bash_detected():
    entries = [
        _assistant_tool_use("r1", "tu1", "Bash", {"command": "safexl inspect work/SimpleMacro.xlsm --json"}),
        _user_tool_result("tu1", "{...}"),
        _assistant_tool_use("r2", "tu2", "Read", {"file_path": "work/SimpleMacro.xlsm"}),
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["safexl_used"] is True
    assert summary["safexl_commands"] == ["safexl inspect work/SimpleMacro.xlsm --json"]
    assert set(summary["tools_called"]) == {"Bash", "Read"}


def test_summarize_transcript_no_safexl():
    entries = [_assistant_tool_use("r1", "tu1", "Bash", {"command": "ls -la"})]
    summary = rec.summarize_transcript(entries)
    assert summary["safexl_used"] is False


def test_summarize_transcript_total_cost_from_cost_state():
    entries = [
        _assistant_tool_use("r1", "tu1", "Read", {}),
        {"type": "cost-state", "totalCostUSD": 0.4123, "sessionId": "s"},
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["total_cost_usd"] == 0.4123


def test_summarize_transcript_ignores_malformed_blocks():
    entries = [
        {"type": "assistant", "requestId": "r1", "message": {"content": ["not-a-dict", {"type": "text"}]}},
        {"type": "user", "message": {"content": [{"type": "tool_result"}]}},
    ]
    summary = rec.summarize_transcript(entries)
    assert summary["tools_called"] == []


# ---------------------------------------------------------------------------
# read_transcript_entries
# ---------------------------------------------------------------------------


def test_read_transcript_entries_tolerant(tmp_path):
    p = tmp_path / "t.jsonl"
    p.write_text('{"a":1}\n\nnot json\n{"b":2}\n')
    assert rec.read_transcript_entries(p) == [{"a": 1}, {"b": 2}]


def test_read_transcript_entries_missing_file_returns_empty(tmp_path):
    assert rec.read_transcript_entries(tmp_path / "nope.jsonl") == []


# ---------------------------------------------------------------------------
# sha256_file / build_manifest
# ---------------------------------------------------------------------------


def test_sha256_file_matches_hashlib(tmp_path):
    p = tmp_path / "f.bin"
    p.write_bytes(b"hello world" * 1000)
    expected = hashlib.sha256(p.read_bytes()).hexdigest()
    assert rec.sha256_file(p) == expected


def test_build_manifest_structure_save_as_unchanged_original(tmp_path):
    # a "save as new filename" task: the original is never touched, so
    # before == after and modified_in_place is False.
    wdir = tmp_path / "work"
    wdir.mkdir()
    (wdir / "SimpleMacro.xlsm").write_bytes(b"orig")
    (wdir / "SimpleMacro-updated.xlsm").write_bytes(b"new")
    cast = tmp_path / "session.cast"
    cast.write_text('{"version":2}\n')
    jsonl = tmp_path / "session.jsonl"
    jsonl.write_text('{"type":"assistant"}\n')
    before_hash = rec.sha256_file(wdir / "SimpleMacro.xlsm")

    manifest = rec.build_manifest(
        "SimpleMacro.xlsm",
        before_hash,
        wdir,
        ["SimpleMacro-updated.xlsm"],
        cast,
        jsonl,
        before_hash,  # workbook_sha256_after -- untouched, so same as before
    )
    assert manifest["workbook"]["path"] == "SimpleMacro.xlsm"
    assert manifest["workbook"]["sha256_before"] == before_hash
    assert manifest["workbook"]["sha256_after"] == before_hash
    assert manifest["workbook"]["modified_in_place"] is False
    assert manifest["output_files"] == [
        {"path": "SimpleMacro-updated.xlsm", "sha256": rec.sha256_file(wdir / "SimpleMacro-updated.xlsm")}
    ]
    assert manifest["session_cast"] == {"path": "session.cast", "sha256": rec.sha256_file(cast)}
    assert manifest["session_jsonl"] == {"path": "session.jsonl", "sha256": rec.sha256_file(jsonl)}


def test_build_manifest_structure_modified_in_place(tmp_path):
    wdir = tmp_path / "work"
    wdir.mkdir()
    (wdir / "budget.xlsx").write_bytes(b"edited-bytes")
    cast = tmp_path / "session.cast"
    jsonl = tmp_path / "session.jsonl"
    before_hash = "deadbeef" * 8
    after_hash = rec.sha256_file(wdir / "budget.xlsx")

    manifest = rec.build_manifest("budget.xlsx", before_hash, wdir, [], cast, None, after_hash)
    assert manifest["workbook"]["sha256_before"] == before_hash
    assert manifest["workbook"]["sha256_after"] == after_hash
    assert manifest["workbook"]["modified_in_place"] is True


def test_build_manifest_workbook_after_defaults_to_none(tmp_path):
    wdir = tmp_path / "work"
    wdir.mkdir()
    manifest = rec.build_manifest("gone.xlsx", "abc123", wdir, [], tmp_path / "session.cast", None)
    assert manifest["workbook"]["sha256_after"] is None
    # unknown, not False -- we genuinely don't know without an after-hash
    assert manifest["workbook"]["modified_in_place"] is None


def test_build_manifest_modified_in_place_unknown_without_before_hash(tmp_path):
    wdir = tmp_path / "work"
    wdir.mkdir()
    (wdir / "x.xlsx").write_bytes(b"data")
    after = rec.sha256_file(wdir / "x.xlsx")
    manifest = rec.build_manifest("x.xlsx", None, wdir, [], tmp_path / "session.cast", None, after)
    assert manifest["workbook"]["modified_in_place"] is None


def test_build_manifest_missing_output_file_skipped(tmp_path):
    wdir = tmp_path / "work"
    wdir.mkdir()
    (wdir / "orig.xlsx").write_bytes(b"x")
    cast = tmp_path / "session.cast"
    manifest = rec.build_manifest("orig.xlsx", rec.sha256_file(wdir / "orig.xlsx"), wdir, ["missing.xlsx"], cast, None)
    assert manifest["output_files"] == []
    assert manifest["session_cast"] is None
    assert manifest["session_jsonl"] is None


# ---------------------------------------------------------------------------
# _tmux_session_alive / _tmux_disable_remain_on_exit
#
# Regression coverage for a real bug: an earlier tmux server on this box had
# `remain-on-exit` set globally ON, so a pane whose process had already died
# still passed `tmux has-session` (the session object lingers showing its
# last frame). A real recording then polled a corpse for the full 20-minute
# timeout instead of getting caught by record_cell's spurious-early-death
# retry. `_tmux_session_alive` must treat a dead pane as "not alive" even
# when has-session alone would say otherwise.
# ---------------------------------------------------------------------------


def test_tmux_session_alive_false_when_no_session(monkeypatch):
    def fake_run(argv, **k):
        class R:
            returncode = 1
            stdout = ""
        return R()
    monkeypatch.setattr(rec.subprocess, "run", fake_run)
    assert rec._tmux_session_alive("nope") is False


def test_tmux_session_alive_false_when_pane_dead_despite_session_existing(monkeypatch):
    calls = []

    def fake_run(argv, **k):
        calls.append(argv)
        class R:
            pass
        r = R()
        if argv[:2] == ["tmux", "has-session"]:
            r.returncode = 0
            r.stdout = ""
        elif argv[:2] == ["tmux", "list-panes"]:
            r.returncode = 0
            r.stdout = "1\n"  # pane_dead == 1 -- the remain-on-exit-lingering case
        return r

    monkeypatch.setattr(rec.subprocess, "run", fake_run)
    assert rec._tmux_session_alive("lingering") is False


def test_tmux_session_alive_true_when_session_and_pane_both_alive(monkeypatch):
    def fake_run(argv, **k):
        class R:
            pass
        r = R()
        if argv[:2] == ["tmux", "has-session"]:
            r.returncode = 0
            r.stdout = ""
        elif argv[:2] == ["tmux", "list-panes"]:
            r.returncode = 0
            r.stdout = "0\n"
        return r

    monkeypatch.setattr(rec.subprocess, "run", fake_run)
    assert rec._tmux_session_alive("live") is True


def test_tmux_session_alive_false_when_list_panes_fails(monkeypatch):
    def fake_run(argv, **k):
        class R:
            pass
        r = R()
        if argv[:2] == ["tmux", "has-session"]:
            r.returncode = 0
            r.stdout = ""
        elif argv[:2] == ["tmux", "list-panes"]:
            r.returncode = 1
            r.stdout = ""
        return r

    monkeypatch.setattr(rec.subprocess, "run", fake_run)
    assert rec._tmux_session_alive("gone-between-checks") is False


def test_tmux_disable_remain_on_exit_calls_set_option(monkeypatch):
    calls = []
    monkeypatch.setattr(rec.subprocess, "run", lambda argv, **k: calls.append(argv))
    rec._tmux_disable_remain_on_exit("my-session")
    assert calls == [["tmux", "set-option", "-t", "my-session", "remain-on-exit", "off"]]


# ---------------------------------------------------------------------------
# verify_dontask_supported
# ---------------------------------------------------------------------------


def test_verify_dontask_supported_true(monkeypatch):
    class FakeResult:
        stdout = "some help text with dontAsk in it"

    monkeypatch.setattr(rec.subprocess, "run", lambda *a, **k: FakeResult())
    assert rec.verify_dontask_supported("claude") is True


def test_verify_dontask_supported_false(monkeypatch):
    class FakeResult:
        stdout = "some help text without the mode"

    monkeypatch.setattr(rec.subprocess, "run", lambda *a, **k: FakeResult())
    assert rec.verify_dontask_supported("claude") is False


def test_verify_dontask_supported_handles_oserror(monkeypatch):
    def boom(*a, **k):
        raise OSError("no such binary")

    monkeypatch.setattr(rec.subprocess, "run", boom)
    assert rec.verify_dontask_supported("nope") is False


# ---------------------------------------------------------------------------
# record_cell retry-on-spurious-early-death (no real tmux/asciinema/claude --
# _launch_and_wait is monkeypatched out entirely)
# ---------------------------------------------------------------------------


def _fake_task(tmp_path):
    workbook = tmp_path / "src.xlsx"
    workbook.write_bytes(b"fake-workbook-bytes")
    return runner.Task(id="t1", workbook=str(workbook), prompt="do something")


def _good_attempt_result(transcript_path):
    return {
        "started_iso": "2026-01-01T00:00:00Z",
        "started_monotonic": 0.0,
        "wall_seconds": 5.0,
        "finished_iso": "2026-01-01T00:00:05Z",
        "dialogs_seen": [],
        "transcript_path": transcript_path,
        "aborted_reason": None,
        "timed_out": False,
        "session_ended_early": False,
    }


def _spurious_death_result():
    return {
        "started_iso": "2026-01-01T00:00:00Z",
        "started_monotonic": 0.0,
        "wall_seconds": 2.5,
        "finished_iso": "2026-01-01T00:00:02Z",
        "dialogs_seen": [],
        "transcript_path": None,
        "aborted_reason": None,
        "timed_out": False,
        "session_ended_early": True,
    }


def test_record_cell_retries_once_on_spurious_early_death(tmp_path, monkeypatch):
    task = _fake_task(tmp_path)
    cdir = tmp_path / "cell"
    cdir.mkdir()
    transcript = tmp_path / "session.jsonl"
    transcript.write_text(
        json.dumps({"type": "assistant", "message": {"model": "m", "stop_reason": "end_turn",
                    "content": [{"type": "text", "text": "done"}]}}) + "\n"
    )

    session_names_used = []

    def fake_launch_and_wait(claude_argv, wdir, arm, session_name, timeout, poll_interval, quiet_seconds,
                              asciinema_bin, cast_path):
        session_names_used.append(session_name)
        if len(session_names_used) == 1:
            return _spurious_death_result()
        return _good_attempt_result(transcript)

    monkeypatch.setattr(rec, "_launch_and_wait", fake_launch_and_wait)

    result = rec.record_cell(task, "A", 1, cdir, "claude", "disallowed", 1200, "xlsx-rec-test",
                              relaunch_backoff_seconds=0)
    assert len(session_names_used) == 2
    # a distinct session name per attempt -- see record_cell's comment on why
    assert session_names_used[0] != session_names_used[1]
    assert session_names_used[0] == "xlsx-rec-test"
    assert result["tmux_session"] == session_names_used[1]
    assert result["relaunch_events"] and result["relaunch_events"][0]["attempt"] == 1
    assert result["session_ended_early"] is False
    assert result["final_text"] == "done"


def test_record_cell_does_not_retry_on_timeout(tmp_path, monkeypatch):
    task = _fake_task(tmp_path)
    cdir = tmp_path / "cell"
    cdir.mkdir()
    calls = []

    def fake_launch_and_wait(*a, **k):
        calls.append(1)
        r = _spurious_death_result()
        r["session_ended_early"] = False
        r["timed_out"] = True
        return r

    monkeypatch.setattr(rec, "_launch_and_wait", fake_launch_and_wait)
    result = rec.record_cell(task, "A", 1, cdir, "claude", "disallowed", 1200, "xlsx-rec-test",
                              relaunch_backoff_seconds=0)
    assert len(calls) == 1  # never retried
    assert result["timed_out"] is True
    assert result["relaunch_events"] == []


def test_record_cell_does_not_retry_on_dialog_abort(tmp_path, monkeypatch):
    task = _fake_task(tmp_path)
    cdir = tmp_path / "cell"
    cdir.mkdir()
    calls = []

    def fake_launch_and_wait(*a, **k):
        calls.append(1)
        r = _spurious_death_result()
        r["session_ended_early"] = False
        r["aborted_reason"] = "unexpected interactive dialog (not folder-trust) -- see dialogs_seen"
        r["dialogs_seen"] = [{"type": "other", "accepted": False, "text": "...", "at": "t"}]
        return r

    monkeypatch.setattr(rec, "_launch_and_wait", fake_launch_and_wait)
    result = rec.record_cell(task, "A", 1, cdir, "claude", "disallowed", 1200, "xlsx-rec-test",
                              relaunch_backoff_seconds=0)
    assert len(calls) == 1  # never retried
    assert result["aborted_reason"] is not None
    assert result["relaunch_events"] == []


def test_record_cell_no_retry_when_early_death_has_a_transcript(tmp_path, monkeypatch):
    # session_ended_early but a transcript DID exist -- not the "zero
    # evidence" signature, so this must not be treated as spurious/retried.
    task = _fake_task(tmp_path)
    cdir = tmp_path / "cell"
    cdir.mkdir()
    transcript = tmp_path / "session.jsonl"
    transcript.write_text(json.dumps({"type": "assistant", "message": {"stop_reason": "tool_use", "content": []}}) + "\n")
    calls = []

    def fake_launch_and_wait(*a, **k):
        calls.append(1)
        r = _spurious_death_result()
        r["transcript_path"] = transcript
        return r

    monkeypatch.setattr(rec, "_launch_and_wait", fake_launch_and_wait)
    result = rec.record_cell(task, "A", 1, cdir, "claude", "disallowed", 1200, "xlsx-rec-test",
                              relaunch_backoff_seconds=0)
    assert len(calls) == 1
    assert result["relaunch_events"] == []
    assert result["session_ended_early"] is True


# ---------------------------------------------------------------------------
# CLI: --status, resumability skip, argument validation (no live runs)
# ---------------------------------------------------------------------------


def _write_tasks(tmp_path, ids=("t1", "t2")):
    data = [{"id": i, "workbook": "research/safexl-moat/fixtures/SimpleMacro.xlsm", "prompt": f"do {i}"} for i in ids]
    p = tmp_path / "tasks.json"
    p.write_text(json.dumps(data))
    return p


def test_main_status_on_empty_out_dir(tmp_path, capsys):
    tasks = _write_tasks(tmp_path)
    rc_code = rec.main(["--tasks", str(tasks), "--arms", "A,B", "--reps", "1",
                         "--out", str(tmp_path / "recordings"), "--status"])
    assert rc_code == 0
    out = capsys.readouterr().out
    assert "Arm A: 0/2 done" in out
    assert "Arm B: 0/2 done" in out


def test_main_rejects_bad_tasks_file(tmp_path, capsys):
    bad = tmp_path / "tasks.json"
    bad.write_text("[]")
    rc_code = rec.main(["--tasks", str(bad), "--status", "--out", str(tmp_path / "recordings")])
    assert rc_code == 2
    assert "at least one task" in capsys.readouterr().err


def test_main_single_mode_requires_task_and_arm(tmp_path, capsys, monkeypatch):
    tasks = _write_tasks(tmp_path)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)
    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(tmp_path / "recordings"),
        "--claude-bin", sys.executable,
    ])
    assert rc_code == 2
    assert "--task and --arm are required" in capsys.readouterr().err


def test_main_all_and_task_mutually_exclusive(tmp_path, capsys, monkeypatch):
    tasks = _write_tasks(tmp_path)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)
    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(tmp_path / "recordings"),
        "--claude-bin", sys.executable, "--all", "--task", "t1", "--arm", "A",
    ])
    assert rc_code == 2
    assert "cannot be combined" in capsys.readouterr().err


def test_main_unknown_task_id(tmp_path, capsys, monkeypatch):
    tasks = _write_tasks(tmp_path)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)
    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(tmp_path / "recordings"),
        "--claude-bin", sys.executable, "--task", "nope", "--arm", "A",
    ])
    assert rc_code == 2
    assert "unknown task id" in capsys.readouterr().err


def test_main_rejects_missing_claude_bin(tmp_path, capsys):
    tasks = _write_tasks(tmp_path)
    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(tmp_path / "recordings"),
        "--claude-bin", str(tmp_path / "no-such-claude-binary"), "--task", "t1", "--arm", "A",
    ])
    assert rc_code == 2
    assert "claude binary not found" in capsys.readouterr().err


def test_main_rejects_when_dontask_unsupported(tmp_path, capsys, monkeypatch):
    tasks = _write_tasks(tmp_path)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: False)
    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(tmp_path / "recordings"),
        "--claude-bin", sys.executable, "--task", "t1", "--arm", "A",
    ])
    assert rc_code == 2
    assert "dontAsk" in capsys.readouterr().err


def test_main_single_mode_skips_when_already_done(tmp_path, monkeypatch, capsys):
    tasks = _write_tasks(tmp_path)
    out_dir = tmp_path / "recordings"
    cdir = runner.run_dir(out_dir, "t1", "A", 1)
    cdir.mkdir(parents=True)
    runner.write_run_json(cdir, {"done": True})

    def boom(*a, **k):
        raise AssertionError("record_cell must not be called when the cell is already done")

    monkeypatch.setattr(rec, "record_cell", boom)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)

    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(out_dir),
        "--claude-bin", sys.executable, "--task", "t1", "--arm", "A",
    ])
    assert rc_code == 0
    assert "already done" in capsys.readouterr().out


def test_main_single_mode_exits_nonzero_on_session_ended_early(tmp_path, monkeypatch):
    # regression: session_ended_early with no aborted_reason/timed_out/
    # rate_limited must still be a failing exit code -- exhausting every
    # retry and never getting a transcript is not a success.
    tasks = _write_tasks(tmp_path)
    out_dir = tmp_path / "recordings"

    def fake_record_cell(*a, **k):
        return {
            "session_ended_early": True, "aborted_reason": None, "timed_out": False,
            "rate_limited": False, "dialogs_seen": [], "wall_seconds": 3.0, "model": None,
        }

    monkeypatch.setattr(rec, "record_cell", fake_record_cell)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)

    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(out_dir),
        "--claude-bin", sys.executable, "--task", "t1", "--arm", "A",
    ])
    assert rc_code == 1


def test_main_all_mode_skips_done_cells_and_stops_on_stopped_json(tmp_path, monkeypatch, capsys):
    tasks = _write_tasks(tmp_path, ids=("t1",))
    out_dir = tmp_path / "recordings"
    out_dir.mkdir(parents=True)
    (out_dir / "STOPPED.json").write_text("{}")
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)

    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(out_dir), "--all",
        "--claude-bin", sys.executable,
    ])
    assert rc_code == 2
    assert "STOPPED.json exists" in capsys.readouterr().err


def test_main_all_mode_skips_already_done_cells(tmp_path, monkeypatch, capsys):
    tasks = _write_tasks(tmp_path, ids=("t1",))
    out_dir = tmp_path / "recordings"
    for arm in ("A", "B"):
        cdir = runner.run_dir(out_dir, "t1", arm, 1)
        cdir.mkdir(parents=True)
        runner.write_run_json(cdir, {"done": True})

    def boom(*a, **k):
        raise AssertionError("record_cell must not be called -- both cells are already done")

    monkeypatch.setattr(rec, "record_cell", boom)
    monkeypatch.setattr(rec, "verify_dontask_supported", lambda claude_bin: True)

    rc_code = rec.main([
        "--tasks", str(tasks), "--out", str(out_dir), "--all", "--arms", "A,B",
        "--claude-bin", sys.executable,
    ])
    assert rc_code == 0
    out = capsys.readouterr().out
    assert "Done: 2/2" in out


def test_build_tmux_argv_passes_arm_env_with_dash_e():
    """A new session on a running tmux server ignores the client's env, so the
    arm's variables must be passed with -e."""
    env = rec.pane_env({"PATH": "/p", "CLAUDE_CODE_DISABLE_CLAUDE_MDS": "1",
                           "CLAUDE_ENV_FILE": "/e.sh", "HOME": "/h"})
    argv = rec.build_tmux_argv("s", Path("/w"), ["asciinema", "rec"], env=env)
    assert ["-e", "PATH=/p"] == argv[argv.index("-e"):argv.index("-e") + 2]
    assert "CLAUDE_CODE_DISABLE_CLAUDE_MDS=1" in argv and "CLAUDE_ENV_FILE=/e.sh" in argv
    assert not any(a.startswith("HOME=") for a in argv)
    assert argv[-2:] == ["asciinema", "rec"]
