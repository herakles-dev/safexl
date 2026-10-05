"""Unit tests for research/xlsx-bench/runner.py.

Covers the pure logic only: tasks-file parsing, the A/B command/env builders,
stream-json event parsing, rate-limit detection, and resumability (skip a
cell whose run.json already exists). No live `claude` invocation anywhere in
this file -- that's the job of the smoke run, not pytest.

runner.py is loaded directly from disk by path (same pattern as
tests/test_mrr.py at the repo root) so these tests run the exact file the
CLI runs, without needing an __init__.py or sys.path surgery.
"""

from __future__ import annotations

import importlib.util
import json
import os
import sys
from pathlib import Path

import pytest

_SCRIPT_PATH = Path(__file__).resolve().parent.parent / "runner.py"
_spec = importlib.util.spec_from_file_location("xlsx_bench_runner", _SCRIPT_PATH)
runner = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
sys.modules["xlsx_bench_runner"] = runner
_spec.loader.exec_module(runner)


# ---------------------------------------------------------------------------
# parse_tasks / load_tasks
# ---------------------------------------------------------------------------


def test_parse_tasks_valid():
    data = [
        {"id": "a", "workbook": "x.xlsx", "prompt": "do something"},
        {"id": "b", "workbook": "y.xlsm", "prompt": "do something else"},
    ]
    tasks = runner.parse_tasks(data)
    assert [t.id for t in tasks] == ["a", "b"]
    assert tasks[0].workbook == "x.xlsx"
    assert tasks[1].prompt == "do something else"


def test_parse_tasks_rejects_non_list():
    with pytest.raises(ValueError, match="JSON list"):
        runner.parse_tasks({"id": "a"})


def test_parse_tasks_rejects_empty_list():
    with pytest.raises(ValueError, match="at least one task"):
        runner.parse_tasks([])


def test_parse_tasks_rejects_missing_fields():
    with pytest.raises(ValueError, match="missing required field"):
        runner.parse_tasks([{"id": "a", "workbook": "x.xlsx"}])


def test_parse_tasks_rejects_bad_id_chars():
    with pytest.raises(ValueError, match="directory name"):
        runner.parse_tasks([{"id": "a/b", "workbook": "x.xlsx", "prompt": "p"}])


def test_parse_tasks_rejects_duplicate_ids():
    data = [
        {"id": "a", "workbook": "x.xlsx", "prompt": "p1"},
        {"id": "a", "workbook": "y.xlsx", "prompt": "p2"},
    ]
    with pytest.raises(ValueError, match="duplicate task id"):
        runner.parse_tasks(data)


def test_parse_tasks_rejects_blank_prompt():
    with pytest.raises(ValueError, match="prompt"):
        runner.parse_tasks([{"id": "a", "workbook": "x.xlsx", "prompt": "   "}])


def test_load_tasks_reads_real_smoke_file():
    smoke = Path(__file__).resolve().parent.parent / "smoke_tasks.json"
    tasks = runner.load_tasks(smoke)
    assert len(tasks) == 2
    ids = {t.id for t in tasks}
    assert ids == {"simplemacro", "withdrawing"}
    for t in tasks:
        assert "reviewed" in t.prompt.lower() or "Reviewed" in t.prompt
        # no tool named in the prompt
        for banned in ("SafeXL", "safexl", "xlsx skill", "openpyxl"):
            assert banned not in t.prompt


def test_load_tasks_bad_json(tmp_path: Path):
    p = tmp_path / "tasks.json"
    p.write_text("{not json")
    with pytest.raises(ValueError, match="not valid JSON"):
        runner.load_tasks(p)


def test_load_tasks_missing_file(tmp_path: Path):
    with pytest.raises(ValueError, match="cannot read"):
        runner.load_tasks(tmp_path / "nope.json")


def test_resolve_workbook_relative_uses_repo_root():
    task = runner.Task(id="a", workbook="research/safexl-moat/fixtures/SimpleMacro.xlsm", prompt="p")
    resolved = runner.resolve_workbook(task)
    assert resolved.is_absolute()
    assert resolved == runner.REPO_ROOT / "research/safexl-moat/fixtures/SimpleMacro.xlsm"


def test_resolve_workbook_absolute_passthrough(tmp_path: Path):
    task = runner.Task(id="a", workbook=str(tmp_path / "x.xlsx"), prompt="p")
    assert runner.resolve_workbook(task) == tmp_path / "x.xlsx"


# ---------------------------------------------------------------------------
# parse_arms / build_matrix
# ---------------------------------------------------------------------------


def test_parse_arms_basic():
    assert runner.parse_arms("A,B") == ["A", "B"]
    assert runner.parse_arms("b") == ["B"]
    assert runner.parse_arms(" a , b ") == ["A", "B"]


def test_parse_arms_dedupes_preserving_order():
    assert runner.parse_arms("B,A,B") == ["B", "A"]


def test_parse_arms_rejects_unknown():
    with pytest.raises(ValueError, match="unknown arm"):
        runner.parse_arms("A,C")


def test_parse_arms_rejects_empty():
    with pytest.raises(ValueError, match="at least one arm"):
        runner.parse_arms("")


def test_build_matrix_cross_product():
    tasks = [runner.Task(id="t1", workbook="w1", prompt="p1"),
             runner.Task(id="t2", workbook="w2", prompt="p2")]
    cells = runner.build_matrix(tasks, ["A", "B"], 2)
    assert len(cells) == 8
    # task-major, then arm, then rep -- matches runs/<task>/<arm>/<rep> layout
    assert cells[0] == runner.Cell(task=tasks[0], arm="A", rep=1)
    assert cells[1] == runner.Cell(task=tasks[0], arm="A", rep=2)
    assert cells[2] == runner.Cell(task=tasks[0], arm="B", rep=1)
    assert cells[-1] == runner.Cell(task=tasks[1], arm="B", rep=2)


def test_build_matrix_rejects_bad_reps():
    with pytest.raises(ValueError, match="--reps"):
        runner.build_matrix([runner.Task(id="t", workbook="w", prompt="p")], ["A"], 0)


# ---------------------------------------------------------------------------
# run directory layout + resumability
# ---------------------------------------------------------------------------


def test_run_dir_layout(tmp_path: Path):
    cdir = runner.run_dir(tmp_path, "simplemacro", "B", 3)
    assert cdir == tmp_path / "simplemacro" / "B" / "3"
    assert runner.work_dir(cdir) == cdir / "work"
    assert runner.run_json_path(cdir) == cdir / "run.json"


def test_is_done_false_when_missing(tmp_path: Path):
    cdir = runner.run_dir(tmp_path, "t", "A", 1)
    assert runner.is_done(cdir) is False


def test_is_done_true_when_valid_json(tmp_path: Path):
    cdir = runner.run_dir(tmp_path, "t", "A", 1)
    cdir.mkdir(parents=True)
    runner.write_run_json(cdir, {"ok": True})
    assert runner.is_done(cdir) is True


def test_is_done_false_when_corrupt(tmp_path: Path):
    cdir = runner.run_dir(tmp_path, "t", "A", 1)
    cdir.mkdir(parents=True)
    (cdir / "run.json").write_text("{not valid json")
    assert runner.is_done(cdir) is False


def test_write_run_json_is_atomic_no_leftover_tmp(tmp_path: Path):
    cdir = runner.run_dir(tmp_path, "t", "A", 1)
    cdir.mkdir(parents=True)
    runner.write_run_json(cdir, {"a": 1})
    assert (cdir / "run.json").exists()
    assert not (cdir / "run.json.tmp").exists()
    assert json.loads((cdir / "run.json").read_text()) == {"a": 1}


def test_pending_cells_skips_done(tmp_path: Path):
    tasks = [runner.Task(id="t1", workbook="w1", prompt="p1")]
    cells = runner.build_matrix(tasks, ["A", "B"], 1)
    # mark the A cell done, leave B pending
    a_dir = runner.run_dir(tmp_path, "t1", "A", 1)
    a_dir.mkdir(parents=True)
    runner.write_run_json(a_dir, {"done": True})

    pending = runner.pending_cells(cells, tmp_path)
    assert len(pending) == 1
    assert pending[0].arm == "B"


def test_status_report_counts(tmp_path: Path):
    tasks = [runner.Task(id="t1", workbook="w1", prompt="p1"),
             runner.Task(id="t2", workbook="w2", prompt="p2")]
    a_dir = runner.run_dir(tmp_path, "t1", "A", 1)
    a_dir.mkdir(parents=True)
    runner.write_run_json(a_dir, {"done": True})

    report = runner.status_report(tasks, ["A", "B"], 1, tmp_path)
    assert "Arm A: 1/2 done" in report
    assert "Arm B: 0/2 done" in report
    assert "Total: 1/4 done" in report
    assert "t2/rep1" in report  # still pending for A
    assert "STOPPED" not in report


def test_status_report_flags_stopped_file(tmp_path: Path):
    tasks = [runner.Task(id="t1", workbook="w1", prompt="p1")]
    tmp_path.mkdir(exist_ok=True)
    (tmp_path / "STOPPED.json").write_text("{}")
    report = runner.status_report(tasks, ["A"], 1, tmp_path)
    assert "STOPPED" in report


# ---------------------------------------------------------------------------
# skill_block_args / build_command
# ---------------------------------------------------------------------------


def test_skill_block_args_disallowed():
    args = runner.skill_block_args("disallowed")
    assert args == ["--disallowedTools", "Skill(anthropic-skills:xlsx)"]


def test_skill_block_args_settings():
    args = runner.skill_block_args("settings")
    assert args[0] == "--settings"
    payload = json.loads(args[1])
    assert payload == {"permissions": {"deny": ["Skill(anthropic-skills:xlsx)"]}}


def test_skill_block_args_rejects_unknown_method():
    with pytest.raises(ValueError, match="unknown block-skill method"):
        runner.skill_block_args("nonsense")


def test_build_command_arm_a_has_no_plugin_or_deny():
    cmd = runner.build_command("claude", "do the thing", "A", "disallowed")
    assert cmd[0] == "claude"
    assert "-p" in cmd and "do the thing" in cmd
    assert "--allowedTools" in cmd
    idx = cmd.index("--allowedTools")
    assert cmd[idx + 1] == runner.ALLOWED_TOOLS
    assert "--plugin-dir" not in cmd
    assert "--disallowedTools" not in cmd
    # Clean-room overlay only: hooks off, only the bench-tree deny rules (no Skill rule),
    # so Anthropic's skill stays usable.
    settings = json.loads(cmd[cmd.index("--settings") + 1])
    assert settings["disableAllHooks"] is True
    assert settings["permissions"]["deny"] == runner.bench_tree_deny_rules()
    assert cmd[cmd.index("--model") + 1] == runner.BENCH_MODEL
    assert cmd[cmd.index("--setting-sources") + 1] == "user,local"
    assert cmd[-2:] == ["--output-format", "stream-json"] or "--verbose" in cmd


def test_build_command_arm_b_has_plugin_and_deny():
    cmd = runner.build_command("claude", "do the thing", "B", "disallowed")
    assert "--plugin-dir" in cmd
    plugin_idx = cmd.index("--plugin-dir")
    assert cmd[plugin_idx + 1] == runner.SAFEXL_PLUGIN_DIR
    assert "--disallowedTools" in cmd
    deny_idx = cmd.index("--disallowedTools")
    assert cmd[deny_idx + 1] == f"Skill({runner.ANTHROPIC_XLSX_SKILL})"


def test_build_command_arm_b_settings_method():
    cmd = runner.build_command("claude", "prompt", "B", "settings")
    assert "--settings" in cmd
    assert "--disallowedTools" not in cmd


def test_build_command_allowed_tools_identical_both_arms():
    cmd_a = runner.build_command("claude", "p", "A", "disallowed")
    cmd_b = runner.build_command("claude", "p", "B", "disallowed")
    idx_a = cmd_a.index("--allowedTools")
    idx_b = cmd_b.index("--allowedTools")
    assert cmd_a[idx_a + 1] == cmd_b[idx_b + 1] == runner.ALLOWED_TOOLS


def test_build_command_prompt_never_mentions_a_tool():
    # guards the runner itself, not just the tasks file: build_command must
    # not inject any tool name into the prompt string it was given.
    prompt = "add a Reviewed column"
    cmd = runner.build_command("claude", prompt, "B", "disallowed")
    p_idx = cmd.index("-p")
    assert cmd[p_idx + 1] == prompt


# ---------------------------------------------------------------------------
# build_env
# ---------------------------------------------------------------------------


def test_build_env_strips_api_key_both_arms():
    base = {"PATH": "/usr/bin", "ANTHROPIC_API_KEY": "sk-should-not-survive"}
    for arm in ("A", "B"):
        env = runner.build_env(base, arm)
        assert "ANTHROPIC_API_KEY" not in env


def test_build_env_disables_claude_mds_both_arms():
    base = {"PATH": "/usr/bin"}
    for arm in ("A", "B"):
        env = runner.build_env(base, arm)
        assert env["CLAUDE_CODE_DISABLE_CLAUDE_MDS"] == "1"


def test_build_env_arm_a_strips_safexl_from_path_even_if_inherited():
    base = {"PATH": f"/usr/bin{os.pathsep}{runner.SAFEXL_VENV_BIN}{os.pathsep}/bin"}
    env = runner.build_env(base, "A")
    parts = env["PATH"].split(os.pathsep)
    assert runner.SAFEXL_VENV_BIN not in parts


def test_build_env_arm_b_prepends_safexl_bin():
    base = {"PATH": "/usr/bin:/bin"}
    env = runner.build_env(base, "B")
    parts = env["PATH"].split(os.pathsep)
    assert parts[0] == runner.SAFEXL_VENV_BIN
    assert "/usr/bin" in parts


def test_build_env_arm_b_does_not_duplicate_safexl_bin():
    base = {"PATH": f"{runner.SAFEXL_VENV_BIN}{os.pathsep}/usr/bin"}
    env = runner.build_env(base, "B")
    parts = env["PATH"].split(os.pathsep)
    assert parts.count(runner.SAFEXL_VENV_BIN) == 1
    assert parts[0] == runner.SAFEXL_VENV_BIN


def test_build_env_does_not_mutate_input():
    base = {"PATH": "/usr/bin", "ANTHROPIC_API_KEY": "x"}
    runner.build_env(base, "A")
    assert base["ANTHROPIC_API_KEY"] == "x"  # original dict untouched


# ---------------------------------------------------------------------------
# summarize_events
# ---------------------------------------------------------------------------


def _assistant_tool_use(tool_use_id, name, tool_input, text=None):
    content = []
    if text is not None:
        content.append({"type": "text", "text": text})
    content.append({"type": "tool_use", "id": tool_use_id, "name": name, "input": tool_input})
    return {"type": "assistant", "message": {"role": "assistant", "content": content}}


def _tool_result(tool_use_id, text, is_error=False):
    return {
        "type": "user",
        "message": {
            "role": "user",
            "content": [
                {
                    "type": "tool_result",
                    "tool_use_id": tool_use_id,
                    "content": [{"type": "text", "text": text}],
                    "is_error": is_error,
                }
            ],
        },
    }


def test_summarize_events_model_from_init():
    events = [{"type": "system", "subtype": "init", "model": "claude-opus-5-5"}]
    summary = runner.summarize_events(events)
    assert summary["model"] == "claude-opus-5-5"


def test_summarize_events_result_fields_total_cost_usd_num_turns():
    events = [
        {"type": "result", "subtype": "success", "is_error": False,
         "total_cost_usd": 0.51, "num_turns": 9, "result": "all done"},
    ]
    summary = runner.summarize_events(events)
    assert summary["total_cost_usd"] == 0.51
    assert summary["num_turns"] == 9
    assert summary["result_subtype"] == "success"
    assert summary["is_error"] is False
    assert summary["final_text"] == "all done"


def test_summarize_events_result_fields_fallback_names():
    # some CLI versions/paths might use cost_usd/turns instead -- must not crash
    events = [{"type": "result", "cost_usd": 0.12, "turns": 3, "result": "ok"}]
    summary = runner.summarize_events(events)
    assert summary["total_cost_usd"] == 0.12
    assert summary["num_turns"] == 3


def test_summarize_events_final_text_falls_back_to_last_assistant_text():
    events = [
        {"type": "assistant", "message": {"role": "assistant",
         "content": [{"type": "text", "text": "first"}]}},
        {"type": "assistant", "message": {"role": "assistant",
         "content": [{"type": "text", "text": "last words"}]}},
        # no "result" event at all (e.g. killed mid-run)
    ]
    summary = runner.summarize_events(events)
    assert summary["final_text"] == "last words"


def test_summarize_events_skill_invoked_and_succeeded():
    events = [
        _assistant_tool_use("tu1", "Skill", {"skill": "safexl"}),
        _tool_result("tu1", "ran safexl edit ok", is_error=False),
    ]
    summary = runner.summarize_events(events)
    assert summary["skills_invoked"] == ["safexl"]
    assert summary["skills_succeeded"] == ["safexl"]
    assert summary["skills_denied"] == []


def test_summarize_events_skill_invoked_and_denied():
    events = [
        _assistant_tool_use("tu1", "Skill", {"skill": "anthropic-skills:xlsx"}),
        _tool_result("tu1", "Permission denied: this skill is disallowed for this session", is_error=True),
    ]
    summary = runner.summarize_events(events)
    assert summary["skills_invoked"] == ["anthropic-skills:xlsx"]
    assert summary["skills_denied"] == ["anthropic-skills:xlsx"]
    assert summary["skills_succeeded"] == []
    assert "anthropic-skills:xlsx" in summary["skill_denial_evidence"]


def test_summarize_events_skill_denied_by_text_even_without_is_error_flag():
    # be defensive: some denial paths might not set is_error, but do say so in text
    events = [
        _assistant_tool_use("tu1", "Skill", {"skill": "anthropic-skills:xlsx"}),
        _tool_result("tu1", "This request was not permitted by disallowedTools.", is_error=False),
    ]
    summary = runner.summarize_events(events)
    assert summary["skills_denied"] == ["anthropic-skills:xlsx"]


def test_summarize_events_safexl_bash_detected():
    events = [
        _assistant_tool_use("tu1", "Bash", {"command": "safexl inspect work/SimpleMacro.xlsm --json"}),
        _tool_result("tu1", "{...}", is_error=False),
        _assistant_tool_use("tu2", "Read", {"file_path": "work/SimpleMacro.xlsm"}),
    ]
    summary = runner.summarize_events(events)
    assert summary["safexl_used"] is True
    assert summary["safexl_commands"] == ["safexl inspect work/SimpleMacro.xlsm --json"]
    assert set(summary["tools_called"]) == {"Bash", "Read"}


def test_summarize_events_no_safexl_bash():
    events = [_assistant_tool_use("tu1", "Bash", {"command": "ls -la"})]
    summary = runner.summarize_events(events)
    assert summary["safexl_used"] is False
    assert summary["safexl_commands"] == []


def test_summarize_events_tools_called_dedup_preserves_first_seen_order():
    events = [
        _assistant_tool_use("tu1", "Read", {}),
        _assistant_tool_use("tu2", "Bash", {"command": "ls"}),
        _assistant_tool_use("tu3", "Read", {}),
    ]
    summary = runner.summarize_events(events)
    assert summary["tools_called"] == ["Read", "Bash"]
    assert summary["tool_call_count"] == 3


def test_summarize_events_ignores_malformed_blocks():
    events = [
        {"type": "assistant", "message": {"content": ["not-a-dict", {"type": "text"}]}},
        {"type": "user", "message": {"content": [{"type": "tool_result"}]}},  # no tool_use_id
    ]
    # should not raise
    summary = runner.summarize_events(events)
    assert summary["tools_called"] == []


# ---------------------------------------------------------------------------
# iter_json_lines
# ---------------------------------------------------------------------------


def test_iter_json_lines_skips_blank_and_garbage():
    text = '{"a":1}\n\n   \nnot json at all\n{"b":2}\n'
    parsed = list(runner.iter_json_lines(text))
    assert parsed == [{"a": 1}, {"b": 2}]


# ---------------------------------------------------------------------------
# detect_rate_limit
# ---------------------------------------------------------------------------


def test_detect_rate_limit_none_when_clean():
    events = [{"type": "result", "subtype": "success", "result": "all good"}]
    assert runner.detect_rate_limit(events, "") is None


def test_detect_rate_limit_matches_in_result_event():
    events = [{"type": "result", "subtype": "error", "result": "You have hit your usage limit for this session."}]
    evidence = runner.detect_rate_limit(events, "")
    assert evidence is not None
    assert "usage limit" in evidence.lower()


def test_detect_rate_limit_matches_429_in_stderr():
    evidence = runner.detect_rate_limit([], "upstream returned 429 Too Many Requests")
    assert evidence is not None
    assert "429" in evidence


def test_detect_rate_limit_ignores_tool_result_text():
    # False positive 2026-09-29: a session Read another run's run.json, whose text
    # says "rate_limited" / "rate_limit_evidence"; the batch stopped as if limited.
    events = [_tool_result("tu1", '"rate_limited": false,\n"rate_limit_evidence": null, "usage limit"')]
    assert runner.detect_rate_limit(events, "") is None


# ---------------------------------------------------------------------------
# snapshot_dir / new_files
# ---------------------------------------------------------------------------


def test_snapshot_and_new_files(tmp_path: Path):
    (tmp_path / "SimpleMacro.xlsm").write_bytes(b"orig")
    before = runner.snapshot_dir(tmp_path)
    assert before == {"SimpleMacro.xlsm"}

    (tmp_path / "SimpleMacro-updated.xlsm").write_bytes(b"new")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "scratch.tmp").write_bytes(b"x")

    after = runner.snapshot_dir(tmp_path)
    assert runner.new_files(before, after) == ["SimpleMacro-updated.xlsm", "sub/scratch.tmp"]


def test_snapshot_dir_missing_returns_empty(tmp_path: Path):
    assert runner.snapshot_dir(tmp_path / "does-not-exist") == set()


# ---------------------------------------------------------------------------
# argument parser wiring
# ---------------------------------------------------------------------------


def test_arg_parser_defaults():
    parser = runner.build_arg_parser()
    args = parser.parse_args([])
    assert args.arms == "A,B"
    assert args.reps == 1
    assert args.timeout == runner.DEFAULT_TIMEOUT_SECONDS
    assert args.block_skill_method == "disallowed"
    assert args.tasks == runner.SCRIPT_DIR / "smoke_tasks.json"
    assert args.out == runner.SCRIPT_DIR / "runs"


def test_arg_parser_rejects_bad_block_skill_method():
    parser = runner.build_arg_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["--block-skill-method", "nonsense"])


def test_main_status_on_empty_out_dir(tmp_path: Path, capsys):
    smoke = Path(__file__).resolve().parent.parent / "smoke_tasks.json"
    rc = runner.main(["--tasks", str(smoke), "--arms", "A,B", "--reps", "1",
                       "--out", str(tmp_path / "runs"), "--status"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "Arm A: 0/2 done" in out
    assert "Arm B: 0/2 done" in out


def test_main_rejects_bad_tasks_file(tmp_path: Path, capsys):
    bad = tmp_path / "tasks.json"
    bad.write_text("[]")
    rc = runner.main(["--tasks", str(bad), "--status", "--out", str(tmp_path / "runs")])
    assert rc == 2
    err = capsys.readouterr().err
    assert "at least one task" in err


def test_build_env_arm_b_env_file_puts_safexl_on_bash_path():
    """Claude Code's Bash tool rebuilds PATH from shell startup files, so arm B
    needs CLAUDE_ENV_FILE; arm A must never get one (even if inherited)."""
    env_b = runner.build_env({"PATH": "/usr/bin", "CLAUDE_ENV_FILE": "/x"}, "B")
    assert env_b["CLAUDE_ENV_FILE"] == str(runner.arm_b_env_file())
    assert runner.SAFEXL_VENV_BIN in Path(env_b["CLAUDE_ENV_FILE"]).read_text()
    env_a = runner.build_env({"PATH": "/usr/bin", "CLAUDE_ENV_FILE": "/x"}, "A")
    assert "CLAUDE_ENV_FILE" not in env_a


def test_operator_hint_is_opt_in_and_symmetric(monkeypatch):
    monkeypatch.delenv("XLSX_BENCH_OPERATOR_HINT", raising=False)
    plain = runner.build_command("claude", "p", "A", "disallowed")
    assert "--append-system-prompt" not in plain
    monkeypatch.setenv("XLSX_BENCH_OPERATOR_HINT", "1")
    for arm in ("A", "B"):
        cmd = runner.build_command("claude", "p", arm, "disallowed")
        i = cmd.index("--append-system-prompt")
        assert cmd[i + 1] == runner.OPERATOR_HINTS[arm]
        assert cmd[-3:] == ["--output-format", "stream-json", "--verbose"]  # record.py slices [3:-3]
