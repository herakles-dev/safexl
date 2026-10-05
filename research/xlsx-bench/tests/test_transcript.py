"""Unit tests for research/xlsx-bench/transcript.py.

Uses synthetic events.jsonl / session.jsonl fixtures shaped like the two
real formats (see record.py's module docstring for how they were verified
against real Claude Code output). No live `claude` invocation, no network.
"""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

_RUNNER_PATH = Path(__file__).resolve().parent.parent / "runner.py"
_spec_runner = importlib.util.spec_from_file_location("xlsx_bench_runner_for_transcript", _RUNNER_PATH)
_runner_mod = importlib.util.module_from_spec(_spec_runner)
assert _spec_runner.loader is not None
sys.modules["xlsx_bench_runner_for_transcript"] = _runner_mod
_spec_runner.loader.exec_module(_runner_mod)

_RECORD_PATH = Path(__file__).resolve().parent.parent / "record.py"
_spec_record = importlib.util.spec_from_file_location("xlsx_bench_record_for_transcript", _RECORD_PATH)
_record_mod = importlib.util.module_from_spec(_spec_record)
assert _spec_record.loader is not None
sys.modules["xlsx_bench_record_for_transcript"] = _record_mod
_spec_record.loader.exec_module(_record_mod)

_TRANSCRIPT_PATH = Path(__file__).resolve().parent.parent / "transcript.py"
_spec_t = importlib.util.spec_from_file_location("xlsx_bench_transcript", _TRANSCRIPT_PATH)
t = importlib.util.module_from_spec(_spec_t)
assert _spec_t.loader is not None
sys.modules["xlsx_bench_transcript"] = t
_spec_t.loader.exec_module(t)

runner = t.runner


# ---------------------------------------------------------------------------
# redact_text
# ---------------------------------------------------------------------------


def test_redact_text_keeps_relative_paths():
    assert t.redact_text("work/SimpleMacro.xlsm", "/home/h/work") == "work/SimpleMacro.xlsm"


def test_redact_text_rewrites_in_work_absolute_path():
    work = "/home/alice/project/.claude/worktrees/x/work"
    text = f"saved to {work}/SimpleMacro-updated.xlsm"
    out = t.redact_text(text, work)
    assert out == "saved to <work>/SimpleMacro-updated.xlsm"
    assert "alice" not in out


def test_redact_text_rewrites_work_dir_itself():
    work = "/home/h/work"
    out = t.redact_text(f"cwd is {work}", work)
    assert out == "cwd is <work>"


def test_redact_text_strips_outside_absolute_path():
    work = "/home/h/work"
    out = t.redact_text("config lives at /home/h/.secrets/prod.env", work)
    assert out == "config lives at <path>"
    assert "secrets" not in out


def test_redact_text_does_not_falsely_match_prefix_ancestor_as_inside():
    # "/home/h" is a *prefix string* of the work dir but is NOT the work dir
    # or a path under it -- must still be redacted, not kept.
    work = "/home/h/work"
    out = t.redact_text("see /home/h for details", work)
    assert out == "see <path> for details"


def test_redact_text_handles_none_and_empty():
    assert t.redact_text(None, "/x") == ""
    assert t.redact_text("", "/x") == ""


def test_redact_text_multiple_paths_in_one_string():
    work = "/w"
    text = "/w/out.xlsx and also /etc/passwd"
    out = t.redact_text(text, work)
    assert out == "<work>/out.xlsx and also <path>"


def test_redact_text_strips_bare_username_with_no_path_involved():
    # regression: `ls -la` output leaks the OS username in the ownership
    # columns ("drwxr-xr-x 2 <user> <user> ...") with zero absolute paths
    # present, so _ABS_PATH_RE alone never touches it.
    username = t._USERNAME
    assert username, "test machine must have a resolvable home directory name"
    text = f"drwxr-xr-x 2 {username} {username} 4096 Sep 26 02:36 ."
    out = t.redact_text(text, "/some/work/dir")
    assert username not in out
    assert out.count("<user>") == 2


def test_redact_text_username_substring_not_over_matched():
    # word-boundary match only -- a longer word that merely contains the
    # username as a substring must survive untouched.
    username = t._USERNAME
    text = f"{username}xyz and xyz{username}"
    out = t.redact_text(text, "/w")
    assert out == text


# ---------------------------------------------------------------------------
# detect_kind / load_entries
# ---------------------------------------------------------------------------


def test_detect_kind_recording(tmp_path):
    (tmp_path / "session.jsonl").write_text("{}\n")
    assert t.detect_kind(tmp_path) == "recording"


def test_detect_kind_headless(tmp_path):
    (tmp_path / "events.jsonl").write_text("{}\n")
    assert t.detect_kind(tmp_path) == "headless"


def test_detect_kind_raises_when_neither(tmp_path):
    try:
        t.detect_kind(tmp_path)
        assert False, "expected FileNotFoundError"
    except FileNotFoundError:
        pass


def test_load_entries_reads_matching_file(tmp_path):
    (tmp_path / "session.jsonl").write_text('{"a":1}\n{"b":2}\n')
    entries = t.load_entries(tmp_path, "recording")
    assert entries == [{"a": 1}, {"b": 2}]


# ---------------------------------------------------------------------------
# extract_model
# ---------------------------------------------------------------------------


def test_extract_model_from_headless_init_event():
    entries = [{"type": "system", "subtype": "init", "model": "claude-opus-5-5"}]
    assert t.extract_model(entries, {}) == "claude-opus-5-5"


def test_extract_model_from_recording_assistant_entry():
    entries = [{"type": "assistant", "message": {"model": "claude-opus-5-5"}}]
    assert t.extract_model(entries, {}) == "claude-opus-5-5"


def test_extract_model_falls_back_to_run_json():
    assert t.extract_model([], {"model": "fallback-model"}) == "fallback-model"


# ---------------------------------------------------------------------------
# work_dir_from_run_json
# ---------------------------------------------------------------------------


def test_work_dir_from_run_json_prefers_invocation_cwd(tmp_path):
    run_json = {"invocation": {"cwd": "/some/original/worktree/path/work"}}
    assert t.work_dir_from_run_json(tmp_path, run_json) == "/some/original/worktree/path/work"


def test_work_dir_from_run_json_falls_back_to_current_location(tmp_path):
    assert t.work_dir_from_run_json(tmp_path, {}) == str((tmp_path / "work").resolve())


# ---------------------------------------------------------------------------
# build_conversation / extract_final_message
# ---------------------------------------------------------------------------


def test_build_conversation_extracts_text_tool_call_and_result():
    entries = [
        {
            "type": "assistant",
            "message": {
                "content": [
                    {"type": "thinking", "thinking": "hmm"},
                    {"type": "tool_use", "id": "tu1", "name": "Bash", "input": {"command": "ls"}},
                ]
            },
        },
        {
            "type": "user",
            "message": {
                "content": [
                    {"type": "tool_result", "tool_use_id": "tu1", "content": "file1\nfile2", "is_error": False}
                ]
            },
        },
        {"type": "assistant", "message": {"content": [{"type": "text", "text": "done"}]}},
    ]
    items = t.build_conversation(entries)
    kinds = [i["kind"] for i in items]
    assert kinds == ["tool_call", "tool_result", "text"]
    assert items[0]["name"] == "Bash"
    assert items[1]["text"] == "file1\nfile2"
    assert items[2]["text"] == "done"


def test_build_conversation_skips_non_conversation_line_types():
    entries = [
        {"type": "queue-operation", "operation": "enqueue"},
        {"type": "system", "subtype": "init", "model": "m", "tools": ["Read", "Write"], "skills": ["a", "b"]},
        {"type": "attachment", "attachment": {}},
        {"type": "cost-state", "totalCostUSD": 1.0},
        {"type": "assistant", "message": {"content": [{"type": "text", "text": "hello"}]}},
    ]
    items = t.build_conversation(entries)
    assert items == [{"kind": "text", "text": "hello"}]


def test_extract_final_message_uses_last_text_item():
    items = [{"kind": "text", "text": "first"}, {"kind": "tool_call", "name": "Bash", "input": {}, "id": "x"},
             {"kind": "text", "text": "last"}]
    assert t.extract_final_message(items, {}) == "last"


def test_extract_final_message_falls_back_to_run_json():
    assert t.extract_final_message([], {"final_text": "fallback"}) == "fallback"


# ---------------------------------------------------------------------------
# format_input_json / trim_result_text
# ---------------------------------------------------------------------------


def test_format_input_json_is_pretty_printed():
    out = t.format_input_json({"skill": "safexl"})
    assert json.loads(out) == {"skill": "safexl"}
    assert "\n" in out


def test_trim_result_text_under_limit_unchanged():
    text = "\n".join(f"line{i}" for i in range(5))
    trimmed, more = t.trim_result_text(text, limit=40)
    assert trimmed == text
    assert more == 0


def test_trim_result_text_over_limit_truncates():
    text = "\n".join(f"line{i}" for i in range(50))
    trimmed, more = t.trim_result_text(text, limit=40)
    assert len(trimmed.splitlines()) == 40
    assert more == 10


# ---------------------------------------------------------------------------
# gather_manifest
# ---------------------------------------------------------------------------


def test_gather_manifest_prefers_existing_manifest_json(tmp_path):
    (tmp_path / "manifest.json").write_text(json.dumps({
        "workbook": {"path": "orig.xlsx", "sha256": "deadbeef"},
        "output_files": [{"path": "out.xlsx", "sha256": "cafebabe"}],
        "session_cast": {"path": "session.cast", "sha256": "feedface"},
        "session_jsonl": {"path": "session.jsonl", "sha256": "abcdef01"},
    }))
    manifest = t.gather_manifest(tmp_path, {}, "recording")
    assert manifest["workbook"]["sha256"] == "deadbeef"
    assert manifest["cast_file"]["sha256"] == "feedface"
    assert manifest["transcript_file"]["sha256"] == "abcdef01"


def test_gather_manifest_computes_when_missing(tmp_path):
    wdir = tmp_path / "work"
    wdir.mkdir()
    # the workbook's OWN copy in this run's work/ -- gather_manifest must
    # hash this, never the (possibly since-changed) repo fixture path.
    (wdir / "SimpleMacro.xlsm").write_bytes(b"whatever-is-in-this-runs-workdir")
    (wdir / "out.xlsm").write_bytes(b"hello")
    (tmp_path / "session.jsonl").write_text('{"type":"assistant"}\n')
    run_json = {
        "workbook": "research/safexl-moat/fixtures/SimpleMacro.xlsm",
        "output_files": ["out.xlsm"],
    }
    manifest = t.gather_manifest(tmp_path, run_json, "recording")
    assert manifest["workbook"]["path"] == "SimpleMacro.xlsm"
    assert manifest["workbook"]["sha256_before"] is None  # no manifest.json -> unknown
    assert manifest["workbook"]["sha256_after"] == _record_mod.sha256_file(wdir / "SimpleMacro.xlsm")
    assert manifest["output_files"] == [
        {"path": "out.xlsm", "sha256": _record_mod.sha256_file(wdir / "out.xlsm")}
    ]
    assert manifest["transcript_file"]["path"] == "session.jsonl"
    assert manifest["cast_file"] is None


def test_gather_manifest_skips_missing_output_files(tmp_path):
    (tmp_path / "session.jsonl").write_text("{}\n")
    run_json = {"workbook": "does/not/exist.xlsx", "output_files": ["missing.xlsx"]}
    manifest = t.gather_manifest(tmp_path, run_json, "recording")
    assert manifest["workbook"] is None
    assert manifest["output_files"] == []


# ---------------------------------------------------------------------------
# render_markdown / render_html: redaction + catalog stripping, end to end
# ---------------------------------------------------------------------------


def _sample_run_json(work_dir_abs):
    return {
        "task_id": "budget-t1-electric-increase",
        "arm": "B",
        "rep": 1,
        "model": "claude-opus-5-5",
        "wall_seconds": 42.1,
        "prompt": "My electric bill went up this month — change it to $135.",
        "workbook": "research/xlsx-bench/corpus/simple-monthly-budget.xlsx",
        "output_files": ["simple-monthly-budget-updated.xlsx"],
        "invocation": {"cwd": work_dir_abs},
        "final_text": "Updated the electric bill to $135.",
    }


def _sample_items(work_dir_abs):
    secret_path = "/home/alice/.secrets/prod.env"
    return [
        {"kind": "tool_call", "name": "Bash",
         "input": {"command": f"safexl edit {work_dir_abs}/simple-monthly-budget.xlsx --set 'C12=135'"},
         "id": "tu1"},
        {"kind": "tool_result", "id": "tu1", "name": "Bash",
         "text": f"wrote {work_dir_abs}/simple-monthly-budget-updated.xlsx (see also {secret_path})",
         "is_error": False},
        {"kind": "text", "text": "Updated the electric bill to $135."},
    ]


def test_render_markdown_redacts_and_omits_catalog():
    work_dir_abs = "/home/alice/project/.claude/worktrees/x/research/xlsx-bench/recordings/t/B/1/work"
    run_json = _sample_run_json(work_dir_abs)
    items = _sample_items(work_dir_abs)
    manifest = {
        "workbook": {"path": "simple-monthly-budget.xlsx", "sha256_before": "aa" * 32,
                     "sha256_after": "aa" * 32, "modified_in_place": False},
        "output_files": [{"path": "simple-monthly-budget-updated.xlsx", "sha256": "bb" * 32}],
        "transcript_file": {"path": "session.jsonl", "sha256": "cc" * 32},
        "cast_file": {"path": "session.cast", "sha256": "dd" * 32},
    }
    md = t.render_markdown(run_json, "recording", "claude-opus-5-5", manifest, items, "Updated the electric bill to $135.", work_dir_abs)

    assert "$135" in md
    assert "sha256:" + "aa" * 32 in md
    assert "<work>/simple-monthly-budget.xlsx" in md
    assert "<path>" in md  # the redacted secrets path
    assert "alice" not in md  # no leaked username/homedir anywhere
    assert ".secrets" not in md
    # never leaked a raw skill/agent/tool catalog (nothing in this fixture
    # data mentions one, but assert the literal marker text a system/init
    # event would carry is genuinely absent, confirming build_conversation's
    # skip actually reached the renderer)
    assert "hugging-face" not in md
    assert "example-skill" not in md


def test_render_html_is_self_contained_and_redacts():
    work_dir_abs = "/home/alice/project/.claude/worktrees/x/work"
    run_json = _sample_run_json(work_dir_abs)
    items = _sample_items(work_dir_abs)
    manifest = {
        "workbook": {"path": "simple-monthly-budget.xlsx", "sha256_before": "aa" * 32,
                     "sha256_after": "aa" * 32, "modified_in_place": False},
        "output_files": [], "transcript_file": None, "cast_file": None,
    }
    out_html = t.render_html(run_json, "recording", "claude-opus-5-5", manifest, items, "Updated the electric bill to $135.", work_dir_abs)

    assert out_html.startswith("<!doctype html>")
    assert "<script" not in out_html  # no JS
    assert "http://" not in out_html and "https://" not in out_html  # no external requests
    assert "alice" not in out_html
    assert "aa" * 32 in out_html  # the manifest hash actually rendered
    # HTML-escaped, since redaction runs before html.escape()
    assert "&lt;work&gt;/simple-monthly-budget.xlsx" in out_html
    assert "&lt;path&gt;" in out_html


# ---------------------------------------------------------------------------
# render_run_dir: full integration on a synthetic recording folder
# ---------------------------------------------------------------------------


def test_render_run_dir_recording_end_to_end(tmp_path):
    run_dir = tmp_path / "budget-t1-electric-increase" / "B" / "1"
    wdir = run_dir / "work"
    wdir.mkdir(parents=True)
    (wdir / "simple-monthly-budget.xlsx").write_bytes(b"orig")
    (wdir / "simple-monthly-budget-updated.xlsx").write_bytes(b"new")

    session_lines = [
        {"type": "queue-operation", "operation": "enqueue", "content": "prompt text"},
        {"type": "assistant", "requestId": "r1",
         "message": {"model": "claude-opus-5-5", "stop_reason": "tool_use",
                     "content": [{"type": "tool_use", "id": "tu1", "name": "Skill", "input": {"skill": "safexl"}}]}},
        {"type": "user", "message": {"content": [
            {"type": "tool_result", "tool_use_id": "tu1", "content": "Launching skill: safexl", "is_error": False}
        ]}},
        {"type": "assistant", "requestId": "r2",
         "message": {"model": "claude-opus-5-5", "stop_reason": "end_turn",
                     "content": [{"type": "text", "text": "Updated the electric bill to $135."}]}},
        {"type": "cost-state", "totalCostUSD": 0.12},
    ]
    (run_dir / "session.jsonl").write_text("\n".join(json.dumps(l) for l in session_lines) + "\n")

    run_json = {
        "task_id": "budget-t1-electric-increase", "arm": "B", "rep": 1,
        "model": "claude-opus-5-5", "wall_seconds": 12.3,
        "prompt": "My electric bill went up this month — change it to $135.",
        "workbook": "research/xlsx-bench/corpus/simple-monthly-budget.xlsx",
        "output_files": ["simple-monthly-budget-updated.xlsx"],
        "invocation": {"cwd": str(wdir)},
    }
    (run_dir / "run.json").write_text(json.dumps(run_json))

    md_path, html_path = t.render_run_dir(run_dir)
    assert md_path == run_dir / "transcript.md"
    assert html_path == run_dir / "transcript.html"
    md = md_path.read_text()
    out_html = html_path.read_text()

    assert "$135" in md
    assert "safexl" in md
    assert "Updated the electric bill to $135." in md
    assert str(wdir) not in md  # in-work absolute path was rewritten to <work>/...
    assert "<work>" in md or "work" in md

    assert "$135" in out_html
    assert str(wdir) not in out_html


def test_render_run_dir_missing_run_json_raises(tmp_path):
    run_dir = tmp_path / "empty"
    run_dir.mkdir()
    try:
        t.render_run_dir(run_dir)
        assert False, "expected FileNotFoundError"
    except FileNotFoundError:
        pass


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def test_main_reports_error_for_bad_dir(tmp_path, capsys):
    rc = t.main([str(tmp_path / "does-not-exist")])
    assert rc == 1
    assert "error" in capsys.readouterr().err


def test_main_writes_files_for_good_dir(tmp_path, capsys):
    run_dir = tmp_path / "run"
    wdir = run_dir / "work"
    wdir.mkdir(parents=True)
    (run_dir / "session.jsonl").write_text(
        json.dumps({"type": "assistant", "message": {"content": [{"type": "text", "text": "hi"}]}}) + "\n"
    )
    (run_dir / "run.json").write_text(json.dumps({"task_id": "x", "arm": "A", "rep": 1, "prompt": "p"}))

    rc = t.main([str(run_dir)])
    assert rc == 0
    assert (run_dir / "transcript.md").exists()
    assert (run_dir / "transcript.html").exists()
    assert "wrote transcript.md" in capsys.readouterr().out
