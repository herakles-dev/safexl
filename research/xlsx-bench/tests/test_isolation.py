"""Tests for the clean-room isolation added 2026-09-29: temp cwd, deny rules,
contamination detector, rate-limit detector (false-positive case), and
grade.py's --exclude-contaminated. Pure logic only; no live claude/tmux."""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent.parent


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, _HERE / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


runner = _load("xlsx_bench_runner_iso", "runner.py")
rec = _load("xlsx_bench_record_iso", "record.py")
grade = _load("xlsx_bench_grade_iso", "grade.py")
scan = _load("xlsx_bench_scan_iso", "scan_contamination.py")

BENCH = str(_HERE)
RUN_DIR = f"{BENCH}/recordings/round/task/A/1"


def _tool_use(tid, name, inp):
    return {"type": "assistant", "message": {"role": "assistant", "content": [
        {"type": "tool_use", "id": tid, "name": name, "input": inp}]}}


def _result(tid, text):
    return {"type": "user", "message": {"role": "user", "content": [
        {"type": "tool_result", "tool_use_id": tid, "content": [{"type": "text", "text": text}]}]}}


def _text(text):
    return {"type": "assistant", "message": {"role": "assistant", "content": [{"type": "text", "text": text}]}}


# --- deny rules -----------------------------------------------------------


def test_deny_rules_cover_bench_tree_absolute_and_bash():
    rules = runner.bench_tree_deny_rules(Path("/x/research/xlsx-bench"))
    assert "Read(//x/research/xlsx-bench/**)" in rules
    assert "Edit(//x/research/xlsx-bench/**)" in rules
    assert "Bash(*xlsx-bench*)" in rules


def test_deny_rules_do_not_touch_safexl_package():
    joined = " ".join(runner.bench_tree_deny_rules())
    assert "packages/safexl" not in joined


@pytest.mark.parametrize("method", ["disallowed", "settings"])
def test_both_arms_get_identical_bench_deny_rules(method):
    def denies(arm):
        cmd = runner.build_command("claude", "p", arm, method)
        return json.loads(cmd[cmd.index("--settings") + 1])["permissions"]["deny"]

    a, b = denies("A"), denies("B")
    bench = runner.bench_tree_deny_rules()
    assert all(r in a for r in bench) and all(r in b for r in bench)
    assert [r for r in a if r not in bench] == []
    assert [r for r in b if r not in bench] == ([f"Skill({runner.ANTHROPIC_XLSX_SKILL})"] if method == "settings" else [])


# --- isolated workdir -----------------------------------------------------


def test_isolated_workdir_is_fresh_unique_and_outside_repo(tmp_path):
    a = rec.make_isolated_workdir(str(tmp_path))
    b = rec.make_isolated_workdir(str(tmp_path))
    assert a != b and a.is_dir() and b.is_dir()
    assert a.name.startswith("xlsxb-") and a.parent == tmp_path.resolve()
    assert list(a.iterdir()) == []
    assert _HERE.parent.parent not in a.parents


def test_copy_back_replaces_destination(tmp_path):
    src = tmp_path / "tmpwork"
    src.mkdir()
    (src / "out.xlsx").write_bytes(b"new")
    dest = tmp_path / "cell" / "work"
    dest.mkdir(parents=True)
    (dest / "stale.txt").write_text("old")
    rec.copy_back(src, dest)
    assert (dest / "out.xlsx").read_bytes() == b"new"
    assert not (dest / "stale.txt").exists()


def test_record_cell_runs_in_temp_dir_and_copies_back(tmp_path, monkeypatch):
    wb = tmp_path / "src.xlsx"
    wb.write_bytes(b"orig")
    task = runner.Task(id="t1", workbook=str(wb), prompt="do it")
    cdir = tmp_path / "cell"
    cdir.mkdir()
    transcript = tmp_path / "s.jsonl"
    transcript.write_text(json.dumps(_tool_use("t", "Bash", {"command": "ls /home/alice/project/research/xlsx-bench/recordings"})) + "\n")
    seen = {}

    def fake_launch(claude_argv, wdir, arm, session_name, timeout, poll, quiet, asciinema_bin, cast_path):
        seen["cwd"] = wdir
        assert wdir.is_dir() and (wdir / "src.xlsx").exists()
        (wdir / "src.xlsx").write_bytes(b"edited")
        (wdir / "extra.txt").write_text("x")
        return {"started_iso": "a", "started_monotonic": 0.0, "wall_seconds": 1.0, "finished_iso": "b",
                "dialogs_seen": [], "transcript_path": transcript, "aborted_reason": None,
                "timed_out": False, "session_ended_early": False}

    monkeypatch.setattr(rec, "_launch_and_wait", fake_launch)
    orig = rec.make_isolated_workdir
    monkeypatch.setattr(rec, "make_isolated_workdir", lambda: orig(str(tmp_path)))

    result = rec.record_cell(task, "A", 1, cdir, "claude", "disallowed", 60, "s", relaunch_backoff_seconds=0)
    cwd = seen["cwd"]
    assert not cwd.exists()  # temp dir removed
    assert cwd.name.startswith("xlsxb-")
    assert (cdir / "work" / "src.xlsx").read_bytes() == b"edited"
    assert (cdir / "work" / "extra.txt").exists()
    assert result["invocation"]["cwd"] == str(cwd) and result["invocation"]["isolated"] is True
    assert result["output_files"] == ["extra.txt"]
    assert result["contamination"]["contaminated"] is True  # the transcript lists the bench dir
    manifest = json.loads((cdir / "manifest.json").read_text())
    assert manifest["workbook"]["modified_in_place"] is True


# --- contamination detector ----------------------------------------------


def test_clean_isolated_run_is_not_contaminated():
    entries = [
        _tool_use("1", "Bash", {"command": "ls /tmp/xlsxb-abc123 && safexl inspect /tmp/xlsxb-abc123/a.xlsx"}),
        _result("1", "a.xlsx\n"),
        _tool_use("2", "Read", {"file_path": "/home/alice/project/packages/safexl/claude-plugin/skills/x/SKILL.md"}),
        _text("Done, saved /tmp/xlsxb-abc123/a.xlsx"),
    ]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-abc123")
    assert v == {"contaminated": False, "evidence": []}


@pytest.mark.parametrize("path", [
    f"{BENCH}/recordings/r/t/A/1/manifest.json",
    f"{BENCH}/runs/x/run.json",
    f"{BENCH}/results/r/grades.jsonl",
    f"{BENCH}/corpus/book.xlsx",
    f"{BENCH}/tasks_hard8.json",
    f"{BENCH}/grade.py",
])
def test_any_bench_path_is_flagged(path):
    v = runner.detect_contamination([_tool_use("1", "Read", {"file_path": path})], own_tmp="/tmp/xlsxb-abc123")
    assert v["contaminated"] and v["evidence"][0]["match"] == path
    assert v["evidence"][0]["kind"] == "tool_use" and v["evidence"][0]["tool"] == "Read"


def test_other_runs_tmp_dir_is_flagged_own_is_not():
    entries = [_tool_use("1", "Bash", {"command": "ls /tmp/xlsxb-abc123 /tmp/xlsxb-zzz999"})]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-abc123")
    assert [e["match"] for e in v["evidence"]] == ["/tmp/xlsxb-zzz999"]


def test_output_only_hits_are_warnings_not_contamination():
    entries = [_tool_use("1", "Bash", {"command": "find / -name '*.xlsx'"}),
               _result("1", f"{BENCH}/corpus/a.xlsx\n/tmp/xlsxb-zzz999/b.xlsx"),
               _text(f"see {BENCH}/corpus/a.xlsx")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is False and v["evidence"] == []
    assert {e["kind"] for e in v["warnings"]} == {"tool_result", "assistant_text"}
    assert "/tmp/xlsxb-zzz999" in {e["match"] for e in v["warnings"]}


def test_arm_b_reading_safexl_test_that_names_the_bench_is_a_warning():
    src = 'BUDGET = REPO / "research" / "xlsx-bench" / "corpus" / "simple-monthly-budget.xlsx"'
    entries = [_tool_use("1", "Read", {"file_path": "/home/alice/project/packages/safexl/tests/test_relationships.py"}),
               _result("1", f"    15→{src}\n")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is False and [e["match"] for e in v["warnings"]] == ["xlsx-bench"]


def test_grep_style_bench_content_in_output_is_contamination():
    entries = [_tool_use("1", "Bash", {"command": "grep -rn Ladakh ~"}),
               _result("1", f"{BENCH}/tasks_hard8.json:12:  \"expected\": \"Ladakh\"\n")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is True and v["evidence"][0]["kind"] == "tool_result"


def test_grep_content_after_a_prose_mention_is_still_contamination():
    path = f"{BENCH}/tasks_hard8.json"
    entries = [_tool_use("1", "Bash", {"command": "grep -rn -C1 Ladakh ~"}),
               _result("1", f"see {path}\n  {path}:12:  \"expected\"\n{path}-13-  \"x\"\n")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is True and v["evidence"][0]["kind"] == "tool_result"


def test_find_listing_with_dates_in_paths_is_only_a_warning():
    entries = [_tool_use("1", "Bash", {"command": "find / -name manifest.json"}),
               _result("1", f"{BENCH}/recordings/haiku-2026-09-29/kyc/A/1/manifest.json\n")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is False and len(v["warnings"]) == 1


def test_same_path_in_input_and_output_is_contamination():
    path = f"{BENCH}/corpus/a.xlsx"
    entries = [_tool_use("1", "Read", {"file_path": path}), _result("1", f"read {path}")]
    v = runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")
    assert v["contaminated"] is True and [e["kind"] for e in v["evidence"]] == ["tool_use"]


def test_arm_b_env_file_mention_is_ignored():
    entries = [_result("1", f"CLAUDE_ENV_FILE={BENCH}/.arm_b_env.sh")]
    assert not runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")["contaminated"]


def test_legacy_layout_allows_own_cell_and_flags_siblings():
    own = RUN_DIR.replace("/A/1", "/A/1")
    ok = [
        _tool_use("1", "Bash", {"command": f"ls {own}/work"}),
        _result("1", f"/home/alice/.claude/projects/{own.replace('/', '-').replace('.', '-')}-work/s/tool-results/x.txt"),
        _result("2", f"research/xlsx-bench/recordings/round/task/A/1/work/a.xlsx"),
    ]
    assert not runner.detect_contamination(ok, own_dirs=[own])["contaminated"]
    for sib in (f"{BENCH}/recordings/round/task/B/1/manifest.json", f"{BENCH}/recordings/round/task"):
        v = runner.detect_contamination([_tool_use("1", "Bash", {"command": f"cat {sib}"})], own_dirs=[own])
        assert v["contaminated"], sib
    v = runner.detect_contamination([_tool_use("1", "Bash", {"command": "cat ../../B/1/manifest.json"})], own_dirs=[own])
    assert v["contaminated"]


def test_evidence_is_capped_and_deduped():
    entries = [_tool_use(str(i), "Read", {"file_path": f"{BENCH}/corpus/f{i}.xlsx"}) for i in range(50)]
    entries += [_tool_use("d", "Read", {"file_path": f"{BENCH}/corpus/f1.xlsx"})]
    assert len(runner.detect_contamination(entries, own_tmp="/tmp/xlsxb-a")["evidence"]) == 20


# --- rate-limit detector --------------------------------------------------


def test_rate_limit_false_positive_from_run_json_text():
    text = '  "rate_limited": false,\n  "rate_limit_evidence": null,\n  "note": "usage limit try again later"'
    entries = [_tool_use("1", "Read", {"file_path": "/x/run.json"}), _result("1", text), _text("rate limit docs say so")]
    assert runner.detect_rate_limit(entries, "") is None


def test_rate_limit_real_synthetic_api_error_message():
    entry = {"type": "assistant", "isApiErrorMessage": True, "error": "rate_limit",
             "message": {"model": "<synthetic>", "content": [{"type": "text", "text": "You've hit your session limit · resets 7am"}]}}
    assert "session limit" in runner.detect_rate_limit([entry], "").lower()


def test_rate_limit_quota_rejected_and_stream_json_events():
    q = {"type": "assistant", "quotaLimits": {"status": "rejected"},
         "message": {"model": "m", "content": [{"type": "text", "text": "Usage limit reached"}]}}
    assert runner.detect_rate_limit([q], "") is not None
    ev = {"type": "rate_limit_event", "rate_limit_info": {"status": "rejected", "rateLimitType": "five_hour"}}
    assert runner.detect_rate_limit([ev], "") is not None
    ok = {"type": "rate_limit_event", "rate_limit_info": {"status": "allowed"}}
    assert runner.detect_rate_limit([ok], "") is None
    assert runner.detect_rate_limit([{"type": "result", "subtype": "success", "is_error": False,
                                      "result": "explained the usage limit docs"}], "") is None


# --- scan + grade option --------------------------------------------------


def _make_cell(root, task, arm, rep, entries, contaminated=None):
    cdir = root / task / arm / str(rep)
    (cdir / "work").mkdir(parents=True)
    (cdir / "session.jsonl").write_text("\n".join(json.dumps(e) for e in entries) + "\n")
    meta = {"rate_limited": False}
    (cdir / "run.json").write_text(json.dumps(meta))
    return cdir


def test_scan_cell_and_write(tmp_path):
    bad = _make_cell(tmp_path, "t", "B", 1, [_tool_use("1", "Bash", {"command": f"cat {BENCH}/recordings/o/t/A/1/run.json"})])
    good = _make_cell(tmp_path, "t", "A", 1, [_tool_use("1", "Bash", {"command": "ls"})])
    assert scan.scan_cell(bad)["contamination"]["contaminated"] is True
    assert scan.scan_cell(good)["contamination"]["contaminated"] is False
    assert scan.main([str(tmp_path), "--write"]) == 0
    assert json.loads((bad / "run.json").read_text())["contamination"]["contaminated"] is True
    assert json.loads((good / "run.json").read_text())["contamination"] == {"contaminated": False, "evidence": []}


def test_grade_run_contaminated_flag(tmp_path):
    d = tmp_path / "c"
    d.mkdir()
    assert grade._run_contaminated(d) is False  # no run.json
    (d / "run.json").write_text(json.dumps({}))
    assert grade._run_contaminated(d) is False
    (d / "run.json").write_text(json.dumps({"contamination": {"contaminated": True, "evidence": []}}))
    assert grade._run_contaminated(d) is True
