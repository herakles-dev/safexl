#!/usr/bin/env python3
"""Regenerate table.json and SUMMARY.md for the agnostic eval.

Reads grade_*.jsonl here and the run.json records under
../../recordings/agnostic-2026-10-04/. Run from anywhere:
    python3 make_summary.py
"""
import glob
import json
import os
import statistics
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
REC = os.path.normpath(os.path.join(HERE, "..", "..", "recordings", "agnostic-2026-10-04"))
FULL_N = 21  # 7 unaffected hard tasks x 3 reps
ALL_N = 24  # all 8 hard tasks x 3 reps
# This task's own rows and formula were in the system prompt every model received, so its runs are not evidence.
EXCLUDED_TASK = "benefits-t4-insert-duplicate-exclusion"
EXCLUDED_REASON = "the system prompt held this task's own rows and total formula as a usage example"

# display order: paid first, then free tier; values are filled from data
def parse_name(fn):
    base = os.path.basename(fn)[len("grade_"):-len(".jsonl")]
    if base.startswith("paid_"):
        return base[5:], "paid", os.path.join(REC, "paid", base[5:])
    free = base.endswith("_free")
    return base, ("free" if free else "free-gateway"), os.path.join(REC, base)


def med(xs):
    return round(statistics.median(xs), 1) if xs else None


rows = []
for fn in sorted(glob.glob(os.path.join(HERE, "grade_*.jsonl"))):
    key, tier, rdir = parse_name(fn)
    grades = [json.loads(l) for l in open(fn) if l.strip()]
    all_n = len(grades)
    all_pass = sum(g["pass"] for g in grades)
    all_int = sum(g["preserved"] for g in grades)
    ex = [g for g in grades if g["task_id"] == EXCLUDED_TASK]
    all_cost = []
    for g in grades:
        p = os.path.join(rdir, g["task_id"], g["arm"], str(g["rep"]), "run.json")
        if os.path.exists(p):
            c = json.load(open(p)).get("cost_usd")
            if c is not None:
                all_cost.append(c)
    grades = [g for g in grades if g["task_id"] != EXCLUDED_TASK]
    runs, missing = [], []
    for g in grades:
        p = os.path.join(rdir, g["task_id"], g["arm"], str(g["rep"]), "run.json")
        if os.path.exists(p):
            runs.append((g, json.load(open(p))))
        else:
            missing.append(f'{g["task_id"]}#{g["rep"]}')
    n = len(grades)
    npass = sum(g["pass"] for g in grades)
    nint = sum(g["preserved"] for g in grades)
    nvalid = sum(g["valid"] for g in grades)
    ncorr = sum(g["correct"] for g in grades)
    stops = Counter(r["stop_reason"] for _, r in runs)
    walls = [r["wall_seconds"] for _, r in runs if r.get("wall_seconds") is not None]
    costs = [r["cost_usd"] for _, r in runs if r.get("cost_usd") is not None]
    calls = sum(r.get("tool_call_count", 0) for _, r in runs)
    errs = sum(r.get("tool_error_count", 0) for _, r in runs)
    # reps per task actually graded
    per_task = Counter(g["task_id"] for g in grades)
    row = {
        "model": next((r["model"] for _, r in runs), key),
        "tier": tier,
        "n_graded": n,
        "partial": n < FULL_N,
        "pass": npass,
        "intact": nint,
        "correct": ncorr,
        "valid": nvalid,
        "median_wall_s": med(walls),
        "cost_runs_recorded": len(costs),
        "cost_per_run_usd": round(sum(costs) / len(costs), 5) if costs else None,
        "cost_total_usd": round(sum(costs), 4) if costs else None,
        "stop_reasons": dict(stops),
        "tool_calls": calls,
        "tool_errors": errs,
        "tasks_covered": len(per_task),
        "reps_per_task": dict(sorted(per_task.items())),
        "runs_without_run_json": missing,
        "failed_tasks": sorted(f'{g["task_id"]}#{g["rep"]}' for g in grades if not g["pass"]),
        "all8": {"n_graded": all_n, "pass": all_pass, "intact": all_int},
        "excluded_task": {"task_id": EXCLUDED_TASK, "n": len(ex), "pass": sum(g["pass"] for g in ex)},
    }
    row["_all_cost"] = sum(all_cost)
    rows.append(row)

order = {"paid": 0, "free-gateway": 1, "free": 2}
rows.sort(key=lambda r: (r["partial"], -r["pass"] / r["n_graded"]))

# spend: the provider's usage meter at the end of the eval minus its reading at the start
log_spent = json.load(open(os.path.join(HERE, "paid_spend_total.json")))["openrouter_spend_usd"]
paid_cost_sum = 0.0
paid_runs = 0
for p in glob.glob(os.path.join(REC, "paid", "**", "run.json"), recursive=True):
    d = json.load(open(p))
    if d.get("cost_usd") is not None:
        paid_cost_sum += d["cost_usd"]
        paid_runs += 1
graded_cost = sum(r.pop("_all_cost") for r in rows)
table = {
    "task_set": "tasks_hard8.json minus one excluded task (7 tasks, up to 3 reps)",
    "full_n": FULL_N,
    "excluded_task": {"task_id": EXCLUDED_TASK, "reason": EXCLUDED_REASON},
    "definitions": {
        "pass": "grade.py: correct and preserved and valid",
        "intact": "grade.py 'preserved': no losses against the original workbook",
    },
    "models": rows,
    "spend": {
        "paid_spend_log_final_usd": log_spent,
        "paid_run_json_cost_sum_usd": round(paid_cost_sum, 4),
        "paid_run_json_count": paid_runs,
        "graded_runs_cost_sum_usd": round(graded_cost, 4),
        "free_tier_runs_billed": 0,
    },
}
with open(os.path.join(HERE, "table.json"), "w") as f:
    # same call the public export uses when it re-dumps .json files
    f.write(json.dumps(table, indent=2, ensure_ascii=False) + "\n")

# blocked: generated once from the private recordings, kept as data so no _ directory is needed
blocked = json.load(open(os.path.join(HERE, "blocked_runs.json")))


def fmt(x, nd=1):
    return "-" if x is None else f"{x:.{nd}f}"


L = []
L.append("# Non-Claude models through SafeXL's MCP tools only: hard-8 results, 7 tasks counted\n")
L.append("Date of runs: 2026-10-04. Each model was given only SafeXL's 11 MCP tools and one task prompt. "
         "There is no arm without SafeXL in this eval, so it says nothing about what these models do without it.\n")
L.append("Tasks: 7 of the 8 hard tasks in `research/xlsx-bench/tasks_hard8.json`, up to 3 reps each (21 runs is a full set). "
         "One task is left out of every headline number; see \"Why one task is left out\" below. "
         "n is the number of runs graded. Any model under 21 is partial and is not comparable to a 21-run row.\n")
L.append("## Definitions (from `grade.py`)\n")
L.append("A run counts as **pass** when the grader finds every expected cell, formula or sheet change correct, finds no new error cells, can open and recalculate the file, and finds nothing lost from the original workbook. A run counts as **intact** when the grader's `preserved` check finds no losses against the original workbook (formatting, validations, named ranges, extensions and similar). A run can be intact and still fail the task.\n")
L.append("## Results (7 tasks, excluded task left out)\n")
L.append("| model | tier | n | pass | intact | median wall s | $ per run | note |")
L.append("|---|---|---|---|---|---|---|---|")
for r in rows:
    notes = []
    if r["partial"]:
        notes.append(f'partial ({r["n_graded"]} of {FULL_N}; {r["tasks_covered"]} tasks)')
    if r["stop_reasons"].get("empty_reply"):
        notes.append(f'{r["stop_reasons"]["empty_reply"]} empty replies')
    if r["stop_reasons"].get("api_failure"):
        notes.append(f'{r["stop_reasons"]["api_failure"]} api failure(s)')
    if r["stop_reasons"].get("cap_tool_calls"):
        notes.append(f'{r["stop_reasons"]["cap_tool_calls"]} hit the tool-call cap')
    if r["tool_errors"]:
        notes.append(f'{r["tool_errors"]} of {r["tool_calls"]} tool calls errored')
    if r["cost_per_run_usd"] is None:
        cost = "not billed" if r["tier"] == "free" else "none recorded"
    else:
        cost = f'{r["cost_per_run_usd"]:.4f}'
    L.append(f'| {r["model"]} | {r["tier"]} | {r["n_graded"]} | {r["pass"]} | {r["intact"]} | '
             f'{fmt(r["median_wall_s"])} | {cost} | {"; ".join(notes) or "-"} |')
L.append("")
L.append("Tier: `paid` = OpenRouter, billed per request, cost is the measured `usage.cost`. "
         "`free` = OpenRouter free-tier id. `free-gateway` = free gateway, no cost recorded. "
         "Median wall time is per run, from the run record.\n")
EX = EXCLUDED_TASK
L.append("## All 8 tasks, including the affected one (for reference; do not quote)\n")
L.append(f"| model | n | pass | intact | `{EX}` alone |")
L.append("|---|---|---|---|---|")
for r in rows:
    e = r["excluded_task"]
    a8 = r["all8"]
    L.append(f'| {r["model"]} | {a8["n_graded"]} | {a8["pass"]} | {a8["intact"]} | {e["pass"]} of {e["n"]} |')
L.append("")
L.append("## Why one task is left out\n")
L.append("The system prompt every model received held two lines as a usage example:\n")
L.append("```\n"
         "- A total just above an inserted row does not grow by itself. =SUM(C25:C28) with a row inserted at 29 needs the formula =SUM(C25:C29) written to the total's new cell.\n"
         "- To show figures that already exist elsewhere in the workbook, write formulas that point at those cells (='Other Sheet'!C30), not copied values.\n"
         "```\n")
full = [r for r in rows if not r["partial"]]
part = [r for r in rows if r["partial"]]
def tot(rs):
    return sum(r["excluded_task"]["pass"] for r in rs), sum(r["excluded_task"]["n"] for r in rs)
L.append(f"These are the rows and the total of task `{EX}`. "
         "The runs of that task are excluded from every headline number above, because the prompt held what the task asks for. "
         f"On that task the {len(full)} models with a full set passed {tot(full)[0]} of {tot(full)[1]} runs. "
         f"The {len(part)} partial models passed {tot(part)[0]} of {tot(part)[1]}. "
         "The per-model figures are in the reference table above.\n")
L.append("The recordings do not hold the system prompt. Its full text as run is in `system_prompt_as_run.txt` in this folder. "
         "I took it from the harness source as it stood during the round. The leak scan finds only the two lines above in it.\n")
L.append("The rest of the prompt's guidance came from SafeXL's skill text, which was written after watching models fail on these same tasks. "
         "So this round is in-sample for that guidance and is not a held-out test.\n")
s = table["spend"]
L.append("## Spend (whole round, all 8 tasks)\n")
L.append(f"- OpenRouter spend during the eval, from the spend log (usage minus the baseline taken before the paid runs): ${s['paid_spend_log_final_usd']:.4f}.\n"
         f"- Sum of per-run measured cost over all {s['paid_run_json_count']} paid-tier run records: ${s['paid_run_json_cost_sum_usd']:.4f}.\n"
         f"- Sum over all graded paid runs, including the excluded task: ${s['graded_runs_cost_sum_usd']:.4f}.\n"
         "- Free-tier runs were not billed.\n")
L.append("## Models that could not be run\n")
L.append("Three models were tried on the Groq free tier and could not be measured there. "
         f"The run records show Groq rejecting requests with a limit of {blocked['groq_tpm_limit']:,} tokens per minute against "
         f"requests of {blocked['rejected_request_tokens'][0]:,} and {blocked['rejected_request_tokens'][1]:,} tokens, and gateway 503s when its keys were in backoff. "
         "Their runs are not graded here:\n")
for m in blocked["models"]:
    L.append(f"- {m['model']}: {m['runs_recorded']} runs recorded, stop reasons {m['stop_reasons']}.")
L.append("\ngpt-oss-120b and gpt-oss-20b were then run on OpenRouter's paid tier, and Qwen 3.8 27B on OpenRouter's free tier. "
         "Those runs are the rows in the tables above.\n")
L.append("## Per-model failures (7 tasks)\n")
for r in rows:
    L.append(f'- {r["model"]} ({r["n_graded"]} runs): failed {len(r["failed_tasks"])}: {", ".join(r["failed_tasks"]) or "none"}')
L.append("")
L.append("Regenerate with `python3 make_summary.py` in this directory. `table.json` holds the 7-task numbers as the main fields, plus an `all8` object, the excluded task's own result and per-task rep counts.\n")
open(os.path.join(HERE, "SUMMARY.md"), "w").write("\n".join(L))
for r in rows:
    print(r["model"], r["tier"], r["n_graded"], r["pass"], r["intact"], r["median_wall_s"], r["cost_per_run_usd"], r["stop_reasons"], r["tool_errors"], r["tool_calls"], r["runs_without_run_json"], r["tasks_covered"])
print(table["spend"])
