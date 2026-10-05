#!/usr/bin/env python3
"""Build results/agnostic-2026-10-04/{table.json,SUMMARY.md} and the outbox report from run.json + grade files.
Usage: summarize.py STAGE   (3 or 4). Prices: per-million-token USD (in, out) with source, looked up 2026-10-04."""
import json, os, statistics as st, sys, collections, subprocess
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import harness as H
STAGE = sys.argv[1] if len(sys.argv) > 1 else "3"
PRICES = {  # model: (in, out, source)
 "openai/gpt-oss-120b": (0.15, 0.60, "Groq model docs console.groq.com/docs/models, 2026-10-04"),
 "openai/gpt-oss-20b": (0.075, 0.30, "Groq model docs console.groq.com/docs/models, 2026-10-04"),
 "qwen/qwen3.8-27b": (0.80, 4.00, "Groq model docs console.groq.com/docs/models, 2026-10-04 (OpenRouter paid lists 0.42/3.00)"),
 "command-a-plus-05-2026": (0.30, 1.50, "OpenRouter API /models cohere/command-a-plus, 2026-10-04 (no Cohere pricing page entry)"),
 "command-a-03-2025": (2.50, 10.00, "OpenRouter API /models cohere/command-a and Cohere list price, 2026-10-04"),
 "command-r-plus-08-2024": (2.50, 10.00, "cohere.com/pricing, 2026-10-04"),
 "poolside/laguna-s-2.1:free": (0.09, 0.18, "OpenRouter paid variant poolside/laguna-s-2.1 (API /models, 2026-10-04); tested id is free tier"),
 "nvidia/nemotron-3-super-120b-a12b:free": (0.08, 0.45, "OpenRouter paid variant nvidia/nemotron-3-super-120b-a12b (API /models, 2026-10-04); tested id is free tier"),
 "qwen/qwen3.8-27b:free": (0.42, 3.00, "OpenRouter paid variant qwen/qwen3.8-27b (API /models, 2026-10-04); tested id is free tier"),
 "ministral-8b-latest": (0.10, 0.10, "third-party aggregator pricing pages (pricepertoken.com, anotherwrapper.com), 2026-10-04; not confirmed on mistral.ai; Mistral key here is Experiment tier (free)"),
 "open-mistral-nemo": (0.15, 0.15, "third-party aggregator pricing pages, 2026-10-04; not confirmed on mistral.ai; Experiment tier (free) here"),
}
BASE = {"haiku": {"pass": "21/24", "per_pass_usd": 0.0875}, "opus": {"pass": "8/8", "per_pass_usd": 0.76}}
def med(x): return round(st.median(x), 1) if x else None
FREE_REC = H.BENCH / "recordings" / "agnostic-2026-10-04"; PAID_REC = FREE_REC / "paid"
def grade_rows(slug, tag=""):
    f = H.RES / f"grade_{tag}{slug}.jsonl"
    if not f.exists(): return {}
    return {(r["task_id"], r.get("rep")): r for r in (json.loads(l) for l in open(f) if l.strip())}
models, per_task, fails, fail_lines = [], {}, {}, {}
order = [(a, *l.split(None, 1)) for a, f in (("free gateway", "order.txt"), ("OpenRouter paid", "order_paid.txt")) for l in (Path(__file__).parent / f).read_text().splitlines() if l.strip()]
for access, prov, mid in order:
    paid = access.endswith("paid")
    root = (PAID_REC if paid else FREE_REC) / H.slug(mid)
    runs = [json.loads(p.read_text()) for p in sorted(root.glob("*/B/*/run.json"))] if root.exists() else []
    if not runs: continue
    gf = H.RES / f"grade_{'paid_' if paid else ''}{H.slug(mid)}.jsonl"
    newest = max(p.stat().st_mtime for p in root.glob("*/B/*/run.json"))
    if not gf.exists() or gf.stat().st_mtime < newest or sum(1 for _ in open(gf)) != len(runs):
        subprocess.run([sys.executable, str(H.BENCH / "grade.py"), "--tasks", str(H.BENCH / "tasks_hard8.json"), "--runs", str(root), "--out", str(gf), "--summary"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    g = grade_rows(H.slug(mid), "paid_" if paid else "")
    gr = [g.get((r["task_id"], str(r["rep"]))) or g.get((r["task_id"], r["rep"])) for r in runs]
    graded = [x for x in gr if x]
    n = len(runs); ps = sum(1 for x in graded if x.get("pass")); it = sum(1 for x in graded if x.get("preserved"))
    stops = collections.Counter(r["stop_reason"] for r in runs)
    ti = [r["tokens_in"] for r in runs]; to = [r["tokens_out"] for r in runs]
    pr = PRICES.get(mid)
    cost = [r.get("cost_usd", 0.0) for r in runs] if paid else ([(a * pr[0] + b * pr[1]) / 1e6 for a, b in zip(ti, to)] if pr else None)
    tot = sum(cost) if cost else None
    models.append({"model": mid, "provider": prov, "access": access, "provider_served": sorted({p for r in runs for p in r.get("providers", [])}), "cost_basis": "measured (OpenRouter usage.cost)" if paid else "list-price arithmetic (run was free)", "n": n, "graded": len(graded), "pass": ps, "intact": it,
        "median_wall_s": med([r["wall_seconds"] for r in runs]), "median_tokens_in": med(ti), "median_tokens_out": med(to),
        "list_price_per_run_usd": round(tot / n, 5) if tot is not None else None,
        "list_price_per_pass_usd": round(tot / ps, 4) if tot is not None and ps else None,
        "price_source": ("OpenRouter measured usage.cost per request" if paid else pr[2]) if (paid or pr) else "free tier only", "price_in_out_per_M": pr[:2] if (pr and not paid) else None, "stop_reasons": dict(stops)})
    kinds = collections.Counter(); lazy = 0; errs = 0; calls = 0
    for r, x in zip(runs, gr):
        for t in r["tool_calls"]:
            calls += 1
            if t["error"]: errs += 1; kinds[t.get("kind") or "tool_error"] += 1
        if r["stop_reason"] == "finished" and r["tool_call_count"] == 0: kinds["answered_without_any_tool_call"] += 1
        wr = [t for t in r["tool_calls"] if t["name"] in ("safexl_edit", "safexl_insert_rows", "safexl_insert_cells", "safexl_rename_sheet", "safexl_add_sheet", "safexl_resize_table") and not t["error"]]
        if r["stop_reason"] == "finished" and not wr: kinds["finished_with_no_successful_write"] += 1
        if x and not x.get("pass") and wr: kinds["wrote_but_wrong_result"] += 1
        if x and not x.get("pass") and r["stop_reason"] == "api_failure": kinds["rate_limited_or_api_failure"] += 0
    key_ = mid + (" (paid)" if paid else "")
    models[-1]["empty_reply_runs"] = sum(1 for r in runs if r["stop_reason"] == "empty_reply")
    models[-1]["wrong_answer_runs"] = sum(1 for r, x in zip(runs, gr) if x and not x.get("pass") and r["stop_reason"] != "empty_reply")
    fl = []
    for r, x in zip(runs, gr):
        if x and x.get("pass"): continue
        bad = [c for c in (x or {}).get("checks", []) if not c.get("ok")]
        why = "; ".join(f"{c['name']}: {str(c.get('detail'))[:110]}" for c in bad[:2]) or ("not graded" if not x else "no failed check listed")
        last = r["tool_calls"][-1] if r["tool_calls"] else None
        fl.append(f"{r['task_id']}#{r['rep']} [{r['stop_reason']}, {r['tool_call_count']} calls, tool errors {r['tool_error_count']}]: grader: {why}. Final text: {(r.get('final_text') or '(empty)')[:140]!r}")
    fail_lines[key_] = fl
    fails[key_] = {"tool_calls": calls, "tool_errors": errs, "kinds": dict(kinds)}
    for r, x in zip(runs, gr):
        per_task.setdefault(r["task_id"], {})[f"{mid}{' (paid)' if paid else ''}#{r['rep']}"] = ("PASS" if x and x.get("pass") else ("fail/intact" if x and x.get("preserved") else "fail")) + f" ({r['stop_reason']}, {r['tool_call_count']} calls)"
J = {"task": "agnostic-eval", "stage": STAGE, "status": "done"}
probe = json.load(open(H.RES / "probe.json"))
J["probe"] = {"ok": [f"{p['provider']} {p['model']}" for p in probe if p["ok"]], "failed": [f"{p['provider']} {p['model']}: {p['detail'][:60]}" for p in probe if not p["ok"]]}
J["models"] = models; J["baselines"] = BASE
good = [m for m in models if m["pass"] >= 0.75 * m["n"] * 21 / 24 and m["graded"]]
J["verdict"] = ("Models within reach of Haiku's 21/24 pass rate: " + (", ".join(f"{m['model']} ({m['pass']}/{m['n']})" for m in good) or "none") + ".")
J["limits"] = ["free-tier rate limits distort wall time; the paid OpenRouter arm is the real speed number (wall times reported per arm, not mixed)", "paid arm capped at $1.50 spend; Haiku control ran in this tools-only harness, not Claude Code", "n is small (stage 3: 1 rep x 8 tasks + smoke cell; stage 4: 3 reps)", "no openpyxl baseline arm run for these models", "tool surface is MCP only", "list-price cost is computed from tokens at public per-token prices; the runs themselves were free", "11 of 22 models unavailable on the gateway at probe time"]
J["blockers"] = []; J["notes"] = "Invalid first attempt (isError bridge bug) excluded; see README."; J["contract_issues"] = []
(H.RES / "table.json").write_text(json.dumps(J, indent=1))
L = ["# SafeXL on non-Claude models (agnostic eval) - interim" if STAGE == "3" else "# SafeXL on non-Claude models (agnostic eval) - final", "",
     f"Stage {STAGE}. Tools: SafeXL MCP only (11 tools). Grader: grade.py, arm B. Pass/intact are out of n runs graded.", "",
     "| model | access | pass | intact | stop reasons | med wall s | med tok in/out | $/run | $/pass | cost basis |", "|---|---|---|---|---|---|---|---|---|---|"]
for m in models:
    L.append(f"| {m['model']} | {m['access']} | {m['pass']}/{m['n']} | {m['intact']}/{m['n']} | {m['stop_reasons']} | {m['median_wall_s']} | {m['median_tokens_in']:.0f}/{m['median_tokens_out']:.0f} | {m['list_price_per_run_usd'] if m['list_price_per_run_usd'] is not None else 'free tier only'} | {m['list_price_per_pass_usd'] if m['list_price_per_pass_usd'] is not None else 'n/a'} | {m['cost_basis']} |")
L += ["| Haiku 4.5 + SafeXL plugin (baseline) | Anthropic | 21/24 | 23/24 | - | ~60 | - | 0.0766 | 0.0875 | measured |", "| Opus 5.5 (baseline) | Anthropic | 8/8 | - | - | - | - | - | 0.76 | measured |", "",
      "## Price sources (per million tokens in/out)"]
for m in models: L.append(f"- {m['model']} ({m['access']}): {m['price_in_out_per_M']} - {m['price_source']}" + (f"; served by {m['provider_served']}" if m["provider_served"] else ""))
L.append("- Groq models (gpt-oss-120b, gpt-oss-20b, qwen3.8-27b on the free gateway): not measurable on the free tier: one request exceeds the per-minute token limit (413, TPM 8000 vs ~8280 requested); runs kept under recordings/.../_blocked-free-tier-tpm/. Measured on the paid OpenRouter arm instead where listed.")
L += ["", "## Per-task results"]
for t, d in per_task.items():
    L.append(f"- **{t}**: " + "; ".join(f"{k}: {v}" for k, v in d.items()))
L += ["", "## Failure modes (from run records; tool errors counted over all tool calls)"]
for mid, f in fails.items(): L.append(f"- {mid}: {f['tool_errors']}/{f['tool_calls']} tool calls errored; {f['kinds']}")
top = sorted([m for m in models if m["access"].endswith("paid") and not m["model"].startswith("anthropic")], key=lambda m: -m["pass"])[:3]
L += ["", "## Why the failed runs of the top three paid models failed (grader check + transcript)"]
for m in top:
    L.append(f"### {m['model']} (paid), provider served: {m['provider_served']}")
    L += [f"- {x}" for x in fail_lines.get(m["model"] + " (paid)", [])]
L += ["", "## Empty replies vs wrong answers"] + [f"- {m['model']} ({m['access']}): {m['empty_reply_runs']} empty-reply runs, {m['wrong_answer_runs']} wrong/other failures, {m['pass']} passes of {m['n']}" for m in models]
L += ["", "## Verdict", J["verdict"], "", "## Limits"] + [f"- {x}" for x in J["limits"]]
(H.RES / "SUMMARY.md").write_text("\n".join(L) + "\n")
# optional extra copy of the verdict JSON: written only if AGN_SUMMARY_OUT names a file
if os.environ.get("AGN_SUMMARY_OUT"):
    out = Path(os.environ["AGN_SUMMARY_OUT"])
    out.parent.mkdir(parents=True, exist_ok=True); out.write_text(json.dumps(J))
print("\n".join(L))
