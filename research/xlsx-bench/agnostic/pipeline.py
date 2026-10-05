#!/usr/bin/env python3
"""Stage 2+3 driver: per model (fastest probe latency first) run smoke, then hard8 x1 if the smoke made a valid tool loop."""
import asyncio, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import harness as H
reps = [int(x) for x in (sys.argv[1] if len(sys.argv) > 1 else "1").split(",")]
models = [sys.argv[2].split(None, 1)]  # one worker per model; parallelism is across processes (xargs -P)
for prov, mid in models:
    valid = False
    for rep in reps:
        tasks = H.load_tasks("smoke") if rep == reps[0] and len(reps) == 1 and not (H.REC / H.slug(mid)).exists() else None
        if tasks:
            r = asyncio.run(H.run_task(mid, prov, tasks[0], 1))
            print(json.dumps({"model": mid, "task": tasks[0]["id"], "stop": r["stop_reason"], "calls": r.get("tool_call_count"), "s": r.get("wall_seconds")}), flush=True)
            if r.get("daily_cap"): print("DAILY CAP"); sys.exit(3)
            if not any(not t["error"] for t in r["tool_calls"]) and r["stop_reason"] in ("api_failure", "empty_reply", "harness_error") :
                print(json.dumps({"model": mid, "skip": "no valid tool loop in smoke"}), flush=True); break
        for t in H.load_tasks("hard8"):
            r = asyncio.run(H.run_task(mid, prov, t, rep))
            print(json.dumps({"model": mid, "task": t["id"], "rep": rep, "stop": r["stop_reason"], "calls": r.get("tool_call_count"), "errs": r.get("tool_error_count"), "s": r.get("wall_seconds")}), flush=True)
            if r.get("daily_cap"): print("DAILY CAP"); sys.exit(3)

import subprocess
for prov, mid in models:
    root = H.REC / H.slug(mid)
    if root.exists():
        H.RES.mkdir(parents=True, exist_ok=True)
        subprocess.run([sys.executable.replace("python", "python"), str(H.BENCH / "grade.py"), "--tasks", str(H.BENCH / "tasks_hard8.json"),
                        "--runs", str(root), "--out", str(H.RES / f"grade_{H.GTAG}{H.slug(mid)}.jsonl"), "--summary"], stdout=open(H.RES / f"grade_{H.GTAG}{H.slug(mid)}.txt", "w"), stderr=subprocess.STDOUT)
        print(json.dumps({"model": mid, "graded": True}), flush=True)
