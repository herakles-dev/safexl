#!/usr/bin/env python3
"""Retro-scan finished recordings for clean-room contamination.

    python3 scan_contamination.py recordings/<round> [...]  [--write]

For every <task>/<arm>/<rep>/ with a session.jsonl, run
runner.detect_contamination over the transcript and print one line per run.
--write stores the verdict in that cell's run.json under "contamination".

Runs recorded in the old in-repo layout (cwd = <cell>/work) may mention their
own cell dir; anything else under research/xlsx-bench/ is flagged. Runs that
recorded an isolated cwd (run.json invocation.isolated) are scanned with
own_tmp instead. Also re-runs the rate-limit detector and reports any run
where the stored flag disagrees (never rewrites it).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import record  # noqa: E402
import runner  # noqa: E402


def scan_cell(cdir: Path) -> dict | None:
    jsonl = cdir / "session.jsonl"
    run_json = cdir / "run.json"
    if not jsonl.exists() or not run_json.exists():
        return None
    entries = record.read_transcript_entries(jsonl)
    meta = json.loads(run_json.read_text())
    inv = meta.get("invocation") or {}
    if inv.get("isolated"):
        verdict = runner.detect_contamination(entries, own_tmp=inv.get("cwd"))
    else:
        verdict = runner.detect_contamination(entries, own_dirs=[str(cdir.resolve())])
    return {
        "meta": meta,
        "contamination": verdict,
        "rate_limit_now": runner.detect_rate_limit(entries, ""),
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("rounds", nargs="+", type=Path)
    ap.add_argument("--write", action="store_true", help="store the verdict in each run.json")
    args = ap.parse_args(argv)

    n = dirty = warned = 0
    for round_dir in args.rounds:
        for task_dir in sorted(p for p in round_dir.iterdir() if p.is_dir() and not p.name.startswith("_")):
            for arm_dir in sorted(p for p in task_dir.iterdir() if p.is_dir()):
                for cdir in sorted(p for p in arm_dir.iterdir() if p.is_dir() and p.name.isdigit()):
                    r = scan_cell(cdir)
                    if r is None:
                        continue
                    n += 1
                    c = r["contamination"]
                    label = f"{round_dir.name}/{task_dir.name}/{arm_dir.name}/{cdir.name}"
                    print(f"{'CONTAMINATED' if c['contaminated'] else 'clean       '}  {label}")
                    for ev in c["evidence"]:
                        print(f"      {ev['kind']}/{ev['tool']}: {ev['match']}")
                    for ev in c.get("warnings", []):
                        print(f"      warning {ev['kind']}/{ev['tool']}: {ev['match']}")
                    warned += bool(c.get("warnings"))
                    if bool(r["meta"].get("rate_limited")) != (r["rate_limit_now"] is not None):
                        print(f"      note: rate_limited stored={r['meta'].get('rate_limited')} "
                              f"redetected={r['rate_limit_now'] is not None}")
                    dirty += c["contaminated"]
                    if args.write:
                        meta = r["meta"]
                        meta["contamination"] = c
                        runner.write_run_json(cdir, meta)
    print(f"\n{n} runs scanned, {dirty} contaminated, {warned} with output-hit warnings")
    return 0


if __name__ == "__main__":
    sys.exit(main())
