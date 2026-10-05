#!/usr/bin/env python3
"""Model-agnostic SafeXL eval harness.

Runs a non-Claude model (via an OpenAI-compatible gateway) against the SafeXL MCP server
(stdio) on xlsx-bench tasks. The model has only the 11 SafeXL MCP tools. Output layout matches
what grade.py expects: <root>/<task>/B/<rep>/{work/<workbook>, run.json}.

Usage (run with packages/safexl/.venv/bin/python):
  harness.py probe  --models-file M --out probe.json
  harness.py run    --model ID --provider P --tasks hard8|smoke --reps 1,2,3
  harness.py grade  --model ID
"""
import argparse, asyncio, json, os, re, shutil, sys, time, urllib.request, urllib.error
from pathlib import Path

BENCH = Path(__file__).resolve().parent.parent
REPO = BENCH.parent.parent
BACKEND = os.environ.get("AGN_BACKEND", "gateway")  # gateway (a local OpenAI-compatible server) | openrouter (paid, capped)
PAID = BACKEND == "openrouter"
REC = BENCH / "recordings" / "agnostic-2026-10-04" / ("paid" if PAID else "")
RES = BENCH / "results" / "agnostic-2026-10-04"
GTAG = "paid_" if PAID else ""
OR_URL = "https://openrouter.ai/api/v1/chat/completions"
OR_CREDITS = "https://openrouter.ai/api/v1/credits"
BUDGET_USD = 1.50
BASELINE_FILE = RES / "paid_baseline_usage.json"
MCP_BIN = os.environ.get("AGN_MCP_BIN", str(REPO / "packages" / "safexl" / ".venv" / "bin" / "safexl-mcp"))
GATEWAY = "http://localhost:8132/v1/chat/completions"
MAX_TOOL_CALLS = 40
MAX_WALL = 12 * 60
REQ_TIMEOUT = 120 if os.environ.get("AGN_BACKEND") == "openrouter" else 60
API_RETRIES = 3
TOOL_RESULT_CAP = 30000

# The 2026-10-04 round ran with an earlier wording of two "usage slips" lines, whose examples were one
# task's own rows and formula. That task's runs are left out of the round's published results; the earlier
# wording is quoted in results/agnostic-2026-10-04/SUMMARY.md. The text below is the corrected prompt.
SYSTEM_PROMPT = """# SafeXL: edit an existing workbook without wrecking it

You edit Excel workbooks (.xlsx/.xlsm) using ONLY the safexl_* tools you have been given. You have no shell, no
Python and no other way to read or write files. Saving a workbook with openpyxl or pandas rebuilds the whole file
and silently drops charts, pivot tables, macros, sparklines, data validation and defined names; the safexl tools
edit the file in place so everything you do not touch stays byte-identical.

The workbook to edit is: {work_path}
A pristine copy of the original, for comparison only (never edit it), is: {ref_path}

## The routine (do these in order)

1. LOOK. safexl_inspect for what the workbook contains. safexl_view for an overview, safexl_view with a sheet and
   range for exact cells, safexl_view with find to locate a label. It shows values, types (text is quoted, numbers
   are bare) and formulas.
2. PLAN. Say the exact calls you will make: sheet, row number, cells, values, formulas, and why.
3. WRITE with the safexl tools only.
   - safexl_edit writes cell values and formulas.
   - safexl_insert_rows inserts whole rows (the row you name and everything below move down).
   - safexl_insert_cells inserts blank cells in a column span only and shifts just those columns down. Use it
     instead of safexl_insert_rows when other lists or tables sit side by side on the sheet and must not move.
   - safexl_rename_sheet renames a sheet. safexl_add_sheet adds an empty sheet.
   - safexl_resize_table resizes a table; safexl_edit already extends an adjacent table by default.
   - Deliver the result in the workbook at the path given above: pass that same path as output_path with
     overwrite true on every write call, so each call builds on the previous one and the final file keeps the
     original name. The tool never modifies the file until the output is validated.
4. CHECK. safexl_verify (pristine copy, edited file) says whether the file is intact. Then safexl_view the cells
   you changed and the rows around them.
5. REPAIR at most once if the check shows a mistake, then stop and report.

## Usage slips to avoid

- Write numbers, booleans and ordinary text as plain values. Use a text-typed value only for codes that must stay
  text (007, 1E5). A number stored as text is ignored by SUM.
- Match the type of the neighbouring cells. In a lookup formula the key must have the same type as the lookup
  column: a number is not found in a column of text, and text is not found in a column of numbers.
- To add a line inside a list: insert rows first, then write the new row. Where neighbouring rows use a formula in
  a column, write the same formula pattern for the new row.
- A total just above an inserted row does not grow by itself. =SUM(B2:B6) with a row inserted at 7 needs
  the formula =SUM(B2:B7) written to the total's new cell.
- To show figures that already exist elsewhere in the workbook, write formulas that point at those cells
  (='Other Sheet'!F20), not copied values.
- Formulas are not calculated by safexl; they compute when the file is next opened. Do not read a new formula's
  value back from the file.
- Row inserts move whole rows; if other lists share those rows, use safexl_insert_cells. No column inserts or row
  deletes.

## A file that was already damaged

safexl_salvage(original, damaged) rebuilds lost parts from the original, only for a workbook that was re-saved by
another library. Then run safexl_verify.

## Safety

SafeXL never runs VBA. .xlsm macros come out byte-identical. When you are done, reply with a short summary of what
you changed.
"""

def key():
    name = "OPENROUTER_API_KEY" if PAID else "FREE_LLM_APP_KEY"
    k = os.environ.get(name)
    if not k:
        raise RuntimeError(f"environment variable {name} is not set")
    return k

class DailyCap(Exception):
    pass

class BudgetStop(Exception):
    pass

def or_usage():
    req = urllib.request.Request(OR_CREDITS, headers={"Authorization": f"Bearer {key()}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return float(json.loads(r.read())["data"]["total_usage"])

def spent():
    """USD spent on the paid key since the eval started (baseline recorded once, shared by all workers)."""
    cur = or_usage()
    if not BASELINE_FILE.exists():
        BASELINE_FILE.write_text(json.dumps({"baseline_total_usage": cur, "at": time.strftime("%FT%TZ", time.gmtime())}))
    return cur - json.loads(BASELINE_FILE.read_text())["baseline_total_usage"]

def check_budget():
    if PAID:
        sp = spent()
        with open(RES / "paid_spend.log", "a") as f:
            f.write(f"{time.strftime('%T')} spent={sp:.4f}\n")
        if sp >= BUDGET_USD:
            raise BudgetStop(f"paid spend {sp:.4f} >= cap {BUDGET_USD}")

def chat(model, messages, tools, max_tokens=4096, retries=None):
    """One gateway call with retry. Returns (json, retries_log). Raises RuntimeError after retries."""
    body = {"model": model, "messages": messages, "max_tokens": max_tokens, "temperature": 0}
    if PAID:
        body["usage"] = {"include": True}
    if tools:
        body["tools"] = tools
        body["tool_choice"] = "auto"
    data = json.dumps(body).encode()
    log = []
    last = None
    nret = API_RETRIES if retries is None else retries
    for attempt in range(nret + 1):
        req = urllib.request.Request(OR_URL if PAID else GATEWAY, data=data, headers={
            "Content-Type": "application/json", "Authorization": f"Bearer {key()}"})
        try:
            with urllib.request.urlopen(req, timeout=REQ_TIMEOUT) as r:
                return json.loads(r.read()), log
        except urllib.error.HTTPError as e:
            txt = e.read().decode("utf-8", "replace")[:500]
            last = f"HTTP {e.code}: {txt}"
            if "daily_cap_exceeded" in txt:
                raise DailyCap(last)
            if e.code in (400, 401, 403, 404, 422):
                log.append(last)
                raise RuntimeError(last)
            ra = e.headers.get("Retry-After")
            wait = min(float(ra), 90) if ra and ra.replace(".", "").isdigit() else min(15 * 2 ** attempt, 120)
        except Exception as e:  # timeout / connection
            last = f"{type(e).__name__}: {e}"
            wait = min(15 * 2 ** attempt, 120)
        log.append(last)
        if attempt < nret:
            time.sleep(wait)
    raise RuntimeError(last)

def mcp_tools_to_openai(tools):
    return [{"type": "function", "function": {"name": t.name, "description": t.description or "",
                                              "parameters": t.input_schema}} for t in tools]

def is_err(res):
    """mcp 2.x CallToolResult uses is_error (1.x: isError)."""
    v = getattr(res, "is_error", None)
    if v is None:
        v = getattr(res, "isError", False)
    return bool(v)

def result_text(res):
    parts = []
    for c in res.content:
        parts.append(getattr(c, "text", None) or str(c))
    s = "\n".join(parts)
    if len(s) > TOOL_RESULT_CAP:
        s = s[:TOOL_RESULT_CAP] + f"\n[truncated by harness at {TOOL_RESULT_CAP} chars]"
    return s

async def with_session(fn):
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client
    async with stdio_client(StdioServerParameters(command=MCP_BIN, args=[])) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            return await fn(s)

def probe_one(model):
    tools = [{"type": "function", "function": {"name": "echo", "description": "Echo a word back.",
              "parameters": {"type": "object", "properties": {"word": {"type": "string"}}, "required": ["word"]}}}]
    msgs = [{"role": "user", "content": "Call the echo tool with word 'ping'. Do not answer in text."}]
    t0 = time.time()
    try:
        resp, log = chat(model, msgs, tools, max_tokens=200, retries=2)
        m = resp["choices"][0]["message"]
        tc = m.get("tool_calls") or []
        if tc:
            args = json.loads(tc[0]["function"]["arguments"])
            ok = tc[0]["function"]["name"] == "echo" and isinstance(args, dict) and "word" in args
            return {"model": model, "ok": ok, "detail": "well-formed" if ok else f"bad call {tc[0]}", "s": round(time.time() - t0, 1)}
        return {"model": model, "ok": False, "detail": "no tool call; content=" + repr((m.get("content") or "")[:80]), "s": round(time.time() - t0, 1)}
    except DailyCap as e:
        return {"model": model, "ok": False, "detail": "DAILY CAP " + str(e)[:150], "s": round(time.time() - t0, 1)}
    except Exception as e:
        return {"model": model, "ok": False, "detail": str(e)[:200], "s": round(time.time() - t0, 1)}

def slug(model):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", model)

async def run_task(model, provider, task, rep, chat_fn=None, rec_root=None):
    root = (rec_root or REC) / slug(model)
    rd = root / task["id"] / "B" / str(rep)
    if (rd / "run.json").exists():
        return json.loads((rd / "run.json").read_text())  # never re-run a finished cell
    work = rd / "work"; ref = rd / "ref"
    work.mkdir(parents=True, exist_ok=True); ref.mkdir(exist_ok=True)
    src = REPO / task["workbook"]
    name = src.name
    shutil.copy2(src, work / name); shutil.copy2(src, ref / name)
    sysmsg = SYSTEM_PROMPT.format(work_path=str(work / name), ref_path=str(ref / name))
    msgs = [{"role": "system", "content": sysmsg}, {"role": "user", "content": task["prompt"]}]
    rec = {"task_id": task["id"], "arm": "B", "rep": rep, "workbook": task["workbook"], "prompt": task["prompt"],
           "model": model, "provider": provider, "started_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
           "output_files": [], "contamination": {"contaminated": False, "evidence": []},
           "timed_out": False, "tool_calls": [], "api_retries": [], "tokens_in": 0, "tokens_out": 0,
           "steps": 0, "first_tool_call_s": None, "final_text": "", "stop_reason": None}
    transcript = []
    t0 = time.time()

    async def loop(session):
        tools = mcp_tools_to_openai((await session.list_tools()).tools)
        while True:
            if time.time() - t0 > MAX_WALL:
                rec["stop_reason"] = "cap_wall"; rec["timed_out"] = True; return
            if len(rec["tool_calls"]) >= MAX_TOOL_CALLS:
                rec["stop_reason"] = "cap_tool_calls"; return
            try:
                check_budget()
            except BudgetStop:
                raise
            try:
                resp, log = await asyncio.get_event_loop().run_in_executor(None, chat_fn or chat, model, msgs, tools)
            except DailyCap:
                raise
            except Exception as e:
                rec["stop_reason"] = "api_failure"; rec["api_error"] = str(e)[:300]; return
            rec["api_retries"] += log
            rec["steps"] += 1
            u = resp.get("usage") or {}
            rec["tokens_in"] += u.get("prompt_tokens") or 0
            rec["tokens_out"] += u.get("completion_tokens") or 0
            if PAID:
                rec["cost_usd"] = rec.get("cost_usd", 0.0) + float(u.get("cost") or 0)
                if resp.get("provider"): rec.setdefault("providers", [])
                if resp.get("provider") and resp["provider"] not in rec["providers"]: rec["providers"].append(resp["provider"])
            ch = (resp.get("choices") or [{}])[0]
            m = ch.get("message") or {}
            tcs = m.get("tool_calls") or []
            content = m.get("content") or ""
            transcript.append({"role": "assistant", "content": content, "tool_calls": tcs, "finish": ch.get("finish_reason")})
            if not tcs:
                rec["final_text"] = content
                rec["stop_reason"] = "finished" if content.strip() else "empty_reply"
                return
            am = {"role": "assistant", "content": content or None, "tool_calls": tcs}
            msgs.append(am)
            for tc in tcs:
                fn = tc.get("function", {})
                tname = fn.get("name", "")
                if rec["first_tool_call_s"] is None:
                    rec["first_tool_call_s"] = round(time.time() - t0, 2)
                err, kind = False, None
                try:
                    args = json.loads(fn.get("arguments") or "{}")
                    if not isinstance(args, dict):
                        raise ValueError("arguments not an object")
                except Exception as e:
                    out = f"Error: tool arguments were not valid JSON ({e})"; err, kind = True, "malformed_args"
                else:
                    try:
                        res = await session.call_tool(tname, args)
                        out = result_text(res)
                        if is_err(res):
                            err = True
                            kind = "unknown_tool" if "nknown tool" in out else "tool_error"
                    except Exception as e:
                        out = f"Error calling tool: {e}"; err, kind = True, "tool_exception"
                rec["tool_calls"].append({"name": tname, "error": err, "kind": kind,
                                          "args": (fn.get("arguments") or "")[:600], "result": out[:600]})
                transcript.append({"role": "tool", "name": tname, "args": fn.get("arguments"), "result": out[:4000]})
                msgs.append({"role": "tool", "tool_call_id": tc.get("id"), "content": out})

    try:
        await with_session(loop)
    except BudgetStop as e:
        rec["stop_reason"] = "budget_stop"; rec["api_error"] = str(e); rec["daily_cap"] = True  # not a finished cell
    except DailyCap as e:
        rec["stop_reason"] = "api_failure"; rec["api_error"] = "DAILY CAP: " + str(e)[:200]; rec["daily_cap"] = True
    except Exception as e:
        rec["stop_reason"] = rec["stop_reason"] or "harness_error"; rec["api_error"] = f"{type(e).__name__}: {e}"[:300]
    rec["wall_seconds"] = round(time.time() - t0, 2)
    rec["finished_at"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    rec["tool_call_count"] = len(rec["tool_calls"])
    rec["tool_error_count"] = sum(1 for t in rec["tool_calls"] if t["error"])
    (rd / "transcript.json").write_text(json.dumps(transcript, indent=1))
    if rec.get("daily_cap"):
        return rec  # not written as a finished cell; may be re-run after reset
    (rd / "run.json").write_text(json.dumps(rec, indent=1))
    return rec

def load_tasks(which):
    f = BENCH / ("tasks_hard8.json")
    ts = json.load(open(f))
    if which == "smoke":
        ts = [t for t in ts if t["id"] == "benefits-t4-insert-duplicate-exclusion"]
    return ts

def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd")
    p = sub.add_parser("probe"); p.add_argument("--models-file"); p.add_argument("--out")
    r = sub.add_parser("run"); r.add_argument("--model"); r.add_argument("--provider"); r.add_argument("--tasks", default="hard8"); r.add_argument("--reps", default="1")
    a = ap.parse_args()
    if a.cmd == "probe":
        ms = [l.strip().split(None, 1) for l in open(a.models_file) if l.strip()]
        out = []
        for prov, mid in ms:
            res = probe_one(mid); res["provider"] = prov; out.append(res)
            print(json.dumps(res), flush=True)
        json.dump(out, open(a.out, "w"), indent=1)
    elif a.cmd == "run":
        for rep in [int(x) for x in a.reps.split(",")]:
            for t in load_tasks(a.tasks):
                rec = asyncio.run(run_task(a.model, a.provider, t, rep))
                print(json.dumps({"model": a.model, "task": t["id"], "rep": rep, "stop": rec["stop_reason"],
                                  "calls": rec.get("tool_call_count"), "s": rec.get("wall_seconds")}), flush=True)
                if rec.get("daily_cap"):
                    print("DAILY CAP HIT; stopping", flush=True); sys.exit(3)

if __name__ == "__main__":
    main()
