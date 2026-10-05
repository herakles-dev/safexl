#!/usr/bin/env python3
"""No-model self-test of the tool bridge. (a) every one of the 11 MCP tools called with known-good args through the
bridge code (result_text/is_err) must come back non-error; (b) an oracle scripted 'model' run through run_task's real
loop must solve pivotrpt-t5-rename-sheet2 and grade as pass. Run before any model run."""
import asyncio, json, os, re, shutil, subprocess, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import harness as H
S = H.Path(tempfile.gettempdir()) / "xlsx-bench-agnostic-selftest"
S.mkdir(parents=True, exist_ok=True)
C = H.BENCH / "corpus"
PY = os.environ.get("AGN_PYTHON", str(H.REPO / "packages" / "safexl" / ".venv" / "bin" / "python"))

async def part_a():
    import uuid
    d = S / ("a-" + uuid.uuid4().hex[:6]); d.mkdir()
    pv = d / "pivot-chart-report.xlsx"; shutil.copy2(C / "pivot-chart-report.xlsx", pv)
    ref = d / "ref.xlsx"; shutil.copy2(C / "pivot-chart-report.xlsx", ref)
    lk = d / "structured-refs-lookup-tool.xlsx"; shutil.copy2(C / "structured-refs-lookup-tool.xlsx", lk)
    res = {}
    async def go(s):
        tools = {t.name for t in (await s.list_tools()).tools}
        calls = [
            ("safexl_inspect", {"path": str(pv)}),
            ("safexl_view", {"path": str(pv), "sheet": "Sheet2", "cell_range": "A1:C6"}),
            ("safexl_edit", {"path": str(pv), "changes": [{"sheet": "Sheet1", "cell": "Z90", "value": 1}], "output_path": str(d / "e.xlsx"), "overwrite": True}),
            ("safexl_insert_rows", {"path": str(pv), "sheet_name": "Sheet1", "row": 3, "output_path": str(d / "r.xlsx"), "overwrite": True}),
            ("safexl_insert_cells", {"path": str(pv), "sheet_name": "Sheet1", "cell_range": "H3:H3", "output_path": str(d / "c.xlsx"), "overwrite": True}),
            ("safexl_rename_sheet", {"path": str(pv), "old_name": "Sheet2", "new_name": "Renamed", "output_path": str(d / "n.xlsx"), "overwrite": True}),
            ("safexl_add_sheet", {"path": str(pv), "sheet_name": "NewS", "output_path": str(d / "a.xlsx"), "overwrite": True}),
            ("safexl_verify", {"before_path": str(ref), "after_path": str(d / "n.xlsx")}),
            ("safexl_diff", {"before_path": str(ref), "after_path": str(d / "n.xlsx")}),
            ("safexl_salvage", {"original_path": str(ref), "damaged_path": str(d / "n.xlsx"), "output_path": str(d / "s.xlsx"), "overwrite": True}),
        ]
        for name, args in calls:
            r = await s.call_tool(name, args)
            txt = H.result_text(r)
            res[name] = (H.is_err(r), txt[:150])
        r = await s.call_tool("safexl_resize_table", {"path": str(lk), "table": "Table_Query_from_RDS2", "ref": "L1:S16", "output_path": str(d / "t.xlsx"), "overwrite": True})
        res["safexl_resize_table"] = (H.is_err(r), H.result_text(r)[:150])
        assert tools == set(n for n, _ in calls) | {"safexl_resize_table"}, tools
    await H.with_session(go)
    bad = {k: v for k, v in res.items() if v[0] or '"ok": false' in v[1].replace('"ok":false', '"ok": false')}
    for k, v in res.items(): print("A", k, "ERR" if v[0] else "ok", v[1][:110].replace("\n", " "))
    return bad

async def part_b():
    task = [t for t in json.load(open(H.BENCH / "tasks_hard8.json")) if t["id"] == "pivotrpt-t5-rename-sheet2"][0]
    state = {"i": 0}
    def oracle(model, msgs, tools):
        work = re.search(r"workbook to edit is: (\S+)", msgs[0]["content"]).group(1)
        ref = re.search(r"original, for comparison only \(never edit it\), is: (\S+)", msgs[0]["content"]).group(1)
        i = state["i"]; state["i"] += 1
        seq = [("safexl_inspect", {"path": work}),
               ("safexl_rename_sheet", {"path": work, "old_name": "Sheet2", "new_name": "Yearly Totals", "output_path": work, "overwrite": True}),
               ("safexl_verify", {"before_path": ref, "after_path": work})]
        if i < len(seq):
            n, a = seq[i]
            return {"choices": [{"message": {"content": "", "tool_calls": [{"id": f"c{i}", "type": "function", "function": {"name": n, "arguments": json.dumps(a)}}]}, "finish_reason": "tool_calls"}], "usage": {"prompt_tokens": 1, "completion_tokens": 1}}, []
        return {"choices": [{"message": {"content": "Renamed Sheet2 to Yearly Totals.", "tool_calls": None}, "finish_reason": "stop"}], "usage": {"prompt_tokens": 1, "completion_tokens": 1}}, []
    import uuid
    root = S / ("b-" + uuid.uuid4().hex[:6])
    rec = await H.run_task("oracle", "none", task, 1, chat_fn=oracle, rec_root=root)
    print("B oracle:", rec["stop_reason"], [(t["name"], t["error"]) for t in rec["tool_calls"]])
    out = root / "grade.jsonl"
    subprocess.run([PY, str(H.BENCH / "grade.py"), "--tasks", str(H.BENCH / "tasks_hard8.json"), "--runs", str(root / "oracle"), "--out", str(out), "--summary"], check=False)
    row = [json.loads(l) for l in open(out) if l.strip()][0]
    print("B grade:", {k: row.get(k) for k in ("task_id", "passed", "pass", "intact", "verdict") if k in row}, list(row)[:20])
    return row

bad = asyncio.run(part_a())
print("PART A", "FAIL " + str(bad) if bad else "PASS")
row = asyncio.run(part_b())
