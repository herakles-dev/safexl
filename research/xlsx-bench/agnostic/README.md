# SafeXL model-agnostic eval (2026-10-04)

Question: can non-Claude models, given only SafeXL's real MCP tools, do the hard-8 xlsx tasks and keep files intact?

- `harness.py`: probe + per-run agent loop. Models reached via an OpenAI-compatible gateway (local server, address in `GATEWAY` in `harness.py`; tool calling, temperature 0).
  The MCP server is `packages/safexl/.venv/bin/safexl-mcp` (0.2.0, 073862d) over stdio; its 11 tools are converted programmatically to OpenAI function specs (real names, descriptions, schemas).
  The model has no other tools. Tool errors are returned to the model as tool results. Caps: 40 tool calls, 12 min wall, 60 s per API request on the gateway and 120 s on OpenRouter (3 retries with backoff; exhaustion = stop reason `api_failure`, treated as rate-limited out). Tool results truncated at 30000 chars.
- Commit hashes in this folder (such as 073862d) refer to my private development history.
- `pipeline.py`: stage 2 (smoke: benefits-t4-insert-duplicate-exclusion) then stage 3 (hard 8 x rep 1) per model, fastest probe latency first; finished cells are skipped on restart.
- Outputs: `recordings/agnostic-2026-10-04/<model>/<task>/B/<rep>/{work,ref,run.json,transcript.json}`; grades via `grade.py` (arm B) into `results/agnostic-2026-10-04/`.

## Fixed system prompt (same for every model and task, written before stage 3)
`SYSTEM_PROMPT` in `harness.py` is the SKILL.md routine with CLI commands swapped for MCP tool names. Inspect, view, edit, insert-rows, insert-cells, rename-sheet, add-sheet, resize-table, verify and salvage become `safexl_*`. The prompt gives the workbook path to edit and the path of a pristine reference copy for `safexl_verify`. It also says to pass the same path as `output_path` with overwrite true. That way the result lands under the original file name. It was meant to hold no task-specific advice. In the 2026-10-04 round it did hold one: two usage-slip lines copied from the skill text held the benefits task's own rows and formula. The prompt as run is in `../results/agnostic-2026-10-04/system_prompt_as_run.txt`. That round's benefits runs are left out of the published results, and the prompt in `harness.py` now uses a neutral example. The user message is the task prompt, word for word.

## Notes
- 11 of 22 listed models were unavailable at probe time (provider keys in backoff, 402/502, unavailable); recorded as "not available on the gateway today", not model failures, not retried.

## Self-test (run before any model run; `selftest.py`, no model)
An earlier attempt was invalid: the bridge read `res.isError` but mcp 2.x names it `is_error`, so every call was shown to the model as failed (those 20B runs are not included here, excluded from all results).
Fixed and verified: (a) all 11 tools called through the bridge with known-good arguments return non-error results; (b) an oracle scripted run through the real `run_task` loop (inspect, rename Sheet2 -> Yearly Totals, verify) finishes and grades as pass on pivotrpt-t5-rename-sheet2.
## Execution
One worker process per model (never two on the same model), up to 5 in parallel (`xargs -P 5`), in the order of `order.txt`.
- `probe.json` and `paid_baseline_usage.json` from my own run are not included. `harness.py probe` and the paid stage write them when you run the harness; `summarize.py` needs them. The published numbers come from `../results/agnostic-2026-10-04/make_summary.py`, which needs neither.
