# Non-Claude models through SafeXL's MCP tools only: hard-8 results, 7 tasks counted

Date of runs: 2026-10-04. Each model was given only SafeXL's 11 MCP tools and one task prompt. There is no arm without SafeXL in this eval, so it says nothing about what these models do without it.

Tasks: 7 of the 8 hard tasks in `research/xlsx-bench/tasks_hard8.json`, up to 3 reps each (21 runs is a full set). One task is left out of every headline number; see "Why one task is left out" below. n is the number of runs graded. Any model under 21 is partial and is not comparable to a 21-run row.

## Definitions (from `grade.py`)

A run counts as **pass** when the grader finds every expected cell, formula or sheet change correct, finds no new error cells, can open and recalculate the file, and finds nothing lost from the original workbook. A run counts as **intact** when the grader's `preserved` check finds no losses against the original workbook (formatting, validations, named ranges, extensions and similar). A run can be intact and still fail the task.

## Results (7 tasks, excluded task left out)

| model | tier | n | pass | intact | median wall s | $ per run | note |
|---|---|---|---|---|---|---|---|
| deepseek/deepseek-v4-flash | paid | 21 | 21 | 21 | 28.6 | 0.0110 | - |
| qwen/qwen3.8-27b:free | free | 21 | 18 | 19 | 150.5 | not billed | 2 api failure(s) |
| nvidia/nemotron-3-super-120b-a12b:free | free | 21 | 17 | 18 | 43.9 | not billed | 2 empty replies |
| poolside/laguna-s-2.1:free | free | 21 | 16 | 16 | 173.5 | not billed | 4 api failure(s); 2 of 167 tool calls errored |
| mistralai/mistral-small-3.2-24b-instruct | paid | 21 | 13 | 19 | 25.6 | 0.0030 | - |
| openai/gpt-oss-120b | paid | 21 | 12 | 15 | 39.5 | 0.0021 | 5 empty replies; 2 of 104 tool calls errored |
| anthropic/claude-haiku-4.5 | paid | 7 | 6 | 6 | 24.8 | 0.0933 | partial (7 of 21; 7 tasks) |
| deepseek/deepseek-v3.2 | paid | 5 | 3 | 4 | 86.6 | 0.0284 | partial (5 of 21; 5 tasks) |
| command-a-plus-05-2026 | free-gateway | 7 | 4 | 5 | 295.9 | none recorded | partial (7 of 21; 7 tasks); 3 hit the tool-call cap; 42 of 162 tool calls errored |
| openai/gpt-oss-20b | paid | 7 | 4 | 5 | 50.4 | 0.0011 | partial (7 of 21; 7 tasks); 3 empty replies; 6 of 31 tool calls errored |
| command-a-03-2025 | free-gateway | 7 | 2 | 6 | 22.8 | none recorded | partial (7 of 21; 7 tasks) |

Tier: `paid` = OpenRouter, billed per request, cost is the measured `usage.cost`. `free` = OpenRouter free-tier id. `free-gateway` = free gateway, no cost recorded. Median wall time is per run, from the run record.

## All 8 tasks, including the affected one (for reference; do not quote)

| model | n | pass | intact | `benefits-t4-insert-duplicate-exclusion` alone |
|---|---|---|---|---|
| deepseek/deepseek-v4-flash | 24 | 24 | 24 | 3 of 3 |
| qwen/qwen3.8-27b:free | 24 | 21 | 22 | 3 of 3 |
| nvidia/nemotron-3-super-120b-a12b:free | 24 | 20 | 21 | 3 of 3 |
| poolside/laguna-s-2.1:free | 24 | 19 | 19 | 3 of 3 |
| mistralai/mistral-small-3.2-24b-instruct | 24 | 16 | 22 | 3 of 3 |
| openai/gpt-oss-120b | 24 | 15 | 18 | 3 of 3 |
| anthropic/claude-haiku-4.5 | 8 | 7 | 7 | 1 of 1 |
| deepseek/deepseek-v3.2 | 6 | 4 | 5 | 1 of 1 |
| command-a-plus-05-2026 | 8 | 4 | 6 | 0 of 1 |
| openai/gpt-oss-20b | 8 | 4 | 6 | 0 of 1 |
| command-a-03-2025 | 8 | 2 | 7 | 0 of 1 |

## Why one task is left out

The system prompt every model received held two lines as a usage example:

```
- A total just above an inserted row does not grow by itself. =SUM(C25:C28) with a row inserted at 29 needs the formula =SUM(C25:C29) written to the total's new cell.
- To show figures that already exist elsewhere in the workbook, write formulas that point at those cells (='Other Sheet'!C30), not copied values.
```

These are the rows and the total of task `benefits-t4-insert-duplicate-exclusion`. The runs of that task are excluded from every headline number above, because the prompt held what the task asks for. On that task the 6 models with a full set passed 18 of 18 runs. The 5 partial models passed 2 of 5. The per-model figures are in the reference table above.

The recordings do not hold the system prompt. Its full text as run is in `system_prompt_as_run.txt` in this folder. I took it from the harness source as it stood during the round. The leak scan finds only the two lines above in it.

The rest of the prompt's guidance came from SafeXL's skill text, which was written after watching models fail on these same tasks. So this round is in-sample for that guidance and is not a held-out test.

## Spend (whole round, all 8 tasks)

- OpenRouter spend during the eval, from the spend log (usage minus the baseline taken before the paid runs): $1.3236.
- Sum of per-run measured cost over all 94 paid-tier run records: $1.3126.
- Sum over all graded paid runs, including the excluded task: $1.3126.
- Free-tier runs were not billed.

## Models that could not be run

Three models were tried on the Groq free tier and could not be measured there. The run records show Groq rejecting requests with a limit of 8,000 tokens per minute against requests of 8,280 and 8,345 tokens, and gateway 503s when its keys were in backoff. Their runs are not graded here:

- openai/gpt-oss-120b: 6 runs recorded, stop reasons {'api_failure': 5, 'finished': 1}.
- openai/gpt-oss-20b: 8 runs recorded, stop reasons {'api_failure': 7, 'finished': 1}.
- qwen/qwen3.8-27b: 8 runs recorded, stop reasons {'api_failure': 8}.

gpt-oss-120b and gpt-oss-20b were then run on OpenRouter's paid tier, and Qwen 3.8 27B on OpenRouter's free tier. Those runs are the rows in the tables above.

## Per-model failures (7 tasks)

- deepseek/deepseek-v4-flash (21 runs): failed 0: none
- qwen/qwen3.8-27b:free (21 runs): failed 3: pivotrpt-t5-rename-sheet2#2, sparkext-t5-add-lookup-check-sheet#3, structrefs-t4-add-guam-to-dropdowns#3
- nvidia/nemotron-3-super-120b-a12b:free (21 runs): failed 4: sparkext-t5-add-lookup-check-sheet#3, structrefs-t4-add-guam-to-dropdowns#1, structrefs-t4-add-guam-to-dropdowns#2, structrefs-t4-add-guam-to-dropdowns#3
- poolside/laguna-s-2.1:free (21 runs): failed 5: pension-t4-insert-cpi-linked-asset-row#2, pension-t4-insert-cpi-linked-asset-row#3, structrefs-t4-add-guam-to-dropdowns#1, structrefs-t4-add-guam-to-dropdowns#2, structrefs-t4-add-guam-to-dropdowns#3
- mistralai/mistral-small-3.2-24b-instruct (21 runs): failed 8: kyc-t4-insert-ladakh-state#1, kyc-t4-insert-ladakh-state#2, sparkext-t5-add-lookup-check-sheet#1, sparkext-t5-add-lookup-check-sheet#2, sparkext-t5-add-lookup-check-sheet#3, structrefs-t4-add-guam-to-dropdowns#1, structrefs-t4-add-guam-to-dropdowns#2, structrefs-t4-add-guam-to-dropdowns#3
- openai/gpt-oss-120b (21 runs): failed 9: pivotrpt-t4-insert-games-row#1, pivotrpt-t4-insert-games-row#2, pivotrpt-t4-insert-games-row#3, sparkext-t5-add-lookup-check-sheet#1, sparkext-t5-add-lookup-check-sheet#2, sparkext-t5-add-lookup-check-sheet#3, structrefs-t4-add-guam-to-dropdowns#1, structrefs-t4-add-guam-to-dropdowns#2, structrefs-t4-add-guam-to-dropdowns#3
- anthropic/claude-haiku-4.5 (7 runs): failed 1: structrefs-t4-add-guam-to-dropdowns#1
- deepseek/deepseek-v3.2 (5 runs): failed 2: pension-t4-insert-cpi-linked-asset-row#1, structrefs-t4-add-guam-to-dropdowns#1
- command-a-plus-05-2026 (7 runs): failed 3: pension-t4-insert-cpi-linked-asset-row#1, pivotrpt-t4-insert-games-row#1, structrefs-t4-add-guam-to-dropdowns#1
- openai/gpt-oss-20b (7 runs): failed 3: kyc-t4-insert-ladakh-state#1, sparkext-t5-add-lookup-check-sheet#1, structrefs-t4-add-guam-to-dropdowns#1
- command-a-03-2025 (7 runs): failed 5: kyc-t4-insert-ladakh-state#1, pension-t4-insert-cpi-linked-asset-row#1, pivotrpt-t4-insert-games-row#1, sparkext-t5-add-lookup-check-sheet#1, structrefs-t4-add-guam-to-dropdowns#1

Regenerate with `python3 make_summary.py` in this directory. `table.json` holds the 7-task numbers as the main fields, plus an `all8` object, the excluded task's own result and per-task rep counts.
