# Haiku 4.5 + SafeXL plugin v0.2 (guard hooks on), NO operator hint: hard 8 x 3, arm B only, final skill text (2026-10-05)

**Method.** `record.py --tasks tasks_hard8.json --all --arms B --reps 3`, model `claude-haiku-4-5-20251001`, Claude subscription (no API key in the environment), `XLSX_BENCH_PLUGIN_HOOKS=1` (the `user` setting source is dropped so only the plugin's hooks can fire), no `--append-system-prompt`, no tuning, no cherry-picking. The `safexl` CLI is the editable install of the working tree at the pinned commit. No cell was re-run, rate-limited or timed out. Every session ran in a fresh temp directory holding only its input workbook; every `run.json` records `invocation.isolated: true`. `scan_contamination.py`: 24 scanned, 0 contaminated, 0 output-hit warnings. Graded with `grade.py` (grader fix #17), 0 excluded.

**What "no hint" means.** There was no operator instruction. The plugin's SessionStart hook still tells the model, in every session, to use the safexl skill and never to save with openpyxl or pandas. By my check of the raw session logs, which are not published, it reached the model in 24/24 sessions.

**Pinned inputs.** The plugin was taken with `git archive` of the repository at the pinned commit (the plugin directory only), into a fresh directory. Content hashes (sha256), checked against the pinned copy and against the files the runs loaded:

| file | sha256 |
|---|---|
| `skills/safexl/SKILL.md` | `97545d71187cd41e15073f75d385d68ecc7e533d9a43a4b26677c8589a76768d` |
| `hooks/guard_hook.py` | `1f0924ff0b08a5ad2d4dff42693a990e9d82e5699ba3d0479e6221427ff38cf5` |
| `hooks/hooks.json` | `be10416202250b00cad932e5f5ad3f59364f500f9833b8e56ee178f3efc61b94` |
| `.claude-plugin/plugin.json` | `6cc28c310651116c6332189148e314d2510cd080368329405c077f782081f67e` |
| `.mcp.json` | `010569da8df17c4c2f087c09effec180772852d444fbead69507363d4a61d309` |

These are the same files as `packages/safexl/claude-plugin/` in this repository; `sha256sum` on them should give the values above. The runs loaded a copy outside the repository (each `run.json` records it as `plugin_dir`); its hashes were taken then, and that copy is not included.

Before any run, the task-leak test (`tests/test_no_task_leak.py`, 10 tests when the round ran and 11 now, covers task text, ranges, formulas, short codes, cell targets and insert rows) passed, and the skill text and the SessionStart and guard messages were read against all 8 task prompts and expectations; none gives a task's sheet name, row, cell, label, value or formula.

## Why this round exists

Two earlier plugin rounds on these tasks are superseded and must not be quoted. Neither is in this repository.
- **2026-10-04.** The plugin's skill examples held the literal answers to two tasks (`pivotrpt-t5-rename-sheet2` and the first step of `sparkext-t5-add-lookup-check-sheet`). That round scored 23/24 intact, 21/24 pass, 22/24 correct, 22/24 valid.
- **First 2026-10-05 re-run.** The two examples were replaced, but the skill's totals example was still the benefits task's own rows and formula, and the leak check did not cover formulas. That round scored 22/24 intact, 19/24 pass, 20/24 correct, 22/24 valid.

The leak test now catches ranges, formulas, short codes, cell targets and insert rows. This round is the same run with the final skill text.

| round | intact | pass | correct | valid |
|---|---:|---:|---:|---:|
| 2026-10-04 (superseded; leaked text) | 23/24 | 21/24 | 22/24 | 22/24 |
| first 2026-10-05 re-run (superseded; formula still in the skill) | 22/24 | 19/24 | 20/24 | 22/24 |
| **this round (final skill text)** | **22/24** | **21/24** | 23/24 | 23/24 |

## Headline (n = 24, grader fix #17)

| round | intact | pass | correct | valid |
|---|---:|---:|---:|---:|
| **plugin + hooks, no operator hint, final skill text** | **22/24** | **21/24** | 23/24 | 23/24 |

Cost: $1.862 total, $0.078 per run, $0.089 per passing run (total cost divided by 21 passes; Claude subscription token-equivalent as reported by the CLI).

`safexl verify original output --json`, run by the author after the round, reports "intact" for all 24 outputs, including the two the grader marks not preserved (structrefs reps 1 and 2). The output workbooks are not in this repository, so that check cannot be repeated from it.

## Per task (3 reps)

| task | pass | intact | cost/run $ |
|---|---:|---:|---|
| benefits-t4-insert-duplicate-exclusion | 3/3 | 3/3 | .069 .060 .066 |
| kyc-t4-insert-ladakh-state | 3/3 | 3/3 | .078 .072 .070 |
| parking-t5-rename-detailed-report | 3/3 | 3/3 | .063 .042 .051 |
| pension-t4-insert-cpi-linked-asset-row | 3/3 | 3/3 | .087 .083 .093 |
| pivotrpt-t4-insert-games-row | 3/3 | 3/3 | .063 .066 .075 |
| pivotrpt-t5-rename-sheet2 | 3/3 | 3/3 | .099 .046 .045 |
| sparkext-t5-add-lookup-check-sheet | 2/3 | 3/3 | .069 .074 .071 |
| structrefs-t4-add-guam-to-dropdowns | 1/3 | 1/3 | .150 .139 .130 |

## Writer and guard, per run

- Writer: all 24 outputs were written with `safexl` commands. No run saved with openpyxl (no `.save(` call in any tool input) and none used the Write or Edit tools on a workbook.
- Guard: SessionStart reached the model in 24/24 sessions, and the guard's state files show no warning and no derived copy in any session. Both come from my check of logs and state files that are not published. No guard notice appears in any transcript, which you can check here: **the guard did not act in any of the 24 runs.** The guard's damage-recovery path (salvage or restore) was not exercised by this round.

## Failures (3), from transcripts and grader output

1. **sparkext rep 3** (correct F, valid F, preserved T). The formula `Lookup Check!B2 = VLOOKUP(A2,Sheet2!A:L,12,FALSE)` recalculates to `#N/A`. That is a new error cell. The grader expects the Sheet2 value 1.86546e-05 to come through a formula. The file itself is intact. The lookup is wrong. In its final text the model said it had retrieved the value.
2. **structrefs rep 1** (correct T, preserved F, valid T). The model ran `insert-rows --row 13` on the dropdowns sheet. That shifted whole rows. Two defined names changed: `dropdown` went from $E$2:$E$13 to $E$14, and `droptable` went from $E$1:$G$13 to $E$1:$G$14. The grader counts that as lost identity. The model did not use `insert-cells`. It corrected the cell contents afterward.
3. **structrefs rep 2** (same checks). The model ran the same `insert-rows --row 13`, with the same two defined-name changes. It then reordered column C with `edit`. The content is right and the names changed. Only rep 3 passed.

In both structrefs failures, `safexl verify` reports intact while the grader reports the defined-name change. For these runs the receipt does not agree with the grader.

## What this is not

This is not a held-out test. The skill's guidance was written after watching models fail on these same 8 tasks (for example the guidance on lookup key types, on text-typed codes, on totals above an inserted row, and on `insert-cells` for side-by-side lists), so the 8 tasks are in-sample for the plugin. The skill text and messages contain none of the tasks' answers (checked by the leak test and by reading), but the advice itself was shaped by these tasks' failures. n = 3 per task, one model, one run of each cell; two tasks (sparkext, structrefs) account for all three failures. A number from this round should be quoted only with this caveat.

## How this arm differs from the skill arm (`../haiku-skill-2026-10-05/`)

Same model, tasks, reps, grader, day and isolation. Different Claude Code settings: the skill arm ran with `--setting-sources user,local`, hooks off, no plugin and a one-line operator hint to use the xlsx skill; this arm ran with `--setting-sources local`, the plugin loaded with its hooks on, the Anthropic xlsx skill denied, and no operator hint. The plugin's start-up message tells the model to use safexl and never to save with openpyxl or pandas. The two arms ran as separate batches.

## Differences from the earlier plugin rounds

- Plugin pinned by the five content hashes above; the plugin skill text is final.
- The CLI is the editable install of the working tree at the pinned commit. Since the earlier rounds the engine changed in four files (`engine.py`, `features.py`, `inspect.py`, `validate.py`): docstrings, comments and one message string. The parsed code is otherwise identical.
- The runner has had path-default and comment-only edits (default paths resolve from the checkout, plus an env override for the SafeXL venv directory).
- The smoke cell (`pivotrpt-t5-rename-sheet2`, rep 1) was recorded straight into this round's directory and counts as that cell; it was not re-run.
