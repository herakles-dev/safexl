# Haiku 4.5 + Anthropic's xlsx skill, one-line operator hint: hard 8 x 3, arm A only, isolated (2026-10-05)

**Method.** `record.py --tasks tasks_hard8.json --all --arms A --reps 3`, model `claude-haiku-4-5-20251001`, Claude subscription (no API key in the environment), `XLSX_BENCH_OPERATOR_HINT=1`. The hint, appended with `--append-system-prompt`, is: "When a task involves an Excel workbook, use the xlsx skill for the edit." Arm A is the default environment: no plugin, SafeXL not on PATH, `--setting-sources user,local` (the skill lives in the user skills directory), hooks off, `--strict-mcp-config`, `--permission-mode dontAsk`, a block on reading the benchmark tree. Every session ran in a fresh `/tmp` directory holding only its input workbook. No tuning, no cherry-picking. No cell was re-run, rate-limited or timed out. Graded with `grade.py` (grader fix #17) with `--exclude-contaminated`, 0 excluded. `scan_contamination.py`: 24 scanned, 0 contaminated, 0 output-hit warnings. In every `run.json`, `invocation.isolated` is `true`.

**Pinned input: what "Anthropic's xlsx skill" is here.** The user-level synced copy of the `xlsx` skill (installed 2026-09-16). It has no version field in its frontmatter.
- `SKILL.md` sha256 `6712b39718fe815054abca9c3ee72f989613e6f771c8a5c58f30a65a6c905622` (99 lines).
- `LICENSE.txt` sha256 `79f6d8f5b427252fa3b1c11ecdbdb6bf610b944f7530b4de78f770f38741cfaa`.
- Directory hash (sha256 of the sorted per-file sha256 list) `cd9720398ccdf4cbcd1fd4169fe14d76659d352296c574c784930077232b9334`.
- The model invoked it under two names: `anthropic-skills:xlsx` in 20 runs and `xlsx` in 4 runs (kyc rep 1, parking rep 1, pivotrpt-t5 rep 1, sparkext rep 2). Both names loaded the same directory in all 24 runs.
- SafeXL was not available and not used in any run.

**Runner.** The runner at this commit differs from the one used for the 2026-09-29 arm-A round. The 09-29 argv had no deny rules. The session cwd was inside the repository. This round uses the isolated harness: temp-dir cwd, deny rules, contamination detector. Since the 10-05 arm-B round I changed the runner in three places: default paths, an env override for the SafeXL venv directory, and comments. Nothing else changed.

## Headline (n = 24)

| round | intact | pass | correct | valid |
|---|---:|---:|---:|---:|
| **this round (isolated)** | **6/24** | **5/24** | 11/24 | 22/24 |
| 2026-09-29 arm A (a different round; 11 of its 24 runs were not isolated) | 8/24 | 6/24 | 12/24 | 21/24 |

The 09-29 figures are a different, non-isolated round and are not a replication; they are shown only for reference. That round is not in this repository.

Cost: $2.351 total, $0.098 per run, $0.470 per passing run (total cost divided by 5 passes; Claude subscription token-equivalent as reported by the CLI).

## Per task (3 reps)

| task | pass | intact | cost/run $ |
|---|---:|---:|---|
| benefits-t4-insert-duplicate-exclusion | 0/3 | 0/3 | .116 .116 .097 |
| kyc-t4-insert-ladakh-state | 3/3 | 3/3 | .092 .092 .119 |
| parking-t5-rename-detailed-report | 0/3 | 0/3 | .064 .054 .064 |
| pension-t4-insert-cpi-linked-asset-row | 0/3 | 1/3 | .174 .095 .109 |
| pivotrpt-t4-insert-games-row | 2/3 | 2/3 | .082 .067 .072 |
| pivotrpt-t5-rename-sheet2 | 0/3 | 0/3 | .111 .068 .063 |
| sparkext-t5-add-lookup-check-sheet | 0/3 | 0/3 | .117 .114 .090 |
| structrefs-t4-add-guam-to-dropdowns | 0/3 | 0/3 | .134 .158 .083 |

## Which library wrote each output

All 24 outputs were saved with openpyxl (`.save()` in a tool call in 24/24). 10 sessions also ran the skill's `recalc.py` (LibreOffice recalculation); pension rep 1 also used zipfile and XML tools. No other writer was used and SafeXL was not available.

## Failures (19), from grader output on each run's graded workbook

- **benefits-t4 reps 1 to 3** (correct T, preserved F). The row is right. The macro (`vbaProject.bin` hash changed) and the data-validation blocks were lost. Each run has 4 losses.
- **parking-t5 reps 1 to 3** (correct F, preserved F). The sheet was renamed. The chart series and defined names still point to the old sheet name `Detailed_report`. Each run has 67 losses.
- **pension-t4 reps 1 and 2** (correct F, preserved F). Pivot tables and pivot caches were dropped. The table range stays `A1:O3` instead of `A1:O4`.
- **pension-t4 rep 3** (correct F, preserved T, valid F). The table range was not extended (`A1:O3`, expected `A1:O4`). The file also has dangling pivot-cache relationships.
- **pivotrpt-t4 rep 1** (correct T, preserved F). The row is right. 2 of 5 pivot tables and 2 of 3 pivot caches were dropped, and merged cells on Sheet5 were lost.
- **pivotrpt-t5 reps 1 to 3** (correct F, preserved F). The chart series still reference `Sheet2` after the tab was renamed to `Yearly Totals`.
- **sparkext-t5 reps 1 and 3** (correct T, preserved F). The new sheet is right. Data validations, a filter-database defined name and sparklines were removed. Each run has 5 losses.
- **sparkext-t5 rep 2** (correct F, preserved F, valid F). The VLOOKUP gives `#N/A`, a new error cell. The expected Sheet2 value is absent. The filter-database name and sparkline group are lost.
- **structrefs-t4 reps 1 to 3** (correct F, preserved F). The defined names `state` and `states` were not extended to cover the new row. Conditional-formatting blocks were removed (6 to 10 losses).

`safexl verify original output`, run afterwards as an independent receipt (the model never had it): "damaged" 18, "rewritten" 5, "invalid" 1. The 5 "rewritten" files are the 5 passes (kyc reps 1 to 3, pivotrpt-t4 reps 2 and 3); "invalid" is pension rep 3. On this round the receipt and the grader agree on all 24 files.

## Differences from the 2026-09-29 arm-A method

- Isolated: each session ran in a fresh `/tmp` directory with only its input workbook; in 09-29, 11 of the 24 runs had their cwd inside the repository the benchmark lives in. The `--settings` overlay now also denies Read/Edit/Bash on the benchmark tree.
- The contamination detector is on and reports 0.
- The runner has had path-default and comment-only edits since (no logic change affecting arm A).
- The operator hint, model, task set, reps, grader and `--setting-sources user,local` are the same.
- Run-to-run variation at n = 3 per task is large; the per-task pass counts differ from 09-29 in only one task (pivotrpt-t4, 3/3 to 2/3).
