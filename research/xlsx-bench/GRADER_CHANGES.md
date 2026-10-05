# Grader and task corrections made during the full run

Every change below was triggered by a specific graded result during the recorded run, found by
reading the actual output files before trusting the score. Each one fixes a flaw in the benchmark
itself, applies to both arms identically, and is in git history with its trigger. Re-grading all
finished runs after each change is how every reported number was produced.

| # | Trigger | Flaw | Fix | Commit |
|---|---|---|---|---|
| 1 | pilot: every output "missing" | prompts don't say "save as", so Claude edits in place; the grader only looked for new files | output = edited in-place copy when its hash changed | ef2b29b |
| 2 | recording verify3 arm B | a dangling calcChain relationship (Excel repair trigger) passed validity | validity checks every internal relationship target | 81fb1d8 |
| 3 | budget-t5 rename, both arms | preservation matched features by sheet NAME, so a correct rename read as "removed" | compare across renames via the worksheet part | 9380fd5 |
| 4 | benefits-t2, arm A (arm B later did the same) | task fixed the new column at M; the prompt names no column; Claude chose EA to keep the macros' fixed column reads intact | position-independent `column_values` | 9380fd5 |
| 5 | benefits-t5, both arms | task fixed summary values at B1/B2; both arms used a header row | position-independent `sheet_contains` (live formula required) | 46fdd37 |
| 6 | structrefs-t1, both arms | LibreOffice returns #VALUE! for Excel-only Table[#All] lookups even in the original | such cells graded UNVERIFIABLE (disclosed); edited input checked directly | 46fdd37 |
| 7 | structrefs-t2, arm A | formula check required the literal "M3"; Claude used the table's structured reference | `contains_any` accepts the structured ref | (this commit) |
| 8 | pivotrpt-t4, arm A (arm B identical) | a pivot cache's source growing over the inserted row (A1:D13 -> A1:D14, correct) read as "missing" | cache identity = type + sheet + name; a SHRINKING range is still a loss | (this commit) |
| 9 | structrefs-t4, arm A | run.json lists every file Claude touched, incl. XML parts it unzipped into x/; the grader took one as the output | only a real .xlsx/.xlsm counts as output | (this commit) |
| 10 | pivotrpt-t5 / parking-t5 tasks | chart checks only knew val_ref; tasks also assert cat_ref / name_ref, and the grader crashed | all three ref kinds supported; a grader error on one run is recorded, never stops the run | (this commit) |

| 11 | pension-t2, arm A | conditional formatting extended over the new table column (A2:G2 -> A2:H2, correct) read as "removed" | a CF/DV block whose range grows to a superset is preserved; shrinking is still a loss | (this commit) |
| 12 | parking-t1sub, arm A | `sheet_names` demanded the new sheet LAST; the prompt names no position and Claude put 'Pilot KPIs' before three empty leftover sheets | an ADDED sheet may go anywhere; existing sheets must keep their order; renames still need the exact order | (this commit) |
| 13 | parking-t1sub, arm A | `formula_contains` demanded the source sheet's name in the answering cell itself; Claude computed =B6/B5 with B5/B6 pulling from Summary_report | the link may run through same-sheet helper cells (up to 3 hops); a typed-in chain still fails | (this commit) |
| 14 | sparkext-t2, both arms | task flaw: the row holds FOUR date cells (G1, H1, I1 = J1), so "the two dates in the row" is ambiguous; both arms used I1-H1 = 364, the task wanted G1-H1 = 50 | `column_values.any_of`: any pair of the distinct dates passes (50 / 364 / 414), sign-insensitive | (this commit) |
| 15 | parking-t2sub, both arms | task flaw: "total abandons" was pinned to Detailed_report!S8 (first data row); both arms used Summary_report!E33, the Summe (total) row under 'Abbrüche (gesamt)', same value 0 | `formula_contains` may list acceptable sources; this value accepts either report sheet | (this commit) |
| 16 | Haiku round, kyc-t4 (both arms) | a defined name rewritten as 'KYC'!$1:$1 (same reference, optional quotes) read as "changed" | defined names compare equal up to needless sheet-name quotes and TRUE()/FALSE() vs TRUE/FALSE; needed quotes still matter | (this commit) |
| 17 | Qwen3-Coder-Next round (local GPU), kyc-t4 B1 / pivotrpt-t4 B3 / sparkext-t5 B2 | run.json's file list (a harness listing, intermediates first) outranked the workbook itself: sessions that built `output.xlsx` / `temp.xlsx`, then moved the final file onto the original name, were graded on the intermediate. Also Haiku hint A3 pension-t4, graded on a leftover `.bak` | an edited-in-place original is the output; run.json's list and new files count only when the original is untouched. Re-graded haiku-2026-09-29 (plain: 0 of 16 cells changed; hint 3-rep: 2 of 48 changed, A preserved 9/24 -> 8/24, no pass changed) and the Qwen round (7 of 48 changed, B pass 8/23 -> 11/23). The 09-26 rounds (in bundles) were not re-graded. Known limit: a session that saves its answer to a new file AND also changes the original is graded on the original | (this commit) |

Runner fix (not grading): a bare `429` pattern matched a 429-byte file in an `unzip -l` listing and stopped the run as a false "rate limit" (c. 75 min idle). 429 now needs rate-limit context.
