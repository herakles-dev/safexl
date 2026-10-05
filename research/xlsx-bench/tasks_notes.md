# xlsx-bench hard-mode task suite — notes

40 tasks across the 8 hard-mode corpus workbooks, 5 per workbook. Every expected value in
`tasks.json` was derived from a LibreOffice headless recalculation of the *original* file
(`soffice --headless --calc --convert-to xlsx`) read back with openpyxl `data_only=True`,
plus the task's own arithmetic computed in Python (see the generator notes inline in each
task's `notes` field for the exact derivation). Corpus files were never modified — only
read (openpyxl `data_only=False`/`True` on separate LibreOffice-recalculated copies in a
scratch directory, never saved back over the originals).

Two schema gaps surfaced repeatedly and are called out per-workbook below rather than
silently worked around: **pivot tables/caches** and **hyperlinks** have no dedicated key
in `TASK_SCHEMA.md`'s `expect` object (only `cf_sqrefs`, `dv_sqrefs`, `tables`,
`chart_series`, `defined_names`, `sheet_names`, `cells`/`text_cells`/`formulas` do). Tasks
that ripple through a pivot cache's source range or a hyperlink's target text can only be
checked for gross survival via the generic `preserve` count-identity rule, not for the
deep "did the reference text update correctly" property. This is noted explicitly in the
affected tasks' `notes` fields.

## simple-monthly-budget.xlsx

The whole hard-mode value here is packed into one sheet: a chart whose series range
(`$E$9:$F$9`) sits *above* the expense table it doesn't reference, a data-bar CF stored in
both classic and x14 form on the same cell, two Excel Tables (`tblIncome`, `tblExpenses`)
feeding two defined names by `SUM(...[Amount])`, and 5 merged-cell blocks in the header
area. All 5 tiers apply organically (no substitution): raising one expense (t1) moves the
same two defined-name-driven cells the chart doesn't touch structurally; the % column
(t2) grows `tblExpenses`; the 3% bulk raise (t3) touches all 13 expense rows; inserting a
Parking row (t4) grows the table from `B10:C23` to `B10:C24` while the CF/chart/defined
names (all anchored at rows 4-9, above the table) stay byte-for-byte the same *ref*, only
their *computed value* moves; and the single-sheet rename (t5) is the only sheet_op shape
this workbook supports (no second sheet exists for a cross-sheet summary), rippling into
chart1's val_ref, which is the one formula in the file that textually embeds the sheet
name.

## benefits-contribution-calculator.xlsm

All 5 tiers are native. The "calculator" itself (`MV Calculator` sheet) is sheet-protected
with a legacy password hash we cannot recover, is gated by 8 data-validation blocks, and
is really driven by VBA macros that don't run under a plain recalculation — so every task
here deliberately avoids it and instead targets the two genuinely formula-driven,
unprotected areas: the `Enrollment Restrictions` reconciliation (real `SUM()` formulas
over literal exclusion counts) for t1 (fix one exclusion figure) and t4 (insert a new
exclusion line, which requires the `SUM` range to auto-expand), and the three large
`MV Cont. Table - *` continuance tables (84 literal data rows each, real actuarial data)
for t2 (a %-of-total column) and t3 (trending column D forward by the exact 6.5% rate the
workbook's own User Guide documents). t5 adds a `Summary` sheet with cross-sheet formulas
because this workbook has zero pre-existing cross-sheet formulas to rename-and-ripple —
a rename here would have nothing to break.

## structured-refs-lookup-tool.xlsx

All 5 tiers are native — this is the richest workbook in the corpus (3798-row and 729-row
tables, 3882 cross-sheet formula references, 426 merged cells, an external link, 7 list
data validations, a chart). The trick was finding the *small*, tractable pieces buried
inside the huge composite-key VLOOKUP machinery: `Table_Query_from_RDS2` (a clean 14-row,
per-year nationwide summary table on `query2`) supports t1 (fix one year's `total1`,
rippling through `Profile`'s `VLOOKUP`-driven C8/D8/C24), t2 (a YoY-%-change column,
growing the table), and t3 (a 5%-bulk raise on `total2`, rippling through row8's E/D/F/H).
t4 exploits a genuine spreadsheet-debt pattern: the `dropdowns` sheet keeps **two**
independent, row-offset copies of the state list (`state`=A2:A53, `states`=C2:C54,
offset by one row because of an extra "All States" header entry) — a real single
full-row insert cannot correctly add "Guam" to both at once, so the task requires two
independent column-scoped inserts, each growing its own defined name by exactly one cell.
t5 renames `List`, whose name is textually embedded in all 3798 `Query!A2:A3799` running-
counter formulas (confirmed via a targeted grep of `cross_sheet_formulas` — 100% of the
3798 `List!`-referencing formulas live on `Query`, none elsewhere) plus the sheet's own
Print_Area/Print_Titles defined names; we spot-check 3 of the 3798 rows (first, middle,
last) rather than asserting all of them.

## pivot-chart-report.xlsx

**Substitution: tier1 → an additional tier5.** This workbook has *zero* live formulas
anywhere (verified by grepping every worksheet XML part for `<f>` elements across all 6
sheets: 0 hits) — its 5 pivot tables are frozen cached snapshots, not formulas, and
Sheet1's raw data is plain literals. Tier1's own definition ("change an input that feeds
formulas, totals must move") has no live formula to move, so it's replaced with a second
tier5 task: adding a `Category Totals` sheet with real `SUM()` formulas over the raw data,
which *does* exercise live-formula creation even though none pre-existed. t2 (a Profit
column) and t3 (a 4% bulk raise on Cost) are native and land on the same plain Sheet1
range. t4 inserts a new `Games` row into that range, which is where the schema gap bites:
whether the 3 pivot caches' stored `source_ref` (`Sheet1!A1:D13`, a plain range not an
Excel Table) grows to `A1:D14` isn't independently assertable (no `pivot_caches` expect
key), so this task's checkable ripple is the shifted-row cell values, not the pivot
internals. t5 (native, not a substitute) renames `Sheet2`, which drives chart1's single
series — a small, clean, fully `chart_series`-assertable ripple, unlike a pivot-cache
rename would be.

## pension-hedging-stress-model.xlsx

**Substitution: tier3 → an additional tier5.** This is a genuinely near-empty actuarial
*template* (`TemplateWithTableSourcePivotTables.xlsx` upstream) — all 14 Excel Tables have
at most 2 data rows of placeholder zeros/blanks, so there is no real block of existing
numbers anywhere to bulk-adjust by a percentage; tier3 is replaced with a second sheet
rename. t1 sets `HedgeResults[NotionalsPerR1]` (a single-row table), which ripples through
the `Checks` sheet's real `Recommended Trades`/`Total error` formulas (`SUM` over
currently-blank ranges elsewhere, confirmed 0 via recalc) — a genuine, if thin, tier1.
t2 adds a formula column to `PerSchemeRequiredBonds` (also a value of 0 by construction,
since the template's one data row is unpopulated placeholder data — still a valid test of
table-structured-formula-column mechanics). t4 inserts a new bond-type row into
`AssetRandsPerPoint` (the one table with 2 real data rows), which correctly ripples a
workbook-scoped defined name's range. Both tier5 tasks (the tier3 substitute renaming
`Hedge detail - Nominal`, and the native tier5 renaming `Hedge detail - ILB`) exploit the
workbook's one real formula cluster: 5 `Checks` formulas each reference both
`Hedge detail - Nominal` and `Hedge detail - ILB` by name, so each rename independently
ripples 5 formulas.

## sparkline-external-reference.xlsx

**Substitutions: tier3 → tier5, tier4 → tier5** (this workbook ends up with 3 tier5 tasks
plus t1/t2). It's a genuinely minimal 2-sheet unit-test-style fixture: Sheet1 is a 2×2
grid (1 header row, 1 data row) and Sheet2 is a single data row spread across many
columns — there is no block of multiple existing values to bulk-adjust (tier3) and no
range with 2+ rows to insert a new row into the middle of (tier4), so both are replaced
with sheet-rename tasks. t1 edits Sheet2's one real numeric cell (`L1`, a calculated
factor), rippling through Sheet1's exact-match `VLOOKUP`; we deliberately avoid editing
the *lookup key* (`B2`) since Sheet2 has only one matching row and changing the key
without a matching new row would produce a `#N/A` — an ambiguity trap this task sidesteps
by design. t2 adds a days-between-dates column. The three tier5 tasks are kept genuinely
distinct: rename `Sheet1` (ripples the sheet's own `_FilterDatabase` defined name), rename
`Sheet2` (ripples Sheet1's VLOOKUP formula — note the sparkline's own source formula
points at an *external* workbook, `'[1]Contract Tail YLT'!B46:E46`, and is untouched by
either rename), and add a new `Lookup Check` summary sheet (the workbook's only
cross-sheet-formula-addition task).

## kyc-intake-data-dictionary.xlsx

**Substitutions: tier1 → tier5, tier3 → tier4** (giving this workbook tiers
[2, 4, 4, 5, 5]). This is a pure 30-sheet code dictionary (Code/Description pairs) with
*zero* formulas anywhere (grepped every sheet's XML for `<f>`: 0 hits) and no numeric
quantities to scale by a percentage (codes are ID-like strings, not amounts) — the exact
"pure dictionary with no numbers for tier 3" case `TASK_SCHEMA.md` itself anticipates, and
tier1 fails for the identical reason (no formula to make move). Both are replaced. The
workbook's real KYC intake form (`KYC` sheet) and 3 related sheets are protected with only
an unrecoverable legacy password hash (`B5EC`/SHA-512 hashes with no plaintext), so per
the brief's "give the password or skip it" rule, **no task in this entire 40-task suite
edits a protected sheet's cell content** (a global design choice, not unique to this
workbook). t2 adds a numeric-code-extraction column to the 131-row `Rejection Reasons`
list (a real `RIGHT()`/`VALUE()` text formula, not a business number, since that's the
only formula shape this dictionary supports). t4 (native) inserts Ladakh into
`State Master`, authentically placed right after Jammu and Kashmir (its real 2019 parent
state) with an explicit code to avoid any numbering ambiguity. The tier3→tier4 substitute
inserts a second, distinct code (`08`/"Re-KYC") into `KYC Update Type`, using a genuine
gap already present in that list's numbering (`07` then `11`). Both t1's substitute
(renaming `Occupation`) and the native t5 (renaming `Country Master`, the richer target —
3 hyperlinks vs. Occupation's 1) ripple into internal hyperlink targets stored on the
*protected* `KYC` sheet — legal and checkable-for-gross-survival because renaming a sheet
is a workbook-structure operation Excel permits without unprotecting (no
workbook-structure lock is set in this file) and updating a hyperlink's target text is an
automatic side effect of the rename, not a direct edit to a locked cell — but the deep
"did the target text update" property hits the hyperlinks schema gap noted above.

## parking-study-report.xlsm

**Substitutions: tier1, tier2, tier3, tier4 → tier5 (×4)**, leaving this workbook's whole
suite as 5 sheet_op tasks. All 6 content-bearing sheets (`CCSInfo`, `General`,
`Detailed_report`, `Summary_report`, `Detailed_graph`, `Summary_graph`) carry per-sheet
protection with only a legacy password hash (`EBF6`, same hash reused across all 6 —
plausibly one shared password, but still unrecoverable), and the workbook's only
unprotected sheets (`Tabelle1`/`2`/`3`) are empty leftover defaults with zero cells of
data. That leaves no unprotected editable data anywhere for a value/formula-column/bulk/
insert-row task. What Excel *does* allow without unprotecting: renaming a sheet (governed
by workbook-structure protection, which this file doesn't set) and reading protected
sheets from formulas on a brand-new, unprotected sheet (protection blocks editing, not
being read elsewhere). So all 5 tasks are sheet_op: two new summary sheets that read
`Summary_report`/`Detailed_report` cells via cross-sheet formulas (substituting for
tier1/tier2), and three sheet renames — `General` (117 of 213 defined names anchor here,
the single largest share; substituting for tier3), `Summary_report` (drives chart2's 5
series plus 19 defined names; substituting for tier4), and `Detailed_report` (drives
chart1's 5 series plus 62 defined names; kept as the native, non-substituted tier5 since
it's the richest fully `chart_series`-assertable ripple in the workbook). The `General`
rename's 117-defined-name ripple is asserted only via `sheet_names` in `expect` since
listing all 117 renamed formula strings would be impractical; this is a deliberate
partial-verification design choice, recorded here rather than silently under-tested.
