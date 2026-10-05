# xlsx-bench task schema (shared contract for tasks.json and grade.py)

`tasks.json` is a JSON list. One entry per task:

```json
{
  "id": "budget-t4-insert-row",
  "workbook": "simple-monthly-budget.xlsx",
  "tier": 4,
  "kind": "value | formula_column | bulk | insert_row | sheet_op",
  "prompt": "What a plain user would type. No tool names, no hints about preservation.",
  "expect": {
    "cells":        [{"sheet": "S", "ref": "C12", "value": 1250.0, "tol": 0.005}],
    "text_cells":   [{"sheet": "S", "ref": "B8", "value": "Parking"}],
    "column_values": [{"sheet": "S", "rows": {"5": 0.1297, "88": 0.0000055}, "tol": 0.0005}],
    "formulas":     [{"sheet": "S", "ref": "D9", "contains": "SUM("}],
    "tables":       [{"name": "tblExpenses", "ref": "E4:F10"}],
    "chart_series": [{"part": "xl/charts/chart1.xml", "val_ref": "'Simple Monthly Budget'!$E$9:$F$9"}],
    "cf_sqrefs":    [{"sheet": "S", "sqref": "E4:E10"}],
    "dv_sqrefs":    [{"sheet": "S", "sqref": "C5:C12"}],
    "defined_names": {"TotalMonthlyExpenses": "SUM(tblExpenses[Amount])"},
    "sheet_names":  ["Summary", "Data"],
    "no_new_errors": true
  },
  "preserve": ["vba", "images", "charts", "pivot_tables", "pivot_caches", "comments",
               "hyperlinks", "protection", "sparklines", "external_links", "merged_cells",
               "conditional_formatting", "data_validations", "defined_names", "tables"],
  "notes": "Why this task is hard, and how each expected value was derived."
}
```

## Rules

- **Values are checked after a LibreOffice headless recalculation of the output file**, never
  from cached values. `cells` are numbers, compared within `tol`; `text_cells` are exact strings.
- Every expected value must be derived from the ORIGINAL workbook plus the task arithmetic, and
  `notes` shows the derivation (e.g. "old total 3,450 from LibreOffice recalc of the original,
  plus the inserted row's 250 = 3,700").
- `preserve` lists the feature classes (keys of `inventory.py`'s output) that must survive. A
  feature "survives" when every source instance is still present with the same identity: same
  vbaProject.bin sha256, same media sha256s, same number of charts/pivots/comments/hyperlinks/
  sparkline groups/external links, same protection flags, and CF/DV/defined names/tables either
  unchanged or changed exactly as `expect` says (a row insert legitimately moves a sqref or a
  table ref; that's asserted in `expect`, and every other instance must be unchanged).
- Structural tasks (tier 4 insert_row, tier 5 sheet_op) must say in `expect` what a correct
  result looks like everywhere the change ripples: totals, table refs, chart series refs, CF and
  DV sqrefs, defined names, cross-sheet formulas.
- `no_new_errors`: after recalc, no cell holds #REF!, #NAME?, #VALUE!, #DIV/0! or #N/A unless the
  same cell held it in the original's recalc.

- `column_values` is for a NEW column in a plain range, where the prompt doesn't say which column
  to use: some column must hold these values at these rows after recalc, and must not have held
  them in the original. Use fixed `cells` only when the position is implied (e.g. the new column
  has to extend an Excel Table).

## Tiers

1. **value**: change an input that feeds formulas (totals must move).
2. **formula_column**: add a calculated column; if the data is an Excel Table, the table grows.
3. **bulk**: change a block of values (e.g. raise a column 3%, rounded to cents).
4. **insert_row**: insert a new data row in the middle of a range or table.
5. **sheet_op**: rename a sheet other things reference, or add a summary sheet with
   cross-sheet formulas.

## Grade (per run)

- `correct`: every `expect` check passes.
- `preserved`: every `preserve` class survives.
- `valid`: opens in LibreOffice headless, well-formed XML, `[Content_Types].xml` covers every
  part, `no_new_errors` holds.
- `pass` = correct AND preserved AND valid. Report each separately; a tool that preserves
  everything but doesn't make the edit is not a pass.
