---
name: safexl
description: Use for ANY edit to an existing .xlsx or .xlsm workbook - changing cells or formulas, inserting rows, renaming or adding a sheet, filling a table. Also use when a workbook was already re-saved with openpyxl or pandas and lost charts, macros or formatting. Do not save existing workbooks with openpyxl, pandas or any other library; safexl patches the workbook file directly so everything you do not touch stays byte-identical. Not needed for a brand-new workbook built from scratch.
---

# SafeXL: edit an existing workbook without wrecking it

Saving a workbook with openpyxl or pandas rebuilds the whole file and can silently drop charts, pivot tables,
macros, sparklines, data validation and defined names. Use the `safexl` command for every write.
A guard hook watches the workbooks in this project: if a save damages one, it is restored to its last
good version and you are told. Then redo the edit with safexl.

If `safexl` is not found, run `command -v safexl`; install with
`pipx install "git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl"`.

## The routine (do these in order)

1. LOOK. `safexl inspect FILE` for what the workbook contains. `safexl view FILE` for an overview,
   `safexl view FILE --sheet "Sheet" --range A1:H40` for exact cells, `safexl view FILE --find TEXT` to locate
   a label. It shows values, types (text is quoted, numbers are bare) and formulas. Never print sheets with
   python or read raw XML.
2. PLAN. Say the exact commands you will run: sheet, row number, cells, values, formulas, and why.
3. WRITE with safexl only. `in` is never modified; `out` is written only on success.
   - `safexl edit in.xlsx out.xlsx --set 'Sheet!B5=42' --text 'Sheet!A5=007' --formula 'Sheet!C5=B5*2'`;
     with `--sheet Data` as the default sheet, cells can be written as `--set 'B5=42'`
   - `safexl insert-rows in.xlsx out.xlsx --sheet Data --row 7 [--count 1]` (row 7 and below move down)
   - `safexl insert-cells in.xlsx out.xlsx --sheet Data --range B5:C6` inserts blank cells in that column
     span only and shifts just those columns down. Use it instead of `insert-rows` when other lists or
     tables sit side by side on the sheet and must not move.
   - `safexl rename-sheet in.xlsx out.xlsx "Old Name" "New Name"`
   - `safexl add-sheet in.xlsx out.xlsx "Notes"`
   - `safexl resize-table in.xlsx out.xlsx --table Name --ref A1:D20` resizes a table; `edit` already
     extends an adjacent table by default.
   Chain steps as `step1.xlsx`, `step2.xlsx`, then `mv stepN.xlsx FILE` so the result keeps the original name.
4. CHECK. `safexl verify FILE_BEFORE FILE_AFTER` (exit 0 = intact; 10 rewritten, 11 damaged, 12 invalid) and
   `safexl view` the cells you changed and the rows around them. Keep a copy of the original to compare.
5. REPAIR at most once if the check shows a mistake, then stop and report.

## Usage slips to avoid

- `--set` is for numbers, booleans and ordinary text. `--text` is only for codes that must stay text
  (`007`, `1E5`). A number written with `--text` is stored as text and SUM ignores it.
- Match the type of the neighbouring cells. In a lookup formula the key must have the same type as the
  lookup column: a number is not found in a column of text, and text is not found in a column of numbers.
- To add a line inside a list: `insert-rows` first, then `edit` the new row. Where neighbouring rows use a
  formula in a column, write the same formula pattern for the new row.
- A total just above an inserted row does not grow by itself. `=SUM(B2:B6)` with a row inserted at 7
  needs `--formula 'Sheet!B8=SUM(B2:B7)'`.
- To show figures that already exist elsewhere in the workbook, write formulas that point at those cells
  (`='Other Sheet'!F20`), not copied values.
- Formulas are not calculated by safexl; they compute when the file is next opened. Do not read a new
  formula's value back from the file.
- `insert-rows` moves whole rows; if other lists share those rows, use `insert-cells`. No column inserts or row deletes.

## When the guard acted

A message that starts with "SAFEXL GUARD" means a workbook was saved by another library. Either it was
rebuilt from the original with your cell edits kept (check it with `safexl view`), or it was restored to its
last good version and your change is gone. A damaged new copy is repaired the same way, or left for you
to redo: `cp original new`, then safexl commands. Do not repeat the same save. Make further changes only
with the commands above.

## A file that was already re-saved

Only for a workbook that was re-saved by another library and lost features (it does nothing, and writes no
output, if the file is already intact). If you still have the original: `safexl salvage ORIGINAL DAMAGED -o OUT` rebuilds the lost parts from the original. Then run
`safexl verify ORIGINAL OUT`.

## Safety

SafeXL never runs VBA. `.xlsm` macros come out byte-identical.
