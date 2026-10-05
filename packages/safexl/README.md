# SafeXL

SafeXL edits `.xlsx`/`.xlsm` workbooks in headless environments (no Excel, no
LibreOffice required to run) by patching the underlying OOXML zip directly
instead of loading and re-serializing the whole file. Every part it doesn't
touch — charts, drawings, images, pivot tables/caches, data validations,
conditional formatting (including extLst/x14 blocks), and VBA
(`vbaProject.bin`) — comes out **byte-identical** to the source. SafeXL never
executes or rewrites VBA.

## Status

Engine core, a CLI (`edit`, structural edits, `resize-table`, `inspect`, `diff`,
`verify`, `view`, `salvage`, `guard`), and an MCP server (`safexl-mcp`) so
agents can call the same operations. Not on PyPI. Output has been opened and
recalculated in LibreOffice headless. It has not been opened in desktop Excel.

## Install

SafeXL is not on PyPI. Install it from the git URL:

```bash
pip install "git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl"
```

This installs two commands: `safexl` (CLI) and `safexl-mcp` (MCP server).
The only runtime dependency is the official MCP Python SDK (`mcp>=2.2.0,<3`);
the engine itself is stdlib-only.

## Use it from an agent (MCP)

Install `safexl-mcp` first (see Install above). Then point your client at the
installed executable.

### Claude Code

```bash
# stdio, runs on your machine, takes local file paths
claude mcp add safexl -- safexl-mcp

# if it isn't on PATH, use the executable's full path
claude mcp add safexl -- /path/to/venv/bin/safexl-mcp
```

Once it is on PyPI, `claude mcp add safexl -- uvx safexl-mcp` will also work.

### Claude Desktop

Add this to `claude_desktop_config.json` (Settings → Developer → Edit
Config), then restart Claude Desktop:

```json
{
  "mcpServers": {
    "safexl": {
      "command": "/path/to/venv/bin/safexl-mcp"
    }
  }
}
```

Claude Desktop doesn't always inherit your shell's `PATH`, so use the
executable's absolute path (`which safexl-mcp`). Once it is on PyPI, you can
use `"command": "uvx", "args": ["safexl-mcp"]` instead.

### HTTP (streamable HTTP)

```bash
safexl-mcp --http --port 8000          # serves http://127.0.0.1:8000/mcp
claude mcp add --transport http safexl http://127.0.0.1:8000/mcp
```

### Tools

| Tool | What it does | stdio input | HTTP input |
|---|---|---|---|
| `safexl_inspect` | Read-only report: sheets, declared vs actual used range, formula counts, defined names, charts, images, pivot tables, validation, conditional formatting, VBA presence. | `path` | `file_base64` |
| `safexl_view` | Read cells from a window (`cell_range`) or by text search (`find`), with type, value and formula. Capped by `max_cells`; `next_range` gives the next window. Read-only. | `path`, `sheet`, `cell_range`, `find`, `max_cells`, `formulas` | same, with `file_base64` |
| `safexl_edit` | Write values and/or formulas into an existing workbook. Every part it doesn't touch is copied byte-for-byte; the result lists which parts changed. | `path`, `changes`, `output_path` (default `<name>.edited.<ext>`), `overwrite` (default `false`), `extend_tables` | `file_base64`, `changes`, `extend_tables`; returns `file_base64` |
| `safexl_insert_rows` | Insert blank rows. Rows at or below shift down and formula references follow, within the sheet and across sheets. | `path`, `sheet_name`, `row`, `count`, `output_path`, `overwrite` | `file_base64`, `sheet_name`, `row`, `count` |
| `safexl_insert_cells` | Insert blank cells over a range and shift only those columns down, for sheets with other lists side by side. Only `shift="down"` is supported. | `path`, `sheet_name`, `cell_range`, `shift`, `output_path`, `overwrite` | `file_base64`, `sheet_name`, `cell_range`, `shift` |
| `safexl_rename_sheet` | Rename a worksheet and update every formula reference to it. | `path`, `old_name`, `new_name`, `output_path`, `overwrite` | `file_base64`, `old_name`, `new_name` |
| `safexl_add_sheet` | Add a new, empty worksheet. | `path`, `sheet_name`, `output_path`, `overwrite` | `file_base64`, `sheet_name` |
| `safexl_resize_table` | Change the range of an Excel table. Only the table definition changes. | `path`, `table`, `ref`, `sheet_name`, `output_path`, `overwrite` | `file_base64`, `table`, `ref`, `sheet_name` |
| `safexl_diff` | Before/after cell diff plus a SHA-256 diff of every internal part. | `before_path`, `after_path`, `max_cells` | `before_base64`, `after_base64`, `max_cells` |
| `safexl_verify` | Compare an edited workbook with its original and return a receipt: verdict (intact, rewritten, damaged or invalid), changed parts, `summary`, `added_risk`, `cell_changes` and `reference_changes`. Read-only. | `before_path`, `after_path` | `before_base64`, `after_base64` |
| `safexl_salvage` | Rebuild a workbook that another library re-saved: re-apply its cell edits to the original. Status is salvaged, partial, unsupported or nothing_to_do. `force` rebuilds even when the file looks intact. | `original_path`, `damaged_path`, `output_path`, `overwrite`, `force` | `original_base64`, `damaged_base64`, `force` |

`changes` is a list of `{"sheet": "Sheet1", "cell": "B7", "value": 42}` or
`{"sheet": "Sheet1", "cell": "B8", "formula": "=B7*2"}`, up to 1,000 per
call. Macros are never executed by any tool. Formulas that `safexl_edit`
writes compute when the file is next opened in Excel or LibreOffice (see the
caveat under Usage).

Every tool returns JSON. Success is `{"ok": true, ...}`. Failure is
`{"ok": false, "error": {"code": "...", "message": "...", "details": {...}}}`
with MCP's `isError` set, so an agent can branch on `error.code`
(`file_not_found`, `sheet_not_found` with `available_sheets`,
`invalid_change`, `extension_mismatch`, `output_exists`, `path_not_allowed`,
`file_too_large`, ...).

### HTTP mode security

In stdio mode the server runs as you on your own machine, so it reads and
writes local paths, the same as the CLI. HTTP mode is meant for callers you
don't control, so the rules are stricter:

- **Base64 content only.** No tool takes a path. Anything that looks like a
  path or URL (`/…`, `~/…`, `C:\…`, `file:`, `https://…`, `*.xlsx`) is
  rejected before decoding. The server never reads its own disk on a
  caller's behalf and never fetches a URL.
- **20 MB cap** per uploaded file (decoded). Uploads that expand past 256 MB
  or hold more than 10,000 zip entries are rejected as zip bombs.
- **Nothing is kept.** Each call works in a private temp directory that's
  deleted when the call returns. The edited workbook comes back as base64,
  and error messages never include server paths.
- **Binds to 127.0.0.1 by default**, with the MCP SDK's DNS-rebinding
  protection on (a foreign `Host` or `Origin` header is refused). There is
  no authentication. If you bind to another interface (`--host 0.0.0.0`),
  put an authenticating reverse proxy in front of it.

### Concurrency

- **HTTP transport, different workbooks.** Each file tool call works in its own private temp directory. Concurrent requests for different workbooks never see each other's data. A test in the suite checks this under real concurrent load. At most `HTTP_MAX_CONCURRENT_JOBS` (4) file jobs run at once. A request beyond that waits up to `HTTP_BUSY_TIMEOUT_S` (120 s) for a free slot. Then it gets a `server_busy` error.
- **`edit()` (stdio, the Python API and `safexl_edit`), same destination.** `edit()` never writes `dst` directly. It writes a private temp file next to `dst`, checks its structure, then moves it into place. The final move is a rename on the same filesystem. Two concurrent calls with the same `dst` leave one caller's complete, valid output, never a mix. No temp file is left behind. Nothing locks between the two calls, so the last rename wins. If you need one writer at a time for a workbook, serialize the edits yourself. SafeXL does not queue or lock concurrent writers.

## Usage (CLI and Python)

```bash
safexl edit in.xlsx out.xlsx \
  --set 'Sheet1!A1=42' \
  --formula 'Sheet1!B1=A1*2'

# Structural edits (every reference follows; see safexl_mcp/structural.py)
safexl insert-rows in.xlsx out.xlsx --sheet Sheet1 --row 5 [--count 1]
safexl insert-cells in.xlsx out.xlsx --sheet Sheet1 --range B5:C6
safexl rename-sheet in.xlsx out.xlsx "Old Name" "New Name"
safexl add-sheet in.xlsx out.xlsx Notes
```

`edit --text 'Sheet1!A1=0451'` writes a string without number coercion.

`edit --sheet Data --set B2=5` is `--set Data!B2=5`: `--sheet` is the default sheet for
references written without a `Sheet!` prefix, and an explicit prefix still wins.

### Command table

Result codes of 10 and up never overlap the error codes. Exit code 1 is an error such as a missing file. Exit code 2 is a usage error.

| Command | What it does | Exit codes |
|---|---|---|
| `edit` | Write values/formulas. Grows an Excel table when a written cell sits next to it; `--no-extend-tables` turns that off. | 0 ok, 1 error, 2 usage |
| `insert-rows`, `rename-sheet`, `add-sheet` | Structural edits; every reference follows. | 0 ok, 1 error, 2 usage |
| `insert-cells` | Insert blank cells in a column span only; those columns shift down. Use it instead of `insert-rows` when the sheet has other lists or tables side by side that must not move. A reference that would be torn is refused with the reason (exit 1). | 0 ok, 1 error, 2 usage |
| `resize-table` | Set the range of an Excel table. | 0 ok, 1 error, 2 usage |
| `inspect`, `diff` | Read-only report and before/after diff. | 0 ok, 1 error |
| `verify BEFORE AFTER` | Receipt for an edit: intact, rewritten, damaged or invalid. | 0 intact, 10 rewritten, 11 damaged, 12 invalid, 13 intact but risky addition (`--strict` only), 14 intact but a defined name or table range changed (`--strict` only) |
| `view PATH` | Compact cell listing, one line per cell. | 0 ok, 1 error |
| `salvage ORIGINAL DAMAGED -o OUT` | Re-apply the edits found in a damaged copy onto the original. | 0 salvaged or nothing to do, 10 partial, 11 unsupported |
| `guard snapshot` / `guard check` | Snapshot workbooks, then detect and undo damage by another writer. | `check`: 0 nothing wrong, 10 a file was flagged |

**Changed in this release:** `edit` now extends an Excel table when you write a cell directly beside it, the way Excel does when you type next to a table. `--no-extend-tables` turns that off and writes the cells only. For each table it extends, `edit` prints a line such as `table Sales extended A1:C20 -> A1:D20 (new column "Margin")`. A column with an empty header cell gets a name, and the line says so. The other commands keep their arguments, output and exit codes.

```bash
# add two blank rows to the list in B:C only; a list in column E stays where it is
safexl insert-cells in.xlsx out.xlsx --sheet Sheet1 --range B5:C6
```

```bash
safexl resize-table in.xlsx out.xlsx --table Sales --ref A1:D20 [--sheet Data]
```

### Verify

```bash
safexl verify before.xlsx after.xlsx [--json] [--strict]
```

Compares an edited workbook with its original. The verdict is `intact` (only
SafeXL-style changes), `rewritten` (another tool re-saved it, so untouched parts
are no longer byte-identical, but no tracked feature was lost), `damaged` (a
chart, image, table, macro project, validation, conditional format or similar
is gone) or `invalid` (not a readable workbook). It lists each lost feature with
the sheet and range where known, and counts the parts unchanged, changed, added
and removed. Exit code 0 intact, 10 rewritten, 11 damaged, 12 invalid; a missing input file is an error (1).

The text output also prints a one-sentence summary after the verdict (for example
"No tracked feature was lost and untouched parts are byte-identical; 1 cell changed on
Sheet1; nothing risky was added."), one `added: <risk>: <detail>` line per risky
addition (an external link, a VBA project, a hidden sheet, sheet protection, a
web-service or DDE formula and similar), and a line per changed sheet such as
`Sheet1: 3 cells changed, 1 row inserted` (`(positional)` when the sheets could not
be aligned). `intact` covers tracked features and untouched parts only; it does not
mean no cell changed, so read the summary. An added risk does not change the verdict
or the exit code, but with `--strict` an intact workbook that gained a risk exits 13
("intact, but something risky was added").

The receipt also lists `reference_changes`: every defined name and table range whose
text differs between the two files. Each entry has `kind` (`defined_name` or `table`),
`name`, `scope`, `before` and `after`; an added or removed name has `null` for the
missing side. A listed change is a fact, not a verdict. A table that grew because you
added a row is listed too, and the receipt cannot know which changes you wanted.
Sheet renames are not listed. With `--strict`, an intact workbook with no added risk
but a non-empty `reference_changes` exits 14; 13 wins if both apply. Data-validation
ranges and chart series references are not compared.

The MCP `safexl_verify` tool returns the same `summary`, `added_risk`,
`cell_changes` (capped at 50 sheets, with `cell_changes_truncated` when cut) and
`reference_changes`.

Every write is read back before the output is kept. If the read-back does not match,
the edit fails naming the cell (CLI exit 1; MCP error `write_not_confirmed`) and no
output file is left. A worksheet written with a namespace prefix on every element (the
.NET OpenXML SDK) is refused the same way (CLI exit 1; MCP error `unsupported_workbook`).

### Salvage

```bash
safexl salvage original.xlsx damaged.xlsx -o salvaged.xlsx [--force] [--json]
```

When another tool (pandas, openpyxl) re-saved a workbook and dropped features,
`salvage` finds the cell edits in the damaged copy and re-applies them to the
original with the SafeXL engine, so the original's charts, tables and macros
stay as they were. It prints the status (`salvaged`, `partial`, `unsupported` or
`nothing_to_do`), how many cells were applied and skipped with the reason for
each skip, and any structural change it detected (renamed, added or removed
sheets, inserted or deleted rows, shifted columns). Exit code 0 salvaged or
nothing to do, 10 partial (output written), 11 unsupported (no output written); a missing
input file is an error (1).

A file that is intact, or was re-saved without losing any feature, is reported as
`nothing_to_do` and no output is written. `--force` also rebuilds such a file, so
its untouched parts are byte-identical to the original again.

### Guard

```bash
safexl guard snapshot --dir .safexl-snap model.xlsx report.xlsx
# ... something else writes to the workbooks ...
safexl guard check --dir .safexl-snap --restore
```

`snapshot` copies each workbook into the snapshot directory (a directory guard
owns, holding byte copies and fingerprints). `check` compares the workbooks with
their snapshots and prints one line per flagged file; with `--restore` it puts
the snapshot back byte-identical. A SafeXL edit is accepted as the new snapshot.
By default a file that was merely rewritten is restored too; `--lenient` leaves
it in place and restores only damaged, invalid or missing files. `--json` prints
the events as JSON. Exit code 0 when nothing was wrong, 10 when a file was
flagged (restored or not); `snapshot` exits 1 when a path is missing.

### View

```bash
safexl view workbook.xlsx --sheet Budget --range A1:D20
safexl view workbook.xlsx --find "total" [--max-cells 400] [--json]
```

Prints one line per non-empty cell, such as `B12  n  1250.5  =SUM(B2:B11)`:
the reference, a type letter (n number, s string, b boolean, e error, d date),
the value, and the formula text. Strings are quoted so numbers stored as text
stay visible. A formula cell shows the value Excel last cached; nothing is
evaluated. A cell in a shared formula shows its own formula, the master's
shifted by the row and column offset. When the listing is cut at `--max-cells`,
the last line gives the `--range` to ask for next; the next window does not
repeat cells already shown. With `--find`, all sheets are searched unless
`--sheet` is given.

```python
from safexl_mcp import edit

edit(
    "in.xlsx",
    "out.xlsx",
    [
        {"sheet": "Sheet1", "cell": "A1", "value": 42},
        {"sheet": "Sheet1", "cell": "B1", "formula": "A1*2"},
    ],
)
```

**Formula caching caveat:** SafeXL does not evaluate formulas itself. A
written formula cell has no cached `<v>` — SafeXL sets
`fullCalcOnLoad="1"` so Excel/LibreOffice recompute it on open. A library
that reads the file without opening it in a calculating application first
will see the formula cell as blank. See `safexl_mcp/engine.py`'s module
docstring for the full explanation and rollback promise.

### Inspect (read-only, never executes macros)

```bash
safexl inspect workbook.xlsm --json
```

Reports per-sheet declared `<dimension>` vs. actual used range, formula
count, defined names, and a feature inventory: charts, images/drawings,
pivot tables + caches, data validation, conditional formatting (native and
extLst/x14), and VBA presence + `vbaProject.bin` size. VBA source is
MS-OVBA-compressed, so stdlib-only mode reports presence/size only; install
the optional `scan` extra for an AutoOpen/Shell/CreateObject/etc. keyword
scan via `oletools` (static analysis only, never executes the macro):

```bash
pip install "safexl-mcp[scan] @ git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl"
```

### Diff (read-only)

```bash
safexl diff before.xlsx after.xlsx --json
```

Per sheet: changed/added/removed cells (value or formula). Plus every OOXML
zip part added, removed, or changed by sha256 — the machine-checkable half
of "nothing else changed."

## Development

The MCP server needs `mcp>=2.2`, so run the tests in a virtualenv rather
than against a system Python that may carry an older `mcp`:

```bash
python3 -m venv .venv
.venv/bin/pip install -e '.[scan,test]'
.venv/bin/python -m pytest -q --cov=safexl_mcp
```

## License

MIT. See [LICENSE](LICENSE).
