# safexl (Claude Code plugin)

## What it does

Installing the plugin adds a guard hook, a skill and an MCP server for existing `.xlsx`/`.xlsm`
workbooks. You don't have to tell Claude to use it:

- **Guard hook** (`hooks/`). At session start it snapshots the workbooks in the project. After every tool
  call that can write files (Bash, Write, Edit, MultiEdit, NotebookEdit) it checks them. If another
  writer (openpyxl, pandas) re-saved a tracked workbook, even losslessly, the hook first runs `safexl salvage --force`:
  the model's own cell edits are re-applied to the snapshot original and the result is put in place if it
  verifies intact; Claude is told what was kept. If salvage cannot do it, the workbook is restored
  byte-identical from its snapshot and Claude is told to redo the edit with `safexl`. A new workbook that
  is a damaged copy of a tracked one (same sheets, saved by another library) is repaired the same way and
  the original is not touched; if it cannot be repaired it is left as is and Claude is told to
  `cp <original> <new>` and use `safexl`. A clean `safexl` edit is accepted and becomes the new snapshot.
  Any other new workbook is tracked silently from then on. A tracked file you move or delete is dropped
  from tracking, never resurrected.
- **Skill** (`skills/safexl/`). A short routine for small models: look, plan, write with `safexl`, verify,
  repair once; plus the common usage slips and `safexl salvage`.
- **MCP server** (`.mcp.json`). Wires `safexl-mcp` (command must be on PATH).

The hook is quiet: it prints nothing when no workbook is present, and does no work when nothing changed
(it compares file size, mtime, inode and ctime, and calls `safexl` only for a changed file).

Caps: a workbook over 50 MB is not tracked; at most 20 are tracked per session; discovery looks in the
project directory and 2 levels below it (hidden and dependency directories are skipped).

## Install

1. Put `safexl` on PATH. The hook calls the `safexl` command and does nothing useful without it
   (it tells you once per session if it is missing):
   ```bash
   pip install "git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl"
   command -v safexl
   ```
   Hooks need `python3` on PATH.
2. Install the plugin:
   ```bash
   claude plugin marketplace add herakles-dev/claude-plugins
   claude plugin install safexl@herakles-plugins
   ```
   Locally, for a trial: `claude --plugin-dir packages/safexl/claude-plugin`.

Validate without touching your real config:
`CLAUDE_CONFIG_DIR=$(mktemp -d) claude plugin validate packages/safexl/claude-plugin`

## Turn the guard off

To turn the guard off for one shell or session, start Claude with `SAFEXL_GUARD=off claude`. To turn it off for one project, put an empty file named `.safexl-guard-off` in its directory. Set `SAFEXL_GUARD_MODE=restore` to get restore only, with no salvage and no copy repair. To turn it off everywhere, run `claude plugin disable safexl`, which disables the plugin.

## What it writes, and where

- Snapshots and a small `state.json`, under `$CLAUDE_PLUGIN_DATA/sessions/<session id>/` (a temp
  directory, `safexl-guard/` under the system temp dir, if that variable is unset). Never inside your
  project. `$CLAUDE_PLUGIN_DATA` is removed when the plugin is uninstalled.
- In your project the hook changes exactly one thing: a damaged tracked workbook, or a damaged new copy of
  one, is replaced by its repaired (salvaged) or restored version. It deletes nothing and makes no network calls.
- A hook error is swallowed (exit 0); a tool call is never blocked.
