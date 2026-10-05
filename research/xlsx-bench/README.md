# xlsx-bench — two-arm Excel-editing benchmark

Head-to-head comparison of Claude Code editing an Excel workbook two ways:

- **Arm A**: Claude Code with Anthropic's `xlsx` skill (`anthropic-skills:xlsx`).
  No SafeXL plugin, and the SafeXL venv is taken off `PATH`.
- **Arm B**: Claude Code with the SafeXL plugin loaded from
  `packages/safexl/claude-plugin` (its skill, the `safexl` CLI on `PATH`, and its
  MCP server, whose tools are not in `--allowedTools`). The `xlsx` skill is
  denied. Plugin hooks are off unless `XLSX_BENCH_PLUGIN_HOOKS=1`.

Both arms get the same prompt and the same `--allowedTools`
(`Skill Read Write Edit Bash Glob Grep`). The published rounds set `XLSX_BENCH_MODEL=claude-haiku-4-5-20251001`, and the plugin round also set `XLSX_BENCH_PLUGIN_HOOKS=1`. Without the variable the model is `claude-opus-5-5`. Each run gets a fresh copy of the workbook.

Two drivers use the same arm logic (`runner.py`: `build_command`, `build_env`):

- `runner.py` runs `claude -p` headlessly and writes `run.json` plus the raw stream.
- `record.py` runs the session interactively in tmux under asciinema, in a temp
  directory, and writes `run.json`, `session.cast` and the session transcript.

`runner.py` and `record.py` use the standard library only (Python 3.11).
`asciinema` and `tmux` are needed for `record.py`.

## Getting the workbooks

No test workbook is in this repo. `python3 research/xlsx-bench/fetch_corpus.py`
downloads them from their pinned upstream sources (Apache POI, ClosedXML) and
checks their hashes. Licences are in `THIRD_PARTY.md`.

## Running it

```bash
# smoke test: 2 tasks x 2 arms x 1 rep = 4 runs, ~20 min ceiling each
python3 research/xlsx-bench/runner.py --arms A,B --reps 1

# the 36-task run set (tasks_run.json, also record.py's default) instead of
# the 2-task smoke_tasks.json
python3 research/xlsx-bench/runner.py \
  --tasks research/xlsx-bench/tasks_run.json \
  --arms A,B --reps 3 \
  --out research/xlsx-bench/runs/corpus

# run inside a detached tmux session (survives you closing the terminal)
python3 research/xlsx-bench/runner.py --tmux
tmux attach -t xlsx-bench

# check progress without running anything
python3 research/xlsx-bench/runner.py --status
```

Flags:

| Flag | Default | Meaning |
|---|---|---|
| `--tasks PATH` | `smoke_tasks.json` | JSON list of `{id, workbook, prompt}` |
| `--arms A,B` | `A,B` | which arm(s) to run |
| `--reps N` | `1` | repetitions per (task, arm) |
| `--out DIR` | `research/xlsx-bench/runs` | where run folders go |
| `--timeout SECONDS` | `1200` (20 min) | per-run wall-clock ceiling |
| `--claude-bin PATH` | `~/.local/bin/claude` | which `claude` binary to invoke |
| `--block-skill-method` | `disallowed` | how arm B blocks the xlsx skill — see below |
| `--tmux` | off | relaunch itself in a detached tmux session `xlsx-bench` |
| `--status` | off | print done/total per arm and exit (no runs) |

Runs are **resumable**: a `(task, arm, rep)` cell is skipped if its
`run.json` already exists and parses as JSON. `run.json` is written
atomically (temp file + rename), so a killed-mid-write run is never
mistaken for done.

## Output layout

What a run directory holds when you run the kit:

```
<out>/<task_id>/<arm>/<rep>/
  work/           the run's workbook and whatever the model wrote beside it
  run.json        parsed summary (see fields below)
```

`runner.py` also writes `events.jsonl` (the raw `--output-format stream-json`
stream) and `stderr.txt`, and runs claude with `work/` as its cwd. `record.py`
runs claude in a temp directory, copies it back to `work/`, and also writes
`session.cast` (asciinema), `session.jsonl` (Claude Code's own session
transcript) and `manifest.json` (sha256 of the input workbook, the outputs,
the cast and the transcript). `transcript.py` renders a run directory into
`transcript.md` and `transcript.html`. The model-agnostic harness in
`agnostic/` writes `run.json` and `transcript.json`.

The published recordings under `recordings/` hold less. Each run has
`run.json` and `session.cast` plus `transcript.md` (Claude Code rounds), or
`run.json` and `transcript.json` (the model-agnostic round). They have no
`work/`, no workbooks, no `events.jsonl`, no `stderr.txt`, no `session.jsonl`
and no `manifest.json`. There is no `runs/` directory in this repo. Grades are in
`results/`.

Commit hashes quoted in the summaries under `results/` refer to my private development history. Each summary also gives a content hash where it matters.

### run.json fields

A run.json holds these fields: `model`, `total_cost_usd`, `num_turns`, `result_subtype`, `is_error`, `final_text` (a truncated excerpt), `wall_seconds`, `exit_code` and `timed_out`.

Tool use is in `tools_called` (unique names, first-seen order) and `tool_call_count`. `skills_invoked` lists every Skill name the model tried, whether or not it was blocked. `skills_succeeded` and `skills_denied` split that list. `skill_denial_evidence` holds the raw tool_result text for denied skill calls. `safexl_used` is true when any Bash command matched `/safexl/i`. `safexl_commands` lists those commands.

`output_files` lists new files that appeared under `work/` during the run. `rate_limited` and `rate_limit_evidence` say whether the run hit a rate limit. `invocation` holds the exact command, the cwd, the PATH prepend and the block method. It also records that `ANTHROPIC_API_KEY` was unset and CLAUDE.md was disabled for this run.

`total_cost_usd` and `num_turns` are read from the `result` event. The runner accepts `total_cost_usd` and `num_turns`, or `cost_usd` and `turns`, because both spellings turned up across CLI versions. `model` comes from the `system`/`init` event.

## Arm definitions and how isolation is enforced

Both arms get the identical, single `--allowedTools` string:

```
Skill Read Write Edit Bash Glob Grep
```

**Arm A (default environment).** No `--plugin-dir`. The SafeXL venv's `bin/`
directory is stripped out of `PATH` even if it were already present
(`build_env`), so `safexl` does not resolve. SafeXL's plugin is not
installed globally (checked against `~/.claude/plugins/installed_plugins.json`
before writing this), so there's no marketplace-install path for it to leak
in either.

**Arm B.** Adds:
- `--plugin-dir <repo>/packages/safexl/claude-plugin` (override: `XLSX_BENCH_PLUGIN_DIR`)
  (loads the SafeXL skill + its `.mcp.json` for this session only; the
  directory is read, never written).
- The SafeXL venv `bin/` prepended to `PATH`
  (`<repo>/packages/safexl/.venv/bin`; override: `XLSX_BENCH_SAFEXL_BIN`), so the
  `safexl` CLI resolves for Bash calls. (The plugin also registers an MCP
  server, but MCP tool names aren't in `--allowedTools`, so the model is
  steered onto the same Bash-CLI path the SafeXL skill's own instructions
  teach — which is also what `safexl_used` scans for.)
- A block on the Anthropic xlsx skill (see next section) — same
  `--allowedTools` as arm A, but the one specific skill is denied.

**Both arms**, in the subprocess environment (`build_env`):
- `ANTHROPIC_API_KEY` is stripped even if inherited — auth is OAuth
  (subscription plan) only, and the runs use a Claude subscription and no API
  key. `--bare` (which would force API-key auth) is not used, and
  because `--bare` also disables CLAUDE.md auto-discovery,
  which we need to disable a different way (below) without touching skills
  or plugins.
- `CLAUDE_CODE_DISABLE_CLAUDE_MDS=1`. The fresh `work/` folder lives inside
  this git checkout, so without this, Claude Code would auto-discover the
  repo's and the user's `CLAUDE.md` (which names SafeXL and holds unrelated
  private project notes) and inject them into the
  system prompt. That's two problems at once: it tips off arm A that
  SafeXL exists (contaminating the "naive baseline" the comparison depends
  on), and it risks leaking internal business context into transcripts
  that go to customers. This env var is real — confirmed by disassembling
  the installed CLI binary (`strings` on
  `~/.local/share/claude/versions/2.1.283`): it's the same flag `--safe-mode`
  sets internally (`CLAUDE_CODE_SAFE_MODE=1` + `CLAUDE_CODE_DISABLE_CLAUDE_MDS=1`
  together), but set alone it *only* suppresses CLAUDE.md/AGENTS.md/`.claude/rules`
  loading — unlike `--safe-mode`, it leaves skills, plugins and hooks alone,
  which is what we need since arm B's whole point is a plugin being loaded.
  No `~/.claude` or other global config was edited to get any of this —
  it's all per-subprocess environment, set by `runner.py` itself.

None of this required editing `~/.claude` or any global config, per the
task constraint — everything above is either a CLI flag or a variable in
the subprocess's own `os.environ` copy.

### Blocking the xlsx skill for arm B

The runner uses `--disallowedTools "Skill(anthropic-skills:xlsx)"`. This is `--block-skill-method disallowed`, the default. The `Skill(<name>)` rule grammar is the same one Claude Code uses for its built-in allow rule `Skill(anthropic-skills:*)`, and deny rules take the same syntax. The rule denies that one skill. `Skill` stays in `--allowedTools`, so arm B can still invoke the `safexl` skill through the same tool.

If a future CLI version stops honoring `--disallowedTools`, `--block-skill-method settings` is the fallback. It passes the same rule string through `--settings`:

```json
{"permissions": {"deny": ["Skill(anthropic-skills:xlsx)"]}}
```

### Isolation of the recorded sessions

Recorded sessions (`record.py`) run with `--strict-mcp-config`, so no project or
user MCP server loads. They use `--permission-mode dontAsk`, a temp-dir cwd,
all hooks off except the plugin's own when `XLSX_BENCH_PLUGIN_HOOKS=1` (the plugin round), the user's other plugins off, `CLAUDE.md` discovery off, and
deny rules for the bench tree (`interactive_command`, `clean_room_settings`,
`make_isolated_workdir`). The headless `runner.py` does not pass
`--strict-mcp-config` and runs in `<cell>/work`, so its raw `events.jsonl` lists
every tool and skill the CLI saw. Use `run.json`, not `events.jsonl`, as the
summary of a headless run.

## Rate-limit handling

Runs execute strictly sequentially. After each run, `runner.py` scans
`events.jsonl` and `stderr.txt` for usage/rate-limit signals (`usage limit`,
`rate limit`, `429`, `session limit`, `quota exceeded`, …) in the session's
own status output only: CLI stderr, error `result` events, and synthetic
API-error assistant messages (`isApiErrorMessage`, `quotaLimits` rejected).
Text that came from a tool result (a file the model read) is never
scanned; that produced a false stop on 2026-09-29. On a match it
writes `runs/STOPPED.json` (reason + a text snippet around the match, which
usually includes any reset-time wording) and exits non-zero *before*
starting the next cell. Every run that already finished keeps its
`run.json` — nothing already-done is touched or re-run. `runner.py` refuses
to start a new batch while a `STOPPED.json` is present; remove it (after
reading it) once ready to resume.

## Session isolation and contamination (added 2026-09-29)

Sessions used to run with cwd `recordings/<round>/<task>/<arm>/<rep>/work`,
inside the repo next to every other run. Looking for the workbook with
`find`/`ls`, some sessions wandered into sibling runs and read another arm's
`manifest.json`, `run.json` or output workbook. Fixes in `record.py` /
`runner.py` (identical for both arms):

- **Temp cwd.** Each recording runs in a fresh `/tmp/xlsxb-<8 hex>/` holding
  only the task's input workbook. When the session ends, the whole dir is
  copied back to `<cell>/work` (so `grade.py` is unchanged) and the temp dir
  is removed. `run.json` records `invocation.cwd` (the temp path),
  `invocation.work_dir_copied_to` and `invocation.isolated`.
  `XLSX_BENCH_TMP_ROOT` overrides `/tmp`.
- **Deny rules.** `runner.clean_room_settings()` adds
  `Read(//…/research/xlsx-bench/**)`, `Edit(//…)` and `Bash(*xlsx-bench*)` to
  the `--settings` overlay. The SafeXL plugin dir and venv (under
  `packages/safexl`) stay usable. `Bash(*xlsx-bench*)` matches the command
  string, so it blocks commands that name the tree but not, say, a
  `find /` whose output lists it; the detector below covers that.
- **Contamination detector.** After each run the transcript is scanned
  (tool inputs, tool results, assistant text) for any path under
  `research/xlsx-bench/` or another run's `/tmp/xlsxb-*`; the verdict is
  stored as `run.json` `contamination: {contaminated, evidence[]}`. An
  attempt counts even when it was denied or returned nothing. Reading the
  SafeXL package, plugin/skill dirs or `~/.claude` is fine.
- **Grading.** `grade.py --summary --exclude-contaminated` lists contaminated
  runs and leaves them out of the rates. Per-run lines in the JSONL carry a
  `contaminated` flag either way.
- **Retro-scan.** `scan_contamination.py recordings/<round> [--write]`
  applies the detector to older runs (old layout: a session may name its own
  cell dir, and Claude Code's own per-cwd state dirs; anything else is
  flagged). Quarantine flagged reps by moving them under
  `<round>/_quarantine/` (`grade.py` and `record.py` ignore it).

## Testing

Use the virtualenv from `packages/safexl` (its README shows how to make it):

```bash
packages/safexl/.venv/bin/python -m pytest research/xlsx-bench/tests -q
```

Unit tests cover tasks-file parsing/validation, the arm A/B command and
environment builders, `stream-json` event parsing (model/cost/turns,
skill invocation/success/denial, safexl-via-Bash detection, rate-limit
detection), file-diffing, and resumability (`is_done` / `pending_cells` /
atomic `run.json` writes). No test invokes `claude` — that only happens via
a real `runner.py` run, which costs plan usage.
