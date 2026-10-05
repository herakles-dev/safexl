# SafeXL

SafeXL lets an AI agent edit Excel workbooks and leave everything it doesn't touch as it was. It changes the file's XML directly instead of loading and re-saving the whole workbook. Charts, pivots, macros, data validation and the rest come out byte-identical unless an edit has to change them. I wrote it for people whose agents work on spreadsheets that other people depend on.

It ships as a CLI (`safexl`), an MCP server (`safexl-mcp`) and a Claude Code plugin. A hosted version is on Apify: https://apify.com/herakles-dev/safexl-workbook

## What I measured

The benchmark is 8 hard tasks on real workbooks, 3 runs each, so 24 runs per arm. A deterministic grader scores them. "Intact" means the grader's preserved check passed. "Passed" means the result was correct, preserved and valid.

This is not a held-out test. I built SafeXL against these 8 tasks, and I wrote the plugin's guidance after watching models fail on them. The guidance holds none of the answers, and a test in the repo checks that. Expect less on tasks I haven't seen.

Claude Haiku 4.5 with Anthropic's xlsx skill, told in one line to use it (2026-10-05): 6 of 24 intact, 5 of 24 passed.

Claude Haiku 4.5 with the SafeXL plugin 0.2 and no operator instruction (2026-10-05): 22 of 24 intact, 21 of 24 passed. The 24 runs cost $1.862 in total, which is $0.078 a run and $0.089 per passing run. Those are API list prices as the CLI reported them. The runs used a Claude subscription.

The two arms are not the same setup. The skill arm had Anthropic's skill, a one-line hint and no hooks. The plugin arm had the plugin with its guard hook on and Anthropic's skill blocked. Its start-up message tells the model to use safexl and never to save with openpyxl or pandas. So the model got a nudge, and it was not a blind test. The arms ran as separate batches on the same day.

The three runs that failed. Once the model wrote a lookup formula that evaluates to #N/A. Twice it inserted a whole row where the task needed an insert limited to some columns, and that moved two defined names.

One miss belongs to SafeXL itself. In my own check after the runs, `safexl verify` said "intact" for all 24 outputs, including the two that moved defined names. You can't repeat that check from this repo, because the output files aren't in it. Version 0.2.0, which these runs used, did not flag a reference that moved. Version 0.2.1 lists every defined name and table range that changed. I ran it again on the recorded outputs: the two failed runs now name the two moved names, `dropdown` and `droptable`, along with two names the task meant to grow. The verdict is still "intact", because the receipt reports the change and cannot know whether it was wanted. The guard hook had nothing to act on in any of the 24 runs.

I also ran other models through SafeXL's MCP tools only. I did not run any of them without SafeXL. These rows show what each model managed with the tools. They do not show what the tools added.

The table covers 7 of the 8 tasks, 21 runs a model. The system prompt of that round held one task's own formula as an example, so I left that task out. The round's summary quotes the lines.

| Model | Passed | Cost per run |
|---|---|---|
| DeepSeek V4 Flash (OpenRouter) | 21 of 21 | about $0.011, about 29 s |
| Qwen 3.8 27B (free tier) | 18 of 21 | not billed |
| Nemotron 3 Super 120B (free tier) | 17 of 21 | not billed |
| Laguna S 2.1 (free tier) | 16 of 21 | not billed |
| Mistral Small 3.2 | 13 of 21 | $0.0030 |
| GPT-OSS 120B | 12 of 21 | $0.0021 |
| Claude Haiku 4.5, same harness | 6 of 7 (one repetition, partial) | $0.0933 |

The prompt's other guidance was written after watching models fail on these tasks, so these rows are in-sample too. Four more models ran only 5 to 7 times and passed 2 to 4 of those. They are in that round's summary. Some rows include API failures and empty replies.

Twenty-four runs per arm on 8 tasks is a small sample, and the benchmark is mine.

## Check it yourself

```bash
git clone https://github.com/herakles-dev/safexl
cd safexl
pip install -e 'packages/safexl[scan,test]'
python3 research/xlsx-bench/fetch_corpus.py
python3 research/xlsx-bench/grade.py --tasks research/xlsx-bench/tasks_hard8.json --runs <runs dir> --summary
```

`fetch_corpus.py` downloads the 14 test workbooks from their upstream sources and checks their hashes. None of them are in this repo. Grading needs LibreOffice (`soffice`) on the PATH. Running a round needs the `claude` CLI.

The grades are in `research/xlsx-bench/results/`. Each round has its own folder. The raw runs are in `research/xlsx-bench/recordings/`. You can read every run and re-run the benchmark. You can't re-grade my exact files, because the output workbooks of my runs are not in the repo. Local paths, my user name and the subscription tier are rewritten in the recordings.

## Install

SafeXL is not on PyPI. Install it from the git URL:

```bash
pip install "git+https://github.com/herakles-dev/safexl#subdirectory=packages/safexl"
command -v safexl
```

For Claude Code, add the plugin:

```bash
claude plugin marketplace add herakles-dev/claude-plugins
claude plugin install safexl@herakles-plugins
```

Or add only the MCP server:

```bash
claude mcp add safexl -- safexl-mcp
```

## What it cannot do

It can't delete rows or columns. It can't insert columns. It can't create charts, pivots or formatting. It doesn't recalculate formulas, it flags the workbook to recalculate on open. Worksheets with a namespace prefix (.NET OpenXML SDK files) get refused. Output was opened in LibreOffice. It has never been opened in desktop Excel.

## What's in the repo

- `packages/safexl/` holds the engine, the CLI, the MCP server and the Claude Code plugin in `claude-plugin/`.
- `research/xlsx-bench/` holds the benchmark: tasks, grader, runner, `fetch_corpus.py`, and the recordings and results of three rounds.
- `research/safexl-moat/fixtures/` holds a manifest and one small workbook of mine.

## Licence

MIT. The third-party test workbooks stay under their upstream licences. See `research/xlsx-bench/THIRD_PARTY.md`.
