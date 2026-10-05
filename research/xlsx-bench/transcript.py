#!/usr/bin/env python3
"""Render one xlsx-bench run folder into transcript.md + transcript.html.

Handles both flavors of run folder produced in this directory:

  - a headless runner.py run: ``run.json`` + ``events.jsonl`` (raw
    ``--output-format stream-json``, includes a ``system``/``init`` event
    that dumps this box's *entire* ambient skill/agent/MCP catalog -- see
    runner.py README's "Known limitation" section);
  - an interactive record.py recording: ``run.json`` + ``session.jsonl``
    (Claude Code's own session transcript, copied by record.py) + optionally
    ``session.cast`` + ``manifest.json``.

Both formats share the same ``assistant``/``user`` message shapes (confirmed
empirically -- see record.py's module docstring), so one conversation
extractor (``build_conversation``) handles both; only the surrounding
metadata line types differ (``system``/``result`` for headless,
``queue-operation``/``attachment``/``atis-latch``/``last-prompt``/
``cost-state`` for a recording) and both are simply skipped -- this is what
strips the tool/skill/agent catalog: the ``system``/``init`` event is never
walked for conversation content, only peeked at for its ``model`` field.

Redaction: any absolute filesystem path is rewritten to ``<work>/rel/path``
if it falls under the run's own work directory (kept, since it's the file
this task actually produced) or to ``<path>`` otherwise (stripped). The work
directory is read from the run's own ``invocation.cwd`` field when present
(the exact path Claude Code actually ran in, at record time) rather than
recomputed from this run folder's current location, so rendering still
redacts correctly after the folder is merged from a worktree into main (a
different absolute path than where the run happened).

Output is self-contained (inline CSS, no external requests, no JS) and
readable at phone width.
"""

from __future__ import annotations

import argparse
import html as html_mod
import json
import re
import sys
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import runner  # noqa: E402
import record  # noqa: E402  (reused for sha256_file only -- no live-run code)

LONG_INPUT_THRESHOLD = 500
RESULT_LINE_LIMIT = 40

# Matches an absolute filesystem path: a "/" not itself preceded by a
# path-safe character (so "work/SimpleMacro.xlsm" is left alone -- that "/"
# is a relative-path separator, not the start of an absolute path), followed
# by path-safe segments. See module docstring for the character-set
# limitation (no spaces/URLs handling).
_ABS_PATH_RE = re.compile(r"(?<![A-Za-z0-9_.\-])/(?:[A-Za-z0-9_.\-]+/)*[A-Za-z0-9_.\-]+")

# The OS username running this box, e.g. from `ls -la` output ("drwxr-xr-x 2
# alice alice ... ") -- a bare word, so it survives _ABS_PATH_RE
# untouched (found by inspecting a real rendered transcript: a plain `ls -la`
# tool result leaked the username three times with zero absolute paths
# involved). Computed from the home directory rather than hardcoded so this
# works on any machine this ever runs on.
_USERNAME = Path.home().name
_USERNAME_RE = re.compile(r"\b" + re.escape(_USERNAME) + r"\b") if _USERNAME else None


# ---------------------------------------------------------------------------
# Redaction
# ---------------------------------------------------------------------------


def redact_text(text: Optional[str], work_dir_abs: Optional[str]) -> str:
    if not text:
        return ""
    work_dir_norm = (work_dir_abs or "").rstrip("/")

    def repl(m: re.Match) -> str:
        p = m.group(0)
        if work_dir_norm and (p == work_dir_norm or p.startswith(work_dir_norm + "/")):
            return "<work>" + p[len(work_dir_norm):]
        return "<path>"

    out = _ABS_PATH_RE.sub(repl, text)
    if _USERNAME_RE is not None:
        out = _USERNAME_RE.sub("<user>", out)
    return out


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------


def detect_kind(run_dir: Path) -> str:
    if (run_dir / "session.jsonl").exists():
        return "recording"
    if (run_dir / "events.jsonl").exists():
        return "headless"
    raise FileNotFoundError(f"no session.jsonl or events.jsonl in {run_dir}")


def load_entries(run_dir: Path, kind: str) -> list[dict]:
    fname = "session.jsonl" if kind == "recording" else "events.jsonl"
    path = run_dir / fname
    if not path.exists():
        return []
    return list(runner.iter_json_lines(path.read_text(errors="replace")))


def extract_model(entries: list[dict], run_json: dict) -> Optional[str]:
    for e in entries:
        if e.get("type") == "system" and e.get("subtype") == "init":
            m = e.get("model")
            if m:
                return m
        if e.get("type") == "assistant":
            m = (e.get("message") or {}).get("model")
            if m:
                return m
    return run_json.get("model")


def work_dir_from_run_json(run_dir: Path, run_json: dict) -> str:
    cwd = (run_json.get("invocation") or {}).get("cwd")
    if cwd:
        return cwd
    return str((run_dir / "work").resolve())


# ---------------------------------------------------------------------------
# Conversation extraction (shared by both run-folder flavors)
# ---------------------------------------------------------------------------


def build_conversation(entries: list[dict]) -> list[dict]:
    """Ordered [{"kind": "text"|"tool_call"|"tool_result", ...}]. Deliberately
    skips "thinking" blocks (not asked for; keeps phone-width transcripts
    short) and every non-assistant/user line type (system/init's catalog,
    result, queue-operation, attachment, atis-latch, last-prompt, cost-state
    -- none of it is "what Claude did" on this task)."""
    items: list[dict] = []
    tool_names_by_id: dict[str, str] = {}

    for e in entries:
        etype = e.get("type")
        if etype == "assistant":
            message = e.get("message") or {}
            for block in message.get("content") or []:
                if not isinstance(block, dict):
                    continue
                bt = block.get("type")
                if bt == "text":
                    text = block.get("text")
                    if text:
                        items.append({"kind": "text", "text": text})
                elif bt == "tool_use":
                    name = block.get("name")
                    tool_id = block.get("id")
                    tool_input = block.get("input") or {}
                    if tool_id:
                        tool_names_by_id[tool_id] = name
                    items.append({"kind": "tool_call", "name": name, "input": tool_input, "id": tool_id})
        elif etype == "user":
            message = e.get("message") or {}
            content = message.get("content")
            if not isinstance(content, list):
                continue
            for block in content:
                if not isinstance(block, dict) or block.get("type") != "tool_result":
                    continue
                tool_id = block.get("tool_use_id")
                text = runner._tool_result_text(block)
                items.append(
                    {
                        "kind": "tool_result",
                        "id": tool_id,
                        "name": tool_names_by_id.get(tool_id),
                        "text": text,
                        "is_error": bool(block.get("is_error")),
                    }
                )
    return items


def extract_final_message(items: list[dict], run_json: dict) -> Optional[str]:
    for item in reversed(items):
        if item["kind"] == "text":
            return item["text"]
    return run_json.get("final_text")


# ---------------------------------------------------------------------------
# Manifest (prefer record.py's manifest.json when present; else compute)
# ---------------------------------------------------------------------------


def gather_manifest(run_dir: Path, run_json: dict, kind: str) -> dict:
    existing = None
    mpath = run_dir / "manifest.json"
    if mpath.exists():
        try:
            existing = json.loads(mpath.read_text())
        except (OSError, json.JSONDecodeError):
            existing = None

    workbook_entry = (existing or {}).get("workbook")
    if not workbook_entry:
        # No manifest.json (a headless runner.py run never writes one). The
        # only trustworthy hash we can compute after the fact is of the copy
        # that's still sitting in this run's own work/ -- NOT the corpus
        # fixture at its repo path, which can (and, empirically on this
        # multi-agent box, does) change over time and would silently hash
        # the wrong bytes. sha256_before is left null: we have no record of
        # what the file looked like before the run, only whatever is there
        # now (already-edited, if the task edited in place).
        wb_rel = run_json.get("workbook")
        if wb_rel:
            wb_path = run_dir / "work" / Path(wb_rel).name
            if wb_path.exists():
                workbook_entry = {
                    "path": Path(wb_rel).name,
                    "sha256_before": None,
                    "sha256_after": record.sha256_file(wb_path),
                    "modified_in_place": None,
                }

    output_entries = (existing or {}).get("output_files")
    if not output_entries:
        output_entries = []
        wdir = run_dir / "work"
        for rel in run_json.get("output_files") or []:
            p = wdir / rel
            if p.is_file():
                output_entries.append({"path": rel, "sha256": record.sha256_file(p)})

    transcript_entry = (existing or {}).get("session_jsonl")
    if not transcript_entry:
        tpath = run_dir / ("session.jsonl" if kind == "recording" else "events.jsonl")
        if tpath.exists():
            transcript_entry = {"path": tpath.name, "sha256": record.sha256_file(tpath)}

    cast_entry = (existing or {}).get("session_cast")

    return {
        "workbook": workbook_entry,
        "output_files": output_entries,
        "transcript_file": transcript_entry,
        "cast_file": cast_entry,
    }


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------


def workbook_hash_lines(workbook_entry: Optional[dict]) -> list[str]:
    """Plain-text (no markdown/HTML markup) lines describing the workbook's
    manifest hash(es). Shared by both renderers so the before/after/
    modified-in-place/unavailable cases are described identically in
    markdown and HTML.

    A task that edits the workbook in place (the common case in this
    corpus -- no "save as" to a new filename) leaves only one file on disk,
    which IS the delivered result; sha256_after is what a customer should
    check it against. sha256_before exists only to prove what the run
    started from. When there's no manifest.json at all (a headless
    runner.py run), sha256_before is unavailable and that's said plainly
    rather than silently omitted."""
    if not workbook_entry:
        return []
    name = workbook_entry.get("path", "?")
    before = workbook_entry.get("sha256_before")
    after = workbook_entry.get("sha256_after")
    if before and after:
        if workbook_entry.get("modified_in_place"):
            return [
                f"workbook `{name}` before: sha256:{before}",
                f"workbook `{name}` after (delivered): sha256:{after} -- edited in place",
            ]
        return [f"workbook `{name}` (unchanged by this task): sha256:{after}"]
    if after:
        return [f"workbook `{name}` (current): sha256:{after} -- before-edit hash unavailable"]
    if before:
        return [f"workbook `{name}` (before only, run did not finish): sha256:{before}"]
    return []


def format_input_json(tool_input: dict) -> str:
    return json.dumps(tool_input, indent=2, sort_keys=False, ensure_ascii=False)


def trim_result_text(text: str, limit: int = RESULT_LINE_LIMIT) -> tuple[str, int]:
    lines = (text or "").splitlines()
    if len(lines) <= limit:
        return text, 0
    return "\n".join(lines[:limit]), len(lines) - limit


# ---------------------------------------------------------------------------
# Markdown rendering
# ---------------------------------------------------------------------------


def render_markdown(
    run_json: dict, kind: str, model: Optional[str], manifest: dict,
    items: list[dict], final_message: Optional[str], work_dir_abs: str,
) -> str:
    task_id = run_json.get("task_id", "?")
    arm = run_json.get("arm", "?")
    rep = run_json.get("rep", "?")
    lines: list[str] = []
    lines.append(f"# xlsx-bench transcript -- {task_id} / arm {arm} / rep {rep}")
    lines.append("")
    lines.append(f"**Model:** {model or 'unknown'}  ")
    lines.append(f"**Run type:** {'interactive recording' if kind == 'recording' else 'headless run'}  ")
    if run_json.get("wall_seconds") is not None:
        lines.append(f"**Wall time:** {run_json['wall_seconds']}s  ")
    lines.append("")
    lines.append("**Manifest hashes:**")
    for wb_line in workbook_hash_lines(manifest.get("workbook")):
        lines.append(f"- {wb_line}")
    for of in manifest.get("output_files") or []:
        lines.append(f"- output `{of['path']}`: `sha256:{of['sha256']}`")
    if manifest.get("transcript_file"):
        lines.append(
            f"- transcript `{manifest['transcript_file']['path']}`: "
            f"`sha256:{manifest['transcript_file']['sha256']}`"
        )
    if manifest.get("cast_file"):
        lines.append(
            f"- recording `{manifest['cast_file']['path']}`: "
            f"`sha256:{manifest['cast_file']['sha256']}`"
        )
    lines.append("")
    lines.append("## Prompt")
    lines.append("")
    lines.append(redact_text(run_json.get("prompt", ""), work_dir_abs))
    lines.append("")
    lines.append("## Conversation")
    lines.append("")
    for item in items:
        if item["kind"] == "text":
            lines.append(redact_text(item["text"], work_dir_abs))
            lines.append("")
        elif item["kind"] == "tool_call":
            redacted = redact_text(format_input_json(item["input"]), work_dir_abs)
            lines.append(f"**Tool call: `{item['name']}`**")
            lines.append("")
            if len(redacted) > LONG_INPUT_THRESHOLD:
                lines.append(f"<details><summary>input ({len(redacted)} chars, click to expand)</summary>")
                lines.append("")
                lines.append("```json")
                lines.append(redacted)
                lines.append("```")
                lines.append("</details>")
            else:
                lines.append("```json")
                lines.append(redacted)
                lines.append("```")
            lines.append("")
        elif item["kind"] == "tool_result":
            trimmed, more = trim_result_text(redact_text(item["text"], work_dir_abs))
            label = "Tool result" + (" (error)" if item.get("is_error") else "")
            lines.append(f"*{label}*")
            lines.append("")
            lines.append("```")
            lines.append(trimmed)
            if more:
                lines.append(f"... ({more} more lines truncated)")
            lines.append("```")
            lines.append("")
    lines.append("## Final message")
    lines.append("")
    lines.append(redact_text(final_message, work_dir_abs) or "(no final message found)")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

_CSS = """
:root { color-scheme: light dark; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
       max-width: 760px; margin: 0 auto; padding: 16px; line-height: 1.45;
       background: #ffffff; color: #1a1a1a; }
@media (prefers-color-scheme: dark) {
  body { background: #14161a; color: #e6e6e6; }
  code, pre { background: #22252b !important; color: #e6e6e6 !important; }
  .call, .result { border-color: #33363d !important; }
  .result.error { border-color: #7a2f2f !important; }
  a { color: #8ab4f8; }
}
h1 { font-size: 1.25rem; }
h2 { font-size: 1.05rem; margin-top: 1.6em; border-bottom: 1px solid #ddd; padding-bottom: 0.2em; }
dl.meta { display: grid; grid-template-columns: auto 1fr; gap: 2px 10px; font-size: 0.9rem; }
dl.meta dt { font-weight: 600; }
ul.hashes { font-size: 0.82rem; word-break: break-all; padding-left: 1.2em; }
pre { white-space: pre-wrap; word-break: break-word; background: #f4f4f4; padding: 8px 10px;
      border-radius: 6px; font-size: 0.82rem; overflow-x: auto; }
code { background: #f0f0f0; padding: 1px 4px; border-radius: 4px; font-size: 0.85em; }
.call, .result { border: 1px solid #ddd; border-radius: 8px; padding: 6px 10px; margin: 10px 0; }
.result.error { border-color: #c66; }
.result p, .call p { margin: 2px 0 6px; font-weight: 600; font-size: 0.85rem; }
.truncated { font-size: 0.8rem; opacity: 0.7; }
.text { margin: 10px 0; white-space: pre-wrap; }
details > summary { cursor: pointer; font-weight: 600; font-size: 0.85rem; }
"""


def render_html(
    run_json: dict, kind: str, model: Optional[str], manifest: dict,
    items: list[dict], final_message: Optional[str], work_dir_abs: str,
) -> str:
    task_id = run_json.get("task_id", "?")
    arm = run_json.get("arm", "?")
    rep = run_json.get("rep", "?")
    title = f"xlsx-bench: {task_id} / arm {arm} / rep {rep}"

    def esc(s: Optional[str]) -> str:
        return html_mod.escape(s or "")

    parts: list[str] = []
    parts.append("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">")
    parts.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    parts.append(f"<title>{esc(title)}</title><style>{_CSS}</style></head><body>")
    parts.append(f"<h1>{esc(title)}</h1>")
    parts.append("<dl class=\"meta\">")
    parts.append(f"<dt>Model</dt><dd>{esc(model or 'unknown')}</dd>")
    parts.append(f"<dt>Run type</dt><dd>{'interactive recording' if kind == 'recording' else 'headless run'}</dd>")
    if run_json.get("wall_seconds") is not None:
        parts.append(f"<dt>Wall time</dt><dd>{esc(str(run_json['wall_seconds']))}s</dd>")
    parts.append("</dl>")
    parts.append("<h2>Manifest hashes</h2><ul class=\"hashes\">")
    for wb_line in workbook_hash_lines(manifest.get("workbook")):
        parts.append(f"<li>{esc(wb_line)}</li>")
    for of in manifest.get("output_files") or []:
        parts.append(f"<li>output <code>{esc(of['path'])}</code>: <code>{esc(of['sha256'])}</code></li>")
    if manifest.get("transcript_file"):
        parts.append(
            f"<li>transcript <code>{esc(manifest['transcript_file']['path'])}</code>: "
            f"<code>{esc(manifest['transcript_file']['sha256'])}</code></li>"
        )
    if manifest.get("cast_file"):
        parts.append(
            f"<li>recording <code>{esc(manifest['cast_file']['path'])}</code>: "
            f"<code>{esc(manifest['cast_file']['sha256'])}</code></li>"
        )
    parts.append("</ul>")
    parts.append("<h2>Prompt</h2>")
    parts.append(f"<pre class=\"prompt\">{esc(redact_text(run_json.get('prompt', ''), work_dir_abs))}</pre>")
    parts.append("<h2>Conversation</h2>")
    for item in items:
        if item["kind"] == "text":
            parts.append(f"<div class=\"text\">{esc(redact_text(item['text'], work_dir_abs))}</div>")
        elif item["kind"] == "tool_call":
            redacted = redact_text(format_input_json(item["input"]), work_dir_abs)
            name = esc(item["name"])
            if len(redacted) > LONG_INPUT_THRESHOLD:
                parts.append(
                    f"<details class=\"call\"><summary>Tool call: <code>{name}</code> "
                    f"({len(redacted)} chars, tap to expand)</summary>"
                    f"<pre>{esc(redacted)}</pre></details>"
                )
            else:
                parts.append(
                    f"<div class=\"call\"><p>Tool call: <code>{name}</code></p>"
                    f"<pre>{esc(redacted)}</pre></div>"
                )
        elif item["kind"] == "tool_result":
            trimmed, more = trim_result_text(redact_text(item["text"], work_dir_abs))
            cls = "result error" if item.get("is_error") else "result"
            label = "Tool result (error)" if item.get("is_error") else "Tool result"
            more_html = f"<p class=\"truncated\">... ({more} more lines truncated)</p>" if more else ""
            parts.append(f"<div class=\"{cls}\"><p>{label}</p><pre>{esc(trimmed)}</pre>{more_html}</div>")
    parts.append("<h2>Final message</h2>")
    parts.append(
        f"<pre class=\"final\">{esc(redact_text(final_message, work_dir_abs) or '(no final message found)')}</pre>"
    )
    parts.append("</body></html>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def render_run_dir(run_dir: Path) -> tuple[Path, Path]:
    run_json_path = run_dir / "run.json"
    if not run_json_path.exists():
        raise FileNotFoundError(f"no run.json in {run_dir}")
    run_json = json.loads(run_json_path.read_text())

    kind = detect_kind(run_dir)
    entries = load_entries(run_dir, kind)
    model = extract_model(entries, run_json)
    items = build_conversation(entries)
    final_message = extract_final_message(items, run_json)
    manifest = gather_manifest(run_dir, run_json, kind)
    work_dir_abs = work_dir_from_run_json(run_dir, run_json)

    md = render_markdown(run_json, kind, model, manifest, items, final_message, work_dir_abs)
    out_html = render_html(run_json, kind, model, manifest, items, final_message, work_dir_abs)

    md_path = run_dir / "transcript.md"
    html_path = run_dir / "transcript.html"
    md_path.write_text(md)
    html_path.write_text(out_html)
    return md_path, html_path


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="transcript.py",
        description="Render xlsx-bench run folder(s) into transcript.md + transcript.html.",
    )
    parser.add_argument("run_dirs", nargs="+", type=Path, help="one or more <task>/<arm>/<rep> run folders")
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    rc = 0
    for rd in args.run_dirs:
        try:
            md_path, html_path = render_run_dir(rd)
            print(f"{rd}: wrote {md_path.name}, {html_path.name}")
        except (FileNotFoundError, ValueError, json.JSONDecodeError) as e:
            print(f"error: {rd}: {e}", file=sys.stderr)
            rc = 1
    return rc


if __name__ == "__main__":
    sys.exit(main())
