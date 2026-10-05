"""Benchmark hygiene, added 2026-10-05: no text the tested model is given
(plugin skill, hook messages, package README, CLI help, MCP tool
descriptions) may contain a distinctive string from a benchmark task.

Why: the skill's `rename-sheet` and `add-sheet` examples were the literal
answers to two hard-8 tasks, and an earlier `insert-rows` example was a
third. A run whose instructions hold a task answer proves nothing.

SAFEXL_PKG overrides the package location (default: packages/safexl beside
this repo's research/ dir). Package tests are not scanned: the model never
sees them."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent.parent

TASK_FILES = ("tasks.json", "tasks_hard8.json", "tasks_pilot.json", "tasks_run.json", "smoke_tasks.json")
# Keys that hold file names or our own notes, not anything a model could copy.
_SKIP_KEYS = {"workbook", "notes", "kind", "preserve"}
# Too generic to mean a leak: default sheet names, one lowercase word, file names.
_GENERIC = re.compile(r"sheet\d+|[a-z]+|[\w.-]+\.xlsx")
_TEXT_SUFFIXES = {".md", ".py", ".json", ".sh", ".txt", ".toml", ".yaml", ".yml"}
_SKIP_DIRS = ("__pycache__", "/build/", "egg-info", ".venv", "/tests/")


def _package() -> Path:
    return Path(os.environ.get("SAFEXL_PKG") or _HERE.parent.parent / "packages" / "safexl")


def _tasks() -> dict[str, dict]:
    tasks: dict[str, dict] = {}
    for name in TASK_FILES:
        path = _HERE / name
        if not path.exists():
            continue
        data = json.loads(path.read_text())
        for task in data if isinstance(data, list) else data.get("tasks", []):
            tasks.setdefault(task["id"], task)
    return tasks


def task_needles(tasks: dict[str, dict]) -> dict[str, set[str]]:
    """Distinctive lowercase strings per task: quoted text in the prompt and
    every string the grader expects (sheet names, labels, cell text)."""
    needles: dict[str, set[str]] = {}

    def add(text: str, task_id: str) -> None:
        text = text.strip().lower()
        if 6 <= len(text) <= 60 and not text.startswith("=") and not _GENERIC.fullmatch(text):
            needles.setdefault(text, set()).add(task_id)

    def walk(node, task_id: str) -> None:
        if isinstance(node, dict):
            for key, value in node.items():
                if key not in _SKIP_KEYS:
                    walk(value, task_id)
        elif isinstance(node, list):
            for value in node:
                walk(value, task_id)
        elif isinstance(node, str):
            add(node, task_id)

    for task_id, task in tasks.items():
        walk({k: v for k, v in task.items() if k not in ("id", "prompt")}, task_id)
        for single, double in re.findall(r"'([^']{5,60})'|\"([^\"]{5,60})\"", task.get("prompt", "")):
            add(single or double, task_id)
    return needles


def instruction_files(package: Path) -> list[Path]:
    roots = [package / "claude-plugin", package / "safexl_mcp", package / "README.md"]
    files: list[Path] = []
    for root in roots:
        if root.is_file():
            files.append(root)
        elif root.is_dir():
            files += [
                p for p in root.rglob("*")
                if p.is_file() and p.suffix in _TEXT_SUFFIXES and not any(s in p.as_posix() for s in _SKIP_DIRS)
            ]
    return sorted(files)


def find_leaks(files: list[Path], needles: dict[str, set[str]]) -> list[str]:
    leaks = []
    for path in files:
        text = path.read_text(errors="replace").lower()
        for needle, task_ids in needles.items():
            if needle in text:
                leaks.append(f"{path}: {needle!r} is from {', '.join(sorted(task_ids))}")
    return sorted(leaks)


# ---------------------------------------------------------------------------
# Structural leaks (added 2026-10-05): ranges, formulas, cell targets, short
# codes and insert-row numbers. The needle scan above skips strings that start
# with "=" and anything under 6 characters, so "C30=SUM(C25:C29)" and the code
# "038" slipped through once.

_ADDR = r"\$?[a-z]{1,3}\$?\d+"
_RANGE_RE = re.compile(r"(?:(?:'[^']+'|[\w.]+)!)?(" + _ADDR + ":" + _ADDR + ")", re.I)
_CALL_RE = re.compile(r"\b[a-z_][a-z0-9_.]*\(([^()]*)\)", re.I)
_ADDR_RE = re.compile(r"(?<![\w$])(\$?[a-z]{1,3}\$?\d+)(?![\w(])", re.I)

# Short literals that are generic in the docs. Each one is listed with the reason.
# Add a token only with a comment saying why the docs may keep it.
_ALLOW_LITERALS: set[str] = {
    # "2006" is the year in the OOXML namespace URI (.../spreadsheetml/2006/main)
    # that the engine code must name; it is also a value in pivotrpt-t4-insert-games-row.
    "2006",
}


def _norm(text: str) -> str:
    return text.replace("$", "").lower()


def _strings(node):
    if isinstance(node, dict):
        for value in node.values():
            yield from _strings(value)
    elif isinstance(node, list):
        for value in node:
            yield from _strings(value)
    elif isinstance(node, str):
        yield node


def _expect_cells(task: dict):
    """(ref, value) for every cell the grader expects a value, text or formula in."""
    exp = task.get("expect", {})
    for key in ("text_cells", "cells", "formulas"):
        for item in exp.get(key, []):
            if "ref" in item:
                yield item["ref"], item.get("value")


def _fmt_number(value) -> str:
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return repr(value)


def task_structure(tasks: dict[str, dict]) -> dict[str, dict]:
    """Per task: ranges, call bodies, addresses, short literals, insert row."""
    out: dict[str, dict] = {}
    for task_id, task in tasks.items():
        texts = list(_strings({k: v for k, v in task.items() if k not in ("id", "workbook")}))
        ranges, calls = set(), set()
        for text in texts:
            ranges |= {_norm(m.group(1)) for m in _RANGE_RE.finditer(text)}
            calls |= {_norm(m.group(0)).replace(" ", "") for m in _CALL_RE.finditer(text)}
        addrs = {_norm(ref) for ref, _ in _expect_cells(task)}
        literals: set[str] = set()
        exp = task.get("expect", {})
        values = [v for _, v in _expect_cells(task)]
        for item in exp.get("column_values", []):
            values += list(item.get("rows", {}).values())
        for item in exp.get("sheet_contains", []):
            values += [v.get("value") for v in item.get("values", [])]
        for value in values:
            if isinstance(value, bool) or value is None:
                continue
            text = value if isinstance(value, str) else _fmt_number(value)
            text = text.strip().lower()
            if len(text) >= 3 and re.search(r"\d", text) and not _GENERIC.fullmatch(text):
                literals.add(text)
        row = None
        if task.get("kind") == "insert_row":
            rows = [int(re.sub(r"\D", "", ref)) for ref, _ in _expect_cells(task)]
            if rows and min(rows) >= 10:
                row = min(rows)
        out[task_id] = {"ranges": ranges, "calls": calls, "addrs": addrs, "literals": literals, "row": row}
    return out


def find_structural_leaks(files: list[Path], tasks: dict[str, dict]) -> list[str]:
    structure = task_structure(tasks)
    plain = task_needles(tasks)
    leaks = []
    for path in files:
        text = _norm(path.read_text(errors="replace"))
        squashed = text.replace(" ", "")
        for task_id, info in structure.items():
            for rng in info["ranges"]:
                if re.search(r"(?<![\w$])" + re.escape(rng) + r"(?![\w])", text):
                    leaks.append(f"{path}: range {rng!r} is from {task_id}")
            for call in info["calls"]:
                if call in squashed:
                    leaks.append(f"{path}: formula {call!r} is from {task_id}")
            for lit in info["literals"]:
                if lit in _ALLOW_LITERALS:
                    continue
                if re.search(r"(?<![\w.])" + re.escape(lit) + r"(?![\w])", text):
                    leaks.append(f"{path}: literal {lit!r} is from {task_id}")
            if info["row"] is not None:
                if re.search(r"(?:--row[=\s]+|\brow\s*)" + str(info["row"]) + r"(?!\d)", text):
                    leaks.append(f"{path}: insert row {info['row']} is from {task_id}")
            own_lits = info["literals"] | {n for n, ids in plain.items() if task_id in ids}
            for line in text.splitlines():
                found = {_norm(m.group(1)) for m in _ADDR_RE.finditer(line)} & info["addrs"]
                if not found:
                    continue
                row_hit = info["row"] is not None and re.search(r"\brow\s*" + str(info["row"]) + r"\b", line)
                lit_hit = any(lit in line for lit in own_lits if lit not in _ALLOW_LITERALS)
                if len(found) >= 2 or row_hit or lit_hit:
                    leaks.append(f"{path}: cell target(s) {sorted(found)} with task data on one line, from {task_id}: {line.strip()[:80]!r}")
    return sorted(set(leaks))


def prompt_files() -> list[Path]:
    """Prompt text that lives in the kit itself, not in the package: the system prompt the
    model-agnostic harness gives every model (it was copied from the skill text, leak included)."""
    return [p for p in (_HERE / "agnostic" / "harness.py",) if p.is_file()]


def test_no_task_leak_in_kit_prompts():
    files = prompt_files()
    if not files:
        pytest.skip("no kit-side prompt files found")
    leaks = find_leaks(files, task_needles(_tasks())) + find_structural_leaks(files, _tasks())
    assert not leaks, "benchmark task text in a prompt the kit gives the model:\n" + "\n".join(leaks)


def test_scan_has_something_to_scan():
    package = _package()
    if not package.is_dir():
        pytest.skip(f"SafeXL package not found at {package}")
    assert len(task_needles(_tasks())) >= 50
    assert len(instruction_files(package)) >= 5


def test_scanner_catches_a_planted_answer(tmp_path):
    (tmp_path / "README.md").write_text('safexl rename-sheet in.xlsx out.xlsx Sheet2 "Yearly Totals"\n')
    leaks = find_leaks(instruction_files(tmp_path), task_needles(_tasks()))
    assert len(leaks) == 1 and "pivotrpt-t5-rename-sheet2" in leaks[0]


def test_no_task_string_in_model_instructions():
    package = _package()
    if not package.is_dir():
        pytest.skip(f"SafeXL package not found at {package}")
    leaks = find_leaks(instruction_files(package), task_needles(_tasks()))
    assert not leaks, "benchmark task text in what the model is shown:\n" + "\n".join(leaks)


def test_no_structural_task_leak_in_model_instructions():
    package = _package()
    if not package.is_dir():
        pytest.skip(f"SafeXL package not found at {package}")
    leaks = find_structural_leaks(instruction_files(package), _tasks())
    assert not leaks, "benchmark ranges, formulas, cells, codes or rows in what the model is shown:\n" + "\n".join(leaks)


def _planted(tmp_path, text):
    (tmp_path / "README.md").write_text(text + "\n")
    return find_structural_leaks(instruction_files(tmp_path), _tasks())


def test_structural_scanner_catches_planted_range_and_formula(tmp_path):
    leaks = _planted(tmp_path, "`=SUM(C25:C28)` with a row inserted needs more")
    assert any("benefits-t4-insert-duplicate-exclusion" in leak and "c25:c28" in leak for leak in leaks)
    assert any("formula 'sum(c25:c28)'" in leak for leak in leaks)


def test_structural_scanner_catches_planted_assignment_with_sheet(tmp_path):
    leaks = _planted(tmp_path, "safexl edit f.xlsx --formula 'Enrollment Restrictions!C30=SUM(C25:C29)'")
    assert any("benefits-t4-insert-duplicate-exclusion" in leak and "cell target" in leak for leak in leaks)


def test_structural_scanner_catches_planted_short_code(tmp_path):
    cells = _tasks()["kyc-t4-insert-ladakh-state"]["expect"]["text_cells"]
    code = next(c["value"] for c in cells if c["ref"] == "A3")
    leaks = _planted(tmp_path, f"safexl edit f.xlsx --text 'Sheet1!A1={code}'")
    assert any(f"literal {code!r}" in leak and "kyc-t4-insert-ladakh-state" in leak for leak in leaks)


def test_structural_scanner_catches_planted_insert_row(tmp_path):
    cells = _tasks()["benefits-t4-insert-duplicate-exclusion"]["expect"]["cells"]
    row = min(int(re.sub(r"\D", "", c["ref"])) for c in cells)
    leaks = _planted(tmp_path, f"safexl insert-rows f.xlsx --sheet Data --row {row}")
    assert any(f"insert row {row}" in leak for leak in leaks)


def test_scanners_catch_planted_new_sheet_formula(tmp_path):
    (tmp_path / "README.md").write_text("safexl edit f.xlsx --formula 'Lookup Check!B2=Sheet2!L1'\n")
    plain = find_leaks(instruction_files(tmp_path), task_needles(_tasks()))
    assert plain and "sparkext-t5-add-lookup-check-sheet" in plain[0]


def test_structural_scanner_is_quiet_on_generic_text(tmp_path):
    assert _planted(tmp_path, "safexl edit book.xlsx --set 'Sheet1!A1=5' and SUM(A1:A3) in row 2") == []
