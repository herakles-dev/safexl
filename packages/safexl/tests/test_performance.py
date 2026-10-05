"""Scaling guard: each operation plus verify() must stay roughly linear in the size of the sheet.

It times the same operation on N and 4N rows and asserts the ratio, so it does not depend on machine
speed. Linear is 4; the bound (6) leaves room for noise and fixed costs but fails on a quadratic step
(a ratio near 16). Skip with SAFEXL_SKIP_PERF=1.
"""
from __future__ import annotations

import os
import time

import pytest

from safexl_mcp import edit, insert_rows
from safexl_mcp.verify import verify
from _pkg import make_package, sheet_xml

pytestmark = pytest.mark.skipif(os.environ.get("SAFEXL_SKIP_PERF") == "1", reason="SAFEXL_SKIP_PERF=1")

SMALL, LARGE = 30_000, 120_000
MAX_RATIO = 6.0


def _book(path, n):
    rows = "".join(
        f'<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}" t="inlineStr"><is><t>name {r}</t></is></c>'
        f'<c r="C{r}"><f>A{r}*2</f><v>{r * 2}</v></c></row>'
        for r in range(1, n + 1)
    )
    return make_package(path, [("Sheet1", sheet_xml(rows))])


def _timed(op, tmp_path, n):
    src = _book(tmp_path / f"src{n}.xlsx", n)
    out = tmp_path / f"out{n}.xlsx"
    best = None
    for _ in range(2 if n <= SMALL else 1):  # best of two on the small run: noise there inflates the ratio's denominator
        t0 = time.perf_counter()
        op(src, out, n)
        receipt = verify(src, out)
        dt = time.perf_counter() - t0
        best = dt if best is None else min(best, dt)
    return best, receipt


def _edit(src, out, n):
    edit(src, out, [{"sheet": "Sheet1", "cell": f"A{1 + i * (n // 100)}", "value": 9} for i in range(100)])


def _insert_1(src, out, n):
    insert_rows(src, out, "Sheet1", n // 2, 1)


def _insert_5000(src, out, n):
    insert_rows(src, out, "Sheet1", n // 2, 5000)


@pytest.mark.parametrize(
    "name,op,rows_inserted",
    [("edit 100 cells", _edit, 0), ("insert 1 row", _insert_1, 1), ("insert 5000 rows", _insert_5000, 5000)],
)
def test_operation_plus_verify_scales_linearly(tmp_path, name, op, rows_inserted):
    small, r_small = _timed(op, tmp_path, SMALL)
    large, r_large = _timed(op, tmp_path, LARGE)
    assert r_large["verdict"] == "intact"
    if rows_inserted:
        assert r_large["cell_changes"]["Sheet1"]["rows_inserted"]["count"] == rows_inserted
        assert r_large["cell_changes"]["Sheet1"]["values_changed"] == 0
    else:
        assert r_large["cell_changes"]["Sheet1"]["cells_changed"] == 100
    assert large / small < MAX_RATIO, f"{name}: {SMALL} rows {small:.2f}s, {LARGE} rows {large:.2f}s, ratio {large / small:.1f}"
