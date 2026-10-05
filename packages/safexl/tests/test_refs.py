"""safexl_mcp.refs: the formula tokenizer behind insert_rows / rename_sheet."""
from __future__ import annotations

import pytest

from safexl_mcp.refs import (
    rename_sheet_in_formula,
    shift_rows_in_formula,
    shift_sqref,
    translate_formula,
)


def shift(f, at=5, n=1, cur="S", target="S"):
    return shift_rows_in_formula(f, cur, target, at, n)


@pytest.mark.parametrize(
    "formula, expected",
    [
        ("A4+A5", "A4+A6"),
        ("$A$5:$B$9", "$A$6:$B$10"),
        ("SUM(A1:A5)", "SUM(A1:A6)"),  # end at the insertion row moves: range grows
        ("SUM(A1:A4)", "SUM(A1:A4)"),  # ends just above: does not grow
        ("A:A+B:C", "A:A+B:C"),  # whole columns have no rows
        ("$1:$3+4:6", "$1:$3+4:7"),  # row ranges
        ("LOG10(A9)", "LOG10(A10)"),  # function name is not a cell
        ('"A9"&A9', '"A9"&A10'),  # string literal untouched
        ('"say ""A9"""&A9', '"say ""A9"""&A10'),  # escaped quotes inside a literal
        ("Table1[[#This Row],[Q9]]+A9", "Table1[[#This Row],[Q9]]+A10"),
        ("Table1[Col 'A9]+A9", "Table1[Col 'A9]+A10"),
        ("Table1[Col '[A9']]+A9", "Table1[Col '[A9']]+A10"),  # escaped brackets stay inside the name
        ("[1]S!A9+A9", "[1]S!A9+A10"),  # external workbook
        ("ZZZ9+A9", "ZZZ9+A10"),  # ZZZ is past XFD: not a cell
        ("MyName+A9", "MyName+A10"),
        ("#REF!+A9", "#REF!+A10"),
        ("S!A9+'S'!A9+s!A9", "S!A10+'S'!A10+s!A10"),
        ("Other!A9+'Other Sheet'!A9", "Other!A9+'Other Sheet'!A9"),
    ],
)
def test_shift_rows_in_formula(formula, expected):
    assert shift(formula) == expected


def test_shift_unqualified_refs_only_on_home_sheet_and_never_for_names():
    assert shift("A9", cur="Other") == "A9"
    assert shift_rows_in_formula("A9+S!A9", None, "S", 5, 1) == "A9+S!A10"


def test_shift_count_and_clamp_at_last_row():
    assert shift("A9", n=3) == "A12"
    assert shift("A1:A1048576") == "A1:A1048576"


def test_shift_sqref():
    assert shift_sqref("A1:B5 D7 9:12 C:C", 5, 2) == "A1:B7 D9 11:14 C:C"


def test_translate_moves_relative_parts_only():
    assert translate_formula("A1+$A$1+$A1+A$1", 2, 3) == "D3+$A$1+$A3+D$1"
    assert translate_formula("S!A1:B2", 1, 0) == "S!A2:B3"
    assert translate_formula("A:B+1:2", 1, 1) == "B:C+2:3"
    assert translate_formula("A1", -1, 0) == "#REF!"
    assert translate_formula('"A1"&Table1[A1]', 5, 5) == '"A1"&Table1[A1]'


@pytest.mark.parametrize(
    "formula, expected",
    [
        ("Old!A1+'Old'!B2", "New!A1+New!B2"),
        ("OLD!A1", "New!A1"),
        ('"Old!A1"&Old!A1', '"Old!A1"&New!A1'),
        ("[1]Old!A1+Old!A1", "[1]Old!A1+New!A1"),
        ("Old!MyName+Older!A1+XOld!A1", "New!MyName+Older!A1+XOld!A1"),
        ("Old!A:A+Old!$1:$3", "New!A:A+New!$1:$3"),
        ("Table1[Old!x]+Old!A1", "Table1[Old!x]+New!A1"),
    ],
)
def test_rename_sheet_in_formula(formula, expected):
    assert rename_sheet_in_formula(formula, "Old", "New") == expected


def test_rename_quotes_only_when_needed():
    assert rename_sheet_in_formula("Old!A1", "Old", "New Name") == "'New Name'!A1"
    assert rename_sheet_in_formula("Old!A1", "Old", "It's") == "'It''s'!A1"
    assert rename_sheet_in_formula("Old!A1", "Old", "AB12") == "'AB12'!A1"
    assert rename_sheet_in_formula("Old!A1", "Old", "R1C1") == "'R1C1'!A1"
    assert rename_sheet_in_formula("Old!A1", "Old", "Data_2") == "Data_2!A1"
    assert rename_sheet_in_formula("Old!A1", "Old", "2020") == "'2020'!A1"
