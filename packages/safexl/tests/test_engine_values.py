"""safexl_mcp.engine -- cell VALUE writing: int/float/bool/str, including
XML-special characters. See engine._make_cell_xml.
"""
from __future__ import annotations

import re
from xml.etree import ElementTree as ET

from safexl_mcp import edit

from _util import build_minimal_workbook, sheet1_xml


def test_int_value_written_as_plain_numeric_v(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 42}])
    xml = sheet1_xml(dst)
    assert '<c r="B2"><v>42</v></c>' in xml


def test_negative_int_value(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": -17}])
    xml = sheet1_xml(dst)
    assert '<c r="B2"><v>-17</v></c>' in xml


def test_float_value_written_as_plain_numeric_v(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 3.14}])
    xml = sheet1_xml(dst)
    assert '<c r="B2"><v>3.14</v></c>' in xml


def test_zero_value_written(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": 0}])
    xml = sheet1_xml(dst)
    assert '<c r="B2"><v>0</v></c>' in xml


def test_bool_true_written_with_t_b_and_v_1(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": True}])
    xml = sheet1_xml(dst)
    assert '<c r="B2" t="b"><v>1</v></c>' in xml


def test_bool_false_written_with_t_b_and_v_0(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": False}])
    xml = sheet1_xml(dst)
    assert '<c r="B2" t="b"><v>0</v></c>' in xml


def test_bool_checked_before_int_bool_is_int_subclass(tmp_path):
    """Regression guard: engine._make_cell_xml checks isinstance(value, bool)
    BEFORE isinstance(value, (int, float)) -- bool is an int subclass in
    Python, so a naive numeric-first check would silently drop the t="b"
    marker and write True as a bare numeric 1."""
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "C3", "value": True}])
    xml = sheet1_xml(dst)
    cell_tag = re.search(r'<c r="C3"[^>]*>', xml).group(0)
    assert 't="b"' in cell_tag
    assert '<c r="C3"><v>1</v></c>' not in xml


def test_str_value_written_as_inline_str(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": "hello"}])
    xml = sheet1_xml(dst)
    assert '<c r="B2" t="inlineStr"><is><t xml:space="preserve">hello</t></is></c>' in xml


def test_empty_str_value_written_as_inline_str(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": ""}])
    xml = sheet1_xml(dst)
    assert '<c r="B2" t="inlineStr"><is><t xml:space="preserve"></t></is></c>' in xml


def test_str_value_with_xml_special_characters_round_trips_exactly(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    raw = "<tag> & \"quotes\" 'apostrophe' </tag> éü中文"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": raw}])

    # The output must still be well-formed XML (this is the safety property
    # that matters -- not any particular escaped byte sequence).
    xml = sheet1_xml(dst)
    root = ET.fromstring(xml)  # raises ET.ParseError if malformed

    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    found = None
    for c in root.iter(f"{ns}c"):
        if c.get("r") == "B2":
            t_el = c.find(f"{ns}is/{ns}t")
            found = t_el.text
            break
    assert found == raw


def test_str_value_amp_lt_gt_escaped_in_raw_bytes(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(src, dst, [{"sheet": "Sheet1", "cell": "B2", "value": "a<b>c&d"}])
    xml = sheet1_xml(dst)
    assert "a&lt;b&gt;c&amp;d" in xml
    assert "a<b>c&d" not in xml  # the raw, unescaped form must never appear


def test_multiple_value_types_in_one_edit_call(tmp_path):
    src = build_minimal_workbook(tmp_path / "src.xlsx")
    dst = tmp_path / "out.xlsx"
    edit(
        src,
        dst,
        [
            {"sheet": "Sheet1", "cell": "A1", "value": 1},
            {"sheet": "Sheet1", "cell": "A2", "value": 1.5},
            {"sheet": "Sheet1", "cell": "A3", "value": True},
            {"sheet": "Sheet1", "cell": "A4", "value": "text"},
        ],
    )
    xml = sheet1_xml(dst)
    assert '<c r="A1"><v>1</v></c>' in xml
    assert '<c r="A2"><v>1.5</v></c>' in xml
    assert '<c r="A3" t="b"><v>1</v></c>' in xml
    assert '<c r="A4" t="inlineStr"><is><t xml:space="preserve">text</t></is></c>' in xml
