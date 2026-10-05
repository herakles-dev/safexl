"""verify() reports defined names and table ranges whose text differs (reference_changes); verdict unchanged."""
from __future__ import annotations

import zipfile

from safexl_mcp.cli import main
from safexl_mcp.server import stdio_verify
from safexl_mcp.verify import verify
from _pkg import make_package, sheet_xml

BODY = '<row r="1"><c r="A1"><v>1</v></c></row>'


def _pkg(tmp_path, name, names="", sheets=None):
    return make_package(tmp_path / name, sheets or [("Sheet1", sheet_xml(BODY))], defined_names=names)


def _dn(*items):
    inner = "".join(
        f'<definedName name="{n}"' + (f' localSheetId="{s}"' if s is not None else "") + f">{t}</definedName>"
        for n, t, s in items
    )
    return f"<definedNames>{inner}</definedNames>"


def test_moved_name_is_reported_and_verdict_stays_intact(tmp_path):
    a = _pkg(tmp_path, "a.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$13", None)))
    b = _pkg(tmp_path, "b.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$14", None)))
    r = verify(a, b)
    assert r["reference_changes"] == [{"kind": "defined_name", "name": "dropdown", "scope": "workbook",
                                       "before": "Sheet1!$E$2:$E$13", "after": "Sheet1!$E$2:$E$14"}]
    assert "1 reference changed: dropdown" in r["summary"]


def test_unchanged_name_is_not_reported(tmp_path):
    a = _pkg(tmp_path, "a.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$13", None)))
    r = verify(a, a)
    assert r["reference_changes"] == [] and "reference" not in r["summary"]


def test_sheet_scoped_name_reports_sheet_name(tmp_path):
    sheets = [("Sheet1", sheet_xml(BODY)), ("Data", sheet_xml(BODY))]
    a = _pkg(tmp_path, "a.xlsx", _dn(("local", "Data!$A$1:$A$3", 1)), sheets)
    b = _pkg(tmp_path, "b.xlsx", _dn(("local", "Data!$A$1:$A$4", 1)), sheets)
    (c,) = verify(a, b)["reference_changes"]
    assert c["scope"] == "Data" and c["name"] == "local" and c["after"] == "Data!$A$1:$A$4"


def test_removed_and_added_names(tmp_path):
    a = _pkg(tmp_path, "a.xlsx", _dn(("gone", "Sheet1!$A$1", None)))
    b = _pkg(tmp_path, "b.xlsx", _dn(("fresh", "Sheet1!$B$1", None)))
    got = {c["name"]: (c["before"], c["after"]) for c in verify(a, b)["reference_changes"]}
    assert got == {"gone": ("Sheet1!$A$1", None), "fresh": (None, "Sheet1!$B$1")}


def test_malformed_workbook_xml_does_not_raise(tmp_path):
    a = _pkg(tmp_path, "a.xlsx", _dn(("n", "Sheet1!$A$1", None)))
    bad = tmp_path / "bad.xlsx"
    with zipfile.ZipFile(a) as zi, zipfile.ZipFile(bad, "w") as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            zo.writestr(item, b"<workbook><oops" if item.filename == "xl/workbook.xml" else data)
    r = verify(a, bad)
    assert r["reference_changes"] == []
    assert isinstance(r["verdict"], str)


def test_table_ref_change_is_reported(tmp_path):
    def tbl(ref):
        return {"xl/tables/table1.xml": '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                                         f'id="1" name="T1" displayName="T1" ref="{ref}"/>'}
    a = make_package(tmp_path / "a.xlsx", [("Sheet1", sheet_xml(BODY))], extra_parts=tbl("A1:B3"))
    b = make_package(tmp_path / "b.xlsx", [("Sheet1", sheet_xml(BODY))], extra_parts=tbl("A1:B4"))
    (c,) = verify(a, b)["reference_changes"]
    assert (c["kind"], c["name"], c["before"], c["after"]) == ("table", "T1", "A1:B3", "A1:B4")


def test_strict_exit_code_14_for_reference_changes_only(tmp_path, capsys):
    a = _pkg(tmp_path, "a.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$13", None)))
    b = _pkg(tmp_path, "b.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$14", None)))
    assert main(["verify", str(a), str(b)]) == 0
    capsys.readouterr()
    assert main(["verify", str(a), str(b), "--strict"]) == 14
    out = capsys.readouterr().out
    assert "reference: defined_name dropdown (workbook): Sheet1!$E$2:$E$13 -> Sheet1!$E$2:$E$14" in out
    assert main(["verify", str(a), str(a), "--strict"]) == 0


def test_strict_13_for_added_risk_and_13_wins_when_both(tmp_path):
    names13 = _dn(("dropdown", "Sheet1!$E$2:$E$13", None))
    names14 = _dn(("dropdown", "Sheet1!$E$2:$E$14", None))
    calc = '<calcPr calcMode="manual"/>'
    a = _pkg(tmp_path, "a.xlsx", names13)
    risky = _pkg(tmp_path, "risky.xlsx", names13 + calc)  # calculation settings changed: an added risk
    both = _pkg(tmp_path, "both.xlsx", names14 + calc)
    assert main(["verify", str(a), str(risky), "--strict"]) == 13
    r = verify(a, both)
    assert r["reference_changes"] and r["added_risk"]
    assert main(["verify", str(a), str(both), "--strict"]) == 13
    assert main(["verify", str(a), str(both)]) == 0



def test_sheet_rename_is_not_a_reference_change(tmp_path):
    old = [("Detailed_report", sheet_xml(BODY)), ("Other", sheet_xml(BODY))]
    new = [("Detail Report", sheet_xml(BODY)), ("Other", sheet_xml(BODY))]
    a = _pkg(tmp_path, "a.xlsx", _dn(("x", "Detailed_report!$Z$6", None), ("y", "Detailed_report!$A$1:$B$2", 0)), old)
    b = _pkg(tmp_path, "b.xlsx", _dn(("x", "'Detail Report'!$Z$6", None), ("y", "'Detail Report'!$A$1:$B$2", 0)), new)
    assert verify(a, b)["reference_changes"] == []


def test_rename_plus_moved_range_is_listed_once_with_real_text(tmp_path):
    old = [("Detailed_report", sheet_xml(BODY))]
    new = [("Detail Report", sheet_xml(BODY))]
    a = _pkg(tmp_path, "a.xlsx", _dn(("x", "Detailed_report!$Z$6", None), ("z", "Detailed_report!$A$1", None)), old)
    b = _pkg(tmp_path, "b.xlsx", _dn(("x", "'Detail Report'!$Z$7", None), ("z", "'Detail Report'!$A$1", None)), new)
    (c,) = verify(a, b)["reference_changes"]
    assert (c["name"], c["before"], c["after"]) == ("x", "Detailed_report!$Z$6", "'Detail Report'!$Z$7")


def test_quoted_and_unquoted_forms_compare_equal(tmp_path):
    sheets = [("My Sheet", sheet_xml(BODY)), ("Plain", sheet_xml(BODY))]
    a = _pkg(tmp_path, "a.xlsx", _dn(("q", "'My Sheet'!$A$1", None), ("p", "Plain!$A$1", None)), sheets)
    b = _pkg(tmp_path, "b.xlsx", _dn(("q", "'My Sheet'!$A$1", None), ("p", "'Plain'!$A$1", None)), sheets)
    assert verify(a, b)["reference_changes"] == []
    # a quoted name with an escaped apostrophe, renamed
    s1 = [("It's", sheet_xml(BODY))]
    s2 = [("Its", sheet_xml(BODY))]
    c = _pkg(tmp_path, "c.xlsx", _dn(("n", "'It''s'!$A$1", None)), s1)
    d = _pkg(tmp_path, "d.xlsx", _dn(("n", "Its!$A$1", None)), s2)
    assert verify(c, d)["reference_changes"] == []


def test_mcp_verify_returns_reference_changes(tmp_path):
    a = _pkg(tmp_path, "a.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$13", None)))
    b = _pkg(tmp_path, "b.xlsx", _dn(("dropdown", "Sheet1!$E$2:$E$14", None)))
    body = stdio_verify(str(a), str(b)).structured_content
    assert body["verdict"] == "intact"
    assert body["reference_changes"][0]["name"] == "dropdown"
