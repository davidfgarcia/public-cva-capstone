"""The registry's own tests (code/filekinds). Invented data only.

    python -m pytest -q code/filekinds/tests
"""
from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import cva_filekinds as fk  # noqa: E402
from cva_filekinds import wage_table as wt  # noqa: E402
from cva_filekinds.merged_wage_table import MERGED_TABLE_HEADER  # noqa: E402


def _rows(rows):
    return [(i + 1, {j + 1: v for j, v in enumerate(r) if v is not None}) for i, r in enumerate(rows)]


def test_the_package_imports_only_the_standard_library_and_itself():
    """It is copied into both components and must load in either, on Python 3.9 (the AWS server)."""
    allowed = {"__future__", "re", "dataclasses", "typing",
               "hashlib", "json", "time"}       # settings_file.py (#151): standard library too
    for path in (HERE.parent / "cva_filekinds").glob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Import):
                assert {a.name.split(".")[0] for a in node.names} <= allowed, path.name
            elif isinstance(node, ast.ImportFrom) and node.level == 0:
                assert node.module.split(".")[0] in allowed, (path.name, node.module)
            assert not isinstance(node, getattr(ast, "Match", ())), path.name


def test_columns_refuse_unknown_types_and_crossings():
    with pytest.raises(ValueError):
        fk.Column("x", "x", type="person")
    with pytest.raises(ValueError):
        fk.Column("x", "x", crosses="maybe")


def test_the_layouts_give_back_the_readers_header_maps_in_order():
    assert list(fk.WAGE_TABLE.headers("lcp").items()) == list(wt.LCP_HEADERS.items())
    assert list(fk.WAGE_TABLE.headers("table").items()) == list(wt.TABLE_HEADERS.items())
    assert wt.WAGE_HEADERS == sorted({c.header_pattern for c in fk.WAGE_TABLE.columns})


def test_columns_read_agrees_with_the_tokenizers_wage_column_test():
    headers = ["Trade", "Job  Classification", "Base Rate", "Basic Hourly", "Fringe Rate", "Total Wage",
               "Expiration Date", "Vac/PTO in gross wage", "Worker", "Apprentice", "Notes", "Week End Date",
               "Standard Hours", "Std Day", "Mod", "Jurisdiction", "County", 42, None]
    header = {i + 1: h for i, h in enumerate(headers)}
    read = fk.columns_read(fk.WAGE_TABLE, header)
    expected = {c for c, h in header.items() if isinstance(h, str) and wt.WAGE_COL.fullmatch(fk.norm_header(h))}
    assert set(read) == expected
    assert {header[c] for c in set(header) - expected} == {"Worker", "Apprentice", "Notes", "Week End Date", 42, None}


def test_crosses_to_aws_fails_closed():
    assert fk.crosses_to_aws(fk.WAGE_TABLE, "Base Rate") == fk.AS_IS
    assert fk.crosses_to_aws(fk.WAGE_TABLE, "Worker") == fk.TEXT_TOKEN
    assert fk.crosses_to_aws(fk.WAGE_TABLE, 12) == fk.TEXT_TOKEN
    assert fk.crosses_to_aws(fk.WAGE_TABLE, "Fringe Rate", layout="lcp") == fk.TEXT_TOKEN   # a table header only


def test_recognise_workbook():
    lcp = _rows([["Location: King County"], ["Trade", "Classification", "Total Rate", "Crew"],
                 ["LABORER", "Group 1", 54.8, "Dana Wexley"]])
    found = fk.recognise_workbook([lcp], ["Wages"])
    assert (found.kind, found.sheet, found.header_row, found.layout, set(found.columns)) == \
        ("wage_table", 0, 2, "lcp", {1, 2, 3})

    table = _rows([["County", "Trade", "Classification", "Base Rate", "Comments"]])
    found = fk.recognise_workbook([_rows([["Notes"]]), table], ["Notes", "Rates"], active=1)
    assert (found.kind, found.sheet, found.layout, set(found.columns)) == ("wage_table", 1, "table", {1, 2, 3, 4})
    assert fk.recognise_workbook([_rows([["Notes"]]), table], ["Notes", "Rates"], active=0) is None
    assert fk.recognise_workbook([table], ["Rates"], active=5).sheet == 0           # kept inside the workbook

    merged = _rows([MERGED_TABLE_HEADER])
    assert wt.recognise([merged], ["rates"]) is not None                            # it looks like a table too...
    assert fk.recognise_workbook([merged], ["rates"]).kind == "merged_wage_table"   # ...the agent's own wins

    assert fk.recognise_workbook([_rows([["Name", "Phone"], ["Dana Wexley", "555-0142"]])], ["Roster"]) is None
    assert fk.recognise_workbook([], []) is None


def test_the_lcp_window_is_forty_rows():
    pad = [["filler"]] * 39
    assert wt.locate_wage_header(_rows(pad + [["Craft", "Classification"]]))[0] == 40
    assert wt.locate_wage_header(_rows(pad + [["x"], ["Craft", "Classification"]])) is None


def test_the_location_line():
    for v, keep in (("Location: Grays Harbor", True), ("Location: Walla Walla County, WA", True),
                    ("Location: Dana Wexley", False), ("Location:", False)):
        assert bool(wt.LOCATION_RE.fullmatch(v)) == keep, v


def test_the_registry_lists_its_kinds():
    assert [k.kind for k in fk.KINDS] == ["merged_wage_table", "payroll_export", "wage_table", "tracker", "project_info",
                                         "project_data"]
    assert fk.BY_NAME["wage_table"] is fk.WAGE_TABLE
    assert re.fullmatch(r"\d+", fk.REGISTRY_VERSION)
    assert all(isinstance(k.version, int) and k.version >= 1 for k in fk.KINDS)


def test_the_copy_check_sees_every_file(tmp_path):
    """Review of PR #117: a changed byte, an extra file, a subfolder or a non-.py file in a copy is a difference."""
    import importlib.util
    import shutil
    spec = importlib.util.spec_from_file_location("filekinds_sync", HERE.parent / "sync.py")
    sync = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(sync)
    copy = tmp_path / "cva_filekinds"
    shutil.copytree(sync.COPIES[0], copy, ignore=shutil.ignore_patterns("__pycache__"))
    assert sync.differences(copy) == []
    (copy / "core.py").write_bytes((copy / "core.py").read_bytes().replace(b"\n", b"\r\n", 1))
    (copy / "extra.json").write_text("{}")
    (copy / "sub").mkdir()
    (copy / "sub" / "x.py").write_text("")
    assert sync.differences(copy) == ["core.py", "extra.json", "sub/x.py"]


def test_amounts_and_the_strict_wage_columns():
    """#118 (strict wage tables): what counts as an amount, and which header cells, rate and money columns a header
    row has."""
    for v, ok in ((54.8, True), (8, True), ("54.80", True), ("$1,234.50", True), (" .5 ", True), (True, False),
                  ("$50.00/hr", True), ("70.00 /hr", True), ("40 per hour", True),
                  ("12 Dana", False), ("ask Dana", False), ("", False), (None, False)):
        assert wt.is_amount(v) is ok, v
    header = {1: "Trade", 2: "Classification Description", 3: "Base Rate", 4: "Fringes", 5: "Total Wage",
              6: "Jurisdiction", 7: "Worker", 8: 12, 9: "Std Day", 10: "Classification (ask Dana Wexley)",
              11: "Classification Dana Wexley", 12: "Classification - Dana", 13: "Job Classification Title"}
    head, rates, money, roles = wt.wage_columns(header)
    assert head == {1, 2, 3, 4, 5, 6, 9, 13} and rates == {3, 5} and money == {3, 4, 5}
    assert roles == {1: "craft", 3: "basic", 4: "fringe", 5: "total", 6: "jur", 9: "std"}
    assert wt.RATE_ROLES == ("basic", "total")          # parse_lcp_export: a row needs a Total or Basic amount


def test_values_of_each_type():
    ok = {"jur": ["Federal", "State PW", "Union (Local 32)", "State - Pierce", "Federal - Pierce County",
                  "State - King County", "Federal \u2013 WA20250091", "Federal WA20250091 Mod 3",
                  "CLARK - UNION RATES - WA", "Federal (Davis-Bacon)", "Union - Local 46",
                  "Union - IBEW Local 46"],
          "wd": ["WA20250091", "WA20250091 Mod 3", "Union CBA", "#2025-11", "N/A", "State PW", "WA2025-0091",
                 "WA State Prevailing Wage", "State PW 08/31/2025"],
          "county": ["King County", "king", "Walla Walla County, WA", "Statewide", "Multnomah County",
                     "King, Pierce, Snohomish", "King/Pierce", "Clark County WA", "Statewide (all counties)",
                     "Eastern WA"],
          "dates": ["12/12/2025\n01/01/2100", "2025-12-12", "Dec 12, 2025", "12/12/2025\nOpen"],
          "std": ["8", "8.0"], "basic": ["$40.10", "40.10/hr", "$54.80*", "54.80 USD"], "cls": ["Group 1 (any text)"],
          "craft": ["LABORER"]}
    bad = {"jur": ["Dana Wexley said state", "see Dana", "State - Dana Wexley", "state and federal"],
           "wd": ["WA Dana Wexley", "per Dana"],
           "county": ["Dana Wexley", "call Dana at King County office", "ask Dana Wexley re King County"],
           "dates": ["per Dana", "12/12 Dana"], "std": ["Dana", "8 Dana"], "basic": ["ask Dana"]}
    for role, values in ok.items():
        for v in values:
            assert wt.value_ok(role, v), (role, v)
    for role, values in bad.items():
        for v in values:
            assert not wt.value_ok(role, v), (role, v)
    assert wt.value_ok("std", 8) and wt.value_ok("dates", None)


def test_the_tracker_kinds():
    tracker = _rows([["Project: Eddleton"], ["Status", "Contractor", "Hiring Contractor", "Intent ID", "Notes"]])
    found = fk.recognise_workbook([_rows([["Prime Intent", 1]]), tracker], ["Project Info", "Eddleton"])
    assert (found.kind, found.sheet, found.header_row, set(found.columns)) == ("tracker", 1, 2, {1, 2, 3, 4})
    assert fk.recognise_workbook([_rows([["Prime Intent", 1]])], ["Project Info"]).kind == "project_info"
    cover = _rows([["Project Data"], ["Field", "Key", "Value", "Format / allowed values"], ["Status", "status", "Active"]])
    assert fk.recognise_workbook([cover], ["Project Data"]).kind == "project_data"
    assert fk.recognise_workbook([_rows([["x"], ["Jurisdiction Override:", "State"]])], ["Data"]) is None   # Nov 12
    assert fk.crosses_to_aws(fk.TRACKER, "Intent ID") == fk.AS_IS
    assert fk.crosses_to_aws(fk.TRACKER, "Section 3 Employees") == fk.TEXT_TOKEN
    from cva_filekinds import project_data, tracker as tk
    assert tk.is_tracker_header(["status", "contractor", "affidavit filed"]) and not tk.is_tracker_header(["contractor"])
    assert project_data.is_project_data_name("Project Data - 26084") and not project_data.is_project_data_name(
        "Project Data (old)") and not project_data.is_project_data_name("26084 Project Data")   # starts with it
    assert tk.AGENT_AUTHORS == {"Compliance Validation Agent", "Compliance Agent"}
    from cva_filekinds import project_info as pi
    assert pi.INFO_LABEL_RE.pattern == r"prime intent|l\s*&\s*i project id|lni project id"     # as before PR 3


PAYROLL_HDR = ["Project Name", "Contractor Name", "Week End Date", "Employee Name", "Employee ID (LCPtracker)",
               "Craft", "Classification", "ST Hours", "ST Mon", "DT Sun", "Base Hourly", "Notes"]


def test_the_payroll_kind():
    from cva_filekinds import payroll_export as pe
    cols = pe.map_columns(PAYROLL_HDR)
    assert cols["employee"] == 3 and cols["emp_id"] == 4 and cols["st_mon"] == 8 and cols["dt_sun"] == 9
    assert "notes" not in cols and len(pe.COLUMN_PATTERNS) == 25 + 21
    # uploads' test: employee or ID, week, and classification / ST hours / craft
    assert pe.is_payroll_header(PAYROLL_HDR)
    assert pe.is_payroll_header(["Employee ID (LCPtracker) 2024", "Week End Date", "ST Hours"])
    assert not pe.is_payroll_header(["Employee Name", "Classification", "ST Hours"])            # no week
    assert pe.payroll_header([["x"]] * 14 + [PAYROLL_HDR]) and not pe.payroll_header([["x"]] * 15 + [PAYROLL_HDR])
    # the reader's test: an exact "classification" cell in rows 1-7, rows read only up to it
    seen = []

    def rows():
        for r in range(1, 20):
            seen.append(r)
            yield r, (PAYROLL_HDR if r == 3 else [None])
    assert pe.find_header(rows())[0] == 3 and seen == [1, 2, 3]
    with pytest.raises(ValueError, match="Could not find the header row"):
        pe.find_header((r, ["x"] if r < 8 else ["Classification"]) for r in range(1, 10))
    assert pe.is_payroll_sheet(_rows([["x"]] * 6 + [[" CLASSIFICATION "]]))
    assert not pe.is_payroll_sheet(_rows([["x"]] * 7 + [["Classification"]]))
    assert not pe.is_payroll_sheet(_rows([["Job Classification"]]))
    assert pe.PAYROLL_EXTS == {".xlsx", ".xlsm", ".csv"} and pe.DAYS[0] == "mon" and len(pe.DAYS) == 7


def test_the_payroll_kind_in_the_registry():
    payroll = _rows([["Payroll report"], PAYROLL_HDR, ["P", "Acme", "2026-01-04", "Dana Wexley", 1001]])
    found = fk.recognise_workbook([payroll], ["Payroll"])
    assert (found.kind, found.sheet, found.header_row) == ("payroll_export", 0, 2)
    assert found.columns == frozenset(range(1, 12))                    # every column but Notes
    # a roster with an exact Classification header is not a payroll by shape (the tokenizer's test says it is: D1)
    roster = _rows([["Name", "Classification", "Phone"], ["Dana Wexley", "Laborer", "555-0142"]])
    assert fk.recognise_workbook([roster], ["Crew"]) is None
    # only the first tab counts
    assert fk.recognise_workbook([_rows([["Notes"]]), payroll], ["Notes", "Payroll"]) is None
    # a header in row 8 is out of the window
    assert fk.recognise_workbook([_rows([["x"]] * 7 + [PAYROLL_HDR])], ["P"]) is None
    k = fk.PAYROLL_EXPORT
    assert fk.crosses_to_aws(k, "Employee Name") == fk.TOKEN and fk.crosses_to_aws(k, "Base Hourly") == fk.AS_IS
    assert fk.crosses_to_aws(k, "Notes") == fk.TEXT_TOKEN                            # descriptive until phase 3
    types = {c.role: c.type for c in k.columns}
    assert types["st_mon"] == types["dt_sun"] == types["st_hours"] == "number" and types["dt_hourly"] == "money"
    assert types["week"] == "date" and types["employee"] == "name" and types["emp_id"] == "id"
    from cva_filekinds import payroll_export as pe
    assert pe.IMPORTANT == ("jurisdiction", "week", "craft", "classification", "base", "st_hours", "ot_hours",
                            "ot_hourly", "cash_fringe_rate", "tp_fringe_rate", "grand_total")
    assert {c.role for c in k.columns if c.required} == set(pe.IMPORTANT)
    # whole headers only: a header that merely contains a known one maps to nothing
    assert pe.map_columns(["Job Classification", "ST Hours Total", "Craft Code", "Employee Names"]) == {}
    assert pe.is_payroll_header(["Employee Name", "Week End Date", "Craft"])
