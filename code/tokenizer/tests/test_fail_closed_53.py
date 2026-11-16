"""#53, the name gap: a worker's name the tokenizer has never learnt must not reach AWS from a payroll export, a wage
table, a sheet's name, a document property, or a file or folder name. Three vNext settings, each off by default (v1
unchanged): fail_closed_tables, opaque_sheet_names, opaque_file_names. Invented names only."""
from __future__ import annotations

import csv
import dataclasses
import io
import json
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest

from helpers import CFG, KEY, SHIPPED, FakeDrive
from pii_tokenizer import names as NM
from pii_tokenizer.cva_filekinds import names as known
from pii_tokenizer.cva_filekinds import payroll_export as PE
from pii_tokenizer.files import Context, detokenize_file, tokenize_file
from pii_tokenizer.mirror import Mirror
from pii_tokenizer.tokens import TOKEN_RE, Vault

AGENT = Path(__file__).resolve().parents[2] / "compliance_agent"
NEVER = "Wexvorlan Pudgett"                        # never in a name column: never learnt
ALL = {"fail_closed_tables": True, "opaque_sheet_names": True, "opaque_file_names": True, "doubt_numbers": True,
       "doubt_typed_cells": True}
ON = dataclasses.replace(CFG, vnext=ALL)
HEADER = ["Project Name", "Contractor Name", "Week End Date", "Employee Name", "Employee ID (LCPtracker)", "Craft",
          "Classification", "ST Hours", "Base Hourly", "Foreman"]


def _xml(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return "".join(z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith((".xml", ".rels")))


def _book(build) -> bytes:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _payroll(wb):
    ws = wb.active
    ws.title = "Payroll"
    ws.append([f"Certified payroll, prepared by {NEVER}"])
    ws.append(HEADER)
    ws.append(["Bridge", "ACME", "2026-10-31", "Rivera, Ana", "7319001", "Laborer", "Group 1", 8, 40, NEVER])
    ws.append([None] * 9 + [f"Signed: {NEVER}"])
    ws["J3"].comment = openpyxl.comments.Comment(f"ask {NEVER}", "Payroll office")
    wb.create_sheet("Notes").append([NEVER])
    d = wb.create_sheet(PE.DECISIONS_SHEET)
    d.append(["Review decisions"])
    d.append(PE.DECISION_HEADERS)
    d.append([1, "R08", "state", "Laborer / Group 1", "WA", "ACME", 2, "line 3", "Keep", "ok", "x", f"see {NEVER}",
              "p1"])


def _tok(data, cfg=ON, name="p.xlsx", vault=None, rel=None):
    vault = vault if vault is not None else Vault()
    _, out, rep = tokenize_file(name, data, Context.build(cfg, KEY, vault), rel=rel)
    return out, vault


def test_off_by_default_and_on_in_the_shipped_config():
    """Off unless set (v1 behaviour); on in the shipped config since Nov 15, 2026 (David)."""
    for k in ("fail_closed_tables", "opaque_sheet_names", "opaque_file_names"):
        assert getattr(CFG, k) is False and getattr(SHIPPED, k) is True, k


# ---- 1. payroll exports and wage tables ----------------------------------------------------------------------------
def test_a_payroll_keeps_only_the_readers_columns_on_its_lines():
    data = _book(_payroll)
    assert NEVER in _xml(_tok(data, CFG)[0])                                 # v1: the name gap (#53)
    out, vault = _tok(data)
    wb = openpyxl.load_workbook(io.BytesIO(out))
    rows = [[c.value for c in r] for r in wb["Payroll"].iter_rows()]
    assert TOKEN_RE.fullmatch(rows[0][0])                                    # above the header
    assert rows[1][1:9] == HEADER[1:9] and TOKEN_RE.fullmatch(rows[1][9])    # header: the reader's cells only
    assert TOKEN_RE.fullmatch(rows[1][0]) and TOKEN_RE.fullmatch(rows[2][0])  # "project": located, never read
    assert rows[2][5:9] == ["Laborer", "Group 1", 8, 40] and rows[2][1] == "ACME"
    assert TOKEN_RE.fullmatch(rows[2][3]) and TOKEN_RE.fullmatch(rows[2][9])  # name (field rules), unread column
    assert TOKEN_RE.fullmatch(rows[3][9])                                    # a footer below the table
    assert TOKEN_RE.fullmatch(wb["Notes"]["A1"].value)                       # another tab
    dec = [[c.value for c in r] for r in wb[PE.DECISIONS_SHEET].iter_rows()]
    assert dec[1] == [h if h in PE.DECISIONS_READ else dec[1][i] for i, h in enumerate(PE.DECISION_HEADERS)]
    assert dec[2][PE.DECISION_HEADERS.index("Decision (Keep / Dismiss)")] == "Keep"
    assert TOKEN_RE.fullmatch(dec[2][PE.DECISION_HEADERS.index("AI explanation and sources")])
    xml = _xml(out)
    assert "Pudgett" not in xml
    back, _ = detokenize_file("p.xlsx", out, vault)
    assert [[c.value for c in r] for r in openpyxl.load_workbook(io.BytesIO(back))["Payroll"].iter_rows()] == \
        [[c.value for c in r] for r in openpyxl.load_workbook(io.BytesIO(data))["Payroll"].iter_rows()]


def test_the_rule_keeps_exactly_the_columns_the_readers_patterns_find():
    """The drift test: the tokenizer's payroll rule keeps the columns the registry's map_columns finds (the agent's
    reader uses the same), the first per role, and nothing else on the header row."""
    from pii_tokenizer import doubt
    sample = lambda pat: pat.split("|")[0].replace("\\(", "(").replace(".*", "").replace("\\", "")
    header = [sample(p) for p in PE.COLUMN_PATTERNS.values()] + ["Classification", "Notes", NEVER]
    assert all(PE.map_columns(header).get(r) == i for i, r in enumerate(PE.COLUMN_PATTERNS)), \
        "every role's pattern must match the header written for it"
    rule = doubt._payroll_rule([(1, {i + 1: v for i, v in enumerate(header)})])
    read = [r for r in PE.COLUMN_PATTERNS if r not in PE.UNREAD_ROLES]
    assert rule.keep_cols == {list(PE.COLUMN_PATTERNS).index(r) + 1 for r in read}   # not the 2nd Classification,
    #                                                                                     not Notes, not 'project'
    assert set(rule.roles.values()) == set(read)
    assert rule.line_cols == {list(PE.COLUMN_PATTERNS).index(r) + 1 for r in PE.LINE_ROLES}


def test_the_agents_payroll_code_reads_no_role_the_tokenizer_drops():
    """The drift test the other way: every column role agent/payroll.py and agent/upload_checks.py read (g("..."),
    cols["..."], "..." in cols) is one the tokenizer keeps readable."""
    import re
    if not (AGENT / "agent" / "payroll.py").exists():
        pytest.skip("the agent's code is not next to the tokenizer")
    kept = set(PE.COLUMN_PATTERNS) - set(PE.UNREAD_ROLES)
    for f in ("payroll.py", "upload_checks.py"):
        src = (AGENT / "agent" / f).read_text()
        used = set(re.findall(r'\bg\("(\w+)"\)|cols\["(\w+)"\]|"(\w+)" in cols', src))
        used = {x for t in used for x in t if x}
        assert used and used <= kept, f"{f} reads {sorted(used - kept)}, which the tokenizer doesn't keep (#53)"


def test_number_columns_keep_only_number_like_text():
    out, _ = _tok(_book(lambda wb: [wb.active.append(r) for r in
                                    [HEADER, ["P", "ACME", "2026-10-31", "Rivera, Ana", "1", "Laborer", "Group 1",
                                              NEVER, "N/A", "x"], ["P", "ACME", "2026-10-31", "Rivera, Ana", "1",
                                                                   "Laborer", "Group 1", "8 hrs", "$40.00", "x"]]]))
    rows = [[c.value for c in r] for r in openpyxl.load_workbook(io.BytesIO(out)).active.iter_rows()]
    assert TOKEN_RE.fullmatch(rows[1][7]) and rows[1][8] == "N/A" and rows[2][7:9] == ["8 hrs", "$40.00"]


COMMON_NAMES = """James Mary John Patricia Robert Jennifer Michael Linda William Elizabeth David Barbara Richard Susan
    Joseph Jessica Thomas Sarah Charles Karen Maria Ana Jose Juan Luis Carlos Rosa Carmen Miguel Elena Jan April May June
    July August Dawn Summer Autumn Grace Hope Joy Faith Rose Lily Mark Marks Page Day Link New Davis Bacon Smith Johnson
    Williams Brown Jones Garcia Miller Wilson Moore Taylor Anderson Jackson White Harris Martin Thompson Martinez
    Robinson Clark Rodriguez Lewis Lee Walker Hall Allen Young King Wright Lopez Hill Scott Green Adams Baker Nelson
    Carter Mitchell Perez Roberts Turner Phillips Campbell Parker Evans Edwards Collins Stewart Sanchez Morris Rogers
    Reed Cook Morgan Bell Murphy Bailey Rivera Cooper Richardson Cox Howard Ward Torres Peterson Gray Ramirez Watson Brooks
    Kelly Sanders Price Bennett Wood Barnes Ross Henderson Coleman Jenkins Perry Powell Long Patterson Hughes Flores
    Washington Butler Simmons Foster Gonzales Bryant Alexander Russell Griffin Diaz Hayes Mason Carpenter Painter Glazier
    Craft Sun Allan Allison Rowan Oran Anand Andor Newall Mayall Dayan March Weeks""".split()


def test_no_common_given_name_or_surname_is_a_known_word():
    """Review of the #53 PR: a name made of known words would cross as written."""
    passing = [n for n in COMMON_NAMES if known.is_plain(n)]
    assert not passing, f"names that pass as known words: {passing}"


def test_the_agent_reads_the_same_column_definition():
    if not (AGENT / "agent" / "payroll.py").exists():
        pytest.skip("the agent's code is not next to the tokenizer")
    sys.path.insert(0, str(AGENT))
    try:
        from agent import knowledge, payroll, upload_checks
    except ImportError as e:                                                  # pragma: no cover
        pytest.skip(f"the agent's requirements are not installed ({e})")
    finally:
        sys.path.remove(str(AGENT))
    assert payroll.COLUMN_PATTERNS == PE.COLUMN_PATTERNS and payroll.LINE_ROLES == PE.LINE_ROLES
    assert upload_checks.LINE_ROLES == PE.LINE_ROLES
    assert knowledge.DECISION_HEADERS == PE.DECISION_HEADERS and knowledge.DECISIONS_SHEET == PE.DECISIONS_SHEET
    src = (AGENT / "agent" / "knowledge.py").read_text()
    read = set(__import__("re").findall(r'get\("([^"]+)"\)', src[src.index("def read_decisions"):
                                                                  src.index("def decision_chunks")]))
    assert read == set(PE.DECISIONS_READ), "read_decisions reads columns the tokenizer doesn't keep (or the reverse)"


def test_a_merged_wage_table_and_a_payroll_like_sheet_are_in_doubt_mode():
    from pii_tokenizer.doubt import MERGED_TABLE_HEADER

    def merged(wb):
        wb.active.title = "rates"
        wb.active.append(MERGED_TABLE_HEADER)
        wb.active.append(["state", "LABORER", "Group 1"] + [None] * (len(MERGED_TABLE_HEADER) - 4) + ["State"])
        wb.active["C2"].comment = openpyxl.comments.Comment(NEVER, "Wage team")
        wb.active["Z1"] = NEVER
        wb.create_sheet("sources").append(["merged_at_pt", "2026-11-01"])

    def crew(wb):
        wb.active.append(["Craft", "Job Classification", "Total Rate", "Worker", "ST Hours"])
        wb.active.append(["LABORER", "Group 1", 54.8, NEVER, 8])
    for build, kept in ((merged, "LABORER"), (crew, None)):    # (the agent reads no rates from the second: as v1)
        data = _book(build)
        assert NEVER in _xml(_tok(data, CFG)[0])
        out, _ = _tok(data)
        assert "Pudgett" not in _xml(out) and (kept is None or kept in _xml(out))


def test_csv_payroll_and_lni_wages():
    pay = (f"Prepared by {NEVER}\r\nWeek End Date,Employee Name,Classification,ST Hours,Foreman\r\n"
           f"10/31/2026,\"Rivera, Ana\",Laborer,8,{NEVER}\r\n,,,,note {NEVER}\r\n").encode()
    lni = f"County,Trade,Job Classification,Wage,Holiday,Overtime,Notes,Contact\nKing,Laborer,G1,54.8,5,1,x,{NEVER}\n"
    for data in (pay, lni.encode()):
        assert NEVER in _tok(data, CFG, "w.csv")[0].decode()
        out, vault = _tok(data, name="w.csv")
        text = out.decode()
        assert "Pudgett" not in text and "Laborer" in text
        assert detokenize_file("w.csv", out, vault)[0] == data
    rows = list(csv.reader(io.StringIO(_tok(pay, name="w.csv")[0].decode())))
    assert rows[1][:4] == ["Week End Date", "Employee Name", "Classification", "ST Hours"]
    assert rows[2][0] == "10/31/2026" and rows[2][2:4] == ["Laborer", "8"]


def test_the_agents_own_lni_data_is_left_as_before():
    data = b"intent_id,contractor\n123,ACME Builders\n"
    rel = "Projects/P/Files/L&I data/2026-11-01-00-00-00/intents.csv"
    assert _tok(data, CFG, "intents.csv", rel=rel)[0] == _tok(data, name="intents.csv", rel=rel)[0]


# ---- 2. sheet names and document properties ------------------------------------------------------------------------
def test_sheet_names_and_properties():
    def build(wb):
        wb.active.title = "Project Info"
        wb.active.append(["Prime Intent", 5])
        wb.create_sheet(f"{NEVER} hours").append([8])
        wb.create_sheet("Audit")
        wb.active["C1"] = f"='{NEVER} hours'!A1*2"
        wb.properties.creator = NEVER
        wb.properties.lastModifiedBy = NEVER
        wb.properties.title = f"Hours of {NEVER}"
    data = _book(build)
    assert NEVER in _xml(_tok(data, CFG)[0])
    # (doubt_numbers off: in doubt mode a formula with a number in it is a token as a whole, sheet name included)
    out, vault = _tok(data, dataclasses.replace(CFG, vnext={**ALL, "doubt_numbers": False}))
    assert "Pudgett" not in _xml(out) and "Pudgett" not in _xml(_tok(data)[0])
    wb = openpyxl.load_workbook(io.BytesIO(out))
    assert wb.sheetnames[0] == "Project Info" and wb.sheetnames[2] == "Audit"
    assert TOKEN_RE.fullmatch(wb.sheetnames[1].split()[0]) and wb.sheetnames[1].endswith(" hours")
    assert wb["Project Info"]["C1"].value == f"='{wb.sheetnames[1]}'!A1*2"
    back, _ = detokenize_file("p.xlsx", out, vault)
    with zipfile.ZipFile(io.BytesIO(back)) as zb, zipfile.ZipFile(io.BytesIO(data)) as zo:
        for n in zo.namelist():
            assert zb.read(n) == zo.read(n), n                              # the way back: byte for byte


def test_a_glued_info_tab_keeps_info_and_comes_back_exactly():
    """Review of the #53 PR: "ProjInfo" is the Project Info tab (agent/intent.py: "info" in its name)."""
    def build(wb):
        wb.active.title = "OrtegaInfo"
        wb.active.append(["Prime Intent", 5])
        wb.create_sheet("Zorquillapayroll")
    data = _book(build)
    out, vault = _tok(data)
    wb = openpyxl.load_workbook(io.BytesIO(out))
    assert wb.sheetnames[0].endswith(" Info") and "Ortega" not in _xml(out) and "Zorquilla" not in _xml(out)
    assert openpyxl.load_workbook(io.BytesIO(detokenize_file("p.xlsx", out, vault)[0])).sheetnames == \
        ["OrtegaInfo", "Zorquillapayroll"]


def test_number_tokens_alone_in_a_sheet_come_back_as_numbers():
    """A bug found in the review (on main since #173, doubt_numbers): a sheet whose only tokens were number tokens was
    not rewritten on the way back."""
    def build(wb):
        wb.active.title = "Tracker"
        wb.active.append(["Contractor", "Intent ID"])
        wb.active.append(["ACME", 1])
        wb.create_sheet("Other")["A1"] = 7
    data = _book(build)
    out, vault = _tok(data, dataclasses.replace(CFG, vnext={"doubt_numbers": True}))
    assert openpyxl.load_workbook(io.BytesIO(detokenize_file("p.xlsx", out, vault)[0]))["Other"]["A1"].value == 7


@pytest.mark.parametrize("tab", ["Zorquilla Quist", "Kim hours", "ZorquillaHours", "ProjInfo", "Marks"])
def test_the_agents_sheets_get_the_tokens_of_names_back(tab):
    """Review and verification of the #53 PR: the agent writes a tab's name into its Audit ("sheet=...;") and Review
    Log (column B); restored on the staff side, it must cross as the tab's AWS-side name again (and match the tab),
    short and glued names too, and nothing else in those sheets may change ("Marks (agent use)" stays)."""
    def build(wb):
        wb.active.title = tab
        wb.active.append(["Contractor", "Intent ID"])
        wb.active.append(["ACME", 1])
        a = wb.create_sheet("Audit")
        a.append(["When", "Task", "Topic", "Note"])
        a.append(["x", "intent", "Marks (agent use)", f"sheet={tab}; B2|s1|s2|n"])
        lg = wb.create_sheet("Review Log")
        lg.append(["When", "Sheet", "Cell"])
        lg.append(["x", tab, "B2"])
    for cfg in (ON, dataclasses.replace(CFG, vnext={"opaque_sheet_names": True})):
        out, vault = _tok(_book(build), cfg)
        wb = openpyxl.load_workbook(io.BytesIO(out))
        aws = wb.sheetnames[0]
        assert wb["Audit"]["D2"].value == f"sheet={aws}; B2|s1|s2|n" and wb["Audit"]["C2"].value == "Marks (agent use)"
        assert wb["Review Log"]["B2"].value == aws
        if tab != "Marks":
            assert tab not in _xml(out)
        back = openpyxl.load_workbook(io.BytesIO(detokenize_file("p.xlsx", out, vault)[0]))
        assert back["Audit"]["D2"].value == f"sheet={tab}; B2|s1|s2|n" and back["Review Log"]["B2"].value == tab


def test_a_merged_tables_source_files_with_opaque_file_names_alone():
    from pii_tokenizer.doubt import MERGED_TABLE_HEADER

    def build(wb):
        wb.active.title = "rates"
        wb.active.append(MERGED_TABLE_HEADER)
        row = ["state", "LABORER", "Group 1"] + [None] * (len(MERGED_TABLE_HEADER) - 3)
        row[MERGED_TABLE_HEADER.index("source_file")] = f"Rates from {NEVER}.xlsx"
        wb.active.append(row)
        wb.create_sheet("sources").append(["merged_at_utc", "file"])
        wb["sources"].append(["2026-11-01", f"Rates from {NEVER}.xlsx"])
    out, _ = _tok(_book(build), dataclasses.replace(CFG, vnext={"opaque_file_names": True}))
    assert "Pudgett" not in _xml(out)


def test_an_apostrophe_in_a_tab_name_comes_back_exactly():
    from openpyxl.workbook.defined_name import DefinedName

    def build(wb):
        wb.active.title = "Tracker"
        wb.active.append(["Contractor", "Intent ID"])
        wb.create_sheet("O'Zorquilla hours")["A1"] = 1
        wb.active["C1"] = "='O''Zorquilla hours'!A1+1"
        wb.active["D1"].hyperlink = "#'O''Zorquilla hours'!A1"
        wb.defined_names["crew"] = DefinedName("crew", attr_text="'O''Zorquilla hours'!$A$1")
    data = _book(build)
    out, vault = _tok(data, dataclasses.replace(CFG, vnext={**ALL, "doubt_numbers": False}))
    assert "Zorquilla" not in _xml(out)
    with zipfile.ZipFile(io.BytesIO(detokenize_file("p.xlsx", out, vault)[0])) as zb, \
            zipfile.ZipFile(io.BytesIO(data)) as zo:
        for n in zo.namelist():
            if n.endswith(".rels") or n == "xl/workbook.xml" or n.startswith("docProps"):
                assert zb.read(n) == zo.read(n), n
        assert "'O''Zorquilla hours'!A1+1" in zb.read("xl/worksheets/sheet1.xml").decode()


def test_tables_charts_and_defined_names():
    from openpyxl.chart import BarChart, Reference
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.table import Table

    def build(wb):
        ws = wb.active
        ws.title = "Tracker"
        for r in (["Contractor", "Intent ID", f"{NEVER} notes"], ["ACME", 1, "x"], ["B", 2, "y"]):
            ws.append(r)
        ws.add_table(Table(displayName="Pudgett_tbl", ref="A1:C3"))
        ws["E1"] = "=SUM(Pudgett_tbl[Intent ID])+Pudgett_range"
        wb.defined_names["Pudgett_range"] = DefinedName("Pudgett_range", attr_text="Tracker!$B$2")
        wb.create_sheet(f"{NEVER} hours")["A1"] = 1
        ws["D2"].hyperlink = f"#'{NEVER} hours'!A1"
        ch = BarChart()
        ch.title = f"Hours of {NEVER}"
        ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=3), titles_from_data=True)
        ws.add_chart(ch, "G2")
    data = _book(build)
    out, vault = _tok(data, dataclasses.replace(CFG, vnext={**ALL, "doubt_numbers": False}))
    assert "Pudgett" not in _xml(out)
    wb = openpyxl.load_workbook(io.BytesIO(out))
    t = list(wb["Tracker"].tables)[0]
    d = list(wb.defined_names)[0]
    assert wb["Tracker"]["E1"].value == f"=SUM({t}[Intent ID])+{d}"
    back = detokenize_file("p.xlsx", out, vault)[0]
    assert "Pudgett_tbl[Intent ID]" in _xml(back) and f"#'{NEVER} hours'!A1" in _xml(back)


def test_a_long_or_clashing_sheet_name():
    long = "Wexvorlan Pudgett and Qorvantine Blamwick crew"              # too long with two tokens: one token
    assert len(NM.sheet_name(long, Context.build(ON, KEY, Vault()))) <= NM.MAX_SHEET_NAME

    def build(wb):
        wb.active.title = "Notes"
        wb.create_sheet(long)
    out, _ = _tok(_book(build))
    assert "Pudgett" not in _xml(out)


# ---- 3. file and folder names ----------------------------------------------------------------------------------------
def test_known_words_pass_and_unknown_runs_become_one_token():
    seen = []
    mk = lambda run, glue: seen.append(run) or f"FIL-{len(seen):010x}"
    assert known.opaque("2026-11-14 Week 36 - REVIEWED.xlsx", mk) == "2026-11-14 Week 36 - REVIEWED.xlsx"
    assert known.opaque("Federal 2026-11-08-13-34-03 PST.xlsx", mk) == "Federal 2026-11-08-13-34-03 PST.xlsx"
    assert known.opaque("WA20250093.txt", mk) == "WA20250093.txt" and known.opaque("L&I data", mk) == "L&I data"
    out = known.opaque(f"{NEVER} union rates 6.26.25.xlsx", mk)
    assert seen == [NEVER] and out == "FIL-0000000001 union rates 6.26.25.xlsx"
    assert known.opaque("Badge 7319001.xlsx", mk).startswith("FIL-")       # a long number can be an ID
    assert known.is_plain("ProjectInfo") and not known.is_plain("Ortega")


class World:
    def __init__(self, cfg=ON):
        self.d = FakeDrive()
        self.raw, self.tok = self.d.root("compliance"), self.d.root("compliance_tokenized")
        self.state, self.cfg = {}, cfg
        self.ctx = Context.build(cfg, KEY, Vault())

    def run(self):
        return Mirror(self.d, self.ctx, self.state, lambda: None, lambda: None).run("compliance", self.raw, self.tok)


def _all_paths(d, root):
    files, folders = d.tree(root)
    return [f["rel"] for f in files.values()] + list(folders)


def test_mirror_names_copies_opaquely_and_restores_the_words():
    w = World()
    w.d.put(w.raw, f"Projects/Bridge {NEVER} Park/Files/{NEVER} notes/readme.txt", b"invented")
    w.d.put(w.raw, f"Projects/Bridge {NEVER} Park/Uploads/Wage rates from {NEVER}.txt", b"General Decision Number")
    w.run()
    paths = _all_paths(w.d, w.tok)
    assert not any("Pudgett" in p for p in paths if not p.startswith("Projects/Bridge"))
    assert all("Pudgett" not in p.split("/", 2)[2] for p in paths if p.count("/") >= 2)   # (the project keeps its name)
    up = next(p for p in paths if "/Uploads/" in p)
    assert up.endswith(" rates from FIL-" + up.rsplit("FIL-", 1)[1]) or "rates" in up
    links = json.loads(w.d.read(w.tok, "_staff_links.json"))["links"]
    assert links and all(not v.startswith("Projects") for v in links.values())
    # AWS writes a result named from the copy: it comes back with the real words
    tdir = next(p for p in paths if p.endswith(" notes"))
    w.d.put(w.tok, f"{tdir}/summary.txt", b"done")
    stem = up.rsplit("/", 1)[1].rsplit(".", 1)[0]
    w.d.put(w.tok, f"Projects/Bridge {NEVER} Park/Payroll/2026-11-14 {stem} - REVIEWED.txt", b"done")
    w.run()
    raw = _all_paths(w.d, w.raw)
    assert f"Projects/Bridge {NEVER} Park/Files/{NEVER} notes/summary.txt" in raw
    assert f"Projects/Bridge {NEVER} Park/Payroll/2026-11-14 Wage rates from {NEVER} - REVIEWED.txt" in raw


def test_turning_it_on_renames_copies_already_on_aws():
    w = World(dataclasses.replace(CFG, vnext={}))
    w.d.put(w.raw, f"Projects/P/Files/{NEVER} notes/readme.txt", b"invented")
    w.run()
    assert any("Pudgett" in p for p in _all_paths(w.d, w.tok))               # v1: the folder's name crosses
    w.ctx = Context.build(ON, KEY, w.ctx.vault)
    w.run()
    assert not any("Pudgett" in p for p in _all_paths(w.d, w.tok))
    assert f"Projects/P/Files/{NEVER} notes/readme.txt" in _all_paths(w.d, w.raw)
