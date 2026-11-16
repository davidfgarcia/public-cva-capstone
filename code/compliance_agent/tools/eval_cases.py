"""The evaluation set for tools/run_evals.py: invented projects, the files staff would upload, and the answer key.

**Everything here is invented** (workers, contractors, rates, projects). Never add a real name, a real payroll or a
copy of a staff file (HANDOFF section 6). Every worker is "Qzvorak, Evaltest nnn" with an employee ID 7319nnn: a
surname and an ID range nothing else uses, so run_evals.py can search the whole run for them (the privacy check).

The payroll lines come from the unit-level acceptance test (tests/test_seeded_violations.py: one deliberate error
per case, or a correct control line), so the two stay one set: that test runs the rules on rows in memory; this set
runs the same lines through the whole pipeline (Uploads, the tokenizer, wage tables merged from an uploaded wage
rates file, the payroll review, the reviewed workbook, the emails).

Each project below is one scenario. ``build(dest)`` writes the input snapshot (``dest/config.yaml`` and
``dest/Projects/...``) and returns the answer key, which run_evals.py saves as ``dest/answer_key.json``:

    {"version": 2,
     "files": {"<project>/<upload file name>": {
         "outcome": "reviewed" | "skipped",
         "reason_contains": "",                    # a skip: words its reason must contain (in the run and in the
                                                   # administrator's email)
         "reason_code": "",                        # a skip: its code in the structured run log (agent/events.py)
         "header_row": 1,                          # the sheet row holding the header (CSV: its line number)
         "lines": {"<sheet row>": {"case": "...", "area": "...", "group": "...", "must": ["R08"], "may": ["R07"],
                                   "notes_contain": [], "notes_absent": []}}}},
     "known_gaps": [{"id": "...", "issue": 67, "files": ["<project>/<file>"], "currently": "not recognized",
                     "why": "..."}]}

A line's ``must``: the rules it must raise (empty: a clean control, no flag at all). ``may``: related rules it may also
raise without counting as a false flag (a short base also leaves base + fringes short). Lines are scored one by one,
except lines sharing a ``group`` (a duplicate is two lines; the flag may sit on either): those are scored together.
``notes_contain`` / ``notes_absent``: words the agent's notes on that line must / must not hold (e.g. which wage type
a check used). A known gap's file is expected to end as ``currently`` says (an open issue): its misses don't fail the
run, but any other outcome, and any false flag, does; when the file gets its real expected result, the report says so
and the gap can be removed.

To add a case: add it to CASES in tests/test_seeded_violations.py when it is about one payroll line (it then runs in
both places), or add a scenario here when it is about files, layouts or whole runs.
"""
from __future__ import annotations

import csv
import datetime as dt
import shutil
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
for p in (AGENT, AGENT / "tests"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

import eval_lni as LNI                                                  # noqa: E402 - the invented tracker and L&I
import test_seeded_violations as SV                                    # noqa: E402 - the shared, invented cases
from agent.layout import Paths, ensure_layout                          # noqa: E402

KEY_VERSION = 4                         # 3: "trackers" (the intent/affidavit check); 4: "reason_code" (the run log)
SURNAME = "Qzvorak"                     # every invented worker's surname: used nowhere else in the set
GIVEN = "Evaltest"                      # ... and given name, followed by a 3-digit number
ID_BASE = 7319000                       # invented employee IDs: 7319001, 7319002, ...
CONTRACTOR = "Example Builders Inc."    # the title rows' contractor (not a worker; may stay readable)
# #53: a worker's name the tokenizer never learns (it is in no payroll's name column): only the fail-closed rules keep
# it off the AWS side (tokenizer vnext.fail_closed_tables, opaque_sheet_names, opaque_file_names; run_evals.py turns
# them on). Searched for everywhere in the run, as the workers are.
NEVER_LEARNT = "Wexvorlan Pudgett"
D = dt.datetime

# The wage rates file every scenario uploads (the firm's template, David Nov 8: "Trade", "Classification", "Base
# Rate"): the same rates as the acceptance test's WageBook (SV.BOOK), so a line means the same in both places.
WAGE_HEADER = ["County", "Trade", "Classification", "Total Wage", "Base Rate", "Fringe Rate", "Jurisdiction",
               "Decision", "Effective Date", "Issue Date", "Expiration Date"]
WAGE_ROWS = [
    ["Pierce County", "Laborers", "General Laborer", 55.0, 40.0, 15.0, "State", "WA L&I 2026", D(2026, 1, 1),
     D(2026, 1, 1), None],
    ["Pierce County", "Electricians", "Journey Level", 85.0, 60.0, 25.0, "State", "WA L&I 2026", D(2026, 1, 1),
     D(2026, 1, 1), None],
    ["Pierce County", "Cement Masons", "Journey Level", 63.0, 45.0, 18.0, "State", "WA L&I 2026", D(2026, 1, 1),
     D(2026, 1, 1), D(2026, 8, 31)],
    ["Pierce County", "Cement Masons", "Journey Level", 66.0, 48.0, 18.0, "State", "WA L&I 2026", D(2026, 9, 1),
     D(2026, 9, 1), None],
    ["Pierce County", "Laborers", "General Laborer", 62.0, 42.0, 20.0, "Union", "Laborers CBA 2026", D(2026, 1, 1),
     D(2026, 1, 1), None],
]

# A small, representative subset of the cases for the layout scenarios (one per area, plus controls)
SUBSET = ["Base rate below the prevailing base", "Fringe benefits short", "Overtime rate below 1.5x base",
          "11 straight-time hours in a day", "45 straight-time hours in the week", "Union: Sunday not paid double time",
          "Classification not in the wage data", "Grand total doesn't add up", "Correct state line (control)",
          "Correct union line (control)", "Overtime correctly paid (control)"]

# Cases of several lines: how each line is scored. "together": the lines are one group (the flag may sit on any of
# them). A list: one entry per line, "case" (the case's must/may) or "control" (a correct line: no flag at all).
# A multi-line case not listed here is scored together.
LINE_ROLES = {
    "Duplicate line": "together",
    # the second line is another worker's correct line, there so the base rate column isn't empty on every line
    # (a fully empty column is one column note: test_an_export_without_any_base_rate_is_noted_once)
    "Base rate missing": ["case", "control"],
    "Base rate written N/A": ["case", "control"],
}

NAME, EID = SV.HEADER.index("Employee Name"), SV.HEADER.index("Employee ID (LCPtracker)")


class Builder:
    """Collects the answer key while the scenarios write their files."""

    def __init__(self, dest: Path):
        self.dest = dest
        self.key = {"version": KEY_VERSION, "files": {}, "trackers": {}, "known_gaps": []}
        self._worker = 0

    def project(self, name: str, wages: bool = True) -> Paths:
        p = Paths(self.dest / "Projects" / name)
        ensure_layout(p.project)
        if wages:
            wage_xlsx(p.upload / "Wage rates 2026.xlsx")
        return p

    def lines(self, cases, notes=None):
        """The cases' lines with invented workers (one per case line, except where a case's lines are the same worker
        on purpose, as in the duplicate line, so cases never add up across each other) and each line's expectation.
        Returns [(case name, [(row values, expectation)])]. ``notes``: {case name: (notes_contain, notes_absent)}."""
        out = []
        for name, area, rows, must, may in cases:
            ids = {}
            roles = LINE_ROLES.get(name, "together")
            if roles != "together" and len(roles) != len(rows):
                raise ValueError(f"LINE_ROLES[{name!r}] needs one role per line")
            contain, absent = (notes or {}).get(name, ((), ()))
            entries = []
            for i, r in enumerate(rows):
                r = list(r)
                if r[NAME] not in ids:
                    self._worker += 1
                    ids[r[NAME]] = (f"{SURNAME}, {GIVEN} {self._worker:03d}", ID_BASE + self._worker)
                r[NAME], r[EID] = ids[r[NAME]]
                case = roles == "together" or roles[i] == "case"
                exp = {"case": name if case else f"{name} (line {i + 1}: a correct line)",
                       "area": area if case else "control",
                       "group": name if roles == "together" else f"{name} #{i + 1}",
                       "must": sorted(must) if case else [], "may": sorted(may) if case else [],
                       "notes_contain": list(contain) if case else [], "notes_absent": list(absent) if case else []}
                entries.append((r, exp))
            out.append((name, entries))
        return out

    def expect(self, project: Paths, file_name: str, outcome: str, header_row: int = 1, lines=None, reason="",
               reason_code=""):
        rel = f"{project.project.name}/{file_name}"
        stems = {Path(f).stem.lower() for f in self.key["files"] if f.startswith(project.project.name + "/")}
        if Path(file_name).stem.lower() in stems:          # the reviewed workbook is named from the stem alone
            raise ValueError(f"two uploads named {Path(file_name).stem!r} in one project can't be told apart")
        self.key["files"][rel] = {"outcome": outcome, "header_row": header_row, "reason_contains": reason,
                                  "reason_code": reason_code,
                                  "lines": {str(row): e for row, e in (lines or {}).items()}}

    def tracker(self, project: Paths, outcome: str, lines=None, header_row: int = 1, lni_only=(), not_due=None,
                never_ask=(), reason=""):
        """The intent/affidavit check's expected result for a project's tracker: ``outcome`` "checked" or "not
        applicable" (``reason``: words the run's line must hold); per line (sheet row) the marks it must make; the L&I
        contractors the Review Log must list as not in the tracker; the not-due-yet counts; IDs L&I must never be asked
        about (a tracker that isn't checked). The run's summary counts follow from the lines: findings other than
        hiring contractor mismatches, and those."""
        lines = {str(r): e for r, e in (lines or {}).items()}
        marks = [m for e in lines.values() for m in e["must"]]
        self.key["trackers"][project.project.name] = {
            "file": LNI.TRACKER_FILE, "outcome": outcome, "reason_contains": reason, "header_row": header_row,
            "lines": lines, "lni_only": sorted(lni_only), "not_due": dict(not_due or {}),
            "never_ask": sorted(str(x) for x in never_ask),
            "summary": {"marked": sum(1 for m in marks if not m.startswith("hiring:")),
                        "hiring": sum(1 for m in marks if m.startswith("hiring:")), "lni_only": len(lni_only)}}

    def gap(self, gap_id: str, issue: int, files, currently: str, why: str, rows_currently=None):
        """``rows_currently``: {file: {sheet row: [rules raised today]}} for a gap on payroll lines (the file's
        outcome is already right): a line ending exactly so is tolerated, any other result is a difference. (A
        tracker line's gap says its marks today as ``lines_currently``: eval_lni.GAPS.)"""
        g = {"id": gap_id, "issue": issue, "files": list(files), "currently": currently, "why": why}
        if rows_currently:
            g["rows_currently"] = {f: {str(r): sorted(v) for r, v in rows.items()}
                                   for f, rows in rows_currently.items()}
        self.key["known_gaps"].append(g)


def wage_xlsx(path: Path, rows=WAGE_ROWS, header=WAGE_HEADER):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wage rates"
    ws.append(header)
    for r in rows:
        ws.append(r)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_payroll(path: Path, lines, header=SV.HEADER, title_rows=(), order=None, extra=None, sheet="Payroll",
                  footer=()):
    """Writes the lines (Builder.lines) as an LCPtracker-style export. ``order``: the header in another column order;
    ``extra``: (header, value) of one more column, at the end; ``title_rows``: rows above the header; ``footer``: rows
    below the lines, as written (not in the key: any flag on them is a false flag). Returns ({sheet row:
    expectation}, the header's row)."""
    order = list(order or header)
    pos = [header.index(h) for h in order]
    head = order + ([extra[0]] if extra else [])
    rows, key = [], {}
    first = len(title_rows) + 2
    for _, entries in lines:
        for r, exp in entries:
            key[first + len(rows)] = exp
            rows.append([r[i] for i in pos] + ([extra[1]] if extra else []))
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".csv":
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            for t in title_rows:
                w.writerow(t)
            w.writerow(head)
            for r in rows + [list(x) for x in footer]:
                w.writerow([v.isoformat() if isinstance(v, (dt.date, dt.datetime)) else
                            ("" if v is None else v) for v in r])
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet
        for t in title_rows:
            ws.append(list(t))
        ws.append(head)
        for r in rows + [list(x) for x in footer]:
            ws.append(r)
        wb.save(path)
    return key, len(title_rows) + 1


def write_tracker(path: Path, lines, info, staff_notes=()):
    """A compliance tracker as staff keep it: the contractor table on the first tab, a Project Info tab.
    ``staff_notes``: [(sheet row, column role)] cells that already hold a staff member's note. Returns {sheet row:
    line} (the header is row 1)."""
    from openpyxl.comments import Comment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = LNI.TRACKER_SHEET
    ws.append(LNI.HEADER)
    rows = {}
    for i, values in enumerate(lines):
        ws.append(values)
        rows[i + 2] = values
    for r, role in staff_notes:
        ws.cell(r, LNI.COLUMN_OF[role]).comment = Comment(LNI.STAFF_NOTE, "Compliance staff")
    pi = wb.create_sheet("Project Info")
    for r in info:
        pi.append(list(r))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return rows


def tracker_key(cases):
    """{sheet row: expectation} for LNI.CASES written in order from row 2."""
    out = {}
    for i, c in enumerate(cases):
        out[i + 2] = {"case": c["case"], "area": c["area"], "must": sorted(f"{r}:{k}" for r, k in c["must"]),
                      "ids": c["ids"], "not_due": c["not_due"], "staff_note": c["staff_note"]}
        gap = LNI.GAPS.get(c["case"])
        if gap:
            out[i + 2]["known_gap"] = gap["id"]
    return out


def clean_week(week: dt.date):
    """Six correct lines for one week: state laborers and an electrician, overtime paid right, and two union
    laborers (one with Saturday hours paid as overtime, as the agreement requires)."""
    L, U = SV.line, SV.U
    sat_ot = (0, 0, 0, 0, 0, 8, 0)
    return [
        ("Clean: state laborer, 40 h", "control", [L(week=week, emp="a")], set(), set()),
        ("Clean: state laborer, 2 h overtime", "control",
         [L(week=week, emp="b", st=(8, 8, 8, 8, 6, 0, 0), ot=(2, 0, 0, 0, 0, 0, 0))], set(), set()),
        ("Clean: electrician, 32 h", "control",
         [L(week=week, emp="c", craft="Electricians", cls="Journey Level", base=60.0, tp=25.0,
            st=(8, 8, 8, 8, 0, 0, 0))], set(), set()),
        ("Clean: state laborer, part week", "control", [L(week=week, emp="d", st=(4, 4, 0, 0, 0, 0, 0))], set(), set()),
        ("Clean: union laborer, 40 h", "control", [L(week=week, emp="e", **U)], set(), set()),
        ("Clean: union laborer, Saturday as overtime", "control",
         [L(week=week, emp="f", st=(8, 8, 8, 8, 0, 0, 0), ot=sat_ot, **U)], set(), set()),
    ]


def county_only_cases():
    """Lines whose Jurisdiction names only a county: each is checked once as each wage type the project's own tables
    give (state and union here; David, Nov 12, 2026), and each note says which ("Checked as state: ..."). Built so
    each pass shows: a laborer paid state rates fails only the union check; an electrician (a state rate only) paid
    short fails the state check; a laborer paid union rates passes both."""
    L, U = SV.line, SV.U
    county = "Pierce County"
    return [
        ("County only: laborer paid state rates (fails the union check)", "jurisdiction",
         [L(jur=county)], {"R07", "R08"}, set()),
        ("County only: electrician paid below the state rate (fails the state check)", "jurisdiction",
         [L(jur=county, craft="Electricians", cls="Journey Level", base=55.0, tp=25.0)], {"R08"}, {"R07", "S02"}),
        ("County only: laborer paid union rates (passes both: control)", "control",
         [L(**{**U, "jur": county})], set(), set()),
    ]


def bad_value_cases():
    """Lines with a value that isn't what its column needs (item 1, upload checks; David, Nov 13, 2026: the line is
    marked S15 and left out of the other checks, the rest of the payroll is checked). Fixed dates: 2062 is always in
    the future, 2019 always more than 5 years ago. The controls are values the checks must still read."""
    L = SV.line
    return [
        ("Hours typed as a word", "values", [L(**{"ST Mon": "eight"})], {"S15"}, set()),
        ("Base rate typed as words", "values", [L(**{"Base Hourly": "see attached"})], {"S15"}, set()),
        ("Week End Date isn't a date", "values", [L(**{"Week End Date": "Week 37"})], {"S15"}, set()),
        ("Week End Date in the future (year typo)", "values", [L(week=dt.date(2062, 9, 5))], {"S15"}, set()),
        ("Week End Date more than 5 years ago", "values", [L(week=dt.date(2019, 9, 7))], {"S15"}, set()),
        ("Hours with a unit, '8 hrs' (control: read as 8)", "control", [L(**{"ST Mon": "8 hrs"})], set(), set()),
        ("Base rate as text with a $ (control)", "control", [L(**{"Base Hourly": "$40.00"})], set(), set()),
        ("Week End Date as text m/d/yyyy (control)", "control", [L(**{"Week End Date": "9/5/2026"})], set(), set()),
        ("A word for nothing, 'N/A', in a money column (control: left alone)", "control",
         [L(**{"Total DT Wages": "N/A"})], set(), set()),
    ]


def build(dest: Path) -> dict:
    """Writes the input snapshot into ``dest`` (empty) and returns the answer key."""
    if dest.exists() and any(dest.iterdir()):
        raise SystemExit("the snapshot folder must be empty")
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copy(AGENT / "config.yaml", dest / "config.yaml")
    b = Builder(dest)
    by_name = {c[0]: c for c in SV.CASES}
    subset = [by_name[n] for n in SUBSET]

    # A: every seeded case in one payroll export, as LCPtracker writes it
    p = b.project("Eval A - Seeded rules")
    key, head = write_payroll(p.upload / "Week 37.xlsx", b.lines(SV.CASES))
    b.expect(p, "Week 37.xlsx", "reviewed", head, key)

    # B: three clean weeks, three files: nothing may be flagged
    p = b.project("Eval B - Clean weeks")
    for week in (dt.date(2026, 8, 30), dt.date(2026, 9, 6), dt.date(2026, 9, 13)):
        f = f"Week ending {week:%m-%d}.xlsx"
        key, head = write_payroll(p.upload / f, b.lines(clean_week(week)))
        b.expect(p, f, "reviewed", head, key)

    # C: the same lines in the layouts exports really come in
    p = b.project("Eval C - Messy layouts")
    order = SV.HEADER[4:8] + SV.HEADER[:4] + SV.HEADER[8:]              # employee and trade columns first
    key, head = write_payroll(p.upload / "Certified payroll with title rows.xlsx", b.lines(subset),
                              title_rows=[["Certified Payroll Report"], [CONTRACTOR], []],
                              order=order, extra=("Notes", "invented"), sheet="Certified Payroll")
    b.expect(p, "Certified payroll with title rows.xlsx", "reviewed", head, key)
    key, head = write_payroll(p.upload / "Payroll export.csv", b.lines(subset))
    b.expect(p, "Payroll export.csv", "reviewed", head, key)
    deep = [["Certified Payroll Report"], [CONTRACTOR], ["Week ending 09-06-2026"], [],
            ["Prepared by: payroll office"], ["Page 1 of 1"], [], ["Lines follow"]]
    key, head = write_payroll(p.upload / "Header on row 9.xlsx", b.lines(subset), title_rows=deep)
    b.expect(p, "Header on row 9.xlsx", "reviewed", head, key)
    b.gap("header-below-row-7", 67, ["Eval C - Messy layouts/Header on row 9.xlsx"], "not recognized",
          "The header is looked for in rows 1-7 only (payroll_export.HEADER_ROWS), so the file isn't recognized as a "
          "payroll and stays in Uploads (staff are told \"1 uploaded file was not recognized\"); #67 is to find the "
          "header from the data.")

    # D: a payroll with no wage rates in the project: not checked, stays in Uploads, the emails say why
    p = b.project("Eval D - No wage rates", wages=False)
    key, head = write_payroll(p.upload / "Week 36.xlsx", b.lines(subset[:3]))
    b.expect(p, "Week 36.xlsx", "skipped", head, reason="no wage rates file in the project",
             reason_code="no_wage_rates")

    # E: a Jurisdiction naming only a county: checked as each of the project's wage types (David, Nov 12, 2026)
    p = b.project("Eval E - Jurisdiction unclear")
    cases = county_only_cases()
    notes = {cases[0][0]: (["Checked as union"], ["Checked as state"]),
             cases[1][0]: (["Checked as state"], [])}
    key, head = write_payroll(p.upload / "Week 35.xlsx", b.lines(cases, notes=notes))
    b.expect(p, "Week 35.xlsx", "reviewed", head, key)

    # F: a wage rates file without its Jurisdiction column: the rates' type can't be told, so the payroll isn't
    # checked and stays in Uploads (David, Nov 7, 2026)
    p = b.project("Eval F - Wage rates without Jurisdiction", wages=False)
    wage_xlsx(p.upload / "Wage rates 2026.xlsx", rows=[r[:6] + r[7:] for r in WAGE_ROWS],
              header=WAGE_HEADER[:6] + WAGE_HEADER[7:])
    key, head = write_payroll(p.upload / "Week 34.xlsx", b.lines(subset[:3]))
    b.expect(p, "Week 34.xlsx", "skipped", head, reason="Jurisdiction column not found in the wage rates file",
             reason_code="no_jurisdiction_column")

    # I-K: upload checks (item 1 of the reliability plan; docs/evals.md "Plan"; agent/upload_checks.py, measured with
    # payroll.upload_checks on, as config.yaml has it since Nov 15, 2026)
    # I: values that aren't what their column needs: the line is marked (S15), the rest of the file checked; a totals
    # footer row and readable values written as text are left alone
    p = b.project("Eval I - Unreadable values")
    st = SV.HEADER.index("ST Hours")
    totals = ["Totals"] + [None] * (st - 1) + [40 * 8] + [None] * (len(SV.HEADER) - st - 1)
    key, head = write_payroll(p.upload / "Week 38.xlsx", b.lines(bad_value_cases()), footer=[[], totals])
    b.expect(p, "Week 38.xlsx", "reviewed", head, key)

    # J: a payroll with a header and no lines: nothing to check, not "reviewed (0 lines)"
    p = b.project("Eval J - No payroll lines")
    key, head = write_payroll(p.upload / "Week 39.xlsx", [])
    b.expect(p, "Week 39.xlsx", "skipped", head, reason="no lines below its header row",
             reason_code="no_payroll_lines")

    # K: two exports pasted into one sheet (the header again below the first export's lines): not checked as one
    p = b.project("Eval K - Two exports in one file")
    two = b.lines(clean_week(SV.WEEK)[:2]) + [("(the header again)", [(list(SV.HEADER), None)])] + \
        b.lines(clean_week(SV.WEEK)[2:4])
    key, head = write_payroll(p.upload / "Week 40.xlsx", two)
    key = {r: e for r, e in key.items() if e is not None}         # the second header is no payroll line
    b.expect(p, "Week 40.xlsx", "skipped", head, reason="header row appears again",
             reason_code="repeated_header")

    # L (#53): a worker's name the tokenizer never learns, in each place it used to cross readable: an unread payroll
    # column, a second sheet, a sheet's name, a document property, a file's name, a folder's name and a wage table's
    # note. The payroll is still reviewed (clean lines: no flag), the wage file still read, and the name nowhere in
    # the run (run_evals.py's privacy check)
    p = b.project("Eval L - Names never learnt", wages=False)
    wage = p.upload / f"Wage rates from {NEVER_LEARNT}.xlsx"
    wage_xlsx(wage)
    wb = openpyxl.load_workbook(wage)
    from openpyxl.comments import Comment
    wb.active["B2"].comment = Comment(f"checked with {NEVER_LEARNT}", "Wage team")
    wb.save(wage)
    f = "Week 41.xlsx"
    key, head = write_payroll(p.upload / f, b.lines(clean_week(SV.WEEK)), extra=("Foreman", NEVER_LEARNT))
    wb = openpyxl.load_workbook(p.upload / f)
    wb.create_sheet("Notes").append([f"Crew lead: {NEVER_LEARNT}"])
    wb.create_sheet(f"{NEVER_LEARNT} hours").append([8])
    wb.properties.creator = NEVER_LEARNT
    wb.properties.lastModifiedBy = NEVER_LEARNT
    wb.save(p.upload / f)
    b.expect(p, f, "reviewed", head, key)
    folder = p.files / f"{NEVER_LEARNT} notes"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "readme.txt").write_text("Invented notes for the evaluation harness.\n")

    # G: the intent/affidavit check: a tracker against invented L&I records (tools/eval_lni.py serves them as L&I's
    # live lookup would): every kind of finding once, correct lines, lines not due yet, an L&I-only contractor
    p = b.project("Eval G - Tracker checks", wages=False)
    write_tracker(p.tracker / LNI.TRACKER_FILE, [c["line"] for c in LNI.CASES],
                  [["L&I Project ID", LNI.PROJECT], ["Prime Intent", LNI.PRIME_INTENT]],
                  staff_notes=[(i + 2, c["staff_note"]) for i, c in enumerate(LNI.CASES) if c["staff_note"]])
    due = {}
    for c in LNI.CASES:
        if c["not_due"]:
            due[c["not_due"]] = due.get(c["not_due"], 0) + 1
    b.tracker(p, "checked", tracker_key(LNI.CASES), lni_only=[r["CoName"] for r in LNI.LNI_ONLY], not_due=due)
    for g in LNI.GAPS.values():
        b.key["known_gaps"].append({"id": g["id"], "issue": g["issue"], "files": [], "trackers": [p.project.name],
                                    "currently": "; ".join(g["currently"]), "lines_currently": g["currently"],
                                    "why": g["why"]})

    # H: a tracker whose Project Info says no L&I filing is required: not checked, L&I not asked about it
    p = b.project("Eval H - No L&I filing", wages=False)
    write_tracker(p.tracker / LNI.TRACKER_FILE, LNI.NOT_APPLICABLE_LINES,
                  [["L&I Project ID", LNI.PROJECT + 1], ["Notes", "No L&I Filing Required"]])
    b.tracker(p, "not applicable", {i + 2: {"case": "Not checked (no L&I filing required)", "area": "control",
                                            "must": [], "ids": {}, "not_due": None, "staff_note": None}
                                    for i in range(len(LNI.NOT_APPLICABLE_LINES))},
              never_ask=[LNI.PROJECT + 1, *(r[4] for r in LNI.NOT_APPLICABLE_LINES)],
              reason="Project Info says no L&I filing required")
    return b.key
