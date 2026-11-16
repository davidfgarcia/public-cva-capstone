"""The payroll export kind: a certified payroll (LCPtracker export, one line per worker and week), as a workbook or a
CSV. Moved here from code/compliance_agent/agent/payroll.py, agent/uploads.py and code/tokenizer/pii_tokenizer/doubt.py
(vNext PR 4); all three import it under their old names. Behaviour unchanged: the three recognition tests still differ
(docs/filekinds_inventory.md D1, D2, pinned by tests), and nothing here changes what crosses to AWS (no narrowing
until phase 3: #53).

The three tests, all built on the same pieces:
  - Uploads (agent/uploads.py ``_payroll_header``): one of the first SNIFF_ROWS rows (from the first SNIFF_SHEETS
    sheets, joined; a CSV's first rows) where ``map_columns`` finds an employee name or ID, the week-end date, and the
    classification, straight-time hours or craft (``is_payroll_header``).
  - The reader (agent/payroll.py ``_find_header``): rows 1-HEADER_ROWS of the first sheet, a cell that says exactly
    "classification"; otherwise the payroll fails.
  - The tokenizer (pii_tokenizer/doubt.py ``is_payroll_sheet``): the same test on the first tab's rows as it reads
    them; a match takes the whole workbook out of doubt mode (#53).
"""
from __future__ import annotations

import re

from .core import AS_IS, TOKEN, Column, FileKind, Recognition, norm_header

# The files that can be payroll exports (agent/uploads.py, agent/payroll.py)
PAYROLL_EXTS = frozenset({".xlsx", ".xlsm", ".csv"})
# The reader and the tokenizer look for the header in rows 1-7 of the first sheet
HEADER_ROWS = 7
# Uploads looks at the first 15 rows of each of the first 3 sheets (joined, then the first 15 of the list)
SNIFF_ROWS = 15
SNIFF_SHEETS = 3
# The reader's header cell: exactly "classification" (whitespace and case aside)
HEADER_CELL = "classification"

DAYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")

# The columns the payroll review reads: {role: pattern the whole normalised header must match}, in the reader's
# order (map_columns takes, per role, the first column that matches)
COLUMN_PATTERNS = {
    "project": r"project name", "contractor": r"contractor name", "jurisdiction": r"jurisdiction",
    "week": r"week end date", "employee": r"employee name",
    "emp_id": r"employee id \(lcptracker.*|employee id lcptracker.*",
    "craft": r"craft", "classification": r"classification",
    "st_hours": r"st hours", "ot_hours": r"ot hours", "dt_hours": r"dt hours",
    "total_hours": r"total hours", "base": r"base hourly", "ot_hourly": r"overtime hourly",
    "dt_hourly": r"doubletime hourly", "cash_fringe_rate": r"cash fringe rate",
    "tp_fringe_rate": r"3rd party fringe rate", "st_pw_total": r"st prevailing wage rate total",
    "total_st_wages": r"total st wages", "total_ot_wages": r"total ot wages",
    "total_dt_wages": r"total dt wages", "grand_total": r"grand total wages",
    "gross_this": r"gross pay this project", "gross_all": r"gross pay all projects", "per_diem": r"per diem",
}
for _p in ("st", "ot", "dt"):
    for _d in DAYS:
        COLUMN_PATTERNS[f"{_p}_{_d}"] = rf"{_p} {_d}"
del _p, _d

# The columns the review needs: a payroll without one is reviewed, the checks that need it skipped, and staff told
# (agent/payroll.py _review_one)
IMPORTANT = ("jurisdiction", "week", "craft", "classification", "base", "st_hours", "ot_hours", "ot_hourly",
             "cash_fringe_rate", "tp_fringe_rate", "grand_total")


def map_columns(header):
    """{role: 0-based position} of a header row (a list of cell values). Every cell counts, whatever its type (D2)."""
    hdr = [norm_header(h) for h in header]
    cols = {}
    for key, pat in COLUMN_PATTERNS.items():
        for i, h in enumerate(hdr):
            if h and re.fullmatch(pat, h):
                cols[key] = i
                break
    return cols


def is_payroll_header(header) -> bool:
    """Uploads' test on one row: an employee name or ID, the week-end date, and the classification, straight-time
    hours or craft."""
    cols = map_columns(list(header))
    return ("employee" in cols or "emp_id" in cols) and "week" in cols and \
        ("classification" in cols or "st_hours" in cols or "craft" in cols)


def payroll_header(rows) -> bool:
    """agent/uploads.py _payroll_header: one of the first SNIFF_ROWS rows of ``rows`` (lists of cell values) passes
    ``is_payroll_header``."""
    return any(is_payroll_header(r) for r in rows[:SNIFF_ROWS])


def _is_header_cell(v) -> bool:
    return norm_header(v) == HEADER_CELL


def find_header(rows):
    """agent/payroll.py _find_header: (row number, values) of the first row, among rows 1-HEADER_ROWS, with a cell
    that says exactly "classification". ``rows``: (row number, list of values) in order; read only up to the match
    (the agent passes a generator over openpyxl rows, so rows below the header are never touched here)."""
    for r, vals in rows:
        if r > HEADER_ROWS:
            break
        if any(_is_header_cell(v) for v in vals):
            return r, vals
    raise ValueError("Could not find the header row (no 'Classification' column) in the payroll file")


def is_payroll_sheet(rows) -> bool:
    """pii_tokenizer/doubt.py is_payroll_sheet: a "Classification" header in rows 1-HEADER_ROWS. ``rows``: [(row
    number, {col: value})] as the tokenizer reads them."""
    return any(r <= HEADER_ROWS and any(_is_header_cell(v) for v in vals.values()) for r, vals in rows)


def recognise(heads, sheet_titles, active=0):
    """The first tab, when one of its rows 1-HEADER_ROWS is a header both the reader and Uploads accept: an exact
    "classification" cell, an employee name or ID and the week-end date (shape only: the three components still
    decide by their own tests, D1)."""
    if not heads:
        return None
    for r, vals in heads[0]:
        if r > HEADER_ROWS:
            break
        if not vals:
            continue
        row = [vals.get(c) for c in range(1, max(vals) + 1)]
        cols = map_columns(row)
        if any(_is_header_cell(v) for v in row) and is_payroll_header(row):
            return Recognition("payroll_export", 0, r, frozenset(i + 1 for i in cols.values()))
    return None


_TYPES = {"project": "text", "contractor": "text", "jurisdiction": "text", "week": "date", "employee": "name",
          "emp_id": "id", "craft": "text", "classification": "text"}
_HOURS = {"st_hours", "ot_hours", "dt_hours", "total_hours", "dt_sat", "dt_sun"}


# The reviewed payroll's "Review decisions" sheet (agent/ai_review.py writes it, agent/knowledge.py read_decisions
# reads staff's decisions back): its name, its header row (the agent's own), and the columns read_decisions reads.
# Moved here from agent/knowledge.py (#53): the tokenizer keeps exactly these readable in a payroll workbook
DECISIONS_SHEET = "Review decisions"
DECISION_HEADERS = ["#", "Rule", "Wage type", "Craft / classification", "Determination / Jurisdiction", "Contractor",
                    "Lines", "Example", "Decision (Keep / Dismiss)", "Note", "AI suggestion",
                    "AI explanation and sources", "Pattern ID"]
DECISIONS_READ = ("Pattern ID", "Decision (Keep / Dismiss)", "Rule", "Wage type", "Craft / classification",
                  "Contractor", "Note", "Example")
DECISIONS_HEADER_ROWS = 10          # read_decisions: the header is the first of the sheet's first 10 rows with "Pattern ID"


def decisions_header(rows):
    """(header row, {column: header}) of a Review decisions sheet, as agent/knowledge.py read_decisions finds it (the
    first of rows 1-10 with a cell that says exactly "Pattern ID"), or None. ``rows``: [(row number, {col: value})]."""
    for r, vals in rows:
        if r > DECISIONS_HEADER_ROWS:
            break
        if "Pattern ID" in [str(v if v is not None else "") for v in vals.values()]:
            return r, vals
    return None


def lines_and_columns(header):
    """The payroll reader's columns on a header row ({1-based column: role}, map_columns: the first column per role;
    not 'project', which is located but never read) and the columns that make a row a payroll line (agent/payroll.py
    review_rows: a row with none of employee, craft, classification or week is a footer or notes line, skipped)."""
    cols = {i + 1: role for role, i in map_columns(header).items() if role not in UNREAD_ROLES}
    return cols, frozenset(c for c, role in cols.items() if role in LINE_ROLES)


LINE_ROLES = ("employee", "craft", "classification", "week")
UNREAD_ROLES = ("project",)
# The payroll's text columns (agent/upload_checks.py); every other column the review reads holds a number (hours) or an
# amount (money)
TEXT_ROLES = frozenset({"project", "contractor", "jurisdiction", "week", "employee", "emp_id", "craft",
                        "classification"})
# Text that means "nothing here" in an export (agent/upload_checks.py says_empty): an accounting zero, "N/A"
EMPTY_WORDS = frozenset({"-", "–", "—", "$ -", "$-", "-0-", "n/a", "na", "n.a.", "none"})
_NUMBER_UNITS = re.compile(r"\b(?:h|hr|hrs|hour|hours|usd)\b", re.I)


def number_text_ok(v) -> bool:
    """#53: a text in an hours or money column the tokenizer may keep readable: one of EMPTY_WORDS, or no letters but
    the review's units ("8 hrs", "$1,400.00", "8:00"). Anything else (a name typed into ST Hours) is a
    token; the upload checks then find it "not a usable number", as they would the text."""
    if not isinstance(v, str):
        return True
    s = re.sub(r"\s+", " ", v).strip().lower()
    return s in EMPTY_WORDS or not re.search(r"[^\W\d_]", _NUMBER_UNITS.sub("", s))


def _type(role: str) -> str:
    if role in _TYPES:
        return _TYPES[role]
    if role in _HOURS or role.split("_")[0] in ("st", "ot", "dt") and role.split("_")[-1] in DAYS:
        return "number"
    return "money"


PAYROLL_EXPORT = FileKind(
    kind="payroll_export",
    recognise=recognise,
    # The employee's name and ID cross as tokens (the tokenizer's field rules, tokenizer_config.yaml); every other
    # column as is. Until phase 3 the tokenizer doesn't read this: a workbook its payroll test accepts crosses whole,
    # every column and tab, under the field rules and name scrubbing only (#53).
    columns=tuple(Column(role, pat, _type(role), role in IMPORTANT, TOKEN if role in ("employee", "emp_id") else AS_IS)
                  for role, pat in COLUMN_PATTERNS.items()),
    version=2,
    notes="'project' is located but never read. The wage type also uses config.yaml payroll.jurisdiction_patterns, "
          "jurisdiction_map and the Project Data override.",
)
