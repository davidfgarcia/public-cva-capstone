"""Upload checks: what a payroll export must hold before its lines are checked (item 1 of the reliability plan,
docs/evals.md "Plan"; #65). Behind ``payroll.upload_checks.enabled`` in config.yaml (on since Nov 15, 2026; off unless it says true: then
nothing here runs and the review is as before).

David's decisions (Nov 13, 2026):
  - **A bad value marks its line, the rest of the payroll is checked.** A number column holding text that isn't a
    number ("eight", "see attached"), or a Week End Date that isn't a date or is out of range, gets an orange S15 on
    the cell, and that line is left out of every other check (a base rate read as 0.00 would otherwise fail R08).
    Readable text stays readable: "8 hrs", "$40.00", "1,400.00", "40.00/hr", a date written "9/5/2026".
  - **Out of range:** more than ``future_days`` after the run's date, or more than ``max_week_age_years`` before it
    (a year typed wrong: 2062 or 2002 for 2026).
  - **The header row again below the lines** (two exports pasted into one sheet) skips the whole payroll: checked as
    one, a worker's hours would be counted twice.
  - A payroll with no lines below its header is skipped too ("reviewed (0 lines)" said nothing useful).

Skips raise ``payroll.PayrollSkipped`` with a run-log reason code (agent/events.py REASONS); the payroll stays in
Uploads and both emails give the words. The words name a column and a row, never a cell's text: on AWS a cell can
hold a token, and in the staff's files a name (HANDOFF section 6).
"""
from __future__ import annotations

import datetime as dt
import re

from .cva_filekinds.payroll_export import COLUMN_PATTERNS, LINE_ROLES
from .cva_filekinds.payroll_export import EMPTY_WORDS as _EMPTY_WORDS, TEXT_ROLES as _TEXT_ROLES
from .util import norm_header, to_date, to_float

DEFAULTS = {"enabled": False, "future_days": 6, "max_week_age_years": 5}

# The payroll's text columns; every other column the review reads holds a number (hours) or an amount (money)
TEXT_ROLES = _TEXT_ROLES              # (the registry's: the tokenizer keeps text in the others only if number-like, #53)
NUMBER_ROLES = tuple(r for r in COLUMN_PATTERNS if r not in TEXT_ROLES)

# A number written as text that the review reads as that number (util.to_float): "$1,400.00", "8 hrs", "8.",
# "40.00/hr", "-$12". Not: "(12.50)", ".5" or "12.50-", which to_float would read as nothing or with the wrong sign
# (independent review, Nov 13: such a cell is read as 0.00, the false R08 this check is there to prevent).
_UNIT = r"(?:h|hr|hrs|hour|hours|/\s*(?:h|hr|hour))?\.?"
NUMBER_TEXT = re.compile(rf"-?\s*\$?\s*-?\s*(?:\d{{1,3}}(?:,\d{{3}})+|\d+)(?:\.\d*)?\s*{_UNIT}", re.I)
# Text that means "nothing here" in an export: an accounting zero, "N/A". Left alone, as an empty cell is (S08 says
# when a value the review needs is missing); the review reads it as 0, as before.
EMPTY_WORDS = _EMPTY_WORDS            # (the registry's: kept readable by the tokenizer, #53)


def settings(cfg) -> dict | None:
    """The upload checks' settings, or None when they are off."""
    s = {**DEFAULTS, **(((cfg.get("payroll") or {}).get("upload_checks")) or {})}
    return s if s.get("enabled") is True else None


def _blank(x) -> bool:
    return x is None or (isinstance(x, str) and not x.strip())


def says_empty(v) -> bool:
    """A text cell that means "nothing here" (EMPTY_WORDS): not this check's business."""
    return isinstance(v, str) and re.sub(r"\s+", " ", v).strip().lower() in EMPTY_WORDS


def is_number(v) -> bool:
    """True when the review reads the cell as the number it shows: a number, or text in one of NUMBER_TEXT's forms
    that util.to_float reads. A time of day ("8:00" formatted as a time) is not a number of hours."""
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if not isinstance(v, str):
        return False
    s = v.strip()
    return NUMBER_TEXT.fullmatch(s) is not None and to_float(s) is not None      # as the review reads it


def _label(header, cols, role) -> str:
    """The column's header as the payroll writes it ("ST Mon"), for the note."""
    i = cols.get(role)
    h = header[i] if i is not None and i < len(header) else None
    return re.sub(r"\s+", " ", str(h)).strip() if not _blank(h) else role.replace("_", " ").title()


def is_payroll_line(v, cols) -> bool:
    """A line the review checks: not empty and not a footer / notes line (review_rows skips those the same way)."""
    g = lambda k: v[cols[k]] if k in cols and cols[k] < len(v) else None
    if all(_blank(x) for x in v):
        return False
    return not all(_blank(g(k)) for k in LINE_ROLES if k in cols)


def line_problems(v, cols, header, today: dt.date, s: dict) -> list:
    """[(role, words)] for one payroll line: number columns that don't hold a usable number, a Week End Date that
    isn't a date or is out of range. Empty cells are not this check's business (S08 says "is empty")."""
    g = lambda k: v[cols[k]] if k in cols and cols[k] < len(v) else None
    out = []
    week = g("week")
    if "week" in cols and not _blank(week):
        d = None if isinstance(week, bool) or isinstance(week, (int, float)) else to_date(week)
        label = _label(header, cols, "week")
        if d is None:
            out.append(("week", f"{label} isn't a date"))
        elif d > today + dt.timedelta(days=int(s["future_days"])):
            out.append(("week", f"{label} {d:%m-%d-%Y} is after the run's date ({today:%m-%d-%Y}): check the year"))
        elif d < _years_before(today, int(s["max_week_age_years"])):
            out.append(("week", f"{label} {d:%m-%d-%Y} is more than {int(s['max_week_age_years'])} years before "
                                f"the run's date ({today:%m-%d-%Y}): check the year"))
    for role in NUMBER_ROLES:
        x = g(role)
        if role in cols and not _blank(x) and not says_empty(x) and not is_number(x):
            out.append((role, f"{_label(header, cols, role)} isn't a number"))
    return out


def _years_before(d: dt.date, n: int) -> dt.date:
    try:
        return d.replace(year=d.year - n)
    except ValueError:                                  # Feb 29
        return d.replace(year=d.year - n, day=28)


def file_problem(rows, cols, header):
    """(reason code, words) when the whole payroll can't be checked as it stands, or None. ``rows``: [(sheet row,
    values)] below the header."""
    lines = [(rn, v) for rn, v in rows if is_payroll_line(v, cols)]
    if not lines:
        return "no_payroll_lines", "the payroll has no lines below its header row"
    # The header again: a line holding most of the header's own labels for the columns the review reads, in any
    # column (two exports pasted into one sheet; the second may have its columns in another order). Compared as
    # labels, so a tokenized cell (the employee's) doesn't hide it.
    known = {norm_header(header[i]) for i in cols.values() if i < len(header) and not _blank(header[i])}
    need = max(3, (len(known) + 1) // 2)
    for rn, v in lines:
        same = len(known & {norm_header(x) for x in v if isinstance(x, str) and x.strip()})
        if same >= need:
            return ("repeated_header", f"the header row appears again at row {rn} (two exports in one file?): upload "
                                       "each export as its own file")
    return None


def line_marks(rows, cols, header, today: dt.date, s: dict) -> dict:
    """{sheet row: [(role, words)]} for the lines with a value the review can't use."""
    out = {}
    for rn, v in rows:
        if is_payroll_line(v, cols):
            p = line_problems(v, cols, header, today, s)
            if p:
                out[rn] = p
    return out
