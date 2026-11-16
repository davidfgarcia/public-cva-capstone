# GENERATED COPY of code/filekinds/cva_filekinds/project_data.py: edit the source, then run code/filekinds/sync.py
"""The Project Data kind: each project's cover sheet (#151 step 3; David, Nov 12, 2026).

The cover sheet is a Google Sheet (or .xlsx) whose name starts with "Project Data", directly in the project's folder
in the **staff** drive (the folder the tracker step finds: pii_tokenizer/staff_trackers.py). The tokenizer reads it
on the staff side and the CVA needs only one thing from it: the project's **status**. Closed and Not started projects
are not checked on the Saturday run and are greyed out in Run tasks; Active or blank = active. Only a status code
crosses to AWS (never the word typed, the file's name or anything else in it); the workbook itself is never copied.

History: from Nov 7 to Nov 12, 2026 the agent read the Jurisdiction Override of a "Project Data" workbook in the CVA's
own project folder, and the tokenizer kept that label and its value readable in any workbook. Since Nov 12 the
project's own wage tables give the wage types (PR "own-wage-tables"), and the cover-sheet PR removed both.
"""
from __future__ import annotations

import re

from .core import FileKind, Recognition

# The file, by name: its name starts with "Project Data" (any case), and no old-copy word; "~$" lock files never
NAME_RE = re.compile(r"\s*project\s+data(?![a-z])", re.I)
OLD_COPY_RE = re.compile(r"(?<![a-z])(copy|old|archived?|backup)(?![a-z])", re.I)   # as the tracker step's

# The status codes (the only thing that crosses to AWS; the agent's agent/runner.py reads them: keep in step)
ACTIVE, CLOSED, NOT_STARTED = "active", "closed", "not_started"
STATUS_WORDS = {"active": ACTIVE, "closed": CLOSED, "not started": NOT_STARTED}
SHOWN = {ACTIVE: "Active", CLOSED: "Closed", NOT_STARTED: "Not started"}
NOT_SET = {"", "n a", "na", "placeholder"}         # empty, N/A, <PLACEHOLDER>: not set, so active

# Why a project counts as what it counts as (codes, AWS side too)
OK = "ok"                      # a status word that is one of the three
BLANK = "blank"                # the Status row is there, its value empty / N/A: active
NONE = "none"                  # no "Project Data..." workbook in the staff project folder: active
SEVERAL = "several"            # more than one: not guessed, active
NO_STATUS = "no_status"        # no Status row: active
UNRECOGNISED = "unrecognised"  # a word that isn't Active, Closed or Not started: active
CONFLICT = "conflict"          # two Status rows that disagree: active
UNREADABLE = "unreadable"      # not a workbook the tokenizer can read (.xls, damaged...): active
NO_FOLDER = "no_folder"        # the project's staff folder isn't known (no tracker match, tracker_source.txt "none")
ERROR = "error"                # the staff drive couldn't be read this time: the last status read is used
WHYS = (OK, BLANK, NONE, SEVERAL, NO_STATUS, UNRECOGNISED, CONFLICT, UNREADABLE, NO_FOLDER, ERROR)

HEAD_ROWS = 200                # rows of the tab read for the Status row
LABEL_COLS = (1, 2)            # without the template's header row, the label must be in column A or B


def norm(v) -> str:
    return " ".join(re.findall(r"[a-z0-9]+", str(v if v is not None else "").lower()))


def is_cover_sheet_name(name: str) -> bool:
    """A file name (with its extension, or a Google Sheet's title) that can be the project's cover sheet."""
    name = str(name or "")
    if name.startswith("~$") or not NAME_RE.match(name):
        return False
    stem = name.rsplit(".", 1)[0] if re.search(r"\.[A-Za-z0-9]{2,5}$", name) else name
    return not OLD_COPY_RE.search(stem)


def is_project_data_name(stem: str) -> bool:
    """Kept for the registry's tests (a name starting "Project Data", no old-copy word)."""
    return is_cover_sheet_name(stem)


def status_of(value) -> tuple[str, str]:
    """(status code, why) for the Status row's value. Only plain text is matched; capitals and spaces don't count."""
    if value is None or (isinstance(value, str) and norm(value) in NOT_SET):
        return ACTIVE, BLANK
    if isinstance(value, str) and norm(value) in STATUS_WORDS:
        return STATUS_WORDS[norm(value)], OK
    return ACTIVE, UNRECOGNISED


def _pick_tab(sheets):
    """The tab named "Project Data" if there is one, else the first tab (the template has one tab)."""
    for title, rows in sheets:
        if norm(title) == "project data":
            return rows
    return sheets[0][1] if sheets else []


def read_status(sheets) -> tuple[str, str, object]:
    """``sheets``: [(tab title, [(row number, {column: value})])] in tab order (the first rows of each).
    Returns (status code, why, the value as typed or None). The value as typed is for the staff side only.

    The Status row: a cell that says exactly "Status" (the template's Field column, or its Key column's ``status``).
    With the template's header row (``Field | Key | Value | ...``) the value is in its Value column; without it, the
    label must be in column A or B and the value is the next cell to its right that isn't the key."""
    rows = _pick_tab(sheets)
    value_col, found = None, []
    for _r, vals in rows:
        words = {c: norm(v) for c, v in vals.items() if isinstance(v, str)}
        if "key" in words.values() and "value" in words.values():
            value_col = min(c for c, w in words.items() if w == "value")
            continue
        label = next((c for c in sorted(words) if words[c] == "status"), None)
        if label is None:
            continue
        if value_col is not None:
            if value_col <= label:
                continue                     # "Status" right of the Value column: a note, not the row's label
            value = vals.get(value_col)
        else:
            if label not in LABEL_COLS:
                continue                     # a "Status" column header in some table: not the cover sheet's row
            value = next((vals[c] for c in sorted(vals) if c > label and words.get(c) != "status"), None)
        found.append(value)
    if not found:
        return ACTIVE, NO_STATUS, None
    results = [status_of(v) for v in found]
    for (code, why), v in zip(results, found):
        if why == UNRECOGNISED:
            return ACTIVE, UNRECOGNISED, v
    if len({code for code, _ in results}) > 1:
        return ACTIVE, CONFLICT, None
    code, why = results[0]
    return code, why, found[0]


def recognise(heads, sheet_titles, active=0):
    """The first sheet with the template's header row (Field | Key | Value) in its first rows (shape only: the
    tokenizer finds the file by its name in the staff folder)."""
    for i, rows in enumerate(heads):
        for r, vals in rows:
            words = {norm(v) for v in vals.values() if isinstance(v, str)}
            if {"field", "key", "value"} <= words:
                return Recognition("project_data", i, r, frozenset())
    return None


PROJECT_DATA = FileKind(
    kind="project_data",
    recognise=recognise,
    columns=(),
    above_header=(),
    version=2,
    notes="The cover sheet in the staff project folder: only its Status row is read (staff side, by the tokenizer); "
          "a status code crosses to AWS, nothing of the workbook itself.",
)
