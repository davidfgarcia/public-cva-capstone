"""
"If in doubt, assume it is sensitive" (David, Nov 4, 2026): in a workbook that isn't a payroll export or a wage
table, the text of every cell the agent doesn't read is replaced by a reversible text token (``TXT-…``,
tokens.Keyring.text_token) on its way to AWS, and comes back exactly as written. A worker's name the tokenizer has
never learnt (a note, a "Section 3 Employees" column, a trainees sheet) then has nowhere to travel.

What the agent reads, and so what stays readable, mirrors code/compliance_agent (tests/test_doubt.py checks that
these patterns stay in step with the agent's own):

  tracker table  the sheet whose header row (one of rows 1-11) has "Contractor" and an Intent or Affidavit column
                 (agent/intent.py locate_table): the header row, and the columns of HEADER_ROLES below
  Project Info   the first sheet with "info" in its name (agent/intent.py project_info): a "Prime Intent" or
                 "L&I Project ID" label and the value after it, and a cell that says "No L&I Filing Required"
  agent's sheets "Audit", "Review Log" and the like, which the agent itself writes: as before

Everything else in such a workbook (other columns, rows above the header, other sheets, cell notes) is a text
token. Numbers and dates stay: they can't be names. A formula keeps its place: only its cached text (and its own
text, when it holds a quoted string) is replaced, so the staff file gets its formulas back.

Payroll exports and wage tables (a sheet the agent reads as one) keep the earlier handling unless
vnext.fail_closed_tables is on (#53): their own columns are tokenized by header (tokenizer_config.yaml) and known
names are replaced elsewhere. With it on, they get the same rule (``sheet_rules``, fail_closed):
  payroll export the first tab with the reader's header (a "Classification" cell in rows 1-7): the header row's cells
                 of the reader's columns (cva_filekinds.payroll_export map_columns: the first column per role), and
                 those columns below it on payroll lines only (a row with an employee, craft, classification or week:
                 LINE_ROLES; any other row is a footer or a note, which the review skips: its text is a token)
  Review decisions  in a reviewed payroll, its header row and the columns agent/knowledge.py read_decisions reads
                 (cva_filekinds.payroll_export DECISIONS_READ)
  merged table   the agent's own Files/Wage rates/<Kind>.xlsx: its "rates" header (MERGED_TABLE_HEADER) and columns;
                 its "sources" sheet as the agent writes it (rule "agent")
  wage table     as above (#101); a sheet that looks like one but has a personal-data or hours column, or passes only the
                 pre-#101 wage test, is no longer taken out of doubt mode: every text in it is a token (its personal-data
                 columns tokenized by header first, as before)
"""
from __future__ import annotations

from dataclasses import dataclass, field

# The wage table kinds come from the file-kind registry (code/filekinds, copied here as pii_tokenizer/cva_filekinds);
# the agent imports the same definitions (agent/cva_filekinds), so the two can't drift apart (vNext PR 2)
from .cva_filekinds.merged_wage_table import MERGED_TABLE_HEADER, is_merged_table  # noqa: F401 - used by xlsx/tests
from .cva_filekinds.wage_table import (BASE_HDR, CLASS_HDR, CRAFT_HDR, LOCATION_RE, WA_COUNTIES,  # noqa: F401
                                       WAGE_HEADERS, locate_wage_header)
from .cva_filekinds.wage_table import HEADER_ROWS as WAGE_HEADER_ROWS
from .cva_filekinds.wage_table import PAYROLLISH as _PAYROLLISH
from .cva_filekinds.wage_table import WAGE_COL as _WAGE_COL
from .cva_filekinds.wage_table import old_is_wage_sheet as _old_is_wage_sheet
from .cva_filekinds.wage_table import is_amount, value_ok, wage_columns  # noqa: F401 - used by xlsx (#118)

# The tracker, Project Info and Project Data kinds: from the registry (vNext PR 3; the agent imports the same)
from .cva_filekinds.core import norm_header  # noqa: F401 - agent/util.py norm_header; used here and by tests
from .cva_filekinds.project_info import INFO_LABEL_RE, NO_FILING_RE, keep_info_cell  # noqa: F401
from .cva_filekinds.tracker import (AGENT_AUTHORS, AGENT_SHEETS, TRACKER_ROLES, TRACKER_VOCAB,  # noqa: F401
                                    locate_tracker_header, locate_tracker_roles)
from .cva_filekinds.tracker import value_ok as tracker_value_ok  # noqa: F401 - used by xlsx (#128)
from .cva_filekinds.tracker import HEADER_ROWS as TRACKER_HEADER_ROWS

# The payroll export kind (vNext PR 4; agent/payroll.py and agent/uploads.py import the same): the reader's header
# test, rows 1-7 of the first tab
from .cva_filekinds.payroll_export import HEADER_ROWS as PAYROLL_HEADER_ROWS
from .cva_filekinds.payroll_export import is_payroll_sheet  # noqa: F401 - used below and by xlsx/tests
from .cva_filekinds.payroll_export import (DECISIONS_READ, DECISIONS_SHEET, HEADER_CELL, decisions_header,  # noqa: F401
                                           lines_and_columns, number_text_ok)
from .cva_filekinds.payroll_export import TEXT_ROLES as PAYROLL_TEXT_ROLES
# WAGE_HEADER_ROWS (40, agent/wages.py parse_lcp_export: rows 1-40): from the registry, above
HEAD_ROWS = max(TRACKER_HEADER_ROWS, PAYROLL_HEADER_ROWS, WAGE_HEADER_ROWS)


@dataclass
class SheetRule:
    kind: str                            # tracker | wage | info | agent | other
    header_row: int | None = None        # tracker, wage: the header row
    keep_cols: set = field(default_factory=set)      # tracker, wage: the agent's columns (1-based)
    vocab: dict = field(default_factory=dict)        # tracker: {column: TRACKER_VOCAB pattern} of the fixed-word ones
    # wage, strict (#118, vnext.strict_wage_tables): header cells kept, rate columns, money columns, {column: the
    # reader's role} (cva_filekinds.wage_table.wage_columns); strict False = v1
    # tracker, strict (#128, vnext.strict_trackers): roles = {column: the intent check's role}; strict False = v1
    strict: bool = False
    head_cols: set = field(default_factory=set)
    rate_cols: set = field(default_factory=set)
    money_cols: set = field(default_factory=set)
    roles: dict = field(default_factory=dict)
    status_words: frozenset = frozenset()            # tracker, strict: the Status words kept readable (#128)
    line_cols: frozenset = frozenset()               # payroll (#53): a row is a payroll line if one of these has a value
    head_only_read: bool = False                     # tracker (#53, fail_closed): header row: only the read columns' cells
    head_labels: dict = field(default_factory=dict)  # payroll (#53): {column: its header}: a repeated header row is kept


# The sheets the agent writes into a merged wage table besides its rates (agent/wages.py): kept as written (rule "agent")
MERGED_AGENT_SHEETS = {"sources"}


def _payroll_rule(head) -> SheetRule | None:
    """#53 (vnext.fail_closed_tables): the rule for a payroll export's first tab: its header row (the reader's,
    cva_filekinds.payroll_export find_header) with the reader's columns, and the payroll-line columns."""
    for r, vals in head:
        if r > PAYROLL_HEADER_ROWS:
            break
        if vals and any(norm_header(v) == HEADER_CELL for v in vals.values()):
            header = [vals.get(c) for c in range(1, max(vals) + 1)]
            cols, line = lines_and_columns(header)
            return SheetRule("payroll", r, set(cols), line_cols=line, roles=cols,
                             head_labels={c: norm_header(vals.get(c)) for c in cols})
    return None


def _decisions_rule(head) -> SheetRule:
    """#53: a reviewed payroll's Review decisions sheet: its header row's cells and the columns read_decisions reads;
    no header found: nothing kept (the agent finds none either)."""
    found = decisions_header(head)
    if not found:
        return SheetRule("other")
    r, vals = found
    keep = {c for c, v in vals.items() if isinstance(v, str) and v in DECISIONS_READ}
    return SheetRule("decisions", r, keep)


# locate_tracker_header: from the registry (cva_filekinds.tracker), above


# is_payroll_sheet (agent/payroll.py _find_header: a "Classification" header in rows 1-7): from the registry, above


# The wage tables (CRAFT_HDR, CLASS_HDR, BASE_HDR, WAGE_HEADERS, locate_wage_header, MERGED_TABLE_HEADER) are defined
# in the registry (cva_filekinds.wage_table, cva_filekinds.merged_wage_table) and imported above.


def is_wage_sheet(rows) -> bool:
    return locate_wage_header(rows) is not None


# _old_is_wage_sheet, _PAYROLLISH, WA_COUNTIES, LOCATION_RE and is_merged_table: from the registry, above


# keep_info_cell: from the registry, above (the Jurisdiction Override is no longer kept readable: #151 step 3)


def sheet_rules(titles: list[str], heads: list[list], active: int = 0, personal=None,
                strict_wage: bool = False, strict_tracker: bool = False,
                fail_closed: bool = False) -> tuple[bool, list[SheetRule]]:
    """(doubt mode for this workbook, one rule per sheet), the sheets in TAB order (as the agent sees them).
    ``heads``: each sheet's first HEAD_ROWS rows as [(row number, {col: value})]. ``personal(header)``: True for a
    header the tokenizer's field rules treat as personal data (an employee name or ID column...).
    - A wage table (the active tab, which the wage readers read) whose header has no personal-data column: doubt
      mode, with only the agent's wage columns kept readable on that tab (rule "wage"); the other tabs as usual.
    - The agent's own merged wage tables, a payroll export (the first tab), or a wage table with a personal-data
      column: doubt mode off, as before (their personal columns are tokenized by header).
    ``strict_wage`` (vNext setting vnext.strict_wage_tables, #118; False = v1): the pre-#101 wage test no longer takes
    a sheet out of doubt mode (Craft in column B, or a wage-like sheet with hours or a personal-data column that isn't
    a payroll: a "Crew" column of names crossed as is), and a wage tab keeps only what the readers use (xlsx._keep).
    ``strict_tracker`` (vnext.strict_trackers, #128; False = v1): a tracker tab's read columns keep a text only when it
    is a value of the column's type (xlsx._keep, cva_filekinds.tracker.value_ok).
    ``fail_closed`` (vnext.fail_closed_tables, #53; False = v1): no workbook leaves doubt mode. A payroll export, the
    agent's merged wage tables and a payroll-like wage sheet get the rules in this module's docstring."""
    at = min(active, len(heads) - 1) if heads else 0
    wage = locate_wage_header(heads[at]) if heads else None
    if wage:
        hdr = [v for r, vals in heads[at] if r == wage[0] for v in vals.values() if isinstance(v, str)]
        if (personal is not None and any(personal(v) for v in hdr)) or any(
                _PAYROLLISH.search(norm_header(v)) and not _WAGE_COL.fullmatch(norm_header(v)) for v in hdr):
            # a payroll-like sheet (a personal-data column, or hours / weeks): decided exactly as before Nov 8, 2026,
            # never more open than that (review of PR #101)
            wage = None
    if fail_closed:
        return True, _fail_closed_rules(titles, heads, at, wage, strict_wage, strict_tracker)
    if heads and (is_merged_table(heads[at]) or
                  (not wage and (is_payroll_sheet(heads[0]) or (not strict_wage and _old_is_wage_sheet(heads[at]))))):
        return False, [SheetRule("other") for _ in titles]
    return True, _rules(titles, heads, at, wage, strict_wage, strict_tracker)


def _fail_closed_rules(titles, heads, at, wage, strict_wage, strict_tracker) -> list[SheetRule]:
    """#53: the rules of a workbook that would have left doubt mode (see sheet_rules), and of every other one."""
    if heads and is_merged_table(heads[at]):
        return [SheetRule("merged", 1, set(range(1, len(MERGED_TABLE_HEADER) + 1))) if i == at else
                SheetRule("agent") if t.strip().lower() in MERGED_AGENT_SHEETS else SheetRule("other")
                for i, t in enumerate(titles)]
    if heads and not wage and is_payroll_sheet(heads[0]):
        first = _payroll_rule(heads[0]) or SheetRule("other")
        return [first] + [_decisions_rule(h) if t.strip() == DECISIONS_SHEET else SheetRule("other")
                          for t, h in zip(titles[1:], heads[1:])]
    rules = _rules(titles, heads, at, wage, strict_wage, strict_tracker)
    for rule in rules:
        if rule.kind == "tracker":
            rule.head_only_read = True        # (review of the #53 PR: "Ana notes" over an unread column crossed)
    return rules


def _rules(titles, heads, at, wage, strict_wage, strict_tracker) -> list[SheetRule]:
    rules, info_done = [], False
    for i, (title, head) in enumerate(zip(titles, heads)):
        t = title.strip().lower()
        found = locate_tracker_header(head)
        if wage and i == at:
            if strict_wage:
                header = next((vals for r, vals in heads[at] if r == wage[0]), {})
                kept_head, rates, money, roles = wage_columns(header)
                rules.append(SheetRule("wage", wage[0], wage[1], strict=True, head_cols=kept_head, rate_cols=rates,
                                       money_cols=money, roles=roles))
            else:
                rules.append(SheetRule("wage", wage[0], wage[1]))
        elif t in AGENT_SHEETS:
            rules.append(SheetRule("agent"))
        elif found:
            if strict_tracker:
                roles = {c: role for role, c in (locate_tracker_roles(head) or (0, {}))[1].items()}
                rules.append(SheetRule("tracker", found[0], found[1], found[2], strict=True, roles=roles))
            else:
                rules.append(SheetRule("tracker", found[0], found[1], found[2]))
        elif "info" in t and not info_done:
            info_done = True
            rules.append(SheetRule("info"))
        else:
            rules.append(SheetRule("other"))
    return rules
