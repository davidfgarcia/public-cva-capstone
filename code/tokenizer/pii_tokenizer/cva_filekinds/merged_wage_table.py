# GENERATED COPY of code/filekinds/cva_filekinds/merged_wage_table.py: edit the source, then run code/filekinds/sync.py
"""The merged wage table kind: the agent's own Files/Wage rates/<Kind>.xlsx (a "rates" sheet whose row 1 is these
fields, plus a "sources" sheet). Moved here from agent/wages.py FIELDS and pii_tokenizer/doubt.py MERGED_TABLE_HEADER
(vNext PR 2). The agent reads ``worksheets[0]``; the tokenizer checks the active tab (D8, unchanged).

In v1 the tokenizer handles these as the agent's own output: no doubt mode, personal columns tokenized by their header
(tokenizer_config.yaml), known names replaced. Narrowing them to these columns is phase 3.
"""
from __future__ import annotations

from .core import AS_IS, Column, FileKind, Recognition, norm_header, sheet_index

MERGED_TABLE_HEADER = ["kind", "craft", "classification", "wage_decision", "mod", "issue_date", "expire_date", "basic",
                       "fringe", "total", "ot_total", "dt_total", "std_hours", "vac_pto", "county", "holiday_code",
                       "ot_code", "notes", "source_type", "source_file", "jurisdiction"]

_TYPES = {"kind": "code", "wage_decision": "code", "mod": "number", "issue_date": "date", "expire_date": "date",
          "basic": "money", "fringe": "money", "total": "money", "ot_total": "money", "dt_total": "money",
          "std_hours": "number", "vac_pto": "money", "holiday_code": "code", "ot_code": "code", "source_type": "code",
          "jurisdiction": "code"}


def is_merged_table(rows) -> bool:
    """Row 1 of the sheet is exactly MERGED_TABLE_HEADER (pii_tokenizer/doubt.py)."""
    first = next((vals for r, vals in rows if r == 1), {})
    return [norm_header(first.get(c)) for c in range(1, len(MERGED_TABLE_HEADER) + 1)] == MERGED_TABLE_HEADER


def recognise(heads, sheet_titles, active=0):
    if not heads:
        return None
    at = sheet_index(heads, active)
    if not is_merged_table(heads[at]):
        return None
    return Recognition("merged_wage_table", at, 1, frozenset(range(1, len(MERGED_TABLE_HEADER) + 1)))


MERGED_WAGE_TABLE = FileKind(
    kind="merged_wage_table",
    recognise=recognise,
    columns=tuple(Column(f, f, _TYPES.get(f, "text"), f in ("kind", "classification"), AS_IS)
                  for f in MERGED_TABLE_HEADER),
    version=1,
    notes="v1: the whole workbook outside doubt mode (field rules and scrubbing); 'crosses' describes phase 3's "
          "intended narrowing, not what the tokenizer does today.",
)
