"""The Project Info kind: the tracker workbook's first tab with "info" in its title (prime intent, L&I project ID,
"No L&I Filing Required"). Moved here from agent/intent.py and pii_tokenizer/doubt.py (vNext PR 3). Behaviour
unchanged: the agent and the tokenizer still read it differently where they did (inventory D10-D12, pinned by tests).
"""
from __future__ import annotations

import re

from .core import FileKind, Recognition

# The labels the agent reads the value after (agent/intent.py project_info): "prime intent" anywhere in the label;
# the L&I project ID
PRIME_INTENT_LABEL = "prime intent"
PROJECT_ID_LABEL_RE = re.compile(r"l\s*&\s*i project id|lni project id")
# Either label (the tokenizer keeps such a label and the value after it readable): built from the two above
INFO_LABEL_RE = re.compile(PRIME_INTENT_LABEL + "|" + PROJECT_ID_LABEL_RE.pattern)
# "No L&I Filing Required": the agent searches the tab's whole text for it (agent/intent.py is_applicable); the
# tokenizer keeps readable only a cell that says just that (D11)
NO_FILING_SEARCH_RE = re.compile(r"no l\s*&\s*i filing required", re.I)
NO_FILING_RE = re.compile(r"\s*no l\s*&\s*i filing required\W*", re.I)


def is_info_title(title: str) -> bool:
    return "info" in title.lower()


def keep_info_cell(row_vals: dict, col: int) -> bool:
    """Project Info, as the tokenizer keeps it: a cell that says only "No L&I Filing Required", or, when the row's
    first non-blank cell is a label above, that cell and the next non-blank one (D10: the agent counts an empty-string
    cell as a cell)."""
    v = row_vals.get(col)
    if isinstance(v, str) and NO_FILING_RE.fullmatch(v):
        return True
    filled = [c for c in sorted(row_vals) if row_vals[c] is not None and str(row_vals[c]).strip()]
    if not filled or not isinstance(row_vals[filled[0]], str) or not INFO_LABEL_RE.search(row_vals[filled[0]].lower()):
        return False
    return col in filled[:2]             # the label and the value after it


def recognise(heads, sheet_titles, active=0):
    """The first tab whose title has "info" in it, as the agent finds it (the tokenizer skips a tab that is already a
    wage, agent or tracker tab: D12)."""
    i = next((n for n, t in enumerate(sheet_titles) if is_info_title(t)), None)
    return None if i is None else Recognition("project_info", i, 0, frozenset())


PROJECT_INFO = FileKind(
    kind="project_info",
    recognise=recognise,
    columns=(),
    above_header=("label rows: a Prime Intent or L&I Project ID label and the value after it",
                  "a cell saying No L&I Filing Required"),
    version=1,
    notes="Read by label, not by column: crosses_to_aws doesn't apply (keep_info_cell decides).",
)
