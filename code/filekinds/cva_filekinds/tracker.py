"""The tracker kind: the staff's contractor table (Contractor, Intent and Affidavit columns), and the agent's own
sheets and notes in it. Moved here from code/compliance_agent/agent/intent.py and
code/tokenizer/pii_tokenizer/doubt.py (vNext PR 3); both now import it. Behaviour unchanged: where the two components
recognise a tracker differently, they still do (docs/filekinds_inventory.md D9, D16, pinned by tests).
"""
from __future__ import annotations

import re

from .core import AS_IS, Column, FileKind, Recognition, norm_header

# The columns the intent/affidavit check reads (whole header, lower case, spaces collapsed; agent/intent.py
# locate_table takes the first column per role, a column used once). The tokenizer keeps the header row and these
# columns readable and turns every other text cell of a tracker tab into a text token.
TRACKER_ROLES = {
    "status": r"status",
    "contractor": r"contractor",
    "hiring": r"hiring contractor",
    "intent_flag": r"intent (filed|approved)",
    # "Intent ID", or "L&I Intent" (Atrovan's tracker), "L&I Intent #", "Intent Number"
    "intent_id": r"(l ?& ?i )?intent (id|#|number|no\.?)|l ?& ?i intent",
    "affidavit_flag": r"affidavit filed",
    "affidavit_id": r"(l ?& ?i )?affidavit (id|#|number|no\.?)|l ?& ?i affidavit",
    # (No "Affidavit Good" or "CP Status" since Nov 7, 2026: staff's Affidavit Good is about the document's contents,
    # not its ID, so the agent ignores it, and CP Status was read only for it. Their cells are text tokens like any
    # column the agent doesn't read.)
    "start": r"(we )?start date",
    "end": r"end date",
}
# Columns the agent reads only for a few fixed words: a cell there stays readable only when it is a plain value (no
# formula) whose whole text is one of them. None now (Affidavit Good and CP Status were the two). Kept for the next.
TRACKER_VOCAB: dict = {}
# The header is in rows 1-11 (agent/intent.py locate_table)
HEADER_ROWS = 11
# The sheets the agent writes into a tracker and reads back (config.yaml intent.audit_sheet, review_log_sheet): as
# before doubt mode. (A reviewed payroll's sheets are in payroll workbooks, never in doubt mode.)
AGENT_SHEETS = {"audit", "review log"}
# The agent's name on the notes it writes (agent/intent.py), and the name it used before the rename (cleared too)
AUTHOR = "Compliance Validation Agent"
OLD_AUTHORS = ("Compliance Agent",)
AGENT_AUTHORS = {AUTHOR, *OLD_AUTHORS}

_TYPES = {"status": "code", "contractor": "text", "hiring": "text", "intent_flag": "flag", "intent_id": "id",
          "affidavit_flag": "flag", "affidavit_id": "id", "start": "date", "end": "date"}


def is_tracker_header(headers) -> bool:
    """A row of normalised headers that is a contractor table's: a "contractor" cell and a cell with "intent" or
    "affidavit" in it. Each component normalises the row its own way first (the agent every cell, the tokenizer
    text cells only)."""
    headers = list(headers)
    return "contractor" in headers and any("intent" in h or "affidavit" in h for h in headers)


def locate_tracker_roles(rows):
    """(header row, {role: column}) of one sheet's contractor table, as the tokenizer finds it (agent/intent.py
    locate_table, from text cells only), or None. ``rows``: [(row number, {col: value})]."""
    for r, vals in rows:
        if r > HEADER_ROWS:
            break
        hdr = {c: norm_header(v) for c, v in vals.items() if isinstance(v, str)}
        if is_tracker_header(hdr.values()):
            used = {}
            for role, pat in TRACKER_ROLES.items():
                for c in sorted(hdr):
                    if re.fullmatch(pat, hdr[c]) and c not in used.values():
                        used[role] = c
                        break
            return r, used
    return None


def locate_tracker_header(rows):
    """(header row, the agent's columns, {column: TRACKER_VOCAB pattern}) of one sheet, as the tokenizer finds it
    (agent/intent.py locate_table, from text cells only), or None. ``rows``: [(row number, {col: value})]."""
    found = locate_tracker_roles(rows)
    if found is None:
        return None
    r, used = found
    return r, set(used.values()), {c: TRACKER_VOCAB[role] for role, c in used.items() if role in TRACKER_VOCAB}


# Strict trackers (#128, tokenizer setting vnext.strict_trackers): in the columns the intent check reads, a text value
# stays readable only when it is one the check can use, of its column's type; anything else (a name typed into a
# Status, ID or date cell) becomes a text token. Contractor and Hiring Contractor hold business names: kept as before.
# Flags: agent/util.py truthy's words, and the usual "no" words (any other text reads as "not ticked", as a token does).
FLAG_WORDS = frozenset({"true", "yes", "y", "x", "filed", "approved", "done", "complete",
                        "false", "no", "n", "n/a", "na", "-", "none", "not filed", "pending"})
# IDs: what agent/intent.py _num_id reads (util.to_float: "$" and "," removed, then a leading number), followed only by
# punctuation, more numbers or a few ID words ("1234567 (amended)", "1234567 / 1234568", "1234567-1"); or NO_ID_WORDS.
# Any other words after the number (a name) make the cell a token: the agent then says "not a valid number". With
# intent.strict_ids on (agent/intent.py _cell_id) it also says so for a letter or other character inside the number
# ("88000O8"); a number followed by a space or separator and words ("8800050-Filed") is still read by the number.
NO_ID_WORDS = frozenset({"no", "n/a", "na", "false"})
ID_WORDS = frozenset({"amended", "amendment", "amend", "pending", "revised", "rev", "filed", "approved", "submitted",
                      "received", "and", "or", "new", "old", "original", "tbd", "na", "n", "a", "see", "also",
                      "intent", "affidavit", "id", "no", "number", "replaced", "replaces", "by", "was"})
_ID_LEAD = re.compile(r"\s*#?\s*-?\d+(?:\.\d+)?")
# Dates: digits and separators, month names, times, and the words agent/intent.py _START_NOT_YET reads (with any
# punctuation after them, as it allows: "TBD!", "Not started?").
DATE_WORDS = frozenset({"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
                        "january", "february", "march", "april", "june", "july", "august", "september", "october",
                        "november", "december", "am", "pm", "t", "not", "yet", "started", "tbd", "to", "be",
                        "determined", "pending", "open", "ongoing", "n", "a", "na", "none"})
_DATE_CHARS = re.compile(r"[A-Za-z0-9\s/.,:\-!?()]*")
# Status: besides the status words, common words staff add around them ("Not started - awaiting NTP", "Not started
# (TBD)"), so agent/intent.py _not_started still finds "not started" in them. No digits; no other words (a name).
STATUS_FILLER = frozenset({"awaiting", "waiting", "ntp", "notice", "to", "proceed", "per", "for", "the", "of", "on",
                           "in", "at", "a", "an", "and", "or", "yet", "contract", "signed", "mobilization",
                           "mobilizing", "schedule", "scheduled", "tbd", "pending", "job", "site", "work", "phase",
                           "complete", "completed", "start", "started", "not", "until", "after", "before",
                           "approval", "paperwork", "submitted", "received", "subcontract", "agreement", "date"})
_STATUS_CHARS = re.compile(r"[^\d]*")


def status_ok(v: str, words) -> bool:
    """A Status the strict mode keeps readable: exactly one of ``words`` (the agent's status words: the tokenizer's
    list and the Settings page's), or, with no digits, made only of their words and STATUS_FILLER's; case, spacing
    and punctuation ignored."""
    norm = " ".join(v.split()).lower()
    if not norm:
        return True
    phrases = {" ".join(str(w).split()).lower() for w in words}
    if norm in phrases:
        return True
    vocab = {x for p in phrases for x in re.findall(r"[^\W\d_]+", p)} | STATUS_FILLER
    found = re.findall(r"[^\W\d_]+", norm)
    return bool(found) and bool(_STATUS_CHARS.fullmatch(norm)) and set(found) <= vocab


def id_ok(v: str) -> bool:
    """An Intent or Affidavit ID cell the strict mode keeps readable (see ID_WORDS)."""
    low = " ".join(v.split()).lower()
    if low in NO_ID_WORDS:
        return True
    s = v.replace("$", "").replace(",", "")
    m = _ID_LEAD.match(s)
    if not m:
        return False
    rest = re.sub(r"^-\d{1,2}\b", "", s[m.end():])          # "1234567-1": a short suffix right after the ID
    # anything else: more IDs (6 digits or more: not a phone number's or a date's short groups), ID words, punctuation
    return all(len(d) >= 6 for d in re.findall(r"\d+", rest)) and \
        all(w in ID_WORDS for w in re.findall(r"[^\W\d_]+", rest.lower()))


def value_ok(role: str, v, status_words=()) -> bool:
    """Is this value of a tracker's read column (by role) one the strict mode keeps readable? Non-text always: a
    number, date or yes/no can't be a name (numbers and dates in other cells: #112)."""
    if not isinstance(v, str):
        return True
    s = v.strip()
    if not s or role in ("contractor", "hiring"):
        return True
    low = " ".join(s.split()).lower()
    if role == "status":
        return status_ok(s, status_words)
    if role in ("intent_flag", "affidavit_flag"):
        return low in FLAG_WORDS
    if role in ("intent_id", "affidavit_id"):
        return id_ok(s)
    if role in ("start", "end"):
        return bool(_DATE_CHARS.fullmatch(s)) and all(w in DATE_WORDS for w in re.findall(r"[a-z]+", low))
    return False


def recognise(heads, sheet_titles, active=0):
    """The first tab, in tab order, with a contractor table's header (shape only; the tokenizer keeps the role columns
    of every such tab readable, the agent reads the first: D9)."""
    for i, rows in enumerate(heads):
        found = locate_tracker_header(rows)
        if found:
            return Recognition("tracker", i, found[0], frozenset(found[1]))
    return None


TRACKER = FileKind(
    kind="tracker",
    recognise=recognise,
    columns=tuple(Column(role, pat, _TYPES[role], role == "contractor", AS_IS) for role, pat in TRACKER_ROLES.items()),
    version=3,
    notes="Contractor names cross as is (business information). 'hiring' is read by the intent check since #121 "
          "(compared with the L&I intent's hiring company); 'end' is located but not read yet.",
)
