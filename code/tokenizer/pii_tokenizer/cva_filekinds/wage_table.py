# GENERATED COPY of code/filekinds/cva_filekinds/wage_table.py: edit the source, then run code/filekinds/sync.py
"""The wage table kind: an LCPtracker wage export ("lcp" layout) or a table with its header in row 1, the union table
or the firm's wage table template ("table" layout). Moved here from code/compliance_agent/agent/wages.py and
code/tokenizer/pii_tokenizer/doubt.py (vNext PR 2); both now import it. Behaviour unchanged: where the two components
recognised these tables differently, they still do (docs/filekinds_inventory.md D3-D8, each pinned by a test).

Wage-table headers (David, Nov 8, 2026): "Trade" replaces "Craft" and "Base Rate" replaces "Basic Rate" from now on;
the old names are still read, so earlier files and the union reference keep working.
"""
from __future__ import annotations

import re

from .core import AS_IS, Column, FileKind, Recognition, norm_header, sheet_index

CRAFT_HDR = r"craft|trade"
CLASS_HDR = r"(job )?classification"
BASE_HDR = r"basic.*|base rate"
# The columns each reader uses (whole header, lower case, spaces collapsed): agent/wages.py parse_lcp_export and
# parse_union_table. The tokenizer keeps exactly these columns of a wage table readable and turns every other text
# cell into a text token (review of PR #101): a name typed into some other column of a wage table never reaches AWS.
LCP_HEADERS = {"craft": CRAFT_HDR, "cls": CLASS_HDR, "wd": "wage decision", "dates": "issue date.*",
               "mod": "mod", "basic": BASE_HDR, "fringe": "fringes", "total": "total rate", "ot": "ot rate",
               "dt": "dt rate", "std": "std day|standard day", "jur": "jurisdiction"}
TABLE_HEADERS = {"craft": CRAFT_HDR, "cls": CLASS_HDR, "basic": BASE_HDR, "fringe": "fringes|fringe rate",
                 "total": "total rate|total wage", "ot": "overtime total", "dt": "doubletime total",
                 "std": "standard hours", "issue": "issue date", "effective": "effective date",
                 "expire": "expire date|expiration date", "vac": "vac/pto.*", "jur": "jurisdiction",
                 "wd": "decision|wage decision", "county": "county"}
WAGE_HEADERS = sorted(set(LCP_HEADERS.values()) | set(TABLE_HEADERS.values()))
WAGE_COL = re.compile("|".join(f"(?:{p})" for p in WAGE_HEADERS))

# An LCPtracker export's header is in rows 1-40 of the active tab; a table's in row 1
HEADER_ROWS = 40

# A header that says the sheet is a payroll rather than a wage table (hours worked, a week or a work date), unless it
# is one of the wage columns ("Standard Hours"). The tokenizer then decides the sheet as before Nov 8, 2026.
PAYROLLISH = re.compile(r"\b(weeks?|hours?|hrs|work ?date|period|pay ?date)\b")

WA_COUNTIES = ["Adams", "Asotin", "Benton", "Chelan", "Clallam", "Clark", "Columbia", "Cowlitz", "Douglas", "Ferry",
               "Franklin", "Garfield", "Grant", "Grays Harbor", "Island", "Jefferson", "King", "Kitsap", "Kittitas",
               "Klickitat", "Lewis", "Lincoln", "Mason", "Okanogan", "Pacific", "Pend Oreille", "Pierce", "San Juan",
               "Skagit", "Skamania", "Snohomish", "Spokane", "Stevens", "Thurston", "Wahkiakum", "Walla Walla",
               "Whatcom", "Whitman", "Yakima"]
# The LCPtracker export's "Location:" line above the header: kept readable only when it names a county (a WA county,
# or words ending in "County"); the agent reads nothing else from it (the agent itself takes any text after the
# colon: D6, intended)
LOCATION_RE = re.compile(r"Location:\s*(?:(?:" + "|".join(WA_COUNTIES) + r")\b(?:\s+County)?|[A-Z][A-Za-z.' ]{0,30}?"
                         r"\s+County)(?:\s*,\s*[A-Z]{2})?\s*")

_TYPES = {"craft": "text", "cls": "text", "wd": "code", "dates": "date", "mod": "number", "basic": "money",
          "fringe": "money", "total": "money", "ot": "money", "dt": "money", "std": "number", "jur": "code",
          "issue": "date", "effective": "date", "expire": "date", "vac": "money", "county": "text"}
_REQUIRED = {"lcp": {"craft", "cls"}, "table": {"cls", "basic"}}


def _columns(layout: str, headers: dict) -> tuple:
    return tuple(Column(role, pat, _TYPES[role], role in _REQUIRED[layout], AS_IS, (layout,))
                 for role, pat in headers.items())


def locate_wage_header_layout(rows):
    """(header row, the columns the agent reads, layout) of a wage table on one sheet, as the tokenizer has found it
    since #101: an LCPtracker export (column A says Craft or Trade, a column with "classification" in it; rows 1-40)
    or a table with its header in row 1 (a Classification or Job Classification column and a Base Rate or Basic...
    column). None if it isn't one. ``rows``: [(row number, {col: value})]."""
    for r, vals in rows:
        if r > HEADER_ROWS:
            break
        hdr = {c: norm_header(v) for c, v in vals.items() if isinstance(v, str)}
        a = hdr.get(1, "")
        lcp = bool(re.fullmatch(CRAFT_HDR, a)) and any("classification" in h for h in hdr.values())
        table = r == 1 and any(re.fullmatch(CLASS_HDR, h) for h in hdr.values()) \
            and any(re.fullmatch(BASE_HDR, h) for h in hdr.values())
        if lcp or table:
            return r, {c for c, h in hdr.items() if WAGE_COL.fullmatch(h)}, "lcp" if lcp else "table"
    return None


def locate_wage_header(rows):
    """(header row, the columns the agent reads), or None (pii_tokenizer/doubt.py)."""
    found = locate_wage_header_layout(rows)
    return found[:2] if found else None


def old_is_wage_sheet(rows) -> bool:
    """The wage-table test before Nov 8, 2026 (Craft first; or "classification" and a "basic" column in row 1), kept
    only so the tokenizer decides its fallback exactly as before (doubt.sheet_rules; D7, intended)."""
    for r, vals in rows:
        if r > HEADER_ROWS:
            break
        texts = [str(v).strip().lower() for v in vals.values() if isinstance(v, str)]
        first = vals.get(min(vals)) if vals else None
        if isinstance(first, str) and first.strip().lower() == "craft" and any("classification" in t for t in texts):
            return True
        if r == 1 and "classification" in texts and any("basic" in t for t in texts):
            return True
    return False


def recognise(heads, sheet_titles, active=0):
    """A wage table's header shape on the active tab (the tab the wage readers read): locate_wage_header only.
    **Not the tokenizer's decision:** doubt.sheet_rules also refuses a wage-shaped sheet with an hours/week header
    (PAYROLLISH) or a personal-data header (and then decides it with the payroll test and the pre-#101 fallback), so
    this says "wage_table" for sheets the tokenizer handles otherwise (tests/test_filekinds.py pins the difference).
    Nothing decides what crosses to AWS from it yet (phase 3)."""
    if not heads:
        return None
    at = sheet_index(heads, active)
    found = locate_wage_header_layout(heads[at])
    if not found:
        return None
    return Recognition("wage_table", at, found[0], frozenset(found[1]), found[2])


# --- #118, strict wage tables (vNext setting vnext.strict_wage_tables in the tokenizer, off by default) ---------
# Only what the wage readers use stays readable. Below the header: in a Jurisdiction column only its words (any row);
# otherwise only on a *rate row* (a Base/Basic or Total amount: parse_lcp_export skips a row without one), and there
# only a value of the column's type: an amount in a money or number column, a date in a date column, a decision code,
# a county; Trade and Classification any text (the agent needs it). A footer ("Prepared by ..."), a note, or a name
# typed into any of these columns then becomes a text token instead of crossing readable.
RATE_ROLES = ("basic", "total")
# An amount as text, as agent/util.py to_float reads a rate ("$54.80", "1,234.5", "$50.00/hr"); nothing else
AMOUNT_RE = re.compile(r"\s*\$?\s*-?(?:\d[\d,]*(?:\.\d+)?|\.\d+)\s*\*?(?:\s*usd)?"
                       r"(?:\s*(?:/\s*|per\s+)?(?:hrs?|hours?|hourly|h))?\s*", re.I)
_MONTHS = re.compile(r"\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|"
                     r"sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b\.?", re.I)
_DATE_CHARS = re.compile(r"[\d/.\-:,\s]*\d[\d/.\-:,\s]*")
_OPEN_END = re.compile(r"\b(?:open|current|present|none|n/a|ongoing|until further notice)\b", re.I)
# Words a Jurisdiction cell may hold besides its wage type and county names (agent/wages.py jurisdiction_kind takes
# the one of federal, state or union in it). A cell without exactly one of the three is never used by the agent, so
# it can become a token at no cost.
KIND_WORDS = {"federal", "state", "union"}
JURISDICTION_WORDS = KIND_WORDS | {"pw", "prevailing", "wage", "wages", "rate", "rates", "cba", "davis", "bacon",
                                   "dbra", "l", "i", "lni", "dol", "wa", "washington", "local", "journey", "level",
                                   "and", "or", "only", "county", "counties", "mod", "statewide", "all",
                                   # union names ("Union - IBEW Local 46")
                                   "ibew", "ua", "iuoe", "liuna", "carpenters", "laborers", "teamsters", "operating",
                                   "engineers", "ironworkers", "plumbers", "pipefitters", "electricians", "painters",
                                   "cement", "masons", "roofers", "sheet", "metal", "workers", "district", "council",
                                   "nw", "northwest", "pacific"}
_COUNTY_WORDS = {w.lower() for c in WA_COUNTIES for w in c.split()}
_JUR_CHARS = re.compile(r"[A-Za-z0-9&/,.()\-\u2013\u2014\s]*")
# A decision: a federal decision number (WA20250091), any code that starts with a digit, or these words
DECISION_WORDS = {"union", "cba", "mod", "mod.", "modification", "rev", "rev.", "revision", "state", "federal", "l&i",
                  "lni", "local", "no", "no.", "n/a", "na", "none", "journey", "level", "wa", "dol", "-", "/", "\u2013",
                  "pw", "prevailing", "wage", "wages", "washington", "rates"}
_DECISION_PART = re.compile(r"[A-Z]{2}\d[\w./#-]*|#?\d[\w./#-]*", re.I)
# One county: a WA county (any case, "County" optional) or one or two words and "County"; a state code may follow
_ONE_COUNTY = (r"(?:(?:" + "|".join(WA_COUNTIES) + r")\b(?:\s+county)?|[a-z][a-z.']*(?:\s[a-z][a-z.']*)?\s+county)"
               r"(?:\s*,?\s*[a-z]{2})?")
_COUNTY_RE = re.compile(r"\s*(?:" + _ONE_COUNTY + r"(?:\s*(?:,|/|;|&|and)\s*" + _ONE_COUNTY + r")*"
                        r"|statewide(?:\s*\(all counties\))?|all counties|(?:eastern|western|central)\s+(?:wa|washington))"
                        r"\s*", re.I)


def is_amount(v) -> bool:
    """A number, or text that is only an amount (an "/hr" after it allowed)."""
    if isinstance(v, bool):
        return False
    return isinstance(v, (int, float)) or (isinstance(v, str) and bool(AMOUNT_RE.fullmatch(v)))


def is_date_text(v: str) -> bool:
    """A date, or two (an LCP export's "issue\\nexpire" cell); an open end written as a word ("Open")."""
    return bool(_DATE_CHARS.fullmatch(_OPEN_END.sub(" ", _MONTHS.sub(" ", v))))


def is_jurisdiction_text(v: str) -> bool:
    """Exactly one of federal, state or union, and otherwise only county names, decision codes and the words above
    ("State - Pierce", "Federal - WA20250091 Mod 3", "CLARK - UNION RATES - WA")."""
    if not _JUR_CHARS.fullmatch(v):
        return False
    words = re.findall(r"[a-z]+", v.lower())
    if len(set(words) & KIND_WORDS) != 1:
        return False
    rest = [p for p in re.split(r"[\s,/()\-\u2013\u2014]+", v) if p]
    return all(p.lower() in JURISDICTION_WORDS or p.lower() in _COUNTY_WORDS or p == "&"
               or _DECISION_PART.fullmatch(p) for p in rest)


def is_decision_text(v: str) -> bool:
    parts = v.split()
    return bool(parts) and all(p.lower() in DECISION_WORDS or _DECISION_PART.fullmatch(p) for p in parts)


def value_ok(role: str, v) -> bool:
    """Is this value of a wage table's column (by the reader's role) one the strict mode keeps readable? Non-text
    values always (numbers and dates cross as they are in doubt mode: #112)."""
    if not isinstance(v, str):
        return True
    col = next((c for c in _ALL_COLUMNS if c.role == role), None)
    if role == "jur":
        return is_jurisdiction_text(v)
    if role == "wd":
        return is_decision_text(v)
    if role == "county":
        return bool(_COUNTY_RE.fullmatch(v))
    if col is not None and col.type in ("money", "number"):
        return bool(AMOUNT_RE.fullmatch(v))
    if col is not None and col.type == "date":
        return is_date_text(v)
    return True                                   # craft, classification: the text the agent needs


_LABEL_WORDS = {"job", "classification", "classifications", "description", "group", "level", "title", "code", "name",
                "craft", "trade", "&", "/", "-", "and"}


def _short_label(h: str) -> bool:
    """A header that is only a classification label ("Job Classification", "Classification Description")."""
    return all(w in _LABEL_WORDS for w in re.split(r"\s+", h.strip()))


def wage_columns(header: dict) -> tuple:
    """(header cells kept, rate columns, money columns, {column: role}) of a wage table's header row ({col: value}).
    Header cells kept: the agent's columns, and a short label with "classification" in it (parse_lcp_export looks
    for one to find the header)."""
    head, rates, money, roles = set(), set(), set(), {}
    for c, v in header.items():
        if not isinstance(v, str):
            continue
        h = norm_header(v)
        cols = [k for k in _ALL_COLUMNS if re.fullmatch(k.header_pattern, h)]
        if cols or ("classification" in h and _short_label(h)):
            head.add(c)
        if cols:
            roles[c] = cols[0].role
        if any(k.role in RATE_ROLES for k in cols):
            rates.add(c)
        if any(k.type == "money" for k in cols):
            money.add(c)
    return head, rates, money, roles


_ALL_COLUMNS = _columns("lcp", LCP_HEADERS) + _columns("table", TABLE_HEADERS)

WAGE_TABLE = FileKind(
    kind="wage_table",
    recognise=recognise,
    columns=_ALL_COLUMNS,
    above_header=("lcp: column A of a 'Location: <county>' line in rows 1-40 (LOCATION_RE); the county",),
    version=3,
    notes="v2 adds rate rows, value checks per column and the strict header cells (#118), used only with the "
          "tokenizer's vnext.strict_wage_tables. The tokenizer keeps every column whose header matches any layout's pattern "
          "(WAGE_COL), not only the recognised layout's: unchanged from #101.",
)


# WA L&I's "Journey level wages" CSV (agent/wages.py parse_lni_csv): its first row has these headers (lower case,
# stripped). The tokenizer's vnext.hold_unknown_csv (#111) lets a CSV through only when it is one of these or a payroll
# export.
LNI_CSV_NEEDS = ("county", "trade", "job classification", "wage")
# ... and every column parse_lni_csv reads (#53: with vnext.fail_closed_tables the tokenizer keeps only these readable)
LNI_CSV_READ = LNI_CSV_NEEDS + ("holiday", "overtime", "notes")


def is_lni_csv_header(row) -> bool:
    """The first row of an L&I journey-level wages CSV, as agent/wages.py parse_lni_csv reads it."""
    hdr = {str(h).strip().lower() for h in row}
    return all(n in hdr for n in LNI_CSV_NEEDS)
