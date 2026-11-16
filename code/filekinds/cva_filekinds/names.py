"""Names the AWS side may see as written (#53, part 2 and 3): file and folder names, sheet names, document properties.

"Only known values pass" (David, Nov 15, 2026; "if in doubt, assume it's sensitive", Nov 4): a name is cut into words,
and only these stay readable:
  - a word of KNOWN_WORDS: the words the agent itself writes or reads in names (its folders, sheets, run files, the
    wage-file hints of agent/uploads.py WAGE_NAME_HINT, rate table names) and a few generic ones (months, "copy",
    "final", "week"...). No word that is also a common given name or surname (review of the #53 PR: no trade words such
    as Mason or Craft, no Jan, April, May, June, July, August, Sun, Day, Page, Link, New, Davis or Bacon: "Davis-Bacon"
    is known only as the pair); tests/test_fail_closed_53.py checks a list of common names;
  - "L&I" and "Davis-Bacon" (or "Davis Bacon") as one unit;
  - a word made only of known words glued together, cut at its capitals ("ProjectInfo", "WageRates") or, in lower
    case, into known words of 4+ letters ("wagerates");
  - a number of up to 4 digits (a longer one can be an employee ID or an SSN), and the parts of a date written with
    separators ("10-31-26", "2026-11-14-03-02-11": each part is such a number);
  - a federal wage determination ID ("WA20250093": agent/wages.py reads it from a .txt file's name), an ordinal
    ("1st"), a short version or week mark ("v2", "wk36");
  - a token already in the name ("EMP-3f2a91c0d1", "FIL-…"), and the separators between words.
Every run of other words (with the separators inside it) is replaced by ONE reversible token (``opaque``, with the
tokenizer's ``make_token``): "Ana Rivera timesheet 10-31.xlsx" -> "FIL-3f2a91c0d1 timesheet 10-31.xlsx". The agent then
still finds what it reads in a name (a date, "union", "REVIEWED", "info" in a sheet's name), and what it writes around
a name it was given ("2026-11-14 FIL-3f2a91c0d1 timesheet 10-31 - REVIEWED.xlsx") comes back to staff with the real
words in place of the token (the token is an ordinary vault token: every restore puts it back).

Pure: no key, no vault. The tokenizer (pii_tokenizer/names.py) gives ``make_token``; the agent's tests check that every
name the agent makes itself passes unchanged (tests/test_names_vocabulary.py), so the AWS side never renames them.
"""
from __future__ import annotations

import re

# A token (pii_tokenizer/tokens.py TOKEN_RE: 2-5 capitals, a hyphen, 10 hex digits)
TOKEN_PART = r"[A-Z]{2,5}-[0-9a-f]{10}(?![0-9a-f])"

# Lower case. Grouped by where they come from; keep each group sorted.
KNOWN_WORDS = frozenset("""
    uploads upload payroll payrolls tracker trackers files file archived archive unused logs log wage wages rates rate
    reviewed review data tasks task intent intents affidavit affidavits replaced removed uploaded migrated being moved
    reference knowledge union federal state cba determination determinations prevailing journey journeyman level dol
    lni wd dbra agent compliance validation cva held back aws copy sync backup deleted purge sources
    comments comment audit decisions decision run runs events email preview staff admin pipeline settings summary
    readme notes note info information metadata project projects sheet sheets table tables previous
    template inferred scan tmp lcptracker lcp certified fringe fringes benefit benefits hours hour classification
    classifications contractor contractors subcontractor subcontractors prime sub general schedule
    schedules export exports report reports list lists source sources current progress links model check checks
    before after header title row rows total totals
    pst pdt utc gmt
    feb mar apr jun jul aug sep sept oct nov dec february september october november december
    mon tue wed thu fri sat monday tuesday wednesday thursday friday saturday sunday
    week weekly ending end start month months monthly year years yearly daily quarter period date dated
    final draft old original updated update revised revision rev version ver copy v wk fy q
    the and of for to from by in on at with no not or as is all per via
    xlsx xlsm xls csv tsv pdf txt md json jsonl eml html htm log zip docx doc png jpg jpeg
    """.split())

_UNIT_RE = re.compile(r"(?P<tok>" + TOKEN_PART + r")|(?P<li>(?i:l&i|davis[- ]bacon)(?![^\W_]))|(?P<word>[^\W_]+)"
                      r"|(?P<sep>[\W_]+)")
_SHORT_NUMBER = re.compile(r"\d{1,4}")
_WD_ID = re.compile(r"[A-Za-z]{2}\d{8}")               # agent/wages.py: [A-Z]{2}\d{8} in a .txt file's name
_MARK = re.compile(r"(?:\d{1,2}(?:st|nd|rd|th)|(?:v|wk|q|rev|ver|fy|p)\d{1,3})", re.I)
_CAMEL = re.compile(r"[A-Z]?[a-z]+|[A-Z]+(?![a-z])|\d+")


def _glued(w: str) -> bool:
    """A word made only of known words: "ProjectInfo" (by its capitals), "wagerates" (by trying every split)."""
    parts = _CAMEL.findall(w)
    if len(parts) > 1 and "".join(parts) == w and all(_one(p) for p in parts):
        return True
    low = w.lower()
    if len(low) > 40:
        return False
    ok = [True] + [False] * len(low)                    # ok[i]: low[:i] is made of known words of 2+ letters
    for i in range(1, len(low) + 1):
        ok[i] = any(ok[j] and len(low[j:i]) >= 4 and low[j:i] in KNOWN_WORDS for j in range(max(0, i - 20), i))
    return ok[-1]


def _one(w: str) -> bool:
    return w.lower() in KNOWN_WORDS or bool(_SHORT_NUMBER.fullmatch(w) or _WD_ID.fullmatch(w) or _MARK.fullmatch(w))


def is_known_word(w: str) -> bool:
    return _one(w) or _glued(w)


def units(text: str) -> list:
    """The name cut into (kind, text): "tok" (a token), "known" (a known word, "L&I" included), "word" (any other
    word), "sep" (what separates words)."""
    out = []
    for m in _UNIT_RE.finditer(text):
        k = m.lastgroup
        if k == "li":
            out.append(("known", m.group(0)))
        elif k == "word":
            w = m.group(0)
            if is_known_word(w):
                out.append(("known", w))
                continue
            # an unknown word with a known word glued in ("ProjInfo", "OrtegaPayroll", "jobinfo"): cut there, so the
            # known part ("Info": Project Info is found by it) stays readable (review of the #53 PR)
            parts = _CAMEL.findall(w) if "".join(_CAMEL.findall(w)) == w else [w]
            if not any(_one(p) for p in parts):
                parts = [p for p in re.split(r"(?i)(info)", w) if p]
            out += [("known" if _one(p) else "word", p) for p in parts]
        else:
            out.append((k, m.group(0)))
    return out


def unknown_runs(text: str) -> list:
    """(start, end) of each run of unknown words in ``text`` (separators between two of them included)."""
    runs, pos, cur = [], 0, None
    for kind, s in units(text):
        if kind == "word":
            cur = (cur[0] if cur else pos, pos + len(s))
        elif kind != "sep" and cur:
            runs.append(cur)
            cur = None
        pos += len(s)
    if cur:
        runs.append(cur)
    return runs


def is_plain(text: str) -> bool:
    """True when every word of ``text`` may be seen as written (``opaque`` leaves it unchanged)."""
    return not unknown_runs(text)


def opaque(text: str, make_token) -> str:
    """``text`` with each run of unknown words replaced by ``make_token(run, glue)`` (a token: 2-5 capitals, a hyphen,
    hex). A token glued to a letter or digit on either side gets a space there, so it stays readable as a token;
    ``glue`` says where ("b" before, "a" after, "ba" both, "" none), so the way back takes the space out again (the
    tokenizer gives such a token its own kind: pii_tokenizer/tokens.py unglue)."""
    out, pos = [], 0
    for start, end in unknown_runs(text):
        before = " " if start and text[start - 1].isalnum() else ""
        after = " " if end < len(text) and text[end].isalnum() else ""
        tok = make_token(text[start:end], ("b" if before else "") + ("a" if after else ""))
        out += [text[pos:start], before, tok, after]
        pos = end
    return "".join(out) + text[pos:]


_TOKEN_IN_NAME = re.compile(TOKEN_PART)


def skeleton(name: str) -> str:
    """What of a name survives the opaque rule, with every token as "…": the same for a name written before the
    rename ("Zorvex_PayrollDetailReport Week 37.xlsx") and after it ("FIL-…_Payroll FIL-… Report Week 37.xlsx"), so
    the agent can find a file it saved by its old name (agent/payroll.py last_reviewed_input, #195)."""
    return _TOKEN_IN_NAME.sub("…", opaque(name, lambda run, glue: "…"))
