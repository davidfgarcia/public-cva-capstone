"""Task 2: download WA L&I intent/affidavit records for each applicable project (data.wa.gov,
Socrata API, no login) and validate the project's compliance tracker against them.
Invalid/missing tracker values are marked red, commented, and listed on a "Review Log" sheet (the latest run only).
Only what staff can act on is marked (David, Nov 7, 2026): an intent before work starts, or an affidavit L&I doesn't
have yet, is counted in the Audit sheet, not marked."""
import csv
import json
import os
import contextvars
import re
import shutil
import datetime as dt
from pathlib import Path

import openpyxl
import requests

from .util import (err_text, is_blank, log, norm_header, presave, run_timestamp, similarity, to_date, to_float,
                   truthy, visible_files, file_kinds)

# The tracker's columns (HEADER_ROLES), the 11-row header window, the header test and the agent's note authors come
# from the file-kind registry (code/filekinds, copied here as agent/cva_filekinds); the tokenizer imports the same
# definitions (pii_tokenizer/cva_filekinds), so what it keeps readable is what is read here (vNext PR 3). Not reading
# "Affidavit Good" or "CP Status" also means the tokenizer sends AWS those columns as text tokens.
from .cva_filekinds.project_info import NO_FILING_SEARCH_RE, PRIME_INTENT_LABEL, PROJECT_ID_LABEL_RE  # noqa: E402
from .cva_filekinds.tracker import AUTHOR, OLD_AUTHORS, is_tracker_header  # noqa: E402,F401
from .cva_filekinds.tracker import HEADER_ROWS as TRACKER_HEADER_ROWS  # noqa: E402
from .cva_filekinds.tracker import TRACKER_ROLES as HEADER_ROLES  # noqa: E402
FIELD_HINTS = {   # first entries are the data.wa.gov field names as of Oct 2026
    "intent_id": [r"^intent_id(_number)?$", r"^intent$", r"intent.*id"],
    "affidavit_id": [r"^affidavit_id(_number)?$", r"^id$", r"affidavit.*id"],
    "project_id": [r"^project_id$", r"project.*id"],
    "contractor": [r"^companyname$", r"^name$", r"^contractor_name$", r"^(business|company|employer)_name$",
                   r"contractor.*name"],
    "status": [r"^(intent_|affidavit_)?status$", r"status"],
    # who hired the contractor, as filed on the intent (#121: compared with the tracker's Hiring Contractor)
    "hiring": [r"^hiringcompanyname$", r"^hiring_?company_?name$", r"hiring.*name"],
    "ubi": [r"^ubi$", r"ubi"],              # (not "hiringcomapanyubi" when "ubi" exists; review of #121)
    # when L&I received the form: picks the newest of several affidavits filed under one intent
    "received": [r"^formreceiveddate$", r"^application_received_date$", r"received.*date|date.*received"],
}


# ------------------------------------------------------------------ client
class LniClient:
    def __init__(self, cfg):
        c = cfg["intent"]
        self.base = c["socrata_base"].rstrip("/")
        self.ds = c["datasets"]
        self.token = c.get("app_token") or os.environ.get("SOCRATA_APP_TOKEN") or ""
        self.overrides = c.get("field_overrides") or {}
        self.fields = {}

    def _get(self, dataset, params):
        h = {"X-App-Token": self.token} if self.token else {}
        from . import net                     # retries dropped connections, timeouts, 429 and 5xx; not 4xx
        r = net.get(f"{self.base}/{dataset}.json", params=params, headers=h, timeout=60, what="data.wa.gov query")
        return r.json()

    def discover(self, which):
        """Map our roles (intent_id, project_id, contractor, ...) to the dataset's real field names."""
        if which in self.fields:
            return self.fields[which]
        sample = self._get(self.ds[which], {"$limit": 50})   # several rows: Socrata leaves out a row's null fields
        keys = list(dict.fromkeys(k for r in sample for k in r))
        m = {}
        for role, pats in FIELD_HINTS.items():
            ov = (self.overrides.get(which) or {}).get(role)
            if ov:
                m[role] = ov
                continue
            for p in pats:
                hit = next((k for k in keys if re.search(p, k) and not k.startswith(":")), None)
                if hit and hit not in m.values():
                    m[role] = hit
                    break
        log.debug("L&I %s fields detected", which, extra={"details": {**m, "available": ", ".join(keys)}})
        missing = [r for r in ("intent_id", "project_id", "contractor") if r not in m]
        if which == "affidavits":
            missing = [r for r in ("affidavit_id", "intent_id") if r not in m]
        if missing and not keys:
            log.debug("No %s records to detect field names from (none filed yet)", which)
        elif missing:
            log.warning("Some L&I fields were not recognised; set intent.field_overrides in config.yaml",
                        extra={"details": {"dataset": which, "missing": ", ".join(missing)}})
        self.fields[which] = m
        return m

    def _where_in(self, which, field, values):
        vals = [str(v) for v in values if v is not None]
        if not vals:
            return []
        out = []
        for i in range(0, len(vals), 50):
            chunk = vals[i:i + 50]
            err = None
            for quote in ("'", ""):
                try:
                    lst = ",".join(f"{quote}{v}{quote}" for v in chunk)
                    out += self._get(self.ds[which], {"$where": f"{field} in ({lst})", "$limit": 50000})
                    err = None
                    break
                except requests.HTTPError as e:
                    err = e
            if err is not None:
                # don't carry on without these records: every ID in the chunk would be marked "not found".
                # Raising lets download() say the check wasn't run (data.wa.gov stopped answering).
                raise err
        return out

    def fetch(self, which, role, values):
        f = self.discover(which).get(role)
        values = list(values)
        out = self._where_in(which, f, values) if f else []
        log.debug("L&I query", extra={"details": {"dataset": which, "field": f, "values": len(values),
                                                   "records": len(out), "source": type(self).__name__}})
        return out


    def last_updated(self, which):
        """When L&I last published records to this dataset: data.wa.gov's own "rows updated" time (UTC). Raises
        requests.RequestException when data.wa.gov can't be reached or refuses (net.get retries what is temporary),
        and LniUnreadable when it answers with something that doesn't say when (not JSON, no or a bad date)."""
        from urllib.parse import urlsplit
        from . import net
        u = urlsplit(self.base)
        h = {"X-App-Token": self.token} if self.token else {}
        r = net.get(f"{u.scheme}://{u.netloc}/api/views/{self.ds[which]}.json", headers=h, timeout=30,
                    what="data.wa.gov dataset date")
        try:
            t = int(r.json().get("rowsUpdatedAt"))
            if t <= 0:
                raise ValueError("no date")
            return dt.datetime.fromtimestamp(t, dt.timezone.utc)
        except Exception as e:                 # not JSON, a list, no date, not a number, out of range
            raise LniUnreadable(f"{which} ({self.ds[which]}): no publish date in the answer ({type(e).__name__})")


class LniUnreadable(Exception):
    """data.wa.gov answered, but not with what the check needs (a date, records)."""


# Whether L&I's records can be used this run, worked out once per run (all projects share the two datasets; a check
# per project couldn't tell a quiet project from a feed that stopped). L&I stopped publishing on Oct 28, 2026 while
# saying "daily": every intent filed since looked "not found". So the intent/affidavit check is not run in any project
# when (David, Nov 8, 2026; three situations, each said differently in both emails):
#   LNI_DOWN       data.wa.gov can't be reached, even after net.get's retries;
#   LNI_STALE      it is up, but L&I hasn't published for more than intent.lni_max_age_days (calendar days, Pacific
#                  time, as the emails show dates);
#   LNI_UNREADABLE it is up but refuses (4xx: a dataset ID L&I replaced, a bad app token) or answers without a date:
#                  a setup problem for the administrator, not an outage ("if in doubt": no date, no check).
# No saved copy is used instead: it would be at least as old. "NOT CHECKED", not "SKIPPED": none of these is a
# failure of the run (SKIPPED would make every run a PARTIAL FAILURE for as long as L&I doesn't publish); staff see
# Not checked, the administrator a run-wide warning.
LNI_STALE = ("intent: NOT CHECKED - data.wa.gov is up, but L&I's records there were last updated {when}, {days} days ago "
             "(the limit is {limit}), so intents and affidavits filed since then can't be seen; the tracker was not "
             "checked or changed")
LNI_DOWN = ("intent: NOT CHECKED - data.wa.gov (L&I's public records) could not be reached, even after retries; the "
            "tracker was not checked or changed")
LNI_UNREADABLE = ("intent: NOT CHECKED - data.wa.gov answered, but L&I's records couldn't be read ({why}); the tracker "
                  "was not checked or changed")
# With intent.lni_source "live" the same lines name L&I's live lookup (for_source), and one more applies to a single
# project: the lookup answered for others but not usably for this one (an ID it refuses, the request cap or time
# budget, none of the tracker's intents found). Only that project isn't checked (review of #150, F2, F8).
LNI_PROJECT_UNREADABLE = ("intent: NOT CHECKED - L&I's live lookup answered, but this project's records couldn't be "
                          "read ({why}); the tracker was not checked or changed")


def for_source(line, source):
    """The NOT CHECKED line naming the source used (data.wa.gov unless intent.lni_source is "live")."""
    if not line or source == "data.wa.gov":
        return line
    return line.replace("data.wa.gov (L&I's public records)", source).replace("data.wa.gov", source)


LNI_TZ = "America/Los_Angeles"           # the emails' time zone (notifications.config `timezone`)
_LNI_USABLE = {}                         # run (util.run_number) -> None (usable) or the NOT CHECKED line
RUN_WIDE = {"run_wide": True}            # listed once in the administrator's email, for the whole run (notify.py)


def _run_wide():
    """Log outside any project: the administrator's email lists it once, for the whole run (tracelog.context)."""
    from .tracelog import context
    return context("", "")


def lni_max_age_days(cfg):
    """intent.lni_max_age_days as a whole number of 0 or more (0: never too old); anything else: 3, with a warning."""
    v = cfg["intent"].get("lni_max_age_days", 3)
    try:
        if isinstance(v, bool) or int(v) != float(v) or int(v) < 0:
            raise ValueError
        return int(v)
    except (TypeError, ValueError, OverflowError):
        with _run_wide():
            log.warning("intent.lni_max_age_days in config.yaml must be a whole number of days (0 or more); 3 is used",
                        extra={"run_wide": True, "details": {"value": str(v)[:40]}})
        return 3


def lni_unusable(client, cfg):
    """None when L&I's records are current enough to check against, else the NOT CHECKED line for this run."""
    from .util import run_number
    run = run_number()
    if run in _LNI_USABLE:
        return _LNI_USABLE[run]
    days = lni_max_age_days(cfg)
    out = None
    with _run_wide():
        try:
            times = {w: client.last_updated(w) for w in ("intents", "affidavits")}
        except requests.HTTPError as e:
            code = getattr(e.response, "status_code", None)
            if code is not None and 400 <= code < 500 and code != 429:      # 429: busy, as net.py treats it
                log.warning("data.wa.gov refused the L&I dataset request (HTTP %s); check the dataset IDs and the app "
                            "token in config.yaml. The intent/affidavit check is not run in any project", code,
                            extra={"run_wide": True, "details": {"datasets": ", ".join(client.ds.values())}})
                out = LNI_UNREADABLE.format(why=f"HTTP {code}")
            else:
                log.warning("data.wa.gov could not be reached (HTTP %s, after retries); the intent/affidavit check is "
                            "not run in any project", code, extra=RUN_WIDE)
                out = LNI_DOWN
        except requests.RequestException as e:
            log.warning("data.wa.gov could not be reached (%s, after retries); the intent/affidavit check is not run "
                        "in any project", type(e).__name__, extra=RUN_WIDE)
            out = LNI_DOWN
        except LniUnreadable as e:
            log.warning("data.wa.gov answered without L&I's publish date: %s. The intent/affidavit check is not run in "
                        "any project", e, extra=RUN_WIDE)
            out = LNI_UNREADABLE.format(why="no publish date")
        else:
            from zoneinfo import ZoneInfo
            tz = ZoneInfo(LNI_TZ)
            oldest = min(times.values())
            age = (dt.datetime.now(tz).date() - oldest.astimezone(tz).date()).days      # calendar days, Pacific
            when = oldest.astimezone(tz)
            details = {**{f"{w} ({client.ds[w]})": t.strftime("%Y-%m-%d %H:%M UTC") for w, t in times.items()},
                       "max_age_days": days}
            if days and age > days:
                log.warning("L&I's records on data.wa.gov were last published %s, %d days ago (limit %d); the "
                            "intent/affidavit check is not run in any project", f"{when:%Y-%m-%d %H:%M %Z}", age, days,
                            extra={"run_wide": True, "details": details})
                out = LNI_STALE.format(when=f"{when:%b} {when.day}, {when.year}", days=age, limit=days)
            else:
                log.info("L&I records last published", extra={"details": details})
    _LNI_USABLE.clear()                  # one run at a time (the app runs several in one process)
    _LNI_USABLE[run] = out
    return out


def lni_stopped_answering(cfg, line=None, source="data.wa.gov"):
    """data.wa.gov stopped answering (or answering usably) during this run's downloads: the rest of the run is not
    checked, and the administrator is told that projects checked before stand (review #4). Returns the line."""
    from .util import run_number
    line = line or LNI_DOWN
    if _LNI_USABLE.get(run_number()) is None:
        with _run_wide():
            log.warning("%s stopped answering during the run: any project checked before then was checked "
                        "against L&I's records as usual; the intent/affidavit check of the rest was not run", source,
                        extra=RUN_WIDE)
    _LNI_USABLE.clear()
    _LNI_USABLE[run_number()] = line
    return line


def _num_id(v):
    f = to_float(v)
    return str(int(f)) if f is not None and f > 0 else None


# intent.strict_ids (#175; on in config.yaml since Nov 15, 2026, off unless it says true): a tracker line's
# Intent/Affidavit ID cell must start with a whole number that ends at the end of the cell or at a space or a separator
# ("-", "–", "/", "(", "," or "." not followed by a digit, "," between two numbers of 5 digits or more, ...): "8800050-Filed", "8800040 (amended)", "8800040, 8800041"
# and "8800040 / 8800041" are read by that number, as staff
# write them (the trial run of Nov 14, 2026: "-Filed" on correct IDs); a letter or another character inside the number
# ("88000O8", "8800040A", "8,800,040", "8800040.5") is "not a valid number" instead of being read by its first digits
# (David, Nov 14, 2026, after the trial). Set for the length of a check (_run_intent_validation, validate), per thread
# (the web app serves requests in threads); the Project Info sheet's numbers and L&I's own records are read with
# _num_id, as before.
_strict_ids = contextvars.ContextVar("cva_strict_ids", default=False)
_WHOLE_ID = re.compile(r"#?\s*(\d+)(?:\.0+)?(?=$|[\s\-\u2013\u2014/()\[\];:&+_|*]|[.,](?!\d)|(?<=\d{5}),(?=\s*\d{5}))")


def _cell_id(v):
    """A tracker line's Intent/Affidavit ID cell, as text, or None (Project Info and L&I records use _num_id). With
    intent.strict_ids the number at the start of the cell must end at the cell's end or at a space or separator
    ("8800040", 8800040, 8800040.0, "#8800040", "8800050-Filed", "8800040 (amended)"); "88000O8" or "8800040A" is
    not a number (the check then says "not a valid number"). Without it: its leading digits, as before."""
    if not _strict_ids.get():
        return _num_id(v)
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() and v > 0 else None
    m = _WHOLE_ID.match(str(v).strip())
    return str(int(m.group(1))) if m and int(m.group(1)) > 0 else None


class _StrictIds:
    """``with _StrictIds(cfg):`` turns intent.strict_ids on for the block, as config.yaml says."""

    def __init__(self, cfg):
        self.on = bool(((cfg or {}).get("intent") or {}).get("strict_ids"))

    def __enter__(self):
        self.token = _strict_ids.set(self.on or _strict_ids.get())

    def __exit__(self, *exc):
        _strict_ids.reset(self.token)


# ------------------------------------------------- the agent's marks, kept in the tracker
MARKS_TOPIC = "Marks (agent use)"


def marks_notes(sheet, marks, limit=30000):
    """The cells marked in this run, as Audit-sheet notes, so whichever side runs next (desktop app or AWS) can
    undo them: 'sheet=<name>' then 'A5|s12|s34|n' per cell (original style, marked style, note added)."""
    items = [f"{ref}|s{o if o is not None else ''}|s{n if n is not None else ''}|{'n' if noted else '-'}"
             for ref, (o, n, noted) in sorted(marks.items())]
    out, cur = [], f"sheet={sheet};"
    for it in items:
        if len(cur) + len(it) + 1 > limit:
            out.append(cur)
            cur = f"sheet={sheet};"
        cur += " " + it
    out.append(cur)
    return out


def marks_from_tracker(wb, sheet_name):
    """The last run's marks recorded in the tracker's Audit sheet: {"sheet": ..., "cells": {ref: [o, n, noted]},
    "run": that run's timestamp}."""
    if sheet_name not in wb.sheetnames:
        return None
    rows = [r for r in wb[sheet_name].iter_rows(min_row=2, values_only=True) if len(r) >= 4 and r[2] == MARKS_TOPIC]
    if not rows:
        return None
    last = max(str(r[0]) for r in rows)
    out = {"sheet": None, "cells": {}, "run": last}
    for r in rows:
        if str(r[0]) != last:
            continue
        head, _, body = str(r[3] or "").partition(";")
        out["sheet"] = head[len("sheet="):]
        for it in body.split():
            ref, o, n, noted = (it.split("|") + ["", "", "", ""])[:4]
            conv = lambda x: (int(x[1:]) if x[1:].isdigit() else (x[1:] or None))
            out["cells"][ref] = [conv(o), conv(n), noted == "n"]
    return out


# ----------------------------------------------------------- the tracker
# Change 1 (#93): the Google-side tokenizer copies each project's working tracker from the staff drive into Tracker/
# and leaves, at the top of the tokenized folder, where each one came from: a status code per project folder and the
# working file's last change, no file names (code/tokenizer/pii_tokenizer/staff_trackers.py; keep the codes in step).
TRACKER_SOURCES = "_tracker_sources.json"
SOURCE_NOT_CHECKED = {
    "several": "the staff folder has more than one file that looks like the tracker",
    "not_found": "the tracker named in tracker_source.txt wasn't found",
    "no_folder": "no project folder in the staff drive matches this project",
    "unsupported": "the tracker in the staff folder is in a format the CVA can't read",
    "ambiguous": "several project folders in the staff drive could be this project's",
}
MATCH_WORDS = {"tracker_name": "the tracker's name", "project_number": "the project number",
               "folder_name": "the folder name"}
NEW_MATCH_DAYS = 7


def new_match_line(source, now=None):
    """For the administrator, for a week after the staff folder was matched (or matched again): check it once."""
    by, at = (source or {}).get("matched_by"), (source or {}).get("matched_at")
    if by not in MATCH_WORDS or not at:
        return None
    try:
        when = dt.datetime.fromisoformat(str(at).replace("Z", "+00:00"))
    except ValueError:
        return None
    now = now or dt.datetime.now(dt.timezone.utc)
    if now - when > dt.timedelta(days=NEW_MATCH_DAYS):
        return None
    return (f"WARNING intent: the tracker's folder in the staff drive was matched by {MATCH_WORDS[by]} on "
            f"{when:%Y-%m-%d}; check it once in Tracker sources.csv")
STAFF_DRIVE_UNREADABLE = ("WARNING intent: the staff drive couldn't be read this time, so the copy already in the "
                          "CVA's Tracker folder was checked (it may be out of date)")


def tracker_source(root, project: str):
    """{"status": ..., "modified": ...} for this project, or None (no summary: the desktop app, or Change 1 off)."""
    try:
        data = json.loads((Path(root) / TRACKER_SOURCES).read_text())
        e = data.get("projects", {}).get(project)
    except (OSError, ValueError, AttributeError):
        return None
    return e if isinstance(e, dict) and isinstance(e.get("status"), str) else None


def tracker_source_note(name: str, source, tz="UTC") -> str:
    """The Audit sheet's line on which tracker was checked (#93)."""
    status = (source or {}).get("status")
    if status in ("copied", "unchanged", "deferred"):
        when = _local_time(source.get("modified"), tz)
        return f"{name}, copied from the staff folder" + (f" (last changed there {when})" if when else "")
    if status == "kept":
        return f"{name}, the copy in the CVA's Tracker folder (tracker_source.txt says none)"
    if status == "error":
        return f"{name}, the copy in the CVA's Tracker folder (the staff drive couldn't be read this time)"
    return f"{name}, the copy in the CVA's Tracker folder"


def _local_time(iso, tz="UTC") -> str:
    try:
        from zoneinfo import ZoneInfo
        t = dt.datetime.fromisoformat(str(iso).replace("Z", "+00:00")).astimezone(ZoneInfo(tz))
        return f"{t:%b} {t.day}, {t.year} {t.hour % 12 or 12}:{t:%M} {'AM' if t.hour < 12 else 'PM'}"
    except Exception:  # noqa: BLE001 - a missing or odd time: the line goes without it
        return ""


def find_tracker(project_dir: Path):
    from .layout import Paths
    files = [f for f in visible_files(Paths(project_dir).tracker, {".xlsx", ".xlsm"})
             if "(AWS copy " not in f.name]               # a tokenizer conflict copy is never the tracker
    if len(files) > 1:
        log.warning("Tracker/ holds more than one tracker; the most recently changed is used. Keep only the current "
                    "one there (older versions belong in Tracker/Archived)",
                    extra={"details": {"files": file_kinds(files)}})
    return max(files, key=lambda p: (p.stat().st_mtime, p.name)) if files else None


def locate_table(wb):
    """Find (sheet, header_row, {role: column_index}) of the contractor table."""
    for ws in wb.worksheets:
        for r in range(1, TRACKER_HEADER_ROWS + 1):
            hdr = [norm_header(c.value) for c in ws[r]]
            if is_tracker_header(hdr):
                cols = {}
                for role, pat in HEADER_ROLES.items():
                    for i, h in enumerate(hdr):
                        if re.fullmatch(pat, h) and i not in cols.values():
                            cols[role] = i
                            break
                return ws, r, cols
    return None, None, None


def _kind(issue):
    """An issue with its numbers, dates and lists taken out, for counting kinds of problems."""
    issue = re.sub(r"\s*\((received|the newest of)[^)]*\)", "", issue)
    return re.sub(r"\b\d{4,}\b", "#", issue)


def other_tables(wb, ws):
    """Titles of the other tabs that also look like a contractor table (they are not checked: named in Audit)."""
    out = []
    for s in wb.worksheets:
        if s is ws or ws is None:
            continue
        for r in range(1, TRACKER_HEADER_ROWS + 1):
            hdr = [norm_header(c.value) for c in s[r]]
            if is_tracker_header(hdr):
                out.append(s.title)
                break
    return out


def project_info(wb):
    """Text of the 'Project Info' sheet + prime intent / L&I project id if listed there."""
    info = {"text": "", "prime_intent": None, "project_id": None}
    ws = next((s for s in wb.worksheets if "info" in s.title.lower()), None)
    if not ws:
        return info
    lines = []
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row if c is not None]
        if not cells:
            continue
        lines.append(" ".join(str(c) for c in cells))
        label = str(cells[0]).lower()
        if PRIME_INTENT_LABEL in label and len(cells) > 1:
            info["prime_intent"] = _num_id(cells[1])      # free text ("8800040 - ABC"): as before
        if PROJECT_ID_LABEL_RE.search(label) and len(cells) > 1:
            info["project_id"] = _num_id(cells[1])
    info["text"] = "\n".join(lines)
    return info


def is_applicable(project_name, wb, cfg):
    pc = (cfg.get("projects") or {}).get(project_name, {}) or {}
    if "intent_applicable" in pc:
        return bool(pc["intent_applicable"]), "set in config.yaml"
    ws, _, cols = locate_table(wb)
    if ws is None or not ({"intent_id", "intent_flag", "affidavit_id", "affidavit_flag"} & set(cols)):
        return False, "tracker has no intent/affidavit columns"
    if NO_FILING_SEARCH_RE.search(project_info(wb)["text"]):
        return False, "Project Info says no L&I filing required"
    return True, "tracker has intent/affidavit columns"


# --------------------------------------------------------- mark handling
def validate(rows, cols, lni, cfg, counts=None):
    """The check's findings (_validate), with intent.strict_ids as config.yaml says."""
    with _StrictIds(cfg):
        return _validate(rows, cols, lni, cfg, counts)


def _validate(rows, cols, lni, cfg, counts=None):
    """rows: [(excel_row, values)]. lni: {'intents': {id: rec}, 'affidavits': {id: rec},
    'intents_by_name': [...], 'affidavits_by_intent': {intent_id: rec}, 'project_ids': set, 'ok': bool}.
    Returns [(row, role, value, lni_value, issue)]. counts (a dict, optional) receives the lines left unmarked
    because nothing is due yet: "not_started" (no intent anywhere, work not started) and "no_affidavit_yet" (work
    completed, no affidavit in the tracker or at L&I) and "filed_not_at_lni" (Affidavit Filed ticked, no Affidavit ID,
    none under the line's intent at L&I yet), for the Audit sheet."""
    counts = {} if counts is None else counts
    counts.setdefault("not_started", 0)
    counts.setdefault("no_affidavit_yet", 0)
    counts.setdefault("filed_not_at_lni", 0)
    c = cfg["intent"]
    skip = [s.lower() for s in c["skip_statuses"]]
    started = [s.lower() for s in c["work_started_statuses"]]
    done = [s.lower() for s in c["completed_statuses"]]
    thr = c.get("name_mismatch_threshold", 0.35)
    out = []
    for rn, v in rows:
        g = lambda role: v[cols[role]] if role in cols and cols[role] < len(v) else None
        name = g("contractor")
        if is_blank(name) or (isinstance(name, bool)):
            continue
        status = str(g("status") or "").strip().lower()
        if status in skip:
            continue
        iid, aid = _cell_id(g("intent_id")), _cell_id(g("affidavit_id"))
        iflag = truthy(g("intent_flag")) if "intent_flag" in cols else bool(iid)
        aflag = truthy(g("affidavit_flag")) if "affidavit_flag" in cols else bool(aid)
        raw_iid = g("intent_id")

        # --- intent
        if "intent_id" in cols:
            if not is_blank(raw_iid) and not isinstance(raw_iid, bool) and iid is None:
                out.append((rn, "intent_id", raw_iid, None, "Intent ID is not a valid number"))
            if iflag and iid is None and is_blank(raw_iid):
                out.append((rn, "intent_id", raw_iid, None, "Intent marked filed/approved but Intent ID is missing"))
        if iid and lni["ok"]:
            rec = lni["intents"].get(iid)
            if rec is None:
                out.append((rn, "intent_id", iid, None, "Intent ID not found in L&I records for this project"))
            else:
                if lni["project_ids"] and rec.get("_project_id") not in lni["project_ids"]:
                    out.append((rn, "intent_id", iid, rec.get("_project_id"), "Intent belongs to a different L&I project"))
                if rec.get("_contractor") and similarity(name, rec["_contractor"]) < thr:
                    out.append((rn, "contractor", name, rec["_contractor"], "Contractor name doesn't match the L&I intent"))
                # #121 (David, Nov 8-9, 2026): the line's Hiring Contractor against the hiring company on the line's
                # own intent, only when that intent is valid for this line: it is this project's, isn't withdrawn,
                # cancelled... at L&I, and names this line's contractor (a subcontractor can have several parent
                # contractors, one per intent: never compared with an intent found by name). Both sides must name one.
                hiring = g("hiring")
                valid = ((not lni["project_ids"] or rec.get("_project_id") in lni["project_ids"])
                         and not ignored_at_lni(rec.get("_status"), cfg)
                         and not (rec.get("_contractor") and similarity(name, rec["_contractor"]) < thr))
                if valid and not names_no_hiring(rec.get("_hiring")) and not names_no_hiring(hiring):
                    same, how = compare_companies(hiring, rec["_hiring"])
                    h = counts.setdefault("hiring", {"compared": 0, "same": 0, "matched": [], "why": {}})
                    h["compared"] += 1
                    if not same:
                        h["why"][rn] = how
                        out.append((rn, "hiring", hiring, rec["_hiring"], HIRING_MISMATCH))
                    elif how:                     # matched, not by the same name: listed in the Audit tab
                        h["matched"].append((rn, name, hiring, rec["_hiring"], how))
                    else:
                        h["same"] += 1
                st = str(rec.get("_status") or "").lower()
                if "intent_flag" in cols and not iflag and st and "approv" in st:
                    out.append((rn, "intent_flag", g("intent_flag"), rec.get("_status"), "L&I shows this intent approved; tracker says no"))
                if "intent_flag" in cols and iflag and st and not re.search(r"approv|filed|received|accept", st):
                    out.append((rn, "intent_flag", g("intent_flag"), rec.get("_status"), "Tracker says filed/approved but L&I status differs"))
        match = None                       # the L&I intent found by contractor name, for a line with no intent ID
        counted_not_started = False
        # A cell already marked "not a valid number" gets no second finding that calls it empty ("tracker has none",
        # "no intent on file", "add it to the tracker": the trial run, Nov 14, 2026): one finding per cell, which
        # names the number L&I has when there is one (_with_lni_number).
        bad_iid = "intent_id" in cols and not is_blank(raw_iid) and not isinstance(raw_iid, bool) and iid is None
        bad_aid = ("affidavit_id" in cols and not is_blank(g("affidavit_id")) and not isinstance(g("affidavit_id"), bool)
                   and aid is None and str(g("affidavit_id")).strip().lower() not in NO_ID_WORDS)
        if not iid and "intent_id" in cols:
            match = next((r for r in lni["intents_by_name"] if similarity(name, r.get("_contractor")) >= NAME_MATCH
                          and not ignored_at_lni(r.get("_status"), cfg)), None) if lni["ok"] else None
            if bad_iid:
                if match:
                    _with_lni_number(out, rn, "intent_id", match["_id"],
                                     f"L&I has intent {match['_id']} for this contractor")
            elif match:
                out.append((rn, "intent_id", g("intent_id"), match["_id"],
                            f"L&I has intent {match['_id']} for this contractor; tracker has none"))
            elif status in started:
                out.append((rn, "intent_id", g("intent_id"), None, f"Status is '{g('status')}' but no intent is on file"))
            elif (lni["ok"] and lni["project_ids"] and is_blank(raw_iid) and _is_name(name)
                  and (status or not _NOT_A_CONTRACTOR.search(name))):    # not a total or a note under the table
                if _not_started(status, g("start"), "start" in cols, c):
                    counts["not_started"] += 1     # nothing is due before work starts (David, Nov 7, 2026)
                    counted_not_started = True
                else:
                    out.append((rn, "intent_id", raw_iid, None, NOT_AT_LNI))

        # --- affidavit
        raw_aid = g("affidavit_id")
        if "affidavit_id" in cols:
            if not is_blank(raw_aid) and not isinstance(raw_aid, bool) and aid is None and str(raw_aid).strip().lower() not in NO_ID_WORDS:
                out.append((rn, "affidavit_id", raw_aid, None, "Affidavit ID is not a valid number"))
            # Affidavit Filed ticked but no Affidavit ID: marked only when L&I has the affidavit, with its number
            # ("Missing affidavit found in L&I's records", below); with none at L&I yet there is nothing for staff
            # to act on, so it is counted in the Audit sheet (David, Nov 8, 2026).
        rec = lni["affidavits"].get(aid) if aid and lni["ok"] else None
        if aid and lni["ok"]:
            if rec is None:
                out.append((rn, "affidavit_id", aid, None, "Affidavit ID not found in L&I records for this project"))
            else:
                if lni["project_ids"] and rec.get("_project_id") and rec["_project_id"] not in lni["project_ids"]:
                    out.append((rn, "affidavit_id", aid, rec["_project_id"], "Affidavit belongs to a different L&I project"))
                if iid and rec.get("_intent_id") and rec["_intent_id"] != iid:
                    out.append((rn, "affidavit_id", aid, rec["_intent_id"], f"Affidavit is linked to intent {rec['_intent_id']}, not {iid}"))
                elif rec.get("_intent_id"):  # several affidavits under one intent: the newest is the one that counts
                    under = rec["_intent_id"]
                    newest = (_affs_for(lni, under) or [rec])[0]
                    if newest["_id"] != aid:
                        out.append((rn, "affidavit_id", aid, newest["_id"],
                                    f"L&I has a newer affidavit {newest['_id']} for intent {under}"
                                    + (f" (received {newest['_received']})" if newest.get("_received") else "")
                                    + f"; tracker has {aid}"))
        found = _affs_for(lni, iid) if lni["ok"] and not aid and iid else []
        by_name = _affs_for(lni, match["_id"]) if lni["ok"] and not aid and not found and match else []
        if bad_aid and (found or by_name):
            a = (found or by_name)[0]
            _with_lni_number(out, rn, "affidavit_id", a["_id"],
                             f"L&I has affidavit {a['_id']} (intent {iid if found else match['_id']})")
        elif (found or by_name) and not bad_aid:
            a, n = (found or by_name)[0], len(found or by_name)
            role = "affidavit_id" if "affidavit_id" in cols else "affidavit_flag"
            more = f" (the newest of {n}: {', '.join(str(x['_id']) for x in (found or by_name)[1:])} too)" if n > 1 else ""
            under = iid if found else match["_id"]            # the intent found by name: its affidavit too
            out.append((rn, role, raw_aid, a["_id"], missing_affidavit(a["_id"], under) + more))
        elif (aflag and aid is None and iid and "affidavit_id" in cols and lni["ok"] and lni["project_ids"]
              and (is_blank(raw_aid) or isinstance(raw_aid, bool) or str(raw_aid).strip().lower() in NO_ID_WORDS)
              and not counted_not_started):
            # ticked, no number, and none under the line's intent at L&I: counted, not marked (David, Nov 8, 2026).
            # Only when L&I was compared (else the run already says it wasn't), for a line with an intent (one with
            # none is marked on its Intent ID cell), not when the cell holds text already marked "not a valid
            # number", and once per line (not also as a not-started line).
            counts["filed_not_at_lni"] += 1
        elif (lni["ok"] and status in done and not aflag and not aid and not bad_aid
              and ("affidavit_flag" in cols or "affidavit_id" in cols)):
            # work completed and no affidavit anywhere: filing can take long, nothing for staff to do yet (David,
            # Nov 7, 2026): counted in the Audit sheet, not marked. (Once L&I has one, the branch above marks it.)
            counts["no_affidavit_yet"] += 1
    # de-duplicate
    seen, uniq = set(), []
    for x in out:
        if (x[0], x[1], x[4]) not in seen:
            seen.add((x[0], x[1], x[4]))
            uniq.append(x)
    return uniq


NOT_AT_LNI = "No intent for this contractor in L&I's records for this project"
HIRING_MISMATCH = "Hiring contractor doesn't match the L&I intent"
# How names are compared, in staff's words (the Review Log's mismatch heading and the Audit tab; David, Nov 9, 2026)
HIRING_RULE = ("How names are compared: legal and trade words (Inc, LLC, Co, Construction, Contractors, Builders, "
               "Electric, General...) are left out, short forms are written out (NW = Northwest, Bros = Brothers), "
               "and the rest must be the same (in any order) or nearly, word for word; one name must be part of the "
               "other (a place or family word alone, like Pacific, Northwest or Brothers, isn't enough); or initials "
               "written as initials must spell the name (JCI or J.C.I. = Jones Construction Inc). Lines are compared only with their own L&I intent, when it "
               "is this project's and names the line's contractor, and only when both name a hiring company")
# Words that don't tell two companies apart (after util.norm: lower case, no punctuation, plural "s" dropped).
# Review of #121: "Quillmere Construction" and "Norvexa Construction Inc" must not match on "construction".
_GENERIC = {"construction", "constructor", "contractor", "contracting", "corporation", "incorporated", "group",
            "service", "builder", "building", "general", "enterprise", "ent", "industrie", "industry", "electric",
            "electrical", "mechanical", "plumbing", "concrete", "paving", "roofing", "excavation", "excavating",
            "steel", "glass", "painting", "lp", "pllc", "llp", "pc", "usa", "us"}
# What a tracker's Hiring Contractor cell holds when it names nobody: treated as blank when every word is one of
# these ("N/A - Prime", "Same as above", "Owner/GC", "G.C.") or it is only digits (review of #121)
_NO_HIRING_WORDS = {"n", "a", "na", "tbd", "tba", "none", "self", "owner", "prime", "g", "c", "gc", "general",
                    "contractor", "same", "as", "above", "see", "direct", "pending", "ditto", "x", "unknown", "no",
                    "the", "to", "be", "determined", "s", "perform", "performed", "performing"}
# Short forms written out before comparing (review of #121): "NG Electric" = "NORTHGLEN ELECTRIC INC"
_SHORT_FORMS = {"nw": "northwest", "ne": "northeast", "sw": "southwest", "se": "southeast", "const": "construction",
                "constr": "construction", "bro": "brother", "bros": "brother", "elec": "electric",
                "mech": "mechanical", "svc": "service", "svcs": "service", "intl": "international", "natl": "national",
                "one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7",
                "eight": "8", "nine": "9", "ten": "10", "and": "", "&": ""}


def _core(part: str) -> str:
    """A company name's distinctive words: "acme paving" -> "acme"; all of it when every word is generic."""
    words = [w for w in part.split() if w not in _GENERIC]
    return " ".join(words) if words else part


def names_no_hiring(v) -> bool:
    """Nobody: blank, a placeholder ("-", "TBD", "Self", "GC"...), or only legal words ("Co.")."""
    from .util import _name_parts, norm
    if is_blank(v) or isinstance(v, bool) or not _name_parts(v):
        return True
    words = norm(v).split()
    return all(w in _NO_HIRING_WORDS for w in words) or all(w.isdigit() for w in words)


def _expand(part: str) -> str:
    return " ".join(x for x in (_SHORT_FORMS.get(w, w) for w in part.split()) if x)


# Region and family words written out from short forms: never enough on their own for "part of" (review of #121, F1:
# "NW Construction" isn't "TAVERNO NORTHWEST", "Bros Paving" isn't "KELDORN BROTHERS INC")
_WEAK = {"northwest", "northeast", "southwest", "southeast", "brother", "north", "south", "east", "west", "pacific",
         "american", "national", "international", "united", "western", "eastern", "northern", "southern", "coast",
         "coastal", "puget", "sound", "washington", "seattle", "tacoma", "spokane", "county", "city", "state"}


def _raw_parts(raw):
    """The names of a "dba" chain as written ("ACME CO dba ZETA" -> ["ACME CO", "ZETA"])."""
    return [p.strip(" ,;-") for p in re.split(r"\bd\.?\s?b\.?\s?a\b\.?", str(raw or ""), flags=re.I) if p.strip(" ,;-")]


def _shown(core: str, raw) -> str:
    """The distinctive words of one name as written in raw (that dba part's text), quoted, each once: core
    "prime builder" of "PRIME BUILDERS, INC." -> '"PRIME BUILDERS"'; the core itself if its words can't be found."""
    from .util import norm
    want, out, seen = set(core.split()), [], set()
    for w in re.findall(r"[A-Za-z0-9&]+(?:'[A-Za-z]+)?", str(raw or "")):       # "O'Brien", "Mike's" stay whole
        nw = _expand(norm(re.sub(r"['’]s\b", "s", w).replace("'", "").replace("’", ""))).split()
        if nw and all(x in want for x in nw) and w.lower() not in seen:
            seen.add(w.lower())
            out.append(w)
    return '"' + (" ".join(out) if out else core) + '"'


def _initials(words) -> str:
    return "".join(w[0] for w in words if w)


def compare_companies(tracker_name, lni_name):
    """(same, how): the hiring company check's name rule (#121), stricter than the contractor one, and in plain words
    why it decided (for the Review Log and the Audit tab: David, Nov 9, 2026, "so staff know why"). Legal and trade
    words are left out ("Inc", "Construction", "Electric"...) and short forms written out ("NW" = Northwest), then:
    the distinctive words are the same (in any order) or nearly ("Prime Builders" = "PRIME BUILDERS, INC." = "Primee
    Builders"), one name's distinctive words are all in the other (whole words; four letters or more as written, and
    not only region or family words), or one is the other's initials ("JCI" or "J.C.I." = "Jones Construction Inc").
    Each name of a "dba" chain counts; a mismatch quotes the closest pair. how is "" for the same name."""
    from difflib import SequenceMatcher
    from .util import _name_parts, norm
    # "Mike's" is "Mikes" (review of #121, F4)
    # and "O'Brien" is "OBrien" (verification N1): apostrophes go, so the name stays one word; reasons quote the
    # names as written
    orig_t, orig_l = tracker_name, lni_name
    tracker_name = re.sub(r"['’]", "", re.sub(r"'s\b", "s", str(tracker_name or ""), flags=re.I))
    lni_name = re.sub(r"['’]", "", re.sub(r"'s\b", "s", str(lni_name or ""), flags=re.I))
    tw, lw = norm(tracker_name).split(), norm(lni_name).split()
    t_all = "".join(tw)
    raw_t = str(orig_t).strip()
    as_initials = all(len(w) == 1 for w in tw) or (len(tw) == 1 and raw_t.isupper())   # "JCI", "J.C.I.", not "Bolt"
    if as_initials and 2 <= len(t_all) <= 6 and \
            t_all in (_initials(lw), _initials([w for w in lw if w not in _GENERIC])):
        return True, f'"{str(orig_t).strip()}" taken as the initials of "{str(orig_l).strip()}"'
    tp, lp = _name_parts(tracker_name), _name_parts(lni_name)
    traw, lraw = _raw_parts(orig_t), _raw_parts(orig_l)
    traw = traw if len(traw) == len(tp) else [orig_t] * len(tp)
    lraw = lraw if len(lraw) == len(lp) else [orig_l] * len(lp)
    best = None                                     # (closeness, ca, cb, raw a, raw b): the closest pair, for the reason
    for a0, ra in zip(tp, traw):
        for b0, rb in zip(lp, lraw):
            a, b = _expand(a0), _expand(b0)
            ca, cb = _core(a), _core(b)
            wa, wb = ca.split(), cb.split()
            r = SequenceMatcher(None, ca, cb).ratio()
            if best is None or r > best[0]:
                best = (r, ca, cb, ra, rb)
            if ca == cb:
                return True, ""
            if set(wa) == set(wb):
                return True, f"{_shown(ca, ra)} and {_shown(cb, rb)} have the same words in another order"
            if len(wa) == len(wb) and r >= 0.8:        # word for word (review of #121, F2: not "Pacific" ~ "Brivantel Pacific")
                return True, f"{_shown(ca, ra)} taken as a spelling of {_shown(cb, rb)}"
            short_raw = min(len(_core(a0).replace(" ", "")), len(_core(b0).replace(" ", "")))
            for x, y, rx, ry in ((wa, wb, ra, rb), (wb, wa, rb, ra)):
                # not when one name is only place or family words and the other adds a real name ("NW" isn't
                # "TAVERNO NORTHWEST"); fine when it only adds place words ("Seattle" = "CITY OF SEATTLE")
                ga, gb = set(a.split()) & _GENERIC, set(b.split()) & _GENERIC    # their trades: not Paving vs Roofing
                same_trade = not ga or not gb or bool(ga & gb)
                if set(x) < set(y) and short_raw >= 4 and (not set(x) <= _WEAK or (set(y) - set(x) <= _WEAK
                                                                                   and same_trade)):
                    sx, sy = " ".join(x), " ".join(y)
                    return True, f"{_shown(sx, rx)} is part of {_shown(sy, ry)}"
            for x, y, rx, ry in ((a, b, ra, rb), (b, a, rb, ra)):
                xs = x.replace(" ", "")
                if (" " not in x and 2 <= len(xs) <= 6 and str(rx).strip().replace(".", "").isupper()
                        and xs in (_initials(y.split()), _initials(_core(y).split()))):
                    return True, f'"{str(rx).strip()}" taken as the initials of "{str(ry).strip()}"'
    if best is None:
        return False, "no name to compare"
    _r, ca, cb, ra, rb = best
    return False, f"the distinctive words differ: {_shown(ca, ra)} vs {_shown(cb, rb)}"


def same_company(tracker_name, lni_name) -> bool:
    return compare_companies(tracker_name, lni_name)[0]


def _hiring_issue(values, cols) -> str:
    """The Review Log's words for a hiring contractor mismatch, with the line's intent: "Hiring contractor doesn't
    match L&I intent 123456"."""
    iid = _cell_id(values[cols["intent_id"]]) if "intent_id" in cols and cols["intent_id"] < len(values) else None
    return f"Hiring contractor doesn't match L&I intent {iid}" if iid else HIRING_MISMATCH
NO_ID_WORDS = ("no", "n/a", "na", "false")    # in an Affidavit ID cell: no number, not "not a valid number"


def _with_lni_number(out, rn, role, lni_id, words):
    """Add what L&I has to the line's "... is not a valid number" finding: "Intent ID is not a valid number; L&I has
    intent 8800040 for this contractor" (one finding per cell, not a second one calling the cell empty)."""
    for k, (r, ro, v, _l, issue) in enumerate(out):
        if r == rn and ro == role and issue.endswith("is not a valid number"):
            out[k] = (r, ro, v, lni_id, f"{issue}; {words}")
            return


def missing_affidavit(aff_id, intent_id) -> str:
    """An affidavit L&I has and the tracker doesn't (David's wording, Nov 7, 2026)."""
    return f"Missing affidavit found in L&I's records: {aff_id} (intent {intent_id}); add it to the tracker"


_START_NOT_YET = re.compile(r"\s*(not\s*(yet\s*)?started|tbd|to be determined|pending)\W*", re.I)


def _start_date(v):
    """A start date cell as a date: a date, a date written as text, or an Excel date serial number (a cell that
    lost its date format)."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 <= v <= 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    d = to_date(v)
    if d is None and isinstance(v, str):
        for fmt in ("%b %d, %Y", "%B %d, %Y", "%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return d


def _not_started(status, start, has_start, c) -> bool:
    """A line whose contractor hasn't started work (David, Nov 7, 2026): its status says so
    (intent.not_started_statuses), or its start date says "not started" / TBD or is in the future, or it has no
    status at all and a blank start date. Any other status with a blank start date counts as started (Closed,
    Completed, In Progress...: a finished line must not go unmarked), and so does a start date the agent can't read.
    A tracker with no start date column says nothing about the date (only the status counts)."""
    words = c.get("not_started_statuses")
    if words is None:                        # not set; [] (the Settings page's "(none)") means no words (#151 review N2)
        words = ["not started"]
    if any(re.search(rf"\b{re.escape(x.lower())}\b", status) for x in words):
        return True
    if not has_start:
        return False
    if is_blank(start):
        return not status                    # no status and no start date: nothing says work has begun
    if isinstance(start, str) and _START_NOT_YET.fullmatch(start):
        return True
    d = _start_date(start)
    return d is not None and d > dt.date.today()


def _by_intent(affidavits):
    """{intent id: [affidavits filed under it, newest first]} (newest: latest received date, then highest ID)."""
    out = {}
    for a in affidavits:
        if a.get("_intent_id"):
            out.setdefault(a["_intent_id"], []).append(a)
    num = lambda a: int(a["_id"]) if str(a["_id"]).isdigit() else 0
    for v in out.values():
        if all(a.get("_received") for a in v):
            v.sort(key=lambda a: (a["_received"], num(a)), reverse=True)
        else:                           # a date is missing: L&I numbers forms in order, so the highest ID is newest
            v.sort(key=num, reverse=True)
    return out


def _affs_for(lni, intent_id):
    """The affidavits L&I has under an intent, newest first ([] for none). Accepts a single record too."""
    v = (lni.get("affidavits_by_intent") or {}).get(intent_id) if intent_id else None
    return [] if not v else (v if isinstance(v, list) else [v])


NAME_MATCH = 0.85               # how alike a tracker contractor and an L&I contractor must be to be the same
NO_NAME_AT_LNI = "(no contractor name at L&I)"


def _is_name(v) -> bool:
    """A contractor cell that can be a contractor: text with a letter (not a number, a date or a flag)."""
    return isinstance(v, str) and bool(re.search(r"[A-Za-z]", v))


_NOT_A_CONTRACTOR = re.compile(r"^\s*(sub-?\s*)?totals?\b|^\s*notes?\b|:\s*$", re.I)


def ignored_at_lni(status, cfg) -> bool:
    """An L&I intent whose status starts a word with one of intent.lni_ignore_statuses (withdrawn, cancelled...)."""
    st = str(status or "").lower()
    return any(re.search(rf"\b{re.escape(x.lower())}", st) for x in (cfg["intent"].get("lni_ignore_statuses") or []))


def lni_only(rows, cols, lni, cfg, prime_intent=None):
    """L&I intents filed on this project whose contractor isn't in the tracker: neither the intent number (nor the
    Project Info's prime intent) nor a contractor name (NAME_MATCH) appears on any tracker line (lines with a skipped
    status count: the contractor is there). Intents whose L&I status contains one of intent.lni_ignore_statuses
    (withdrawn, cancelled...) are left out. [(L&I contractor, [(intent id, L&I status)])], sorted by contractor.
    Empty without L&I data or without a known L&I project (nothing to compare with)."""
    if not lni["ok"] or not lni["project_ids"] or "contractor" not in cols:
        return []
    ids, names = {_num_id(prime_intent)} - {None}, []
    for _rn, v in rows:
        g = lambda role: v[cols[role]] if role in cols and cols[role] < len(v) else None
        name = g("contractor")
        if _is_name(name):
            names.append(name)
        if (i := _cell_id(g("intent_id"))):
            ids.add(i)
    out = {}
    for r in lni["intents_by_name"]:
        st = r.get("_status")
        if r["_id"] in ids or ignored_at_lni(st, cfg):
            continue
        c = r.get("_contractor")
        if not is_blank(c) and any(similarity(c, n) >= NAME_MATCH for n in names):
            continue
        out.setdefault(NO_NAME_AT_LNI if is_blank(c) else str(c).strip(), []).append((r["_id"], st))
    return sorted((c, sorted(v, key=lambda x: str(x[0]))) for c, v in out.items())


def _normalise(recs, fields, id_role):
    out = {}
    for r in recs:
        rid = _num_id(r.get(fields.get(id_role)))
        if not rid:
            continue
        r["_id"] = rid
        r["_project_id"] = _num_id(r.get(fields.get("project_id")))
        r["_contractor"] = r.get(fields.get("contractor"))
        r["_hiring"] = r.get(fields.get("hiring")) if fields.get("hiring") else None
        r["_status"] = r.get(fields.get("status"))
        r["_received"] = str(r.get(fields.get("received")) or "")[:10]   # ISO date: sorts as text
        r["_intent_id"] = _num_id(r.get(fields.get("intent_id"))) if id_role != "intent_id" else rid
        out[rid] = r
    return out


def _collect(client, tracker_ids, info, pcfg):
    fi = client.discover("intents")
    fa = client.discover("affidavits")
    seed = {x for x in tracker_ids + [info.get("prime_intent")] if x}
    pids = {str(pcfg["lni_project_id"])} if pcfg.get("lni_project_id") else set()
    if info.get("project_id"):
        pids.add(info["project_id"])
    if not pids and seed:
        from collections import Counter
        found = {_num_id(r.get(fi.get("intent_id"))): _num_id(r.get(fi.get("project_id")))
                 for r in client.fetch("intents", "intent_id", sorted(seed))}
        prime = found.get(info.get("prime_intent"))
        log.debug("Tracker intent IDs looked up", extra={"details": {
            "in_tracker": len(seed), "found_in_lni": sum(1 for v in found.values() if v),
            "prime_intent": info.get("prime_intent"), "prime_project": prime}})
        if prime:
            pids.add(prime)              # the prime contractor's intent defines the project
        else:
            common = Counter(p for p in found.values() if p).most_common(1)
            if common:
                pids.add(common[0][0])   # otherwise the project most tracker intents belong to
    log.info("L&I project identified", extra={"details": {"project_ids": ", ".join(sorted(pids)) or "none",
                                                           "how": "config/tracker" if (pcfg.get("lni_project_id") or info.get("project_id")) else "from intents"}})
    intents = client.fetch("intents", "project_id", sorted(pids)) if pids else []
    intents += client.fetch("intents", "intent_id", sorted(seed))
    affs = client.fetch("affidavits", "project_id", sorted(pids)) if pids and fa.get("project_id") else []
    ii = _normalise(intents, fi, "intent_id")
    affs += client.fetch("affidavits", "intent_id", sorted(ii)) if fa.get("intent_id") else []
    aa = _normalise(affs, fa, "affidavit_id")
    # de-duplicate raw rows for saving
    seen, raw_i = set(), []
    for r in intents:
        if r.get("_id") not in seen:
            seen.add(r.get("_id")); raw_i.append(r)
    seen, raw_a = set(), []
    for r in affs:
        if r.get("_id") not in seen:
            seen.add(r.get("_id")); raw_a.append(r)
    return pids, ii, aa, raw_i, raw_a


def _project_only():
    from .lni_live import LiveProjectUnreadable
    return LiveProjectUnreadable


def download(project_dir, tracker_ids, info, client, cfg, pcfg):
    """Fetch intents/affidavits for the project from data.wa.gov, or from L&I's live lookup when intent.lni_source is
    "live" (client: agent/lni_live.py, whose docstring sets out both sources field by field). Returns the lni dict
    used by validate(). When
    data.wa.gov can't be reached, or L&I hasn't published for too long (lni_unusable), nothing is downloaded and
    lni["skip"] holds the NOT CHECKED line: the check is skipped, never run on older records (David, Nov 8, 2026). What
    was used is saved in .agent/lni_data/ and Files/L&I data/<ts>/."""
    import json
    lni = {"intents": {}, "affidavits": {}, "intents_by_name": [], "affidavits_by_intent": {},
           "project_ids": set(), "ok": False, "error": None, "note": None}
    from .layout import Paths
    d = Paths(project_dir).lni_data
    snap_path = d / "lni_snapshot.json"
    tz = cfg.get("timezone", "UTC")
    saved_at, source = None, getattr(client, "name", "data.wa.gov")
    lni["skip"] = for_source(lni_unusable(client, cfg), source)
    if lni["skip"]:
        return lni
    try:
        try:
            pids, ii, aa, raw_i, raw_a = _collect(client, tracker_ids, info, pcfg)
            from .util import utc_stamp
            saved_at = utc_stamp()
        except _project_only() as e:
            # the live lookup answered, but not usably for this project only: the next project is still checked
            log.warning("%s didn't answer usably for this project (%s); its intent/affidavit check was not run",
                        source, e)
            lni["skip"] = LNI_PROJECT_UNREADABLE.format(why="see the administrator's email")
            return lni
        except requests.RequestException as e:
            # reachable a moment ago, but not now: this project and the rest of the run are not checked (no older
            # copy is used); projects checked before stand. A refusal (4xx but 429) or an answer that isn't JSON is a
            # setup problem ("not readable"), not an outage.
            code = getattr(getattr(e, "response", None), "status_code", None)
            refused = isinstance(e, requests.HTTPError) and code is not None and 400 <= code < 500 and code != 429
            from .lni_live import LiveUnreadable
            unreadable = refused or isinstance(e, (getattr(requests, "JSONDecodeError", ()), LiveUnreadable))
            log.warning("%s stopped answering the L&I queries (%s%s)", source, type(e).__name__,
                        f", HTTP {code}" if code else "")
            line = LNI_UNREADABLE.format(why=f"HTTP {code}" if refused else "not a data answer") if unreadable \
                else LNI_DOWN
            lni["skip"] = lni_stopped_answering(cfg, for_source(line, source), source)
            return lni
        lni["source"] = source
        lni.update(intents=ii, affidavits=aa, project_ids=pids, ok=True,
                   intents_by_name=[r for r in ii.values() if r["_project_id"] in pids],
                   affidavits_by_intent=_by_intent(aa.values()))
        # save what was used (pre-saving the previous files first)
        d.mkdir(parents=True, exist_ok=True)
        clean = lambda rows: [{k: v for k, v in r.items() if not k.startswith(("_", ":"))} for r in rows]
        snap_path.write_text(json.dumps({"saved_utc": saved_at, "source": source,
                                         "intents": clean(raw_i), "affidavits": clean(raw_a)}, indent=1))
        for name, recs in (("intents.csv", clean(raw_i)), ("affidavits.csv", clean(raw_a))):
            keys = sorted({k for r in recs for k in r})
            with open(d / name, "w", newline="") as f:
                w = csv.DictWriter(f, fieldnames=keys or ["none"], extrasaction="ignore")
                w.writeheader()
                for r in recs:
                    w.writerow(r)
        from .layout import Paths as _P
        keep = _P(project_dir).lni_run()                                  # what this run used, kept with the run
        keep.mkdir(parents=True, exist_ok=True)
        for f in visible_files(d):
            shutil.copy2(f, keep / f.name)
        if not pids:
            lni["error"] = "could not determine the L&I project id (set lni_project_id in config.yaml)"
    except Exception as e:
        # a file error's text names a path: on AWS its type only (#56); a download error keeps its explanation
        msg = (err_text(e) if isinstance(e, OSError) else str(e))[:300]
        lni["error"] = f"L&I download failed ({msg}); only checks that don't need L&I data were run"
        log.warning(lni["error"], exc_info=True)
    return lni


OVERWRITTEN = ("WARNING intent: the tracker was replaced by an older copy after the last check (its marks are from "
               "{seen}, the last check that marked it was {last}), so that check's marks were lost; marked again")


def run_intent_validation(project_dir: Path, cfg, state, client=None, force=False):
    """The intent/affidavit check; the administrator's notes about where the tracker came from (#93) are added
    whatever the check's outcome, early ones included (not applicable, open in Excel, not checked, ...)."""
    with _StrictIds(cfg):
        out = _run_intent_validation(project_dir, cfg, state, client, force)
    source = tracker_source(cfg["root"], project_dir.name)
    extra = [STAFF_DRIVE_UNREADABLE] if (source or {}).get("status") == "error" else []
    extra += [m for m in [new_match_line(source)] if m]
    return out + "".join("\n" + x for x in extra) if extra and isinstance(out, str) else out


def _run_intent_validation(project_dir: Path, cfg, state, client=None, force=False):
    tz = cfg.get("timezone", "UTC")
    pcfg = (cfg.get("projects") or {}).get(project_dir.name, {}) or {}
    if "intent_applicable" in pcfg and not pcfg["intent_applicable"]:
        # turned off by the administrator: say so even when there is no tracker (Cora Vale: "Not applicable", not
        # "No tracker" or "Tracker held back")
        return "intent: not applicable (set in config.yaml)"
    source = tracker_source(cfg["root"], project_dir.name)
    status = (source or {}).get("status")
    if status in SOURCE_NOT_CHECKED:
        return f"intent: NOT CHECKED - {SOURCE_NOT_CHECKED[status]}"
    if status == "no_tracker":
        return "intent: no compliance tracker in the staff folder"
    tracker = find_tracker(project_dir)
    if not tracker:
        return "intent: no compliance tracker in Tracker/"
    wb = openpyxl.load_workbook(tracker, data_only=True)  # read-only use; writes go through XlsxPatcher
    ok, why = is_applicable(project_dir.name, wb, cfg)
    log.debug("Tracker", extra={"details": {"applicable": ok, "why": why}})   # no path: it can hold a name (#56)
    if not ok:
        return f"intent: not applicable ({why})"
    if (tracker.parent / f"~${tracker.name}").exists():
        return (f"intent: SKIPPED - {tracker.name} is open in Excel (or Excel left its ~$ lock file). "
                "Close it and run again; nothing was changed")
    ws, hrow, cols = locate_table(wb)
    others = other_tables(wb, ws)
    info = project_info(wb)
    log.debug("Tracker table found", extra={"details": {"header_row": hrow,
                                                          "columns": ", ".join(f"{k}={openpyxl.utils.get_column_letter(v + 1)}" for k, v in cols.items()),
                                                          "prime_intent": info.get("prime_intent")}})
    for need in ("contractor", "intent_id"):
        if need not in cols:
            log.warning("Tracker has no '%s' column; related checks were skipped", need.replace("_", " "))
    rows = []
    for r in ws.iter_rows(min_row=hrow + 1):
        vals = [c.value for c in r]
        rows.append((r[0].row, vals))
    tracker_ids = [i for i in (_cell_id(v[cols["intent_id"]]) for _, v in rows if "intent_id" in cols and cols["intent_id"] < len(v)) if i]
    from .lni_live import make_client                 # data.wa.gov, or L&I's live lookup (intent.lni_source)
    lni = download(project_dir, tracker_ids, info, client or make_client(cfg), cfg, pcfg)
    if lni.get("skip"):
        return lni["skip"]                   # before anything is archived or marked: the tracker is left as it is
    counts = {}
    findings = validate(rows, cols, lni, cfg, counts)
    hiring_found = [f for f in findings if f[4] == HIRING_MISMATCH]     # #121: told apart everywhere below
    hiring_seen = counts.get("hiring") or {"compared": 0, "same": 0, "matched": [], "why": {}}
    hiring_why = hiring_seen["why"]
    missing_here = lni_only(rows, cols, lni, cfg, info.get("prime_intent"))   # at L&I, not in the tracker
    for rn, role, val, lval, issue in findings:
        log.debug("Tracker issue", extra={"details": {"row": rn, "field": role, "tracker_value": val,
                                                        "lni_value": lval, "issue": issue}})

    # write: pre-save the tracker, undo last run's marks, mark, log. The tracker is edited in
    # place at the XML level so nothing else in it (threaded comments, tasks, tables...) changes.
    from .layout import TRACKER, archive_files
    archive_files(project_dir, [tracker], TRACKER, move=False, label="before review")    # the tracker before marking
    from .xlsx_patch import XlsxPatcher
    xp = XlsxPatcher(tracker)
    audit_name = cfg["intent"].get("audit_sheet", "Audit")
    in_file = marks_from_tracker(wb, audit_name)            # recorded in the tracker itself (shared by app and AWS)
    remembered = state.get("intent_marks") or {}
    # The last check that marked this tracker is newer than the marks the tracker holds: someone saved an older copy
    # over it (Nimbrel, Nov 6, 2026: a stale Google Sheets tab undid a run). The marks the file holds are undone and
    # it is marked again as usual; the warning tells both emails (David, Nov 7, 2026).
    # Only when the file holds marks of an older check: a tracker with no marks record at all (a new tracker put
    # straight into Tracker/, an Audit tab deleted) is not called an older copy, and the remembered marks are undone
    # as before (only cells still in the agent's style are touched).
    last = state.get("intent_last") if remembered.get("file") == tracker.name else None
    overwritten = None
    if last and in_file is not None and str(in_file.get("run") or "") < str(last):
        overwritten = OVERWRITTEN.format(seen=in_file.get("run"), last=last)
        log.warning("The tracker was replaced by an older copy after the last check; its marks were lost and it is "
                    "marked again", extra={"details": {"marks_in_file": in_file.get("run"), "last_check": last}})
    # the marks to undo: those the file holds (an older copy holds that older check's), else the remembered ones
    prev = {**in_file, "file": tracker.name} if in_file else remembered
    if prev.get("file") == tracker.name and prev.get("sheet") in xp.sheet_names():
        cells = prev.get("cells", {})
        for ref, (orig_s, new_s, _) in cells.items():
            if xp.get_style(prev["sheet"], ref) == new_s:      # untouched since we marked it
                xp.set_style(prev["sheet"], ref, orig_s)
        for who in (AUTHOR, *OLD_AUTHORS):
            xp.remove_notes(prev["sheet"], {r for r, (_, _, noted) in cells.items() if noted}, who)
    colors = cfg["payroll"]["colors"]
    by_ref = {}
    for rn, role, val, lval, issue in findings:
        col = cols.get(role, cols.get("contractor"))
        ref = f"{openpyxl.utils.get_column_letter(col + 1)}{rn}"
        by_ref.setdefault(ref, []).append(issue + (f" (L&I: {lval})" if lval not in (None, "") else ""))
    marks = {}
    unnoted = 0
    for ref, msgs in by_ref.items():
        old_s, new_s = xp.fill_cell(ws.title, ref, colors["red"])
        noted = xp.add_note(ws.title, ref, "Compliance Validation Agent:\n" + "\n".join(f"- {m}" for m in msgs), AUTHOR)
        if not noted:
            unnoted += 1
            log.debug("Cell already has a note or comment; the message went to the Review Log only", extra={"details": {"cell": ref}})
        marks[ref] = [old_s, new_s, noted]
    by_row = {rn: v for rn, v in rows}
    ts = run_timestamp(tz)
    log_name = cfg["intent"]["review_log_sheet"]
    # notes about the whole tracker (too repetitive for each cell) go to the Audit sheet
    audit = [("Tracker", tracker_source_note(tracker.name, source, tz)),
             ("Applicability", f"L&I intent/affidavit review applies: {why}"),
             ("L&I data", f"{len(lni['intents'])} intent(s) and {len(lni['affidavits'])} affidavit(s) downloaded from "
                          f"the L&I records ({lni.get('source') or 'data.wa.gov'}) and compared with the tracker's IDs")]
    if overwritten:
        audit.append(("Warning", "This tracker had been replaced by an older copy after the check of "
                                 f"{last} (the marks in it were from {in_file.get('run')}), so "
                                 "that check's marks were lost; this check marked it again. Don't edit the CVA's copy "
                                 "of the tracker"))
    if lni.get("error"):
        audit.append(("L&I data", f"WARNING: {lni['error']}"))
    if lni.get("note"):
        audit.append(("L&I data", lni["note"]))
    missing = [r for r in ("contractor", "intent_id", "affidavit_id") if r not in cols]
    if missing:
        audit.append(("Columns", "Tracker columns not found, so their checks were skipped: "
                                 + ", ".join(m.replace("_", " ") for m in missing)))
    red = findings
    red_other = [f for f in red if f[4] != HIRING_MISMATCH]       # the mismatches have their own line (#121)
    if red_other:
        from collections import Counter
        kinds = Counter(_kind(f[4]) for f in red_other)
        cells = {(f[0], cols.get(f[1], cols.get("contractor"))) for f in red_other}
        audit.append(("Findings", f"{len(red_other)} invalid/missing value(s) on {len(cells)} cell(s), marked red: "
                                  + "; ".join(f"{k}: {n}" for k, n in kinds.most_common(8))))
    if not findings:
        audit.append(("Findings", "No invalid or missing values found"))
    # not due yet, so not marked (David, Nov 7, 2026): counted here, so a reader can see they were looked at
    if counts.get("not_started"):
        audit.append(("Not due yet", f"{counts['not_started']} line(s) with no intent in the tracker or at L&I whose "
                                     "work hasn't started (status, or a blank or future start date): not marked "
                                     "until work starts"))
    if counts.get("no_affidavit_yet"):
        audit.append(("Not due yet", f"{counts['no_affidavit_yet']} completed line(s) with no affidavit in the tracker "
                                     "or at L&I yet (filing can take a while): not marked until L&I has one"))
    if counts.get("filed_not_at_lni"):
        audit.append(("Not due yet", f"{counts['filed_not_at_lni']} line(s) ticked as affidavit filed, with no "
                                     "Affidavit ID and none under its intent at L&I yet: not marked until L&I has it "
                                     "(then its number is given)"))
    if hiring_seen["compared"]:
        ok = hiring_seen["same"] + len(hiring_seen["matched"])
        n_bad = len(hiring_found)
        audit.append(("Hiring contractor", f"{hiring_seen['compared']} line(s) compared with the hiring company on "
                      f"their L&I intent: {ok} match ({hiring_seen['same']} by the same name, "
                      f"{len(hiring_seen['matched'])} by the rule on the next line), {n_bad} "
                      + ("doesn't" if n_bad == 1 else "don't")
                      + (f" (marked red and listed first in the '{log_name}' sheet, each with why)" if n_bad else "")))
        audit.append(("Hiring contractor", HIRING_RULE))
        for rn, who, h, lh, how in hiring_seen["matched"][:30]:    # not the same name: say why, so staff can check
            audit.append(("Hiring contractor", f"Row {rn} ({who}): matched: tracker {h!s}, L&I {lh!s}: {how}"))
        if len(hiring_seen["matched"]) > 30:
            audit.append(("Hiring contractor", f"... and {len(hiring_seen['matched']) - 30} more line(s) matched by "
                                               "the rule"))
    if missing_here:
        audit.append(("Findings", f"{len(missing_here)} contractor(s) in L&I's records for this project are not in "
                                  f"the tracker; listed in the '{log_name}' sheet"))
    if unnoted:
        audit.append(("Notes", f"{unnoted} marked cell(s) already had a note or comment, so the agent's message is "
                               f"only in the '{log_name}' sheet"))
    if others:
        audit.append(("Columns", "Other tabs with a contractor table were not checked (only the first one is): "
                                 + ", ".join(f"'{t}'" for t in others)))
    skip = cfg["intent"].get("skip_statuses") or []
    if skip:
        audit.append(("Settings", "Tracker lines with status " + ", ".join(f"'{x}'" for x in skip) + " are not checked"))
    # The Review Log holds the latest check only (David, Nov 7, 2026: staff never type there; earlier checks are in
    # Tracker/Archived and the run logs): one summary line, then one line per finding.
    n_found = len(findings) + len(missing_here)
    if n_found:
        parts = ([f"{len(hiring_found)} hiring contractor mismatch(es)"] if hiring_found else []) \
            + ([f"{len(missing_here)} contractor(s) not in the tracker"] if missing_here else [])
        head = f"{n_found} finding(s)" + (f" ({'; '.join(parts)})" if parts else "")
    elif lni.get("error"):
        head = "No issues found in the checks that ran; L&I's records couldn't be compared this time"
    else:
        head = "No issues found"
    log_rows = [[ts, ws.title, "", "", "", "", "", head]]
    # Hiring contractor mismatches first, under their own heading, so they aren't buried under the cell rows (David,
    # Nov 9, 2026, #121); then the other findings, then the contractors not in the tracker
    others_found = [f for f in findings if f[4] != HIRING_MISMATCH]
    if hiring_found:
        log_rows.append([ts, ws.title, "", "", "", "", "",
                         f"HIRING CONTRACTOR MISMATCHES ({len(hiring_found)}): the tracker's Hiring Contractor isn't "
                         "the hiring company on the line's L&I intent. Check which is right; each line says why. How "
                         "names are compared: see the Audit tab"])
    for rn, role, val, lval, issue in hiring_found:
        col = cols.get(role, cols.get("contractor"))
        log_rows.append([ts, ws.title, f"{openpyxl.utils.get_column_letter(col + 1)}{rn}",
                         by_row[rn][cols["contractor"]], "hiring contractor",
                         "" if val is None else str(val), "" if lval is None else str(lval),
                         _hiring_issue(by_row[rn], cols)
                         + (f": {hiring_why[rn]}" if hiring_why.get(rn) else "")])
    if hiring_found and (others_found or missing_here):
        log_rows.append([ts, ws.title, "", "", "", "", "", f"OTHER FINDINGS ({len(others_found) + len(missing_here)})"])
    for rn, role, val, lval, issue in others_found:
        col = cols.get(role, cols.get("contractor"))
        log_rows.append([ts, ws.title, f"{openpyxl.utils.get_column_letter(col + 1)}{rn}",
                         by_row[rn][cols["contractor"]], role.replace("_", " "),
                         "" if val is None else str(val), "" if lval is None else str(lval), issue])
    for contractor, intents in missing_here:            # no tracker line to mark: the Review Log lists them
        log_rows.append([ts, ws.title, "", contractor, "contractor", "",
                         ", ".join(str(i) for i, _ in intents),
                         "In L&I's records for this project but not in the tracker ("
                         + "; ".join(f"intent {i}" + (f", {st}" if not is_blank(st) else "") for i, st in intents)
                         + "). If it is in the tracker under another name, make the names match"])
    xp.replace_log(log_name, ["Run (UTC)", "Sheet", "Cell", "Contractor", "Field", "Tracker value", "L&I value", "Issue"],
                   log_rows, widths=[20, 18, 7, 34, 14, 16, 16, 80])
    audit += [(MARKS_TOPIC, n) for n in marks_notes(ws.title, marks)]
    xp.append_log(audit_name, ["Run (UTC)", "Task", "Topic", "Note"],
                  [[ts, "intent/affidavit review", t, n] for t, n in audit], widths=[20, 22, 15, 120])
    xp.save(tracker)
    log.debug("Tracker saved", extra={"details": {"red_cells": len(marks), "log_rows": len(log_rows)}})
    state.set("intent_marks", {"file": tracker.name, "sheet": ws.title, "cells": marks})
    state.set("intent_last", ts)
    # the hiring contractor mismatches have their own line below (review of #121: not counted twice in the emails)
    summary = (f"intent: {len(red_other)} invalid/missing value(s) marked in {tracker.name}"
               + (f"; WARNING {lni['error']}" if lni.get("error") else "")
               + (f"; {lni['note']}" if lni.get("note") else "")
               + f" ({len(lni['intents'])} intents, {len(lni['affidavits'])} affidavits downloaded)")
    if red_other:                          # what kind of problems, most frequent first (numbers taken out)
        from collections import Counter
        kinds = Counter(_kind(f[4]) for f in red_other)
        summary += "\nintent:   " + "; ".join(f"{k}: {n}" for k, n in kinds.most_common(6)) \
            + (f"; and {len(kinds) - 6} other kind(s)" if len(kinds) > 6 else "")
    if missing_here:                       # a count only: the names are in the tracker's Review Log
        summary += (f"\nintent: {len(missing_here)} contractor(s) in L&I's records for this project are not in the "
                    "tracker; listed in its Review Log")
    if hiring_found:                       # #121: a count only; the names are at the top of the tracker's Review Log
        summary += (f"\nintent: {len(hiring_found)} hiring contractor(s) in the tracker don't match the L&I intent; "
                    "listed at the top of its Review Log")
    not_due = [f"{counts[k]} {what}" for k, what in (
        ("not_started", "line(s) not started with no intent"),
        ("no_affidavit_yet", "completed line(s) with no affidavit at L&I yet"),
        ("filed_not_at_lni", "line(s) ticked as affidavit filed with none under its intent at L&I yet"))
        if counts.get(k)]
    if not_due:
        summary += "\nintent: not due yet, not marked: " + ", ".join(not_due)
    if overwritten:
        summary += "\n" + overwritten
    return summary
