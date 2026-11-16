"""L&I's live intent/affidavit lookup: a second source for the intent/affidavit check (David, Nov 10, 2026).

WHY
---
The check compares each tracker line with L&I's records of intents and affidavits. Its usual source, L&I's public
datasets on data.wa.gov, stopped updating on Oct 28, 2026 (they still say "posted daily"), so since then the check is
skipped ("L&I public data out of date (10-28-2026)"). L&I's own public lookup, "Search Approved Intents and
Affidavits" (https://secure.lni.wa.gov/wagelookup/intent/<intent id>), stays current. This module reads that lookup's
data instead, when `intent.lni_source: live` is set in config.yaml: **on since Nov 10, 2026** (David, after the
probe below). It is not an official feed; `intent.lni_source: data.wa.gov` goes back to the public datasets.

WHAT THE CHECK DOES TODAY (data.wa.gov; agent/intent.py LniClient)
------------------------------------------------------------------
Two datasets, queried with Socrata's API (`<socrata_base>/<dataset>.json?$where=<field> in (...)`):

    intents     t9je-9qwa  "L&I Intent Project Details"     one row per intent
    affidavits  9ncw-tqjn  "L&I Affidavit Project Details"  one row per affidavit

Before downloading, `intent.lni_unusable` reads each dataset's "rows updated" time (`/api/views/<dataset>.json`,
`rowsUpdatedAt`); older than `intent.lni_max_age_days` (3): the check is skipped in every project.

Per project (`intent._collect`): the L&I project ID comes from config (`lni_project_id`), the tracker's Project Info
sheet, or the project most of the tracker's intents belong to. Then: every intent of the project (by project ID),
every intent the tracker names (by intent ID, in case one belongs elsewhere), every affidavit of the project and of
those intents. `intent.discover` maps the dataset's field names to roles (FIELD_HINTS); `intent._normalise` keeps, per
record, only these roles, and the check reads nothing else:

    role          intents field                 affidavits field   what the check does with it
    intent_id     intent_id                     intent_id          finds the tracker line's intent; links an
                                                                   affidavit to its intent
    affidavit_id  -                             affidavit_id/id    finds the tracker line's affidavit
    project_id    project_id                    project_id         "belongs to a different L&I project"; the
                                                                   project's list for "not in the tracker"
    contractor    companyname                   companyname        "Contractor name doesn't match the L&I
                                                                   intent"; matching a line by name; "not in
                                                                   the tracker" rows
    hiring        hiringcompanyname             -                  the Hiring Contractor check (#121)
    status        status                        status             approved / withdrawn... (lni_ignore_statuses;
                                                                   the "filed/approved" flags)
    received      application_received_date     formreceiveddate   the newest of several affidavits of an intent;
                                                                   "(received <date>)" in a note
    ubi           ubi                           ubi                found, not used by a check yet

WHAT THE LIVE SOURCE DOES INSTEAD (this module)
-----------------------------------------------
The lookup page (a Vue app, pwse.js) gets all its data from one endpoint of L&I's PWIAPublicService:

    POST https://secure.lni.wa.gov/wagelookup/GatewayProxy/UIProxy.aspx?op=UiGatewayRestfulOperation
    Content-Type: application/json
    {"ServiceName": "PWIAPublicService", "UrlData": "api/Public/<operation>/<id>", "RestHttpMethod": "GET"}

(the page builds this body in oneLni.js from {operation: "/Public/<op>/<id>", httpMethod: "GET", ServiceName}; seen
Nov 9-10, 2026, in the browser, read-only). Operations the page uses:

    GetIntent/<intent id>                one intent (the answer seen for a real intent: fields below)
    GetAffidavit/<affidavit id>          one affidavit
    GetProjectStructureTree/<project id> the Project Dashboard: the prime contractor and nested
                                         SubContractorTreeNodes, each with its Intents and Affidavits
    GetProjectMissingIntents/<id>, GetProjectMissingAffidavits/<id>, DownloadAllProjectContractors, SearchForms
                                         (not used)

`LiveLniClient` answers the check's three questions (`discover`, `fetch`, `last_updated`) the way `LniClient` does,
with rows in data.wa.gov's own field names, so `intent._normalise`, `validate` and everything after are unchanged:

    check asks for                       live client does
    intents by intent ID                 GetIntent per ID
    intents by project ID                GetProjectStructureTree, every "IntentId" in the tree, GetIntent per ID
    affidavits by project ID             the same tree, every "AffidavitId", GetAffidavit per ID
    affidavits by intent ID              the intent's "AffidavitList", GetAffidavit per ID
    last_updated                         now (L&I's own system: no publish date to check)

Answer fields kept, by role (`intent_row`, `affidavit_row`):

    role        GetIntent answer                         GetAffidavit answer (from the page's code; probe it)
    intent_id   IntentId                                 IntentId
    affidavit   -                                        AffidavitId (or AfdvtId)
    project_id  ProjId (or ProjectInfo.ProjId)           ProjId (or ProjectInfo/AffidavitProject.ProjId)
    contractor  CoName                                   CoName (or ContractorName)
    hiring      HiringCoName                             -
    status      IntentStatusText (or StatusDescription)  AffidavitStatus (or AffidavitStatusText / AfdvtStatusText /
                                                         StatusDescription / StatusText
    received    DocRecdDt                                DocRecdDt
                (an ISO date-time or "/Date(ms)/" becomes "YYYY-MM-DD"; anything else: no date, never a status date)
    ubi         CoUbiNum                                 CoUbiNum (or ContractorUbi)

**Everything else in an answer is dropped at once**, never kept, saved or logged: the filer's name (FiledByCoName),
the company's email, phone and address, public notes, the agency contact, wages and worker counts per trade. (The
answer also has HiringCoUbiNum, HiringCoRegiNum, HiringCoIntentId, PrimeCntrctIntentId, the contractor's licence,
project name and contract number, intent details per trade: not used yet.)

Each run creates one client per project; within it, answers are cached (one GetIntent per intent), requests are
paced (`intent.live.pause_seconds`, 0.25 s), capped (`intent.live.max_requests_per_project`, 400) and time-limited
(`intent.live.max_seconds_per_project`, 600). What happens when an answer can't be used:

    network errors, 429, 5xx       net.post's retries first; then the service is down: this project and the rest
    (after the retries)            of the run aren't checked (LNI_DOWN naming the live lookup; staff see "L&I
                                   website down"); projects checked before stand
    an answer that isn't JSON      a setup problem for everyone (an error page): the rest of the run isn't checked
                                   (LNI_UNREADABLE; "L&I website not readable")
    a 4xx but 429; an error        this project only (LiveProjectUnreadable -> LNI_PROJECT_UNREADABLE; "L&I
    message instead of a record;   website: this project not readable"); the next project is checked
    a project tree with no intent;
    none of the tracker's intents
    returned; the cap; the time
    budget

An intent or affidavit is "not found" only when the service answered with data that isn't that record (an empty or
different object) while others of the project did come back; that's when a staff mark is made, as with data.wa.gov.
The probe below shows how a made-up ID answers: if it's an error (4xx or a message), a tracker typo skips its
project rather than being marked not found. If it's HTTP 5xx, don't switch the live source on: every tracker typo
would look like an outage (about 22 s of retries) and stop the rest of the run (review of #150, N2). The time limit
is checked before each request, so a project can take up to one request's retries (about 4 min) longer than it.

CONFIRMED, AND HOW TO CHECK AGAIN
--------------------------------
Probe run by David on Nov 10, 2026 from his Mac (a request not from the page: answered), on project 9900901 (a
prime's intent 8829901, a subcontractor's 8829902, affidavit 8839901):
  * GetIntent: a record; every role found (the hiring company only on a subcontractor's intent: a prime has none);
    DocRecdDt reads as a date.
  * GetProjectStructureTree: a list with the prime's node (ContractorName, IsPrimeContractor, Intents, Affidavits,
    SubContractorTreeNodes, HiringCoIntentId...); 56 intents and 7 affidavits found in it.
  * GetAffidavit: a record (AfdvtId, IntentId, DocRecdDt, CoName, CoUbiNum, AffidavitProject...); the status is
    AffidavitStatus ("Approved"), with AffidavitStatusCode ("A").
  * An intent ID that doesn't exist: HTTP 200 with null, so it's simply not found (a tracker typo is marked, as
    with data.wa.gov; it doesn't skip the project or stop the run).
Switched on the same day (David). If L&I changes the service, run the probe again:

    python -m agent.lni_live --probe <intent id> --project <L&I project id> --affidavit <affidavit id>

It prints each answer's status, field names and types and which roles were found, never a value, and tries an ID
that can't exist. Then set `intent.lni_source: live` (and back to data.wa.gov to stop).
"""
from __future__ import annotations

import re
import time

import requests

from .util import log

DEFAULT_URL = "https://secure.lni.wa.gov/wagelookup/GatewayProxy/UIProxy.aspx?op=UiGatewayRestfulOperation"
SERVICE = "PWIAPublicService"
SOURCES = ("data.wa.gov", "live")
# The record shape given to intent.py: data.wa.gov's own field names (intent.FIELD_HINTS finds them)
INTENT_FIELDS = {"intent_id": "intent_id", "project_id": "project_id", "contractor": "companyname",
                 "hiring": "hiringcompanyname", "status": "status", "ubi": "ubi", "received": "application_received_date"}
AFFIDAVIT_FIELDS = {"affidavit_id": "affidavit_id", "intent_id": "intent_id", "project_id": "project_id",
                    "contractor": "companyname", "status": "status", "ubi": "ubi", "received": "formreceiveddate"}


def source(cfg) -> str:
    """intent.lni_source: "data.wa.gov" (default) or "live"; anything else: data.wa.gov, with a warning."""
    v = str((cfg.get("intent") or {}).get("lni_source") or "data.wa.gov").strip().lower()
    if v not in SOURCES:
        log.warning("intent.lni_source in config.yaml must be data.wa.gov or live; data.wa.gov is used")
        return "data.wa.gov"
    return v


class LiveUnreadable(requests.RequestException):
    """L&I's live lookup answered, but not with data (an error page, not JSON, an error message): it doesn't work as
    expected for anyone, so the rest of the run isn't checked (as when data.wa.gov refuses)."""


class LiveProjectUnreadable(LiveUnreadable):
    """Only this project's lookups can't be used (an ID the service refuses, the request cap or time budget, nothing
    of the tracker's found): this project isn't checked; the next one is (review of #150, F2)."""


def _date(v):
    """An ISO day "YYYY-MM-DD" from an ISO date-time or a .NET "/Date(ms)/"; None otherwise (review of #150, F4)."""
    import datetime as dt
    s = str(v or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        return s[:10]
    m = re.match(r"^/Date\((-?\d+)([+-]\d{4})?\)", s)
    if m:
        off = m[2] or "+0000"
        tz = dt.timezone(int(off[0] + "1") * dt.timedelta(hours=int(off[1:3]), minutes=int(off[3:])))
        return dt.datetime.fromtimestamp(int(m[1]) / 1000, tz).strftime("%Y-%m-%d")
    return None


def _error_answer(j) -> bool:
    """An error message instead of a record ({"Message": "An error has occurred."} and the like)."""
    return isinstance(j, dict) and bool({"Message", "ExceptionMessage", "ExceptionType", "error", "Error"} & set(j)) \
        and not any(k in j for k in ("IntentId", "AffidavitId", "AfdvtId", "ProjId"))


def _first(d, *keys):
    for k in keys:
        if isinstance(d, dict) and d.get(k) not in (None, ""):
            return d.get(k)
    return None


def _ids(node, key, out):
    """Every value of `key` anywhere in a nested answer (the project tree: SubContractorTreeNodes, Intents, ...)."""
    if isinstance(node, dict):
        for k, v in node.items():
            if k == key and isinstance(v, (int, str)) and str(v).strip().isdigit():
                out.add(str(v).strip())
            else:
                _ids(v, key, out)
    elif isinstance(node, list):
        for v in node:
            _ids(v, key, out)
    return out


def intent_row(j):
    """The fields the check uses, from a GetIntent answer (field names seen Nov 9, 2026)."""
    proj = _first(j, "ProjId") or _first(j.get("ProjectInfo") or {}, "ProjId")
    return {"intent_id": _first(j, "IntentId"), "project_id": proj, "companyname": _first(j, "CoName"),
            "hiringcompanyname": _first(j, "HiringCoName"), "status": _first(j, "IntentStatusText", "StatusDescription"),
            "ubi": _first(j, "CoUbiNum"), "application_received_date": _date(_first(j, "DocRecdDt")),
            "_affidavit_ids": sorted(_ids(j.get("AffidavitList") or [], "AffidavitId", set()))}


def affidavit_row(j):
    """The fields the check uses, from a GetAffidavit answer (names from the lookup page's code; the probe confirms)."""
    proj = _first(j, "ProjId") or _first(j.get("ProjectInfo") or j.get("AffidavitProject") or {}, "ProjId")
    return {"affidavit_id": _first(j, "AffidavitId", "AfdvtId"), "intent_id": _first(j, "IntentId"),
            "project_id": proj, "companyname": _first(j, "CoName", "ContractorName"),
            "status": _first(j, "AffidavitStatus", "AffidavitStatusText", "AfdvtStatusText", "StatusDescription",
                             "StatusText"),        # AffidavitStatus ("Approved"): confirmed Nov 10, 2026
            "ubi": _first(j, "CoUbiNum", "ContractorUbi"),
            "formreceiveddate": _date(_first(j, "DocRecdDt"))}


class LiveLniClient:
    """The interface of intent.LniClient (discover, fetch, last_updated), answered from L&I's live lookup."""
    name = "L&I's live lookup"
    ds = {"intents": "live lookup", "affidavits": "live lookup"}       # (intent.lni_unusable's log details)

    def __init__(self, cfg, post=None, sleep=time.sleep):
        c = (cfg.get("intent") or {}).get("live") or {}
        self.url = c.get("url") or DEFAULT_URL
        self.pause = float(c.get("pause_seconds", 0.25))
        self.cap = int(c.get("max_requests_per_project", 400))
        self.budget = float(c.get("max_seconds_per_project", 600))      # review of #150, F7
        self.started = None
        self.requests = 0
        self.sleep = sleep
        if post is None:
            from . import net
            post = net.post
        self._post = post
        self._intents, self._affidavits, self._trees = {}, {}, {}

    # ---- the service
    def _call(self, op, ident):
        if self.requests >= self.cap:
            raise LiveProjectUnreadable(f"more than {self.cap} requests for one project "
                                        "(intent.live.max_requests_per_project)")
        if self.started is None:
            self.started = time.monotonic()
        elif time.monotonic() - self.started > self.budget:
            raise LiveProjectUnreadable(f"more than {self.budget:.0f} s for one project "
                                        "(intent.live.max_seconds_per_project)")
        if self.requests:
            self.sleep(self.pause)
        self.requests += 1
        try:
            r = self._post(self.url, json={"ServiceName": SERVICE, "UrlData": f"api/Public/{op}/{ident}",
                                           "RestHttpMethod": "GET"},
                           headers={"Accept": "application/json", "User-Agent": "CVA compliance check"},
                           timeout=60, what="L&I live lookup")
        except requests.HTTPError as e:
            code = getattr(e.response, "status_code", None)
            if code is not None and 400 <= code < 500 and code != 429:   # refuses this ID: this project only (F2)
                raise LiveProjectUnreadable(f"{op}: HTTP {code}")
            raise                                                         # down (after retries): the rest of the run
        try:
            j = r.json()
        except ValueError:
            raise LiveUnreadable(f"{op}: not a data answer")       # an error page: the type only, never its text
        if _error_answer(j):
            raise LiveProjectUnreadable(f"{op}: an error message, not data")   # review of #150, F1
        return j

    def intent(self, iid):
        iid = str(iid)
        if iid not in self._intents:
            j = self._call("GetIntent", iid)
            self._intents[iid] = intent_row(j) if isinstance(j, dict) and _first(j, "IntentId") else None
        return self._intents[iid]

    def affidavit(self, aid):
        aid = str(aid)
        if aid not in self._affidavits:
            j = self._call("GetAffidavit", aid)
            self._affidavits[aid] = affidavit_row(j) if isinstance(j, dict) and _first(j, "AffidavitId", "AfdvtId") \
                else None
        return self._affidavits[aid]

    def tree(self, pid):
        pid = str(pid)
        if pid not in self._trees:
            j = self._call("GetProjectStructureTree", pid)
            ii = _ids(j, "IntentId", set())
            if not isinstance(j, (dict, list)) or not ii:     # a project has at least its prime's intent (F1)
                raise LiveProjectUnreadable("GetProjectStructureTree: no intent in the answer")
            self._trees[pid] = (ii, _ids(j, "AffidavitId", set()))
        return self._trees[pid]

    # ---- intent.LniClient's interface
    def discover(self, which):
        return dict(INTENT_FIELDS if which == "intents" else AFFIDAVIT_FIELDS)

    def last_updated(self, which):
        """The live lookup is L&I's own system: current by definition (no publish date to check)."""
        import datetime as dt
        return dt.datetime.now(dt.timezone.utc)

    def fetch(self, which, role, values):
        values = [str(v) for v in values if v not in (None, "")]
        ids = set()
        if which == "intents":
            if role == "intent_id":
                ids = set(values)
            elif role == "project_id":
                for pid in values:
                    ids |= self.tree(pid)[0]
            rows = [self.intent(i) for i in sorted(ids, key=int)]
            # none of the tracker's intents (or of the tree's) came back: the service isn't answering with records (or
            # every ID is wrong); marking them all would be false marks, so this project isn't checked (F1; review N1)
            if ids and not any(rows):
                raise LiveProjectUnreadable(f"none of {len(ids)} intent(s) came back as a record")
        else:
            if role == "affidavit_id":
                ids = set(values)
            elif role == "project_id":
                for pid in values:
                    ids |= self.tree(pid)[1]
            elif role == "intent_id":
                for iid in values:
                    rec = self.intent(iid)
                    ids |= set((rec or {}).get("_affidavit_ids") or [])
            rows = [self.affidavit(a) for a in sorted(ids, key=int)]
        out = [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows if r]
        log.debug("L&I query", extra={"details": {"dataset": which, "field": role, "values": len(values),
                                                   "records": len(out), "source": "live", "requests": self.requests}})
        return out


def make_client(cfg):
    """The client for intent.lni_source: intent.LniClient (data.wa.gov) or LiveLniClient."""
    if source(cfg) == "live":
        return LiveLniClient(cfg)
    from .intent import LniClient
    return LniClient(cfg)


# ------------------------------------------------------------------ probe: field names and types, never values
_ONE_WORD_FIELDS = {"Intents", "Affidavits", "Message", "Data", "Items", "Results"}


def _field_name(k):
    """A key that looks like one of the service's field names (CamelCase with an inner capital, like IntentId or
    CoName, or a known one-word field), not a value used as a key (a company or trade name)."""
    k = str(k)
    return k in _ONE_WORD_FIELDS or bool(re.match(r"^[A-Z][a-z0-9]+(?:[A-Z][a-z0-9]*)+$", k)) and len(k) <= 40


def _shape(j, depth=0):
    if isinstance(j, dict):
        if not all(_field_name(k) for k in j):                    # data as keys (names, trades): count only (F6, N3)
            return f"<{len(j)} keys>"
        return {k: _shape(v, depth + 1) if depth < 2 else type(v).__name__ for k, v in j.items()}
    if isinstance(j, list):
        return [f"list of {len(j)}"] + ([_shape(j[0], depth + 1)] if j and depth < 2 else [])
    return type(j).__name__


def probe(intent_id, project_id=None, affidavit_id=None, post=None):
    """For each lookup: whether it answered with data, and the answer's shape. Returns a dict (printable)."""
    c = LiveLniClient({"intent": {"live": {"pause_seconds": 0.5, "max_requests_per_project": 10}}}, post=post)
    out = {}
    # the last: an ID that can't exist, to see how "not found" answers (data, or an error that would stop a project's
    # check: then a tracker typo would stop it too)
    for op, ident in (("GetIntent", intent_id), ("GetProjectStructureTree", project_id),
                      ("GetAffidavit", affidavit_id), ("GetIntent (an ID that doesn't exist)", "999999999")):
        if not ident:
            continue
        try:
            j = c._call(op.split(" ")[0], ident)
            out[op] = {"answered": "data", "shape": _shape(j)}
            if op.startswith("GetIntent") and isinstance(j, dict):
                row = intent_row(j)
                out[op]["is_record"] = bool(_first(j, "IntentId"))                          # (F5)
                out[op]["fields_found"] = {k: v is not None for k, v in row.items() if not k.startswith("_")}
                out[op]["received_is_a_date"] = row["application_received_date"] is not None
            if op == "GetAffidavit" and isinstance(j, dict):
                out[op]["is_record"] = bool(_first(j, "AffidavitId", "AfdvtId"))
                out[op]["fields_found"] = {k: v is not None for k, v in affidavit_row(j).items()}
            if op == "GetProjectStructureTree":
                out[op]["intent_ids"], out[op]["affidavit_ids"] = (len(_ids(j, "IntentId", set())),
                                                                   len(_ids(j, "AffidavitId", set())))
        except requests.HTTPError as e:
            out[op] = {"answered": f"HTTP {getattr(e.response, 'status_code', '?')}"}
        except LiveUnreadable as e:
            out[op] = {"answered": f"{type(e).__name__}: {e}"}        # our own message: no record content
        except requests.RequestException as e:
            out[op] = {"answered": type(e).__name__}
    return out


if __name__ == "__main__":
    import argparse
    import json
    ap = argparse.ArgumentParser(description="Check L&I's live lookup answers (field names and types only).")
    ap.add_argument("--probe", required=True, help="an intent ID")
    ap.add_argument("--project", help="an L&I project ID")
    ap.add_argument("--affidavit", help="an affidavit ID")
    a = ap.parse_args()
    print(json.dumps(probe(a.probe, a.project, a.affidavit), indent=1))
