"""The intent/affidavit part of the evaluation set: an invented compliance tracker, the invented L&I records it is
checked against, and a stand-in for L&I's live lookup that serves them (tools/run_evals.py starts it for the run).

**Everything here is invented** (contractors, intent and affidavit numbers, L&I project numbers). Contractors are
companies, not workers. Each company name is built on a coined word (Brivantel, Cedrowan...) so it can't be a real
firm's by chance; never replace one with a real name. Intent numbers are 88000nn, affidavit numbers 88100nn, L&I
project numbers 99001nn / 99002nn.

The stand-in is also the run's only proxy (snapshot_run.run_checkout ``proxy``): every request either step makes goes to
it. It answers exactly one kind: a POST to L&I's live lookup address (LNI_URL: the production address, in plain http so
the proxy can read it) with production's body, ``{"ServiceName": "PWIAPublicService", "UrlData": "api/Public/<operation>
/<id>", "RestHttpMethod": "GET"}``, for the three operations agent/lni_live.py calls (GetIntent, GetAffidavit,
GetProjectStructureTree), in the answer shapes lni_live reads (IntentId, ProjId, CoName, HiringCoName,
IntentStatusText, DocRecdDt, AffidavitList; AfdvtId, AffidavitStatus...). An ID it doesn't have answers ``null`` with
HTTP 200, as L&I's service does (confirmed Nov 10, 2026). Anything else (another address or body, CONNECT for https,
GET...) is refused (403 or 400). Every request is recorded in full (method, address, headers, body); run_evals.py
saves the record in the run folder, where its privacy search reads it too.

Each tracker case is one line: the tracker's values, the L&I records it needs, and the marks the check must make on it
(``must``: (column role, kind of finding)), with the record numbers its words must give (``ids``). A case with no
``must`` is a control: no mark at all. ``not_due``: a line that must not be marked, only counted as not due yet (David,
Nov 7-8, 2026). ``staff_note``: the cell to be marked already holds a staff member's note, so the check's message goes
to the Review Log only. The kinds are the check's own wording (agent/intent.py validate), recognised by KINDS below.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit

PROJECT = 9900101                      # the tracker's L&I project (its Project Info sheet says so)
OTHER_PROJECT = 9900202                # another L&I project, for "belongs to a different L&I project"
PRIME = "Evalco General Contractors Inc"
PRIME_INTENT = 8800001
TRACKER_FILE = "Compliance Tracker.xlsx"
TRACKER_SHEET = "Compliance Tracker"
HEADER = ["Status", "Contractor", "Hiring Contractor", "Intent Filed", "Intent ID", "Affidavit Filed", "Affidavit ID",
          "Start Date"]
# the tracker's columns, as the check calls them (code/filekinds/cva_filekinds/tracker.py TRACKER_ROLES)
ROLE_OF = {"Status": "status", "Contractor": "contractor", "Hiring Contractor": "hiring", "Intent Filed": "intent_flag",
           "Intent ID": "intent_id", "Affidavit Filed": "affidavit_flag", "Affidavit ID": "affidavit_id",
           "Start Date": "start"}
COLUMN_OF = {role: HEADER.index(h) + 1 for h, role in ROLE_OF.items()}          # 1-based sheet columns
STAFF_NOTE = "Checked with the contractor by phone (staff)"

# L&I's live lookup (agent/lni_live.py DEFAULT_URL), in plain http so the stand-in, as the run's proxy, can read it
LNI_URL = "http://secure.lni.wa.gov/wagelookup/GatewayProxy/UIProxy.aspx?op=UiGatewayRestfulOperation"
OPS = ("GetIntent", "GetAffidavit", "GetProjectStructureTree")

# The check's findings, by the start of its words (agent/intent.py validate and _hiring_issue; NOT_AT_LNI,
# HIRING_MISMATCH, missing_affidavit). A note line that matches none of these is "unknown" (a false mark).
KINDS = {
    "intent_not_a_number": r"Intent ID is not a valid number",
    "intent_flag_without_id": r"Intent marked filed/approved but Intent ID is missing",
    "intent_not_found": r"Intent ID not found in L&I records",
    "intent_other_project": r"Intent belongs to a different L&I project",
    "contractor_mismatch": r"Contractor name doesn't match the L&I intent",
    "hiring_mismatch": r"Hiring contractor doesn't match (the )?L&I intent",
    "approved_but_tracker_says_no": r"L&I shows this intent approved; tracker says no",
    "filed_but_lni_status_differs": r"Tracker says filed/approved but L&I status differs",
    "intent_found_by_name": r"L&I has intent \d+ for this contractor; tracker has none",
    "started_without_intent": r"Status is '.*' but no intent is on file",
    "no_intent_at_lni": r"No intent for this contractor in L&I's records for this project",
    "affidavit_not_a_number": r"Affidavit ID is not a valid number",
    "affidavit_not_found": r"Affidavit ID not found in L&I records",
    "affidavit_other_project": r"Affidavit belongs to a different L&I project",
    "affidavit_other_intent": r"Affidavit is linked to intent \d+, not \d+",
    "affidavit_newer_at_lni": r"L&I has a newer affidavit \d+",
    "affidavit_missing": r"Missing affidavit found in L&I's records: \d+",
}


def kind_of(text: str) -> str:
    for k, pat in KINDS.items():
        if re.match(pat, text.strip()):
            return k
    return "unknown"


def intent(iid, co, *, project=PROJECT, hiring=PRIME, status="Approved", received="2026-05-04", affidavits=()):
    return {"IntentId": iid, "ProjId": project, "CoName": co, "HiringCoName": hiring, "IntentStatusText": status,
            "CoUbiNum": f"60{iid}", "DocRecdDt": f"{received}T00:00:00",
            "AffidavitList": [{"AffidavitId": a} for a in affidavits]}


def affidavit(aid, iid, co, *, project=PROJECT, status="Approved", received="2026-09-01"):
    return {"AfdvtId": aid, "IntentId": iid, "ProjId": project, "CoName": co, "AffidavitStatus": status,
            "CoUbiNum": f"60{iid}", "DocRecdDt": f"{received}T00:00:00"}


def line(co, *, status="Onsite", hiring=PRIME, filed="Yes", iid=None, aff_filed="No", aid=None, start="2026-06-01"):
    """A tracker line; ``start``: an ISO date, a date, or None (blank)."""
    if isinstance(start, str):
        start = dt.datetime.fromisoformat(start)
    elif isinstance(start, dt.date):
        start = dt.datetime(start.year, start.month, start.day)
    return [status, co, hiring, filed, iid, aff_filed, aid, start]


def case(name, area, tracker_line, intents=(), affidavits=(), must=(), ids=None, not_due=None, staff_note=None):
    """``must``: [(role, kind)]; ``ids``: {(role, kind): [numbers its words must give]}; ``staff_note``: the role of
    the cell that already has a staff note."""
    return {"case": name, "area": area, "line": tracker_line, "intents": list(intents), "affidavits": list(affidavits),
            "must": list(must), "ids": {f"{r}:{k}": [str(x) for x in v] for (r, k), v in (ids or {}).items()},
            "not_due": not_due, "staff_note": staff_note}


SOON = dt.date.today() + dt.timedelta(days=90)          # a start date always in the future
CASES = [
    case("Prime contractor, intent approved (control)", "control",
         line(PRIME, hiring="N/A - Prime", iid=PRIME_INTENT), [intent(PRIME_INTENT, PRIME, hiring=None)]),
    case("Subcontractor finished, affidavit on file (control)", "control",
         line("Brivantel Paving LLC", status="Complete", iid=8800002, aff_filed="Yes", aid=8810002),
         [intent(8800002, "Brivantel Paving LLC", affidavits=(8810002, 8810014))],
         [affidavit(8810002, 8800002, "Brivantel Paving LLC", received="2026-09-10")]),
    case("Hiring contractor written as initials (control: matched by the rule)", "control",
         line("Cedrowan Electric Inc", hiring="EGC", iid=8800003), [intent(8800003, "Cedrowan Electric Inc")]),
    case("Intent ID not at L&I (a typo)", "intent",
         line("Harvetto Concrete", iid=8800099), must=[("intent_id", "intent_not_found")]),
    case("Intent ID not at L&I, the cell already has a staff note (Review Log only)", "intent",
         line("Halvorne Masonry Supply", iid=8800098), must=[("intent_id", "intent_not_found")], staff_note="intent_id"),
    case("Intent filed under another L&I project", "intent",
         line("Irnovak Steel Erectors", iid=8800005),
         [intent(8800005, "Irnovak Steel Erectors", project=OTHER_PROJECT)], must=[("intent_id", "intent_other_project")],
         ids={("intent_id", "intent_other_project"): [OTHER_PROJECT]}),
    case("Contractor isn't the one on the intent", "intent",
         line("Junoveth Roofing Co", iid=8800006), [intent(8800006, "Zentravo Drywall LLC")],
         must=[("contractor", "contractor_mismatch")]),
    case("Hiring contractor isn't the one on the intent", "hiring",
         line("Kestravin Glass and Glazing", iid=8800007),
         [intent(8800007, "Kestravin Glass and Glazing", hiring="Brivantel Paving LLC")],
         must=[("hiring", "hiring_mismatch")]),
    case("Intent ID is not a number", "intent",
         line("Lakmoren Plumbing Inc", status="Mobilizing", iid="88000O8"), must=[("intent_id", "intent_not_a_number")]),
    case("Intent ID with a letter after a real intent's number (one finding, naming L&I's number)", "intent",
         line("Kovastren Drywall", iid="8800040A"), [intent(8800040, "Kovastren Drywall")],
         must=[("intent_id", "intent_not_a_number")], ids={("intent_id", "intent_not_a_number"): [8800040]}),
    case("Intent and affidavit IDs with a note after the number (\"-Filed\": control, read by the number)", "control",
         line("Brennoval Shotcrete", status="Complete", iid="8800050-Filed", aff_filed="Yes", aid="8810050-Filed"),
         [intent(8800050, "Brennoval Shotcrete", affidavits=(8810050,))],
         [affidavit(8810050, 8800050, "Brennoval Shotcrete")]),
    case("Intent ticked as filed, no Intent ID, none at L&I", "intent",
         line("Calvenni Railings", status="Mobilizing"),
         must=[("intent_id", "intent_flag_without_id"), ("intent_id", "no_intent_at_lni")]),
    case("No intent in the tracker; L&I has one for the contractor", "intent",
         line("Mapolund Painting", filed="No"), [intent(8800009, "Mapolund Painting")],
         must=[("intent_id", "intent_found_by_name")], ids={("intent_id", "intent_found_by_name"): [8800009]}),
    case("On site, no intent anywhere", "intent",
         line("Norvexa Excavation LLC", filed="No"), must=[("intent_id", "started_without_intent")]),
    case("Status not a 'started' word, start date passed, no intent at L&I", "intent",
         line("Orcavelle Interiors", status="Mobilizing", filed="No"), must=[("intent_id", "no_intent_at_lni")]),
    case("Not started, no intent yet (not due: counted, not marked)", "not due",
         line("Pinvarra Masonry", status="Not started", filed="No", start=None), not_due="not_started"),
    case("Start date in the future, no intent yet (not due: counted, not marked)", "not due",
         line("Dunmarra Elevators", status="Mobilizing", filed="No", start=SOON), not_due="not_started"),
    case("Affidavit ID not at L&I", "affidavit",
         line("Quorvane Landscaping", status="Complete", iid=8800013, aff_filed="Yes", aid=8810099),
         [intent(8800013, "Quorvane Landscaping")], must=[("affidavit_id", "affidavit_not_found")]),
    case("Affidavit ID is not a number", "affidavit",
         line("Zefrano Tile", status="Complete", iid=8800022, aff_filed="Yes", aid="see email"),
         [intent(8800022, "Zefrano Tile")], must=[("affidavit_id", "affidavit_not_a_number")]),
    case("Affidavit filed under another L&I project", "affidavit",
         line("Ambrevo Coatings", status="Complete", iid=8800023, aff_filed="Yes", aid=8810023),
         [intent(8800023, "Ambrevo Coatings", affidavits=(8810023,))],
         [affidavit(8810023, 8800023, "Ambrevo Coatings", project=OTHER_PROJECT)],
         must=[("affidavit_id", "affidavit_other_project")],
         ids={("affidavit_id", "affidavit_other_project"): [OTHER_PROJECT]}),
    case("Affidavit filed under another contractor's intent", "affidavit",
         line("Rivolta Insulation", status="Complete", iid=8800014, aff_filed="Yes", aid=8810014),
         [intent(8800014, "Rivolta Insulation")],
         [affidavit(8810014, 8800002, "Brivantel Paving LLC", received="2026-08-20")],
         must=[("affidavit_id", "affidavit_other_intent")],
         ids={("affidavit_id", "affidavit_other_intent"): [8800002, 8800014]}),
    case("A newer affidavit at L&I than the tracker's", "affidavit",
         line("Silvanek Fencing", status="Complete", iid=8800015, aff_filed="Yes", aid=8810015),
         [intent(8800015, "Silvanek Fencing", affidavits=(8810015, 8810016))],
         [affidavit(8810015, 8800015, "Silvanek Fencing", received="2026-08-01"),
          affidavit(8810016, 8800015, "Silvanek Fencing", received="2026-09-15")],
         must=[("affidavit_id", "affidavit_newer_at_lni")],
         ids={("affidavit_id", "affidavit_newer_at_lni"): [8810016, 8800015]}),
    case("Affidavit at L&I, not in the tracker", "affidavit",
         line("Timbravo Flooring", status="Complete", iid=8800017),
         [intent(8800017, "Timbravo Flooring", affidavits=(8810017,))],
         [affidavit(8810017, 8800017, "Timbravo Flooring")], must=[("affidavit_id", "affidavit_missing")],
         ids={("affidavit_id", "affidavit_missing"): [8810017, 8800017]}),
    case("Finished, no affidavit anywhere yet (not due: counted, not marked)", "not due",
         line("Upvarna Sheet Metal", status="Complete", iid=8800018), [intent(8800018, "Upvarna Sheet Metal")],
         not_due="no_affidavit_yet"),
    case("Affidavit ticked as filed, none at L&I yet (not due: counted, not marked)", "not due",
         line("Valtrisk Survey", status="Complete", iid=8800019, aff_filed="Yes"),
         [intent(8800019, "Valtrisk Survey")], not_due="filed_not_at_lni"),
    case("L&I approved the intent; the tracker says not filed", "intent",
         line("Wesmarro Mechanical", filed="No", iid=8800020), [intent(8800020, "Wesmarro Mechanical")],
         must=[("intent_flag", "approved_but_tracker_says_no")]),
    case("Tracker says filed; L&I's status is something else", "intent",
         line("Yelvanti Signs", iid=8800021), [intent(8800021, "Yelvanti Signs", status="Pending Review")],
         must=[("intent_flag", "filed_but_lni_status_differs")]),
]
# Lines the check gets wrong today (an open issue): scored with their correct expectation; ending as ``currently``
# says is tolerated (reported, not counted), any other result is a difference; passing means: close the gap.
GAPS = {}           # none open (#175 closed by intent.strict_ids, which the harness turns on: Nov 13, 2026)
# L&I intents on the project with no tracker line: listed in the Review Log ("not in the tracker"), unless withdrawn
LNI_ONLY = [intent(8800030, "Foxtrennic Traffic Control")]
LNI_ONLY_IGNORED = [intent(8800031, "Havrolde Demolition", status="Withdrawn")]

# The second tracker: its Project Info says no L&I filing is required, so it is not checked (and L&I not asked)
NOT_APPLICABLE_LINES = [line("Aspenvik Grading", iid=8800090), line("Birvanto Cabinets", iid=8800091)]


def records():
    """(intents {id: answer}, affidavits {id: answer}, trees {project: answer}) the stand-in serves."""
    ii, aa = {}, {}
    for c in CASES:
        for r in c["intents"]:
            ii[r["IntentId"]] = r
        for r in c["affidavits"]:
            aa[r["AfdvtId"]] = r
    for r in LNI_ONLY + LNI_ONLY_IGNORED:
        ii[r["IntentId"]] = r
    trees = {}
    for p in (PROJECT, OTHER_PROJECT):
        mine = [i for i, r in ii.items() if r["ProjId"] == p]
        affs = [a for a, r in aa.items() if r["ProjId"] == p]
        if mine or affs:
            # the Project Dashboard's shape: the prime's node, its intents and affidavits, subcontractors nested
            trees[p] = [{"ContractorName": PRIME, "IsPrimeContractor": True,
                         "Intents": [{"IntentId": PRIME_INTENT}] if PRIME_INTENT in mine else [],
                         "Affidavits": [{"AffidavitId": a} for a in affs if aa[a]["IntentId"] not in mine],
                         "SubContractorTreeNodes": [{"ContractorName": ii[i]["CoName"], "Intents": [{"IntentId": i}],
                                                     "Affidavits": [{"AffidavitId": a} for a in affs
                                                                    if aa[a]["IntentId"] == i]}
                                                    for i in mine if i != PRIME_INTENT]}]
    return ii, aa, trees


class FakeLookup:
    """The run's proxy and L&I's live lookup, on 127.0.0.1, serving records(). ``proxy``, once started: the address
    for HTTP(S)_PROXY. ``requests``: every request, in full, as dicts (method, url, headers, body, op, id, answered)."""

    def __init__(self):
        self.intents, self.affidavits, self.trees = records()
        self.requests = []
        self.proxy = None
        self._server = None
        self._lock = threading.Lock()

    def answer(self, op: str, ident: str):
        n = int(ident)
        if op == "GetIntent":
            return self.intents.get(n)
        if op == "GetAffidavit":
            return self.affidavits.get(n)
        return self.trees.get(n)                                        # GetProjectStructureTree

    @staticmethod
    def understand(method: str, url: str, raw: bytes):
        """(operation, id) when the request is exactly what lni_live sends to L&I, else None."""
        if method != "POST" or url != LNI_URL:
            return None
        try:
            body = json.loads(raw)
        except (ValueError, UnicodeDecodeError):
            return None
        if not isinstance(body, dict) or set(body) != {"ServiceName", "UrlData", "RestHttpMethod"} \
                or body["ServiceName"] != "PWIAPublicService" or body["RestHttpMethod"] != "GET":
            return None
        m = re.fullmatch(r"api/Public/(\w+)/(\d+)", str(body["UrlData"]))
        return (m[1], m[2]) if m and m[1] in OPS else None

    def _record(self, rec):
        with self._lock:
            self.requests.append(rec)

    def start(self):
        lookup = self

        class Handler(BaseHTTPRequestHandler):
            def _handle(self):
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    raw = self.rfile.read(min(n, 1_000_000)) if n > 0 else b""   # (larger: refused, first MB kept)
                    got = lookup.understand(self.command, self.path, raw)
                    lookup._record({"method": self.command, "url": self.path, "headers": dict(self.headers.items()),
                                    "body": raw.decode("utf-8", "replace"), "op": got[0] if got else None,
                                    "id": got[1] if got else None, "answered": bool(got)})
                    if not got:
                        other = self.command == "CONNECT" or urlsplit(self.path).hostname != urlsplit(LNI_URL).hostname
                        self.send_response(403 if other else 400)
                        self.send_header("Content-Length", "0")
                        self.end_headers()
                        return
                    data = json.dumps(lookup.answer(*got)).encode()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                except Exception as e:                                  # noqa: BLE001 - never unrecorded
                    lookup._record({"method": self.command, "url": self.path, "op": None, "id": None,
                                    "answered": False, "error": type(e).__name__})
                    try:
                        self.send_response(400)
                        self.end_headers()
                    except Exception:                                   # noqa: BLE001
                        pass

            do_POST = do_GET = do_PUT = do_DELETE = do_HEAD = do_PATCH = do_OPTIONS = do_CONNECT = _handle

            def log_message(self, *a):                                  # quiet
                pass

        self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        threading.Thread(target=self._server.serve_forever, daemon=True).start()
        self.proxy = f"http://127.0.0.1:{self._server.server_address[1]}"
        return self

    def stop(self):
        if self._server:
            self._server.shutdown()
            self._server.server_close()
            self._server = None

    def __enter__(self):
        return self.start()

    def __exit__(self, *exc):
        self.stop()
