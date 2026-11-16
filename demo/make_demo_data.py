"""Builds the CVA demo: two invented projects, their uploads and trackers, invented L&I records, and an answer key.

**Everything here is invented.** Companies, projects, workers, IDs, rates and L&I numbers are made up for the
capstone demo; none is a copy of, or derived from, real staff files. Company and project names are built on coined
words (Quillmere, Norvanta, Pellavine...) so they can't match a real firm by chance. Every invented worker has an
employee ID in the range 7320001-7320999, used nowhere else (the evaluation harness uses 7319nnn), so a search can
find every one of them (``demo/check_demo.py``). Intent IDs are 88200nn, affidavit IDs 88300nn and L&I project IDs
99005nn, in no range L&I uses for real records today. The wage rates are sample values in the firm's wage table
layout, not a real determination or union agreement.

    python demo/make_demo_data.py ~/Downloads/CVA_demo            # build (the folder must be empty or not exist)
    python demo/make_demo_data.py ~/Downloads/CVA_demo --reset    # delete what an earlier build made, build again

The output folder holds:
    config.yaml, notifications.config   the demo's own settings (L&I's lookup through the local stand-in,
                                        emails saved as .eml previews, no AWS, no error reporting)
    Projects/<GC> - <Project>/          the two projects, as staff hand files in (Uploads/, Tracker/)
    To upload/                          the large project's last weekly payroll, to add in the video
    demo_data/answer_key.json           what every seeded line and tracker line must show (tools/run_evals.py format)
    demo_data/lni_records.json          the invented L&I records demo/lni_standin.py serves
    demo_data/workers.json              every invented worker (name, ID, contractor, trade): for the privacy search

The same seed always gives the same files (except the tracker's "start date in the future" lines, which are
relative to today so they stay in the future).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import random
import shutil
import sys
from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import Font, PatternFill

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
AGENT = REPO / "code" / "compliance_agent"
for p in (AGENT, AGENT / "tools", AGENT / "tests"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

import eval_cases as EC                                  # noqa: E402 - the harness's writers (payroll, wage table)
import eval_lni as LNI                                   # noqa: E402 - L&I record shapes, the check's finding kinds
import test_seeded_violations as SV                      # noqa: E402 - the payroll layout and a consistent line
from agent.layout import Paths, ensure_layout            # noqa: E402

SEED = 20261123
MARKER = "cva-demo"                       # written in every folder this script makes (reset deletes only those)
ID_BASE = 7320000                         # invented employee IDs: 7320001...
D = dt.datetime

# ------------------------------------------------------------------------------------------------- the projects
SMALL = {
    "folder": "Quillmere Construction - Brightfen Library Annex",
    "full_name": "Brightfen Library Annex (sample project)",
    "prime": "Quillmere Construction Inc",
    "county": "King",
    "lni_project": 9900501,
    "prime_intent": 8820001,
    "owner": "Brightfen Public Library District (invented)",
    "contract": "BLA-2026-01",
    "address": "100 Sample Way, Brightfen WA (invented)",
}
LARGE = {
    "folder": "Norvanta Builders - Kestrel Point Housing Phase 2",
    "full_name": "Kestrel Point Housing Phase 2 (sample project)",
    "prime": "Norvanta Builders Inc",
    "county": "Pierce",
    "lni_project": 9900502,
    "prime_intent": 8820101,
    "owner": "Kestrel Point Housing Authority (invented)",
    "contract": "KPH-2026-02",
    "address": "2200 Example Avenue, Kestrel Point WA (invented)",
}
OTHER_LNI_PROJECT = 9900599                # "belongs to another L&I project" (not used in the demo's lines today)

# Sample wage rates in the firm's wage table layout (Trade, Classification, Base Rate...). Invented values.
STATE_RATES = {   # (trade, classification): (base, fringe)
    ("Carpenters", "Journey Level"): (52.00, 22.50),
    ("Laborers", "General Laborer"): (41.00, 16.50),
    ("Electricians", "Journey Level"): (61.00, 27.00),
    ("Plumbers & Pipefitters", "Journey Level"): (63.00, 28.00),
    ("Drywall Applicators", "Journey Level"): (50.00, 21.00),
    ("Painters", "Journey Level"): (40.00, 17.50),
    ("Cement Masons", "Journey Level"): (49.00, 19.00),
    ("Ironworkers", "Journey Level"): (58.00, 30.00),
    ("Roofers", "Journey Level"): (46.00, 15.50),
    ("Power Equipment Operators", "Group 2"): (57.00, 24.00),
    ("Glaziers", "Journey Level"): (54.00, 23.00),
    ("Landscape Construction", "Landscape Laborer"): (36.00, 12.00),
    ("Sheet Metal Workers", "Journey Level"): (59.00, 29.00),
    ("Insulation Applicators", "Journey Level"): (44.00, 16.00),
}
UNION_RATES = {("Laborers", "General Laborer"): (43.00, 21.00)}
UNION_DECISION = "Laborers agreement 2026 (sample)"
STATE_DECISION = "WA L&I 2026 (sample)"


def jur_state(county):
    return f"Washington State Wages - {county} County"


JUR_UNION = "Union rates - Laborers (sample agreement)"

# Contractors: name, tier ("prime", 1, 2), who hired it, crew [(trade, classification, union?)], weeks on the payroll
SMALL_CONTRACTORS = [
    ("Quillmere Construction Inc", "prime", None,
     [("Carpenters", "Journey Level", False)] * 3 + [("Laborers", "General Laborer", False)] * 2, (1, 2)),
    ("Pellavine Electric LLC", 1, "Quillmere Construction Inc",
     [("Electricians", "Journey Level", False)] * 4, (1, 2)),
    ("Ormsgate Plumbing & Heating", 1, "Quillmere Construction Inc",
     [("Plumbers & Pipefitters", "Journey Level", False)] * 3, (1,)),
    ("Kelvarro Drywall Inc", 1, "Quillmere Construction Inc",
     [("Drywall Applicators", "Journey Level", False)] * 4, (1,)),
    ("Fenrowe Taping Co", 2, "Kelvarro Drywall Inc",
     [("Drywall Applicators", "Journey Level", False)] * 3, (2,)),
    ("Tamsvik Painting", 2, "Kelvarro Drywall Inc",
     [("Painters", "Journey Level", False)] * 2, (2,)),
]
LARGE_CONTRACTORS = [
    ("Norvanta Builders Inc", "prime", None,
     [("Carpenters", "Journey Level", False)] * 6 + [("Laborers", "General Laborer", True)] * 5, (1, 2, 3, 4)),
    ("Dravello Concrete LLC", 1, "Norvanta Builders Inc",
     [("Cement Masons", "Journey Level", False)] * 6 + [("Laborers", "General Laborer", True)] * 4, (1, 2, 3, 4)),
    ("Rindlemoor Rebar Co", 2, "Dravello Concrete LLC",
     [("Ironworkers", "Journey Level", False)] * 4, (1, 2, 3)),
    ("Tessivar Steel Erectors", 1, "Norvanta Builders Inc",
     [("Ironworkers", "Journey Level", False)] * 6, (1, 2)),
    ("Corvanne Electric Inc", 1, "Norvanta Builders Inc",
     [("Electricians", "Journey Level", False)] * 7, (1, 2, 3, 4)),
    ("Holtavane Mechanical LLC", 1, "Norvanta Builders Inc",
     [("Plumbers & Pipefitters", "Journey Level", False)] * 5 + [("Sheet Metal Workers", "Journey Level", False)] * 3,
     (1, 2, 3, 4)),
    ("Velkorin Insulation", 2, "Holtavane Mechanical LLC",
     [("Insulation Applicators", "Journey Level", False)] * 3, (1, 2)),
    ("Zubrina Roofing Co", 1, "Norvanta Builders Inc",
     [("Roofers", "Journey Level", False)] * 5, (1, 2)),
    ("Marlowick Painting", 1, "Norvanta Builders Inc",
     [("Painters", "Journey Level", False)] * 5, (1, 2, 3, 4)),
    ("Fennacre Excavation", 1, "Norvanta Builders Inc",
     [("Power Equipment Operators", "Group 2", False)] * 4 + [("Laborers", "General Laborer", True)] * 2, (1, 2, 3)),
    ("Quarvelle Glass & Glazing", 1, "Norvanta Builders Inc",
     [("Glaziers", "Journey Level", False)] * 4, (2, 3, 4)),
    ("Kelbranne Drywall", 1, "Norvanta Builders Inc",
     [("Drywall Applicators", "Journey Level", False)] * 6, (1, 2, 3, 4)),
    ("Yarrowind Landscaping", 1, "Norvanta Builders Inc", [], ()),          # not started: no payroll yet
]

FIRST = ["Marisol", "Dmitri", "Aaliyah", "Tomasz", "Keoni", "Priya", "Jamal", "Ingrid", "Mateo", "Hana", "Darnell",
         "Sofia", "Ravi", "Lucia", "Ethan", "Amara", "Oskar", "Yesenia", "Tariq", "Bethany", "Kwame", "Noor", "Callum",
         "Ximena", "Hector", "Leilani", "Viktor", "Grace", "Rafael", "Mei", "Andre", "Fatima", "Colin", "Paola", "Jin",
         "Rosalind", "Emeka", "Ana", "Brody", "Svetlana", "Tevita", "Imani", "Gustavo", "Nadia", "Wesley", "Ayumi",
         "Desmond", "Lorena", "Mikhail", "Talia", "Rohan", "Isabel", "Cormac", "Esperanza", "Kofi", "Annika",
         "Joaquin", "Shauna", "Bao", "Felipe", "Marguerite", "Dante", "Yusuf", "Elena", "Tobias", "Carmen", "Niko",
         "Renata", "Malik", "Astrid", "Diego", "Harriet", "Anil", "Juniper", "Luka", "Rosa", "Kai", "Zainab",
         "Bartholomew", "Inez", "Quentin", "Maeve", "Arjun", "Corinne", "Salvador", "Thandiwe", "Emil", "Odessa"]
LAST = ["Okafor", "Lindqvist", "Tran", "Castellano", "Brightwell", "Haddad", "Novak", "Kealoha", "Mendoza", "Sato",
        "Abernathy", "Petrov", "Achterberg", "Villanueva", "Oyelaran", "Fairweather", "Kowalczyk", "Delacroix",
        "Nakashima", "Ferreira", "Ibarra", "Holloway", "Szabo", "Quintero", "Mbeki", "Larkin", "Vasquez", "Ostrowski",
        "Thorsen", "Aguilar", "Whitcombe", "Banerjee", "Calloway", "Duarte", "Eriksen", "Galloway", "Hargrove",
        "Iwasaki", "Jaramillo", "Kristoffersen", "Lombardi", "Moreau", "Nwosu", "Ortega", "Pemberton", "Rasmussen",
        "Salazar", "Takahashi", "Umberger", "Valdivia", "Wojcik", "Yoon", "Zamora", "Alvarado", "Bergstrom",
        "Chaudhry", "Dunleavy", "Esposito", "Fonseca", "Gallardo", "Hoang", "Ishikawa", "Jovanovic", "Kerrigan",
        "Lindgren", "Montoya", "Nordahl", "Obi", "Pacheco", "Quiroga", "Rinaldi", "Sandoval", "Tamura", "Underhill",
        "Vance", "Whitaker", "Ximenez", "Yamada", "Zeller", "Arcand", "Bustamante", "Corrigan", "Dalca", "Elsworth"]


# ----------------------------------------------------------------------------------------------- the crew
class Crew:
    """Invented workers: unique names and IDs (the 7320nnn range), the same worker in every week."""

    def __init__(self, rng: random.Random):
        self.rng, self.used, self.next_id, self.workers = rng, set(), ID_BASE, []

    def hire(self, contractor, trade, cls, union):
        while True:
            name = f"{self.rng.choice(LAST)}, {self.rng.choice(FIRST)}"
            if name not in self.used:
                break
        self.used.add(name)
        self.next_id += 1
        w = {"name": name, "id": self.next_id, "contractor": contractor, "craft": trade, "cls": cls, "union": union}
        self.workers.append(w)
        return w


def rates(project, w):
    table = UNION_RATES if w["union"] else STATE_RATES
    return table[(w["craft"], w["cls"])]


def pay_line(project, w, week, *, st=(8, 8, 8, 8, 8, 0, 0), ot=(0,) * 7, dt_=(0,) * 7, base=None, tp=None, otr=None,
             dtr=None, cls=None, craft=None, **override):
    """One consistent payroll line (SV.line: days add up, wages = hours x rates) for an invented worker."""
    b, f = rates(project, w)
    jur = JUR_UNION if w["union"] else jur_state(project["county"])
    row = SV.line(week=week, craft=craft or w["craft"], cls=cls or w["cls"], st=st, ot=ot, dt_=dt_,
                  base=b if base is None else base, tp=f if tp is None else tp, otr=otr, dtr=dtr, jur=jur,
                  **{"Project Name": project["full_name"], "Contractor Name": w["contractor"],
                     "Employee Name": w["name"], "Employee ID (LCPtracker)": w["id"]}, **override)
    return row


def normal_hours(rng):
    """A plausible clean week: mostly 8 x 5, sometimes a short week or a little overtime paid right."""
    r = rng.random()
    if r < 0.62:
        return {"st": (8, 8, 8, 8, 8, 0, 0)}
    if r < 0.78:
        return {"st": (8, 8, 8, 8, 0, 0, 0)}
    if r < 0.88:
        return {"st": (8, 8, 8, 8, 8, 0, 0), "ot": (0, 2, 0, 0, 0, 0, 0)}
    if r < 0.95:
        return {"st": (0, 8, 8, 8, 8, 0, 0)}
    return {"st": (8, 8, 8, 8, 6, 0, 0)}


# ------------------------------------------------------------------------------------ seeded payroll findings
# Each: (week number, contractor, the crew member's position in that contractor's crew, case, area, must, may, how)
# "how": keyword arguments for pay_line that make the deliberate error (the same errors as the harness's seeded
# cases, tests/test_seeded_violations.py). "dup": the line appears twice (S09). Every other line is correct.
SMALL_SEEDS = [
    (2, "Pellavine Electric LLC", 1, "Base rate below the prevailing base", "rates", {"R08"}, {"R07"},
     {"base": 57.0}),
    (2, "Quillmere Construction Inc", 0, "Overtime rate below 1.5x base", "overtime", {"R03"}, {"S12"},
     {"ot": (2, 0, 0, 0, 0, 0, 0), "otr": 70.0}),
    (2, "Fenrowe Taping Co", 2, "45 straight-time hours in the week", "overtime", {"S05"}, {"S01"},
     {"st": (9, 9, 9, 9, 9, 0, 0)}),
]
LARGE_SEEDS = [
    # week 1
    (1, "Kelbranne Drywall", 2, "Base rate below the prevailing base", "rates", {"R08"}, {"R07"}, {"base": 45.0}),
    (1, "Marlowick Painting", 1, "Fringe benefits short", "fringes", {"R07"}, {"R08"}, {"tp": 11.0}),
    (1, "Holtavane Mechanical LLC", 3, "9 straight-time hours in a day", "overtime", {"S01"}, set(),
     {"st": (9, 8, 8, 8, 0, 0, 0)}),
    (1, "Corvanne Electric Inc", 4, "Duplicate line", "data", {"S09"}, {"S05"}, {"dup": True}),
    # week 2
    (2, "Dravello Concrete LLC", 2, "11 straight-time hours in a day", "overtime", {"R01"}, set(),
     {"st": (11, 8, 8, 8, 0, 0, 0)}),
    (2, "Norvanta Builders Inc", 7, "Union: Saturday paid straight time", "overtime", {"R05"}, set(),
     {"st": (8, 8, 8, 8, 0, 8, 0)}),
    (2, "Dravello Concrete LLC", 7, "Union: Sunday not paid double time", "overtime", {"R06"}, set(),
     {"st": (8, 8, 8, 8, 0, 0, 8)}),
    (2, "Fennacre Excavation", 1, "Classification not in the wage data", "classification", {"S02"}, set(),
     {"cls": "Grade Checker"}),
    (2, "Tessivar Steel Erectors", 3, "Total hours don't add up", "data", {"S03"}, set(), {"Total Hours": 41}),
    # week 3
    (3, "Quarvelle Glass & Glazing", 0, "45 straight-time hours in the week", "overtime", {"S05"}, {"S01"},
     {"st": (9, 9, 9, 9, 9, 0, 0)}),
    (3, "Corvanne Electric Inc", 5, "Classification found only under another trade", "classification", {"S11"},
     {"S02", "R08", "R07"}, {"cls": "General Laborer"}),
    (3, "Rindlemoor Rebar Co", 1, "13 hours in a day", "overtime", {"R02"}, {"S01"},
     {"st": (10, 8, 8, 8, 0, 0, 0), "ot": (3, 0, 0, 0, 0, 0, 0)}),
    (3, "Kelbranne Drywall", 4, "Grand total doesn't add up", "data", {"S04"}, set(), {"Grand Total Wages": 1500.0}),
    (3, "Marlowick Painting", 3, "Base rate missing", "data", {"S08"}, {"R08", "R07", "R03", "S04", "S06"},
     {"Base Hourly": None}),
    # week 4 (the file added in the video)
    (4, "Norvanta Builders Inc", 8, "Double-time rate below 2x base", "overtime", {"R04"}, set(),
     {"dt_": (0, 0, 0, 0, 0, 0, 4), "dtr": 70.0}),
    (4, "Holtavane Mechanical LLC", 0, "Overtime rate below 1.5x base", "overtime", {"R03"}, {"S12"},
     {"ot": (2, 0, 0, 0, 0, 0, 0), "otr": 80.0}),
    (4, "Corvanne Electric Inc", 2, "Hours typed as a word", "values", {"S15"}, set(), {"ST Mon": "eight"}),
]


def build_payrolls(project, contractors, seeds, weeks, crew, rng):
    """{week number: [(case name, [(row, expectation)])]}: every contractor's crew for each week it worked; the seeded
    lines carry their must/may, every other line is a clean control."""
    staff = {}
    for name, _tier, _hired_by, members, _w in contractors:
        staff[name] = [crew.hire(name, t, c, u) for t, c, u in members]
    seeded = {(wk, co, i): (case, area, must, may, how) for wk, co, i, case, area, must, may, how in seeds}
    out = {}
    for n, week in weeks.items():
        lines = []
        for name, _tier, _hired_by, _members, on in contractors:
            if n not in on:
                continue
            for i, w in enumerate(staff[name]):
                s = seeded.get((n, name, i))
                if s is None:
                    lines.append((f"Correct line: {name}", [(pay_line(project, w, week, **normal_hours(rng)), {
                        "case": "Correct line (control)", "area": "control", "group": f"w{n} {w['id']}",
                        "must": [], "may": [], "notes_contain": [], "notes_absent": []})]))
                    continue
                case, area, must, may, how = s
                how = dict(how)
                dup = how.pop("dup", False)
                row = pay_line(project, w, week, **how)
                exp = {"case": case, "area": area, "group": f"{case} (w{n})", "must": sorted(must),
                       "may": sorted(may), "notes_contain": [], "notes_absent": []}
                lines.append((case, [(row, exp)] + ([(list(row), dict(exp))] if dup else [])))
        out[n] = lines
    used = {(wk, co, i) for wk, co, i, *_ in seeds}
    for wk, co, i in used:                                     # a seed must land on a worker who exists that week
        crew_of = dict((c[0], c) for c in contractors)[co]
        if wk not in crew_of[4] or i >= len(crew_of[3]):
            raise SystemExit(f"seed for {co} week {wk} member {i} lands on nobody")
    return out


# ------------------------------------------------------------------------------------------ trackers and L&I
TRACKER_HEADER = ["Status", "Contractor", "Tier", "Hiring Contractor", "Scope of Work", "Intent Filed", "Intent ID",
                  "Affidavit Filed", "Affidavit ID", "Affidavit Good", "Start Date", "End Date"]
ROLE_COLUMNS = {h: LNI.ROLE_OF[h] for h in TRACKER_HEADER if h in LNI.ROLE_OF}


def tline(co, tier, hiring, scope, *, status="Onsite", filed="Yes", iid=None, aff_filed="No", aid=None, good=None,
          start="2026-09-14", end=None):
    def d(x):
        return dt.datetime.fromisoformat(x) if isinstance(x, str) else (
            dt.datetime(x.year, x.month, x.day) if isinstance(x, dt.date) else x)
    return [status, co, {"prime": "Prime", 1: "1st tier", 2: "2nd tier"}[tier], hiring, scope, filed, iid, aff_filed,
            aid, good, d(start), d(end)]


def small_tracker(p):
    P, prime = p["lni_project"], p["prime"]
    I, A = LNI.intent, LNI.affidavit
    soon = dt.date.today() + dt.timedelta(days=60)
    return [
        dict(case="Prime contractor, intent approved (control)", area="control",
             line=tline(prime, "prime", "N/A - Prime", "General contractor", iid=8820001),
             intents=[I(8820001, prime, project=P, hiring=None)]),
        dict(case="Intent ID doesn't match L&I (a typo: L&I has 8820012)", area="intent",
             line=tline("Pellavine Electric LLC", 1, prime, "Electrical", iid=8820021),
             intents=[I(8820012, "Pellavine Electric LLC", project=P, hiring=prime)],
             must=[("intent_id", "intent_not_found")]),
        dict(case="Affidavit at L&I, missing from the tracker (Affidavit Good ticked: ignored)", area="affidavit",
             line=tline("Ormsgate Plumbing & Heating", 1, prime, "Plumbing", status="Complete", iid=8820013,
                        good="Yes", end="2026-11-08"),
             intents=[I(8820013, "Ormsgate Plumbing & Heating", project=P, hiring=prime, affidavits=(8830013,))],
             affidavits=[A(8830013, 8820013, "Ormsgate Plumbing & Heating", project=P, received="2026-11-12")],
             must=[("affidavit_id", "affidavit_missing")], ids={("affidavit_id", "affidavit_missing"): [8830013, 8820013]}),
        dict(case="Finished, no affidavit at L&I yet (not due: counted, not marked)", area="not due",
             line=tline("Kelvarro Drywall Inc", 1, prime, "Drywall", status="Complete", iid=8820014,
                        end="2026-11-08"),
             intents=[I(8820014, "Kelvarro Drywall Inc", project=P, hiring=prime)], not_due="no_affidavit_yet"),
        dict(case="Hiring contractor doesn't match L&I (L&I: hired by Kelvarro Drywall Inc)", area="hiring",
             line=tline("Fenrowe Taping Co", 2, prime, "Taping", iid=8820015, start="2026-11-02"),
             intents=[I(8820015, "Fenrowe Taping Co", project=P, hiring="Kelvarro Drywall Inc")],
             must=[("hiring", "hiring_mismatch")]),
        dict(case="On site, no intent anywhere (intent missing)", area="intent",
             line=tline("Tamsvik Painting", 2, "Kelvarro Drywall Inc", "Painting", filed="No", start="2026-11-09"),
             must=[("intent_id", "started_without_intent")]),
        dict(case="Not started yet, no intent (not due: counted, not marked)", area="not due",
             line=tline("Ashcombe Millwork", 1, prime, "Casework", status="Not started", filed="No",
                        start=soon), not_due="not_started"),
    ]


def large_tracker(p):
    P, prime = p["lni_project"], p["prime"]
    I, A = LNI.intent, LNI.affidavit
    soon = dt.date.today() + dt.timedelta(days=45)
    return [
        dict(case="Prime contractor, intent approved (control)", area="control",
             line=tline(prime, "prime", "N/A - Prime", "General contractor", start="2026-09-01", iid=8820101),
             intents=[I(8820101, prime, project=P, hiring=None)]),
        dict(case="Subcontractor on site, intent approved (control)", area="control",
             line=tline("Dravello Concrete LLC", 1, prime, "Concrete", iid=8820102, start="2026-09-08"),
             intents=[I(8820102, "Dravello Concrete LLC", project=P, hiring=prime)]),
        dict(case="2nd-tier sub: hiring contractor doesn't match L&I (L&I: hired by Dravello Concrete LLC)",
             area="hiring", line=tline("Rindlemoor Rebar Co", 2, prime, "Rebar", iid=8820103, start="2026-09-15"),
             intents=[I(8820103, "Rindlemoor Rebar Co", project=P, hiring="Dravello Concrete LLC")],
             must=[("hiring", "hiring_mismatch")]),
        dict(case="Finished, affidavit on file, Affidavit Good ticked (control)", area="control",
             line=tline("Tessivar Steel Erectors", 1, prime, "Structural steel", status="Complete", iid=8820104,
                        aff_filed="Yes", aid=8830104, good="Yes", start="2026-09-15", end="2026-11-01"),
             intents=[I(8820104, "Tessivar Steel Erectors", project=P, hiring=prime, affidavits=(8830104,))],
             affidavits=[A(8830104, 8820104, "Tessivar Steel Erectors", project=P, received="2026-11-05")]),
        dict(case="Intent ID doesn't match L&I (a typo: L&I has 8820105)", area="intent",
             line=tline("Corvanne Electric Inc", 1, prime, "Electrical", iid=8820150, start="2026-09-22"),
             intents=[I(8820105, "Corvanne Electric Inc", project=P, hiring=prime)],
             must=[("intent_id", "intent_not_found")]),
        dict(case="Subcontractor on site, intent approved (control)", area="control",
             line=tline("Holtavane Mechanical LLC", 1, prime, "Plumbing and HVAC", iid=8820106, start="2026-09-22"),
             intents=[I(8820106, "Holtavane Mechanical LLC", project=P, hiring=prime)]),
        dict(case="Finished, no affidavit at L&I yet (not due: counted, not marked)", area="not due",
             line=tline("Velkorin Insulation", 2, "Holtavane Mechanical LLC", "Insulation", status="Complete",
                        iid=8820107, start="2026-10-06", end="2026-11-01"),
             intents=[I(8820107, "Velkorin Insulation", project=P, hiring="Holtavane Mechanical LLC")],
             not_due="no_affidavit_yet"),
        dict(case="Affidavit at L&I, missing from the tracker", area="affidavit",
             line=tline("Zubrina Roofing Co", 1, prime, "Roofing", status="Complete", iid=8820108,
                        start="2026-10-06", end="2026-11-01"),
             intents=[I(8820108, "Zubrina Roofing Co", project=P, hiring=prime, affidavits=(8830108,))],
             affidavits=[A(8830108, 8820108, "Zubrina Roofing Co", project=P, received="2026-11-10")],
             must=[("affidavit_id", "affidavit_missing")], ids={("affidavit_id", "affidavit_missing"): [8830108, 8820108]}),
        dict(case="On site, no intent anywhere (intent missing)", area="intent",
             line=tline("Marlowick Painting", 1, prime, "Painting", filed="No", start="2026-10-13"),
             must=[("intent_id", "started_without_intent")]),
        dict(case="A newer affidavit at L&I than the tracker's (Affidavit Good says No: ignored)", area="affidavit",
             line=tline("Fennacre Excavation", 1, prime, "Excavation", status="Complete", iid=8820110,
                        aff_filed="Yes", aid=8830110, good="No", start="2026-09-01", end="2026-11-08"),
             intents=[I(8820110, "Fennacre Excavation", project=P, hiring=prime, affidavits=(8830110, 8830111))],
             affidavits=[A(8830110, 8820110, "Fennacre Excavation", project=P, received="2026-10-20"),
                         A(8830111, 8820110, "Fennacre Excavation", project=P, received="2026-11-12")],
             must=[("affidavit_id", "affidavit_newer_at_lni")],
             ids={("affidavit_id", "affidavit_newer_at_lni"): [8830111, 8820110]}),
        dict(case="L&I approved the intent; the tracker says not filed", area="intent",
             line=tline("Quarvelle Glass & Glazing", 1, prime, "Glazing", filed="No", iid=8820111,
                        start="2026-10-27"),
             intents=[I(8820111, "Quarvelle Glass & Glazing", project=P, hiring=prime)],
             must=[("intent_flag", "approved_but_tracker_says_no")]),
        dict(case="No intent in the tracker; L&I has one for the contractor", area="intent",
             line=tline("Kelbranne Drywall", 1, prime, "Drywall", filed="No", start="2026-10-13"),
             intents=[I(8820112, "Kelbranne Drywall", project=P, hiring=prime)],
             must=[("intent_id", "intent_found_by_name")], ids={("intent_id", "intent_found_by_name"): [8820112]}),
        dict(case="Not started yet, no intent (not due: counted, not marked)", area="not due",
             line=tline("Yarrowind Landscaping", 1, prime, "Landscaping", status="Not started", filed="No",
                        start=soon), not_due="not_started"),
    ]


LARGE_LNI_ONLY = [("Brokkvale Traffic Control", 8820130, "Approved"), ("Sendrapol Demolition", 8820131, "Withdrawn")]


def write_tracker(path, project, cases):
    """The tracker as staff keep it: the contractor table on the first tab, a Project Info tab (what the check reads
    for the L&I project), and the firm's Metadata tab (project_metadata/Metadata.xlsx's layout, invented values)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compliance Tracker"
    ws.append(TRACKER_HEADER)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFD9E1F2")
    for c in cases:
        ws.append(c["line"])
    for col, width in zip("ABCDEFGHIJKL", (12, 30, 9, 28, 20, 11, 11, 13, 12, 13, 12, 12)):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        for c in row[10:12]:
            c.number_format = "mm/dd/yyyy"
    pi = wb.create_sheet("Project Info")
    for r in (["Project", project["full_name"]], ["General Contractor", project["prime"]],
              ["L&I Project ID", project["lni_project"]], ["Prime Intent", project["prime_intent"]],
              ["County", project["county"]], ["Notes", "Invented project for the capstone demo"]):
        pi.append(r)
    md = wb.create_sheet("Metadata")
    tpl = next((f for f in (REPO / "project_metadata" / "Metadata.xlsx",
                            AGENT.resolve().parents[1] / "project_metadata" / "Metadata.xlsx") if f.is_file()), None)
    template = openpyxl.load_workbook(tpl).active if tpl else None
    values = {"project_name": project["folder"], "project_full_name": project["full_name"],
              "prime_contractor": project["prime"], "awarding_agency": project["owner"],
              "contract_number": project["contract"], "lni_project_id": project["lni_project"],
              "project_address": project["address"], "state": "WA", "county": project["county"],
              "project_type": "Commercial", "status": "Active", "state_rates_date": "2026-06-01",
              "federal_wage_determination": "N/A", "federal_wd_mod": "N/A", "payroll_submission": "LCPtracker",
              "payroll_frequency": "Weekly", "intent_applicable": "Yes", "prime_intent": project["prime_intent"],
              "doc_intent": "Yes", "doc_fbs": "Yes", "doc_apprentice": "No",
              "cva_comments": "Invented project for the capstone demo", "metadata_updated": "2026-11-16"}
    if template is not None:
        for row in template.iter_rows(values_only=True):
            row = list(row)
            if len(row) > 2 and row[1] in values:
                row[2] = values[row[1]]
            md.append(row)
    else:                                                   # the public repo without the template: the key fields
        md.append(["Project Data"])
        md.append(["Field", "Key", "Value"])
        for k, v in values.items():
            md.append([k.replace("_", " ").capitalize(), k, v])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def tracker_expectation(cases):
    out = {}
    for i, c in enumerate(cases):
        out[str(i + 2)] = {"case": c["case"], "area": c["area"],
                           "must": sorted(f"{r}:{k}" for r, k in c.get("must", [])),
                           "ids": {f"{r}:{k}": [str(x) for x in v] for (r, k), v in (c.get("ids") or {}).items()},
                           "not_due": c.get("not_due"), "staff_note": None}
    return out


def lni_tree(project, intents, affidavits):
    """GetProjectStructureTree's answer for one L&I project (eval_lni.records' shape)."""
    P, prime_id = project["lni_project"], project["prime_intent"]
    mine = [i for i, r in intents.items() if r["ProjId"] == P]
    affs = [a for a, r in affidavits.items() if r["ProjId"] == P]
    return [{"ContractorName": project["prime"], "IsPrimeContractor": True,
             "Intents": [{"IntentId": prime_id}] if prime_id in mine else [],
             "Affidavits": [{"AffidavitId": a} for a in affs if affidavits[a]["IntentId"] == prime_id],
             "SubContractorTreeNodes": [{"ContractorName": intents[i]["CoName"], "Intents": [{"IntentId": i}],
                                         "Affidavits": [{"AffidavitId": a} for a in affs
                                                        if affidavits[a]["IntentId"] == i]}
                                        for i in mine if i != prime_id]}]


# ---------------------------------------------------------------------------------------------- settings
def demo_config(dest: Path):
    """The repository's config.yaml with the demo's settings: no per-project entries (they name real projects), L&I's
    live lookup at its production address in plain http (the local stand-in, as the only proxy, answers it), no
    SAM.gov, no local model, no error reporting, the structured run log on, the AI review as --ai chooses (set by
    run_demo.sh: demo/ai_mode.py)."""
    cfg = yaml.safe_load((AGENT / "config.yaml").read_text()) or {}
    cfg["root"] = "."
    cfg["projects"] = {}
    cfg.setdefault("wages", {}).setdefault("sam_gov", {})["enabled"] = False
    cfg.setdefault("llm", {})["enabled"] = False
    cfg.setdefault("error_reporting", {})["sentry_dsn"] = ""
    cfg.setdefault("run_log", {})["events"] = True
    it = cfg.setdefault("intent", {})
    it["lni_source"] = "live"
    it.setdefault("live", {}).update({"url": LNI.LNI_URL, "pause_seconds": 0})
    ai = cfg.setdefault("ai_review", {})
    ai.update({"enabled": True, "mode": "local", "shadow": False, "download_regulations": False,
               "knowledge_dir": "reference/knowledge", "cache_dir": "./.home/ai_review"})
    (dest / "config.yaml").write_text(
        "# The CVA demo's settings (made by demo/make_demo_data.py from the repository's config.yaml).\n"
        "# Invented data only. demo/run_demo.sh sets the AI review's mode (--ai evidence | claude).\n"
        + yaml.safe_dump(cfg, sort_keys=False))
    (dest / "notifications.config").write_text(
        "# The demo's outcome emails: saved as .eml previews in run_logs/ (run_demo.sh sets COMPLIANCE_NOTIFY=dry-run),\n"
        "# never sent. Invented addresses on the reserved .invalid domain.\n"
        "to: [compliance-staff@cva-demo.invalid]\n"
        "admin_to: [cva-admin@cva-demo.invalid]\n"
        "cc: []\n"
        "from: \"Compliance Validation Agent <cva@cva-demo.invalid>\"\n"
        "links: {google_drive: false, local_paths: true}\n"
        "attach_logs: problems\n"
        "transport: ses        # never used: COMPLIANCE_NOTIFY=dry-run saves a preview instead of sending\n")


KNOWLEDGE = {
    # Short, plain summaries written for the demo (not copies of the regulations), so the AI review has evidence to
    # retrieve offline. Each says where the rule comes from; the CVA's own run downloads the official texts.
    "regulations/state/overtime_daily_and_weekly.md":
        "# Overtime on Washington public works (demo summary)\n\n"
        "Source: RCW 39.12 and WAC 296-127 (summary written for the demo; read the official text).\n\n"
        "Prevailing wage rates on public works include overtime provisions. Where the applicable rates say overtime "
        "is paid after 8 hours in a day or 40 hours in a week, straight-time hours above 8 in a day or 40 in a week "
        "should have been paid at the overtime rate, at least 1.5 times the base rate. A payroll showing more than "
        "40 straight-time hours in a week, or more than 8 in a day, needs checking against the trade's overtime "
        "provisions; some agreements allow a 4 x 10 schedule, which must be documented.\n",
    "regulations/state/classifications.md":
        "# Using the right classification (demo summary)\n\n"
        "Source: WAC 296-127 (summary written for the demo).\n\n"
        "Each worker must be paid at least the prevailing rate for the work actually performed, in the trade and "
        "classification that describes it. When a payroll names a classification that the wage determination lists "
        "only under another trade, the classification may be wrong or the work may belong to another trade: check "
        "the scope of work and the contractor's statement before accepting it.\n",
    "guidance/staff/review_practice.md":
        "# Staff practice notes for payroll review (invented for the demo)\n\n"
        "Orange flags are questions, not violations. For hours above 8 in a day, ask whether the contractor has an "
        "approved alternative schedule on file. For an approximate classification match, ask the contractor which "
        "trade the work belongs to and keep the answer in the Review decisions sheet.\n",
}


# ---------------------------------------------------------------------------------------------- build
def build(dest: Path) -> dict:
    rng = random.Random(SEED)
    crew = Crew(rng)
    dest.mkdir(parents=True, exist_ok=True)
    (dest / f".{MARKER}").write_text("Made by demo/make_demo_data.py (invented data for the CVA capstone demo).\n")
    demo_config(dest)
    for rel, text in KNOWLEDGE.items():
        f = dest / "reference" / "knowledge" / rel
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(text)
    key = {"version": EC.KEY_VERSION, "files": {}, "trackers": {}, "known_gaps": []}
    intents, affidavits = {}, {}

    def wages(p, project, union):
        rows = [[f"{project['county']} County", t, c, round(b + f, 2), b, f, "State", STATE_DECISION,
                 D(2026, 6, 1), D(2026, 6, 1), None] for (t, c), (b, f) in STATE_RATES.items()]
        if union:
            rows += [[f"{project['county']} County", t, c, round(b + f, 2), b, f, "Union", UNION_DECISION,
                      D(2026, 6, 1), D(2026, 6, 1), None] for (t, c), (b, f) in UNION_RATES.items()]
        EC.wage_xlsx(p.upload / f"Wage rates - {project['full_name'].split(' (')[0]}.xlsx", rows=rows)

    def payroll(p, project, fname, lines, target=None, footer=()):
        path = (target or p.upload) / fname
        exp, head = EC.write_payroll(path, lines, footer=footer)
        key["files"][f"{project['folder']}/{fname}"] = {
            "outcome": "reviewed", "header_row": head, "reason_contains": "", "reason_code": "",
            "lines": {str(r): e for r, e in exp.items()}}

    def tracker(p, project, cases, lni_only=()):
        write_tracker(p.tracker / f"{project['full_name'].split(' (')[0]} - Compliance Tracker.xlsx", project, cases)
        for c in cases:
            for r in c.get("intents", []) + c.get("affidavits", []):
                r["CoUbiNum"] = f"999{r.get('IntentId', 0) % 1000000:06d}"   # invented UBIs: 999nnnnnn, no real range
            for r in c.get("intents", []):
                intents[r["IntentId"]] = r
            for r in c.get("affidavits", []):
                affidavits[r["AfdvtId"]] = r
        for co, iid, status in lni_only:
            intents[iid] = LNI.intent(iid, co, project=project["lni_project"], hiring=project["prime"], status=status)
            intents[iid]["CoUbiNum"] = f"999{iid % 1000000:06d}"
        due = {}
        for c in cases:
            if c.get("not_due"):
                due[c["not_due"]] = due.get(c["not_due"], 0) + 1
        lines = tracker_expectation(cases)
        marks = [m for e in lines.values() for m in e["must"]]
        listed = sorted(co for co, _i, status in lni_only if status.lower() not in ("withdrawn",))
        key["trackers"][project["folder"]] = {
            "file": f"{project['full_name'].split(' (')[0]} - Compliance Tracker.xlsx", "outcome": "checked",
            "reason_contains": "", "header_row": 1, "lines": lines, "lni_only": listed, "not_due": due,
            "never_ask": [], "summary": {"marked": sum(1 for m in marks if not m.startswith("hiring:")),
                                         "hiring": sum(1 for m in marks if m.startswith("hiring:")),
                                         "lni_only": len(listed)}}

    # --- the small project: two weekly payrolls, a few clear findings
    p = Paths(dest / "Projects" / SMALL["folder"])
    ensure_layout(p.project)
    wages(p, SMALL, union=False)
    weeks = {1: dt.date(2026, 11, 8), 2: dt.date(2026, 11, 15)}
    lines = build_payrolls(SMALL, SMALL_CONTRACTORS, SMALL_SEEDS, weeks, crew, rng)
    for n, week in weeks.items():
        payroll(p, SMALL, f"Brightfen PayrollDetailReport WE {week:%m-%d-%Y}.xlsx", lines[n])
    tracker(p, SMALL, small_tracker(SMALL))

    # --- the large project: four weeks (the last one to add in the video), a skipped upload, a union trade
    p = Paths(dest / "Projects" / LARGE["folder"])
    ensure_layout(p.project)
    wages(p, LARGE, union=True)
    weeks = {1: dt.date(2026, 10, 25), 2: dt.date(2026, 11, 1), 3: dt.date(2026, 11, 8), 4: dt.date(2026, 11, 15)}
    lines = build_payrolls(LARGE, LARGE_CONTRACTORS, LARGE_SEEDS, weeks, crew, rng)
    for n, week in weeks.items():
        fname = f"Kestrel Point PayrollDetailReport WE {week:%m-%d-%Y}.xlsx"
        payroll(p, LARGE, fname, lines[n], target=(dest / "To upload") if n == 4 else None)
    # a resubmitted week with two exports pasted into one sheet: skipped, with its plain reason
    two = lines[3][:3] + [("(the header again)", [(list(SV.HEADER), None)])] + lines[3][3:6]
    fname = "Kestrel Point PayrollDetailReport WE 11-08-2026 (resubmitted).xlsx"
    exp, head = EC.write_payroll(p.upload / fname, two)
    key["files"][f"{LARGE['folder']}/{fname}"] = {
        "outcome": "skipped", "header_row": head, "reason_contains": "header row appears again",
        "reason_code": "repeated_header", "lines": {}}
    tracker(p, LARGE, large_tracker(LARGE), LARGE_LNI_ONLY)

    # --- L&I's invented records, for the stand-in
    trees = {str(pr["lni_project"]): lni_tree(pr, intents, affidavits) for pr in (SMALL, LARGE)}
    data = dest / "demo_data"
    data.mkdir(exist_ok=True)
    (data / "lni_records.json").write_text(json.dumps(
        {"intents": {str(k): v for k, v in intents.items()}, "affidavits": {str(k): v for k, v in affidavits.items()},
         "trees": trees}, indent=1, default=str))
    (data / "answer_key.json").write_text(json.dumps(key, indent=1, default=str))
    (data / "workers.json").write_text(json.dumps(crew.workers, indent=1))
    (dest / "To upload").mkdir(exist_ok=True)
    return key


def reset(dest: Path):
    """Deletes a folder this script made (it has the marker file), and only such a folder."""
    if not dest.exists():
        return
    if not (dest / f".{MARKER}").is_file():
        raise SystemExit(f"{dest} wasn't made by this script (no .{MARKER} file): not deleting it")
    shutil.rmtree(dest)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Build the CVA demo's invented projects.")
    ap.add_argument("dest", type=Path, help="the demo folder (empty, or made by this script with --reset)")
    ap.add_argument("--reset", action="store_true", help="delete the demo folder's earlier contents first")
    a = ap.parse_args(argv)
    dest = a.dest.expanduser().resolve()
    if a.reset:
        reset(dest)
    if dest.exists() and any(dest.iterdir()):
        raise SystemExit(f"{dest} isn't empty (use --reset to rebuild a demo folder)")
    key = build(dest)
    n = sum(len(f["lines"]) for f in key["files"].values())
    print(f"demo built in {dest}: {len(key['files'])} payroll files ({n} lines), {len(key['trackers'])} trackers")
    return 0


if __name__ == "__main__":
    sys.exit(main())
