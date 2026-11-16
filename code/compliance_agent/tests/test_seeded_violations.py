"""Acceptance test: a payroll with known, seeded violations and clean control lines.

Each case is one deliberate error (or a correct line that must stay clean), written from the rule's meaning, with
the flags it must raise ("must") and the related flags it may also raise ("may", e.g. a short fringe also leaves
base + fringes below the total). The test measures, per rule:
  detection rate      seeded errors whose "must" flags were all raised
  false positives     flags raised on a case outside its "must" and "may" sets (clean controls: any flag)
and fails when any seeded error is missed or any false positive appears. Run it on its own for the table:

    python tests/test_seeded_violations.py          (prints the acceptance table, Markdown)

Every name, number and company here is invented; employees are tokens, as on AWS.
"""
import datetime as dt
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
sys.path.insert(0, str(AGENT))

from agent.config import load_config                                        # noqa: E402
from agent.payroll import RULE_TITLES, map_columns, review_rows             # noqa: E402
from agent.wages import WageBook                                            # noqa: E402

CFG = load_config(AGENT / "config.yaml", AGENT.parent)
DAY_NAMES = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
HEADER = (["Project Name", "Contractor Name", "Jurisdiction", "Week End Date", "Employee Name",
           "Employee ID (LCPtracker)", "Craft", "Classification"]
          + [f"{p} {d}" for p in ("ST", "OT", "DT") for d in DAY_NAMES]
          + ["ST Hours", "OT Hours", "DT Hours", "Total Hours", "Base Hourly", "Overtime Hourly",
             "Doubletime Hourly", "Cash Fringe Rate", "3rd Party Fringe Rate", "ST Prevailing Wage Rate Total",
             "Total ST Wages", "Total OT Wages", "Total DT Wages", "Grand Total Wages"])
COLS = map_columns(HEADER)
STATE = "Washington State Wages - Pierce County"
UNION = "Union rates - Laborers Local 901 (sample)"
WEEK = dt.date(2026, 9, 6)


def rate(kind, craft, cls, basic, fringe, issued=dt.date(2026, 1, 1), expire=None, decision=None):
    return {"kind": kind, "craft": craft, "classification": cls,
            "wage_decision": decision or {"state": "WA L&I 2026", "union": "Laborers CBA 2026"}[kind],
            "county": None, "issue_date": issued, "expire_date": expire, "mod": 0, "basic": basic, "fringe": fringe,
            "total": basic + fringe, "ot_total": 1.5 * basic + fringe, "dt_total": 2 * basic + fringe,
            "std_hours": 8, "vac_pto": 0, "source_type": "lni_csv"}


# The project's wage data: state rates for two trades, a state rate that rises on Sep 1, a union rate
BOOK = WageBook([
    rate("state", "Laborers", "General Laborer", 40.0, 15.0),
    rate("state", "Electricians", "Journey Level", 60.0, 25.0),
    rate("state", "Cement Masons", "Journey Level", 45.0, 18.0, expire=dt.date(2026, 8, 31)),
    rate("state", "Cement Masons", "Journey Level", 48.0, 18.0, issued=dt.date(2026, 9, 1)),
    rate("union", "Laborers", "General Laborer", 42.0, 20.0),
])


def line(week=WEEK, craft="Laborers", cls="General Laborer", st=(8, 8, 8, 8, 8, 0, 0), ot=(0,) * 7, dt_=(0,) * 7,
         base=40.0, otr=None, dtr=None, cash=0.0, tp=15.0, emp="EMP-0a1b2c3d4e", jur=STATE, **override):
    """One payroll line, internally consistent (days add up, wages = hours x rates) unless overridden."""
    otr = 1.5 * base if otr is None else otr
    dtr = 2 * base if dtr is None else dtr
    s, o, d = sum(st), sum(ot), sum(dt_)
    stw, otw, dtw = round(s * base, 2), round(o * otr, 2), round(d * dtr, 2)
    v = {"Project Name": "Test", "Contractor Name": "Acme Paving", "Jurisdiction": jur, "Week End Date": week,
         "Employee Name": emp, "Employee ID (LCPtracker)": "EID-0a1b2c3d4e", "Craft": craft, "Classification": cls,
         **{f"ST {n}": h for n, h in zip(DAY_NAMES, st)}, **{f"OT {n}": h for n, h in zip(DAY_NAMES, ot)},
         **{f"DT {n}": h for n, h in zip(DAY_NAMES, dt_)},
         "ST Hours": s, "OT Hours": o, "DT Hours": d, "Total Hours": s + o + d, "Base Hourly": base,
         "Overtime Hourly": otr, "Doubletime Hourly": dtr, "Cash Fringe Rate": cash, "3rd Party Fringe Rate": tp,
         "ST Prevailing Wage Rate Total": base + cash + tp, "Total ST Wages": stw, "Total OT Wages": otw,
         "Total DT Wages": dtw, "Grand Total Wages": round(stw + otw + dtw, 2)}
    v.update(override)
    return [v[h] for h in HEADER]


U = {"jur": UNION, "base": 42.0, "tp": 20.0}      # a correct union line's pay

# (case, area, rows, must, may). An empty "must" = a clean control: no flag at all may be raised.
CASES = [
    # --- incorrect rates
    # R07 too: for state rates it checks base + fringes against the total, which a low base also misses
    ("Base rate below the prevailing base", "rates", [line(base=35.0)], {"R08"}, {"R07"}),
    ("Fringe benefits short", "fringes", [line(tp=10.0)], {"R07"}, {"R08"}),
    ("Fringe paid as cash instead of to a plan (control: allowed)", "fringes", [line(cash=15.0, tp=0.0)], set(), set()),
    ("Union base below the agreement", "rates", [line(**{**U, "base": 40.0})], {"R08"}, {"R07"}),
    # --- overtime
    ("Overtime rate below 1.5x base", "overtime", [line(ot=(2, 0, 0, 0, 0, 0, 0), otr=50.0)], {"R03"}, {"S12"}),
    ("Overtime correctly paid (control)", "overtime", [line(ot=(2, 0, 0, 0, 0, 0, 0))], set(), set()),
    ("Double-time rate below 2x base", "overtime", [line(dt_=(0, 0, 0, 0, 0, 0, 4), dtr=70.0, **U, )], {"R04"}, set()),
    ("11 straight-time hours in a day", "overtime", [line(st=(11, 8, 8, 8, 0, 0, 0))], {"R01"}, set()),
    ("9 straight-time hours in a day", "overtime", [line(st=(9, 8, 8, 8, 0, 0, 0))], {"S01"}, set()),
    ("13 hours in a day", "overtime", [line(st=(10, 8, 8, 8, 0, 0, 0), ot=(3, 0, 0, 0, 0, 0, 0))], {"R02"}, {"S01"}),
    ("45 straight-time hours in the week", "overtime", [line(st=(9, 9, 9, 9, 9, 0, 0))], {"S05"}, {"S01"}),
    ("Union: Saturday paid straight time", "overtime", [line(st=(8, 8, 8, 8, 0, 8, 0), **U)], {"R05"}, set()),
    ("Union: Sunday not paid double time", "overtime", [line(st=(8, 8, 8, 8, 0, 0, 8), **U)], {"R06"}, set()),
    ("State: Saturday at straight time, 40 h week (control: no weekend rule)", "overtime",
     [line(st=(8, 8, 8, 8, 0, 8, 0))], set(), set()),
    # --- classification
    ("Classification not in the wage data", "classification", [line(cls="Tunnel Specialist")], {"S02"}, set()),
    ("Classification found only under another trade", "classification",
     [line(craft="Electricians", cls="General Laborer")], {"S11"}, {"S02", "R08", "R07"}),
    # --- wage change during the period
    # the week ending Sep 13 lies wholly after the change (the week ending Sep 6 straddles it)
    ("Rate rose Sep 1, old rate still paid", "rate change",
     [line(week=dt.date(2026, 9, 13), craft="Cement Masons", cls="Journey Level", base=45.0, tp=18.0)], {"R08"}, {"R07"}),
    ("Old rate paid before the change (control)", "rate change",
     [line(week=dt.date(2026, 8, 23), craft="Cement Masons", cls="Journey Level", base=45.0, tp=18.0)], set(), set()),
    ("New rate paid after the change (control)", "rate change",
     [line(week=dt.date(2026, 9, 13), craft="Cement Masons", cls="Journey Level", base=48.0, tp=18.0)], set(), set()),
    # --- arithmetic and data quality
    ("Total hours don't add up", "data", [line(**{"Total Hours": 41})], {"S03"}, set()),
    ("Grand total doesn't add up", "data", [line(**{"Grand Total Wages": 1500.0})], {"S04"}, set()),
    # the duplicate doubles the week's hours to 80 straight time: S05 is a fair consequence
    ("Duplicate line", "data", [line(), line()], {"S09"}, {"S05"}),
    ("Base rate missing", "data", [line(**{"Base Hourly": None}), line(emp="EMP-1f2e3d4c5b", **{"Employee ID (LCPtracker)": "EID-1f2e3d4c5b"})], {"S08"},
     {"R08", "R07", "R03", "S04", "S06"}),
    # "N/A" is a word for nothing: the same as an empty cell (it was read as 0.00: a red R08 and no S08)
    ("Base rate written N/A", "data", [line(**{"Base Hourly": "N/A"}), line(emp="EMP-2a3b4c5d6e", **{"Employee ID (LCPtracker)": "EID-2a3b4c5d6e"})], {"S08"},
     {"R08", "R07", "R03", "S04", "S06"}),
    # --- clean controls
    ("Correct state line (control)", "control", [line()], set(), set()),
    ("Correct union line (control)", "control", [line(**U)], set(), set()),
    ("Correct electrician line (control)", "control",
     [line(craft="Electricians", cls="Journey Level", base=60.0, tp=25.0)], set(), set()),
]


def evaluate(cases=CASES):
    """Runs every case through the payroll review. Returns (per-case results, per-rule stats)."""
    results = []
    for name, area, rows, must, may in cases:
        issues = review_rows([(i + 2, r) for i, r in enumerate(rows)], COLS, BOOK, CFG, project="Alpha")
        got = {i.rule for i in issues}
        results.append({"case": name, "area": area, "must": must, "got": got, "missed": must - got,
                        "extra": got - must - may})
    per_rule = defaultdict(lambda: {"seeded": 0, "detected": 0, "false_pos": 0})
    for r in results:
        for rule in r["must"]:
            per_rule[rule]["seeded"] += 1
            per_rule[rule]["detected"] += rule in r["got"]
        for rule in r["extra"]:
            per_rule[rule]["false_pos"] += 1
    return results, dict(per_rule)


def report(results, per_rule) -> str:
    seeded = [r for r in results if r["must"]]
    controls = [r for r in results if not r["must"]]
    caught = sum(1 for r in seeded if not r["missed"])
    clean = sum(1 for r in controls if not r["got"])
    fp = sum(len(r["extra"]) for r in results)
    out = [f"Seeded errors detected: {caught} of {len(seeded)}. Clean controls with no flag: {clean} of "
           f"{len(controls)}. False-positive flags: {fp}.", "",
           "| Rule | Meaning | Seeded | Detected | False positives |", "|---|---|---|---|---|"]
    for rule in sorted(per_rule):
        s = per_rule[rule]
        out.append(f"| {rule} | {RULE_TITLES.get(rule, rule)} | {s['seeded']} | {s['detected']} | {s['false_pos']} |")
    out += ["", "| Case | Area | Expected | Raised |", "|---|---|---|---|"]
    for r in results:
        out.append(f"| {r['case']} | {r['area']} | {', '.join(sorted(r['must'])) or 'none (control)'} | "
                   f"{', '.join(sorted(r['got'])) or 'none'} |")
    return "\n".join(out)


def test_every_seeded_violation_is_caught_and_controls_stay_clean():
    results, per_rule = evaluate()
    missed = [(r["case"], sorted(r["missed"])) for r in results if r["missed"]]
    extra = [(r["case"], sorted(r["extra"])) for r in results if r["extra"]]
    assert not missed, f"seeded violations missed: {missed}\n\n{report(results, per_rule)}"
    assert not extra, f"false positives: {extra}\n\n{report(results, per_rule)}"


def test_the_set_covers_the_reviewers_areas():
    """Incorrect rates, overtime, fringe benefits, classification, wage changes during the period, and controls."""
    areas = {c[1] for c in CASES}
    assert {"rates", "overtime", "fringes", "classification", "rate change", "control"} <= areas
    assert sum(1 for c in CASES if not c[3]) >= 6                 # clean controls


def test_an_export_without_any_base_rate_is_noted_once():
    """A Base Hourly column empty on every line (e.g. a redacted export) is one column note, not one per line."""
    rows = [line(**{"Base Hourly": None}), line(emp="EMP-1f2e3d4c5b", **{"Base Hourly": None, "Employee ID (LCPtracker)": "EID-1f2e3d4c5b"})]
    s08 = [i for i in review_rows([(i + 2, r) for i, r in enumerate(rows)], COLS, BOOK, CFG, project="Alpha")
           if i.rule == "S08"]
    assert len(s08) == 1 and "empty on every line" in s08[0].detail, [i.detail for i in s08]


def test_an_export_with_na_for_every_base_rate_is_noted_once():
    """"N/A" or "-" in Base Hourly on every line reads as an empty column: one note, not one S08 per line."""
    rows = [line(**{"Base Hourly": "N/A"}), line(emp="EMP-1f2e3d4c5b", **{"Base Hourly": " - ", "Employee ID (LCPtracker)": "EID-1f2e3d4c5b"})]
    s08 = [i for i in review_rows([(i + 2, r) for i, r in enumerate(rows)], COLS, BOOK, CFG, project="Alpha")
           if i.rule == "S08"]
    assert len(s08) == 1 and "no rate on any line" in s08[0].detail, [i.detail for i in s08]


def test_na_in_base_says_no_rate_and_a_blank_says_empty():
    """An accounting dash ("$ -") or "N/A" is not called empty: the words say there is no rate, read as 0.00."""
    for value, words in (("$ -", "has no rate"), ("N/A", "has no rate"), (None, "is empty")):
        rows = [line(**{"Base Hourly": value}), line(emp="EMP-2a3b4c5d6e", **{"Employee ID (LCPtracker)": "EID-2a3b4c5d6e"})]
        s08 = [i for i in review_rows([(i + 2, r) for i, r in enumerate(rows)], COLS, BOOK, CFG, project="Alpha")
               if i.rule == "S08"]
        assert len(s08) == 1 and s08[0].row == 2 and words in s08[0].detail, (value, [i.detail for i in s08])


def test_na_only_counts_as_empty_in_base_hourly():
    """The word-for-nothing reading is for Base Hourly only: "N/A" in Classification is not "Classification is empty"
    (it is looked up like any classification), and a real base rate on a line is never an S08."""
    issues = review_rows([(2, line(cls="N/A")), (3, line())], COLS, BOOK, CFG, project="Alpha")
    assert not [i for i in issues if i.rule == "S08" and i.col in ("classification", "base")], \
        [(i.rule, i.col, i.detail) for i in issues]


if __name__ == "__main__":
    print(report(*evaluate()))
