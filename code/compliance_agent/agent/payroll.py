"""Task 1: review the latest certified-payroll export against the in-use wage data and the
reference-workbook rules. Red = breaks a rule (definitely wrong); orange = suspected issue."""
import csv
import re
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import events, upload_checks
from .util import err_text, file_hash, file_kinds, norm, log, presave, run_timestamp, to_date, to_float, visible_files
from .config import project_cfg
from .layout import Paths
from .wages import jurisdiction_target, load_wagebook, match_key

# The payroll export kind: its columns, recognition and file types come from the file-kind registry
# (code/filekinds; vNext PR 4); the tokenizer and agent/uploads.py import the same definitions
from .cva_filekinds.payroll_export import (COLUMN_PATTERNS, DAYS, IMPORTANT, PAYROLL_EXTS,  # noqa: E402,F401
                                           find_header, map_columns)
from .cva_filekinds.payroll_export import HEADER_ROWS as PAYROLL_HEADER_ROWS  # noqa: E402
from .cva_filekinds.payroll_export import LINE_ROLES  # noqa: E402 - what makes a row a payroll line (#53)
from .cva_filekinds.tracker import AUTHOR, OLD_AUTHORS  # noqa: E402 - the agent's note authors (registry)

RULE_TITLES = {
    "R01": "More than 10 straight-time hours in a day",
    "R02": "More than 12 straight-time + overtime hours in a day",
    "R03": "Overtime rate below 1.5x base",
    "R04": "Double-time rate below 2x base",
    "R05": "Saturday worked without overtime",
    "R06": "Sunday worked without double time",
    "R07": "Fringe requirement not met",
    "R08": "Base / total prevailing rate not met",
    "S01": "More than 8 straight-time hours in a day",
    "S02": "No wage rate found for this craft/classification and date",
    "S03": "Hours don't add up",
    "S04": "Wage amounts don't add up",
    "S05": "Over 40 straight-time hours in the week with no overtime",
    "S06": "Payroll's prevailing-rate total doesn't equal its base + fringes",
    "S07": "Gross pay this project exceeds gross pay all projects",
    "S08": "Missing or negative value",
    "S09": "Duplicate payroll line",
    "S10": "Federal and state lines for the same week disagree",
    "S11": "Classification matched only approximately (or under another craft)",
    "S12": "Overtime pay below published overtime rate",
    "S13": "Wage determination or county named in Jurisdiction not in the wage data",
    "S14": "Wage determination not named and the project has several; one was chosen",
    "S15": "Not a usable number or date",
}
def names_breakdown(by_how) -> str:
    """'names: 230 lines matched by another name (name match (config.yaml) 179, ...); 882 lines not matched'."""
    matched = {h: n for h, n in by_how.items() if h != "no match"}
    out = []
    if matched:
        n = sum(matched.values())
        out.append(f"{n:,} line{'s' if n != 1 else ''} matched to the wage data by another name ("
                   + ", ".join(f"{h} {n:,}" for h, n in sorted(matched.items(), key=lambda x: -x[1])) + ")")
    if by_how.get("no match"):
        n = by_how["no match"]
        out.append(f"{n:,} line{'s' if n != 1 else ''} with no matching wage name")
    return "names: " + "; ".join(out) + " - listed in the Comments sheet"


def rule_breakdown(by_rule, top=5) -> str:
    """'rule-breaking: R08 Base / total prevailing rate not met 34, ... | suspected: S01 ... 7,886, ...'"""
    def part(prefix, label):
        items = sorted(((r, n) for r, n in by_rule.items() if r.startswith(prefix)), key=lambda x: -x[1])
        if not items:
            return None
        shown = ", ".join(f"{r} {RULE_TITLES.get(r, r)}: {n:,}" for r, n in items[:top])
        more = f", and {len(items) - top} more rule(s)" if len(items) > top else ""
        return f"{label}: {shown}{more}"
    return " | ".join(x for x in (part("R", "rule-breaking"), part("S", "suspected")) if x)


RULE_KEYS = {"R01": "R01_over_10_st_day", "R02": "R02_over_12_total_day", "R03": "R03_ot_rate_1_5x",
             "R04": "R04_dt_rate_2x", "R05": "R05_saturday_no_ot", "R06": "R06_sunday_no_dt",
             "R07": "R07_fringe_met", "R08": "R08_base_met"}
SUSPECT_KEYS = {"S01": "S01_over_8_st_day", "S02": "S02_rate_not_found", "S03": "S03_hours_dont_add_up",
                "S04": "S04_wages_dont_add_up", "S05": "S05_over_40_week_no_ot",
                "S06": "S06_payroll_rate_differs_from_wage_data", "S07": "S07_gross_this_project_gt_all",
                "S08": "S08_missing_or_negative", "S09": "S09_duplicate_row",
                "S10": "S10_fed_state_pair_mismatch", "S11": "S11_rate_matched_approximately",
                "S12": "S12_ot_below_published_ot_rate", "S13": "S13_jurisdiction_wd_not_loaded",
                "S14": "S14_determination_not_named", "S15": "S15_unusable_value"}


@dataclass
class Issue:
    severity: str          # red | orange
    rule: str
    row: int
    col: str               # column key
    detail: str
    ctx: dict = field(default_factory=dict)


# ------------------------------------------------------------ column map
# COLUMN_PATTERNS, DAYS, map_columns and PAYROLL_EXTS: from the registry (cva_filekinds.payroll_export), above


def inbox_payrolls(project_dir: Path):
    """Payroll exports waiting in Uploads/, oldest first."""
    from .uploads import uploads_of
    return uploads_of(project_dir, "payroll")


def last_reviewed_input(project_dir: Path, state):
    """The archived copy of the last payroll reviewed (used for re-reviews when nothing new was uploaded)."""
    last = state.get("payroll_last") or {}
    rel = last.get("archived_input")
    if rel and (project_dir / rel).is_file():
        return project_dir / rel
    from .layout import ARCHIVED
    arch = Paths(project_dir).payroll / ARCHIVED
    cands = [f for d in sorted(x for x in arch.iterdir() if x.is_dir()) for f in visible_files(d, PAYROLL_EXTS)
             if f.name == last.get("file")] if arch.exists() and last.get("file") else []
    if cands:
        return cands[-1]
    # Renamed since (the tokenizer's opaque_file_names, #53: words of a staff file's name become tokens on the AWS
    # side, and its first pass renames the copies already there): the run folder it was archived in keeps its name,
    # and so does what of the name survives the rename (its known words, numbers and dates: names.skeleton); the
    # payroll there with the same skeleton is the one, when there is exactly one
    if rel:
        from .cva_filekinds.names import skeleton
        folder = project_dir / Path(rel).parent
        found = [f for f in (visible_files(folder, PAYROLL_EXTS) if folder.is_dir() else [])
                 if "REVIEWED" not in f.name.upper()]
        same = [f for f in found if skeleton(f.name) == skeleton(Path(rel).name)]
        if len(same) == 1:
            return same[0]
        if len(found) == 1 and not same:
            return found[0]
    return None


def reviewed_name(src: Path, cfg, ts):
    """'2026-10-24 XHA 9.1.26-9.30.26 - REVIEWED.xlsx': the run's date first, so the folder lists by date."""
    return f"{ts[:10]} {src.stem}{cfg['payroll']['output_suffix']}.xlsx"


def earlier_reviews(project_dir: Path, src: Path, cfg):
    """Reviews of the same payroll file from earlier runs (dated, or named the old way without a date)."""
    suffix = re.escape(f"{src.stem}{cfg['payroll']['output_suffix']}.xlsx")
    pat = re.compile(rf"^(\d{{4}}-\d{{2}}-\d{{2}} )?{suffix}$")
    return [f for f in visible_files(Paths(project_dir).reviewed) if pat.match(f.name)]

def _load_payroll(path: Path):
    if path.suffix.lower() == ".csv":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Payroll"
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            for r in csv.reader(f):
                ws.append([to_float(v) if to_float(v) is not None and re.fullmatch(r"-?[\d.,$]+", v.strip() or "x") else v for v in r])
        return wb
    return openpyxl.load_workbook(path)


def _find_header(ws):
    """(header row, its values): rows 1-7 of the sheet, a cell that says exactly "classification" (the registry's
    find_header). A generator, so rows below the header are never touched (openpyxl creates the cells it is asked
    for)."""
    return find_header((r, [c.value for c in ws[r]]) for r in range(1, PAYROLL_HEADER_ROWS + 1))


_NEGATED_BEFORE = re.compile(r"(\bnon|\bnot|\bno)\b[\w\s&-]{0,25}$|non-?$", re.I)   # "not (Davis-Bacon) exempt"
_NEGATED_AFTER = re.compile(r"^\s*[:=?-]?\s*(no|n|false|0)\b", re.I)                      # "exempt: no"


def says_exempt(text, exempt_re):
    """True when the text says exempt: not "non-exempt", "not exempt", "not Davis-Bacon exempt" or "exempt: no"."""
    if not exempt_re:
        return False
    t = str(text or "")
    for m in exempt_re.finditer(t):
        before = re.split(r"[.;,()/|]", t[:m.start()])[-1]          # the same clause only
        if not _NEGATED_BEFORE.search(before) and not _NEGATED_AFTER.search(t[m.end():]):
            return True
    return False


def _kind(jurisdiction_text, patterns, exempt_re=None):
    t = str(jurisdiction_text or "").lower()
    # "State PW (Davis-Bacon exempt)": exempt from the federal rules, so a state line (whatever word comes first)
    if "state" in patterns and re.search(patterns["state"], t) and says_exempt(t, exempt_re):
        return "state"
    for kind, pat in patterns.items():
        if re.search(pat, t):
            return kind
    return None


# ------------------------------------------------- lines whose Jurisdiction doesn't say their wage type
class PayrollSkipped(Exception):
    """The payroll can't be checked as it stands (no wage rates, no way to tell a line's wage type, an upload check:
    agent/upload_checks.py): the message says why, for both emails. The payroll stays in Uploads. ``code``: the
    structured run log's reason (docs/run_log.md, agent/events.py REASONS)."""
    code = "wage_type_unknown"

    def __init__(self, message="", code=None):
        super().__init__(message)
        if code:
            self.code = code


def review_with_override(rows, cols, book, cfg, llm, project, stats, matches, kinds, unusable=None):
    """The review when some lines' Jurisdiction doesn't say federal, state or union: those lines are checked once per
    wage type the project has its own rates for (David, Nov 12, 2026; until then the types named by the Project Data
    file's Jurisdiction Override, Nov 7), each pass on its own. Every pass sees every line (so
    checks across lines, such as weekly hours, see the whole payroll), but only the first pass's flags count for the
    lines whose Jurisdiction says their type. On the override lines, a flag every pass raises is listed once; one
    only some passes raise starts with the wage type it was checked as ("Checked as federal: ..."). The tallies
    (stats, matches) count each override line once per pass, under "<Jurisdiction> (as <type>, ...)"."""
    cache, passes = {}, []
    for i, k in enumerate(kinds):
        st, mt = {}, Counter()
        iss = review_rows(rows, cols, book, cfg, llm, project=project, stats=st, matches=mt, forced_kind=k,
                          only_forced_stats=i > 0, llm_cache=cache, **({"unusable": unusable} if unusable else {}))
        passes.append((k, iss, st, mt))
    forced_rows = {rn for rn, v in rows if _forced_row(v, cols, cfg, project) and rn not in (unusable or {})}
    for i, (k, _iss, st, mt) in enumerate(passes):
        for jt, js in st.items():
            if js.get("override"):
                stats[f"{jt or '(empty)'} (as {k}, one of the project's wage types)"] = {**js, "kind": k}
            elif i == 0:
                stats[jt] = js
        if matches is not None:
            matches.update(mt)
    key = lambda x: (x.severity, x.rule, x.row, x.col, x.detail)
    if len(passes) == 1:
        return passes[0][1]
    on_forced = [{key(x) for x in iss if x.row in forced_rows} for _k, iss, _s, _m in passes]
    common = set.intersection(*on_forced)
    out, told = [], set()
    for i, (k, iss, _s, _m) in enumerate(passes):
        for x in iss:
            if x.row not in forced_rows:
                if i == 0:
                    out.append(x)                # a line that says its type: the first pass's flags only
            elif key(x) in common:
                if key(x) not in told:
                    told.add(key(x))
                    out.append(x)
            else:
                out.append(Issue(x.severity, x.rule, x.row, x.col, f"Checked as {k}: {x.detail}", dict(x.ctx)))
    return out


def _forced_row(v, cols, cfg, project=None) -> bool:
    """A payroll line whose Jurisdiction doesn't say federal, state or union (nor a mapping in config.yaml)."""
    pc = cfg["payroll"]
    jmap = {**(pc.get("jurisdiction_map") or {}), **(project_cfg(cfg, project).get("jurisdiction_map") or {})}
    exempt_re = re.compile(pc["federal_exempt_pattern"], re.I) if pc.get("federal_exempt_pattern") else None
    g = lambda k: v[cols[k]] if k in cols and cols[k] < len(v) else None
    jtext = str(g("jurisdiction") or "").strip()
    return not (jurisdiction_target(jtext, jmap)["kind"] or _kind(jtext, pc["jurisdiction_patterns"], exempt_re))


def undetermined_lines(rows, cols, cfg, project=None) -> int:
    """How many payroll lines have a Jurisdiction that doesn't say federal, state or union (nor a mapping in
    config.yaml): those are checked as each wage type the project has its own rates for."""
    n = 0
    for _rn, v in rows:
        g = lambda k: v[cols[k]] if k in cols and cols[k] < len(v) else None
        if all(x in (None, "") for x in v) or \
                all(g(k) in (None, "") for k in LINE_ROLES if k in cols):
            continue                             # empty or a footer line, as review_rows skips them
        n += _forced_row(v, cols, cfg, project)
    return n


# ------------------------------------------------------------- the review
RESIDENTIAL_SURE = {"exact", "same name, split differently", "same name without county/WA words"}


def upload_checks_today(tz):
    """The run's date in the configured time zone (the upload checks' "after today")."""
    import datetime as _dt
    try:
        from zoneinfo import ZoneInfo
        return _dt.datetime.now(ZoneInfo(tz)).date()
    except Exception:                          # an unknown zone: the server's date
        return _dt.date.today()


def rates_date_of(cfg, project_dir):
    return to_date(project_cfg(cfg, Path(project_dir).name).get("state_rates_date"))


def review_rows(rows, cols, book, cfg, llm=None, project=None, stats=None, matches=None, forced_kind=None,
                only_forced_stats=False, llm_cache=None, unusable=None):
    """rows: list of (excel_row_number, values). Returns list[Issue].
    Each line's Jurisdiction picks the wage type (federal/state/union) and, when it names one, the wage determination
    (e.g. WA20250091) and county used for that line; payroll.jurisdiction_map / projects.<name>.jurisdiction_map map
    a Jurisdiction text explicitly. `stats` (a dict) receives, per Jurisdiction text, the lines and what was applied.
    `matches` (a Counter) receives, per (kind, payroll craft, payroll classification, wage craft, wage classification,
    how), the lines whose names weren't an exact match to the wage data (wage side None: no match).
    forced_kind: the wage type for lines whose Jurisdiction doesn't say (each type the project has its own rates for,
    one pass per type: see review_with_override). only_forced_stats: stats and matches count only those lines (a later pass:
    the other lines were counted in the first). llm_cache: the local model's answers, shared between passes.
    unusable: {sheet row: [(column role, words)]} from the upload checks (agent/upload_checks.py): those lines get S15
    on each cell and no other check."""
    pc = cfg["payroll"]
    jmap = {**(pc.get("jurisdiction_map") or {}), **(project_cfg(cfg, project).get("jurisdiction_map") or {})}
    # the project's county (projects.<name>.county), used for state lines whose Jurisdiction names none
    pcounty = (norm(project_cfg(cfg, project).get("county") or "").split() or [None])[0]
    # the date whose L&I rates apply to every state line (projects.<name>.state_rates_date: the bid due date)
    rates_date = to_date(project_cfg(cfg, project).get("state_rates_date"))
    # residential projects: a state line uses the residential classification when there is one
    # (projects.<name>.prefer_residential), the commercial one otherwise
    prefer_res = bool(project_cfg(cfg, project).get("prefer_residential"))
    book_kinds = {kd for kd, _ in book.by_key}
    targets = {}
    stats = stats if stats is not None else {}
    tol = pc.get("money_tolerance", 0.02)
    rules = pc["rules"]
    suspects = pc["suspect_checks"]
    pats = pc["jurisdiction_patterns"]
    issues = []
    llm_cache = {} if llm_cache is None else llm_cache
    exempt_re = re.compile(pc["federal_exempt_pattern"], re.I) if pc.get("federal_exempt_pattern") else None

    def on(rule, kind):
        if rule.startswith("R"):
            r = rules.get(RULE_KEYS[rule], {})
            return r.get("severity", "red") != "off" and kind in r.get("applies_to", [])
        return bool(suspects.get(SUSPECT_KEYS[rule], True))

    def sev(rule):
        return rules.get(RULE_KEYS[rule], {}).get("severity", "red") if rule.startswith("R") else "orange"

    def add(rule, rn, col, detail, ctx):
        if col not in cols:
            col = "classification" if "classification" in cols else next(iter(cols))
        issues.append(Issue(sev(rule), rule, rn, col, detail, ctx))

    # A column that is empty on every line (e.g. employee names in a redacted export) is reported once, not per line
    data_rows = [v for _, v in rows if any(x not in (None, "") for x in v)]
    # "N/A", "-" and the like in Base Hourly mean no rate was entered (upload_checks.says_empty): empty, as a blank
    nothing = lambda key, x: x in (None, "") or (key == "base" and upload_checks.says_empty(x))
    all_empty = {key for key in ("week", "craft", "classification", "employee", "base") if key in cols and data_rows
                 and all(nothing(key, v[cols[key]] if cols[key] < len(v) else None) for v in data_rows)}
    if all_empty and on("S08", "state"):
        first = next(rn for rn, v in rows if any(x not in (None, "") for x in v))
        for key in sorted(all_empty):
            words = key == "base" and any(upload_checks.says_empty(v[cols[key]] if cols[key] < len(v) else None)
                                          for v in data_rows)
            issues.append(Issue("orange", "S08", first, key,
                                f"{key.title()} column has no rate on any line (\"N/A\", a dash or empty); not flagged "
                                "line by line" if words else
                                f"{key.title()} column is empty on every line (redacted export?); not flagged line by line",
                                {}))
    groups_week = defaultdict(list)
    partial_weeks = set()                    # (contractor, employee, week) with a line left out (upload checks)
    dup_seen = {}
    pairs = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0]))

    for rn, v in rows:
        g = lambda k: v[cols[k]] if k in cols and cols[k] < len(v) else None
        f = lambda k: to_float(g(k)) or 0.0
        if all(x in (None, "") for x in v):
            continue
        if all(g(k) in (None, "") for k in LINE_ROLES if k in cols):
            continue  # report footer / notes line (the registry's LINE_ROLES: the tokenizer tokenizes its text, #53)
        if unusable and rn in unusable:
            # a value the review can't use (upload checks): marked, and the line left out of every other check; the
            # worker's week is then incomplete, so the checks across that week's lines (S05, S10) skip it
            partial_weeks.add((g("contractor"), g("emp_id") or g("employee"), to_date(g("week"))))
            if suspects.get(SUSPECT_KEYS["S15"], True):
                bad_ctx = {"contractor": g("contractor"), "employee": g("employee"), "week": to_date(g("week")),
                           "jurisdiction": g("jurisdiction"), "kind": None, "craft": g("craft"),
                           "classification": g("classification")}
                for role, words in unusable[rn]:
                    add("S15", rn, role, words, bad_ctx)
            continue
        jtext = str(g("jurisdiction") or "").strip()
        if jtext not in targets:
            targets[jtext] = jurisdiction_target(jtext, jmap)
        target = targets[jtext]
        said = target["kind"] or _kind(jtext, pats, exempt_re)
        kind = said or forced_kind
        week = to_date(g("week"))
        if only_forced_stats and said:          # counted by the first pass: this pass's tallies are thrown away
            js, line_matches = {"lines": 0, "kind": kind, "asked": target, "applied": Counter()}, None
        else:
            js = stats.setdefault(jtext, {"lines": 0, "kind": kind, "asked": target, "applied": Counter()})
            line_matches = matches
        if forced_kind and not said:
            js["override"] = True
        js["lines"] += 1
        ctx = {"contractor": g("contractor"), "employee": g("employee"), "week": week,
               "jurisdiction": g("jurisdiction"), "kind": kind, "craft": g("craft"),
               "classification": g("classification")}
        k = kind or "state"

        # ---- S08 missing / negative
        if on("S08", k):
            for key in ("week", "craft", "classification", "employee"):
                if key in cols and key not in all_empty and g(key) in (None, ""):
                    add("S08", rn, key, f"{key.replace('_', ' ').title()} is empty", ctx)
            # a line with hours but no base rate: say it's missing (it is also read as 0, so the rate checks flag it)
            hours = sum(to_float(g(h)) or 0 for h in ("st_hours", "ot_hours", "dt_hours") if h in cols)
            if "base" in cols and "base" not in all_empty and hours > 0 and nothing("base", g("base")):
                add("S08", rn, "base", "Base Hourly is empty on a line with hours" if g("base") in (None, "") else
                    "Base Hourly has no rate (\"N/A\", a dash or similar) on a line with hours; read as 0.00", ctx)
            for key in [c for c in cols if re.match(r"(st|ot|dt)_(hours|mon|tue|wed|thu|fri|sat|sun)$|base|.*_rate$|.*hourly$", c)]:
                x = to_float(g(key))
                if x is not None and x < 0:
                    add("S08", rn, key, f"Negative value {x}", ctx)
            if kind is None and on("S08", k):
                add("S08", rn, "jurisdiction", "Can't tell if this line is federal, state or union from the Jurisdiction text; treated as state", ctx)

        # ---- daily hours (R01, R02, S01, R05, R06)
        for d in DAYS:
            st, ot = f(f"st_{d}"), f(f"ot_{d}")
            if st > 10 and on("R01", k):
                add("R01", rn, f"st_{d}", f"{st:g} straight-time hours on {d.title()} (max 10)", ctx)
            elif st > 8 and on("S01", k):
                add("S01", rn, f"st_{d}", f"{st:g} straight-time hours on {d.title()} (over 8)", ctx)
            # reference rules workbook "Over 12 total": ST + OT > 12 (hours past 12 must be paid as double time)
            if st + ot > 12 and on("R02", k):
                add("R02", rn, f"ot_{d}" if ot else f"st_{d}",
                    f"{st + ot:g} straight-time + overtime hours on {d.title()} (max 12; later hours are double time)", ctx)
        # reference rules workbook "Saturday no OT": Saturday straight time, unless both OT and DT were also paid that day
        if on("R05", k) and f("st_sat") > 0 and (f("ot_sat") == 0 or f("dt_sat") == 0):
            add("R05", rn, "st_sat", f"{f('st_sat'):g} Saturday hours paid as straight time", ctx)
        if on("R06", k) and (f("st_sun") > 0 or f("ot_sun") > 0):
            sun_dt = f("dt_sun") if "dt_sun" in cols else f("dt_hours")
            if sun_dt == 0:
                add("R06", rn, "st_sun" if f("st_sun") else "ot_sun", "Sunday hours not paid as double time", ctx)

        # ---- hour arithmetic (S03)
        if on("S03", k):
            for p in ("st", "ot", "dt"):
                if f"{p}_hours" in cols and any(f"{p}_{d}" in cols for d in DAYS):
                    s = sum(f(f"{p}_{d}") for d in DAYS)
                    if abs(s - f(f"{p}_hours")) > 0.01:
                        add("S03", rn, f"{p}_hours", f"{p.upper()} Hours is {f(f'{p}_hours'):g} but the days add up to {s:g}", ctx)
            if "total_hours" in cols:
                s = f("st_hours") + f("ot_hours") + f("dt_hours")
                if abs(s - f("total_hours")) > 0.01:
                    add("S03", rn, "total_hours", f"Total Hours is {f('total_hours'):g} but ST+OT+DT = {s:g}", ctx)

        # ---- wage arithmetic (S04)
        if on("S04", k):
            for p, rate in (("st", "base"), ("ot", "ot_hourly"), ("dt", "dt_hourly")):
                wk = f"total_{p}_wages"
                if wk in cols and f"{p}_hours" in cols and rate in cols:
                    hrs = f(f"{p}_hours")
                    exp_cash = hrs * (f(rate) + f("cash_fringe_rate"))
                    exp_plain = hrs * f(rate)
                    got = f(wk)
                    if hrs and min(abs(got - exp_cash), abs(got - exp_plain)) > tol * hrs + 0.05:
                        add("S04", rn, wk, f"{wk.replace('_', ' ').title()} is {got:.2f}; {hrs:g} h x {f(rate):.2f} = {exp_plain:.2f}", ctx)
            if "grand_total" in cols:
                s = f("total_st_wages") + f("total_ot_wages") + f("total_dt_wages") + f("per_diem")
                if abs(s - f("grand_total")) > 0.05:
                    add("S04", rn, "grand_total", f"Grand Total is {f('grand_total'):.2f} but ST+OT+DT wages + per diem = {s:.2f}", ctx)

        if on("S07", k) and "gross_this" in cols and "gross_all" in cols and f("gross_all") and f("gross_this") > f("gross_all") + 0.01:
            add("S07", rn, "gross_this", f"{f('gross_this'):.2f} this project vs {f('gross_all'):.2f} all projects", ctx)

        # ---- rates against wage data
        rate, how, in_range = (None, None, False)
        use_pc = False
        # federal lines marked exempt from Davis-Bacon (e.g. "EXEMPT - federal Davis-Bacon does not apply") skip the
        # federal wage check; state lines are always checked
        exempt = kind == "federal" and any(says_exempt(g(c), exempt_re) for c in ("classification", "craft", "jurisdiction"))
        if exempt:
            js["exempt"] = js.get("exempt", 0) + 1
        if g("classification") and not exempt:          # unrecognized Jurisdiction: checked as state (k)
            use_pc = bool(pcounty) and k == "state" and not target["county"]
            when = rates_date if (k == "state" and rates_date) else week
            rate, how, in_range = book.lookup(k, g("craft"), g("classification"), when,
                                              target["wage_decision"], pcounty if use_pc else target["county"])
            if k == "state" and prefer_res and g("craft") and not norm(g("craft")).startswith("residential"):
                r2, how2, ir2 = book.lookup(k, f"Residential {g('craft')}", g("classification"), when,
                                            target["wage_decision"], pcounty if use_pc else target["county"])
                if r2 is not None and norm(r2["craft"]).startswith("residential") and \
                        (how2.split(" | ")[0] in RESIDENTIAL_SURE or how2.startswith("name match")):
                    rate, how, in_range = r2, "residential preferred, " + how2, ir2
                    js["residential"] = js.get("residential", 0) + 1
            if k == "state" and rates_date:
                js["rates_date"] = js.get("rates_date", 0) + 1
            if rate is None and k not in book_kinds:
                js["no_data"] = js.get("no_data", 0) + 1        # no wage data of this kind in the project at all
            if rate is not None and not target["wage_decision"] and not (k == "state" and rates_date):
                # the Jurisdiction names no determination, and this classification is in several: one was chosen
                decs = sorted({str(c["wage_decision"]) for c in book.by_key.get((k, match_key(rate)), [])})
                if len(decs) > 1:
                    js["unclear"] = js.get("unclear", 0) + 1
                    js.setdefault("unclear_decs", set()).update(decs)
                    if on("S14", k):
                        add("S14", rn, "jurisdiction",
                            f"Jurisdiction '{jtext}' names no wage determination and this classification is in "
                            f"{len(decs)} of the project's {k} ones ({', '.join(decs)}); checked against "
                            f"{rate['wage_decision']}, the one in effect on {week}. "
                            + ("Set projects.<project>.state_rates_date (the bid due date) or name the determination "
                               "in projects.<project>.jurisdiction_map in config.yaml" if k == "state" else
                               "Name the determination in projects.<project>.jurisdiction_map in config.yaml"), ctx)
            if use_pc and rate is not None:
                js["project_county"] = js.get("project_county", 0) + 1
            if line_matches is not None and not (how or "").startswith("exact"):
                line_matches[(k, str(g("craft") or ""), str(g("classification") or ""),
                         rate["craft"] if rate else None, rate["classification"] if rate else None,
                         (how or "").split(" | ")[0] or "no match")] += 1
            if rate is not None:
                js["applied"][f"{rate['wage_decision']}" + (f" ({rate['county']})" if rate.get("county") else "")] += 1
                if " | " in how:
                    how, note = how.split(" | ", 1)
                    if on("S13", k):
                        add("S13", rn, "jurisdiction",
                            f"Jurisdiction '{jtext}': {note}; checked against {rate['wage_decision']}"
                            + (f" ({rate['county']})" if rate.get("county") else ""), ctx)
        if exempt:
            pass
        elif rate is None:
            if on("S02", k) and g("classification"):
                hint = ""
                key = (k, g("craft"), g("classification"))
                if llm is not None and llm.available():
                    if key not in llm_cache:
                        llm_cache[key] = llm.suggest_classification(g("craft"), g("classification"), book.candidates(k))
                    if llm_cache[key]:
                        hint = f" Possible match (local model suggestion, verify): {llm_cache[key]}."
                add("S02", rn, "classification",
                    f"No {k} wage rate found for '{g('craft')} / {g('classification')}'" +
                    ("" if book else " (no wage data loaded for this project)") + "." + hint +
                    (" If it's the same work under another name, map it under payroll.name_matches in config.yaml."
                     if book else ""), ctx)
        else:
            ctx = {**ctx, "wage_decision": rate["wage_decision"],
                   "rate": f"{rate['wage_decision']} {rate['craft']} / {rate['classification']}"
                   + (f" ({str(rate.get('county') or pcounty).title()} County: the project's county, as the "
                      "Jurisdiction names none)" if use_pc else "")}
            if how.startswith("approx") and on("S11", k):
                add("S11", rn, "classification", f"Matched to '{rate['craft']} / {rate['classification']}' ({how}); rule checks used that rate", ctx)
            elif how.startswith("classification name only") and on("S11", k):
                add("S11", rn, "classification",
                    f"Matched on the classification name only, under another craft: '{rate['craft']} / {rate['classification']}'; "
                    "rule checks used that rate. If that's right, map it under payroll.name_matches in config.yaml", ctx)
            if not in_range and on("S02", k):
                add("S02", rn, "week", (f"No state rate in effect on {rates_date}, the project's state rates date "
                                        "(projects.<project>.state_rates_date)" if k == "state" and rates_date
                                        else f"No {k} rate effective on {week}")
                    + f"; checked against {rate['wage_decision']} issued {rate['issue_date']}", ctx)
            base, cashf, tpf = f("base"), f("cash_fringe_rate"), f("tp_fringe_rate")
            vac = rate.get("vac_pto") or 0.0
            basic, fringe, total = rate.get("basic"), rate.get("fringe"), rate.get("total")
            # R08 base / total
            if on("R08", k):
                # union: base hourly vs CBA base (reference rules workbook); federal/state: cash paid per hour
                # (base + cash in lieu of fringes) vs required base
                cash = base if kind == "union" else base + cashf
                if basic:
                    if cash + tol < basic:
                        what = "Base" if kind == "union" else "Base + cash fringe"
                        add("R08", rn, "base", f"{what} {cash:.2f} < required base {basic:.2f} ({rate['wage_decision']})", ctx)
                elif total and base + cashf + tpf + tol < total:
                    add("R08", rn, "base", f"Base + fringes {base + cashf + tpf:.2f} < prevailing total {total:.2f} ({rate['wage_decision']})", ctx)
            # R07 fringes
            if on("R07", k) and fringe:
                if kind == "union":
                    if tpf + tol < fringe:
                        add("R07", rn, "tp_fringe_rate", f"3rd-party fringe {tpf:.2f} < required {fringe:.2f}", ctx)
                elif basic and total and base + cashf + tpf + tol < total:
                    add("R07", rn, "tp_fringe_rate", f"Base + fringes {base + cashf + tpf:.2f} < required total {total:.2f} (fringe {fringe:.2f})", ctx)
            # R03 / R04 overtime multipliers (master: OT >= 1.5 x (base - vac/PTO))
            if on("R03", k) and f("ot_hours") > 0 and "ot_hourly" in cols:
                need = (base - vac) * 1.5
                if f("ot_hourly") + tol < need:
                    add("R03", rn, "ot_hourly", f"OT rate {f('ot_hourly'):.2f} < 1.5 x {base - vac:.2f} = {need:.2f}", ctx)
            if on("R04", k) and f("dt_hours") > 0 and "dt_hourly" in cols:
                need = (base - vac) * 2
                if f("dt_hourly") + tol < need:
                    add("R04", rn, "dt_hourly", f"DT rate {f('dt_hourly'):.2f} < 2 x {base - vac:.2f} = {need:.2f}", ctx)
            if on("S06", k) and "st_pw_total" in cols and f("st_pw_total") and abs(f("st_pw_total") - (base + cashf + tpf)) > tol:
                add("S06", rn, "st_pw_total", f"Shows {f('st_pw_total'):.2f}; base {base:.2f} + cash fringe {cashf:.2f} + 3rd-party fringe {tpf:.2f} = {base + cashf + tpf:.2f}", ctx)
            if on("S12", k) and rate.get("ot_total") and f("ot_hours") > 0 and "ot_hourly" in cols:
                paid = f("ot_hourly") + cashf + tpf
                if paid + tol < rate["ot_total"]:
                    add("S12", rn, "ot_hourly", f"OT pay incl. fringes {paid:.2f} < published OT rate {rate['ot_total']:.2f}", ctx)

        # ---- cross-row bookkeeping
        emp = g("emp_id") or g("employee")
        groups_week[(g("contractor"), emp, week, k)].append((rn, f("st_hours"), f("ot_hours"), f("dt_hours"), ctx))
        dkey = (g("contractor"), emp, week, k, g("craft"), g("classification"), f("st_hours"), f("ot_hours"), f("grand_total"))
        if on("S09", k) and dkey in dup_seen:
            add("S09", rn, "employee", f"Same as row {dup_seen[dkey]}", ctx)
        dup_seen.setdefault(dkey, rn)
        if kind in ("federal", "state"):
            p = pairs[(g("contractor"), emp, week)][kind]
            p[0] += f("st_hours"); p[1] += f("ot_hours"); p[2] += f("grand_total")
            pairs[(g("contractor"), emp, week)].setdefault("_rows", []).append((rn, ctx))

    if suspects.get("S05_over_40_week_no_ot", True):
        for (c_, e_, w_, k), lines in groups_week.items():
            if (c_, e_, w_) in partial_weeks:
                continue                     # a line of that week was left out: its total isn't known
            st = sum(x[1] for x in lines)
            ot = sum(x[2] + x[3] for x in lines)
            if st > 40 and ot == 0:
                for rn, *_, ctx in lines:
                    add("S05", rn, "st_hours", f"{st:g} straight-time hours this week for this employee, no overtime", ctx)
    if suspects.get("S10_fed_state_pair_mismatch", True):
        for key, d in pairs.items():
            if key in partial_weeks:
                continue
            if "federal" in d and "state" in d:
                fe, sta = d["federal"], d["state"]
                if any(abs(a - b) > 0.05 for a, b in zip(fe, sta)):
                    for rn, ctx in d["_rows"]:
                        add("S10", rn, "st_hours",
                            f"Federal lines: ST {fe[0]:g} / OT {fe[1]:g} / ${fe[2]:.2f}; state lines: ST {sta[0]:g} / OT {sta[1]:g} / ${sta[2]:.2f}", ctx)
    return issues


# ------------------------------------------------------------ write output
def _apply_marks(ws, issues, cols, colors):
    by_cell = defaultdict(list)
    for i in issues:
        by_cell[(i.row, cols[i.col] + 1)].append(i)
    red = PatternFill("solid", fgColor=colors["red"])
    orange = PatternFill("solid", fgColor=colors["orange"])
    for (r, c), lst in by_cell.items():
        cell = ws.cell(row=r, column=c)
        lst.sort(key=lambda i: (i.severity != "red", i.rule))
        cell.fill = red if lst[0].severity == "red" else orange
        text = "\n".join(f"[{'RULE BROKEN' if i.severity == 'red' else 'SUSPECTED'}] {i.rule} {RULE_TITLES[i.rule]}: {i.detail}" for i in lst)
        if cell.comment and cell.comment.author not in (AUTHOR, *OLD_AUTHORS):
            text = cell.comment.text + "\n---\n" + text
        cm = Comment(text[:30000], AUTHOR)
        cm.width, cm.height = 360, min(400, 60 + 40 * len(lst))
        cell.comment = cm


def _comments_sheet(wb, ws, issues, cols, header, meta):
    name = "Comments" if "Comments" not in wb.sheetnames else "Comments (agent)"
    if name in wb.sheetnames:
        del wb[name]
    cs = wb.create_sheet(name, 0)
    bold = Font(bold=True)
    cs["A1"] = f"Payroll review: {meta['file']}"
    cs["A1"].font = Font(bold=True, size=14)
    cs["A2"] = f"Run {meta['run']} (UTC) | {meta['rows']} payroll lines reviewed | wage data: {meta['wages']}"
    reds = [i for i in issues if i.severity == "red"]
    oranges = [i for i in issues if i.severity != "red"]
    cs["A3"] = f"{len(reds)} rule-breaking issue(s) (red) on {len({i.row for i in reds})} line(s); " \
               f"{len(oranges)} suspected issue(s) (orange) on {len({i.row for i in oranges})} line(s)"
    cs["A3"].font = bold
    if meta.get("jurisdictions"):
        cs["A4"] = f"Wage determinations applied, by Jurisdiction: {meta['jurisdictions']}"
    if meta.get("exempt"):
        cs["A4"] = (cs["A4"].value + " | " if cs["A4"].value else "") + \
            f"{meta['exempt']} federal line(s) marked exempt from Davis-Bacon: federal wage rate not checked"
    r = 5
    cs.cell(r, 1, "Summary by rule").font = bold
    r += 1
    counts = defaultdict(int)
    for i in issues:
        counts[(i.severity, i.rule)] += 1
    for (s, rule), n in sorted(counts.items(), key=lambda x: (x[0][0] != "red", x[0][1])):
        cs.cell(r, 1, "Rule-breaking" if s == "red" else "Suspected")
        cs.cell(r, 2, rule)
        cs.cell(r, 3, RULE_TITLES[rule])
        cs.cell(r, 4, n)
        r += 1
    wrap_rows = set()
    matches = meta.get("matches") or {}
    if matches:
        r += 1
        c = cs.cell(r, 1, "NAMES MATCHED TO THE WAGE DATA: payroll craft/classification names that aren't written exactly "
                          "as in the wage determination, and the wage rate each was checked against")
        c.font = Font(bold=True, size=12)
        r += 1
        cs.cell(r, 1, "Exact matches aren't listed. To match a name differently, or one with no match, add it under "
                      "payroll.name_matches (all projects) or projects.<project>.name_matches in config.yaml, e.g. "
                      "{payroll: \"PAYROLL CRAFT / PAYROLL CLASSIFICATION\", wage: \"WAGE CRAFT / WAGE CLASSIFICATION\", kind: state}")
        r += 1
        for j, h in ((1, "Wage type"), (2, "Lines"), (3, "Payroll craft / classification"),
                     (4, "Wage table craft / classification (rate used)"), (9, "How matched")):
            cs.cell(r, j, h).font = bold
        cs.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        r += 1
        for (kind, pcraft, pcls, wcraft, wcls, how), n in sorted(
                matches.items(), key=lambda x: (x[0][3] is None, x[0][0] or "", -x[1], x[0][1], x[0][2])):
            cs.cell(r, 1, kind)
            cs.cell(r, 2, n)
            cs.cell(r, 3, f"{pcraft} / {pcls}")
            cs.cell(r, 4, f"{wcraft} / {wcls}" if wcraft is not None else "NO MATCH (rule S02: no wage rate found)")
            cs.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
            cs.cell(r, 9, how if wcraft is not None else "-")
            if wcraft is None:
                cs.cell(r, 4).font = Font(bold=True)
            wrap_rows.add(r)
            r += 1
    cols_out = ["Severity", "Rule", "Issue", "Cell", "Payroll row", "Contractor", "Employee",
                "Week end", "Jurisdiction", "Craft", "Classification", "Rate used", "Detail"]
    fills = {"red": PatternFill("solid", fgColor=meta["colors"]["red"]),
             "orange": PatternFill("solid", fgColor=meta["colors"]["orange"])}
    for title, lst, sev in (("RULE-BREAKING ISSUES (red): definitely wrong per the reference rules workbook rules", reds, "red"),
                            ("SUSPECTED ISSUES (orange): possibly wrong, please check", oranges, "orange")):
        r += 1
        c = cs.cell(r, 1, title)
        c.font = Font(bold=True, size=12)
        c.fill = fills[sev]
        r += 1
        for j, h in enumerate(cols_out, 1):
            cs.cell(r, j, h).font = bold
        r += 1
        if not lst:
            cs.cell(r, 1, "None found")
            r += 1
        for i in sorted(lst, key=lambda i: (i.rule, i.row)):
            ref = f"{get_column_letter(cols[i.col] + 1)}{i.row}"
            vals = ["Rule-breaking" if sev == "red" else "Suspected", i.rule, RULE_TITLES[i.rule], ref, i.row,
                    i.ctx.get("contractor"), i.ctx.get("employee"), i.ctx.get("week"), i.ctx.get("jurisdiction"),
                    i.ctx.get("craft"), i.ctx.get("classification"), i.ctx.get("rate"), i.detail]
            for j, val in enumerate(vals, 1):
                cs.cell(r, j, val)
            link = cs.cell(r, 4)
            link.hyperlink = f"#'{ws.title}'!{ref}"
            link.font = Font(color="FF0563C1", underline="single")
            r += 1
    widths = [13, 6, 44, 8, 8, 28, 26, 11, 30, 18, 30, 34, 80]
    for j, w in enumerate(widths, 1):
        cs.column_dimensions[get_column_letter(j)].width = w
    for row in cs.iter_rows(min_row=6):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=c.column == 13 or (c.row in wrap_rows and c.column in (3, 4, 9)))
    return name


def _audit_sheet(wb, run, notes):
    """'Audit' sheet: timestamped notes about the whole file, too repetitive to put on each cell."""
    name = "Audit" if "Audit" not in wb.sheetnames else "Audit (agent)"
    if name in wb.sheetnames:
        del wb[name]
    a = wb.create_sheet(name, 1)
    for j, h in enumerate(["Run (UTC)", "Task", "Topic", "Note"], 1):
        a.cell(1, j, h).font = Font(bold=True)
    for i, (topic, note) in enumerate(notes or [("Review", "Nothing to note for the whole file")], 2):
        for j, v in enumerate([run, "payroll review", topic, note], 1):
            a.cell(i, j, v).alignment = Alignment(vertical="top", wrap_text=j == 4)
    for j, w in enumerate([20, 15, 15, 120], 1):
        a.column_dimensions[get_column_letter(j)].width = w
    a.freeze_panes = "A2"
    return name


def _review_one(src: Path, project_dir: Path, cfg, book, llm, tz):
    log.info("Reviewing payroll file", extra={"details": {"type": src.suffix.lower(), "size_bytes": src.stat().st_size}})
    wb = _load_payroll(src)
    ws = wb.worksheets[0]
    hrow, header = _find_header(ws)
    cols = map_columns(header)
    important = list(IMPORTANT)
    missing = [c for c in important if c not in cols]
    log.debug("Payroll columns found", extra={"details": {"sheet": ws.title, "header_row": hrow,
                                                            "mapped": len(cols), "columns": ", ".join(sorted(cols))}})
    audit = []                 # (topic, note): what applies to the whole file, for the Audit sheet
    gaps = []                  # what kept lines from being fully validated: listed in the email, which then says Failure
    if missing:
        gaps.append(f"columns not found: {', '.join(missing)}; the checks that need them were skipped on every line. "
                    "Export the payroll with these columns and upload it again")
        log.info("Payroll file is missing expected columns; the checks that need them were skipped",
                    extra={"details": {"missing": ", ".join(missing)}})
        audit.append(("Columns", f"Columns not found in the payroll: {', '.join(missing)}. The checks that need them "
                                 "were skipped for every line."))
    rows = [(r[0].row, [c.value for c in r]) for r in ws.iter_rows(min_row=hrow + 1)]
    # Upload checks (agent/upload_checks.py; payroll.upload_checks, on in config.yaml since Nov 15, 2026, kept on
    # after the trial): a payroll that can't be checked as it stands is skipped; a line with a value the review can't
    # use is marked S15 and left out of the rest
    unusable = {}
    uc = upload_checks.settings(cfg)
    if uc:
        problem = upload_checks.file_problem(rows, cols, header)
        if problem:
            raise PayrollSkipped(problem[1], code=problem[0])
        unusable = upload_checks.line_marks(rows, cols, header, upload_checks_today(tz), uc)
        if unusable:
            n_bad = sum(len(x) for x in unusable.values())
            audit.append(("Values", f"{n_bad:,} value(s) on {len(unusable):,} line(s) are not a usable number or date "
                                    "(flagged S15); those lines were left out of every other check. Fix them in "
                                    "the export and upload it again"))
            log.info("Payroll lines with an unusable value were left out of the checks",
                     extra={"details": {"lines": len(unusable), "values": n_bad}})
        else:           # shown every time, so a reviewed payroll says the checks ran (David, Nov 14, 2026)
            audit.append(("Values", "Every value the review uses is a usable number or date (checked before the review)"))
    jstats = {}
    matches = Counter()
    # Lines whose Jurisdiction doesn't say federal, state or union are checked as each wage type the project's own
    # tables give (David, Nov 12, 2026: only the rates the staff provided for the project)
    unsure = undetermined_lines([(rn, v) for rn, v in rows if rn not in unusable], cols, cfg, project_dir.name)
    extra_checks = 0                         # line checks beyond one per line: the later passes
    if unsure:
        # the project's types: its typed tables, and a table given without its Jurisdiction column (that pass has no
        # rates and says so, as a gap: review R2), never the firm-wide union table
        own = set(getattr(book, "own_kinds", [])) | set(getattr(book, "untyped_kinds", set()))
        kinds = [k for k in cfg["wages"]["kinds"] if k in own]
        if not kinds:                        # (a safety net: a project without typed rates is skipped before this)
            raise PayrollSkipped(f"{unsure:,} payroll line(s) don't say in their Jurisdiction whether federal, state or "
                                 "union rates apply, and the project's wage rates don't say which type they are")
        audit.append(("Jurisdiction", f"{unsure:,} line(s) whose Jurisdiction doesn't say federal, state or union were "
                                      "checked against the project's own wage rates: "
                                      + ", ".join(k if k in getattr(book, "own_kinds", []) else f"{k} (no usable "
                                                  "rates: no Jurisdiction column)" for k in kinds) + (" (one pass each; a flag only some passes raise says "
                                                            "which, and the counts include each pass's flags)"
                                                            if len(kinds) > 1 else "")))
        issues = review_with_override(rows, cols, book, cfg, llm, project_dir.name, jstats, matches, kinds,
                                      unusable=unusable)
        extra_checks = unsure * (len(kinds) - 1)
    else:
        issues = review_rows(rows, cols, book, cfg, llm, project=project_dir.name, stats=jstats, matches=matches,
                             unusable=unusable)
    if getattr(book, "untyped", 0):
        audit.append(("Wage data", f"{book.untyped:,} rate(s) in the wage rates file have no Jurisdiction (merged from a "
                                   "file without the column) and were not used"))
    audit.append(("Wage data", "Wage data in use: " + ", ".join(
        f"{k} ({len(book.decisions(k))} determination(s): {', '.join(book.decisions(k)[:6])})" if book.decisions(k)
        else f"{k} (none)" for k in cfg["wages"]["kinds"])))
    for bad in getattr(book, "bad_name_matches", []):
        log.warning("A name match in config.yaml wasn't used", extra={"details": {"name_match": bad}})
        audit.append(("Name matches", f"A name match in config.yaml wasn't used: {bad}"))
    by_how = Counter()
    for key, n in matches.items():
        by_how[key[5].split(":")[0] if key[3] is not None else "no match"] += n
    exempt = sum(js.get("exempt", 0) for js in jstats.values())
    if exempt:
        log.info("Federal lines marked exempt from Davis-Bacon: federal wage rate not checked",
                 extra={"details": {"lines": exempt, "setting": "payroll.federal_exempt_pattern in config.yaml"}})
        audit.append(("Exempt lines", f"{exempt} federal line(s) marked exempt from Davis-Bacon (Classification, Craft "
                                      "or Jurisdiction says 'exempt') were skipped in the federal wage-rate validation: "
                                      "no 'no wage rate found' or rate-rule flags. Their other checks (hours, totals, "
                                      "duplicates) still ran. State lines are always validated."))
    if by_how:
        log.info("Payroll names matched to the wage data", extra={"details": dict(by_how.most_common())})
        audit.append(("Name matches", names_breakdown(by_how)[len("names: "):].replace(" - listed in the Comments sheet", "")
                      + ". Each name, the rate used and how it was matched are listed in the Comments sheet."
                      + (" Lines matched on the classification name only (under another craft) are flagged S11."
                         if by_how.get("classification name only") else "")))
    for jtext, js in jstats.items():
        log.info("Jurisdiction resolved", extra={"details": {
            "jurisdiction": jtext or "(empty)", "lines": js["lines"], "wage_type": js["kind"] or "unknown (state)",
            "named_decision": js["asked"]["wage_decision"] or "-", "named_county": js["asked"]["county"] or "-",
            "applied": ", ".join(f"{k}: {n}" for k, n in js["applied"].most_common()) or "no rate found"}})
        audit.append(("Jurisdiction", f"'{jtext or '(empty)'}': {js['lines']} line(s) checked as {js['kind'] or 'state (not stated)'}"
                      + (f", wage determination {js['asked']['wage_decision']}" if js["asked"]["wage_decision"] else "")
                      + (f", county {js['asked']['county']}" if js["asked"]["county"] else "")
                      + "; rates used: " + (", ".join(f"{k} ({n})" for k, n in js["applied"].most_common()) or "none found")))
        if js.get("project_county"):
            pc_name = project_cfg(cfg, project_dir.name).get("county")
            audit.append(("County", f"'{jtext or '(empty)'}' names no county: {js['project_county']} state line(s) were "
                                    f"checked against {pc_name} County rates, the project's county "
                                    f"(projects.{project_dir.name}.county in config.yaml)."))
        if js.get("residential"):
            audit.append(("Residential rates", f"'{jtext or '(empty)'}': {js['residential']} state line(s) named with a "
                                               "commercial craft were checked against the residential classification "
                                               f"of the same name (projects.{project_dir.name}.prefer_residential); "
                                               "lines with no residential classification use the commercial one."))
        if js.get("rates_date"):
            audit.append(("State rates date", f"'{jtext or '(empty)'}': {js['rates_date']} state line(s) were checked "
                                              f"against the L&I rates in effect on {rates_date_of(cfg, project_dir)}, "
                                              f"the project's state rates date (projects.{project_dir.name}."
                                              "state_rates_date in config.yaml, the bid due date)."))
        if js.get("unclear"):
            kd = js["kind"] or "state"
            fix = (f"set projects.{project_dir.name}.state_rates_date (the bid due date) in config.yaml, or name the "
                   f"determination in projects.{project_dir.name}.jurisdiction_map" if kd == "state" else
                   f"name the determination in projects.{project_dir.name}.jurisdiction_map in config.yaml")
            g_ = (f"Jurisdiction '{jtext or '(empty)'}' names no wage determination and the project has several {kd} "
                  f"ones ({', '.join(sorted(js['unclear_decs']))}); {js['unclear']:,} line(s) were checked against the "
                  f"one in effect in each line's week, which may not be the one the contract uses (flagged S14). "
                  f"To fix: {fix}")
            gaps.append(g_)
            audit.append(("Data gap", g_))
    no_data, unclear_no_data = Counter(), Counter()     # all; of them, lines whose Jurisdiction doesn't say the type
    for js in jstats.values():
        if js.get("no_data"):
            no_data[js["kind"] or "state"] += js["no_data"]
            if js.get("override"):
                unclear_no_data[js["kind"] or "state"] += js["no_data"]
    for kd, n in no_data.items():
        if kd in getattr(book, "untyped_kinds", ()):     # it has some, but without a Jurisdiction: not usable
            said, unclear = n - unclear_no_data[kd], unclear_no_data[kd]
            if said:
                g_ = (f"{said:,} {kd} line(s) could not be checked: the project's {kd} wage rates don't say they "
                      f"are {kd} rates (no Jurisdiction column). Upload the {kd} wage rates again with their "
                      "Jurisdiction column")
                gaps.append(g_)
                audit.append(("Data gap", g_))
            if unclear:                                  # their pass as this type (verification N1)
                g_ = (f"{unclear:,} line(s) whose Jurisdiction doesn't say federal, state or union could not be "
                      f"checked as {kd}: the project's {kd} wage rates don't say they are {kd} rates (no "
                      f"Jurisdiction column). Upload the {kd} wage rates again with their Jurisdiction column")
                gaps.append(g_)
                audit.append(("Data gap", g_))
            continue
        g_ = (f"{n:,} {kd} line(s) but the project has no {kd} wage rates, so they could not be checked against a "
              f"wage rate (S02). Upload the {kd} wage determination(s) to the project's Uploads folder"
              + (f" or list them in projects.{project_dir.name}.federal_wage_determinations in config.yaml"
                 if kd == "federal" else ""))
        gaps.append(g_)
        audit.append(("Data gap", g_))
    n_lines = sum(1 for _, v in rows if any(x not in (None, "") for x in v))
    # per line and wage type checked (an override line is checked once per type, and no_data counts it per pass)
    no_rate = len({(i.row, i.ctx.get("kind") or "state") for i in issues if i.rule == "S02"}) - sum(no_data.values())
    checks = n_lines - len(unusable) + extra_checks     # lines left out by the upload checks have no rate check
    pct = float(((cfg["payroll"].get("data_gaps") or {}).get("no_rate_lines_percent", 1.0)))
    if no_rate > 0 and checks and no_rate * 100.0 / checks >= pct:
        g_ = (f"{no_rate:,} of {checks:,} line(s){' checks' if extra_checks else ''} "
              f"({no_rate * 100.0 / checks:.1f}%) had no wage rate for their "
              "classification and date, so they could not be checked against one (S02; listed in the Comments "
              "sheet). Add the missing wage data, or map payroll names to the wage data under name_matches in "
              f"config.yaml (failure from {pct:g}% of lines: payroll.data_gaps.no_rate_lines_percent)")
        gaps.append(g_)
        audit.append(("Data gap", g_))
    by_rule = Counter(f"{i.severity}:{i.rule}" for i in issues)
    log.info("Payroll checks done", extra={"details": {"lines": len(rows), **dict(sorted(by_rule.items()))}})
    nf = Counter((i.ctx.get("kind"), i.ctx.get("craft"), i.ctx.get("classification")) for i in issues if i.rule == "S02")
    for (k, craft, cls), n in nf.most_common(25):
        log.debug("No wage rate for classification", extra={"details": {"kind": k, "craft": craft, "classification": cls, "lines": n}})
    unknown = sum(1 for i in issues if i.rule == "S08" and i.col == "jurisdiction")
    if unknown:
        log.warning("Lines whose Jurisdiction text didn't say federal/state/union were treated as state",
                    extra={"details": {"lines": unknown, "setting": "payroll.jurisdiction_patterns in config.yaml"}})
        audit.append(("Jurisdiction", f"{unknown} line(s) whose Jurisdiction didn't say federal, state or union were "
                                      "checked as state."))
    off = [r for r, key in RULE_KEYS.items() if cfg["payroll"]["rules"].get(key, {}).get("severity") == "off"] + \
          [r for r, key in SUSPECT_KEYS.items() if cfg["payroll"]["suspect_checks"].get(key) is False]
    if off:
        audit.append(("Settings", "Checks turned off in config.yaml (not run on any line): "
                                  + ", ".join(f"{r} {RULE_TITLES.get(r, r)}" for r in sorted(off))))
    # AI review of judgment-type orange flags (agent/ai_review.py): never stops the payroll review
    from . import ai_review
    names = sorted({str(v[cols["employee"]]) for _, v in rows
                    if "employee" in cols and cols["employee"] < len(v) and v[cols["employee"]]})
    try:
        ai = ai_review.review(project_dir, cfg, issues, kinds={kd for kd, _ in book.by_key}, titles=RULE_TITLES,
                              names=names)
    except Exception as e:                                       # noqa: BLE001
        log.warning("AI review failed: %s", err_text(e), exc_info=True)
        ai = {"patterns": [], "lines": [f"WARNING ai review: failed ({err_text(e)[:160]})"], "audit": [], "shadow": True,
              "mode": "local"}
    ai_review.annotate_issues(ai)
    audit += ai["audit"]
    _apply_marks(ws, issues, cols, cfg["payroll"]["colors"])
    ws.freeze_panes = ws.cell(hrow + 1, 1)
    meta = {"file": src.name, "run": run_timestamp(tz),
            "rows": sum(1 for _, v in rows if any(x not in (None, "") for x in v)),
            "wages": ", ".join(f"{k} ({'loaded' if any(kd == k for kd, _ in book.by_key) else 'none'})" for k in cfg["wages"]["kinds"]),
            "colors": cfg["payroll"]["colors"],
            "jurisdictions": "; ".join(
                f"'{j or '(empty)'}' ({js['lines']} lines) -> {js['kind'] or 'state?'}"
                + (": " + ", ".join(k for k, _ in js["applied"].most_common(3)) if js["applied"] else ": no rate found")
                for j, js in jstats.items()),
            "matches": matches, "exempt": exempt}
    _comments_sheet(wb, ws, issues, cols, header, meta)
    try:
        ai_review.decisions_sheet(wb, ai)
    except Exception as e:                                       # noqa: BLE001 - never lose the review over it
        log.warning("The Review decisions sheet could not be written: %s", err_text(e), exc_info=True)
    _audit_sheet(wb, meta["run"], audit)
    out = Paths(project_dir).reviewed / reviewed_name(src, cfg, run_timestamp(tz))
    from .safesave import save_workbook
    save_workbook(wb, out)           # never half-written, even if the run is killed
    log.debug("Reviewed payroll written", extra={"details": {"type": out.suffix.lower()}})   # no name: #56
    reds = sum(1 for i in issues if i.severity == "red")
    return out, meta["rows"], reds, len(issues) - reds, Counter(i.rule for i in issues), by_how, exempt, gaps, ai["lines"]


def wage_rates_problem(project_dir: Path, cfg):
    """The words of wage_rates_reason (None when the wage rates can be used)."""
    return wage_rates_reason(project_dir, cfg)[1]


def wage_rates_reason(project_dir: Path, cfg):
    """(code, words), or (None, None): the code is the structured run log's reason (docs/run_log.md). Why the project's wage rates can't be used for a payroll check, or None. Its own wage rates file (the
    Federal, State or Union table in Files/Wage rates/, layout.rate_file) must hold at least one rate whose wage type
    was given: a Jurisdiction column in the uploaded wage rates file, or an official determination (DOL / SAM.gov,
    L&I). (The firm-wide union reference doesn't count: it is not this project's wage determination.) David, Nov 7,
    2026."""
    from .wages import read_table
    rates = typed = 0
    for kind in cfg["wages"]["kinds"]:
        f = Paths(project_dir).rate_file(kind)
        try:
            got = read_table(f) if f.is_file() else []
        except Exception:                      # unreadable: not a usable rates file
            continue
        rates += len(got)
        typed += sum(1 for r in got if r.get("jurisdiction"))
    if not rates:
        from .uploads import uploads_of
        if any(f.suffix.lower() in (".xlsx", ".xlsm") for f in uploads_of(project_dir, "wages")):
            return ("no_jurisdiction_column", "Jurisdiction column not found in the wage rates file (it was left in "
                    "Uploads: none of its rates says State, Federal or Union)")
        return "no_wage_rates", "no wage rates file in the project (Files/Wage rates)"
    if not typed:
        return ("no_jurisdiction_column", "Jurisdiction column not found in the wage rates file (none of its rates "
                "says State, Federal or Union)")
    return None, None


def run_payroll_review(project_dir: Path, cfg, state, llm=None, force=False):
    """Review every payroll export in Uploads/ (oldest first): write 'Payroll/<date> <name> - REVIEWED.xlsx'
    and move the export to Payroll/Archived/<ts>/. An earlier review of the same export moves there too. With nothing new uploaded, the last payroll is reviewed again only if
    the wage data changed since (or with force)."""
    tz = cfg.get("timezone", "UTC")
    wage_stamp = {k: state.get(f"wages_{k}_updated") for k in cfg["wages"]["kinds"]}
    todo = inbox_payrolls(project_dir)
    log.debug("Payroll exports in Uploads", extra={"details": {"files": file_kinds(todo) or "(empty)",
                                                    "wage_data_versions": wage_stamp}})
    rerun = False
    if not todo:
        prev = last_reviewed_input(project_dir, state)
        last = state.get("payroll_last") or {}
        if prev and (force or last.get("wages") != wage_stamp):
            todo, rerun = [prev], True
            log.info("Re-reviewing the last payroll", extra={"details": {
                "reason": "forced" if force else "wage data changed since the last review"}})
        elif prev:
            return f"payroll: no new payroll uploaded; {prev.name} already reviewed with the current wage data (use force to redo)"
        else:
            return "payroll: no new payroll uploaded, nothing to review"
    # No wage rates file for the project at all: nothing to check the pay against, so the payroll isn't reviewed
    # (it stays in Uploads and is reviewed once the wage rates arrive) and both emails say so (David, Nov 7, 2026)
    code, problem = wage_rates_reason(project_dir, cfg)
    if problem:
        for src in todo:
            events.emit("file", file_n=events.next_file_n(project_dir.name), file_kind="payroll",
                        file_ext=events.file_ext(src), outcome="skipped", reason=code)
        return (f"payroll: SKIPPED - {problem}; the payroll was not "
                + ("checked again (the last review stands)" if rerun else "checked and stays in Uploads"))
    book = load_wagebook(project_dir, cfg)
    from collections import Counter
    log.debug("Wage rates loaded", extra={"details": dict(Counter(k for k, _ in book.by_key))})
    if not book:
        log.warning("No wage data in use for this project; every line will show 'no wage rate found'",
                    extra={"details": {"fix": "upload the wage determinations to the project's Uploads folder and run again"}})
    from .layout import PAYROLL, REVIEWED_DIR, archive_files
    Paths(project_dir).reviewed.mkdir(parents=True, exist_ok=True)
    lines = []
    for src in todo:
        file_ev = {"file_n": events.next_file_n(project_dir.name), "file_kind": "payroll",
                   "file_ext": events.file_ext(src)}
        t0 = time.monotonic()
        try:
            out, n, reds, oranges, by_rule, by_how, exempt, gaps, ai_lines = _review_one(src, project_dir, cfg, book, llm, tz)
        except PayrollSkipped as e:
            events.emit("file", **file_ev, outcome="skipped", reason=e.code,
                        duration_ms=int((time.monotonic() - t0) * 1000))
            log.info("Payroll not checked: its lines' wage type can't be told" if e.code == "wage_type_unknown" else
                     "Payroll not checked: an upload check", extra={"details": {"reason": e.code, "why": str(e)}})
            lines.append(f"payroll: SKIPPED - {e}; the payroll was not checked"
                         + (" again (the last review stands)" if rerun else " and stays in Uploads"))
            continue
        except Exception as e:
            events.emit("file", **file_ev, outcome="error", error_type=type(e).__name__,
                        duration_ms=int((time.monotonic() - t0) * 1000))
            log.error("Could not review a payroll file (%s); it stays in the inbox: %s", src.suffix.lower(),
                      type(e).__name__, exc_info=True,      # (no file name, and the error's text can hold one: #56)
                      extra={"details": {"hint": "check that it opens in Excel and is an LCPtracker payroll export"}})
            lines.append(f"ERROR (payroll): could not review {src.name} ({e}); it stays in Uploads")
            continue
        try:
            archive_files(project_dir, [f for f in earlier_reviews(project_dir, src, cfg) if f != out], PAYROLL, move=True)
            if rerun:
                archived = src
            else:
                archived = archive_files(project_dir, [src], PAYROLL, move=True)[src]
            state.set("payroll_last", {"file": src.name, "archived_input": str(archived.relative_to(project_dir)),
                                       "wages": wage_stamp, "output": out.name, "hash": file_hash(archived)})
        except Exception as e:                     # the task stops here as before; the run log says so (no text)
            events.emit("file", **file_ev, outcome="error", error_type=type(e).__name__,
                        duration_ms=int((time.monotonic() - t0) * 1000))
            raise
        events.emit("file", **file_ev, outcome="reviewed", rows=n, flags=dict(by_rule),
                    duration_ms=int((time.monotonic() - t0) * 1000))
        lines.append(f"payroll: {'re-' if rerun else ''}reviewed {src.name} ({n} lines): {reds} rule-breaking, "
                     f"{oranges} suspected -> {REVIEWED_DIR}/{out.name}")
        if by_rule:
            lines.append("payroll:   " + rule_breakdown(by_rule))
        if by_how:
            lines.append("payroll:   " + names_breakdown(by_how))
        if by_rule.get("S02"):
            lines.append(f"payroll:   no wage rate: {by_rule['S02']:,} line(s) could not be checked against a wage rate (S02)")
        if exempt:
            lines.append(f"payroll:   exempt: {exempt:,} federal line{'s' if exempt != 1 else ''} marked exempt from "
                         "Davis-Bacon; federal wage rate not checked (state lines still checked)")
        lines += [l if l.startswith("payroll:") else l.replace("WARNING ai review: ", f"WARNING ai review: {src.name}: ", 1)
                  for l in ai_lines]
        lines += [f"GAP payroll: {src.name}: {g_}" for g_ in gaps]
    return "\n".join(lines)
