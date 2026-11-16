"""One-time conversion of items written before October 22, 2026 by earlier versions, which used Pacific Time and
named timestamped items <artifact>__<ts>. Everything is converted to UTC and the yyyy-MM-dd-HH-mm-ss__<artifact_name>
pattern. Only items still in the old format are touched, so running it again does nothing."""
import csv
import datetime as dt
import json
import re
from pathlib import Path
from zoneinfo import ZoneInfo

from .util import TS_FORMAT, TS_REGEX, UTC, log

PACIFIC = ZoneInfo("America/Los_Angeles")
NAME_MAP = {"new_wages_data": "inbox", "current_wages_data": "current", "input": "inbox", "output": "reviewed",
            "compliance_tracker": "tracker", "intent_affidavit_data": "lni_data"}
OLD_DIR = re.compile(rf"^(?P<name>[A-Za-z][A-Za-z_]*?)__(?P<ts>{TS_REGEX})$")
OLD_IN_TEXT = re.compile(rf"\b(?P<name>[A-Za-z][A-Za-z_]*?)__(?P<ts>{TS_REGEX})\b")
OLD_TASK_LOG = re.compile(rf"^(?P<ts>{TS_REGEX})_log\.txt$")
OLD_RUN_SUMMARY = re.compile(rf"^run__(?P<ts>{TS_REGEX})\.json$")


def pt_to_utc(ts):
    """'2026-10-22-09-38-33' (Pacific) -> '2026-10-22-16-38-33' (UTC)."""
    return dt.datetime.strptime(ts, TS_FORMAT).replace(tzinfo=PACIFIC).astimezone(UTC).strftime(TS_FORMAT)


def pt_text_to_utc(s):
    """'2026-10-22 09:45', '2026-10-22 09:45:00' or '2026-10-22 09:45:00.123' (Pacific) -> UTC pattern."""
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", TS_FORMAT):
        try:
            return dt.datetime.strptime(s, fmt).replace(tzinfo=PACIFIC).astimezone(UTC).strftime(TS_FORMAT)
        except ValueError:
            continue
    return s


def new_name(old):
    m = OLD_DIR.match(old)
    if not m:
        return None
    return f"{pt_to_utc(m['ts'])}__{NAME_MAP.get(m['name'], m['name'])}"


def convert_text(s):
    """Rewrite old folder names mentioned inside text (log messages, paths)."""
    if not isinstance(s, str):
        return s
    s = OLD_IN_TEXT.sub(lambda m: f"{pt_to_utc(m['ts'])}__{NAME_MAP.get(m['name'], m['name'])}", s)
    return re.sub(rf"\b({TS_REGEX})_log\.txt\b", lambda m: f"{pt_to_utc(m[1])}__log.txt", s)


def _rename(path: Path, target: Path, changes):
    if target.exists():
        target = target.with_name(target.name + "__2")
    path.rename(target)
    changes.append(f"{path.name} -> {target.name}")
    return target


def _convert_log_csv(path: Path):
    rows = list(csv.reader(open(path, newline="", encoding="utf-8")))
    if not rows or rows[0][0] != "timestamp_pt":
        return False
    rows[0][0] = "timestamp_utc"
    for r in rows[1:]:
        if r:
            r[0] = pt_text_to_utc(r[0])
            if len(r) > 2 and re.fullmatch(TS_REGEX, r[2] or ""):
                r[2] = pt_to_utc(r[2])
            r[:] = [convert_text(x) for x in r]
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    return True


def _state(p):
    sp = p / ".agent_state.json"
    try:
        return json.loads(sp.read_text() or "{}")
    except Exception:
        return {}


def convert_project(p: Path):
    if _state(p).get("legacy_converted"):
        return []
    changes, pending = [], False
    # 1. archive folders and unused/: <name>__<ts> -> <ts utc>__<name>
    parents = [p / "tasks" / "payroll" / "archive", p / "tracker" / "archive", p / "tasks" / "intent" / "archive", p / "unused"] + \
              [k / "archive" for k in (p / "tasks" / "payroll" / "wages").glob("*") if k.is_dir()]
    for parent in parents:
        if not parent.is_dir():
            continue
        for d in sorted(parent.iterdir()):
            nn = new_name(d.name) if d.is_dir() else None
            if nn:
                _rename(d, parent / nn, changes)
    # 2. task logs: <ts>_log.txt -> <ts utc>__log.txt, entries converted
    for logs in (p / "tasks" / "payroll" / "logs", p / "tasks" / "payroll" / "wages" / "logs", p / "tasks" / "intent" / "logs"):
        if not logs.is_dir():
            continue
        for f in sorted(logs.iterdir()):
            m = OLD_TASK_LOG.match(f.name)
            if m:
                g = _rename(f, logs / f"{pt_to_utc(m['ts'])}__log.txt", changes)
                _convert_log_csv(g)
    # 3. run state
    sp = p / ".agent_state.json"
    if sp.is_file():
        st = json.loads(sp.read_text() or "{}")
        if st and st.get("timestamps") != "utc":
            for k, v in list(st.items()):
                if (k.startswith("wages_") and k.endswith("_updated") or k == "intent_last") and isinstance(v, str) \
                        and re.fullmatch(TS_REGEX, v):
                    st[k] = pt_to_utc(v)
            pl = st.get("payroll_last") or {}
            if pl.get("archived_input"):
                pl["archived_input"] = "/".join(new_name(x) or x for x in pl["archived_input"].split("/"))
            if isinstance(pl.get("wages"), dict):
                pl["wages"] = {k: (pt_to_utc(v) if isinstance(v, str) and re.fullmatch(TS_REGEX, v) else v)
                               for k, v in pl["wages"].items()}
            st["timestamps"] = "utc"
            sp.write_text(json.dumps(st, indent=2, default=str))
            changes.append(".agent_state.json timestamps -> UTC")
    # 4. L&I snapshot
    snap = p / "tasks" / "intent" / "lni_data" / "lni_snapshot.json"
    if snap.is_file():
        d = json.loads(snap.read_text())
        if "saved_pt" in d:
            d = {("saved_utc" if k == "saved_pt" else k): (pt_text_to_utc(v) if k == "saved_pt" else v) for k, v in d.items()}
            snap.write_text(json.dumps(d, indent=1))
            changes.append("lni_snapshot.json -> saved_utc")
    # 5. wage tables (sources sheet) and reviewed payrolls (Comments tab), both written by the agent
    import openpyxl
    for cur in (p / "tasks" / "payroll" / "wages").glob("*/current/current_wage_rates.xlsx"):
        wb = openpyxl.load_workbook(cur)
        if "sources" in wb.sheetnames and wb["sources"]["A1"].value == "merged_at_pt":
            ws = wb["sources"]
            ws["A1"] = "merged_at_utc"
            for (c,) in ws.iter_rows(min_row=2, max_col=1):
                if isinstance(c.value, str) and re.fullmatch(TS_REGEX, c.value):
                    c.value = pt_to_utc(c.value)
            wb.save(cur)
            changes.append(f"{cur.parent.parent.name}/current/{cur.name}: sources -> UTC")
    for rv in (p / "tasks" / "payroll" / "reviewed").glob("*.xlsx"):
        wb = openpyxl.load_workbook(rv)
        if "Comments" in wb.sheetnames and "(Pacific Time)" in str(wb["Comments"]["A2"].value or ""):
            c = wb["Comments"]["A2"]
            c.value = re.sub(TS_REGEX, lambda m: pt_to_utc(m[0]), c.value, count=1).replace("(Pacific Time)", "(UTC)")
            wb.save(rv)
            changes.append(f"tasks/payroll/reviewed/{rv.name}: Comments run time -> UTC")
    # 6. tracker Review Log (edited losslessly, skipped while the tracker is open in Excel)
    for tr in (p / "tracker").glob("*.xlsx"):
        if tr.name.startswith("~$"):
            continue
        from .xlsx_patch import XlsxPatcher
        xp = XlsxPatcher(tr)
        if "Review Log" not in xp.sheet_names() or xp.cell_text("Review Log", "A1") != "Run (PT)":
            continue
        if (tr.parent / f"~${tr.name}").exists():
            log.warning("A tracker is open in Excel; its Review Log times will be converted to UTC on a later run")
            pending = True
            continue
        xp.rewrite_column("Review Log", "A", lambda v, r: "Run (UTC)" if r == 1 else
                          (pt_to_utc(v) if v and re.fullmatch(TS_REGEX, v) else v))
        xp.save(tr)
        changes.append(f"tracker/{tr.name}: Review Log times -> UTC")
    if not pending:
        st = _state(p)
        st["timestamps"] = "utc"
        st["legacy_converted"] = True
        (p / ".agent_state.json").write_text(json.dumps(st, indent=2, default=str))
    if changes:                              # (each change names a file: counted only, #56)
        log.info("[%s] converted to UTC naming: %d item(s)", p.name, len(changes))
    return changes


def convert_root(root: Path):
    changes = []
    logs = Path(root) / "run_logs"
    if logs.is_dir():
        for f in sorted(logs.iterdir()):
            m = OLD_RUN_SUMMARY.match(f.name)
            if m:
                g = _rename(f, logs / f"{pt_to_utc(m['ts'])}__run_summary.json", changes)
                d = json.loads(g.read_text())
                if "run_pt" in d:
                    d = {("run_utc" if k == "run_pt" else k): (pt_to_utc(v) if k == "run_pt" else v) for k, v in d.items()}
                d = json.loads(convert_text(json.dumps(d)))
                g.write_text(json.dumps(d, indent=2))
                continue
            m = OLD_TASK_LOG.match(f.name)
            if m:
                g = _rename(f, logs / f"{pt_to_utc(m['ts'])}__run_log.txt", changes)
                _convert_log_csv(g)
        mid = logs / "model_ids.json"
        if mid.is_file():
            d = json.loads(mid.read_text())
            touched = False
            for v in d.values():
                if isinstance(v, dict) and "first_seen_pt" in v:
                    v["first_seen_utc"] = pt_text_to_utc(v.pop("first_seen_pt"))
                    touched = True
            if touched:
                mid.write_text(json.dumps(d, indent=2))
                changes.append("model_ids.json -> UTC")
    if changes:
        log.info("[run_logs] converted to UTC naming: %d item(s)", len(changes))
    return changes
