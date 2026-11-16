"""The CVA's evaluation harness: run a checkout end to end on invented projects with known answers, and score it.

    python tools/run_evals.py                          # build, run this checkout, score, print the report
    python tools/run_evals.py --json evals.json        # the same, and the numbers as JSON
    python tools/run_evals.py --code /tmp/cva-main/code   # another checkout (the answer key is this checkout's)
    python tools/run_evals.py --make /tmp/eval-snap     # only build the input snapshot and its answer_key.json
    python tools/run_evals.py --score /tmp/eval/run --snapshot /tmp/eval/snapshot   # score an earlier run again

**Invented data only** (tools/eval_cases.py builds it; HANDOFF section 6). The report quotes the run's own lines
(file names, skip reasons): never point ``--score`` at a run on staff files. Each run uses tools/snapshot_run.py's
pipeline: the checkout's tokenizer on every uploaded file (what AWS would see), then the checkout's agent, every task,
offline (no AI review, no network, no keys, the emails saved as previews), its printed output kept in the run folder.

What is measured, against the answer key (eval_cases.py explains it):
  - **run health**: no task crashed, no "ERROR (...)" line, no run error, both emails written (as previews);
  - each payroll file's **outcome**: reviewed, or skipped (left in Uploads) for the reason the key gives, that reason
    also in the administrator's email (also seen: "not recognized" as a payroll, "error", or "missing": gone from
    Uploads unreviewed, e.g. held back by the tokenizer);
  - each line's **flags**, read from the reviewed workbook's cell notes (the marks staff see): a seeded error is
    **detected** when every rule it must raise was raised on its line(s); a flag outside the line's "must" and "may"
    rules is a **false flag** (any flag on a clean control line, the header or a title row is one); the notes' words
    where the key asks (e.g. which wage type a check used);
  - per rule: seeded, detected, false flags; overall: detection rate, clean controls, flag precision, run time;
  - **privacy**: nothing in the run folder (what AWS would hold, everything the agent wrote, the emails decoded part
    by part, spreadsheets and archives opened, file and folder names, both steps' printed output) may hold an invented
    worker's name or employee ID: the tokenizer must have replaced each one. A file that can't be read counts as a
    find (fail closed).

Exit code: 0 when everything matches the key (a known gap's misses are reported, not counted), 1 when anything else
differs, 2 on bad arguments, 3 when the run itself failed. With $GITHUB_STEP_SUMMARY set (GitHub Actions) the report
is also added to the job's summary page.

  - the **intent/affidavit check** (Eval G, H): the checkout's own check runs against invented L&I records that
    tools/eval_lni.py serves on 127.0.0.1 as L&I's live lookup would (the source in use); each tracker's outcome, each
    line's marks (column and kind of finding, from the agent's notes), the Review Log's contractors at L&I but not in
    the tracker, the not-due-yet counts, and that L&I is never asked about a tracker that isn't checked.

  - the **structured run log** (run_logs/<ts>__events.jsonl; the run turns run_log.events on): every line passes the
    schema, one run_start and one run_end, one run ID, one task line per task run, per project the file lines'
    outcomes and skip reason codes against what the harness saw and their flag counts against the reviewed workbooks'
    notes, nothing dropped (a checkout without agent/events.py: "not measured").
Not measured yet (docs/evals.md, "Next"): the AI review's suggestions (it needs the Claude API; offline here), the
wage updates, the intent check's data.wa.gov source and its run-wide "not checked" cases (unit tests cover them).
"""
from __future__ import annotations

import argparse
import datetime as dt
import email
import gzip
import io
import json
import os
import re
import sys
import tempfile
import time
import zipfile
from collections import defaultdict
from email import policy
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
for p in (HERE, AGENT):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

import eval_cases                                                       # noqa: E402
import eval_lni                                                         # noqa: E402
import snapshot_run                                                     # noqa: E402
from agent.payroll import RULE_TITLES                                   # noqa: E402

FLAG = re.compile(r"\[(?:RULE BROKEN|SUSPECTED)\] ([RS]\d\d)\b")       # payroll._apply_marks' note lines
# The invented workers (eval_cases): their surname, "Evaltest nnn", and their employee IDs, as text
WORKER = re.compile(re.escape(eval_cases.SURNAME) + "|" + re.escape(eval_cases.GIVEN) + r" \d{3}\b|(?<![\d.])"
                    + str(eval_cases.ID_BASE // 1000) + r"\d{3}(?!\d)"
                    + "".join("|" + re.escape(w) for w in getattr(eval_cases, "NEVER_LEARNT", "").split()), re.I)
# #53: the tokenizer's fail-closed settings, on for the run (the never-learnt name of Eval L needs them)
TOKENIZER_VNEXT = {"fail_closed_tables": True, "opaque_sheet_names": True, "opaque_file_names": True}
UNLISTED = "(a row the key doesn't list)"
AGENT_SHEETS = {"Comments", "Comments (agent)", "Audit", "Audit (agent)", "Review log", "Review Log"}


# ------------------------------------------------------------------------------------------------ reading a run
def reviewed_for(project_dir: Path, file_name: str, suffix: str):
    """The reviewed workbook made from this upload ('<date> <stem><suffix>.xlsx', payroll.reviewed_name), or None.
    (Named from the stem alone: eval_cases refuses two uploads with one stem in a project.)"""
    stem = Path(file_name).stem
    pat = re.compile(rf"^\d{{4}}-\d{{2}}-\d{{2}} {re.escape(stem)}{re.escape(suffix)}\.xlsx$")
    found = sorted(f for f in (project_dir / "Payroll").glob("*.xlsx") if pat.match(f.name))
    return found[-1] if found else None


def notes_by_row(workbook: Path) -> dict:
    """{sheet row: (set of rule IDs, the notes' text)} from the agent's notes on the payroll sheet (the first sheet
    that isn't one the agent adds). Rows without a rule are left out."""
    wb = openpyxl.load_workbook(workbook)
    sheets = [ws for ws in wb.worksheets if ws.title not in AGENT_SHEETS]
    rules, texts = defaultdict(set), defaultdict(list)
    for row in (sheets[0].iter_rows() if sheets else ()):
        for c in row:
            if c.comment is not None:
                rules[c.row].update(FLAG.findall(c.comment.text or ""))
                texts[c.row].append(c.comment.text or "")
    return {r: (s, "\n".join(texts[r])) for r, s in rules.items() if s}


def run_summary(results: Path) -> dict:
    found = sorted((results / "run_logs").glob("*__run_summary.json"))
    return json.loads(found[-1].read_text()) if found else {}


def eml_text(path: Path) -> str:
    """Every text part of a saved email, decoded."""
    msg = email.message_from_bytes(path.read_bytes(), policy=policy.default)
    parts = []
    for part in msg.walk():
        if part.get_content_maintype() == "text":
            try:
                parts.append(part.get_content())
            except Exception:
                parts.append((part.get_payload(decode=True) or b"").decode("utf-8", "replace"))
    return "\n".join(parts)


def health(results: Path, summary: dict) -> list:
    """What went wrong with the run itself: crashed tasks, ERROR lines, run errors, emails not written."""
    problems = []
    if not summary:
        return ["no run summary was written (the agent didn't finish)"]
    for project, tasks in (summary.get("details") or {}).items():
        for task, d in (tasks or {}).items():
            if d.get("crashed") or d.get("errors"):
                problems.append(f"{project}: the {task} task " + ("crashed" if d.get("crashed") else "had errors"))
    for project, lines in (summary.get("projects") or {}).items():
        problems += [f"{project}: {l[:200]}" for l in lines if l.startswith("ERROR (")]
    problems += [f"run error: {str(e)[:200]}" for e in summary.get("run_errors") or []]
    if not str(summary.get("email", "")).startswith("dry run"):
        problems.append(f"the emails were not written: {str(summary.get('email'))[:200]}")
    for who in ("staff", "admin"):
        if not list((results / "run_logs").glob(f"*__email_preview_{who}.eml")):
            problems.append(f"no {'staff' if who == 'staff' else 'administrator'} email preview")
    return problems


def aws_file_name(results: Path, project: str, file_name: str) -> str:
    """The upload's name on the AWS side: the same, unless the tokenizer renamed it (#53, opaque_file_names;
    snapshot_run.py records it by a hash of the staff path, so the record holds no name)."""
    import hashlib
    f = results.parent / "aws_names.json"
    names = json.loads(f.read_text()) if f.is_file() else {}
    aws = names.get(hashlib.sha256(f"Projects/{project}/Uploads/{file_name}".encode()).hexdigest())
    return aws.rsplit("/", 1)[-1] if aws else file_name


def outcome_of(results: Path, project: str, file_name: str, suffix: str, summary: dict):
    """('reviewed' | 'skipped' | 'not recognized' | 'error' | 'missing', reviewed workbook or None, reason). The
    agent gives a skip's reason per project ("payroll: SKIPPED - ..."), not per file: with two uploads skipped for
    different reasons in one project, each would get the first reason (the set has none such)."""
    pdir = results / "Projects" / project
    lines = (summary.get("projects") or {}).get(project) or []
    file_name = aws_file_name(results, project, file_name)
    out = reviewed_for(pdir, file_name, suffix)
    if out is not None:
        return "reviewed", out, ""
    why = next((l for l in lines if l.startswith("ERROR (payroll)")), "")
    if why:
        return "error", None, why
    if (pdir / "Uploads" / file_name).is_file():
        why = next((l for l in lines if l.startswith("WARNING upload:") and file_name in l
                    and "not recognized" in l), "")
        if why:
            return "not recognized", None, why
        skipped = ((summary.get("details") or {}).get(project) or {}).get("payroll", {}).get("skipped") or []
        why = next((l for l in list(skipped) + lines if "payroll: SKIPPED" in l), "")
        return "skipped", None, why
    return "missing", None, "no reviewed workbook, and the upload is gone (held back by the tokenizer?)"


# ------------------------------------------------------------------------------------------- the run log
RUN_LOG_OUTCOME = {"reviewed": "reviewed", "skipped": "skipped", "not recognized": "not_recognized",
                   "error": "error"}                       # the harness's outcome -> the file line's (missing: none)
TS_RE = re.compile(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ")


def flags_in(workbook: Path) -> dict:
    """{rule: number of flags} in the reviewed workbook's notes on the payroll sheet (one per note line, as the agent
    counts its issues)."""
    wb = openpyxl.load_workbook(workbook)
    sheets = [ws for ws in wb.worksheets if ws.title not in AGENT_SHEETS]
    out = defaultdict(int)
    for row in (sheets[0].iter_rows() if sheets else ()):
        for c in row:
            if c.comment is not None:
                for rule in FLAG.findall(c.comment.text or ""):
                    out[rule] += 1
    return dict(out)


def events_module(code: Path = None):
    """The events schema (agent/events.py) of the checkout being scored (``code``: its code/ folder; default: this
    one), loaded on its own (it imports nothing of the agent at load time); None when that checkout has none. It runs
    in the harness's process: score only checkouts of this repository (the harness trusts them anyway)."""
    import importlib.util
    f = (code or AGENT.parent) / "compliance_agent" / "agent" / "events.py"
    if not f.is_file():
        return None
    spec = importlib.util.spec_from_file_location("cva_checkout_events", f)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run_log_on(results: Path) -> bool:
    """Whether the run's own config.yaml turned the structured run log on (run_log.events: true)."""
    import yaml
    cfg = yaml.safe_load((results / "config.yaml").read_text()) or {}
    return (cfg.get("run_log") or {}).get("events") is True


def score_run_log(results: Path, key: dict, files: list, summary: dict, ev) -> dict:
    """The structured run log (run_logs/<ts>__events.jsonl; docs/run_log.md) against the run itself: the lines'
    schema (``ev``: the scored checkout's agent/events.py module; None: not measured); one run_start first and one
    run_end last, one run ID; run_start's tasks and project count, and run_end's outcome and counts, as the run
    summary gives them; one task line per task run, with the summary's status for it (what the emails say); per
    project, the file lines' outcomes and skip reason codes against what the harness itself saw, file numbers 1..n,
    and each reviewed file's flags against one reviewed workbook's notes (file by file, not just summed); nothing
    dropped, no failed write. Not measured: the ai_review line's numbers (the AI review is off in these runs)."""
    if ev is None:
        return {"measured": False, "problems": [], "lines": 0}
    found = sorted((results / "run_logs").glob("*__events.jsonl"))
    if len(found) != 1:
        return {"measured": True, "lines": 0,
                "problems": [f"{len(found)} events files in run_logs (expected 1)"]}
    problems = []
    projects_dir = results / "Projects"
    project_names = {p.name for p in projects_dir.iterdir() if p.is_dir()} if projects_dir.is_dir() else set()
    lines = []
    for i, raw in enumerate(found[0].read_text(encoding="utf-8").splitlines(), 1):
        try:
            d = json.loads(raw)
        except ValueError:
            problems.append(f"line {i} is not JSON")
            continue
        if not isinstance(d, dict):
            problems.append(f"line {i} is not an object")
            continue
        event = d.get("event")
        if not isinstance(event, str) or event not in ev.ALLOWED:
            problems.append(f"line {i}: unknown event")
            continue
        base = {"v", "ts", "run_id", "event", "dropped"}
        extra = sorted(k for k in d if k not in base | ev.ALLOWED[event])
        bad = sorted(k for k in d if k in ev.ALLOWED[event] and not (
            (k == "project" and isinstance(d[k], str) and d[k] in project_names) or
            (k == "outcome" and isinstance(d[k], str) and d[k] in ev.OUTCOMES.get(event, ())) or
            (k not in ("project", "outcome") and ev.SCHEMA[k](d[k]))))
        if d.get("v") != ev.V or not TS_RE.fullmatch(str(d.get("ts"))) or \
                not ev.RUN_ID_RE.fullmatch(str(d.get("run_id"))) or extra or bad:
            problems.append(f"line {i} ({event}) fails the schema" + (f": fields {', '.join(extra + bad)}"
                                                                       if extra or bad else ""))
            continue                         # (verification N1: not used further, whatever its types)
        if event in ("task", "file", "ai_review") and not (isinstance(d.get("project"), str)
                                                           and d["project"] in project_names):
            problems.append(f"line {i} ({event}) names no project of the run")
            continue
        lines.append(d)
    kinds = [d["event"] for d in lines]
    if kinds.count("run_start") != 1 or kinds[:1] != ["run_start"]:
        problems.append("the first line isn't the only run_start")
    if kinds.count("run_end") != 1 or kinds[-1:] != ["run_end"]:
        problems.append("the last line isn't the only run_end")
    if len({d.get("run_id") for d in lines}) > 1:
        problems.append("more than one run ID")
    start = next((d for d in lines if d["event"] == "run_start"), {})
    end = next((d for d in reversed(lines) if d["event"] == "run_end"), {})
    summary_projects = summary.get("projects") or {}
    if start and (start.get("tasks") != sorted(summary.get("tasks") or []) or
                  start.get("projects") != len(summary_projects)):
        problems.append("run_start's tasks or project count differ from the run summary")
    # one task line per task run, with the run summary's status
    want = sorted((p, t, "crashed" if d_.get("crashed") else d_.get("status"))
                  for p, ts_ in (summary.get("details") or {}).items() for t, d_ in ts_.items()
                  if t in ((summary.get("durations") or {}).get(p) or {}))
    got = sorted((str(d.get("project")), str(d.get("task")), str(d.get("outcome"))) for d in lines
                 if d["event"] == "task")
    if got != want:
        problems.append(f"task lines: {len(got)} for {len(want)} tasks run"
                        + ("" if len(got) != len(want) else " (a task's project, name or outcome differs)"))
    # per project: file lines against what the harness saw
    by_project = defaultdict(list)
    for f in files:
        by_project[f["file"].split("/", 1)[0]].append(f)
    file_lines = defaultdict(list)
    for d in lines:
        if d["event"] == "file":
            file_lines[d["project"]].append(d)
    suffix = _output_suffix(results)
    for project in sorted(set(by_project) | set(file_lines)):
        mine = file_lines.get(project, [])
        want_out = sorted(RUN_LOG_OUTCOME[f["got"]] for f in by_project.get(project, ()) if f["got"] in RUN_LOG_OUTCOME)
        got_out = sorted(str(d.get("outcome")) for d in mine)
        if want_out != got_out:
            problems.append(f"{project}: file lines say {', '.join(got_out) or 'nothing'}; the run shows "
                            f"{', '.join(want_out) or 'nothing'}")
        if sorted(d.get("file_n") for d in mine if isinstance(d.get("file_n"), int)) != list(range(1, len(mine) + 1)):
            problems.append(f"{project}: file numbers aren't 1 to {len(mine)}")
        reasons = [d.get("reason") for d in mine if d.get("outcome") == "skipped"]
        if any(r is None for r in reasons):
            problems.append(f"{project}: a skipped file line has no reason code")
        for f in by_project.get(project, ()):
            code = (key["files"].get(f["file"]) or {}).get("reason_code")
            if f["got"] == "skipped" and code and code not in reasons:
                problems.append(f"{project}: no skipped file line with reason {code}")
        books = []
        for f in by_project.get(project, ()):
            if f["got"] == "reviewed":
                wb = reviewed_for(projects_dir / project, aws_file_name(results, project, f["file"].split("/", 1)[1]),
                                  suffix)
                books.append(sorted((flags_in(wb) if wb else {}).items()))
        logged = [sorted((d.get("flags") or {}).items()) for d in mine if d.get("outcome") == "reviewed"]
        if sorted(books) != sorted(logged):
            diff = sorted({r for b in books + logged for r, _ in b
                           if sum(dict(b_).get(r, 0) for b_ in books) != sum(dict(l_).get(r, 0) for l_ in logged)})
            problems.append(f"{project}: flag counts differ from the reviewed workbook's notes"
                            + (f" ({', '.join(diff)})" if diff else " (file by file)"))
    if end:
        if end.get("dropped") != 0 or end.get("events_failed") != 0:
            problems.append(f"run_end: {end.get('dropped')} field(s) dropped, {end.get('events_failed')} failed write(s)")
        if end.get("files") != sum(len(v) for v in file_lines.values()):
            problems.append("run_end's file count doesn't match the file lines")
        errors = sum(1 for v in summary_projects.values() for l in v if str(l).startswith("ERROR"))
        if end.get("outcome") != ev.RUN_OUTCOMES.get(summary.get("outcome")) or \
                end.get("projects") != len(summary_projects) or end.get("errors") != errors or \
                end.get("warnings") != int(summary.get("warning_count") or 0):
            problems.append("run_end's outcome or counts differ from the run summary")
    return {"measured": True, "problems": problems, "lines": len(lines),
            "events": {k: kinds.count(k) for k in sorted(set(kinds))}}


# ------------------------------------------------------------------------------------------- the intent check
TRACKER_AUTHORS = ("Compliance Validation Agent", "Compliance Agent")  # the agent's note authors (filekinds tracker)
NOTE_HEAD = re.compile(r"^(Compliance Validation Agent|Compliance Agent):\s*$")
TRACKER_AGENT_SHEETS = {"review log", "audit"}                          # sheets the check writes (config.yaml intent)
CELL = re.compile(r"^([A-Z]{1,3})(\d+)$")
NUMBER = re.compile(r"(?<![\d.])\d{6,8}(?![\d.])")                       # the IDs in a note or a Review Log line


def intent_outcome(lines) -> tuple:
    """('checked' | 'not applicable' | 'not checked' | 'no tracker' | 'skipped' | 'error' | 'none' | 'other', the
    line) from a project's run lines (intent.py's first "intent: ..." line)."""
    err = next((l for l in lines if l.startswith("ERROR (intent)")), None)
    if err:
        return "error", err
    first = next((l for l in lines if l.startswith("intent:")), None)
    if first is None:
        return "none", ""
    for pat, what in ((r"intent: \d+ invalid/missing value", "checked"), (r"intent: not applicable", "not applicable"),
                      (r"intent: NOT CHECKED", "not checked"), (r"intent: no compliance tracker", "no tracker"),
                      (r"intent: SKIPPED", "skipped")):
        if re.match(pat, first):
            return what, first
    return "other", first


def read_tracker(tracker: Path, header_row: int, red: str) -> dict:
    """What the check left in a tracker, as staff see it:
      notes   {row: {"role:kind": text}} from the agent's notes on the contractor table (the first tab): every line of
              an agent note but its "Compliance Validation Agent:" head is a mark (eval_lni.KINDS; else "unknown");
      log     {row: {"role:kind": text}} from the Review Log's rows that name a cell (the complete list: a cell that
              already had a note gets its message there only);
      lni_only the Review Log's contractors at L&I but not in the tracker (no cell, Field "contractor");
      red     {row: {role}} cells of the contractor table filled with the check's red;
      stray   agent notes on any other tab ("Sheet!A1")."""
    wb = openpyxl.load_workbook(tracker)
    ws = wb.worksheets[0]
    roles = {c.column_letter: eval_lni.ROLE_OF.get(str(c.value or "").strip(), f"column {c.column_letter}")
             for c in ws[header_row]}
    out = {"notes": defaultdict(dict), "log": defaultdict(dict), "lni_only": [], "red": defaultdict(set), "stray": []}
    for s in wb.worksheets:
        if s is not ws and s.title.lower() in TRACKER_AGENT_SHEETS:
            continue
        for row in s.iter_rows():
            for c in row:
                if s is ws and (getattr(c.fill, "fgColor", None) is not None and c.fill.fill_type == "solid"
                                and str(c.fill.fgColor.rgb or "").upper() == red.upper()):
                    out["red"][c.row].add(roles.get(c.column_letter, "?"))
                if c.comment is None or (c.comment.author or "") not in TRACKER_AUTHORS:
                    continue
                if s is not ws:
                    out["stray"].append(f"{s.title}!{c.coordinate}")
                    continue
                for ln in (c.comment.text or "").splitlines():
                    ln = ln.strip()
                    if not ln or NOTE_HEAD.match(ln):
                        continue
                    ln = ln[2:] if ln.startswith("- ") else ln
                    out["notes"][c.row][f"{roles.get(c.column_letter, '?')}:{eval_lni.kind_of(ln)}"] = ln
    log = next((w for w in wb.worksheets if w.title.lower() == "review log"), None)
    for r in (log.iter_rows(min_row=2, values_only=True) if log is not None else ()):
        r = list(r) + [None] * 8
        m = CELL.match(str(r[2] or "").strip())
        if m:                                       # Cell, Contractor, Field, Tracker value, L&I value, Issue
            text = str(r[7] or "")
            out["log"][int(m[2])][f"{roles.get(m[1], '?')}:{eval_lni.kind_of(text)}"] = f"{text} (L&I: {r[6] or ''})"
        elif not r[2] and str(r[4] or "").lower() == "contractor" and r[3]:
            out["lni_only"].append(str(r[3]))
    out["lni_only"].sort()
    return out


NOT_DUE_WORDS = {"not_started": "line(s) not started with no intent",
                 "no_affidavit_yet": "completed line(s) with no affidavit at L&I yet",
                 "filed_not_at_lni": "line(s) ticked as affidavit filed with none under its intent at L&I yet"}


def not_due_counts(lines) -> dict:
    """The run line "intent: not due yet, not marked: 1 line(s) not started ..." as {kind: count}."""
    line = next((l for l in lines if l.startswith("intent: not due yet")), "")
    out = {}
    for k, words in NOT_DUE_WORDS.items():
        m = re.search(r"(\d+) " + re.escape(words), line)
        if m:
            out[k] = int(m[1])
    return out


def record_numbers(text: str) -> list:
    """The record numbers a finding's words give, in their order: those in the sentence itself, else (when it has
    none, as "Intent belongs to a different L&I project") those in its "(L&I: ...)" value."""
    head, sep, tail = text.rpartition(" (L&I: ")
    if not sep:
        head, tail = text, ""
    return NUMBER.findall(head) or NUMBER.findall(tail)


def summary_counts(lines) -> dict:
    """The run's intent counts: findings marked (other than hiring), hiring mismatches, L&I contractors not in the
    tracker (0 when the line isn't there)."""
    def n(pat):
        m = next((re.match(pat, l) for l in lines if re.match(pat, l)), None)
        return int(m[1]) if m else 0
    return {"marked": n(r"intent: (\d+) invalid/missing value"), "hiring": n(r"intent: (\d+) hiring contractor"),
            "lni_only": n(r"intent: (\d+) contractor\(s\) in L&I's records")}


def _red(results: Path) -> str:
    import yaml
    cfg = yaml.safe_load((results / "config.yaml").read_text()) or {}
    return str((((cfg.get("payroll") or {}).get("colors") or {}).get("red")) or "FFFF4D4D")


def score_trackers(results: Path, key: dict, summary: dict, requests) -> tuple:
    """(trackers, tracker lines, request problems): each tracker's outcome, run lines, Review Log, red fill and
    not-due counts against the key; each line's marks (in the Review Log, in the notes, with the record numbers the
    key asks for); requests L&I mustn't get, or that weren't L&I's (``requests``: the stand-in's record, or None)."""
    trackers_key = key.get("trackers") or {}
    gap_ids = {g["id"]: g for g in key.get("known_gaps", [])}
    red = _red(results) if trackers_key else ""
    trackers, lines = [], []
    for project, exp in sorted(trackers_key.items()):
        plines = (summary.get("projects") or {}).get(project) or []
        got, said = intent_outcome(plines)
        path = results / "Projects" / project / "Tracker" / exp["file"]
        t = read_tracker(path, exp["header_row"], red) if path.is_file() else \
            {"notes": {}, "log": {}, "lni_only": [], "red": {}, "stray": []}
        checked = exp["outcome"] == "checked"
        problems = []
        if got != exp["outcome"]:
            problems.append(f"the check ended \"{got}\" ({said[:160]}), not \"{exp['outcome']}\"")
        elif exp.get("reason_contains") and exp["reason_contains"] not in said:
            problems.append(f"its line doesn't say \"{exp['reason_contains']}\" ({said[:160]})")
        if not path.is_file():
            problems.append("the tracker is gone from Tracker/")
        if t["stray"]:
            problems.append("agent notes outside the contractor table: " + ", ".join(t["stray"][:5]))
        if checked and got == "checked":
            if "WARNING" in said:
                problems.append(f"a warning on a clean check: {said[:200]}")
            want, have = exp.get("summary") or {}, summary_counts(plines)
            # a known gap's line ending as the gap says changes the counts by its marks today
            now = dict(want)
            for e in exp["lines"].values():
                g = gap_ids.get(e.get("known_gap"))
                if g:
                    cur = list(g.get("lines_currently") or ())
                    for k, hiring in (("marked", False), ("hiring", True)):
                        if k in now:
                            now[k] += sum(m.startswith("hiring:") == hiring for m in cur) - \
                                sum(m.startswith("hiring:") == hiring for m in e["must"])
            for k in sorted(want):
                if have.get(k) != want[k] and have.get(k) != now[k]:
                    problems.append(f"the run counts {have.get(k)} {k.replace('_', ' ')}, expected {want[k]}")
            if t["lni_only"] != exp["lni_only"]:
                problems.append(f"Review Log's contractors not in the tracker: {t['lni_only'] or 'none'}, expected "
                                f"{exp['lni_only'] or 'none'}")
            due = not_due_counts(plines)
            if due != exp["not_due"]:
                problems.append(f"not due yet: {due or 'none'}, expected {exp['not_due'] or 'none'}")
        # red fill: on each cell the Review Log names, and nowhere else
        logged = {(r, m.split(":", 1)[0]) for r, ms in t["log"].items() for m in ms}
        filled = {(r, role) for r, roles in t["red"].items() for role in roles}
        if filled - logged:
            problems.append("red cells without a finding: " + ", ".join(f"row {r} {c}" for r, c in sorted(filled - logged)))
        if logged - filled:
            problems.append("findings without the red fill: " + ", ".join(f"row {r} {c}" for r, c in sorted(logged - filled)))
        trackers.append({"project": project, "expected": exp["outcome"], "got": got, "ok": not problems,
                         "problems": problems})
        for r, e in sorted(exp["lines"].items(), key=lambda x: int(x[0])):
            row = int(r)
            must = set(e["must"])
            in_log, in_notes = t["log"].get(row, {}), t["notes"].get(row, {})
            raised = set(in_log) | set(in_notes)
            want_notes = set() if e.get("staff_note") else must
            gap = gap_ids.get(e.get("known_gap"))
            entry = {"tracker": project, "row": row, "case": e["case"], "area": e["area"], "must": sorted(must),
                     "raised": sorted(raised), "missed": [], "extra": [], "words": [], "not_due": e.get("not_due"),
                     "known_gap": gap["id"] if gap else None, "tolerated": False}
            if got == exp["outcome"] or not checked:        # a tracker that wasn't checked: the outcome is the difference
                entry["missed"] = sorted([m for m in must - set(in_log) if checked] +
                                         [f"{m} (note)" for m in want_notes - set(in_notes) if checked])
                entry["extra"] = sorted((raised - must) | {f"{m} (note)" for m in set(in_notes) - want_notes
                                                            if m in must})
                for m, ids in (e.get("ids") or {}).items():         # each place staff read must give them
                    for where, texts in (("Review Log", in_log), ("note", in_notes)):
                        if m in texts and record_numbers(texts[m])[:len(ids)] != ids:
                            entry["words"].append(f"{m}: {where} gives {record_numbers(texts[m]) or 'no number'}, "
                                                  f"not {ids}")
                if gap and (entry["missed"] or entry["extra"]) and raised == set(gap.get("lines_currently") or ()):
                    entry["tolerated"] = True
            lines.append(entry)
        for r in sorted((set(t["log"]) | set(t["notes"])) - {int(x) for x in exp["lines"]}):
            raised = sorted(set(t["log"].get(r, {})) | set(t["notes"].get(r, {})))
            lines.append({"tracker": project, "row": r, "case": UNLISTED, "area": "-", "must": [], "raised": raised,
                          "missed": [], "extra": raised, "words": [], "not_due": None, "known_gap": None,
                          "tolerated": False})
    asked = []
    if trackers_key:
        if requests is None:
            asked = ["(what L&I was asked wasn't recorded: lni_requests.json is missing)"]
        else:
            never = {i for t in trackers_key.values() for i in t.get("never_ask") or ()}
            for q in requests:
                if not q.get("answered"):
                    asked.append(f"refused: {q.get('method')} {str(q.get('url'))[:120]}")
                elif q.get("id") in never:
                    asked.append(f"{q['op']} {q['id']} (a tracker that isn't checked)")
            asked = sorted(set(asked))
    return trackers, lines, asked


# ------------------------------------------------------------------------------------------------------ privacy
def _hits(data: bytes) -> bool:
    if WORKER.search(data.decode("utf-8", "ignore")):
        return True
    if b"\x00" in data:                                                 # UTF-16 text
        return any(WORKER.search(data.decode(enc, "ignore")) for enc in ("utf-16-le", "utf-16-be"))
    return False


def _search_blob(data: bytes, depth: int = 0) -> bool:
    """True when the bytes, or anything packed inside them (zip members and their names, gzip, email parts and
    attachments' names), hold an invented worker's name or ID."""
    if depth > 3:
        return _hits(data)
    if data[:4] == b"PK\x03\x04":
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            return any(_hits(n.encode()) or _search_blob(z.read(n), depth + 1) for n in z.namelist())
    if data[:2] == b"\x1f\x8b":
        return _search_blob(gzip.decompress(data), depth + 1)
    if _hits(data):
        return True
    head = data[:2000].lower()
    if b"mime-version:" in head or b"content-type:" in head:            # an email: decode every part
        msg = email.message_from_bytes(data, policy=policy.default)
        for part in msg.walk():
            if part.is_multipart():
                continue
            name = part.get_filename() or ""
            payload = part.get_payload(decode=True) or b""
            if _hits(name.encode()) or _search_blob(payload, depth + 1):
                return True
    return False


def worker_names_found(run: Path) -> list:
    """Everything in the run folder (tokenized/: what AWS would hold; results/: what the agent wrote; step_output/:
    what both steps printed; home/) that still holds an invented worker's name or employee ID. A file that can't be
    read is reported too (fail closed). Returns 'path (where)' strings, relative to ``run``."""
    found = []
    for f in sorted(run.rglob("*")):
        rel = f.relative_to(run).as_posix()
        if _hits(rel.encode()):
            found.append(f"{WORKER.sub('<name>', rel)} (name)")         # the report never prints the name found
            continue
        if not f.is_file():
            continue
        try:
            if _search_blob(f.read_bytes()):
                found.append(f"{rel} (content)")
        except Exception as e:
            found.append(f"{rel} (could not be read: {type(e).__name__})")
    return found


# #53: each place Eval L puts the never-learnt name, and where to look for it in the AWS-side copy (tokenized/)
NEVER_PLACES = ("unread payroll column", "second sheet", "sheet name", "document property", "file name",
                "folder name", "wage table note")


def never_learnt_places(run: Path) -> dict:
    """{place: True if the never-learnt name is still there in the AWS-side copy} for Eval L (NEVER_PLACES); {} when
    the set has no such case. A file that can't be found or read counts as the name being there (fail closed)."""
    name = getattr(eval_cases, "NEVER_LEARNT", "")
    pdir = run / "tokenized" / "Projects" / "Eval L - Names never learnt"
    if not name or not pdir.is_dir():
        return {}
    has = lambda text: name.split()[-1].lower() in str(text or "").lower()
    out = {p: True for p in NEVER_PLACES}
    up = pdir / "Uploads"
    files = [f for f in up.iterdir() if f.is_file()] if up.is_dir() else []
    out["file name"] = any(has(f.name) for f in files) or len(files) != 2
    out["folder name"] = any(has(d.name) for d in pdir.rglob("*") if d.is_dir()) or \
        not any(d.name.endswith(" notes") for d in (pdir / "Files").iterdir() if d.is_dir())
    for f in files:
        try:
            wb = openpyxl.load_workbook(f)
        except Exception:
            continue
        if f.name.startswith("Week 41"):
            ws = wb.worksheets[0]
            out["unread payroll column"] = any(has(c.value) for row in ws.iter_rows() for c in row)
            out["second sheet"] = any(has(c.value) for w in wb.worksheets[1:] for row in w.iter_rows() for c in row)
            out["sheet name"] = any(has(w.title) for w in wb.worksheets) or len(wb.worksheets) != 3
            props = wb.properties
            out["document property"] = any(has(getattr(props, k, "")) for k in ("creator", "lastModifiedBy", "title"))
        else:
            out["wage table note"] = any(has(c.comment.text) or has(c.comment.author)
                                         for row in wb.active.iter_rows() for c in row if c.comment) or \
                not any(c.comment for row in wb.active.iter_rows() for c in row)
    return out


# ------------------------------------------------------------------------------------------------------ scoring
def score(results: Path, key: dict, run: Path = None, requests=None, events=False) -> dict:
    """Compares a run with the answer key. ``results``: the agent's folder; ``run``: the whole run folder for the
    privacy check (default: the parent of ``results``); ``requests``: what the stand-in for L&I was asked (default:
    ``run``/lni_requests.json, when there). Returns the numbers and the differences."""
    run = run or results.parent
    if requests is None and (run / "lni_requests.json").is_file():
        requests = json.loads((run / "lni_requests.json").read_text())
    suffix = _output_suffix(results)
    summary = run_summary(results)
    admin_mail = next(iter(sorted((results / "run_logs").glob("*__email_preview_admin.eml"))), None)
    admin_text = eml_text(admin_mail) if admin_mail else ""
    gap_of = {f: g for g in key.get("known_gaps", []) for f in g["files"]}
    files, cases, per_rule = [], [], defaultdict(lambda: {"seeded": 0, "detected": 0, "false_flags": 0})
    for rel, exp in sorted(key["files"].items()):
        project, name = rel.split("/", 1)
        got, wb, why = outcome_of(results, project, name, suffix, summary)
        gap = gap_of.get(rel)
        want = exp.get("reason_contains") or ""
        ok = got == exp["outcome"] and want in why and (not want or want in admin_text)
        files.append({"file": rel, "expected": exp["outcome"], "got": got, "ok": ok, "reason": why,
                      "expected_reason": want, "reason_in_admin_email": bool(want) and want in admin_text,
                      "known_gap": gap["id"] if gap else None,
                      # a known gap: its current (wrong) outcome is tolerated, any other isn't
                      "tolerated": bool(gap) and not ok and got == gap.get("currently")})
        # Flags are judged on files the key expects reviewed (a file reviewed when it should have been skipped is
        # one difference, its outcome, not one per flag); a file not reviewed misses every seeded rule.
        notes = notes_by_row(wb) if wb is not None and exp["outcome"] == "reviewed" else {}
        groups = {}
        for r, e in exp["lines"].items():
            groups.setdefault(e.get("group") or e["case"], {**e, "rows": []})["rows"].append(int(r))
        for e in sorted(groups.values(), key=lambda g: min(g["rows"])):
            must, may = set(e["must"]), set(e["may"])
            raised = set().union(*(notes.get(r, (set(), ""))[0] for r in e["rows"]))
            text = "\n".join(notes.get(r, (set(), ""))[1] for r in e["rows"])
            missed = must - raised if exp["outcome"] == "reviewed" else set()
            extra = raised - must - may
            words = [f'notes lack "{w}"' for w in e.get("notes_contain") or [] if must and w not in text] + \
                    [f'notes say "{w}"' for w in e.get("notes_absent") or [] if w in text]
            if exp["outcome"] != "reviewed" or got != "reviewed":
                words = []                                           # the outcome is the difference already
            # a known gap on lines (rows_currently): a line raising exactly today's rules is tolerated
            today = ((gap or {}).get("rows_currently") or {}).get(rel) or {}
            tolerated = bool(missed or extra) and len(e["rows"]) == 1 and str(e["rows"][0]) in today and \
                raised == set(today[str(e["rows"][0])])
            cases.append({"file": rel, "rows": sorted(e["rows"]), "case": e["case"], "area": e["area"],
                          "must": sorted(must), "raised": sorted(raised), "missed": sorted(missed),
                          "extra": sorted(extra), "notes": words, "known_gap": gap["id"] if gap else None,
                          "tolerated": tolerated})
            # a known gap's file: its misses aren't counted (they are the gap); its false flags are
            for rule in (must if gap is None else ()):
                per_rule[rule]["seeded"] += 1
                per_rule[rule]["detected"] += rule in raised
            for rule in (extra if not tolerated else ()):
                per_rule[rule]["false_flags"] += 1
        # flags on rows the key doesn't list (the header, a title row, an empty row): false flags too
        listed = {int(r) for r in exp["lines"]}
        for r, (raised, _) in sorted(notes.items()):
            if r not in listed:
                cases.append({"file": rel, "rows": [r], "case": UNLISTED, "area": "-", "must": [],
                              "raised": sorted(raised), "missed": [], "extra": sorted(raised), "notes": [],
                              "known_gap": gap["id"] if gap else None, "tolerated": False})
                for rule in raised:
                    per_rule[rule]["false_flags"] += 1
    trackers, tlines, asked = score_trackers(results, key, summary, requests)
    out = _totals(files, cases, per_rule, key, trackers, tlines)
    out["lni_asked_wrongly"] = asked
    out["lni_requests"] = None if requests is None else len(requests)
    out["failures"] += len(asked)
    out["health"] = problems = health(results, summary)
    if events is False:                     # not given: the run's own setting decides, with this checkout's schema
        events = events_module() if run_log_on(results) else None
    out["run_log"] = rl = score_run_log(results, key, files, summary, events)
    out["failures"] += len(rl["problems"])
    out["worker_names_on_aws_side"] = leaks = worker_names_found(run)
    out["totals"]["worker_names_on_aws_side"] = len(leaks)
    out["failures"] += len(leaks) + len(problems)
    out["never_learnt"] = places = never_learnt_places(run)
    out["failures"] += sum(places.values())
    return out


def _totals(files, cases, per_rule, key, trackers=(), tlines=()) -> dict:
    # gaps on payroll lines: each of their lines must end as the gap says (any other result is a difference)
    line_gaps = {g["id"] for g in key.get("known_gaps", []) if g.get("rows_currently")}
    counted = [c for c in cases if c["known_gap"] is None]
    seeded = [c for c in counted if c["must"]]
    controls = [c for c in counted if not c["must"] and c["case"] != UNLISTED]
    flags = sum(len(c["raised"]) for c in counted)
    # a known gap's false flags count too, unless the gap says the line ends so today (rows_currently)
    false_flags = sum(len(c["extra"]) for c in cases if not c.get("tolerated"))
    rule_instances = sum(len(c["must"]) for c in seeded)
    rule_hits = sum(len(c["must"]) - len(c["missed"]) for c in seeded)
    failures = [c for c in counted if c["missed"] or c["extra"] or c["notes"]] + \
               [c for c in cases if c["known_gap"] and not c.get("tolerated") and
                (c["extra"] or (c["known_gap"] in line_gaps and (c["missed"] or c["notes"])))] + \
               [f for f in files if not f["ok"] and not f["tolerated"]] + \
               [t for t in trackers if not t["ok"]] + \
               [ln for ln in tlines if (ln["missed"] or ln["extra"] or ln["words"]) and not ln["tolerated"]]
    gaps = []
    for g in key.get("known_gaps", []):
        mine = [c for c in cases if c["known_gap"] == g["id"]]
        fs = [f for f in files if f["known_gap"] == g["id"]]
        tl = [ln for ln in tlines if ln["known_gap"] == g["id"]]
        if tl:                                                   # a gap on tracker lines
            if any(ln["raised"] == [] and ln["missed"] == [] for ln in tl if ln["must"]):
                status = "not measured (the tracker wasn't checked)"
            elif not any(ln["missed"] or ln["extra"] for ln in tl):
                status = "passes now: close the gap (remove it from the key)"
            elif all(ln["tolerated"] or not (ln["missed"] or ln["extra"]) for ln in tl):
                status = "still open"
            else:
                status = "changed: " + ", ".join(f"{ln['tracker']} row {ln['row']} raised {ln['raised'] or 'nothing'}"
                                                 for ln in tl if (ln["missed"] or ln["extra"]) and not ln["tolerated"])
            gaps.append({**g, "status": status})
            continue
        # on a gap's lines (rows_currently) every line must end as today; elsewhere, only false flags differ
        off = [c for c in mine if not c.get("tolerated") and (c["extra"] or (g.get("rows_currently") and
                                                                              (c["missed"] or c["notes"])))]
        if all(f["ok"] for f in fs) and not any(c["missed"] or c["extra"] or c["notes"] for c in mine):
            status = "passes now: close the gap (remove it from the key)"
        elif all(f["ok"] or f["tolerated"] for f in fs) and not off:
            status = "still open"
        else:
            status = "changed: " + ", ".join([f"{f['file']} ended {f['got']}" for f in fs if not f["ok"] and
                                              not f["tolerated"]] +
                                             [f"{c['file']} row {c['rows'][0]} raised {c['raised'] or 'nothing'}"
                                              for c in off])
        gaps.append({**g, "status": status})
    counted_t = [ln for ln in tlines if ln["known_gap"] is None]
    t_seeded = [ln for ln in counted_t if ln["must"]]
    t_controls = [ln for ln in counted_t if not ln["must"] and ln["case"] != UNLISTED]
    t_marks = sum(len(ln["raised"]) for ln in counted_t)
    t_false = sum(len(ln["extra"]) for ln in tlines if not ln["tolerated"])
    t_words = sum(len(ln["words"]) for ln in tlines)
    return {
        "files": files, "cases": cases, "per_rule": {k: dict(v) for k, v in sorted(per_rule.items())},
        "trackers": list(trackers), "tracker_lines": list(tlines),
        "known_gaps": gaps,
        "totals": {
            "files": len(files), "files_as_expected": sum(f["ok"] for f in files),
            "files_known_gap": sum(1 for f in files if f["tolerated"]),
            "seeded_cases": len(seeded), "seeded_cases_detected": sum(1 for c in seeded if not c["missed"]),
            "seeded_rules": rule_instances, "seeded_rules_detected": rule_hits,
            "control_cases": len(controls), "control_cases_clean": sum(1 for c in controls if not c["raised"]),
            "flags_raised": flags, "false_flags": false_flags,
            "flag_precision": round((flags - false_flags) / flags, 4) if flags else None,
            "detection_rate": round(rule_hits / rule_instances, 4) if rule_instances else None,
            "trackers": len(trackers), "trackers_as_expected": sum(t["ok"] for t in trackers),
            "tracker_findings_seeded": len(t_seeded),
            "tracker_findings_detected": sum(1 for ln in t_seeded if not ln["missed"]),
            "tracker_controls": len(t_controls),
            "tracker_controls_clean": sum(1 for ln in t_controls if not ln["raised"]),
            "tracker_not_due_lines": sum(1 for ln in t_controls if ln["not_due"]),
            "tracker_marks": t_marks, "tracker_false_marks": t_false, "tracker_wrong_numbers": t_words,
            "tracker_lines_known_gap": sum(1 for ln in tlines if ln["tolerated"]),
        },
        "failures": len(failures),
    }


def _output_suffix(results: Path) -> str:
    import yaml
    cfg = yaml.safe_load((results / "config.yaml").read_text()) or {}
    return ((cfg.get("payroll") or {}).get("output_suffix")) or " - REVIEWED"


# ------------------------------------------------------------------------------------------------------ report
def _pct(a, b):
    return f"{a} of {b}" + (f" ({100 * a / b:.0f}%)" if b else "")


def report(s: dict, meta: dict = None) -> str:
    t = s["totals"]
    meta = meta or {}
    verdict = "PASS: every result matches the answer key" if not s["failures"] else \
        f"FAIL: {s['failures']} difference(s) from the answer key"
    out = [f"## CVA evaluation: {verdict}", ""]
    if meta:
        out += [" · ".join(f"{k}: {v}" for k, v in meta.items()), ""]
    leaks = s["worker_names_on_aws_side"]
    out += [
        "- Run health: " + ("no task crashed, no errors, both emails written" if not s["health"] else
                            f"**{len(s['health'])} problem(s)**: " + "; ".join(s["health"][:5])),
        f"- Payroll files with the expected outcome: {_pct(t['files_as_expected'], t['files'])}"
        + (f" ({t['files_known_gap']} known gap)" if t["files_known_gap"] else ""),
        f"- Seeded errors detected (cases): {_pct(t['seeded_cases_detected'], t['seeded_cases'])}; "
        f"rule by rule: {_pct(t['seeded_rules_detected'], t['seeded_rules'])}",
        f"- Clean control cases with no flag: {_pct(t['control_cases_clean'], t['control_cases'])}",
        f"- Flags raised: {t['flags_raised']}; false flags: {t['false_flags']}"
        + (f"; flag precision {100 * t['flag_precision']:.1f}%" if t["flag_precision"] is not None else ""),
        f"- Run files still holding an invented worker's name or ID: {len(leaks)}"
        + (" (" + ", ".join(leaks[:5]) + ")" if leaks else ""),
    ]
    places = s.get("never_learnt") or {}
    if places:
        held = [p for p, v in places.items() if v]
        out.append(f"- A name never learnt (Eval L, #53), on the AWS side: in {len(held)} of {len(places)} places"
                     + (" (" + ", ".join(held) + ")" if held else " (" + ", ".join(places) + ": none)"))
    if t.get("trackers"):
        asked = s.get("lni_asked_wrongly") or []
        n_req = s.get("lni_requests")
        out += [
            f"- Intent/affidavit check, trackers with the expected outcome: {_pct(t['trackers_as_expected'], t['trackers'])}"
            " (checked against invented L&I records served by a local stand-in for L&I's live lookup)",
            f"- Tracker findings marked: {_pct(t['tracker_findings_detected'], t['tracker_findings_seeded'])}; correct "
            f"and not-yet-due lines with no mark: {_pct(t['tracker_controls_clean'], t['tracker_controls'])} "
            f"({t['tracker_not_due_lines']} not due yet)",
            f"- Tracker marks: {t['tracker_marks']}; false marks: {t['tracker_false_marks']}; findings with the wrong record "
            f"numbers: {t['tracker_wrong_numbers']}"
            + (f"; {t['tracker_lines_known_gap']} line(s) of a known gap" if t["tracker_lines_known_gap"] else ""),
            "- Network: requests refused or that L&I mustn't get: " + ("; ".join(asked[:5]) if asked else "none")
            + (f" ({n_req} requests, all to the stand-in for L&I)" if n_req is not None else ""),
        ]
    out += [
        "- Run log (one JSON line per step, schema-checked): " + (
            "not measured (this checkout has no structured run log)" if not s.get("run_log", {}).get("measured") else
            (f"{s['run_log']['lines']} lines, matches the run" if not s["run_log"]["problems"] else
             f"**{len(s['run_log']['problems'])} problem(s)**: " + "; ".join(s["run_log"]["problems"][:5]))),
        "- Known gaps (scored, misses not counted): " + (", ".join(f"{g['id']} (#{g['issue']}): {g['status']}"
                                                          for g in s["known_gaps"]) or "none"),
        "- Not measured here: the AI review (needs the Claude API), the wage updates; the intent check's other L&I "
        "source (data.wa.gov) and its run-wide \"not checked\" cases (L&I down or out of date: unit tests)",
        "", "| File | Expected | Got | |", "|---|---|---|---|"]
    for f in s["files"]:
        mark = "ok" if f["ok"] else ("known gap" if f["tolerated"] else "**DIFFERENT**")
        out.append(f"| {f['file']} | {f['expected']} | {f['got']} | {mark} |")
    out += ["", "| Rule | Meaning | Seeded | Detected | False flags |", "|---|---|---|---|---|"]
    for rule, v in s["per_rule"].items():
        out.append(f"| {rule} | {RULE_TITLES.get(rule, rule)} | {v['seeded']} | {v['detected']} | {v['false_flags']} |")
    bad = [c for c in s["cases"] if c["missed"] or c["extra"] or c["notes"]]
    if bad:
        out += ["", "| File | Row(s) | Case | Must | Raised | Missed | False flags | Notes | |",
                "|---|---|---|---|---|---|---|---|---|"]
        for c in bad:
            out.append(f"| {c['file']} | {', '.join(map(str, c['rows']))} | {c['case']} | {', '.join(c['must']) or '-'} | "
                       f"{', '.join(c['raised']) or '-'} | {', '.join(c['missed']) or '-'} | "
                       f"{', '.join(c['extra']) or '-'} | {'; '.join(c['notes']) or '-'} | "
                       f"{'known gap' if c['known_gap'] else ''} |")
    if s.get("trackers"):
        out += ["", "| Tracker | Expected | Got | |", "|---|---|---|---|"]
        for tr in s["trackers"]:
            out.append(f"| {tr['project']} | {tr['expected']} | {tr['got']} | "
                       + ("ok" if tr["ok"] else "**DIFFERENT**: " + "; ".join(tr["problems"])) + " |")
    tbad = [ln for ln in s.get("tracker_lines") or [] if ln["missed"] or ln["extra"] or ln["words"]]
    if tbad:
        out += ["", "| Tracker | Row | Case | Must | Marked | Missed | False marks | Numbers | |",
                "|---|---|---|---|---|---|---|---|---|"]
        for ln in tbad:
            out.append(f"| {ln['tracker']} | {ln['row']} | {ln['case']} | {', '.join(ln['must']) or '-'} | "
                       f"{', '.join(ln['raised']) or '-'} | {', '.join(ln['missed']) or '-'} | "
                       f"{', '.join(ln['extra']) or '-'} | {'; '.join(ln['words']) or '-'} | "
                       f"{'known gap' if ln['tolerated'] else ''} |")
    wrong = [f for f in s["files"] if not f["ok"]]
    if wrong:
        out += ["", "Files without the expected outcome:"]
        for f in wrong:
            said = []
            if f["expected_reason"]:
                said.append(f"the reason should say \"{f['expected_reason']}\""
                            + ("" if f["reason_in_admin_email"] else ", in the administrator's email too"))
            said.append(f"the run said: {f['reason'][:300]}" if f["reason"] else "the run gave no reason")
            out.append(f"- {f['file']}: " + "; ".join(said) + (" (known gap)" if f["tolerated"] else ""))
    return "\n".join(out) + "\n"


# --------------------------------------------------------------------------------------------------------- main
def make(dest: Path) -> dict:
    key = eval_cases.build(dest)
    (dest / "answer_key.json").write_text(json.dumps(key, indent=1, sort_keys=True))
    return key


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Run the CVA end to end on invented projects and score it.")
    ap.add_argument("--make", type=Path, help="only build the input snapshot (and answer_key.json) in this folder")
    ap.add_argument("--score", type=Path, help="score this earlier run folder (snapshot_run.py --out) again; "
                                               "invented runs only: the report quotes the run's lines")
    ap.add_argument("--snapshot", type=Path, help="with --score: the snapshot it ran on (for answer_key.json)")
    ap.add_argument("--code", type=Path, default=AGENT.parent, help="the checkout's code/ folder (default: this one; "
                                                                     "with --score: the one that made the run)")
    ap.add_argument("--keep", type=Path, help="keep the snapshot and the run in this (empty) folder")
    ap.add_argument("--json", type=Path, help="also write the numbers and differences to this JSON file")
    ap.add_argument("--agent-python", default=sys.executable)
    ap.add_argument("--tokenizer-python", default=sys.executable)
    ap.add_argument("--shipped-tokenizer", action="store_true",
                    help="run the tokenizer as tokenizer_config.yaml ships it, without turning on #53's settings "
                         "(Eval L's never-learnt name then reaches the AWS side: reported as a leak)")
    a = ap.parse_args(argv)
    if a.make:
        make(a.make)
        print(f"snapshot and answer key written to {a.make}")
        return 0
    if a.score:
        if not a.snapshot:
            ap.error("--score needs --snapshot")
        key = json.loads((a.snapshot / "answer_key.json").read_text())
        s = score(a.score / "results", key, run=a.score,
                  events=events_module(a.code.resolve()) if run_log_on(a.score / "results") else None)
        return _finish(s, {"run": "scored again from " + str(a.score)}, a.json)
    if a.keep and a.keep.exists() and any(a.keep.iterdir()):
        ap.error("--keep: the folder must be empty")
    with tempfile.TemporaryDirectory(prefix="cva-eval-") as tmp:
        base = a.keep or Path(tmp)
        snap, out = base / "snapshot", base / "run"
        key = make(snap)
        start = time.monotonic()
        lookup = eval_lni.FakeLookup().start()
        try:
            snapshot_run.run_checkout(a.code.resolve(), snap.resolve(), out.resolve(), a.agent_python,
                                      a.tokenizer_python, capture=True,
                                      config_extra={**lni_settings(),
                                                    ("run_log", "events"): True},
                                      proxy=lookup.proxy,
                                      tokenizer_vnext=None if a.shipped_tokenizer else TOKENIZER_VNEXT)
        except SystemExit as e:
            print(f"the run failed: {e}", file=sys.stderr)
            for f in sorted((out / "step_output").glob("*.txt")):           # invented data only: safe to show
                print(f"--- {f.name} (last lines)\n" + "\n".join(f.read_text(errors="replace").splitlines()[-30:]),
                      file=sys.stderr)
            return 3
        finally:
            lookup.stop()
            if out.is_dir():                    # every request either step made, in full (searched for names too)
                (out / "lni_requests.json").write_text(json.dumps(lookup.requests, indent=1))
        took = time.monotonic() - start
        s = score(out / "results", key, run=out, events=events_module(a.code.resolve()))
        meta = {"checkout": _describe(a.code), "run": f"{took:.0f} s",
                "date": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")}
        return _finish(s, meta, a.json)


def lni_settings() -> dict:
    """The intent check's settings for the run: L&I's live lookup (the source in use since Nov 10, 2026) at its
    production address in plain http (eval_lni.LNI_URL; the stand-in, as the run's proxy, answers it), without the pause
    between requests."""
    return {("intent", "lni_source"): "live", ("intent", "live", "url"): eval_lni.LNI_URL,
            ("intent", "live", "pause_seconds"): 0}


def _describe(code: Path) -> str:
    import subprocess
    try:
        sha = subprocess.run(["git", "-C", str(code), "rev-parse", "--short", "HEAD"], capture_output=True,
                             text=True, timeout=10).stdout.strip()
    except Exception:
        sha = ""
    return sha or str(code)


def _finish(s: dict, meta: dict, json_path) -> int:
    text = report(s, meta)
    print(text)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a") as f:
            f.write(text)
    if json_path:
        json_path.write_text(json.dumps({"meta": meta, **s}, indent=1, sort_keys=True, default=str))
    return 1 if s["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
