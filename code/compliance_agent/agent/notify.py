"""Run-outcome emails: one email after every task run (scheduled or manual), saying whether it succeeded.

Recipients, sender and transport come from notifications.config (next to config.yaml). When a run fails, fully or
partially, the email describes each failure and links to the log files that show it.

Outcome of a run:
  SUCCESS          every selected task finished for every project with no errors (warnings are listed, not failures)
  PARTIAL FAILURE  at least one task failed, logged errors, or was skipped for a project, but not all of them
  FAILED           nothing completed: every task failed, or the run stopped before or while the agent ran
  NOT RUN          the run did not start because another run was still in progress

Where the email is sent from:
  * runs started on this computer (desktop app or command line): at the end of run()
  * runs on AWS (EC2 timer, Lambda, desktop app's "Run on AWS", trigger.sh): by deploy/run_pipeline.sh after the
    results are pushed back to Google Drive, so the links open the log files on the shared drive. The pipeline sets
    COMPLIANCE_NOTIFY=defer so the agent itself doesn't send a second email.
An email that can't be sent never fails the run; the reason is logged.
"""
from __future__ import annotations

import datetime as dt
import html
import json
import logging
import os
import platform
import re
import shutil
import socket
import subprocess
import urllib.parse
from email.message import EmailMessage
from email.utils import formataddr, make_msgid, parseaddr
from pathlib import Path

import yaml

from .tracelog import TASK_DIRS, _project, _task
from .util import TS_FORMAT, err_text, shown_name

log = logging.getLogger("compliance.notify")

AGENT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG = AGENT_DIR / "notifications.config"
TASK_NAMES = {"wages": "Wage data merge", "payroll": "Payroll review", "intent": "Intent/affidavit validation"}
SKIP_MARKERS = ("SKIPPED",)                # a task line with this means the task didn't do its job for that project
DEFAULTS = {
    "enabled": True,
    "send_on": "all",                     # all | problems (partial failure, failed, not run, gaps): the staff email
    "admin_send_on": "all",               # the same for the administrator's email (Settings page, Nov 12, 2026)
    "to": [],                             # staff: the short email (agent/staff_email.py)
    "admin_to": [],                       # administrator: the full email with every detail
    "cc": [],
    "bcc": [],
    "reply_to": [],
    "from": "",
    "subject_prefix": "[Compliance Validation Agent]",
    "timezone": "America/Los_Angeles",    # times in the email are shown in this zone, with UTC alongside
    "transport": "ses",                   # ses | smtp | none
    "ses": {"region": "", "configuration_set": ""},
    "smtp": {"host": "", "port": 587, "security": "starttls", "username": "", "password_env": "SMTP_PASSWORD",
             "password_secret": ""},
    "links": {"google_drive": True, "local_paths": True},     # (s3_console: retired Nov 15, 2026, ignored)
    "attach_logs": "problems",            # all | problems | never
    "max_attachments_mb": 8,
    "max_issues_listed": 50,
}


# ----------------------------------------------------------------------------------------------- configuration
def _merge(base, extra):
    out = dict(base)
    for k, v in (extra or {}).items():
        out[k] = _merge(base[k], v) if isinstance(base.get(k), dict) and isinstance(v, dict) else v
    return out


def _addresses(value):
    if not value:
        return []
    if isinstance(value, str):
        value = value.replace(";", ",").split(",")
    return [a.strip() for a in value if a and str(a).strip()]


def load_notify_config(path=None, root=None):
    """notifications.config (YAML). COMPLIANCE_NOTIFY_CONFIG overrides the path; COMPLIANCE_NOTIFY_TO/CC/ADMIN_TO add
    addresses (comma-separated) without editing the file, e.g. for a test, or set by the deployment in the server's
    /etc/compliance-agent/env; COMPLIANCE_NOTIFY_FROM sets the sender. Placeholder addresses (@example.com) are
    dropped once a real one is given. With `root`, the Settings page's recipient lists replace these (#151,
    agent/admin_settings.py)."""
    path = Path(path or os.environ.get("COMPLIANCE_NOTIFY_CONFIG") or DEFAULT_CONFIG)
    data = {}
    if path.exists():
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    cfg = _merge(DEFAULTS, data)
    cfg["path"] = str(path)
    cfg["found"] = path.exists()
    for key in ("to", "admin_to", "cc", "bcc", "reply_to"):
        cfg[key] = _addresses(cfg.get(key))
    cfg["admin_to"] += [a for a in _addresses(os.environ.get("COMPLIANCE_NOTIFY_ADMIN_TO")) if a not in cfg["admin_to"]]
    cfg["admin_to"] = [a for a in cfg["admin_to"] if not _placeholder(a)]
    cfg["to"] += [a for a in _addresses(os.environ.get("COMPLIANCE_NOTIFY_TO")) if a not in cfg["to"]]
    cfg["cc"] += [a for a in _addresses(os.environ.get("COMPLIANCE_NOTIFY_CC")) if a not in cfg["cc"]]
    for key in ("to", "cc", "bcc"):
        real = [a for a in cfg[key] if not _placeholder(a)]
        if real or key != "to" or os.environ.get("COMPLIANCE_NOTIFY_TO"):
            cfg[key] = real
    if os.environ.get("COMPLIANCE_NOTIFY_FROM"):
        cfg["from"] = os.environ["COMPLIANCE_NOTIFY_FROM"]
    elif _placeholder(cfg["from"]) and cfg["to"] and not _placeholder(cfg["to"][0]):
        cfg["from"] = ""
    from .admin_settings import apply_notify
    apply_notify(cfg, root)
    if not cfg["from"]:
        cfg["from"] = cfg["to"][0] if cfg["to"] else ""
    return cfg


def _placeholder(addr) -> bool:
    return bool(re.search(r"@example\.(com|org|net)\b", str(addr or ""), re.I))


# ----------------------------------------------------------------------------------------- collecting problems
class IssueCollector(logging.Handler):
    """Keeps every WARNING/ERROR logged during a run, with the project and task it belongs to."""

    def __init__(self):
        super().__init__(level=logging.WARNING)
        self.items = []

    def emit(self, record):
        try:
            if record.name.startswith("compliance.notify"):
                return
            self.items.append({"level": record.levelname, "project": _project.get(), "task": _task.get(),
                               "run_wide": bool(getattr(record, "run_wide", False)),
                               "message": record.getMessage(),
                               "time_utc": dt.datetime.fromtimestamp(record.created, dt.timezone.utc).strftime(TS_FORMAT)})
        except Exception:
            pass

    def __enter__(self):
        logging.getLogger().addHandler(self)
        return self

    def __exit__(self, *exc):
        logging.getLogger().removeHandler(self)


def task_log_path(project_dir: Path, task, ts):
    from .layout import Paths
    return Paths(project_dir).task_log(task, ts)


def _words(text):
    return {w.strip(".,;:()'\"").lower() for w in text.split() if len(w) > 2}


def _same_issue(a, b):
    """True when two messages describe the same problem (e.g. the result line and the logged error)."""
    wa, wb = _words(a), _words(b)
    return bool(wa and wb) and len(wa & wb) / min(len(wa), len(wb)) >= 0.6


def _ordered_tasks(tasks):
    return [t for t in TASK_NAMES if t in tasks] + [t for t in tasks if t not in TASK_NAMES]


def _unique(items):
    out = []
    for i in items:
        if i not in out:
            out.append(i)
    return out


def classify(report, issues, projects):
    """Adds report['outcome'] and report['details'][project][task] = {status, errors, warnings, skipped, log}."""
    ts = report["run_utc"]
    details, counts = {}, {"ok": 0, "failed": 0, "errors": 0, "skipped": 0}
    for name, lines in report["projects"].items():
        pdir = projects.get(name)
        per = {}
        for task in (report.get("plan") or {}).get(name, report["tasks"]):     # a plan: only this project's tasks
            prefix = f"ERROR ({task}): "
            err_lines = [l[len(prefix):] for l in lines if l.startswith(prefix)]
            done_lines = [l for l in lines if l.startswith(f"{task}:")]
            skipped = [l for l in done_lines if any(m in l for m in SKIP_MARKERS)]
            logged_errs = [i["message"] for i in issues if i["project"] == name and i["task"] == task
                           and i["level"] in ("ERROR", "CRITICAL")]
            # errors logged by the task that its result lines don't already say
            extra_errs = [m for m in logged_errs if not m.startswith("Task failed:")
                          and not any(_same_issue(m, e) for e in err_lines)]
            errors = err_lines + extra_errs
            categorised = set(lines) | {prefix + e for e in err_lines}
            warns = [i["message"] for i in issues if i["project"] == name and i["task"] == task
                     and i["level"] == "WARNING" and i["message"] not in categorised]
            if err_lines and not [l for l in done_lines if l not in skipped]:
                status = "failed"            # the task stopped, or every item it tried failed
            elif errors:
                status = "errors"            # it finished, but some items failed
            elif skipped:
                status = "skipped"
            else:
                status = "ok"
            counts[status] += 1
            lp = task_log_path(pdir, task, ts) if pdir else None
            per[task] = {"status": status, "errors": errors,
                         "crashed": any(m.startswith("Task failed:") for m in logged_errs), "warnings": _unique(warns), "skipped": skipped,
                         "log": str(lp) if lp and lp.exists() else ""}
        layout = [l for l in lines if l.startswith("ERROR (layout)")]
        project_errs = [i["message"] for i in issues if i["project"] == name and not i["task"] and i["level"] in ("ERROR", "CRITICAL")]
        if layout or project_errs:
            per["_project"] = {"status": "errors", "errors": _unique([l[len("ERROR (layout): "):] for l in layout]
                                                                      + project_errs),
                               "warnings": [], "skipped": [], "log": ""}
            counts["errors"] += 1
        details[name] = per
    run_errs = [i["message"] for i in issues if not i["project"] and i["level"] in ("ERROR", "CRITICAL")]
    total = sum(counts[k] for k in ("ok", "failed", "errors", "skipped"))
    if total and counts["failed"] == total:
        outcome = "FAILED"
    elif counts["failed"] or counts["errors"] or counts["skipped"] or run_errs:
        outcome = "PARTIAL FAILURE"
    elif not report["projects"] and report.get("not_run"):
        outcome = "SUCCESS"                  # every project is Closed or Not started (#151 step 3): nothing to do
    elif not report["projects"]:
        outcome = "FAILED"
        run_errs.append("No projects were selected or found under the projects folder.")
    else:
        outcome = "SUCCESS"
    report["outcome"] = outcome
    report["details"] = details
    report["run_errors"] = run_errs
    # warnings about the whole run, marked run_wide (L&I's records out of date or down: agent/intent.py); other
    # warnings logged outside a project stay out of the email, as before
    report["run_warnings"] = _unique([i["message"] for i in issues if i.get("run_wide") and i["level"] == "WARNING"])
    report["warning_count"] = sum(len(t["warnings"]) for p in details.values() for t in p.values()) \
        + len(report["run_warnings"])
    report["level"], report["level_reason"] = email_level(report)
    return report


# ------------------------------------------------------------------------------------------------------- links
# The tokenizer's list of each AWS-side copy's staff file (code/tokenizer/pii_tokenizer/screen.py STAFF_LINKS)
STAFF_LINKS = "_staff_links.json"
_DRIVE_ID = re.compile(r"[A-Za-z0-9_-]{10,100}")


class Linker:
    """Turns a file in the run's root folder into links people can open."""

    def __init__(self, root: Path, ncfg):
        self.root = Path(root).resolve()
        self.cfg = ncfg["links"]
        self.remote = os.environ.get("SOURCE_REMOTE", "")
        self.mirror = os.environ.get("MIRROR_REMOTE", "")
        self.region = os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION") or ""
        self._staff_ids = None

    def rel(self, path):
        try:
            return Path(path).resolve().relative_to(self.root).as_posix()
        except ValueError:
            return None

    def _drive_link(self, rel):
        """A link to the STAFF's own file for the AWS-side file at ``rel`` (David, Nov 15, 2026: no email link may lead
        to the AWS side, not even the tokenized copy on Drive): the tokenizer writes, after each restore, each copy's
        staff file by Drive ID (code/tokenizer: screen.STAFF_LINKS, mirror._publish_links). No entry, no link."""
        if not (self.cfg.get("google_drive") and self.remote and shutil.which("rclone")):
            return None
        if self._staff_ids is None:
            self._staff_ids = {}
            try:
                out = subprocess.run(["rclone", "cat", f"{self.remote.rstrip('/')}/{STAFF_LINKS}"],
                                     capture_output=True, text=True, timeout=60)
                links = json.loads(out.stdout or "{}").get("links") if out.returncode == 0 else {}
                self._staff_ids = {str(k): str(v) for k, v in (links or {}).items()
                                   if isinstance(v, str) and _DRIVE_ID.fullmatch(v)}
            except Exception as e:
                log.debug("Staff file links not read: %s", err_text(e))   # no path: #56
        fid = self._staff_ids.get(rel)
        return f"https://drive.google.com/file/d/{fid}/view" if fid else None

    def links(self, path):
        """[(label, url_or_text)] for a log file."""
        rel = self.rel(path)
        out = []
        if rel:
            d = self._drive_link(rel)
            if d:
                out.append(("Google Drive (staff copy)", d))
            base = self.remote.split(":", 1)[-1].rstrip("/").split("/")[-1] if self.remote else self.root.name
            out.append(("Shared folder path", shown_name(f"{base}/{rel}")))
        if self.cfg.get("local_paths") and not self.remote:
            out.append(("On this computer", Path(path).resolve().as_uri()))
        return out


# ------------------------------------------------------------------------------------------------------ message
STATUS_LABEL = {"ok": "OK", "failed": "FAILED", "errors": "COMPLETED WITH ERRORS", "skipped": "SKIPPED"}
OUTCOME_COLOR = {"SUCCESS": "#1e7a46", "PARTIAL FAILURE": "#b35c00", "FAILED": "#b3261e", "NOT RUN": "#b3261e"}
LEVEL_COLOR = {"Success": "#1e7a46", "Warning": "#9a6700", "Failure": "#b35c00", "Critical": "#b3261e"}
_REVIEWED = re.compile(r"^payroll: (?:re-)?reviewed .*?\((\d+) lines\): (\d+) rule-breaking, (\d+) suspected")
_MARKED = re.compile(r"^intent: (\d+) invalid/missing value")
_LNI_ONLY = re.compile(r"^intent: (\d+) contractor\(s\) in L&I's records for this project are not in the tracker")
_HIRING = re.compile(r"^intent: (\d+) hiring contractor\(s\) in the tracker don't match the L&I intent")   # #121
_NORATE = re.compile(r"^payroll:   no wage rate: ([\d,]+)")
_MERGED = re.compile(r"^(federal|state|union): merged (\d+) rows from (\d+) file")


def _pretty(line: str) -> str:
    """A breakdown line ('payroll:   rule-breaking: ...') shown indented under the result it belongs to."""
    m = re.match(r"^(\w+):   (.*)$", line)
    return f"   -> {m.group(2)}" if m else line


def _dur(secs):
    secs = float(secs)
    return f"{secs:.0f} s" if secs < 90 else f"{int(secs // 60)} min {int(secs % 60):02d} s"


def email_level(report, extra_failures=()):
    """(level, reason) for the subject and the first line of the email:
    Success  - everything finished, nothing needs attention
    Warning  - everything finished, but something needs a look: rule-breaking issues or lines with no wage rate in a
               payroll, values marked in a tracker, warnings, a file that could not be read, the tokenizer not run
    Failure  - some tasks failed or finished with errors, or a data gap kept payroll lines from being fully validated
               (no wage data, determination not named, missing columns, many lines with no wage rate); the rest finished
    Critical - the run could not do its job: a pipeline step failed, it timed out, nothing ran, or every task failed"""
    details = report.get("details") or {}
    pairs = [(p, t, d) for p, per in details.items() for t, d in per.items() if not t.startswith("_")]
    bad = [(p, t) for p, t, d in pairs if d["status"] in ("failed", "errors", "skipped")]
    outcome = report.get("outcome", "FAILED")
    if extra_failures or report.get("run_errors") or outcome in ("FAILED", "NOT RUN"):
        why = (list(extra_failures) and extra_failures[0].get("text")) or (report.get("run_errors") or [None])[0] \
            or ("every task failed" if pairs else "nothing was run")
        return "Critical", str(why)
    gaps = data_gaps(report)
    gap_txt = (f"{len(gaps)} data gap{'s' if len(gaps) != 1 else ''} kept payroll lines from being fully validated"
               if gaps else "")
    broken = [p for p, per in details.items() if (per.get("_project") or {}).get("status") in ("failed", "errors")]
    if bad or outcome == "PARTIAL FAILURE":
        what = (f"{len(bad)} of {len(pairs)} task{'s' if len(pairs) != 1 else ''} failed or finished with errors"
                if bad or not broken else
                f"{len(broken)} project{'s' if len(broken) != 1 else ''} couldn't be prepared")
        return "Failure", what + (f"; {gap_txt}" if gaps else "")
    if gaps:
        return "Failure", gap_txt
    notes = sum(len(d["warnings"]) for _, _, d in pairs) + sum(
        1 for lines in report.get("projects", {}).values() for l in lines
        if l.startswith("WARNING") or "could not read" in l or l.startswith("wages: no Jurisdiction column")) \
        + len(report.get("pipeline_warnings") or []) + len(report.get("run_warnings") or [])
    reds = norate = marked = absent = hiring = 0
    for lines in report.get("projects", {}).values():
        for l in lines:
            if m := _REVIEWED.match(l):
                reds += int(m[2])
            elif m := _MARKED.match(l):
                marked += int(m[1])
            elif m := _LNI_ONLY.match(l):
                absent += int(m[1])
            elif m := _HIRING.match(l):          # #121: counted apart from the marked values (review N1)
                hiring += int(m[1])
            elif m := _NORATE.search(l):
                norate += int(m[1].replace(",", ""))
    found = [f"{reds:,} rule-breaking payroll issue{'s' if reds != 1 else ''}" if reds else "",
             f"{norate:,} payroll line{'s' if norate != 1 else ''} with no wage rate found" if norate else "",
             f"{marked:,} tracker value{'s' if marked != 1 else ''} marked" if marked else "",
             f"{absent:,} L&I contractor{'s' if absent != 1 else ''} not in a tracker" if absent else "",
             f"{hiring:,} hiring contractor mismatch{'es' if hiring != 1 else ''}" if hiring else "",
             f"{notes} warning{'s' if notes != 1 else ''}" if notes else ""]
    found = [f for f in found if f]
    if found:
        return "Warning", ", ".join(found)
    return "Success", f"all {len(pairs)} task{'s' if len(pairs) != 1 else ''} OK"


def data_gaps(report):
    """[(project, text)] from the 'GAP <task>: ...' result lines: what kept payroll lines from being fully validated."""
    return [(p, l.split(": ", 1)[1]) for p, lines in (report.get("projects") or {}).items() for l in lines
            if l.startswith("GAP ")]


# Written by the Google-side tokenizer at the top of the tokenized folder (code/tokenizer/pii_tokenizer/screen.py):
# the files it kept in the staff drive instead of copying them to AWS, as anonymous keys only. Their names and reasons
# are in "Held back from AWS.csv" in the staff drive: a name may itself be the reason, so it never comes here.
HELD_BACK_SUMMARY = "_held_back_from_aws.json"
HELD_BACK_REPORTED = Path(".agent") / "held_back_reported.json"     # keys already mentioned in an email


def held_back(root) -> tuple[list[str], list[str] | None]:
    """(keys not mentioned in an email yet, all keys). A file changed and held back again gets a new key.
    All keys is None when the summary can't be read: then nothing is said, and the record of what was already
    mentioned is kept as it is (not emptied, which would mention every file again next time)."""
    root = Path(root)
    try:
        data = json.loads((root / HELD_BACK_SUMMARY).read_text())
        keys = [e["key"] for e in data.get("held_back", []) if isinstance(e, dict) and isinstance(e.get("key"), str)]
    except (OSError, ValueError, AttributeError):
        return [], None
    try:
        seen = set(json.loads((root / HELD_BACK_REPORTED).read_text()))
    except (OSError, ValueError, TypeError):
        seen = set()
    return [k for k in keys if k not in seen], keys


def _held_back_projects(root, kind) -> list[str]:
    """One project name per held-back entry of ``kind`` ("tracker" or "upload"), as the summary gives them (the
    project folder's name as the AWS side has it; no file name, no reason). [] when it can't be read or has none."""
    try:
        data = json.loads((Path(root) / HELD_BACK_SUMMARY).read_text())
        entries = data.get("held_back", [])
    except (OSError, ValueError, AttributeError):
        return []
    if not isinstance(entries, list):
        return []
    return [e["project"] for e in entries if isinstance(e, dict) and e.get("kind") == kind
            and isinstance(e.get("project"), str) and 0 < len(e["project"]) <= 200]


def held_back_trackers(root) -> list[str]:
    """The projects whose tracker (a workbook directly in Tracker/) the tokenizer kept in the staff drive. A project
    with a held-back tracker has no tracker here, but it isn't a project without one: the emails say so."""
    return sorted(set(_held_back_projects(root, "tracker")))


def held_back_uploads(root) -> dict:
    """{project: how many files directly in its Uploads/ the tokenizer kept in the staff drive}: the staff email
    names them on that project's payroll review or wage update when the run did one (David, Nov 12, 2026)."""
    out: dict = {}
    for p in _held_back_projects(root, "upload"):
        out[p] = out.get(p, 0) + 1
    return out


_DRIVE_ID_RE = re.compile(r"[A-Za-z0-9_-]{10,120}")


def held_back_link(root) -> str | None:
    """A link to the staff drive's "Held back from AWS.csv", from the Drive ID the tokenizer leaves in its summary;
    None if there is none or it isn't a Drive ID. A plain file link: it grants nothing, Drive asks anyone without
    access to the staff drive (or the file) to request it."""
    try:
        sid = json.loads((Path(root) / HELD_BACK_SUMMARY).read_text()).get("sheet_id")
    except (OSError, ValueError, AttributeError):
        return None
    if not isinstance(sid, str) or not _DRIVE_ID_RE.fullmatch(sid):
        return None
    return f"https://drive.google.com/file/d/{sid}/view"


def held_back_line(new, keys) -> str | None:
    if not new:
        return None
    return (f"Held back from AWS: {len(new)} new file{'s' if len(new) != 1 else ''} ({len(keys)} in all): kept in "
            "the staff drive, so the agent didn't check them. \"Held back from AWS.csv\" at the top of the shared drive "
            "says which and why (for example an SSN in a file name, or a payroll sent as a PDF), and how to fix each.")


def _remember_held_back(root, keys) -> None:
    if keys is None:                        # the summary couldn't be read: keep the record as it is
        return
    try:
        path = Path(root) / HELD_BACK_REPORTED
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(sorted(keys)))
    except OSError as e:
        log.warning("Could not remember the held-back files already reported: %s", err_text(e))


# #151 step 3: the statuses that keep a project from being checked (agent/runner.py split_inactive)
INACTIVE = {"closed": "Closed", "not_started": "Not started"}


def not_run_line(not_run):
    """The technical email's line on projects not checked because of their status (None when there are none)."""
    parts = [f"{INACTIVE[s]}: {', '.join(n)}" for s, n in (not_run or {}).items() if n and s in INACTIVE]
    return ("Not checked (status on the project's Project Data sheet): " + "; ".join(parts)) if parts else None


def at_a_glance(report) -> list[str]:
    """The key numbers of the run, from its result lines."""
    rev = lines_n = reds = oranges = trackers = marked = absent = hiring = 0
    merged, skipped, unread, unrec, nojur = {}, 0, 0, 0, 0
    for lines in report.get("projects", {}).values():
        for l in lines:
            if m := _REVIEWED.match(l):
                rev += 1; lines_n += int(m[1]); reds += int(m[2]); oranges += int(m[3])
            elif m := _MARKED.match(l):
                trackers += 1; marked += int(m[1])
            elif m := _LNI_ONLY.match(l):
                absent += int(m[1])
            elif m := _HIRING.match(l):
                hiring += int(m[1])
            elif m := _MERGED.match(l):
                merged[m[1]] = merged.get(m[1], 0) + int(m[2])
            elif "skipped: already processed" in l:
                skipped += 1
            elif l.startswith("wages: could not read"):
                unread += l.split(";")[0].count(",") + 1
            elif l.startswith("wages: no Jurisdiction column in"):
                nojur += l.split(";")[0].count(",") + 1
            elif l.startswith("WARNING upload:"):
                unrec += 1
    out = []
    if rev:
        out.append(f"Payroll: {rev} file{'s' if rev != 1 else ''} reviewed, {lines_n:,} lines: {reds:,} rule-breaking "
                   f"and {oranges:,} suspected issue(s)")
    if trackers:
        out.append(f"Trackers: {trackers} checked against L&I, {marked} invalid or missing value(s) marked"
                   + (f", {hiring} hiring contractor mismatch(es)" if hiring else "")
                   + (f", {absent} L&I contractor(s) not in the tracker" if absent else ""))
    if merged or skipped or unread or nojur:
        parts = [f"{n:,} {k} rows merged" for k, n in sorted(merged.items())]
        parts += [f"{skipped} file(s) skipped (already processed)"] if skipped else []
        parts += [f"{unread} file(s) could not be read"] if unread else []
        parts += [f"{nojur} file(s) without a Jurisdiction column (left in Uploads)"] if nojur else []
        out.append("Wage data: " + ", ".join(parts))
    if unrec:
        out.append(f"Uploads: {unrec} file(s) not recognized (left in Uploads)")
    return out


def _local_time(ts, tzname):
    try:
        from zoneinfo import ZoneInfo
        t = dt.datetime.strptime(ts, TS_FORMAT).replace(tzinfo=dt.timezone.utc).astimezone(ZoneInfo(tzname))
        return t.strftime("%a %b %-d, %-I:%M %p %Z")
    except Exception:
        return None


def describe_trigger(trigger):
    t = trigger or os.environ.get("COMPLIANCE_TRIGGER") or "manual (command line)"
    return t


def result_files(report, root):
    """[(project, label, path)] for files a run produced: reviewed payrolls and marked trackers."""
    import re as _re
    out = []
    from .config import projects_folder
    projects = projects_folder(root)
    for project, lines in (report.get("projects") or {}).items():
        for line in lines:
            m = _re.match(r"payroll: .*-> ((?:Payroll|Reviewed payroll)/.+\.xlsx)$", line)
            if m:
                out.append((project, "Reviewed payroll", projects / project / m.group(1)))
            # the tracker's name ends at " (", ";" or the line's end (#82: it used to need a ";", so a run with no
            # warning didn't attach its marked tracker)
            m = _re.match(r"intent: \d+ invalid/missing value\(s\) marked in (.+?\.xlsx?m?)(?: \(|;|$)", line)
            if m:
                out.append((project, "Marked tracker", projects / project / "Tracker" / m.group(1)))
    return [(p, l, f) for p, l, f in out if f.exists()]


def run_options(report) -> str:
    """The options the run was started with, as a command line would give them (#30): the pipeline's arguments when
    the AWS pipeline started it, else the agent's own (--task, --project, --force). "none" for a plain full run."""
    args = str(report.get("run_args") or "").strip()
    if not args and report.get("plan"):          # each project its own tasks (run.py --plan)
        args = " ".join(f'--plan {",".join(t)}:"{p}"' for p, t in report["plan"].items()) + \
            (" --force" if report.get("force") else "")
    if not args:
        tasks = sorted(report.get("tasks") or [])
        parts = [] if not tasks or set(tasks) == {"wages", "payroll", "intent"} else ["--task " + ",".join(tasks)]
        wanted = report.get("requested") or []
        parts += [f'--project "{w}"' for w in (wanted if isinstance(wanted, list) else [wanted])]
        parts += ["--force"] if report.get("force") else []
        args = " ".join(parts)
    return args or "none"


def build_message(report, ncfg, linker, trigger, extra_failures=(), extra_files=()):
    outcome = report.get("outcome", "FAILED")
    ts = report.get("run_utc") or dt.datetime.now(dt.timezone.utc).strftime(TS_FORMAT)
    tasks = _ordered_tasks(report.get("tasks") or [])
    details = report.get("details", {})
    limit = int(ncfg.get("max_issues_listed") or 50)
    host = os.environ.get("COMPLIANCE_HOST_LABEL") or socket.gethostname()
    where = f"AWS ({host})" if os.environ.get("SOURCE_REMOTE") else f"{host} ({platform.system()})"
    task_txt = ", ".join(TASK_NAMES.get(t, t) for t in tasks) or "no tasks"
    n_proj = len(report.get("projects", {}))
    proj_txt = f"{n_proj} project{'s' if n_proj != 1 else ''}" if n_proj else report.get("requested_projects", "all projects")
    proj_list = ", ".join(report.get("projects", {})) or f"{report.get('requested_projects', 'all projects')} (not reached)"
    level, reason = email_level(report, extra_failures)
    local = _local_time(ts, ncfg.get("timezone") or "UTC")
    when = f"{local} ({ts} UTC)" if local else f"{ts} UTC"
    # the subject's few words: cut at a word, so the subject stays readable in an inbox list
    short = reason if len(reason) <= 55 else (reason[:52].rsplit(" ", 1)[0] if " " in reason[:52] else reason[:52]
                                              ).rstrip(" ,;:") + "..."
    # two verdicts (agent/staff_email.py): Run, did the tasks work (the administrator's); Findings, did the projects pass
    from .staff_email import ALL, COLOR as FINDINGS_COLOR, FAILED, email_rows, findings_verdict, run_verdict
    new_hb, _keys = held_back(linker.root)
    report = {**report, "held_back_trackers": held_back_trackers(linker.root),
              "held_back_uploads": {}}            # the whole-drive line instead (David, Nov 12: admin email as is)
    # the same rows as the staff email: the whole run's (didn't start or finish, files held back) and every project's
    rows, _problem, _late, _links, unreached, _stopped = email_rows(
        report, level, extra_failures, (len(new_hb), held_back_link(linker.root)) if new_hb else None, admin=True)
    run_word, findings = run_verdict(report, rows, extra_failures), findings_verdict(rows, unreached)
    n_review = len({r.project for r in rows if r.findings == "Needs review" and r.project != ALL})
    subject = (f"{ncfg['subject_prefix']} Run {run_word.lower()} · {findings}: {task_txt} for {proj_txt}"
               f" - {short} ({local or ts + ' UTC'})").strip()
    if level == "Critical":
        run_lead = f"The run could not do its job: {reason.rstrip('.')}. See Failures below."
    elif run_word == FAILED:
        failed = [f"{r.project} · {r.task}" for r in rows if r.run == FAILED and r.project != ALL]
        why = (", ".join(failed[:3]) + (f" and {len(failed) - 3} more" if len(failed) > 3 else "")) if failed \
            else reason.split("; ")[0]
        run_lead = f"Part of the run did not work: {why}. The rest finished; see Failures below."
    elif run_word != "Completed":
        run_lead = "Every task finished, with something to look at before it breaks: see Warnings below."
    else:
        run_lead = "Every task finished normally."
    findings_lead = {"Needs review": (f"{n_review} {'project needs' if n_review == 1 else 'projects need'} review"
                                      if n_review else "Something needs review (see What to do)")
                                     + (" (including data gaps: see Data gaps below)." if data_gaps(report) else "."),
                     "Not checked": "Some projects couldn't be checked.",
                     "No issues": "No issues found."}[findings]
    lead = f"{run_lead} Findings: {findings_lead}"
    level_word, level_color = f"Run {run_word.lower()}", {FAILED: LEVEL_COLOR["Critical"]}.get(
        run_word, LEVEL_COLOR["Warning"] if run_word != "Completed" else LEVEL_COLOR["Success"])
    took = report.get("duration_s")

    # failures, in the order people should read them
    failures = list(extra_failures)
    failures += [{"title": "Run", "text": e, "log": report.get("run_log", "")} for e in report.get("run_errors", [])]
    attach = []
    for project, per in details.items():
        for task, d in per.items():
            title = f"{project} - {TASK_NAMES.get(task, 'Project setup' if task == '_project' else task)}"
            for e_ in d["errors"]:
                pre = "The task stopped with an error: " if d.get("crashed") else ""
                failures.append({"title": title, "text": pre + e_, "log": d["log"] or report.get("run_log", "")})
            for s in d["skipped"]:
                why = s.split(":", 1)[-1].strip()
                why = why[len("SKIPPED - "):] if why.startswith("SKIPPED - ") else why
                failures.append({"title": title, "text": f"Not done: {why}",
                                 "log": d["log"] or report.get("run_log", "")})
            if d["status"] != "ok" and d["log"]:
                attach.append(d["log"])
    if outcome != "SUCCESS" and report.get("run_log"):
        attach.append(report["run_log"])
    attach += [f for f in extra_files if f]

    warnings = [(f"{p} - {TASK_NAMES.get(t, t)}", w, d["log"]) for p, per in details.items()
                for t, d in per.items() for w in d["warnings"]]
    warnings = [("All projects - Run", w, report.get("run_log", "")) for w in report.get("run_warnings") or []] + warnings
    warnings = [(f"Pipeline - {w.split(':', 1)[0]}", w.split(":", 1)[-1].strip(), "")
                for w in report.get("pipeline_warnings") or []] + warnings

    link_cache = {}

    def links_for(path):
        if not path:
            return []
        if path not in link_cache:
            link_cache[path] = linker.links(path)
        return link_cache[path]

    # warnings: the same message for several projects is listed once, with the projects
    grouped = {}
    for title, w, log_ in warnings:
        proj, _, task = title.partition(" - ")
        grouped.setdefault((task, w), []).append(proj)
    glance = at_a_glance(report)
    if not_run_line(report.get("not_run") or {}):
        glance.insert(0, not_run_line(report["not_run"]))       # #151 step 3: Closed / Not started, not checked
    hb_line = held_back_line(*held_back(linker.root))
    hb_url = held_back_link(linker.root) if hb_line else None
    if hb_line:
        glance.append(hb_line)
    gaps = data_gaps(report)
    from .staff_email import GRID_ORDER, render_grid     # the staff email's grid, with the technical warnings too
    grid_text, grid_html = render_grid([x for x in GRID_ORDER if x in tasks], list(report.get("projects") or {}),
                                       [r for r in rows if r.project != ALL], admin=True, unreached=unreached,
                                       plan=report.get("plan"))

    # incremental (only what changed) or a full rerun (re-review even if unchanged, "--force"); #30
    action = "Full rerun" if report.get("force") else "Incremental"
    options = run_options(report)                         # the options that started the run (#30)

    # ---- plain text
    t = [f"{level_word.upper()} · {findings.upper()}: {lead}", "",
         f"Run: {when}" + (f", took {_dur(took)}" if took else ""), f"Started by: {describe_trigger(trigger)}",
         f"Ran on: {where}", f"Action: {action}", f"Options: {options}", f"Tasks: {task_txt}", f"Projects: {proj_list}"]
    from .admin_settings import summary_lines
    page_lines = summary_lines(ncfg)                      # the Settings page's choices in effect (#151)
    t += page_lines
    if grid_text:
        t += [""] + grid_text
    if glance:
        t += ["", "AT A GLANCE"] + [f"- {g}" + (f" Open the list: {hb_url}" if hb_url and g == hb_line else "")
                                    for g in glance]
    if report.get("model"):
        t.append(f"Local model: {report['model'].get('name')} (id {report['model'].get('id')})")
    t.append("")
    if gaps:
        t.append(f"DATA GAPS ({len(gaps)}) - payroll lines that could not be fully validated")
        for i, (p, g) in enumerate(gaps[:limit], 1):
            t.append(f"{i}. {p}: {g}")
        t.append("")
    if failures:
        t.append(f"FAILURES ({len(failures)})")
        for i, f in enumerate(failures[:limit], 1):
            t.append(f"{i}. {f['title']}: {f['text']}")
            for label, url in links_for(f.get("log")):
                t.append(f"   {label}: {url}")
        if len(failures) > limit:
            t.append(f"... and {len(failures) - limit} more; see the logs.")
        t.append("")
    if details:
        t.append("WHAT THE RUN DID")
    for project in details:
        t.append(f"{project}")
        for line in report["projects"].get(project, []):
            t.append(f"    {_pretty(line)}")
    produced = result_files(report, linker.root)
    if produced:
        t += ["", "FILES PRODUCED"]
        for project, label, path in produced:
            t.append(f"- {project} - {label}: {shown_name(path.name)}")
            t += [f"    {lb}: {u}" for lb, u in links_for(str(path))]
    if warnings:
        t += ["", f"WARNINGS ({len(warnings)}) - not failures, but worth a look"]
        for (task, w), projs in list(grouped.items())[:limit]:
            t.append(f"- {task} ({', '.join(projs)}): {w}")
    from .whats_new import recent
    news = recent(dt.date(*map(int, ts[:10].split("-"))) if ts[:4].isdigit() else None)
    if news:
        t += ["", "WHAT'S NEW"] + [f"- {d}: {title}. {text}" for d, title, text in news]
    if report.get("run_log"):
        t += ["", "Full run log:"] + [f"  {label}: {url}" for label, url in links_for(report["run_log"])]
    t += ["", f"Sent by the Compliance Validation Agent. Recipients are set {_recipients_where(ncfg)}."]
    text = "\n".join(t)

    # ---- HTML
    e = html.escape
    color = OUTCOME_COLOR.get(outcome, "#333")

    def linkline(path):
        ls = links_for(path)
        parts = [f'<a href="{e(u)}">{e(l)}</a>' if u.startswith(("http", "file:")) else f"{e(l)}: <code>{e(u)}</code>"
                 for l, u in ls]
        return " &middot; ".join(parts)

    h = ['<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1c1f1d;max-width:820px">',
         f'<p style="font-size:20px;font-weight:bold;margin:0 0 4px"><span style="color:{level_color}">{e(level_word)}</span>'
         f' <span style="color:#9aa5a0">·</span> <span style="color:{FINDINGS_COLOR[findings]}">{e(findings)}</span></p>',
         f'<p style="font-size:15px;margin:0 0 12px">{e(lead)}</p>',
         '<table style="border-collapse:collapse;font-size:13px">']
    for k, v in (("Run", when + (f", took {_dur(took)}" if took else "")), ("Started by", describe_trigger(trigger)),
                 ("Ran on", where), ("Action", action), ("Options", options), ("Tasks", task_txt),
                 ("Projects", proj_list)):
        h.append(f'<tr><td style="padding:2px 12px 2px 0;color:#55605b">{e(k)}</td><td>{e(v)}</td></tr>')
    h.append("</table>")
    if page_lines:
        h.append('<p style="font-size:13px;margin:8px 0 0;white-space:pre-wrap">' + e("\n".join(page_lines)) + "</p>")
    if grid_html:
        h.append(f'<div style="margin:16px 0 4px">{grid_html}</div>')
    if glance:
        h.append('<h3 style="margin:18px 0 6px">At a glance</h3><ul style="padding-left:20px;font-size:13px">')
        h += [f"<li>{e(g)}" + (f' <a href="{e(hb_url)}">Open the list</a>' if hb_url and g == hb_line else "")
              + "</li>" for g in glance]
        h.append("</ul>")
    if gaps:
        h.append(f'<h3 style="color:{color};margin:18px 0 6px">Data gaps ({len(gaps)})</h3>'
                 '<p style="margin:0 0 4px;font-size:12px;color:#55605b">Payroll lines that could not be fully '
                 'validated.</p><ol style="padding-left:20px">')
        h += [f'<li style="margin-bottom:6px"><b>{e(p)}</b>: {e(g)}</li>' for p, g in gaps[:limit]]
        h.append("</ol>")
    if failures:
        h.append(f'<h3 style="color:{color};margin:18px 0 6px">Failures ({len(failures)})</h3><ol style="padding-left:20px">')
        for f in failures[:limit]:
            ll = linkline(f.get("log"))
            h.append(f'<li style="margin-bottom:6px"><b>{e(f["title"])}</b>: {e(f["text"])}'
                     + (f'<br><span style="font-size:12px">Log: {ll}</span>' if ll else "") + "</li>")
        h.append("</ol>")
        if len(failures) > limit:
            h.append(f"<p>... and {len(failures) - limit} more; see the logs.</p>")
    if report.get("projects"):
        h.append('<h3 style="margin:18px 0 6px">What the run did</h3>')
    for project, lines in report.get("projects", {}).items():
        h.append(f"<p style='margin:8px 0 2px'><b>{e(project)}</b></p><ul style='margin:0;padding-left:20px;font-size:13px'>")
        h += [f"<li style='list-style:none;color:#55605b'>&#8627; {e(_pretty(l).strip())}</li>" if _pretty(l) != l
              else f"<li>{e(l)}</li>" for l in lines] or ["<li>nothing to do</li>"]
        h.append("</ul>")
    if produced:
        h.append('<h3 style="margin:18px 0 6px">Files produced</h3><ul style="padding-left:20px;font-size:13px">')
        for project, label, path in produced:
            ll = linkline(str(path))
            h.append(f"<li><b>{e(project)}</b> - {e(label)}: {e(shown_name(path.name))}" + (f"<br><span style='font-size:12px'>{ll}</span>" if ll else "") + "</li>")
        h.append("</ul>")
    if warnings:
        h.append(f'<h3 style="margin:18px 0 6px">Warnings ({len(warnings)})</h3>'
                 '<p style="margin:0 0 4px;font-size:12px;color:#55605b">Not failures, but worth a look.</p>'
                 '<ul style="padding-left:20px;font-size:13px">')
        h += [f"<li><b>{e(task)}</b> ({e(', '.join(projs))}): {e(w)}</li>" for (task, w), projs in list(grouped.items())[:limit]]
        h.append("</ul>")
    if news:
        h.append('<h3 style="margin:18px 0 6px">What\'s new</h3><ul style="padding-left:20px;font-size:13px">')
        h += [f"<li><b>{e(title)}</b> <span style='color:#7b8a84'>({e(d)})</span><br>{e(text)}</li>"
              for d, title, text in news]
        h.append("</ul>")
    if report.get("run_log"):
        h.append(f'<p style="font-size:12px">Full run log: {linkline(report["run_log"])}</p>')
    h.append(f'<p style="font-size:11px;color:#7b8a84">Sent by the Compliance Validation Agent. Recipients are set '
             f'{e(_recipients_where(ncfg))}.</p></div>')
    return subject, text, "\n".join(h), _dedupe(attach)


def _recipients_where(ncfg) -> str:
    values = ((ncfg.get("page") or {}).get("values") or {})
    file = Path(ncfg.get("path") or "notifications.config").name
    from .admin_settings import NOTIFY_KEYS
    on_page = [k for k in NOTIFY_KEYS if k in values]
    if not on_page:
        return f"in {file}"
    elsewhere = [f for k, f in NOTIFY_KEYS.items() if k not in values and ncfg.get(f)]   # a list set outside the page
    return f"on the web app's Settings page and in {file}" if elsewhere else "on the web app's Settings page"


def _dedupe(items):
    out = []
    for i in items:
        if i and i not in out and Path(i).exists():
            out.append(i)
    return out


# ------------------------------------------------------------------------------------------------------ sending
LOGO = Path(__file__).resolve().parents[1] / "ui" / "static" / "brand" / "cva-demo-lockup-email-light.png"


def _logo() -> bytes | None:
    """The logo for the top of the emails (build_tools/brand_assets.py makes it); none if the file is missing."""
    try:
        return LOGO.read_bytes()
    except OSError:
        return None


def _compose(ncfg, subject, text, html_body, attachments, outcome):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = ncfg["from"] if "<" in ncfg["from"] else formataddr(("Compliance Validation Agent", ncfg["from"]))
    msg["To"] = ", ".join(ncfg["to"])
    if ncfg["cc"]:
        msg["Cc"] = ", ".join(ncfg["cc"])
    if ncfg["reply_to"]:
        msg["Reply-To"] = ", ".join(ncfg["reply_to"])
    msg["Message-ID"] = make_msgid(domain=parseaddr(msg["From"])[1].split("@")[-1] or None)
    msg["X-Compliance-Outcome"] = outcome
    msg.set_content(text)
    logo = _logo()
    if logo:
        # the company logo above the email, as an image inside the message (cid:): mail programs show those without
        # asking to load remote images, and nothing is fetched from anywhere when the email is opened
        cid = make_msgid(domain="cva.logo")
        html_body = (f'<div style="font-family:Arial,Helvetica,sans-serif;padding:0 0 14px;margin:0 0 18px;'
                     f'border-bottom:1px solid #e3e8e5"><img src="cid:{cid[1:-1]}" alt="CVA" '
                     f'width="192" height="56" style="display:block;width:192px;height:56px;border:0"></div>') + html_body
        msg.add_alternative(html_body, subtype="html")
        msg.get_payload()[1].add_related(logo, "image", "png", cid=cid, filename="cva-logo.png",
                                         disposition="inline")
    else:
        msg.add_alternative(html_body, subtype="html")
    mode = ncfg.get("attach_logs", "problems")
    if os.environ.get("SOURCE_REMOTE") and not ncfg.get("attach_logs_on_aws"):
        # on AWS the logs hold tokens instead of employee names: link to them on the drive (where the tokenizer has
        # put the names back) rather than attach them
        mode = "never"
    if mode == "all" or (mode == "problems" and outcome != "SUCCESS"):
        budget = float(ncfg.get("max_attachments_mb") or 8) * 1024 * 1024
        for path in attachments:
            p = Path(path)
            size = p.stat().st_size
            if size > budget:
                log.info("Not attaching a %s file (%d bytes): over the attachment limit", p.suffix.lower(), size)  # #56
                continue
            budget -= size
            parts = p.parts
            name = p.name
            if "Logs" in parts and parts.index("Logs") >= 2:   # <project>__<ts>__<task>__log.txt
                i = parts.index("Logs")
                proj, ts = (parts[i - 2], parts[i + 1]) if parts[i - 1] == "Files" else (parts[i - 3], parts[i - 1])
                name = f"{proj}__{ts}__{p.name}"
            msg.add_attachment(p.read_bytes(), maintype="text", subtype="csv", filename=name.replace(" ", "_"))
    return msg


def _send_ses(ncfg, msg):
    import boto3
    region = ncfg["ses"].get("region") or os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION")
    if not region:                          # e.g. the desktop app with an AWS profile that names no region
        try:
            region = boto3.session.Session().region_name
        except Exception:                   # noqa: BLE001
            region = None
        region = region or "us-east-2"      # where the firm's sender is verified
    client = boto3.client("sesv2", region_name=region) if region else boto3.client("sesv2")
    kw = {"FromEmailAddress": parseaddr(msg["From"])[1],
          "Destination": {"ToAddresses": ncfg["to"], "CcAddresses": ncfg["cc"], "BccAddresses": ncfg["bcc"]},
          "Content": {"Raw": {"Data": msg.as_bytes()}}}
    if ncfg["ses"].get("configuration_set"):
        kw["ConfigurationSetName"] = ncfg["ses"]["configuration_set"]
    try:
        resp = client.send_email(**kw)
    except Exception as e:                  # noqa: BLE001
        refused = _refused_addresses(e, kw)
        if not refused:
            raise
        # Amazon SES in test mode ("sandbox") refuses the whole email when one address isn't verified (David, Nov 11:
        # an address added on the Settings page): it goes to the others, and the administrator's email names it
        d = kw["Destination"]
        for k in ("ToAddresses", "CcAddresses", "BccAddresses"):
            d[k] = [a for a in d[k] if a.lower() not in refused]
        if not (d["ToAddresses"] or d["CcAddresses"] or d["BccAddresses"]):
            raise
        for header, key in (("To", "ToAddresses"), ("Cc", "CcAddresses")):     # the headers match who gets it
            if msg[header] is not None:
                del msg[header]
                if d[key]:
                    msg[header] = ", ".join(d[key])
        kw["Content"]["Raw"]["Data"] = msg.as_bytes()
        resp = client.send_email(**kw)      # recorded only once this one is sent (review H1)
        log.warning("The email service refused %d address(es) (not verified); sent to the others", len(refused))
        ncfg.setdefault("refused", [])
        ncfg["refused"] += [a for a in sorted(refused) if a not in ncfg["refused"]]
    return resp.get("MessageId", "")


def _refused_addresses(e, kw) -> set:
    """The recipients Amazon SES refused as not verified (its MessageRejected names them), or an empty set: for
    anything else, or when the sender itself is the one refused."""
    err = getattr(e, "response", {}).get("Error", {}) if hasattr(e, "response") else {}
    text = str(err.get("Message") or e)
    if err.get("Code") != "MessageRejected" or "not verified" not in text:
        return set()
    named = {a.lower().rstrip(".") for a in re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+", text)}
    sender = parseaddr(kw.get("FromEmailAddress", ""))[1].lower()
    d = kw["Destination"]
    ours = {a.lower() for k in ("ToAddresses", "CcAddresses", "BccAddresses") for a in d[k]}
    if sender in named:
        return set()
    return named & ours


def _smtp_password(scfg):
    if scfg.get("password_env") and os.environ.get(scfg["password_env"]):
        return os.environ[scfg["password_env"]]
    if scfg.get("password_secret"):
        import boto3
        return boto3.client("secretsmanager").get_secret_value(SecretId=scfg["password_secret"])["SecretString"]
    return ""


def _send_smtp(ncfg, msg):
    import smtplib
    import ssl
    s = ncfg["smtp"]
    if not s.get("host"):
        raise ValueError("smtp.host is empty in notifications.config")
    ctx = ssl.create_default_context()
    port = int(s.get("port") or 587)
    if s.get("security") == "ssl":
        server = smtplib.SMTP_SSL(s["host"], port, context=ctx, timeout=60)
    else:
        server = smtplib.SMTP(s["host"], port, timeout=60)
        if s.get("security", "starttls") == "starttls":
            server.starttls(context=ctx)
    with server:
        if s.get("username"):
            server.login(s["username"], _smtp_password(s))
        server.send_message(msg, to_addrs=ncfg["to"] + ncfg["cc"] + ncfg["bcc"])
    return ""


def _staff_message(report, ncfg, root, linker, extra_failures):
    from .staff_email import build_staff_message
    from .whats_new import staff_unsent
    report = {**report, "held_back_trackers": held_back_trackers(root),     # a held-back tracker isn't "no tracker"
              "held_back_uploads": held_back_uploads(root), "held_back_link": held_back_link(root)}
    level, _ = email_level(report, extra_failures)
    ts = report.get("run_utc") or dt.datetime.now(dt.timezone.utc).strftime(TS_FORMAT)
    # staff: no whole-drive line, only what kept this run's own tasks from their files (David, Nov 12): the link
    hb = (0, held_back_link(root))
    return build_staff_message(report, ncfg, root, level, extra_failures, produced=result_files(report, root),
                               held_back=hb, news=staff_unsent(root),
                               when=_local_time(ts, ncfg.get("timezone") or "UTC") or f"{ts} UTC")


def send(report, root, trigger=None, ncfg=None, extra_failures=(), extra_files=(), dry_run=False):
    """Build and send the outcome emails for a finished (or failed) run: the short one to the staff (`to`, `cc`,
    `bcc`) and, when `admin_to` is set, the full one to the administrator. Returns a short status string starting
    with "not sent:" when the staff email could not be sent; never raises."""
    try:
        ncfg = ncfg or load_notify_config(root=root)
        if not ncfg.get("enabled", True) or ncfg.get("transport") == "none":
            return "notifications disabled"
        if not ncfg["to"] and not ncfg["admin_to"]:
            log.warning("No email sent: no recipients in %s", ncfg["path"])
            return "no recipients"
        outcome = report.get("outcome", "FAILED")
        # send_on / admin_send_on "problems": that email is held back on a clean run (no failure, no gap; findings
        # don't count); the recipient-change notices always go (#151 review F2; David, Nov 12, 2026)
        clean = outcome == "SUCCESS" and not data_gaps(report)
        staff_held = clean and ncfg.get("send_on") == "problems"
        admin_held = clean and ncfg.get("admin_send_on") == "problems"
        # every project Closed or Not started (#151 step 3): nothing was checked, so staff get nothing (only the
        # administrator's email says which projects weren't checked)
        nothing = outcome == "SUCCESS" and not report.get("projects") and bool(report.get("not_run"))
        linker = Linker(root, ncfg)
        msgs = []                     # (who, config with its recipients, the message or a function that builds it)
        if ncfg["to"] and not staff_held and not nothing:
            subject, text, html_body = _staff_message(report, ncfg, root, linker, extra_failures)
            msgs.append(("staff", ncfg, _compose(ncfg, subject, text, html_body, [], outcome)))
        from . import admin_settings
        import html as _html
        for audience, subj, body, everyone in admin_settings.recipient_notices(ncfg, root):   # the page (#151)
            ncc = {**ncfg, "to": everyone, "cc": [], "bcc": [], "refused": []}
            nmsg = _compose(ncc, f"{ncfg['subject_prefix']} {subj}".strip(), body,
                            f'<pre style="font-family:Arial,Helvetica,sans-serif;font-size:14px;white-space:pre-wrap">'
                            f'{_html.escape(body)}</pre>', [], "SUCCESS")
            msgs.append((f"recipients change ({audience})", ncc, nmsg))
        if ncfg["admin_to"]:              # last: it names any address refused on the emails before it
            acfg = {**ncfg, "to": ncfg["admin_to"], "cc": list(ncfg.get("admin_cc") or []), "bcc": []}

            def admin_msg(acfg=acfg):
                acfg["refused"] = list(dict.fromkeys(a for _w, c, _m in msgs if c is not acfg
                                                     for a in (c.get("refused") or [])))
                subject, text, html_body, files = build_message(report, acfg, linker, trigger, extra_failures,
                                                                extra_files)
                return _compose(acfg, subject, text, html_body, files, outcome)
            msgs.append(("admin", acfg, admin_msg))
        # the administrator's email held back (admin_send_on) still goes when an email before it had an address
        # refused: it is the only one that says so. Decided when its turn comes (refusals are known by then).
        held = ["staff"] if staff_held and ncfg["to"] else []
        if not [w for w, _c, _m in msgs if not (w == "admin" and admin_held)]:
            return "not sent (send_on: problems)"
        if dry_run or os.environ.get("COMPLIANCE_NOTIFY") == "dry-run":
            out = []
            for who, _c, msg in msgs:
                if who == "admin" and admin_held:
                    continue
                msg = msg() if callable(msg) else msg
                f = Path(root) / "run_logs" / f"{report.get('run_utc', 'run')}__email_preview_{who}.eml"
                f.parent.mkdir(parents=True, exist_ok=True)
                f.write_bytes(msg.as_bytes())
                out.append(str(f))
            log.info("Email not sent (dry run); saved to %s", ", ".join(out))
            return f"dry run: {', '.join(out)}"
        sent, failed = [], []
        for who, c, msg in msgs:
            if who == "admin" and admin_held and not any(x.get("refused") for _w, x, _m in msgs if x is not c):
                held.append("admin")
                continue
            try:
                msg = msg() if callable(msg) else msg
                before = list(c.get("refused") or [])
                mid = _send_ses(c, msg) if c.get("transport", "ses") == "ses" else _send_smtp(c, msg)
            except Exception as e:                                  # noqa: BLE001
                log.error("Could not send the %s outcome email: %s", who, e, exc_info=True)
                failed.append((who, e))
                continue
            recipients = [a for a in c["to"] + c["cc"] if a.lower() not in (c.get("refused") or [])]
            sent.append(f"{', '.join(recipients) or 'Bcc only'} ({who})")
            log.info("Outcome email sent", extra={"details": {"outcome": outcome, "email": who,
                                                              "to": ", ".join(recipients),
                                                              "transport": c.get("transport"), "message_id": mid}})
            if who == "admin":          # the email that names them (the staff email only names a run's own: Nov 12)
                _remember_held_back(root, held_back(root)[1])        # each held-back file is mentioned once
            if who == "staff":
                from .whats_new import staff_remember
                staff_remember(root)
            if who == "admin" and len(c.get("refused") or []) > len(before):
                # refused on the administrator's own list: nothing else would say so (review M1); "not sent:" sends
                # the short alert. A count only: the address is in SES's answer, never in a log line
                n = len(c["refused"]) - len(before)
                failed.append((who, f"the email service refused {n} address(es) on the technical email's list "
                                    "(not verified in Amazon SES); it went to the others"))
        # a notice that failed goes again next time; one that missed a refused address doesn't (it would go to the
        # others on every run: review N1): the administrator's email names that address
        if sent and not any(who.startswith("recipients change") for who, _e in failed):
            admin_settings.remember_recipients(ncfg, root)
        held_note = "".join(f"; {who} email not sent (send_on: problems)" for who in held)
        if not sent and not failed:
            return "not sent (send_on: problems)"
        if not sent:
            return f"not sent: {failed[0][1]}"
        if failed:              # "not sent:" makes run.py send the short alert (to the administrator)
            return f"not sent: {failed[0][0]} email: {failed[0][1]} (sent to {'; '.join(sent)}){held_note}"
        return f"sent to {'; '.join(sent)}{held_note}"
    except Exception as e:
        log.error("Could not send the outcome email: %s", e, exc_info=True)
        return f"not sent: {e}"


# ------------------------------------------------------------------------------------- pipeline (AWS) entry
def pipeline_report(root, started_utc, failed_step="", run_rc=0, args="", run_projects=""):
    """Report for a pipeline run: the agent's run summary when the agent got far enough to write one, plus any
    pipeline step that failed (download, upload, mirror)."""
    root = Path(root)
    summary = None
    logs = root / "run_logs"
    if logs.exists() and failed_step not in ("lock", "config", "pull", "timeout"):
        cands = sorted(p for p in logs.glob("*__run_summary.json") if p.name[:19] >= started_utc)
        if cands:
            summary = cands[-1]
    if summary:
        report = json.loads(summary.read_text())
    else:
        report = {"run_utc": started_utc, "tasks": [], "projects": {}, "details": {}, "outcome": "FAILED",
                  "run_errors": []}
        report["tasks"] = _tasks_from_args(args)
    toks = (args or "").split()
    report["run_args"] = args or ""                      # the Options line (#30)
    report.setdefault("force", "--force" in toks)       # a run that stopped before its summary: the Action line (#30)
    # the pipeline's resolved projects, one per line; else from its arguments (one string: a name runs to the next
    # option, spaces and all)
    wanted = [x for x in str(run_projects or "").splitlines() if x.strip()] or \
        [m.strip() for m in re.findall(r"(?:^|\s)--(?:project\s+|plan\s+[a-z,]+:)(.+?)(?=\s--[A-Za-z]|$)", args or "")]
    report["requested_projects"] = (f"project{'s' if len(wanted) > 1 else ''} " + ", ".join(wanted)) if wanted \
        else "all projects"
    return report, summary


def _tasks_from_args(args):
    toks = (args or "").split()
    plans = [toks[i + 1].partition(":")[0] for i, t in enumerate(toks[:-1]) if t == "--plan"]
    if plans:                                  # --plan TASKS:PROJECT (one run, each project its own tasks)
        from .runner import parse_tasks
        try:
            return sorted(set().union(*(parse_tasks(t) for t in plans)))
        except ValueError:
            return []
    if "--task" in toks and toks.index("--task") + 1 < len(toks):
        from .runner import parse_tasks
        try:
            return sorted(parse_tasks(toks[toks.index("--task") + 1]))
        except ValueError:
            return []
    from .runner import TASKS
    return list(TASKS)


STEP_TEXT = {
    "lock": "Not run: another run was still in progress, so this one did not start. Start it again when that run "
            "has finished.",
    "config": "The run could not read its Google Drive settings (rclone.conf) from AWS Secrets Manager.",
    "pull": "The run could not download the compliance folder from Google Drive, so no task ran.",
    "agent": "The agent stopped unexpectedly before finishing, so nothing was uploaded (not even the projects it had "
             "finished: one may be half-done) and no file changed on the shared drive. Run it again; if it happens "
             "again, tell the administrator.",
    "start": "Not run: the agent stopped without finishing its run (usually before checking anything: it refused "
             "its arguments, couldn't find a project folder, or couldn't read a setting), so nothing was uploaded and "
             "no file changed on the shared drive. See the pipeline log, fix the cause, and start the run again.",
    "push": "The results could not be uploaded back to Google Drive (or the run was stopped while uploading them), "
            "so this run's reviewed files, tracker marks and logs are not on the shared drive, or only partly: "
            "projects are pushed one after another and the push stops at the first failure. Staff were told the "
            "results were not saved and to skip the review for now. Run the same projects again with Full rerun "
            "(--force): an Incremental run would skip payrolls whose upload already reached the drive as reviewed. "
            "Full rerun redoes only the newest payroll per project: if this run reviewed several for a project, "
            "the findings of the earlier ones are only in this email (Results below); pass them on to staff.",
    "mirror": "The copy to S3 failed (the results did reach Google Drive).",
    "timeout": "The run was stopped because it reached the Lambda time limit (15 minutes). Results of this run were "
               "not uploaded. Run fewer projects at a time, or move to EC2 (installation guide, section 1).",
    "tokenize": "Not run: the tokenizer did not finish before the sync (it failed, or was still running, possibly "
                "waiting for its nightly run), so nothing was synced and no task ran; all files are as they were. "
                "The next run tries again.",
    "stopped": "The run was stopped before it finished (it reached its time limit, or the server was shut down). "
               "Results of this run may be incomplete; run it again.",
}


TOKENIZER_WARNING = {
    "before": "Tokenizer: it could not be run before the sync, so this run reviewed the copy from the tokenizer's last "
              "run; files uploaded since then were not processed",
    "after": "Tokenizer: it could not be run after the sync, so the results are not in the staff folders yet; they "
             "will appear after the tokenizer's next run (it also runs nightly), or start it from the Cloud Run page",
}


def notify_pipeline(root, started_utc, failed_step="", run_rc=0, args="", pipeline_log="", trigger=None,
                    tokenizer_failed="", run_projects="", empty_projects="", unpulled_projects=""):
    ncfg = load_notify_config(root=root)
    report, summary = pipeline_report(root, started_utc, failed_step, run_rc, args, run_projects)
    report["pipeline_warnings"] = [TOKENIZER_WARNING[w] for w in str(tokenizer_failed or "").split()
                                   if w in TOKENIZER_WARNING and failed_step != "tokenize"]
    empty = [x for x in str(empty_projects or "").splitlines() if x.strip()]
    if empty and failed_step not in ("start",):
        report["pipeline_warnings"].append(
            f"Pull: {len(empty)} project(s) had no files on the tokenized drive (only empty folders or Google Docs), so "
            f"they were checked as empty folders: {', '.join(empty)}")
    # a full run: project folders with no file to copy aren't on the server at all, so the agent never saw them
    unpulled = [x for x in str(unpulled_projects or "").splitlines() if x.strip()]
    if unpulled and failed_step not in ("start",):
        report["pipeline_warnings"].append(
            f"Pull: {len(unpulled)} project(s) had no files on the tokenized drive (only empty folders or Google Docs), "
            f"so this run didn't check them: {', '.join(unpulled)}")
    extra = []
    if failed_step:
        extra.append({"title": "Pipeline", "text": STEP_TEXT.get(failed_step, f"Step '{failed_step}' failed."),
                      "log": "", "step": failed_step})
        if failed_step in ("lock", "tokenize", "start"):
            report["outcome"] = "NOT RUN"
        elif failed_step in ("config", "pull") or not summary:
            report["outcome"] = "FAILED"
        elif report.get("outcome") == "SUCCESS":
            report["outcome"] = "PARTIAL FAILURE"
    elif not summary:
        report["outcome"] = "FAILED"
        extra.append({"title": "Pipeline", "text": "The agent finished without writing a run summary.", "log": ""})
    if summary and not report.get("run_log"):
        rl = summary.with_name(summary.name.replace("__run_summary.json", "__run_log.txt"))
        report["run_log"] = str(rl) if rl.exists() else ""
    files = [pipeline_log] if pipeline_log and report.get("outcome") != "SUCCESS" else []
    if pipeline_log and failed_step:
        errs = []
        try:
            errs = [l.strip() for l in Path(pipeline_log).read_text(errors="replace").splitlines()
                    if ("error" in l.lower() or "failed" in l.lower() or l.startswith("Tokenizer:"))
                    and not l.strip().startswith(('"', "{", "}"))      # pieces of a multi-line JSON body
                    and "asking again" not in l                         # retries that later worked
                    and "sending the outcome email" not in l and " step '" not in l
                    and "error reporting:" not in l][-3:]
        except OSError:
            pass
        if errs:
            extra[-1]["text"] += " Error: " + " | ".join(errs)
        extra[-1]["text"] += f" Full output: the attached pipeline log ({Path(pipeline_log).name})."
    return send(report, root, trigger=trigger, ncfg=ncfg, extra_failures=extra, extra_files=files)


def send_test(root, ncfg=None, dry_run=False):
    """Sends a short test email with the current notifications.config (run.py --test-email)."""
    ncfg = ncfg or load_notify_config(root=root)
    ts = dt.datetime.now(dt.timezone.utc).strftime(TS_FORMAT)
    msg = EmailMessage()
    msg["Subject"] = f"{ncfg['subject_prefix']} Test email ({ts} UTC)".strip()
    msg["From"] = ncfg["from"] if "<" in ncfg["from"] else formataddr(("Compliance Validation Agent", ncfg["from"]))
    everyone = list(dict.fromkeys(ncfg["to"] + ncfg["admin_to"]))
    msg["To"] = ", ".join(everyone)
    if ncfg["cc"]:
        msg["Cc"] = ", ".join(ncfg["cc"])
    msg.set_content(f"This is a test from the Compliance Validation Agent on {socket.gethostname()}.\n"
                    f"Settings: {ncfg['path']} (transport: {ncfg.get('transport')}).\n"
                    f"Run-outcome emails: the staff email to {', '.join(ncfg['to']) or 'nobody'}; the full email to "
                    f"{', '.join(ncfg['admin_to']) or 'nobody'}.")
    if dry_run:
        return msg.as_string()
    if not everyone:
        raise ValueError(f"No recipients in {ncfg['path']}")
    tcfg = {**ncfg, "to": everyone, "refused": []}
    mid = _send_ses(tcfg, msg) if ncfg.get("transport", "ses") == "ses" else _send_smtp(tcfg, msg)
    refused = tcfg["refused"]               # a test must say who didn't get it (review M2)
    got = [a for a in everyone + ncfg["cc"] if a.lower() not in refused]
    return (f"sent to {', '.join(got)}" + (f" (SES message {mid})" if mid else "")
            + (f"; REFUSED (not verified in Amazon SES): {', '.join(refused)}" if refused else ""))
