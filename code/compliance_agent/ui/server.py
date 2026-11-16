"""Desktop UI: a tiny local web server (standard library only) + one HTML page.

Runs on this computer and opens the default browser. The page talks to a small JSON API.

  python run.py --ui            (or the packaged ComplianceAgent app)

The same server is also the hosted web app (CVA_HOSTED=1, on Google Cloud Run behind Identity-Aware Proxy): see
ui/hosted.py for what changes there (who may connect, the shared drive instead of a local folder, runs on AWS only).
"""
from __future__ import annotations

import json
import re
import socket
import logging
import os
import platform
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

BUNDLE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))   # PyInstaller or source
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else BUNDLE
sys.path.insert(0, str(BUNDLE))

from agent.config import load_config          # noqa: E402
from agent.layout import FILES, UPLOAD_DIR, Paths, archive_files, ensure_layout   # noqa: E402
from agent.runner import TASKS, list_projects, parse_tasks, run  # noqa: E402
from agent.util import run_timestamp, visible_files  # noqa: E402
from ui import hosted                                 # noqa: E402

SETTINGS_FILE = Path.home() / ".compliance_agent" / "settings.json"
# The checks offered on the page, in display order, named as in the staff email. "payroll" also adds new wage
# files first.
TASK_MENU = [
    {"key": "payroll", "label": "Payroll review",
     "hint": "Checks new payroll exports in the Uploads folder against the wage rates (new wage files are added first)."},
    {"key": "intent", "label": "Intent/affidavit check",
     "hint": "Compares the project's compliance tracker with L&I's intent and affidavit records."},
]
STATIC = BUNDLE / "ui" / "static"
log = logging.getLogger("compliance")


# ------------------------------------------------------------------ settings
def load_settings():
    try:
        return json.loads(SETTINGS_FILE.read_text())
    except Exception:
        return {}


def save_settings(s):
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(s, indent=2))


def config_path(settings):
    for p in (settings.get("config"), APP_DIR / "config.yaml", BUNDLE / "config.yaml"):
        if p and Path(p).is_file():
            return Path(p)
    return BUNDLE / "config.yaml"


def get_cfg(settings):
    root = settings.get("root")
    return load_config(config_path(settings), root) if root else None


def default_root():
    """The folder that holds projects/ when running from source (compliance_agent's parent)."""
    if getattr(sys, "frozen", False):
        return None
    try:
        cfg = load_config(BUNDLE / "config.yaml")
        return str(cfg["root"]) if cfg["projects_path"].is_dir() else None
    except Exception:
        return None


# ------------------------------------------------------------------- runs
class RunState:
    def __init__(self):
        self.lock = threading.Lock()
        self.lines, self.status, self.report, self.started, self.params = [], "idle", None, None, None

    def snapshot(self, since=0):
        return {"status": self.status, "lines": self.lines[since:], "next": len(self.lines),
                "report": self.report, "started": self.started, "params": self.params}


RUN = RunState()


# ------------------------------------------------------------------- queue
SEVERITY = {"ok": 0, "skipped": 1, "warning": 2, "errors": 3, "failed": 4}


def pair_status(detail: dict | None) -> str:
    """One (project, task) result as the app shows it: ok, warning, failed or skipped."""
    if not detail:
        return "unknown"
    st = detail.get("status", "ok")
    if st in ("failed", "errors"):
        return "failed"
    if st == "ok" and detail.get("warnings"):
        return "warning"
    return st


def _endpoint(path: str) -> str:
    """The request's endpoint without its query string, which can name a project and a file (#158)."""
    return urlparse(path or "").path[:100]


def _what(e) -> str:
    """An error as the log may show it: its type and HTTP status, never its text, which can name a file (#158)."""
    status = getattr(e, "status", None) or getattr(e, "code", None)
    return type(e).__name__ + (f" (HTTP {status})" if isinstance(status, int) else "")


class _PathFreeFormatter(logging.Formatter):
    """The log's traceback keeps the code locations and the error's type, never its text, which can name a file
    or carry a request's details (#158). The same for every handler: the console goes to Cloud Logging."""

    def format(self, record):
        saved = record.exc_text                 # logging caches the traceback text on the record: not this one's
        record.exc_text = None
        try:
            return super().format(record)
        finally:
            record.exc_text = saved

    def formatException(self, ei):
        frames = "".join(traceback.format_tb(ei[2])) if ei and ei[2] else ""
        return f"Traceback (most recent call last):\n{frames}{ei[0].__name__ if ei and ei[0] else 'error'} (text left out)"

    def formatStack(self, stack_info):
        return ""


class Server(ThreadingHTTPServer):
    """The HTTP server, whose own error handler goes through the log (and so the formatter that leaves out an error's
    text) instead of printing the whole traceback to the console (#158 review F1)."""
    daemon_threads = True

    def handle_error(self, request, client_address):
        log.error("A request failed: %s", _what(sys.exc_info()[1]), exc_info=True)


def setup_logging():
    """The console (Cloud Logging when hosted): UTC times, tracebacks without the error's text (#158)."""
    logging.basicConfig(level=logging.INFO)
    for h in logging.getLogger().handlers:
        h.setFormatter(log_formatter())


def log_formatter():
    """The console format: UTC time, and tracebacks without the error's text (#158)."""
    from agent.util import TS_FORMAT
    f = _PathFreeFormatter("%(asctime)s %(message)s", datefmt=TS_FORMAT)
    f.converter = time.gmtime
    return f


def pairs_from_report(report: dict, wanted: list[dict] | None = None) -> list[dict]:
    """[{project, task, status, seconds}] from a run summary. A payroll review includes its wage merge, so the
    worse of the two is shown for payroll when only payroll was asked for."""
    out = []
    details, took = report.get("details") or {}, report.get("durations") or {}
    try:
        from agent.staff_email import cell_results, merge_cells
        cells = cell_results(report)
    except Exception as e:                      # an old or odd summary: show what is known
        log.warning("Could not judge a run summary: %s", type(e).__name__)
        cells, merge_cells = {}, (lambda a, b: a)
    asked = wanted or [{"project": p, "task": t} for p, per in details.items() for t in per
                       if not t.startswith("_") and not (t == "wages" and "payroll" in per)]   # folded into payroll
    for w in asked:
        per = details.get(w["project"]) or {}
        st = pair_status(per.get(w["task"]))
        secs = (took.get(w["project"]) or {}).get(w["task"])
        folded = w["task"] == "payroll" and "wages" in per and (not wanted or not any(
            x["project"] == w["project"] and x["task"] == "wages" for x in wanted))
        if folded:
            ws = pair_status(per["wages"])
            if SEVERITY.get(ws, 0) > SEVERITY.get(st, 0):
                st = ws
            if secs is not None:
                secs = round(secs + ((took.get(w["project"]) or {}).get("wages") or 0), 1)
        if per.get("_project") and SEVERITY.get(st, 0) < SEVERITY["failed"]:
            st = "failed"                       # the project itself could not be prepared
        # the two verdicts the emails give (agent/staff_email.py): Run (did the task work) and Findings (did it pass)
        judged = cells.get((w["project"], w["task"])) or {}
        if folded and judged:                   # the wage update shown inside the payroll review: its problems too
            judged = merge_cells(judged, cells.get((w["project"], "wages")))
        out.append({**w, "status": st, "seconds": secs, "run": judged.get("run"), "findings": judged.get("findings"),
                    "note": judged.get("note", "")})
    return out


class JobQueue:
    """Runs are queued instead of refused: one runs at a time, the rest wait their turn, in order."""

    def __init__(self):
        self.lock = threading.Lock()
        self.waiting, self.current, self.history, self._n = [], None, [], 0
        self.wake = threading.Event()
        threading.Thread(target=self._worker, daemon=True).start()

    def submit(self, pairs, groups, where, force, arg0) -> dict:
        with self.lock:
            self._n += 1
            job = {"id": self._n, "pairs": pairs, "groups": groups, "where": where, "force": force,
                   "submitted": time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime()), "status": "waiting",
                   "_arg0": arg0}
            if hosted.HOSTED:
                # the web app: start it now, within this request (Cloud Run gives no CPU between requests, and
                # the server on AWS queues runs itself)
                position = 0
            else:
                self.waiting.append(job)
                position = len(self.waiting) + (1 if self.current else 0)
        if hosted.HOSTED:
            with self.lock:
                job["status"], job["started"] = "running", time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
                self.current = job
            self._process(job)
            return {"ok": RUN.status == "done", "job": job["id"], "position": 0, "runs": 1,
                    **({} if RUN.status == "done" else {"error": " ".join(RUN.lines[-1:]) or "The run could not be started."})}
        self.wake.set()
        return {"ok": True, "job": job["id"], "position": position, "runs": 1}

    def cancel(self, job_id) -> bool:
        with self.lock:
            for j in self.waiting:
                if j["id"] == job_id:
                    self.waiting.remove(j)
                    j["status"] = "removed"
                    return True
        return False

    def busy(self) -> bool:
        return bool(self.current or self.waiting)

    def snapshot(self) -> dict:
        pub = lambda j: {k: v for k, v in j.items() if not k.startswith("_") and k != "groups"} if j else None
        with self.lock:
            return {"running": pub(self.current), "waiting": [pub(j) for j in self.waiting],
                    "finished": [pub(j) for j in self.history[-15:]][::-1]}

    def _worker(self):
        while True:
            self.wake.wait()
            with self.lock:
                if not self.waiting:
                    self.wake.clear()
                    continue
                job = self.waiting.pop(0)
                job["status"], job["started"] = "running", time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
                self.current = job
            self._process(job)

    def _process(self, job):
        t0 = time.monotonic()
        try:
            self._run(job)
        except Exception as e:                       # never leave the queue stuck
            RUN.lines.append(f"ERROR: {e}")
            RUN.status = "error"
        job["seconds"] = round(time.monotonic() - t0, 1)
        job["finished"] = time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
        job["status"] = "done"
        rep = RUN.report or {}
        if job["where"] == "aws":
            st = "ok" if RUN.status == "done" else "failed"
            job["results"] = [{**w, "status": st, "seconds": None} for w in job["pairs"]]
            job["outcome"] = "sent to AWS" if RUN.status == "done" else "failed"
        else:
            reports = rep.get("runs") or ([rep] if rep else [])
            res = {}
            for r in reports:
                names = set((r.get("details") or {}))
                for x in pairs_from_report(r, [w for w in job["pairs"] if w["project"] in names
                                               and w["task"] in (r.get("tasks") or [])]):
                    res[(x["project"], x["task"])] = x
            job["results"] = [res.get((w["project"], w["task"]), {**w, "status": "failed" if RUN.status == "error"
                                                                     else "unknown", "seconds": None})
                              for w in job["pairs"]]
            worst = max((SEVERITY.get(x["status"], 0) for x in job["results"]), default=0)
            job["outcome"] = {0: "ok", 1: "ok", 2: "warning"}.get(worst, "failed")
        with self.lock:
            self.history.append(job)
            self.current = None

    def _run(self, job):
        RUN.lock.acquire()
        RUN.lines, RUN.report, RUN.status = [], None, "running"
        RUN.started = job["started"]
        RUN.params = {"job": job["id"], "pairs": job["pairs"], "force": job["force"], "where": job["where"],
                      "groups": [{"tasks": t, "projects": p} for t, p in job["groups"]]}
        if job["where"] == "aws" and not hosted.HOSTED:   # a Lambda call waits for the whole run (minutes), not 15 s
            try:
                RUN.params["waits_for_run"] = bool((load_settings().get("aws") or {}).get("lambda_function"))
            except Exception:                              # noqa: BLE001 - only the page's wording depends on it
                RUN.params["waits_for_run"] = False
        target = _run_aws if job["where"] == "aws" else _run_local
        target(job["_arg0"], None, None, job["force"], job["groups"])     # releases RUN.lock when done


QUEUE = JobQueue()


def recent_pairs(root: Path, runs=12) -> list[dict]:
    """The latest (project, task) results from the run summaries in run_logs/, whoever started them (this app,
    another computer, AWS, the schedule)."""
    out = []
    logs = Path(root) / "run_logs"
    for f in sorted(logs.glob("*__run_summary.json"), reverse=True)[:runs] if logs.exists() else []:
        try:
            r = json.loads(f.read_text())
        except Exception:
            continue
        for x in pairs_from_report(r):
            out.append({**x, "run_utc": r.get("run_utc"), "trigger": r.get("trigger") or "",
                        "run_seconds": r.get("duration_s")})
    return out


class _Capture(logging.Handler):
    def emit(self, record):
        from agent.tracelog import _project, _task
        where = " / ".join(x for x in (_project.get(), _task.get()) if x)
        RUN.lines.append(time.strftime("%Y-%m-%d-%H-%M-%S ", time.gmtime()) + (f"[{where}] " if where else "") + record.getMessage())


def plan_groups(plan):
    """[{project, tasks}] -> [(tasks, [projects])]: projects that share the same tasks (shown together on the page)."""
    groups = {}
    for item in plan:
        groups.setdefault(tuple(sorted(item["tasks"])), []).append(item["project"])
    return [(list(t), ps) for t, ps in groups.items()]


def groups_plan(groups):
    """[(tasks, [projects])] with more than one group -> [{project, tasks}]: one run in which each project has its own
    tasks (one summary, one outcome email), instead of one run per group. None for a single group (a plain run)."""
    if not groups or len(groups) < 2:
        return None
    return [{"project": p, "tasks": list(t)} for t, ps in groups for p in ps]


def _run_local(cfg, tasks, projects, force, groups=None):
    handler = _Capture(level=logging.INFO)     # the UI shows INFO and above; the CSV logs keep everything
    lg = logging.getLogger("compliance")
    lg.addHandler(handler)
    try:
        reports = []
        plan = groups_plan(groups)
        trigger = f"manual (desktop app on {socket.gethostname()})"
        if plan:                           # each project its own tasks: still one run, one email
            reports.append(run(force=force, cfg=cfg, trigger=trigger,
                               plan={i["project"]: ",".join(i["tasks"]) for i in plan}))
        for g_tasks, g_projects in ([] if plan else groups or [(tasks, projects)]):
            reports.append(run(",".join(g_tasks), g_projects, force, cfg=cfg, trigger=trigger))
        RUN.report = reports[0] if len(reports) == 1 else {
            "outcome": "SUCCESS" if all(x.get("outcome") == "SUCCESS" for x in reports) else "ERRORS",
            "runs": reports, "projects": {k: v for x in reports for k, v in x.get("projects", {}).items()},
            "email": "; ".join(x["email"] for x in reports if x.get("email")) or None}
        RUN.status = "done" if RUN.report.get("outcome") == "SUCCESS" else "error"
        if RUN.report.get("email"):
            RUN.lines.append(f"Outcome email: {RUN.report['email']}")
    except Exception as e:
        RUN.lines.append(f"ERROR: {e}")
        RUN.lines.append(traceback.format_exc())
        RUN.status = "error"
    finally:
        lg.removeHandler(handler)
        RUN.lock.release()


def _run_aws(settings, tasks, projects, force, groups=None):
    """Trigger the AWS deployment (Lambda function or EC2 instance via SSM): one run, with one outcome email, even
    when the projects have different tasks (a plan)."""
    try:
        groups = groups or [(tasks, projects)]
        plan = groups_plan(groups)
        g_tasks, g_projects = groups[0] if not plan else (None, None)
        _aws_once(settings, g_tasks, g_projects, force, plan=plan)
        if RUN.status != "error":
            RUN.status = "done"
    finally:
        RUN.lock.release()


def ec2_command(tasks, projects, force, who, plan=None):
    """The shell command SSM runs on the EC2 instance. Every value is quoted for the shell it lands in, so
    project names with spaces, dashes or apostrophes ("Brennick - 77Q") arrive as one argument each. plan:
    [{project, tasks}], each project its own tasks (run.py --plan), instead of tasks and projects."""
    import shlex
    if plan:
        args = [a for i in plan for a in ("--plan", f"{','.join(i['tasks'])}:{i['project']}")]
    else:
        args = ["--task", ",".join(tasks)] + [a for p in projects for a in ("--project", p)]
    args += ["--force"] if force else []
    # started in the background on the server (deploy/start_run.sh): Session Manager's time limit can't cut it short
    return shlex.join(["bash", "/opt/compliance-agent/deploy/start_run.sh", who, *args])


def _aws_once(settings, tasks, projects, force, plan=None):
    try:
        if hosted.HOSTED:
            try:
                cid = hosted.start_on_aws(tasks, projects, force, settings.get("user") or "unknown", plan=plan)
            except (RuntimeError, ValueError):
                raise                        # hosted.py words these for staff
            except Exception as e:           # AWS or network errors: the details go to the log, not to staff
                log.error("Could not send the run to AWS: %s", _what(e))
                RUN.lines.append("ERROR: The Compliance Validation Agent couldn't be reached. Please try again in a few minutes; "
                                 "if it keeps happening, tell the administrator.")
                RUN.status = "error"
                return
            RUN.lines.append(f"Started on the Compliance Validation Agent's server (command {cid}). Results appear in the shared drive "
                             "and in the outcome email when it finishes.")
            RUN.status = "done"
            return
        import boto3
        aws = settings.get("aws") or {}
        region = aws.get("region") or None
        if aws.get("lambda_function"):
            payload = {**({"plan": plan} if plan else {"task": ",".join(tasks), "project": projects}), "force": force,
                       "trigger": f"manual (desktop app on {socket.gethostname()}, Run on AWS)"}
            RUN.lines.append(f"Invoking Lambda {aws['lambda_function']} ... (this can take several minutes)")
            r = boto3.client("lambda", region_name=region).invoke(
                FunctionName=aws["lambda_function"], Payload=json.dumps(payload).encode())
            body = json.loads(r["Payload"].read() or b"{}")
            out = json.loads(body.get("body", "{}"))
            RUN.lines += (out.get("output") or "").splitlines()
            RUN.status = "done" if body.get("statusCode") == 200 else "error"
        elif aws.get("ec2_instance_id"):
            cmd = ec2_command(tasks, projects, force, f"manual (desktop app on {socket.gethostname()}, Run on AWS)",
                              plan=plan)
            r = boto3.client("ssm", region_name=region).send_command(
                InstanceIds=[aws["ec2_instance_id"]], DocumentName="AWS-RunShellScript",
                Parameters={"commands": [cmd]})
            RUN.lines.append(f"Started on EC2 {aws['ec2_instance_id']} (SSM command {r['Command']['CommandId']}). "
                             "Results appear in the shared folder when it finishes.")
            RUN.status = "done"
        else:
            RUN.lines.append("AWS is not configured (Settings > AWS).")
            RUN.status = "error"
    except Exception as e:
        RUN.lines.append(f"ERROR: {e}")
        RUN.status = "error"


# --------------------------------------------------------------- project info
def _files(d, exts=None):
    return [{"name": f.name, "size": f.stat().st_size, "modified": int(f.stat().st_mtime)}
            for f in visible_files(d, exts)] if d.exists() else []


_NA_CACHE: dict = {}                   # project -> ((tracker file, size, mtime, config setting), reason or None)
_NA_LOCK = threading.Lock()            # one tracker read at a time: a cold start doesn't open them all at once


def intent_not_applicable(p: Path, cfg, tracker_uploaded=False) -> str | None:
    """Why the intent/affidavit check doesn't apply to this project, in the staff's few words (the Results tab's
    words, agent/staff_email.py NA_REASON), or None when it applies or can't be told. Decided by the agent's own rule
    (agent/intent.py is_applicable) on the tracker the agent would use (the newest in Tracker/, as find_tracker).
    None, so the check stays on offer and the run says what happened, when: there is no tracker; it can't be read (in
    the hosted app, a tracker whose download failed); or a tracker is waiting in Uploads (only the intent check puts
    it in Tracker/, so a corrected tracker must never be locked out by the old one)."""
    from agent.intent import is_applicable
    from agent.staff_email import NA_OTHER, NA_REASON
    if tracker_uploaded:
        return None
    pc = (cfg.get("projects") or {}).get(p.name, {}) or {}
    tdir = Paths(p).tracker
    try:
        files = [f for f in visible_files(tdir, {".xlsx", ".xlsm"}) if "(AWS copy " not in f.name] \
            if tdir.is_dir() else []
        if not files:
            return None
        tracker = max(files, key=lambda f: (f.stat().st_mtime, f.name))
        st = tracker.stat()
        # the inode too: in the hosted mirror a placeholder (failed download) and the real file later downloaded
        # have the same name, size and date, but never the same inode (an unchanged file keeps its, hard-linked)
        key = (tracker.name, st.st_size, st.st_mtime_ns, st.st_ino, repr(pc.get("intent_applicable")))
    except OSError:
        return None
    with _NA_LOCK:
        hit = _NA_CACHE.get(p.name)
        if hit and hit[0] == key:
            return hit[1]
        try:
            import openpyxl
            # loaded exactly as the agent loads it (a streamed read can cut a header row short when the file's
            # stored sheet size is stale, and then disagree with the run); one at a time, under the lock
            wb = openpyxl.load_workbook(tracker, data_only=True)
            try:
                ok, why = is_applicable(p.name, wb, cfg)
            finally:
                wb.close()
            out = None if ok else NA_REASON.get(why, NA_OTHER)[1]
        except Exception:                         # never a path or a message in the log: names can be in them
            log.warning("Could not tell whether the intent/affidavit check applies to a project")
            out = None                            # remembered too: an unreadable file isn't reopened every refresh
        _NA_CACHE[p.name] = (key, out)            # only the latest version of each project's tracker is kept
        return out


# #151 step 3: each project's status from its cover sheet, as the tokenizer writes it in "Tracker sources.csv" at the
# top of the compliance folder (staff side; pii_tokenizer/staff_trackers.py staff_sheet). Closed and Not started
# projects are greyed out in Run tasks; anything else, or no file (the desktop app, the tracker step off), is active.
TRACKER_SOURCES_SHEET = "Tracker sources.csv"
INACTIVE_WORDS = {"closed": "Closed", "not started": "Not started"}


def project_statuses(root) -> dict:
    """{project folder name: "Closed" | "Not started"} for the projects that aren't active (others are left out)."""
    import csv
    import io
    try:
        text = (Path(root) / TRACKER_SOURCES_SHEET).read_bytes()[:2_000_000].decode("utf-8-sig", "replace")
        out = {}
        for row in csv.DictReader(io.StringIO(text)):
            word = " ".join(str(row.get("Project status") or "").lower().split())
            name = str(row.get("Project") or "")
            if name[:1] == "'" and name[1:2] in ("=", "+", "-", "@"):
                name = name[1:]               # the CSV's guard against a name read as a formula (staff_trackers._cell)
            if word in INACTIVE_WORDS and name:
                out[name] = INACTIVE_WORDS[word]
        return out
    except (OSError, csv.Error, ValueError):
        return {}


def project_summary(p: Path, cfg, statuses=None):
    paths = Paths(p)
    try:
        state = json.loads((p / ".agent_state.json").read_text())
    except Exception:
        state = {}
    from agent.layout import old_layout_dirs
    old_layout = bool(old_layout_dirs(p))     # files still in an old layout: runs skip it until migrate_layout.py
    from agent.uploads import classify
    inbox = {}                         # files waiting in Uploads/, by what the agent recognizes them as
    for f in visible_files(paths.upload) if paths.upload.is_dir() else []:
        kind = classify(f)[0] or "unknown"
        inbox.setdefault(kind, []).append({"name": f.name, "size": f.stat().st_size, "modified": int(f.stat().st_mtime)})
    current = {k: paths.rate_file(k).is_file() for k in ("federal", "state", "union")}
    pc = (cfg.get("projects") or {}).get(p.name, {}) or {}
    return {"name": p.name, "inbox": inbox, "wages_in_use": current,
            "tracker": [f["name"] for f in _files(paths.tracker, {".xlsx", ".xlsm"})],
            "reviewed": sorted((f["name"] for f in _files(paths.reviewed)), reverse=True),
            "last_payroll": (state.get("payroll_last") or {}).get("file"),
            "last_intent": state.get("intent_last"),
            "intent_applicable": pc.get("intent_applicable", "auto"),
            # tasks that don't apply to this project: {task key: why, in a few words}; the page greys them out
            "unavailable": {k: v for k, v in (("intent", intent_not_applicable(p, cfg, bool(inbox.get("tracker")))),)
                            if v},
            "needs_migration": old_layout,
            # Closed / Not started on the project's Project Data sheet: greyed out in Run tasks (#151 step 3)
            "inactive": (statuses or {}).get(p.name)}


def _inside(root: Path, p: Path):
    try:
        p.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _open(path: Path):
    if platform.system() == "Darwin":
        subprocess.Popen(["open", str(path)])
    elif platform.system() == "Windows":
        os.startfile(str(path))  # noqa
    else:
        subprocess.Popen(["xdg-open", str(path)])


def _pick_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog
        r = tk.Tk()
        r.withdraw()
        r.attributes("-topmost", True)
        d = filedialog.askdirectory(title="Choose the compliance folder (the one that contains 'Projects')")
        r.destroy()
        return d or None
    except Exception:
        return None


# ------------------------------------------------------------------ handler
LAST_PING = [time.time()]


class Handler(BaseHTTPRequestHandler):
    server_version = "ComplianceAgent/1.0"

    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        b = json.dumps(obj, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(b)

    def _err(self, msg, code=400):
        try:
            self._json({"error": msg}, code)
        except OSError:                         # the page went away meanwhile: nothing to tell it (#158 review F1)
            pass

    def _body(self, peek=False):
        """The request body (read once and kept, so a handler can look at it before another reads it)."""
        if not hasattr(self, "_cached_body"):
            n = int(self.headers.get("Content-Length") or 0)
            self._cached_body = self.rfile.read(n) if n else b""
        return self._cached_body

    def _ctx(self):
        if hosted.HOSTED:
            root = hosted.VIEW.refresh()
            settings = {"root": str(root), "aws": {"hosted": "1"}, "user": getattr(self, "user", "")}
            return settings, load_config(config_path(settings), root)
        settings = load_settings()
        if not settings.get("root") and default_root():
            settings["root"] = default_root()
        return settings, get_cfg(settings)

    def _allowed(self, change=False):
        """Desktop: this computer only. Hosted: a user signed in through IAP (and, for changes, the app's own page)."""
        if not hosted.HOSTED:
            return self._local_ok()
        try:
            self.user = hosted.verify_iap(self.headers)
        except hosted.Unauthorized as e:
            log.warning("Refused a request: %s", type(e).__name__)
            return False
        return hosted.origin_ok(self.headers) if change else True

    def _local_ok(self):
        # the API only answers this computer, and only the page served from it: another web page open in the
        # browser (Origin) or a site pointing its name at 127.0.0.1 (Host) is refused
        if self.client_address[0] not in ("127.0.0.1", "::1"):
            return False
        host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]").lower()
        if host and host not in ("127.0.0.1", "localhost", "::1"):
            return False
        origin = (self.headers.get("Origin") or "").lower()
        if origin == "null":                       # sandboxed frames and pages opened from a file
            return False
        if origin and not re.match(r"^https?://(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$", origin):
            return False
        return True

    def _admin_settings(self, change=None):
        """The Settings page (#151): web app only, administrators only (Managers of the shared drive). GET: the
        settings and the last changes; POST {"base": the file's rev seen, "changes": {setting: value or null}}."""
        from agent.cva_filekinds import settings_file as SF
        if not hosted.HOSTED:
            return self._err("The Settings page is in the web app only.", 404)
        if not hosted.VIEW.is_admin(getattr(self, "user", "")):
            return self._err("Only administrators (Managers of the compliance shared drive) can open the "
                             "Settings page.", 403)
        fid, raw = hosted.VIEW.read_settings()
        try:
            doc = SF.parse(raw) if raw else None
        except SF.Invalid:
            doc = None                     # unreadable: the page starts again from it (saving writes a new file)
        if change is not None:
            changes = change.get("changes")
            if not isinstance(changes, dict) or (not changes and not SF.page_values(doc)[1]):
                return self._err("Nothing to save.")         # (a save with no change removes unusable values)
            rev = (doc or {}).get("rev", 0)
            if (rev if type(rev) is int else 0) != (change.get("base") or 0):
                return self._err("The settings were changed meanwhile (by someone else, or in another window). "
                                 "Reload them and make your change again.", 409)
            try:
                new = SF.updated(doc, changes, self.user)
            except SF.Invalid as e:
                return self._err(str(e))
            hosted.VIEW.write_settings(json.dumps(new, indent=1, sort_keys=True).encode(), fid)
            log.info("Settings page: %d setting(s) saved", len(changes))       # never the values
            doc = new
        # The rev check is best-effort: a save between this read and the write (two administrators in the same
        # second, or two web app instances) can still be lost; the history shows what each save did.
        values, bad = SF.page_values(doc)          # a hand-edited bad value: shown as "can't be used", not a crash
        from agent.admin_settings import server_values
        server = server_values()                   # what each setting is without the page (review F3)
        return self._json({
            "invalid": bad,
            "settings": [{"key": k, "label": s["label"], "help": s["help"], "kind": s["kind"],
                          "choices": s.get("choices"), "min": s.get("min"), "max": s.get("max"),
                          "set": k in values, "value": values.get(k), "server_value": server.get(k)}
                         for k, s in SF.SETTINGS.items()],
            "rev": (doc or {}).get("rev", 0) if type((doc or {}).get("rev", 0)) is int else 0,
            "changed_at": (doc or {}).get("changed_at", ""), "changed_by": (doc or {}).get("changed_by", ""),
            "edited_outside_page": bool(doc) and SF.edited_outside_page(doc),
            "unreadable": bool(raw) and doc is None,
            "history": list(reversed(SF.history(doc)))[:15],
            "firm_domains": list(SF.FIRM_DOMAINS)})

    def do_GET(self):
        try:
            return self._get()
        except Exception as e:                       # e.g. the shared drive can't be read right now
            log.error("GET %s failed: %s", _endpoint(self.path), _what(e), exc_info=True)
            if getattr(e, "rate_limited", False):     # Google Drive said "slow down" for about a minute
                return self._err("Google Drive is busy right now. Wait a minute, then try again.", 503)
            if getattr(e, "quota", False):            # a daily quota (review N1): not a minute's wait
                return self._err("Google Drive's daily limit has been reached. Try again later; if it keeps "
                                 "happening, tell the administrator.", 503)
            return self._err(f"Something went wrong: {str(e)[:200]}", 500)

    def do_POST(self):
        try:
            return self._post()
        except Exception as e:
            log.error("POST %s failed: %s", _endpoint(self.path), _what(e), exc_info=True)
            if getattr(e, "rate_limited", False):     # Google Drive said "slow down" for about a minute
                return self._err("Google Drive is busy right now. Wait a minute, then try again.", 503)
            if getattr(e, "quota", False):            # a daily quota (review N1): not a minute's wait
                return self._err("Google Drive's daily limit has been reached. Try again later; if it keeps "
                                 "happening, tell the administrator.", 503)
            return self._err(f"Something went wrong: {str(e)[:200]}", 500)

    def _get(self):
        if not self._allowed():
            return self._err("forbidden", 403)
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        if u.path in ("/tasks", "/tasks.html"):        # the old task page is now part of the one page
            self.send_response(302)
            self.send_header("Location", "/")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if re.fullmatch(r"/brand/[a-z-]+\.png", u.path) and (STATIC / u.path.lstrip("/")).is_file():
            # the logo (ui/static/brand/, made by build_tools/brand_assets.py): fixed names, nothing else served
            b = (STATIC / u.path.lstrip("/")).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(b)))
            self.send_header("Cache-Control", "max-age=86400")
            self.end_headers()
            return self.wfile.write(b)
        if u.path in ("/", "/index.html"):
            b = (STATIC / "index.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            return self.wfile.write(b)
        if u.path == "/api/ping":
            LAST_PING[0] = time.time()
            return self._json({"ok": True})
        if u.path == "/api/state":
            settings, cfg = self._ctx()
            out = {"root": settings.get("root"), "aws": settings.get("aws") or {}, "tasks": list(TASKS),
                   "hosted": hosted.HOSTED, "user": getattr(self, "user", None),
                   "admin": bool(hosted.HOSTED and hosted.VIEW.is_admin(getattr(self, "user", ""))),
                   "task_menu": TASK_MENU,
                   "upload_kinds": ["auto"], "projects": [], "problem": None,
                   "run": RUN.snapshot(10**9)}
            if not cfg:
                out["problem"] = "Choose the compliance folder (the one that contains 'Projects')."
            elif not cfg["projects_path"].is_dir():
                out["problem"] = f"No 'Projects' folder inside {cfg['root']}."
            else:
                projs = list_projects(cfg)
                if not hosted.HOSTED:      # the hosted app only reads its mirror of the drive
                    from agent.legacy import convert_root
                    try:
                        convert_root(cfg["root"])
                    except Exception as e:
                        log.warning("Could not update run_logs: %s", _what(e))
                    for p in projs:        # create missing folders (a project in an old layout is only reported)
                        try:
                            ensure_layout(p)
                        except Exception as e:
                            log.warning("Could not update folders of %s: %s", p.name, _what(e))
                statuses = project_statuses(cfg["root"])
                out["projects"] = [project_summary(p, cfg, statuses) for p in projs]
                logs = cfg["root"] / "run_logs"
                out["recent_runs"] = []
                for f in sorted(logs.glob("*__run_summary.json"), reverse=True)[:8] if logs.exists() else []:
                    try:
                        out["recent_runs"].append(json.loads(f.read_text()))
                    except Exception:
                        pass
            return self._json(out)
        if u.path == "/api/admin/settings":
            return self._admin_settings()
        if u.path == "/api/run":
            return self._json(RUN.snapshot(int(q.get("since", 0))))
        if u.path == "/api/queue":
            settings, cfg = self._ctx()
            out = QUEUE.snapshot()
            out["recent"] = recent_pairs(cfg["root"]) if cfg else []
            return self._json(out)
        return self._err("not found", 404)

    def _post(self):
        # every change must come from the app's own page: it sends X-Compliance-App, which a browser only lets
        # another page send after asking the app (a CORS preflight this server never approves)
        if not self._allowed(change=True) or self.headers.get("X-Compliance-App") != "1":
            return self._err("forbidden", 403)
        u = urlparse(self.path)
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        settings, cfg = self._ctx()
        if hosted.HOSTED and u.path in ("/api/settings", "/api/pick-folder", "/api/quit", "/api/open"):
            return self._err("Not available in the web app: open the file in Google Drive instead.", 403)
        if u.path == "/api/admin/settings":
            return self._admin_settings(json.loads(self._body() or b"{}"))

        if u.path == "/api/settings":
            data = json.loads(self._body() or b"{}")
            if "root" in data:
                root = Path(os.path.expanduser(data["root"].strip()))
                from agent.config import projects_folder
                if projects_folder(root).is_dir():
                    pass
                elif root.name.lower() == "projects" and root.is_dir():
                    root = root.parent
                else:
                    return self._err(f"'{root}' has no 'Projects' folder inside it.")
                settings["root"] = str(root)
            if "aws" in data:
                settings["aws"] = {k: v.strip() for k, v in data["aws"].items() if isinstance(v, str)}
            save_settings(settings)
            return self._json({"ok": True, "root": settings.get("root")})

        if u.path == "/api/pick-folder":
            d = _pick_folder()
            return self._json({"path": d}) if d else self._err("Folder picker isn't available here; type the path instead.")

        if u.path == "/api/quit":
            self._json({"ok": True})
            threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
            return

        if not cfg:
            return self._err("Choose the compliance folder first.")
        projects = {p.name: p for p in list_projects(cfg)}

        if u.path == "/api/projects":
            name = (json.loads(self._body() or b"{}").get("name") or "").strip()
            if not name or any(c in name for c in '/\\:*?"<>|') or name.startswith((".", "_")):
                return self._err("Enter a valid project name (no / \\ : * ? \" < > |).")
            if hosted.HOSTED and not hosted.project_name_ok(name):
                return self._err("Use letters, digits, spaces and . , & ( ) ' _ - only (the server runs only such names).")
            p = cfg["projects_path"] / name
            if p.exists():
                return self._err("A project with that name already exists.")
            if hosted.HOSTED:
                try:
                    hosted.VIEW.new_project(name)
                except FileExistsError:
                    return self._err("A project with that name already exists.")
                return self._json({"ok": True, "name": name})
            p.mkdir(parents=True)
            ensure_layout(p)
            return self._json({"ok": True, "name": name})

        if u.path == "/api/upload":
            # Every file goes to the project's Uploads/ folder; the next run recognizes what it is.
            proj, name = q.get("project"), unquote(q.get("name", ""))
            if proj not in projects:
                return self._err(f"Unknown project '{proj}'.")
            name = Path(name).name
            if not name or name.startswith((".", "~$")):
                return self._err("Invalid file name.")
            ok_ext = {".xlsx", ".xlsm", ".csv", ".txt", ".pdf"}
            if Path(name).suffix.lower() not in ok_ext:
                return self._err("only Excel, CSV, text and PDF files can be added (.xlsx, .xlsm, .csv, .txt, .pdf).")
            p = projects[proj]
            if hosted.HOSTED:
                from agent.layout import old_layout_dirs
                if old_layout_dirs(p):
                    return self._err(f"{proj}'s folders need fixing before files can be added: please ask the administrator.")
                note = hosted.VIEW.upload(proj, name, self._body(), time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime()))
                return self._json({"ok": True, "saved_to": f"{proj}/{UPLOAD_DIR}/{name}", "note": note})
            try:
                ensure_layout(p)
            except Exception as e:                   # still in the old folder layout: run migrate_layout.py first
                log.warning("Upload refused, %s's folders: %s", proj, _what(e))
                return self._err(f"{proj}'s folders need fixing before files can be added: please ask the administrator.")
            dest_dir = Paths(p).upload
            dest_dir.mkdir(parents=True, exist_ok=True)
            data = self._body()
            from agent.util import reset_run_timestamp
            reset_run_timestamp()
            note = ""
            if (dest_dir / name).exists():
                archive_files(p, [dest_dir / name], FILES, move=True, label="replaced upload")
                note = "an unprocessed file with the same name was already in Uploads; it was moved to the archive"
            (dest_dir / name).write_bytes(data)
            return self._json({"ok": True, "saved_to": f"{proj}/{UPLOAD_DIR}/{name}", "note": note})

        if u.path == "/api/remove":
            data = json.loads(self._body() or b"{}")
            proj, name = data.get("project"), Path(data.get("name", "")).name
            if proj not in projects:
                return self._err("Unknown project.")
            f = Paths(projects[proj]).upload / name
            if not f.is_file():
                return self._err("Only unprocessed files in Uploads can be removed, and this one is not there.")
            if hosted.HOSTED:
                try:
                    dest = hosted.VIEW.remove(proj, name, time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime()))
                except FileNotFoundError:
                    return self._err("Only unprocessed files in Uploads can be removed, and this one is not there.")
                return self._json({"ok": True, "note": f"moved to {proj}/{dest}"})
            from agent.util import reset_run_timestamp
            reset_run_timestamp()
            dest = archive_files(projects[proj], [f], FILES, move=True, label="removed")[f]
            return self._json({"ok": True, "note": f"moved to {proj}/{dest.relative_to(projects[proj]).as_posix()}"})

        if u.path == "/api/open":
            data = json.loads(self._body() or b"{}")
            proj = data.get("project")
            target = cfg["root"] if not proj else projects.get(proj)
            if target is None:
                return self._err("Unknown project.")
            if data.get("path"):
                target = target / data["path"]
            if not _inside(cfg["root"], target) or not target.exists():
                return self._err("Not found.")
            _open(target)
            return self._json({"ok": True})

        if u.path == "/api/run" and "plan" in json.loads(self._body(peek=True) or b"{}"):
            # Task page: a plan of tasks per project, e.g. [{"project": "A", "tasks": ["payroll", "intent"]}, ...]
            data = json.loads(self._body() or b"{}")
            known = {t["key"] for t in TASK_MENU}
            plan, problems = [], []
            for item in data.get("plan") or []:
                name, tasks = item.get("project"), [t for t in item.get("tasks") or [] if t in known]
                if name not in projects:
                    problems.append(f"unknown project '{name}'")
                elif tasks:
                    plan.append({"project": name, "tasks": tasks})
            if problems:
                return self._err("Can't run: " + "; ".join(problems) + ". Reload the page.")
            if not plan:
                return self._err("Pick at least one task for at least one project.")
            groups = plan_groups(plan)
            where = "aws" if hosted.HOSTED else data.get("where", "local")
            pairs = [{"project": i["project"], "task": t} for i in plan for t in i["tasks"]]
            res = QUEUE.submit(pairs, groups, where, bool(data.get("force")), settings if where == "aws" else cfg)
            return self._err(res["error"], 409) if res.get("error") else self._json(res)

        if u.path == "/api/run":
            data = json.loads(self._body() or b"{}")
            try:
                tasks = sorted(parse_tasks(",".join(data.get("tasks") or [])) if data.get("tasks") else [])
            except ValueError as e:
                return self._err(str(e))
            sel = [p for p in data.get("projects") or [] if p in projects]
            if not tasks:
                return self._err("Pick at least one task.")
            if not sel:
                return self._err("Pick at least one project.")
            where = "aws" if hosted.HOSTED else data.get("where", "local")
            pairs = [{"project": p, "task": t} for p in sel for t in tasks]
            res = QUEUE.submit(pairs, [(tasks, sel)], where, bool(data.get("force")), settings if where == "aws" else cfg)
            return self._err(res["error"], 409) if res.get("error") else self._json(res)

        if u.path == "/api/queue/remove":
            data = json.loads(self._body() or b"{}")
            if not QUEUE.cancel(int(data.get("job") or 0)):
                return self._err("That run has already started or finished.")
            return self._json({"ok": True})

        return self._err("not found", 404)


def _watchdog(server, idle_minutes=20):
    """Quit when the page has been closed for a while (no window to close on a packaged app)."""
    while True:
        time.sleep(30)
        if time.time() - LAST_PING[0] > idle_minutes * 60 and RUN.status != "running" and not QUEUE.busy():
            server.shutdown()
            return


def main(root=None, config=None, port=None, open_browser=True):
    setup_logging()
    s = load_settings()
    changed = False
    if root:
        s["root"], changed = str(Path(root).expanduser().resolve()), True
    if config:
        s["config"], changed = str(Path(config).resolve()), True
    if changed:
        save_settings(s)
    if hosted.HOSTED:                          # Cloud Run: every interface, the port it gives; no browser, no idle exit
        from run import _error_reporting       # Sentry (agent/errors.py): the hosted app only, never the desktop app
        _error_reporting("web", config)
        server = Server(("0.0.0.0", int(os.environ.get("PORT", 8080))), Handler)
        print(f"Compliance Validation Agent web app on port {server.server_address[1]}")
        server.serve_forever()
        return 0
    port = port or int(os.environ.get("COMPLIANCE_UI_PORT", 0)) or 8765
    for p in range(port, port + 20):
        try:
            server = Server(("127.0.0.1", p), Handler)
            break
        except OSError:
            continue
    else:
        raise SystemExit("No free port for the UI")
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"Compliance Validation Agent running at {url}  (close with the Quit button or Ctrl+C)")
    threading.Thread(target=_watchdog, args=(server,), daemon=True).start()
    if open_browser and not os.environ.get("COMPLIANCE_UI_NO_BROWSER"):
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
