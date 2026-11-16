"""Run the tasks for one, several or all projects. Writes run_logs/<ts>__run_summary.json, run_logs/<ts>__run_log.txt
(full trace) and a <ts>__log.txt in each task folder's logs/ (see tracelog.py). <ts> = yyyy-MM-dd-HH-mm-ss, UTC.
At the end it emails the outcome (see notify.py), unless COMPLIANCE_NOTIFY=defer (AWS pipeline: it emails after
uploading the results) or off."""
import json
import logging
import os
import platform
import sys
import time

from . import __version__, events
from .config import load_config
from .intent import run_intent_validation
from .layout import OldLayoutError, ensure_layout
from .llm import LocalLLM
from .notify import INACTIVE, IssueCollector, classify, describe_trigger, email_level, load_notify_config, send
from .payroll import run_payroll_review
from .tracelog import context, tracing
from .util import State, err_text, log, reset_run_timestamp, run_timestamp, start_run
from .wages import update_wage_data

TASKS = ("wages", "payroll", "intent")   # "all" = all three; payroll always merges new wage data first


def parse_tasks(task):
    items = {t.strip().lower() for t in (task or "all").replace("+", ",").split(",") if t.strip()}
    if "all" in items:
        return set(TASKS)
    bad = items - set(TASKS)
    if bad:
        raise ValueError(f"Unknown task(s): {', '.join(sorted(bad))}. Use: all, {', '.join(TASKS)}")
    if "payroll" in items:
        items.add("wages")
    return items


def list_projects(cfg):
    from .config import projects_folder
    root = cfg["projects_path"] = projects_folder(cfg["root"], cfg.get("projects_dir", "Projects"))   # renamed meanwhile?
    ignore = {x.lower() for x in cfg.get("ignore_project_dirs", [])}
    if not root.exists():
        return []
    return sorted(p for p in root.iterdir()
                  if p.is_dir() and not p.name.startswith((".", "_")) and p.name.lower() not in ignore)


def select_projects(cfg, wanted):
    projects = list_projects(cfg)
    if not wanted:
        return projects
    if isinstance(wanted, str):
        wanted = [wanted]
    out = []
    for w in wanted:
        hits = [p for p in projects if p.name == w] or \
               [p for p in projects if p.name.lower().startswith(w.lower())]
        if not hits:
            raise ValueError(f"No project folder matches '{w}' under {cfg['projects_path']}")
        if len(hits) > 1:                       # a short name must pick exactly one project
            raise ValueError(f"'{w}' matches several projects ({', '.join(h.name for h in hits)}); "
                             "use the full folder name")
        out += [h for h in hits if h not in out]
    return out


# #151 step 3: each project's status from its cover sheet ("Project Data..." in its staff folder), as the Google-side
# tokenizer leaves it in _tracker_sources.json (codes only: pii_tokenizer/cover_sheet.py). Closed and Not started
# projects are not checked when no project is named (the Saturday run) nor from a Run tasks plan (the page greys
# them out); a project named on the command line still runs. Anything unknown counts as active.
COVER_PROBLEMS = {    # why -> the technical email's words (Tracker sources.csv, staff side, names the files)
    "none": "no Project Data sheet in the staff project folder",
    "several": "more than one Project Data sheet in the staff project folder (see Tracker sources.csv)",
    "no_status": "the Project Data sheet has no Status row",
    "unrecognised": "the Project Data sheet's Status isn't Active, Closed or Not started (Tracker sources.csv shows "
                    "the word)",
    "conflict": "the Project Data sheet has two Status rows that disagree",
    "unreadable": "the Project Data sheet can't be read (save it as a Google Sheet or .xlsx)",
    "no_folder": "the staff project folder isn't known (no tracker match, or tracker_source.txt says none), so no "
                 "Project Data sheet was looked for",
}


def project_status(root, name):
    """(status, why) for a project folder: status is active, closed or not_started; why is the tokenizer's code, or
    None when there is no summary (the desktop app, the tracker step off, or a tokenizer from before Nov 12)."""
    from .intent import tracker_source
    entry = (tracker_source(root, name) or {}).get("project")
    if not isinstance(entry, dict):
        return "active", None
    status = entry.get("status") if entry.get("status") in ("active", *INACTIVE) else "active"
    why = entry.get("why") if isinstance(entry.get("why"), str) else None
    return status, why


def split_inactive(root, projects):
    """(projects to run, {status: [names] not run}, {why: [names]} cover-sheet problems), in the given order."""
    keep, not_run, problems = [], {}, {}
    for p in projects:
        status, why = project_status(root, p.name)
        if why in COVER_PROBLEMS or why == "error":
            problems.setdefault(why, []).append(p.name)
        if status in INACTIVE:
            not_run.setdefault(status, []).append(p.name)
        else:
            keep.append(p)
    return keep, not_run, problems


def cover_warnings(problems):
    """The technical email's lines on cover sheets that didn't give a status (the projects counted as active)."""
    out = [f"{COVER_PROBLEMS[why]}: {', '.join(names)}; counted as active" for why, names in problems.items()
           if why in COVER_PROBLEMS]
    if problems.get("error"):
        out.append("the Project Data sheet couldn't be read this time; the last status read was used (active if "
                   "none was read before): " + ", ".join(problems["error"]))
    return out


def _run_task(name, fn, project, lines, ran=None):
    """Run one task in its own log context; an error stops this task only. ``ran``: {(project, task): {"duration_ms",
    "error_type", "finished"}} for the structured run log's task lines (written after classify, from the summary)."""
    with context(project.name, name):
        t0, m0 = time.time(), time.monotonic()
        log.debug("Task started")                # (the project is in the record's project column; no path, #56)
        error_type, finished = None, False
        try:
            out = fn()
            out = out if isinstance(out, list) else str(out).split("\n")
            for line in out:
                level = logging.WARNING if any(w in line for w in ("WARNING", "GAP ", "SKIPPED", "could not", "none could", "ERROR")) else logging.INFO
                log.log(level, line)
            lines += out
            finished = True
        except Exception as e:
            log.error("Task failed: %s", err_text(e), exc_info=True)
            lines.append(f"ERROR ({name}): {e}")
            error_type, finished = type(e).__name__, True
        except BaseException as e:               # an interrupt, KeyboardInterrupt or SystemExit: not finished
            error_type = type(e).__name__
            raise
        finally:
            log.debug("Task finished", extra={"details": {"duration_s": round(time.time() - t0, 2)}})
            if ran is not None:
                ran[(project.name, name)] = {"duration_ms": int((time.monotonic() - m0) * 1000), "finished": finished,
                                             "error_type": error_type}


def task_events(report, ran, crashed=False):
    """The structured run log's task lines (docs/run_log.md), one per task run: the run summary's status for it
    (notify.classify: what the emails say), "crashed" when it raised. In a run that crashed (no summary), the tasks
    it ran before: "crashed" for the one interrupted, "unknown" for the others."""
    try:                                       # never stops a run (verification N5)
        details = report.get("details") or {}
        for (project, task), r in ran.items():
            d = (details.get(project) or {}).get(task) or {}
            if not r["finished"] or d.get("crashed"):
                outcome = "crashed"
            elif d.get("status"):              # classified (also when the run crashed later: verification N4)
                outcome = d["status"]
            else:
                outcome = "crashed" if r["error_type"] else "unknown"
            with context(project, task):
                events.emit("task", outcome=outcome, duration_ms=r["duration_ms"],
                            error_type=r["error_type"] if outcome == "crashed" else None)
    except Exception:                                                  # noqa: BLE001
        pass


def notify_config_path(cfg):
    """notifications.config next to the config.yaml in use, else the one in the agent folder."""
    p = cfg.get("config_path")
    if p and (p.parent / "notifications.config").exists():
        return p.parent / "notifications.config"
    return None


def parse_plan(items):
    """--plan TASKS:PROJECT values -> {project: "tasks"}, or None when there are none. A project named twice gets the
    tasks of both."""
    plan = {}
    for item in items or []:
        task, sep, name = str(item).partition(":")
        if not sep or not name.strip() or not task.strip():
            raise ValueError(f"--plan takes TASKS:PROJECT (e.g. payroll,intent:Brovane), not {item!r}")
        parse_tasks(task)                       # unknown task names: refused here
        name = name.strip()
        plan[name] = ",".join(x for x in (plan.get(name), task) if x)
    return plan or None


def plan_projects(cfg, plan):
    """{project name: tasks} (a "Run tasks" plan: each project its own tasks) -> {project folder: task set}, in the
    plan's order. Names are resolved as select_projects does; tasks as parse_tasks does (payroll brings wages)."""
    out = {}
    for name, task in plan.items():
        task = task if isinstance(task, str) else ",".join(task)
        for p in select_projects(cfg, name):
            out.setdefault(p, set()).update(parse_tasks(task))
    return out


def run(task="all", project=None, force=False, config_path=None, root=None, cfg=None, trigger=None, notify=True,
        plan=None):
    """One run: ``task`` for every project in ``project`` (all when None); or, with ``plan`` ({project: tasks}),
    each project its own tasks, still one run with one summary and one outcome email (task and project are then
    not used)."""
    cfg = cfg or load_config(config_path, root)
    tz = cfg.get("timezone", "UTC")
    start_run()
    ts = run_timestamp(tz)
    logs = cfg["root"] / "run_logs"
    logs.mkdir(exist_ok=True)
    # Two runs started within the same second would share a timestamp (and overwrite each other's log and
    # summary): wait for the next free second.
    while (logs / f"{ts}__run_log.txt").exists() or (logs / f"{ts}__run_summary.json").exists():
        time.sleep(0.25)
        reset_run_timestamp()
        ts = run_timestamp(tz)
    started = time.monotonic()
    not_run, problems = {}, {}
    if plan:
        asked = plan_projects(cfg, plan)
        keep, not_run, problems = split_inactive(cfg["root"], list(asked))     # a stale page can't run them
        asked = {p: t for p, t in asked.items() if p in keep}
        projects = list(asked)
        tasks = set().union(*asked.values()) if asked else set()
    else:
        asked = None
        tasks = parse_tasks(task)
        projects = select_projects(cfg, project)
        if not project:                       # every project (the Saturday run): not the Closed / Not started ones
            projects, not_run, problems = split_inactive(cfg["root"], projects)
    report = {"run_utc": ts, "tasks": sorted(tasks), "force": force, "trigger": describe_trigger(trigger),
              "requested": list(plan) if plan else
              ([project] if isinstance(project, str) else list(project or [])),   # the Options line (#30)
              "projects": {}}
    if asked is not None:       # each project's own tasks: the emails show only those (a blank cell for the others)
        report["plan"] = {p.name: sorted(t) for p, t in asked.items()}
    if not_run:
        report["not_run"] = not_run           # {status: [project names]} (#151 step 3)
    from .legacy import convert_root
    convert_root(cfg["root"])                     # older items: Pacific-time names -> UTC pattern (once)
    # Only the selected projects are touched (their folder layout is checked in the loop below), so one task
    # can be run for one project without changing anything in the others.
    report["run_log"] = str(logs / f"{ts}__run_log.txt")
    # the structured run log (docs/run_log.md; run_log.events, off by default): run_logs/<ts>__events.jsonl
    events.start(cfg, logs, ts, [p.name for p in projects], tasks)
    ran = {}                                      # the tasks run so far, for the run log's task lines
    try:
        return _run_projects(cfg, report, projects, asked, tasks, problems, logs, ts, tz, started, force, trigger,
                             notify, ran)
    except BaseException as e:
        task_events(report, ran, crashed=True)
        events.end("crashed", error_type=type(e).__name__)
        raise


def _run_projects(cfg, report, projects, asked, tasks, problems, logs, ts, tz, started, force, trigger, notify, ran):
    """The rest of run(): every project's tasks, the summary and the outcome email."""
    collector = IssueCollector()                  # warnings and errors, for the outcome email
    with tracing(ts, logs / f"{ts}__run_log.txt", {p.name: p for p in projects}, tz), collector:
        log.info("Run started", extra={"details": {
            "tasks": ",".join(sorted(tasks)), "projects": " | ".join(p.name for p in projects), "force": force,
            "root": cfg["root"], "agent_version": __version__, "python": sys.version.split()[0],
            "platform": platform.platform()}})
        llm = LocalLLM(cfg)
        if llm.available():
            report["model"] = {"name": llm.model, "id": llm.model_id}
        for p in projects:
            lines = []
            ptasks = asked[p] if asked is not None else tasks     # this project's tasks
            with context(p.name, ""):
                try:
                    for m in ensure_layout(p):
                        lines.append(f"layout: {m}")
                except OldLayoutError as e:
                    log.error("%s", err_text(e))
                    lines.append(f"ERROR (layout): {e}")
                    lines += [f"{t}: SKIPPED - the project is still in the old folder layout" for t in sorted(ptasks)]
                    report["projects"][p.name] = lines
                    continue                  # nothing runs on a project still in the old layout
                except Exception as e:
                    log.error("Folder layout check failed: %s", err_text(e), exc_info=True)
                    lines.append(f"ERROR (layout): {e}")
            state = State(p)          # loaded after the layout check, which may convert old timestamps
            with context(p.name, "upload"):
                try:
                    from .uploads import install_tracker_upload, unrecognized
                    if "intent" in ptasks:            # the tracker is only used (and replaced) by the intent check
                        installed = install_tracker_upload(p, tz)
                        lines += installed
                        if any(x.startswith("tracker: ") and "is now the current tracker" in x for x in installed):
                            # a new tracker, not an older copy of the last one: the marks remembered for the old
                            # one don't apply (and mustn't raise "replaced by an older copy", which needs them)
                            state.data.pop("intent_marks", None)     # (intent_last stays: the app's last check)
                    for f, why in unrecognized(p):
                        log.warning("Uploads: a %s file was not recognized; it stays in Uploads",
                                    f.suffix.lower() or "no-extension")    # (no file name or reason text: #56)
                        lines.append(f"WARNING upload: {f.name} not recognized ({why}); it stays in Uploads")
                        events.emit("file", file_n=events.next_file_n(p.name), file_kind="unrecognized",
                                    file_ext=events.file_ext(f), outcome="not_recognized")
                except Exception as e:
                    log.error("Could not sort the Uploads folder: %s", err_text(e), exc_info=True)
                    lines.append(f"ERROR (upload): {e}")
            took = report.setdefault("durations", {}).setdefault(p.name, {})   # seconds per task (for the app)

            def timed(name, fn):
                t0 = time.monotonic()
                _run_task(name, fn, p, lines, ran)
                took[name] = round(time.monotonic() - t0, 1)
            if "wages" in ptasks:
                timed("wages", lambda: update_wage_data(p, cfg, state, llm, tz) or ["wages: no new wage files"])
            if "payroll" in ptasks:
                timed("payroll", lambda: run_payroll_review(p, cfg, state, llm, force))
            if "intent" in ptasks:
                timed("intent", lambda: run_intent_validation(p, cfg, state, force=force))
            try:
                state.save()
            except Exception as e:
                with context(p.name, ""):
                    log.error("Could not save run state: %s", err_text(e), exc_info=True)
            report["projects"][p.name] = lines
        errors = sum(1 for v in report["projects"].values() for l in v if l.startswith("ERROR"))
        classify(report, collector.items, {p.name: p for p in projects})
        extra = cover_warnings(problems)      # technical email only; the projects were checked as active
        if extra:
            report["run_warnings"] = report.get("run_warnings", []) + extra
            report["warning_count"] = report.get("warning_count", 0) + len(extra)
            report["level"], report["level_reason"] = email_level(report)
        log.info("Run finished", extra={"details": {"projects": len(projects), "errors": errors,
                                                    "outcome": report["outcome"]}})
    try:                # the app's Results tab judges cells from this summary as the staff email does (review N1)
        from .notify import held_back_trackers, held_back_uploads
        report["held_back_trackers"] = held_back_trackers(cfg["root"])
        report["held_back_uploads"] = held_back_uploads(cfg["root"])      # project names and counts only
    except Exception as e:                                                  # noqa: BLE001
        log.warning("Could not read the held-back summary: %s", type(e).__name__)
    report["finished_utc"] = time.strftime("%Y-%m-%d-%H-%M-%S", time.gmtime())
    report["duration_s"] = round(time.monotonic() - started, 1)
    summary = logs / f"{ts}__run_summary.json"
    summary.write_text(json.dumps(report, indent=2, default=str))
    task_events(report, ran)
    events.end(report.get("outcome"), projects=len(projects),
               errors=sum(1 for v in report["projects"].values() for l in v if str(l).startswith("ERROR")),
               warnings=int(report.get("warning_count") or 0))
    if notify and os.environ.get("COMPLIANCE_NOTIFY", "").lower() not in ("defer", "off"):
        report["email"] = send(report, cfg["root"], trigger, load_notify_config(notify_config_path(cfg), cfg["root"]))
        summary.write_text(json.dumps(report, indent=2, default=str))
    return report
