"""The structured run log (docs/run_log.md; item 3 of the reliability plan): one JSON line per step of a run, in
run_logs/<ts>__events.jsonl, when the setting run_log.events is on (off by default).

Privacy is enforced here, not promised by the callers: every field of every line is checked against a fixed schema
(SCHEMA) and against the fields its event may carry (ALLOWED). Unknown keys, values that don't match their type, enum
or pattern, and project names other than the run's own project folders are dropped and counted ("dropped" on the
line, and in run_end). So no path, file name, cell text, worker name, token or exception message can get in, whatever
a caller passes. Writing never raises and never changes a run: an error here is counted in run_end as
"events_failed".

    start(cfg, logs_dir, ts, projects, tasks)   at the start of a run (runner.run)
    emit("task" | "file" | "ai_review", **fields) during it; project and task default to the log context
    end(outcome, **counts)                      at the end (also after a crash)
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import secrets
import threading
import time

V = 1
RUN_ID_RE = re.compile(r"[0-9a-f]{12}")
TASKS = {"wages", "payroll", "intent", "upload"}
OUTCOMES = {
    "file": {"reviewed", "skipped", "not_recognized", "error"},
    "task": {"ok", "skipped", "errors", "failed", "crashed", "unknown"},     # notify.classify's statuses
    "run_end": {"success", "partial_failure", "failed", "not_run", "crashed"},
}
REASONS = {"no_wage_rates", "no_jurisdiction_column", "wage_type_unknown",
           "no_payroll_lines", "repeated_header"}                    # the last two: agent/upload_checks.py
FILE_KINDS = {"payroll", "unrecognized"}
FILE_EXTS = {"xlsx", "xlsm", "xls", "csv", "pdf", "none", "other"}
AI_MODES = {"claude", "local"}
AI_STOPPED = {"call_failed", "client_failed", "call_limit", "time_limit"}
RULE_RE = re.compile(r"[RS]\d\d")
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]{0,99}")
VERSION_RE = re.compile(r"[\w.+-]{1,40}")
MODEL_RE = re.compile(r"[a-z0-9][a-z0-9.-]{0,63}")
EXECUTION_RE = re.compile(r"[a-z0-9][a-z0-9-]{0,62}")
MAX_INT = 10 ** 12


def _int(v):
    return isinstance(v, int) and not isinstance(v, bool) and 0 <= v < MAX_INT


def _enum(values):
    return lambda v: isinstance(v, str) and v in values


def _pattern(rx):
    return lambda v: isinstance(v, str) and rx.fullmatch(v) is not None


def _flags(v):
    return (isinstance(v, dict) and len(v) <= 40
            and all(isinstance(k, str) and RULE_RE.fullmatch(k) and _int(n) for k, n in v.items()))


def _tasks(v):
    return isinstance(v, (list, tuple)) and len(v) <= len(TASKS) and all(isinstance(t, str) and t in TASKS for t in v)


# field -> check (the project check needs the run's projects: see _ok)
SCHEMA = {
    "task": _enum(TASKS), "tasks": _tasks, "file_n": _int, "file_kind": _enum(FILE_KINDS),
    "file_ext": _enum(FILE_EXTS), "reason": _enum(REASONS), "rows": _int, "flags": _flags, "duration_ms": _int,
    "error_type": _pattern(IDENT_RE), "gen_ai.request.model": _pattern(MODEL_RE), "ai.mode": _enum(AI_MODES),
    "ai.patterns": _int, "ai.calls": _int, "ai.cached": _int, "ai.skipped": _int, "ai.no_evidence": _int,
    "ai.unclear": _int, "gen_ai.usage.input_tokens": _int, "gen_ai.usage.output_tokens": _int,
    "ai.stopped": _enum(AI_STOPPED), "tokenizer_before": _pattern(EXECUTION_RE),
    "agent_version": _pattern(VERSION_RE), "python": _pattern(VERSION_RE),
    "projects": _int, "files": _int, "errors": _int, "warnings": _int, "events_failed": _int,
    "project": None, "outcome": None,
}
ALLOWED = {
    "run_start": {"tasks", "projects", "agent_version", "python", "tokenizer_before"},
    "task": {"project", "task", "outcome", "duration_ms", "error_type"},
    "file": {"project", "task", "file_n", "file_kind", "file_ext", "outcome", "reason", "rows", "flags",
             "duration_ms", "error_type"},
    "ai_review": {"project", "task", "duration_ms", "gen_ai.request.model", "ai.mode", "ai.patterns", "ai.calls",
                  "ai.cached", "ai.skipped", "ai.no_evidence", "ai.unclear", "gen_ai.usage.input_tokens",
                  "gen_ai.usage.output_tokens", "ai.stopped"},
    "run_end": {"outcome", "duration_ms", "error_type", "projects", "files", "errors", "warnings", "events_failed"},
}
EVENTS = set(ALLOWED)
# The run summary's outcome (notify.classify) -> run_end's
RUN_OUTCOMES = {"SUCCESS": "success", "PARTIAL FAILURE": "partial_failure", "FAILED": "failed", "NOT RUN": "not_run"}

_lock = threading.Lock()
_run: dict = {}


def enabled(cfg) -> bool:
    """The setting run_log.events (off unless it is exactly true)."""
    section = (cfg or {}).get("run_log") if isinstance(cfg, dict) else None
    return isinstance(section, dict) and section.get("events") is True


def run_id_from_env() -> str:
    """CVA_RUN_ID (run_pipeline.sh, D1) when it is a valid ID, else ""."""
    rid = os.environ.get("CVA_RUN_ID", "").strip().lower()
    return rid if RUN_ID_RE.fullmatch(rid) else ""


def file_ext(path) -> str:
    """A file's extension as one of FILE_EXTS (never the name)."""
    ext = os.path.splitext(str(path))[1].lower().lstrip(".")
    if not ext:
        return "none"
    return ext if ext in FILE_EXTS else "other"


def active() -> bool:
    return bool(_run.get("on"))


def _sentry_run_id(rid):
    """D9: a Sentry error names its run. ``rid`` "" puts back the pipeline's ID (CVA_RUN_ID), or none, so a later
    error in a long-lived process (the web app) never carries an earlier run's ID. No-op when Sentry is off."""
    try:
        from .errors import clear_run_id, set_run_id
        if rid:
            set_run_id(rid)
        else:
            clear_run_id()
    except Exception:                                                  # noqa: BLE001
        pass


def start(cfg, logs_dir, ts, projects, tasks=()) -> str:
    """Begins the run's events file (when the setting is on). Returns the run ID ("" when off). Never raises."""
    try:
        with _lock:
            _run.clear()
            on = enabled(cfg)
            if on:
                rid = run_id_from_env() or secrets.token_hex(6)
                _run.update(on=True, run_id=rid, path=os.path.join(str(logs_dir), f"{ts}__events.jsonl"),
                            projects={str(p) for p in projects}, file_n={}, dropped=0, failed=0, files=0,
                            t0=time.monotonic(), ended=False)
        _sentry_run_id(rid if on else "")
        if not on:
            return ""
        import platform
        from . import __version__
        emit("run_start", tasks=sorted(t for t in tasks if t in TASKS), projects=len(_run["projects"]),
             agent_version=str(__version__), python=platform.python_version(),
             tokenizer_before=os.environ.get("CVA_TOKENIZER_BEFORE", "").strip() or None)
        return rid
    except Exception:                                                  # noqa: BLE001 - never stops a run
        with _lock:
            _run.clear()
        return ""


def next_file_n(project) -> int:
    """1, 2, 3... per project in this run (0 when off). Never raises."""
    try:
        if not active():
            return 0
        with _lock:
            n = _run["file_n"].get(project, 0) + 1
            _run["file_n"][project] = n
            return n
    except Exception:                                                  # noqa: BLE001
        return 0


def _context():
    try:
        from .tracelog import _project, _task
        return _project.get(), _task.get()
    except Exception:                                                  # noqa: BLE001
        return "", ""


def _ok(key, value, event) -> bool:
    if key == "project":
        return isinstance(value, str) and value in _run.get("projects", ())
    if key == "outcome":
        return isinstance(value, str) and value in OUTCOMES.get(event, ())
    check = SCHEMA.get(key)
    return check is not None and check(value)


def emit(event, **fields) -> None:
    """Writes one line (when on). Fields that fail the schema are dropped and counted; None values are left out
    without counting. Never raises."""
    if not active():
        return
    try:
        if event not in EVENTS:
            with _lock:
                _run["dropped"] += 1
            return
        if event in ("task", "file", "ai_review"):
            project, task = _context()
            fields.setdefault("project", project)
            if task:
                fields.setdefault("task", task)
        line, dropped = {}, 0
        for k, v in fields.items():
            if v is None:
                continue
            if k in ALLOWED[event] and _ok(k, v, event):
                line[k] = dict(v) if isinstance(v, dict) else (list(v) if isinstance(v, tuple) else v)
            else:
                dropped += 1
        with _lock:
            if not _run.get("on"):
                return
            _run["dropped"] += dropped
            if event == "file":
                _run["files"] += 1
            line.update(v=V, ts=dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                        run_id=_run["run_id"], event=event)
            if event == "run_end":                 # the whole run's dropped fields, this line's included
                line["dropped"] = _run["dropped"]
            elif dropped:
                line["dropped"] = dropped
            text = json.dumps(line, sort_keys=True, separators=(",", ":"), ensure_ascii=True)
            with open(_run["path"], "a", encoding="utf-8") as f:
                f.write(text + "\n")
    except Exception:                                                  # noqa: BLE001 - never stops a run
        try:
            with _lock:
                _run["failed"] = _run.get("failed", 0) + 1
        except Exception:                                              # noqa: BLE001
            pass


def end(outcome, error_type=None, **counts) -> None:
    """Writes run_end once (when on) and closes the run. ``outcome``: a run summary outcome ("SUCCESS"...), or one
    of OUTCOMES["run_end"]. Never raises."""
    try:
        if not active() or _run.get("ended"):
            return
        out = RUN_OUTCOMES.get(outcome, outcome)
        took = int((time.monotonic() - _run["t0"]) * 1000)
        emit("run_end", outcome=out, error_type=error_type, duration_ms=took, files=_run["files"],
             events_failed=_run["failed"], **counts)
    except Exception:                                                  # noqa: BLE001 - never stops a run
        pass
    finally:
        try:
            with _lock:
                was_on = bool(_run.get("on"))
                _run["ended"] = True
                _run["on"] = False
            if was_on:
                _sentry_run_id("")
        except Exception:                                              # noqa: BLE001
            pass


def totals() -> dict:
    """What run_end will count (for tests): dropped fields and failed writes so far."""
    return {"dropped": _run.get("dropped", 0), "events_failed": _run.get("failed", 0)}


BASE_KEYS = {"v", "ts", "run_id", "event", "dropped"}
TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z")


def _no_repeats(pairs):
    keys = [k for k, _ in pairs]
    if len(keys) != len(set(keys)):
        raise ValueError("a key twice")           # json.loads would keep the last value, the raw text both
    return dict(pairs)


def checked_lines(path, projects):
    """The lines of an events file, each re-written from its checked values, when every line is one this module would
    have written for a run on ``projects`` (the run's project folder names): a JSON object with a known event, only
    its allowed fields, each passing the schema, no key twice, one run ID throughout, at most 10,000 lines. None
    otherwise. For a copy leaving the server (run_pipeline.sh: CloudWatch), so that a file the agent didn't write, or
    one changed since, is never sent, and what is sent is only text made from checked values. Never raises."""
    try:
        projects = {str(p) for p in projects}
        rid, out = None, []
        with open(path, encoding="utf-8") as f:
            for i, raw in enumerate(f):
                if i >= 10000:
                    return None
                line = json.loads(raw, object_pairs_hook=_no_repeats)
                if not isinstance(line, dict):
                    return None
                event = line.get("event")
                if event not in EVENTS or line.get("v") != V or not isinstance(line.get("ts"), str) \
                        or not TS_RE.fullmatch(line["ts"]):
                    return None
                if not isinstance(line.get("run_id"), str) or not RUN_ID_RE.fullmatch(line["run_id"]):
                    return None
                if rid is None:
                    rid = line["run_id"]
                elif line["run_id"] != rid:
                    return None
                if "dropped" in line and not _int(line["dropped"]):
                    return None
                for k, v in line.items():
                    if k in BASE_KEYS:
                        continue
                    if k not in ALLOWED[event]:
                        return None
                    if k == "project":
                        if not (isinstance(v, str) and v in projects):
                            return None
                    elif k == "outcome":
                        if not (isinstance(v, str) and v in OUTCOMES.get(event, ())):
                            return None
                    elif not SCHEMA[k](v):
                        return None
                out.append(json.dumps(line, sort_keys=True, separators=(",", ":"), ensure_ascii=True))
        return out if rid is not None else None
    except Exception:                                                  # noqa: BLE001
        return None


def file_ok(path, projects) -> bool:
    """True when checked_lines accepts the file."""
    return checked_lines(path, projects) is not None


if __name__ == "__main__":
    # python3 -m agent.events --check FILE PROJECTS_DIR [DEST]: exit 0 if the file is fine; with DEST, the checked
    # lines (re-written from their values) are written there (mode 644, replaced in one step)
    import sys
    if len(sys.argv) not in (4, 5) or sys.argv[1] != "--check":
        sys.exit(2)
    pdir = sys.argv[3]
    names = [n for n in os.listdir(pdir) if os.path.isdir(os.path.join(pdir, n))] if os.path.isdir(pdir) else []
    got = checked_lines(sys.argv[2], names)
    if got is None:
        sys.exit(1)
    if len(sys.argv) == 5:
        dest = sys.argv[4]
        tmp = dest + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write("".join(x + "\n" for x in got))
        os.chmod(tmp, 0o644)
        os.replace(tmp, dest)
    sys.exit(0)
