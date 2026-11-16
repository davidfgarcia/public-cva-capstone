"""Error reporting to Sentry, privacy first. Shared by the agent on AWS, the hosted web app and the tokenizer
(code/tokenizer/pii_tokenizer/errors.py is an identical copy; tests/test_errors.py checks they stay the same).

Client data must never reach Sentry, so a report carries only:
  * what failed: the exception type (KeyError, TimeoutError, ...), or "error logged by <logger>" for a logged error
  * where: the stack trace (file, function and line, and the program's own source code around each line)
  * which part and version: component (aws-agent, web, tokenizer), environment, release, server name, Python version
  * which run: the run ID (12 random hex characters, docs/run_log.md D9), when the pipeline or the run set one
Never sent: exception messages, log messages and their arguments, local variables, command-line arguments, request
data, breadcrumbs, file paths outside the program and its libraries. Two layers make sure of it:
  * each error event is rebuilt from an allow-list of fields (strip below), each value checked, so a field a future
    Sentry version adds is dropped too;
  * the connection itself (_events_only) sends error events and nothing else: every other kind of item a library or
    a future SDK might add (attachments, sessions, metrics, logs, transactions, check-ins, profiles, client reports)
    is dropped there, and the envelope's own headers are rebuilt too.
No tracing headers are added to the program's own requests to Google, AWS, L&I, SAM.gov or Anthropic (child processes
do get a random trace ID in their environment, SUBPROCESS_SENTRY_TRACE: it identifies nothing and stays on the
machine). The one message the program writes itself: report_step_failure("pull") sends "Pipeline step failed: pull".

Reporting is on only when both hold:
  1. the program calls setup(component). The AWS pipeline, the web app (CVA_HOSTED=1) and the tokenizer's Cloud Run
     job do; the desktop app never does (and setup() refuses to start in a packaged app).
  2. a DSN is configured: the SENTRY_DSN environment variable, else error_reporting.sentry_dsn in the config file.
     SENTRY_DSN=off (or empty) turns reporting off.
Within one process, the same error is reported at most once every SENTRY_THROTTLE_MINUTES (default 10): this matters
in the web app, which runs for hours; each run on AWS is a process of its own. Nothing here may stop a run.
"""
from __future__ import annotations

import logging
import os
import re
import site
import sys
import sysconfig
import threading
import time

log = logging.getLogger("compliance.errors")

OFF = {"", "off", "none", "false", "0", "no"}
# The pipeline's steps (deploy/run_pipeline.sh) that can fail, and the Lambda time limit. A closed list: the step
# message is the one text the program writes itself, so nothing else can pass as one. Not here: "lock" (another run
# in progress) and "projects" (a project name that matches nothing), which aren't bugs.
STEPS = {"config", "tokenize", "pull", "agent", "start", "push", "mirror", "restore", "timeout"}
STEP_MESSAGES = {f"Pipeline step failed: {s}" for s in STEPS}
SAFE_TAG = "cva_safe_message"
KEEP_FRAME = {"filename", "abs_path", "function", "module", "lineno", "colno", "in_app",
              "context_line", "pre_context", "post_context"}
LEVELS = {"fatal", "error", "warning", "info", "debug"}
NAME_RE = re.compile(r"[A-Za-z0-9_.+-]{1,60}")             # SDK and runtime names and versions
HOST_RE = re.compile(r"[A-Za-z0-9.-]{1,253}")
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]{0,99}")
MODULE_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]{0,199}")
COMPONENT_RE = re.compile(r"[a-z][a-z0-9-]{0,39}")      # aws-agent, web, tokenizer
RUN_ID_RE = re.compile(r"[0-9a-f]{12}")                 # the run ID (docs/run_log.md D1, D9): random, names nothing
SOURCE_RE = re.compile(r"__cause__|__context__|exceptions\[\d{1,4}\]")
# How an error was caught, as the SDK's integrations in use here write it
MECHANISMS = {"generic", "chained", "logging", "excepthook", "threading", "atexit", "unraisablehook", "aws_lambda"}
ENV_RE = re.compile(r"[a-z][a-z0-9_-]{0,39}")            # production, staging, ...
RELEASE_RE = re.compile(r"compliance-agent@[A-Za-z0-9._-]{1,40}")   # compliance-agent@<commit>
# Logger names are kept only down to their first part, and only for the program's loggers and loaded libraries
PROGRAM_LOGGERS = {"compliance", "pii_tokenizer", "root"}
APP_ROOT = os.path.realpath(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))) + os.sep
_BASE_TRANSPORT = None                    # tests put a capturing transport here; None = Sentry's HTTP transport
# Frames from these folders (the program, Python and installed libraries) keep their path and source lines; any
# other file (generated or exec'd code, whose "file name" can be anything) keeps only function and line number.
def _site_dirs():
    try:
        return [*site.getsitepackages(), site.getusersitepackages()]
    except Exception:                       # some embedded Pythons have no site helpers
        return []


CODE_ROOTS = tuple(sorted({os.path.realpath(p) + os.sep for p in (
    APP_ROOT,
    sys.prefix, sys.base_prefix, sysconfig.get_paths().get("stdlib") or "",
    sysconfig.get_paths().get("purelib") or "", sysconfig.get_paths().get("platlib") or "", *_site_dirs()) if p}))

_state = {"on": False}
_cleaned: dict[str, None] = {}            # event IDs before_send produced (the last 1,000): sent as they are
_last_sent: dict[str, float] = {}
_lock = threading.Lock()


def _throttle_minutes() -> float:
    try:
        return float(os.environ.get("SENTRY_THROTTLE_MINUTES", "10"))
    except ValueError:
        return 10.0


# ------------------------------------------------------------------------------------- what may leave
def _from_code(f, roots=None) -> bool:
    """A frame from a real Python file of the program, Python or an installed library (not exec'd or generated
    code, whose "file name" can be anything)."""
    path = f.get("abs_path") or ""
    if not (path and os.path.isabs(path) and path.endswith(".py")):
        return False
    real = os.path.realpath(path)
    return real.startswith(roots or CODE_ROOTS) and os.path.isfile(real)


def _frames(stacktrace) -> dict:
    frames = []
    for f in (stacktrace or {}).get("frames") or []:
        if not isinstance(f, dict):
            continue
        if _from_code(f):
            frames.append({k: f[k] for k in KEEP_FRAME if k in f})
        else:
            frames.append({k: f[k] for k in ("lineno", "in_app") if k in f} | {"filename": "<other>"})
    return {"frames": frames}


def _logger(name) -> str:
    """The first part of a logger's name, for the program's loggers and loaded libraries; else "other"."""
    top = str(name or "").split(".")[0]
    if IDENT_RE.fullmatch(top) and (top in PROGRAM_LOGGERS or top in sys.modules):
        return top
    return "other"


def _exception_type(v) -> tuple[str, str]:
    """(type, module) of an exception, kept only for a class that really lives in a loaded module under that name
    (a class made at run time could be named after data)."""
    t, m = str(v.get("type") or ""), str(v.get("module") or "builtins")
    if m == "botocore.errorfactory":       # AWS service errors (NoSuchKey, ...) are made at run time from ClientError
        return "ClientError", "botocore.exceptions"
    mod = sys.modules.get(m) if MODULE_RE.fullmatch(m) else None
    if IDENT_RE.fullmatch(t) and mod is not None and isinstance(getattr(mod, t, None), type):
        return t, ("" if m == "builtins" else m)
    return "Error", ""


def _name(value, pattern=NAME_RE, default=""):
    value = str(value or "")
    return value if pattern.fullmatch(value) else default


def _top_frame(event) -> dict:
    """The innermost frame of the program's own code (else the innermost frame): from the exception's stack trace,
    or, for an error logged without one, from the logging thread's. "The program's own code" is decided by where
    the file is, not by Sentry's in_app flag, which depends on the folder a run starts in."""
    stacks = [v.get("stacktrace") for v in (event.get("exception") or {}).get("values") or []]
    if not any(stacks):
        stacks = [v.get("stacktrace") for v in (event.get("threads") or {}).get("values") or []]
    frames = [f for s in stacks if s for f in s.get("frames") or [] if isinstance(f, dict)]
    own = [f for f in frames if _from_code(f, (APP_ROOT,))]
    # else the innermost frame that isn't Python's logging or Sentry itself (the caller of log.error)
    callers = [f for f in frames if str(f.get("module") or "").split(".")[0] not in ("logging", "sentry_sdk")]
    return (own or callers or frames or [{}])[-1]


def strip(event: dict) -> dict:
    """The event Sentry receives: only the fields listed here, never any text that came from data."""
    out = {"platform": "python"}
    if re.fullmatch(r"[0-9a-f]{32}", str(event.get("event_id") or "")):
        out["event_id"] = event["event_id"]
    if isinstance(event.get("timestamp"), (str, int, float)) and re.fullmatch(r"[0-9T:.+Z-]{1,40}", str(event["timestamp"])):
        out["timestamp"] = event["timestamp"]
    out["level"] = event.get("level") if event.get("level") in LEVELS else "error"
    if _name(event.get("release"), RELEASE_RE):
        out["release"] = event["release"]
    if _name(event.get("environment"), ENV_RE):
        out["environment"] = event["environment"]
    if _name(event.get("server_name"), HOST_RE):
        out["server_name"] = event["server_name"]
    sdk = event.get("sdk") or {}
    if isinstance(sdk, dict) and _name(sdk.get("name")) and _name(sdk.get("version")):
        out["sdk"] = {"name": sdk["name"], "version": sdk["version"]}
    if event.get("logger"):
        out["logger"] = _logger(event["logger"])
    excs = []
    for v in (event.get("exception") or {}).get("values") or []:
        if not isinstance(v, dict):
            continue
        t, m = _exception_type(v)
        e = {"type": t, "value": ""}
        if m:
            e["module"] = m
        mech = v.get("mechanism") or {}
        if isinstance(mech, dict) and mech:
            m2 = {k: mech[k] for k in ("handled", "is_exception_group") if isinstance(mech.get(k), bool)}
            if mech.get("type") in MECHANISMS:
                m2["type"] = mech["type"]
            # how chained and grouped exceptions fit together (small numbers, and where each came from)
            m2.update({k: mech[k] for k in ("exception_id", "parent_id")
                       if type(mech.get(k)) is int and 0 <= mech[k] <= 9999})
            if _name(mech.get("source"), SOURCE_RE):
                m2["source"] = mech["source"]
            e["mechanism"] = m2
        if v.get("stacktrace"):
            e["stacktrace"] = _frames(v["stacktrace"])
        excs.append(e)
    if excs:
        out["exception"] = {"values": excs}
    threads = []
    for t in (event.get("threads") or {}).get("values") or []:
        if t.get("stacktrace"):
            threads.append({"id": t.get("id") if isinstance(t.get("id"), int) else 0,
                            "current": bool(t.get("current")), "crashed": bool(t.get("crashed")),
                            "stacktrace": _frames(t["stacktrace"])})
    if threads:
        out["threads"] = {"values": threads}
    tags = event.get("tags") or {}
    msg = (event.get("logentry") or {}).get("message") or event.get("message") or ""
    if isinstance(msg, dict):
        msg = msg.get("formatted") or msg.get("message") or ""
    if tags.get(SAFE_TAG) == "1" and str(msg) in STEP_MESSAGES:
        out["message"] = str(msg)                        # written by report_step_failure, never data
    elif excs:
        out["message"] = excs[-1]["type"]
    else:
        f = _top_frame(event)
        where = f"{f.get('module') or f.get('filename') or '?'}.{f.get('function') or '?'}" if f else ""
        if not _from_code(f):
            where = ""
        name = _logger(event["logger"]) if event.get("logger") else "the program"
        out["message"] = f"Error logged by {name}" + (f" in {where}" if where else "")
    out["tags"] = {"component": tags["component"]} if _name(tags.get("component"), COMPONENT_RE) else {}
    if _name(tags.get("run_id"), RUN_ID_RE):             # which run it came from (D9): 12 random hex characters
        out["tags"]["run_id"] = tags["run_id"]
    ctx = event.get("contexts") or {}
    out["contexts"] = {}
    for k in ("runtime",):
        c = ctx.get(k) if isinstance(ctx, dict) else None
        if isinstance(c, dict):
            kept = {f: c[f] for f in ("name", "version") if _name(c.get(f))}
            if kept:
                out["contexts"][k] = kept
    return out


# ------------------------------------------------------------------------------------- how often
def _key(event: dict) -> str:
    f = _top_frame(event)
    types = ",".join(str(v.get("type")) for v in (event.get("exception") or {}).get("values") or [])
    msg = event.get("message") if (event.get("tags") or {}).get(SAFE_TAG) == "1" else ""
    return f"{event.get('logger')}|{types}|{f.get('module')}.{f.get('function')}:{f.get('lineno')}|{msg}"


def _throttled(event: dict, now: float | None = None) -> bool:
    minutes = _throttle_minutes()
    if minutes <= 0:
        return False
    key, now = _key(event), time.time() if now is None else now
    with _lock:
        last = _last_sent.get(key)
        if last is not None and now - last < minutes * 60:
            return True
        _last_sent[key] = now
        return False


def before_send(event, hint):
    try:
        if event.get("type") not in (None, "error", "default"):   # e.g. a check-in: not an error report
            return None
        if isinstance(hint, dict):
            hint.pop("attachments", None)
        if _throttled(event):
            return None
        out = strip(event)
        if out.get("event_id"):
            with _lock:
                _cleaned[out["event_id"]] = None
                while len(_cleaned) > 1000:
                    _cleaned.pop(next(iter(_cleaned)))
        return out
    except Exception:                       # never send an event this code could not clean
        return None


def _refuse(*args, **kwargs):
    return None


def _events_only(envelope):
    """The envelope actually sent: its error events and nothing else, with headers rebuilt from scratch. An event
    that didn't come through before_send (something handed straight to the connection) is cleaned here."""
    from sentry_sdk.envelope import Envelope, Item, PayloadRef
    items = []
    for item in envelope.items or []:
        event = item.get_event() if item.type == "event" else None
        if not isinstance(event, dict):
            continue
        with _lock:
            ours = _cleaned.pop(str(event.get("event_id")), 0) is None
        if not ours:
            if event.get("type") not in (None, "error", "default"):
                continue
            event = strip({k: v for k, v in event.items() if k != "tags"})   # no step message can be forged here
        items.append(Item(payload=PayloadRef(json=event), type="event"))
    if not items:
        return None
    headers = {k: envelope.headers[k] for k in ("event_id", "sent_at")
               if k in (envelope.headers or {}) and _name(envelope.headers[k], re.compile(r"[0-9a-fA-FTZ:.+-]{1,40}"))}
    return Envelope(headers=headers, items=items)


def _transport_class():
    from sentry_sdk.transport import HttpTransport
    base = _BASE_TRANSPORT or HttpTransport

    class EventsOnly(base):
        # Sentry's HTTP transport adds a "client report" (counts of dropped items) to envelopes after
        # capture_envelope, and its records come from here: neither exists for this connection
        def _fetch_pending_client_report(self, *args, **kwargs):
            return None

        def record_lost_event(self, *args, **kwargs):
            return None

        def capture_envelope(self, envelope):
            try:
                kept = _events_only(envelope)
            except Exception:               # never send what could not be checked
                kept = None
            if kept is not None:
                super().capture_envelope(kept)

    return EventsOnly


# ------------------------------------------------------------------------------------- set-up
def _section(cfg) -> dict:
    sec = cfg.get("error_reporting") if isinstance(cfg, dict) else None
    return sec if isinstance(sec, dict) else {}


def dsn_from_config(cfg) -> str:
    return str(_section(cfg).get("sentry_dsn") or "").strip()


def environment_from_config(cfg) -> str:
    return str(_section(cfg).get("environment") or "").strip()


def release_name(code_dir=None) -> str:
    """CVA_RELEASE (set by the deploy scripts), else the commit in DEPLOYED_COMMIT next to the code (EC2)."""
    rel = os.environ.get("CVA_RELEASE", "").strip()
    if not rel and code_dir:
        try:
            with open(os.path.join(str(code_dir), "DEPLOYED_COMMIT")) as f:
                rel = f.read().strip()
        except OSError:
            pass
    return f"compliance-agent@{rel}" if rel else "compliance-agent@unknown"   # never let the SDK guess it


def setup(component: str, dsn: str = "", environment: str = "", release: str = "") -> bool:
    """Start reporting for this process. Returns True when on. Never raises."""
    if _state["on"]:
        return True
    if getattr(sys, "frozen", False):        # a packaged app (the desktop app): never, whatever the settings say
        return False
    env_dsn = os.environ.get("SENTRY_DSN")
    dsn = (env_dsn if env_dsn is not None else dsn or "").strip()
    if dsn.lower() in OFF:
        return False
    try:
        import sentry_sdk
        from sentry_sdk.integrations.logging import LoggingIntegration
    except ImportError:
        log.warning("error reporting: a Sentry DSN is set but the sentry-sdk package is not installed")
        return False
    try:
        import socket
        sentry_sdk.init(
            dsn=dsn,
            environment=os.environ.get("SENTRY_ENVIRONMENT") or environment or "production",
            release=release or "compliance-agent@unknown",
            server_name=socket.gethostname(),
            transport=_transport_class(),         # sends error events only (see _events_only)
            integrations=[LoggingIntegration(level=None, event_level=logging.ERROR)],
            auto_enabling_integrations=False,    # no anthropic/boto3/... hooks: 1.8 s and 84 MB less at start-up
            project_root=APP_ROOT.rstrip(os.sep),  # the program's folder, whatever folder a run starts in
            include_local_variables=False,
            attach_stacktrace=True,
            max_breadcrumbs=0,                   # no breadcrumbs at all
            max_request_body_size="never",
            send_client_reports=False,
            before_send=before_send,
            before_breadcrumb=_refuse,
            # every other kind of item is refused here too (and dropped by the transport anyway)
            before_send_metric=_refuse,          # (metrics are always collected by the SDK; this drops them all)
            enable_logs=False,
            before_send_log=_refuse,
            traces_sample_rate=None,
            before_send_transaction=_refuse,
            auto_session_tracking=False,
            # no sentry-trace / baggage headers on the program's own requests to other services
            trace_propagation_targets=[],
            spotlight=False,                     # no copy to a local debugging sidecar either
        )
        sentry_sdk.set_tag("component", component)
        rid = os.environ.get("CVA_RUN_ID", "").strip().lower()   # set by run_pipeline.sh (docs/run_log.md D1, D9)
        if RUN_ID_RE.fullmatch(rid):
            sentry_sdk.set_tag("run_id", rid)
    except Exception as e:                               # a bad DSN must never stop a run
        log.warning("error reporting: Sentry could not start (%s)", type(e).__name__)
        return False
    _state["on"] = True
    log.info("error reporting: Sentry on for %s", component)
    return True


def set_run_id(run_id: str) -> bool:
    """Tags this process's later Sentry events with the run ID (docs/run_log.md D9), when reporting is on and the ID
    is valid (12 hex characters). Returns True when set. Never raises."""
    if not _state["on"] or not RUN_ID_RE.fullmatch(str(run_id or "")):
        return False
    try:
        import sentry_sdk
        sentry_sdk.set_tag("run_id", str(run_id))
        return True
    except Exception:                                    # noqa: BLE001 - error reporting must never stop a run
        return False


def clear_run_id() -> bool:
    """After a run (docs/run_log.md D9): the tag goes back to the pipeline's run ID (CVA_RUN_ID), or is removed, so a
    later error in a long-lived process doesn't name an earlier run. Never raises."""
    if not _state["on"]:
        return False
    try:
        import sentry_sdk
        rid = os.environ.get("CVA_RUN_ID", "").strip().lower()
        if RUN_ID_RE.fullmatch(rid):
            sentry_sdk.set_tag("run_id", rid)
        else:
            sentry_sdk.get_isolation_scope().remove_tag("run_id")
        return True
    except Exception:                                    # noqa: BLE001 - error reporting must never stop a run
        return False


def report_step_failure(step: str) -> bool:
    """A failed pipeline step (a shell step, so no Python error to report): sends 'Pipeline step failed: <step>'."""
    if not _state["on"] or step not in STEPS:
        return False
    try:
        import sentry_sdk
        with sentry_sdk.new_scope() as scope:
            scope.set_tag(SAFE_TAG, "1")
            sentry_sdk.capture_message(f"Pipeline step failed: {step}", level="error")
        return True
    except Exception:
        return False


def flush(timeout: float = 5.0) -> None:
    if not _state["on"]:
        return
    try:
        import sentry_sdk
        sentry_sdk.flush(timeout)
    except Exception:
        pass
