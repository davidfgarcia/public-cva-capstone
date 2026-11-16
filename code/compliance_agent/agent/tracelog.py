"""Execution traces as CSV files, for troubleshooting.

Every run writes, using the run's UTC timestamp (yyyy-MM-dd-HH-mm-ss):
  <project>/Files/Logs/<ts>/payroll__log.txt     payroll review of that project
  <project>/Files/Logs/<ts>/wages__log.txt       wage-data merge of that project
  <project>/Files/Logs/<ts>/intent__log.txt      intent/affidavit validation of that project
  run_logs/<ts>__run_log.txt             everything in the run, all projects

Columns: timestamp_utc (yyyy-MM-dd-HH-mm-ss), level, run, project, task, step, source, message, details
  step    = the function that logged the line
  source  = module:line in the agent code
  details = extra data (file paths, counts, field maps) or the full error traceback
"""
import contextlib
import contextvars
import csv
import datetime as dt
import logging
import traceback
from pathlib import Path
from .util import TS_FORMAT

TASK_DIRS = {"payroll": "payroll", "wages": "wages", "intent": "intent"}   # tasks that get a log file of their own
COLUMNS = ["timestamp_utc", "level", "run", "project", "task", "step", "source", "message", "details"]

_project = contextvars.ContextVar("project", default="")
_task = contextvars.ContextVar("task", default="")
NOISY = ("pdfminer", "pdfplumber", "PIL", "charset_normalizer", "botocore", "boto3", "s3transfer")


@contextlib.contextmanager
def context(project="", task=""):
    tp = _project.set(project)
    tt = _task.set(task)
    try:
        yield
    finally:
        _project.reset(tp)
        _task.reset(tt)


class _Files:
    def __init__(self):
        self.open = {}

    def writer(self, path: Path):
        if path not in self.open:
            path.parent.mkdir(parents=True, exist_ok=True)
            new = not path.exists()
            f = open(path, "a", newline="", encoding="utf-8")
            w = csv.writer(f)
            if new:
                w.writerow(COLUMNS)
            self.open[path] = (f, w)
        return self.open[path][1]

    def flush(self):
        for f, _ in self.open.values():
            f.flush()

    def close(self):
        for f, _ in self.open.values():
            f.close()
        self.open.clear()


def _traceback(ei) -> str:
    """The traceback for the run log. On AWS (CVA_COMPONENT) the run log is copied to Drive and S3, so only where it
    happened and the error's type are kept, never its message, which can name a file (#56)."""
    import os
    if not os.environ.get("CVA_COMPONENT"):
        return "".join(traceback.format_exception(*ei)).strip()
    out, e = [], ei[1]
    while e is not None and len(out) < 10:
        out.append("".join(traceback.format_tb(e.__traceback__)).rstrip() + f"\n{type(e).__name__} (message left out)")
        e = e.__cause__ or e.__context__
    return "\n-- caused by / during:\n".join(reversed(out)).strip()


class TraceRouter(logging.Handler):
    """Writes each log record to the run log and to the log of the project/task it belongs to."""

    def __init__(self, run_ts, run_log: Path, project_dirs: dict, tz="UTC"):
        super().__init__(level=logging.DEBUG)
        self.run_ts, self.run_log, self.dirs = run_ts, run_log, project_dirs
        self.files = _Files()

    def emit(self, record):
        try:
            project, task = _project.get(), _task.get()
            details = getattr(record, "details", "")
            if isinstance(details, dict):
                details = "; ".join(f"{k}={v}" for k, v in details.items())
            if record.exc_info:
                details = (str(details) + "\n" if details else "") + _traceback(record.exc_info)
            row = [dt.datetime.fromtimestamp(record.created, dt.timezone.utc).strftime(TS_FORMAT),
                   record.levelname, self.run_ts, project, task, record.funcName,
                   f"{record.module}:{record.lineno}" if record.name.startswith("compliance") else f"{record.name}:{record.lineno}",
                   record.getMessage(), details]
            self.files.writer(self.run_log).writerow(row)
            if project in self.dirs and task in TASK_DIRS:
                from .layout import Paths
                p = Paths(self.dirs[project]).task_log(task, self.run_ts)
                self.files.writer(p).writerow(row)
            if record.levelno >= logging.WARNING:
                self.files.flush()
        except Exception:  # logging must never break a run
            self.handleError(record)

    def close(self):
        self.files.close()
        super().close()


@contextlib.contextmanager
def tracing(run_ts, run_log: Path, project_dirs: dict, tz="UTC"):
    """Install the CSV trace handler for the duration of a run. Console handlers keep their level."""
    root = logging.getLogger()
    old_level = root.level
    for h in root.handlers:                    # keep existing console output at its previous verbosity
        if h.level == logging.NOTSET:
            h.setLevel(old_level or logging.WARNING)
            h._trace_reset = True
    router = TraceRouter(run_ts, run_log, project_dirs, tz)
    root.addHandler(router)
    root.setLevel(logging.DEBUG)
    logging.getLogger("compliance").setLevel(logging.DEBUG)
    quiet = {n: logging.getLogger(n).level for n in NOISY}
    for n in NOISY:
        # the PDF readers warn about drawing details of a page (a colour they can't set...), never about its text:
        # not worth a warning in the outcome email
        logging.getLogger(n).setLevel(logging.ERROR if n in ("pdfminer", "pdfplumber") else logging.WARNING)
    try:
        yield router
    finally:
        root.removeHandler(router)
        router.close()
        root.setLevel(old_level)
        for n, lv in quiet.items():
            logging.getLogger(n).setLevel(lv)
        for h in root.handlers:
            if getattr(h, "_trace_reset", False):
                h.setLevel(logging.NOTSET)
                h._trace_reset = False
