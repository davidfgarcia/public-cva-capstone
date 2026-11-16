import datetime as dt
import hashlib
import json
import os
import logging
import re
import shutil
import time
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

log = logging.getLogger("compliance")

_RUN_TS = None
TS_FORMAT = "%Y-%m-%d-%H-%M-%S"          # yyyy-MM-dd-HH-mm-ss, 24-hour, always UTC
TS_REGEX = r"\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}"
UTC = dt.timezone.utc


def run_timestamp(tz=None):
    """One timestamp per run, UTC: yyyy-MM-dd-HH-mm-ss. (`tz` is ignored; kept for old callers.)"""
    global _RUN_TS
    if _RUN_TS is None:
        _RUN_TS = now_utc().strftime(TS_FORMAT)
    return _RUN_TS


def reset_run_timestamp():
    """A new timestamp from now on (a new run; the app also resets it around uploads and removals)."""
    global _RUN_TS
    _RUN_TS = None


_RUN_NO = 0


def start_run():
    """A new run starts (runner.run): a new timestamp and a new run number."""
    global _RUN_NO
    reset_run_timestamp()
    _RUN_NO += 1


def run_number():
    """Which run this is in this process. Changes only when a run starts (start_run), not when the app resets the
    timestamp for an upload during a run, and even within one second (two runs can share a timestamp). For things
    worked out once per run."""
    return _RUN_NO


def now_utc():
    return dt.datetime.now(UTC)


def utc_stamp(when=None):
    """A moment as yyyy-MM-dd-HH-mm-ss in UTC (now if not given)."""
    when = when or now_utc()
    if when.tzinfo is None:
        when = when.replace(tzinfo=UTC)
    return when.astimezone(UTC).strftime(TS_FORMAT)


def stamped(name, ts=None):
    """yyyy-MM-dd-HH-mm-ss__<artifact_name> using the run timestamp."""
    return f"{ts or run_timestamp()}__{name}"


def file_kinds(files) -> str:
    """Files described for a log record without their names: "2 .xlsx, 1 .pdf" (#56: a file name can hold a
    worker's name, and the run log goes to AWS). "" when there are none."""
    from collections import Counter
    n = Counter(Path(f).suffix.lower() or "no extension" for f in files)
    return ", ".join(f"{c} {k}" for k, c in sorted(n.items()))


class _PathFreeFormatter(logging.Formatter):
    """Console format on AWS: a traceback is left out, since an exception's message can name a file (on AWS the
    console output goes to the pipeline log and CloudWatch, #56). The run log keeps the full traceback."""

    def format(self, record):
        saved = record.exc_text                 # logging caches the traceback text on the record: not this one's
        record.exc_text = None
        try:
            return super().format(record)
        finally:
            record.exc_text = saved

    def formatException(self, ei):
        return f"({ei[0].__name__ if ei and ei[0] else 'error'}: traceback in the run log)"

    def formatStack(self, stack_info):
        return ""


def on_aws() -> bool:
    """Running on AWS (CVA_COMPONENT is set there only): what is logged goes to CloudWatch and the AWS-side run log."""
    return bool(os.environ.get("CVA_COMPONENT"))


def err_text(e) -> str:
    """An exception for a log message: on AWS its type only (its text can name a file, #56), elsewhere its text."""
    return type(e).__name__ if on_aws() else str(e)


def utc_log_formatter(fmt="%(asctime)s %(levelname)s %(message)s", path_free=False):
    """Console log format with the same UTC timestamp pattern. path_free: no traceback (see _PathFreeFormatter)."""
    f = (_PathFreeFormatter if path_free else logging.Formatter)(fmt, datefmt=TS_FORMAT)
    f.converter = time.gmtime
    return f


# ------------------------------------------------------------------ backups
def backup_dir(folder: Path, history_parent: Path, tz=None) -> Path:
    """<history_parent>/yyyy-MM-dd-HH-mm-ss__<folder name> (created on demand)."""
    d = history_parent / stamped(folder.name)
    d.mkdir(parents=True, exist_ok=True)
    return d


def presave(files, folder: Path, history_parent: Path, move=False, tz=None):
    """Copy (or move) files out of `folder` into a timestamped backup folder before they are
    replaced/updated. Returns the backup folder, or None if there was nothing to save."""
    files = [Path(f) for f in files if Path(f).exists()]
    if not files:
        return None
    dest = backup_dir(folder, history_parent, tz)
    for f in files:
        target = dest / f.name
        if move:
            shutil.move(str(f), str(target))
        else:
            shutil.copy2(f, target)
    # no folder or file names: a folder below a project's can hold a worker's name (#56)
    log.info("%s %d file(s) to a dated backup folder", "Moved" if move else "Copied", len(files),
             extra={"details": {"files": file_kinds(files)}})
    return dest


def visible_files(folder: Path, exts=None):
    if not folder.exists():
        return []
    out = []
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.name.startswith(".") or p.name.startswith("~$"):
            continue
        if exts and p.suffix.lower() not in exts:
            continue
        out.append(p)
    return out


def file_hash(path: Path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# -------------------------------------------------------------- run state
class State:
    """Small JSON file per project remembering what has already been processed, so scheduled
    runs don't redo (and re-backup) identical work."""

    def __init__(self, project_dir: Path):
        self.path = project_dir / ".agent_state.json"
        try:
            self.data = json.loads(self.path.read_text())
        except Exception:
            self.data = {}

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value

    def save(self):
        self.path.write_text(json.dumps(self.data, indent=2, default=str))


# ------------------------------------------------------------ text helpers
_STOP = {"the", "and", "of", "inc", "llc", "co", "corp", "company", "dba", "ltd", "a"}


def norm(s):
    """Normalise a craft/classification for matching."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).lower()
    s = s.replace("&", " and ").replace("/", " ")
    s = re.sub(r"[^a-z0-9%]+", " ", s)
    words = []
    for w in s.split():
        if len(w) > 3 and w.endswith("s") and not w.endswith("ss"):
            w = w[:-1]
        words.append(w)
    return " ".join(words)


def norm_company(s):
    return " ".join(w for w in norm(s).split() if w not in _STOP)


def _name_parts(s):
    """'Brivantel Pacific dba Quorvane dba Irnovak' -> each name in the dba chain, normalised."""
    parts = re.split(r"\bdba\b|\bd b a\b", norm(s))
    return [p for p in (" ".join(w for w in x.split() if w not in _STOP) for x in parts) if p]


def _sim1(a, b):
    if a in b or b in a:
        return 1.0
    ca, cb = a.replace(" ", ""), b.replace(" ", "")      # "Q R X CONSTRUCTION" vs "QRX"
    if len(min(ca, cb, key=len)) >= 3 and (ca.startswith(cb) or cb.startswith(ca)):
        return 1.0
    ta, tb = set(a.split()), set(b.split())
    jac = len(ta & tb) / max(1, len(ta | tb))
    return max(SequenceMatcher(None, a, b).ratio(), jac)


def similarity(a, b):
    pa, pb = _name_parts(a), _name_parts(b)
    if not pa or not pb:
        return 0.0
    return max(_sim1(x, y) for x in pa for y in pb)


def norm_header(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def to_float(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("$", "").replace(",", "").strip()
    m = re.match(r"^-?\d+(\.\d+)?", s)
    return float(m.group(0)) if m else None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m-%d-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


# #53 (the tokenizer's vnext.opaque_file_names): words of a staff file or folder name that the AWS side only sees as a
# token (code/tokenizer/pii_tokenizer/names.py FILE_PREFIX)
FILE_TOKEN = re.compile(r"\bFIL-[0-9a-f]{10}\b")


def shown_name(name) -> str:
    """A file or folder name as an email shows it: words the AWS side only has as tokens shown as "…" (the email is
    written on AWS: it can't have them; staff open the file from its link, or find it by the words left)."""
    return FILE_TOKEN.sub("…", str(name))


def date_from_filename(name):
    """'..._6.26.25 LNI.csv' -> 2025-06-26 ; '8_29_2026' -> 2026-08-29."""
    m = re.search(r"(\d{1,2})[._-](\d{1,2})[._-](\d{2,4})", name)
    if not m:
        return None
    mo, d, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def truthy(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    return str(v).strip().lower() in {"true", "yes", "y", "x", "filed", "approved", "done", "complete"}


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")
