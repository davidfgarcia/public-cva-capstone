"""Compare two CVA result folders made from the same input snapshot (docs/vNext.md section 12; vNext phase 1).

    python tools/compare_runs.py <run A> <run B> [--with-logs] [--json]

Each run folder is what tools/snapshot_run.py writes: ``tokenized/`` (the tokenizer's copy of the snapshot, what AWS
would see) and ``results/`` (the agent's run on it: reviewed payrolls, trackers with their Audit and Review Log
sheets, merged wage tables, run summaries and the two outcome emails, saved as .eml previews). Any two folders with
the same layout can be compared.

What is compared:
  - spreadsheets (.xlsx, .xlsm): every sheet and cell: value, fill (colour, theme, palette, tint), note (comment) and
    format (number format, bold, italic, font colour); sheets present on one side only, and their order;
  - outcome emails (.eml): subject, sender and recipients, the text and the HTML part line by line, and each
    attachment by name and content (a run log attachment only with --with-logs);
  - JSON (run summaries, the agent's state): key by key, types included (``true`` is not ``1``);
  - CSV and text: line by line (run logs only with --with-logs: they hold timings);
  - anything else: byte for byte;
  - files present on one side only; a symbolic link is reported, never followed.

**Run times are ignored, nothing else.** Each run's time window (ten minutes either side) comes from its
``run_window.json`` (snapshot_run.py), else its run summaries' ``run_utc`` and ``finished_utc``, else its newest
file; a window longer than six hours is warned about on stderr. Inside a value or a name, only these are masked: a ``yyyy-MM-dd-HH-mm-ss`` stamp (or ISO
date and time) inside a window; a date (``yyyy-MM-dd``, or a date cell) on a day a window touches; the outcome emails'
run time ("Mon Nov 9, 6:21 AM PST" on such a day) and the "took 22 s" right after it; the JSON keys ``duration_s`` and ``durations``; the run folders' own paths; an email's random
Content-IDs. Clock times, weekdays, durations and dates outside the windows are compared. (Blind spot: an *input*
date on the day the runs were made is masked too; a snapshot's input dates are older.)

**It never prints a value.** A difference is reported as a path, a sheet and a cell reference (or a line number, a
JSON key path, a header name) and the kind of difference ("value", "fill", "missing on B"...). A JSON key is printed
only in the form code gives its keys (lower case letters, digits, ``_``); any other key (it could be a name, an ID
or a token) is printed as ``{key#<8 hex>}``, keyed with a random salt made for each comparison: it tells keys apart
within one output and can't be looked up. A lower-case one-word key is printed as it is. A file that can't be read is reported by
its error type alone. Paths and sheet names are printed as they are, so run it on invented fixtures or on tokenized
copies only (HANDOFF section 6).

Exit code: 0 when the runs are the same, 1 when they differ, 2 on bad arguments, 3 when the comparison itself
failed (only the error type is printed).
"""
from __future__ import annotations

import argparse
import datetime as dt
import email
import hashlib
import json
import hmac
import os
import re
import secrets
import sys
from email import policy
from pathlib import Path
from urllib.parse import quote

SHEET_EXTS = {".xlsx", ".xlsm"}
TEXT_EXTS = {".csv", ".txt", ".md", ".log"}
MARGIN = dt.timedelta(minutes=10)
# JSON keys whose values are run durations (seconds): skipped, with everything under them
TIME_KEYS = {"duration_s", "durations"}
# The structured run log's fields that differ on every run (agent/events.py; docs/run_log.md): not compared
EVENT_VOLATILE = {"ts", "run_id", "duration_ms", "tokenizer_before"}
# A JSON key is printed only in the form code gives its keys (lower case, digits, "_"); a name or a token never is
_PLAIN_KEY = re.compile(r"[a-z][a-z0-9_]{0,63}\Z")
MAX_WINDOW = dt.timedelta(hours=6)

# A run stamp, as the agent and the tokenizer write them: 2026-11-09-14-21-48 (config.yaml: UTC), or ISO
# (UTC unless an offset or a Pacific zone follows: wage tables are named in Pacific time, "Federal <stamp> PST.xlsx")
_STAMP = re.compile(r"(\d{4})-(\d{2})-(\d{2})[-T _](\d{2})[-:](\d{2})[-:](\d{2})(?:\.\d+)?"
                    r"(Z|([+-])(\d{2}):?(\d{2})| ?(PST|PDT)\b)?")
_PACIFIC = {"PST": 8, "PDT": 7}                     # hours to add to reach UTC
_DATE = re.compile(r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?![\d-])")
# The outcome emails' run time, as written (notify.py, staff_email.py): "Mon Nov 9, 6:21 AM PST", in the staff
# email after "Checked:", in the administrator email followed by "(<stamp> UTC), took 22 s" (notify._dur)
_LOCAL = re.compile(r"\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun) ([A-Z][a-z]{2}) (\d{1,2}), \d{1,2}:\d{2} [AP]M [A-Z]{2,4}\b")
_TOOK = re.compile(r"(<T>(?: \(<T> UTC\))?), took \d+(?: min \d{2})? s\b")
_MONTHS = {m: i + 1 for i, m in enumerate("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split())}
_CID = re.compile(r"(cid:)[^\s\"'<>)]+")


class _Ctx:
    """The two runs' time windows (UTC) and folder paths, set once per comparison."""
    windows: list = []
    days: set = set()
    salt: bytes = b""
    warnings: list = []
    roots: re.Pattern | None = None


def _parse_stamp(text):
    m = _STAMP.search(str(text or ""))
    if not m:
        return None
    try:
        return dt.datetime(*(int(m.group(i)) for i in range(1, 7)))
    except ValueError:
        return None


def _window(root: Path):
    """The run's time window (UTC): from ``run_window.json`` (snapshot_run.py writes it), else from the agent's run
    summaries (``run_utc`` .. ``finished_utc``), else the newest file's time alone. Never from the oldest file: files
    copied with their original times (the agent does that) would stretch it over months."""
    times = []
    rw = root / "run_window.json"
    if rw.is_file() and not rw.is_symlink():
        try:
            w = json.loads(rw.read_text(encoding="utf-8"))
            times = [t for t in (_parse_stamp(w.get("start_utc")), _parse_stamp(w.get("end_utc"))) if t]
        except (ValueError, AttributeError, OSError):
            times = []
    if not times:
        for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
            for n in filenames:
                f = Path(dirpath) / n
                if n.endswith("run_summary.json") and not f.is_symlink():
                    try:
                        summ = json.loads(f.read_text(encoding="utf-8"))
                    except (ValueError, OSError):
                        continue
                    if isinstance(summ, dict):
                        times += [t for t in (_parse_stamp(summ.get("run_utc")), _parse_stamp(summ.get("finished_utc")))
                                  if t]
    if not times:
        newest = 0.0
        for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
            for n in filenames:
                try:
                    newest = max(newest, os.lstat(os.path.join(dirpath, n)).st_mtime)
                except OSError:
                    pass
        if not newest:
            return None
        times = [dt.datetime.fromtimestamp(newest, dt.timezone.utc).replace(tzinfo=None)]
    return min(times) - MARGIN, max(times) + MARGIN


def _setup(a: Path, b: Path) -> None:
    _Ctx.windows = [w for w in (_window(a), _window(b)) if w]
    _Ctx.salt = secrets.token_bytes(16)             # key labels match within one comparison, and can't be looked up
    _Ctx.warnings = [f"run {side}: its run times span more than {MAX_WINDOW.seconds // 3600} hours; dates in that span "
                     "are not compared" for side, w in zip("AB", (_window(a), _window(b)))
                     if w and w[1] - w[0] > MAX_WINDOW + 2 * MARGIN]
    days = set()
    for lo, hi in _Ctx.windows:
        d = lo.date() - dt.timedelta(days=1)         # a local date can be a day off UTC either way
        while d <= hi.date() + dt.timedelta(days=1):
            days.add(d)
            d += dt.timedelta(days=1)
    _Ctx.days = days
    forms = set()
    for r in (a, b):
        p = r.resolve()
        forms |= {str(p), quote(str(p)), p.as_uri()}
    # A run folder's own path is masked only as a whole path (not followed by more of a name)
    alts = "|".join(re.escape(f) for f in sorted(forms, key=len, reverse=True))
    _Ctx.roots = re.compile(f"(?<![\\w./\\\\-])(?:{alts})(?![^/\\\\\\s\"'<>|,;)\\]])")


def _in_window(t: dt.datetime) -> bool:
    return any(lo <= t <= hi for lo, hi in _Ctx.windows)


def _stamp(m) -> str:
    try:
        t = dt.datetime(*(int(m.group(i)) for i in range(1, 7)))
    except ValueError:
        return m.group(0)
    if m.group(11):
        t += dt.timedelta(hours=_PACIFIC[m.group(11)])
    elif m.group(8):
        off = dt.timedelta(hours=int(m.group(9)), minutes=int(m.group(10)))
        t += -off if m.group(8) == "+" else off
    return "<T>" if _in_window(t) else m.group(0)


def _date(m) -> str:
    try:
        d = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return m.group(0)
    return "<D>" if d in _Ctx.days else m.group(0)


def _local(m) -> str:
    month, day = _MONTHS.get(m.group(1)), int(m.group(2))
    if month and any(d.month == month and d.day == day for d in _Ctx.days):
        return "<T>"
    return m.group(0)


def mask(text: str) -> str:
    """``text`` with the run folders' own paths and the run times (see the module docstring) replaced."""
    if _Ctx.roots is not None:
        text = _Ctx.roots.sub("<RUN>", text)
    text = _STAMP.sub(_stamp, text)
    text = _DATE.sub(_date, text)
    text = _LOCAL.sub(_local, text)
    return _TOOK.sub(r"\1, took <DUR>", text)


def key_label(k) -> str:
    """A JSON key as printed: itself when a plain identifier, else a short hash (a key can be a name)."""
    s = str(k)
    if _PLAIN_KEY.match(s):
        return s
    digest = hmac.new(_Ctx.salt, mask(s).encode("utf-8", "surrogatepass"), hashlib.sha256).hexdigest()
    return "{key#" + digest[:8] + "}"


def _files(root: Path, with_logs: bool) -> tuple:
    """({masked relative path: file}, [masked relative paths of symbolic links])."""
    out, links = {}, []
    for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
        d = Path(dirpath)
        for n in sorted(dirnames):
            if (d / n).is_symlink():
                links.append(mask((d / n).relative_to(root).as_posix()))
        dirnames.sort()
        for n in sorted(filenames):
            f = d / n
            rel = f.relative_to(root).as_posix()
            if f.is_symlink():
                links.append(mask(rel))
                continue
            if n.startswith(".~") or n.startswith("~$"):
                continue
            if not with_logs and f.suffix.lower() in (".txt", ".log") and "log" in n.lower():
                continue
            key = mask(rel)
            if key in out:                   # two files differing only by a run time: keep both apart
                k = 2
                while f"{key}#{k}" in out:
                    k += 1
                key = f"{key}#{k}"
            out[key] = f
    return out, links


# ------------------------------------------------------------------ spreadsheets
def _cell_value(v):
    if isinstance(v, dt.datetime):
        return "<T>" if _in_window(v.replace(tzinfo=None)) or (v.time() == dt.time() and v.date() in _Ctx.days) else v
    if isinstance(v, dt.date):
        return "<D>" if v in _Ctx.days else v
    if isinstance(v, str):
        return mask(v)
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _color(c):
    if c is None:
        return None
    return (getattr(c, "type", None), getattr(c, "rgb", None) if getattr(c, "type", None) == "rgb" else None,
            getattr(c, "theme", None) if getattr(c, "type", None) == "theme" else None,
            getattr(c, "indexed", None) if getattr(c, "type", None) == "indexed" else None, getattr(c, "tint", None))


def _fill(c):
    f = c.fill
    if f is None or f.fill_type is None:
        return None
    return (f.fill_type, _color(f.fgColor), _color(f.bgColor))


def _format(c):
    ft = c.font
    return (c.number_format, bool(ft.b) if ft else None, bool(ft.i) if ft else None, _color(ft.color) if ft else None)


def _note(c):
    return mask(c.comment.text) if c.comment is not None and c.comment.text else None


def _same_type_eq(x, y) -> bool:
    if isinstance(x, bool) or isinstance(y, bool):
        return type(x) is type(y) and x == y
    return x == y


def compare_workbooks(a: Path, b: Path, rel: str) -> list:
    import openpyxl
    out = []
    wa, wb = openpyxl.load_workbook(a), openpyxl.load_workbook(b)
    ta, tb = [mask(t) for t in wa.sheetnames], [mask(t) for t in wb.sheetnames]
    for t in ta:
        if t not in tb:
            out.append(f"{rel} [{t}]: sheet missing on B")
    for t in tb:
        if t not in ta:
            out.append(f"{rel} [{t}]: sheet missing on A")
    if [t for t in ta if t in tb] != [t for t in tb if t in ta]:
        out.append(f"{rel}: sheets in a different order")
    for ws_a, t in zip(wa.worksheets, ta):
        if t not in tb:
            continue
        ws_b = wb.worksheets[tb.index(t)]
        ca = {c.coordinate: c for row in ws_a.iter_rows() for c in row}
        cb = {c.coordinate: c for row in ws_b.iter_rows() for c in row}
        for ref in sorted(set(ca) | set(cb), key=_ref_key):
            x, y = ca.get(ref), cb.get(ref)
            kinds = []
            if not _same_type_eq(_cell_value(x.value) if x is not None else None,
                                 _cell_value(y.value) if y is not None else None):
                kinds.append("value")
            for kind, fn in (("fill", _fill), ("note", _note), ("format", _format)):
                if (fn(x) if x is not None else None) != (fn(y) if y is not None else None):
                    kinds.append(kind)
            if kinds:
                out.append(f"{rel} [{t}]!{ref}: {', '.join(kinds)}")
    return out


def _ref_key(ref: str):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


# ------------------------------------------------------------------ emails, JSON, text
def _email_parts(path: Path, with_logs: bool) -> dict:
    msg = email.message_from_bytes(path.read_bytes(), policy=policy.default)
    parts = {h: mask(str(msg.get(h, ""))) for h in ("Subject", "From", "To", "Cc", "Reply-To")}
    for kind in ("plain", "html"):
        p = msg.get_body(preferencelist=(kind,))
        body = p.get_content() if p is not None and p.get_content_type() == f"text/{kind}" else ""
        parts[kind] = [_CID.sub(r"\1<ID>", mask(line.rstrip())) for line in body.splitlines()]
    att = []
    for p in msg.iter_attachments():
        name = mask(p.get_filename() or "")
        if not with_logs and "log" in name.lower():
            continue
        data = p.get_content()
        data = mask(data).encode("utf-8", "surrogatepass") if isinstance(data, str) else bytes(data)
        att.append((name, hashlib.sha256(data).hexdigest()))
    parts["attachments"] = sorted(att)
    return parts


def compare_emails(a: Path, b: Path, rel: str, with_logs: bool = False) -> list:
    pa, pb = _email_parts(a, with_logs), _email_parts(b, with_logs)
    out = []
    for h in ("Subject", "From", "To", "Cc", "Reply-To"):
        if pa[h] != pb[h]:
            out.append(f"{rel}: header {h}")
    for kind in ("plain", "html"):
        out += [f"{rel}: {kind} {d}" for d in _line_diffs(pa[kind], pb[kind])]
    na, nb = [n for n, _ in pa["attachments"]], [n for n, _ in pb["attachments"]]
    if na != nb:
        out.append(f"{rel}: attachment names")
    elif pa["attachments"] != pb["attachments"]:
        out.append(f"{rel}: attachment contents")
    return out


def _line_diffs(la: list, lb: list) -> list:
    out = [f"line {i + 1}" for i, (x, y) in enumerate(zip(la, lb)) if x != y]
    if len(la) != len(lb):
        out.append(f"{abs(len(la) - len(lb))} more line(s) on {'A' if len(la) > len(lb) else 'B'}")
    return out


def compare_text(a: Path, b: Path, rel: str) -> list:
    la = [mask(l.rstrip()) for l in a.read_text(encoding="utf-8", errors="replace").splitlines()]
    lb = [mask(l.rstrip()) for l in b.read_text(encoding="utf-8", errors="replace").splitlines()]
    return [f"{rel}: {d}" for d in _line_diffs(la, lb)]


def _masked_keys(d: dict) -> dict:
    """{masked key: key}; keys that look the same once masked (two run stamps) are numbered in the order of the keys
    themselves, so none is dropped."""
    out = {}
    for k in sorted(d, key=str):
        m = mask(str(k))
        if m in out:
            n = 2
            while f"{m}#{n}" in out:
                n += 1
            m = f"{m}#{n}"
        out[m] = k
    return out


def _json_diffs(x, y, path="$") -> list:
    if isinstance(x, dict) and isinstance(y, dict):
        out = []
        mx, my = _masked_keys(x), _masked_keys(y)
        for k in sorted(set(mx) | set(my)):
            if k in TIME_KEYS:
                continue
            kp = f"{path}.{key_label(k)}"
            if k not in my:
                out.append(f"{kp}: missing on B")
            elif k not in mx:
                out.append(f"{kp}: missing on A")
            else:
                out += _json_diffs(x[mx[k]], y[my[k]], kp)
        return out
    if isinstance(x, list) and isinstance(y, list):
        out = [d for i, (p, q) in enumerate(zip(x, y)) for d in _json_diffs(p, q, f"{path}[{i}]")]
        if len(x) != len(y):
            out.append(f"{path}: {abs(len(x) - len(y))} more item(s) on {'A' if len(x) > len(y) else 'B'}")
        return out
    if isinstance(x, str) and isinstance(y, str):
        return [] if mask(x) == mask(y) else [f"{path}: value"]
    return [] if type(x) is type(y) and x == y else [f"{path}: value"]


def compare_json(a: Path, b: Path, rel: str) -> list:
    try:
        x, y = json.loads(a.read_text(encoding="utf-8")), json.loads(b.read_text(encoding="utf-8"))
    except ValueError:
        return compare_text(a, b, rel)
    return [f"{rel}: {d}" for d in _json_diffs(x, y)]


def compare_jsonl(a: Path, b: Path, rel: str) -> list:
    """The structured run log (run_logs/<ts>__events.jsonl): line by line, as JSON, without the fields that differ on
    every run (EVENT_VOLATILE: the run ID, times, durations)."""
    def lines(f):
        out = []
        for ln in f.read_text(encoding="utf-8", errors="replace").splitlines():
            if ln.strip():
                d = json.loads(ln)
                out.append({k: v for k, v in d.items() if k not in EVENT_VOLATILE} if isinstance(d, dict) else d)
        return out
    try:
        x, y = lines(a), lines(b)
    except ValueError:
        return compare_text(a, b, rel)
    return [f"{rel}: {d}" for d in _json_diffs(x, y)]


def compare_bytes(a: Path, b: Path, rel: str) -> list:
    ha = hashlib.sha256(a.read_bytes()).hexdigest()
    hb = hashlib.sha256(b.read_bytes()).hexdigest()
    return [] if ha == hb else [f"{rel}: contents"]


def compare_file(a: Path, b: Path, rel: str, with_logs: bool = False) -> list:
    """The differences between two files; a file that can't be read is reported by its error type alone (an error
    message can quote a path or a value)."""
    ext = a.suffix.lower()
    try:
        if ext in SHEET_EXTS:
            return compare_workbooks(a, b, rel)
        if ext == ".eml":
            return compare_emails(a, b, rel, with_logs)
        if ext == ".json":
            return compare_json(a, b, rel)
        if ext == ".jsonl":
            return compare_jsonl(a, b, rel)
        if ext in TEXT_EXTS:
            return compare_text(a, b, rel)
        return compare_bytes(a, b, rel)
    except (Exception, RecursionError) as e:                 # noqa: BLE001 - reported, never raised
        return [f"{rel}: could not be compared ({type(e).__name__})"]


def compare_runs(a: Path, b: Path, with_logs: bool = False) -> list:
    """Every difference between run folders ``a`` and ``b``, as lines with no values in them."""
    a, b = Path(a), Path(b)
    _setup(a, b)
    (fa, la), (fb, lb) = _files(a, with_logs), _files(b, with_logs)
    out = [f"{rel}: symbolic link on {side}, not compared" for side, links in (("A", la), ("B", lb))
           for rel in links]
    for rel in sorted(set(fa) | set(fb)):
        if rel not in fb:
            out.append(f"{rel}: missing on B")
        elif rel not in fa:
            out.append(f"{rel}: missing on A")
        else:
            out += compare_file(fa[rel], fb[rel], rel, with_logs)
    return out


class _Parser(argparse.ArgumentParser):
    def error(self, message):                      # argparse's own message repeats the arguments (a file name)
        self.exit(2, "bad arguments: give two run folders, and optionally --with-logs and --json (see --help)\n")


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="backslashreplace")
        except (AttributeError, ValueError):
            pass
    ap = _Parser(description="Compare two CVA result folders (no values are printed).")
    ap.add_argument("a", type=Path)
    ap.add_argument("b", type=Path)
    ap.add_argument("--with-logs", action="store_true", help="also compare run logs (they hold timings)")
    ap.add_argument("--json", action="store_true", help="print the differences as a JSON list")
    args = ap.parse_args(argv)
    for p in (args.a, args.b):
        if not p.is_dir():
            print("not a folder (give two run folders)", file=sys.stderr)
            return 2
    try:
        diffs = compare_runs(args.a, args.b, args.with_logs)
    except (Exception, RecursionError) as e:                 # noqa: BLE001 - the type only: no path, no value
        print(f"the comparison failed ({type(e).__name__})", file=sys.stderr)
        return 3
    for w in _Ctx.warnings:
        print(f"warning: {w}", file=sys.stderr)
    if args.json:
        print(json.dumps(diffs, indent=1))
    else:
        for d in diffs:
            print(d)
        print(f"{len(diffs)} difference(s)" if diffs else "The two runs are the same.")
    return 1 if diffs else 0


if __name__ == "__main__":
    sys.exit(main())
