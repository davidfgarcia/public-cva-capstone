"""Project folder layout: four folders, named for what people do with them.

Projects/<Project>/
  Uploads/             the only place anyone adds files: payroll exports, a new compliance tracker, wage determinations.
                       Each run recognizes every file by its contents (agent/uploads.py) and moves it once it is used;
                       files it doesn't recognize stay here (the email names them).
  Payroll/             every reviewed payroll report, named "<date> <payroll file> - REVIEWED.xlsx"
      Archived/<ts>/       the payroll exports as uploaded, and earlier reviews replaced in that run
  Tracker/             the current compliance tracker (the one people open and edit), and nothing else
      Archived/<ts>/       "<name> (before review).xlsx" (the tracker before the agent marked it),
                           "<name> (replaced).xlsx" (replaced by a newer upload), "<name> (as uploaded).xlsx"
  Files/               everything else, by type:
      Wage rates/          the current merged wage rates, one table per kind, named for the run that last changed it,
                           in Pacific time: "Federal 2026-11-08-13-34-03 PST.xlsx" (PDT in summer), "State …",
                           "Union …". An older plain "Federal.xlsx" still works until its next merge. If several
                           tables of one kind are there, the newest one is used (rate_files below)
      L&I data/<ts>/       the intents and affidavits downloaded from L&I in that run
      Logs/<ts>/           one trace per task (payroll__log.txt, wages__log.txt, intent__log.txt)
      Archived/<ts>/       wage files as uploaded, "<Kind> (previous).xlsx" wage tables, "(removed)" / "(replaced
                           upload)" files taken out of Uploads with the app
      Unused/              anything parked by hand; ignored by the agent
  .agent/              the agent's own working files (the last L&I download, used when data.wa.gov is down)
  .agent_state.json    what the agent remembers between runs

<ts> is the run's UTC timestamp (yyyy-MM-dd-HH-mm-ss), the same in every folder for one run. Only the wage tables'
own names use Pacific time (David, Nov 8, 2026), with the label (PST or PDT) that applied at that moment.
`ensure_layout()` creates the four folders. A project still in an older layout is not changed by a run: it is moved
into this one once, by migrate_layout.py (next to run.py), which uses migrate_project() below."""
import os
import re
import shutil
from pathlib import Path

from .util import err_text, file_kinds, log

KINDS = ("federal", "state", "union")
UPLOAD_DIR, PAYROLL_DIR, TRACKER_DIR, FILES_DIR, AGENT_DIR = "Uploads", "Payroll", "Tracker", "Files", ".agent"
RATES_DIR, LNI_DIR, LOGS_DIR, ARCHIVED, UNUSED = "Wage rates", "L&I data", "Logs", "Archived", "Unused"
REVIEWED_DIR = PAYROLL_DIR                     # where reviewed payroll reports are (used in messages)
# areas with an Archived/ folder
PAYROLL, TRACKER, FILES = "payroll", "tracker", "files"
UPLOAD_TARGETS = {kind: UPLOAD_DIR for kind in ("payroll", "tracker", *KINDS)}   # every upload goes to Uploads/


PACIFIC = "America/Los_Angeles"
STAMP_FORMAT = "%Y-%m-%d-%H-%M-%S"
STAMP_LABELS = {"PST": -8, "PDT": -7}         # the label decides the offset (the November repeated hour sorts right)
_RATE_NAME = re.compile(r"^(federal|state|union)(?: (\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}) (PST|PDT))?\.xlsx$",
                        re.IGNORECASE)


def rate_file_name(kind, run_ts):
    """'federal', '2026-11-08-21-34-03' (the run's UTC timestamp) -> 'Federal 2026-11-08-13-34-03 PST.xlsx': the
    same moment in Pacific time, labelled PST or PDT as it was then."""
    import datetime as dt
    from zoneinfo import ZoneInfo
    t = dt.datetime.strptime(run_ts, STAMP_FORMAT).replace(tzinfo=dt.timezone.utc).astimezone(ZoneInfo(PACIFIC))
    label = t.tzname()
    if label not in STAMP_LABELS:              # never expected; a wrong label would sort the table wrongly
        raise ValueError(f"unexpected Pacific time label {label!r}")
    return f"{kind.capitalize()} {t.strftime(STAMP_FORMAT)} {label}.xlsx"


def rate_table_stamp(name):
    """'Federal 2026-11-08-13-34-03 PST.xlsx' -> (kind, the moment in UTC); 'Federal.xlsx' (the old plain name) ->
    (kind, None); anything else (an '(AWS copy …)', a '(previous)' table, another file) -> None."""
    import datetime as dt
    m = _RATE_NAME.match(str(name))
    if not m:
        return None
    kind = m[1].lower()
    if not m[2]:
        return kind, None
    try:
        local = dt.datetime.strptime(m[2], STAMP_FORMAT)
        return kind, (local - dt.timedelta(hours=STAMP_LABELS[m[3].upper()])).replace(tzinfo=dt.timezone.utc)
    except (ValueError, OverflowError):        # a date that doesn't exist (month 13, year 9999 + 8 h): not one of ours
        return None


def rate_files(rates_dir, kind):
    """Every wage table of one kind in Files/Wage rates/, oldest first, so the last one is the table in use. An old
    plain 'Federal.xlsx' counts as older than any time-stamped one."""
    import datetime as dt
    d = Path(rates_dir)
    if not d.is_dir():
        return []
    oldest = dt.datetime.min.replace(tzinfo=dt.timezone.utc)
    found = []
    for f in d.iterdir():
        st = rate_table_stamp(f.name) if f.is_file() else None
        if st and st[0] == kind.lower():
            found.append(((st[1] or oldest), f.name, f))
    return [f for _, _, f in sorted(found)]


class Paths:
    def __init__(self, project_dir: Path):
        p = Path(project_dir)
        self.project = p
        self.upload = p / UPLOAD_DIR
        self.payroll = p / PAYROLL_DIR
        self.reviewed = self.payroll
        self.tracker = p / TRACKER_DIR
        self.files = p / FILES_DIR
        self.rates = self.files / RATES_DIR
        self.lni = self.files / LNI_DIR
        self.logs = self.files / LOGS_DIR
        self.unused = self.files / UNUSED
        self.agent = p / AGENT_DIR
        self.lni_data = self.agent / "lni_data"

    def rate_file(self, kind):
        """The wage table of this kind in use: the newest in Files/Wage rates/ (rate_files). With none there, the
        old plain name ('Federal.xlsx'), which doesn't exist: callers check .exists() / .is_file()."""
        found = rate_files(self.rates, kind)
        return found[-1] if found else self.rates / f"{kind.capitalize()}.xlsx"

    def area(self, area):
        return {PAYROLL: self.payroll, TRACKER: self.tracker, FILES: self.files}[area]

    def archived(self, area, ts=None):
        from .util import run_timestamp
        return self.area(area) / ARCHIVED / (ts or run_timestamp())

    def lni_run(self, ts=None):
        from .util import run_timestamp
        return self.lni / (ts or run_timestamp())

    def task_log(self, task, ts):
        return self.logs / ts / f"{task}__log.txt"

    def archived_dirs(self):
        return [self.payroll / ARCHIVED, self.tracker / ARCHIVED, self.files / ARCHIVED]

    def all_dirs(self):
        return [self.upload, self.payroll, self.tracker, self.files, self.rates]


def labelled(name: str, label=None) -> str:
    """'Tracker.xlsx' + 'before review' -> 'Tracker (before review).xlsx'."""
    if not label:
        return name
    ext = Path(name).suffix
    return f"{name[:-len(ext)] if ext else name} ({label}){ext}"


def _free_name(dst_dir: Path, name: str) -> Path:
    """dst_dir/name, or 'name (2).ext', 'name (3).ext'... if that is taken."""
    target = dst_dir / name
    n = 2
    while target.exists():
        stem, ext = (name[:-len(Path(name).suffix)], Path(name).suffix) if Path(name).suffix else (name, "")
        target = dst_dir / f"{stem} ({n}){ext}"
        n += 1
    return target


def archive_files(project_dir: Path, files, area, move=True, label=None, ts=None):
    """Move (or copy) files into <area>/Archived/<run ts>/, optionally as '<name> (<label>).<ext>'. area is
    PAYROLL, TRACKER or FILES. Returns {original path: archived path}."""
    files = [Path(f) for f in files if Path(f).exists()]
    if not files:
        return {}
    dest = Paths(project_dir).archived(area, ts)
    dest.mkdir(parents=True, exist_ok=True)
    out = {}
    for f in files:
        target = _free_name(dest, labelled(f.name, label))
        (shutil.move if move else shutil.copy2)(str(f), str(target))
        out[f] = target
    log.info("%s %d file(s) to %s", "Moved" if move else "Copied", len(files),
             "/".join(dest.relative_to(Path(project_dir)).parts),
             extra={"details": {"files": file_kinds(files)}})
    return out


class OldPaths:
    """The previous layout (tasks/payroll/..., tasks/intent/..., tracker/, unused/), only used to migrate it."""
    def __init__(self, project_dir: Path):
        p = Path(project_dir)
        self.project = p
        self.tasks = p / "tasks"
        self.payroll = self.tasks / "payroll"
        self.payroll_inbox = self.payroll / "inbox"
        self.payroll_reviewed = self.payroll / "reviewed"
        self.payroll_archive = self.payroll / "archive"
        self.wages = self.payroll / "wages"
        self.intent = self.tasks / "intent"
        self.lni_data = self.intent / "lni_data"
        self.intent_archive = self.intent / "archive"
        self.tracker = p / "tracker"
        self.tracker_archive = self.tracker / "archive"
        self.unused = p / "unused"

    def wage(self, kind):
        b = self.wages / kind
        return b / "inbox", b / "current", b / "archive"



def _move_contents(src: Path, dst: Path, rename=None):
    """Move everything inside src into dst (optionally renaming entries), skipping .DS_Store."""
    if not src.is_dir():
        return 0
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for item in sorted(src.iterdir()):
        if item.name == ".DS_Store":
            continue
        name = rename(item.name) if rename else item.name
        target = dst / name
        if target.exists():
            target = _free_name(dst, f"{Path(name).stem}__migrated{Path(name).suffix}")
        shutil.move(str(item), str(target))
        n += 1
    return n


def _merge_dir(src: Path, dst: Path):
    """Move a folder tree into dst, merging with folders already there; a file that already exists at the
    destination is kept and the incoming one is renamed <name>__migrated. Returns the number of files moved."""
    if not src.is_dir():
        return 0
    dst.mkdir(parents=True, exist_ok=True)
    n = 0
    for item in sorted(src.iterdir()):
        if item.name == ".DS_Store":
            continue
        target = dst / item.name
        if item.is_dir() and target.is_dir():
            n += _merge_dir(item, target)
            continue
        if target.exists():
            target = _free_name(dst, f"{item.stem}__migrated{item.suffix}")
        files = 1 if item.is_file() else sum(1 for f in item.rglob("*") if f.is_file())
        shutil.move(str(item), str(target))
        n += files
    _prune(src)
    return n


def _fix_state_paths(project: Path, old_prefix, new_prefix):
    """Keep paths saved in .agent_state.json (relative to the project) pointing at moved files."""
    import json
    sp = project / ".agent_state.json"
    if not sp.is_file():
        return
    try:
        st = json.loads(sp.read_text() or "{}")
    except ValueError:
        return
    pl = st.get("payroll_last") or {}
    rel = pl.get("archived_input")
    if isinstance(rel, str) and rel.startswith(old_prefix):
        pl["archived_input"] = new_prefix + rel[len(old_prefix):]
        sp.write_text(json.dumps(st, indent=2, default=str))


def _prune(d: Path):
    """Remove an old folder tree if nothing but .DS_Store files is left in it."""
    if not d.is_dir():
        return True
    for sub in sorted(d.iterdir(), reverse=True):
        if sub.is_dir():
            _prune(sub)
    try:
        for f in d.iterdir():
            if f.name == ".DS_Store":
                f.unlink()
        d.rmdir()
        return True
    except OSError:
        return False


def _rn(old_prefix, new_prefix):
    return lambda name: re.sub(rf"^{re.escape(old_prefix)}__", f"{new_prefix}__", name)


def _migrate_v1(project_dir: Path, moved, leftovers):
    """Earliest layouts (payroll_reports/, <kind>_wages_data/, compliance_tracker/, intent_affidavit_data/, and
    payroll/ intent/ wages/ at project level) -> the tasks/ layout, which _migrate_v2 then converts."""
    p = OldPaths(project_dir)
    # exact (lower-case) names only: on a Mac, Windows or Google Drive "payroll" would also find the new "Payroll"
    names = {n for n in (os.listdir(p.project) if p.project.is_dir() else []) if (p.project / n).is_dir()}
    has = names.__contains__
    old_tracker = has("tracker")
    if not any(has(n) for n in ("payroll_reports", "intent_affidavit_data", "compliance_tracker",
                                "payroll", "intent", "wages", *[f"{k}_wages_data" for k in KINDS])) \
            and not (old_tracker and ((p.tracker / "lni_data").is_dir() or (p.tracker / "logs").is_dir())):
        return
    old_payroll = p.project / "payroll_reports"
    if old_payroll.is_dir():
        n = _move_contents(old_payroll / "input" / "history", p.payroll_archive, _rn("input", "inbox"))
        n += _move_contents(old_payroll / "output" / "history", p.payroll_archive, _rn("output", "reviewed"))
        if (old_payroll / "input" / "history").is_dir():
            _prune(old_payroll / "input" / "history")
        if (old_payroll / "output" / "history").is_dir():
            _prune(old_payroll / "output" / "history")
        n += _move_contents(old_payroll / "input", p.payroll_inbox)
        n += _move_contents(old_payroll / "output", p.payroll_reviewed)
        moved.append(f"payroll_reports -> payroll ({n} item(s))")
        if not _prune(old_payroll):
            leftovers.append(old_payroll.name)
    for k in KINDS:
        old = p.project / f"{k}_wages_data"
        if not old.is_dir():
            continue
        inbox, current, archive = p.wage(k)
        n = _move_contents(old / "new_wages_data" / "processed_wages_data", archive, _rn("new_wages_data", "inbox"))
        n += _move_contents(old / "current_wages_data" / "previous_wages_data", archive, _rn("current_wages_data", "current"))
        _prune(old / "new_wages_data" / "processed_wages_data")
        _prune(old / "current_wages_data" / "previous_wages_data")
        n += _move_contents(old / "new_wages_data", inbox)
        n += _move_contents(old / "current_wages_data", current)
        moved.append(f"{old.name} -> wages/{k} ({n} item(s))")
        if not _prune(old):
            leftovers.append(old.name)
    old_ia = p.project / "intent_affidavit_data"
    if old_ia.is_dir():
        n = _move_contents(old_ia / "history", p.intent_archive, _rn("intent_affidavit_data", "lni_data"))
        _prune(old_ia / "history")
        n += _move_contents(old_ia, p.lni_data)
        moved.append(f"intent_affidavit_data -> tasks/intent/lni_data ({n} item(s))")
        if not _prune(old_ia):
            leftovers.append(old_ia.name)
    old_tr = p.project / "compliance_tracker"
    if old_tr.is_dir():
        n = _move_contents(old_tr / "history", p.tracker_archive, _rn("compliance_tracker", "tracker"))
        _prune(old_tr / "history")
        n += _move_contents(old_tr, p.tracker)
        moved.append(f"compliance_tracker -> tracker ({n} item(s))")
        if not _prune(old_tr):
            leftovers.append(old_tr.name)
    # task folders at project level -> tasks/
    for name, dest in (("payroll", p.payroll), ("intent", p.intent)):
        if has(name):
            n = _merge_dir(p.project / name, dest)
            moved.append(f"{name} -> tasks/{name} ({n} file(s))")
            _fix_state_paths(p.project, f"{name}/", f"tasks/{name}/")
    # earlier layout: wages/ at project level -> tasks/payroll/wages/; intent data out of tracker/ -> tasks/intent/
    if has("wages"):
        n = _merge_dir(p.project / "wages", p.wages)
        moved.append(f"wages -> tasks/payroll/wages ({n} file(s))")
        _fix_state_paths(p.project, "wages/", "tasks/payroll/wages/")
    if old_tracker and (p.tracker / "lni_data").is_dir():
        n = _merge_dir(p.tracker / "lni_data", p.lni_data)
        moved.append(f"tracker/lni_data -> tasks/intent/lni_data ({n} file(s))")
    if old_tracker and (p.tracker / "logs").is_dir():
        n = _merge_dir(p.tracker / "logs", p.intent / "logs")
        moved.append(f"tracker/logs -> tasks/intent/logs ({n} file(s))")
    if old_tracker and p.tracker_archive.is_dir():
        old = [d for d in p.tracker_archive.iterdir() if d.is_dir() and d.name.endswith("__lni_data")]
        for d in old:
            _merge_dir(d, p.intent_archive / d.name)
        if old:
            moved.append(f"tracker/archive/<ts>__lni_data -> tasks/intent/archive ({len(old)} folder(s))")


_STAMPED = re.compile(r"^(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})__(.+)$")
_TASK_LOG = re.compile(r"^(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})__log\.txt$")


def _move_file(src: Path, dst_dir: Path, name=None):
    dst_dir.mkdir(parents=True, exist_ok=True)
    target = _free_name(dst_dir, name or src.name)
    shutil.move(str(src), str(target))
    return target


def _files_in(d: Path):
    return [f for f in sorted(d.iterdir()) if f.is_file() and f.name != ".DS_Store"] if d.is_dir() else []


# ---- the five-folder layout (Upload, Reviewed payroll, Tracker, Wage rates, Archive), only used to migrate it
V3_UPLOAD, V3_REVIEWED, V3_TRACKER, V3_RATES, V3_ARCHIVE = "Upload", "Reviewed payroll", "Tracker", "Wage rates", "Archive"
V3_UPLOADED, V3_REPLACED, V3_LNI, V3_LOGS, V3_REMOVED = "Uploaded files", "Replaced", "L&I data", "Logs", "Removed"


class _Five:
    def __init__(self, project_dir):
        p = Path(project_dir)
        self.upload, self.reviewed, self.tracker = p / V3_UPLOAD, p / V3_REVIEWED, p / V3_TRACKER
        self.rates, self.archive = p / V3_RATES, p / V3_ARCHIVE
        self.lni_data = p / AGENT_DIR / "lni_data"
        self.unused = self.archive / UNUSED


def _migrate_v2(project_dir: Path, moved, leftovers):
    """The tasks/ layout (tasks/payroll/..., tasks/intent/..., tracker/, unused/) -> the five-folder layout, which
    _migrate_v3 then converts."""
    new, proj = _Five(project_dir), Path(project_dir)
    T = proj / "tasks"
    names = set(os.listdir(proj)) if proj.is_dir() else set()
    old_tracker = "tracker" in names                  # exact lower-case name: Mac and Drive ignore case, so
    old_unused = "unused" in names                    # "tracker" and "Tracker" are the same folder there
    t_tmp, u_tmp = proj / "_tracker_being_moved", proj / "_unused_being_moved"   # left by an interrupted run
    if not (T.is_dir() or old_tracker or old_unused or t_tmp.is_dir() or u_tmp.is_dir()):
        return
    n = 0
    arch = new.archive

    def run(ts, *sub):
        return arch.joinpath(ts, *sub)

    # Files waiting to be processed -> Upload/
    for d in [T / "payroll" / "inbox"] + [T / "payroll" / "wages" / k / "inbox" for k in KINDS]:
        for f in _files_in(d):
            _move_file(f, new.upload); n += 1
    # Reviewed reports and current wage rates
    for f in _files_in(T / "payroll" / "reviewed"):
        _move_file(f, new.reviewed); n += 1
    for k in KINDS:
        cur = T / "payroll" / "wages" / k / "current"
        for f in _files_in(cur):
            if f.name == "current_wage_rates.xlsx":
                _move_file(f, new.rates, f"{k.capitalize()}.xlsx")
            else:
                _move_file(f, new.unused / "Wage rates (old current folder)" / k)
            n += 1
    # Archives: <parent>/<ts>__<what>/ -> Archive/<ts>/<folder>/
    def from_archive(parent: Path, mapping, rename=None):
        nonlocal n
        if not parent.is_dir():
            return
        for d in sorted(parent.iterdir()):
            m = _STAMPED.match(d.name) if d.is_dir() else None
            if not m:
                continue
            sub = mapping.get(m.group(2), (V3_UPLOADED,))
            for f in [x for x in d.rglob("*") if x.is_file() and x.name != ".DS_Store"]:
                rel = f.relative_to(d).parent
                _move_file(f, run(m.group(1), *sub, *rel.parts), rename(f.name) if rename else None)
                n += 1
    from_archive(T / "payroll" / "archive", {"inbox": (V3_UPLOADED,), "reviewed": (V3_REPLACED, V3_REVIEWED)})
    for k in KINDS:
        from_archive(T / "payroll" / "wages" / k / "archive",
                     {"inbox": (V3_UPLOADED,), "current": (V3_REPLACED, V3_RATES)},
                     rename=lambda nm, k=k: f"{k.capitalize()}.xlsx" if nm == "current_wage_rates.xlsx" else nm)
    from_archive(T / "intent" / "archive", {"lni_data": (V3_LNI,)})
    # Current L&I download -> .agent/lni_data
    for f in _files_in(T / "intent" / "lni_data"):
        _move_file(f, new.lni_data); n += 1
    # Task logs: tasks/<...>/logs/<ts>__log.txt -> Archive/<ts>/Logs/<task>__log.txt
    for task, d in (("payroll", T / "payroll" / "logs"), ("wages", T / "payroll" / "wages" / "logs"),
                    ("intent", T / "intent" / "logs")):
        for f in _files_in(d):
            m = _TASK_LOG.match(f.name)
            if m:
                _move_file(f, run(m.group(1), V3_LOGS), f"{task}__log.txt"); n += 1
    # tracker/ -> Tracker/ (through a temporary name, since the two names are the same folder on a Mac)
    if old_tracker or t_tmp.is_dir():
        tmp = t_tmp
        if old_tracker:
            if tmp.is_dir():
                _merge_dir(proj / "tracker", tmp)
            else:
                (proj / "tracker").rename(tmp)
        from_archive(tmp / "archive", {"tracker": (V3_REPLACED, V3_TRACKER)})
        new.tracker.mkdir(exist_ok=True)
        for f in _files_in(tmp):
            _move_file(f, new.tracker); n += 1
        _merge_dir(tmp, new.unused / "Tracker (left over)") if any(tmp.rglob("*.*")) else _prune(tmp)
        _prune(tmp)
    if old_unused or u_tmp.is_dir():
        tmp = u_tmp
        if old_unused:
            if tmp.is_dir():
                _merge_dir(proj / "unused", tmp)
            else:
                (proj / "unused").rename(tmp)
        n += _merge_dir(tmp, new.unused)
    # Anything else still in tasks/ (not expected) is kept, under Archive/Unused/tasks/
    if T.is_dir():
        if any(f.is_file() and f.name != ".DS_Store" for f in T.rglob("*")):
            n += _merge_dir(T, new.unused / "tasks (left over)")
            moved.append("unexpected files from tasks/ kept in Archive/Unused/tasks (left over)")
        if not _prune(T):
            leftovers.append("tasks")
    # Paths remembered in .agent_state.json
    import json
    sp = proj / ".agent_state.json"
    if sp.is_file():
        try:
            st = json.loads(sp.read_text() or "{}")
            pl = st.get("payroll_last") or {}
            m = re.match(r"^tasks/payroll/archive/(\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2})__inbox/(.+)$",
                         str(pl.get("archived_input") or ""))
            if m:
                pl["archived_input"] = f"{V3_ARCHIVE}/{m.group(1)}/{V3_UPLOADED}/{m.group(2)}"
            if isinstance(st.get("intent_marks"), dict):
                pass                                           # refers to the tracker by name; unchanged
            sp.write_text(json.dumps(st, indent=2, default=str))
        except ValueError:
            pass
    moved.append(f"moved to the new folder layout (Upload, Reviewed payroll, Tracker, Wage rates, Archive): "
                 f"{n} file(s)")


def _kind_of(f: Path):
    try:
        from .uploads import classify
        return classify(f)[0]
    except Exception:                                 # unreadable: treat as a wage/other file
        return None


def _migrate_v3(project_dir: Path, moved, leftovers):
    """The five-folder layout (Upload, Reviewed payroll, Tracker, Wage rates, Archive/<ts>/...) -> Uploads, Payroll,
    Tracker, Files. Each run's archive is split by what the files are: payroll exports and replaced reviews ->
    Payroll/Archived/<ts>/; tracker versions -> Tracker/Archived/<ts>/; wage files, replaced wage tables and
    files removed from Upload -> Files/Archived/<ts>/; L&I downloads -> Files/L&I data/<ts>/; logs -> Files/Logs/<ts>/."""
    old, new, proj = _Five(project_dir), Paths(project_dir), Path(project_dir)
    where = {}                                         # old path -> new path (relative to the project)

    def mv(f, dst_dir, name=None):
        target = _move_file(f, dst_dir, name)
        where[f.relative_to(proj).as_posix()] = target.relative_to(proj).as_posix()
        return target
    names = set(os.listdir(proj)) if proj.is_dir() else set()
    if not ({V3_UPLOAD, V3_REVIEWED, V3_RATES, V3_ARCHIVE} & names):
        return
    n = 0
    for f in _files_in(old.upload):
        mv(f, new.upload); n += 1
    for f in _files_in(old.reviewed):
        mv(f, new.payroll); n += 1
    for f in _files_in(old.rates):
        mv(f, new.rates); n += 1
    if old.archive.is_dir():
        for run in sorted(old.archive.iterdir()):
            if not run.is_dir():
                continue
            if run.name == UNUSED:
                n += _merge_dir(run, new.unused)
                continue
            if not re.match(r"^\d{4}-\d{2}-\d{2}-\d{2}-\d{2}-\d{2}$", run.name):
                n += _merge_dir(run, new.files / ARCHIVED / run.name)
                continue
            ts = run.name
            for f in [x for x in run.rglob("*") if x.is_file() and x.name != ".DS_Store"]:
                rel = f.relative_to(run).parts
                top, sub = rel[0], rel[1] if len(rel) > 2 else None
                if top == V3_UPLOADED:
                    k = _kind_of(f)
                    if k == "payroll":
                        mv(f, new.archived(PAYROLL, ts))
                    elif k == "tracker":
                        mv(f, new.archived(TRACKER, ts), labelled(f.name, "as uploaded"))
                    else:
                        mv(f, new.archived(FILES, ts))
                elif top == V3_REPLACED and sub == V3_REVIEWED:
                    mv(f, new.archived(PAYROLL, ts))
                elif top == V3_REPLACED and sub == V3_TRACKER:
                    mv(f, new.archived(TRACKER, ts), labelled(f.name, "previous"))
                elif top == V3_REPLACED and sub == V3_RATES:
                    mv(f, new.archived(FILES, ts), labelled(f.name, "previous"))
                elif top == V3_REPLACED and sub == V3_UPLOAD:
                    mv(f, new.archived(FILES, ts), labelled(f.name, "replaced upload"))
                elif top == V3_REMOVED:
                    mv(f, new.archived(FILES, ts), labelled(f.name, "removed"))
                elif top == V3_LNI:
                    mv(f, new.lni_run(ts).joinpath(*rel[1:-1]))
                elif top == V3_LOGS:
                    mv(f, new.logs.joinpath(ts, *rel[1:-1]))
                else:
                    mv(f, new.archived(FILES, ts).joinpath(*rel[:-1]))
                n += 1
            _prune(run)
    for d in (old.upload, old.reviewed, old.rates, old.archive):
        if d.is_dir() and not _prune(d):
            leftovers.append(d.name)
    import json
    sp = proj / ".agent_state.json"
    if sp.is_file():
        try:
            st = json.loads(sp.read_text() or "{}")
            pl = st.get("payroll_last") or {}
            old_rel = str(pl.get("archived_input") or "")
            if old_rel in where:                     # wherever that file actually went
                pl["archived_input"] = where[old_rel]
                sp.write_text(json.dumps(st, indent=2, default=str))
        except ValueError:
            pass
    moved.append(f"moved to the new folder layout (Uploads, Payroll, Tracker, Files): {n} file(s)")


# folders of the older layouts (exact names: on a Mac or Google Drive "tracker" and "Tracker" are the same folder)
OLD_LAYOUT_DIRS = ("Upload", "Reviewed payroll", "Wage rates", "Archive",                    # five-folder layout
                   "tasks", "tracker", "unused", "_tracker_being_moved", "_unused_being_moved",  # tasks/ layout
                   "payroll_reports", "intent_affidavit_data", "compliance_tracker", "payroll", "intent", "wages",
                   *[f"{k}_wages_data" for k in KINDS])


class OldLayoutError(RuntimeError):
    pass


def old_layout_dirs(project_dir: Path, with_files_only=True):
    """Folders of an older layout in the project (by their exact names); by default only those holding a file."""
    proj = Path(project_dir)
    names = [n for n in (os.listdir(proj) if proj.is_dir() else []) if n in OLD_LAYOUT_DIRS and (proj / n).is_dir()]
    if not with_files_only:
        return sorted(names)
    return sorted(n for n in names if any(f.is_file() and f.name != ".DS_Store" for f in (proj / n).rglob("*")))


def ensure_layout(project_dir: Path):
    """Create the four folders. A project whose files are still in an older layout's folders raises OldLayoutError
    and is left as it is: it is moved into this layout once, with migrate_layout.py. Empty old folders are ignored."""
    old = old_layout_dirs(project_dir)
    if old:
        raise OldLayoutError(
            f"this project is still in the old folder layout ({', '.join(old)}); nothing was run or changed. "
            "Move it to the new layout (Uploads, Payroll, Tracker, Files) once with migrate_layout.py "
            "(see the installation guide, 'Moving projects to the new folder layout')")
    for d in Paths(project_dir).all_dirs():
        d.mkdir(parents=True, exist_ok=True)
    return []


def migrate_project(project_dir: Path):
    """Move a project in any older layout into the current one, moving files without changing them (only
    .agent_state.json is updated to the new paths). Returns (what moved, old folders that could not be removed)."""
    moved, leftovers = [], []
    _migrate_v1(project_dir, moved, leftovers)
    from .legacy import convert_project
    proj = Path(project_dir)
    try:
        # only for the tasks/ layout: it has files under tasks/, or a folder named exactly "tracker" (on a Mac or Drive,
        # (proj / "tracker") would also open the current "Tracker/")
        has_tasks = (proj / "tasks").is_dir() and any(f.is_file() and f.name != ".DS_Store"
                                                      for f in (proj / "tasks").rglob("*"))
        if not (has_tasks or "tracker" in (set(os.listdir(proj)) if proj.is_dir() else set())):
            raise StopIteration                  # nothing in the older layout to convert
        convert_project(Path(project_dir))   # older Pacific-time names -> UTC, in the tasks/ layout (logs its own changes)
    except StopIteration:
        pass
    except Exception as e:                   # a damaged file must not stop the folders from being moved
        log.warning("[%s] older timestamps could not be converted: %s", Path(project_dir).name, err_text(e))
    _migrate_v2(project_dir, moved, leftovers)
    _migrate_v3(project_dir, moved, leftovers)
    for d in Paths(project_dir).all_dirs():
        d.mkdir(parents=True, exist_ok=True)
    if moved:                                # (the moves themselves name files: counted only, #56)
        log.info("[%s] layout: %d item(s) moved to the current folder layout", Path(project_dir).name, len(moved))
    return moved, leftovers
