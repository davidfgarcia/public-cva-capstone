"""What is each file in a project's Uploads/ folder? Decided from the contents, not the name.

    payroll   a certified payroll export: a sheet whose header has the employee, week-end date and hours columns
    tracker   the compliance tracker: a sheet with a Contractor column and an intent or affidavit column
    wages     a wage determination or rate table that one of the wage parsers can read (LCPtracker wage export,
              L&I journey-level CSV, DOL decision as text or PDF, union rate table). Whether its rows are federal,
              state or union is decided row by row when it is merged (agent/wages.py)
    None      not recognized: the file stays in Uploads/ and the run's email says so

The same file is only examined once per run (results are cached by path, size and modification time).
"""
from __future__ import annotations

import re
from pathlib import Path

from .cva_filekinds.payroll_export import PAYROLL_EXTS, SNIFF_ROWS, SNIFF_SHEETS  # noqa: F401 - the registry (PR 4)
from .cva_filekinds.payroll_export import payroll_header as _payroll_header  # noqa: F401
from .util import log, visible_files

TRACKER_EXTS = {".xlsx", ".xlsm"}
WAGE_EXTS = {".xlsx", ".xlsm", ".csv", ".txt", ".pdf"}
# A file the parsers can't read but whose name says it is a wage determination still counts as one, so the
# local model (if enabled) gets a chance to read it during the merge.
WAGE_NAME_HINT = re.compile(r"wage|determination|prevailing|journey|rates?\b|\b[A-Z]{2}\d{8}\b|L&?I\b|LNI|DOL|union|CBA",
                            re.I)
_cache: dict = {}


# _payroll_header(rows): from the registry (cva_filekinds.payroll_export.payroll_header), above


def _rows_of(path: Path, limit=SNIFF_ROWS):
    if path.suffix.lower() == ".csv":
        import csv
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            return [r for _, r in zip(range(limit), csv.reader(f))]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets[:SNIFF_SHEETS]:
        out += [list(r) for r in ws.iter_rows(max_row=limit, values_only=True)]
    return out


def _is_tracker(path: Path) -> bool:
    import openpyxl
    from .intent import locate_table
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    ws, _, _ = locate_table(wb)
    return ws is not None


PDF_WAGE_MARK = re.compile(r"General\s*Decision\s*Number|Prevailing\s*Wage|Journey\s*Level|Wage\s*Determination|"
                           r"Davis[- ]Bacon|Department\s*of\s*Labor", re.I)


def _is_wage_file(path: Path) -> bool:
    from .wages import pdf_first_page_text, sniff_and_parse
    if path.suffix.lower() == ".pdf":
        # Only the first page is read here; the whole document is read once, when it is merged.
        return bool(PDF_WAGE_MARK.search(pdf_first_page_text(path))) or bool(WAGE_NAME_HINT.search(path.name))
    if sniff_and_parse(path, "state", llm=None):
        return True
    return bool(WAGE_NAME_HINT.search(path.name))


def classify(path: Path):
    """'payroll' | 'tracker' | 'wages' | None, with a short reason."""
    path = Path(path)
    st = path.stat()
    key = (str(path), st.st_size, st.st_mtime)
    if key in _cache:
        return _cache[key]
    ext = path.suffix.lower()
    kind, why = None, f"'{ext or 'no extension'}' files are not read"
    log_why = None                           # the reason for the run log, when `why` may hold a path (#56)
    try:
        if ext in PAYROLL_EXTS and _payroll_header(_rows_of(path)):
            kind, why = "payroll", "payroll export (employee, week and hours columns)"
        elif ext in TRACKER_EXTS and _is_tracker(path):
            kind, why = "tracker", "compliance tracker (Contractor and intent/affidavit columns)"
        elif ext in WAGE_EXTS and _is_wage_file(path):
            kind, why = "wages", "wage determination or rate table"
        elif ext in WAGE_EXTS:
            why = "not a payroll export, tracker or wage file the agent recognizes"
    except Exception as e:  # unreadable or damaged: leave it for a person to look at
        why = f"could not be read ({type(e).__name__}: {e})"
        log_why = f"could not be read ({type(e).__name__})"
    _cache[key] = (kind, why)
    log.debug("Upload recognized", extra={"details": {"type": ext or "no extension", "kind": kind or "not recognized",
                                                         "why": log_why or why}})
    return kind, why


def uploads_of(project_dir: Path, kind: str):
    """Files of one kind waiting in the project's Uploads/ folder, oldest first."""
    from .layout import Paths
    files = [f for f in visible_files(Paths(project_dir).upload) if classify(f)[0] == kind]
    return sorted(files, key=lambda p: p.stat().st_mtime)


def unrecognized(project_dir: Path):
    from .layout import Paths
    return [(f, classify(f)[1]) for f in visible_files(Paths(project_dir).upload) if classify(f)[0] is None]


def install_tracker_upload(project_dir: Path, tz="UTC"):
    """A compliance tracker in Uploads/ becomes the current tracker: the one in Tracker/ moves to
    Tracker/Archived/<ts>/ as '<name> (replaced)', a copy of the upload goes there as '<name> (as uploaded)',
    and the upload moves into Tracker/. Returns summary lines."""
    import shutil
    from .layout import TRACKER, Paths, archive_files
    paths = Paths(project_dir)
    new = uploads_of(project_dir, "tracker")
    if not new:
        return []
    newest, older = new[-1], new[:-1]
    lines = []
    old = visible_files(paths.tracker, TRACKER_EXTS)
    archive_files(project_dir, old, TRACKER, move=True, label="replaced")
    archive_files(project_dir, [newest], TRACKER, move=False, label="as uploaded")
    paths.tracker.mkdir(parents=True, exist_ok=True)
    shutil.move(str(newest), str(paths.tracker / newest.name))
    lines.append(f"tracker: {newest.name} is now the current tracker"
                 + (f" (replaced {', '.join(f.name for f in old)}, kept in the archive)" if old else ""))
    if older:
        archive_files(project_dir, older, TRACKER, move=True, label="not used, newer upload")
        lines.append("WARNING tracker: more than one tracker was uploaded; only the newest is used, the others "
                     f"were archived: {', '.join(f.name for f in older)}")
    return lines
