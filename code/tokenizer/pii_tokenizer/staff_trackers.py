"""
Change 1 (#93; design approved by David on Nov 4, 2026): each project's tracker is read from the staff's working
folder, not from a copy someone placed in the CVA's folder by hand.

Where the working tracker is (checked on every tokenizer pass, before anything is copied to AWS):

    1. Projects/<project>/Tracker/tracker_source.txt in the CVA's (raw) folder, when there is one. Its first line that
       isn't empty or a "#" comment is either
          - a Google Drive link to the tracker (or its bare file ID),
          - a path inside the staff drive ("Corvath/Corvath XHA 50001/XHA 50002 ... Tracker.xlsx"), or
          - "none": the project has no working tracker; the copy already in Tracker/ is used, as before Change 1.
       It always wins over the rule, and a pointer that leads nowhere never falls back to the rule.
    2. Otherwise the rule: staff drive > <main contractor> > <project> > the one file with "tracker" in its name.
       The CVA project "Contractor - Project" is matched to the contractor folder (same name, or one starts with the
       other) and to the project folder whose name holds every word of "Project". Only files directly in that
       folder count, .xlsx or Google Sheets, not "~$..." lock files, not names with copy/old/archive/backup.

What is done with it: when the working tracker changed since it was last copied (or was never copied), it is
downloaded (a Google Sheet exported as .xlsx) and written into Projects/<project>/Tracker/ of the CVA's folder,
under the working file's name. Any other workbook directly in Tracker/ is moved to Tracker/Archived/<time>/. The
mirror then tokenizes it for AWS as usual; the agent marks the CVA's copy. **The staff drive is only ever read:**
this module is handed a read-only view of Drive for it (ReadOnlyDrive), and the tokenizer's account should be a
Viewer there.

A copy is not made while AWS's marks on the current copy are still on their way back (the tokenized copy changed
and the mirror hasn't brought it back yet): the next pass makes it.

What each side is told:
    staff side  "Tracker sources.csv" at the top of the CVA's folder: per project, the staff file used (or why
                none), its last change, and what to do. Real names: staff side only.
    AWS side    "_tracker_sources.json": per project folder name, a status code and the working file's last
                change. No file names (a name may hold a worker's name; the agent names the tracker from its
                own copy, whose name the mirror has already handled).
"""
from __future__ import annotations

import csv
import io
import json
import os
import posixpath
import re
from dataclasses import dataclass

from .gcp import GSHEET_MIME, XLSX_MIME, GoogleError

SOURCE_FILE = "tracker_source.txt"
STAFF_SHEET = "Tracker sources.csv"          # staff side (real names)
AWS_SUMMARY = "_tracker_sources.json"         # AWS side (codes only)
FOLDER_MIME = "application/vnd.google-apps.folder"
TRACKER_WORD = re.compile(r"tracker", re.I)
OLD_COPY = re.compile(r"(?<![a-z])(copy|old|archived?|backup)(?![a-z])", re.I)   # also "Tracker_old", "archived"
LINK_ID = re.compile(r"/d/([A-Za-z0-9_-]{20,})|[?&]id=([A-Za-z0-9_-]{20,})|^([A-Za-z0-9_-]{25,})$")
WORKBOOK_EXTS = {".xlsx", ".xlsm"}                     # an .xlsm is copied as .xlsx, its macros removed
OTHER_WORKBOOK_EXTS = {".xls", ".xlsb", ".ods", ".numbers"}   # a tracker in a format the CVA can't read
RUN_MARKER = "_aws_run_in_progress.json"               # written by the AWS pipeline at its pull, removed after its push
RUN_MARKER_HOURS = 3                                   # an older marker is a stopped run's: ignored

# Status codes (the agent's agent/intent.py reads them; keep in step)
COPIED, UNCHANGED, DEFERRED = "copied", "unchanged", "deferred"    # the working tracker is (or will be) in Tracker/
KEPT = "kept"                  # tracker_source.txt says "none": the copy in Tracker/ is used, as before Change 1
NO_TRACKER = "no_tracker"      # the staff project folder has no tracker
SEVERAL = "several"            # more than one file looks like the tracker: not guessed
NOT_FOUND = "not_found"        # tracker_source.txt points at something that isn't there (or isn't a workbook)
NO_FOLDER = "no_folder"        # the rule found no staff project folder (or more than one) for this project
UNSUPPORTED = "unsupported"    # the tracker is in a format the CVA can't read (.xls, ...)
AMBIGUOUS = "ambiguous"        # several staff folders could be this project's, none clearly: not guessed
ERROR = "error"                # couldn't read the staff drive this time
GOOD = {COPIED, UNCHANGED, DEFERRED}


class ReadOnlyDrive:
    """The only view of Drive used on the staff drive: listing, reading and downloading, nothing else."""

    def __init__(self, drive):
        self._d = drive

    def get(self, file_id):
        return self._d.get(file_id)

    def children(self, folder_id):
        # listed once per pass (a ReadOnlyDrive lives for one pass): matching looks at the same folders many times
        cache = self.__dict__.setdefault("_kids", {})
        if folder_id not in cache:
            cache[folder_id] = self._d.children(folder_id)
        return cache[folder_id]

    def download(self, f) -> bytes:
        return self._d.download(f)


@dataclass
class Outcome:
    status: str
    source: dict | None = None        # the staff file's metadata (id, name, mimeType, modifiedTime, ...)
    candidates: list | None = None    # names, when there are several (staff side only)
    note: str = ""                    # staff side only
    folder: dict | None = None        # the staff project folder the tracker was found in (the rule)
    matched_by: str = ""              # pointer | pinned | tracker_name | project_number | folder_name
    rematched: bool = False           # the folder pinned earlier is gone and another was matched
    pin_missing: bool = False         # the pinned folder wasn't found this pass (gone only after two passes)
    drop_pin: bool = False            # the pinned folder was renamed into something that no longer fits


def norm(s: str) -> str:
    return " ".join(re.findall(r"[a-z0-9]+", str(s or "").lower()))


def split_project(name: str) -> tuple[str, str]:
    """'Brennick - Merivane 12th & Example' -> ('Brennick', 'Merivane 12th & Example'); no ' - ': (name, '')."""
    contractor, sep, project = name.partition(" - ")
    return (contractor.strip(), project.strip()) if sep else (name.strip(), "")


def is_candidate(f: dict) -> bool:
    name = f.get("name", "")
    if f.get("mimeType") == FOLDER_MIME or name.startswith("~$") or not TRACKER_WORD.search(name):
        return False
    if OLD_COPY.search(os.path.splitext(name)[0]):
        return False
    return f.get("mimeType") == GSHEET_MIME or os.path.splitext(name)[1].lower() in WORKBOOK_EXTS


def is_other_workbook(f: dict) -> bool:
    return os.path.splitext(f.get("name", ""))[1].lower() in OTHER_WORKBOOK_EXTS


def is_workbook(f: dict) -> bool:
    """A file named in tracker_source.txt: any name will do (it was chosen on purpose), but it must be a workbook."""
    if f.get("mimeType") == FOLDER_MIME or f.get("name", "").startswith("~$"):
        return False
    return f.get("mimeType") == GSHEET_MIME or os.path.splitext(f.get("name", ""))[1].lower() in WORKBOOK_EXTS


def version(f: dict) -> str:
    return f.get("md5Checksum") or f.get("modifiedTime") or ""


def copy_name(f: dict) -> str:
    """The CVA copy's name: always .xlsx (a Google Sheet is exported, an .xlsm loses its macros)."""
    name = f["name"]
    stem, ext = os.path.splitext(name)
    if f.get("mimeType") == GSHEET_MIME:
        return name if ext.lower() == ".xlsx" else name + ".xlsx"
    return stem + ".xlsx" if ext.lower() == ".xlsm" else name


def read_pointer(text: str) -> str | None:
    for line in text.splitlines():
        line = line.strip().lstrip("﻿")
        if line and not line.startswith("#"):
            return line
    return None


class StaffIndex:
    """The staff drive's top two levels (contractor folders, their project folders), listed once per pass."""

    def __init__(self, drive: ReadOnlyDrive, root_id: str):
        self.drive, self.root_id = drive, root_id
        self.last_strong = 0
        self._contractors = None
        self._projects = {}

    def contractors(self) -> list[dict]:
        if self._contractors is None:
            self._contractors = [f for f in self.drive.children(self.root_id) if f.get("mimeType") == FOLDER_MIME]
        return self._contractors

    def projects(self, contractor: dict) -> list[dict]:
        if contractor["id"] not in self._projects:
            self._projects[contractor["id"]] = [f for f in self.drive.children(contractor["id"])
                                                if f.get("mimeType") == FOLDER_MIME]
        return self._projects[contractor["id"]]

    def candidates(self) -> list[tuple[dict, dict]]:
        """Every (contractor folder, project folder) pair two levels down in the staff drive."""
        return [(cf, pf) for cf in self.contractors() for pf in self.projects(cf)]

    def match_folder(self, cva_project: str, copy_names=()) -> tuple[dict | None, str, list]:
        """(the staff project folder for this CVA project, how it was matched, other possible folders' names).

        A wrong match would check another project's tracker against this project's L&I records, silently, so a folder
        is used only on strong evidence and only when exactly one folder has it (see folder_evidence). Folders under a
        contractor whose name matches the project's are the only ones looked at when there are any."""
        cands = self.candidates()
        contractor = split_project(cva_project)[0]
        mine = [(cf, pf) for cf, pf in cands if contractor_matches(contractor, cf["name"])]
        if mine:
            cands = mine
        copies = {_stem(n) for n in copy_names if n}
        # how many staff folders hold each tracker name (a name in several folders proves nothing)
        name_count = {}
        if copies:
            for cf, pf in self.candidates():
                for f in self.drive.children(pf["id"]):
                    if is_candidate(f):
                        name_count[_stem(f["name"])] = name_count.get(_stem(f["name"]), 0) + 1
        strong, weak = [], []
        for cf, pf in cands:
            by = folder_evidence(cva_project, cf["name"], pf["name"])
            if by is None:
                continue                                       # ruled out (another project number, phase, stage)
            if copies and mine:
                # (only under the project's own contractor folder) the same tracker name as the CVA's copy, found in no other staff folder, and saying which project it
                # is in words that are also this folder's ("Eddleton Boiler Compliance Tracker" in "Eddleton Boiler")
                pwords = set(norm(pf["name"]).split()) - STOP_WORDS
                cwords = set(norm(cf["name"]).split())
                project_part = set(norm(split_project(cva_project)[1]).split()) - STOP_WORDS - cwords
                for f in self.drive.children(pf["id"]):
                    n = _stem(f["name"]) if is_candidate(f) else None
                    dw = distinctive_words(f["name"], cf["name"]) if n else set()
                    # both ways: the tracker's distinctive words are the folder's, and the folder's words are the
                    # tracker's (or the contractor's); and a word of the project's own name is in the folder's
                    if n in copies and name_count.get(n) == 1 and dw and dw <= pwords and pwords <= dw | cwords \
                            and (not project_part or project_part & pwords):
                        by = "tracker_name"
            if by in ("tracker_name", "project_number", "folder_name"):
                strong.append((by, cf, pf))
            elif by == "weak":
                weak.append((cf, pf))
        self.last_strong = len(strong)
        if len(strong) == 1:
            return strong[0][2], strong[0][0], []
        others = strong or [(None, cf, pf) for cf, pf in weak]
        return None, "", [f"{cf['name']} / {pf['name']}" for _b, cf, pf in others[:3]]

    def inside(self, f: dict, depth: int = 12) -> bool:
        """Is this file inside the staff drive (or folder) the tokenizer was told to read? Walks up its parents."""
        seen, todo = set(), list(f.get("parents") or [])
        for _ in range(depth):
            if self.root_id in todo:
                return True
            nxt = []
            for pid in todo:
                if pid in seen:
                    continue
                seen.add(pid)
                try:
                    nxt += self.drive.get(pid).get("parents") or []
                except GoogleError as e:
                    if e.status != 404:
                        raise                    # 403 is also Drive's "slow down": "couldn't read", not "not inside"
                    # not visible to the tokenizer: not inside
            if not nxt:
                return False
            todo = nxt
        return False

    def by_path(self, path: str) -> dict | None:
        """A file by its path inside the staff drive, matching each part's name without regard to case."""
        parts = [x for x in path.strip().strip("/").split("/") if x]
        fid, f = self.root_id, None
        for i, part in enumerate(parts):
            kids = [k for k in self.drive.children(fid) if k.get("name", "").strip().lower() == part.strip().lower()]
            if len(kids) != 1:
                return None
            f = kids[0]
            if i < len(parts) - 1 and f.get("mimeType") != FOLDER_MIME:
                return None
            fid = f["id"]
        return f


# Folder matching (the rule): see StaffIndex.match_folder and folder_evidence
ID_TOKEN = re.compile(r"(?=[a-z0-9]*[a-z])(?=[a-z0-9]*\d{3})[a-z0-9]+|\d{5,}")   # xe9901, wa20250091, 14378, 900001
PHASE_RE = re.compile(r"\b(phase|ph|stage|stg|part|bldg|building|unit)\s+([0-9]+|[ivx]+|[a-z])\b"
                      r"|\b(ph|stg)([0-9]+)\b|\b(demo|demolition)\b|\b([0-9][a-z])\b")
ROMAN = {"i": "1", "ii": "2", "iii": "3", "iv": "4", "v": "5", "vi": "6"}
STOP_WORDS = {"the", "and", "of", "a", "at", "for"}
# Words that say nothing about which project a tracker is
GENERIC_WORDS = {"compliance", "tracker", "tracking", "sheet", "payroll", "project", "projects", "new", "final",
                 "copy", "xlsx", "updated", "update", "current", "latest", "certified", "master", "working", "rev",
                 "revised", "version", "draft", "log", "report", "wages", "wage", "intent", "affidavit"}
_NOISE = re.compile(r"(19|20)\d{2}|\d{1,2}|v\d+|\d{6,8}")      # years, small numbers, versions, dates


def _stem(name: str) -> str:
    return norm(os.path.splitext(name)[0])


def distinctive_words(name: str, contractor: str = "") -> set:
    """The words of a tracker's name that say which project it is: not generic words, years, dates or versions,
    not the contractor's name ("Brennick Compliance Tracker v4" has none)."""
    skip = set(norm(contractor).split())
    return {t for t in _stem(name).split()
            if t not in GENERIC_WORDS and t not in STOP_WORDS and t not in skip and not _NOISE.fullmatch(t)}


def distinctive(name: str) -> bool:
    return bool(distinctive_words(name))


def contractor_matches(contractor: str, folder: str) -> bool:
    """The same contractor: the same words, or one name the other's start ("Anvold Construction" / "Anvold";
    "Kremlow-Scarvio" / "Kremlow-Scarvio"). Not just a common first word ("The …", "WA …", "City …")."""
    c, f = norm(contractor), norm(folder)
    return bool(c and f) and (c == f or c.startswith(f + " ") or f.startswith(c + " "))


def _phases(text: str) -> set:
    out = set()
    for m in PHASE_RE.finditer(norm(text)):
        if m.group(5):
            out.add("demo")
        elif m.group(6):
            out.add(f"st {m.group(6)}")                          # "2a", "1b": a stage of a road project
        else:
            word, val = (m.group(1), m.group(2)) if m.group(1) else (m.group(3), m.group(4))
            out.add(f"{'ph' if word in ('phase', 'ph') else word[:2]} {ROMAN.get(val, val)}")
    return out


def words_overlap(cva_project: str, contractor: str, folder: str) -> bool:
    p = set(norm(cva_project).split()) - STOP_WORDS
    return bool(p & (set(norm(folder).split()) | set(norm(contractor).split())))


def folder_evidence(cva_project: str, contractor: str, folder: str) -> str | None:
    """What a staff folder's name says about a CVA project folder name:
      None             ruled out: the folder names a project number the project doesn't, or the project names one
                       the folder doesn't, or their phases/stages differ (Phase I vs DEMO, Stage 1 vs Stage 2)
      "project_number" every project number of the CVA project is in the folder's name (and nothing rules it out)
      "folder_name"    the same words, both ways (contractor words may sit in the contractor folder)
      "weak"           some words in common: shown as a possibility, never used
      ""               nothing in common"""
    p_all = set(norm(cva_project).split()) - STOP_WORDS
    f_all = set(norm(folder).split()) - STOP_WORDS
    c_all = set(norm(contractor).split()) - STOP_WORDS
    pid = {t for t in p_all if ID_TOKEN.fullmatch(t)}
    fid = {t for t in f_all | c_all if ID_TOKEN.fullmatch(t)}
    if fid - pid or (_phases(cva_project) != _phases(folder)):
        return None
    if pid:
        # every number of the project, and nothing else in the folder's name but the project's or contractor's words
        # ("XE9901 SR 901 2b" is not "XE9901 SR 901 2a"; "10050 Sample Way" is not "10050 Example")
        return "project_number" if pid <= fid and f_all <= p_all | c_all else None
    if f_all and f_all <= p_all and p_all <= f_all | c_all:
        return "folder_name"
    return "weak" if p_all & (f_all | c_all) else ""


def resolve(index: StaffIndex, cva_project: str, pointer: str | None, copy_names=(), pinned: dict | None = None
            ) -> Outcome:
    """Where this project's working tracker is (no download). ``copy_names``: the workbooks now in the CVA's
    Tracker/ folder (the strongest evidence for the rule). ``pinned``: the staff folder matched on an earlier pass
    ({id, matched_by}), used while it exists, so a rename elsewhere can't move the project to another folder."""
    d = index.drive
    if pointer is not None:
        if pointer.strip().lower() == "none":
            return Outcome(KEPT, note=f"{SOURCE_FILE} says none: the copy in Tracker/ is used")
        if not pointer.strip():
            return Outcome(NOT_FOUND, note=f"{SOURCE_FILE} is empty: put the tracker's Drive link, its path in the "
                           "staff drive, or none in it")
        m = LINK_ID.search(pointer.strip())
        f = None
        if m:
            fid = next(g for g in m.groups() if g)
            try:
                f = d.get(fid)
            except GoogleError as e:
                if e.status not in (403, 404) or e.rate_limited:
                    raise                        # a passing Google error: "couldn't read" this time, retried
                f = None                         # not there, or not shared with the tokenizer
        else:
            f = index.by_path(pointer)
        if f is not None and not index.inside(f):
            return Outcome(NOT_FOUND, note=f"the file named in {SOURCE_FILE} isn't in the staff drive")
        if f is not None and is_other_workbook(f):
            return Outcome(UNSUPPORTED, source=f, note="the tracker is in a format the CVA can't read; save it as "
                           ".xlsx or a Google Sheet")
        if f is None or f.get("trashed") or not is_workbook(f):
            return Outcome(NOT_FOUND, note=f"the tracker named in {SOURCE_FILE} wasn't found (or isn't an .xlsx "
                           "or Google Sheet the tokenizer can read)")
        return Outcome(COPIED, source=f, matched_by="pointer")
    folder, by, rematched = None, "", False
    if pinned and pinned.get("id"):
        gone = False
        try:
            f = d.get(pinned["id"])
            if f.get("mimeType") == FOLDER_MIME and not f.get("trashed") and index.inside(f):
                parent = (f.get("parents") or [None])[0]
                pname = d.get(parent).get("name", "") if parent and parent != index.root_id else ""
                if folder_evidence(cva_project, pname, f.get("name", "")) is None:
                    # renamed since it was matched into something that rules it out (another phase, another
                    # project number): not used, not guessed again; the administrator says which folder it is
                    return Outcome(NO_FOLDER, drop_pin=True, note="the staff folder matched earlier was renamed to "
                                   f"{f.get('name', '')!r}, which no longer fits this project; add Tracker/"
                                   f"{SOURCE_FILE} with the tracker's Drive link")
                folder, by = f, "pinned"
            else:
                gone = True
        except GoogleError as e:
            if e.status != 404:
                raise                  # 403 is also how Drive says "slow down": not "gone"; this pass is an error
            gone = True
        if gone and not pinned.get("missing"):
            # Gone only when it is still gone on the next pass: one odd answer from Drive must not move the project
            return Outcome(ERROR, note="the staff folder matched earlier wasn't found this time; checked again on "
                           "the next pass", pin_missing=True)
        rematched = folder is None
    if folder is None:
        folder, by, runners, several = (*index.match_folder(cva_project, copy_names), False)
        if folder is None:
            several = len(runners) > 1 and index.last_strong > 1
            if several:
                return Outcome(AMBIGUOUS, candidates=runners, note="several staff folders could be this project's: "
                               + "; ".join(runners) + f". Add Tracker/{SOURCE_FILE} with the tracker's Drive link")
            return Outcome(NO_FOLDER, note="no project folder in the staff drive clearly matches this project"
                           + (" (possible: " + "; ".join(runners) + ")" if runners else "")
                           + f"; add Tracker/{SOURCE_FILE} with the tracker's Drive link (or none)")
    out = _in_folder(d, folder)
    out.folder, out.matched_by, out.rematched = folder, by, rematched
    return out


def _in_folder(d, folder) -> Outcome:
    kids = d.children(folder["id"])
    found = [f for f in kids if is_candidate(f)]
    if not found:
        other = [f for f in kids if is_other_workbook(f) and TRACKER_WORD.search(f.get("name", ""))
                 and not OLD_COPY.search(os.path.splitext(f.get("name", ""))[0])]
        if other:
            return Outcome(UNSUPPORTED, note="the tracker is in a format the CVA can't read ("
                           + "; ".join(sorted(f["name"] for f in other)) + "); save it as .xlsx or a Google Sheet")
        return Outcome(NO_TRACKER, note="the staff project folder has no tracker")
    if len(found) > 1:
        names = sorted(f["name"] for f in found)
        return Outcome(SEVERAL, candidates=names, note="more than one file looks like the tracker: "
                       + "; ".join(names) + f". Keep one, or name the current one in Tracker/{SOURCE_FILE}")
    return Outcome(COPIED, source=found[0])


def project_dirs(raw_dirs: dict, projects_dir: str) -> dict[str, str]:
    """{project folder name: its rel dir} for Projects/<name> (not names starting with '_' or '.')."""
    out = {}
    for rel in raw_dirs:
        parent, name = posixpath.split(rel)
        if parent == projects_dir and name and not name.startswith(("_", ".")):
            out[name] = rel
    return out


def staff_sheet(rows: list[tuple]) -> bytes:
    out = io.StringIO(newline="")
    w = csv.writer(out, lineterminator="\r\n")
    # the last three: #151 step 3, the project's cover sheet (pii_tokenizer/cover_sheet.py); the web app reads
    # "Project status" to grey out Closed and Not started projects in Run tasks
    w.writerow(["Project", "Status", "Staff tracker", "Last changed (UTC)", "Note", "Project status",
                "Project Data sheet", "Project Data note"])
    for r in rows:
        w.writerow([_cell(x) for x in r])
    return out.getvalue().encode("utf-8-sig")


def _cell(v) -> str:
    """A value that a spreadsheet would read as a formula is written as text."""
    v = "" if v is None else str(v)
    return "'" + v if v[:1] in ("=", "+", "-", "@") else v


STATUS_WORDS = {COPIED: "copied (changed since the last pass)", UNCHANGED: "up to date",
                DEFERRED: "changed; copied on the next pass (the last check's marks are on their way back)",
                KEPT: "not read from the staff drive", NO_TRACKER: "no tracker", SEVERAL: "several trackers",
                AMBIGUOUS: "several staff folders could be this project's",
                NOT_FOUND: "tracker_source.txt leads nowhere", NO_FOLDER: "staff folder not found",
                UNSUPPORTED: "tracker in a format the CVA can't read",
                ERROR: "staff drive not readable this time"}


def summary_json(statuses: dict[str, dict]) -> bytes:
    return json.dumps({"projects": statuses, "note": "where each project's tracker came from (pii_tokenizer/"
                       "staff_trackers.py); codes only, the staff side's Tracker sources.csv says which files"},
                      indent=1, sort_keys=True).encode()


__all__ = ["ReadOnlyDrive", "StaffIndex", "Outcome", "resolve", "read_pointer", "project_dirs", "copy_name",
           "version", "staff_sheet", "summary_json", "STATUS_WORDS", "SOURCE_FILE", "STAFF_SHEET", "AWS_SUMMARY",
           "GOOD", "XLSX_MIME"]
