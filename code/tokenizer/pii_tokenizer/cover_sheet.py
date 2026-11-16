"""
#151 step 3 (David, Nov 12, 2026): each project's cover sheet, for its **status** only.

The cover sheet is a Google Sheet (or .xlsx) whose name starts with "Project Data", directly in the project's folder
in the **staff** drive: the folder the tracker step finds (staff_trackers.py), or, for a project whose tracker comes
from tracker_source.txt, the folder that tracker is in. It is read here, on the staff side, through the same
read-only view of Drive as the trackers, and only when it changed since the last pass. **Nothing of it is copied**:
the staff side's "Tracker sources.csv" shows its name, the status and any problem (real names, staff side only; never
the word typed, which could be a name), and
the AWS side's "_tracker_sources.json" gets two codes per project (status: active | closed | not_started; why: see
cva_filekinds/project_data.py), never the word typed or the file's name. Closed and Not started projects are then
not checked on the Saturday run (agent/runner.py) and are greyed out in Run tasks (the web app reads the staff
side's CSV). Anything that can't be read counts as active, and the technical email says why.
"""
from __future__ import annotations

import io
import os
import re
import zipfile

from . import xlsx
from .cva_filekinds import project_data as PD
from .gcp import GSHEET_MIME

FOLDER_MIME = "application/vnd.google-apps.folder"
MAX_FILE_MB = 5                    # a cover sheet is a few KB; anything larger isn't downloaded (review P-M1)
MAX_UNPACKED_MB = 30               # nor read when it unpacks to more than this (a zip bomb would stop every pass)
WORKBOOK_EXTS = {".xlsx", ".xlsm"}
OTHER_WORKBOOK_EXTS = {".xls", ".xlsb", ".ods", ".numbers"}


def _ext(f: dict) -> str:
    return os.path.splitext(f.get("name", ""))[1].lower()


def candidates(kids: list[dict]) -> tuple[list[dict], list[dict]]:
    """(cover sheets the tokenizer can read, ones in another workbook format) among a folder's files."""
    named = [f for f in kids if f.get("mimeType") != FOLDER_MIME and not f.get("trashed")
             and PD.is_cover_sheet_name(f.get("name", ""))]
    readable = [f for f in named if f.get("mimeType") == GSHEET_MIME or _ext(f) in WORKBOOK_EXTS]
    other = [f for f in named if f not in readable and _ext(f) in OTHER_WORKBOOK_EXTS]
    return readable, other


def version(f: dict) -> str:
    return f.get("md5Checksum") or f.get("modifiedTime") or ""


def _unzip(data: bytes) -> dict:
    """Every part of the workbook, counting the bytes really unpacked (a zip's own sizes can lie: verification N1)
    and stopping at MAX_UNPACKED_MB in all."""
    left, parts = MAX_UNPACKED_MB * 1_000_000, {}
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                chunks = []
                with z.open(info) as f:
                    while True:
                        chunk = f.read(min(1_000_000, left + 1))
                        if not chunk:
                            break
                        left -= len(chunk)
                        if left < 0:
                            raise xlsx.TokenError("too large for a cover sheet once unpacked")
                        chunks.append(chunk)
                parts[info.filename] = b"".join(chunks)
    except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError, ValueError, EOFError) as e:
        raise xlsx.TokenError("not a valid .xlsx file") from e
    return parts


def heads(data: bytes) -> list[tuple[str, list]]:
    """The tab read for the Status row (the one named "Project Data", else the first) as [(title, first rows)].
    xlsx.TokenError when the workbook can't be read. The values stay in memory here; nothing is written or logged."""
    if len(data) > MAX_FILE_MB * 1_000_000:
        raise xlsx.TokenError("too large for a cover sheet")
    parts = _unzip(data)
    for name in parts:
        if name.endswith((".xml", ".rels")):
            try:
                parts[name] = xlsx.canonical_attributes(parts[name].decode("utf-8")).encode("utf-8")
            except UnicodeDecodeError:
                continue
    sheet_parts, sst_part = xlsx._layout(parts)
    tabs = xlsx._tab_order(parts, list(dict.fromkeys(sheet_parts)))
    titles = dict(zip(tabs, xlsx._sheet_titles(parts, tabs)))
    pick = next((t for t in tabs if PD.norm(titles.get(t, "")) == "project data"), tabs[0])
    sst_xml = xlsx._text(parts, sst_part) if sst_part else ""
    sst = [xlsx._si_text(m.group(1)) for m in xlsx.SI_RE.finditer(sst_xml)]
    return [(titles.get(pick, ""), xlsx._head(xlsx._text(parts, pick), sst, PD.HEAD_ROWS))]


def read(data: bytes) -> tuple[str, str, object]:
    """(status code, why, the value as typed) from the workbook's bytes; unreadable: (active, unreadable, None)."""
    try:
        return PD.read_status(heads(data))
    except Exception:  # noqa: BLE001 - whatever is wrong with the file, it can't be read (not "this time": review P-L1)
        return PD.ACTIVE, PD.UNREADABLE, None


def look(drive, folder_id: str | None, cached: dict | None) -> dict:
    """This pass's record for one project: {status, why, id, ver, name, names}. ``drive``: the read-only view
    (staff_trackers.ReadOnlyDrive). ``cached``: last pass's record; the file is downloaded again only when it changed.
    Google errors are raised (the caller counts them as "couldn't read this time")."""
    if not folder_id:
        return {"status": PD.ACTIVE, "why": PD.NO_FOLDER}
    readable, other = candidates(drive.children(folder_id))
    if len(readable) > 1:
        return {"status": PD.ACTIVE, "why": PD.SEVERAL, "names": sorted(f["name"] for f in readable)}
    if not readable:
        if other:
            return {"status": PD.ACTIVE, "why": PD.UNREADABLE, "names": sorted(f["name"] for f in other)}
        return {"status": PD.ACTIVE, "why": PD.NONE}
    f = readable[0]
    rec = {"id": f["id"], "ver": version(f), "name": f.get("name", "")}
    c = cached or {}
    # unchanged since the last pass: not downloaded again
    if c.get("id") == rec["id"] and c.get("ver") == rec["ver"] and rec["ver"] \
            and c.get("why") in PD.WHYS and c.get("why") != PD.ERROR:
        return {**rec, "status": c.get("status", PD.ACTIVE), "why": c["why"]}
    try:
        size = int(f.get("size") or 0)
    except (TypeError, ValueError):
        size = 0
    if size > MAX_FILE_MB * 1_000_000:                  # (a Google Sheet has no size: its export is small)
        return {**rec, "status": PD.ACTIVE, "why": PD.UNREADABLE}
    status, why, _typed = read(drive.download(f))     # the word typed is never kept or shown (it could be a name)
    return {**rec, "status": status, "why": why}


# The staff side's CSV: the status as shown, and a note on anything that didn't go as planned
NOTES = {PD.BLANK: "Project Data: Status is empty; counted as Active",
         PD.NONE: "no Project Data sheet in the staff project folder; counted as Active",
         PD.SEVERAL: "more than one Project Data sheet ({names}); keep one; counted as Active",
         PD.NO_STATUS: "the Project Data sheet has no Status row; counted as Active",
         PD.UNRECOGNISED: "the Project Data sheet's Status isn't Active, Closed or Not started; counted as Active",
         PD.CONFLICT: "the Project Data sheet has two Status rows that disagree; counted as Active",
         PD.UNREADABLE: "the Project Data sheet can't be read{names} (too large, damaged, or not a Google Sheet or "
                        ".xlsx); counted as Active",
         PD.NO_FOLDER: "the project's staff folder isn't known, so its Project Data sheet wasn't looked for; counted "
                       "as Active",
         PD.ERROR: "the Project Data sheet couldn't be read this time; the last status read is used"}


def staff_note(rec: dict) -> str:
    why = rec.get("why")
    if why in (PD.OK, None):
        return ""
    names = rec.get("names") or []
    joined = "; ".join(names)
    return NOTES.get(why, why).format(names=(f" ({joined})" if why == PD.UNREADABLE and names else joined))


def aws_entry(rec: dict) -> dict:
    """The two codes that cross to AWS (nothing typed, no name)."""
    status = rec.get("status") if rec.get("status") in PD.SHOWN else PD.ACTIVE
    why = rec.get("why") if rec.get("why") in PD.WHYS else PD.UNREADABLE
    return {"status": status, "why": why}


__all__ = ["candidates", "heads", "read", "look", "staff_note", "aws_entry"]
