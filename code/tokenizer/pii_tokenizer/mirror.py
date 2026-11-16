"""
Keeps each raw Google Drive folder (the one staff use) in step with its tokenized copy (the only one AWS sees).

    raw folder  ──tokenize──▶  tokenized folder  ──▶  AWS reads it, and writes its results back into it
    raw folder  ◀─detokenize─  tokenized folder

Every file that has been copied is recorded as a *pair* (raw file ID <-> tokenized file ID, with the version
of each as last seen). On each run, for every pair:

    raw changed, copy unchanged     -> tokenize again and update the copy          (staff replaced a file)
    copy changed, raw unchanged     -> detokenize and update the raw file           (AWS wrote a result)
    both changed                    -> AWS's version is saved next to the raw file as "... (AWS copy <time>)",
                                       then the raw file is tokenized over the copy (staff's version wins)
    raw gone (deleted by staff)     -> the copy is moved to the bin
    copy gone (deleted on AWS)      -> the raw file is moved to _sync_backup/<time>__deleted_on_aws/, never deleted
    moved/renamed on one side       -> moved/renamed the same way on the other side

New raw files are tokenized into the copy; new files in the copy (AWS output) are detokenized into the raw
folder. Nothing under _sync_backup/ is copied as a new file in either direction.

Safety: a run refuses to act if it would treat more than half of the pairs (and more than 5) as deleted on
one side; that is what a wrong folder ID or a removed share looks like, not real work.
"""
from __future__ import annotations

import csv
import dataclasses
import io
from collections import Counter
import json
import logging
import time
import os
import posixpath
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

import hashlib

from . import screen
from . import names as NM
from . import cover_sheet as CS
from .cva_filekinds import names as known_names
from . import staff_trackers as ST
from .files import Context, detokenize_file, tokenize_file
from .cva_filekinds import settings_file
from .gcp import GSHEET_MIME, MIME_BY_EXT
from .tokens import TOKEN_HEX, TokenError
from .xlsx import strip_macros

log = logging.getLogger("pii_tokenizer")
GOOGLE_NATIVE = "application/vnd.google-apps."
BACKUP_DIR = "_sync_backup"
BATCH = 25
BATCH_BYTES = 150_000_000   # ... or fewer, once the tokenized copies waiting to be written reach this size (memory)


def version(f: dict) -> str:
    return f.get("md5Checksum") or f.get("modifiedTime") or ""


MD5_HEX = re.compile(r"[0-9a-f]{32}")
EMPTY_MD5 = hashlib.md5(b"").hexdigest()   # an empty file: never listed for purging (it is everywhere, holds nothing)
PURGE_KINDS = ("content", "path")          # what held the name: the file's content, or only its path


def mime_for(name: str) -> str:
    return MIME_BY_EXT.get(os.path.splitext(name)[1].lower(), "application/octet-stream")


@dataclass
class RunReport:
    forwarded: list = field(default_factory=list)      # (rel, detail)
    returned: list = field(default_factory=list)
    refused: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    moved: list = field(default_factory=list)
    removed: list = field(default_factory=list)
    conflicts: list = field(default_factory=list)
    converted: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    folders: list = field(default_factory=list)        # (rel, detail): Uploads/ created, old-layout folders removed

    def lines(self, mirror: str) -> list[str]:
        out = []
        for label, items in (("to AWS (tokenized)", self.forwarded), ("back from AWS (names restored)", self.returned),
                             ("NOT copied to AWS", self.refused), ("left out by the file rules", self.skipped),
                             ("moved", self.moved), ("macro workbooks converted", self.converted), ("removed", self.removed), ("conflicts", self.conflicts), ("folders", self.folders)):
            for rel, detail in items:
                out.append(f"[{mirror}] {label}: {rel}: {detail}")
        out += [f"[{mirror}] warning: {w}" for w in self.warnings]
        return out



# The Compliance Validation Agent's project layout (Projects/<project>/...): the folders of the current layout, and those of the
# previous one, emptied by the one-off migrate_layout.py on the staff side; the mirror removes them on AWS (_tidy_layout).
NEW_LAYOUT = ("Uploads", "Payroll", "Tracker", "Files")
OLD_LAYOUT = ("Upload", "Reviewed payroll", "Wage rates", "Archive")
RECENT_HOURS = 12        # a copy made less than this long ago that vanishes on AWS is sent again (see Mirror.run)
RECHECK_SECONDS = 600    # at most this long per run re-reading PDFs already on AWS for new names; the rest next run
BACKUP_RECHECK_SECONDS = 300   # the same for PDFs in the AWS-side backups (_sync_backup/)
PURGE_KEEP_DAYS = 35     # an entry of _purge_on_aws.json is kept this long (just over S3's 30 days for old versions)
TAKE_BACK_MIN = 5        # one run takes back at most max(this, a tenth of the copies); more means something's off
BACKUP_KEEP_DAYS = 30    # rclone's backups on the AWS side (_sync_backup/<time>__sync_backup) are kept this long
PROJECT_INPUT_DIRS = {"uploads", "upload", "tracker"}


def macro_archive_dir(rel_dir: str, stamp: str) -> str:
    """Where an uploaded macro workbook is parked, in the Compliance Validation Agent's project layout:
    <project>/Uploads/x.xlsm -> <project>/Files/Archived/<stamp>/; <project>/Tracker/x.xlsm ->
    <project>/Tracker/Archived/<stamp>/. Outside a project's Uploads/ or Tracker/ folder, an Archived/ folder
    beside the file is used."""
    parts = [p for p in rel_dir.split("/") if p]
    for i in range(len(parts) - 1, -1, -1):
        if parts[i].lower() in PROJECT_INPUT_DIRS:
            area = "Tracker" if parts[i].lower() == "tracker" else "Files"
            return posixpath.join(*parts[:i], area, "Archived", stamp)
    return posixpath.join(*parts, "Archived", stamp)

# A held-back file that is a project's tracker (a workbook directly in Projects/<project>/Tracker/): the AWS-side
# summary says so, with the project folder's name, so the outcome emails can say "the tracker was held back" instead
# of "no tracker". Only a fixed word and the project folder name (already the folder's name on AWS): never the file's
# name or why it was held back, which may hold a worker's name (#53 would make project names opaque on AWS: then
# this field has to follow).
_TRACKER_REL_RE = re.compile(r"(?i)^projects/([^/]+)/tracker/[^/]+\.xls[xm]$")
_UPLOAD_REL_RE = re.compile(r"(?i)^projects/([^/]+)/uploads/[^/]+$")     # a file directly in a project's Uploads/


class Mirror:
    def __init__(self, drive, ctx: Context, state: dict, save_vault, save_state, now=None):
        self.drive, self.ctx, self.state = drive, ctx, state
        self.save_vault, self.save_state = save_vault, save_state
        self.now = now or (lambda: datetime.now(timezone.utc))

    # ---- helpers --------------------------------------------------------------------------------
    def _tracker_entry(self, rel: str) -> dict:
        """{"kind": "tracker", "project": <the project folder as the AWS side names it>} for a held-back workbook
        directly in Projects/<project>/Tracker/; {"kind": "upload", "project": ...} for a file directly in its
        Uploads/ (the staff email says so on that project's payroll review: David, Nov 12, 2026); else {}. The name
        goes through _tp like every AWS-side path. A project folder whose own name holds a phone number, email or SSN
        is not on AWS at all: it gets no entry, never its name."""
        try:
            rel = rel.replace("\\", "/").lstrip("/")
            kind = "tracker" if _TRACKER_REL_RE.match(rel) else "upload" if _UPLOAD_REL_RE.match(rel) else None
            if kind is None:
                return {}
            project_dir = "/".join(rel.split("/")[:2])
            if screen.name_reason(project_dir, self.ctx.scrubber, True):
                return {}              # the folder's own name holds a phone number, email or SSN: it is never on
                                       # AWS (held back for that), so its name must not reach AWS in a summary either
            return {"kind": kind, "project": posixpath.basename(self._tp(project_dir, True))}
        except Exception:              # noqa: BLE001 - this is a hint for an email; when in doubt, say nothing
            return {}

    def _excluded(self, rel: str) -> bool:
        return screen.is_own_file(rel) or self.ctx.config.excluded(rel)

    def _tp(self, rel: str, folder: bool = False) -> str:
        """A staff path as it is on the AWS side: workers' names replaced by their tokens (screen.tokenize_path)."""
        return screen.tokenize_path(rel, self.ctx.scrubber, folder) if rel else rel

    def _copy_path(self, r: dict, name: str | None = None) -> str:
        """Where the copy of raw file ``r`` belongs on the AWS side, workers' names replaced by their tokens
        (``name``: the copy's own file name, if it differs). A clash is held back: see _path_problem."""
        rel = (posixpath.join(posixpath.dirname(r["rel"]), name) if name else _copy_rel(r)).lstrip("/")
        d, b = posixpath.split(rel)
        td, tb = self._tok_dir(d), self._aws_part(b)
        return posixpath.join(td, tb) if td else tb

    def _aws_part(self, part: str) -> str:
        """One staff file or folder name (not a project's own folder) as the AWS side names it: learnt workers' names
        as their tokens, and with vnext.opaque_file_names (#53) every other word not known as a FIL- token."""
        out = screen.tokenize_name(part, self.ctx.scrubber)
        return NM.aws_name(out, self.ctx.scrubber) if self.ctx.config.opaque_file_names else out

    def _aws_tp(self, rel: str) -> str:
        """An AWS-side path (a copy made on AWS keeps its name) with its file name as _aws_part would name it: the
        same as _tp, and with vnext.opaque_file_names its file name's unknown words as FIL- tokens too (#53; its
        folders are renamed by _recheck_folders)."""
        out = self._tp(rel)
        if not self.ctx.config.opaque_file_names:
            return out
        d, b = posixpath.split(out)
        b = NM.aws_name(b, self.ctx.scrubber)
        return posixpath.join(d, b) if d else b

    def _tok_dir(self, d: str) -> str:
        """The AWS-side name of staff folder ``d``."""
        if not d:
            return ""
        self._map_folders()
        hit = self._tokdir.get(d)
        if hit is None:
            parent, last = posixpath.split(d)
            hit = posixpath.join(self._tok_dir(parent), self._tok_part(d)) if parent else self._tok_part(d)
            self._tokdir[d] = hit
        return hit

    def _tok_part(self, d: str) -> str:
        parts = d.split("/")
        project = len(parts) == 2 and parts[0].casefold() == "projects"     # a project's own folder keeps its name
        return parts[-1] if project else self._aws_part(parts[-1])

    def _map_folders(self):
        """Staff folders and their AWS-side names, again whenever new names have been learnt. Two staff folders whose
        names become one ("Ana Rivera/", "ana_rivera/"): the one whose files are already there keeps it, the files
        of the other are held back (see _path_problem)."""
        if self._map_names == self.ctx.scrubber.name_count:
            return
        self._map_names = self.ctx.scrubber.name_count
        self._tokdir, self._dirmap, self._ambiguous, low = {}, {}, {}, {}
        settled, lives_at = {}, {}
        for q in self._st["pairs"]:
            if q.get("tok_rel") and q.get("rel") and not q.get("taken_back"):
                td, rd = posixpath.dirname(q["tok_rel"]), posixpath.dirname(q["rel"])
                while td and rd:                       # every level: a folder used only through its subfolders too
                    settled.setdefault(td.lower(), rd)
                    lives_at.setdefault(rd, td)
                    td, rd = posixpath.dirname(td), posixpath.dirname(rd)
        for d in sorted((d for d in self._raw_dirs if d), key=lambda d: (d.count("/"), d)):
            parent = posixpath.dirname(d)
            t = posixpath.join(self._tok_dir(parent), self._tok_part(d)) if parent else self._tok_part(d)
            here = lives_at.get(d)                     # its copies are already there under another spelling
            if here and here != t and here.lower() != t.lower() and posixpath.dirname(here) == posixpath.dirname(t) \
                    and not screen.name_left(here, self.ctx.scrubber, folder=True) and here.lower() not in low \
                    and (not self.ctx.config.opaque_file_names or known_names.is_plain(posixpath.basename(here))):
                t = here                               # (e.g. "EMP-…hours/" made on AWS; staff "Ana Rivera hours/")
            self._tokdir[d] = t
            other = low.get(t.lower())
            if other is None:
                low[t.lower()], self._dirmap[t] = d, d
            elif settled.get(t.lower()) == d:          # this one was there first: the other gives way
                self._ambiguous[other], low[t.lower()], self._dirmap[t] = d, d, d
            else:
                self._ambiguous[d] = other

    def _owner(self, path: str):
        """The staff file whose copy is at AWS-side ``path`` (any case: AWS's disk and rclone can't hold two names
        that differ only so), or None. An index of the copies, built once per mirror run and kept up to date by _own
        wherever a copy's AWS-side path changes; rebuilt after the pairs loop."""
        if self.__dict__.get("_owners") is None:       # built once, then kept up to date (_own)
            self._owners = {}
            for q in self._st["pairs"]:
                if not q.get("taken_back") and q.get("tok_rel"):
                    owner = q.get("raw_id") or q.get("tok_id")
                    self._owners.setdefault(q["tok_rel"].lower(), owner)
                    now = (self.__dict__.get("_tok", {}).get(q.get("tok_id")) or {}).get("rel")
                    if now and now.lower() != q["tok_rel"].lower():     # moved on AWS since: both places taken
                        self._owners.setdefault(now.lower(), owner)
        return self._owners.get(path.lower())

    def _own(self, old: str | None, new: str | None, owner) -> None:
        """A copy's AWS-side path changed (or appeared, or went): the owner index follows."""
        if self.__dict__.get("_owners") is None:
            return
        if old and self._owners.get(old.lower()) == owner:
            del self._owners[old.lower()]
        if new:
            self._owners[new.lower()] = owner

    def _target(self, r: dict, p: dict | None, name: str | None = None) -> str:
        """Where the copy of ``r`` is written: where it belongs, or, for a copy already there, its folder."""
        want = self._copy_path(r, name)
        if p and p.get("tok_rel") and p.get("tok_id"):
            base = self._aws_tp(posixpath.basename(p["tok_rel"])) if p.get("aws_named") else posixpath.basename(want)
            want = posixpath.join(posixpath.dirname(p["tok_rel"]), base)       # made on AWS: keeps AWS's name
        return want

    def _clash(self, r: dict, p: dict | None, path: str) -> str | None:
        owner = self._owner(path)
        if owner is None or owner == r["id"] or (p and owner == p.get("raw_id")):
            return None
        return ("another file has the same name on the AWS side once the worker's name is replaced by its token "
                "(for example 'Ana Rivera notes' and 'ana_rivera notes'); rename one of them")

    def _old_name_copy(self, p, unpaired_by_rel):
        """The copy of ``p`` an AWS run pushed again under the name it had before the tokenizer renamed it (the run
        pulled before the rename), within RECENT_HOURS of the rename; None otherwise."""
        old, when = p.get("prev_tok_rel"), p.get("prev_tok_utc") or ""
        try:
            made = datetime.strptime(when, "%Y-%m-%d-%H-%M-%S").replace(tzinfo=timezone.utc)
        except ValueError:
            return None
        if not old or (self.now() - made).total_seconds() > RECENT_HOURS * 3600:
            return None
        return unpaired_by_rel.pop(old, None)

    def _detok_name(self, text: str) -> str:
        """Workers' names put back in a file or folder name made on the AWS side ("EMP-3f2a91c0 notes" ->
        "Ana Rivera notes"; "EMP-3f2a91c0_hours" too). Only name tokens, and only names that would be replaced by a
        token again on the way out: an ID, SSN, date or a one-word name stays a token in the staff name, so the
        value never goes to AWS in a name. A name glued to other letters is set apart by spaces."""
        def put_back(m):
            tok = m.group(0)
            e = self.ctx.vault.entries.get(tok)
            if e and str(e[0]).split("+")[0] == NM.FILE_KIND:
                return str(e[2])                   # #53: words of a staff file or folder name, exactly as written
            if not e or e[0] not in self.ctx.scrubber.name_kinds:
                return tok
            name = str(self.ctx.vault.original(tok))
            before = " " if m.start() and text[m.start() - 1].isalnum() else ""
            after = " " if m.end() < len(text) and text[m.end()].isalnum() else ""
            again = screen.tokenize_name(before + name + after, self.ctx.scrubber)
            if again == before + name + after or screen.name_left(again, self.ctx.scrubber):
                return tok
            return before + name + after
        from .tokens import unglue
        return _NAME_TOKEN_RE.sub(put_back, unglue(text, self.ctx.vault))

    def _path_problem(self, r: dict, p: dict | None, name: str | None = None, target: str | None = None):
        """Why raw file ``r`` can't go to the AWS side under its name (names replaced by tokens), or None: a name
        that can't be replaced, its folder sharing one AWS-side name with another, or another file there."""
        left = screen.name_left(self._copy_path(r, name), self.ctx.scrubber)
        if left:
            return left
        self._map_folders()
        d = posixpath.dirname(r["rel"])
        while d:
            if d in self._ambiguous:
                return (f"its folder '{d}' has the same name on the AWS side as '{self._ambiguous[d]}' once the "
                        "worker's name is replaced by its token; merge the two folders or rename one")
            d = posixpath.dirname(d)
        return self._clash(r, p, target or self._target(r, p, name))

    def _raw_dir(self, tok_dir: str) -> str:
        """The staff folder of an AWS-side folder: the one it was made from, else its tokens put back."""
        if not tok_dir:
            return ""
        self._map_folders()
        hit = self._dirmap.get(tok_dir)
        if hit is None:
            parent, last = posixpath.split(tok_dir)
            parts = tok_dir.split("/")
            if not (len(parts) == 2 and parts[0].casefold() == "projects"):   # a project's own folder goes to AWS
                last = self._detok_name(last)                                # as it is: a token there stays one
            hit = posixpath.join(self._raw_dir(parent), last) if parent else last
            self._dirmap[tok_dir] = hit
            self._tokdir.setdefault(hit, tok_dir)      # and back: the staff folder made from it maps to it
        return hit

    def _folder(self, folders: dict, root_rel_dir: str) -> str:
        """ID of folder ``root_rel_dir`` on one side, creating any missing folders."""
        if root_rel_dir in folders:
            return folders[root_rel_dir]
        parent = posixpath.dirname(root_rel_dir)
        pid = self._folder(folders, parent)
        folders[root_rel_dir] = self.drive.mkdir(pid, posixpath.basename(root_rel_dir))
        return folders[root_rel_dir]

    def _tokenize(self, raw_f: dict, data: bytes | None = None):
        """(name, bytes, detail); raises TokenError or screen.Held when the file must not go to AWS; (None, None,
        reason) when the file rules leave it out. ``data``: the file's bytes when already downloaded."""
        mime = raw_f.get("mimeType", "")
        if mime.startswith(GOOGLE_NATIVE) and mime != GSHEET_MIME:
            return None, None, "Google Docs/Slides/Forms files are not copied"
        sheet = mime == GSHEET_MIME
        if raw_f.get("rel") == settings_file.REL and not sheet:
            # the administrator's settings (issue #151): not tokenized (its addresses are the firm's, chosen on the
            # Settings page) but rebuilt from its checked values only; anything else in it holds the file back
            try:
                return raw_f["name"], settings_file.for_aws(self.drive.download(raw_f) if data is None else data), \
                    "settings checked: known settings only, firm addresses; history left out"
            except settings_file.Invalid:        # never the reason's text: it can name an address
                raise screen.Held("the settings file holds something the Settings page doesn't write") from None
        how = "tokenize" if sheet else self.ctx.config.handling(raw_f["name"])
        if how == "skip":
            return tokenize_file(raw_f["name"], b"", self.ctx)          # the file rules' reason
        _progress(raw_f)
        data = self.drive.download(raw_f) if data is None else data
        if how == "scan_pdf":
            screen.screen_pdf(raw_f["rel"], data, self.ctx.scrubber)
            return raw_f["name"], data, "PDF checked: no worker data, copied as it is"
        name, out, detail = tokenize_file(raw_f["name"], data, self.ctx, google_sheet=sheet, rel=raw_f.get("rel"))
        if out is not None:
            why = screen.residual(name, out, self.ctx.scrubber)
            if why:
                raise screen.Held(why)
        return name, out, detail

    def _held(self, r: dict, reason: str, rep) -> None:
        """Record that raw file ``r`` was held back (listed in the staff side's Held back from AWS.csv)."""
        hb = self._st["held_back"]
        old = hb.get(r["id"])
        since = old["since"] if old and old.get("ver") == version(r) else self.now().strftime("%Y-%m-%d %H:%M")
        hb[r["id"]] = {"rel": r["rel"], "reason": reason, "ver": version(r), "since": since}
        rep.refused.append((r["rel"], reason))

    # ---- one mirror -----------------------------------------------------------------------------
    def run(self, name: str, raw_root: str, tok_root: str) -> RunReport:
        rep = RunReport()
        st = self.state.setdefault("mirrors", {}).setdefault(name, {})
        self._owners = None                            # per mirror: one Mirror runs them all (run_all)
        for k in ("pairs", "retired", "not_copied", "retired_pair", "held_back"):
            st.setdefault(k, [] if k == "pairs" else {})
        self._st = st
        self._taken, self._take_limit, self._cap_warned = 0, max(TAKE_BACK_MIN, len(st["pairs"]) // 10), False
        raw, raw_dirs = self.drive.tree(raw_root)
        tok, tok_dirs = self.drive.tree(tok_root)
        own = {"raw": {f["rel"]: f for f in raw.values() if f["rel"] in screen.OWN_FILES},
               "tok": {f["rel"]: f for f in tok.values() if f["rel"] in screen.OWN_FILES}}
        theirs = [f for f in tok.values() if f["rel"] not in screen.OWN_FILES]
        if st.get("roots") and st["roots"] != [raw_root, tok_root]:
            raise RuntimeError(f"[{name}] the raw or tokenized folder ID differs from the one this mirror was started "
                               "with; nothing was changed. To start over with new folders, remove this mirror's "
                               "state first.")
        if not st.get("started") and not st["pairs"] and theirs:
            raise RuntimeError(f"[{name}] the tokenized folder must be empty the first time it is used "
                               f"(it holds {len(theirs)} files); nothing was changed.")
        # all copies may be taken back later: that doesn't make it a first run again
        st["started"], st["roots"] = True, [raw_root, tok_root]
        stamp = self.now().strftime("%Y-%m-%d-%H-%M-%S")
        if self._convert_macros(raw, raw_dirs, tok, st, stamp, rep):
            raw, raw_dirs = self.drive.tree(raw_root)       # see the converted files and the moved originals
        if self._sync_staff_trackers(name, raw, raw_dirs, tok, st, stamp, own, raw_root, tok_root, rep):
            raw, raw_dirs = self.drive.tree(raw_root)       # see the trackers just copied from the staff drive
        self._load_status_words(raw, rep)
        raw = {i: f for i, f in raw.items() if not self._excluded(f["rel"])}
        tok_all = tok                                  # unfiltered: a take-back still pending under an excluded path
        tok = {i: f for i, f in tok.items() if not self._excluded(f["rel"])}
        # Pairs under a path excluded since they were made are forgotten here, before the "looks deleted" check:
        # their files are filtered out of both sides just above, so they would otherwise count as deleted on both
        # and could stop every run. Nothing is moved or binned on either side; the copies stay where they are.
        before = len(st["pairs"])
        # A copy the tokenizer had decided to take back from AWS but hadn't yet (pending_take_back) keeps its pair,
        # so the take-back still finishes (the copy binned, its content put on the purge list).
        st["pairs"] = [q for q in st["pairs"] if q.get("pending_take_back") or not (
            self._excluded(q.get("rel") or "") or self._excluded(q.get("tok_rel") or ""))]
        if len(st["pairs"]) < before:
            rep.warnings.append(f"{before - len(st['pairs'])} file(s) in folders the tokenizer now leaves out "
                                "(tokenizer_config.yaml, exclude) are no longer kept in step; their copies were left "
                                "where they were")
        self._tok = tok
        self._raw_dirs, self._map_names = raw_dirs, None
        self._map_folders()
        tb = st.setdefault("taken_back_md5", {})      # contents of copies taken back -> the staff file, where it was
        if not st.get("names_tokenized"):
            # Once, after the update that replaces names in file names by tokens: files held back only for a
            # worker's name in their name are checked (and copied) again.
            for rid, h in list(st["held_back"].items()):
                why = h.get("reason", "")
                if "name contains a worker's name;" in why or "name is a worker's name" in why:
                    del st["held_back"][rid]
                    st["not_copied"].pop(rid, None)
                    for k in [k for k, e in tb.items() if isinstance(e, dict) and e.get("raw") == rid]:
                        del tb[k]
            st["names_tokenized"] = True
        for rid, h in st["held_back"].items():
            if "same name on the AWS side" in h.get("reason", ""):
                st["not_copied"].pop(rid, None)        # cheap to check again (no download): see _path_problem
        live = {p["raw_id"] for p in st["pairs"] if not p.get("taken_back") and not p.get("pending_take_back")}
        for k in [k for k, e in tb.items() if not isinstance(e, dict) or e["raw"] in live or e["raw"] not in raw]:
            del tb[k]                                  # copied again (released) or deleted by staff: forgotten

        gone_raw = sum(1 for p in st["pairs"] if p["raw_id"] not in raw)
        gone_tok = sum(1 for p in st["pairs"] if p["tok_id"] not in tok)
        n = len(st["pairs"])
        if n > 5 and max(gone_raw, gone_tok) > n / 2:
            raise RuntimeError(f"[{name}] {max(gone_raw, gone_tok)} of {n} files look deleted on one side. "
                               "This usually means a wrong folder ID or a removed share; nothing was changed.")

        # Copies a stopped run was writing (see _run_forward): ones it made are binned (their staff files are not
        # paired, so they are copied again below); ones it updated are its own, not a change made on AWS.
        inflight = st.pop("in_flight", None) or {}
        if inflight:
            known = {p["tok_id"] for p in st["pairs"]}
            for p in st["pairs"]:
                if p["raw_id"] in inflight and p["tok_id"] in tok:
                    p["tok_ver"] = version(tok[p["tok_id"]])
            want = set(inflight.values())
            for i, f in list(tok.items()):
                if i not in known and f["rel"] in want:
                    try:
                        self.drive.trash(i)
                        del tok[i]
                    except Exception as e:             # noqa: BLE001 - kept: it would be restored as a staff file
                        st["in_flight"] = inflight
                        # the Drive ID, not the path: this message is printed to Cloud Logging (#21)
                        raise RuntimeError(f"[{name}] tokenized file {i}: a copy left by a stopped run could not be "
                                           f"removed ({type(e).__name__}); nothing else was changed") from e
            rep.warnings.append(f"the previous run stopped while writing {len(inflight)} copies; they are made again")

        paired_raw = {p["raw_id"] for p in st["pairs"]}
        paired_tok = {p["tok_id"] for p in st["pairs"]}
        unpaired_tok_by_rel = {f["rel"]: f for f in tok.values() if f["id"] not in paired_tok}
        unpaired_tok_by_ver = {}          # to recognise a move that arrived as "moved aside + uploaded elsewhere"
        for f in tok.values():
            if f["id"] not in paired_tok and not f["rel"].startswith(BACKUP_DIR + "/"):
                unpaired_tok_by_ver.setdefault(version(f), []).append(f)
        forward_jobs = []          # (pair or None, raw file, existing tok id or None)
        keep = []

        for p in st["pairs"]:
            if p.get("taken_back"):            # its copy was taken back by an earlier run: nothing left to follow
                continue
            r, t = raw.get(p["raw_id"]), tok.get(p["tok_id"])
            if p.get("pending_take_back"):     # the copy step held the staff file back but couldn't bin the old copy
                t = t or tok_all.get(p["tok_id"])              # under a path excluded since: still finished
                if t is None:
                    continue                   # gone already
                if r is not None and version(r) != st["not_copied"].get(r["id"]):
                    # staff changed it again since: bin the old copy, and check and copy the new version as usual
                    try:
                        kind = self._kind_of_old(t["id"], p.get("tok_rel"))   # (before the bin: it is read)
                        self.drive.trash(t["id"])
                    except Exception as e:     # noqa: BLE001
                        rep.warnings.append(f"{r['rel']}: its previous copy could not be taken back from the AWS "
                                            f"side this run ({e})")
                        keep.append(p)
                        continue
                    self._stale_copy(kind, t.get("md5Checksum"), p)   # only if a name caused it (_kind_of_old)
                    rep.removed.append((r["rel"], "previous tokenized copy moved to the bin"))
                    st["not_copied"].pop(r["id"], None)
                    st["held_back"].pop(r["id"], None)
                    paired_raw.discard(r["id"])
                    self.save_state()
                    continue
                if r is None:
                    r = {"id": p["raw_id"], "rel": p["rel"], "md5Checksum": None}
                if not self._take_back(p, r, t["id"], p["pending_take_back"], rep):
                    keep.append(p)             # tried again next run; nothing else is done with it meanwhile
                continue
            if t is not None and t["rel"].startswith(BACKUP_DIR + "/") and \
                    not p["tok_rel"].startswith(BACKUP_DIR + "/"):
                # rclone --backup-dir: AWS moved the old copy aside and (when it replaced the file) uploaded a
                # new one at the same path. That is a change (or a deletion), not a move to follow; the copy
                # left under _sync_backup/ stays on the AWS side only.
                t = unpaired_tok_by_rel.pop(p["tok_rel"], None) or self._old_name_copy(p, unpaired_tok_by_rel)
                if t is not None:
                    paired_tok.add(t["id"])
                    self._own(p["tok_rel"], t["rel"], p.get("raw_id"))
                    p["tok_id"], p["tok_ver"], p["tok_rel"] = t["id"], None, t["rel"]   # None: counts as changed
                else:
                    # Nothing new at the old path, but the same bytes were uploaded elsewhere: AWS moved the file
                    # (an rclone sync without --track-renames). Follow it as a move, so the staff original is
                    # moved too instead of being set aside and replaced by a detokenized copy.
                    same = [f for f in unpaired_tok_by_ver.get(p["tok_ver"], []) if f["id"] not in paired_tok]
                    same.sort(key=lambda f: posixpath.basename(f["rel"]) != posixpath.basename(p["tok_rel"]))
                    if same:
                        t = same[0]
                        paired_tok.add(t["id"])
                        unpaired_tok_by_rel.pop(t["rel"], None)
                        p["tok_id"] = t["id"]
            if t is None:                              # gone: maybe pushed again under the name it had before a rename
                t = self._old_name_copy(p, unpaired_tok_by_rel)
                if t is not None:
                    paired_tok.add(t["id"])
                    self._own(p["tok_rel"], t["rel"], p.get("raw_id"))
                    p["tok_id"], p["tok_ver"], p["tok_rel"] = t["id"], None, t["rel"]
            if r is None and t is None:
                continue
            r_changed = r is not None and version(r) != p["raw_ver"]
            t_changed = t is not None and version(t) != p["tok_ver"]
            if t_changed and r is not None and version(t) in p.get("tok_hist", [])[:-1]:
                # an older copy this tokenizer wrote itself, pushed back by an AWS run that pulled before the latest
                # one: not a change made on AWS (taking it would undo a staff edit). The current version goes again.
                t_changed, r_changed = False, True
                rep.warnings.append(f"{r['rel']}: an AWS run pushed back an older copy; the current version was sent "
                                    "again")
            if settings_file.REL in (p.get("rel"), (r or {}).get("rel")):
                # The Settings page's file (#151; review F1, N1-N3): nothing done to its copy on the AWS side is ever
                # taken (no move, rename, change or deletion reaches the staff folder); the staff file decides.
                if r is None or r["rel"] != settings_file.REL or p.get("rel") != settings_file.REL:
                    # staff deleted it, moved it away (then it's an ordinary .json, not copied), or moved another
                    # file in (then it is checked as the settings file, from scratch): the old copy goes
                    if t is not None and not self._bin_settings_copy(t, rep):
                        keep.append(p)                    # tried again next run
                        continue
                    if t is not None:
                        rep.removed.append((t["rel"], "the old copy at or of the settings file: moved to the bin"))
                    if r is not None:
                        paired_raw.discard(r["id"])       # looked at again below as a new staff file
                    continue
                if t is None or t["rel"] != settings_file.REL:
                    # AWS deleted, moved or renamed the copy: the moved one goes to the bin, a new copy is made
                    if t is not None and not self._bin_settings_copy(t, rep):
                        keep.append(p)
                        continue
                    rep.warnings.append(f"{settings_file.REL}: its copy was deleted, moved or renamed on the AWS "
                                        "side; nothing was changed in the staff folder, and it was copied again")
                    forward_jobs.append((p, r, None))
                elif t_changed:
                    rep.warnings.append(f"{settings_file.REL}: changed on the AWS side; not copied back to the staff "
                                        "folder, whose file is sent again")
                    forward_jobs.append((p, r, t["id"]))
                elif r_changed:
                    forward_jobs.append((p, r, t["id"]))
                keep.append(p)
                continue
            if r is None:
                if t_changed:                       # staff deleted it, but AWS has since changed it: bring it back
                    self._back(t, None, p, raw_dirs, rep)
                    keep.append(p)
                else:
                    self.drive.trash(t["id"])
                    rep.removed.append((t["rel"], "deleted in the raw folder, so moved to the bin in the copy"))
                continue
            if t is None:
                if r_changed:
                    forward_jobs.append((p, r, None))
                    keep.append(p)
                elif self._recent(p):
                    # copied only hours ago and gone already: most likely an AWS run that pulled before the copy
                    # existed pushed its older view back. The agent moves files, it doesn't delete them, so send
                    # the file again rather than take the staff file away.
                    forward_jobs.append((p, r, None))
                    keep.append(p)
                    rep.warnings.append(f"{r['rel']}: its copy disappeared on the AWS side within {RECENT_HOURS} h of "
                                        "being made (probably a run already in progress); copied again")
                else:
                    dest = f"{BACKUP_DIR}/{stamp}__deleted_on_aws/{posixpath.dirname(r['rel'])}".rstrip("/")
                    self.drive.move(r["id"], self._folder(raw_dirs, dest), r["parents"][0])
                    rep.removed.append((r["rel"], f"deleted on the AWS side; raw file moved to {dest}/"))
                continue
            # Both present: follow moves and renames first.
            if r["rel"] != p["rel"]:                       # staff moved or renamed it (wins over AWS)
                why = screen.name_reason(_copy_rel(r), self.ctx.scrubber) or \
                    self._path_problem(r, p, target=self._copy_path(r))
                if why:                                    # ... to a name that can't go to AWS: take the copy back
                    if not self._take_back(p, r, t["id"], why, rep):
                        keep.append(p)                     # not this run: the copy keeps its old name meanwhile
                    continue
                self._follow(t, tok_dirs, self._copy_path(r), p, rep, "the raw folder")
                p["rel"] = r["rel"]
                p.pop("aws_named", None)                   # renamed by staff: named from the staff side now
            elif t["rel"] != p["tok_rel"]:                 # AWS moved or renamed it
                td, tn = posixpath.split(t["rel"])
                if tn == posixpath.basename(self._copy_path(r)):
                    rn = r["name"]                         # moved only: the staff file keeps its own name
                else:
                    rn = self._detok_name(posixpath.basename(_raw_rel({"rel": tn}, r)))
                rd = self._raw_dir(td)
                new = posixpath.join(rd, rn) if rd else rn
                self._follow(r, raw_dirs, new, p, rep, "the copy")
                self._own(p["tok_rel"], t["rel"], p.get("raw_id"))
                p["rel"], p["tok_rel"] = new, t["rel"]
            if r_changed and t_changed:
                self._back(t, None, p, raw_dirs, rep, conflict_stamp=stamp)
                forward_jobs.append((p, r, t["id"]))
            elif r_changed:
                forward_jobs.append((p, r, t["id"]))
            elif t_changed:
                self._back(t, r, p, raw_dirs, rep)
            keep.append(p)
        st["pairs"] = keep
        self._owners = None                            # rebuilt from the pairs as they are now

        # New files on either side.
        for f in sorted(tok.values(), key=lambda f: f["rel"]):
            if f["id"] not in paired_tok and not f["rel"].startswith(BACKUP_DIR + "/"):
                if self._was_taken_back(f) or (screen.has_name(f["rel"], self.ctx.scrubber) and any(
                        q.get("tok_ver") == version(f) and (q.get("tok_rel") == self._tp(f["rel"]) or
                                                             q.get("prev_tok_rel") == f["rel"])
                        for q in st["pairs"])):
                    # an AWS run pulled it before it was taken back and has pushed it again: not restored (that would
                    # make a duplicate in the staff folder); removed from the AWS side again
                    try:
                        self.drive.trash(f["id"])
                        if screen.has_name(f["rel"], self.ctx.scrubber):   # pushed again under a name: its path
                            self._stale("path", f.get("md5Checksum"))
                        else:                          # (listed already, if it should be, when taken back: its date)
                            self._refresh(f.get("md5Checksum"))
                        rep.removed.append((f["rel"], "held back, but uploaded again from AWS: moved to the bin there"))
                        rep.warnings.append(f"{f['rel']}: uploaded from AWS with the name or content of a file held "
                                            "back; moved to the bin on the AWS side, not copied to the staff folder")
                    except Exception as e:             # noqa: BLE001
                        rep.warnings.append(f"{f['rel']}: held back, uploaded again from AWS, not removed ({e})")
                    continue
                if f["rel"] == settings_file.REL or f["rel"].startswith(settings_file.REL.split("/")[0] + "/"):
                    # the settings folder: only the staff side has its files (#151; focused review F2)
                    try:
                        self.drive.trash(f["id"])
                        rep.warnings.append(f"{settings_file.REL.split('/')[0]}/: a file found on the AWS side only; "
                                            "moved to the bin there, never copied to the staff folder")
                    except Exception as e:             # noqa: BLE001
                        rep.warnings.append(f"{settings_file.REL}: found on the AWS side only, not removed ({e})")
                    continue
                p = {"raw_id": None, "tok_id": f["id"], "rel": None, "tok_rel": f["rel"], "raw_ver": None,
                     "tok_ver": None}
                st["pairs"].append(p)
                target = posixpath.join(self._raw_dir(posixpath.dirname(f["rel"])),
                                        self._detok_name(f["name"])).lstrip("/").lower()
                clash = any(g["rel"].lower() == target for g in raw.values())   # e.g. a copy pushed under an old name
                if not self._back(f, None, p, raw_dirs, rep, conflict_stamp=stamp if clash else None, new=True):
                    st["pairs"].remove(p)
        for f in sorted(raw.values(), key=lambda f: f["rel"]):
            if f["id"] in paired_raw or f["rel"].startswith(BACKUP_DIR + "/"):
                continue
            if f["id"] in st["retired"]:
                if version(f) == st["retired"][f["id"]]:
                    continue
                pair = next((p for p in st["pairs"] if p["tok_id"] == st["retired_pair"].get(f["id"])), None)
                if pair is None:
                    continue           # Google-side-only file (an "AWS copy" saved in a conflict): never copied
                # Staff edited a Google Sheet whose readable copy had been saved alongside as .xlsx: the Sheet
                # becomes the source again, over the same tokenized copy; the .xlsx stays in Drive only.
                old = raw.get(pair["raw_id"])
                if old is not None:
                    st["retired"][old["id"]] = version(old)
                del st["retired"][f["id"]]
                st["retired_pair"].pop(f["id"], None)
                forward_jobs.append((pair, f, pair["tok_id"]))
                continue
            hb = st["held_back"].get(f["id"])
            renamed = hb is not None and hb.get("rel") != f["rel"]       # tried again after a rename or a move too
            if (version(f) == st["not_copied"].get(f["id"]) and not renamed) or \
                    any(p["raw_id"] == f["id"] for p in st["pairs"]):
                continue
            forward_jobs.append((None, f, None))

        # Names are learnt from the name columns of payroll files, and free text elsewhere (a tracker, a log)
        # is only scrubbed for names already learnt. So payroll files go first, then other spreadsheets, then
        # text; otherwise a project's Tracker/ (alphabetically before Upload/) would pass with workers' names.
        forward_jobs.sort(key=lambda j: _forward_order(j[1]["rel"], self.ctx.config))
        self._recheck_folders(st, raw, tok, tok_dirs, rep)     # names known already: before copies go into them
        self._run_forward(forward_jobs, st, tok, tok_dirs, rep)
        self._recheck_folders(st, raw, tok, tok_dirs, rep)
        self._recheck(st, raw, tok, tok_dirs, rep)
        self._clean_backups(tok, tok_dirs, rep)
        # Folders: after the files, so that the names this run's payrolls taught are known (a folder named after a
        # worker is never made on the AWS side).
        # Make sure each side has the other's folder structure (empty folders included).
        # A folder whose name differs only in case from one already there (e.g. the old "projects" after it was
        # renamed "Projects") is not created: Macs and Windows treat the two as one folder.
        tok_low, raw_low = {r.lower() for r in tok_dirs}, {r.lower() for r in raw_dirs}
        # nor a folder of the agent's previous layout in a project already in the current one (_tidy_layout)
        migrated = _migrated_projects(set(raw_dirs) | set(tok_dirs))
        old = lambda rel: _in_old_layout(rel, migrated)
        losing = lambda rel: any(rel == a or rel.startswith(a + "/") for a in self._ambiguous)
        as_tok = {self._tok_dir(rel): rel for rel in raw_dirs if rel and not losing(rel)}
        for t_rel, rel in sorted(as_tok.items()):
            if t_rel not in tok_dirs and t_rel.lower() not in tok_low and not old(rel) \
                    and not self._excluded(rel + "/") and not rel.startswith(BACKUP_DIR) \
                    and not screen.name_reason(rel, self.ctx.scrubber, folder=True) \
                    and not screen.name_left(t_rel, self.ctx.scrubber, folder=True):
                self._folder(tok_dirs, t_rel)
        for rel in sorted(set(tok_dirs) - set(as_tok)):
            r_rel = self._raw_dir(rel) if rel else rel
            if rel and r_rel not in raw_dirs and r_rel.lower() not in raw_low and not old(rel) \
                    and not self._excluded(rel + "/") and not rel.startswith(BACKUP_DIR) \
                    and not screen.has_name(rel, self.ctx.scrubber, folder=True):
                self._folder(raw_dirs, r_rel)

        self._tidy_layout(raw_root, tok_root, rep)
        self._publish_held_back(st, raw, raw_root, tok_root, own, rep)
        self._publish_purge(st, tok_root, own, rep)
        self._publish_links(st, tok_root, own, rep)
        return rep

    def _publish_links(self, st, tok_root, own, rep):
        """AWS side: {the copy's path there: the Drive ID of the staff file it stands for}, for the outcome email's
        links (agent/notify.py Linker): they open the staff's own file, never the copy AWS works on. Drive IDs only:
        an ID names nothing and grants nothing (opening it needs access to the staff drive). Written after the
        restore, so this run's results are in it."""
        links = {p["tok_rel"]: p["raw_id"] for p in st["pairs"]
                 if p.get("raw_id") and p.get("tok_rel") and not p.get("taken_back") and not p.get("pending_take_back")
                 and not self._excluded(p.get("rel") or "")}
        data = json.dumps({"links": dict(sorted(links.items())),
                           "note": "each AWS-side copy's staff file, by Drive ID"}, indent=1).encode()
        self._write_own(own["tok"].get(screen.STAFF_LINKS), tok_root, screen.STAFF_LINKS, data, True, rep)

    # ---- held back -------------------------------------------------------------------------------
    def _load_status_words(self, raw, rep):
        """Strict trackers (#128): the tracker Status words kept readable this pass are tokenizer_config.yaml's and
        the Settings page's four tracker status lists (read from the staff side's settings file, checked as the page
        checks them). Nothing else is read from it here. Without the setting, or if the file can't be read: the
        config's words only (a Status that is only on the page's list then reaches the agent as a token)."""
        self.ctx.status_words = frozenset(self.ctx.config.tracker_status_words)
        if not self.ctx.config.strict_trackers:
            return
        f = next((f for f in raw.values() if f.get("rel") == settings_file.REL
                  and not f.get("mimeType", "").startswith(GOOGLE_NATIVE)), None)
        if f is None:
            return
        try:
            values, _ = settings_file.read(self.drive.download(f))
        except Exception as e:                         # noqa: BLE001 - never the reason's text: it can name an address
            rep.warnings.append(f"{settings_file.REL}: its tracker status words couldn't be read this run "
                                f"({type(e).__name__}); the tokenizer's own list is used")
            return
        words = set(self.ctx.status_words)
        for key in ("intent.not_started_statuses", "intent.work_started_statuses", "intent.completed_statuses",
                    "intent.skip_statuses"):
            for w in values.get(key) or []:
                if isinstance(w, str) and w.strip():
                    words.add(" ".join(w.split()).lower())
        self.ctx.status_words = frozenset(words)

    def _recheck(self, st, raw, tok, tok_dirs, rep):
        """Copies already on the AWS side, checked again for names learnt since: a worker's name in a copy's file
        name is replaced by its token there (the copy is renamed); a PDF that now holds one is taken back."""
        count = self.ctx.scrubber.name_count
        changed, start, later = False, time.monotonic(), 0
        for p in list(st["pairs"]):                    # names: every run (no download, only the path)
            if p.get("taken_back") or p.get("pending_take_back") or not p.get("raw_id") or not p.get("tok_rel"):
                continue
            r = raw.get(p["raw_id"])
            if r is None:
                continue
            want = self._aws_tp(p["tok_rel"]) if p.get("aws_named") else self._copy_path(r)
            left = screen.name_left(want, self.ctx.scrubber)
            if left:
                if self._take_back(p, r, p["tok_id"], left, rep):      # a PDF: its text is checked too
                    st["pairs"].remove(p)
                continue
            if want.lower() == p["tok_rel"].lower():
                continue
            clash = self._path_problem(r, p, target=want)
            if clash:                                  # renamed, it would share a name: taken back, listed
                if self._take_back(p, r, p["tok_id"], clash, rep):     # listed only if a name is in the copy
                    st["pairs"].remove(p)
                continue
            t = tok.get(p["tok_id"])
            try:
                t = dict(t) if t is not None else {**self.drive.get(p["tok_id"])}
                t["rel"] = p["tok_rel"]
                self._follow(t, tok_dirs, want, p, rep, "the raw folder", why="a worker's name replaced by its token")
                self._stale("path", t.get("md5Checksum"), *self._hist(p))   # the old path held a name
                changed = True
            except Exception as e:                     # noqa: BLE001 - tried again next run
                rep.warnings.append(f"{p.get('rel')}: its copy on the AWS side could not be renamed this run ({e})")
        due = [p for p in st["pairs"] if p.get("names_at_check", 0) < count and p.get("raw_id") and p.get("tok_id")
               and not p.get("pending_take_back")]
        redo = []                                      # #54: copies to tokenize again from the staff file
        for p in sorted(due, key=lambda p: p.get("names_at_check", 0)):  # least recently checked first: none starves
            r, t = raw.get(p["raw_id"]), tok.get(p["tok_id"])
            if r is not None and t is None:            # copied in this run: not in the snapshot taken at its start
                try:
                    t = self.drive.get(p["tok_id"])
                except Exception:                      # noqa: BLE001
                    t = None
            if r is None or t is None:
                continue
            kind, read = None, False               # (None: _take_back decides, see _kind_of_old)
            try:
                why = screen.name_reason(p["tok_rel"], self.ctx.scrubber)
                content = self._recheck_content(p, r, t, start, redo, rep) if why is None else None
                if content == "later":
                    later += 1                         # its names_at_check stays behind: checked next run
                    continue
                if content == "found":
                    # its names_at_check stays behind until a copy made with the names known now is on AWS (set by
                    # _run_forward when it writes one): a re-copy that fails, waits or is capped is tried again
                    changed = True
                    continue
                if why is None and self.ctx.config.handling(r["name"]) == "scan_pdf":
                    if time.monotonic() - start > RECHECK_SECONDS:
                        later += 1                     # its names_at_check stays behind: checked next run
                        continue
                    _progress(t)
                    read = True
                    screen.screen_pdf(r["rel"], self.drive.download(t), self.ctx.scrubber)
            except screen.Found as e:
                why, kind = str(e), "content"              # a name (or SSN...) in the text on AWS
            except screen.Held as e:
                why = str(e)                           # can't be checked: the copy is no evidence of a name
            except Exception as e:                     # noqa: BLE001 - e.g. a download failed; next run
                rep.warnings.append(f"{r['rel']}: not checked again this run ({e})")
                continue
            if why:
                if self._take_back(p, r, t["id"], why, rep, kind=kind, read=not read):
                    st["pairs"].remove(p)
            else:
                p["names_at_check"] = count
            changed = True
        if redo:
            # #54: a copy that holds a name learnt since it was made goes again, tokenized with the names known now
            # (its old content is listed for purging from S3 already: _recheck_content)
            redo.sort(key=lambda j: _forward_order(j[1]["rel"], self.ctx.config))
            self._run_forward(redo, st, tok, tok_dirs, rep)
            changed = True
        if later:
            rep.warnings.append(f"{later} file(s) on the AWS side not yet checked again for newly learnt names; "
                                "the next run continues")
        if changed:
            self.save_state()

    def _recheck_content(self, p, r, t, start, redo, rep) -> str | None:
        """#54 (vnext.recheck_spreadsheets): a spreadsheet or text copy already on AWS, read again for names learnt
        since it was made, as PDFs are. One that now holds a known name (or an SSN, email, phone number...) is listed
        for purging from S3 (its content, and what this tokenizer wrote before at its path: kind content, every old S3
        version with that content under any key) and tokenized again from the staff file (added to ``redo``), when
        AWS hasn't changed it since (else it comes back to the staff side first, and is checked again next run), at
        most a tenth of the copies per run (as take-backs: a wrongly learnt name can't rewrite everything at once).
        "found" then (its names_at_check must stay behind), "later" when the time budget is spent, else None."""
        cfg = self.ctx.config
        if not cfg.recheck_spreadsheets or settings_file.REL in (p.get("rel"), r.get("rel")):
            return None
        sheet = r.get("mimeType") == GSHEET_MIME
        if not sheet and cfg.handling(r["name"]) not in ("tokenize", "scrub_text"):
            return None
        if time.monotonic() - start > RECHECK_SECONDS:
            return "later"
        _progress(t)
        why = screen.residual(t.get("name") or posixpath.basename(p["tok_rel"]), self.drive.download(t),
                              self.ctx.scrubber)
        if not why:
            return None
        self._stale("content", t.get("md5Checksum"), *[m for m in self._hist(p) if m])
        if version(t) != p.get("tok_ver"):
            rep.warnings.append(f"{r['rel']}: its copy on the AWS side holds a worker's name learnt since it was made, "
                                "and was changed on AWS since: it comes back first, then is tokenized again; its old "
                                "content is listed for purging")
        elif any(j[0] is p for j in redo):
            pass
        elif len(redo) >= self._take_limit:
            rep.warnings.append(f"{r['rel']}: its copy on the AWS side holds a worker's name learnt since it was made; "
                                "not tokenized again this run (at most a tenth of the copies per run), next run")
        else:
            try:                                    # the staff file as it is now: a run's _back may have updated it
                fresh = {**r, **self.drive.get(r["id"])}
            except Exception as e:                  # noqa: BLE001 - next run
                rep.warnings.append(f"{r['rel']}: not tokenized again this run, will retry ({type(e).__name__})")
                return "found"
            redo.append((p, fresh, t["id"]))
            rep.warnings.append(f"{r['rel']}: its copy on the AWS side held a worker's name learnt since it was made; "
                                "tokenized again, and its old content is listed for purging")
        return "found"

    def _was_taken_back(self, f) -> bool:
        """A new file on the AWS side with the content of a copy taken back, in the same folder or under the same
        name: an AWS run that pulled the copy before then pushing it again. Both are needed, so that a common
        content (an empty file, a blank template) elsewhere is never mistaken for it."""
        e = self._st["taken_back_md5"].get(f.get("md5Checksum"))
        return isinstance(e, dict) and (posixpath.dirname(f["rel"]) == e["dir"] or
                                        posixpath.basename(f["rel"]) == e["name"])

    def _md5_of(self, tok_id):
        """The Drive MD5 of an AWS-side file (None if it has none or can't be read)."""
        f = self._tok.get(tok_id)
        if f is None:                                  # made in this run: not in the snapshot taken at its start
            try:
                f = self.drive.get(tok_id)
            except Exception:                          # noqa: BLE001
                f = None
        return (f or {}).get("md5Checksum")

    def _remember_taken(self, tok_id, raw_id, tok_rel):
        """Keep the content (md5) and place of a copy taken back from the AWS side (see _was_taken_back)."""
        md5 = self._md5_of(tok_id)
        if md5 and tok_rel:
            tb = self._st["taken_back_md5"]
            tb[md5] = {"raw": raw_id, "dir": posixpath.dirname(tok_rel), "name": posixpath.basename(tok_rel)}
            while len(tb) > 5000:                      # oldest first
                del tb[next(iter(tb))]

    def _take_back(self, p, r, tok_id, why, rep, kind=None, read=True) -> bool:
        """Move a copy to the bin on the AWS side and hold the staff file back. Saved at once, so a run that stops
        later never mistakes the missing copy for one deleted on AWS (which would move the staff file aside).
        At most a tenth of the copies per run (at least TAKE_BACK_MIN): more points to a wrong name, not real work.
        Listed for purging on AWS as ``kind``, or as _kind_of_old finds (``read``: a PDF's text is checked)."""
        if self._taken >= self._take_limit:
            if not self._cap_warned:
                rep.warnings.append(f"more than {self._take_limit} copies would be taken back from the AWS side in one "
                                    "run; the rest wait for the next run. Check the names in Held back from AWS.csv")
                self._cap_warned = True
            return False
        try:
            if kind is None:
                kind = self._kind_of_old(tok_id, p.get("tok_rel"), read=read)   # (before the bin: it is read)
            self.drive.trash(tok_id)
        except Exception as e:                         # noqa: BLE001 - try again next run
            rep.warnings.append(f"{r['rel']}: its copy could not be taken back from the AWS side this run ({e})")
            return False
        self._taken += 1
        self._stale_copy(kind, self._md5_of(tok_id), p)   # it reached AWS: purged there if a name caused this
        self._own(p.get("tok_rel"), None, r["id"])
        self._remember_taken(tok_id, r["id"], p.get("tok_rel"))
        p["taken_back"] = True
        p.pop("pending_take_back", None)
        self._st["not_copied"][r["id"]] = version(r)
        self._held(r, why, rep)
        rep.removed.append((r["rel"], "taken back from the AWS side (moved to the bin there)"))
        self.save_state()
        return True

    def _clean_backups(self, tok, tok_dirs, rep):
        """rclone's backups on the AWS side (_sync_backup/<time>__sync_backup/, one per AWS run that replaced or
        deleted files): those older than BACKUP_KEEP_DAYS are moved to the bin (Drive keeps them there another 30
        days). The rest are never renamed: a file or folder there whose name holds a worker's name (kept from before
        the name was learnt) is moved to the bin. The staff side's own _sync_backup/ is never pruned."""
        cutoff = self.now() - timedelta(days=BACKUP_KEEP_DAYS)
        for d in sorted(d for d in tok_dirs if d.count("/") == 1 and d.startswith(BACKUP_DIR + "/")):
            try:
                made = datetime.strptime(d.split("/")[1][:19], "%Y-%m-%d-%H-%M-%S").replace(tzinfo=timezone.utc)
            except ValueError:
                continue                               # not one of rclone's backups: left alone
            if made < cutoff:
                try:
                    self.drive.trash(tok_dirs[d])
                except Exception as e:                 # noqa: BLE001
                    rep.warnings.append(f"{d}: old backup not moved to the bin this run ({e})")
                    continue
                for x in [x for x in tok_dirs if x == d or x.startswith(d + "/")]:
                    del tok_dirs[x]
                rep.removed.append((d, f"AWS-side backup older than {BACKUP_KEEP_DAYS} days: moved to the bin"))
        self._recheck_backup_pdfs(tok, rep)
        inside = lambda rel: rel.split("/", 2)[2] if rel.count("/") >= 2 else ""   # the path the backup was at
        for d in sorted((d for d in tok_dirs if d.startswith(BACKUP_DIR + "/")), key=lambda d: d.count("/")):
            if d in tok_dirs and inside(d) and screen.name_left(inside(d), self.ctx.scrubber, folder=True):
                try:
                    self.drive.trash(tok_dirs[d])
                except Exception as e:                 # noqa: BLE001
                    rep.warnings.append(f"{d}: not moved to the bin on the AWS side this run ({e})")
                    continue
                self._stale("path", *self._md5s_under(d))   # their content was live on AWS once: see _recheck_backup_pdfs
                for x in [x for x in tok_dirs if x == d or x.startswith(d + "/")]:
                    del tok_dirs[x]
                rep.removed.append((d, "backup folder with a worker's name: moved to the bin on the AWS side"))
        for i, f in list(tok.items()):
            if f["rel"].startswith(BACKUP_DIR + "/") and posixpath.dirname(f["rel"]) in tok_dirs \
                    and inside(f["rel"]) and screen.name_left(inside(f["rel"]), self.ctx.scrubber):
                try:
                    self.drive.trash(i)
                    self._stale("path", f.get("md5Checksum"))
                    rep.removed.append((f["rel"], "backup with a worker's name: moved to the bin on the AWS side"))
                except Exception as e:                 # noqa: BLE001
                    rep.warnings.append(f"{f['rel']}: not moved to the bin on the AWS side this run ({e})")

    def _recheck_backup_pdfs(self, tok, rep):
        """PDFs in the AWS-side backups (kept up to BACKUP_KEEP_DAYS) are read again when names were learnt since
        they were last read, as live copies are: one that now holds a worker's name is moved to the bin.
        The backups themselves never reach the server or S3 (the AWS run's filter skips _sync_backup/), but what
        is in them does, earlier: each is a file an AWS run replaced or deleted, so it was a live file the server
        pulled and, usually, an S3 version of its key that is now an old one. A binned backup's MD5 is therefore
        listed for purging too (_stale), as are those binned by _clean_backups for a name in their path."""
        count = self.ctx.scrubber.name_count
        seen = self._st.setdefault("backup_checked", {})          # file ID -> names known when last read
        start, later = time.monotonic(), 0
        for i, f in sorted(tok.items(), key=lambda kv: seen.get(kv[0], 0)):
            if not f["rel"].startswith(BACKUP_DIR + "/") or seen.get(i, 0) >= count \
                    or self.ctx.config.handling(f["name"]) != "scan_pdf":
                continue
            if time.monotonic() - start > BACKUP_RECHECK_SECONDS:
                later += 1
                continue
            try:
                at = f["rel"].split("/", 2)[2] if f["rel"].count("/") >= 2 else f["rel"]   # where it was before
                screen.screen_pdf(at, self.drive.download(f), self.ctx.scrubber)
                seen[i] = count
            except screen.Held as e:
                try:
                    self.drive.trash(i)
                    if isinstance(e, screen.Found):    # (one that can't be checked holds no known name)
                        self._stale("content", f.get("md5Checksum"))
                    seen.pop(i, None)
                    rep.removed.append((f["rel"], "backup PDF holding a worker's name: moved to the bin on the AWS side"))
                except Exception as e:                 # noqa: BLE001
                    rep.warnings.append(f"{f['rel']}: not moved to the bin on the AWS side this run ({e})")
            except Exception as e:                     # noqa: BLE001 - e.g. a download failed; next run
                rep.warnings.append(f"{f['rel']}: backup not checked again this run ({e})")
        for i in [i for i in seen if i not in tok]:    # gone (pruned, binned): forgotten
            del seen[i]
        if later:
            rep.warnings.append(f"{later} backup PDF(s) on the AWS side not yet checked again for newly learnt names; "
                                "the next run continues")

    def _hist(self, p):
        """The MD5s of the AWS-side copy of pair ``p`` as last seen and as this tokenizer wrote it before: the old
        S3 versions of a path that held a name may hold any of them."""
        return [p.get("tok_ver"), *p.get("tok_hist", [])] if p else []

    def _md5s_under(self, prefix):
        """The MD5s of every AWS-side file at or below ``prefix`` (a folder's path), with their pairs' history."""
        under = lambda rel: bool(rel) and (rel == prefix or rel.startswith(prefix + "/"))
        out = [f.get("md5Checksum") for f in self._tok.values() if under(f["rel"])]
        for p in self._st["pairs"]:
            if under(p.get("tok_rel")):
                out += self._hist(p)
        return out

    def _kind_of_old(self, tok_id, tok_rel, read=True):
        """What of an AWS-side copy holds a worker's name, with the names known now: "content" (a PDF whose text
        does; ``read``: it is read to know), "path" (its file or folder name does), or None (nothing does: binned
        for another reason, it is not listed). Spreadsheets and text were tokenized: their content is not judged."""
        f = self._tok.get(tok_id)
        if f is None:
            try:
                f = self.drive.get(tok_id)
            except Exception:                          # noqa: BLE001
                f = None
        name = posixpath.basename(tok_rel) if tok_rel else (f or {}).get("name", "")
        if read and f is not None and self.ctx.config.handling(name) == "scan_pdf":
            try:
                screen.screen_pdf(tok_rel or name, self.drive.download(f), self.ctx.scrubber)
            except screen.Found:
                return "content"
            except Exception:                          # noqa: BLE001 - can't be checked: no evidence of a name
                pass
        if tok_rel and (screen.has_name(tok_rel, self.ctx.scrubber) or screen.name_left(tok_rel, self.ctx.scrubber)):
            return "path"
        return None

    def _stale_copy(self, kind, md5, p):
        """List a binned copy: its own content as ``kind``, the copies written before it at the same path as "path"
        (their content was not checked, only the path is known to be gone). Nothing if ``kind`` is None."""
        if kind:
            self._stale(kind, md5)
            self._stale("path", *[m for m in self._hist(p) if m != md5])

    def _refresh(self, *md5s):
        """Keep listed entries from today (a copy seen on AWS again); nothing new is listed."""
        listed = {e["md5"] for e in self._st.get("purge", []) if isinstance(e, dict)}
        for m in md5s:
            for e in self._st.get("purge", []):
                if m in listed and e.get("md5") == m:
                    self._stale(e["kind"], m)
                    break

    def _stale(self, kind, *md5s):
        """Remember the content of AWS-side files renamed, moved, binned or taken back because of a worker's name
        learnt after they reached AWS, by Drive MD5 only (no path: a path may hold the name, and a hash of one can
        be guessed from lists of names). S3's ETag of a file uploaded in one part is its MD5: the AWS run deletes
        old S3 versions by it (deploy/purge_stale.py). ``kind``: "content" when the file's content held the name
        (every copy of it, under any key, holds it too), "path" when only its path did (only keys that are gone,
        and only when all they held is listed, are purged: the same content elsewhere, a blank template, say, is
        left alone). Both recorded: "content" wins. A file without an MD5 (a native Google doc, never pulled) or
        an empty one is not listed."""
        assert kind in PURGE_KINDS
        today = self.now().strftime("%Y-%m-%d")
        purge = {e["md5"]: e for e in self._st.get("purge", []) if isinstance(e, dict) and isinstance(e.get("md5"), str)}
        added = False
        for m in md5s:
            if isinstance(m, str) and MD5_HEX.fullmatch(m) and m != EMPTY_MD5:
                was = purge.get(m, {}).get("kind")
                purge[m] = {"md5": m, "kind": "content" if "content" in (kind, was) else "path", "at": today}
                added = True                           # seen again: kept from today
        if added:
            self._st["purge"] = list(purge.values())
            self.save_state()

    def _publish_purge(self, st, tok_root, own, rep):
        """AWS side: the MD5s of the content to purge (no path, no name), for the AWS run (deploy/purge_stale.py).
        Entries older than PURGE_KEEP_DAYS are dropped: by then S3 has expired old versions anyway (30 days) and
        every AWS run has had its chance."""
        cutoff = (self.now() - timedelta(days=PURGE_KEEP_DAYS)).strftime("%Y-%m-%d")
        st["purge"] = [e for e in st.get("purge", []) if isinstance(e, dict) and isinstance(e.get("md5"), str)
                       and MD5_HEX.fullmatch(e["md5"]) and e["md5"] != EMPTY_MD5 and e.get("kind") in PURGE_KINDS
                       and isinstance(e.get("at"), str) and e["at"] >= cutoff]
        if not st["purge"] and own["tok"].get(screen.PURGE_SUMMARY) is None:
            return
        data = json.dumps({"purge": [{"md5": e["md5"], "kind": e["kind"], "at": e["at"]}
                                     for e in sorted(st["purge"], key=lambda e: e["md5"])],
                           "note": "Drive MD5s of AWS-side files renamed, moved or taken back because of a worker's "
                           "name learnt after they reached AWS. kind content: the content held it, so the AWS run "
                           "deletes old S3 versions with that ETag under any key; kind path: only the path did, so "
                           "only keys that are gone and held nothing else listed are purged. "
                           f"Kept {PURGE_KEEP_DAYS} days"}, indent=1).encode()
        self._write_own(own["tok"].get(screen.PURGE_SUMMARY), tok_root, screen.PURGE_SUMMARY, data, True, rep)
        self.save_state()

    def _recheck_folders(self, st, raw, tok, tok_dirs, rep):
        """A folder on the AWS side whose name holds a worker's name (made before the name was learnt) is renamed
        there, the name replaced by its token; the copies in it follow."""
        while True:
            todo = [d for d in sorted(tok_dirs, key=lambda d: d.count("/"))
                    if d and not d.startswith(BACKUP_DIR) and self._aws_dir_name(d) != posixpath.basename(d)]
            if not todo:
                return
            rel = todo[0]
            parent, last = posixpath.split(rel)
            new = posixpath.join(parent, self._aws_dir_name(rel)) if parent else self._aws_dir_name(rel)
            renamed = False
            try:
                if new in tok_dirs:                    # the token-named folder exists already: move what's in it over
                    if not self._merge_folder(rel, new, tok_dirs, rep):
                        return
                else:
                    self.drive.rename(tok_dirs[rel], posixpath.basename(new))
                    renamed = True
            except Exception as e:                     # noqa: BLE001 - tried again next run
                rep.warnings.append(f"{rel}: folder not renamed on the AWS side this run ({e})")
                return
            if renamed:                                # (a merge lists what it moves itself)
                self._stale("path", *self._md5s_under(rel))   # every file in it had the name in its path: purged on AWS
            move = lambda x: new + x[len(rel):] if x == rel or x.startswith(rel + "/") else x
            for d in [d for d in tok_dirs if move(d) != d]:
                fid = tok_dirs.pop(d)
                tok_dirs.setdefault(move(d), fid)
            for p in st["pairs"]:
                if p.get("tok_rel"):
                    p["tok_rel"] = move(p["tok_rel"])
            self._owners = None
            for f in tok.values():
                f["rel"] = move(f["rel"])
            rep.moved.append((new, "folder on the AWS side: a worker's name, or words not known (#53), replaced by "
                                   "tokens"))
            self.save_state()

    def _aws_dir_name(self, d: str) -> str:
        """The name AWS-side folder ``d`` should have: learnt names as tokens (as before), and with
        vnext.opaque_file_names (#53) its other unknown words as FIL- tokens, except a project's own folder."""
        b = posixpath.basename(self._tp(d, folder=True))
        parts = d.split("/")
        if self.ctx.config.opaque_file_names and not (len(parts) == 2 and parts[0].casefold() == "projects"):
            b = NM.aws_name(b, self.ctx.scrubber)
        return b

    def _merge_folder(self, old, new, tok_dirs, rep) -> bool:
        """Move everything in AWS-side folder ``old`` into ``new`` (same name with the token), then bin ``old``.
        A subfolder that exists in both is merged the same way. False if something is in the way (left as is).
        What is moved is listed for purging on AWS as it goes, so a merge that fails part-way still lists it."""
        for c in self.drive.children(tok_dirs[old]):
            sub_old, sub_new = f"{old}/{c['name']}", f"{new}/{c['name']}"
            if c.get("mimeType") == "application/vnd.google-apps.folder":
                if sub_new in tok_dirs:
                    if not self._merge_folder(sub_old, sub_new, tok_dirs, rep):
                        return False
                    continue
            elif any(f["rel"] == sub_new for f in self._tok.values()):
                rep.warnings.append(f"{sub_old}: not moved on the AWS side, {sub_new} is already there")
                return False
            self.drive.move(c["id"], tok_dirs[new], tok_dirs[old])
            self._stale("path", c.get("md5Checksum"), *self._md5s_under(sub_old))   # (a folder: every file in it)
        self.drive.trash(tok_dirs[old])
        return True

    def _publish_held_back(self, st, raw, raw_root, tok_root, own, rep):
        """Staff side: Held back from AWS.csv, with each file, why, and since when. AWS side: only how many, as
        anonymous keys, for the outcome email (which is written on AWS and so can't name them)."""
        hb = st["held_back"]
        paired = {p["raw_id"] for p in st["pairs"] if not p.get("pending_take_back")}
        for rid in list(hb):
            r = raw.get(rid)
            if r is None or rid in paired or version(r) != hb[rid]["ver"] or self._excluded(r["rel"]):
                del hb[rid]                          # deleted, copied since, changed (checked again), excluded
            else:
                hb[rid]["rel"] = r["rel"]            # follow a move or rename
        rows = sorted(hb.values(), key=lambda h: h["rel"].lower())
        out = io.StringIO(newline="")
        w = csv.writer(out, lineterminator="\r\n")
        w.writerow(["File", "Why it was held back", "Since (UTC)"])
        for h in rows:
            w.writerow([_cell(h["rel"]), _cell(h["reason"]), h["since"]])
        if not rows:
            w.writerow(["(nothing is held back)", "", ""])
        sheet = out.getvalue().encode("utf-8-sig")
        keys = [{"key": hashlib.sha256(f"{rid}|{h['ver']}".encode()).hexdigest()[:16], "since": h["since"],
                 **self._tracker_entry(h["rel"])} for rid, h in sorted(hb.items())]
        sheet_id = self._write_own(own["raw"].get(screen.HELD_BACK_SHEET), raw_root, screen.HELD_BACK_SHEET, sheet,
                                   bool(rows), rep)
        if sheet_id and rows:
            self._check_not_shared(sheet_id, rep)
        # The sheet's Drive ID lets the outcome email link to it. A Drive ID grants nothing: opening the link needs
        # access to the staff drive (or to the file), as for any Drive file; no name or reason goes to AWS.
        summary = json.dumps({"held_back": keys, "sheet_id": sheet_id,
                              "note": "files kept in the staff drive and not copied to AWS; "
                              f"the staff drive's {screen.HELD_BACK_SHEET} says which and why"}, indent=1).encode()
        self._write_own(own["tok"].get(screen.HELD_BACK_SUMMARY), tok_root, screen.HELD_BACK_SUMMARY, summary,
                        True, rep)
        self.save_state()

    def _write_own(self, f, root, name, data, needed, rep):
        """Create or update one of the tokenizer's own files; its Drive ID (None if there is none)."""
        try:
            if f is None:
                return self.drive.create(root, name, data, mime_for(name))["id"] if needed else None
            if f.get("md5Checksum") != hashlib.md5(data).hexdigest():
                self.drive.update(f["id"], data, mime_for(name))
            return f["id"]
        except Exception as e:                         # noqa: BLE001 - never stops the run
            rep.warnings.append(f"{name}: not updated this run ({e})")
            return None                                # (no link to a list that may be out of date)

    def _check_not_shared(self, file_id, rep):
        """The held-back list names staff files: warn if someone turned on "anyone with the link" for it (the
        tokenizer never changes sharing; this only tells)."""
        try:
            if getattr(self.drive, "anyone_can_open", None) and self.drive.anyone_can_open(file_id):
                rep.warnings.append(f"{screen.HELD_BACK_SHEET} is shared with anyone who has the link; it lists staff "
                                    "file names. In Drive: Share > General access > Restricted")
        except Exception:                              # noqa: BLE001 - a check only
            pass

    # ---- the agent's project layout ---------------------------------------------------------------
    def _tidy_layout(self, raw_root, tok_root, rep):
        """For each project in the agent's current layout (see _migrated_projects):
        * make sure Uploads/ exists on both sides: a run on AWS creates it empty, and rclone doesn't copy empty
          folders, so staff would otherwise not see the folder they are meant to use;
        * on the AWS copy only, trash the previous layout's folders (Upload/, Reviewed payroll/, Wage rates/,
          Archive/) once they hold nothing at all: they are left behind there when the staff side is migrated (by
          the one-off migrate_layout.py), whose moves the mirror follows. The staff side is never touched here.
        A folder that can't be trashed stays, with one warning; nothing here ever stops the run."""
        sides = {"raw": raw_root, "tok": tok_root}
        refused = self._st.setdefault("tidy_refused", {})            # folder id -> reason (warned about once)
        self._st.pop("empty_old_folders", None)                       # used by an earlier version
        try:
            trees = {k: self.drive.tree(r) for k, r in sides.items()}
        except Exception as e:                                         # noqa: BLE001
            rep.warnings.append(f"project folders not tidied this run ({e})")
            return
        projects = _migrated_projects(set(trees["raw"][1]) | set(trees["tok"][1]))
        for side, (files, folders) in trees.items():
            where = "staff side" if side == "raw" else "AWS side"
            low = {r.lower() for r in folders}
            busy = set()
            for f in files.values():                                   # any file at all (Docs, shortcuts, .zip...)
                d = posixpath.dirname(f["rel"])
                while d:
                    busy.add(d)
                    d = posixpath.dirname(d)
            for proj in sorted(projects):
                if proj not in folders or self._excluded(proj + "/"):
                    continue
                if f"{proj}/uploads".lower() not in low:
                    try:
                        folders[f"{proj}/Uploads"] = self.drive.mkdir(folders[proj], "Uploads")
                        rep.folders.append((f"{proj}/Uploads", f"folder created ({where})"))
                    except Exception as e:                             # noqa: BLE001
                        rep.warnings.append(f"{proj}/Uploads: could not be created ({where}): {e}")
                if side != "tok":
                    continue
                old = [rel for rel in folders if _in_old_layout(rel, {proj}) and not self._excluded(rel + "/")]
                for rel in sorted(old, key=lambda r: r.count("/")):    # top-most first: one trash takes the tree
                    if rel in busy or any(rel.startswith(o + "/") for o in old if o != rel and o not in busy):
                        continue
                    fid = folders[rel]
                    try:
                        if _has_files(self.drive, fid):
                            continue                                   # something arrived meanwhile: keep it
                        self.drive.trash(fid)
                    except Exception as e:                             # noqa: BLE001
                        if fid not in refused:
                            rep.warnings.append(f"{rel}: empty folder of the previous layout could not be removed "
                                                f"({where}): {e}; it can be deleted by hand")
                        refused[fid] = str(e)[:200]
                        continue
                    refused.pop(fid, None)
                    rep.folders.append((rel, f"empty folder of the previous layout removed ({where})"))

    # ---- macro-enabled workbooks -----------------------------------------------------------------
    def _sync_staff_trackers(self, name, raw, raw_dirs, tok, st, stamp, own, raw_root, tok_root, rep) -> bool:
        """Change 1 (#93): copy each project's working tracker from the staff drive into Projects/<project>/Tracker/
        of the raw folder when it changed (pii_tokenizer/staff_trackers.py), and tell both sides where each tracker
        came from. The staff drive is only read (ST.ReadOnlyDrive). True when a copy was written (the caller lists
        the raw folder again)."""
        sc = self.ctx.config.staff_trackers or {}
        if sc.get("mirror") != name or not sc.get("source_folder_id"):
            return False
        ro = ST.ReadOnlyDrive(self.drive)
        index = ST.StaffIndex(ro, sc["source_folder_id"])
        rec = st.setdefault("staff_trackers", {})
        by_raw = {p["raw_id"]: p for p in st["pairs"] if p.get("raw_id")}
        tok_by_rel = {}
        for t in tok.values():
            tok_by_rel.setdefault(t["rel"], []).append(t)
        aws_busy = self._aws_run_in_progress(tok)
        changed, statuses, rows = False, {}, []
        pins = st.setdefault("staff_tracker_folders", {})
        covers = st.setdefault("cover_sheets", {})          # #151 step 3: last pass's status per project
        # 1. Where each project's tracker is: resolved for every project first, nothing written yet
        plans = []
        for proj, rel in sorted(ST.project_dirs(raw_dirs, sc.get("projects_dir") or "Projects").items()):
            tdir = f"{rel}/Tracker"
            here = [f for f in raw.values() if posixpath.dirname(f["rel"]) == tdir]
            pointer_f = next((f for f in here if f["name"].strip().lower() == ST.SOURCE_FILE), None)
            current = [f for f in here if f is not pointer_f and not f["name"].startswith("~$")
                       and "(AWS copy " not in f["name"]
                       and (os.path.splitext(f["name"])[1].lower() in (".xlsx", ".xlsm") or
                            f.get("mimeType") == GSHEET_MIME)]
            pointer = None
            try:
                if pointer_f is not None:
                    pointer = ST.read_pointer(self.drive.download(pointer_f).decode("utf-8-sig", "replace")) or ""
                out = ST.resolve(index, proj, pointer, [f["name"] for f in current], pins.get(proj))
            except Exception as e:  # noqa: BLE001 - try again next pass; the copy in Tracker/ stays as it is
                out = ST.Outcome(ST.ERROR, note=f"couldn't read the staff drive this time ({type(e).__name__})")
                rep.warnings.append(f"staff trackers: {type(e).__name__}")
            plans.append([proj, rel, tdir, current, pointer, out])
        # 2. Two projects can't have the same staff folder or the same tracker: one of them would be checking the other's.
        #    All of them are "ambiguous" (Not checked) until the administrator says which is which (tracker_source.txt).
        seen = {}
        for plan in plans:
            out = plan[5]
            for key in ([("folder", out.folder["id"])] if out.folder else []) + \
                       ([("file", out.source["id"])] if out.source and out.status == ST.COPIED else []):
                seen.setdefault(key, []).append(plan)
        clash = {id(p) for group in seen.values() if len(group) > 1 for p in group}
        for plan in plans:
            if id(plan) in clash:
                others = sorted(q[0] for g in seen.values() if plan in g and len(g) > 1 for q in g if q is not plan)
                plan[5] = ST.Outcome(ST.AMBIGUOUS, note="the same staff folder or tracker was found for "
                                     + "; ".join(sorted(set(others))) + f" too; add Tracker/{ST.SOURCE_FILE} with the "
                                     "right tracker's Drive link for each")
                pins.pop(plan[0], None)
        # 3. Pins, copies and reports
        for proj, rel, tdir, current, pointer, out in plans:
            src, modified = out.source, None
            try:
                if out.pin_missing:
                    pins[proj]["missing"] = 1
                elif out.drop_pin:
                    pins.pop(proj, None)
                elif out.folder is not None and out.matched_by == "pinned":
                    pins[proj].pop("missing", None)
                if out.folder is not None and out.matched_by not in ("pinned", ""):
                    # a new match (or a new one after the pinned folder went): pinned from now on, and said
                    pins[proj] = {"id": out.folder["id"], "matched_by": out.matched_by,
                                  "matched_at": self.now().strftime("%Y-%m-%dT%H:%M:%SZ")}
                elif pointer is not None:
                    pins.pop(proj, None)                 # tracker_source.txt decides: no pin to fall back on
                if out.status == ST.COPIED:
                    r = rec.get(proj) or {}
                    modified = src.get("modifiedTime")
                    if r.get("src_id") == src["id"] and r.get("ver") == ST.version(src) and r.get("copy_id") in raw:
                        out.status = ST.UNCHANGED
                    elif aws_busy or any(self._results_pending(f, by_raw, tok, tok_by_rel) for f in current):
                        # an AWS run is between its pull and its push, or its marks are still on their way back:
                        # the copy it checked must not be replaced under it (the next pass copies)
                        out.status, modified = ST.DEFERRED, r.get("modified")
                        out.note = "changed in the staff folder; copied on the next pass (an AWS check is running " \
                                   "or its marks are on their way back)"
                    else:
                        try:
                            self._write_staff_tracker(proj, tdir, src, current, raw_dirs, stamp, ro, rec)
                            changed = True
                        except TokenError as e:          # an .xlsm whose macros can't be removed safely
                            # (the folder kept: the cover sheet is still read there; review M1)
                            out = dataclasses.replace(out, status=ST.UNSUPPORTED, source=src, note="the tracker's "
                                                      f"macros can't be removed safely ({e}); save it as .xlsx or a "
                                                      "Google Sheet")
                            modified = None
            except Exception as e:  # noqa: BLE001 - try again next pass; the copy in Tracker/ stays as it is
                out = ST.Outcome(ST.ERROR, note=f"couldn't read the staff drive this time ({type(e).__name__})")
                rep.warnings.append(f"staff trackers: {type(e).__name__}")
            cover = self._cover_sheet(proj, out, pointer, index, ro, covers, rep)
            pin = pins.get(proj) or {}
            how = {"pointer": "tracker_source.txt", "tracker_name": "same tracker name", "project_number":
                   "project number", "folder_name": "folder name"}.get(pin.get("matched_by") if out.matched_by ==
                                                                      "pinned" else out.matched_by, "")
            where = f"{out.folder['name']} (matched by {how}" + (", again: the earlier folder is gone" if
                                                              out.rematched else "") + ")" if out.folder else ""
            rows.append((proj, ST.STATUS_WORDS.get(out.status, out.status), (src or {}).get("name", ""),
                         (src or {}).get("modifiedTime", ""), "; ".join(x for x in (where, out.note) if x),
                         # #151 step 3: the project's status from its cover sheet (the web app reads this column)
                         CS.PD.SHOWN.get(cover.get("status"), "Active"), cover.get("name", ""), CS.staff_note(cover)))
            # The AWS side's name for the project folder, as for the held-back summary (_tracker_entry): a folder whose
            # own name holds a phone number, email or SSN is never on AWS, so it gets no entry
            if not screen.name_reason(rel, self.ctx.scrubber, True):
                pin = (pins.get(proj) or {}) if out.folder is not None else {}
                statuses[posixpath.basename(self._tp(rel, True))] = {
                    "status": out.status, "modified": modified if out.status in ST.GOOD else None,
                    # how the staff folder was found, and when (codes only): the administrator email says so for a
                    # week after a new match, so the match is checked once (Tracker sources.csv names the folder)
                    "matched_by": "pointer" if out.matched_by == "pointer" else pin.get("matched_by"),
                    "matched_at": pin.get("matched_at") if out.matched_by != "pointer" else None,
                    # #151 step 3: two codes from the cover sheet (status, why); nothing typed, no file name
                    "project": CS.aws_entry(cover)}
        live = set(ST.project_dirs(raw_dirs, sc.get("projects_dir") or "Projects"))
        for gone in [p for p in st.get("staff_tracker_folders", {}) if p not in live]:
            del st["staff_tracker_folders"][gone]     # a renamed or removed project: its pin goes with it
        for gone in [p for p in rec if p not in live]:
            del rec[gone]
        for gone in [p for p in covers if p not in live]:
            del covers[gone]
        sheet = ST.staff_sheet(rows)
        self._write_own(own["raw"].get(ST.STAFF_SHEET), raw_root, ST.STAFF_SHEET, sheet, True, rep)
        self._write_own(own["tok"].get(ST.AWS_SUMMARY), tok_root, ST.AWS_SUMMARY, ST.summary_json(statuses), True,
                        rep)
        self.save_state()
        return changed

    @staticmethod
    def _cover_sheet(proj, out, pointer, index, ro, covers, rep) -> dict:
        """#151 step 3: this project's cover sheet ("Project Data...") in its staff folder: the folder the tracker
        step found, or the folder of the tracker tracker_source.txt names (pii_tokenizer/cover_sheet.py). Read-only;
        downloaded only when it changed. Google errors: the last status read is used, said as "error"."""
        folder = out.folder["id"] if out.folder is not None else None
        if folder is None and pointer is not None and out.source:     # also an .xls or macro tracker it names
            parent = (out.source.get("parents") or [None])[0]
            folder = parent if parent and parent != index.root_id else None   # a tracker at the drive's top: none
        last = covers.get(proj) or {}
        if folder is None and out.status == ST.ERROR:
            cover = {**last, "status": last.get("status", CS.PD.ACTIVE), "why": CS.PD.ERROR}
        else:
            try:
                cover = CS.look(ro, folder, last)
            except Exception as e:  # noqa: BLE001 - the last status read stands; tried again next pass
                cover = {**last, "status": last.get("status", CS.PD.ACTIVE), "why": CS.PD.ERROR}
                rep.warnings.append(f"cover sheets: {type(e).__name__}")
        if cover.get("why") != CS.PD.ERROR:
            covers[proj] = {k: cover[k] for k in ("id", "ver", "status", "why") if cover.get(k)}   # codes, no word
        return cover

    def _aws_run_in_progress(self, tok) -> bool:
        """The AWS pipeline's marker (deploy/run_pipeline.sh writes it right before its pull and removes it after its
        push) at the top of the tokenized folder, if it is recent: a run older than ST.RUN_MARKER_HOURS has stopped."""
        m = next((t for t in tok.values() if t["rel"] == ST.RUN_MARKER), None)
        if m is None:
            return False
        try:
            when = datetime.fromisoformat(str(m.get("modifiedTime")).replace("Z", "+00:00"))
        except ValueError:
            return True                                      # can't tell its age: assume the run is going on
        return self.now() - when < timedelta(hours=ST.RUN_MARKER_HOURS)

    @staticmethod
    def _results_pending(f, by_raw, tok, tok_by_rel) -> bool:
        """AWS's results for this raw file are on the Drive but not yet brought back: its tokenized copy changed since
        the mirror last saw it, or (rclone sync --backup-dir, as AWS pushes) the copy the mirror knows was moved to
        _sync_backup/ and a new file now sits at its path."""
        p = by_raw.get(f["id"])
        if not p:
            return False
        t = tok.get(p["tok_id"])
        at_path = [x for x in tok_by_rel.get(p.get("tok_rel") or "", []) if t is None or x["id"] != t["id"]]
        if t is None:
            return bool(at_path)
        if t["rel"].startswith(BACKUP_DIR + "/") or at_path:
            return True
        return version(t) != p.get("tok_ver")

    def _write_staff_tracker(self, proj, tdir, src, current, raw_dirs, stamp, ro, rec):
        data = ro.download(src)                      # a Google Sheet comes back as .xlsx (gcp.Drive.download)
        if os.path.splitext(src.get("name", ""))[1].lower() == ".xlsm":
            data = strip_macros(data)                # TokenError if it can't be done safely: reported as an error
        cname = ST.copy_name(src)
        target = next((f for f in current if f["name"] == cname and f.get("mimeType") != GSHEET_MIME), None)
        if target is not None:
            copy_id = self.drive.update(target["id"], data, ST.XLSX_MIME)["id"]
        else:
            copy_id = self.drive.create(self._folder(raw_dirs, tdir), cname, data, ST.XLSX_MIME)["id"]
        for f in current:                            # one tracker in Tracker/: the others are archived, not deleted
            if f is not target:
                self.drive.move(f["id"], self._folder(raw_dirs, f"{tdir}/Archived/{stamp}"), f["parents"][0])
        rec[proj] = {"src_id": src["id"], "ver": ST.version(src), "copy_id": copy_id,
                     "modified": src.get("modifiedTime")}

    def _convert_macros(self, raw, raw_dirs, tok, st, stamp, rep) -> bool:
        """Every macro-enabled workbook in the staff folder becomes an ordinary .xlsx beside it; the original is
        moved to the project's archive (<project>/Files/Archived/<ts>/ or <project>/Tracker/Archived/<ts>/) as <name>.xlsm<suffix>, which nothing opens by double-click and which is never
        copied to AWS (its extension is in no file rule). A workbook that cannot be converted safely stays where it
        is, is not copied, and is reported."""
        cfg = self.ctx.config
        if not cfg.macro_exts:
            return False
        changed = False
        for f in sorted(raw.values(), key=lambda f: f["rel"]):
            rel, name = f["rel"], f["name"]
            if os.path.splitext(name)[1].lower() not in cfg.macro_exts or self._excluded(rel) \
                    or rel.startswith(BACKUP_DIR + "/") or "/archive/" in "/" + rel.lower():
                continue
            if st["not_copied"].get(f["id"]) == version(f):
                continue                                  # already refused; not tried again until it changes
            try:
                plain = strip_macros(self.drive.download(f))
            except TokenError as e:
                rep.refused.append((rel, f"macros not removed: {e}"))
                st["not_copied"][f["id"]] = version(f)
                continue
            except Exception as e:  # noqa: BLE001 - try again next run
                rep.warnings.append(f"{rel}: macros not removed this run, will retry ({e})")
                continue
            rel_dir = posixpath.dirname(rel)
            parent = raw_dirs.get(rel_dir) or self._folder(raw_dirs, rel_dir)
            siblings = {g["name"].lower() for g in raw.values() if posixpath.dirname(g["rel"]) == rel_dir}
            stem = os.path.splitext(name)[0]
            new_name = stem + ".xlsx"
            if new_name.lower() in siblings:
                new_name = f"{stem} (no macros).xlsx"
            self.drive.create(parent, new_name, plain, mime_for(new_name))
            arch_rel = macro_archive_dir(rel_dir, stamp)
            self.drive.move(f["id"], self._folder(raw_dirs, arch_rel), f["parents"][0],
                            name=name + cfg.macro_archive_suffix)
            # A copy of the macro workbook sent to AWS earlier is taken back: the .xlsx replaces it.
            for p in [p for p in st["pairs"] if p["raw_id"] == f["id"]]:
                if p["tok_id"] in tok:
                    self.drive.trash(p["tok_id"])
                st["pairs"].remove(p)
            rep.converted.append((rel, f"macros removed: {new_name} saved beside it; the original moved to "
                                       f"{arch_rel}/{name}{cfg.macro_archive_suffix}"))
            changed = True
        if changed:
            self.save_state()
        return changed

    def _recent(self, p) -> bool:
        try:
            made = datetime.strptime(p.get("copied_utc") or "", "%Y-%m-%d-%H-%M-%S").replace(tzinfo=timezone.utc)
        except ValueError:
            return False
        return (self.now() - made).total_seconds() < RECENT_HOURS * 3600

    # ---- directions -----------------------------------------------------------------------------
    def _bin_settings_copy(self, t, rep) -> bool:
        """The settings file's AWS copy to the bin (#151); False (and a warning) if Drive refuses: tried next run."""
        try:
            self.drive.trash(t["id"])
            return True
        except Exception as e:                         # noqa: BLE001
            rep.warnings.append(f"{settings_file.REL}: its copy on the AWS side couldn't be moved to the bin this run "
                                f"({e}); tried again next run")
            return False

    def _run_forward(self, jobs, st, tok, tok_dirs, rep):
        # A file byte-for-byte identical to one already tokenized in this run (same Drive checksum and type: an
        # archived copy of an upload, a re-uploaded file) gets the same tokenized copy without being read again.
        # Safe: AWS has already received exactly that content.
        def same_key(r):
            return (r.get("md5Checksum"), os.path.splitext(r["name"])[1].lower()) \
                if r.get("md5Checksum") and r.get("mimeType") != GSHEET_MIME and \
                self.ctx.config.handling(r["name"]) not in ("scan_pdf", "skip") and \
                r.get("rel") != settings_file.REL else None    # a PDF's check, and the settings file's, depend on
                                                                # where it is (#151 review F3)
        # Kept only while another file still to come has the same content, so memory stays bounded by the batch
        # (keeping every copy of the run made the first full run run out of memory).
        to_come = Counter(same_key(r) for _, r, _ in jobs)
        same = {}
        j = 0
        while j < len(jobs):
            ready = []
            while j < len(jobs) and len(ready) < BATCH and sum(len(x[4] or b"") for x in ready) < BATCH_BYTES:
                p, r, tok_id = jobs[j]
                j += 1
                key = same_key(r)
                to_come[key] -= 1
                try:
                    why = screen.name_reason(_copy_rel(r), self.ctx.scrubber) or self._path_problem(r, p)  # check 2
                    if why:
                        raise screen.Held(why)
                    if key in same:
                        data, detail, first = same[key] if to_come[key] > 0 else same.pop(key)
                        ready.append((p, r, tok_id, r["name"], data, f"{detail} (same content as {first}: copy reused)"))
                        continue
                    name, data, detail = self._tokenize(r)
                    if key and data is not None and to_come[key] > 0:
                        same[key] = (data, detail, r["rel"])
                except (TokenError, screen.Held) as e:
                    self._held(r, str(e), rep)
                    ready.append((p, r, tok_id, None, None, None))
                    continue
                except Exception as e:  # noqa: BLE001 - e.g. Drive refused the download; try again next run
                    rep.warnings.append(f"{r['rel']}: not copied this run, will retry ({e})")
                    continue
                if data is None:
                    rep.skipped.append((r["rel"], detail))
                ready.append((p, r, tok_id, name, data, detail))
            if self.ctx.vault.dirty:          # tokens must be recoverable before any copy carrying them exists
                self.save_vault()
            # Which copies this batch is about to write: if the run is stopped before the state is saved (out of
            # memory, time limit), the next run recognises them as its own (see run), not as files from AWS.
            targets = {r["id"]: self._target(r, q, name) for q, r, tok_id, name, data, _ in ready if data is not None}
            st["in_flight"] = dict(targets)
            if st["in_flight"]:
                self.save_state()
            for p, r, tok_id, name, data, detail in ready:
                if data is not None:
                    # checked again now that the whole batch is tokenized: a payroll later in it may have taught
                    # the name of a worker this file is named after
                    why = screen.name_reason(_copy_rel(r), self.ctx.scrubber) or \
                        self._path_problem(r, p, name, target=targets[r["id"]])
                    if why:
                        self._held(r, why, rep)
                        data = None
                if data is None:
                    if tok_id:                # the raw file can no longer be copied safely: take the old copy away
                        try:
                            # what of the OLD copy (on AWS) holds a name now, if anything: not the new version's
                            # reason (that version never went), and a hold for anything else lists nothing
                            kind = self._kind_of_old(tok_id, p.get("tok_rel") if p else None)
                            self.drive.trash(tok_id)
                        except Exception as e:     # noqa: BLE001 - finished first thing next run
                            rep.warnings.append(f"{r['rel']}: its previous copy could not be taken back from the "
                                                f"AWS side this run ({e})")
                            p["pending_take_back"] = st["held_back"].get(r["id"], {}).get("reason", "held back")
                            st["not_copied"][r["id"]] = version(r)
                            continue
                        self._stale_copy(kind, self._md5_of(tok_id), p)
                        self._remember_taken(tok_id, r["id"], p.get("tok_rel") if p else None)
                        rep.removed.append((r["rel"], "previous tokenized copy moved to the bin"))
                    if p in st["pairs"]:
                        self._own(p.get("tok_rel"), None, r["id"])
                        st["pairs"].remove(p)
                    st["not_copied"][r["id"]] = version(r)    # not tried again until the file changes
                    continue
                tok_rel = targets[r["id"]]                  # workers' names in it replaced by their tokens
                tdir, tname = posixpath.split(tok_rel)
                if tok_id:                  # updated where it is (a folder renamed for a new name follows: _recheck)
                    old_rel = p.get("tok_rel") if p else None
                    old_md5 = self._md5_of(tok_id)
                    meta = self.drive.update(tok_id, data, mime_for(name), name=tname)
                    if old_rel and posixpath.basename(old_rel) != tname and \
                            screen.has_name(old_rel, self.ctx.scrubber):
                        # renamed here, a name in its old name replaced by its token (_recheck then has nothing
                        # to do): the old path's versions are purged on AWS like any other such rename
                        self._stale("path", old_md5, *self._hist(p))
                else:
                    meta = self.drive.create(self._folder(tok_dirs, tdir), tname, data, mime_for(name))
                if p is None:
                    p = {}
                    st["pairs"].append(p)
                self._own(p.get("tok_rel"), tok_rel, r["id"])
                p.update(raw_id=r["id"], tok_id=meta["id"], rel=r["rel"], tok_rel=tok_rel,
                         raw_ver=version(r), tok_ver=version(meta), names_at_check=self.ctx.scrubber.name_count)
                p["tok_hist"] = (p.get("tok_hist", []) + [version(meta)])[-6:]   # copies written (see the pairs loop)
                if not tok_id:          # first copy only: a copy sent again doesn't restart the 12 h (no endless re-copy)
                    p.setdefault("copied_utc", self.now().strftime("%Y-%m-%d-%H-%M-%S"))
                st["not_copied"].pop(r["id"], None)
                rep.forwarded.append((r["rel"], detail))
            st["in_flight"] = {}
            self.save_state()

    def _back(self, t, r, p, raw_dirs, rep, conflict_stamp=None, new=False) -> bool:
        """Detokenize copy ``t`` into the raw folder: over raw file ``r``, or as a new file.
        Returns False (and reports it) if that failed; the pair is then left as it was, so it is retried."""
        try:
            self._back_one(t, r, p, raw_dirs, rep, conflict_stamp, new)
        except Exception as e:  # noqa: BLE001
            rep.warnings.append(f"{t['rel']}: not restored this run, will retry ({e})")
            return False
        self.save_state()                    # a restored file must never be restored twice
        return True

    def _back_one(self, t, r, p, raw_dirs, rep, conflict_stamp, new=False):
        data = self.drive.download(t)
        out, drep = detokenize_file(t["name"], data, self.ctx.vault)
        rel_dir = self._raw_dir(posixpath.dirname(t["rel"]))
        parent = self._folder(raw_dirs, rel_dir)
        name = self._detok_name(t["name"])        # "EMP-3f2a91c0 notes.xlsx" -> "Ana Rivera notes.xlsx"
        detail = drep.summary()
        if drep.unknown_tokens:
            rep.warnings.append(f"{t['rel']}: tokens not in the vault, left as they are: "
                                + ", ".join(sorted(drep.unknown_tokens)[:5]))
        if conflict_stamp and new:
            # a new file from AWS whose name is taken in the staff folder: saved beside it, and followed from now on
            stem, ext = os.path.splitext(name)
            cname = f"{stem} (AWS copy {conflict_stamp}){ext}"
            meta = self.drive.create(parent, cname, out, mime_for(t["name"]))
            p.update(raw_id=meta["id"], rel=posixpath.join(rel_dir, cname).lstrip("/"), tok_rel=t["rel"],
                     raw_ver=version(meta), tok_ver=version(t), aws_named=True)
            self._own(None, t["rel"], meta["id"])
            rep.returned.append((t["rel"], f"{detail}; a staff file has its name, so saved as {cname}"))
            return
        if conflict_stamp:
            stem, ext = os.path.splitext(name)
            cname = f"{stem} (AWS copy {conflict_stamp}){ext}"
            meta = self.drive.create(parent, cname, out, mime_for(t["name"]))
            self._st["retired"][meta["id"]] = version(meta)      # Google-side only: never copied to AWS
            rep.conflicts.append((t["rel"], f"changed on both sides; AWS's version saved as {cname}, "
                                            "the raw file was copied over it"))
            return
        if r is not None and r.get("mimeType") == GSHEET_MIME:
            # A Google Sheet cannot be overwritten with an .xlsx: write it alongside and pair that instead.
            meta = self.drive.create(parent, name, out, mime_for(t["name"]))
            self._st["retired"][r["id"]] = version(r)            # the Sheet stays, but is not copied again
            self._st["retired_pair"][r["id"]] = p["tok_id"]      # ... unless staff edit it (see run)
        elif r is not None:
            meta = self.drive.update(r["id"], out, mime_for(t["name"]))
        else:
            meta = self.drive.create(parent, name, out, mime_for(t["name"]))
        same = r is not None and r.get("mimeType") != GSHEET_MIME
        self._own(p.get("tok_rel"), t["rel"], meta["id"])
        p.update(raw_id=meta["id"], rel=r["rel"] if same else posixpath.join(rel_dir, name).lstrip("/"),
                 tok_rel=t["rel"], raw_ver=version(meta), tok_ver=version(t))
        if r is None:
            p["aws_named"] = True                  # made on AWS: its name there is AWS's, not tokenized from staff
        rep.returned.append((t["rel"], detail))

    def _follow(self, f, dirs, new_rel, p, rep, where, why=None):
        new_dir, new_name = posixpath.split(new_rel)
        if new_rel == f["rel"]:
            return
        if posixpath.dirname(f["rel"]) == new_dir:
            self.drive.rename(f["id"], new_name)
        else:
            self.drive.move(f["id"], self._folder(dirs, new_dir), f["parents"][0], name=new_name)
        old_rel = f["rel"]
        f["rel"], f["name"] = new_rel, new_name
        if where == "the raw folder":
            p["tok_rel"] = new_rel
            p["prev_tok_rel"] = old_rel                # an AWS run that pulled before may push the old name again
            p["prev_tok_utc"] = self.now().strftime("%Y-%m-%d-%H-%M-%S")
            self._own(old_rel, new_rel, p.get("raw_id") or p.get("tok_id"))
        rep.moved.append((new_rel, why or f"moved or renamed in {where}; followed on the other side"))


def _cell(text: str) -> str:
    """A spreadsheet cell that is shown, never run: a file named "=HYPERLINK(...)" must not become a formula."""
    return "'" + text if text[:1] in ("=", "+", "-", "@", "\t", "\r") else text


MIGRATED_MARKERS = ("Payroll", "Files/Logs", "Files/Wage rates", "Files/L&I data")


def _migrated_projects(dirs) -> set:
    """Projects/<project> folders the agent has already moved to its current layout: they have Payroll/ or one of
    the Files/ folders only the agent makes. Not Files/Archived/ or Tracker/Archived/ alone: the tokenizer itself
    parks converted macro workbooks there, in projects the agent hasn't migrated yet."""
    out = set()
    for rel in dirs:
        parts = rel.split("/")
        if len(parts) >= 3 and parts[0].lower() == "projects" and "/".join(parts[2:4]) in MIGRATED_MARKERS:
            out.add("/".join(parts[:2]))
    return out


def _has_files(drive, folder_id) -> bool:
    """True if anything other than a folder is somewhere under folder_id."""
    todo = [folder_id]
    while todo:
        for c in drive.children(todo.pop()):
            if c.get("mimeType") == "application/vnd.google-apps.folder":
                todo.append(c["id"])
            else:
                return True
    return False


def _in_old_layout(rel: str, migrated) -> bool:
    parts = rel.split("/")
    return len(parts) >= 3 and "/".join(parts[:2]) in migrated and parts[2] in OLD_LAYOUT


def _forward_order(rel: str, cfg) -> tuple:
    low, ext = rel.lower(), os.path.splitext(rel)[1].lower()
    tabular = ext in cfg.tokenize_exts
    payroll = tabular and ("/uploads/" in "/" + low or "/upload/" in "/" + low or "payroll" in low)
    return (0 if payroll else 1 if tabular else 2, rel)


_NAME_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9])[A-Z]{2,5}-[0-9a-f]{%d}(?![0-9a-f])" % TOKEN_HEX)


def _copy_rel(raw_f: dict) -> str:
    """Where the tokenized copy of a raw file lives: same path; Google Sheets get '.xlsx'."""
    rel = raw_f["rel"]
    if raw_f.get("mimeType") == GSHEET_MIME and not rel.lower().endswith(".xlsx"):
        rel += ".xlsx"
    return rel


def _raw_rel(tok_f: dict, raw_f: dict) -> str:
    rel = tok_f["rel"]
    if raw_f.get("mimeType") == GSHEET_MIME and rel.lower().endswith(".xlsx") and \
            not raw_f["name"].lower().endswith(".xlsx"):
        rel = rel[:-5]
    return rel


def run_all(drive, ctx: Context, state: dict, save_vault, save_state, now=None, only=None) -> list[str]:
    lines = []
    m = Mirror(drive, ctx, state, save_vault, save_state, now)
    for mc in ctx.config.mirrors:
        if only and mc.name != only:
            continue
        if not mc.raw_folder_id or not mc.tokenized_folder_id:
            lines.append(f"[{mc.name}] skipped: folder IDs not set")
            continue
        if mc.raw_folder_id == mc.tokenized_folder_id:
            raise RuntimeError(f"[{mc.name}] the raw and tokenized folder IDs are the same")
        rep = m.run(mc.name, mc.raw_folder_id, mc.tokenized_folder_id)
        lines += rep.lines(mc.name)
    save_state()
    return lines


_LOGGED_EXTS = {".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".tsv", ".pdf", ".json", ".docx", ".doc"}


def _progress(f: dict) -> None:
    """One line in the job's log per file read: its Drive ID, type and size, and the memory used so far. No name
    or path (they may hold a worker's name). If the job runs out of memory, the last line names the file."""
    try:
        import resource
        peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024
    except Exception:                          # noqa: BLE001 - not on every system
        peak = "?"
    ext = os.path.splitext(f.get("name", ""))[1].lower()
    kind = ext if ext in _LOGGED_EXTS else f.get("mimeType", "?")     # never part of a name ("J.Garcia")
    print(f"reading {f.get('id')} ({kind}, "
          f"{f.get('size', '?')} bytes); peak memory so far {peak} MB", flush=True)


def dumps_state(state: dict) -> bytes:
    return json.dumps(state, indent=1, sort_keys=True).encode()
