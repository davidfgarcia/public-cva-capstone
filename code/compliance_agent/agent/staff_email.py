"""The short outcome email for the firm's staff (notifications.config `to`), and the classification both emails share.

Every result is judged on two separate scales:
  Run       did the task itself work? Completed, Failed (a technical problem: the administrator's business),
            Skipped (nothing to do, e.g. no new payroll) or Not applicable (the task doesn't apply to the project:
            the intent/affidavit check of a project whose tracker has no intent or affidavit columns, whose Project
            Info says no L&I filing is required, or that the administrator turned it off for).
  Findings  did the project pass? No issues, Needs review (rule breaks or lines to review, tracker entries missing,
            data gaps, files to fix), Not checked (the task failed or was skipped) or Not applicable (as Run).
The staff email speaks only of Findings: its subject ("Needs review: 4 of 5 projects"), a grid (one line per project,
one column per task) with an "At a glance" column, and "What to do". A technical failure reaches staff only as
"Not checked ... the administrator has the details", never as a red Failure. Technical warnings (a download that
failed, an AI review that got no answer) go to the administrator's email only. Nothing technical: no rule codes,
settings, servers or pipeline steps.
"""
from __future__ import annotations

import datetime as dt
import html
import re
from pathlib import Path

from .util import shown_name

# Findings, the staff's verdict, and its colour
NO_ISSUES, NEEDS_REVIEW, NOT_CHECKED = "No issues", "Needs review", "Not checked"
# Both verdicts of a task that doesn't apply to the project (never a problem, never counted as not checked)
NOT_APPLICABLE = "Not applicable"
COLOR = {NO_ISSUES: "#1e7a46", NEEDS_REVIEW: "#9a6700", NOT_CHECKED: "#55605b", NOT_APPLICABLE: "#55605b"}
# Run, the administrator's verdict
COMPLETED, FAILED, SKIPPED = "Completed", "Failed", "Skipped"
# What one reason means: a technical failure; something for staff to look at or fix; a normal skip (no new payroll);
# a task that doesn't apply to the project; a technical warning for the administrator only; or a note for staff that
# changes no verdict (results arrive later)
KINDS = ("failed", "review", "skipped", "na", "tech", "info")
NOTIFIED = "The administrator has the details."

# A run that didn't do its job at all: what staff need to know, by the pipeline step that failed.
STEP_STAFF = {
    "lock": "This run didn't start because another run was still going. Your files will be checked in the next run.",
    "start": "This run didn't start because of a technical problem. Nothing was changed. " + NOTIFIED,
    "tokenize": "This run didn't start: a preparation step wasn't ready. Nothing was changed, and the next run will "
                "try again.",
    "stopped": "The run stopped before it finished, so some projects may not have been checked. " + NOTIFIED,
    "agent": "The run stopped before it finished, so nothing was changed. " + NOTIFIED,
    "push": "Results not saved: skip the review for now. " + NOTIFIED,
    "timeout": "The run ran out of time, so some projects may not have been checked. " + NOTIFIED,
}
# The same, in a few words for the subject and "At a glance": only these steps mean the run didn't start or finish
STEP_SHORT = {"lock": "the run didn't start", "tokenize": "the run didn't start", "start": "the run didn't start", "stopped": "the run didn't finish",
              "agent": "the run didn't finish", "timeout": "the run didn't finish", "push": "results not saved"}
STAFF_SILENT_STEPS = {"mirror"}       # only the administrator's copy failed: nothing changes for staff
# The tokenizer not run before / after the sync (notify.TOKENIZER_WARNING keys), in staff terms
LATE = {
    "before": "Files uploaded since last night may not have been included; they will be checked in the next run.",
    "after": "The results will appear in the shared drive a little later (by tomorrow morning).",
}
GENERIC_PROBLEM = "The run hit a technical problem, so some results may be missing. " + NOTIFIED
# The staff email's Run words (#77): never "failed" (#44), and only what staff can act on
RUN_COMPLETED = "Run completed"
RUN_WORDS = {"the run didn't start": "Run didn't start", "the run didn't finish": "Run didn't finish",
             "results not saved": "Run completed, results not saved"}
RUN_PROBLEM = "Run had a technical problem"

_REVIEWED = re.compile(r"^payroll: (?:re-)?reviewed (.+?) \(([\d,]+) lines\): ([\d,]+) rule-breaking, ([\d,]+) suspected")
_MARKED = re.compile(r"^intent: ([\d,]+) invalid/missing value\(s\) marked in (.+?)(?: \(|;|$)")
_LNI_ONLY = re.compile(r"^intent: ([\d,]+) contractor\(s\) in L&I's records for this project are not in the tracker")
_HIRING = re.compile(r"^intent: ([\d,]+) hiring contractor\(s\) in the tracker don't match the L&I intent")
_OVERWRITTEN = "WARNING intent: the tracker was replaced by an older copy"
_UNREAD = re.compile(r"^wages: could not read (.+?); left in Uploads")
_NOJUR = re.compile(r"^wages: no Jurisdiction column in (.+?); left in Uploads")
_PARTJUR = re.compile(r"^wages: (.+?): ([\d,]+) rate\(s\) with a blank or unrecognised Jurisdiction were not used")
_NORATE = re.compile(r"([\d,]+) of ([\d,]+) line\(s\) \(([\d.]+)%\) had no wage rate")
_NODATA = re.compile(r"([\d,]+) (federal|state|union) line\(s\) but the project has no \w+ wage rates")
_UNCLEAR = re.compile(r"Jurisdiction '(.*?)' names no wage determination and the project has several (\w+) ones "
                      r"\((.*?)\); ([\d,]+) line")
_COLUMNS = re.compile(r"columns not found: (.+?); ")
# In this order in the grid: wage updates last (skipped in most runs), just left of "At a glance"
TASK_LABEL = {"payroll": "Payroll review", "intent": "Intent/affidavit check", "wages": "Wage updates",
              "uploads": "Uploads", "project": "Project", "run": "Run", "results": "Results"}
GRID_ORDER = ("payroll", "intent", "wages")
RANK = {"failed": 4, "review": 3, "info": 2, "skipped": 1, "na": 1, "tech": 0}
# Skips that need someone to do something (shown in "What to do"; the verdict is then Not checked). The others
# (no new payroll, no new wage files, no tracker, not needed) are normal and only shown in the grid.
ACTION_SKIPS = {"Tracker open in Excel", "Folders not in expected layout", "Tracker held back", "Skipped",
                "Several trackers in the staff folder", "Tracker link not found", "Staff folder not found",
                "Tracker format not readable", "Several possible working folders",
                "No wage rates file", "Wage type unclear", "No Jurisdiction in wage rates",
                "No payroll lines", "Two exports in one file"}             # the last two: agent/upload_checks.py
# Skips that leave a project not checked although nobody needs to act (the check couldn't be done at all, L&I's
# records being down or out of date): the Findings verdict and the summary say Not checked, never No issues; they are
# not listed in "What to do" (review of #119).
# (David, Nov 9, 2026: "public data", so staff don't read it as L&I's own website, which may be newer)
UNCHECKED_SKIPS = {"L&I public data out of date", "L&I public data site down", "L&I public data not readable"}


def _unchecked_skip(short: str) -> bool:
    """One of UNCHECKED_SKIPS, whatever follows it ("L&I public data out of date (10-28-2026)"): matched by its start,
    so a date in any form can't make the run look checked (review of #144)."""
    return any(short == u or short.startswith(u + " ") for u in UNCHECKED_SKIPS)


def _lni_date(text: str) -> str:
    """L&I's last update as the emails' intent line gives it ("Oct 28, 2026") in MM-dd-yyyy (David, Nov 9, 2026);
    the text as it is if it isn't such a date."""
    try:
        return dt.datetime.strptime(text, "%b %d, %Y").strftime("%m-%d-%Y")
    except ValueError:
        return text


ALL = "All projects"
# Why the intent/affidavit check doesn't apply (agent/intent.py is_applicable): (reason, few words), in staff terms.
# The app's Run tasks tab greys the check out with the same words (ui/server.py).
NA_REASON = {
    "Project Info says no L&I filing required": ("The tracker's Project Info says no L&I filing is required.",
                                                 "No L&I filing required"),
    "tracker has no intent/affidavit columns": ("The tracker has no intent or affidavit columns to check.",
                                                "No intent/affidavit columns in the tracker"),
    "set in config.yaml": ("Turned off for this project by the administrator.", "Turned off by the administrator"),
}
NA_OTHER = ("This check doesn't apply to this project.", "Doesn't apply to this project")


def _n(s) -> int:
    return int(str(s).replace(",", ""))


def _plural(n, one, many=None):
    return one if n == 1 else (many or one + "s")


def gap_text(gap: str, name_file=False) -> str:
    """A data gap (agent/payroll.py), told the way staff would: what wasn't checked and what to do about it.
    name_file: say which payroll (when the project had several this run)."""
    file, sep, rest = gap.partition(": ")
    if not sep or not re.search(r"\.(xlsx|xlsm|xls|csv)$", file, re.I):
        file, rest = "", gap
    where = f" ({file})" if file and name_file else ""
    if m := _NORATE.search(rest):
        return (f"{_n(m[1]):,} of {_n(m[2]):,} payroll lines{where} ({float(m[3]):.0f}%) have no wage rate for their "
                "classification and date, so their pay wasn't checked. They're listed in the Comments sheet: upload "
                "the missing wage rates, or tell the administrator which classification applies.")
    if m := _NODATA.search(rest):
        return (f"{_n(m[1]):,} {m[2]} payroll lines{where} weren't checked: the project has no {m[2]} wage rates. "
                f"Upload the {m[2]} wage determination.")
    if m := _UNCLEAR.search(rest):
        return (f"{_n(m[4]):,} payroll lines{where} with Jurisdiction \"{m[1] or '(empty)'}\" don't name a {m[2]} "
                f"wage determination, and the project has several ({m[3]}). Tell the administrator which applies.")
    if m := _COLUMNS.search(rest):
        return (f"The payroll{where} is missing columns ({m[1]}), so some checks couldn't run. Export it again with "
                "them and upload it.")
    # anything else: the agent's own sentence, without settings or rule codes
    rest = re.sub(r"\s*\((?:[^()]*config\.yaml[^()]*|S\d\d[^()]*|failure from[^()]*)\)", "", rest)
    rest = re.sub(r",? or (?:list|name|map|set) [^.]*config\.yaml[^.]*", "", rest)
    return (rest[:1].upper() + rest[1:]).rstrip(". ") + "."


def gap_short(line: str) -> str:
    """A data gap in a few words, for the at-a-glance column."""
    if m := _NORATE.search(line):
        return f"{float(m[3]):.0f}% of lines have no wage rate"
    if m := _NODATA.search(line):
        return f"No {m[2]} wage rates"
    if _UNCLEAR.search(line):
        return "Wage determination unclear"
    if _COLUMNS.search(line):
        return "Payroll columns missing"
    return "Some lines not checked"


class Row:
    """One task's result for one project (or for the whole run, project ALL): its reasons, each with its kind."""
    __slots__ = ("project", "key", "items", "held_uploads")

    def __init__(self, project, key):
        self.project, self.key, self.items = project, key, []      # items: (kind, reason, short)
        self.held_uploads = 0                                       # files held back in its Uploads/ (said here)

    @property
    def task(self):
        return TASK_LABEL.get(self.key, self.key)

    @property
    def kinds(self):
        return {k for k, _r, _s in self.items}

    @property
    def run(self):
        k = self.kinds
        if "failed" in k:
            return FAILED
        if "na" in k and not k & {"review", "skipped"}:
            return NOT_APPLICABLE
        if "skipped" in k and not k & {"review"}:
            return SKIPPED
        return COMPLETED

    @property
    def findings(self):
        """Needs review whenever something was found, even if the task then failed (Run says that); Not checked when
        the task failed or was skipped before finding anything."""
        if "review" in self.kinds:
            return NEEDS_REVIEW
        if self.run == NOT_APPLICABLE:
            return NOT_APPLICABLE
        return NOT_CHECKED if self.run != COMPLETED else NO_ISSUES

    def visible(self, admin=False):
        """(kind, reason, short) a reader is told about, worst first: staff never see technical warnings."""
        return sorted((i for i in self.items if admin or i[0] != "tech"), key=lambda i: -RANK[i[0]])

    def reason(self, admin=False):
        return " ".join(r for _k, r, _s in self.visible(admin))

    def short(self, admin=False):
        return "; ".join(s for _k, _r, s in self.visible(admin))

    def needs_action(self, admin=False):
        """Listed in "What to do": failures, things to review, notes, and skips someone has to act on."""
        return any(k in ("failed", "review", "info") or (k == "skipped" and s in ACTION_SKIPS) or (admin and k == "tech")
                   for k, _r, s in self.items)

    def unchecked(self, admin=False):
        """Not checked, whether or not someone has to act: a skip that needs someone, or one of UNCHECKED_SKIPS."""
        return self.needs_action(admin) or any(k == "skipped" and _unchecked_skip(s) for k, _r, s in self.items)

    def __repr__(self):
        return f"Row({self.project!r}, {self.key!r}, {self.run}, {self.findings}, {self.items!r})"


def held_uploads_text(n: int, linked: bool = True) -> str:
    """Said on the first of a project's payroll review, wage update or intent check that ran when files in its
    Uploads folder were kept out of the check; the email links the list to it (else names it: review L1)."""
    where = "The list says" if linked else "The \"Held back\" list at the top of the shared drive says"
    return (f"{n} {_plural(n, 'file')} in the Uploads folder {_plural(n, 'was', 'were')} not checked (for example, "
            f"an SSN in the file name, or a payroll saved as a PDF). {where} which and how to fix "
            f"{_plural(n, 'it', 'them')}.")


def task_rows(lines, details=None, results_late=False, project="", held_tracker=False, held_uploads=0,
              held_linked=True) -> list[Row]:
    """The results of one project's tasks that did not simply finish with no issues, with their reasons.
    held_tracker: the tokenizer kept this project's tracker in the staff drive (couldn't read it safely): "no
    tracker" would be wrong, the tracker exists. held_uploads: how many of its Uploads/ files it kept there, said on
    the payroll review (else the wage update) when the run did one: those files weren't part of it."""
    rows: dict[str, Row] = {}

    def add(task, kind, reason, short, first=False):
        # a file's name the AWS side has only as tokens (#53): "…" in their place; staff find it by the words left
        reason, short = shown_name(reason), shown_name(short)
        r = rows.setdefault(task, Row(project, task))
        if not any(reason == x for _k, x, _s in r.items):
            r.items.insert(0, (kind, reason, short)) if first else r.items.append((kind, reason, short))

    where = "will appear in the Payroll folder" if results_late else "is in the Payroll folder"
    several = sum(1 for line in lines if _REVIEWED.match(line)) > 1
    unrec = 0
    for line in lines:
        if line.startswith("GAP "):
            # lines that couldn't be checked at all: told before the flags on the lines that were
            add("payroll", "review", gap_text(line.split(": ", 1)[1], name_file=several), gap_short(line), first=True)
        elif m := _REVIEWED.match(line):
            reds, oranges = _n(m[3]), _n(m[4])
            if reds or oranges:
                found = ([f"{reds:,} break a rule (red)"] if reds else []) + \
                        ([f"{oranges:,} need a review (orange)"] if oranges else [])
                add("payroll", "review", f"{_n(m[2]):,} lines checked{f' in {m[1]}' if several else ''}: "
                    + " and ".join(found) + f". The reviewed file {where}.",
                    ", ".join(([f"{reds:,} rule {_plural(reds, 'break')}"] if reds else [])
                              + ([f"{oranges:,} to review"] if oranges else [])))
        elif line.startswith("payroll: no new payroll uploaded;"):
            add("payroll", "skipped", "No new payroll file in the Uploads folder (the last one was already reviewed).",
                "No new payroll")
        elif line.startswith("payroll: no new payroll uploaded"):
            add("payroll", "skipped", "No payroll file in the Uploads folder.", "No payroll uploaded")
        elif line == "wages: no new wage files":
            add("wages", "skipped", "No new wage files in the Uploads folder.", "No new wage files")
        elif re.match(r"^wages: .+ skipped: already processed", line):
            add("wages", "skipped", "An uploaded wage file had already been added before; it was archived.",
                "Wage file already added")
        elif m := _UNREAD.match(line):
            names = [x.strip() for x in m[1].split(",") if x.strip()]
            add("wages", "review", f"{len(names)} wage {_plural(len(names), 'file')} couldn't be read and "
                f"{_plural(len(names), 'is', 'are')} still in Uploads: {', '.join(names)}. Check the format "
                "(a wage determination as text or PDF, the L&I rates as CSV or PDF).",
                f"{len(names)} wage {_plural(len(names), 'file')} unreadable")
        elif m := _NOJUR.match(line):
            names = [x.strip() for x in m[1].split(",") if x.strip()]
            add("wages", "review", f"{len(names)} wage rates {_plural(len(names), 'file has', 'files have')} no "
                f"Jurisdiction column and {_plural(len(names), 'is', 'are')} still in Uploads: {', '.join(names)}. "
                "Add a Jurisdiction column (State, Federal or Union for each rate) and upload again.",
                f"{len(names)} wage {_plural(len(names), 'file')} without Jurisdiction")
        elif m := _PARTJUR.match(line):
            n = _n(m[2])
            add("wages", "review", f"{n:,} {_plural(n, 'rate')} in {m[1]} {_plural(n, 'has', 'have')} a blank or "
                "unrecognised Jurisdiction and {} not used; the rest were added. Fill in State, Federal or Union for "
                "{} and upload the file again.".format(_plural(n, "was", "were"), _plural(n, "it", "them")),
                f"{n:,} {_plural(n, 'rate')} without Jurisdiction")
        elif line.startswith("WARNING federal: could not download"):
            add("wages", "tech", "The federal wage rates couldn't be downloaded from SAM.gov this time; the rates "
                "already in the project were used.", "Federal rates not downloaded")
        elif re.match(r"^wages: \d+ older (federal|state|union) table\(s\) in ", line):
            add("wages", "tech", "Older wage tables of the same kind were found next to the one in use; they were "
                "moved to Files/Archived without merging their rates.", "Older wage tables set aside")
        elif line.startswith("WARNING federal:"):
            add("wages", "tech", "A federal wage determination listed for this project isn't valid.",
                "Federal determination invalid")
        elif line == "intent: no compliance tracker in Tracker/" and held_tracker:
            add("intent", "skipped", "The tracker couldn't be read safely, so it wasn't checked. Tell the "
                "administrator.", "Tracker held back")
        elif line == "intent: no compliance tracker in Tracker/":
            add("intent", "skipped", "No tracker in the Tracker folder.", "No tracker")
        elif line == "intent: no compliance tracker in the staff folder":
            add("intent", "skipped", "No tracker in the project's working folder.", "No tracker")
        elif line.startswith("intent: NOT CHECKED - the staff folder has more than one"):
            add("intent", "skipped", "The project's working folder has more than one file that looks like the "
                "tracker, so none was checked. Keep only the current one there, or tell the administrator which it "
                "is.", "Several trackers in the staff folder")
        elif line.startswith("intent: NOT CHECKED - the tracker named in tracker_source.txt"):
            add("intent", "skipped", "The tracker the administrator pointed the CVA to wasn't found, so it wasn't "
                "checked. Tell the administrator.", "Tracker link not found")
        elif line.startswith("intent: NOT CHECKED - the tracker in the staff folder is in a format"):
            add("intent", "skipped", "The tracker in the project's working folder is in a format the CVA can't read, "
                "so it wasn't checked. Save it as an Excel workbook (.xlsx) or a Google Sheet.",
                "Tracker format not readable")
        elif line.startswith("intent: NOT CHECKED - several project folders in the staff drive"):
            add("intent", "skipped", "More than one project folder could be this project's working folder, so the "
                "tracker wasn't checked. Tell the administrator which folder it is.",
                "Several possible working folders")
        elif line.startswith("WARNING intent: the tracker's folder in the staff drive was matched"):
            add("intent", "tech", line.split("WARNING intent: ", 1)[1][:1].upper()
                + line.split("WARNING intent: ", 1)[1][1:] + ".", "Working folder matched")
        elif line.startswith("WARNING intent: the staff drive couldn't be read"):
            add("intent", "tech", "The project's working folder couldn't be read this time, so the CVA's last copy "
                "of the tracker was checked; the next run will read the folder again.", "Working folder not read")
        elif m := re.match(r"^intent: NOT CHECKED - data\.wa\.gov is up, but L&I's records there were last updated "
                           r"(.+?), (\d+) days ago", line):
            add("intent", "skipped", f"L&I hasn't updated its public data, which the CVA reads, since "
                f"{_lni_date(m[1])} ({m[2]} days ago), so intents and affidavits filed since then can't be seen "
                "there. L&I's own website may show newer filings. The tracker wasn't checked or changed; it will be "
                "checked once L&I updates its public data.",
                f"L&I public data out of date ({_lni_date(m[1])})")
        elif line.startswith("intent: NOT CHECKED - L&I's live lookup answered, but this project's records couldn't"):
            add("intent", "skipped", "L&I's website didn't return this project's records usably this time, so the "
                "tracker wasn't checked or changed; the next run will try again. The administrator has the details.",
                "L&I website: this project not readable")
        elif line.startswith("intent: NOT CHECKED - L&I's live lookup answered, but L&I's records couldn't be read"):
            add("intent", "skipped", "L&I's website couldn't be read this time, so the tracker wasn't checked or "
                "changed; the administrator has the details.", "L&I website not readable")
        elif line.startswith("intent: NOT CHECKED - L&I's live lookup could not be reached"):
            add("intent", "skipped", "L&I's website couldn't be reached (it seems to be down), so the tracker wasn't "
                "checked or changed; the next run will try again.", "L&I website down")
        elif line.startswith("intent: NOT CHECKED - data.wa.gov answered, but L&I's records couldn't be read"):
            add("intent", "skipped", "L&I's public data couldn't be read this time, so the tracker wasn't checked or "
                "changed; the administrator has the details.", "L&I public data not readable")
        elif line.startswith("intent: NOT CHECKED - data.wa.gov (L&I's public records) could not be reached"):
            add("intent", "skipped", "L&I's public data site couldn't be reached (it seems to be down), so the "
                "tracker wasn't checked or changed; the next run will try again.", "L&I public data site down")
        elif line.startswith("intent: NOT CHECKED - no project folder in the staff drive"):
            add("intent", "skipped", "The CVA couldn't find this project's working folder, so the tracker wasn't "
                "checked. Tell the administrator.", "Staff folder not found")
        elif m := re.match(r"^intent: not applicable \((.*)\)$", line):
            add("intent", "na", *NA_REASON.get(m[1], NA_OTHER))
        elif m := _MARKED.match(line):
            n = _n(m[1])
            no_lni = "WARNING L&I download failed" in line or "could not determine the L&I project id" in line
            if n:
                add("intent", "review", f"{n:,} intent or affidavit {_plural(n, 'entry is', 'entries are')} "
                    + ("missing" if no_lni else "missing or don't match L&I's records") + "; marked in red in the tracker.",
                    f"{n:,} tracker {_plural(n, 'entry', 'entries')} to fix")
            if "project id" in line:            # staff know the number: they can tell the administrator
                add("intent", "review", "The tracker couldn't be compared with L&I's records this time (the L&I "
                    "project number isn't known: tell the administrator which it is).", "L&I project number unknown")
            elif no_lni:                        # L&I's site: the administrator's business, the next run tries again
                add("intent", "tech", "The tracker couldn't be compared with L&I's records this time (L&I's site "
                    "couldn't be reached); the next run will.", "L&I not reachable")
        elif line.startswith(_OVERWRITTEN):
            # someone saved an older copy of the tracker over the CVA's (Nov 6, 2026): a note, it was marked again
            add("intent", "info", "The tracker in the CVA's folder had been replaced by an older copy, which removed "
                "the last check's marks; this check marked it again. Please don't edit the CVA's copy of the tracker.",
                "Tracker replaced by an older copy")
        elif m := _LNI_ONLY.match(line):
            n = _n(m[1])
            add("intent", "review", f"{n:,} {_plural(n, 'contractor')} in L&I's records for this project "
                f"{_plural(n, 'is', 'are')} not in the tracker; {_plural(n, 'it is', 'they are')} listed in the "
                "tracker's Review Log sheet.", f"{n:,} {_plural(n, 'contractor')} not in the tracker")
        elif m := _HIRING.match(line):      # #121 (David, Nov 9, 2026)
            n = _n(m[1])
            add("intent", "review", f"{n:,} {_plural(n, 'line')} in the tracker {_plural(n, 'names', 'name')} a "
                f"Hiring Contractor that isn't the hiring company on the L&I intent; {_plural(n, 'it is', 'they are')} "
                "marked in red and listed at the top of the tracker's Review Log sheet.",
                f"{n:,} hiring contractor {_plural(n, 'mismatch', 'mismatches')}")
        elif line.startswith("WARNING tracker: more than one tracker"):
            add("uploads", "review", "More than one tracker was uploaded: only the newest is used, the others "
                "were archived.", "Several trackers uploaded")
        elif line.startswith("WARNING upload:"):
            unrec += 1
        elif line.startswith("ERROR (upload)"):
            add("uploads", "failed", "The Uploads folder couldn't be sorted, so new files may not have been used. "
                + NOTIFIED, "Uploads not sorted")
        elif line.startswith("ERROR (layout)") and "_project" in (details or {}):
            pass                                            # told once, from details["_project"] below
        elif line.startswith("ERROR (") and not any(line.startswith(f"ERROR ({t})") for t in (details or {})):
            add("project", "failed", "Part of this project couldn't be checked. " + NOTIFIED, "Not fully checked")
    if unrec:
        add("uploads", "review", f"{unrec} uploaded {_plural(unrec, 'file was', 'files were')} not recognized and "
            f"{_plural(unrec, 'is', 'are')} still in Uploads (a payroll, tracker or wage file is expected).",
            f"{unrec} {_plural(unrec, 'file')} not recognized")
    for task, d in (details or {}).items():
        key = "project" if task == "_project" else task
        if d.get("errors"):
            add(key, "failed", ("This project couldn't be checked. " if key == "project" else
                                "A technical problem stopped this task. ") + NOTIFIED,
                "Project problem" if key == "project" else "Technical problem")
        for s in d.get("skipped", []):
            why = s.split(":", 1)[-1].strip()
            why = why[len("SKIPPED - "):] if why.startswith("SKIPPED - ") else why
            short = "Skipped"
            if "old folder layout" in why:
                why = "The project's folders aren't in the expected layout (Uploads, Payroll, Tracker, Files). " \
                      + "Tell the administrator."
                short = "Folders not in expected layout"
            elif "no wage rates file" in why:
                where = ("the last review stands" if "last review stands" in why else "it stays in the Uploads folder")
                why = (f"This project has no wage rates file, so the payroll wasn't checked; {where}. Upload the "
                       "project's wage determination or rate table, and it will be checked.")
                short = "No wage rates file"
            elif "Jurisdiction column not found in the wage rates file" in why:
                where = ("the last review stands" if "last review stands" in why else "it stays in the Uploads folder")
                why = ("Jurisdiction column not found in the wage rates file, so the payroll wasn't checked; "
                       f"{where}. Since Nov 7 the wage rates file needs a Jurisdiction column saying State, Federal "
                       "or Union for each rate (official federal and L&I determinations don't). Upload the wage "
                       "rates again with that column, and the payroll will be checked.")
                short = "No Jurisdiction in wage rates"
            elif "whether federal, state or union rates apply" in why:
                m = re.match(r"([\d,]+) payroll line\(s\) .*?, and (.+?); the payroll was not checked", why)
                n_, problem = (m[1], m[2]) if m else ("Some", "the project's wage rates don't say which type they are")
                lines_ = "payroll line doesn't" if n_ == "1" else "payroll lines don't"
                its = "its" if n_ == "1" else "their"
                where = ("the last review stands" if "last review stands" in why else "it stays in the Uploads folder")
                why = (f"{n_} {lines_} say in {its} Jurisdiction whether federal, state or union rates apply, and "
                       f"{problem}, so the payroll wasn't checked; {where}. Upload the project's wage rates "
                       "with their Jurisdiction column (State, Federal or Union), or fix the payroll's Jurisdiction "
                       "column.")
                short = "Wage type unclear"
            elif "the payroll has no lines below its header row" in why:
                where = ("the last review stands" if "last review stands" in why else "it stays in the Uploads folder")
                why = (f"The payroll has no lines below its header row, so it wasn't checked; {where}. Export the "
                       "payroll again with its lines and upload it.")
                short = "No payroll lines"
            elif "the header row appears again at row" in why:
                m = re.search(r"appears again at row (\d+)", why)
                where = ("the last review stands" if "last review stands" in why else "it stays in the Uploads folder")
                why = (f"The payroll's header row appears again at row {m[1] if m else '?'}, as if two exports were "
                       f"pasted into one sheet, so it wasn't checked; {where}. Upload each export as its own file.")
                short = "Two exports in one file"
            elif "open in Excel" in why:
                why = "The tracker was open in Excel. Close it and run again; nothing was changed."
                short = "Tracker open in Excel"
            add(key, "skipped", why.rstrip(". ") + ".", short)
    order = list(TASK_LABEL)
    ran = [t for t in ("payroll", "wages", "intent") if t in (details or {})]   # an upload can be any of these
    if held_uploads and ran:                                                  # (a tracker: the intent check, M1)
        n, key = held_uploads, ran[0]
        add(key, "review", held_uploads_text(n, held_linked), f"{n} {_plural(n, 'upload')} not checked", first=True)
        rows[key].held_uploads = n
        rows[key].items = [(k, "No other payroll file in the Uploads folder.", "No other payroll")
                           if r.startswith("No payroll file in the Uploads") else (k, r, sh)
                           for k, r, sh in rows[key].items]                  # the held one may be the payroll (L2)
    return sorted(rows.values(), key=lambda r: order.index(r.key) if r.key in order else 99)


def _worst(r: Row) -> int:
    return max((RANK[k] for k in r.kinds), default=0)


def table_rows(report, results_late=False) -> list[Row]:
    """Every project's rows, worst first (failures, things to review, notes, skips; stable within each)."""
    held = set(report.get("held_back_trackers") or [])
    uploads = report.get("held_back_uploads") or {}
    rows = [r for p, lines in (report.get("projects") or {}).items()
            for r in task_rows(lines, (report.get("details") or {}).get(p), results_late, project=p,
                               held_tracker=p in held, held_uploads=uploads.get(p, 0),
                               held_linked=bool(report.get("held_back_link")))]
    return sorted(rows, key=lambda r: -_worst(r))


def merge_cells(main: dict, extra: dict | None) -> dict:
    """One cell for two tasks shown as one (the wage update inside the payroll review): the worse Run (a skipped
    wage update doesn't make the payroll review skipped), Needs review if either found something, and both notes."""
    if not extra:
        return main
    run = main["run"]
    if extra["run"] == FAILED:
        run = FAILED
    if NEEDS_REVIEW in (main["findings"], extra["findings"]):
        findings = NEEDS_REVIEW
    elif main["findings"] == NOT_CHECKED or extra["run"] == FAILED:
        findings = NOT_CHECKED
    else:
        findings = main["findings"]
    useful = extra["run"] == FAILED or extra["findings"] == NEEDS_REVIEW     # "No new wage files" adds nothing
    note = "; ".join(x for x in (main["note"], extra["note"] if useful else "") if x)
    return {"run": run, "findings": findings, "note": note}


def cell_results(report) -> dict[tuple[str, str], dict]:
    """{(project, task key): {run, findings, note}} for every task asked for in every project, the way the emails
    judge them (the app's Results tab shows the same). note: the staff's few words (empty when nothing to say)."""
    out = {}
    rows = {(r.project, r.key): r for r in table_rows(report)}
    for p, per in (report.get("details") or {}).items():
        for t in per:
            if t.startswith("_"):
                continue
            r = rows.get((p, t))
            proj = rows.get((p, "project"))
            if proj and proj.run == FAILED and (not r or r.findings != NEEDS_REVIEW):
                r = proj                                 # the project itself couldn't be prepared (as in the emails)
            out[(p, t)] = {"run": r.run if r else COMPLETED, "findings": r.findings if r else NO_ISSUES,
                           "note": r.short() if r else ""}
    return out


def findings_verdict(rows, unreached=()) -> str:
    """Both emails' Findings, from the grid itself, so the two agree, with one difference (David, Nov 12, 2026):
    files held back in a project's Uploads/ make that project's row Needs review in the staff email only; the
    administrator email has the whole-drive held-back line instead. Needs review if anything was found;
    Not checked if something failed, was never reached, or was skipped and needs someone; else No issues."""
    if any(r.findings == NEEDS_REVIEW for r in rows):
        return NEEDS_REVIEW
    if unreached or any(r.run == FAILED or (r.unchecked() and r.run == SKIPPED) for r in rows):
        return NOT_CHECKED
    return NO_ISSUES


def run_verdict(report, rows, extra_failures=()) -> str:
    """The administrator's verdict on the run itself: Failed (a task or a step failed), Completed, with warnings
    (something to look at before it breaks), or Completed."""
    from .notify import email_level
    if email_level(report, extra_failures)[0] == "Critical" or any(r.run == FAILED for r in rows):
        return FAILED
    details = report.get("details") or {}
    warned = any("tech" in r.kinds for r in rows) or report.get("pipeline_warnings") or report.get("run_warnings") \
        or any(d.get("warnings") for per in details.values() for d in per.values()) \
        or any(l.startswith("WARNING") and not l.startswith(("WARNING upload:", "WARNING tracker:"))
               for lines in (report.get("projects") or {}).values() for l in lines)
    return f"{COMPLETED}, with warnings" if warned else COMPLETED


SYMBOL = {NO_ISSUES: "✓", NEEDS_REVIEW: "!", NOT_CHECKED: "–", NOT_APPLICABLE: "–", FAILED: "✕"}
FAILED_COLOR = "#b3261e"


def render_grid(tasks, projects, rows, links=None, admin=False, unreached=(), plan=None):
    """(text lines, html): a grid of Findings (one line per project, one column per task run: No issues, Needs
    review, or Not checked, marked ✕ when the task failed), an "At a glance" column on the right (the project's
    issues in a few words, worst first), then "What to do": each reason and suggested action, numbered.
    admin: also show the technical warnings. links: {reason: (label, url)} shown after a reason. unreached: projects
    the run never got to (it stopped first): Not checked. plan: {project: its tasks} when each project had its own
    tasks (report["plan"]): a task not asked for that project is a blank cell."""
    e = html.escape
    links = links or {}
    keys = list(tasks)
    cell = {(r.project, r.key): r for r in rows}
    lines = ([ALL] if any(r.project == ALL for r in rows) else []) + list(projects)
    def joined(parts):
        parts = list(dict.fromkeys(x for x in parts if x))
        return "; ".join([parts[0]] + [x[:1].lower() + x[1:] if x[1:2].islower() else x for x in parts[1:]]) \
            if parts else ""
    glance = {p: joined(s for r in rows if r.project == p and not _plain_skip(r, admin) for s in r.short(admin).split("; "))
              for p in lines}
    todo = [(r.project, r.task, r.reason(admin)) for r in rows if r.needs_action(admin) and r.reason(admin)]

    def word(p, key):
        if p == ALL:
            return "", ""
        if plan and p in plan and key not in plan[p]:      # not asked for this project in this run
            return "", ""
        if p in unreached:                                  # the run stopped before this project
            return NOT_CHECKED, SYMBOL[NOT_CHECKED]
        r = cell.get((p, key))
        proj = cell.get((p, "project"))
        if proj and proj.run == FAILED and (not r or r.findings != NEEDS_REVIEW):
            r = proj                                         # the project itself couldn't be prepared
        if not r:
            return NO_ISSUES, SYMBOL[NO_ISSUES]
        if _plain_skip(r, admin):                     # "Skipped · No new wage files / No new payroll", not "– Not checked"
            return f"{SKIPPED} · {r.short(admin)}", ""
        return r.findings, SYMBOL[FAILED] if r.run == FAILED else SYMBOL[r.findings]

    # ---- text
    text = []
    labels = [TASK_LABEL.get(t, t) for t in keys]
    if lines and labels:
        w0 = max(len("Project"), *(len(p) for p in lines))
        cells_of = {p: [(f"{word(p, k)[1]} {word(p, k)[0]}".strip() if word(p, k)[0] else "") for k in keys]
                    for p in lines}
        # each column as wide as its label or its widest cell ("Skipped · No payroll uploaded": review of #141)
        w = [max(len(lb), 14, *(len(cells_of[p][i]) for p in lines)) for i, lb in enumerate(labels)]
        text.append(f"{'Project':<{w0}}  " + "  ".join(f"{lb:<{w[i]}}" for i, lb in enumerate(labels)) + "  At a glance")
        for p in lines:
            cells = cells_of[p]
            text.append(f"{p:<{w0}}  " + "  ".join(f"{c:<{w[i]}}" for i, c in enumerate(cells))
                        + (f"  {glance[p]}" if glance[p] else ""))
    if todo:
        text += ["", "What to do:"]
        text += [f"  {i}. {p} · {t}: {r}" + (f" {links[r][0]}: {links[r][1]}" if r in links else "")
                 for i, (p, t, r) in enumerate(todo, 1)]

    # ---- HTML
    th = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #c9d1cd;font-size:12px;color:#55605b"'
    td = 'style="padding:6px 10px;border-bottom:1px solid #e3e8e5;vertical-align:top;white-space:nowrap"'
    tdg = 'style="padding:6px 10px;border-bottom:1px solid #e3e8e5;vertical-align:top;color:#33403a"'
    h = []
    if lines and labels:
        h += ['<table style="border-collapse:collapse;font-size:13px">',
              f"<tr><th {th}>Project</th>" + "".join(f"<th {th}>{e(lb)}</th>" for lb in labels)
              + f"<th {th}>At a glance</th></tr>"]
        for p in lines:
            h.append(f"<tr><td {td}><b>{e(p)}</b></td>")
            for k in keys:
                w_, sym = word(p, k)
                if not w_:
                    h.append(f"<td {td}></td>")
                    continue
                color = FAILED_COLOR if sym == SYMBOL[FAILED] else COLOR.get(w_, COLOR[NOT_CHECKED])
                h.append(f'<td {td}><b style="color:{color}">{e(f"{sym} {w_}".strip())}</b></td>')
            h.append(f"<td {tdg}>{e(glance[p])}</td></tr>")
        h.append("</table>")
    if todo:
        h.append('<p style="margin:18px 0 6px;font-weight:bold">What to do</p><ol style="margin:0;padding-left:22px">')
        for p, t, r in todo:
            link = f' <a href="{e(links[r][1])}">{e(links[r][0])}</a>' if r in links else ""
            h.append(f'<li style="margin-bottom:6px"><b>{e(p)} · {e(t)}</b>: {e(r)}{link}</li>')
        h.append("</ol>")
    return text, "\n".join(h)


# A payroll review with nothing to do: no new payroll since the last review, or none uploaded yet (#141)
PAYROLL_PLAIN_SKIPS = {"No new payroll", "No payroll uploaded"}


def _plain_skip(r: Row, admin=False) -> bool:
    """A task with nothing to do: its cell says so ("Skipped · No new wage files", "Skipped · No new payroll") instead
    of "Not checked", which read as if something had gone wrong (David, Nov 7, 2026; payroll: #141, Nov 9), and its
    words leave At a glance. A wage update: any skip nobody has to act on. A payroll review: only its two normal
    skips (every other payroll skip needs someone, or means the review couldn't be done)."""
    if r.run != SKIPPED or r.needs_action(admin) or not r.items:
        return False
    if r.key == "wages":
        return True
    return r.key == "payroll" and all(k == "skipped" and s in PAYROLL_PLAIN_SKIPS for k, _r, s in r.items)


def _whole_run(kind, task, reason, short) -> Row:
    r = Row(ALL, task)
    r.items.append((kind, reason, short))
    return r


def email_rows(report, level, extra_failures=(), held_back=None, admin=False):
    """Everything both emails judge: the whole run's rows (didn't start or finish, results late, files held back)
    and every project's. -> (rows worst first, problem or None, late keys, links, unreached projects).
    level: notify.email_level()'s word (Critical: the run didn't do its job); held_back: (count, url or None)."""
    from .notify import TOKENIZER_WARNING, email_level
    late = [k for k in LATE if any(w == TOKENIZER_WARNING.get(k) for w in report.get("pipeline_warnings") or [])]
    steps = [f.get("step") for f in extra_failures if f.get("step")]
    if steps and all(x in STAFF_SILENT_STEPS for x in steps) and len(steps) == len(extra_failures):
        level, steps = email_level(report)[0], []          # only the administrator's copy failed: judged as if fine
    problem = None
    if level == "Critical":
        problem = STEP_STAFF.get(steps[0], GENERIC_PROBLEM) if steps else GENERIC_PROBLEM
    elif report.get("run_errors"):
        problem = GENERIC_PROBLEM
    # what happened to the run as a whole, in a few words: only a step that stopped it says it didn't start or finish
    stopped = STEP_SHORT.get(steps[0], "") if problem and steps else ""
    rows = []
    if problem:
        short = stopped[0].upper() + stopped[1:] if stopped else "Technical problem"
        rows.append(_whole_run("failed", "run", problem, short.replace("The run", "Run")))
    rows += [_whole_run("info", "results", LATE[k],
                        "Results arrive later" if k == "after" else "Newest uploads may wait") for k in late]
    links = {}
    hb_n, hb_url = held_back or (0, None)
    if hb_n and admin:          # staff: only what touched this run's own tasks, on those projects (David, Nov 12)
        hb = (f"{hb_n} new {_plural(hb_n, 'file was', 'files were')} not checked (for example, an SSN in a file name, "
              "or a payroll saved as a PDF). " + ("The list says which and how to fix them." if hb_url else
                                                  "The list \"Held back from AWS\" at the top of the shared drive "
                                                  "says which and how to fix them."))
        rows.append(_whole_run("review", "uploads", hb, f"{hb_n} {_plural(hb_n, 'file')} not checked"))
        if hb_url:
            links[hb] = ("Open the list", hb_url)
    project_rows = table_rows(report, results_late="after" in late)
    if stopped == STEP_SHORT["push"] and not admin:
        # staff: no project's findings (they may not be on the drive); the administrator keeps them, and the whole
        # run's line above says once that the results may not have been saved
        project_rows = []
        for p in report.get("projects") or {}:
            r = Row(p, "project")
            r.items.append(("failed", "", "Results not saved"))
            project_rows.append(r)
    if hb_url:
        for r in project_rows:              # the link goes after the row's whole reason (render_grid)
            if r.held_uploads:
                links[r.reason(admin)] = ("Open the list", hb_url)
    rows += project_rows
    details = report.get("details") or {}
    # a run that didn't start or finish: the projects it holds nothing for were never reached
    unreached = [p for p in (report.get("projects") or {}) if not details.get(p) and not report["projects"][p]] \
        if stopped.startswith("the run") else []
    return sorted(rows, key=lambda r: -_worst(r)), problem, late, links, unreached, stopped


def build_staff_message(report, ncfg, root, level, extra_failures=(), produced=(), held_back=None,
                        news=(), when=""):
    """(subject, text, html) of the staff email. level: notify.email_level()'s word (Critical: the run didn't do its
    job); held_back: (count, url or None)."""
    projects = report.get("projects") or {}
    rows, problem, late, links, unreached, stopped = email_rows(report, level, extra_failures, held_back)

    word = NOT_CHECKED if stopped == STEP_SHORT["push"] else findings_verdict(rows, unreached)
    review = list(dict.fromkeys(r.project for r in rows if r.findings == NEEDS_REVIEW and r.project != ALL))
    unchecked = list(dict.fromkeys([r.project for r in rows if r.project != ALL and r.unchecked()
                                    and r.findings == NOT_CHECKED and r.project not in review]
                                   + [p for p in unreached if p not in review]))
    whole_run = any(r.project == ALL and r.needs_action() for r in rows)
    n = len(projects)
    def some(names):
        return names[0] if len(names) == 1 and n == 1 else f"{len(names)} of {n} projects" if n > 1 else names[0]
    # the run itself, in the staff's words (#77); the summary after it no longer repeats it
    not_saved = stopped == STEP_SHORT["push"]
    if not_saved and not projects:
        run_words = RUN_PROBLEM                # no summary: the run itself didn't get far enough to say it completed
    elif stopped in RUN_WORDS:
        run_words = RUN_WORDS[stopped]
    elif problem:
        run_words = RUN_PROBLEM
    else:
        run_words = RUN_COMPLETED
    if problem and not projects:
        summary = "nothing was checked"
    elif review:
        summary = some(review) + (f"; {len(unchecked)} not checked" if unchecked else "")
    elif unchecked:
        summary = some(unchecked)
    elif whole_run:
        summary = "see What to do"
    else:
        summary = (next(iter(projects)) if n == 1 else f"{n} projects checked") if projects else "nothing to check"
    subject = f"{ncfg.get('subject_prefix', '[Compliance Validation Agent]')} {run_words} · {word}: {summary}".strip()
    tasks = [t for t in GRID_ORDER if t in (report.get("tasks") or [])]
    if problem and not projects:
        lead = (f"{stopped[0].upper() + stopped[1:]}, so nothing was checked." if stopped.startswith("the run")
                else "Nothing was checked.")         # say it stopped only when a step stopped it
    elif not_saved:
        lead = f"Results not saved for {n} {_plural(n, 'project')}: skip the review for now."
    elif review or unchecked:
        parts = ([f"{len(review)} {_plural(len(review), 'needs', 'need')} review"] if review else []) + \
                ([f"{len(unchecked)} couldn't be fully checked"] if unchecked else [])
        lead = f"{n} {_plural(n, 'project')} checked: " + " and ".join(parts) + "."
    elif whole_run:
        lead = f"{n} {_plural(n, 'project')} checked. See What to do."
    else:
        lead = f"{n} {_plural(n, 'project')} checked: no issues found." if projects else "Nothing to check."
    grid_text, grid_html = render_grid(tasks, list(projects), rows, links, unreached=unreached,
                                       plan=report.get("plan"))

    # ---- plain text
    t = [f"{run_words.upper()} · {word.upper()}", lead, ""] + grid_text + [""]
    if news:
        t.append("New")
        t += [f"  - {title}: {text}" for _d, title, text in news]
        t.append("")
    if when:
        t.append(f"Checked: {when}")
    text = "\n".join(t).rstrip() + "\n"

    # ---- HTML
    e = html.escape
    c = COLOR[word]
    run_c = COLOR[NO_ISSUES] if run_words == RUN_COMPLETED else FAILED_COLOR
    h = ['<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1c1f1d;max-width:900px">',
         f'<p style="font-size:20px;font-weight:bold;margin:0 0 4px"><span style="color:{run_c}">{e(run_words)}</span>'
         f' <span style="color:#7b8a84">·</span> <span style="color:{c}">{e(word)}</span></p>',
         f'<p style="font-size:14px;margin:0 0 12px;color:#55605b">{e(lead)}</p>', grid_html]
    if news:
        h.append('<p style="margin:18px 0 4px;font-weight:bold">New</p><ul style="margin:0;padding-left:20px">')
        h += [f"<li><b>{e(title)}</b>: {e(text)}</li>" for _d, title, text in news]
        h.append("</ul>")
    if when:
        h.append(f'<p style="font-size:12px;color:#7b8a84;margin-top:18px">Checked: {e(when)}</p>')
    h.append("</div>")
    return subject, text, "\n".join(h)
