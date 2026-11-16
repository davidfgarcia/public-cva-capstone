"""
The checks a file must pass before it is copied to AWS, on top of the file rules (tokenizer_config.yaml, files:).
A file that fails one is *held back*: it stays in the staff drive, untouched, and is listed with the reason in
"Held back from AWS.csv" at the top of the staff folder (see mirror.py). Nothing here ever changes a file.

    1  file type               the file rules (tokenize / scrub_text / scan_pdf); anything else is not copied
    2  file and folder names   no SSN, email address or phone number (name_reason); workers' names in them are
                               replaced by their tokens on the AWS side (tokenize_path)
    3  is it a file the agent uses?
                               PDFs only: under reference/, or, in a project, a wage determination by its first
                               page and not a certified payroll (pdf_relevant). Spreadsheets and text files always
                               pass: their personal data is replaced in check 4a, and the agent lists any upload it
                               doesn't recognise.
    4a spreadsheets and text   tokenized, then the result is scanned again; anything left holds it back (residual)
    4b PDFs                    one revision only, no attached files, no notes or other annotations but links, readable
                               text on every page that has a picture (no scans), and no SSN or worker's name in the
                               text, form fields, bookmarks or properties (screen_pdf)

"A worker's name" is one the tokenizer has learnt from a payroll. Names and contact details of anyone else (union
reps, officials, contractors' offices) are not looked for in PDFs: reference documents are full of them.
"""
from __future__ import annotations

import html
import io
import posixpath
import re
import zipfile

from .tokens import TOKEN_HEX

_ANY_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9])[A-Z]{2,5}-[0-9a-f]{%d}(?![0-9a-f])" % TOKEN_HEX)

# Which project PDFs may go to AWS. Stricter than the agent's own classification (agent/uploads.py), which also
# counts "Department of Labor" or a file name such as "...wages.pdf": a certified payroll (WH-347) has both, and
# its subcontractors' workers are never in the firm's own payrolls, so their names wouldn't be recognised.
WD_MARK = re.compile(r"General\s*Decision\s*(?:Number|No\.?)|Wage\s*Determination|Prevailing\s*Wage\s*Rates?|"
                     r"Journey\s*Level|Collective\s*Bargaining\s*Agreement|Wage\s*(?:and|&)\s*Fringe|"
                     r"Fringe\s*Benefit\s*Rates?|Wage\s*Rates?\b", re.I)
PAYROLL_MARK = re.compile(r"WH[-\s]?347|Statement\s*of\s*Compliance|Certified\s*Payroll|Payroll\s*(?:No|Number)\b|"
                          r"Social\s*Security|\bSSN\b|Employee\s*(?:Name|ID)|Net\s*(?:Pay|Wages)|Gross\s*(?:Pay|Amount)",
                          re.I)
# What a held-back reason names for each PAYROLL_MARK alternative (fixed words, never the document's own text).
_PAYROLL_LABELS = ((r"WH[-\s]?347", "a WH-347 form"), (r"Statement\s*of\s*Compliance", "a Statement of Compliance"),
                   (r"Certified\s*Payroll", "a certified payroll"), (r"Payroll\s*(?:No|Number)", "a payroll number"),
                   (r"Social\s*Security|SSN", "Social Security numbers"), (r"Employee\s*(?:Name|ID)", "employee details"),
                   (r"Net\s*(?:Pay|Wages)|Gross\s*(?:Pay|Amount)", "pay amounts"))


def _payroll_label(matched: str) -> str:
    for pattern, label in _PAYROLL_LABELS:
        if re.fullmatch(pattern, matched, re.I):
            return label
    return "payroll details"

# The two files the tokenizer itself writes at the top of the folders; never copied either way (mirror.py).
HELD_BACK_SHEET = "Held back from AWS.csv"            # staff side: which files and why (real names)
HELD_BACK_SUMMARY = "_held_back_from_aws.json"         # AWS side: how many, as anonymous keys (no names)
# AWS side: the Drive MD5s of copies renamed, moved or taken back because of a worker's name learnt after they reached
# AWS (no path, no name): the AWS run deletes the old S3 versions with that content (deploy/purge_stale.py)
PURGE_SUMMARY = "_purge_on_aws.json"
# Change 1 (#93, pii_tokenizer/staff_trackers.py): where each project's tracker came from; staff side with names,
# AWS side codes only
TRACKER_SOURCES_SHEET = "Tracker sources.csv"
TRACKER_SOURCES_SUMMARY = "_tracker_sources.json"
# AWS side, written by the AWS pipeline itself (deploy/run_pipeline.sh) between its pull and its push: staff trackers
# aren't copied over the one it is checking meanwhile
AWS_RUN_MARKER = "_aws_run_in_progress.json"
# AWS side: for each copy, the Drive ID of the staff file it stands for (IDs only, no name), so the administrator's
# email can link to the staff's own file, never to the AWS-side copy (David, Nov 15, 2026; mirror._publish_links)
STAFF_LINKS = "_staff_links.json"
OWN_FILES = (HELD_BACK_SHEET, HELD_BACK_SUMMARY, PURGE_SUMMARY, TRACKER_SOURCES_SHEET, TRACKER_SOURCES_SUMMARY,
             AWS_RUN_MARKER, STAFF_LINKS)


def is_own_file(rel: str) -> bool:
    """The tokenizer's two lists, and any copy staff make of them ("Held back from AWS (1).csv", a Google Sheet):
    never copied either way, wherever they are."""
    base = posixpath.basename(rel).casefold()
    return base.startswith(("held back from aws", "_held_back_from_aws", "tracker sources", "_tracker_sources",
                            "_aws_run_in_progress", "_staff_links"))

MAX_PDF_PAGES = 600                 # longer PDFs are held back rather than checked half-way
MAX_PDF_MB = 150                    # larger ones too (memory)
MIN_PAGE_TEXT = 20                  # characters: below this, a page with an image counts as a scan


class Held(Exception):
    """The file must not go to AWS; str(e) says why, in a few plain words (shown to staff)."""


class Found(Held):
    """Held because worker data (a name, an SSN, contact details) was found in the file's content."""


# ------------------------------------------------------------------------------------------------ 2. names
_SEP_RE = re.compile(r"[^\s_.,;()\[\]+&-]+")


def _words(text: str) -> list[str]:
    return [w for _, _, w in _word_spans(text)]


def _word_spans(text: str) -> list[tuple[int, int, str]]:
    """The words of a file or folder name with where they are: "Ana_Rivera", "rivera-ana" and "RiveraAna" are
    each two words. (start, end, lower-case word)."""
    out = []
    for m in _SEP_RE.finditer(text):
        seg, base = m.group(0), m.start()
        cuts = [0] + [k for k in range(1, len(seg)) if "a" <= seg[k - 1] <= "z" and "A" <= seg[k] <= "Z"] + [len(seg)]
        out += [(base + i, base + j, seg[i:j].casefold()) for i, j in zip(cuts, cuts[1:])]
    return out


def tokenize_name(part: str, scrubber) -> str:
    """One file or folder name with every learnt worker's name in it replaced by its token, however it is written
    (Scrubber.name_places): "Ana Rivera timesheet.xlsx" -> "EMP-3f2a91c0 timesheet.xlsx". Nothing else changes."""
    out, pos = [], 0
    for start, end, tok in scrubber.name_places(part):
        before = " " if start and part[start - 1].isalnum() else ""        # "WeekAnaRivera" -> "Week EMP-...":
        after = " " if end < len(part) and part[end].isalnum() else ""     # a token glued to letters can't be read
        out += [part[pos:start], before, tok, after]
        pos = end
    return "".join(out) + part[pos:]


def name_left(rel: str, scrubber, folder: bool = False) -> str | None:
    """A known worker's name still in a path after tokenize_path (written in a way it couldn't replace), or None.
    Such a file is held back: no name ever reaches AWS."""
    parts = rel.split("/")
    for i, part in enumerate(parts):
        if _is_project_folder(parts, i, folder):
            continue
        if scrubber.finds(part, names=True) == "a worker's name" or scrubber.name_places(part) \
                or scrubber.name_in_words(_words(part)) or scrubber.name_letters_in(part):
            last = i == len(parts) - 1 and not folder
            return (f"its {'file' if last else 'folder'} name contains a worker's name written in a way that "
                    "can't be replaced by its token; rename it and it will be copied")
    # across folder levels too ("Ana/Rivera.xlsx"), tokens left out
    rest = [p for i, p in enumerate(parts) if not _is_project_folder(parts, i, folder) and
            not (i == 0 and len(parts) > 1 and parts[0].casefold() == "projects")]
    if scrubber.name_letters_in(" ".join(_ANY_TOKEN_RE.sub(" ", p) for p in rest)):
        return ("its folder and file names together spell a worker's name; rename them and it will be copied")
    return None


def _is_project_folder(parts: list[str], i: int, folder: bool) -> bool:
    return i == 1 and parts[0].casefold() == "projects" and (folder or len(parts) > 2)


def tokenize_path(rel: str, scrubber, folder: bool = False) -> str:
    """Where a staff file or folder lives on the AWS side: every worker's name in its file and folder names
    replaced by its token, except in a project's own folder (Projects/<project>): projects are often named after
    people ("Cesar Chavez Park - Phase 2"), and the project name is business information."""
    parts = rel.split("/")
    return "/".join(p if _is_project_folder(parts, i, folder) else tokenize_name(p, scrubber)
                    for i, p in enumerate(parts))


def has_name(rel: str, scrubber, folder: bool = False) -> bool:
    """True if a worker's name is in the path as it is (tokenize_path would change it)."""
    return tokenize_path(rel, scrubber, folder) != rel


def name_reason(copy_rel: str, scrubber, folder: bool = False) -> str | None:
    """Why a path can't go to AWS even with names replaced: an SSN, email address or phone number in a file or
    folder name (project folders included), or None. Workers' names are not a reason: tokenize_path replaces them.
    ``folder``: the path is a folder's."""
    parts = [p for p in copy_rel.split("/") if p]
    for i, part in enumerate(parts):
        last = i == len(parts) - 1 and not folder
        stem = posixpath.splitext(part)[0] if last else part                     # "call 206-555-0100.xlsx"
        found = scrubber.finds(part, contacts=True, names=False) or scrubber.finds(stem, contacts=True, names=False)
        if found:
            return f"its {'file' if last else 'folder'} name contains {found}; rename it and it will be copied"
    return None


# ------------------------------------------------------------------------------------------------ 3. relevance
def pdf_relevant(rel: str, first_page_text: str) -> str | None:
    """None if the agent reads this PDF, else why not. reference/ holds documents the firm chose for the agent
    (CBAs, L&I guidance, contracts); in a project, only wage determinations are read (as the agent decides)."""
    top = rel.split("/", 1)[0].lower()
    if top == "reference":
        return None
    if top == "projects":
        first = first_page_text or ""
        wd, payroll = WD_MARK.search(first), PAYROLL_MARK.search(first)
        if wd and not payroll:
            return None
        if payroll:
            # a fixed label, never the text as written: the match can sit inside a name ("Bennet Payne": "net Pay")
            said = _payroll_label(payroll.group(0))
            also = (" (it also mentions wage rates, but files with payroll details can't go to AWS as PDFs)"
                    if wd else "")
            return (f"its first page looks like a payroll (it mentions {said}){also}; in a project, only wage "
                    "determinations go to AWS as PDFs; upload payrolls as .xlsx or .csv exports instead")
        return ("its first page isn't recognised as a wage determination, and other PDFs in a project aren't sent "
                "to AWS; if it is one, upload the original from SAM.gov or L&I, the LCPtracker Excel export, or the "
                "union's rate table as a spreadsheet, instead of a PDF that was printed, scanned or converted")
    return "the agent doesn't read PDFs here (only in reference/ and, for wage determinations, in projects)"


# ------------------------------------------------------------------------------------------------ 4a. residual
def _xlsx_texts(data: bytes):
    """Every piece of text in a workbook: cells, shared strings, formulas, comments, sheet and defined names,
    headers and footers, and the properties (author, title)."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for n in z.namelist():
            if not n.lower().endswith((".xml", ".vml")):
                continue
            xml = z.read(n).decode("utf-8", errors="replace")
            for m in re.finditer(r'(?:>([^<]+)<)|(?:\b(?:name|ref|display|tooltip)="([^"]*)")', xml):
                t = m.group(1) if m.group(1) is not None else m.group(2)
                if t and t.strip():
                    yield html.unescape(t)


def residual(name: str, data: bytes, scrubber) -> str | None:
    """After tokenizing: why the result still can't go to AWS (worker data left in it), or None."""
    ext = posixpath.splitext(name)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        texts = _xlsx_texts(data)
    else:
        from .files import _decode                 # UTF-16 and Windows-1252 too, as when it was tokenized
        texts = _decode(data)[0].splitlines()
    for t in texts:
        found = scrubber.finds(t)
        if found:
            return f"{found} was still in it after tokenizing"
    return None


# ------------------------------------------------------------------------------------------------ 4b. PDFs
def _xobjects_hold_image(xo, depth=0) -> bool:
    xo = xo.get_object() if hasattr(xo, "get_object") else xo
    for k in xo or {}:
        o = xo[k].get_object()
        if o.get("/Subtype") == "/Image":
            return True
        if o.get("/Subtype") == "/Form" and depth < 3:
            res = o.get("/Resources")
            res = res.get_object() if hasattr(res, "get_object") else res
            if res and res.get("/XObject") and _xobjects_hold_image(res["/XObject"], depth + 1):
                return True
    return False


def _has_image(page) -> bool:
    """A picture on the page: an image object (also inside a form), or an inline image in its drawing."""
    try:
        res = page.get("/Resources") or {}
        res = res.get_object() if hasattr(res, "get_object") else res
        if res.get("/XObject") and _xobjects_hold_image(res["/XObject"]):
            return True
        c = page.get("/Contents")
        c = c.get_object() if hasattr(c, "get_object") else c
        streams = [x.get_object() for x in c] if isinstance(c, list) else [c]     # one drawing, or several
        return any(st is not None and re.search(rb"(?:^|\s)BI\s", st.get_data()) for st in streams)
    except Exception:                       # noqa: BLE001 - can't tell: assume it has one (held back if no text)
        return True


def _outline_titles(items):
    for it in items or []:
        if isinstance(it, list):
            yield from _outline_titles(it)
        else:
            title = getattr(it, "title", None) or (it.get("/Title") if hasattr(it, "get") else None)
            if title:
                yield str(title)


def pdf_texts(data: bytes) -> list[str]:
    """The text of each page, then the form fields' values and the document properties. Raises Held for a PDF
    that can't be checked: unreadable, password-protected, with attached files, too long, or scanned pages."""
    if len(data) > MAX_PDF_MB * 1_000_000:
        raise Held(f"it is {len(data) // 1_000_000} MB, larger than the {MAX_PDF_MB} MB that are checked")
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted and not reader.decrypt(""):
            raise Held("it is protected by a password, so it can't be checked")
        if reader.attachments:
            raise Held("it has files attached inside it, which can't be checked")
        # one end marker per saved version; a "web-optimized" (linearized) PDF has two from the start
        if data.count(b"%%EOF") > (2 if b"/Linearized" in data[:4096] else 1):
            raise Held("it was edited after it was made and keeps the earlier versions inside it, which can't be "
                       "checked; save it again as a new PDF (Print > Save as PDF) and upload that")
        n = len(reader.pages)
        if n > MAX_PDF_PAGES:
            raise Held(f"it has {n} pages, more than the {MAX_PDF_PAGES} that are checked")
        out = []
        for i, page in enumerate(reader.pages, start=1):
            for a in page.get("/Annots") or []:
                if a.get_object().get("/Subtype") not in ("/Link", "/Widget"):   # links; form fields (scanned below)
                    raise Held(f"page {i} has notes or other annotations, which can't be checked; save it again "
                               "as a new PDF (Print > Save as PDF) and upload that")
            text = page.extract_text() or ""
            if len(text.strip()) < MIN_PAGE_TEXT and _has_image(page):
                raise Held(f"page {i} is a scan or a picture, so it can't be checked")
            out.append(text)
        for f in (reader.get_fields() or {}).values():
            v = f.get("/V") if hasattr(f, "get") else None
            if v:
                out.append(str(v))
        for v in (reader.metadata or {}).values():
            if v:
                out.append(str(v))
        out.extend(_outline_titles(reader.outline))
        xmp = reader.trailer["/Root"].get_object().get("/Metadata")
        if xmp is not None:
            raw = xmp.get_object().get_data().decode("utf-8", errors="replace")
            out.extend(html.unescape(t) for t in re.findall(r">([^<]+)<", raw) if t.strip())
        return out
    except Held:
        raise
    except Exception as e:                  # noqa: BLE001 - anything unexpected: hold back rather than leak
        raise Held(f"it can't be read as a PDF ({type(e).__name__})") from e


def screen_pdf(rel: str, data: bytes, scrubber) -> None:
    """Raises Held unless the PDF may go to AWS as it is (checks 3 and 4b)."""
    texts = pdf_texts(data)
    if not any(t.strip() for t in texts):
        raise Held("it has no readable text, so it can't be checked")
    why = pdf_relevant(rel, texts[0])
    if why:
        raise Held(why)
    for t in texts:
        found = scrubber.finds(t)
        if found:
            raise Found(f"it contains {found}; PDFs can't be tokenized")
