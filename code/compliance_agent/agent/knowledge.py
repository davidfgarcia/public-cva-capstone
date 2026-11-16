"""The knowledge base the AI review searches for evidence (agent/ai_review.py, agent/compliance_rag.py):

1. Public regulations, downloaded and refreshed automatically (REGULATIONS below): 29 CFR Part 5 from ecfr.gov,
   RCW 39.12 and WAC 296-127 from app.leg.wa.gov. Kept in the agent's cache folder (ai_review.cache_dir), refreshed
   every ai_review.refresh_days; if a download fails, the copy already saved is used.
2. Documents staff put in reference/knowledge/ on the shared drive (.pdf, .txt, .md): CBAs, L&I guidance, contract
   clauses. The folder names say what they are (see document_chunks).
3. Staff decisions on earlier AI-review patterns: the Decision column of the "Review decisions" sheet in the
   reviewed payrolls (Keep / Dismiss, with a note). Each becomes a past case.

Every passage becomes a compliance_rag.Chunk. Nothing here sends anything anywhere: the downloads only fetch
public text."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from pathlib import Path

from .util import err_text, log

try:                                  # the library sits next to this module
    from . import compliance_rag as R
except ImportError:                   # pragma: no cover
    import compliance_rag as R

# ------------------------------------------------------------------------------------------- 1. regulations
REGULATIONS = [
    {"id": "CFR-29-5", "title": "29 CFR Part 5 (Davis-Bacon labor standards provisions)", "wage_source": "federal",
     "kind": "ecfr", "cfr_title": 29, "part": 5, "cite": r"§\s*5\.\d+[a-z]?"},
    {"id": "RCW-39-12", "title": "RCW 39.12 (Prevailing wages on public works)", "wage_source": "state",
     "kind": "html", "url": "https://app.leg.wa.gov/RCW/default.aspx?cite=39.12&full=true",
     "cite": r"RCW\s+39\.12\.\d{3}[a-z]?"},
    {"id": "WAC-296-127", "title": "WAC 296-127 (Prevailing wage rules)", "wage_source": "state",
     "kind": "html", "url": "https://app.leg.wa.gov/WAC/default.aspx?cite=296-127&full=true",
     "cite": r"(?:WAC\s+)?296-127-\d{3,4}[a-z]?", "cite_prefix": "WAC "},
]
ECFR = "https://www.ecfr.gov/api/versioner/v1"
UA = {"User-Agent": "cva-capstone (prevailing wage review; public regulations)"}


def html_to_text(html: str) -> str:
    """Plain text with one line per block element (enough for the WA legislature pages)."""
    s = re.sub(r"(?is)<(script|style|head|nav|footer)\b.*?</\1>", " ", html)
    s = re.sub(r"(?i)<br\s*/?>|</(p|div|h\d|li|tr|table|section|span class=\"?hdr)[^>]*>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    import html as h
    s = h.unescape(s)
    s = re.sub(r"[ \t\xa0]+", " ", s)
    return "\n".join(line.strip() for line in s.splitlines() if line.strip())


def split_sections(text: str, cite_re: str, *, min_len=40) -> list:
    """[(citation, body)]: the text cut at each heading line: a short line (at most 160 characters) that starts with
    the citation (e.g. 'RCW 39.12.020'), optionally after punctuation or a 'PDF' link, followed by nothing or a
    capitalised title; a sentence that merely mentions a section ('See RCW 39.12.020 for...') is not a heading. The same citation appearing again (a table of contents, then the
    section) keeps the longest body."""
    pat = re.compile(rf"(?m)^(?=[^\n]{{0,160}}$)[^A-Za-z\n]{{0,12}}(?:PDF[^A-Za-z\n]{{0,6}})?({cite_re})\b"
                     rf"(?=[ \t]*$|[ \t]+[A-Z(\[\d])")
    ms = list(pat.finditer(text))
    out = {}
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
        cite = re.sub(r"\s+", " ", m.group(1)).strip()
        body = text[m.start():end].strip()
        if len(body) >= min_len and len(body) > len(out.get(cite, "")):
            body = body[m.start(1) - m.start():]                    # from the citation on
            out[cite] = body
    return list(out.items())


def ecfr_sections(xml: str) -> list:
    """[(citation, body)] from an eCFR 'full' XML part: one per SECTION (DIV8)."""
    from xml.etree import ElementTree as ET
    root = ET.fromstring(xml)
    out = []
    for div in root.iter():
        if div.tag.upper().startswith("DIV") and div.get("TYPE", "").upper() == "SECTION":
            head = (div.findtext("HEAD") or "").strip()
            paras = [" ".join("".join(p.itertext()).split()) for p in div if p.tag != "HEAD"]
            body = "\n".join([head] + [p for p in paras if p])
            if "[Reserved]" in head and not any(paras):          # a placeholder, nothing to cite
                continue
            cite = f"§ {div.get('N')}" if div.get("N") else (head[:20] or "section")
            out.append((cite, body))
    return out


def _fetch(reg, getter, timeout):
    if reg["kind"] == "ecfr":
        titles = getter(f"{ECFR}/titles.json", headers=UA, timeout=timeout, what="ecfr.gov titles").json()
        t = next(x for x in titles.get("titles", []) if x.get("number") == reg["cfr_title"])
        day = t.get("up_to_date_as_of") or t.get("latest_issue_date")
        xml = getter(f"{ECFR}/full/{day}/title-{reg['cfr_title']}.xml", params={"part": reg["part"]}, headers=UA,
                     timeout=timeout, what=f"ecfr.gov title {reg['cfr_title']} part {reg['part']}").text
        secs = ecfr_sections(xml)
        if not secs:                                            # the XML changed: fall back to its text
            secs = split_sections(html_to_text(xml), reg["cite"])
        return secs, day
    page = getter(reg["url"], headers=UA, timeout=timeout, what=reg["title"]).text
    secs = split_sections(html_to_text(page), reg["cite"])
    pre = reg.get("cite_prefix", "")
    if pre:                                   # the WAC page heads its sections "296-127-010": cite them "WAC 296-127-010"
        secs = [(c if c.startswith(pre) else pre + c, b) for c, b in secs]
    return secs, dt.date.today().isoformat()


def regulation_chunks(cache_dir: Path, refresh_days=30, getter=None, timeout=60, regs=REGULATIONS):
    """(chunks, notes). Downloads a regulation when its saved copy is missing or older than refresh_days."""
    from . import net
    if getter is None:
        getter = lambda *a, **k: net.get(*a, waits=(3,), **k)    # one retry: the run never waits long for this
    cache = Path(cache_dir) / "regulations"
    cache.mkdir(parents=True, exist_ok=True)
    chunks, notes = [], []
    today = dt.date.today().isoformat()
    down = set()                                                # hosts that failed in this call
    for reg in regs:
        f = cache / f"{reg['id']}.json"
        tried = cache / f"{reg['id']}.last_attempt"
        try:
            saved = json.loads(f.read_text()) if f.is_file() else None
            dt.date.fromisoformat(saved["fetched"]) if saved else None
        except (ValueError, KeyError, TypeError, OSError):          # damaged copy: download it again
            saved = None
        fresh = saved and (dt.date.today() - dt.date.fromisoformat(saved["fetched"])).days < refresh_days
        host = re.sub(r"^https?://([^/]+).*$", r"\1", reg.get("url") or ECFR)
        if not fresh and (host in down or (tried.is_file() and tried.read_text().strip() == today)):
            notes.append(f"{reg['title']}: not downloaded (tried earlier today; tried again tomorrow)"
                         + ("; the copy saved on " + saved["fetched"] + " was used" if saved else ""))
            fresh = True                                        # at most one attempt a day
        if not fresh:
            tried.write_text(today)
            try:
                secs, as_of = _fetch(reg, getter, timeout)
                if len(secs) < 3:
                    raise ValueError(f"only {len(secs)} section(s) found; the page layout may have changed")
                saved = {"fetched": dt.date.today().isoformat(), "as_of": as_of, "sections": secs}
                tmp = f.with_suffix(".tmp")
                tmp.write_text(json.dumps(saved))
                tmp.replace(f)
                log.info("Regulation downloaded for the AI review",
                         extra={"details": {"regulation": reg["title"], "sections": len(secs), "as_of": as_of}})
            except Exception as e:                              # noqa: BLE001 - offline, blocked, layout changed
                down.add(host)
                notes.append(f"{reg['title']}: not downloaded ({err_text(e)[:160]})"
                             + ("; the copy saved on " + saved["fetched"] + " was used" if saved else ""))
        for cite, body in (saved or {}).get("sections", []):
            parts = _pieces(body)
            for n, part in enumerate(parts, 1):
                chunks.append(R.Chunk(chunk_id=f"{reg['id']}-{R._slug(cite)}" + (f"-p{n}" if len(parts) > 1 else ""),
                                      text=part, title=f"{reg['title']}, {cite}" + (f" (part {n} of {len(parts)})"
                                                                                    if len(parts) > 1 else ""),
                                      citation=cite, source_type="regulation", wage_source=reg["wage_source"]))
    return chunks, notes


def _pieces(body: str, size=4000) -> list:
    """A long section in parts of at most about `size` characters, cut between paragraphs (lines), and a paragraph
    longer than that at a sentence end. The first line (the heading) is repeated on each part."""
    if len(body) <= size:
        return [body]
    head, _, rest = body.partition("\n")
    head = head[:300]
    units = []
    for line in rest.split("\n"):
        while len(line) > size:
            cut = line.rfind(". ", 0, size)
            cut = cut + 1 if cut > 0 else size
            units.append(line[:cut])
            line = line[cut:].lstrip()
        if line:
            units.append(line)
    parts, cur, budget = [], [], size - len(head) - 1
    for u in units:
        if cur and sum(len(x) + 1 for x in cur) + len(u) > budget:
            parts.append(head + "\n" + "\n".join(cur))
            cur = []
        cur.append(u)
    if cur:
        parts.append(head + "\n" + "\n".join(cur))
    return parts or [body[:size]]


# ------------------------------------------------------------------------------------- 2. staff documents
DOC_EXT = {".pdf", ".txt", ".md"}
_HEADERS = re.compile(r"(?m)^\s*(§\s*[\d.]+[a-z]?(?:\([a-z0-9]+\))*|Section\s+[\w.]+|ARTICLE\s+[\w.]+|"
                      r"RCW\s+[\d.]+[a-z]?|WAC\s+[\d-]+[a-z]?)", re.I)


def _read_doc(f: Path) -> str:
    if f.suffix.lower() == ".pdf":
        from .pdftext import pdf_text
        return pdf_text(f)
    return f.read_text(errors="replace")


def _paragraph_chunks(text: str, size=1500) -> list:
    parts, cur = [], ""
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        if cur and len(cur) + len(para) > size:
            parts.append(cur)
            cur = ""
        cur = f"{cur}\n\n{para}" if cur else para
    if cur:
        parts.append(cur)
    return parts


def document_chunks(knowledge_dir: Path, project_names=()):
    """(chunks, notes) from reference/knowledge/. Folder names (any level, any case) say what a document is:
    'regulations', 'guidance' or 'contracts' (default: guidance); 'federal', 'state' or 'union' (default: all
    three); and a project's folder name makes a contract bind only that project."""
    d = Path(knowledge_dir)
    chunks, notes = [], []
    if not d.is_dir():
        return chunks, notes
    projects = {p.lower(): p for p in project_names}
    for f in sorted(d.rglob("*")):
        if not f.is_file() or f.name.startswith((".", "~$")):
            continue
        if f.suffix.lower() not in DOC_EXT:
            notes.append(f"{f.relative_to(d)}: not read (the AI review reads .pdf, .txt and .md)")
            continue
        parts = [p.lower() for p in f.relative_to(d).parts[:-1]]
        stype = next((t for key, t in (("regulation", "regulation"), ("guidance", "guidance"), ("contract", "contract"))
                      if any(key in p for p in parts)), "guidance")
        sources = [w for w in R.WAGE_SOURCES if any(w in p for p in parts)] or list(R.WAGE_SOURCES)
        project = next((projects[p] for p in parts if p in projects), None)
        try:
            text = _read_doc(f)
        except Exception as e:                                  # noqa: BLE001
            notes.append(f"a {f.suffix.lower() or 'no-extension'} document: could not be read ({err_text(e)[:120]})")  # no name: #56
            continue
        secs = [(h, b) for h, b in R._split_on_headers(text, _HEADERS) if len(b) > 40]
        covered = sum(len(b) for _, b in secs)
        # headers that cover little of the text (a stray "Section 9" in a memo) would drop the rest: use paragraphs
        pieces = secs if covered >= 0.7 * len(text.strip() or "x") else [("", p) for p in _paragraph_chunks(text)]
        rel = f.relative_to(d).as_posix()
        base = "DOC-" + hashlib.sha1(rel.encode()).hexdigest()[:8]
        for i, (head, body) in enumerate(pieces):
            for w in sources:
                chunks.append(R.Chunk(chunk_id=f"{base}-{i}-{w}", text=body[:4000],
                                      title=f"{f.stem}" + (f", {head}" if head else f" (part {i + 1})"),
                                      citation=head, source_type=stype, wage_source=w,
                                      project_id=project if stype == "contract" else None))
    return chunks, notes


# -------------------------------------------------------------------------------------- 3. staff decisions
# The sheet's name and headers: from the registry (cva_filekinds.payroll_export; the tokenizer keeps the columns
# read_decisions reads, DECISIONS_READ, readable in a payroll workbook: #53)
from .cva_filekinds.payroll_export import DECISION_HEADERS, DECISIONS_HEADER_ROWS, DECISIONS_SHEET  # noqa: E402,F401


def read_decisions(projects_dir: Path):
    """{pattern id: decision} from every reviewed payroll's "Review decisions" sheet (newest review wins).
    decision = {project, rule, kind, craft_class, contractor, decision: keep|dismiss, note, decided_on, example}."""
    import openpyxl
    out = {}
    files = []
    for proj in sorted(p for p in Path(projects_dir).iterdir() if p.is_dir() and not p.name.startswith((".", "_"))):
        pay = proj / "Payroll"
        if pay.is_dir():
            files += [(proj.name, f) for f in pay.rglob("*REVIEWED*.xlsx") if not f.name.startswith("~$")]
    def mtime(f):
        try:
            return f.stat().st_mtime
        except OSError:
            return 0
    for proj, f in sorted(files, key=lambda x: (mtime(x[1]), x[1].name)):     # oldest first: the newest wins
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        except Exception:                                       # noqa: BLE001 - open in Excel, damaged, a placeholder
            continue
        if DECISIONS_SHEET not in wb.sheetnames:
            continue
        rows = list(wb[DECISIONS_SHEET].iter_rows(values_only=True))
        hdr = next((i for i, r in enumerate(rows[:DECISIONS_HEADER_ROWS]) if r and "Pattern ID" in [str(x or "") for x in r]), None)
        if hdr is None:
            continue
        col = {str(h or ""): j for j, h in enumerate(rows[hdr])}
        m = re.match(r"^(\d{4}-\d{2}-\d{2})", f.name)
        day = dt.date.fromisoformat(m.group(1)) if m else dt.date.fromtimestamp(mtime(f) or 0)
        for r in rows[hdr + 1:]:
            get = lambda h: (r[col[h]] if h in col and col[h] < len(r) else None)
            pid, dec = get("Pattern ID"), str(get("Decision (Keep / Dismiss)") or "").strip().lower()
            if not pid or dec not in ("keep", "dismiss"):
                continue
            old = out.get(str(pid))
            if old and old["decision"] == dec:                  # carried over: keep the day it was first decided
                day_ = old["decided_on"]
            else:
                day_ = day
            out[str(pid)] = {"project": proj, "rule": str(get("Rule") or ""), "kind": str(get("Wage type") or ""),
                             "craft_class": str(get("Craft / classification") or ""),
                             "contractor": str(get("Contractor") or ""), "decision": dec,
                             "note": str(get("Note") or ""), "decided_on": day_, "example": str(get("Example") or "")}
    return out


def decision_chunks(decisions: dict, with_notes=True):
    """Past cases from staff decisions. Unlike the module's generic past cases, a staff decision counts for every
    week, earlier ones included: the same pattern comes back when a long payroll is reviewed again."""
    chunks = []
    for pid, d in decisions.items():
        kind = d["kind"] if d["kind"] in R.WAGE_SOURCES else "state"
        chunks.append(R.case_to_chunk({
            "case_id": pid, "project_id": d["project"], "wage_source": kind, "contractor": d["contractor"],
            "classification": d["craft_class"], "issue_type": d["rule"],
            "outcome": "corrected" if d["decision"] == "keep" else "dismissed", "decided_on": d["decided_on"],
            "summary": f"{d['rule']} on {d['craft_class']} ({d['contractor']}): {d['example']}",
            "note": d["note"] if with_notes else ""}))
        chunks[-1].valid_from = None
    return chunks
