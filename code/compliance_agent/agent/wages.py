"""Wage determination data: parse every supported format into one normalised table, merge new
files into the project's "in use" table (wages/<kind>/current/current_wage_rates.xlsx) and look up
the rate that applies to a payroll line."""
import csv
import datetime as dt
import hashlib
import re
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

# The wage table kinds come from the file-kind registry (code/filekinds, copied here as agent/cva_filekinds); the
# tokenizer imports the same definitions (pii_tokenizer/cva_filekinds), so the two can't drift apart (vNext PR 2)
from .cva_filekinds.merged_wage_table import MERGED_TABLE_HEADER
from .cva_filekinds.wage_table import (BASE_HDR, CLASS_HDR, CRAFT_HDR, LCP_HEADERS, TABLE_HEADERS,  # noqa: F401
                                       WA_COUNTIES, WAGE_HEADERS)
from .cva_filekinds.wage_table import HEADER_ROWS as LCP_HEADER_ROWS
from .cva_filekinds.wage_table import LNI_CSV_NEEDS
from .util import (date_from_filename, file_kinds, log, norm, presave, to_date, to_float,
                   visible_files)

# The merged table's columns (row 1 of Files/Wage rates/<Kind>.xlsx): the registry's merged wage table kind
FIELDS = list(MERGED_TABLE_HEADER)
# Rates from official determinations say their own wage type, so they need no Jurisdiction column: a DOL / SAM.gov
# federal determination, L&I's journey-level rates (CSV or PDF). Any other wage rates file (a spreadsheet: LCPtracker
# export, union table, or one read by the local model) must have a Jurisdiction column (David, Nov 7, 2026: strict).
SELF_TYPED = {"dol_wd": "federal", "lni_csv": "state", "lni_pdf": "state"}
OPEN_END = dt.date(2100, 1, 1)
SOURCE_RANK = {"lcp_export": 1, "union_table": 1, "dol_wd": 2, "lni_csv": 3, "llm": 9}
WAGE_EXTS = {".xlsx", ".xlsm", ".csv", ".txt", ".pdf"}


def row(**kw):
    r = {f: None for f in FIELDS}
    r.update(kw)
    return r


# ---------------------------------------------------------------- parsers
# Wage-table headers (David, Nov 8, 2026): "Trade" replaces "Craft" and "Base Rate" replaces "Basic Rate"; the old
# names are still read. CRAFT_HDR, CLASS_HDR, BASE_HDR, LCP_HEADERS (parse_lcp_export), TABLE_HEADERS
# (parse_union_table) and WAGE_HEADERS (every header either reader uses; the tokenizer keeps exactly these columns of
# a wage table readable) are the registry's (cva_filekinds.wage_table), imported above. The merged tables the agent
# writes itself keep the tokenizer's earlier handling, as the agent's own sheets do.


def _header_index(values, wanted):
    hdr = [re.sub(r"\s+", " ", str(v or "")).strip().lower() for v in values]
    return {k: next((i for i, h in enumerate(hdr) if re.fullmatch(p, h)), None)
            for k, p in wanted.items()}


def parse_lcp_export(path: Path, kind):
    """LCPtracker 'Project Wage Data Export' / 'Master Wage Data' report (.xlsx)."""
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    county = None
    hi = None
    for i, r in enumerate(rows[:LCP_HEADER_ROWS]):
        first = str(r[0] or "") if r else ""
        if first.startswith("Location:"):
            county = first.split(":", 1)[1].strip()
        if r and re.fullmatch(CRAFT_HDR, str(r[0] or "").strip().lower()) and any(
                "classification" in str(c or "").lower() for c in r):
            hi = i
            break
    if hi is None:
        return []
    ix = _header_index(rows[hi], LCP_HEADERS)
    out = []
    for r in rows[hi + 1:]:
        if not r or r[ix["craft"]] in (None, "") or re.fullmatch(CRAFT_HDR, str(r[ix["craft"]]).strip().lower()):
            continue
        if str(r[0]).startswith(("Construction Type", "Location", "Jurisdiction")):
            continue
        if (ix["total"] is None or to_float(r[ix["total"]]) is None) and \
                (ix["basic"] is None or to_float(r[ix["basic"]]) is None):
            continue  # footers / disclaimers
        dates = str(r[ix["dates"]] or "").split("\n") if ix["dates"] is not None else []
        issue = to_date(dates[0]) if len(dates) > 0 else None
        expire = to_date(dates[1]) if len(dates) > 1 else None
        g = lambda k: r[ix[k]] if ix[k] is not None and ix[k] < len(r) else None
        out.append(row(
            kind=kind, craft=str(g("craft")).strip(), classification=str(g("cls") or "").strip(),
            wage_decision=str(g("wd") or "").strip(), mod=to_float(g("mod")),
            issue_date=issue, expire_date=expire or OPEN_END,
            basic=to_float(g("basic")) or None, fringe=to_float(g("fringe")),
            total=to_float(g("total")), ot_total=to_float(g("ot")) or None,
            dt_total=to_float(g("dt")) or None, std_hours=to_float(g("std")) or None,
            county=county, source_type="lcp_export", source_file=path.name, _jur=g("jur")))
    return out


def parse_lni_csv(path: Path, kind):
    """WA L&I 'Journey level wages' CSV (county, trade, classification, total wage, codes)."""
    issue = date_from_filename(path.name)
    out = []
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        rd = csv.reader(f)
        hdr = [h.strip().lower() for h in next(rd)]
        idx = {h: i for i, h in enumerate(hdr)}
        need = list(LNI_CSV_NEEDS)          # the registry's: the tokenizer lets such a CSV through (#111)
        if not all(n in idx for n in need):
            return []
        for r in rd:
            if not r or len(r) < len(need):
                continue
            g = lambda k: r[idx[k]].strip() if k in idx and idx[k] < len(r) else None
            out.append(row(
                kind=kind, craft=g("trade"), classification=g("job classification"),
                wage_decision=f"L&I Journey Level {issue or ''}".strip(), issue_date=issue,
                expire_date=OPEN_END, total=to_float(g("wage")), county=g("county"),
                holiday_code=g("holiday"), ot_code=g("overtime"), notes=g("notes"),
                source_type="lni_csv", source_file=path.name))
    return out


LNI_PDF_MARK = re.compile(r"Prevailing\s*Wage.*?Journey\s*Level|Journey\s*Level\s*Prevailing\s*Wage", re.I | re.S)
_MONEY = re.compile(r"^\$\s?\d[\d,]*\.\d{2}$")


def _peaks(xs, min_sep=25, min_count=1, k=None):
    """Left edges of columns: the x positions where most words start (counted within 2 points), at least
    ``min_sep`` apart, busiest first. Words in the middle of a cell start anywhere, so they never form a peak."""
    from collections import Counter
    counts = Counter(round(x) for x in xs)
    score = {x: sum(counts.get(x + d, 0) for d in range(-2, 3)) for x in counts}
    out = []
    for x in sorted(score, key=lambda x: (-score[x], x)):
        if score[x] >= min_count and all(abs(x - y) >= min_sep for y in out):
            out.append(x)
            if k and len(out) == k:
                break
    return out


def _pdf_lines(words, tol=3.0):
    """Words grouped into printed lines: words of one row can sit a couple of points apart vertically."""
    lines = []
    for w in sorted(words, key=lambda w: w["top"]):
        if lines and w["top"] - lines[-1][0] <= tol:
            lines[-1][1].append(w)
        else:
            lines.append([w["top"], [w]])
    return [sorted(ws, key=lambda w: w["x0"]) for _, ws in lines]


def _lni_code(words):
    """A holiday / overtime code as printed ('1N', '11P'); the print sometimes shows the digit 1 as 'I'."""
    t = " ".join(words)
    m = re.fullmatch(r"([0-9I]+)([A-Z]?)", t)
    return (m.group(1).replace("I", "1") + m.group(2)) if m and any(c.isdigit() or c == "I" for c in m.group(1)) \
        else (t or None)


def _join_words(words):
    """Words of a cell, joined; a word broken at a hyphen where the cell wraps is put back together:
    'Window Cleaner (Non-' / 'Scaffold)' -> 'Window Cleaner (Non-Scaffold)'."""
    return re.sub(r"(?<=[A-Za-z])- (?=[A-Za-z])", "-", " ".join(words))


def parse_lni_pdf(path: Path, kind):
    """WA L&I 'Journey Level Prevailing Wage Rates' page printed to PDF (the same table as the L&I CSV export):
    County, Trade, Job Classification, Wage, Holiday, Overtime, Note. Columns are found from where words start on
    the page (the table has no lines); a trade or classification that wraps onto a second line is joined to the
    row above. Gives the same rows as parse_lni_csv."""
    first = pdf_first_page_text(path)
    if not LNI_PDF_MARK.search(first):
        return []
    pages = [[{k: w[k] for k in ("text", "x0", "top", "bottom")} for w in ws] for ws in pdf_pages(path, "words")]
    m = re.search(r"Effective\s*Date:?\s*(\d{1,2}/\d{1,2}/\d{4})", first)
    issue = to_date(m.group(1)) if m else date_from_filename(path.name)
    # Columns are found from the rate rows only (lines with a $ amount), never from headings or page furniture.
    page_lines = [_pdf_lines(ws) for ws in pages]
    data_lines = [ln for lines in page_lines for ln in lines if any(_MONEY.match(w["text"]) for w in ln)]
    words = [w for ln in data_lines for w in ln]
    money = [w for w in words if _MONEY.match(w["text"])]
    if not money:
        return []
    wage_x = sorted(round(w["x0"]) for w in money)[len(money) // 2]
    view_x = min((round(w["x0"]) for w in words if w["text"] == "View" and w["x0"] > wage_x), default=None)
    # Two layouts: the L&I print has a County column (County, Trade, Job Classification, ...); the exhibit attached
    # to a contract (Brennick's '900001 - QCV Wages') has none, with the county as a heading ("King County") above the
    # table. Which one: the table heading line, the one with "Trade" and "Classification".
    heads = [" ".join(w["text"] for w in ln) for lines in page_lines for ln in lines]
    head = next((h for h in heads if re.search(r"\bTrade\b", h) and "Classification" in h), "")
    has_county_col = not head or bool(re.search(r"\bCounty\b", head))
    # Columns: the busiest left edges of the rate rows before the wage (county, trade, classification; or trade,
    # classification); codes: the left edges between the wage and the "View" link.
    edges = _peaks([w["x0"] for w in words if w["x0"] < wage_x - 8], k=3 if has_county_col else 2)
    if len(edges) < (3 if has_county_col else 2):
        return []
    edges = sorted(edges)
    county_x = edges[0] if has_county_col else None
    trade_x, class_x = edges[-2:]
    right = view_x or 10 ** 6
    codes = sorted(_peaks([w["x0"] for w in words if wage_x + 20 < w["x0"] < right - 4],
                          min_count=max(1, len(money) // 20), k=3))
    names = ["holiday", "overtime", "notes"][:len(codes)]

    def code_col(x):
        return min(zip(codes, names), key=lambda cn: abs(cn[0] - x))[1] if codes else None

    def is_rate(ln):
        return any(_MONEY.match(w["text"]) and abs(w["x0"] - wage_x) <= 12 for w in ln)
    # A cell that wraps: in the L&I print every cell starts on its rate line and wraps downwards; in the contract
    # exhibit (no County column) each cell is centred on its rate line, so a long trade or classification starts
    # above it, and the rate line of a two-line trade has no trade text at all. Either sign means centred cells.
    centred = not has_county_col or any(not any(trade_x - 6 <= w["x0"] < class_x - 6 for w in ln)
                                        for ln in data_lines if is_rate(ln))
    county_head = re.compile(r"^([A-Z][A-Za-z.' -]*?)\s+County$")
    # The lines in reading order, as one stream of (kind, page, top, line): a county heading, the table heading, a
    # rate line (a row) or a wrapped line. Anything above the table heading (title, notes) and page furniture left of
    # the trade column are dropped.
    stream, in_table = [], not head
    for pi, lines in enumerate(page_lines):
        for ln in lines:
            text, top = " ".join(w["text"] for w in ln), min(w["top"] for w in ln)
            rate = is_rate(ln)
            m = county_head.match(text)
            if m and not has_county_col and not rate:
                stream.append(("county", pi, top, m.group(1)))
            elif re.search(r"\bTrade\b", text) and "Classification" in text:
                in_table = True
                stream.append(("heading", pi, top, None))
            elif in_table and (rate or ln[0]["x0"] >= trade_x - 10):
                stream.append(("rate" if rate else "wrap", pi, top, ln))
    parts = {i: [(i, ln)] for i, (k, _, _, ln) in enumerate(stream) if k == "rate"}
    county_of, county = {}, None
    for i, (k, _, _, v) in enumerate(stream):
        if k == "county":
            county = v
        elif k == "rate":
            county_of[i] = county
    if not centred:
        # The L&I print: a wrapped line joins the row above while each line follows within 22 points.
        last, last_page, last_top = None, None, None
        for i, (k, pi, top, v) in enumerate(stream):
            if k == "rate":
                last, last_page, last_top = i, pi, top
            elif k == "wrap" and last is not None and pi == last_page and 0 < top - last_top <= 22:
                parts[last].append((i, v))
                last_top = top
            elif k != "wrap":
                last = None
    else:
        # The contract exhibit: lines of one row are at most 2.2 text heights apart (a cell is centred on its rate
        # line, so a long trade or classification starts above it); rows are 3.4 text heights or more apart (23 and
        # 34 points in Brennick's exhibit, text 10 points high). So the lines fall into blocks, one rate line each. It is a web page printed to PDF: a row cut by a page break carries
        # on at the top of the next page ("Power Equipment Operators- Underground" / "Sewer & Water"), so a block
        # with no rate line at the top of a page belongs to the last row on the page before (and one at the bottom
        # of a page, to the first row on the next).
        import statistics
        row_gap = 2.8 * statistics.median(w["bottom"] - w["top"] for w in words)
        blocks, prev = [], None
        for i, (k, pi, top, _) in enumerate(stream):
            if k not in ("rate", "wrap"):
                prev = None
                continue
            if prev is None or pi != prev[0] or top - prev[1] > row_gap:
                blocks.append([])
            blocks[-1].append(i)
            prev = (pi, top)
        pages_of = [stream[b[0]][1] for b in blocks]
        has_rate = [any(stream[i][0] == "rate" for i in b) for b in blocks]
        for n, b in enumerate(list(blocks)):
            if has_rate[n]:
                continue
            if n and pages_of[n - 1] != pages_of[n]:                       # top of a page: the row before
                blocks[n - 1].extend(b)
                blocks[n] = []
            elif n + 1 < len(blocks) and pages_of[n + 1] != pages_of[n]:   # bottom of a page: the row after
                blocks[n + 1][:0] = b
                blocks[n] = []
        rates = lambda b: [i for i in b if stream[i][0] == "rate"]        # noqa: E731
        dropped = [" ".join(w["text"] for w in stream[i][3]) for b in blocks if not rates(b) for i in b]
        if dropped:                         # text between rows that belongs to none: say so, never drop it quietly
            # no file name and none of the left-out text: a name can hold a worker's name (#56)
            log.warning("L&I rates PDF: %d line(s) belong to no rate row and were left out", len(dropped))
        for b in blocks:
            rs = rates(b)
            for i in b:
                if stream[i][0] == "wrap" and rs:   # one rate line per block; if ever two, the nearer one
                    j = min(rs, key=lambda j: (abs(stream[i][2] - stream[j][2]) + 1000 * (stream[i][1] != stream[j][1]),
                                               j > i))
                    parts[j].append((i, stream[i][3]))
    out = []
    for i in sorted(parts):
        cur = {"county": [] if has_county_col else [county_of[i] or ""], "trade": [], "classification": [],
               "wage": None, "holiday": [], "overtime": [], "notes": []}
        out.append(cur)
        for w in (w for _, ln in sorted(parts[i], key=lambda p: p[0]) for w in ln):   # in reading order
            x, t = w["x0"], w["text"]
            if view_x and x >= view_x - 4:
                continue
            if _MONEY.match(t) and abs(x - wage_x) <= 12:
                cur["wage"] = t
            elif has_county_col and x < trade_x - 6:
                cur["county"].append(t)
            elif x < class_x - 6:
                cur["trade"].append(t)
            elif x < wage_x - 4:
                cur["classification"].append(t)     # "Mechanics: All (Leadmen - $0.50 per hour over mechanic)"
            elif (col := code_col(x)):
                cur[col].append(t)
    rows = []
    for r in out:
        code = lambda k: _lni_code(r[k])
        rows.append(row(kind=kind, craft=_join_words(r["trade"]), classification=_join_words(r["classification"]),
                        wage_decision=f"L&I Journey Level {issue or ''}".strip(), issue_date=issue,
                        expire_date=OPEN_END, total=to_float(r["wage"]), county=" ".join(r["county"]) or None,
                        holiday_code=code("holiday"), ot_code=code("overtime"), notes=" ".join(r["notes"]) or None,
                        source_type="lni_pdf", source_file=path.name))
    return rows


_NUM = r"\d+(?:\.\d+)?"
# fringe: "19.01", or with a percentage of the base rate: "3%+27.21", "27.21+3%", "3%", "3% + 27.21"
_RATE_LINE = re.compile(r"^(?P<name>.*?)\.{2,}\s*\$\s*(?P<rate>" + _NUM + r")\s*(?:\*\*)?\s*"
                        r"(?P<fringe>(?:" + _NUM + r"%\s*\+\s*)?" + _NUM + r"(?:\s*\+\s*" + _NUM + r"%|%)?)?(?P<rest>.*)$")


def fringe_amount(text, basic):
    """'3%+27.21' with basic 60.00 -> 29.01; '19.01' -> 19.01."""
    if not text:
        return 0.0
    total = 0.0
    for part in re.split(r"\s*\+\s*", text.strip()):
        total += basic * float(part[:-1]) / 100 if part.endswith("%") else float(part)
    return round(total, 4)


def parse_dol_text(text: str, kind, source_name):
    """US DOL / SAM.gov Davis-Bacon general wage determination text."""
    m = re.search(r"General Decision Number:\s*([A-Z]{2}\d{8})\s+(\d{2}/\d{2}/\d{4})", text)
    if not m:
        return []
    wd = m.group(1)
    mods = re.findall(r"^\s*(\d+)\s+(\d{2}/\d{2}/\d{4})\s*$", text, re.M)
    mod, pub = (int(mods[-1][0]), to_date(mods[-1][1])) if mods else (0, to_date(m.group(2)))
    county = (re.search(r"County:\s*(.+)", text) or [None, None])[1]
    out, pending, block = [], [], None
    indents, emitted, skip_to = [], False, -1
    lines = text.splitlines()
    for i, raw in enumerate(lines):
        if i <= skip_to:
            continue
        line = raw.rstrip()
        s = line.strip()
        if not s:
            pending, indents = [], []
            continue
        if s.startswith("---") or s.startswith("FOOTNOTE") or s.startswith("WELDERS"):
            pending = []
            if s.startswith("---"):
                block = None
            continue
        bm = re.match(r"^\*?\s*([A-Z]{4}\d{4}-\d{3}|SU[A-Z]{2}\d{4}-\d{3}|UAVG-\S+)\s+(\d{2}/\d{2}/\d{4})", s)
        if bm:
            block, pending, indents = bm.group(1), [], []
            continue
        if block is None or "Rates" in s and "Fringes" in s:
            continue
        rm = _RATE_LINE.match(line)
        if rm:
            text_part = rm.group("name").strip()
            if line[:1].isspace() and pending:
                name = " ".join(pending) + " - " + text_part
                emitted = True
                # the start of a wrapped name ("CEMENT MASONS ON / SUSPENDED ... / HANGING SCAFFOLD....$") belongs to
                # this rate line only; sub-headings ("AREA 1:") and the craft heading stay for the next lines
                ind = len(line) - len(line.lstrip())
                while len(pending) > 1 and indents and indents[-1] >= ind and not pending[-1].endswith(":"):
                    indents.pop()
                    pending.pop()
            else:
                name = " ".join(pending + [text_part])
                pending, indents = [], []
            name = re.sub(r"\s+", " ", name).strip(" :")
            craft, cls = (name.split(":", 1) + [""])[:2] if ":" in name else (name, "")
            basic = float(rm.group("rate"))
            fringe = fringe_amount(rm.group("fringe"), basic)
            rest = (rm.group("rest") or "").strip()
            out.append(row(
                kind=kind, craft=craft.strip(), classification=cls.strip() or craft.strip(),
                wage_decision=wd, mod=mod, issue_date=pub, expire_date=OPEN_END, basic=basic,
                fringe=fringe, total=round(basic + fringe, 4),
                ot_total=round(basic * 1.5 + fringe, 4), dt_total=round(basic * 2 + fringe, 4),
                county=county, notes=(f"rate block {block}" + (f"; fringe footnote {rest}" if rest else "")),
                source_type="dol_wd", source_file=source_name))
        elif not s.startswith(("*", "Note", "Rates")) and len(s) < 80:
            depth = s.count("(") - s.count(")")
            if depth > 0:
                # a description in brackets running over several lines and closed before any rate line, e.g.
                # "Truck drivers: (AREA 1: SPOKANE ZONE CENTER: Adams, ... Counties)": dropped from the name
                d, j = depth, i
                while d > 0 and j + 1 < len(lines) and j - i < 40:
                    j += 1
                    if _RATE_LINE.match(lines[j].rstrip()):
                        break
                    d += lines[j].count("(") - lines[j].count(")")
                if d <= 0 and j > i:
                    s, skip_to = s[:s.index("(")].strip(), j
            if not s:
                continue
            ind = len(line) - len(line.lstrip())
            if emitted:                         # a new sub-heading after rates: it replaces the sub-heading at its level
                while indents and indents[-1] >= ind and len(pending) > 1:
                    indents.pop()
                    pending.pop()
                emitted = False
            pending.append(s)
            indents.append(ind)
    return out


DOL_MARK = re.compile(r"General\s*Decision\s*Number", re.I)


from .pdftext import pdf_pages                      # noqa: E402


def pdf_first_page_text(path: Path) -> str:
    """Text of a PDF's first page only: enough to tell what the document is, in well under a second, where
    reading a 40-page determination takes minutes."""
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        if not pdf.pages:
            return ""
        page = pdf.pages[0]
        try:
            return page.extract_text() or ""
        finally:
            page.close()


def parse_dol_file(path: Path, kind):
    if path.suffix.lower() == ".pdf":
        try:
            if not DOL_MARK.search(pdf_first_page_text(path)):
                log.debug("PDF is not a DOL wage determination (first page)")
                return []
            text = "\n".join(pdf_pages(path, "text"))
        except Exception as e:  # pragma: no cover
            log.warning("Could not read a wage PDF: %s", type(e).__name__)    # no file name: #56
            return []
    else:
        text = path.read_text(errors="replace")
    return parse_dol_text(text, kind, path.name)


def parse_union_table(ws, kind="union", source_name="reference rules workbook"):
    """A wage table with its header in row 1: the 'WSCCA Wages' union table (Craft, Classification, Basic Rate,
    Fringes, Total Rate, OT, DT, Std hours, Issue, Expire, Vac/PTO in gross wage) or the firm's wage table template
    (County, Trade, Classification, Total Wage, Base Rate, Fringe Rate, Jurisdiction, Decision, Effective Date,
    Issue Date, Expiration Date). A Decision column gives each rate its determination (else "Union CBA"); the
    Effective Date, when there is one, is when the rate starts (else the Issue Date)."""
    rows = ws.iter_rows(values_only=True)
    hdr = next(rows)
    ix = _header_index(hdr, TABLE_HEADERS)
    out = []
    for r in rows:
        if not r or ix["cls"] is None or r[ix["cls"]] in (None, ""):
            continue
        g = lambda k: r[ix[k]] if ix[k] is not None and ix[k] < len(r) else None
        start = to_date(g("effective")) or to_date(g("issue"))
        county = str(g("county") or "").strip() or None
        out.append(row(kind=kind, craft=str(g("craft") or "").strip(),
                       classification=str(g("cls")).strip(), wage_decision=str(g("wd") or "").strip() or "Union CBA",
                       issue_date=start, expire_date=to_date(g("expire")) or OPEN_END,
                       basic=to_float(g("basic")), fringe=to_float(g("fringe")),
                       total=to_float(g("total")), ot_total=to_float(g("ot")),
                       dt_total=to_float(g("dt")), std_hours=to_float(g("std")),
                       vac_pto=to_float(g("vac")) or 0.0, county=county, source_type="union_table",
                       source_file=source_name, _jur=g("jur")))
    return out


def sniff_and_parse(path: Path, kind, llm=None):
    ext = path.suffix.lower()
    rows = []
    parser = None
    try:
        if ext in (".xlsx", ".xlsm"):
            parser = "LCPtracker wage export"
            rows = parse_lcp_export(path, kind)
            if not rows:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                ix = _header_index(next(ws.iter_rows(max_row=1, values_only=True)),
                                   {"cls": CLASS_HDR, "base": BASE_HDR})
                if ix["cls"] is not None and ix["base"] is not None:
                    parser = "union wage table"
                    rows = parse_union_table(ws, kind, path.name)
        elif ext == ".csv":
            parser = "L&I journey-level CSV"
            rows = parse_lni_csv(path, kind)
        elif ext in (".txt", ".pdf"):
            parser = "DOL wage determination"
            rows = parse_dol_file(path, kind)
            if not rows and ext == ".pdf":
                parser = "L&I journey-level rates (PDF)"
                rows = parse_lni_pdf(path, kind)
    except Exception as e:
        log.warning("Parser error on a %s wage file: %s", path.suffix.lower(), type(e).__name__, exc_info=True,
                    extra={"details": {"parser": parser}})
    log.debug("Parsed wage file", extra={"details": {"type": path.suffix.lower(), "parser": parser or "none matched",
                                                       "rows": len(rows)}})
    if not rows and llm is not None and llm.available():
        log.info("No built-in parser understood a %s wage file; asking the local model", path.suffix.lower())
        rows = [row(kind=kind, source_type="llm", source_file=path.name, **r)
                for r in llm.extract_wage_rows(path)]
    return rows


# ------------------------------------------------------------- the table
def match_key(r):
    return norm(f"{r['craft']} {r['classification']}")


def alias_key(craft, classification):
    """Craft + classification without the craft repeated: 'TRUCK DRIVER - GROUP 3' / 'TRUCK DRIVER - GROUP 3' (a wage
    determination) and 'TRUCK DRIVER' / 'Group 3' (a payroll) both give 'truck driver group 3'."""
    c, k = norm(craft), norm(classification)
    if not c:
        return k
    if not k:
        return c
    return k if k.startswith(c) else f"{c} {k}"


def county_key(r):
    c = norm(r.get("county"))
    return c.split()[0] if c else ""


# WA_COUNTIES: the registry's (cva_filekinds.wage_table), imported above
DECISION_RE = re.compile(r"\b([A-Z]{2}\d{8})\b")


def jurisdiction_target(text, mapping=None):
    """What a payroll line's Jurisdiction text says about the wage determination to use:
    {"kind": None|federal|state|union, "wage_decision": None|"WA20250091", "county": None|"king"}.
    An entry in `mapping` (exact Jurisdiction text, case-insensitive) wins over what is read from the text."""
    t = str(text or "")
    out = {"kind": None, "wage_decision": None, "county": None}
    m = DECISION_RE.search(t.upper())
    if m:
        out["wage_decision"] = m.group(1)
    low = f" {t.lower()} "
    for c in sorted(WA_COUNTIES, key=len, reverse=True):
        if re.search(rf"\b{c.lower()}\b", low):
            out["county"] = norm(c).split()[0]
            break
    for k, v in (mapping or {}).items():
        if norm(k) == norm(t):
            v = v or {}
            out.update({x: v[x] for x in ("kind", "wage_decision") if v.get(x)})
            if v.get("county"):
                out["county"] = norm(v["county"]).split()[0]
            break
    return out


def _route(rows, kind, cfg):
    """Keep only rows that belong in this kind's folder (combined State + Federal exports)."""
    if kind == "union":
        return rows
    fed = re.compile(cfg["wages"]["federal_decision_regex"])
    keep = []
    for r in rows:
        is_fed = bool(fed.match(str(r.get("wage_decision") or "")))
        if (kind == "federal") == is_fed:
            keep.append(r)
    return keep


def read_table(path: Path):
    if not path.exists():
        return []
    ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    out = []
    for r in it:
        d = dict(zip(hdr, r))
        for k in ("issue_date", "expire_date"):
            d[k] = to_date(d.get(k))
        # a table merged before the Jurisdiction column existed: official determinations still say their own type
        d["jurisdiction"] = d.get("jurisdiction") or SELF_TYPED.get(d.get("source_type"))
        out.append(row(**{k: d.get(k) for k in FIELDS}))
    return out


# One line per wage file merged into the table. file_sha256 and rates_fingerprint identify a file already processed
# for this project, so an operator re-uploading it does not merge it again (see processed_wage_files).
SOURCE_COLUMNS = ["merged_at_utc", "file", "rows", "file_sha256", "rates_fingerprint"]


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def rates_fingerprint(rows) -> str:
    """Identifies the rates a file holds, whatever its bytes: the same table re-saved, a copy that went through
    the tokenizer, or the same L&I table as CSV or as PDF give the same fingerprint."""
    key = sorted((r["kind"] or "", str(r.get("jurisdiction") or ""), norm(r["craft"]), norm(r["classification"]),
                  str(r["wage_decision"] or ""),
                  str(r["issue_date"] or ""), str(r.get("expire_date") or ""), norm(r["county"]),
                  *(round(float(r.get(k)), 2) if r.get(k) not in (None, "") else None
                    for k in ("total", "basic", "fringe", "ot_total", "dt_total", "vac_pto", "std_hours")))
                 for r in rows)
    return hashlib.sha256(repr(key).encode()).hexdigest()


def processed_wage_files(paths) -> dict:
    """{"sha": {checksum: record}, "fp": {fingerprint: record}} from the sources sheets of the project's rate
    tables: every wage file already merged for this project."""
    out = {"sha": {}, "fp": {}}
    for kind in ("federal", "state", "union"):
        p = paths.rate_file(kind)
        if not p.exists():
            continue
        try:
            ws = openpyxl.load_workbook(p, read_only=True)["sources"]
        except Exception:
            continue
        for r in ws.iter_rows(min_row=2, values_only=True):
            rec = dict(zip(SOURCE_COLUMNS, list(r) + [None] * len(SOURCE_COLUMNS)))
            if rec.get("file_sha256"):
                out["sha"].setdefault(rec["file_sha256"], rec)
            if rec.get("rates_fingerprint"):
                out["fp"].setdefault(rec["rates_fingerprint"], rec)
    return out


def write_table(path: Path, rows, sources):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "rates"
    ws.append(FIELDS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFD9E1F2")
    for r in sorted(rows, key=lambda r: (match_key(r), r["issue_date"] or dt.date.min)):
        ws.append([r[f] for f in FIELDS])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    s = wb.create_sheet("sources")
    s.append(SOURCE_COLUMNS)
    for x in sources:
        s.append(x)
    from .safesave import save_workbook
    save_workbook(wb, path)          # never half-written, even if the run is killed


def merge_rows(existing, new):
    """Union of rows, de-duplicated; newer issue/mod supersedes older for the same
    classification by closing the older row's open-ended expire date."""
    best = {}
    for r in existing + new:
        k = (r["kind"], match_key(r), r["issue_date"], county_key(r), r["basic"], r["total"], _decision_key(r))
        cur = best.get(k)
        rank = lambda x: (0 if x.get("jurisdiction") else 1, SOURCE_RANK.get(x["source_type"], 5), -(x["mod"] or 0))
        cand_rank = rank(r)              # a rate that says its wage type first (an untyped one is never used)
        # on a tie the row read later wins: a corrected file (same rates, new overtime or PTO) replaces the old row
        if cur is None or cand_rank <= rank(cur):
            best[k] = r
    rows = list(best.values())
    groups = {}
    for r in rows:        # typed and untyped rates apart: an unused (untyped) rate never cuts off a used one
        groups.setdefault((r["kind"], match_key(r), county_key(r), bool(r.get("jurisdiction")), _decision_key(r)),
                          []).append(r)
    for g in groups.values():
        g.sort(key=lambda r: (r["issue_date"] or dt.date.min, r["mod"] or 0))
        for older, newer in zip(g, g[1:]):
            if newer["issue_date"] and (older["expire_date"] or OPEN_END) >= OPEN_END \
                    and (older["issue_date"] or dt.date.min) < newer["issue_date"]:
                older["expire_date"] = newer["issue_date"] - dt.timedelta(days=1)
    return rows


def _decision_key(r) -> str:
    """Federal determinations (WA20250091, WA20250093, ...) are different determinations in force side by side: one
    never cuts off or replaces another's rate (only its own later modifications do). State and union rates have no
    such number: their newer issue replaces the older, as before."""
    if r.get("kind") != "federal":
        return ""
    m = DECISION_RE.search(str(r.get("wage_decision") or "").upper())
    return m.group(1) if m else ""


def jurisdiction_kind(text):
    """A wage file's Jurisdiction cell: "federal", "state" or "union" (the word, any case, alone or in a phrase such
    as "State PW"), else None."""
    words = set(re.findall(r"[a-z]+", str(text or "").lower())) & {"federal", "state", "union"}
    return words.pop() if len(words) == 1 else None


def _kind_of_rows(rows, path: Path, cfg):
    """Split parsed rows by wage type. A row from a wage file with a Jurisdiction column takes its kind from it
    (David, Nov 7, 2026: staff's wage rates file says state, federal or union for each rate). Otherwise: a union rate
    table is union; a row whose wage decision looks like a federal decision number (WA20250091) is federal, the rest
    state. Combined State + Federal exports split cleanly."""
    fed = re.compile(cfg["wages"]["federal_decision_regex"])
    out = {}
    for r in rows:
        jk = jurisdiction_kind(r.pop("_jur", None))
        r["jurisdiction"] = jk or SELF_TYPED.get(r.get("source_type"))   # None: no Jurisdiction column (not used)
        if jk:
            k = jk
            if k != "union" and r.get("wage_decision") == "Union CBA":
                r["wage_decision"] = ""             # a union-style table's row that says it is state or federal
        elif r.get("source_type") == "union_table" or (r.get("source_type") == "llm" and re.search(r"union|cba", path.name, re.I)):
            k = "union"
        elif fed.match(str(r.get("wage_decision") or "")):
            k = "federal"
        else:
            k = "state"
        r["kind"] = k
        out.setdefault(k, []).append(r)
    return out


def update_wage_data(project_dir: Path, cfg, state, llm=None, tz="UTC"):
    """Merge the wage determinations waiting in Uploads/ into a new 'Files/Wage rates/<Federal|State|Union>
    <Pacific time> PST|PDT.xlsx' per kind (layout.rate_file_name). The previous rate tables are moved to
    Files/Archived/<ts>/ as '<name> (previous).xlsx' and the uploaded files to Files/Archived/<ts>/. A file nothing
    could be read from stays in Uploads/. Returns a summary list."""
    from .layout import FILES, FILES_DIR, RATES_DIR, Paths, archive_files, rate_file_name, rate_files
    from .uploads import uploads_of
    paths = Paths(project_dir)
    from .samgov import fetch_for_project
    fetched = fetch_for_project(project_dir, cfg)      # federal determinations listed for the project, from SAM.gov
    files = uploads_of(project_dir, "wages")
    if not files:
        log.debug("No wage files in Uploads")
        return fetched
    log.info("Wage files in Uploads: %d", len(files), extra={"details": {"files": file_kinds(files)}})
    # A PDF with a same-decision .txt twin is the same document: read the text version only.
    txt_ids = {re.search(r"[A-Z]{2}\d{8}", f.name).group(0) for f in files
               if f.suffix.lower() == ".txt" and re.search(r"[A-Z]{2}\d{8}", f.name)}
    by_kind, used, unreadable, already, no_jur, partly = {}, [], [], [], [], []
    done = processed_wage_files(paths)                 # wage files already merged for this project
    info = {}                                          # file -> (checksum, fingerprint)
    for f in files:
        wid = re.search(r"[A-Z]{2}\d{8}", f.name)
        if f.suffix.lower() == ".pdf" and wid and wid.group(0) in txt_ids:
            used.append((f, {}, "skipped: same decision as .txt"))
            log.debug("Skipped PDF: a .txt of the same wage decision was uploaded")
            continue
        sha = file_sha256(f)
        if sha in done["sha"]:                         # the very same file: not even read again
            already.append((f, done["sha"][sha]))
            continue
        split = _kind_of_rows(sniff_and_parse(f, "state", llm), f, cfg)
        split = {k: v for k, v in split.items() if k in cfg["wages"]["kinds"]}
        blank = sum(1 for v in split.values() for r in v if not r.get("jurisdiction"))
        if split and blank:
            # some rates with a blank Jurisdiction: those rows are dropped, and the run says how many
            if blank < sum(len(v) for v in split.values()):
                partly.append((f, blank))
                split = {k: [r for r in v if r.get("jurisdiction")] for k, v in split.items()}
                split = {k: v for k, v in split.items() if v}
        if split and not any(r.get("jurisdiction") for v in split.values() for r in v):
            # a wage rates file without a Jurisdiction column: not merged (its rates couldn't be used), it stays in
            # Uploads and both emails say why (David, Nov 7, 2026)
            no_jur.append(f)
            log.warning("A wage rates file has no Jurisdiction column; it stays in Uploads",
                        extra={"details": {"type": f.suffix.lower()}})
            continue
        fp = rates_fingerprint([r for v in split.values() for r in v]) if split else None
        if fp and fp in done["fp"]:                    # different bytes, same rates (re-saved, or via AWS)
            already.append((f, done["fp"][fp]))
            continue
        if split:
            rec = {"merged_at_utc": "this run", "file": f.name}
            done["sha"][sha] = done["fp"][fp] = rec    # the same file twice in one upload: merged once
            info[f] = (sha, fp)
        if not split:
            unreadable.append(f)
            log.warning("Could not read wage file (no rows found); it stays in Uploads", extra={"details": {
                "type": f.suffix.lower(), "hint": "check it is a supported export; enable the local model for other layouts"}})
            continue
        for k, v in split.items():
            by_kind.setdefault(k, []).extend(v)
        used.append((f, {k: len(v) for k, v in split.items()}, "parsed"))
    summary = []
    from .util import run_timestamp
    for kind, parsed in sorted(by_kind.items()):
        # The table in use is the newest of this kind (layout.rate_files); the merged one gets a new name, for this
        # run in Pacific time (layout.rate_file_name, David, Nov 8, 2026), and every older table of the kind is
        # moved to Files/Archived/<ts>/ as '<name> (previous).xlsx' once the new one is safely written.
        olds = rate_files(paths.rates, kind)
        cur_path = olds[-1] if olds else None
        new_path = paths.rates / rate_file_name(kind, run_timestamp(tz))
        paths.rates.mkdir(parents=True, exist_ok=True)
        existing = read_table(cur_path) if cur_path else []
        merged = merge_rows(existing, parsed)
        log.info("Merged %s wage data", kind, extra={"details": {
            "rows_in_use_before": len(existing), "new_rows": len(parsed), "rows_in_use_after": len(merged),
            "closed_by_newer_rate": sum(1 for r in merged if r["expire_date"] and r["expire_date"] < OPEN_END),
            "tables_before": len(olds), "output": new_path.name}})
        old_sources = []
        if cur_path:
            try:
                s = openpyxl.load_workbook(cur_path, read_only=True)["sources"]
                old_sources = [list(r) for r in s.iter_rows(min_row=2, values_only=True)]
            except Exception:
                pass
        if new_path in olds:                   # merged twice in one run: keep the table as it was, then replace it
            archive_files(project_dir, [new_path], FILES, move=False, label="previous")
        sources = [list(x) + [None] * (len(SOURCE_COLUMNS) - len(x)) for x in old_sources] + \
            [[run_timestamp(tz), f.name, n.get(kind, 0), *info.get(f, (None, None))] for f, n, _ in used if n.get(kind)]
        write_table(new_path, merged, sources)
        archive_files(project_dir, [p for p in olds if p != new_path], FILES, move=True, label="previous")
        state.set(f"wages_{kind}_updated", run_timestamp(tz))
        summary.append(f"{kind}: merged {len(parsed)} rows from "
                       f"{sum(1 for _, n, _ in used if n.get(kind))} file(s) -> {len(merged)} rows in {FILES_DIR}/{RATES_DIR}/{new_path.name}")
        if len(olds) > 1:              # only the newest was in use: say the others were set aside, not merged
            summary.append(f"wages: {len(olds) - 1} older {kind} table(s) in {FILES_DIR}/{RATES_DIR} were not in use "
                           f"(only the newest is); moved to {FILES_DIR}/Archived without merging their rates")
    archive_files(project_dir, [f for f, _, _ in used] + [f for f, _ in already], FILES, move=True)
    for f, rec in already:
        when = rec.get("merged_at_utc") or "earlier"
        same = "" if rec.get("file") == f.name else f" as {rec.get('file')}"
        summary.append(f"wages: {f.name} skipped: already processed for this project ({when}{same}); "
                       "its rates are already in use")
        log.info("Wage file already processed for this project; skipped", extra={"details": {
            "type": f.suffix.lower(), "first_processed": when}})
    for f, n in partly:
        summary.append(f"wages: {f.name}: {n} rate(s) with a blank or unrecognised Jurisdiction were not used "
                       "(it must say State, Federal or Union)")
    if no_jur:
        summary.append(f"wages: no Jurisdiction column in {', '.join(f.name for f in no_jur)}; left in Uploads. "
                       "Add a Jurisdiction column (State, Federal or Union for each rate) and upload it again")
    if unreadable:
        summary.append(f"wages: could not read {', '.join(f.name for f in unreadable)}; left in Uploads. "
                       "No reader exists yet for this kind of wage file: send it to the developer so one can be built")
    return fetched + summary


# ----------------------------------------------------------------- lookup
def _name_parts(text):
    """'CRAFT / CLASSIFICATION' (either side may be empty) -> (norm craft, norm classification)."""
    craft, _, cls = str(text or "").partition(" / ") if " / " in str(text or "") else (str(text or ""), "", "")
    return norm(craft), norm(cls)


_COUNTY_WORDS = None


def _starts(name, prefix):
    """name starts with prefix; a prefix ending in a number must end there too ('group 1' is not 'group 10'), one
    ending in letters may be a word cut short ('operator und' matches 'operator underground')."""
    if not name.startswith(prefix):
        return False
    return not (prefix and prefix[-1].isdigit() and len(name) > len(prefix) and name[len(prefix)].isdigit())


def strip_noise(s):
    """A payroll name without WA county names, 'WA'/'County'/'State' and brackets: '(Pierce) Journey Level' ->
    'journey level'."""
    global _COUNTY_WORDS
    if _COUNTY_WORDS is None:
        _COUNTY_WORDS = {norm(c) for c in WA_COUNTIES} | {"wa", "county", "state", "washington"}
    return " ".join(w for w in norm(s).split() if w not in _COUNTY_WORDS)


class WageBook:
    def __init__(self, rows, fuzzy_threshold=0.88, name_matches=None):
        """name_matches: [{payroll: 'CRAFT / CLASSIFICATION', wage: 'CRAFT / CLASSIFICATION', kind: optional}] from
        payroll.name_matches and projects.<name>.name_matches in config.yaml. The payroll side matches a payroll
        line whose craft and classification start with the given text (so truncated names work); an empty side
        matches anything."""
        self.by_key = {}
        self.by_cls = {}
        self.by_alias = {}
        for r in rows:
            self.by_key.setdefault((r["kind"], match_key(r)), []).append(r)
            self.by_alias.setdefault((r["kind"], alias_key(r["craft"], r["classification"])), []).append(r)
            self.by_cls.setdefault((r["kind"], norm(r["classification"])), []).append(r)
        self.keys = {}
        for (kind, k) in self.by_key:
            self.keys.setdefault(kind, []).append(k)
        self.threshold = fuzzy_threshold
        self._fuzzy = {}
        self.name_matches = []
        self.bad_name_matches = []
        for m in name_matches or []:
            if not isinstance(m, dict) or not m.get("payroll") or not m.get("wage"):
                self.bad_name_matches.append(f"{m!r}: needs both 'payroll' and 'wage'")
                continue
            pc, pk = _name_parts(m["payroll"])
            wc, wk = _name_parts(m["wage"])
            kinds = [m["kind"]] if m.get("kind") else sorted({kd for kd, _ in self.by_key})
            for kind in kinds:
                target = self.by_key.get((kind, norm(f"{wc} {wk}")))
                if not target and wc and wk:
                    target = self.by_alias.get((kind, alias_key(wc, wk)))
                if target:
                    self.name_matches.append((kind, pc, pk, target, f"{m['payroll']} -> {m['wage']}"))
            if not any(nm[4] == f"{m['payroll']} -> {m['wage']}" for nm in self.name_matches) and self.by_key:
                self.bad_name_matches.append(f"'{m['wage']}' isn't in this project's wage data"
                                             + (f" ({m['kind']})" if m.get("kind") else ""))

    def __bool__(self):
        return bool(self.by_key)

    @staticmethod
    def _effective(cands, when):
        if not cands:
            return None, False
        if when:
            ok = [c for c in cands if (c["issue_date"] or dt.date.min) <= when <= (c["expire_date"] or OPEN_END)]
            if ok:
                return max(ok, key=lambda c: (c["issue_date"] or dt.date.min, c["mod"] or 0)), True
        return max(cands, key=lambda c: (c["issue_date"] or dt.date.min, c["mod"] or 0)), False

    def decisions(self, kind):
        return sorted({str(r["wage_decision"]) for (kd, _), rs in self.by_key.items() if kd == kind for r in rs})

    @staticmethod
    def _scope(cands, decision, county):
        """Narrows candidates to the wage determination / county the Jurisdiction names. Returns (cands, note):
        note is "" when narrowed (or nothing was asked), else says what wasn't found."""
        note = []
        if cands and decision:
            hit = [c for c in cands if decision.lower() in str(c["wage_decision"] or "").lower()]
            if hit:
                cands = hit
            else:
                note.append(f"no rate in {decision}")
        if cands and county:
            with_county = [c for c in cands if county_key(c)]
            hit = [c for c in with_county if county_key(c) == county]
            if hit:
                cands = hit
            elif with_county:
                note.append(f"no rate for {county.title()} County")
        elif cands:
            # no county named (nor set for the project): fine if every county pays the same, otherwise say so
            by_county = {}
            for c in cands:
                if county_key(c):
                    by_county.setdefault(county_key(c), set()).add((c.get("basic"), c.get("total")))
            if len({frozenset(v) for v in by_county.values()}) > 1:
                note.append(f"no county named and the rate differs by county ({len(by_county)} counties); set "
                            "projects.<project>.county in config.yaml")
        return cands, "; ".join(note)

    def lookup(self, kind, craft, classification, when, decision=None, county=None):
        """Returns (rate_row, how, in_range): how is "exact", "name match (config.yaml): ...", "same name, split
        differently", "same name without county/WA words", "classification name only" or "approx:<score>", followed by
        " | <note>" when the wage determination or county named by the Jurisdiction had no rate (another one was
        used); in_range says whether the rate's dates cover `when`."""
        k = norm(f"{craft} {classification}")
        cands = self.by_key.get((kind, k))
        how = "exact"
        if not cands:
            # names an operator mapped in config.yaml (payroll.name_matches / projects.<name>.name_matches)
            nc, nk = norm(craft), norm(classification)
            for kd, pc, pk, target, rule in self.name_matches:
                if kd == kind and _starts(nc, pc) and _starts(nk, pk):
                    cands, how = target, f"name match (config.yaml): {rule}"
                    break
        if not cands:
            # the same name split differently between craft and classification
            alias = self.by_alias.get((kind, alias_key(craft, classification)))
            if alias and len({match_key(c) for c in alias}) == 1:
                cands, how = alias, "same name, split differently"
        if not cands:
            # the same name once county names, 'WA', 'County' and brackets are left out
            sc, sk = strip_noise(craft), strip_noise(classification)
            if (sc, sk) != (norm(craft), norm(classification)) and sk:
                got = self.by_key.get((kind, norm(f"{sc} {sk}"))) or self.by_alias.get((kind, alias_key(sc, sk)))
                if got and len({match_key(c) for c in got}) == 1:
                    cands, how = got, "same name without county/WA words"
        if not cands:
            cands = self.by_cls.get((kind, norm(classification)))
            if cands and len({match_key(c) for c in cands}) > 1:
                # same classification name under several crafts: prefer the closest craft
                cands = sorted(cands, key=lambda c: -SequenceMatcher(None, norm(c["craft"]), norm(craft)).ratio())
                top = match_key(cands[0])
                cands = [c for c in cands if match_key(c) == top]
            how = "classification name only"
        if not cands:
            if (kind, k) not in self._fuzzy:
                best, score = None, 0.0
                nums = re.findall(r"\d+", k)
                for cand in self.keys.get(kind, []):
                    if re.findall(r"\d+", cand) != nums:     # 'group 1' is never close to 'group 10'
                        continue
                    sm = SequenceMatcher(None, k, cand)
                    if sm.real_quick_ratio() < self.threshold or sm.quick_ratio() < self.threshold:
                        continue
                    s = sm.ratio()
                    if s > score:
                        best, score = cand, s
                self._fuzzy[(kind, k)] = (best, score)
            best, score = self._fuzzy[(kind, k)]
            if best and score >= self.threshold:
                cands, how = self.by_key[(kind, best)], f"approx:{score:.2f}"
        note = ""
        if cands:
            cands, note = self._scope(cands, decision, county)
        r, in_range = self._effective(cands, when)
        return r, how + (f" | {note}" if note else ""), in_range

    def candidates(self, kind, limit=400):
        return sorted({f"{r['craft']} | {r['classification']}" for (kd, _), rs in self.by_key.items()
                       if kd == kind for r in rs})[:limit]


def _reopened(typed, untyped):
    """Typed rates whose end date was set only because a newer untyped rate of the same name followed (the day
    before it starts), in tables merged before the Jurisdiction rule: open-ended again, unless a newer typed rate
    follows (that one closes it)."""
    cut = {(r["kind"], match_key(r), county_key(r), r["issue_date"] - dt.timedelta(days=1))
           for r in untyped if r.get("issue_date")}
    out = []
    for r in typed:
        key = (r["kind"], match_key(r), county_key(r), r.get("expire_date"))
        later_typed = any(x is not r and (x["kind"], match_key(x), county_key(x)) == key[:3] and x.get("issue_date")
                          and r.get("issue_date") and x["issue_date"] > r["issue_date"] for x in typed)
        out.append({**r, "expire_date": OPEN_END} if key in cut and not later_typed else r)
    return out


def load_wagebook(project_dir: Path, cfg):
    rows, untyped, untyped_kinds = [], 0, set()
    for kind in cfg["wages"]["kinds"]:
        from .layout import Paths
        got = read_table(Paths(project_dir).rate_file(kind))
        # only rates whose wage type was given (a Jurisdiction column, or an official determination): rates merged
        # from a file without the column (before Nov 7, 2026) are not used (David: strict)
        rows += _reopened([r for r in got if r.get("jurisdiction")], [r for r in got if not r.get("jurisdiction")])
        n = sum(1 for r in got if not r.get("jurisdiction"))
        untyped += n
        if n:
            untyped_kinds.add(kind)
    # the wage types the project's own tables give (staff provide 1-3 per project; David, Nov 12, 2026: "stick to the
    # rates provided by the staff for that project"): a payroll line whose Jurisdiction doesn't say its type is checked
    # as each of these. Federal rates downloaded earlier from SAM.gov into the project's folder count as its own (they
    # are its tables; downloads are off by default since Nov 12, and no project lists a determination today).
    own_kinds = [k for k in cfg["wages"]["kinds"] if any(r["kind"] == k for r in rows)]
    # the firm-wide union table stands in only when the project has no union rates at all, not for its own union
    # rates that lack a Jurisdiction (those lines are told as such instead); off by default since Nov 12, 2026
    if not any(r["kind"] == "union" for r in rows) and "union" not in untyped_kinds and cfg["wages"].get("union_fallback", False):
        from .reference import load_union_reference      # union table extracted from the reference rules workbook
        rows += load_union_reference(cfg)
    from .config import project_cfg
    matches = list(project_cfg(cfg, project_dir.name).get("name_matches") or []) + list(cfg["payroll"].get("name_matches") or [])
    book = WageBook(rows, cfg["payroll"].get("fuzzy_match_threshold", 0.88), matches)
    book.untyped = untyped                               # rates left out: no Jurisdiction (for the Audit sheet)
    book.untyped_kinds = untyped_kinds
    book.own_kinds = own_kinds                           # never the firm-wide union table
    return book
