"""
Tokenize and detokenize .xlsx / .xlsm files by editing their XML directly.

Why not open the file with openpyxl and save it again? Because openpyxl drops the cached results of
formulas when it saves, so a checker that reads values (data_only) would see empty cells, and it also
loses some formatting. Editing only what holds personal data leaves everything else as it was: column
positions, formulas, cached values, styles, sheet order.

An .xlsx file is a zip of XML parts. The ones that hold cell data:
  xl/sharedStrings.xml   every distinct text value, referenced by number from the cells (t="s")
  xl/worksheets/*.xml    the cells: <c r="G5" s="3" t="s"><v>17</v></c>
Personal-data cells are rewritten as inline text holding the token:
  <c r="G5" s="3" t="inlineStr"><is><t>EMP-…</t></is></c>

Every other part is handled by an ALLOW-LIST (``_part_rule``): parts known to hold no user text are kept,
parts that can hold text (comments, charts, drawings, workbook and sheet names, document properties, sheet
headers and footers, data-validation lists, formulas) have every text value and attribute scrubbed, and
anything else, including pivot-table, slicer and external-link caches (which keep their own hidden copy of
the data) and embedded objects, makes the whole file be refused. Refused files are not copied to AWS.
"""
from __future__ import annotations

import io
import posixpath
import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from . import doubt
from .config import HeaderTracker
from .tokens import TOKEN_HEX, TOKEN_RE, TokenError, normalise, unglue

CELL_RE = re.compile(r'<c\b((?:[^>"/]|"[^"]*"|/(?!>))*)(?:/>|>(.*?)</c>)', re.S)
ATTR_RE = re.compile(r'([\w:]+)="([^"]*)"')
SI_RE = re.compile(r"<si\b[^>]*?(?:/>|>(.*?)</si>)", re.S)
T_RE = re.compile(r"<t\b[^>]*?(?:/>|>(.*?)</t>)", re.S)
RPH_RE = re.compile(r"<rPh\b.*?</rPh>", re.S)
V_RE = re.compile(r"<v\b[^>]*>(.*?)</v>", re.S)
REF_RE = re.compile(r"([A-Z]{1,3})(\d+)$")
TEXT_NODE_RE = re.compile(r">([^<]+)<")
# The same, used when scrubbing, with what can never be personal text skipped inside the pattern itself (so it
# never reaches Python): plain numbers, and tokens written by the first pass. A sheet has hundreds of thousands.
SCRUB_TEXT_NODE_RE = re.compile(r">(?!\s*[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?\s*<)"
                                r"(?![A-Z]{2,5}-[0-9a-f]{%d}<)([^<]+)<" % TOKEN_HEX)
ATTR_VALUE_RE = re.compile(r'(\s[\w:.-]+=")([^"]*)(")')
SCRUB_ATTR_RE = re.compile(r'(\s(?!(?:r|ref|sqref|spans|s|t|ht|cm|vm)=")[\w:.-]+=")([^"]*)(")')
NUMBER_RE = re.compile(r"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?")
# Cell addresses and ranges ("AH2026", "$A$1:B7", "A1 C3:D4"): an ID that looks like one must not change them.
CELL_REF_RE = re.compile(r"(\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?)( \$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?)*")
STRUCTURAL_ATTRS = {"r", "ref", "sqref", "spans", "s", "t", "ht", "cm", "vm", "topLeftCell", "activeCell",
                    "xr:uid", "xr2:uid", "xr3:uid", "x14ac:dyDescent", "customHeight", "customFormat"}
PREFIXED_ROOT_RE = re.compile(r"<\w+:(worksheet|sst|workbook|comments)\b")

REL_SHEET = "/worksheet"
REL_SST = "/sharedStrings"
_ENT = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


_ENTITY_RE = re.compile(r"&(#[xX][0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);")
_INT_RE = re.compile(r"-?\d+")
_COL_CACHE: dict = {}


# A tag, with quoted attribute values read whole: a ">" is legal inside an attribute value (Excel and Google Sheets
# can write it there), so "<[^>]*>" would end the tag too early. Comments and CDATA come first and are kept as they
# are: a quote inside them must not be taken for the start of a value.
_TAG_RE = re.compile(r"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<(?:[^>\"']|\"[^\"]*\"|'[^']*')*>", re.S)
# Quick test: does any tag need rewriting? A ">" inside a quoted value, a single-quoted value, or spaces around "=".
# Most parts have none and are not rewritten. A false alarm (the same shapes in text) only costs a full pass.
_NEEDS_CANON_RE = re.compile(r"=\s*\"[^\"<]*>|=\s*'|\s=|=\s+[\"']")
# One attribute, in any form XML allows: name = "value" or name = 'value'.
_ANY_ATTR_RE = re.compile(r"(\s)([\w:.-]+)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')")


def canonical_attributes(xml: str) -> str:
    """The same XML with every attribute written as name="value", and a ">" or '"' inside a value as "&gt;" or
    "&quot;" (it means the same). The passes below find text with ">...<" and attribute values with 'name="..."' by
    pattern. A raw ">" inside an attribute made them start inside the tag and swallow the rest of it and the next tag
    (QRX's tracker, Nov 4: a drop-down message with ">" became error="TXT-..."list" ... and the sheet could not be
    read); a value in single quotes, or with spaces around "=", was never tokenized at all."""
    if not _NEEDS_CANON_RE.search(xml):
        return xml
    def attr(m):
        v = m.group(3) if m.group(3) is not None else m.group(4)
        v = v.replace(">", "&gt;").replace('"', "&quot;")
        return f'{m.group(1)}{m.group(2)}="{v}"'
    def tag(m):
        t = m.group(0)
        if t.startswith(("<!--", "<![CDATA[", "<?", "<!")):
            return t
        return _ANY_ATTR_RE.sub(attr, t)
    return _TAG_RE.sub(tag, xml)


def unescape(s: str) -> str:
    if "&" not in s:
        return s
    def rep(m):
        g = m.group(1)
        if g[0] == "#":
            return chr(int(g[2:], 16) if g[1] in "xX" else int(g[1:]))
        return _ENT[g]
    return _ENTITY_RE.sub(rep, s)


def escape(s: str) -> str:
    """For text nodes and attribute values alike. Line breaks and tabs as character references: raw ones inside an
    attribute would be read back as spaces, so a restored value would not be exactly the original."""
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;").replace("\n", "&#10;").replace("\r", "&#13;").replace("\t", "&#9;"))


# ------------------------------------------------------------------------------------------------
# Which parts may be in a file, and what happens to them
# ------------------------------------------------------------------------------------------------
_KEEP = [r"\[Content_Types\]\.xml", r"xl/styles\.xml", r"xl/theme/[^/]+\.xml", r"xl/calcChain\.xml",
         r"xl/printerSettings/[^/]+", r"xl/media/[^/]+", r"xl/metadata\.xml", r"xl/vbaProject\.bin",
         r"xl/vbaProjectSignature\.bin", r"xl/sheetMetadata\.xml",
         # Google Sheets writes this into every workbook it saves or exports: a few bytes of binary settings
         # (format version, locale, time zone, default font), no cell text.
         r"xl/metadata"]
_SCRUB = [r"(.*/)?_rels/[^/]*\.rels", r"xl/workbook\.xml", r"docProps/(core|app|custom)\d*\.xml",
          r"xl/drawings/[^/]+\.(xml|vml)", r"xl/charts/[^/]+\.xml", r"xl/tables/[^/]+\.xml",
          r"xl/persons/[^/]+\.xml", r"xl/connections\.xml", r"xl/queryTables/[^/]+\.xml",
          r"customXml/[^/]+\.xml", r"xl/comments[^/]*\.xml", r"xl/comments/[^/]+\.xml",
          r"xl/threadedComments/[^/]+\.xml",
          # Excel "assigned tasks" (task text, assignee names and e-mails): scrubbed like threaded comments
          r"xl/documenttasks/[^/]+\.xml",
          r"xl/ctrlProps/[^/]+\.xml", r"xl/webextensions/[^/]+\.xml"]
_DROP = [r"docProps/thumbnail\.\w+",         # a picture of the first sheet: removed from the copy
         # Google Sheets writes one of these per comments part when it saves or exports a sheet that has comments
         # (xl/commentsmeta0, xl/commentsmeta1 ...; its own relationship type, workbookmetadata). Binary, not part of
         # the Excel format, and it repeats the comments' text and their authors. The tokenizer can't read it, so
         # the copy doesn't have it (the comments themselves stay in xl/comments*.xml, tokenized as before); Google
         # rebuilds it the next time the sheet is saved. Found Nov 5 2026: a tracker saved through Google Sheets
         # was held back as "does not inspect xl/commentsmeta1" and the project showed "No tracker".
         r"xl/commentsmeta\d+"]
_KEEP_RE = [re.compile(p + "$") for p in _KEEP]
_SCRUB_RE = [re.compile(p + "$") for p in _SCRUB]
_DROP_RE = [re.compile(p + "$") for p in _DROP]


def _part_rule(name: str) -> str:
    if any(p.match(name) for p in _KEEP_RE):
        return "keep"
    if any(p.match(name) for p in _SCRUB_RE):
        return "scrub"
    if any(p.match(name) for p in _DROP_RE):
        return "drop"
    return "refuse"


def _inline(attrs: dict, text: str) -> str:
    keep = "".join(f' {k}="{attrs[k]}"' for k in ("r", "s") if k in attrs)
    return f'<c{keep} t="inlineStr"><is><t xml:space="preserve">{escape(text)}</t></is></c>'


def _empty(attrs: dict) -> str:
    keep = "".join(f' {k}="{attrs[k]}"' for k in ("r", "s") if k in attrs)
    return f"<c{keep}/>"


def _col_index(letters: str) -> int:
    n = _COL_CACHE.get(letters)
    if n is None:
        n = 0
        for ch in letters:
            n = n * 26 + ord(ch) - 64
        _COL_CACHE[letters] = n
    return n


@dataclass(slots=True)
class Cell:
    start: int
    end: int
    attrs: dict
    inner: str
    col: int
    row: int
    kind: str                 # s | inlineStr | str | num | other
    value: object             # text or number; None when empty
    sst_index: int | None = None
    has_formula: bool = False


@dataclass
class Report:
    columns: dict = field(default_factory=lambda: defaultdict(Counter))   # sheet part -> {field kind: cells}
    tokenized: int = 0
    dropped: int = 0
    free_text: int = 0
    doubt: int = 0           # text the agent doesn't read, replaced by text tokens (doubt.py)
    restored: int = 0
    unknown_tokens: set = field(default_factory=set)
    warnings: list = field(default_factory=list)

    def summary(self) -> str:
        cols = "; ".join(f"{s.rsplit('/', 1)[-1]}: " + ", ".join(f"{k} {n}" for k, n in c.items())
                         for s, c in self.columns.items() if c)
        out = []
        if self.tokenized or self.dropped:
            out.append(f"{self.tokenized} cells tokenized, {self.dropped} emptied ({cols})")
        if self.free_text:
            out.append(f"{self.free_text} free-text values changed")
        if self.doubt:
            out.append(f"{self.doubt} values the agent doesn't read replaced by text tokens")
        if self.restored:
            out.append(f"{self.restored} tokens restored")
        if self.unknown_tokens:
            out.append(f"{len(self.unknown_tokens)} unknown tokens left as they are")
        out += self.warnings
        return "; ".join(out) or "no personal data found"


# ------------------------------------------------------------------------------------------------
# Reading
# ------------------------------------------------------------------------------------------------
def _si_text(inner: str | None) -> str:
    if not inner:
        return ""
    inner = RPH_RE.sub("", inner)
    return "".join(unescape(m.group(1) or "") for m in T_RE.finditer(inner))


MAX_PART_MB = 600      # unpacked size of one part: checking one takes about 3 times this in memory (the job has 4 GB)
MAX_TOTAL_MB = 1000


def _read_zip(data: bytes):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        infos = z.infolist()
    except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError) as e:
        raise TokenError("not a valid .xlsx file") from e
    biggest, total = max((i.file_size for i in infos), default=0), sum(i.file_size for i in infos)
    if biggest > MAX_PART_MB * 1_000_000 or total > MAX_TOTAL_MB * 1_000_000:
        raise TokenError(f"it is too large to check ({total // 1_000_000} MB of data once unpacked); split it "
                         "into smaller files (for example one per year or per project) and upload those")
    try:
        parts = {i.filename: z.read(i.filename) for i in infos}
    except (zipfile.BadZipFile, OSError, RuntimeError, NotImplementedError) as e:
        raise TokenError("not a valid .xlsx file") from e
    return infos, parts


def _write_zip(infos, parts: dict) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            if info.filename in parts:
                z.writestr(info, parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
    return out.getvalue()


def _text(parts, name) -> str:
    try:
        return parts[name].decode("utf-8")
    except UnicodeDecodeError as e:
        raise TokenError(f"{name} is not UTF-8 XML") from e


def _rels(parts, part: str) -> list[tuple[str, str]]:
    """(relationship type, target part path) for a part, from its .rels file."""
    d, b = posixpath.split(part)
    rels = posixpath.join(d, "_rels", b + ".rels")
    if rels not in parts:
        return []
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", _text(parts, rels)):
        a = dict(ATTR_RE.findall(m.group(1)))
        if a.get("TargetMode") == "External":
            continue
        tgt = a.get("Target", "")
        path = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
        out.append((a.get("Type", ""), path))
    return out


def _layout(parts) -> tuple[list[str], str | None]:
    """The worksheet parts and the shared-strings part, found the way Excel finds them: through the
    relationship files, not by assuming the usual names."""
    wb = next((p for t, p in _rels(parts, "") if t.endswith("/officeDocument")), None) \
        if "_rels/.rels" in parts else None
    if not wb or wb not in parts:
        raise TokenError("the workbook part could not be found")
    rels = _rels(parts, wb)
    sheets = [p for t, p in rels if t.endswith(REL_SHEET) and p in parts]
    sst = next((p for t, p in rels if t.endswith(REL_SST) and p in parts), None)
    if not sheets:
        raise TokenError("no worksheets found")
    for p in [wb, *sheets] + ([sst] if sst else []):
        if PREFIXED_ROOT_RE.search(_text(parts, p)[:2000]):
            raise TokenError("written with namespace prefixes (x:c, x:row), which the tokenizer does not read; "
                             "open it in Excel or Google Sheets, save it again, and upload that copy")
    return sheets, sst


def _sheet_titles(parts, sheet_parts: list[str]) -> list[str]:
    """Each worksheet part's name as shown on its tab ("" if it can't be found)."""
    wb = next((p for t, p in _rels(parts, "") if t.endswith("/officeDocument")), None)
    if not wb or wb not in parts:
        return [""] * len(sheet_parts)
    d = posixpath.dirname(wb)
    rels = posixpath.join(d, "_rels", posixpath.basename(wb) + ".rels")
    by_id = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", _text(parts, rels) if rels in parts else ""):
        a = dict(ATTR_RE.findall(m.group(1)))
        tgt = a.get("Target", "")
        by_id[a.get("Id")] = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
    names = {}
    for m in re.finditer(r"<sheet\b([^>]*)/?>", _text(parts, wb)):
        a = dict(ATTR_RE.findall(m.group(1)))
        rid = next((v for k, v in a.items() if k.endswith(":id")), None)
        if rid in by_id:
            names[by_id[rid]] = unescape(a.get("name", ""))      # dicts keep insertion order: tab order
    return [names.get(sp, "") for sp in sheet_parts]


def _workbook_part(parts) -> str | None:
    wb = next((p for t, p in _rels(parts, "") if t.endswith("/officeDocument")), None) if "_rels/.rels" in parts else None
    return wb if wb in parts else None


def _tab_order(parts, sheet_parts: list[str]) -> list[str]:
    """The worksheet parts in the order of the workbook's tabs (as openpyxl's wb.worksheets lists them)."""
    wb = next((p for t, p in _rels(parts, "") if t.endswith("/officeDocument")), None)
    if not wb or wb not in parts:
        return list(sheet_parts)
    rels = dict((rid, p) for rid, p in _rel_ids(parts, wb))
    order = []
    for m in re.finditer(r"<sheet\b([^>]*)/?>", _text(parts, wb)):
        a = dict(ATTR_RE.findall(m.group(1)))
        rid = next((v for k, v in a.items() if k.endswith(":id")), None)
        if rels.get(rid) in sheet_parts and rels[rid] not in order:
            order.append(rels[rid])
    return order + [p for p in sheet_parts if p not in order]


def _rel_ids(parts, part: str) -> list[tuple[str, str]]:
    """(relationship id, target part path) for a part, internal targets only."""
    d, b = posixpath.split(part)
    rels = posixpath.join(d, "_rels", b + ".rels")
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", _text(parts, rels) if rels in parts else ""):
        a = dict(ATTR_RE.findall(m.group(1)))
        if a.get("TargetMode") == "External":
            continue
        tgt = a.get("Target", "")
        out.append((a.get("Id"), tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))))
    return out


def _head(xml: str, sst: list[str], last_row: int) -> list:
    """The first rows of a sheet as [(row number, {column: value})], up to ``last_row``."""
    out = []
    for r, row in _rows(xml, sst):
        if r > last_row:
            break
        out.append((r, {c.col: c.value for c in row if c.value is not None}))
    return out


_F_RE = re.compile(r"(<f\b[^>]*>)(.*?)(</f>)", re.S)


# A cell's stored value, by its type (#124): what a valid one looks like. Anything else in a typed cell is text.
_ISO_DATE_RE = re.compile(r"-?\d{4,}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?"
                          r"|\d{2}:\d{2}(:\d{2}(\.\d+)?)?")
_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!",
           "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!", "#EXTERNAL!", "#PYTHON!"}


def _raw_value(c: Cell):
    """The cell's stored value (<v>) as written, unescaped, or None."""
    m = V_RE.search(c.inner or "")
    return unescape(m.group(1)) if m else None


def _typed_text(c: Cell) -> str | None:
    """#124: the text stored in a cell typed as a number, date, yes/no or error (or an unknown type) whose value isn't
    one of that type, or None. Excel never writes such a cell; a hand-built or third-party file can."""
    if c.kind in ("s", "inlineStr", "str", "num"):
        return None
    v = _raw_value(c)
    if v is None or v == "":
        return None
    t = c.attrs.get("t", "n")
    if t == "b" and v.strip() in ("0", "1"):
        return None
    if t == "d" and _ISO_DATE_RE.fullmatch(v.strip()):
        return None
    if t == "e" and v.strip() in _ERRORS:
        return None
    return v


# A chart's cached values (#112), with or without a namespace prefix: c:numCache, c:strCache, c:multiLvlStrCache, the
# "value from cells" data labels' c15:dlblRangeCache, and a chartEx's cx:lvl (its data points)
_CHART_CACHE_RE = re.compile(r"<((?:\w+:)?)(numCache|strCache|multiLvlStrCache|dlblRangeCache|lvl)\b[^>]*?"
                             r"(?:/>|>.*?</\1\2>)", re.S)
# A number written in a formula (not a cell reference such as A1, $B$2 or Sheet2!C3): data, under #112
_NUM_LITERAL_RE = re.compile(r"(?<![A-Za-z_$\d.!:])\d+(?:\.\d+)?(?:[eE][+-]?\d+)?(?![\d(!])")


def _doubt_cell(c: Cell, ctx, rep: Report, read: bool = False, strict: bool = False) -> str | None:
    """The cell rewritten with its text as a text token (doubt.py), or None if it holds no text. A formula keeps
    its place: its cached text, and its own text when it has a quoted string, become tokens instead.
    vnext.doubt_typed_cells (#124): text stored in a cell typed as a number, date, yes/no or error becomes a token too
    (always in a sheet with a strict rule, ``strict``). vnext.doubt_numbers (#112): numbers and dates too, outside the
    agent's columns (``read``: a cell of a column the agent reads that its strict rule didn't keep, e.g. a formula
    rate: its cached number stays as before); a number's token comes back as the number on the staff side, a formula's
    number literals make its text a token, its cached number a number token."""
    cfg = ctx.config
    numbers = cfg.doubt_numbers and not read
    def tok(text: str) -> str:
        t = ctx.keyring.text_token(text)
        ctx.vault.add_text(t, text)
        return t
    def num_tok(value) -> str:
        t = ctx.keyring.token("NUM", "number", value)
        ctx.vault.add(t, "number", value)
        return t
    typed = _typed_text(c) if (cfg.doubt_typed_cells or strict) else None
    if typed is not None and not c.has_formula:
        rep.doubt += 1
        return _inline(c.attrs, tok(typed))
    if numbers and not c.has_formula:
        if c.kind == "num" and isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
            rep.doubt += 1
            return _inline(c.attrs, num_tok(c.value))
        if c.attrs.get("t") == "d" and _raw_value(c):
            rep.doubt += 1                                  # (comes back as text: Excel doesn't write t="d")
            return _inline(c.attrs, tok(_raw_value(c)))
    if c.kind in ("s", "inlineStr", "str") and not c.has_formula:
        if not isinstance(c.value, str) or not c.value.strip():
            return None
        rep.doubt += 1
        return _inline(c.attrs, tok(c.value))
    if not c.has_formula:
        return None
    inner, attrs = c.inner, dict(c.attrs)
    if c.kind == "str" and isinstance(c.value, str) and c.value.strip():
        inner = V_RE.sub(lambda m: f"<v>{tok(unescape(m.group(1)))}</v>", inner, count=1)
    elif typed is not None:
        # a cached value that is text in a typed cell (#124): a text token, typed as text
        inner = V_RE.sub(lambda m: f"<v>{tok(unescape(m.group(1)))}</v>", inner, count=1)
        attrs["t"] = "str"
    elif numbers and c.kind == "num" and isinstance(c.value, (int, float)):
        # a cached number (#112): a number token, typed as text here; it comes back as the number (detokenize)
        inner = V_RE.sub(lambda m: f"<v>{num_tok(c.value)}</v>", inner, count=1)
        attrs["t"] = "str"
    elif numbers and attrs.get("t") == "d" and _raw_value(c):
        inner = V_RE.sub(lambda m: f"<v>{tok(unescape(m.group(1)))}</v>", inner, count=1)
        attrs["t"] = "str"
    def ftext(m):
        text = unescape(m.group(2))
        if '"' in text or (numbers and _NUM_LITERAL_RE.search(text)):
            return m.group(1) + tok(text) + m.group(3)
        return m.group(0)
    inner = _F_RE.sub(ftext, inner, count=1)
    if inner == c.inner and attrs == c.attrs:
        return None
    rep.doubt += 1
    attrs = "".join(f' {k}="{v}"' for k, v in attrs.items())
    return f"<c{attrs}>{inner}</c>"


def _keep(rule, r: int, col: int, row_vals: dict, cell=None, sst_rph=frozenset()) -> bool:
    """Does the agent read this cell (doubt.py)? Then it is handled as before; otherwise it becomes a text token."""
    if rule.kind == "agent":
        return True
    if rule.kind == "tracker":
        if rule.strict and r > rule.header_row and col in rule.keep_cols and cell is not None:
            # vNext, strict trackers (#128): a value of the column's type only (cva_filekinds.tracker.value_ok). No
            # hidden phonetic text; a formula only without quoted text and with a cached value that isn't text.
            if "<rPh" in (cell.inner or "") or (cell.kind == "s" and cell.sst_index in sst_rph):
                return False
            if _typed_text(cell) is not None:
                return False
            if cell.has_formula:                     # no quoted text in it; then its cached value, as a plain cell's
                fm = _F_RE.search(cell.inner or "")
                if fm and '"' in unescape(fm.group(2)):
                    return False
            return doubt.tracker_value_ok(rule.roles.get(col, ""), cell.value, rule.status_words)
        if r > rule.header_row and col in rule.vocab and cell is not None:
            # a fixed word only, as a plain value: a formula (its quoted strings could name someone; _doubt_cell keeps
            # it in place and leaves a cached TRUE/FALSE or number readable) or hidden phonetic text is not kept
            if cell.has_formula or "<rPh" in (cell.inner or "") or (cell.kind == "s" and cell.sst_index in sst_rph):
                return False
            if isinstance(cell.value, str):
                return bool(rule.vocab[col].fullmatch(cell.value))   # any other text is a token
        if r == rule.header_row:
            return not rule.head_only_read or col in rule.keep_cols
        return r > rule.header_row and col in rule.keep_cols
    if rule.kind == "wage" and rule.strict:
        # vNext, strict wage tables (#118): only what agent/wages.py reads (cva_filekinds.wage_table). No formula or
        # hidden phonetic text anywhere (_doubt_cell keeps a formula in place and leaves a cached number readable).
        # Above the header: only an LCPtracker export's "Location: <county>" line in column A. Header row: the agent's
        # columns (and a short label with "classification" in it). Below it, in the agent's columns: a Jurisdiction
        # value on any row; anything else only on a rate row (a Base/Basic or Total amount), and only a value of the
        # column's type (value_ok).
        vals = row_vals or {}
        v = vals.get(col)
        if cell is not None and (cell.has_formula or "<rPh" in (cell.inner or "")
                                 or (cell.kind == "s" and cell.sst_index in sst_rph)):
            return False
        if r < rule.header_row:
            return col == 1 and isinstance(v, str) and bool(doubt.LOCATION_RE.fullmatch(v))
        if r == rule.header_row:
            return col in rule.head_cols
        if col not in rule.keep_cols:
            return False
        role = rule.roles.get(col, "")
        if role == "jur":
            return doubt.value_ok(role, v)
        return any(doubt.is_amount(vals.get(c)) for c in rule.rate_cols) and doubt.value_ok(role, v)
    if rule.kind == "wage":
        # the header row and the agent's wage columns below it; above it, only an LCPtracker export's "Location:"
        # line in column A (agent/wages.py parse_lcp_export reads the county from it)
        if r >= rule.header_row:
            return r == rule.header_row or col in rule.keep_cols
        v = (row_vals or {}).get(col)
        return col == 1 and isinstance(v, str) and bool(doubt.LOCATION_RE.fullmatch(v))
    if rule.kind == "info":
        return doubt.keep_info_cell(row_vals, col)
    if rule.kind in ("payroll", "merged", "decisions"):
        # #53 (vnext.fail_closed_tables): the header row's cells and the columns the agent reads below it; in a
        # payroll, only on its lines (a footer or a note below the table is a token). A formula with quoted text, or
        # hidden phonetic text, is not kept (_doubt_cell keeps a formula in place and leaves a cached number readable)
        if r < rule.header_row or col not in rule.keep_cols:
            return False
        if cell is not None and ("<rPh" in (cell.inner or "") or (cell.kind == "s" and cell.sst_index in sst_rph)):
            return False
        if cell is not None and cell.has_formula:
            fm = _F_RE.search(cell.inner or "")
            if fm and '"' in unescape(fm.group(2)):
                return False
        if r == rule.header_row or rule.kind != "payroll":
            return True
        vals = row_vals or {}
        if rule.roles.get(col) not in doubt.PAYROLL_TEXT_ROLES and not doubt.number_text_ok(vals.get(col)) \
                and doubt.norm_header(vals.get(col)) != rule.head_labels.get(col):
            return False                         # a name typed into an hours or money column (its own header label,
            #                                      pasted again with a second export, stays: upload_checks finds it)
        return any(vals.get(c) not in (None, "") for c in rule.line_cols)
    return False


def _opaque_comments(xml: str, ctx, rep: Report) -> str:
    """Cell notes in a workbook in doubt mode: each note's text, and every author but the agent, become text tokens
    (the agent finds and removes its own notes by its author name)."""
    def tok(text: str) -> str:
        t = ctx.keyring.text_token(text)
        ctx.vault.add_text(t, text)
        rep.doubt += 1
        return t
    def author(m):
        name = unescape(m.group(2))
        if not name.strip() or name in doubt.AGENT_AUTHORS or TOKEN_RE.fullmatch(name):
            return m.group(0)
        return m.group(1) + escape(tok(name)) + m.group(3)
    xml = re.sub(r"(<author\b[^>]*>)(.*?)(</author>)", author, xml, flags=re.S)
    def text(m):
        t = _si_text(m.group(2))
        if not t.strip() or TOKEN_RE.fullmatch(t):
            return m.group(0)
        return f'{m.group(1)}<t xml:space="preserve">{tok(t)}</t></text>'
    return re.sub(r"(<text\b[^>]*>)(.*?)</text>", text, xml, flags=re.S)


# Attributes outside the cells that can hold staff text: data-validation messages, hyperlink text, filter values.
_OUTSIDE_ATTRS = r"val|display|tooltip|prompt|promptTitle|error|errorTitle|name|text"
_GUID_RE = re.compile(r"\{[0-9A-Fa-f-]{36}\}")      # Excel's IDs (conditional formats, data validations)


def _opaque_outside_cells(xml: str, ctx, rep: Report) -> str:
    """A worksheet in doubt mode, outside <sheetData>: every text (data-validation lists and formulas, conditional
    formatting formulas, headers and footers) and the attributes above become text tokens. The cells are done."""
    i, j = xml.find("<sheetData"), xml.find("</sheetData>")
    if i < 0:
        return xml
    j = j + len("</sheetData>") if j >= 0 else len(xml)
    def tok(text: str) -> str:
        t = ctx.keyring.text_token(text)
        ctx.vault.add_text(t, text)
        rep.doubt += 1
        return t
    def node(m):
        text = unescape(m.group(1))
        s = text.strip()
        if not s or TOKEN_RE.fullmatch(s) or NUMBER_RE.fullmatch(s) or CELL_REF_RE.fullmatch(s) or _GUID_RE.fullmatch(s):
            return m.group(0)
        return f">{escape(tok(text))}<"
    def attr(m):
        v = unescape(m.group(3))
        s = v.strip()
        if not s or TOKEN_RE.fullmatch(s) or NUMBER_RE.fullmatch(s) or CELL_REF_RE.fullmatch(s):
            return m.group(0)
        return f'{m.group(1)}{m.group(2)}="{escape(tok(v))}"'
    def part(x: str) -> str:
        x = TEXT_NODE_RE.sub(node, x)
        return re.sub(rf'(\s)({_OUTSIDE_ATTRS})="([^"]*)"', attr, x)
    return part(xml[:i]) + xml[i:j] + part(xml[j:])


def _opaque_threaded(xml: str, ctx, rep: Report) -> str:
    """Threaded comments, their people and assigned tasks in doubt mode: every text and every name or user ID is a
    text token."""
    def tok(text: str) -> str:
        t = ctx.keyring.text_token(text)
        ctx.vault.add_text(t, text)
        rep.doubt += 1
        return t
    def node(m):
        text = unescape(m.group(1))
        if not text.strip() or TOKEN_RE.fullmatch(text.strip()) or NUMBER_RE.fullmatch(text.strip()):
            return m.group(0)
        return f">{escape(tok(text))}<"
    xml = TEXT_NODE_RE.sub(node, xml)
    def attr(m):
        v = unescape(m.group(3))
        if not v.strip() or TOKEN_RE.fullmatch(v):
            return m.group(0)
        return f'{m.group(1)}{m.group(2)}="{escape(tok(v))}"'
    return re.sub(r'(\s)(displayName|userId|title|assignee\w*)="([^"]*)"', attr, xml)


def _cells(xml: str, sst: list[str]) -> list[Cell]:
    cells = []
    sheet_data = xml.find("<sheetData")
    if sheet_data < 0:
        if "<c " in xml or "<c>" in xml:
            raise TokenError("a worksheet's cells could not be read")
        return cells
    for m in CELL_RE.finditer(xml, sheet_data):
        cells.append(_make_cell(m, sst))
    return cells


ROW_NUM_RE = re.compile(r'<row\b[^>]*?\sr="(\d+)"')
CELL_ROW_RE = re.compile(r'<c\b[^>]*?\sr="\$?[A-Za-z]{1,3}\$?(\d+)"')


def _in_order(xml: str, start: int) -> bool:
    """True if every row is numbered and ascending, and the cells' own row numbers never go down: then reading
    the cells in file order gives the rows exactly as sorting them would. Checked without keeping anything."""
    rows = [int(r) for r in ROW_NUM_RE.findall(xml)]
    if len(rows) != len(re.findall(r"<row\b", xml)) or any(b <= a for a, b in zip(rows, rows[1:])):
        return False
    last = 0
    for m in CELL_ROW_RE.finditer(xml, start):
        r = int(m.group(1))
        if r < last:
            return False
        last = r
    return True


def _rows(xml: str, sst: list[str]):
    """The cells of a sheet as (row number, [Cell]) in ascending row order, one row at a time, so that a large
    sheet is never held as one object per cell (about 1.6 KB each: a 5 MB workbook needed 2.4 GB).
    Excel writes rows in ascending order, each with its number; a sheet that doesn't is read whole and sorted,
    as before."""
    if not _in_order(xml, max(xml.find("<sheetData"), 0)):
        by_row = defaultdict(list)
        for c in _cells(xml, sst):
            by_row[c.row].append(c)
        for r in sorted(by_row):
            yield r, by_row[r]
        return
    sheet_data = xml.find("<sheetData")
    if sheet_data < 0:
        if "<c " in xml or "<c>" in xml:
            raise TokenError("a worksheet's cells could not be read")
        return
    cur, buf, done = None, [], set()
    for m in CELL_RE.finditer(xml, sheet_data):
        c = _make_cell(m, sst)
        if c.row != cur:
            if buf:
                yield cur, buf
                done.add(cur)
            if c.row in done:             # a cell outside its numbered row
                raise TokenError("a worksheet's cells are not in row order; open it in Excel or Google Sheets, "
                                 "save it again, and upload that copy")
            cur, buf = c.row, []
        buf.append(c)
    if buf:
        yield cur, buf



def _make_cell(m, sst: list[str]) -> Cell:
    """One <c> element (a CELL_RE match) as a Cell."""
    attrs = dict(ATTR_RE.findall(m.group(1)))
    ref = REF_RE.match(attrs.get("r", ""))
    if not ref:
        raise TokenError("a worksheet has cells without a cell reference; not supported")
    inner = m.group(2) or ""
    t = attrs.get("t", "n")
    vm = V_RE.search(inner)
    v = unescape(vm.group(1)) if vm else None
    cell = Cell(m.start(), m.end(), attrs, inner, _col_index(ref.group(1)), int(ref.group(2)),
                "other", None, has_formula="<f" in inner)
    if t == "s" and v is not None:
        cell.kind, cell.sst_index = "s", int(v)
        if cell.sst_index >= len(sst):
            raise TokenError("a cell refers to a shared text that is not in the file")
        cell.value = sst[cell.sst_index]
    elif t == "inlineStr":
        cell.kind, cell.value = "inlineStr", _si_text(inner)
    elif t == "str":
        cell.kind, cell.value = "str", v
    elif t == "n" and v not in (None, ""):
        cell.kind = "num"
        try:
            cell.value = int(v) if _INT_RE.fullmatch(v) else float(v)
        except ValueError:
            cell.kind, cell.value = "other", v
    return cell


_MONEY_RE = re.compile(r"\$?\s*-?[\d,]*\.\d+")


def _could_be_ssn(value) -> bool:
    """False only for values that cannot be an SSN or part of one: an amount with a decimal point (a
    "Social Security/Medicare" withholding column) or a number of fewer than 4 digits. Anything else in an SSN
    column, including a full SSN with or without dashes, the last 4 digits, or text, is tokenized."""
    if isinstance(value, float) and not value.is_integer():
        return False
    s = str(int(value)) if isinstance(value, float) else str(value).strip()
    if _MONEY_RE.fullmatch(s):
        return False
    digits = re.sub(r"\D", "", s)
    return not (s.lstrip("-").isdigit() and len(digits) < 4)


def _apply(xml: str, edits: list[tuple[int, int, str]]) -> str:
    """Apply non-overlapping (start, end, new) edits in one pass (linear, for sheets with 100k+ cells)."""
    if not edits:
        return xml
    out, pos = [], 0
    for start, end, new in sorted(edits):
        if start < pos:
            continue                                   # overlapping edit: the first one wins
        out += [xml[pos:start], new]
        pos = end
    out.append(xml[pos:])
    return "".join(out)


def _remove_part(parts: dict, name: str) -> None:
    parts.pop(name, None)
    ct = parts["[Content_Types].xml"].decode("utf-8")
    parts["[Content_Types].xml"] = re.sub(rf'<Override[^>]*PartName="/{re.escape(name)}"[^>]*/>', "", ct).encode()
    base = re.escape(posixpath.basename(name))
    for rel in [p for p in parts if p.endswith(".rels")]:
        r = parts[rel].decode("utf-8")
        # the relationship whose target IS this part (not "xcommentsmeta0" or "commentsmeta10"), written either as
        # <Relationship .../> or <Relationship ...></Relationship>
        new = re.sub(rf'<Relationship\b[^>]*\bTarget="(?:[^"]*/)?{base}"[^>]*?(?:/>|>\s*</Relationship>)', "", r)
        if new != r:
            parts[rel] = new.encode()


# ------------------------------------------------------------------------------------------------
# Scrubbing text in any XML part
# ------------------------------------------------------------------------------------------------
def _exact_ok(text: str) -> bool:
    """Exact-value replacement only for values that cannot be confused with numbers or codes."""
    return bool(re.search(r"[A-Za-z]", text) or re.fullmatch(r"\d{3}-\d{2}-\d{4}", text.strip()))


class _Scrub:
    def __init__(self, ctx, exact: dict, rep: Report):
        self.ctx, self.exact, self.rep = ctx, exact, rep

    def value(self, text: str, use_exact: bool) -> str:
        if not text.strip():
            return text
        if use_exact and _exact_ok(text):
            hit = self.exact.get(normalise(text))
            if hit is not None:
                return hit
        return self.ctx.scrubber.scrub(text)

    def xml(self, xml: str, use_exact: bool) -> str:
        """Scrub every text node and every attribute value of an XML part."""
        def node(m):
            raw = m.group(1)
            if NUMBER_RE.fullmatch(raw.strip()):
                return m.group(0)               # a number (cell value, index, size): never personal text
            text = unescape(raw)
            new = self.value(text, use_exact)
            if new == text:
                return m.group(0)
            self.rep.free_text += 1
            return f">{escape(new)}<"

        def attr(m):
            if m.group(1)[1:-2] in STRUCTURAL_ATTRS:
                return m.group(0)               # a cell address, style, size: checked first, it is most of them
            raw = m.group(2).strip()
            if not raw or NUMBER_RE.fullmatch(raw) or CELL_REF_RE.fullmatch(raw):
                return m.group(0)               # a number, cell address or range, never personal text
            text = unescape(m.group(2))
            new = self.value(text, use_exact)
            if new == text:
                return m.group(0)
            self.rep.free_text += 1
            return m.group(1) + escape(new) + m.group(3)
        xml = SCRUB_TEXT_NODE_RE.sub(node, xml)
        return SCRUB_ATTR_RE.sub(attr, xml)

    def joined_runs(self, xml: str, container: str) -> str:
        """Rich text split over runs (<r><t>Ana</t></r><r><t> Rivera</t></r>) is joined before scrubbing."""
        def repl(m):
            text = _si_text(m.group(2))
            new = self.value(text, True)
            if new == text:
                return m.group(0)
            self.rep.free_text += 1
            return f'{m.group(1)}<t xml:space="preserve">{escape(new)}</t></{container}>'
        return re.sub(rf"(<{container}\b[^>]*>)(.*?)</{container}>", repl, xml, flags=re.S)


# ------------------------------------------------------------------------------------------------
# Tokenize
# ------------------------------------------------------------------------------------------------
def tokenize(data: bytes, ctx) -> tuple[bytes, Report]:
    """``ctx`` is a ``files.Context`` (config, keyring, vault, scrubber)."""
    cfg, rep = ctx.config, Report()
    infos, parts = _read_zip(data)
    # Every pass below reads tags by pattern: first write every attribute as name="value" (canonical_attributes).
    for name in parts:
        if name.endswith((".xml", ".rels", ".vml")):
            try:
                xml = parts[name].decode("utf-8")
            except UnicodeDecodeError:
                continue        # left to the passes below, as before (e.g. a UTF-16 customXml part)
            fixed = canonical_attributes(xml)
            if fixed is not xml:
                parts[name] = fixed.encode("utf-8")
    sheet_parts, sst_part = _layout(parts)
    sheet_parts = list(dict.fromkeys(sheet_parts))     # a sheet listed twice is still one sheet
    for name in parts:
        if name not in sheet_parts and name != sst_part and _part_rule(name) == "refuse":
            raise TokenError(f"contains '{name}', which the tokenizer does not inspect (pivot tables, slicers, "
                             "external links and embedded objects keep their own copy of the data); save the data "
                             "alone in a new file and upload that")

    sst_xml = _text(parts, sst_part) if sst_part else ""
    si_matches = list(SI_RE.finditer(sst_xml))
    sst = [_si_text(m.group(1)) for m in si_matches]
    sst_rph = {i for i, m in enumerate(si_matches) if "<rPh" in (m.group(1) or "")}   # hidden phonetic text
    sheets = {}
    for n in sheet_parts:                     # (memory: a large sheet unpacks to hundreds of MB; one copy at a time)
        sheets[n] = _text(parts, n)
        parts[n] = b""

    # "If in doubt" (doubt.py): which cells the agent reads, decided per sheet from its name and first rows.
    tabs = _tab_order(parts, sheet_parts)
    titles = dict(zip(sheet_parts, _sheet_titles(parts, sheet_parts)))
    m = re.search(r'<workbookView\b[^>]*\sactiveTab="(\d+)"', _text(parts, _workbook_part(parts)) if _workbook_part(parts) else "")
    active = int(m.group(1)) if m and int(m.group(1)) < len(tabs) else 0
    in_doubt, rules = doubt.sheet_rules([titles[n] for n in tabs], [_head(sheets[n], sst, doubt.HEAD_ROWS) for n in tabs],
                                        active, personal=lambda h: cfg.field_for(h) is not None,
                                        strict_wage=cfg.strict_wage_tables, strict_tracker=cfg.strict_trackers,
                                        fail_closed=cfg.fail_closed_tables)
    for rule in rules:
        if rule.kind == "tracker" and rule.strict:
            rule.status_words = ctx.status_words         # tokenizer_config.yaml's and the Settings page's (#128)
    rules = dict(zip(tabs, rules))
    backs = _names_back(tabs, titles, sheets, sst, ctx) if (cfg.opaque_sheet_names or cfg.opaque_file_names) else {}

    refs_total, refs_removed = Counter(), Counter()
    exact = {}                                # normalised personal value -> replacement
    edits = defaultdict(list)                 # sheet part -> [(start, end, new_xml)]
    replaced_formula = False

    # Pass 1: find header rows and replace every cell in a personal-data column. Rows are read one at a time,
    # and nothing is kept of the other cells (pass 2 reads them again).
    for part, xml in sheets.items():
        replaced_formula |= _pass1(part, _rows(xml, sst), cfg, ctx, rep, edits, exact, refs_total, refs_removed,
                                   rules[part] if in_doubt else None, sst_rph, backs.get(part))

    scrub = _Scrub(ctx, exact, rep)

    # Pass 2: free text in the sheets (inline rich text joined first), then everything else in the sheet XML:
    # formulas, data-validation lists, headers and footers, hyperlinks. No exact-value replacement inside
    # sheets, so a trade called "Mason" is not mistaken for the last name Mason.
    for part in sheets:
        xml, done = sheets[part], {e[0] for e in edits[part]}
        sheet_data = xml.find("<sheetData")
        for m in CELL_RE.finditer(xml, max(sheet_data, 0)) if sheet_data >= 0 else ():
            if m.start() in done or "inlineStr" not in m.group(1):
                continue
            attrs = dict(ATTR_RE.findall(m.group(1)))
            value = _si_text(m.group(2)) if attrs.get("t") == "inlineStr" else ""
            if not value:
                continue
            new = ctx.scrubber.scrub(value)
            if new != value:
                rep.free_text += 1
                edits[part].append((m.start(), m.end(), _inline(attrs, new)))
        del done, xml
        xml = _apply(sheets[part], edits.pop(part, []))
        sheets[part] = None
        if in_doubt and rules[part].kind != "agent":
            xml = _opaque_outside_cells(xml, ctx, rep)
        xml = scrub.xml(xml, use_exact=False)
        parts[part] = xml.encode("utf-8")
        del xml

    # Shared strings: empty the ones only personal-data cells used; scrub the rest (rich text joined).
    if si_matches:
        si_edits = []
        for i, m in enumerate(si_matches):
            text = sst[i]
            if refs_total[i] and refs_removed[i] == refs_total[i]:
                new = ""
            else:
                new = ctx.scrubber.scrub(text)
                if new != text:
                    rep.free_text += 1
            if new != text:
                si_edits.append((m.start(), m.end(), f'<si><t xml:space="preserve">{escape(new)}</t></si>'))
        parts[sst_part] = _apply(sst_xml, si_edits).encode("utf-8")

    # Every other part, by the allow-list.
    for name in list(parts):
        if name in sheet_parts or name == sst_part:
            continue
        rule = _part_rule(name)
        if rule == "drop":
            _remove_part(parts, name)
        elif rule == "scrub":
            xml = _text(parts, name)
            if in_doubt and cfg.doubt_numbers and re.fullmatch(r"xl/charts/[^/]+\.xml", name):
                # #112: a chart keeps its own copy of the numbers and labels it draws (its caches); Excel rebuilds
                # them from the cells (now tokens where the agent doesn't read) when the workbook is opened
                xml = _CHART_CACHE_RE.sub("", xml)
            if in_doubt and re.fullmatch(r"xl/comments[^/]*\.xml|xl/comments/[^/]+\.xml", name):
                xml = _opaque_comments(xml, ctx, rep)
            elif in_doubt and re.match(r"xl/(threadedComments|persons|documenttasks)/", name):
                xml = _opaque_threaded(xml, ctx, rep)
            if re.search(r"comments[^/]*\.xml$", name):
                xml = scrub.joined_runs(xml, "text")
            parts[name] = scrub.xml(xml, use_exact=True).encode("utf-8")

    if replaced_formula and "xl/calcChain.xml" in parts:
        _remove_part(parts, "xl/calcChain.xml")        # Excel rebuilds it
    if in_doubt and cfg.fail_closed_tables:
        _opaque_objects(parts, ctx, rep)
    if cfg.opaque_sheet_names:
        _opaque_workbook_names(parts, ctx, rep)
    return _write_zip(infos, parts), rep


_A_T_RE = re.compile(r"(<a:t\b[^>]*>)([^<]*)(</a:t>)")
_TABLE_EL_RE = re.compile(r'<(?:\w+:)?(table|tableColumn)\b[^>]*>')
_NAME_ATTR_RE = re.compile(r'(\s(?:name|displayName)=")([^"]*)(")')


def _not_plain(text: str) -> bool:
    from .cva_filekinds import names as known
    s = text.strip()
    return bool(s) and not TOKEN_RE.fullmatch(s) and not NUMBER_RE.fullmatch(s) and not known.is_plain(s)


def _opaque_objects(parts: dict, ctx, rep: Report) -> None:
    """#53 (vnext.fail_closed_tables), a workbook in doubt mode: the text of charts and drawings (titles, labels, text
    boxes) and the names of tables and of their columns become text tokens unless made of known words (review of the
    #53 PR: they were only checked for learnt names). A table's new name is used in every structured reference to it
    (Table1[Hours]); a column's token is the one its header cell got (the same exact text)."""
    from . import names
    tables = {}
    for name in list(parts):
        if re.fullmatch(r"xl/(charts|drawings)/[^/]+\.xml", name):
            xml = parts[name].decode("utf-8")
            new = _A_T_RE.sub(lambda m: m.group(1) + escape(names.text_token(ctx, unescape(m.group(2)))) + m.group(3)
                              if _not_plain(unescape(m.group(2))) else m.group(0), xml)
            if new != xml:
                rep.doubt += 1
                parts[name] = new.encode("utf-8")
        elif re.fullmatch(r"xl/tables/[^/]+\.xml", name):
            xml = parts[name].decode("utf-8")

            def element(el):
                def attr(m):
                    old = unescape(m.group(2))
                    if not _not_plain(old):
                        return m.group(0)
                    new = names.text_token(ctx, old)
                    if el.group(1) == "table":
                        tables[old] = new
                    return m.group(1) + escape(new) + m.group(3)
                return _NAME_ATTR_RE.sub(attr, el.group(0))
            new = _TABLE_EL_RE.sub(element, xml)
            if new != xml:
                rep.doubt += 1
                parts[name] = new.encode("utf-8")
    if tables:
        rx = re.compile(r"(?<![\w.$'!])(" + "|".join(re.escape(escape(t)) for t in sorted(tables, key=len, reverse=True))
                        + r")(?=\[)")
        _in_formulas(parts, lambda f: rx.sub(lambda m: escape(tables[unescape(m.group(1))]), f))


def _in_formulas(parts: dict, fn) -> None:
    """Apply ``fn`` to the text of every formula (cells, defined names, charts, validations, conditional formats)."""
    for name in list(parts):
        if not name.endswith(".xml") or name.startswith("docProps/"):
            continue
        xml = parts[name].decode("utf-8")
        new = _FORMULA_EL_RE.sub(lambda m: m.group(1) + fn(m.group(3)) + m.group(4), xml)
        if new != xml:
            parts[name] = new.encode("utf-8")


_SHEET_TAG_RE = re.compile(r'(<sheet\b[^>]*?\sname=")([^"]*)(")')
# Where a formula, and so a reference to a sheet by its name, can be written (any namespace prefix)
_FORMULA_EL_RE = re.compile(r"(<((?:\w+:)?(?:f|formula|formula1|formula2|definedName|calculatedColumnFormula|"
                            r"totalsRowFormula))\b[^>]*>)(.*?)(</\2>)", re.S)
_LOCATION_RE = re.compile(r'(\slocation=")([^"]*)(")')
_TARGET_RE = re.compile(r'(\sTarget="#)([^"]*)(")')
_DEFINED_NAME_RE = re.compile(r'(<definedName\b[^>]*?\sname=")([^"]*)(")')


def _opaque_workbook_names(parts: dict, ctx, rep: Report) -> None:
    """#53 (vnext.opaque_sheet_names): each sheet's name with only known words readable (names.sheet_name), every
    reference to it renamed the same way, and the document properties made opaque (names.opaque_properties). Done
    last, on the parts as tokenized: the rules above were decided on the real names."""
    from . import names
    wb = _workbook_part(parts)
    mapping = {}
    if wb:
        xml = parts[wb].decode("utf-8")

        def sheet(m):
            old = unescape(m.group(2))
            new = names.sheet_name(old, ctx)
            if new != old:
                mapping[old] = new
            return m.group(1) + escape(new) + m.group(3)
        parts[wb] = _SHEET_TAG_RE.sub(sheet, xml).encode("utf-8")
    # In a reference a tab's name is quoted, its apostrophes doubled ('O''Brien hours'!A1): such a name's words get the
    # tokens of their doubled form there, so the way back writes them doubled again (review of the #53 PR)
    refs = {old: (names.sheet_name(old.replace("'", "''"), ctx) if "'" in old else new) for old, new in mapping.items()}
    after = _sheet_names(parts, wb)
    if len({n.lower() for n in after}) < len(after):
        raise TokenError("two of its sheets would have the same name on the AWS side; rename one of them")
    rep.doubt += len(mapping)
    # a defined name (named range) that isn't made of known words: a token, and every formula using it follows
    defined = {}
    if wb:
        xml = parts[wb].decode("utf-8")

        def dname(m):
            old = unescape(m.group(2))
            if old.startswith("_xlnm.") or not _not_plain(old):
                return m.group(0)
            defined[old] = names.text_token(ctx, old)
            return m.group(1) + escape(defined[old]) + m.group(3)
        parts[wb] = _DEFINED_NAME_RE.sub(dname, xml).encode("utf-8")
    if defined:
        rx = re.compile(r"(?<![\w.$'!\[])(" + "|".join(re.escape(escape(d)) for d in sorted(defined, key=len, reverse=True))
                        + r")(?![\w.(\[!])")
        _in_formulas(parts, lambda f: rx.sub(lambda m: escape(defined[unescape(m.group(1))]), f))
        rep.doubt += len(defined)
    for name in list(parts):
        if name.endswith(".rels") and mapping:
            # a hyperlink to a place in the workbook written as a relationship: Target="#'Sheet name'!A1"
            xml = parts[name].decode("utf-8")
            new = _TARGET_RE.sub(lambda m: m.group(1) + names.rename_refs(m.group(2), refs) + m.group(3), xml)
            if new != xml:
                parts[name] = new.encode("utf-8")
            continue
        if not name.endswith((".xml", ".vml")):
            continue
        if re.fullmatch(r"docProps/(core|app|custom)\d*\.xml", name):
            parts[name] = names.opaque_properties(parts[name].decode("utf-8"), ctx, mapping).encode("utf-8")
            continue
        if not mapping:
            continue
        xml = parts[name].decode("utf-8")
        new = _FORMULA_EL_RE.sub(lambda m: m.group(1) + names.rename_refs(m.group(3), refs) + m.group(4), xml)
        new = _LOCATION_RE.sub(lambda m: m.group(1) + names.rename_refs(m.group(2), refs) + m.group(3), new)
        if new != xml:
            parts[name] = new.encode("utf-8")


def _sheet_names(parts: dict, wb) -> list:
    return [unescape(m.group(2)) for m in _SHEET_TAG_RE.finditer(parts[wb].decode("utf-8"))] if wb else []



def _pass1(part, rows, cfg, ctx, rep, edits, exact, refs_total, refs_removed, rule=None,
           sst_rph=frozenset(), back=None) -> bool:
    """Pass 1 of tokenize for one sheet. Returns True if a replaced cell held a formula. ``rule``: the sheet's
    doubt.SheetRule when its workbook is in doubt mode (every cell the agent doesn't read becomes a text token)."""
    headers = HeaderTracker(cfg)
    replaced_formula = False
    for r, row in rows:
        for c in row:
            if c.kind == "s":
                refs_total[c.sst_index] += 1
        rewritten = set()
        todo = headers.row([(c.col, c.value) for c in row])
        for label in headers.last_labels:
            ctx.scrubber.saw_label(label)
        if todo:
            name_parts = {}
            for c in row:
                f = todo.get(c.col)
                if f is None or c.value is None or (isinstance(c.value, str) and not c.value.strip()):
                    continue
                if 't="shared"' in c.inner and "ref=" in c.inner:
                    raise TokenError("a personal-data column holds the master cell of a shared formula; "
                                     "save the file with values instead of formulas in that column")
                if f.kind == "ssn" and not _could_be_ssn(c.value):
                    continue                    # e.g. "Social Security/Medicare" withholding: 152.37
                if f.action == "token":
                    tok = ctx.keyring.token(f.prefix, f.kind, c.value)
                    ctx.vault.add(tok, f.kind, c.value)
                    edits[part].append((c.start, c.end, _inline(c.attrs, tok)))
                    exact[normalise(c.value)] = tok
                    rep.tokenized += 1
                    if f.scan_free_text == "name":
                        ctx.scrubber.add_name(str(c.value), tok, f.kind)
                    else:
                        ctx.scrubber.add_value(c.value, tok)
                    if f.part_of_name:
                        name_parts[f.part_of_name] = str(c.value).strip()
                else:
                    edits[part].append((c.start, c.end, _empty(c.attrs)))
                    exact.setdefault(normalise(c.value), "")
                    rep.dropped += 1
                rep.columns[part][f.kind] += 1
                rewritten.add(c.start)
                if c.kind == "s":
                    refs_removed[c.sst_index] += 1
            if "first" in name_parts and "last" in name_parts:
                # First and last name in separate columns: also look for the whole name in free text.
                full = f"{name_parts['last']}, {name_parts['first']}"
                tok = ctx.keyring.token("EMP", "employee_name", full)
                ctx.vault.add(tok, "employee_name", full)
                ctx.scrubber.add_name(full, tok, "employee_name")

        for c in row:
            if c.start in rewritten:
                replaced_formula |= c.has_formula
            elif back is not None:
                new = back(r, c)
                if new is not None:
                    edits[part].append((c.start, c.end, _inline(c.attrs, new)))
                    rewritten.add(c.start)
                    if c.kind == "s":
                        refs_removed[c.sst_index] += 1
        if rule is not None:
            row_vals = {c.col: c.value for c in row if c.value is not None} \
                if rule.kind in ("info", "other", "wage", "payroll") else None
            for c in row:
                if c.start in rewritten:
                    continue
                if _keep(rule, r, c.col, row_vals, c, sst_rph) and not (
                        cfg.doubt_typed_cells and rule.kind != "agent" and _typed_text(c) is not None):
                    continue                    # (#124: text in a typed cell is never kept, outside the agent's sheets)
                # a cell of a column the agent reads that a strict rule didn't keep: its number stays (#112 is about
                # the cells the agent doesn't read; a strict rule decides these)
                read = rule.kind in ("tracker", "wage", "payroll", "merged", "decisions") and rule.header_row is not None \
                    and r > rule.header_row and c.col in rule.keep_cols
                new = _doubt_cell(c, ctx, rep, read=read, strict=rule.strict)
                if new is not None:
                    edits[part].append((c.start, c.end, new))
                    if c.kind == "s":
                        refs_removed[c.sst_index] += 1
    return replaced_formula

def _names_back(tabs, titles, sheets, sst, ctx) -> dict:
    """#53 (opaque_sheet_names / opaque_file_names; review of the #53 PR): where the agent writes a name the AWS side
    only has as tokens, restored to the real words on the staff side, it is written as the AWS side names it again, so
    it never crosses readable and still matches: {sheet part: back(row, cell) -> new text or None}.
      Audit           a "sheet=<tab>;" note (agent/intent.py marks_notes): this workbook's tab, as renamed
      Review Log      column B below the header (the tab's name, agent/intent.py)
      merged table    its rates' source_file column and its sources sheet's file column (agent/wages.py): the file
                      name as the AWS side names it (names.aws_name)
    Only these places, each an exact value: no other text is looked at."""
    from . import names, screen
    cfg = ctx.config
    mapping = {t: names.sheet_name(t, ctx) for t in titles.values()} if cfg.opaque_sheet_names else {}
    mapping = {k: v for k, v in mapping.items() if k != v}
    merged = any(doubt.is_merged_table(_head(sheets[p], sst, 1)) for p in tabs)
    src_col = doubt.MERGED_TABLE_HEADER.index("source_file") + 1
    out = {}
    for p in tabs:
        t = (titles.get(p) or "").strip().lower()
        if t == "audit" and mapping:
            def audit(r, c):
                v = c.value if isinstance(c.value, str) and not c.has_formula else None
                if v and v.startswith("sheet="):
                    head, sep, rest = v.partition(";")
                    if sep and head[len("sheet="):] in mapping:
                        return f"sheet={mapping[head[len('sheet='):]]};{rest}"
                return None
            out[p] = audit
        elif t == "review log" and mapping:
            out[p] = lambda r, c: mapping.get(c.value) if r > 1 and c.col == 2 and isinstance(c.value, str) \
                and not c.has_formula else None
        elif merged and cfg.opaque_file_names and (t == "sources" or doubt.is_merged_table(_head(sheets[p], sst, 1))):
            col = 2 if t == "sources" else src_col

            def file_name(r, c, col=col):
                if r > 1 and c.col == col and isinstance(c.value, str) and not c.has_formula:
                    new = names.aws_name(screen.tokenize_name(c.value, ctx.scrubber), ctx.scrubber)
                    return new if new != c.value else None
                return None
            out[p] = file_name
    return out


# ------------------------------------------------------------------------------------------------
# Detokenize
# ------------------------------------------------------------------------------------------------
def detokenize(data: bytes, vault) -> tuple[bytes, Report]:
    rep = Report()
    infos, parts = _read_zip(data)
    try:
        sheet_parts, sst_part = _layout(parts)
    except TokenError:
        sheet_parts, sst_part = [n for n in parts if re.match(r"xl/worksheets/[^/]+\.xml$", n)], \
            ("xl/sharedStrings.xml" if "xl/sharedStrings.xml" in parts else None)
    sst = [_si_text(m.group(1)) for m in SI_RE.finditer(_text(parts, sst_part))] if sst_part else []

    def sub(text: str) -> str:
        def r(m):
            val = vault.original(m.group(0))
            if val is None:
                rep.unknown_tokens.add(m.group(0))
                return m.group(0)
            rep.restored += 1
            return escape(str(val))
        text = unglue(text, vault, escape)  # (#53: a glued sheet- or file-name token, its added space taken out)
        return TOKEN_RE.sub(r, text)

    for name in list(parts):
        if not name.endswith((".xml", ".rels", ".vml")):
            continue
        xml = orig = _text(parts, name)
        if "-" not in xml:
            continue
        if name in sheet_parts:
            edits = []
            # A cell that is exactly one token of a value that was a number becomes a number again. Only the cells
            # that can be one are read (a token in the cell itself, or a shared text that is such a token), not
            # every cell: a reviewed payroll has well over a million.
            for c in _numeric_token_cells(xml, sst, vault):
                if c.kind in ("s", "inlineStr") and isinstance(c.value, str) and TOKEN_RE.fullmatch(c.value):
                    val = vault.original(c.value)
                    if isinstance(val, (int, float)):
                        keep = "".join(f' {k}="{c.attrs[k]}"' for k in ("r", "s") if k in c.attrs)
                        edits.append((c.start, c.end, f"<c{keep}><v>{val}</v></c>"))
                        rep.restored += 1
                elif c.kind == "str" and c.has_formula and isinstance(c.value, str) and TOKEN_RE.fullmatch(c.value):
                    # a formula whose cached number became a number token (#112): its cached number again
                    val = vault.original(c.value)
                    if isinstance(val, (int, float)):
                        attrs = "".join(f' {k}="{v}"' for k, v in c.attrs.items() if k != "t")
                        inner = V_RE.sub(lambda m, v=val: f"<v>{v}</v>", c.inner, count=1)
                        edits.append((c.start, c.end, f"<c{attrs}>{inner}</c>"))
                        rep.restored += 1
            xml = _apply(xml, edits)
        new = sub(xml)
        if new != orig:            # (orig: a sheet whose only tokens were number tokens is restored too, #53 review)
            parts[name] = new.encode("utf-8")
    return _write_zip(infos, parts), rep


_V_NUM_RE = re.compile(r"<v>(\d+)</v>")


def _numeric_token_cells(xml: str, sst: list[str], vault):
    """The cells of a sheet that may hold exactly one token of a numeric value, found by searching for the
    tokens (and for the indexes of shared texts that are such tokens) rather than reading every cell."""
    sheet_data = xml.find("<sheetData")
    if sheet_data < 0:
        return []
    numeric = {}

    def is_num(tok):
        if tok not in numeric:
            numeric[tok] = isinstance(vault.original(tok), (int, float))
        return numeric[tok]
    starts = set()
    for m in TOKEN_RE.finditer(xml, sheet_data):           # a token written in the cell itself
        if is_num(m.group(0)):
            starts.add(xml.rfind("<c ", sheet_data, m.start()))
    num_sst = {i for i, t in enumerate(sst) if TOKEN_RE.fullmatch(t) and is_num(t)}
    if num_sst:                                             # a cell pointing to such a shared text
        for m in _V_NUM_RE.finditer(xml, sheet_data):
            if int(m.group(1)) in num_sst:
                starts.add(xml.rfind("<c ", sheet_data, m.start()))
    out = []
    for st in sorted(x for x in starts if x >= 0):
        m = CELL_RE.match(xml, st)
        if m:
            out.append(_make_cell(m, sst))
    return out


# ------------------------------------------------------------------------------------------------
# Macro-enabled workbooks -> ordinary workbooks
# ------------------------------------------------------------------------------------------------
_MACRO_PART_RE = re.compile(r"^(xl/vbaProject[^/]*\.bin|xl/vbaProjectSignature[^/]*\.bin|xl/vbaData\.xml|"
                            r"customUI/|customUi/)", re.I)
_MACRO_MAIN = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
_PLAIN_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"


def strip_macros(data: bytes) -> bytes:
    """An .xlsm workbook as an ordinary .xlsx: the VBA project (and any ribbon customisation that calls it) is
    removed; sheets, cells, formulas, cached values and formatting are left exactly as they were.

    Refused (TokenError) when the macros are woven into the sheets themselves, where removing them would need
    the sheets rewritten: Excel 4.0 macro sheets, dialog sheets and ActiveX controls."""
    infos, parts = _read_zip(data)
    for name in parts:
        low = name.lower()
        if low.startswith(("xl/macrosheets/", "xl/dialogsheets/", "xl/activex/")):
            raise TokenError(f"has macros built into its sheets ({name}); open it in Excel, save it as an Excel "
                             "Workbook (.xlsx) and upload that")
    removed = {n for n in parts if _MACRO_PART_RE.match(n)}
    for n in removed:
        del parts[n]
    ct = parts["[Content_Types].xml"].decode("utf-8").replace(_MACRO_MAIN, _PLAIN_MAIN)
    ct = re.sub(r'<Override[^>]*PartName="/(?:xl/vbaProject|xl/vbaProjectSignature|xl/vbaData|customUI/|customUi/)'
                r'[^"]*"[^>]*/>', "", ct, flags=re.I)
    ct = re.sub(r'<Default[^>]*ContentType="application/vnd\.ms-office\.vbaProject"[^>]*/>', "", ct)
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    gone = {posixpath.basename(n).lower() for n in removed}
    for rel in [n for n in parts if n.endswith(".rels")]:
        xml = parts[rel].decode("utf-8")

        def keep(m):
            tag = m.group(0)
            target = re.search(r'Target="([^"]*)"', tag)
            rtype = re.search(r'Type="([^"]*)"', tag)
            if rtype and re.search(r"vbaProject|ui/extensibility|customUI", rtype.group(1), re.I):
                return ""
            if target and posixpath.basename(target.group(1)).lower() in gone:
                return ""
            return tag
        new = re.sub(r"<Relationship\b[^>]*/>", keep, xml)
        if new != xml:
            parts[rel] = new.encode("utf-8")
    if "xl/workbook.xml" in parts:          # the workbook's code name only matters to VBA; harmless, but tidy
        wb = parts["xl/workbook.xml"].decode("utf-8")
        parts["xl/workbook.xml"] = re.sub(r'\s+codeName="[^"]*"', "", wb, count=1).encode("utf-8")
    return _write_zip(infos, parts)
