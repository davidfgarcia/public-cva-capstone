"""
One entry point per direction, whatever the file type:

    tokenize_file(name, data, ctx)   -> (new name, new bytes, report)   on the way to AWS
    detokenize_file(name, data, vault) -> (new bytes, report)           on the way back to Google Drive

Anything that goes wrong raises ``TokenError`` and the caller must not copy the file (fail closed).
"""
from __future__ import annotations

import csv
import io
import json
import os
import re
from dataclasses import dataclass

from . import xlsx
from .config import Config, HeaderTracker
from .cva_filekinds import payroll_export, wage_table
from .tokens import Keyring, Scrubber, TokenError, Vault, detokenize_text, normalise


@dataclass
class Context:
    config: Config
    keyring: Keyring
    vault: Vault
    scrubber: Scrubber
    # Strict trackers (#128): the Status words kept readable: tokenizer_config.yaml's, and the Settings page's lists
    # (added by the mirror at the start of a pass, mirror.Mirror._load_status_words)
    status_words: frozenset = frozenset()

    @classmethod
    def build(cls, config: Config, key: bytes, vault: Vault | None = None) -> "Context":
        kr = Keyring(key)
        vault = vault if vault is not None else Vault()
        return cls(config, kr, vault, Scrubber(kr, vault, config.patterns, config.name_kinds, config.min_name_words,
                                               config.value_kinds), config.tracker_status_words)


# ------------------------------------------------------------------------------------------------
# Text encodings (CSV exports are often Windows-1252, sometimes UTF-8 with a byte-order mark)
# ------------------------------------------------------------------------------------------------
def _decode(data: bytes) -> tuple[str, str]:
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):        # Excel's "Unicode Text" export
        return data.decode("utf-16"), "utf-16"
    if b"\x00" in data[:4096]:
        raise TokenError("looks like UTF-16/32 text without a byte-order mark; save it as UTF-8 and upload again")
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="strict"), "cp1252"


def _encode(text: str, enc: str) -> bytes:
    return text.encode(enc)


# ------------------------------------------------------------------------------------------------
# CSV
# ------------------------------------------------------------------------------------------------
def _csv_dialect(text: str):
    try:
        d = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|")
    except csv.Error:
        d = csv.excel
    newline = "\r\n" if "\r\n" in text[:20000] else "\n"
    return d, newline


def _writer(out, dialect, nl):
    """Same delimiter as the input, standard quoting (a value with a comma or quote is always quoted)."""
    return csv.writer(out, delimiter=dialect.delimiter, quotechar='"', doublequote=True,
                      quoting=csv.QUOTE_MINIMAL, lineterminator=nl)


def _csv_keep(rows, ctx: Context, rel: str | None):
    """#53 (vnext.fail_closed_tables): keep(row index, column index, row) -> does the agent read this cell of the CSV?
    None (setting off, or the agent's own L&I data: AGENT_CSV_RE): every cell handled as before.
      - WA L&I's wages CSV: its first row's cells, and the columns, that agent/wages.py parse_lni_csv reads;
      - a payroll export: the header row the reader finds (agent/payroll.py: rows 1-7, an exact "Classification" cell)
        with the reader's columns, and those columns on payroll lines below it (doubt.py's payroll rule);
      - anything else (no header the agent's readers find): nothing."""
    if not ctx.config.fail_closed_tables or (rel and AGENT_CSV_RE.search(rel)):
        return None
    if rows and wage_table.is_lni_csv_header(rows[0]):
        cols = {i for i, h in enumerate(rows[0]) if h.strip().lower() in wage_table.LNI_CSV_READ}
        return lambda r, i, row: i in cols
    for h, header in enumerate(rows[:payroll_export.HEADER_ROWS]):
        if any(payroll_export.norm_header(v) == payroll_export.HEADER_CELL for v in header):
            roles, line = payroll_export.lines_and_columns(header)
            cols, line = {c - 1: role for c, role in roles.items()}, {c - 1 for c in line}
            labels = {i: payroll_export.norm_header(header[i]) for i in cols}
            return lambda r, i, row: i in cols and r >= h and (r == h or (
                (cols[i] in payroll_export.TEXT_ROLES or payroll_export.number_text_ok(row[i])
                 or payroll_export.norm_header(row[i]) == labels[i])
                and any(c < len(row) and row[c] not in ("", None) for c in line)))
    return lambda r, i, row: False


def tokenize_csv(data: bytes, ctx: Context, rel: str | None = None) -> tuple[bytes, xlsx.Report]:
    rep = xlsx.Report()
    text, enc = _decode(data)
    dialect, nl = _csv_dialect(text)
    rows = list(csv.reader(io.StringIO(text, newline=""), dialect))
    headers = HeaderTracker(ctx.config)
    done = set()
    for r, row in enumerate(rows):
        todo = headers.row(list(enumerate(row)))
        for label in headers.last_labels:
            ctx.scrubber.saw_label(label)
        name_parts = {}
        for i, f in todo.items():
            value = row[i]
            if not value.strip():
                continue
            if f.action == "token":
                tok = ctx.keyring.token(f.prefix, f.kind, value)
                ctx.vault.add(tok, f.kind, value)
                row[i] = tok
                rep.tokenized += 1
                if f.scan_free_text == "name":
                    ctx.scrubber.add_name(value, tok, f.kind)
                else:
                    ctx.scrubber.add_value(value, tok)
                if f.part_of_name:
                    name_parts[f.part_of_name] = value.strip()
            else:
                row[i] = ""
                rep.dropped += 1
            rep.columns["csv"][f.kind] += 1
            done.add((r, i))
        if "first" in name_parts and "last" in name_parts:
            # First and last name in separate columns: also look for the whole name in free text.
            full = f"{name_parts['last']}, {name_parts['first']}"
            tok = ctx.keyring.token("EMP", "employee_name", full)
            ctx.vault.add(tok, "employee_name", full)
            ctx.scrubber.add_name(full, tok, "employee_name")
    keep = _csv_keep(rows, ctx, rel)
    for r, row in enumerate(rows):
        for i, v in enumerate(row):
            if (r, i) in done or not v:
                continue
            if keep is not None and not keep(r, i, row) and v.strip() and (
                    ctx.config.doubt_numbers or re.search(r"[^\W\d_]", v)):
                # #53 (vnext.fail_closed_tables): a cell the agent doesn't read becomes a text token, as in a workbook
                # in doubt mode (numbers and dates too with vnext.doubt_numbers)
                tok = ctx.keyring.text_token(v)
                ctx.vault.add_text(tok, v)
                row[i] = tok
                rep.doubt += 1
                continue
            new = ctx.scrubber.scrub(v)
            if new != v:
                row[i] = new
                rep.free_text += 1
    out = io.StringIO(newline="")
    w = _writer(out, dialect, nl)
    w.writerows(rows)
    body = out.getvalue()
    if not text.endswith(("\n", "\r")) and body.endswith(nl):
        body = body[: -len(nl)]
    return _encode(body, enc), rep


# The agent's own CSVs: L&I's public intents and affidavits it saved for a project (agent/intent.py, in "Files/L&I
# data/<run>/" since the current layout; "tasks/intent/lni_data/", "Tracker/lni_data/" and their archives'
# "<time>__lni_data/" folders in older ones, agent/layout.py). They come back from AWS (or are written by
# the desktop app on the staff side) and must reach the agent again: #111's hold leaves them alone. Public data, no
# worker's; tokenized as any CSV.
AGENT_CSV_RE = re.compile(r"(?i)(?:^|/)(?:files/l&i data/[^/]+|(?:tasks/intent|tracker)/lni_data"
                          r"|(?:tasks/intent|tracker)/archive/[^/]*lni_data)/(?:intents|affidavits)\.csv$")


def csv_is_read(data: bytes) -> bool:
    """#111 (vnext.hold_unknown_csv): is this CSV one the agent reads? A payroll export (Uploads' header test in its
    first rows: cva_filekinds.payroll_export.payroll_header) or WA L&I's journey-level wages CSV (its first row:
    cva_filekinds.wage_table.is_lni_csv_header). Anything else is held back."""
    try:
        text, _ = _decode(data)
        dialect, _ = _csv_dialect(text)
        rows = []
        for row in csv.reader(io.StringIO(text, newline=""), dialect):
            rows.append(row)
            if len(rows) >= max(payroll_export.SNIFF_ROWS, payroll_export.HEADER_ROWS):
                break
    except Exception:                            # noqa: BLE001 - unreadable: not one the agent reads
        return False
    if not rows:
        return False
    # (only Uploads' payroll test: the agent reads a payroll CSV only when Uploads recognised it, agent/uploads.py)
    return wage_table.is_lni_csv_header(rows[0]) or payroll_export.payroll_header(rows)


_WD_RE = re.compile(r"General Decision Number|Wage Determination|Prevailing Wage", re.I)


def _is_text_table(text: str) -> bool:
    """A tab-separated table, not text with a stray tab: not a wage determination (the agent reads those as text,
    agent/wages.py parse_dol_text), and at least two of its first lines with the same number of tabs."""
    head = text[:4000]
    if _WD_RE.search(head):
        return False
    tabs = [ln.count("\t") for ln in head.splitlines()[:10] if ln.strip()]
    return len(tabs) >= 2 and tabs[0] >= 1 and tabs.count(tabs[0]) >= 2


def _hold_unread_table(data: bytes, ctx: Context, rel: str | None, what: str = "a CSV file") -> None:
    """#111 (vnext.hold_unknown_csv): raise TokenError (held back) for a CSV, or a tab-separated text table, that the
    agent doesn't read: not a payroll export, not L&I's wages CSV, not the agent's own L&I data (AGENT_CSV_RE)."""
    if ctx.config.hold_unknown_csv and not (rel and AGENT_CSV_RE.search(rel)) and not csv_is_read(data):
        raise TokenError(f"{what} the CVA doesn't read (not a payroll export or an L&I wages CSV); if it is "
                         "needed, save it as an Excel workbook (.xlsx) and upload that")


# ------------------------------------------------------------------------------------------------
# Dispatch
# ------------------------------------------------------------------------------------------------
def tokenize_file(name: str, data: bytes, ctx: Context, google_sheet: bool = False, rel: str | None = None):
    """Returns (name for the tokenized copy, bytes, report) or raises TokenError.
    Returns (None, None, reason) for files that are deliberately not copied. ``rel``: the file's path in the staff
    folder, when known (the agent's own CSVs are recognised by it: AGENT_CSV_RE)."""
    how = "tokenize" if google_sheet else ctx.config.handling(name)
    ext = os.path.splitext(name)[1].lower()
    try:
        if how == "skip":
            return None, None, f"not copied: '{ext or 'no extension'}' files are not in the tokenizer's file rules"
        if how == "scan_pdf":            # checked and copied by the mirror (screen.py), never here
            raise TokenError("PDFs are only copied after the checks in screen.py")
        if how == "scrub_text":
            text, enc = _decode(data)
            # a tab-separated table (Excel "Text" export); #53 (fail_closed_tables): also one below a title line
            table = _is_text_table(text) or (ctx.config.fail_closed_tables and not _WD_RE.search(text[:4000]) and
                                             _is_text_table("\n".join(ln for ln in text[:4000].splitlines() if "\t" in ln)))
            first_tab = text.split("\n", 1)[0].count("\t") >= 1
            table = (first_tab and _is_text_table(text)) or (table and ctx.config.fail_closed_tables)
            if first_tab or table:              # (settings off: as before, only with a tab in the first line)
                if table:
                    _hold_unread_table(data, ctx, rel, what="a text table")      # #111: as a CSV
                out, rep = tokenize_csv(data, ctx, rel)
                return name, out, rep.summary()
            rep = xlsx.Report()
            new = ctx.scrubber.scrub(text)
            rep.free_text = int(new != text)
            return name, _encode(new, enc), rep.summary()
        if google_sheet or ext in (".xlsx", ".xlsm"):
            out, rep = xlsx.tokenize(data, ctx)
            if google_sheet and not name.lower().endswith(".xlsx"):
                name += ".xlsx"
            return name, out, rep.summary()
        if ext == ".csv":
            _hold_unread_table(data, ctx, rel)
            out, rep = tokenize_csv(data, ctx, rel)
            return name, out, rep.summary()
        raise TokenError(f"no tokenizer for '{ext}' files")
    except TokenError:
        raise
    except Exception as e:                       # anything unexpected: refuse rather than leak
        raise TokenError(f"could not be read ({type(e).__name__}: {e})") from e


def detokenize_file(name: str, data: bytes, vault: Vault):
    """Returns (bytes, report). Types we cannot look inside come back unchanged."""
    ext = os.path.splitext(name)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        out, rep = xlsx.detokenize(data, vault)
        return out, rep
    rep = xlsx.Report()
    if ext in (".csv", ".tsv"):
        text, enc = _decode(data)
        dialect, nl = _csv_dialect(text)
        rows = list(csv.reader(io.StringIO(text, newline=""), dialect))
        rows = [[detokenize_text(v, vault, rep.unknown_tokens) for v in row] for row in rows]
        out = io.StringIO(newline="")
        _writer(out, dialect, nl).writerows(rows)   # re-quoted: "Smith, John" stays one field
        body = out.getvalue()
        if not text.endswith(("\n", "\r")) and body.endswith(nl):
            body = body[: -len(nl)]
        rep.restored = int(body != text)
        return _encode(body, enc), rep
    if ext == ".json":
        text, enc = _decode(data)
        try:
            obj = json.loads(text)
        except ValueError:
            return data, rep
        def walk(o):
            if isinstance(o, str):
                return detokenize_text(o, vault, rep.unknown_tokens)
            if isinstance(o, list):
                return [walk(v) for v in o]
            if isinstance(o, dict):
                return {walk(k): walk(v) for k, v in o.items()}
            return o
        new = walk(obj)
        if new == obj:
            return data, rep
        rep.restored = 1
        return _encode(json.dumps(new, ensure_ascii=False, indent=1), enc), rep
    if ext in (".txt", ".md", ".html", ".htm", ".log", ".eml"):
        text, enc = _decode(data)
        before = text
        text = detokenize_text(text, vault, rep.unknown_tokens)
        rep.restored = int(text != before)
        return _encode(text, enc), rep
    return data, rep


__all__ = ["Context", "tokenize_file", "detokenize_file", "TokenError", "normalise"]
