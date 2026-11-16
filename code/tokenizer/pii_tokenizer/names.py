"""#53: names the AWS side sees, with only known words readable (cva_filekinds/names.py says which words).

  vnext.opaque_sheet_names  a workbook's sheet names (and every reference to a sheet by its name: formulas, defined
                            names, charts, hyperlinks, the sheet list in its document properties) and its document
                            properties (title, author, last saved by, company, custom properties...): each run of
                            unknown words becomes a TXT- text token (xlsx.tokenize calls ``workbook_names``)
  vnext.opaque_file_names   file and folder names below the drive's top (a project's own folder keeps its name, as
                            before): each run of unknown words becomes a FIL- token (mirror.py calls ``aws_name``)

Every token is an ordinary vault token: the way back (detokenize, mirror._detok_name) puts the words back exactly.
"""
from __future__ import annotations

import re

from .cva_filekinds import names as N
from .tokens import TOKEN_RE

FILE_PREFIX, FILE_KIND = "FIL", "file_name"
SHEET_KIND = "sheet_name"     # TXT- tokens of words of a sheet's name (exact text, Keyring.exact_token)
MAX_SHEET_NAME = 31          # Excel's limit for a sheet's name


def text_token(ctx, text: str) -> str:
    t = ctx.keyring.text_token(text)
    ctx.vault.add_text(t, text)
    return t


def _kind(kind: str, glue: str) -> str:
    """A token's kind, with where opaque put a space next to it (tokens.unglue takes it out on the way back)."""
    return f"{kind}+{glue}" if glue else kind


def file_token(scrubber, text: str, glue: str = "") -> str:
    """A FIL- token for words of a file or folder name, kept exactly as written (Keyring.exact_token)."""
    kind = _kind(FILE_KIND, glue)
    t = scrubber.k.exact_token(FILE_PREFIX, kind, text)
    scrubber.vault.add_exact(t, kind, text)
    return t


def aws_name(part: str, scrubber) -> str:
    """One file or folder name as the AWS side sees it with vnext.opaque_file_names (``part`` has its learnt workers'
    names replaced already: screen.tokenize_name)."""
    return N.opaque(part, lambda run, glue: file_token(scrubber, run, glue))


# ------------------------------------------------------------------------------------------------- workbooks
def sheet_token(ctx, text: str, glue: str = "") -> str:
    kind = _kind(SHEET_KIND, glue)
    t = ctx.keyring.exact_token("TXT", kind, text)
    ctx.vault.add_exact(t, kind, text)
    return t


def sheet_name(title: str, ctx) -> str:
    """A sheet's name as the AWS side sees it: unknown words as text tokens; still within Excel's 31 characters (else
    the whole name is one token)."""
    new = N.opaque(title, lambda run, glue: sheet_token(ctx, run, glue))
    if len(new) > MAX_SHEET_NAME:
        new = sheet_token(ctx, title)
    return new


_Q = r"(?:'|&apos;)"
# A reference to a sheet by its quoted name ('Ana''s hours'!A1, 'Jan:Mar'!B2) or by its plain name (Data!A1)
_QUOTED_REF_RE = re.compile(_Q + r"((?:[^'&]|&(?!apos;)|''|&apos;&apos;)+?)" + _Q + r"!")
_PLAIN_REF_RE = re.compile(r"(?<![\w.'&;])([^\W\d][\w.]*(?::[^\W\d][\w.]*)?)!")


def _unq(inner: str) -> str:
    from .xlsx import unescape
    return unescape(inner.replace("&apos;&apos;", "''")).replace("''", "'")


def rename_refs(xml: str, mapping: dict) -> str:
    """``xml`` with every reference to a renamed sheet (``mapping``: old name -> its new name as written inside a
    quoted reference, apostrophes already doubled) written with the new name, quoted ('TXT-3f2a91c0d1'!A1). A 3D range
    (two names with ':') is renamed at both ends."""
    if not mapping:
        return xml
    from .xlsx import escape

    def new_ref(names: list[str]):
        if not any(n in mapping for n in names):
            return None
        joined = ":".join(mapping[n] if n in mapping else n.replace("'", "''") for n in names)
        return "'" + escape(joined) + "'!"

    def quoted(m):
        r = new_ref(_unq(m.group(1)).split(":"))
        return r if r is not None else m.group(0)

    def plain(m):
        r = new_ref(m.group(1).split(":"))
        return r if r is not None else m.group(0)
    xml = _QUOTED_REF_RE.sub(quoted, xml)
    return _PLAIN_REF_RE.sub(plain, xml)


# Text in the document properties that may stay readable: what the office programs write there themselves
PROPERTY_WORDS = frozenset("""microsoft excel macintosh compatible openpyxl google sheets worksheets worksheet named
    ranges range office libreoffice calc apple numbers wps spreadsheet spreadsheets docs charts chart""".split())
_ISO_RE = re.compile(r"-?\d{4,}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?")
# A text element's tag and text: the typed values of custom and extended properties (vt:bool, vt:i4, vt:r8,
# vt:filetime...) are left as written (a token there makes the file unreadable: review of the #53 PR); they hold a
# number, a date or true/false, never a name
_PROP_TEXT_RE = re.compile(r"<([\w:]+)\b[^>]*>([^<]+)(?=</\1>)")
_TYPED = re.compile(r"vt:(?:bool|i[1248]|ui[1248]|int|uint|r4|r8|decimal|cy|date|filetime|error|clsid)")
_PROP_NAME_RE = re.compile(r'(\sname=")([^"]*)(")')


def _plain_property(text: str) -> bool:
    s = text.strip()
    if not s or TOKEN_RE.fullmatch(s) or _ISO_RE.fullmatch(s) or re.fullmatch(r"[-+]?\d+(\.\d+)*", s):
        return True
    words = [u for k, u in N.units(s) if k not in ("sep", "tok")]
    return all(w.lower() in PROPERTY_WORDS or N.is_known_word(w) for w in words)


def opaque_properties(xml: str, ctx, sheets: dict) -> str:
    """docProps/core.xml, app.xml, custom.xml: every text and custom property name that isn't plain (``_plain_property``)
    becomes one text token; a sheet's name in the sheet list (app.xml TitlesOfParts) becomes its new name."""
    from .xlsx import escape, unescape

    def text(m):
        t = unescape(m.group(2))
        head = m.group(0)[:len(m.group(0)) - len(m.group(2))]
        if _TYPED.fullmatch(m.group(1)):
            return m.group(0)
        if t in sheets:
            return head + escape(sheets[t])
        return m.group(0) if _plain_property(t) else head + escape(text_token(ctx, t))

    def name(m):
        t = unescape(m.group(2))
        return m.group(0) if _plain_property(t) else m.group(1) + escape(text_token(ctx, t)) + m.group(3)
    return _PROP_NAME_RE.sub(name, _PROP_TEXT_RE.sub(text, xml))


def aws_rel(rel: str, scrubber, config) -> str:
    """A staff path as the AWS side names it, one level at a time as the mirror does (mirror._copy_path, _tok_part):
    learnt names as their tokens, and with vnext.opaque_file_names every other unknown word as a FIL- token; a project's
    own folder (Projects/<project>) keeps its name. For tools that copy without the mirror (the evaluation harness)."""
    from . import screen
    parts = rel.split("/")
    out = []
    for i, p in enumerate(parts):
        if i == 1 and parts[0].casefold() == "projects" and len(parts) > 2:
            out.append(p)
            continue
        p = screen.tokenize_name(p, scrubber)
        out.append(aws_name(p, scrubber) if config.opaque_file_names else p)
    return "/".join(out)
