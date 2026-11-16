"""Loads tokenizer_config.yaml and turns it into ready-to-use objects."""
from __future__ import annotations

import fnmatch
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

DEFAULT_PATH = Path(__file__).resolve().parent.parent / "tokenizer_config.yaml"


def norm_header(h) -> str:
    """'Employee ID (LCPtracker)' -> 'employee id lcptracker'."""
    if h is None:
        return ""
    return re.sub(r"[^0-9a-z#$]+", " ", str(h).casefold()).strip()


@dataclass
class Field:
    kind: str
    action: str                       # token | drop
    prefix: str | None
    headers: list[re.Pattern]
    scan_free_text: str | None = None     # name | value | None
    part_of_name: str | None = None   # first | last

    def matches(self, header: str) -> bool:
        return any(p.fullmatch(header) for p in self.headers)


@dataclass
class Mirror:
    name: str
    raw_folder_id: str
    tokenized_folder_id: str


@dataclass
class Config:
    fields: list[Field]
    patterns: dict
    min_name_words: int
    tokenize_exts: set[str]
    scrub_exts: set[str]
    scan_pdf_exts: set[str]
    macro_exts: set[str] = field(default_factory=set)   # converted to .xlsx on the Google side
    macro_archive_suffix: str = ".uploaded"
    mirrors: list[Mirror] = field(default_factory=list)
    exclude: list[str] = field(default_factory=list)
    state_bucket: str = ""
    key_secret: str = ""
    # Change 1 (#93): {mirror, source_folder_id, projects_dir}; off while source_folder_id is empty
    staff_trackers: dict = field(default_factory=dict)
    # vNext settings (docs/vNext.md section 13): each off by default, so the tokenizer behaves as v1
    vnext: dict = field(default_factory=dict)
    # Strict trackers (#128): the tracker Status words that stay readable (with the Settings page's lists)
    tracker_status_words: frozenset = field(default_factory=frozenset)

    @property
    def strict_wage_tables(self) -> bool:
        """#118: a wage table keeps readable only what the readers use (pii_tokenizer/xlsx.py _keep)."""
        return self.vnext.get("strict_wage_tables") is True

    @property
    def strict_trackers(self) -> bool:
        """#128: a tracker's read columns keep a text readable only when it is a value of the column's type."""
        return self.vnext.get("strict_trackers") is True

    @property
    def doubt_numbers(self) -> bool:
        """#112: in doubt mode, numbers and dates in cells the agent doesn't read become tokens too."""
        return self.vnext.get("doubt_numbers") is True

    @property
    def doubt_typed_cells(self) -> bool:
        """#124: in doubt mode, a cell typed as a number, date, yes/no or error that holds text becomes a token."""
        return self.vnext.get("doubt_typed_cells") is True

    @property
    def hold_unknown_csv(self) -> bool:
        """#111: a .csv that is neither a payroll export nor an L&I wages CSV is held back."""
        return self.vnext.get("hold_unknown_csv") is True

    @property
    def recheck_spreadsheets(self) -> bool:
        """#54: copies of spreadsheets and text already on AWS are read again for names learnt since (as PDFs are)."""
        return self.vnext.get("recheck_spreadsheets") is True

    @property
    def fail_closed_tables(self) -> bool:
        """#53: payroll exports, wage tables and the agent's merged wage tables in doubt mode too (doubt.py)."""
        return self.vnext.get("fail_closed_tables") is True

    @property
    def opaque_sheet_names(self) -> bool:
        """#53: sheet names and document properties: only known words readable (names.py)."""
        return self.vnext.get("opaque_sheet_names") is True

    @property
    def opaque_file_names(self) -> bool:
        """#53: file and folder names on the AWS side: only known words readable (names.py, mirror.py)."""
        return self.vnext.get("opaque_file_names") is True

    def field_for(self, header) -> Field | None:
        """Called for every text cell (any of them may be a header label), so results are cached."""
        cache = self.__dict__.setdefault("_field_cache", {})
        hit = cache.get(header, cache)
        if hit is not cache:
            return hit
        h = norm_header(header)
        found = next((f for f in self.fields if f.matches(h)), None) if h else None
        if len(cache) > 200_000:
            cache.clear()
        cache[header] = found
        return found

    @property
    def name_kinds(self) -> set[str]:
        return {f.kind for f in self.fields if f.scan_free_text == "name"}

    @property
    def value_kinds(self) -> set[str]:
        return {f.kind for f in self.fields if f.scan_free_text == "value"}

    def excluded(self, rel_path: str) -> bool:
        # Case-insensitive ("Tmp/", "PROJECTS/_old"): Drive names keep any case, and leaving out more is the safe side.
        rel = rel_path.lstrip("/").casefold()
        for pat in self.exclude:
            pat = pat.casefold()
            # '**/x' should also match 'x' at the top level
            if fnmatch.fnmatchcase(rel, pat) or (pat.startswith("**/") and fnmatch.fnmatchcase(rel, pat[3:])):
                return True
        return False

    def handling(self, name: str) -> str:
        """tokenize | scrub_text | scan_pdf | skip, from the file extension."""
        ext = os.path.splitext(name)[1].lower()
        if ext in self.tokenize_exts:
            return "tokenize"
        if ext in self.scrub_exts:
            return "scrub_text"
        if ext in self.scan_pdf_exts:
            return "scan_pdf"
        return "skip"


class HeaderTracker:
    """Finds header rows anywhere in a sheet, so a second table lower down is covered too.

    Call ``row(cells)`` for each row in order, with (column, value) pairs. It returns {column: Field} for the
    cells of that row that must be tokenized. A text cell matching a header pattern is a header label, unless
    its column is already a personal-data column of the same kind (then it is data, e.g. a person whose name
    happens to look like a header). Mappings from earlier header rows are kept, never dropped, so a stray
    header-like value can only add columns to tokenize, never remove one."""

    def __init__(self, config: "Config"):
        self.config, self.active = config, {}
        self.last_labels: list = []             # the header labels of the last row given (Scrubber.saw_label)

    def row(self, cells) -> dict:
        labels, texts = set(), []
        for col, value in cells:
            if isinstance(value, str) and (f := self.config.field_for(value)):
                if self.active.get(col) is f:
                    continue                    # data that looks like a header: tokenized below
                self.active[col] = f
                labels.add(col)
                texts.append(value)
        # For Scrubber.saw_label, only a real header row's labels: two or more without a digit (a cell such as
        # "EMP#55555" in a column with no personal-data header also matches a field's pattern; independent
        # verification, Nov 13, 2026: taken for a label, it made the tokenizer forget that learnt ID)
        plain = [t for t in texts if not re.search(r"\d", t)]
        self.last_labels = plain if len(plain) >= 2 else []
        return {col: self.active[col] for col, _ in cells if col in self.active and col not in labels}


def _expand(value):
    if isinstance(value, str):
        out = os.path.expandvars(value)
        if re.search(r"\$\{\w+\}", out):
            return ""                         # unset variable -> empty (checked where it matters)
        return out
    if isinstance(value, list):
        return [_expand(v) for v in value]
    if isinstance(value, dict):
        return {k: _expand(v) for k, v in value.items()}
    return value


def _free_text_mode(kind, v):
    if v in (None, False):
        return None
    if v is True:
        return "name"
    if v not in ("name", "value"):
        raise ValueError(f"fields.{kind}.scan_free_text must be name or value")
    return v


def load(path: str | Path | None = None) -> Config:
    raw = yaml.safe_load(Path(path or os.environ.get("TOKENIZER_CONFIG") or DEFAULT_PATH).read_text())
    raw = _expand(raw)
    fields = []
    for kind, f in (raw.get("fields") or {}).items():
        action = f.get("action", "token")
        if action not in ("token", "drop"):
            raise ValueError(f"fields.{kind}.action must be token or drop")
        if action == "token" and not re.fullmatch(r"[A-Z]{2,5}", f.get("prefix") or ""):
            raise ValueError(f"fields.{kind}.prefix must be 2-5 capital letters")
        fields.append(Field(kind=kind, action=action, prefix=f.get("prefix"),
                            headers=[re.compile(h) for h in f.get("headers", [])],
                            scan_free_text=_free_text_mode(kind, f.get("scan_free_text")), part_of_name=f.get("part_of_name")))
    ft = raw.get("free_text") or {}
    files = raw.get("files") or {}
    if files.get("copy_as_is"):
        raise ValueError("files.copy_as_is is no longer supported: nothing goes to AWS unchecked. List PDFs under "
                         "files.scan_pdf instead (see pii_tokenizer/screen.py)")
    if any(e.lower() != ".pdf" for e in files.get("scan_pdf", [])):
        raise ValueError("files.scan_pdf can only list .pdf")
    return Config(
        fields=fields, patterns=raw.get("patterns") or {},
        min_name_words=int(ft.get("min_name_words", 2)),
        tokenize_exts={e.lower() for e in files.get("tokenize", [])},
        scrub_exts={e.lower() for e in files.get("scrub_text", [])},
        scan_pdf_exts={e.lower() for e in files.get("scan_pdf", [])},
        macro_exts={e.lower() for e in files.get("convert_macro_workbooks", [])},
        macro_archive_suffix=files.get("macro_archive_suffix", ".uploaded"),
        mirrors=[Mirror(m["name"], m.get("raw_folder_id") or "", m.get("tokenized_folder_id") or "")
                 for m in raw.get("mirrors") or []],
        exclude=raw.get("exclude") or [], state_bucket=raw.get("state_bucket") or "",
        key_secret=raw.get("key_secret") or "", staff_trackers=raw.get("staff_trackers") or {},
        vnext=_vnext(raw.get("vnext")), tracker_status_words=_status_words(raw.get("tracker_status_words")))


VNEXT_SETTINGS = {"strict_wage_tables", "strict_trackers", "doubt_numbers", "doubt_typed_cells", "hold_unknown_csv",
                  "recheck_spreadsheets", "fail_closed_tables", "opaque_sheet_names", "opaque_file_names"}
# A status word: letters, spaces and hyphens, up to 3 words (as the Settings page's lists, settings_file.WORD)
_STATUS_WORD = re.compile(r"[^\W\d_](?:[^\W\d_]| |-){0,29}")


def _status_words(v) -> frozenset:
    """tracker_status_words: a list of words or short phrases (letters, spaces, hyphens; up to 3 words), never a
    number or anything else (a status kept readable must not be able to carry an ID or a name with digits)."""
    if v is None:
        return frozenset()
    if not isinstance(v, list):
        raise ValueError("tracker_status_words must be a list of words")
    out = set()
    for w in v:
        if not isinstance(w, str) or not _STATUS_WORD.fullmatch(w.strip()) or len(w.split()) > 3:
            raise ValueError("tracker_status_words: each entry is a word or a short phrase (letters, spaces, hyphens)")
        out.add(" ".join(w.split()).lower())
    return frozenset(out)


def _vnext(v) -> dict:
    """The vNext settings (each off by default): only known names, only a literal true or false (a quoted "true" or
    a ${VAR} would be a string), so a setting can't be on by accident or silently off by a typo."""
    if v is None:
        return {}
    if not isinstance(v, dict):
        raise ValueError("vnext must be a mapping of setting: true|false")
    unknown = set(v) - VNEXT_SETTINGS
    if unknown:
        raise ValueError(f"vnext: unknown setting(s) {', '.join(sorted(map(str, unknown)))}")
    for k, val in v.items():
        if not isinstance(val, bool):
            raise ValueError(f"vnext.{k} must be true or false, not {type(val).__name__}")
    return dict(v)
