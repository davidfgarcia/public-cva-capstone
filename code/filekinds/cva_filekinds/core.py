"""The registry's building blocks: a file kind, its columns, and the three questions both components ask of a file.

A sheet's rows are given the way the tokenizer reads them: ``[(row number, {column number: value})]``, rows and
columns counted from 1. A workbook is one such list per sheet, in tab order, with the sheet titles and the index of the
active tab (the tab Excel opens on, which the wage readers read).

Pure Python (standard library only), Python 3.9. This package is copied into both components (code/filekinds/README.md): never
import it from outside its own folder, and import within it only relatively.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Optional

# What happens to a column's cells on their way to AWS
AS_IS = "as_is"              # readable
TOKEN = "token"              # a reversible token chosen by the tokenizer's field rules (tokenizer_config.yaml)
TEXT_TOKEN = "text_token"    # text replaced by a reversible TXT-... token (doubt mode)
DROP = "drop"                # emptied
CROSSES = (AS_IS, TOKEN, TEXT_TOKEN, DROP)

# What a column holds (documentation and later validation, #65; nothing acts on it yet)
TYPES = ("text", "name", "id", "code", "money", "number", "count", "date", "flag")


def norm_header(v) -> str:
    """A header as the readers compare it: whitespace runs to one space, stripped, lower case (agent/util.py
    norm_header). The tokenizer's field rules keep their own normaliser (pii_tokenizer/config.py)."""
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


@dataclass(frozen=True)
class Column:
    role: str                    # the reader's name for it ("craft", "basic")
    header_pattern: str          # regular expression the whole normalised header must match
    type: str = "text"           # one of TYPES
    required: bool = False       # the reader can't use the file without it
    crosses: str = AS_IS         # one of CROSSES
    layouts: tuple = ()          # the layouts of the kind that have this column; () = every layout

    def __post_init__(self):
        if self.type not in TYPES:
            raise ValueError(f"unknown column type {self.type!r}")
        if self.crosses not in CROSSES:
            raise ValueError(f"unknown crossing {self.crosses!r}")

    def matches(self, header) -> bool:
        return isinstance(header, str) and re.fullmatch(self.header_pattern, norm_header(header)) is not None

    def in_layout(self, layout: Optional[str]) -> bool:
        return layout is None or not self.layouts or layout in self.layouts


@dataclass(frozen=True)
class Recognition:
    """Where a kind was found: the sheet (0-based tab index), its header row, the columns the readers use there."""
    kind: str
    sheet: int
    header_row: int
    columns: frozenset
    layout: Optional[str] = None


@dataclass(frozen=True)
class FileKind:
    kind: str
    # recognise(heads, sheet_titles, active) -> Recognition | None
    recognise: Callable
    columns: tuple = ()
    above_header: tuple = ()     # what the readers take from above the header row (a description per item)
    version: int = 1
    notes: str = ""

    def layout_columns(self, layout: Optional[str] = None) -> tuple:
        return tuple(c for c in self.columns if c.in_layout(layout))

    def headers(self, layout: Optional[str] = None) -> dict:
        """{role: header pattern} of one layout (the readers' header maps)."""
        return {c.role: c.header_pattern for c in self.layout_columns(layout)}

    def column_for(self, header, layout: Optional[str] = None) -> Optional[Column]:
        return next((c for c in self.layout_columns(layout) if c.matches(header)), None)


def recognise_workbook(heads: list, sheet_titles: list, active: int = 0, kinds=None) -> Optional[Recognition]:
    """The first kind (in ``kinds`` order, the registry's by default) whose header shape the workbook has, or None.
    Recognition by shape only: the tokenizer's decision (pii_tokenizer/doubt.py sheet_rules) adds checks of its own,
    so this must not decide what crosses to AWS until phase 3 makes the two the same."""
    if kinds is None:
        from . import KINDS
        kinds = KINDS
    for k in kinds:
        found = k.recognise(heads, sheet_titles, active)
        if found is not None:
            return found
    return None


def columns_read(kind: FileKind, header: dict, layout: Optional[str] = None) -> dict:
    """{column number: Column} of a header row ({column number: value}): every column whose header is one the kind's
    readers use (in ``layout``, or in any layout of the kind). Non-text headers are never matched. Which columns the
    readers use, not how their cells cross: every cell of such a column, whatever it says, is readable in v1 (#118)."""
    out = {}
    for c, v in header.items():
        col = kind.column_for(v, layout)
        if col is not None:
            out[c] = col
    return out


def crosses_to_aws(kind: FileKind, header, layout: Optional[str] = None) -> str:
    """How a column with this header is meant to cross to AWS in a file of this kind; a header the kind doesn't list
    fails closed: a text token (#53). Descriptive until phase 3: the tokenizer doesn't read it, and for sheets it
    takes out of doubt mode (payroll-like ones, D1, D7) unlisted columns cross as is today."""
    col = kind.column_for(header, layout)
    return col.crosses if col is not None else TEXT_TOKEN


def sheet_index(heads: list, active: int) -> int:
    """The tab the active-tab readers look at: ``active``, kept inside the workbook (doubt.sheet_rules)."""
    return min(active, len(heads) - 1) if heads else 0


__all__ = ["AS_IS", "TOKEN", "TEXT_TOKEN", "DROP", "CROSSES", "TYPES", "norm_header", "Column", "Recognition",
           "FileKind", "recognise_workbook", "columns_read", "crosses_to_aws", "sheet_index"]
