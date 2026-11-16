"""Reading PDFs one page at a time.

pdfplumber keeps every page it has read in memory, about 6 MB a page: a 300-page L&I print took 1.8 GB and the
4 GB server ran out of memory. Here each page is released as soon as it has been read, and only its result is kept.
"""
from __future__ import annotations

from pathlib import Path


def pdf_pages(path: Path, what: str = "plain"):
    """Each page's words (what='words'), layout text ('text') or plain text ('plain'), one page at a time."""
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            try:
                if what == "words":
                    yield page.extract_words(keep_blank_chars=False)
                elif what == "text":
                    yield page.extract_text(layout=True) or ""
                else:
                    yield page.extract_text() or ""
            finally:
                page.close()                   # frees the page's parsed objects (pdfplumber's cache)


def pdf_text(path: Path) -> str:
    """The whole PDF's plain text, read page by page."""
    return "\n".join(pdf_pages(path, "plain"))
