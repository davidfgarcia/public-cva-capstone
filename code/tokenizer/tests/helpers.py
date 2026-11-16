"""Test helpers: hand-built .xlsx files (with shared strings, like Google Sheets and Excel write them) and an
in-memory Google Drive. Every name, number and company in the tests is invented."""
from __future__ import annotations

import dataclasses
import hashlib
import io
import itertools
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pii_tokenizer import config                       # noqa: E402
from pii_tokenizer.files import Context                # noqa: E402

KEY = bytes(range(32))
# The shipped config with every vNext setting off: the baseline most tests describe (v1 behaviour). Tests of a setting
# turn it on themselves; test_privacy_quick_wins.py checks which ones the shipped config has on.
CFG = dataclasses.replace(config.load(), vnext={k: False for k in config.VNEXT_SETTINGS})
SHIPPED = config.load()


def ctx(vault=None):
    return Context.build(CFG, KEY, vault)


def col(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _put(z, name, data):
    """One zip entry with a fixed time: the same rows always give the same bytes (a test can build the expected file
    again to compare; with the current time in each entry, that failed whenever a second ticked over)."""
    zi = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
    zi.compress_type = zipfile.ZIP_DEFLATED
    z.writestr(zi, data)


def make_xlsx(rows, extra_parts=None, formulas=None, calc_chain=False):
    """rows: lists of values (str -> shared string, int/float -> number, None -> no cell).
    formulas: {"J3": ("H3*1.5", 78.15)} -> formula cell with a cached value."""
    sst, index = [], {}
    cells_xml = []
    for r, row in enumerate(rows, start=1):
        cs = []
        for c, v in enumerate(row, start=1):
            ref = f"{col(c)}{r}"
            if v is None:
                continue
            if isinstance(v, str):
                if v not in index:
                    index[v] = len(sst)
                    sst.append(v)
                cs.append(f'<c r="{ref}" t="s"><v>{index[v]}</v></c>')
            else:
                cs.append(f'<c r="{ref}"><v>{v}</v></c>')
        for ref, (f, cached) in (formulas or {}).items():
            if int("".join(ch for ch in ref if ch.isdigit())) == r:
                t = ' t="str"' if isinstance(cached, str) else ""
                cs.append(f'<c r="{ref}"{t}><f>{esc(f)}</f><v>{esc(cached)}</v></c>')
        cells_xml.append(f'<row r="{r}">{"".join(cs)}</row>')
    sheet = ('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main"><sheetData>' + "".join(cells_xml) + "</sheetData></worksheet>")
    sst_xml = ('<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/'
               f'2006/main" count="{len(sst)}" uniqueCount="{len(sst)}">'
               + "".join(f'<si><t xml:space="preserve">{esc(s)}</t></si>' for s in sst) + "</sst>")
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/'
          'content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.'
          'relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
          'spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/'
          'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.'
          'xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
          + ('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.'
             'spreadsheetml.calcChain+xml"/>' if calc_chain else "") + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/'
            'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/'
          '2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          '<sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/'
               'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
               'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship '
               'Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
               'Target="sharedStrings.xml"/>'
               + ('<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                  'relationships/calcChain" Target="calcChain.xml"/>' if calc_chain else "") + "</Relationships>")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        _put(z, "[Content_Types].xml", ct)
        _put(z, "_rels/.rels", rels)
        _put(z, "xl/workbook.xml", wb)
        _put(z, "xl/_rels/workbook.xml.rels", wb_rels)
        _put(z, "xl/worksheets/sheet1.xml", sheet)
        _put(z, "xl/sharedStrings.xml", sst_xml)
        if calc_chain:
            _put(z, "xl/calcChain.xml", '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/'
                                           '2006/main"><c r="J3" i="1"/></calcChain>')
        for name, data in (extra_parts or {}).items():
            _put(z, name, data)
    return out.getvalue()


def all_xml(data: bytes) -> str:
    z = zipfile.ZipFile(io.BytesIO(data))
    return "".join(z.read(n).decode("utf-8") for n in z.namelist() if n.endswith((".xml", ".rels")))


PAYROLL = [
    ["Certified payroll export"],
    ["Project Name", "Contractor Name", "Week End Date", "Employee Name", "Employee ID (LCPtracker)", "SSN",
     "Gender", "Classification", "Base Hourly", "Notes"],
    ["P07", "Acme Paving", 46000, "Rivera, Ana", 10234, "123-45-6789", "F", "Laborer", 52.1,
     "Ana Rivera says hours on 6/2 were split; call 206-555-0100"],
    ["P07", "Acme Paving", 46000, "Mason, Lee", 10235, "987-65-4321", "M", "Mason", 50, "OK"],
    ["P07", "Acme Paving", 46007, "Rivera, Ana", 10234, "123-45-6789", "F", "Laborer", 52.1, None],
]


# ------------------------------------------------------------------------------------------------
class FakeDrive:
    """Enough of Google Drive for the mirror: folders, files, versions (md5 of the content), bin."""
    FOLDER = "application/vnd.google-apps.folder"

    def __init__(self):
        self.items, self._ids = {}, itertools.count(1)
        self.calls = []

    def _new(self, name, parent, mime, data=None):
        fid = f"id{next(self._ids)}"
        self.items[fid] = {"id": fid, "name": name, "mimeType": mime, "parents": [parent] if parent else [],
                           "data": data, "trashed": False}
        return fid

    def root(self, name):
        return self._new(name, None, self.FOLDER)

    def meta(self, fid):
        it = self.items[fid]
        m = {k: it[k] for k in ("id", "name", "mimeType", "parents")}
        if it["data"] is not None and not it["mimeType"].startswith("application/vnd.google-apps."):
            m["md5Checksum"] = hashlib.md5(it["data"]).hexdigest()
        m["modifiedTime"] = str(it.get("mtime", 0))
        if it["trashed"]:
            m["trashed"] = True
        return m

    # helpers for tests (playing staff or AWS)
    def put(self, parent, path, data, mime="application/octet-stream"):
        *dirs, name = path.split("/")
        for d in dirs:
            parent = self.find_child(parent, d) or self.mkdir(parent, d)
        existing = self.find_child(parent, name)
        if existing:
            self.items[existing]["data"] = data
            self.items[existing]["mtime"] = self.items[existing].get("mtime", 0) + 1
            return existing
        return self._new(name, parent, mime, data)

    def find_child(self, parent, name):
        for i, it in self.items.items():
            if parent in it["parents"] and it["name"] == name and not it["trashed"]:
                return i
        return None

    def path(self, root, path):
        fid = root
        for part in path.split("/"):
            fid = self.find_child(fid, part)
            if fid is None:
                return None
        return fid

    def read(self, root, path):
        fid = self.path(root, path)
        return None if fid is None else self.items[fid]["data"]

    def delete(self, fid):
        self.items[fid]["trashed"] = True

    # the interface the mirror uses
    def get(self, fid):
        return self.meta(fid)

    def tree(self, root):
        files, folders, todo = {}, {"": root}, [("", root)]
        while todo:
            rel_dir, fid = todo.pop()
            for i, it in list(self.items.items()):
                if fid in it["parents"] and not it["trashed"]:
                    rel = f"{rel_dir}/{it['name']}" if rel_dir else it["name"]
                    if it["mimeType"] == self.FOLDER:
                        folders[rel] = i
                        todo.append((rel, i))
                    else:
                        files[i] = {**self.meta(i), "rel": rel}
        return files, folders

    def download(self, f):
        return self.items[f["id"]]["data"]

    def create(self, parent, name, data, mime):
        self.calls.append(("create", name))
        return self.meta(self._new(name, parent, mime, data))

    def update(self, fid, data, mime, name=None):
        self.calls.append(("update", self.items[fid]["name"]))
        self.items[fid]["data"] = data
        if name:
            self.items[fid]["name"] = name
        return self.meta(fid)

    def mkdir(self, parent, name):
        return self._new(name, parent, self.FOLDER)

    def children(self, fid):
        return [self.meta(i) for i, it in self.items.items() if fid in it["parents"] and not it["trashed"]]

    def move(self, fid, new_parent, old_parent, name=None):
        it = self.items[fid]
        it["parents"] = [new_parent if p == old_parent else p for p in it["parents"]]
        if name:
            it["name"] = name
        return self.meta(fid)

    def anyone_can_open(self, fid):
        return self.items[fid].get("anyone", False)

    def rename(self, fid, name):
        self.calls.append(("rename", self.items[fid]["name"]))
        self.items[fid]["name"] = name
        return self.meta(fid)

    def trash(self, fid):
        self.calls.append(("trash", self.items[fid]["name"]))
        self.items[fid]["trashed"] = True


def make_pdf(pages, metadata=None, image_only_page=False, note=None, outline=None, xmp=None, inline_image_page=False,
             edited=False):
    """A real PDF with one page of text per item in ``pages`` (each a list of lines). ``image_only_page``: an extra
    page that holds only an image, as a scan would (``inline_image_page``: drawn inline). ``metadata``: {"/Author": ...};
    ``note``: a sticky note on page 1; ``outline``: one bookmark title; ``xmp``: XMP properties text; ``edited``: saved
    again incrementally (an earlier version stays inside)."""
    objs = []                                   # bodies; object number = index + 1

    def add(body):
        objs.append(body)
        return len(objs)

    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    page_ids, pages_id = [], None
    kids_placeholder = add(b"")                 # the Pages object, filled in below
    for lines in pages:
        text = "BT /F1 11 Tf 72 720 Td 14 TL " + " ".join(
            "(%s) Tj T*" % l.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)") for l in lines) + " ET"
        stream = add(b"<< /Length %d >>\nstream\n%s\nendstream" % (len(text), text.encode("latin-1")))
        annots = b""
        if note and not page_ids:
            n = add(b"<< /Type /Annot /Subtype /Text /Rect [72 700 92 720] /Contents (%s) >>" % note.encode("latin-1"))
            annots = b" /Annots [%d 0 R]" % n
        page_ids.append(add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 %d 0 R >> >> "
                            b"/Contents %d 0 R%s >>" % (kids_placeholder, font, stream, annots)))
    if inline_image_page:
        draw = b"q 100 0 0 100 72 600 cm BI /W 1 /H 1 /CS /G /BPC 8 ID \x80 EI Q"
        cs = add(b"<< /Length %d >>\nstream\n%s\nendstream" % (len(draw), draw))
        page_ids.append(add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] /Contents %d 0 R >>" % (kids_placeholder, cs)))
    if image_only_page:
        img = add(b"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 "
                  b"/Length 1 >>\nstream\n\x80\nendstream")
        draw = b"q 100 0 0 100 72 600 cm /Im1 Do Q"
        cs = add(b"<< /Length %d >>\nstream\n%s\nendstream" % (len(draw), draw))
        page_ids.append(add(b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im1 %d 0 R >> >> "
                            b"/Contents %d 0 R >>" % (kids_placeholder, img, cs)))
    objs[kids_placeholder - 1] = b"<< /Type /Pages /Kids [%s] /Count %d >>" % (
        b" ".join(b"%d 0 R" % i for i in page_ids), len(page_ids))
    extra = b""
    if outline:
        outlines = add(b"")
        item = add(b"<< /Title (%s) /Parent %d 0 R /Dest [%d 0 R /Fit] >>" % (outline.encode("latin-1"), outlines, page_ids[0]))
        objs[outlines - 1] = b"<< /Type /Outlines /First %d 0 R /Last %d 0 R /Count 1 >>" % (item, item)
        extra += b" /Outlines %d 0 R" % outlines
    if xmp:
        x = ('<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
             '<rdf:Description xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator><rdf:Seq><rdf:li>%s</rdf:li>'
             '</rdf:Seq></dc:creator></rdf:Description></rdf:RDF></x:xmpmeta>' % xmp).encode()
        extra += b" /Metadata %d 0 R" % add(b"<< /Type /Metadata /Subtype /XML /Length %d >>\nstream\n%s\nendstream" % (len(x), x))
    catalog = add(b"<< /Type /Catalog /Pages %d 0 R%s >>" % (kids_placeholder, extra))
    info = None
    if metadata:
        info = add(b"<< " + b" ".join(b"%s (%s)" % (k.encode(), v.encode("latin-1")) for k, v in metadata.items()) + b" >>")
    out = io.BytesIO()
    out.write(b"%PDF-1.4\n")
    offsets = []
    for n, body in enumerate(objs, start=1):
        offsets.append(out.tell())
        out.write(b"%d 0 obj\n%s\nendobj\n" % (n, body))
    xref = out.tell()
    out.write(b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1))
    for o in offsets:
        out.write(b"%010d 00000 n \n" % o)
    out.write(b"trailer\n<< /Size %d /Root %d 0 R%s >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objs) + 1, catalog, (b" /Info %d 0 R" % info) if info else b"", xref))
    if edited:                                  # an incremental save: a new trailer and end marker appended
        out.write(b"trailer\n<< /Size %d /Root %d 0 R /Prev %d >>\nstartxref\n%d\n%%%%EOF\n" % (
            len(objs) + 1, catalog, xref, xref))
    return out.getvalue()
