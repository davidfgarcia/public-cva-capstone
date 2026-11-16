"""Minimal, lossless edits to an existing .xlsx: change a cell's fill, add/remove notes, and add or
append to a log sheet. Every other part of the package (threaded comments, tasks, tables, data
validation, drawings, Google-Sheets extras ...) is copied through untouched, unlike a full
load/save with openpyxl which drops what it doesn't understand."""
import re
import zipfile
from copy import deepcopy
from posixpath import dirname, join, normpath

from lxml import etree
from openpyxl.utils import column_index_from_string, get_column_letter

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
T_WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
T_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
T_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
q = lambda t: f"{{{NS}}}{t}"


def _split(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return m.group(1), int(m.group(2))


VML_HEAD = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit">'
            '<o:idmap v:ext="edit" data="{idmap}"/></o:shapelayout>'
            '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
            '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype></xml>')
SHAPE = ('<v:shape id="_x0000_s{sid}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;'
         'margin-top:2pt;width:260pt;height:90pt;z-index:{z};visibility:hidden" fillcolor="#ffffe1" '
         'o:insetmode="auto" alt="compliance-agent"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
         '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div>'
         '</v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
         '<x:Anchor>{c1}, 15, {r1}, 2, {c2}, 15, {r2}, 2</x:Anchor><x:AutoFill>False</x:AutoFill>'
         '<x:Row>{row}</x:Row><x:Column>{col}</x:Column></x:ClientData></v:shape>')


class XlsxPatcher:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.order = [i.filename for i in z.infolist()]
            self.parts = {n: z.read(n) for n in self.order}
        self._xml = {}
        self._fill_xf = {}
        self._bold_xf = None

    # ------------------------------------------------------------ helpers
    def xml(self, name):
        if name not in self._xml:
            self._xml[name] = etree.fromstring(self.parts[name])
        return self._xml[name]

    def _put(self, name, data: bytes, ctype=None):
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = data
        if ctype:
            ct = self.xml("[Content_Types].xml")
            if not ct.xpath(f'//*[local-name()="Override"][@PartName="/{name}"]'):
                etree.SubElement(ct, f"{{{CTNS}}}Override", PartName=f"/{name}", ContentType=ctype)

    def _rels_name(self, part):
        return join(dirname(part), "_rels", part.split("/")[-1] + ".rels")

    def _rels(self, part):
        rn = self._rels_name(part)
        if rn not in self.parts:
            self._put(rn, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PNS}"/>'.encode())
        return self.xml(rn)

    def _add_rel(self, part, rtype, target):
        rels = self._rels(part)
        ids = {r.get("Id") for r in rels}
        n = 1
        while f"rId{n}" in ids:
            n += 1
        etree.SubElement(rels, f"{{{PNS}}}Relationship", Id=f"rId{n}", Type=rtype, Target=target)
        return f"rId{n}"

    def _rel_target(self, part, rtype):
        rn = self._rels_name(part)
        if rn not in self.parts:
            return None
        for r in self.xml(rn):
            if r.get("Type") == rtype:
                t = r.get("Target")
                # an absolute target ("/xl/comments/comment1.xml", as openpyxl and others write it) is from the root
                return t.lstrip("/") if t.startswith("/") else normpath(join(dirname(part), t))
        return None

    def sheet_part(self, sheet_name):
        wb = self.xml("xl/workbook.xml")
        for s in wb.iter(q("sheet")):
            if s.get("name") == sheet_name:
                rid = s.get(f"{{{RNS}}}id")
                for r in self.xml("xl/_rels/workbook.xml.rels"):
                    if r.get("Id") == rid:
                        t = r.get("Target")
                        return t.lstrip("/") if t.startswith("/") else normpath(join("xl", t))
        return None

    def sheet_names(self):
        return [s.get("name") for s in self.xml("xl/workbook.xml").iter(q("sheet"))]

    # ------------------------------------------------------------- cells
    def _cell(self, part, ref, create=True):
        ws = self.xml(part)
        sd = ws.find(q("sheetData"))
        col, rown = _split(ref)
        row = None
        for r in sd.iter(q("row")):
            rr = int(r.get("r"))
            if rr == rown:
                row = r
                break
            if rr > rown:
                if not create:
                    return None
                row = etree.Element(q("row"), r=str(rown))
                r.addprevious(row)
                break
        if row is None:
            if not create:
                return None
            row = etree.SubElement(sd, q("row"), r=str(rown))
        ci = column_index_from_string(col)
        for c in row.iter(q("c")):
            cc = column_index_from_string(_split(c.get("r"))[0])
            if cc == ci:
                return c
            if cc > ci:
                if not create:
                    return None
                new = etree.Element(q("c"), r=ref)
                c.addprevious(new)
                return new
        if not create:
            return None
        return etree.SubElement(row, q("c"), r=ref)

    def get_style(self, sheet, ref):
        c = self._cell(self.sheet_part(sheet), ref, create=False)
        return int(c.get("s", "0")) if c is not None else 0

    def set_style(self, sheet, ref, s):
        c = self._cell(self.sheet_part(sheet), ref)
        if int(s) == 0:
            c.attrib.pop("s", None)
        else:
            c.set("s", str(s))

    def fill_cell(self, sheet, ref, argb):
        """Give the cell a solid fill, keeping its font/border/number format. Returns (old_s, new_s)."""
        old = self.get_style(sheet, ref)
        key = (old, argb)
        if key not in self._fill_xf:
            st = self.xml("xl/styles.xml")
            fills = st.find(q("fills"))
            f = etree.SubElement(fills, q("fill"))
            pf = etree.SubElement(f, q("patternFill"), patternType="solid")
            etree.SubElement(pf, q("fgColor"), rgb=argb)
            etree.SubElement(pf, q("bgColor"), indexed="64")
            fills.set("count", str(len(fills)))
            xfs = st.find(q("cellXfs"))
            base = deepcopy(xfs[old]) if old < len(xfs) else etree.Element(q("xf"), numFmtId="0", fontId="0", borderId="0", xfId="0")
            base.set("fillId", str(len(fills) - 1))
            base.set("applyFill", "1")
            xfs.append(base)
            xfs.set("count", str(len(xfs)))
            self._fill_xf[key] = len(xfs) - 1
        self.set_style(sheet, ref, self._fill_xf[key])
        return old, self._fill_xf[key]

    def _bold(self):
        if self._bold_xf is None:
            st = self.xml("xl/styles.xml")
            fonts = st.find(q("fonts"))
            f = etree.SubElement(fonts, q("font"))
            etree.SubElement(f, q("b"))
            etree.SubElement(f, q("sz"), val="10")
            fonts.set("count", str(len(fonts)))
            xfs = st.find(q("cellXfs"))
            etree.SubElement(xfs, q("xf"), numFmtId="0", fontId=str(len(fonts) - 1), fillId="0",
                             borderId="0", xfId="0", applyFont="1")
            xfs.set("count", str(len(xfs)))
            self._bold_xf = len(xfs) - 1
        return self._bold_xf

    # ------------------------------------------------------------- notes
    def _comments(self, part, create):
        cpart = self._rel_target(part, T_COMMENTS)
        vpart = self._rel_target(part, T_VML)
        if cpart and vpart:
            return cpart, vpart
        if not create:
            return cpart, vpart
        n = 1
        while f"xl/comments{n}.xml" in self.parts or f"xl/comments/comment{n}.xml" in self.parts:
            n += 1
        cpart = cpart or f"xl/comments_agent{n}.xml"
        if cpart not in self.parts:
            self._put(cpart, f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="{NS}"><authors/><commentList/></comments>'.encode(),
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            self._add_rel(part, T_COMMENTS, "../" + cpart.split("xl/", 1)[1])
        if not vpart:
            vpart = f"xl/drawings/vmlDrawing_agent{n}.vml"
            self._put(vpart, VML_HEAD.format(idmap=100 + n).encode())
            ct = self.xml("[Content_Types].xml")
            if not ct.xpath('//*[local-name()="Default"][@Extension="vml"]'):
                etree.SubElement(ct, f"{{{CTNS}}}Default", Extension="vml",
                                 ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing")
            rid = self._add_rel(part, T_VML, "../" + vpart.split("xl/", 1)[1])
            ws = self.xml(part)
            ld = etree.Element(q("legacyDrawing"))
            ld.set(f"{{{RNS}}}id", rid)
            after = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
                     "webPublishItems", "tableParts", "extLst"]
            nxt = next((ch for ch in ws if etree.QName(ch).localname in after), None)
            nxt.addprevious(ld) if nxt is not None else ws.append(ld)
        return cpart, vpart

    def has_note(self, sheet, ref):
        cpart, _ = self._comments(self.sheet_part(sheet), create=False)
        if not cpart:
            return False
        return bool(self.xml(cpart).xpath(f'//*[local-name()="comment"][@ref="{ref}"]'))

    def add_note(self, sheet, ref, text, author):
        """Adds a note unless the cell already has one (an existing note or a threaded comment is
        never overwritten). Returns True if added."""
        part = self.sheet_part(sheet)
        if self.has_note(sheet, ref):
            return False
        cpart, vpart = self._comments(part, create=True)
        cx = self.xml(cpart)
        authors = cx.find(q("authors"))
        names = [a.text for a in authors]
        if author not in names:
            etree.SubElement(authors, q("author")).text = author
            names.append(author)
        cl = cx.find(q("commentList"))
        cm = etree.SubElement(cl, q("comment"), ref=ref, authorId=str(names.index(author)))
        t = etree.SubElement(etree.SubElement(etree.SubElement(cm, q("text")), q("r")), q("t"))
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        vml = self.parts[vpart].decode("utf-8", "replace")
        ids = [int(x) for x in re.findall(r"_x0000_s(\d+)", vml)] or [1024 + 100 * 1024]
        col, rown = _split(ref)
        ci = column_index_from_string(col) - 1
        shape = SHAPE.format(sid=max(ids) + 1, z=len(ids) + 1, c1=ci + 1, r1=max(0, rown - 2),
                             c2=ci + 5, r2=rown + 4, row=rown - 1, col=ci)
        i = vml.rfind("</xml>")
        self.parts[vpart] = (vml[:i] + shape + vml[i:]).encode()
        return True

    def remove_notes(self, sheet, refs, author):
        part = self.sheet_part(sheet)
        cpart, vpart = self._comments(part, create=False)
        if not cpart:
            return
        cx = self.xml(cpart)
        names = [a.text for a in cx.find(q("authors"))]
        if author not in names:
            return
        aid = str(names.index(author))
        vml = self.parts[vpart].decode("utf-8", "replace") if vpart else ""
        for cm in list(cx.find(q("commentList"))):
            ref = cm.get("ref")
            if cm.get("authorId") == aid and ref in refs:
                cm.getparent().remove(cm)
                col, rown = _split(ref)
                pat = (r"<v:shape\b(?:(?!</v:shape>).)*?<x:Row>%d</x:Row>\s*<x:Column>%d</x:Column>"
                       r"(?:(?!</v:shape>).)*?</v:shape>") % (rown - 1, column_index_from_string(col) - 1)
                vml = re.sub(pat, "", vml, count=1, flags=re.S)
        if vpart:
            self.parts[vpart] = vml.encode()

    # ------------------------------------------------------- read / rewrite
    def _text(self, c):
        t = c.get("t")
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter(q("t")))
        v = c.find(q("v"))
        if v is None:
            return None
        if t == "s":
            sst = self.xml("xl/sharedStrings.xml")
            si = sst.findall(q("si"))[int(v.text)]
            return "".join(x.text or "" for x in si.iter(q("t")))
        return v.text

    def cell_text(self, sheet, ref):
        part = self.sheet_part(sheet)
        c = self._cell(part, ref, create=False) if part else None
        return self._text(c) if c is not None else None

    def rewrite_column(self, sheet, col, fn):
        """Replace the text of every cell in a column with fn(text, row_number) (written as an inline string)."""
        sd = self.xml(self.sheet_part(sheet)).find(q("sheetData"))
        for row in sd.iter(q("row")):
            rn = int(row.get("r"))
            for c in row.iter(q("c")):
                if _split(c.get("r"))[0] != col:
                    continue
                old = self._text(c)
                new = fn(old, rn)
                if new == old or new is None:
                    continue
                for ch in list(c):
                    c.remove(ch)
                c.set("t", "inlineStr")
                t = etree.SubElement(etree.SubElement(c, q("is")), q("t"))
                t.text = str(new)
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # ---------------------------------------------------------- log sheet
    def append_log(self, sheet_name, header, rows, widths=None):
        part = self.sheet_part(sheet_name)
        bold = self._bold() if part is None else None      # only a new sheet's header needs it (styles.xml stays put)
        if part is None:
            n = 1
            while f"xl/worksheets/sheet{n}.xml" in self.parts:
                n += 1
            part = f"xl/worksheets/sheet{n}.xml"
            cols = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths or [], 1))
            self._put(part, (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" xmlns:r="{RNS}">'
                             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                             '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                             + (f"<cols>{cols}</cols>" if cols else "") + '<sheetData/></worksheet>').encode(),
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            rid = self._add_rel("xl/workbook.xml", T_WS, f"worksheets/sheet{n}.xml")
            wb = self.xml("xl/workbook.xml")
            sheets = wb.find(q("sheets"))
            sid = max(int(s.get("sheetId")) for s in sheets) + 1
            el = etree.SubElement(sheets, q("sheet"), name=sheet_name, sheetId=str(sid))
            el.set(f"{{{RNS}}}id", rid)
            rows = [header] + list(rows)
            start_bold = True
        else:
            start_bold = False
        sd = self.xml(part).find(q("sheetData"))
        last = max([int(r.get("r")) for r in sd.iter(q("row"))] or [0])
        for k, vals in enumerate(rows):
            rn = last + 1 + k
            row = etree.SubElement(sd, q("row"), r=str(rn))
            for j, v in enumerate(vals, 1):
                if v is None or v == "":
                    continue
                c = etree.SubElement(row, q("c"), r=f"{get_column_letter(j)}{rn}", t="inlineStr")
                if start_bold and k == 0:
                    c.set("s", str(bold))
                t = etree.SubElement(etree.SubElement(c, q("is")), q("t"))
                t.text = str(v)
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def replace_log(self, sheet_name, header, rows, widths=None):
        """Like append_log, but the sheet keeps only the header and these rows: every earlier row is deleted (the
        Review Log holds the latest check only; David, Nov 7, 2026). Only for a sheet the agent alone writes."""
        part = self.sheet_part(sheet_name)
        if part is not None:
            ws = self.xml(part)
            sd = ws.find(q("sheetData"))
            head = next((r for r in sd if r.get("r") == "1"), None)
            keep_s = next((c.get("s") for c in (head if head is not None else []) if c.get("s")), None)
            for row in list(sd):
                sd.remove(row)
            for tag in ("mergeCells", "autoFilter", "hyperlinks"):    # they would point at rows that are gone
                for el in ws.findall(q(tag)):
                    ws.remove(el)
            dim = ws.find(q("dimension"))
            if dim is not None:
                ws.remove(dim)                                       # optional; Excel works it out again
            rows = [header] + list(rows)
            self.append_log(sheet_name, header, rows, widths)
            first = sd.find(q("row"))
            # the header's existing style when it has one: a new bold style each run would grow styles.xml forever
            bold = keep_s or str(self._bold())
            for c in (first if first is not None else []):
                c.set("s", bold)                                     # the header stays bold, as when created
            return
        self.append_log(sheet_name, header, rows, widths)

    # ---------------------------------------------------------------- save
    def save(self, path):
        for name, el in self._xml.items():
            self.parts[name] = etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = str(path) + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.parts[n])
        import os
        os.replace(tmp, path)
