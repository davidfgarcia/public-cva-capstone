"""
Checks for the tokenizer. Run either way:

    python tests/test_tokenizer.py
    pytest tests/
"""
from __future__ import annotations

import csv
import io
import os
import zipfile

import openpyxl

from helpers import CFG, KEY, PAYROLL, FakeDrive, all_xml, ctx, make_pdf, make_xlsx

from pii_tokenizer.files import Context, detokenize_file, tokenize_file
from pii_tokenizer.mirror import Mirror
from pii_tokenizer.tokens import TOKEN_RE, Keyring, TokenError, Vault


def _rows(data):
    return [list(r) for r in openpyxl.load_workbook(io.BytesIO(data)).active.iter_rows(values_only=True)]


# --- tokens and vault ----------------------------------------------------------------------------
def test_tokens_are_deterministic_keyed_and_normalised():
    a, b = Keyring(KEY), Keyring(bytes(reversed(range(32))))
    t = a.token("EMP", "employee_name", "Rivera, Ana")
    assert TOKEN_RE.fullmatch(t)
    assert t == a.token("EMP", "employee_name", "  RIVERA,   ana ")      # same person, same token
    assert t != b.token("EMP", "employee_name", "Rivera, Ana")          # another key, another token
    assert t != a.token("EMP", "employee_id", "Rivera, Ana")            # kind is part of the input
    assert a.token("EID", "employee_id", 10234) == a.token("EID", "employee_id", 10234.0) \
        == a.token("EID", "employee_id", "10234")


def test_vault_is_encrypted_and_refuses_the_wrong_key():
    kr = Keyring(KEY)
    v = Vault()
    v.add("EMP-0123456789", "employee_name", "Rivera, Ana")
    blob = v.to_bytes(kr)
    assert b"Rivera" not in blob
    assert Vault.from_bytes(blob, kr).original("EMP-0123456789") == "Rivera, Ana"
    try:
        Vault.from_bytes(blob, Keyring(bytes(32 * [7])))
        raise AssertionError("wrong key accepted")
    except TokenError:
        pass


def test_vault_refuses_a_collision():
    v = Vault()
    v.add("EMP-0123456789", "employee_name", "A Person")
    try:
        v.add("EMP-0123456789", "employee_name", "Someone Else")
        raise AssertionError("collision accepted")
    except TokenError:
        pass


# --- xlsx ----------------------------------------------------------------------------------------
def test_xlsx_personal_columns_are_tokenized_in_place():
    c = ctx()
    name, out, detail = tokenize_file("payroll.xlsx", make_xlsx(PAYROLL), c)
    rows = _rows(out)
    assert rows[1] == PAYROLL[1]                                     # header row untouched
    hdr = rows[1]
    emp, eid, ssn, gender, cls = (hdr.index(h) for h in ("Employee Name", "Employee ID (LCPtracker)", "SSN",
                                                          "Gender", "Classification"))
    assert all(TOKEN_RE.fullmatch(r[emp]) and TOKEN_RE.fullmatch(r[eid]) and TOKEN_RE.fullmatch(r[ssn])
               for r in rows[2:])
    assert rows[2][emp] == rows[4][emp] and rows[2][eid] == rows[4][eid]   # same person, same token
    assert all(r[gender] is None for r in rows[2:])                  # demographics emptied
    assert [r[cls] for r in rows[2:]] == ["Laborer", "Mason", "Laborer"]   # "Mason" the trade survives
    assert rows[2][hdr.index("Base Hourly")] == 52.1 and rows[2][2] == 46000  # numbers untouched
    xml = all_xml(out)
    for secret in ("Rivera", "Ana", "Mason, Lee", "123-45-6789", "987-65-4321", "206-555-0100", "10234"):
        assert secret not in xml, secret                               # nowhere in the file, not even unused


def test_xlsx_free_text_names_and_patterns_are_replaced():
    c = ctx()
    _, out, _ = tokenize_file("payroll.xlsx", make_xlsx(PAYROLL), c)
    note = _rows(out)[2][9]
    assert "Ana Rivera" not in note and "206-555-0100" not in note
    assert note.startswith("EMP-") and "PHN-" in note and "hours on 6/2 were split" in note


def test_xlsx_round_trip_restores_exact_values_and_types():
    c = ctx()
    _, out, _ = tokenize_file("payroll.xlsx", make_xlsx(PAYROLL), c)
    back, rep = detokenize_file("payroll.xlsx", out, c.vault)
    rows = _rows(back)
    assert not rep.unknown_tokens
    for orig, got in zip(PAYROLL[2:], rows[2:]):
        assert got[3] == orig[3] and got[4] == orig[4] and got[5] == orig[5]   # name, numeric ID, SSN
        assert isinstance(got[4], int)
    assert rows[2][9] == PAYROLL[2][9]                                          # free text exactly restored


def test_xlsx_formulas_and_cached_values_survive():
    data = make_xlsx(PAYROLL, formulas={"K3": ("I3*1.5", 78.15), "L3": ('D3&" (lead)"', "Rivera, Ana (lead)")},
                     calc_chain=True)
    c = ctx()
    _, out, _ = tokenize_file("payroll.xlsx", data, c)
    z = zipfile.ZipFile(io.BytesIO(out))
    sheet = z.read("xl/worksheets/sheet1.xml").decode()
    assert "<f>I3*1.5</f><v>78.15</v>" in sheet                    # formula and its cached result kept
    assert "Rivera" not in sheet                                     # cached text result scrubbed
    assert "xl/calcChain.xml" in z.namelist()                        # no formula was replaced: chain kept
    vals = openpyxl.load_workbook(io.BytesIO(out), data_only=True).active
    assert vals["K3"].value == 78.15


def test_xlsx_with_pivot_cache_is_refused():
    data = make_xlsx(PAYROLL, extra_parts={"xl/pivotCache/pivotCacheRecords1.xml": "<x/>"})
    try:
        tokenize_file("payroll.xlsx", data, ctx())
        raise AssertionError("pivot cache accepted")
    except TokenError as e:
        assert "pivot" in str(e)


def test_file_without_personal_columns_is_unchanged_in_content():
    wage = [["Craft", "Classification", "Basic Rate", "Fringes"], ["Laborers", "Laborer", 45.5, 20.1]]
    _, out, detail = tokenize_file("wages.xlsx", make_xlsx(wage), ctx())
    assert _rows(out) == wage and detail == "no personal data found"


def test_unreadable_file_fails_closed():
    try:
        tokenize_file("broken.xlsx", b"not a zip", ctx())
        raise AssertionError("broken file accepted")
    except TokenError:
        pass


def test_unknown_types_are_not_copied():
    name, data, reason = tokenize_file("photo.jpg", b"...", ctx())
    assert data is None and "not copied" in reason


# --- csv -----------------------------------------------------------------------------------------
def test_csv_round_trip_with_separate_first_and_last_names():
    text = ("First Name,Last Name,Employee ID,Hours,Comment\r\n"
            "Ana,Rivera,10234,40,Ana Rivera left early\r\n"
            "Lee,Mason,10235,38,ok\r\n")
    c = ctx()
    _, out, _ = tokenize_file("p.csv", text.encode("cp1252"), c)
    rows = list(csv.reader(io.StringIO(out.decode())))
    assert rows[0] == ["First Name", "Last Name", "Employee ID", "Hours", "Comment"]
    assert all(TOKEN_RE.fullmatch(v) for r in rows[1:] for v in r[:3]) and rows[1][3] == "40"
    assert "Ana Rivera" not in rows[1][4] and rows[1][4].endswith(" left early")
    back, _ = detokenize_file("p.csv", out, c.vault)
    assert back.decode() == text


# --- mirror --------------------------------------------------------------------------------------
class World:
    """A raw folder, its tokenized copy, the mirror and saved state, as on Google Cloud."""

    def __init__(self):
        self.d = FakeDrive()
        self.raw, self.tok = self.d.root("compliance"), self.d.root("compliance_tokenized")
        self.state, self.vault_saves = {}, 0
        self.ctx = Context.build(CFG, KEY, Vault())

    def run(self):
        def save_vault():
            self.vault_saves += 1
            self.ctx.vault.dirty = False
        return Mirror(self.d, self.ctx, self.state, save_vault, lambda: None,
                      now=getattr(self, "clock", None)).run("compliance", self.raw, self.tok)


def test_mirror_round_trip():
    w = World()
    w.d.put(w.raw, "projects/P07/payroll/inbox/week1.xlsx", make_xlsx(PAYROLL))
    w.d.put(w.raw, "private_reference/WageWorkbookNew.xlsx", b"secret")
    wd = make_pdf([["General Decision Number: WA20250001", "Laborers 52.10"]])
    w.d.put(w.raw, "projects/P07/wages/new/WD.pdf", wd)
    w.d.mkdir(w.d.path(w.raw, "projects/P07"), "tracker")
    w.run()
    copy = w.d.read(w.tok, "projects/P07/payroll/inbox/week1.xlsx")
    assert copy and "Rivera" not in all_xml(copy)
    assert w.d.path(w.tok, "private_reference") is None                     # excluded: never copied
    assert w.d.read(w.tok, "projects/P07/wages/new/WD.pdf") == wd            # checked, copied as it is
    assert w.d.path(w.tok, "projects/P07/tracker") is not None             # empty folders mirrored
    assert w.vault_saves == 1

    # AWS writes a reviewed copy (with comments that mention the employee by token) into the copy.
    emp = _rows(copy)[2][3]
    result = make_xlsx([["Employee Name", "Comment"], [emp, f"{emp}: overtime paid at base rate"]])
    w.d.put(w.tok, "projects/P07/payroll/reviewed/week1__reviewed.xlsx", result)
    w.run()
    readable = _rows(w.d.read(w.raw, "projects/P07/payroll/reviewed/week1__reviewed.xlsx"))
    assert readable[1] == ["Rivera, Ana", "Rivera, Ana: overtime paid at base rate"]

    # A quiet run changes nothing.
    w.d.calls.clear()
    w.run()
    assert w.d.calls == []


def test_mirror_follows_changes_on_each_side():
    w = World()
    raw_id = w.d.put(w.raw, "inbox/week1.xlsx", make_xlsx(PAYROLL))
    w.run()
    # Staff replace the file: the copy is updated.
    w.d.put(w.raw, "inbox/week1.xlsx", make_xlsx(PAYROLL[:3]))
    w.run()
    assert len(_rows(w.d.read(w.tok, "inbox/week1.xlsx"))) == 3
    # AWS moves it to the archive (rclone server-side move): the raw file follows.
    tok_id = w.d.path(w.tok, "inbox/week1.xlsx")
    w.d.move(tok_id, w.d.put(w.tok, "archive/.keep", b"") and w.d.path(w.tok, "archive"), w.d.path(w.tok, "inbox"))
    w.run()
    assert w.d.path(w.raw, "archive/week1.xlsx") == raw_id
    # AWS deletes it (a day later): the raw file is kept, under _sync_backup.
    for p in w.state["mirrors"]["compliance"]["pairs"]:
        p["copied_utc"] = "2000-01-01-00-00-00"
    w.d.delete(tok_id)
    w.run()
    assert w.d.path(w.raw, "archive/week1.xlsx") is None
    assert w.d.items[raw_id]["parents"][0] != w.d.path(w.raw, "archive")
    assert not w.d.items[raw_id]["trashed"]


def test_mirror_conflict_keeps_both_versions():
    w = World()
    hdr = ["Contractor", "Status", "Intent ID"]                       # a tracker: Status is a column the agent reads
    w.d.put(w.raw, "tracker.xlsx", make_xlsx([hdr, ["Acme Paving", "open", 1001]]))
    w.run()
    w.d.put(w.raw, "tracker.xlsx", make_xlsx([hdr, ["Acme Paving", "closed", 1001]]))
    w.d.put(w.tok, "tracker.xlsx", make_xlsx([hdr, ["Acme Paving", "reviewed", 1001]]))
    rep = w.run()
    assert rep.conflicts
    names = sorted(it["name"] for it in w.d.items.values()
                   if w.raw in it["parents"] and not it["trashed"])
    assert any("(AWS copy " in n for n in names)
    assert _rows(w.d.read(w.tok, "tracker.xlsx"))[1][1] == "closed"       # staff's version wins
    w.d.calls.clear()
    w.run()
    assert w.d.calls == []                                                 # the AWS copy is not sent to AWS


def test_mirror_refuses_mass_deletion_and_non_empty_start():
    w = World()
    for i in range(8):
        w.d.put(w.raw, f"f{i}.csv", b"Hours\n1\n")
    w.run()
    for it in w.d.items.values():
        if w.tok in it["parents"]:
            it["trashed"] = True                       # e.g. the tokenized folder ID was changed by mistake
    try:
        w.run()
        raise AssertionError("mass deletion accepted")
    except RuntimeError as e:
        assert "nothing was changed" in str(e)

    w2 = World()
    w2.d.put(w2.tok, "leftover.csv", b"x")
    try:
        w2.run()
        raise AssertionError("non-empty tokenized folder accepted on first run")
    except RuntimeError:
        pass


def test_mirror_does_not_copy_a_file_it_cannot_tokenize_and_removes_the_old_copy():
    w = World()
    w.d.put(w.raw, "inbox/week1.xlsx", make_xlsx(PAYROLL))
    w.run()
    w.d.put(w.raw, "inbox/week1.xlsx", make_xlsx(PAYROLL, extra_parts={"xl/pivotCache/pivotCacheRecords1.xml": "<x/>"}))
    rep = w.run()
    assert rep.refused and w.d.read(w.tok, "inbox/week1.xlsx") is None
    w.d.calls.clear()
    w.run()
    assert w.d.calls == []                             # not retried until the file changes


# --- cases found in review -----------------------------------------------------------------------
import re  # noqa: E402

NEEDLES = ("Rivera", "Mason, Lee", "123-45-6789", "10234")


def _rezip(data, fn):
    zin = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for i in zin.infolist():
            r = fn(i.filename, zin.read(i.filename))
            if r:
                z.writestr(r[0], r[1])
    return out.getvalue()


def _no_leak(out):
    xml = all_xml(out)
    leaked = [n for n in NEEDLES if n in xml]
    assert not leaked, leaked


def _refused(name, data):
    try:
        tokenize_file(name, data, ctx())
    except TokenError as e:
        return str(e)
    raise AssertionError(f"{name} was accepted")


def test_namespace_prefixed_xlsx_is_refused():
    def pref(n, d):
        if n == "xl/worksheets/sheet1.xml":
            d = re.sub(rb"<(/?)(worksheet|sheetData|row|c|v)\b", rb"<\1x:\2", d)
        return n, d
    assert "namespace prefixes" in _refused("p.xlsx", _rezip(make_xlsx(PAYROLL), pref))


def test_shared_strings_found_through_relationships():
    def mv(n, d):
        if n == "xl/sharedStrings.xml":
            return "xl/strings.xml", d
        return n, d.replace(b"sharedStrings.xml", b"strings.xml") if n.endswith((".rels", "].xml")) else d
    _, out, _ = tokenize_file("p.xlsx", _rezip(make_xlsx(PAYROLL), mv), ctx())
    _no_leak(out)


def test_text_outside_cells_is_scrubbed():
    def extra(n, d):
        if n == "xl/worksheets/sheet1.xml":
            d = d.replace(b"</sheetData>", b'</sheetData><dataValidations count="1"><dataValidation sqref="D3">'
                          b'<formula1>"Ana Rivera,Lee Mason"</formula1></dataValidation></dataValidations>'
                          b'<headerFooter><oddHeader>&amp;LPayroll of Rivera, Ana 10234</oddHeader></headerFooter>')
        if n == "xl/workbook.xml":
            d = d.replace(b'name="Data"', b'name="Rivera, Ana"')
        return n, d
    data = _rezip(make_xlsx(PAYROLL, extra_parts={
        "xl/drawings/drawing1.xml": '<wsDr><a:t>Rivera, Ana SSN 123-45-6789</a:t></wsDr>',
        "docProps/core.xml": "<cp><title>Rivera, Ana 10234</title></cp>",
        "xl/comments1.xml": "<comments><commentList><comment ref='J3'><text><r><t>Ana</t></r><r><t> Rivera "
                            "called</t></r></text></comment></commentList></comments>"}), extra)
    _, out, _ = tokenize_file("p.xlsx", data, ctx())
    _no_leak(out)


def test_parts_holding_hidden_copies_are_refused():
    for part in ("xl/pivotCache/pivotCacheDefinition1.xml", "xl/slicerCaches/slicerCache1.xml",
                 "xl/externalLinks/externalLink1.xml", "xl/embeddings/oleObject1.bin", "xl/unknown.xml"):
        assert "does not inspect" in _refused("p.xlsx", make_xlsx(PAYROLL, extra_parts={part: "<x/>"})), part


def _google_saved(rows):
    """A workbook the way Google Sheets writes one that has comments: each comments part has a binary
    xl/commentsmeta<N> next to it (relationship type workbookmetadata, content type application/binary) that
    repeats the comment text and the author. Names and numbers are invented."""
    comment = ("<comments><commentList><comment ref='J3'><text><r><t>Ana</t></r><r><t> Rivera "
               "called</t></r></text></comment></commentList></comments>")
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships"><Relationship Id="rId1" Type="http://customschemas.google.com/relationships/'
            'workbookmetadata" Target="commentsmeta0"/></Relationships>')
    meta = b"\n\x03" + b"1.0" + b"!1" + "Ana Rivera called, SSN 123-45-6789, emp 10234".encode() + b"\x00\xff\x10"
    data = make_xlsx(rows, extra_parts={"xl/comments1.xml": comment, "xl/_rels/comments1.xml.rels": rels,
                                        "xl/commentsmeta0": meta})
    return _rezip(data, lambda n, d: (n, d.replace(
        b"</Types>", b'<Override ContentType="application/binary" PartName="/xl/commentsmeta0"/></Types>'))
        if n == "[Content_Types].xml" else (n, d))


def test_a_workbook_saved_by_google_sheets_with_comments_is_tokenized_not_refused():
    """Nov 5 2026: a tracker saved through Google Sheets was held back ("does not inspect xl/commentsmeta1") and
    the project was reported as having no tracker."""
    out = tokenize_file("tracker.xlsx", _google_saved(PAYROLL), ctx())[1]
    z = zipfile.ZipFile(io.BytesIO(out))
    assert not [n for n in z.namelist() if "commentsmeta" in n]
    for n in z.namelist():                                     # nothing refers to the removed part either
        if n.endswith((".xml", ".rels")):
            assert b"commentsmeta" not in z.read(n), n
    assert "xl/comments1.xml" in z.namelist()                  # the comments stay, tokenized as before
    assert all(needle.encode() not in z.read(n) for n in z.namelist() for needle in NEEDLES), "a name was kept"
    assert zipfile.ZipFile(io.BytesIO(out)).testzip() is None


def test_only_the_known_google_part_is_dropped_other_unknown_parts_are_still_refused():
    for part in ("xl/commentsmeta", "xl/commentsmeta-1", "xl/commentsmetaX", "xl/commentsmeta0.xml.bak",
                 "xl/googlesomething1"):
        assert "does not inspect" in _refused("p.xlsx", make_xlsx(PAYROLL, extra_parts={part: b"\x00\x01"})), part


def test_removing_a_part_removes_only_the_relationship_to_that_exact_part():
    from pii_tokenizer.xlsx import _remove_part
    rels = ('<Relationships><Relationship Id="a" Type="t" Target="commentsmeta1"/>'
            '<Relationship Id="b" Type="t" Target="commentsmeta10"/>'
            '<Relationship Id="c" Type="t" Target="../xcommentsmeta1"/>'
            '<Relationship Id="d" Type="t" Target="/xl/commentsmeta1"></Relationship></Relationships>')
    parts = {"[Content_Types].xml": b"<Types/>", "xl/_rels/c.xml.rels": rels.encode(), "xl/commentsmeta1": b"x"}
    _remove_part(parts, "xl/commentsmeta1")
    left = parts["xl/_rels/c.xml.rels"].decode()
    assert 'Id="a"' not in left and 'Id="d"' not in left
    assert 'Id="b"' in left and 'Id="c"' in left and "commentsmeta1" in left
    assert "xl/commentsmeta1" not in parts


def test_header_variants_and_a_second_table_lower_down():
    for h in ("Name", "Full Name", "Emp Name", "Employee's Name", "Employee Name (Last, First)", "Worker Name"):
        rows = [list(r) for r in PAYROLL]
        rows[1][3] = h
        _, out, _ = tokenize_file("p.xlsx", make_xlsx(rows), ctx())
        assert "Rivera" not in all_xml(out), h
    rows = [list(r) for r in PAYROLL] + [[None]] * 40 + [["Week", "Employee Name"], [46014, "Mason, Lee"]]
    _, out, _ = tokenize_file("p.xlsx", make_xlsx(rows), ctx())
    _no_leak(out)


def test_middle_initials_and_ids_in_free_text():
    rows = [list(r) for r in PAYROLL]
    rows[2][9] = "Ana B. Rivera (emp 10234, SSN 123456789) left early"
    _, out, _ = tokenize_file("p.xlsx", make_xlsx(rows), ctx())
    note = _rows(out)[2][9]
    assert "Rivera" not in note and "10234" not in note and "123456789" not in note, note
    assert _rows(out)[2][8] == 52.1                                # numbers that are not IDs left alone


def test_utf16_text_exports_and_json():
    text = "Employee Name\tEmployee ID\tHours\r\nRivera, Ana\t10234\t40\r\n"
    for name in ("p.txt", "p.csv"):
        _, out, _ = tokenize_file(name, b"\xff\xfe" + text.encode("utf-16-le"), ctx())
        decoded = out.decode("utf-16")
        assert "Rivera" not in decoded and "10234" not in decoded and "\t40" in decoded
    assert "UTF-16" in _refused("p.csv", text.encode("utf-16-le"))   # no byte-order mark
    assert tokenize_file("p.json", b'{"last": "Rivera"}', ctx())[1] is None


def test_csv_restore_keeps_columns_when_names_have_commas():
    c = ctx()
    _, out, _ = tokenize_file("p.csv", b'Employee Name,Hours\n"Rivera, Ana",40\n', c)
    result = out.decode().replace("Hours", "Hours,Note").replace(",40", ',40,"ok"')
    back, _ = detokenize_file("r.csv", result.encode(), c.vault)
    assert list(csv.reader(io.StringIO(back.decode()))) == [["Employee Name", "Hours", "Note"],
                                                             ["Rivera, Ana", "40", "ok"]]


def test_rclone_backup_dir_replacement_is_a_change_not_a_move():
    w = World()
    w.d.put(w.raw, "tracker.xlsx", make_xlsx([["Contractor", "Status"], ["Acme Paving", "open"]]))
    w.run()
    # rclone --backup-dir: old copy moved aside, new copy uploaded at the same path
    old = w.d.path(w.tok, "tracker.xlsx")
    w.d.move(old, w.d.put(w.tok, "_sync_backup/x/.keep", b"") and w.d.path(w.tok, "_sync_backup/x"), w.tok)
    w.d.put(w.tok, "tracker.xlsx", make_xlsx([["Contractor", "Status"], ["Acme Paving", "reviewed"]]))
    w.run()
    assert _rows(w.d.read(w.raw, "tracker.xlsx"))[1][1] == "reviewed"     # updated in place, not hidden away
    assert w.d.path(w.raw, "_sync_backup") is None
    w.d.calls.clear()
    w.run()
    assert w.d.calls == []


def test_a_failing_download_does_not_stop_the_run_or_duplicate_results():
    w = World()
    w.d.put(w.raw, "big.xlsx", make_xlsx(PAYROLL))
    real = w.d.download
    w.d.download = lambda f: (_ for _ in ()).throw(RuntimeError("exportSizeLimitExceeded")) \
        if f["name"] == "big.xlsx" else real(f)
    rep = w.run()
    assert rep.warnings and w.d.read(w.tok, "big.xlsx") is None
    w.d.download = real
    w.run()
    assert w.d.read(w.tok, "big.xlsx") is not None                          # retried and copied



# --- runs started by AWS ---------------------------------------------------------------------------
def test_vault_saves_from_two_jobs_are_merged():
    a, b = Vault(), Vault()
    a.add("EMP-0000000001", "employee_name", "A Person")
    b.add("EMP-0000000002", "employee_name", "B Person")
    a.merge(b)
    assert a.original("EMP-0000000002") == "B Person" and a.dirty
    b.add("EMP-0000000001", "employee_name", "Someone Else")
    try:
        a.merge(b)
        raise AssertionError("collision accepted")
    except TokenError:
        pass


def test_a_run_waits_for_the_lock_instead_of_skipping():
    import main
    from pii_tokenizer.gcp import GoogleError

    class Bucket:
        def __init__(self):
            self.objs, self.gen = {}, 0

        def put(self, name, data, if_generation=None):
            cur = self.objs.get(name)
            if if_generation is not None and (cur[1] if cur else 0) != if_generation:
                raise GoogleError("precondition", 412)
            self.gen += 1
            self.objs[name] = (data, self.gen)
            return self.gen

        def get(self, name):
            return self.objs.get(name, (None, 0))

        def delete(self, name):
            self.objs.pop(name, None)

    b = Bucket()
    assert main._take_lock(b, "locks/x.lock")
    waits = []

    def sleep(s):                       # the other run finishes while we wait
        waits.append(s)
        b.delete("locks/x.lock")
    assert main._take_lock(b, "locks/x.lock", sleep=sleep) and waits
    assert main._take_lock(b, "locks/y.lock")                 # another folder's lock is independent



# --- macro-enabled workbooks ------------------------------------------------------------------------
def _xlsm(rows=PAYROLL, extra=None):
    def macro(n, d):
        if n == "[Content_Types].xml":
            d = d.replace(b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                          b"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            d = d.replace(b"</Types>", b'<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                                       b"</Types>")
        if n == "xl/_rels/workbook.xml.rels":
            d = d.replace(b"</Relationships>", b'<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/'
                                               b'relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
        if n == "xl/workbook.xml":
            d = d.replace(b"<sheets>", b'<workbookPr codeName="ThisWorkbook"/><sheets>')
        return n, d
    parts = {"xl/vbaProject.bin": "MACRO-CODE"}
    parts.update(extra or {})
    return _rezip(make_xlsx(rows, extra_parts=parts), macro)


def test_macros_are_removed_and_nothing_else_changes():
    from pii_tokenizer.xlsx import strip_macros
    data = _xlsm()
    out = strip_macros(data)
    z = zipfile.ZipFile(io.BytesIO(out))
    assert "xl/vbaProject.bin" not in z.namelist()
    xml = all_xml(out)
    assert "macroEnabled" not in xml and "vbaProject" not in xml and "codeName" not in xml
    assert _rows(out) == _rows(data)                           # every cell as it was
    assert "has macros built into its sheets" in _refused_strip(_xlsm(extra={"xl/activeX/activeX1.xml": "<x/>"}))


def _refused_strip(data):
    from pii_tokenizer.xlsx import strip_macros
    try:
        strip_macros(data)
    except TokenError as e:
        return str(e)
    raise AssertionError("accepted")


def test_mirror_converts_macro_workbooks_and_parks_the_original():
    w = World()
    w.d.put(w.raw, "Projects/P1/Uploads/week.xlsm", _xlsm())
    w.d.put(w.raw, "Projects/P1/Tracker/Tracker.xlsm", _xlsm([["Contractor", "Status"], ["Acme Paving", "open"]]))
    w.d.put(w.raw, "Projects/P1/Uploads/bad.xlsm", _xlsm(extra={"xl/activeX/activeX1.xml": "<x/>"}))
    rep = w.run()
    assert len(rep.converted) == 2 and rep.refused
    # staff side: .xlsx beside, original archived as .uploaded
    assert w.d.read(w.raw, "Projects/P1/Uploads/week.xlsx")
    assert w.d.read(w.raw, "Projects/P1/Uploads/week.xlsm") is None
    arch = [it["name"] for it in w.d.items.values() if it["name"].endswith(".xlsm.uploaded") and not it["trashed"]]
    assert sorted(arch) == ["Tracker.xlsm.uploaded", "week.xlsm.uploaded"]
    archives = [i for i, it in w.d.items.items() if it["name"] == "Archived" and w.raw in _ancestors(w.d, i)]
    assert len(archives) == 2                      # Projects/P1/Files/Archived/<ts>/ and Projects/P1/Tracker/Archived/<ts>/
    assert w.d.path(w.raw, "Projects/P1/Files/Archived") and w.d.path(w.raw, "Projects/P1/Tracker/Archived")
    # AWS side: only macro-free, tokenized .xlsx; no .uploaded, no macro workbook at all
    tok_names = [it["name"] for i, it in w.d.items.items() if not it["trashed"] and it["data"] is not None
                 and w.tok in _ancestors(w.d, i)]
    assert "week.xlsx" in tok_names and "Tracker.xlsx" in tok_names
    assert not any(n.endswith((".xlsm", ".uploaded")) for n in tok_names)
    assert "vbaProject" not in all_xml(w.d.read(w.tok, "Projects/P1/Uploads/week.xlsx"))
    # a quiet second run changes nothing
    w.d.calls.clear()
    w.run()
    assert w.d.calls == []


def _ancestors(d, fid):
    out, cur = [], fid
    while d.items[cur]["parents"]:
        cur = d.items[cur]["parents"][0]
        out.append(cur)
    return out



def test_google_sheets_metadata_part_is_accepted():
    """Workbooks saved by Google Sheets carry xl/metadata (binary settings, no text); they must not be refused."""
    from pii_tokenizer.xlsx import _part_rule
    assert _part_rule("xl/metadata") == "keep"
    assert _part_rule("xl/metadata.bin") == "refuse"


def test_parts_seen_in_real_files_are_handled():
    """Parts found in the real project files: LCPtracker exports (docProps/core0.xml), openpyxl notes
    (xl/comments/comment1.xml, as in every reviewed payroll) and Excel assigned tasks in the trackers."""
    from pii_tokenizer.xlsx import _part_rule
    for name in ("docProps/core0.xml", "xl/comments/comment1.xml", "xl/documenttasks/documenttask1.xml"):
        assert _part_rule(name) == "scrub", name


def test_short_numeric_ids_do_not_replace_years_in_free_text():
    """Seen in the real files: an employee ID of 2025 turned 'WA L&I 2025' and '11/10/2025' into tokens."""
    from pii_tokenizer.tokens import Keyring, Scrubber, Vault
    sc = Scrubber(Keyring(b"k" * 32), Vault(), {}, set())
    sc.add_value(2025, "EID-aaaaaaaaaa")
    sc.add_value("4417", "SSN-bbbbbbbbbb")
    sc.add_value("1234567", "EID-cccccccccc")
    sc.add_value("AB1234", "EID-dddddddddd")
    out = sc.scrub("WA L&I 2025, work before 11/10/2025; last 4 4417; ID 1234567; badge AB1234")
    assert "WA L&I 2025" in out and "11/10/2025" in out and "4417" in out
    assert "EID-cccccccccc" in out and "EID-dddddddddd" in out


def test_ids_that_look_like_cell_addresses_or_numbers_do_not_break_the_sheet():
    """Seen in the real files: an employee ID 'AH2026' replaced the address of cell AH2026 (file unreadable)."""
    from pii_tokenizer.files import Context
    from pii_tokenizer.tokens import Vault
    pad = [f"Col {k}" for k in range(31)]                  # so that column AH (34th) exists
    rows = [["Employee Name", "Employee ID", "Hours"] + pad + ["ST Hours"], ["Rivera, Ana", "AH2026", 40] + [0] * 31 + [8]]
    rows += [["Lopez, Ben", f"Z{i}", i] + [0] * 31 + [8] for i in range(3, 2030)]
    ctx = Context.build(CFG, KEY, Vault())
    name, out, _ = tokenize_file("p.xlsx", make_xlsx(rows), ctx)
    wb = openpyxl.load_workbook(io.BytesIO(out))          # must open
    ws = wb.active
    assert ws["C2026"].value == 2026 and ws["AH2026"].value == 0 and ws["AI2026"].value == 8 and str(ws["B2"].value).startswith("EID-")


def test_names_in_a_tracker_are_scrubbed_even_on_the_first_sync():
    """Seen in the real files: Tracker/ sorts before Upload/, so the tracker was copied before the payroll had
    taught the tokenizer the workers' names, and a list of workers in the tracker reached AWS."""
    from pii_tokenizer.tokens import Vault
    w = World()
    w.d.put(w.raw, "projects/P1/Tracker/Tracker.xlsx", make_xlsx([["Contractor", "Notes"],
                                                                 ["Acme Paving", "foreman Ana Rivera on site"]]))
    w.d.put(w.raw, "projects/P1/Upload/week.xlsx", make_xlsx(PAYROLL))
    w.run()
    assert "Ana Rivera" not in all_xml(w.d.read(w.tok, "projects/P1/Tracker/Tracker.xlsx"))


def test_social_security_withholding_amounts_stay_but_real_ssns_are_tokenized():
    """Seen in the real files: the 'Social Security/Medicare' withholding column (amounts) was tokenized as SSNs.
    Amounts stay; anything that could be an SSN or part of one is still tokenized, whatever the header says."""
    from pii_tokenizer.files import Context
    from pii_tokenizer.tokens import Vault
    rows = [["Employee Name", "Social Security\n/Medicare", "SSN", "Notes", "Classification"],   # a payroll export
            ["Rivera, Ana", 152.37, "123-45-6789", "SSN on file 987-65-4321", "Laborer"],
            ["Lopez, Ben", "88.10", 123456789, None, "Laborer"],
            ["Kim, Jo", 1234, "4417", None, "Laborer"]]
    name, out, _ = tokenize_file("p.xlsx", make_xlsx(rows), Context.build(CFG, KEY, Vault()))
    ws = openpyxl.load_workbook(io.BytesIO(out)).active
    assert ws["B2"].value == 152.37 and ws["B3"].value == "88.10"           # withholding amounts kept
    assert str(ws["B4"].value).startswith("SSN-")                           # 4 digits: could be a last 4
    for c in ("C2", "C3", "C4"):
        assert str(ws[c].value).startswith("SSN-"), c
    xml = all_xml(out)
    for real in ("123-45-6789", "987-65-4321", "123456789", "4417"):
        assert real not in xml, real


def test_name_pre_check_never_changes_the_result():
    """The quick pre-check before the name/ID search must give exactly the same text as the full search."""
    from pii_tokenizer.tokens import Keyring, Scrubber, Vault
    def scrubber():
        sc = Scrubber(Keyring(b"k" * 32), Vault(), {}, {"employee_name"})
        for n, t in (("Rivera, Ana", "EMP-1111111111"), ("O'Brien, Sean", "EMP-2222222222"),
                     ("Lee J Mason", "EMP-3333333333"), ("Núñez, José", "EMP-4444444444")):
            sc.add_name(n, t)
        sc.add_value("AB12345", "EID-5555555555")
        sc.add_value("123-45-6789", "SSN-6666666666")
        return sc
    texts = ["Ana Rivera called", "ana  rivera", "Ana B. Rivera", "RIVERA, ANA", "Anabel Riveras", "Sean O'Brien",
             "Lee Mason / Lee J Mason", "José Núñez", "jose nunez", "badge ab12345 ok", "AB12345-2", "123-45-6789",
             "Mason (trade)", "", "40", "Journey Level", "pay 1,234.50"]
    fast, slow = scrubber(), scrubber()
    slow._may_hold_name = lambda t: True
    slow._may_hold_value = lambda t: True
    for t in texts:
        assert fast.scrub(t) == slow.scrub(t), t


def test_remembered_results_are_forgotten_when_a_new_name_is_learnt():
    from pii_tokenizer.tokens import Keyring, Scrubber, Vault
    sc = Scrubber(Keyring(b"k" * 32), Vault(), {}, {"employee_name"})
    assert sc.scrub("call Ana Rivera") == "call Ana Rivera"
    sc.add_name("Rivera, Ana", "EMP-1111111111")
    assert "Ana Rivera" not in sc.scrub("call Ana Rivera")
    sc.add_value("AB12345", "EID-2222222222")
    assert sc.scrub("badge AB12345") == "badge EID-2222222222"


def test_numeric_ids_come_back_as_numbers_in_files_rewritten_on_aws():
    """AWS rewrites workbooks with openpyxl (tokens become shared texts); detokenizing only visits the cells that
    can hold a numeric token, and those must still become numbers again."""
    from pii_tokenizer.files import Context
    from pii_tokenizer.tokens import Vault
    v = Vault()
    rows = [["Employee Name", "Employee ID", "Hours"]] + [[f"Lopez{i}, Ben", 10000 + i, 8] for i in range(50)]
    _, tok, _ = tokenize_file("p.xlsx", make_xlsx(rows), Context.build(CFG, KEY, v))
    wb = openpyxl.load_workbook(io.BytesIO(tok))
    wb.active["D2"] = "note about " + wb.active["B2"].value      # a token inside text stays text
    buf = io.BytesIO(); wb.save(buf)
    out, _ = detokenize_file("p.xlsx", buf.getvalue(), v)
    ws = openpyxl.load_workbook(io.BytesIO(out)).active
    assert ws["B2"].value == 10000 and ws["B51"].value == 10049 and ws["A2"].value == "Lopez0, Ben"
    assert ws["D2"].value == "note about 10000"


def test_agent_state_files_are_not_synced_either_way():
    """Each side keeps its own .agent_state.json and .agent/ folder; AWS's copy must not come back as a duplicate."""
    from pii_tokenizer import config as config_mod
    cfg = config_mod.load()
    for rel in ("projects/P1/.agent_state.json", "projects/P1/.agent/lni_data/lni_snapshot.json", ".agent_state.json"):
        assert cfg.excluded(rel), rel
    assert not cfg.excluded("projects/P1/Upload/week.xlsx")
    w = World()
    w.d.put(w.raw, "projects/P1/.agent_state.json", b'{"staff": 1}')
    w.d.put(w.raw, "projects/P1/Upload/week.xlsx", make_xlsx(PAYROLL))
    w.run()
    w.d.put(w.tok, "projects/P1/.agent_state.json", b'{"aws": 1}')          # AWS writes its own
    w.d.put(w.tok, "projects/P1/.agent/lni_data/intents.csv", b"intent_id\n1\n")
    w.run()
    raw_names = [it["name"] for i, it in w.d.items.items() if not it["trashed"] and it["data"] is not None
                 and w.raw in _ancestors(w.d, i)]
    assert raw_names.count(".agent_state.json") == 1 and "intents.csv" not in raw_names, raw_names
    assert w.d.read(w.raw, "projects/P1/.agent_state.json") == b'{"staff": 1}'


def test_identical_files_are_tokenized_once():
    """Brovane: the upload and its archived copy are the same bytes; the second copy reuses the first's output."""
    from pii_tokenizer import mirror as mirror_mod
    calls = []
    real = mirror_mod.tokenize_file
    mirror_mod.tokenize_file = lambda *a, **k: calls.append(a[0]) or real(*a, **k)
    try:
        w = World()
        data = make_xlsx(PAYROLL)
        w.d.put(w.raw, "projects/P1/Upload/week.xlsx", data)
        w.d.put(w.raw, "projects/P1/Archive/2026-01-01-00-00-00/Uploaded files/week copy.xlsx", data)
        rep = w.run()
    finally:
        mirror_mod.tokenize_file = real
    assert calls == ["week.xlsx"], calls
    a = w.d.read(w.tok, "projects/P1/Upload/week.xlsx")
    b = w.d.read(w.tok, "projects/P1/Archive/2026-01-01-00-00-00/Uploaded files/week copy.xlsx")
    assert a == b and "Rivera" not in all_xml(b)



def test_a_move_synced_as_delete_and_upload_is_followed_as_a_move():
    """rclone sync without --track-renames moves the old copy into --backup-dir and uploads the same bytes at the new
    path. The staff original must be moved too, unchanged (not set aside and replaced by a detokenized copy)."""
    w = World(); d = w.d
    pay = [r + (["Gender"] if i == 0 else ["F"]) for i, r in enumerate(PAYROLL)]
    old, new = "Projects/A/Upload/week1.xlsx", "Projects/A/Payroll/Archived/2026-10-01-00-00-00/week1.xlsx"
    d.put(w.raw, old, make_xlsx(pay))
    w.run()
    original = d.read(w.raw, old)
    fid = d.path(w.tok, old)
    d.put(w.tok, new, d.items[fid]["data"])
    bk = d.put(w.tok, "_sync_backup/TS__sync_backup/.keep", b"") and d.path(w.tok, "_sync_backup/TS__sync_backup")
    parent = bk
    for x in old.split("/")[:-1]:
        parent = d.find_child(parent, x) or d.mkdir(parent, x)
    d.move(fid, parent, d.items[fid]["parents"][0])
    w.run()
    assert d.read(w.raw, old) is None
    assert d.read(w.raw, new) == original                       # the same file, byte for byte
    raw, _ = d.tree(w.raw)
    assert not any("deleted_on_aws" in f["rel"] for f in raw.values())


def test_renaming_projects_to_Projects_is_followed_and_the_old_name_not_recreated():
    w = World(); d = w.d
    for i in range(10):
        d.put(w.raw, f"projects/Acme/Payroll/w{i}.xlsx", make_xlsx(PAYROLL))
    w.run()
    d.items[d.path(w.raw, "projects")]["name"] = "Projects"          # staff rename the folder on Drive
    rep = w.run()
    assert len(rep.moved) == 10
    w.run()
    _, raw_dirs = d.tree(w.raw)
    assert not any(r.split("/")[0] == "projects" for r in raw_dirs), sorted(raw_dirs)
    assert d.read(w.tok, "Projects/Acme/Payroll/w0.xlsx")


def test_a_copy_that_vanishes_on_aws_soon_after_it_was_made_is_sent_again():
    """An AWS run that pulled before the tokenizer copied a new upload pushes its older view back, deleting the copy.
    The staff file must stay where it is and be copied again."""
    w = World()
    raw_id = w.d.put(w.raw, "Projects/A/Uploads/week2.xlsx", make_xlsx(PAYROLL))
    w.run()
    w.d.delete(w.d.path(w.tok, "Projects/A/Uploads/week2.xlsx"))
    rep = w.run()
    assert w.d.path(w.raw, "Projects/A/Uploads/week2.xlsx") == raw_id and not w.d.items[raw_id]["trashed"]
    assert w.d.read(w.tok, "Projects/A/Uploads/week2.xlsx")
    assert any("copied again" in x for x in rep.warnings), rep.warnings


def _staff_migrates(d, root, moves, project):
    """What migrate_layout.py does on the staff side: the same Drive files move (Drive for desktop keeps their IDs),
    Uploads/ is created, and the emptied old folders are moved to _sync_backup/."""
    for old, new_dir in moves:
        fid = d.path(root, old)
        d.put(root, new_dir + "/.keep", b"")
        d.delete(d.path(root, new_dir + "/.keep"))
        d.move(fid, d.path(root, new_dir), d.items[fid]["parents"][0])
    d.put(root, f"{project}/Uploads/.keep", b"")
    d.delete(d.path(root, f"{project}/Uploads/.keep"))
    d.put(root, "_sync_backup/T__layout_migration/.keep", b"")
    hold = d.path(root, "_sync_backup/T__layout_migration")
    for name in ("Upload", "Wage rates", "Reviewed payroll", "Archive"):
        fid = d.path(root, f"{project}/{name}")
        if fid:
            d.move(fid, hold, d.path(root, project))


def test_staff_side_migration_is_followed_on_aws():
    """The one-off migration runs on the staff side: the mirror moves the AWS copies the same way (no new copies),
    trashes the emptied old folders on AWS, and never puts them back on the staff side."""
    w = World(); d = w.d
    d.put(w.raw, "Projects/A/Upload/week1.xlsx", make_xlsx(PAYROLL))
    d.put(w.raw, "Projects/A/Wage rates/State.xlsx", make_xlsx([["Craft", "Classification"], ["Laborers", "GL"]]))
    d.put(w.raw, "Projects/A/Archive/Unused/old.zip", b"zip")                  # never copied to AWS
    d.put(w.raw, "Projects/B/Upload/week9.xlsx", make_xlsx(PAYROLL))           # B is not migrated
    w.run()
    tok_week = d.path(w.tok, "Projects/A/Upload/week1.xlsx")
    _staff_migrates(d, w.raw, (("Projects/A/Upload/week1.xlsx", "Projects/A/Payroll/Archived/2026-10-25-00-00-00"),
                               ("Projects/A/Wage rates/State.xlsx", "Projects/A/Files/Wage rates"),
                               ("Projects/A/Archive/Unused/old.zip", "Projects/A/Files/Unused")), "Projects/A")
    creates = len([c for c in d.calls if c[0] == "create"])
    rep = w.run()
    _, raw_dirs = d.tree(w.raw)
    _, tok_dirs = d.tree(w.tok)
    assert d.path(w.tok, "Projects/A/Payroll/Archived/2026-10-25-00-00-00/week1.xlsx") == tok_week   # moved, same file
    assert d.read(w.tok, "Projects/A/Files/Wage rates/State.xlsx")
    assert len([c for c in d.calls if c[0] == "create"]) == creates, d.calls                        # nothing re-sent
    for dirs in (raw_dirs, tok_dirs):
        assert "Projects/A/Uploads" in dirs
        assert not any(r.split("/")[2:3] and r.split("/")[2] in ("Upload", "Wage rates", "Archive")
                       for r in dirs if r.startswith("Projects/A/")), sorted(dirs)
    assert "Projects/B/Upload" in raw_dirs and "Projects/B/Upload" in tok_dirs and "Projects/B/Uploads" not in raw_dirs
    assert any("removed (AWS side)" in x for _, x in rep.folders), rep.folders
    assert d.read(w.raw, "Projects/A/Files/Unused/old.zip") == b"zip"
    # quiet next run
    rep2 = w.run()
    assert not rep2.folders and not rep2.warnings, (rep2.folders, rep2.warnings)


def test_aws_old_folder_holding_a_file_is_kept_and_a_refusal_is_warned_once():
    w = World(); d = w.d
    d.put(w.raw, "Projects/A/Upload/week1.xlsx", make_xlsx(PAYROLL))
    w.run()
    d.put(w.tok, "Projects/A/Upload/stray.txt", b"x")                        # something on AWS only
    _staff_migrates(d, w.raw, (("Projects/A/Upload/week1.xlsx", "Projects/A/Payroll/Archived/T"),), "Projects/A")
    w.run()
    assert d.path(w.tok, "Projects/A/Upload") is not None                      # holds a file: kept
    d.delete(d.path(w.tok, "Projects/A/Upload/stray.txt"))
    def refuse(*a_, **k_):
        raise RuntimeError("insufficientFilePermissions")
    orig, d.trash = d.trash, refuse
    rep = w.run()
    assert any("could not be removed" in x for x in rep.warnings), rep.warnings
    rep = w.run()
    assert not any("could not be removed" in x for x in rep.warnings), rep.warnings      # warned about once
    d.trash = orig
    w.run()
    assert d.path(w.tok, "Projects/A/Upload") is None


def test_uploads_made_on_aws_appears_for_staff():
    """A project migrated by a run on AWS (the agent creates Uploads/ empty there, which rclone doesn't copy)."""
    w = World(); d = w.d
    d.put(w.raw, "Projects/C/Tracker/T.xlsx", make_xlsx([["Intent ID"], ["123"]]))
    w.run()
    d.put(w.tok, "Projects/C/Files/Logs/T/.keep", b"")
    d.delete(d.path(w.tok, "Projects/C/Files/Logs/T/.keep"))
    w.run()
    _, raw_dirs = d.tree(w.raw)
    assert "Projects/C/Uploads" in raw_dirs, sorted(raw_dirs)


def test_a_project_not_yet_migrated_is_left_alone_even_with_a_converted_macro_workbook():
    """Converting an .xlsm parks the original in Files/Archived/<ts>/ before the agent has migrated the project:
    that alone must not make the project look migrated (no Uploads/ created, no old folders removed)."""
    w = World(); d = w.d
    d.put(w.raw, "Projects/W/Upload/week.xlsm", _xlsm())
    d.put(w.raw, "Projects/W/Reviewed payroll/.keep", b"")
    d.delete(d.path(w.raw, "Projects/W/Reviewed payroll/.keep"))          # an empty old-layout folder
    rep = w.run()
    _, raw_dirs = d.tree(w.raw)
    assert any(r.startswith("Projects/W/Files/Archived") for r in raw_dirs)
    assert "Projects/W/Uploads" not in raw_dirs and "Projects/W/Reviewed payroll" in raw_dirs, sorted(raw_dirs)
    assert not rep.folders, rep.folders


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"\nAll {len(tests)} tests passed.")


def test_backups_and_underscore_project_folders_are_not_copied():
    """tmp/ (staff backups, scratch copies of projects) and Projects/_* folders (templates, previous versions,
    exports; the agent ignores them as projects) never go to AWS: a full AWS run pulls the whole tokenized copy."""
    from pii_tokenizer import config as config_mod
    cfg = config_mod.load()
    for rel in ("tmp/Projects_original/Brennick - 77Q/Uploads/week.xlsx", "tmp/notes.xlsx",
                "Projects/__PREVIOUS/Atrovan - XE9901/Tracker/t.xlsx", "Projects/__Project Template/Uploads/x.csv",
                "Projects/__inferred_trackers/Tots Co Compliance Tracker.xlsx", "projects/_old/x.xlsx"):
        assert cfg.excluded(rel), rel
    for rel in ("Projects/Brennick - Tots Co/Uploads/week.xlsx", "Projects/Atrovan - 9540 SR 901 1b/Tracker/t.xlsx",
                "reference/knowledge/guide.md", "Projects/Corvath - XHA 50001/Files/Logs/2026-11-04__log.txt"):
        assert not cfg.excluded(rel), rel


def test_newly_excluded_folder_does_not_trip_the_safety_stop():
    """Most of the files already in step sit in a folder that is excluded afterwards: they are forgotten (a
    warning says how many), not counted as deleted, so the "more than half look deleted" stop doesn't fire; their
    copies stay where they are and the staff files are untouched."""
    import dataclasses
    w = World()
    for i in range(8):
        w.d.put(w.raw, f"tmp/Projects_original/P1/Uploads/rates{i}.csv", b"County,Trade\nKing,Laborer\n")
    for i in range(2):
        w.d.put(w.raw, f"Projects/P2/Uploads/rates{i}.csv", b"County,Trade\nKing,Laborer\n")
    w.ctx.config = dataclasses.replace(CFG, exclude=[p for p in CFG.exclude if p != "tmp/*"])
    w.run()
    assert len(w.state["mirrors"]["compliance"]["pairs"]) == 10
    w.ctx.config = CFG
    rep = w.run()                                     # would raise "8 of 10 files look deleted" before the fix
    assert len(w.state["mirrors"]["compliance"]["pairs"]) == 2
    assert any("no longer kept in step" in x for x in rep.warnings)
    assert w.d.read(w.tok, "tmp/Projects_original/P1/Uploads/rates0.csv") is not None
    assert w.d.read(w.raw, "tmp/Projects_original/P1/Uploads/rates0.csv") == b"County,Trade\nKing,Laborer\n"


def test_exclusions_ignore_case():
    from pii_tokenizer import config as config_mod
    cfg = config_mod.load()
    for rel in ("Tmp/x.xlsx", "TMP/a/b.csv", "PROJECTS/_old/x.xlsx", "projects/__Template/y.csv"):
        assert cfg.excluded(rel), rel
    assert not cfg.excluded("Projects/Brennick - Tots Co/Uploads/x.csv")


def test_newly_excluded_folder_keeps_a_pending_take_back():
    import dataclasses
    w = World()
    w.d.put(w.raw, "tmp/a.csv", b"County,Trade\nKing,Laborer\n")
    w.d.put(w.raw, "Projects/P2/Uploads/b.csv", b"County,Trade\nKing,Laborer\n")
    w.ctx.config = dataclasses.replace(CFG, exclude=[p for p in CFG.exclude if p != "tmp/*"])
    w.run()
    pairs = w.state["mirrors"]["compliance"]["pairs"]
    next(q for q in pairs if q["rel"] == "tmp/a.csv")["pending_take_back"] = True
    w.ctx.config = CFG
    w.run()
    rels = [q["rel"] for q in w.state["mirrors"]["compliance"]["pairs"]]
    assert rels == ["Projects/P2/Uploads/b.csv"]       # the pending pair was finished, not dropped unfinished:
    assert w.d.read(w.tok, "tmp/a.csv") is None         # its copy was taken back (binned) from the AWS side
    assert w.d.read(w.raw, "tmp/a.csv") == b"County,Trade\nKing,Laborer\n"   # the staff file is untouched


def test_a_new_tracker_status_word_reaches_aws_readable():
    """#151 step 2 (design note section 3, review F9): the Settings page can add status words the agent matches, so
    the tracker's Status column must cross whole and readable, whatever the word."""
    w = World()
    hdr = ["Status", "Contractor", "Intent ID"]
    words = ["Re-Mobilized", "Awaiting Paperwork", "Démobilisé"]
    w.d.put(w.raw, "tracker.xlsx", make_xlsx([hdr] + [[s, "Delta Mechanical", 1000 + i] for i, s in enumerate(words)]))
    w.run()
    rows = _rows(w.d.read(w.tok, "tracker.xlsx"))
    assert [r[0] for r in rows[1:]] == words, rows


def test_a_header_pasted_twice_does_not_teach_its_labels_as_names():
    """Seen in the evaluation set (Nov 13, 2026): a payroll with its header row again below the lines (two exports in
    one sheet) holds "Employee Name" in the Employee Name column. Learnt as a name, it replaced that label in every
    later file of the run, so no payroll was recognized. The cell is still tokenized where it stands."""
    c = ctx()
    header = PAYROLL[1]
    twice = PAYROLL + [header] + PAYROLL[2:]
    _, out, _ = tokenize_file("pasted.xlsx", make_xlsx(twice), c)
    ws = openpyxl.load_workbook(io.BytesIO(out)).active
    col = header.index("Employee Name") + 1
    assert ws.cell(2, col).value == "Employee Name"                          # the real header: a label
    assert str(ws.cell(len(PAYROLL) + 1, col).value).startswith("EMP-")      # the pasted one: tokenized in place
    assert str(ws.cell(len(PAYROLL) + 2, col).value).startswith("EMP-")      # the lines below it still are
    _, out2, _ = tokenize_file("next.xlsx", make_xlsx(PAYROLL), c)          # a later file of the same run
    ws2 = openpyxl.load_workbook(io.BytesIO(out2)).active
    assert ws2.cell(2, col).value == "Employee Name"
    assert c.scrubber.scrub("Employee Name, Rivera, Ana") .startswith("Employee Name, EMP-")
    # a vault that learnt the label before this fix: forgotten as soon as a header row shows it is a label
    tok = c.keyring.token("EMP", "employee_name", "Employee Name")
    c.vault.add(tok, "employee_name", "Employee Name")
    again = Context.build(CFG, KEY, c.vault)
    assert again.scrubber.scrub("Employee Name") == tok                     # learnt from the vault...
    _, out3, _ = tokenize_file("third.xlsx", make_xlsx(PAYROLL), again)
    assert openpyxl.load_workbook(io.BytesIO(out3)).active.cell(2, col).value == "Employee Name"
    assert again.scrubber.scrub("Employee Name") == "Employee Name"         # ...and forgotten


def test_ids_that_match_a_header_pattern_are_still_replaced_in_free_text():
    """Independent review (Nov 13, 2026), high: a first version skipped any value matching a header *pattern*, so an
    employee ID such as "EMP#55555" stayed readable in a note. Only labels seen in a header row are skipped."""
    c = ctx()
    rows = [["Employee Name", "Employee ID", "Classification", "Notes"],
            ["Rivera, Ana", "EMP#55555", "Laborer", "Ask about EMP#55555 hours, also EMPID7777"],
            ["Mason, Lee", "EMPID7777", "Laborer", "Employee Name: Rivera, Ana"]]
    _, out, _ = tokenize_file("ids.xlsx", make_xlsx(rows), c)
    ws = openpyxl.load_workbook(io.BytesIO(out)).active
    note = str(ws.cell(2, 4).value) + " " + str(ws.cell(3, 4).value)
    assert "55555" not in note and "7777" not in note and "Rivera" not in note
    assert ws.cell(1, 1).value == "Employee Name"


def test_a_header_pasted_twice_in_a_csv_does_not_teach_its_labels():
    c = ctx()
    text = "Employee Name,Employee ID,Classification\n\"Rivera, Ana\",10234,Laborer\n" \
           "Employee Name,Employee ID,Classification\n\"Mason, Lee\",10235,Laborer\n"
    _, out, _ = tokenize_file("twice.csv", text.encode(), c)
    got = out.decode()
    assert got.splitlines()[0].startswith("Employee Name,Employee ID") and "Lee" not in got and "Ana" not in got
    assert c.scrubber.scrub("Employee Name") == "Employee Name"


def test_a_data_cell_that_looks_like_a_label_never_unlearns_an_id():
    """Independent verification (Nov 13, 2026), high: "EMP#55555" in a column with no personal-data header (a log's
    Topic) was taken for a header label and made the tokenizer forget that learnt ID: readable in the rest of the
    run. Only a row of two or more labels without digits counts as a header row for this."""
    c = ctx()
    rows = [list(r) for r in PAYROLL]
    rows[2][4] = "EMP#55555"
    tokenize_file("p.xlsx", make_xlsx(rows), c)
    _, out, _ = tokenize_file("log.csv", b"Date,Topic,Note\n11/1/2026,EMP#55555,called re EMP#55555\n", c)
    assert "55555" not in out.decode() and out.decode().count("EID-") == 2
    wb = [["Date", "Topic", "Note"], ["11/1/2026", "EMP#55555", "called re EMP#55555"],
          ["11/2/2026", "Employee Name", "Last Name"]]                      # a stray pair of labels in a data row
    _, out, _ = tokenize_file("log.xlsx", make_xlsx(wb), c)
    ws = openpyxl.load_workbook(io.BytesIO(out)).active
    assert "55555" not in f"{ws['B2'].value} {ws['C2'].value}"
    _, out, _ = tokenize_file("n.txt", b"follow up EMP#55555\n", c)
    assert "55555" not in out.decode()
    assert c.scrubber.scrub("EMP#55555 and Rivera, Ana").count("-") >= 2


def test_only_a_real_header_row_shows_labels():
    from pii_tokenizer.config import HeaderTracker
    h = HeaderTracker(CFG)
    h.row([(1, "Employee Name"), (2, "Employee ID"), (3, "Classification")])
    assert h.last_labels == ["Employee Name", "Employee ID"]
    h2 = HeaderTracker(CFG)
    h2.row([(1, "Date"), (2, "EMP#55555"), (3, "Note")])                    # one label, with a digit
    assert h2.last_labels == []
    h3 = HeaderTracker(CFG)
    h3.row([(1, "Employee Name"), (2, "Notes")])                            # one label alone: not a header row
    assert h3.last_labels == []
    h4 = HeaderTracker(CFG)
    h4.row([(1, "Employee Name"), (2, "EMP#5")])                            # only one without a digit
    assert h4.last_labels == []
    h5 = HeaderTracker(CFG)
    h5.row([(1, "Employee Name"), (2, "Address Line 1"), (3, "Employee ID")])   # labels with a digit left out
    assert h5.last_labels == ["Employee Name", "Employee ID"]
