"""'If in doubt, assume it is sensitive' (pii_tokenizer/doubt.py): in a tracker, every text the agent doesn't read
reaches AWS only as a text token, and comes back exactly as it was, formulas and notes included; what the agent reads
still works. The agent's own code (code/compliance_agent) is used to check that."""
from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import openpyxl
import pytest
from openpyxl.comments import Comment

from helpers import CFG, KEY
from pii_tokenizer import doubt
from pii_tokenizer.files import Context, detokenize_file, tokenize_file
from pii_tokenizer.tokens import Vault

AGENT = Path(__file__).resolve().parents[2] / "compliance_agent"
SECRETS = ("Maria", "Ortega", "Pedro Unknown", "pass: hunter2", "Staff Person", "Linda Example", "Trainee Tom")


def _tracker() -> bytes:
    wb = openpyxl.Workbook()
    info = wb.active
    info.title = "Project Info"
    for row in (["Eddleton Boiler"], ["Prime Intent", 1234567], ["PROJECT TEAM"],
                ["Linda Example - Project Admin"], ["No L&I Filing Required"]):
        info.append(row)
    t = wb.create_sheet("Eddleton Boiler")
    t.append(["Status", "Contractor", "Hiring Contractor", "Intent Filed", "Intent ID", "Section 3 Employees",
              "LCP Account Info", "Notes"])
    t.append(["Onsite", "Acme Paving", "Prime GC", "Yes", 1111111, "Pedro Unknown", "user: acme pass: hunter2",
              "Maria Ortega missing affidavit"])
    t.append(["Completed Work", "Bolt Electric", "Acme Paving", "No", None, None, None, '=CONCAT("Maria ", "Ortega")'])
    t["H2"].comment = Comment("Call Maria Ortega", "Staff Person")
    t["B3"].comment = Comment("Compliance Validation Agent:\n- old mark", "Compliance Validation Agent")
    tr = wb.create_sheet("24092A1 - XE9901 Trainees ")
    tr.append(["Special Training By Classification", "Totals by Trade"])
    tr.append(["Trainee Tom", 3])
    tr.append(["Total", "=SUM(B2:B2)"])
    au = wb.create_sheet("Audit")
    au.append(["Run (UTC)", "Task", "Topic", "Note"])
    au.append(["2026-11-01-00-00-00", "intent/affidavit review", "Marks (agent use)", "Eddleton Boiler!H2 ..."])
    md = wb.create_sheet("Metadata")
    md.append(["Field", "Key", "Value", "Example (from current trackers)"])
    md.append(["Contact", "cp_contact", "", "Linda Example"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _xml(data: bytes) -> str:
    z = zipfile.ZipFile(io.BytesIO(data))
    return "".join(z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith((".xml", ".vml", ".rels")))


def _values(data: bytes, formulas: bool) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=not formulas)
    return {(ws.title, c.coordinate): c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if c.value is not None}


def _notes(data: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data))
    return {(ws.title, c.coordinate): (c.comment.text, c.comment.author) for ws in wb.worksheets
            for row in ws.iter_rows() for c in row if c.comment}


def _tokenize(data: bytes, vault: Vault):
    return tokenize_file("Eddleton Boiler Compliance Tracker.xlsx", data, Context.build(CFG, KEY, vault))


def test_nothing_the_agent_does_not_read_reaches_aws_in_clear():
    """#53's acceptance test, for trackers: an unknown name in a note, a free-text column, another sheet, a formula
    and a cell note never appears in the copy; nor does a staff member's name or what looks like a password."""
    _, out, summary = _tokenize(_tracker(), Vault())
    xml = _xml(out)
    for secret in SECRETS:
        assert secret not in xml, secret
    assert "text tokens" in summary
    assert "Compliance Validation Agent" in xml          # the agent finds its own notes by this author name
    vals = _values(out, formulas=False)
    # what the agent reads is as it was
    for ref, v in ((("Eddleton Boiler", "A2"), "Onsite"), (("Eddleton Boiler", "B2"), "Acme Paving"),
                   (("Eddleton Boiler", "C2"), "Prime GC"), (("Eddleton Boiler", "D2"), "Yes"),
                   (("Eddleton Boiler", "E2"), 1111111), (("Eddleton Boiler", "F1"), "Section 3 Employees"),
                   (("Project Info", "A2"), "Prime Intent"), (("Project Info", "B2"), 1234567),
                   (("Project Info", "A5"), "No L&I Filing Required"), (("Audit", "C2"), "Marks (agent use)"),
                   (("Trainees", None), None)):
        if ref[1]:
            assert vals[ref] == v, ref
    assert vals[("24092A1 - XE9901 Trainees ", "B2")] == 3   # numbers stay


def test_round_trip_gives_back_every_value_formula_and_note():
    vault = Vault()
    src = _tracker()
    _, out, _ = _tokenize(src, vault)
    back, _ = detokenize_file("t.xlsx", out, vault)
    assert _values(back, formulas=True) == _values(src, formulas=True)
    assert _notes(back) == _notes(src)


def test_the_same_text_gives_the_same_token():
    """Deterministic: a tracker tokenized twice gives identical copies (no churn on AWS between runs)."""
    src = _tracker()       # built once: openpyxl stamps the time into docProps/core.xml (a flake across a second)
    a = _tokenize(src, Vault())[1]
    b = _tokenize(src, Vault())[1]
    assert _xml(a) == _xml(b)


@pytest.fixture()
def agent():
    if not (AGENT / "agent" / "intent.py").exists():
        pytest.skip("the agent's code is not next to the tokenizer")
    sys.path.insert(0, str(AGENT))
    try:
        from agent import intent, knowledge, xlsx_patch       # noqa: F401 - CI installs both components' needs
    except ImportError as e:                                  # pragma: no cover
        pytest.skip(f"the agent's requirements are not installed ({e})")
    yield intent, knowledge, xlsx_patch
    sys.path.remove(str(AGENT))


def test_the_agent_reads_the_copy_as_it_reads_the_original(agent, tmp_path):
    intent, _, _ = agent
    _, out, _ = _tokenize(_tracker(), Vault())
    orig = openpyxl.load_workbook(io.BytesIO(_tracker()), data_only=True)
    copy = openpyxl.load_workbook(io.BytesIO(out), data_only=True)
    ws_o, r_o, cols_o = intent.locate_table(orig)
    ws_c, r_c, cols_c = intent.locate_table(copy)
    assert (ws_c.title, r_c, cols_c) == (ws_o.title, r_o, cols_o)
    for r in range(r_o + 1, ws_o.max_row + 1):
        for role, i in cols_o.items():
            assert ws_c.cell(r, i + 1).value == ws_o.cell(r, i + 1).value, (role, r)
    assert intent.project_info(copy)["prime_intent"] == intent.project_info(orig)["prime_intent"]
    assert str(intent.project_info(copy)["prime_intent"]) == "1234567"
    cfg = {"projects": {}}
    assert intent.is_applicable("P", copy, cfg) == intent.is_applicable("P", orig, cfg)


def test_staff_notes_survive_the_agent_editing_the_copy(agent, tmp_path):
    """The agent marks the copy on AWS (fill, note, log) and the result is restored over the staff file: every
    tokenized value and note comes back, next to the agent's marks."""
    _, _, xlsx_patch = agent
    vault = Vault()
    _, out, _ = _tokenize(_tracker(), vault)
    p = tmp_path / "t.xlsx"
    p.write_bytes(out)
    xp = xlsx_patch.XlsxPatcher(p)
    xp.fill_cell("Eddleton Boiler", "E2", "FFFF0000")
    xp.add_note("Eddleton Boiler", "B2", "Compliance Validation Agent:\n- Intent ID not found", "Compliance Validation Agent")
    xp.append_log("Review Log", ["Run (UTC)", "Sheet", "Cell", "Issue"], [["2026-11-04", "Eddleton Boiler", "E2", "x"]])
    xp.save(p)
    back, _ = detokenize_file("t.xlsx", p.read_bytes(), vault)
    vals = _values(back, formulas=True)
    assert vals[("Eddleton Boiler", "H2")] == "Maria Ortega missing affidavit"
    assert vals[("Eddleton Boiler", "G2")] == "user: acme pass: hunter2"
    assert vals[("Eddleton Boiler", "H3")] == '=CONCAT("Maria ", "Ortega")'
    assert vals[("24092A1 - XE9901 Trainees ", "A2")] == "Trainee Tom"
    notes = _notes(back)
    assert notes[("Eddleton Boiler", "H2")] == ("Call Maria Ortega", "Staff Person")
    assert "Intent ID not found" in notes[("Eddleton Boiler", "B2")][0]


def test_payroll_exports_and_wage_tables_keep_the_earlier_handling():
    for rows in ([["Employee Name", "Classification", "Notes"], ["Rivera, Ana", "Laborer", "ok"]],
                 [["Craft", "Classification", "Basic Rate"], ["LABORER", "Group 1", 50.1]]):
        wb = openpyxl.Workbook()
        for r in rows:
            wb.active.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
        assert "TXT-" not in _xml(out)


def test_rules_come_from_the_registry_and_match_the_agents_config(agent):
    """Was a kept-in-step test; since vNext PR 3 both sides import the tracker, Project Info and Project Data kinds
    from the registry (test_filekinds.py checks the copies are the source). The agent's config.yaml part stays."""
    intent, knowledge, _ = agent
    import yaml
    from agent.cva_filekinds import tracker as agent_tracker
    from pii_tokenizer.cva_filekinds import tracker as our_tracker
    assert doubt.TRACKER_ROLES is our_tracker.TRACKER_ROLES and intent.HEADER_ROLES is agent_tracker.TRACKER_ROLES
    assert doubt.AGENT_AUTHORS is our_tracker.AGENT_AUTHORS and intent.AUTHOR is agent_tracker.AUTHOR
    assert doubt.TRACKER_ROLES == intent.HEADER_ROLES
    assert {intent.AUTHOR, *intent.OLD_AUTHORS} <= doubt.AGENT_AUTHORS
    cfg = yaml.safe_load((AGENT / "config.yaml").read_text())
    for name in (cfg["intent"]["audit_sheet"], cfg["intent"]["review_log_sheet"]):
        assert name.lower() in doubt.AGENT_SHEETS, name
    assert doubt.NO_FILING_RE.fullmatch("No L&I Filing Required")
    assert doubt.norm_header("  Intent \n ID ") == "intent id"


def _book(build) -> bytes:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _tracker_sheet(ws):
    ws.append(["Status", "Contractor", "Intent ID", "Notes"])
    ws.append(["Onsite", "Acme Paving", 1111111, "Yolanda Xu called"])


def test_text_outside_the_cells_is_tokenized_too():
    """From the review: dropdown lists, filter values, headers and footers, hyperlink text."""
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        ws = wb.active
        ws.title = "Tracker"
        _tracker_sheet(ws)
        dv = DataValidation(type="list", formula1='"Yolanda Xu,Other"', prompt="Ask Yolanda Xu")
        ws.add_data_validation(dv)
        dv.add("D2:D9")
        ws.oddHeader.center.text = "Yolanda Xu"
        ws.auto_filter.ref = "A1:D2"
        ws.auto_filter.add_filter_column(3, ["Yolanda Xu called"])
        ws["D2"].hyperlink = "#Tracker!A1"
        ws["D2"].hyperlink.tooltip = "Yolanda Xu"
    vault = Vault()
    src = _book(build)
    _, out, _ = _tokenize(src, vault)
    assert "Yolanda" not in _xml(out)
    back, _ = detokenize_file("t.xlsx", out, vault)
    assert _values(back, formulas=True) == _values(src, formulas=True)
    w = openpyxl.load_workbook(io.BytesIO(back)).active
    assert w.data_validations.dataValidation[0].formula1 == '"Yolanda Xu,Other"'
    assert w.oddHeader.center.text == "Yolanda Xu"


def test_a_classification_column_on_another_tab_keeps_the_tracker_in_doubt_mode():
    def build(wb):
        _tracker_sheet(wb.active)
        wb.create_sheet("Apprentices").append(["Craft", "Classification", "Name of apprentice"])
        wb["Apprentices"].append(["LABORER", "Group 1", "Yolanda Xu"])
    _, out, _ = _tokenize(_book(build), Vault())
    assert "Yolanda" not in _xml(out)


def test_a_staff_tab_named_comments_is_not_the_agents():
    def build(wb):
        _tracker_sheet(wb.active)
        wb.create_sheet("Comments").append(["Yolanda Xu: follow up"])
    _, out, _ = _tokenize(_book(build), Vault())
    assert "Yolanda" not in _xml(out)


def test_a_wage_export_with_its_header_below_row_11_is_found():
    """Rows above the header are tokens, except the "Location:" line the agent reads the county from."""
    def build(wb):
        ws = wb.active
        for _ in range(12):
            ws.append(["Project Wage Data Export"])
        ws.append(["Location: Pierce County"])
        ws.append(["Craft", "Classification", "Basic Rate"])
        ws.append(["LABORER", "Group 1", 50.1])
    _, out, _ = _tokenize(_book(build), Vault())
    xml = _xml(out)
    assert all(w in xml for w in ("Location: Pierce County", "Craft", "LABORER", "Group 1"))
    assert "Project Wage Data Export" not in xml


def test_project_info_is_the_first_info_tab_in_tab_order():
    """As the agent sees it (wb.worksheets order), even when the file stores the sheets in another order."""
    def build(wb):
        _tracker_sheet(wb.active)
        a = wb.create_sheet("Old Info")
        a.append(["Prime Intent", 999])
        b = wb.create_sheet("Project Info")
        b.append(["Prime Intent", 1234567])
        wb.move_sheet("Project Info", offset=-1)        # tab order: Tracker, Project Info, Old Info
    _, out, _ = _tokenize(_book(build), Vault())
    vals = _values(out, formulas=False)
    assert vals[("Project Info", "A1")] == "Prime Intent" and vals[("Project Info", "B1")] == 1234567
    assert str(vals[("Old Info", "A1")]).startswith("TXT-")


def test_line_breaks_in_attributes_come_back_exactly():
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        _tracker_sheet(wb.active)
        dv = DataValidation(type="list", formula1='"A,B"', prompt="Line one\nLine two\tend", promptTitle="T")
        wb.active.add_data_validation(dv)
        dv.add("D2")
    vault = Vault()
    _, out, _ = _tokenize(_book(build), vault)
    back, _ = detokenize_file("t.xlsx", out, vault)
    assert openpyxl.load_workbook(io.BytesIO(back)).active.data_validations.dataValidation[0].prompt == \
        "Line one\nLine two\tend"


def test_a_wage_export_saved_on_another_tab_is_found_on_the_active_tab():
    def build(wb):
        wb.active.append(["Summary"])
        ws = wb.create_sheet("Wages")
        ws.append(["Craft", "Classification", "Basic Rate"])
        ws.append(["LABORER", "Group 1", 50.1])
        wb.active = 1
    _, out, _ = _tokenize(_book(build), Vault())
    xml = _xml(out)
    assert "LABORER" in xml and "Group 1" in xml and "Summary</t>" not in xml   # the other tab: text tokens


def _raw_gt(data: bytes, part: str, old: str, new: str) -> bytes:
    """The workbook with ``old`` replaced by ``new`` in one part: openpyxl writes ">" in an attribute as "&gt;",
    but Excel and Google Sheets may write it raw, which is legal XML."""
    src, out = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in src.infolist():
            b = src.read(i.filename)
            if i.filename == part:
                assert old.encode() in b
                b = b.replace(old.encode(), new.encode())
            z.writestr(i, b)
    return out.getvalue()


def _parses(data: bytes):
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(io.BytesIO(data))
    for n in z.namelist():
        if n.endswith((".xml", ".rels")):
            ET.fromstring(z.read(n))             # raises on a part that isn't well-formed


def test_a_raw_gt_inside_an_attribute_does_not_break_the_sheet():
    """QRX's tracker, Nov 4: two drop-downs whose error message held a raw ">". The text-outside-the-cells pass
    started at that ">", swallowed the rest of the tag and the next one, and the agent could not read the copy
    (not well-formed: error="TXT-..."list" ...). Now the copy parses, and the messages come back exactly."""
    from openpyxl.worksheet.datavalidation import DataValidation
    msg = "Pick one > or ask Yolanda Xu"
    def build(wb):
        ws = wb.active
        _tracker_sheet(ws)
        for ref, items in (("A2:A658", '"Onsite,Completed Work"'), ("B2:B658", '"Acme Paving,Other"')):
            dv = DataValidation(type="list", formula1=items, allow_blank=True, showErrorMessage=True,
                                errorStyle="warning", error=msg, errorTitle="Check > this", prompt="Use > list")
            ws.add_data_validation(dv)
            dv.add(ref)
        ws.cell(1, 1).hyperlink = "https://example.com/?a=1"
    src = _book(build)
    src = _raw_gt(src, "xl/worksheets/sheet1.xml", "&gt;", ">")
    assert b'error="Pick one > or' in zipfile.ZipFile(io.BytesIO(src)).read("xl/worksheets/sheet1.xml")
    vault = Vault()
    _, out, _ = _tokenize(src, vault)
    _parses(out)
    assert "Yolanda" not in _xml(out)
    w = openpyxl.load_workbook(io.BytesIO(out)).active
    dvs = w.data_validations.dataValidation
    assert [str(d.sqref) for d in dvs] == ["A2:A658", "B2:B658"]
    assert all(d.type == "list" for d in dvs)
    back, _ = detokenize_file("t.xlsx", out, vault)
    _parses(back)
    dvs = openpyxl.load_workbook(io.BytesIO(back)).active.data_validations.dataValidation
    assert [(d.error, d.errorTitle, d.prompt, d.formula1) for d in dvs] == [
        (msg, "Check > this", "Use > list", '"Onsite,Completed Work"'),
        (msg, "Check > this", "Use > list", '"Acme Paving,Other"')]
    # and again: the copy written back is tokenized again on the next pass
    _, again, _ = _tokenize(back, vault)
    _parses(again)


def test_single_quoted_or_spaced_attributes_are_tokenized_too():
    """From the review of #80: a value in single quotes, or with spaces around "=", is legal XML that Excel and
    Google don't write but other programs may; the attribute passes only knew name="value", so such text went out
    in clear. Now every attribute is rewritten as name="value" first."""
    from openpyxl.worksheet.datavalidation import DataValidation
    def build(wb):
        ws = wb.active
        _tracker_sheet(ws)
        for ref, msg in (("A2", "Ask Yolanda Xu"), ("B2", "Or Ana Rivera")):
            dv = DataValidation(type="list", formula1='"Onsite,Completed Work"', prompt=msg)
            ws.add_data_validation(dv)
            dv.add(ref)
    src = _book(build)
    src = _raw_gt(src, "xl/worksheets/sheet1.xml", 'prompt="Ask Yolanda Xu"', "prompt='Ask \"Yolanda\" Xu'")
    src = _raw_gt(src, "xl/worksheets/sheet1.xml", 'prompt="Or Ana Rivera"', 'prompt = "Or Ana Rivera"')
    vault = Vault()
    _, out, _ = _tokenize(src, vault)
    _parses(out)
    assert "Yolanda" not in _xml(out) and "Rivera" not in _xml(out)
    back, _ = detokenize_file("t.xlsx", out, vault)
    dvs = openpyxl.load_workbook(io.BytesIO(back)).active.data_validations.dataValidation
    assert [d.prompt for d in dvs] == ['Ask "Yolanda" Xu', "Or Ana Rivera"]


def test_canonical_attributes_keeps_comments_cdata_and_vml_intact():
    """From the review of #80: a quote inside a comment or CDATA must not be taken for the start of a value, and
    VML parts (cell-note shapes) get the same treatment."""
    from pii_tokenizer.xlsx import canonical_attributes
    import xml.etree.ElementTree as ET
    x = '<a><!-- don"t --><![CDATA[it\'s]]><b x="1>2" y=\'z\' w = "v"/></a>'
    out = canonical_attributes(x)
    assert out == '<a><!-- don"t --><![CDATA[it\'s]]><b x="1&gt;2" y="z" w="v"/></a>'
    ET.fromstring(out)
    assert canonical_attributes('<a b="c">d</a>') == '<a b="c">d</a>'       # nothing to do: unchanged

    def build(wb):
        _tracker_sheet(wb.active)
        wb.active.cell(2, 4).comment = Comment("Yolanda Xu called back", "Ana Rivera")
    src = _book(build)
    vml = next(n for n in zipfile.ZipFile(io.BytesIO(src)).namelist() if n.endswith(".vml"))
    src = _raw_gt(src, vml, "<ns0:shape ", '<ns0:shape alt="a > b" ')
    vault = Vault()
    _, out, _ = _tokenize(src, vault)
    _parses(out)
    ET.fromstring(zipfile.ZipFile(io.BytesIO(out)).read(vml))
    assert "Yolanda" not in _xml(out)


def test_affidavit_good_and_cp_status_are_text_tokens_like_any_other_column():
    """Nov 7, 2026: the agent no longer reads Affidavit Good or CP Status (David's decision), so their text is a text
    token like any column the agent doesn't read (if in doubt, sensitive); a checkbox stays a checkbox, and every
    cell comes back as written."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Site Work"
    ws.append(["Status", "Contractor", "Intent ID", "Affidavit ID", "Affidavit Good", "CP Status"])
    ws.append(["Closed", "Acme Paving", 1111111, 2222222, True, "Final"])
    ws.append(["Closed", "Bolt Electric", 1111112, 2222223, "No", "Not Final"])
    ws.append(["Closed", "Zeta Concrete", 1111113, 2222224, "No - Marisol Venkataraman's hours", "Final - Odin Pryce"])
    buf = io.BytesIO()
    wb.save(buf)
    vault = Vault()
    _, out, _ = _tokenize(buf.getvalue(), vault)
    xml = _xml(out)
    for secret in ("Marisol", "Venkataraman", "Odin Pryce"):
        assert secret not in xml, secret
    vals = _values(out, formulas=False)
    assert vals[("Site Work", "E2")] is True                                  # a checkbox isn't text: it stays
    for ref in ("F2", "E3", "F3", "E4", "F4"):                                 # every text, the fixed words too
        assert str(vals[("Site Work", ref)]).startswith("TXT-"), ref
    assert not doubt.TRACKER_VOCAB
    back = _values(detokenize_file("Eddleton Boiler Compliance Tracker.xlsx", out, vault)[0],
                   formulas=False)
    assert back[("Site Work", "E4")] == "No - Marisol Venkataraman's hours"
    assert back[("Site Work", "F2")] == "Final" and back[("Site Work", "E3")] == "No"


def test_former_vocab_columns_keep_no_formula_no_phonetic_text_and_shared_strings_work():
    """Review of #89, kept after Nov 7: a formula in Affidavit Good or CP Status could quote a name, and a phonetic run
    hides text behind "Yes": none of it reaches AWS. Shared strings behave like inline ones. (Since Nov 7 neither
    column is read by the agent, so even the plain words are text tokens.)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Site Work"
    ws.append(["Status", "Contractor", "Intent ID", "Affidavit ID", "Affidavit Good", "CP Status", "Notes"])
    ws.append(["Closed", "Acme Paving", 1111111, 2222222, "Yes", "Final", "Yes"])
    ws.append(["Closed", "Bolt Electric", 1111112, 2222223, None, None, None])
    buf = io.BytesIO()
    wb.save(buf)
    # rewrite the sheet: shared strings for row 2, crafted formulas and a phonetic run in row 3
    src = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    sheet = src.read("xl/worksheets/sheet1.xml").decode()
    import re as _re
    texts, sst = [], []
    def to_shared(m):
        t = _re.search(r"<t[^>]*>(.*?)</t>", m.group(2)).group(1)
        sst.append(t)
        return f'<c r="{m.group(1)}" t="s"><v>{len(sst) - 1}</v></c>'
    sheet = _re.sub(r'<c r="([A-Z]+[12])"[^>]*t="inlineStr"[^>]*><is>(.*?)</is></c>', to_shared, sheet)
    row3 = ('<c r="E3" t="b"><f>IF(B3="Marisol Venkataraman",TRUE,FALSE)</f><v>0</v></c>'
            '<c r="F3" t="str"><f>IF(A3="Open","Final","Odin Pryce pending")</f><v>Final</v></c>'
            '<c r="G3" t="inlineStr"><is><t>Yes</t><rPh sb="0" eb="1"><t>Yolanda Quist</t></rPh></is></c>')
    sheet = _re.sub(r'(<row r="3"[^>]*>.*?)(</row>)', lambda m: m.group(1) + row3 + m.group(2), sheet, flags=_re.S)
    # row 4: a phonetic run in the fixed-word columns themselves, inline (E4) and in a shared string (F4)
    sst.append('Final</t><rPh sb="0" eb="1"><t>Quentin Abara')
    row4 = ('<row r="4"><c r="A4" t="inlineStr"><is><t>Closed</t></is></c>'
            '<c r="E4" t="inlineStr"><is><t>Yes</t><rPh sb="0" eb="1"><t>Yolanda Quist</t></rPh></is></c>'
            f'<c r="F4" t="s"><v>{len(sst) - 1}</v></c></row>')
    sheet = sheet.replace("</sheetData>", row4 + "</sheetData>")
    sst_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/'
               f'spreadsheetml/2006/main" count="{len(sst)}" uniqueCount="{len(sst)}">'
               + "".join(f"<si><t>{t}</t></si>" for t in sst) + "</sst>")
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w") as z:
        for n in src.namelist():
            data = src.read(n)
            if n == "xl/worksheets/sheet1.xml":
                data = sheet.encode()
            elif n == "[Content_Types].xml" and b"sharedStrings" not in data:
                data = data.replace(b"</Types>", b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                                    b'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
            elif n == "xl/_rels/workbook.xml.rels" and b"sharedStrings" not in data:
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdSst" Type="http://schemas.openxmlformats'
                                    b'.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
                                    b"</Relationships>")
            z.writestr(n, data)
        if "xl/sharedStrings.xml" not in src.namelist():
            z.writestr("xl/sharedStrings.xml", sst_xml)
    vault = Vault()
    _, out, _ = _tokenize(out_buf.getvalue(), vault)
    xml = _xml(out)
    for secret in ("Marisol", "Venkataraman", "Odin Pryce", "Yolanda", "Quentin", "Abara"):
        assert secret not in xml, secret
    vals = _values(out, formulas=False)
    assert str(vals[("Site Work", "E2")]).startswith("TXT-") and str(vals[("Site Work", "F2")]).startswith("TXT-")
    assert str(vals[("Site Work", "G2")]).startswith("TXT-")                            # not an agent column
    assert vals.get(("Site Work", "E3")) is False                                       # cached boolean stays
    assert str(vals[("Site Work", "E4")]).startswith("TXT-") and str(vals[("Site Work", "F4")]).startswith("TXT-")
    back = _values(detokenize_file("Eddleton Boiler Compliance Tracker.xlsx", out, vault)[0], formulas=True)
    assert back[("Site Work", "F3")] == '=IF(A3="Open","Final","Odin Pryce pending")'


def test_a_project_data_workbook_keeps_nothing_readable():
    """Nov 7-12, 2026 the "Jurisdiction Override" label and a value naming only wage types stayed readable for the
    agent. Since #151 step 3 nothing of a Project Data workbook is read on AWS: every text is a token."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Project Manager", "Marisol Venkataraman"])
    ws.append(["Jurisdiction Override", "Federal, State"])
    ws.append(["Jurisdiction Override", "State - ask Odin Pryce"])
    buf = io.BytesIO()
    wb.save(buf)
    vault = Vault()
    _, out, _ = _tokenize(buf.getvalue(), vault)
    xml = _xml(out)
    for secret in ("Marisol", "Venkataraman", "Odin Pryce", "Jurisdiction Override", "Federal, State"):
        assert secret not in xml, secret
    back = _values(detokenize_file("Project Data.xlsx", out, vault)[0], formulas=False)
    assert back[("Data", "B2")] == "Federal, State" and back[("Data", "B1")] == "Marisol Venkataraman"


def test_wage_tables_with_trade_and_base_rate_keep_the_earlier_handling():
    """David, Nov 8, 2026: "Trade" and "Base Rate" headers (and the firm's wage table template) are wage tables, so
    their cells reach AWS readable, as the agent needs."""
    for rows in ([["County", " Trade", "Classification", "Total Wage", "Base Rate", "Jurisdiction", "Decision"],
                  ["King County", "LABORER", "Group 1", 54.8, 38.05, "Federal", "WA20250091"]],
                 [["Trade", "Classification", "Base Rate"], ["LABORER", "Group 1", 50.1]]):
        wb = openpyxl.Workbook()
        for r in rows:
            wb.active.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
        assert "TXT-" not in _xml(out)


def test_wage_headers_come_from_the_registry_in_both_components(agent):
    """Was a kept-in-step test; since vNext PR 2 both sides import the registry (test_filekinds.py checks the copies
    are the source)."""
    from agent import wages
    from agent.cva_filekinds import wage_table as agent_kind
    from pii_tokenizer.cva_filekinds import wage_table as our_kind
    for name in ("CRAFT_HDR", "CLASS_HDR", "BASE_HDR"):
        assert getattr(doubt, name) is getattr(our_kind, name), name
        assert getattr(wages, name) is getattr(agent_kind, name), name
    assert (doubt.CRAFT_HDR, doubt.CLASS_HDR, doubt.BASE_HDR) == (wages.CRAFT_HDR, wages.CLASS_HDR, wages.BASE_HDR)


def test_only_the_wage_columns_of_a_wage_table_stay_readable():
    """Review of PR #101: a sheet that looks like a wage table keeps readable only the columns the agent reads, so a
    roster with a "Worker", "Apprentice" or "Crew Member" column doesn't send names to AWS. (A sheet with hours or
    weeks is payroll-like and is decided as before: an exact "Classification" header there is the payroll gap #53.)"""
    cases = [
        [["Job Classification", "Base Rate", "Worker", "Notes"], ["Group 1", 40.0, "Dana Wexley", "fine"]],
        [["Trade", "Classification", "Base Rate", "Apprentice"], ["LABORER", "Group 1", 40.0, "Dana Wexley"]],
        [[None]] * 8 + [["Trade", "Job Classification", "Crew Member"], ["LABORER", "Group 1", "Dana Wexley"]],
    ]
    for rows in cases:
        wb = openpyxl.Workbook()
        for r in rows:
            wb.active.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
        xml = _xml(out)
        assert "Dana Wexley" not in xml and "Group 1" in xml, rows[-2]


def test_the_wage_table_template_keeps_its_columns():
    import datetime as dt
    wb = openpyxl.Workbook()
    wb.active.append(["County", "Trade", "Classification", "Total Wage", "Base Rate", "Fringe Rate", "Jurisdiction",
                      "Decision", "Effective Date", "Issue Date", "Expiration Date", "Comments"])
    wb.active.append(["King County", "LABORER", "Group 1", 54.8, 38.05, 16.75, "Federal", "WA20250091",
                      dt.datetime(2025, 12, 12), dt.datetime(2025, 12, 12), dt.datetime(2100, 1, 1), "ask Dana Wexley"])
    buf = io.BytesIO()
    wb.save(buf)
    _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
    vals = _values(out, formulas=False)
    sheet = next(iter({k[0] for k in vals}))
    assert [vals[(sheet, f"{c}2")] for c in "ABCDEFGH"] == ["King County", "LABORER", "Group 1", 54.8, 38.05, 16.75,
                                                             "Federal", "WA20250091"]
    assert str(vals[(sheet, "L2")]).startswith("TXT-")


def test_a_wage_table_with_an_employee_column_and_the_agents_merged_tables_keep_the_earlier_handling():
    for rows in ([["Employee Name", "Craft", "Classification", "Basic Rate"], ["Rivera, Ana", "LABORER", "G1", 50]],
                 [doubt.MERGED_TABLE_HEADER, ["federal", "LABORER", "Group 1", "WA20250091"]]):
        wb = openpyxl.Workbook()
        for r in rows:
            wb.active.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
        xml = _xml(out)
        assert "TXT-" not in xml and "Rivera" not in xml        # the employee column is tokenized by its header


def test_wage_columns_come_from_the_registry_in_both_components(agent):
    from agent import wages
    from agent.cva_filekinds import merged_wage_table as agent_merged, wage_table as agent_kind
    from pii_tokenizer.cva_filekinds import merged_wage_table as our_merged, wage_table as our_kind
    assert doubt.WAGE_HEADERS is our_kind.WAGE_HEADERS and wages.WAGE_HEADERS is agent_kind.WAGE_HEADERS
    assert wages.LCP_HEADERS is agent_kind.LCP_HEADERS and wages.TABLE_HEADERS is agent_kind.TABLE_HEADERS
    assert doubt.MERGED_TABLE_HEADER is our_merged.MERGED_TABLE_HEADER
    assert wages.FIELDS == agent_merged.MERGED_TABLE_HEADER
    assert doubt.WAGE_HEADERS == wages.WAGE_HEADERS and doubt.MERGED_TABLE_HEADER == wages.FIELDS


def test_payroll_like_sheets_are_decided_as_before():
    """Verification of PR #101: a wage-looking sheet with a personal-data column or hours/weeks columns is decided as
    before Nov 8 (never more open): here doubt mode, every unknown cell a token, the Employee column a name token."""
    cases = {"personal": [["Trade", "Job Classification", "Base Rate", "Employee", "Apprentice"],
                          ["LABORER", "Group 1", 40.0, "Rivera, Ana", "Dana Wexley"]],
             "hours": [["Craft", "Classification", "Worker", "ST Hours", "Base Hourly", "Week End Date"],
                       ["LABORER", "Group 1", "Dana Wexley", 8, 40.0, "11/01/2026"]]}
    for why, rows in cases.items():
        in_doubt, _ = doubt.sheet_rules(["S"], [[(i + 1, {j + 1: v for j, v in enumerate(r)}) for i, r in
                                                 enumerate(rows)]], 0, personal=lambda h: CFG.field_for(h) is not None)
        main_rows = [(i + 1, {j + 1: v for j, v in enumerate(r)}) for i, r in enumerate(rows)]
        before = doubt.is_payroll_sheet(main_rows) or doubt._old_is_wage_sheet(main_rows)
        assert in_doubt == (not before), why
    wb = openpyxl.Workbook()
    for r in cases["personal"]:
        wb.active.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    _, out, _ = tokenize_file("x.xlsx", buf.getvalue(), Context.build(CFG, KEY, Vault()))
    assert "Dana Wexley" not in _xml(out) and "Rivera" not in _xml(out)


def test_the_location_line_is_kept_only_when_it_names_a_county():
    for v, keep in (("Location: Pierce County", True), ("Location: King County, WA", True), ("Location: King", True),
                    ("Location: Multnomah County", True), ("Location: Dana Wexley", False),
                    ("Location: call Dana at 206-555-0142", False)):
        assert bool(doubt.LOCATION_RE.fullmatch(v)) == keep, v


def test_wa_counties_come_from_the_registry_in_both_components(agent):
    from agent import wages
    from agent.cva_filekinds import wage_table as agent_kind
    from pii_tokenizer.cva_filekinds import wage_table as our_kind
    assert doubt.WA_COUNTIES is our_kind.WA_COUNTIES and wages.WA_COUNTIES is agent_kind.WA_COUNTIES
    assert doubt.WA_COUNTIES == wages.WA_COUNTIES
