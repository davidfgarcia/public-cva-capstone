"""The checks a file must pass before it is copied to AWS (pii_tokenizer/screen.py, run by the mirror): names,
relevance of PDFs, what is left after tokenizing, PDFs' text, form fields and properties; and the held-back lists.
Every name, number and company is invented."""
from __future__ import annotations

import base64
import copy
import csv
import io
import json
import re
from pathlib import Path

import pytest

from helpers import CFG, KEY, PAYROLL, FakeDrive, ctx, make_pdf, make_xlsx
from pii_tokenizer import config as config_mod
from pii_tokenizer import screen
from pii_tokenizer.files import Context
from pii_tokenizer.mirror import Mirror
from pii_tokenizer.tokens import Vault

WD = [["General Decision Number: WA20250001", "Laborers Group 1 52.10", "Modification 3"]]


class World:
    def __init__(self):
        self.d = FakeDrive()
        self.raw, self.tok = self.d.root("staff"), self.d.root("tokenized")
        self.state = {}
        self.ctx = Context.build(CFG, KEY, Vault())

    def run(self):
        def save_vault():
            self.ctx.vault.dirty = False
        def save_state():                     # what the job keeps if the run stops right after
            self.saved = copy.deepcopy(self.state)
        return Mirror(self.d, self.ctx, self.state, save_vault, save_state).run("compliance", self.raw, self.tok)

    def crash_and_restart(self):              # the run stopped: only the last saved state is left
        self.state = copy.deepcopy(self.saved)

    def sheet(self):
        data = self.d.read(self.raw, screen.HELD_BACK_SHEET)
        return None if data is None else list(csv.reader(io.StringIO(data.decode("utf-8-sig"))))

    def summary(self):
        return json.loads(self.d.read(self.tok, screen.HELD_BACK_SUMMARY))["held_back"]

    def payroll(self):            # teaches the tokenizer the workers Ana Rivera and Lee Mason
        self.d.put(self.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(PAYROLL))


def held(w):
    return {row[0]: row[1] for row in (w.sheet() or [])[1:]}


def aws_paths(w):
    files, folders = w.d.tree(w.tok)
    return [f["rel"] for f in files.values()] + list(folders)


def no_name_on_aws(w, *names):
    low = " ".join(aws_paths(w)).casefold()
    return not any(n.casefold() in low for n in names or ("rivera", "mason"))


def tok_path(w, prefix):
    """The one AWS-side file whose path starts with ``prefix``."""
    hits = [f["rel"] for f in w.d.tree(w.tok)[0].values() if f["rel"].startswith(prefix)]
    assert len(hits) == 1, hits
    return hits[0]


def test_a_worker_name_in_a_file_name_is_replaced_by_its_token_on_aws():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera timesheet.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    rel = tok_path(w, "Projects/Brennick - 77Q/Uploads/EMP-")
    assert re.fullmatch(r"Projects/Brennick - 77Q/Uploads/EMP-[0-9a-f]+ timesheet\.xlsx", rel), rel
    assert no_name_on_aws(w)
    assert set(held(w)) <= {"(nothing is held back)"}
    w.run()                                                     # stable: nothing renamed or copied again
    assert tok_path(w, "Projects/Brennick - 77Q/Uploads/EMP-") == rel
    assert w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera timesheet.xlsx")    # staff side untouched


def test_contact_details_and_ssns_in_names_are_held_back():
    w = World()
    for name in ("call 206-555-0100.xlsx", "from bob@example.com.csv", "SSN 123-45-6789.xlsx"):
        w.d.put(w.raw, f"Projects/P/Uploads/{name}", make_xlsx([["Hours"], [8]]))
    w.run()
    reasons = held(w)
    assert "a phone number" in reasons["Projects/P/Uploads/call 206-555-0100.xlsx"]
    assert "an email address" in reasons["Projects/P/Uploads/from bob@example.com.csv"]
    assert "an SSN" in reasons["Projects/P/Uploads/SSN 123-45-6789.xlsx"]


def _summary_text(w):
    return w.d.read(w.tok, screen.HELD_BACK_SUMMARY).decode()


def test_a_held_back_tracker_is_named_by_project_in_the_aws_summary_and_nothing_else():
    """The outcome emails (written on AWS) say "the tracker was held back" instead of "no tracker": the summary
    carries a fixed word and the project folder, never the file's name or why it was held back."""
    w = World()
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Tracker/call 206-555-0100.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Uploads/call 206-555-0199.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Tracker/Archived/call 206-555-0111.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Tracker/call 206-555-0122.pdf", b"%PDF-1.4 x")
    w.run()
    entries = w.summary()
    trackers = [e for e in entries if e.get("kind") == "tracker"]
    assert [(e["kind"], e["project"]) for e in trackers] == [("tracker", "Corvath - Nimbrel")]
    assert len(entries) >= 3 and all(set(e) <= {"key", "since", "kind", "project"} for e in entries)
    text = _summary_text(w)
    assert "206-555" not in text and "call " not in text and "phone" not in text      # no name, no reason


def test_the_project_in_the_summary_is_the_folder_name_exactly_as_aws_has_it():
    """Project folders keep their names on AWS (screen.tokenize_path: projects are often named after people, and the
    name is business information), so the summary adds no name AWS doesn't already have; it takes the name from the
    same function, so it follows if that rule changes. A worker's name in a FILE name is never in the summary."""
    w = World()
    w.payroll()                                    # teaches the tokenizer Ana Rivera and Lee Mason
    w.d.put(w.raw, "Projects/Cesar Chavez Park/Uploads/hours.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Cesar Chavez Park/Tracker/call 206-555-0100.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Cesar Chavez Park/Tracker/Lee Mason 206-555-0111.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    trackers = [e for e in w.summary() if e.get("kind") == "tracker"]
    assert [e["project"] for e in trackers] == ["Cesar Chavez Park"] * 2
    assert any(p.startswith("Projects/Cesar Chavez Park/") for p in aws_paths(w))     # the same name AWS has
    assert "Mason" not in _summary_text(w)


def test_a_project_folder_whose_own_name_holds_contact_details_gets_no_entry_and_its_name_is_not_sent():
    """Such a folder is held back whole and is not on AWS; its name must not reach AWS through the summary."""
    w = World()
    w.d.put(w.raw, "Projects/Site 206-555-0100/Tracker/t.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Site bob@example.com/Tracker/t.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Site 123-45-6789/Tracker/t.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    text = _summary_text(w)
    assert not [e for e in w.summary() if "project" in e or "kind" in e]
    assert "206-555" not in text and "example.com" not in text and "123-45" not in text and "Site" not in text


def test_only_a_workbook_directly_in_tracker_counts_as_a_tracker_in_the_summary():
    w = World()
    refused = make_xlsx([["Hours"], [8]], extra_parts={"xl/unknown.xml": "<x/>"})
    w.d.put(w.raw, "Projects/P1/Uploads/book.xlsx", refused)
    w.d.put(w.raw, "Projects/P3/Tracker/Archived/old.xlsx", refused)
    w.d.put(w.raw, "Projects/P4/Tracker/book.xlsx", refused)
    w.run()
    assert len(w.summary()) == 3                                  # all three are held back
    assert [e["project"] for e in w.summary() if e.get("kind") == "tracker"] == ["P4"]
    # a file directly in Uploads/ is an upload (David, Nov 12, 2026: said on that project's payroll review)
    assert [e["project"] for e in w.summary() if e.get("kind") == "upload"] == ["P1"]


def test_a_held_back_upload_is_named_by_project_only():
    w = World()
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Uploads/call 206-555-0199.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Uploads/old/call 206-555-0188.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    entries = w.summary()
    assert [(e.get("kind"), e.get("project")) for e in entries if "kind" in e] == [("upload", "Corvath - Nimbrel")]
    assert all(set(e) <= {"key", "since", "kind", "project"} for e in entries)
    text = _summary_text(w)
    assert "206-555" not in text and "call " not in text and "old" not in text


def test_a_new_version_of_a_copied_tracker_that_is_held_back_removes_the_old_copy_and_is_reported():
    w = World()
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Tracker/t.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    assert any(p.endswith("Tracker/t.xlsx") for p in aws_paths(w))
    w.d.put(w.raw, "Projects/Corvath - Nimbrel/Tracker/t.xlsx",
            make_xlsx([["Hours"], [9]], extra_parts={"xl/unknown.xml": "<x/>"}))
    w.run()
    assert not any(p.endswith("Tracker/t.xlsx") for p in aws_paths(w))
    assert [e["project"] for e in w.summary() if e.get("kind") == "tracker"] == ["Corvath - Nimbrel"]


def test_a_folder_named_after_a_worker_is_made_on_aws_with_the_token():
    w = World()
    w.payroll()
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Unused/Lee Mason/notes.txt", b"hours checked")
    w.run()
    rel = tok_path(w, "Projects/Brennick - 77Q/Files/Unused/EMP-")
    assert rel.endswith("/notes.txt") and w.d.read(w.tok, rel) == b"hours checked"
    assert no_name_on_aws(w)


def test_a_copied_file_renamed_after_a_worker_is_renamed_on_aws_with_the_token():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/hours.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    fid = w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/hours.xlsx")
    w.d.items[fid]["name"] = "Lee Mason hours.xlsx"
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/hours.xlsx") is None
    assert tok_path(w, "Projects/Brennick - 77Q/Uploads/EMP-").endswith(" hours.xlsx")
    assert no_name_on_aws(w) and set(held(w)) <= {"(nothing is held back)"}


# ------------------------------------------------------------------------------------------------ PDFs
def test_pdfs_reach_aws_only_when_used_by_the_agent_and_free_of_worker_data():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf(WD))
    w.d.put(w.raw, "reference/knowledge/cba/Laborers CBA.pdf",
            make_pdf([["Agreement between Local 242 and contractors", "Union rep: Pat Quinn, 206-555-0177"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/certified payroll.pdf",
            make_pdf([["Statement of compliance", "Hours 40 regular"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WD with worker.pdf",
            make_pdf([WD[0] + ["Note: Ana Rivera classified as Group 1"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WD ssn.pdf", make_pdf([WD[0] + ["SSN 123456789"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/scanned WD.pdf", make_pdf(WD, image_only_page=True))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WD author.pdf", make_pdf(WD, metadata={"/Author": "Lee Mason"}))
    w.d.put(w.raw, "notes/letter.pdf", make_pdf([["General Decision Number: WA20250001"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/broken.pdf", b"%PDF-1.4 not really")
    w.run()
    sent = lambda rel: w.d.read(w.tok, rel) is not None
    assert sent("Projects/Brennick - 77Q/Uploads/WA20250001.pdf")
    assert sent("reference/knowledge/cba/Laborers CBA.pdf")      # officials' names and phones: not worker data
    reasons = held(w)
    assert "only wage determinations go to AWS as PDFs" in reasons["Projects/Brennick - 77Q/Uploads/certified payroll.pdf"]
    assert "contains a worker's name" in reasons["Projects/Brennick - 77Q/Uploads/WD with worker.pdf"]
    assert "only wage determinations" in reasons["Projects/Brennick - 77Q/Uploads/WD ssn.pdf"]   # "SSN" marks a payroll
    assert "scan or a picture" in reasons["Projects/Brennick - 77Q/Uploads/scanned WD.pdf"]
    assert "contains a worker's name" in reasons["Projects/Brennick - 77Q/Uploads/WD author.pdf"]
    assert "doesn't read PDFs here" in reasons["notes/letter.pdf"]
    assert "can't be read as a PDF" in reasons["Projects/Brennick - 77Q/Uploads/broken.pdf"]
    for rel in reasons:
        assert not sent(rel), rel


def test_a_pdf_already_on_aws_is_taken_back_when_a_new_worker_name_appears_in_it():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()                                       # no payroll yet: the name is unknown, the PDF is copied
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf") is not None
    w.payroll()                                   # the tokenizer learns Ana Rivera
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf") is None
    assert "contains a worker's name" in held(w)["Projects/Brennick - 77Q/Uploads/WA20250001.pdf"]
    w.run()                                       # stays held back, without being checked again every run
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf") is None


# ------------------------------------------------------------------------------------------------ 4a
def test_worker_data_left_after_tokenizing_holds_a_workbook_back(monkeypatch):
    c = ctx()
    c.scrubber.add_name("Rivera, Ana", "EMP-0000000001")
    leaky = make_xlsx([["Contractor", "Notes"], ["Acme Paving", "Ana Rivera"]])
    assert "a worker's name was still in it" in screen.residual("t.xlsx", leaky, c.scrubber)
    assert screen.residual("t.xlsx", make_xlsx([["Contractor"], ["Acme Paving"]]), c.scrubber) is None
    assert "an SSN" in screen.residual("t.csv", b"id,ssn\n1,123-45-6789\n", c.scrubber)
    # in the mirror: a tokenizer that missed something (simulated) never gets its result to AWS
    w = World()
    w.payroll()
    w.run()
    import pii_tokenizer.mirror as mirror_mod
    monkeypatch.setattr(mirror_mod, "tokenize_file", lambda name, data, ctx, google_sheet=False, rel=None: (name, data, "x"))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Tracker/tracker.xlsx", leaky)
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Tracker/tracker.xlsx") is None
    assert "still in it after tokenizing" in held(w)["Projects/Brennick - 77Q/Tracker/tracker.xlsx"]


def test_payrolls_pass_their_own_check():
    w = World()
    w.payroll()
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/week1.xlsx") is not None
    assert w.sheet() is None                      # nothing held back: no list in the staff folder


def test_the_held_back_files_are_never_copied_either_way():
    w = World()
    w.d.put(w.raw, "Projects/P/Uploads/a 206-555-0100.xlsx", make_xlsx([["x"], [1]]))
    w.run()
    w.run()
    assert w.d.path(w.tok, screen.HELD_BACK_SHEET) is None
    assert w.d.path(w.raw, screen.HELD_BACK_SUMMARY) is None
    key = w.summary()[0]["key"]
    fid = w.d.path(w.raw, "Projects/P/Uploads/a 206-555-0100.xlsx")
    w.d.items[fid]["data"] = make_xlsx([["x"], [2]])       # changed (still held back): reported again as new
    w.run()
    assert w.summary()[0]["key"] != key


# ------------------------------------------------------------------------------------------------ pieces
def test_finds_changes_nothing():
    c = ctx()
    c.scrubber.add_name("Rivera, Ana", "EMP-0000000001")
    before = dict(c.vault.entries)
    assert c.scrubber.finds("Ana Rivera called") == "a worker's name"
    assert c.scrubber.finds("call 206-555-0100") is None
    assert c.scrubber.finds("call 206-555-0100", contacts=True) == "a phone number"
    assert c.scrubber.finds("Social Security No: 123456789") == "an SSN"
    assert c.scrubber.finds("Permit 123456789") is None
    assert c.vault.entries == before


def test_copy_as_is_is_no_longer_accepted(tmp_path):
    cfg = Path(config_mod.DEFAULT_PATH).read_text().replace("scan_pdf: ['.pdf']", "copy_as_is: ['.pdf']")
    p = tmp_path / "c.yaml"
    p.write_text(cfg)
    with pytest.raises(ValueError, match="copy_as_is"):
        config_mod.load(p)


def test_review_payrolls_are_not_held_back_by_their_own_ssn_tokens():
    """A token such as SSN-e412240181 is not an SSN: before this was fixed, 1 in 80 tokens looked like "SSN" + 9
    digits, which held back most large payrolls."""
    rows = [PAYROLL[1]]
    for i in range(400):
        ssn = f"{100 + i % 800:03d}-{10 + i % 89:02d}-{1000 + i:04d}"
        rows.append(["P07", "Acme Paving", 46000, f"Worker{i:03d}, Test", 20000 + i, ssn, "F", "Laborer", 50, None])
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/big.xlsx", make_xlsx(rows))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/big.csv", "\n".join(",".join(map(str, r)) for r in rows).encode())
    w.run()
    assert w.sheet() is None, w.sheet()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/big.xlsx") and w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/big.csv")


@pytest.mark.parametrize("name", ["Ana_Rivera.xlsx", "Rivera_Ana.xlsx", "Rivera Ana.xlsx", "ana.rivera hours.xlsx",
                                  "RiveraAna.xlsx", "Rivera,Ana.xlsx", "Ana-Rivera wk42.xlsx"])
def test_worker_names_are_replaced_in_any_common_spelling(name):
    w = World()
    w.payroll()
    w.run()
    w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/{name}", make_xlsx([["Hours"], [8]]))
    w.run()
    assert no_name_on_aws(w, "rivera"), aws_paths(w)
    assert len([r for r in aws_paths(w) if "/Uploads/" in r and "EMP-" in r]) == 1


def test_a_certified_payroll_pdf_is_held_back_even_with_wage_words():
    w = World()
    wh347 = make_pdf([["U.S. Department of Labor", "PAYROLL (WH-347)", "Prevailing Wage Rates apply",
                       "Name of worker: Sam Okafor  Hours 40"]])
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/prevailing wage payroll DOL.pdf", wh347)
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/prevailing wage payroll DOL.pdf") is None
    assert "only wage determinations" in held(w)["Projects/Brennick - 77Q/Uploads/prevailing wage payroll DOL.pdf"]


def test_each_held_back_pdf_reason_fits_the_file():
    """#64: a wage file that isn't recognised must not be told to "upload payrolls"."""
    from pii_tokenizer.screen import pdf_relevant
    rel = "Projects/Brennick - 77Q/Uploads/x.pdf"
    assert pdf_relevant(rel, "General Decision Number: WA20250001") is None
    # payroll words only: today's advice, naming what decided it with a fixed label
    only_payroll = pdf_relevant(rel, "Statement   of\nCompliance  Hours 40")
    assert "looks like a payroll (it mentions a Statement of Compliance)" in only_payroll
    assert "upload payrolls as .xlsx or .csv" in only_payroll and "also mentions" not in only_payroll
    # both: the payroll details decide, and the reason says so
    both = pdf_relevant(rel, "PAYROLL (WH-347)  Prevailing Wage Rates apply")
    assert "a WH-347 form" in both and "also mentions wage rates, but files with payroll details" in both
    # neither (a garbled scan layer, an Oregon book, an LCP report saved as PDF): no payroll advice
    for first in ("Oregon BOLI rate book  Group 3", "LCPtracker report  Craft  Rate", ""):
        why = pdf_relevant(rel, first)
        assert "isn't recognised as a wage determination" in why and "LCPtracker Excel export" in why
        assert "payroll" not in why.lower()
    # every kind of payroll word has its own label
    for text, label in (("WH 347", "a WH-347 form"), ("Certified Payroll", "a certified payroll"),
                        ("Payroll No 4", "a payroll number"), ("SSN", "Social Security numbers"),
                        ("Social Security", "Social Security numbers"), ("Employee ID", "employee details"),
                        ("Gross Amount", "pay amounts"), ("Net Wages", "pay amounts")):
        assert f"it mentions {label})" in pdf_relevant(rel, text), text


def test_a_held_back_reason_never_quotes_the_document():
    """The words that matched can sit inside a worker's name ("Bennet Payne" holds "net Pay")."""
    from pii_tokenizer.screen import pdf_relevant
    rel = "Projects/Brennick - 77Q/Uploads/x.pdf"
    for text, label in (("Wage Determination  Bennet Payne  Group 1", "pay amounts"),
                        ("Prevailing Wage Rates  Employee Idaho Falls", "employee details")):
        why = pdf_relevant(rel, text)
        assert label in why
        assert "Payne" not in why and "net Pay" not in why and "Id" not in why.replace("Idaho", "")


def test_an_unrecognised_wage_pdf_is_listed_with_the_right_advice():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/City wages in LCP.pdf",
            make_pdf([["LCPtracker report", "Craft  Laborer  52.10"]]))
    w.run()
    why = held(w)["Projects/Brennick - 77Q/Uploads/City wages in LCP.pdf"]
    assert "isn't recognised as a wage determination" in why and "upload payrolls" not in why


def test_pdf_notes_bookmarks_xmp_inline_images_and_old_versions():
    w = World()
    w.payroll()
    w.run()
    cases = {"note.pdf": dict(note="ok"), "bookmark.pdf": dict(outline="Lee Mason"), "xmp.pdf": dict(xmp="Ana Rivera"),
             "inline.pdf": dict(inline_image_page=True), "edited.pdf": dict(edited=True)}
    for name, kw in cases.items():
        w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/{name}", make_pdf(WD, **kw))
    w.run()
    reasons = held(w)
    expect = {"note.pdf": "notes or other annotations", "bookmark.pdf": "a worker's name", "xmp.pdf": "a worker's name",
              "inline.pdf": "scan or a picture", "edited.pdf": "earlier versions inside it"}
    for name, why in expect.items():
        assert why in reasons[f"Projects/Brennick - 77Q/Uploads/{name}"], (name, reasons)
        assert w.d.read(w.tok, f"Projects/Brennick - 77Q/Uploads/{name}") is None


def test_a_folder_named_after_a_worker_learnt_in_the_same_run_is_not_made():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt", b"hours checked")
    w.d.mkdir(w.raw, "Lee Mason")                             # an empty one too
    w.run()
    w.run()
    assert w.d.path(w.tok, "Projects/Brennick - 77Q/Files/Ana Rivera") is None
    assert w.d.path(w.tok, "Lee Mason") is None


def test_a_folder_already_on_aws_is_renamed_once_the_name_is_learnt():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt", b"hours checked")
    w.run()                                                   # name unknown yet: copied as it is
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt") is not None
    w.payroll()
    rep = w.run()
    rel = tok_path(w, "Projects/Brennick - 77Q/Files/EMP-")
    assert rel.endswith("/notes.txt") and w.d.read(w.tok, rel) == b"hours checked"
    assert no_name_on_aws(w)
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt") == b"hours checked"   # no duplicate
    assert [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith("notes.txt")] == \
        ["Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt"]
    w.d.put(w.tok, rel.rsplit("/", 1)[0] + "/review.txt", b"EMP-x reviewed")                 # AWS adds a file there
    w.run()
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/review.txt") is not None    # in the staff folder


def test_a_file_copied_before_a_later_batch_taught_the_name_is_taken_back_in_the_same_run(monkeypatch):
    import pii_tokenizer.mirror as mirror_mod
    monkeypatch.setattr(mirror_mod, "BATCH", 1)
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv", b"Hours\n8\n")
    w.payroll()
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv") is None


def test_utf16_and_windows_text_is_checked_after_tokenizing():
    c = ctx()
    c.scrubber.add_name("Garcia, Jose", "EMP-0000000002")
    assert screen.residual("t.txt", "\ufeffSSN 123-45-6789".encode("utf-16"), c.scrubber)
    assert screen.residual("t.txt", "Notes: José García".encode("cp1252"), c.scrubber) is None    # not learnt with accents
    c.scrubber.add_name("García, José", "EMP-0000000003")
    assert "a worker's name" in screen.residual("t.txt", "Notes: José García".encode("cp1252"), c.scrubber)


def test_the_same_pdf_in_a_place_the_agent_doesnt_read_is_not_sent_by_reuse():
    w = World()
    cba = make_pdf([["Agreement between Local 242 and contractors"]])
    w.d.put(w.raw, "reference/knowledge/CBA.pdf", cba)
    w.d.put(w.raw, "scans/CBA copy.pdf", cba)            # after reference/ in the run's order, where the agent doesn't read
    w.run()
    assert w.d.read(w.tok, "reference/knowledge/CBA.pdf") is not None
    assert w.d.read(w.tok, "scans/CBA copy.pdf") is None


def test_the_list_never_runs_a_formula_and_copies_of_it_are_never_sent():
    w = World()
    w.d.put(w.raw, "=HYPERLINK(1) 206-555-0100.xlsx", make_xlsx([["x"], [1]]))
    w.d.put(w.raw, "Held back from AWS (1).csv", b"File,Why\nAna Rivera.xlsx,x\n")
    w.run()
    assert held(w)["'=HYPERLINK(1) 206-555-0100.xlsx"]
    assert w.d.path(w.tok, "Held back from AWS (1).csv") is None


# ------------------------------------------------------------------------------------------------ take-backs
def test_projects_and_places_named_after_people_are_not_taken_off_aws():
    """A worker named like a project or a place must not take the project off AWS."""
    w = World()
    w.d.put(w.raw, "Projects/Lee Mason/Uploads/hours.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Lee Mason Park - Phase 2/Files/notes.txt", b"site walk")
    w.run()
    w.payroll()                                   # a worker called Lee Mason turns up
    w.run()
    assert w.d.read(w.tok, "Projects/Lee Mason/Uploads/hours.xlsx") is not None
    assert w.d.read(w.tok, "Projects/Lee Mason Park - Phase 2/Files/notes.txt") is not None
    w.d.mkdir(w.d.path(w.raw, "Projects"), "Ana Rivera")      # a new project, named like a known worker
    w.run()
    assert w.d.path(w.tok, "Projects/Ana Rivera") is not None
    assert set(held(w)) <= {"(nothing is held back)"}


def test_one_run_takes_back_only_a_few_copies():
    w = World()
    for i in range(12):
        w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/WD {i}.pdf", make_pdf([WD[0] + [f"Ana Rivera, laborer {i}"]]))
    w.run()                                       # name not known yet: all copied
    w.payroll()
    rep = w.run()
    on_aws = lambda: sum(w.d.read(w.tok, f"Projects/Brennick - 77Q/Uploads/WD {i}.pdf") is not None for i in range(12))
    assert on_aws() == 7                          # 5 taken back (at least 5, or a tenth of the copies)
    assert any("Held back from AWS.csv" in x for x in rep.warnings)
    w.run(); w.run()
    assert on_aws() == 0
    assert all(w.d.path(w.raw, f"Projects/Brennick - 77Q/Uploads/WD {i}.pdf") for i in range(12))  # untouched


@pytest.mark.parametrize("stop_in", ["_tidy_layout", "_publish_held_back"])
def test_a_take_back_survives_a_run_that_stops_later(monkeypatch, stop_in):
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    for st in (w.state, w.saved):                              # copied days ago (not a copy an AWS run just missed)
        for pair in st["mirrors"]["compliance"]["pairs"]:
            pair["copied_utc"] = "2020-01-01-00-00-00"
    w.payroll()                                                # the PDF now holds a known name: taken back
    monkeypatch.setattr(Mirror, stop_in, lambda *a: 1 / 0)
    with pytest.raises(ZeroDivisionError):
        w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None   # taken back, then the run stopped
    w.crash_and_restart()
    monkeypatch.undo()
    w.run()
    # not mistaken for a file deleted on AWS: the staff file stays where it is, and is listed as held back
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf") is not None
    assert w.d.path(w.raw, "_sync_backup") is None
    assert "Projects/Brennick - 77Q/Uploads/WA.pdf" in held(w)


def test_a_failed_take_back_is_tried_again_next_run():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    w.payroll()
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    rep = w.run()
    assert any("could not be taken back" in x for x in rep.warnings)
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is not None
    w.d.trash = real
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert "Projects/Brennick - 77Q/Uploads/WA.pdf" in held(w)


def test_a_taken_back_copy_uploaded_again_from_aws_is_not_restored_to_staff():
    """An AWS run that pulled the copy before it was taken back pushes it again: it must not come back to the staff
    folder as a second file (and is removed from the AWS side again)."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    w.run()
    pdf = w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf")
    w.payroll()
    w.run()                                                    # both taken back
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", pdf)            # AWS pushes them again
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001 (1).pdf", pdf)        # even under another name
    rep = w.run()
    staff = [r for r in w.d.tree(w.raw)[0].values() if r["rel"].startswith("Projects/Brennick - 77Q/Uploads/")]
    assert sorted(r["rel"].rsplit("/", 1)[1] for r in staff) == ["Lee Mason note.txt", "WA20250001.pdf", "week1.xlsx"]
    for name in ("WA20250001.pdf", "Lee Mason note.txt", "WA20250001 (1).pdf"):
        assert w.d.read(w.tok, f"Projects/Brennick - 77Q/Uploads/{name}") is None, name
    assert sum("uploaded again from AWS" in d for _, d in rep.removed) == 3


def test_union_rate_sheets_and_fillable_forms_are_wage_pdfs():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Local 242 rates.pdf",
            make_pdf([["Laborers Local 242", "Wage and Fringe Benefit Rates, effective June 1", "Group 1 52.10"]]))
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/Local 242 rates.pdf") is not None


def test_pdfs_left_unchecked_when_time_runs_out_are_checked_next_run(monkeypatch):
    import pii_tokenizer.mirror as mirror_mod
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    w.payroll()
    monkeypatch.setattr(mirror_mod, "RECHECK_SECONDS", -1)
    rep = w.run()
    assert any("not yet checked again" in x for x in rep.warnings)
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf") is not None
    monkeypatch.undo()
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf") is None


def test_folders_below_a_project_get_the_token_for_a_worker_name():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera timesheets/w1.csv", b"Hours\n8\n")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Rivera, Ana - docs/w2.csv", b"Hours\n8\n")
    w.run()
    assert no_name_on_aws(w)
    files = sorted(r for r in aws_paths(w) if r.endswith(".csv"))
    assert len(files) == 2 and all("/Files/EMP-" in r for r in files), files
    assert set(held(w)) <= {"(nothing is held back)"}


def test_a_copy_with_a_worker_name_pushed_again_from_aws_is_not_restored():
    """After the copy was renamed with the token, an AWS run that pulled it before may push the old name again:
    it is not restored to the staff folder as a second file, and is removed from the AWS side."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    w.run()                                                   # name unknown: copied as it is
    w.payroll()
    w.run()                                                   # renamed on AWS with the token
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")   # the old name pushed again
    rep = w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt") is None
    staff = [f["rel"] for f in w.d.tree(w.raw)[0].values() if "note" in f["rel"]]
    assert staff == ["Projects/Brennick - 77Q/Uploads/Lee Mason note.txt"]
    assert any("uploaded again from AWS" in d for _, d in rep.removed)


def test_an_agent_file_identical_to_a_held_back_staff_file_is_kept():
    w = World()
    w.payroll()
    pdf = make_pdf(WD)
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", pdf)
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/WA.pdf", pdf)          # held back (never copied)
    w.run()
    w.d.put(w.tok, "Projects/Brennick - 77Q/Files/WA.pdf", pdf)                     # the agent writes the same bytes
    w.run()
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/WA.pdf") is not None


def test_a_mirror_whose_copies_were_all_taken_back_still_runs():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera note.txt", b"checked")
    w.run()                                                    # copied (name unknown yet)
    w.state["mirrors"]["compliance"]["pairs"].clear()          # as after taking every copy back
    w.d.put(w.tok, "Projects/Brennick - 77Q/Files/review.csv", b"ok\n")            # the agent adds a file
    w.run()                                                    # not refused as "must be empty the first time"
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/review.csv") is not None
    w2 = World()
    w2.d.put(w2.tok, "stray.csv", b"x\n")
    with pytest.raises(RuntimeError, match="must be empty"):   # a real first run still is
        w2.run()


def test_pdfs_are_checked_again_least_recently_checked_first(monkeypatch):
    """With time for only one PDF per run and new names every run, every PDF still gets its turn."""
    import pii_tokenizer.mirror as mirror_mod
    w = World()
    for n in ("A", "B"):
        w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/WD {n}.pdf", make_pdf(WD))
    w.run()
    clock = iter(range(10**6))
    monkeypatch.setattr(mirror_mod, "time", type("T", (), {"monotonic": staticmethod(lambda: next(clock))}))
    monkeypatch.setattr(mirror_mod, "RECHECK_SECONDS", 1.5)      # start 0, first PDF at 1, the next one at 2: later
    for i, (last, first) in enumerate([("Quinn", "Pat"), ("Okafor", "Chidi")]):
        rows = PAYROLL[:2] + [["P07", "Acme Paving", 46000, f"{last}, {first}", 20000 + i, "", "", "Laborer", 50, ""]]
        w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/pay{i}.xlsx", make_xlsx(rows))
        w.run()
    pdfs = [p for p in w.state["mirrors"]["compliance"]["pairs"] if p["rel"].endswith(".pdf")]
    assert all(p["names_at_check"] > 0 for p in pdfs)


def test_a_failed_bin_call_for_a_changed_file_held_back_doesnt_stop_the_run():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))  # changed
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/other.csv", b"Hours\n9\n")
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    rep = w.run()
    assert any("could not be taken back" in x for x in rep.warnings)
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/other.csv") is not None             # the rest went on
    w.d.trash = real
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert "contains a worker's name" in held(w)["Projects/Brennick - 77Q/Uploads/WA.pdf"]


def test_a_mirror_refuses_a_different_tokenized_folder_even_with_no_copies():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera note.txt", b"checked")
    w.payroll()
    w.d.items[w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx")]["trashed"] = True
    w.run()
    w.state["mirrors"]["compliance"]["pairs"].clear()          # e.g. everything held back or taken back
    other = w.d.root("another mirror's tokenized folder")
    w.d.put(other, "Projects/X/Uploads/data.csv", b"[TOKEN]\n")
    w.tok = other
    with pytest.raises(RuntimeError, match="differs"):
        w.run()
    assert w.d.path(w.raw, "Projects/X") is None


def test_a_failed_bin_call_doesnt_make_conflict_copies_and_is_listed():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Modification 4"]]))  # AWS changed it too
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    w.run(); w.run(); w.run()
    staff = [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith(".pdf")]
    assert len(staff) <= 2, staff                             # at most the one conflict copy, not one per run
    assert "contains a worker's name" in held(w)["Projects/Brennick - 77Q/Uploads/WA.pdf"]
    w.d.trash = real
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert "Projects/Brennick - 77Q/Uploads/WA.pdf" in held(w)


def test_common_content_elsewhere_is_not_mistaken_for_a_taken_back_copy():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/blank.txt", b"")
    w.run()
    w.d.items[w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/blank.txt")]["name"] = "Ana Rivera.txt"
    w.run()                                                    # taken back
    w.d.put(w.tok, "Projects/Other/Files/empty.txt", b"")      # an unrelated empty file from the agent
    w.run()
    assert w.d.read(w.raw, "Projects/Other/Files/empty.txt") == b""


def test_a_taken_back_copy_is_still_recognised_after_a_failed_retry():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    pdf = w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.payroll()
    w.run()                                                    # taken back
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, foreman"]]))
    real = w.d.download
    w.d.download = lambda f: (_ for _ in ()).throw(OSError("timeout")) if f["name"] == "WA.pdf" else real(f)
    w.run()                                                    # edited; its check fails for now
    w.d.download = real
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf", pdf)  # AWS pushes the old copy again
    w.run()
    staff = [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith(".pdf")]
    assert staff == ["Projects/Brennick - 77Q/Uploads/WA.pdf"], staff


def test_a_file_fixed_while_its_take_back_is_pending_is_copied_again():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    w.run()                                                    # held back; old copy not binned
    w.d.trash = real
    fixed = make_pdf([WD[0] + ["Modification 4"]])
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", fixed)              # staff remove the name
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") == fixed
    assert "Projects/Brennick - 77Q/Uploads/WA.pdf" not in held(w)


def test_state_from_before_roots_were_remembered_is_not_a_first_run():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/note.txt", b"checked")
    w.run()
    st = w.state["mirrors"]["compliance"]
    st["pairs"].clear(); st.pop("roots")                       # as saved by the previous version, all taken back
    w.d.put(w.tok, "_sync_backup/old.txt", b"x")
    w.run()
    assert st["roots"] == [w.raw, w.tok]


# ------------------------------------------------------------------------------------------------ stopped runs
@pytest.mark.parametrize("first_run", [True, False])
def test_a_run_stopped_while_writing_copies_leaves_no_duplicates(monkeypatch, first_run):
    import pii_tokenizer.mirror as mirror_mod
    monkeypatch.setattr(mirror_mod, "BATCH", 3)
    w = World()
    if not first_run:
        w.payroll()
        w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/old.csv", b"Hours\n1\n")
        w.run()
        w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/old.csv", b"Hours\n2\n")      # changed: its copy is updated
    for i in range(7):
        w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/f{i}.csv", f"Hours\n{i}\n".encode())
    real, n = w.d.create, {"made": 0}

    def create(*a):
        if n["made"] == 4:
            raise MemoryError("killed")                        # as the job being stopped half-way through a batch
        n["made"] += 1
        return real(*a)
    w.d.create = create
    with pytest.raises(MemoryError):
        w.run()
    w.crash_and_restart()
    w.d.create = real
    rep = w.run()
    w.run()
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith(".csv"))
    expected = sorted([f"Projects/Brennick - 77Q/Uploads/f{i}.csv" for i in range(7)] +
                      ([] if first_run else ["Projects/Brennick - 77Q/Uploads/old.csv"]))
    assert staff == expected                                   # nothing came back from AWS as a second file
    aws = sorted(f["rel"] for f in w.d.tree(w.tok)[0].values() if f["rel"].endswith(".csv"))
    assert aws == expected                                     # and each copy is there once
    if not first_run:
        assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/old.csv") == b"Hours\n2\n"


def test_identical_contents_are_kept_in_memory_only_while_needed(monkeypatch):
    import pii_tokenizer.mirror as mirror_mod
    monkeypatch.setattr(mirror_mod, "BATCH", 2)
    seen = []
    real = mirror_mod.Mirror._tokenize

    def spy(self, r):
        out = real(self, r)
        seen.append(r["name"])
        return out
    monkeypatch.setattr(mirror_mod.Mirror, "_tokenize", spy)
    w = World()
    for i in range(6):
        w.d.put(w.raw, f"Projects/P/Uploads/u{i}.csv", f"Hours\n{i}\n".encode())
    w.d.put(w.raw, "Projects/P/Uploads/dup1.csv", b"Hours\n0\n")
    w.d.put(w.raw, "Projects/P/Files/dup2.csv", b"Hours\n0\n")
    w.run()
    assert len(seen) == 6                     # u0, dup1 and dup2 are identical: read once, in different batches
    assert all(w.d.read(w.tok, rel) == b"Hours\n0\n" for rel in ("Projects/P/Uploads/dup1.csv",
                                                                "Projects/P/Files/dup2.csv"))


# ------------------------------------------------------------------------------------------------ large files
def test_a_workbook_too_large_to_check_is_held_back_not_a_crash(monkeypatch):
    from pii_tokenizer import xlsx
    monkeypatch.setattr(xlsx, "MAX_PART_MB", 0)
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/tracker.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/notes.csv", b"Hours\n8\n")
    w.run()
    assert "too large to check" in held(w)["Projects/Brennick - 77Q/Uploads/tracker.xlsx"]
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/notes.csv") is not None          # the run went on


def test_a_pdf_too_large_to_check_is_held_back(monkeypatch):
    monkeypatch.setattr(screen, "MAX_PDF_MB", 0)
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    assert "larger than the 0 MB" in held(w)["Projects/Brennick - 77Q/Uploads/WA.pdf"]


def test_rows_out_of_order_give_the_same_result_as_in_order():
    import zipfile
    from pii_tokenizer.files import tokenize_file
    data = make_xlsx(PAYROLL)
    z = zipfile.ZipFile(io.BytesIO(data))
    parts = {n: z.read(n) for n in z.namelist()}
    x = parts["xl/worksheets/sheet1.xml"].decode()
    rows = re.findall(r"<row\b.*?</row>", x, re.S)
    parts["xl/worksheets/sheet1.xml"] = x.replace("".join(rows), "".join([rows[0]] + rows[1:][::-1])).encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zw:
        for n, d in parts.items():
            zw.writestr(n, d)
    for d in (data, out.getvalue()):
        _, res, detail = tokenize_file("p.xlsx", d, Context.build(CFG, KEY, Vault()))
        assert "9 cells tokenized, 3 emptied" in detail
        text = b"".join(zipfile.ZipFile(io.BytesIO(res)).read(n) for n in zipfile.ZipFile(io.BytesIO(res)).namelist())
        assert b"Rivera" not in text and b"123-45-6789" not in text


def test_the_progress_lines_in_the_log_never_hold_a_name(capsys):
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv", b"Hours\n8\n")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    lines = [l for l in capsys.readouterr().out.splitlines() if l.startswith("reading ")]
    assert len(lines) >= 2
    assert not any(s in l for l in lines for s in ("Rivera", "Brennick", "week1", "WA.pdf", "Projects"))


def test_the_run_report_in_the_log_holds_counts_and_no_path(capsys, monkeypatch):
    """#21: main.py prints the run's report to Cloud Logging: counts per kind of change, never a file path."""
    import main
    w = World()
    w.payroll()
    # a payroll export: main.py runs the shipped config, which holds back a CSV the agent doesn't read
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv",
            b"Employee ID (LCPtracker),Week End Date,Classification,ST Hours\n12345,11/07/2026,Laborer,8\n")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/call 206-555-0100.xlsx", make_xlsx([["Hours"], [8]]))   # held back
    report = []

    class Bucket:
        def __init__(self, *a):
            self.objs = {}

        def put(self, name, data, if_generation=None):
            self.objs[name] = data
            return 1

        def get(self, name):
            return self.objs.get(name), (1 if name in self.objs else 0)

        def delete(self, name):
            self.objs.pop(name, None)

    def run_all(drive, ctx, state, save_vault, save_state, only=None):
        report.extend(Mirror(w.d, ctx, state, save_vault, save_state).run("compliance", w.raw, w.tok).lines("compliance"))
        return list(report)

    monkeypatch.setattr(main.gcp, "Auth", lambda: None)
    monkeypatch.setattr(main.gcp, "Bucket", Bucket)
    monkeypatch.setattr(main.gcp, "Drive", lambda auth: w.d)
    monkeypatch.setattr(main.gcp, "access_secret", lambda *a: base64.b64encode(KEY).decode())
    monkeypatch.setattr(main, "run_all", run_all)
    cfg = config_mod.load()
    monkeypatch.setattr(cfg, "state_bucket", "bucket", raising=False)
    monkeypatch.setattr(cfg, "key_secret", "secret", raising=False)
    capsys.readouterr()
    assert main.cmd_run(cfg, None) == 0
    out = capsys.readouterr().out
    assert any("Projects/Brennick - 77Q" in l for l in report), report       # the report itself names the files
    assert "[compliance] to AWS (tokenized): " in out, out                   # ... the log only counts them
    assert f"{len(report)} changes reported" in out
    for s in ("Projects", "Brennick", "Rivera", "Mason", "week1", "WA.pdf", "hours", "206-555", "phone", "/"):
        assert s not in out, (s, out)
    refused = [l for l in report if "NOT copied to AWS" in l]
    assert refused and "[compliance] NOT copied to AWS: 1" in out, out
    assert "a phone number" in held(w)["Projects/Brennick - 77Q/Uploads/call 206-555-0100.xlsx"]   # staff: which, why


def _workbook_with(sheet_data=None, extra_sheet_rel=False):
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(make_xlsx([["Employee Name"], ["Zelda Quixote"]])))
    parts = {n: z.read(n) for n in z.namelist()}
    if sheet_data is not None:
        x = parts["xl/worksheets/sheet1.xml"].decode()
        parts["xl/worksheets/sheet1.xml"] = re.sub(r"<sheetData>.*</sheetData>", sheet_data, x, flags=re.S).encode()
    if extra_sheet_rel:
        r = parts["xl/_rels/workbook.xml.rels"].decode()
        rel = ('<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
               'worksheet" Target="worksheets/sheet1.xml"/>')
        parts["xl/_rels/workbook.xml.rels"] = r.replace("</Relationships>", rel + "</Relationships>").encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zw:
        for n, d in parts.items():
            zw.writestr(n, d)
    return out.getvalue()


def _all_text(data):
    import zipfile
    z = zipfile.ZipFile(io.BytesIO(data))
    return b"".join(z.read(n) for n in z.namelist())


C = '<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'


@pytest.mark.parametrize("sheet_data", [
    '<sheetData><row r="1">' + C.format("A2", "Zelda Quixote") + C.format("A1", "Employee Name") + '</row></sheetData>',
    '<sheetData><row r="1">' + C.format("A2", "Zelda Quixote") + '</row><row r="2">' + C.format("A1", "Employee Name")
    + '</row></sheetData>',
    '<sheetData><row>' + C.format("A2", "Zelda Quixote") + '</row><row>' + C.format("A1", "Employee Name")
    + '</row></sheetData>',
])
def test_cells_stored_out_of_order_are_still_tokenized(sheet_data):
    from pii_tokenizer.files import tokenize_file
    _, out, detail = tokenize_file("p.xlsx", _workbook_with(sheet_data), Context.build(CFG, KEY, Vault()))
    assert b"Zelda" not in _all_text(out), detail


def test_a_sheet_listed_twice_is_tokenized_once_and_not_wiped():
    from pii_tokenizer.files import tokenize_file
    _, out, detail = tokenize_file("p.xlsx", _workbook_with(extra_sheet_rel=True), Context.build(CFG, KEY, Vault()))
    import zipfile
    assert b"Zelda" not in _all_text(out)
    assert b"EMP-" in zipfile.ZipFile(io.BytesIO(out)).read("xl/worksheets/sheet1.xml")


@pytest.mark.parametrize("name", ["Timesheet J.Garcia", "Payroll - Smith, J. Maria Lopez", "notes.Rivera"])
def test_the_log_line_never_shows_text_after_a_dot_in_a_name(capsys, name):
    from pii_tokenizer.mirror import _progress
    _progress({"id": "id1", "name": name, "mimeType": "application/octet-stream", "size": "10"})
    out = capsys.readouterr().out
    assert "garcia" not in out.lower() and "lopez" not in out.lower() and "rivera" not in out.lower(), out


# ------------------------------------------------------------------------------------------------ names back from AWS
def test_a_file_aws_makes_with_a_token_in_its_name_comes_back_with_the_name():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt", b"hours checked")
    w.run()
    folder = tok_path(w, "Projects/Brennick - 77Q/Files/EMP-").rsplit("/", 1)[0]
    token = folder.rsplit("/", 1)[1]
    w.d.put(w.tok, f"{folder}/{token} review.txt", f"{token} reviewed".encode())          # written by the agent
    w.run()
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "/Files/" in f["rel"])
    assert staff == ["Projects/Brennick - 77Q/Files/Ana Rivera/Ana Rivera review.txt",
                     "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt"], staff
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/Ana Rivera review.txt") == b"Ana Rivera reviewed"
    rep = w.run()                                              # and nothing more happens
    assert not (rep.forwarded or rep.returned or rep.moved or rep.removed), (rep.forwarded, rep.returned, rep.moved)


def test_a_copy_moved_on_aws_moves_the_staff_file_and_keeps_its_own_name():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana_Rivera timesheet.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    rel = tok_path(w, "Projects/Brennick - 77Q/Uploads/EMP-")
    fid = w.d.path(w.tok, rel)
    arch = w.d.mkdir(w.d.mkdir(w.d.path(w.tok, "Projects/Brennick - 77Q"), "Files"), "Archived")
    w.d.move(fid, arch, w.d.items[fid]["parents"][0])                         # the agent archives it
    w.run()
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/Archived/Ana_Rivera timesheet.xlsx") is not None
    assert w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/Ana_Rivera timesheet.xlsx") is None
    assert no_name_on_aws(w)


def test_a_file_in_a_folder_whose_name_is_learnt_later_follows_the_renamed_folder():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/a.txt", b"one")
    w.run()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/a.txt", b"two")      # changed in the same run
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/b.txt", b"new")      # and a new file
    w.run()
    w.run()
    assert no_name_on_aws(w, "mason"), aws_paths(w)
    files = sorted(f["rel"].rsplit("/", 1)[1] for f in w.d.tree(w.tok)[0].values() if "/Files/" in f["rel"])
    assert files == ["a.txt", "b.txt"], files
    assert len({f["rel"].rsplit("/", 1)[0] for f in w.d.tree(w.tok)[0].values() if "/Files/" in f["rel"]}) == 1
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "/Files/" in f["rel"])
    assert staff == ["Projects/Brennick - 77Q/Files/Lee Mason/a.txt", "Projects/Brennick - 77Q/Files/Lee Mason/b.txt"]


def test_a_stale_folder_is_merged_into_the_one_with_the_token(monkeypatch):
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/a.txt", b"one")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/sub/x.txt", b"x")
    w.run()                                                    # name unknown: copied as it is
    monkeypatch.setattr(Mirror, "_recheck_folders", lambda *a: None)
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/b.txt", b"new")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/sub/y.txt", b"y")
    w.run()                                                    # b and y go into folders with the token
    monkeypatch.undo()
    w.run(); w.run()
    assert no_name_on_aws(w, "mason"), aws_paths(w)
    files = sorted(f["rel"].split("/Files/", 1)[1].split("/", 1)[1] for f in w.d.tree(w.tok)[0].values()
                   if "/Files/" in f["rel"])
    assert files == ["a.txt", "b.txt", "sub/x.txt", "sub/y.txt"], files
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "/Files/" in f["rel"])
    assert staff == [f"Projects/Brennick - 77Q/Files/Lee Mason/{n}" for n in ("a.txt", "b.txt", "sub/x.txt", "sub/y.txt")]


def test_nothing_is_ever_created_on_aws_with_a_known_name_not_even_briefly():
    w = World()
    made = []

    def under_tok(fid):
        while fid:
            if fid == w.tok:
                return True
            ps = w.d.items[fid]["parents"]
            fid = ps[0] if ps else None
        return False
    real_create, real_mkdir, real_update = w.d.create, w.d.mkdir, w.d.update
    w.d.create = lambda parent, name, data, mime: (made.append(name) if under_tok(parent) else None,
                                                   real_create(parent, name, data, mime))[1]
    w.d.mkdir = lambda parent, name: (made.append(name) if under_tok(parent) else None, real_mkdir(parent, name))[1]
    w.d.update = lambda fid, data, mime, name=None: (made.append(name or "") if under_tok(fid) else None,
                                                     real_update(fid, data, mime, name))[1]
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv", b"Hours\n8\n")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Lee Mason/notes.txt", b"checked")
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv", b"Hours\n9\n")       # an update too
    w.run()
    assert made and not any(n in m.casefold() for m in made for n in ("rivera", "mason")), made


# ------------------------------------------------------------------------------------------------ review of #17
MORE = PAYROLL + [["P07", "Acme Paving", 46000, "McDonald, John", 10236, "111-22-3333", "M", "Laborer", 50, ""],
                  ["P07", "Acme Paving", 46000, "Garcia-Lopez, Maria", 10237, "222-33-4444", "F", "Laborer", 50, ""]]


@pytest.mark.parametrize("name", ["John McDonald hours.xlsx", "Maria Garcia-Lopez hours.xlsx",
                                  "Ana B. Rivera hours.xlsx", "Lee J. Mason notes.txt", "MCDONALD_JOHN.csv",
                                  "garcia-lopez, maria.csv", "Ana B.Rivera hours.txt", "Ana.B.Rivera hours.txt",
                                  "McDonald John hours.txt", "McDonald_John hours.txt", "Garcia-Lopez Maria hours.txt",
                                  "JohnMcDonald.csv", "Ana_Rivera_hours.txt", "WeekAnaRivera.csv",
                                  "Week42AnaRiveraHours.csv"])
def test_names_written_as_in_payrolls_are_replaced_in_file_names(name):
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(MORE))
    w.run()
    w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/{name}", b"Hours\n8\n" if not name.endswith("xlsx")
            else make_xlsx([["Hours"], [8]]))
    w.run()
    assert no_name_on_aws(w, "mcdonald", "garcia", "rivera", "mason"), aws_paths(w)
    assert len([r for r in aws_paths(w) if "/Uploads/" in r and re.search(r"(^|[^A-Za-z0-9])EMP-", r)]) == 1, \
        aws_paths(w)


def test_a_name_that_cant_be_replaced_holds_the_file_back(monkeypatch):
    monkeypatch.setattr(screen, "tokenize_name", lambda part, scrubber: part)      # as if no spelling matched
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv", b"Hours\n8\n")
    w.run()
    assert no_name_on_aws(w)
    assert ("create", "Ana Rivera hours.csv") not in w.d.calls                  # never copied, not even briefly
    assert "can't be replaced by its token" in held(w)["Projects/Brennick - 77Q/Uploads/Ana Rivera hours.csv"]


def test_two_spellings_of_one_name_never_share_a_name_on_aws():
    """The first is copied; the second is held back with a plain reason (rare: two spellings of one worker's name
    in one folder). Folders too: the one already in use keeps its place."""
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera notes.txt", b"one")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/ana_rivera notes.txt", b"two")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/ANA RIVERA/a.txt", b"a")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana_Rivera/b.txt", b"b")
    w.run()
    rep = w.run()
    assert not (rep.forwarded or rep.moved or rep.returned or rep.removed)          # stable
    files = [f["rel"].lower() for f in w.d.tree(w.tok)[0].values()]
    assert len(files) == len(set(files)) and no_name_on_aws(w), files
    reasons = held(w)
    assert sum("same name on the AWS side" in r for r in reasons.values()) == 2, reasons
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana_Rivera/c.txt", b"c")            # later, a new file there
    w.run()
    assert sum("same name on the AWS side" in r for r in held(w).values()) == 3
    folder = next(f["rel"] for f in w.d.tree(w.tok)[0].values() if f["rel"].endswith("/a.txt")).rsplit("/", 1)[0]
    w.d.put(w.tok, folder + "/review.txt", b"ok")                                  # the agent writes there
    w.run()
    assert w.d.read(w.raw, "Projects/Brennick - 77Q/Files/ANA RIVERA/review.txt") == b"ok"



def test_files_held_back_for_a_name_by_the_previous_version_are_copied_now():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera notes.txt", b"checked")
    w.run()
    st = w.state["mirrors"]["compliance"]
    fid = w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/Ana Rivera notes.txt")
    for q in [q for q in st["pairs"] if q["raw_id"] == fid]:                   # as the previous version left it
        w.d.items[q["tok_id"]]["trashed"] = True
        st["pairs"].remove(q)
    st["not_copied"][fid] = w.d.meta(fid)["md5Checksum"]
    st["held_back"][fid] = {"rel": "Projects/Brennick - 77Q/Uploads/Ana Rivera notes.txt", "since": "x",
                            "ver": w.d.meta(fid)["md5Checksum"],
                            "reason": "its file name contains a worker's name; rename it and it will be copied"}
    st.pop("names_tokenized")
    w.run()
    assert tok_path(w, "Projects/Brennick - 77Q/Uploads/EMP-").endswith(" notes.txt")
    assert set(held(w)) <= {"(nothing is held back)"}


def test_agent_output_in_a_folder_renamed_meanwhile_comes_back_to_staff():
    w = World()
    w.d.put(w.raw, "Projects/P/Files/Lee Mason/notes.txt", b"checked")
    w.run()                                                    # name unknown: folder copied as it is
    w.payroll()
    w.run()                                                    # folder renamed on AWS with the token
    w.d.put(w.tok, "Projects/P/Files/Lee Mason/summary.txt", b"EMP-x summary")   # written by a run that pulled before
    w.d.put(w.tok, "Projects/P/Files/Lee Mason/notes.txt", b"checked")          # and the old copy pushed again
    w.run(); w.run()
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "/P/" in f["rel"])
    assert staff == ["Projects/P/Files/Lee Mason/notes.txt", "Projects/P/Files/Lee Mason/summary.txt"], staff
    assert no_name_on_aws(w, "mason"), aws_paths(w)



def test_a_staff_rename_into_a_clash_is_taken_back_not_a_shared_name():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/Ana Rivera notes.txt", b"one")
    w.d.put(w.raw, "Projects/P/Notes/plan.txt", b"two")
    w.run()
    w.d.items[w.d.path(w.raw, "Projects/P/Notes/plan.txt")]["name"] = "ana_rivera notes.txt"
    w.run(); w.run()
    files = sorted(f["rel"] for f in w.d.tree(w.tok)[0].values() if "/Notes/" in f["rel"])
    assert len(files) == 1 and w.d.read(w.tok, files[0]) == b"one" and no_name_on_aws(w), files
    assert "same name on the AWS side" in held(w)["Projects/P/Notes/ana_rivera notes.txt"]
    w.d.items[w.d.path(w.raw, "Projects/P/Notes/ana_rivera notes.txt")]["name"] = "plan v2.txt"   # renamed: copied
    w.run()
    assert w.d.read(w.tok, "Projects/P/Notes/plan v2.txt") == b"two"


def test_copies_renamed_for_a_newly_learnt_name_never_share_a_name():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/Ana Rivera notes.txt", b"one")
    w.d.put(w.raw, "Projects/P/Notes/ana_rivera notes.txt", b"two")
    w.run()                                                    # names unknown: copied as they are
    w.payroll()
    w.run(); w.run()
    files = sorted(f["rel"] for f in w.d.tree(w.tok)[0].values() if "/Notes/" in f["rel"])
    assert len(files) == 1 and no_name_on_aws(w), files
    assert sum("same name on the AWS side" in r for r in held(w).values()) == 1


def test_a_token_joined_by_underscore_comes_back_as_the_name():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/Ana_Rivera_docs/Ana_Rivera_hours.txt", b"8")
    w.run()
    rel = tok_path(w, "Projects/P/Notes/EMP-")
    folder, name = rel.rsplit("/", 1)
    assert name.endswith("_hours.txt") and folder.endswith("_docs"), rel
    token = name.split("_", 1)[0]
    w.d.put(w.tok, f"{folder}/{token}_checked.txt", b"done")  # written by the agent
    w.run()
    back = [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith("_checked.txt")]
    assert len(back) == 1 and "EMP-" not in back[0] and back[0].startswith("Projects/P/Notes/Ana_Rivera_docs/"), back



@pytest.mark.parametrize("name", ["Hannah Riveras.txt", "Erica Lee.txt", "Riverside plan.txt"])
def test_words_that_only_look_like_a_name_are_left_as_they_are(name):
    w = World()
    w.payroll()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert w.d.read(w.tok, f"Projects/P/Notes/{name}") == b"x"
    assert set(held(w)) <= {"(nothing is held back)"}


@pytest.mark.parametrize("name", ["Banana Rivera notes.txt", "Lee Masonry quote.txt"])
def test_a_name_hidden_inside_other_letters_holds_the_file_back(name):
    """The strict check: rarely a false alarm, never a name sent."""
    w = World()
    w.payroll()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert w.d.read(w.tok, f"Projects/P/Notes/{name}") is None
    assert "can't be replaced" in held(w)[f"Projects/P/Notes/{name}"]



def test_a_token_glued_to_other_words_is_set_apart_so_it_comes_back():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/WeekAnaRiveraHours.csv", b"Hours\n8\n")
    w.run()
    rel = tok_path(w, "Projects/P/Notes/")
    assert re.fullmatch(r"Projects/P/Notes/Week EMP-[0-9a-f]+ Hours\.csv", rel), rel
    w.d.put(w.tok, rel.replace("Hours.csv", "Review.txt"), b"ok")                  # the agent writes a file
    w.run()
    back = [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith("Review.txt")]
    assert len(back) == 1 and "EMP-" not in back[0], back



def test_a_deleted_file_replaced_by_one_with_the_same_name_keeps_the_name():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/notes.txt", b"one")
    w.run()
    w.d.items[w.d.path(w.raw, "Projects/P/Notes/notes.txt")]["trashed"] = True
    w.d.put(w.raw, "Projects/P/Notes/notes.txt", b"two")                         # a new file, same name
    w.run(); w.run()
    assert w.d.read(w.tok, "Projects/P/Notes/notes.txt") == b"two"
    assert set(held(w)) <= {"(nothing is held back)"}


def test_a_copy_moved_on_aws_and_edited_by_staff_keeps_its_name():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/hours.csv", b"Hours\n8\n")
    w.run()
    fid = w.d.path(w.tok, "Projects/P/Notes/hours.csv")
    arch = w.d.mkdir(w.d.path(w.tok, "Projects/P/Notes"), "Archived")
    w.d.move(fid, arch, w.d.items[fid]["parents"][0])
    w.d.put(w.raw, "Projects/P/Notes/hours.csv", b"Hours\n9\n")                # edited the same day
    w.run(); w.run()
    tok_files = sorted(f["rel"] for f in w.d.tree(w.tok)[0].values() if f["rel"].endswith(".csv"))
    assert tok_files == ["Projects/P/Notes/Archived/hours.csv"], tok_files
    assert set(held(w)) <= {"(nothing is held back)"}


@pytest.mark.parametrize("name", ["Erica Lee.txt", "Dana Smith.txt", "Anas Rivera.txt", "Leea Eric.txt"])
def test_other_peoples_names_close_to_a_worker_name_are_left_alone(name):
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(PAYROLL + [
        ["P07", "Acme Paving", 46000, "Lee, Eric", 1, "", "", "Laborer", 50, ""],
        ["P07", "Acme Paving", 46000, "Smith, Dan", 2, "", "", "Laborer", 50, ""]]))
    w.run()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert w.d.read(w.tok, f"Projects/P/Notes/{name}") == b"x"


@pytest.mark.parametrize("name", ["José Peña hours.txt", "Ｊｏｓé Ｐｅñａ.txt", "Peña, José.txt", "JosePena.txt"])
def test_accented_and_full_width_names_are_replaced(name):
    import unicodedata
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(PAYROLL + [
        ["P07", "Acme Paving", 46000, "Peña, José", 1, "", "", "Laborer", 50, ""]]))
    w.run()
    for form in ("NFC", "NFD"):
        w.d.put(w.raw, f"Projects/P/{form}/" + unicodedata.normalize(form, name), b"x")
    w.run()
    paths = " ".join(aws_paths(w))
    assert "EMP-" in paths and not any(x in unicodedata.normalize("NFKC", paths).casefold() for x in ("pena", "peña")), \
        aws_paths(w)


@pytest.mark.parametrize("name", ["Watson, Mary-Jane.txt", "Watson Mary Jane.txt", "WatsonMaryJane.txt"])
def test_a_name_learnt_without_a_comma_is_found_in_other_orders(name):
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(PAYROLL + [
        ["P07", "Acme Paving", 46000, "Mary-Jane Watson", 1, "", "", "Laborer", 50, ""]]))
    w.run()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert no_name_on_aws(w, "watson"), aws_paths(w)


def test_the_folder_already_in_use_keeps_its_place_when_a_second_spelling_appears():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/ana_rivera/x.txt", b"x")
    w.run()
    w.d.put(w.raw, "Projects/P/Notes/Ana-Rivera/y.txt", b"y")       # sorts first, appears later
    w.run()
    x = next(f["rel"] for f in w.d.tree(w.tok)[0].values() if f["rel"].endswith("/x.txt"))
    w.d.put(w.tok, x.rsplit("/", 1)[0] + "/review.txt", b"ok")       # the agent writes next to x
    w.run()
    assert w.d.read(w.raw, "Projects/P/Notes/ana_rivera/review.txt") == b"ok"
    assert "same name on the AWS side" in held(w)["Projects/P/Notes/Ana-Rivera/y.txt"]


# ------------------------------------------------------------------------------------------------ review 4 of #17
@pytest.mark.parametrize("name", ["TSAnaRivera.csv", "CPRAnaRivera hours.txt", "IDLeeMason.txt", "Ana RIVERAhours.txt",
                                  "José" + "AnaRivera.txt", "JoséAnaRivera.txt", "xxAnaRiveraxx.txt"])
def test_names_glued_to_capitals_or_accents_never_reach_aws(name):
    w = World()
    w.payroll()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert no_name_on_aws(w), aws_paths(w)


def test_a_file_held_back_for_a_clash_is_copied_once_the_other_is_renamed():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/Ana Rivera notes.txt", b"one")
    w.d.put(w.raw, "Projects/P/Notes/ana_rivera notes.txt", b"two")
    w.run()
    held_one = next(k for k, v in held(w).items() if "same name on the AWS side" in v)
    other = ({"Projects/P/Notes/Ana Rivera notes.txt", "Projects/P/Notes/ana_rivera notes.txt"} - {held_one}).pop()
    w.d.items[w.d.path(w.raw, other)]["name"] = "plan.txt"                    # staff rename the copied one
    w.run(); w.run()
    assert set(held(w)) <= {"(nothing is held back)"}
    assert {w.d.read(w.tok, f["rel"]) for f in w.d.tree(w.tok)[0].values() if "/Notes/" in f["rel"]} == {b"one", b"two"}


def test_a_new_agent_file_named_like_a_staff_file_comes_back_once_as_an_aws_copy():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/plan.txt", b"staff")
    w.run()
    w.d.put(w.raw, "Projects/P/Notes/Ana Rivera notes.txt", b"mine")         # staff file, not yet copied...
    w.d.put(w.tok, "Projects/P/Notes/" + tok_path(w, "Projects/P/Notes/plan").rsplit("/", 1)[1]
            .replace("plan.txt", "x"), b"")                                   # (unrelated noise)
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/P/Notes/{token} notes.txt", b"agent")         # ...and the agent writes that name
    w.run()
    rep = w.run()
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "notes" in f["rel"])
    assert len(staff) == 2 and any("(AWS copy" in f for f in staff), staff
    assert not (rep.returned or rep.forwarded), (rep.returned, rep.forwarded)


def test_a_folder_used_only_through_its_subfolders_keeps_its_place():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/ana_rivera/Sub/x.txt", b"x")
    w.run()
    w.d.put(w.raw, "Projects/P/Notes/Ana-Rivera/y.txt", b"y")
    w.d.put(w.raw, "Projects/P/Notes/Ana-Rivera/Sub2/z.txt", b"z")
    w.run(); w.run()
    assert "same name on the AWS side" in held(w)["Projects/P/Notes/Ana-Rivera/y.txt"]
    tok_dirs = w.d.tree(w.tok)[1]
    assert not any(d.endswith("/Sub2") for d in tok_dirs), tok_dirs                # nothing of the other folder
    w.d.put(w.raw, "Projects/P/Notes/ana_rivera/Sub/x.txt", b"x2")             # the copy stays
    rep = w.run()
    assert not rep.removed and w.d.read(w.tok, tok_path(w, "Projects/P/Notes/EMP-")) == b"x2"


@pytest.mark.parametrize("name", ["TSAnaRivera.csv", "IDLeeMason.txt", "Ana RIVERAhours.txt", "JoséAnaRivera.txt"])
def test_names_glued_to_capitals_or_accents_are_replaced_not_held_back(name):
    w = World()
    w.payroll()
    w.d.put(w.raw, f"Projects/P/Notes/{name}", b"x")
    w.run()
    assert set(held(w)) <= {"(nothing is held back)"}, held(w)
    assert len([r for r in aws_paths(w) if "/Notes/" in r and "EMP-" in r]) == 1, aws_paths(w)


# ------------------------------------------------------------------------------------------------ review 5 of #17
@pytest.mark.parametrize("payroll_name,file_name", [
    ("Rivera, Ana", "Аna Rivera.xlsx"), ("Rivera, Ana", "Rivеra Ana.xlsx"), ("Rivera, Ana", "Rıvera Ana.xlsx"),
    ("Nowak, Łukasz", "Lukasz Nowak.xlsx"), ("Nowak, Lukasz", "Łukasz Nowak.xlsx"),
    ("Yılmaz, İlker", "Ilker Yilmaz.xlsx"), ("Hansen, Søren", "Soren Hansen.xlsx"),
])
def test_look_alike_and_plain_letters_are_the_same_name(payroll_name, file_name):
    import unicodedata
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/week1.xlsx", make_xlsx(PAYROLL + [
        ["P07", "Acme Paving", 46000, payroll_name, 1, "", "", "Laborer", 50, ""]]))
    w.run()
    w.d.put(w.raw, f"Projects/Brennick - 77Q/Uploads/{file_name}", make_xlsx([["Hours"], [8]]))
    w.run()
    names = [f["name"] for f in w.d.tree(w.tok)[0].values() if f["name"] != "week1.xlsx"]
    assert file_name not in names and not any(f.startswith(file_name.split(" ")[0]) for f in names), names
    assert set(held(w)) <= {"(nothing is held back)"}                                # replaced, not held back


def test_a_name_split_across_folder_levels_is_held_back():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Ana/Rivera.xlsx", make_xlsx([["Hours"], [8]]))
    w.run()
    assert w.d.path(w.tok, "Projects/Brennick - 77Q/Uploads/Ana/Rivera.xlsx") is None
    assert "spell a worker's name" in held(w)["Projects/Brennick - 77Q/Uploads/Ana/Rivera.xlsx"]


# ------------------------------------------------------------------------------------------------ full review of #17
def _token_for(w, kind_prefix):
    return next(t for t, e in w.ctx.vault.entries.items() if t.startswith(kind_prefix + "-"))


@pytest.mark.parametrize("prefix", ["EID", "SSN"])
def test_id_and_ssn_tokens_in_agent_file_names_never_become_values_on_aws(prefix):
    w = World()
    w.payroll()
    w.run()
    tok = _token_for(w, prefix)
    w.d.put(w.tok, f"Projects/Brennick - 77Q/Files/{tok} review.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.tok, f"Projects/Brennick - 77Q/Files/{tok}/notes.txt", b"ok")
    w.run(); w.run()
    paths = " ".join(aws_paths(w))
    assert "10234" not in paths and "123-45-6789" not in paths and "123456789" not in paths, aws_paths(w)
    assert tok in paths
    staff = " ".join(f["rel"] for f in w.d.tree(w.raw)[0].values())
    assert "10234" not in staff and "123-45-6789" not in staff                     # a token stays a token
    assert set(held(w)) <= {"(nothing is held back)"}


def test_a_copy_pushed_back_under_its_old_name_during_a_rename_keeps_the_staff_file():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/Lee Mason notes.txt", b"v1")
    w.run()
    for st in (w.state, w.saved):
        for q in st["mirrors"]["compliance"]["pairs"]:
            q["copied_utc"] = "2020-01-01-00-00-00"
    w.payroll()
    w.run()                                                    # renamed on AWS to the token
    new = tok_path(w, "Projects/P/Notes/EMP-")
    fid = w.d.path(w.tok, new)                                 # an AWS run that pulled before, now pushing an edit:
    bk = w.d.mkdir(w.d.mkdir(w.tok, "_sync_backup"), "x")      # rclone moves the new name aside...
    w.d.move(fid, bk, w.d.items[fid]["parents"][0])
    w.d.put(w.tok, "Projects/P/Notes/Lee Mason notes.txt", b"v2 from AWS")      # ...and uploads the old name
    w.run(); w.run()
    staff = sorted(f["rel"] for f in w.d.tree(w.raw)[0].values() if "notes" in f["rel"].lower())
    assert staff == ["Projects/P/Notes/Lee Mason notes.txt"], staff                # not set aside, no AWS copy
    assert w.d.read(w.raw, "Projects/P/Notes/Lee Mason notes.txt") == b"v2 from AWS"
    assert no_name_on_aws(w, "mason") or all(r.startswith("_sync_backup") for r in aws_paths(w) if "Mason" in r)


def test_agent_output_with_the_same_content_as_some_copy_is_kept():
    w = World()
    w.d.put(w.raw, "Projects/P/Files/Lee Mason/template.txt", b"blank")
    w.d.put(w.raw, "Projects/P/Files/other.txt", b"blank")
    w.run()                                                    # name unknown: folder copied as it is
    w.payroll()
    w.run()                                                    # folder renamed on AWS with the token
    w.d.put(w.tok, "Projects/P/Files/Lee Mason/review.txt", b"blank")          # written by a run that pulled before
    w.run(); w.run()
    assert w.d.read(w.raw, "Projects/P/Files/Lee Mason/review.txt") == b"blank"


def test_names_in_the_aws_sides_own_backups_are_removed():
    w = World()
    w.d.put(w.tok, "_sync_backup/2026-10-01/Lee Mason/old.txt", b"x")
    w.d.put(w.tok, "_sync_backup/2026-10-01/Ana Rivera notes.txt", b"x")
    w.d.put(w.raw, "Projects/P/a.txt", b"a")
    w.state = {"mirrors": {"compliance": {"started": True, "pairs": [], "roots": [w.raw, w.tok]}}}
    w.payroll()
    w.run()
    assert not any("Mason" in r or "Rivera" in r for r in aws_paths(w)), aws_paths(w)


def test_an_agent_file_with_a_token_glued_to_a_word_is_not_taken_back():
    w = World()
    w.payroll()
    w.run()
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/P/Files/{token}hours.txt", b"8")
    w.run(); rep = w.run()
    assert set(held(w)) <= {"(nothing is held back)"} and not rep.removed
    assert any(r.endswith("hours.txt") for r in aws_paths(w)) and no_name_on_aws(w)
    staff = [f["rel"] for f in w.d.tree(w.raw)[0].values() if f["rel"].endswith("hours.txt")]
    assert staff == ["Projects/P/Files/Ana Rivera hours.txt"], staff              # the name, set apart


def test_an_agent_file_saved_as_an_aws_copy_keeps_its_name_on_aws():
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/P/Notes/Ana Rivera notes.txt", b"mine")
    w.run()
    w.d.items[w.d.path(w.tok, tok_path(w, "Projects/P/Notes/EMP-"))]["trashed"] = True   # (its copy gone)
    st = w.state["mirrors"]["compliance"]
    st["pairs"].clear()
    st["not_copied"][w.d.path(w.raw, "Projects/P/Notes/Ana Rivera notes.txt")] = "x"
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/P/Notes/{token} notes.txt", b"agent")
    w.run()
    rep = w.run()
    assert w.d.read(w.tok, f"Projects/P/Notes/{token} notes.txt") == b"agent" and not rep.moved, rep.moved


# ------------------------------------------------------------------------------------------------ verification of #17
def test_a_project_folder_made_on_aws_with_a_token_never_gets_the_name_on_aws():
    w = World()
    w.payroll()
    w.run()
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/{token}/notes.txt", b"x")
    w.d.put(w.tok, f"Projects/{token} Plaza/notes.txt", b"y")
    w.run(); rep = w.run()
    assert no_name_on_aws(w), aws_paths(w)
    assert not (rep.forwarded or rep.returned or rep.moved or rep.removed)
    staff_dirs = w.d.tree(w.raw)[1]
    assert f"Projects/{token}" in staff_dirs and "Projects/Ana Rivera" not in staff_dirs, staff_dirs


def test_the_clash_index_is_kept_per_mirror():
    d = FakeDrive()
    roots = {k: d.root(k) for k in ("ra", "ta", "rb", "tb")}
    ctx = Context.build(CFG, KEY, Vault())
    state = {}
    m = Mirror(d, ctx, state, lambda: None, lambda: None)
    d.put(roots["ra"], "P/x.txt", b"a")
    d.put(roots["rb"], "P/y.txt", b"b")
    m.run("a", roots["ra"], roots["ta"])
    m.run("b", roots["rb"], roots["tb"])
    d.items[d.path(roots["rb"], "P/y.txt")]["name"] = "x.txt"               # same path as mirror a's copy
    d.put(roots["ra"], "P/new.txt", b"n")                                      # (mirror a's index is built late)
    m.run("a", roots["ra"], roots["ta"])
    m.run("b", roots["rb"], roots["tb"])
    assert d.read(roots["tb"], "P/x.txt") == b"b"
    assert not state["mirrors"]["b"]["held_back"]


@pytest.mark.parametrize("aws_moves", ["a", "b"])                    # either one handled first
def test_an_aws_move_and_a_staff_rename_in_one_run_never_share_a_name(aws_moves):
    w = World()
    w.d.put(w.raw, "Projects/P/a.txt", b"a")
    w.d.put(w.raw, "Projects/P/b.txt", b"b")
    w.run()
    staff_renames = "b" if aws_moves == "a" else "a"
    w.d.items[w.d.path(w.tok, f"Projects/P/{aws_moves}.txt")]["name"] = "zz.txt"
    w.d.items[w.d.path(w.raw, f"Projects/P/{staff_renames}.txt")]["name"] = "zz.txt"
    w.run(); w.run()
    files = [f["rel"].lower() for f in w.d.tree(w.tok)[0].values()]
    assert len(files) == len(set(files)), files


def test_backups_of_projects_named_after_people_are_kept():
    w = World()
    w.d.put(w.tok, "_sync_backup/2026-10-01__sync_backup/Projects/Lee Mason Plaza/old.txt", b"x")
    w.d.put(w.raw, "Projects/P/a.txt", b"a")
    w.state = {"mirrors": {"compliance": {"started": True, "pairs": [], "roots": [w.raw, w.tok]}}}
    w.payroll()
    w.run()
    assert w.d.read(w.tok, "_sync_backup/2026-10-01__sync_backup/Projects/Lee Mason Plaza/old.txt") == b"x"


def test_an_aws_folder_with_a_glued_token_is_not_doubled():
    w = World()
    w.payroll()
    w.run()
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/P/Files/{token}hours/a.txt", b"8")
    w.run(); w.run()
    dirs = [d for d in w.d.tree(w.tok)[1] if d.startswith("Projects/P/Files/")]
    assert len(dirs) == 1 and no_name_on_aws(w), dirs


def test_a_staff_edit_of_a_file_made_on_aws_keeps_its_aws_name():
    w = World()
    w.payroll()
    w.run()
    token = screen.tokenize_name("Ana Rivera", w.ctx.scrubber)
    w.d.put(w.tok, f"Projects/P/Files/{token}hours.txt", b"8")
    w.run()
    w.d.put(w.raw, "Projects/P/Files/Ana Rivera hours.txt", b"9")              # staff edit it
    w.run()
    assert w.d.read(w.tok, f"Projects/P/Files/{token}hours.txt") == b"9", aws_paths(w)


# ------------------------------------------------------------------------------------------------ link in the email
def test_the_summary_carries_the_lists_drive_id_and_nothing_else_about_it():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/a.txt", b"a")
    w.run()
    data = json.loads(w.d.read(w.tok, screen.HELD_BACK_SUMMARY))
    assert data["sheet_id"] is None                                            # nothing held back: no list
    w.d.put(w.raw, "Projects/P/Notes/call 206-555-0100.txt", b"x")
    w.run()
    data = json.loads(w.d.read(w.tok, screen.HELD_BACK_SUMMARY))
    assert data["sheet_id"] == w.d.path(w.raw, screen.HELD_BACK_SHEET)
    assert "206-555" not in json.dumps(data) and "call" not in json.dumps(data)  # no name, path or reason
    w.run()
    assert json.loads(w.d.read(w.tok, screen.HELD_BACK_SUMMARY))["sheet_id"] == data["sheet_id"]    # stable


def test_a_list_shared_with_anyone_with_the_link_is_warned_about():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/call 206-555-0100.txt", b"x")
    w.run()
    w.d.items[w.d.path(w.raw, screen.HELD_BACK_SHEET)]["anyone"] = True
    rep = w.run()
    assert any("shared with anyone who has the link" in x for x in rep.warnings)
    w.d.items[w.d.path(w.raw, screen.HELD_BACK_SHEET)]["anyone"] = False
    assert not any("anyone who has the link" in x for x in w.run().warnings)


def test_no_link_to_a_list_that_could_not_be_updated():
    w = World()
    w.d.put(w.raw, "Projects/P/Notes/call 206-555-0100.txt", b"x")
    w.run()
    w.d.items[w.d.path(w.raw, "Projects/P/Notes/call 206-555-0100.txt")]["name"] = "call.txt"
    w.run()                                                    # released: the list says nothing is held back
    w.d.put(w.raw, "Projects/P/Notes/mail bob@example.com.txt", b"y")
    sheet = w.d.path(w.raw, screen.HELD_BACK_SHEET)
    real = w.d.update
    w.d.update = lambda fid, *a, **k: (_ for _ in ()).throw(OSError("503")) if fid == sheet else real(fid, *a, **k)
    w.run()
    data = json.loads(w.d.read(w.tok, screen.HELD_BACK_SUMMARY))
    assert len(data["held_back"]) == 1 and data["sheet_id"] is None


# ------------------------------------------------------------------------------------------------ system review
@pytest.mark.parametrize("with_backup", [True, False])
def test_an_older_copy_pushed_back_by_an_overlapping_aws_run_never_undoes_a_staff_edit(with_backup):
    """The nightly tokenizer runs while an AWS run is between its pull and its push: the push brings back the copy
    it pulled, older than the one the tokenizer just wrote. That is not a change made on AWS."""
    import zipfile
    w = World()
    w.payroll()
    rel = "Projects/Brennick - 77Q/Tracker/tracker.xlsx"
    w.d.put(w.raw, rel, make_xlsx([["Note"], ["v1"]]))
    w.run()
    pulled = w.d.read(w.tok, rel)                                     # the AWS run pulls this
    w.d.put(w.raw, rel, make_xlsx([["Note"], ["v2 staff edit"]]))    # staff edit meanwhile...
    w.run()                                                           # ...and the nightly tokenizer copies it
    fid = w.d.path(w.tok, rel)
    if with_backup:                                                   # rclone --backup-dir: aside, then upload
        bk = w.d.mkdir(w.d.mkdir(w.tok, "_sync_backup"), "2026-10-31__sync_backup")
        w.d.move(fid, bk, w.d.items[fid]["parents"][0])
        w.d.put(w.tok, rel, pulled)
    else:
        w.d.put(w.tok, rel, pulled)                                   # overwritten in place
    rep = w.run(); w.run()
    staff = w.d.read(w.raw, rel)
    xml = b"".join(zipfile.ZipFile(io.BytesIO(staff)).read(n) for n in zipfile.ZipFile(io.BytesIO(staff)).namelist())
    assert b"v2 staff edit" in xml, "the staff edit was undone"
    assert any("pushed back an older copy" in x for x in rep.warnings)
    copy = w.d.read(w.tok, rel)
    xml = b"".join(zipfile.ZipFile(io.BytesIO(copy)).read(n) for n in zipfile.ZipFile(io.BytesIO(copy)).namelist())
    # and AWS gets the current version again (a note the agent doesn't read: as its text token, see doubt.py)
    assert w.ctx.keyring.text_token("v2 staff edit").encode() in xml


def test_aws_side_backups_older_than_30_days_are_pruned_and_the_rest_kept():
    from datetime import datetime, timezone
    w = World()
    w.d.put(w.raw, "Projects/P/a.txt", b"a")
    w.run()
    w.d.put(w.tok, "_sync_backup/2026-09-01-02-00-00__sync_backup/Projects/P/old.txt", b"o")
    w.d.put(w.tok, "_sync_backup/2026-10-20-02-00-00__sync_backup/Projects/P/new.txt", b"n")
    w.d.put(w.tok, "_sync_backup/notes by hand/x.txt", b"x")
    w.d.put(w.raw, "_sync_backup/2026-01-01-00-00-00__deleted_on_aws/a.txt", b"staff")   # staff side: never pruned
    m = Mirror(w.d, w.ctx, w.state, lambda: None, lambda: None,
               now=lambda: datetime(2026, 10, 30, 12, tzinfo=timezone.utc))
    rep = m.run("compliance", w.raw, w.tok)
    assert w.d.read(w.tok, "_sync_backup/2026-09-01-02-00-00__sync_backup/Projects/P/old.txt") is None
    assert w.d.read(w.tok, "_sync_backup/2026-10-20-02-00-00__sync_backup/Projects/P/new.txt") == b"n"
    assert w.d.read(w.tok, "_sync_backup/notes by hand/x.txt") == b"x"
    assert w.d.read(w.raw, "_sync_backup/2026-01-01-00-00-00__deleted_on_aws/a.txt") == b"staff"
    assert any("older than 30 days" in d for _, d in rep.removed)


# ------------------------------------------------------------------------------------------- issue #22
# Every file and folder name the tests below use: none may appear in the list published to AWS.
NAMES_USED = ("Projects", "Brennick - 77Q", "Brennick", "77Q", "Uploads", "Files", "Ana Rivera", "Rivera", "Lee Mason",
              "Mason", "hours.xlsx", "Lee Mason hours.xlsx", "notes.txt", "a.txt", "b.txt", "WA20250001.pdf",
              "WA20250002.pdf", "WA.pdf", "note.txt", "_sync_backup", "__sync_backup", "week1.xlsx")


def purge_entries(w):
    data = w.d.read(w.tok, screen.PURGE_SUMMARY)
    return [] if data is None else json.loads(data)["purge"]


def purged(w, kind=None):
    return {e["md5"] for e in purge_entries(w) if kind in (None, e["kind"])}


def md5_at(w, rel):
    """The Drive MD5 of the AWS-side file at ``rel`` now."""
    return w.d.get(w.d.path(w.tok, rel))["md5Checksum"]


def run_at(w, when):
    return Mirror(w.d, w.ctx, w.state, lambda: None, lambda: None, now=lambda: when).run("compliance", w.raw, w.tok)


def assert_no_path_in_list(w):
    text = w.d.read(w.tok, screen.PURGE_SUMMARY).decode()
    for n in NAMES_USED:
        assert n not in text, n
    data = json.loads(text)
    assert set(data) == {"purge", "note"}
    assert all(set(e) == {"md5", "kind", "at"} and re.fullmatch(r"[0-9a-f]{32}", e["md5"])
               and e["kind"] in ("content", "path") for e in data["purge"])
    assert len(data["purge"]) == len({e["md5"] for e in data["purge"]})      # each content once


def test_content_renamed_or_taken_back_for_a_name_is_listed_by_md5_only():
    """Issue #22: copies renamed, moved or taken back on the AWS side because a worker's name was learnt after they
    reached AWS leave old S3 versions. Their content is listed by its Drive MD5 (S3's ETag), never by path."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/hours.xlsx", make_xlsx([["Hours"], [8]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt", b"hours checked")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/sub/a.txt", b"deeper")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    assert purge_entries(w) == []                                    # nothing learnt yet: nothing to purge
    fid = w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/hours.xlsx")
    w.d.items[fid]["name"] = "Lee Mason hours.xlsx"                  # renamed by staff after a worker...
    w.run()                                                          # (name not known yet: copied as named)
    before = {rel: md5_at(w, rel) for rel in ("Projects/Brennick - 77Q/Uploads/Lee Mason hours.xlsx",
                                              "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt",
                                              "Projects/Brennick - 77Q/Files/Ana Rivera/sub/a.txt",
                                              "Projects/Brennick - 77Q/Uploads/WA20250001.pdf")}
    w.payroll()                                                      # ... then the names are learnt
    w.run()
    assert no_name_on_aws(w)
    assert set(before.values()) <= purged(w), before                 # file renamed, folder renamed, PDF taken back
    pdf = before.pop("Projects/Brennick - 77Q/Uploads/WA20250001.pdf")
    assert pdf in purged(w, "content")                               # its text held the name: any copy of it
    assert set(before.values()) <= purged(w, "path")                 # only their paths did
    payroll_copy = next(f for f in w.d.tree(w.tok)[0].values() if f["rel"].endswith("week1.xlsx"))
    assert payroll_copy["md5Checksum"] not in purged(w)              # an unrelated copy is not listed
    assert_no_path_in_list(w)
    n = len(purge_entries(w))
    w.run()
    assert len(purge_entries(w)) == n                                # listed once, kept for later AWS runs


def test_a_copy_binned_in_the_copy_step_is_listed_by_what_the_old_copy_holds():
    """Staff change a PDF in the run that learns a name the copy on AWS holds: the new version is held back and
    the old copy binned in the copy step. The old copy's text holds the name: listed as content."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()                                                          # name unknown yet: copied
    old = md5_at(w, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, foreman"]]))
    rep = w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert any("previous tokenized copy" in d for _, d in rep.removed)
    assert old in purged(w, "content")
    assert_no_path_in_list(w)


@pytest.mark.parametrize("new", ["name", "scan page"])
def test_a_clean_old_copy_binned_in_the_copy_step_is_not_listed(new):
    """The old copy was screened clean and still is: binned because the new version can't go (it holds a name,
    or can't be checked), it is not listed. Wage determinations are the same bytes in many projects: listing one
    would wipe every project's old versions of it."""
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf",
            make_pdf([WD[0] + ["Ana Rivera, laborer"]]) if new == "name" else make_pdf(WD, image_only_page=True))
    rep = w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert any("previous tokenized copy" in d for _, d in rep.removed)
    assert purge_entries(w) == []


def test_a_pending_take_back_binned_after_staff_changed_the_file_again_is_listed():
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()                                                          # name unknown yet: copied
    old = md5_at(w, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, foreman"]]))
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    w.run()                                                          # held back; the old copy couldn't be binned
    assert any(p.get("pending_take_back") for p in w.state["mirrors"]["compliance"]["pairs"])
    assert old not in purged(w)
    w.d.trash = real
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Lee Mason, laborer"]]))   # changed again
    rep = w.run()
    assert any("previous tokenized copy" in d for _, d in rep.removed)
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert old in purged(w, "content")


@pytest.mark.parametrize("again", [False, True])
def test_a_pending_take_back_of_a_clean_copy_is_not_listed(again):
    """Its bin failed, then it was binned next run (as it was, or after staff changed it again): the copy held no
    name, so nothing is listed, nor when an AWS run pushes those bytes back."""
    w = World()
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD))
    w.run()
    pdf = w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf(WD, image_only_page=True))   # can't be checked
    real = w.d.trash
    w.d.trash = lambda fid: (_ for _ in ()).throw(OSError("Drive said no"))
    w.run()
    w.d.trash = real
    if again:
        w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0], "page two"], image_only_page=True))
    rep = w.run()                                                    # binned now
    assert any("tokenized copy" in d or "taken back" in d for _, d in rep.removed), rep.removed
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert purge_entries(w) == []
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf", pdf)       # an AWS run pushes the old copy back
    w.run()
    assert purge_entries(w) == []


def test_a_backup_pdf_binned_for_another_reason_is_not_listed():
    w = World()
    w.run()
    keep = backup_dir() + "Projects/Brennick - 77Q/Uploads/"
    w.d.put(w.tok, keep + "WA20250001.pdf", make_pdf(WD, image_only_page=True))
    w.payroll()
    w.run()                                                          # names learnt: backup PDFs read again
    assert w.d.read(w.tok, keep + "WA20250001.pdf") is None          # binned: it can't be checked
    assert purge_entries(w) == []


def test_a_binned_copy_is_listed_as_its_kind_and_its_earlier_copies_as_path():
    w = World()
    w.run()
    m = Mirror(w.d, w.ctx, w.state, lambda: None, lambda: None)
    m._st = w.state["mirrors"]["compliance"]
    x, y = "a" * 32, "b" * 32
    m._stale_copy("content", x, {"tok_ver": x, "tok_hist": [y, x]})
    m._stale_copy(None, "c" * 32, {"tok_ver": "c" * 32, "tok_hist": ["d" * 32]})   # no name: nothing
    assert {e["md5"]: e["kind"] for e in m._st["purge"]} == {x: "content", y: "path"}


def test_a_pdf_whose_path_holds_a_name_is_listed_as_content_when_its_text_does_too():
    """A copy taken back for a name in its path that can't be replaced by the token: a PDF is read too."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/AnaRivera.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/LeeMason.pdf", make_pdf(WD))
    w.run()
    named, clean = (md5_at(w, "Projects/Brennick - 77Q/Uploads/AnaRivera.pdf"),
                    md5_at(w, "Projects/Brennick - 77Q/Uploads/LeeMason.pdf"))
    from pii_tokenizer.mirror import Mirror as M
    m = M(w.d, w.ctx, w.state, lambda: None, lambda: None)
    m._tok = w.d.tree(w.tok)[0]
    w.payroll()
    w.run()
    for rel, md5, kind in (("AnaRivera.pdf", named, "content"), ("LeeMason.pdf", clean, "path")):
        fid = next(i for i, f in m._tok.items() if f["rel"].endswith(rel))
        assert m._kind_of_old(fid, f"Projects/Brennick - 77Q/Uploads/{rel}") == kind, rel
    w.d.put(w.tok, "Projects/Brennick - 77Q/Files/scan.pdf", make_pdf(WD, image_only_page=True))   # can't be checked
    m._tok = w.d.tree(w.tok)[0]
    fid = next(i for i, f in m._tok.items() if f["rel"].endswith("scan.pdf"))
    assert m._kind_of_old(fid, "Projects/Brennick - 77Q/Files/scan.pdf") is None


def test_a_pdf_taken_back_for_its_path_is_read_and_listed_as_content(monkeypatch):
    """Staff rename a PDF to a name that can't go (an SSN in it) before its text was checked again for the names
    learnt since: the copy on AWS is read when it is taken back, and its text holds a name."""
    import pii_tokenizer.mirror as mirror_mod
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.run()
    old = md5_at(w, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.payroll()
    monkeypatch.setattr(mirror_mod, "RECHECK_SECONDS", -1)            # its text not checked again yet
    w.run()
    assert old not in purged(w)
    fid = w.d.path(w.raw, "Projects/Brennick - 77Q/Uploads/WA.pdf")
    w.d.items[fid]["name"] = "WA 123-45-6789.pdf"                     # renamed by staff to a name that can't go
    w.run()
    assert w.d.read(w.tok, "Projects/Brennick - 77Q/Uploads/WA.pdf") is None
    assert old in purged(w, "content")


def test_a_copy_with_a_name_pushed_again_from_aws_keeps_its_entry_for_another_35_days():
    from datetime import datetime, timedelta, timezone
    from pii_tokenizer import mirror
    t0 = datetime.now(timezone.utc)
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    run_at(w, t0)                                                    # name unknown: copied as it is
    w.payroll()
    run_at(w, t0)                                                    # renamed on AWS with the token: listed
    md5 = next(iter(purged(w)))
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")   # an AWS run pushes the old name
    later = t0 + timedelta(days=20)
    rep = run_at(w, later)
    assert any("uploaded again from AWS" in d for _, d in rep.removed)
    assert {e["md5"]: e["at"] for e in purge_entries(w)}[md5] == later.strftime("%Y-%m-%d")   # from the re-push
    run_at(w, t0 + timedelta(days=mirror.PURGE_KEEP_DAYS + 1))
    assert md5 in purged(w)                                          # 36 days after the rename, 16 after the push
    w.state["mirrors"]["compliance"]["purge"] = []                   # expired meanwhile: pushed under the name again
    w.d.put(w.tok, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    run_at(w, t0 + timedelta(days=40))
    assert md5 in purged(w, "path")


def test_a_folder_merge_that_fails_part_way_lists_what_it_moved(monkeypatch):
    """The token-named folder exists already, so the files of the name-bearing one are moved into it one by one. A
    merge stopped part-way lists the files it moved (they are now at a new path) and not the others."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/a.txt", b"first")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/b.txt", b"second")
    w.run()                                                          # names unknown: the folder goes as named
    a = md5_at(w, "Projects/Brennick - 77Q/Files/Ana Rivera/a.txt")
    b = md5_at(w, "Projects/Brennick - 77Q/Files/Ana Rivera/b.txt")
    w.payroll()
    monkeypatch.setattr(Mirror, "_recheck_folders", lambda *a: None)  # names learnt, nothing renamed yet
    monkeypatch.setattr(Mirror, "_recheck", lambda *a: None)
    w.run()                                                          # (the token-named folder is made, empty)
    monkeypatch.undo()
    files_dir = w.d.path(w.tok, "Projects/Brennick - 77Q/Files")
    assert len([c for c in w.d.children(files_dir) if c["mimeType"] == FakeDrive.FOLDER]) == 2
    assert purge_entries(w) == []
    real, moves = w.d.move, []

    def move_once(fid, new_parent, old_parent, name=None):
        moves.append(fid)
        if len(moves) > 1:
            raise OSError("Drive said no")
        return real(fid, new_parent, old_parent, name)
    w.d.move = move_once
    rep = w.run()
    assert any("not renamed" in x for x in rep.warnings), rep.warnings
    moved = w.d.items[moves[0]]["data"]
    assert ({a, b} & purged(w)) == {a if moved == b"first" else b}  # the moved one only
    assert_no_path_in_list(w)
    w.d.move = real
    w.run()                                                          # finished next run
    assert {a, b} <= purged(w) and no_name_on_aws(w)


def test_purge_entries_expire_after_35_days_and_the_list_is_one_of_the_tokenizers_own_files():
    from datetime import datetime, timedelta, timezone
    from pii_tokenizer import mirror
    assert mirror.PURGE_KEEP_DAYS == 35
    t0 = datetime.now(timezone.utc)
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    run_at(w, t0)
    w.payroll()
    run_at(w, t0)
    assert len(purge_entries(w)) == 1
    assert screen.PURGE_SUMMARY in screen.OWN_FILES                  # never copied back to staff as a file of AWS's
    assert w.d.read(w.raw, screen.PURGE_SUMMARY) is None
    run_at(w, t0 + timedelta(days=34))
    assert len(purge_entries(w)) == 1                                # still kept
    run_at(w, t0 + timedelta(days=36))
    assert purge_entries(w) == []


def backup_dir():
    from datetime import datetime, timedelta, timezone
    return "_sync_backup/" + (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d-%H-%M-%S") + \
        "__sync_backup/"


def test_a_backup_pdf_holding_a_name_learnt_later_is_binned_and_listed():
    """Issue #22: the AWS-side backups (_sync_backup/) keep files an AWS run replaced. They are never pulled, but
    each was a live file on AWS before, so an old S3 version may hold its content: a PDF there is read again when
    names are learnt, as live copies are, binned when it holds one, and listed."""
    w = World()
    w.run()                                                          # a mirror in use (its first run needs it empty)
    keep = backup_dir() + "Projects/Brennick - 77Q/Uploads/"
    w.d.put(w.tok, keep + "WA20250001.pdf", make_pdf([WD[0] + ["Ana Rivera, laborer"]]))
    w.d.put(w.tok, keep + "WA20250002.pdf", make_pdf([WD[0]]))
    w.run()
    named, clean = md5_at(w, keep + "WA20250001.pdf"), md5_at(w, keep + "WA20250002.pdf")
    assert w.d.read(w.tok, keep + "WA20250001.pdf") is not None      # no name known yet
    w.payroll()
    w.run()
    assert w.d.read(w.tok, keep + "WA20250001.pdf") is None          # now holds a known name: binned
    assert w.d.read(w.tok, keep + "WA20250002.pdf") is not None      # clean: kept
    assert named in purged(w, "content") and clean not in purged(w)
    assert_no_path_in_list(w)


def test_a_backup_whose_path_holds_a_name_learnt_later_is_binned_and_listed():
    w = World()
    w.run()
    bk = backup_dir() + "Projects/Brennick - 77Q/"
    w.d.put(w.tok, bk + "Files/Ana Rivera/notes.txt", b"in a named folder")
    w.d.put(w.tok, bk + "Uploads/Lee Mason note.txt", b"a named file")
    w.d.put(w.tok, bk + "Uploads/note.txt", b"neither")
    w.run()
    md5s = {rel: md5_at(w, bk + rel) for rel in ("Files/Ana Rivera/notes.txt", "Uploads/Lee Mason note.txt",
                                                 "Uploads/note.txt")}
    w.payroll()
    w.run()
    assert w.d.read(w.tok, bk + "Files/Ana Rivera/notes.txt") is None
    assert w.d.read(w.tok, bk + "Uploads/Lee Mason note.txt") is None
    assert md5s["Files/Ana Rivera/notes.txt"] in purged(w, "path")  # binned with its folder
    assert md5s["Uploads/Lee Mason note.txt"] in purged(w, "path")
    assert md5s["Uploads/note.txt"] not in purged(w)
    assert_no_path_in_list(w)


def test_a_copy_renamed_in_the_copy_step_for_a_name_learnt_in_the_same_run_is_listed():
    """Staff edit a file named after a worker in the run that learns the name: the copy step updates the copy and
    gives it the token name at once (so _recheck has nothing left to rename). Its old path's content is listed."""
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked")
    w.run()                                                          # name unknown: copied as it is
    old = md5_at(w, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt")
    w.payroll()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Uploads/Lee Mason note.txt", b"checked again")   # edited, same run
    w.run()
    assert no_name_on_aws(w)
    assert any(c[0] == "update" for c in w.d.calls)                  # updated (and renamed) by the copy step
    assert old in purged(w, "path")
    assert_no_path_in_list(w)


def test_an_empty_file_is_never_listed_and_content_wins_over_path():
    import hashlib
    w = World()
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/blank.txt", b"")
    w.d.put(w.raw, "Projects/Brennick - 77Q/Files/Ana Rivera/notes.txt", b"hours checked")
    w.run()
    w.payroll()
    w.run()                                                          # the folder renamed: both files "path"
    assert hashlib.md5(b"").hexdigest() not in purged(w)
    notes = hashlib.md5(b"hours checked").hexdigest()
    assert notes in purged(w, "path")
    m = Mirror(w.d, w.ctx, w.state, lambda: None, lambda: None)
    m._st = w.state["mirrors"]["compliance"]
    m._stale("content", notes)
    m._stale("path", notes)                                          # recorded both ways: content wins
    m._stale("content", hashlib.md5(b"").hexdigest())
    entries = [e for e in m._st["purge"] if e["md5"] == notes]
    assert [e["kind"] for e in entries] == ["content"]
    assert all(e["md5"] != hashlib.md5(b"").hexdigest() for e in m._st["purge"])
