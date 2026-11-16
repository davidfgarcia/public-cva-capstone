"""Upload checks (agent/upload_checks.py; item 1 of the reliability plan, #65) and whole-number tracker IDs
(intent.strict_ids, #175). End to end they are measured by tools/run_evals.py (Eval I, J, K, and the tracker's
"88000O8" / "8800040A" lines); these tests pin the pieces. Invented data only."""
import copy
import datetime as dt
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
for p in (AGENT, HERE):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

import test_seeded_violations as SV                                         # noqa: E402
from agent import events, intent, upload_checks as U                       # noqa: E402
from agent.payroll import RULE_TITLES, SUSPECT_KEYS, PayrollSkipped, review_rows  # noqa: E402

TODAY = dt.date(2026, 11, 14)
ON = {**U.DEFAULTS, "enabled": True}
HEADER = list(SV.HEADER)
COLS = SV.COLS


def cfg_with(enabled=True, **extra):
    c = copy.deepcopy(SV.CFG)
    c["payroll"]["upload_checks"] = {"enabled": enabled, **extra}
    return c


def rows_of(*lines, first=2):
    return [(first + i, list(v)) for i, v in enumerate(lines)]


# --------------------------------------------------------------------------------------------- the setting
def test_shipped_on_and_on_only_when_true():
    assert U.settings(SV.CFG) is not None                      # the repo's config.yaml: on (Nov 15, 2026)
    assert U.settings(cfg_with(False)) is None
    assert U.settings(cfg_with("yes")) is None                 # only a real true turns it on
    s = U.settings(cfg_with(True, future_days=3))
    assert s["future_days"] == 3 and s["max_week_age_years"] == 5
    assert SV.CFG["intent"].get("strict_ids") is True


def test_s15_is_a_known_rule_and_its_reasons_are_run_log_codes():
    assert RULE_TITLES["S15"] == "Not a usable number or date" and SUSPECT_KEYS["S15"] == "S15_unusable_value"
    assert SV.CFG["payroll"]["suspect_checks"]["S15_unusable_value"] is True
    assert {"no_payroll_lines", "repeated_header"} <= events.REASONS
    assert PayrollSkipped("x").code == "wage_type_unknown"
    assert PayrollSkipped("x", code="repeated_header").code == "repeated_header"


# --------------------------------------------------------------------------------------------- values
@pytest.mark.parametrize("v", [8, 8.5, 0, -2, "8", "8.5", "8.", " 40 ", "$40.00", "$ 1,400.00", "1,400", "8 hrs",
                               "8hrs", "8 Hours", "8 hrs.", "40.00/hr", "40 / hour", "-12", "-$12", "$-12"])
def test_numbers_the_review_reads_as_shown(v):
    from agent.util import to_float
    assert U.is_number(v)
    assert isinstance(v, (int, float)) or to_float(v) is not None


@pytest.mark.parametrize("v", ["eight", "see attached", "8 days", "1,40", "4O", "12-", "12.50-", "$", "hrs", True, "x8",
                               "8 8", "1.2.3", "1 400.00", ".5", "(12.50)", "- $12", dt.time(8, 0)])
def test_text_the_review_would_misread_is_not_a_number(v):
    """Independent review, Nov 13: "(12.50)", ".5" or "12.50-" would be read as nothing or with the wrong sign."""
    assert not U.is_number(v)


@pytest.mark.parametrize("v", ["-", "–", "—", "$ -", "$-", "-0-", "N/A", "n/a", "NA", " na ", "None"])
def test_words_for_nothing_are_left_alone(v):
    """Independent review, Nov 13: an export with "N/A" in Per Diem on every line must not lose every line's checks."""
    line = SV.line(**{"Total DT Wages": v})
    assert U.says_empty(v) and U.line_problems(line, COLS, HEADER, TODAY, ON) == []


def test_a_line_with_bad_values_names_the_column_and_never_the_text():
    line = SV.line(**{"ST Mon": "Qzvorak, Evaltest 900", "Base Hourly": "see attached"})
    got = U.line_problems(line, COLS, HEADER, TODAY, ON)
    assert sorted(got) == [("base", "Base Hourly isn't a number"), ("st_mon", "ST Mon isn't a number")]
    assert not any("Qzvorak" in w or "attached" in w for _, w in got)


def test_week_end_dates_not_a_date_or_out_of_range():
    p = lambda week: U.line_problems(SV.line(**{"Week End Date": week}), COLS, HEADER, TODAY, ON)
    assert p("Week 37") == [("week", "Week End Date isn't a date")]
    assert p(46000) == [("week", "Week End Date isn't a date")]            # a bare number, not a date
    assert p(dt.date(2062, 9, 5)) == [("week", "Week End Date 09-05-2062 is after the run's date (11-14-2026): "
                                               "check the year")]
    assert "more than 5 years before" in p(dt.date(2019, 9, 7))[0][1]
    for ok in (dt.date(2026, 11, 20), dt.date(2021, 11, 15), dt.datetime(2026, 9, 5), "9/5/2026", "2026-09-05",
               None, ""):
        assert p(ok) == [], ok                                              # in range, readable, or empty (S08's)
    assert p(dt.date(2026, 11, 21)) != []                                  # 7 days ahead: out (future_days 6)


def test_the_range_follows_the_settings():
    s = {**ON, "future_days": 30, "max_week_age_years": 1}
    p = lambda week: U.line_problems(SV.line(week=week), COLS, HEADER, TODAY, s)
    assert p(dt.date(2026, 12, 10)) == [] and p(dt.date(2025, 9, 6)) != []
    assert U._years_before(dt.date(2028, 2, 29), 5) == dt.date(2023, 2, 28)


# --------------------------------------------------------------------------------------------- the file
def test_no_lines_below_the_header_skips_the_file():
    totals = ["Totals"] + [None] * (len(HEADER) - 1)
    for rows in ([], rows_of([None] * len(HEADER)), rows_of([None] * len(HEADER), totals)):
        assert U.file_problem(rows, COLS, HEADER) == ("no_payroll_lines", "the payroll has no lines below its header row")


def test_the_header_again_skips_the_file_even_with_its_name_cells_tokenized():
    again = list(HEADER)
    again[HEADER.index("Employee Name")] = "EMP-0a1b2c3d4e"            # as AWS sees it: the cell is tokenized
    again[HEADER.index("Employee ID (LCPtracker)")] = "EID-0a1b2c3d4e"
    rows = rows_of(SV.line(), SV.line(emp="b"), again, SV.line(emp="c"))
    code, words = U.file_problem(rows, COLS, HEADER)
    assert code == "repeated_header" and "row 4" in words and "upload each export as its own file" in words


def test_the_header_again_in_another_column_order_is_found_too():
    """Independent review, Nov 13: the second export may have its columns in another order or one more column."""
    again = list(reversed(HEADER)) + ["Notes"]
    rows = rows_of(SV.line(), again, SV.line(emp="c"))
    assert U.file_problem(rows, COLS, HEADER)[0] == "repeated_header"


def test_a_line_with_a_few_header_words_is_not_a_second_header():
    line = SV.line(craft="Craft", cls="Classification")              # odd, but a line: two labels only
    assert U.file_problem(rows_of(SV.line(), line), COLS, HEADER) is None


def test_line_marks_skip_empty_and_footer_lines():
    totals = ["Totals"] + [None] * 28 + ["n/a"] + [None] * (len(HEADER) - 30)
    rows = rows_of(SV.line(), [None] * len(HEADER), totals, SV.line(**{"OT Hours": "two"}))
    assert U.line_marks(rows, COLS, HEADER, TODAY, ON) == {5: [("ot_hours", "OT Hours isn't a number")]}


# --------------------------------------------------------------------------------------------- the review
def test_a_marked_line_gets_s15_only_and_the_others_are_checked_as_before():
    bad = SV.line(**{"Base Hourly": "see attached"})
    short = SV.line(emp="b", base=35.0, **{"Employee ID (LCPtracker)": "EID-1b2c3d4e5f"})   # R08 as before
    rows = rows_of(bad, short)
    before = review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha")
    assert {i.rule for i in before if i.row == 2} >= {"R08"}              # today: base read as 0.00
    marks = U.line_marks(rows, COLS, HEADER, TODAY, ON)
    after = review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha", unusable=marks)
    assert [(i.rule, i.row, i.col, i.severity) for i in after if i.row == 2] == [("S15", 2, "base", "orange")]
    assert "Base Hourly isn't a number" in next(i.detail for i in after if i.row == 2)
    assert sorted(i.rule for i in after if i.row == 3) == sorted(i.rule for i in before if i.row == 3)


def test_a_marked_line_is_left_out_of_the_weekly_hours_too():
    """By design (David, Nov 13: "left out of every other check"): its hours don't count toward the worker's week."""
    rows = rows_of(SV.line(**{"Base Hourly": "see attached"}), SV.line())          # one worker, 80 hours
    assert "S05" in {i.rule for i in review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha")}
    marks = U.line_marks(rows, COLS, HEADER, TODAY, ON)
    assert "S05" not in {i.rule for i in review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha", unusable=marks)}


def test_a_week_with_a_line_left_out_skips_the_federal_state_comparison():
    """Independent review, Nov 13: a worker's federal and state lines of one week are compared (S10); with one state
    line left out, the comparison would be wrong, so that week isn't compared."""
    fed = "Federal Davis-Bacon - Pierce"
    rows = rows_of(SV.line(jur=fed), SV.line(jur=fed, st=(4, 4, 0, 0, 0, 0, 0)), SV.line(),
                   SV.line(st=(4, 4, 0, 0, 0, 0, 0), **{"Base Hourly": "see attached"}))
    assert "S10" not in {i.rule for i in review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha")}  # agree today
    marks = U.line_marks(rows, COLS, HEADER, TODAY, ON)
    got = review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha", unusable=marks)
    assert "S10" not in {i.rule for i in got} and [i.rule for i in got if i.row == 5] == ["S15"]


def test_s15_can_be_turned_off_like_the_other_suspect_checks():
    c = copy.deepcopy(SV.CFG)
    c["payroll"]["suspect_checks"]["S15_unusable_value"] = False
    rows = rows_of(SV.line(**{"ST Mon": "eight"}))
    marks = U.line_marks(rows, COLS, HEADER, TODAY, ON)
    assert review_rows(rows, COLS, SV.BOOK, c, project="Alpha", unusable=marks) == []   # still left out


def test_without_marks_the_review_is_unchanged():
    rows = rows_of(*[r for _, _, rs, _, _ in SV.CASES for r in rs])
    a = review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha")
    b = review_rows(rows, COLS, SV.BOOK, SV.CFG, project="Alpha", unusable={})
    assert [(i.rule, i.row, i.col, i.detail) for i in a] == [(i.rule, i.row, i.col, i.detail) for i in b]


# --------------------------------------------------------------------------------------------- #175: tracker IDs
def test_tracker_ids_strict_only_with_the_setting():
    cells = {"88000O8": "88000", "8800040A": "8800040", "1234567 (pending)": "1234567", "8800040": "8800040",
             8800040: "8800040", 8800040.0: "8800040", "#8800040": None, " 8800040.0 ": "8800040", "x": None,
             None: None, True: None, 0: None}
    for v, lenient in cells.items():
        assert intent._cell_id(v) == lenient, v
    # a letter or other character inside the number is "not a valid number"; a note after it (a space or a
    # separator first) is read by the number, as staff write them (David, Nov 14, 2026, after the trial run)
    strict = {"88000O8": None, "8800040A": None, "8800040Filed": None, "8,800,040": None, "8800,040": None, "88000,40": None, "8800040.5": None,
              "8800050-Filed": "8800050", "1234567 (pending)": "1234567", "8800040 / 8800041": "8800040",
              "8800040-1": "8800040", "8800040, 8800041": "8800040", "8800040,8800041": "8800040",
              "8800040.": "8800040", "8800040\u2013Filed": "8800040", "8800040\u2014x": "8800040",
              "8800040*": "8800040", "8800040": "8800040", 8800040: "8800040", 8800040.0: "8800040",
              "#8800040": "8800040", " 8800040.0 ": "8800040", "8800040.0-Filed": "8800040", 8800040.5: None,
              "0": None}
    with intent._StrictIds({"intent": {"strict_ids": True}}):
        for v, want in strict.items():
            assert intent._cell_id(v) == want, v
    assert intent._cell_id("8800040A") == "8800040"                        # back to lenient after the block
    with intent._StrictIds({"intent": {"strict_ids": False}}):
        assert intent._cell_id("8800040A") == "8800040"


def _strict_lni():
    ii = {"8800060": {"_id": "8800060", "_project_id": "9", "_contractor": "Qzvorak Drain Works", "_status": None,
                      "_intent_id": "8800060"},
          "8800070": {"_id": "8800070", "_project_id": "9", "_contractor": "Qzvorak Shotcrete", "_status": None,
                      "_intent_id": "8800070"}}
    aa = {"8810070": {"_id": "8810070", "_project_id": "9", "_contractor": None, "_status": None,
                      "_intent_id": "8800070", "_received": "2026-10-01"}}
    return {"ok": True, "intents": ii, "affidavits": aa, "intents_by_name": list(ii.values()),
            "affidavits_by_intent": intent._by_intent(aa.values()), "project_ids": {"9"}, "error": None, "note": None}


def test_a_note_after_a_correct_id_is_no_finding_and_a_bad_id_is_one_finding():
    """The trial run (Nov 14, 2026): "8800050-Filed" (intent) and "8810050-Filed" (affidavit) were correct IDs with a
    note; with "any extra text" each got "not a valid number" plus a second finding saying the cell was empty ("tracker
    has none", "add it to the tracker"). Now the note is read past, and an ID that is really unusable gets one finding."""
    cols = {"status": 0, "contractor": 1, "intent_id": 2, "affidavit_id": 3}
    cfg = copy.deepcopy(SV.CFG)
    cfg["intent"]["strict_ids"] = True
    lni = _strict_lni()

    def issues(row):
        return [(role, issue) for _rn, role, _v, _l, issue in intent.validate([(5, row)], cols, lni, cfg)]

    assert issues(["Onsite", "Qzvorak Drain Works", "8800060-Filed", None]) == []
    assert issues(["Completed Work", "Qzvorak Shotcrete", "8800070", "8810070-Filed"]) == []
    # one finding per cell, naming the number L&I has (not a second one saying the cell is empty)
    assert issues(["Onsite", "Qzvorak Drain Works", "88000G0", None]) == \
        [("intent_id", "Intent ID is not a valid number; L&I has intent 8800060 for this contractor")]
    assert issues(["Completed Work", "Qzvorak Shotcrete", "8800070", "88100T0"]) == \
        [("affidavit_id", "Affidavit ID is not a valid number; L&I has affidavit 8810070 (intent 8800070)")]
    assert issues(["Onsite", "Qzvorak Nobody Else", "88000O8", None]) == [("intent_id", "Intent ID is not a valid number")]
    # an empty cell still gets the finding that says what L&I has
    assert ("intent_id", "L&I has intent 8800060 for this contractor; tracker has none") in \
        issues(["Onsite", "Qzvorak Drain Works", None, None])
    assert any("Missing affidavit found" in i for _r, i in issues(["Completed Work", "Qzvorak Shotcrete", "8800070", None]))


def test_lni_records_are_read_as_before_with_strict_ids():
    with intent._StrictIds({"intent": {"strict_ids": True}}):
        assert intent._num_id("8800040.0") == "8800040" and intent._num_id(8800040) == "8800040"


# --------------------------------------------------------------------------------------------- staff words
@pytest.mark.parametrize("line,short,words", [
    ("payroll: SKIPPED - the payroll has no lines below its header row; the payroll was not checked and stays in "
     "Uploads", "No payroll lines", "The payroll has no lines below its header row, so it wasn't checked; it stays "
     "in the Uploads folder. Export the payroll again with its lines and upload it."),
    ("payroll: SKIPPED - the header row appears again at row 4 (two exports in one file?): upload each export as its "
     "own file; the payroll was not checked and stays in Uploads", "Two exports in one file",
     "The payroll's header row appears again at row 4, as if two exports were pasted into one sheet, so it wasn't "
     "checked; it stays in the Uploads folder. Upload each export as its own file."),
])
def test_the_staff_email_says_what_to_do_in_plain_words(line, short, words):
    from agent.staff_email import ACTION_SKIPS, task_rows
    details = {"payroll": {"status": "skipped", "skipped": [line], "errors": [], "warnings": [], "crashed": False}}
    row = next(r for r in task_rows([line], details, project="Alpha") if r.key == "payroll")
    assert row.short() == short and row.reason() == words and short in ACTION_SKIPS
