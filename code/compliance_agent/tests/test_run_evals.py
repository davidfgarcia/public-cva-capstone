"""The evaluation harness (tools/run_evals.py, tools/eval_cases.py): the invented set, its answer key, the scoring, and
one end-to-end run of this checkout (skipped where the tokenizer can't run, as the compare_runs end-to-end test is).

    python -m pytest -q tests/test_run_evals.py        (or: python tests/test_run_evals.py)

Every name, number and company here is invented.
"""
import base64
import csv
import gzip
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest
from openpyxl.comments import Comment

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
sys.path.insert(0, str(AGENT))
sys.path.insert(0, str(AGENT / "tools"))
sys.path.insert(0, str(HERE))

import eval_cases as E                                       # noqa: E402
import eval_lni as L                                         # noqa: E402
import run_evals as R                                        # noqa: E402
import test_seeded_violations as SV                          # noqa: E402

SEEDED = "Eval A - Seeded rules/Week 37.xlsx"
TRACKER = "Eval G - Tracker checks"
NOT_APPLICABLE = "Eval H - No L&I filing"


def _tokenizer_runs() -> bool:
    probe = subprocess.run([sys.executable, "-c", "import pii_tokenizer.files"], cwd=AGENT.parent / "tokenizer",
                           capture_output=True)
    return probe.returncode == 0


@pytest.fixture(scope="module")
def run(tmp_path_factory):
    """One end-to-end run of this checkout on the invented set: (kept folder, key, scores, exit code)."""
    if not _tokenizer_runs():
        pytest.skip("the tokenizer can't run with this Python")
    base = tmp_path_factory.mktemp("eval")
    keep = base / "kept"
    summary = os.environ.pop("GITHUB_STEP_SUMMARY", None)         # the CI step adds the report once, not this test
    try:
        code = R.main(["--keep", str(keep), "--json", str(base / "evals.json")])
    finally:
        if summary is not None:
            os.environ["GITHUB_STEP_SUMMARY"] = summary
    if code == 3:
        pytest.fail("the end-to-end run itself failed (its output is printed above)")
    key = json.loads((keep / "snapshot" / "answer_key.json").read_text())
    return keep, key, json.loads((base / "evals.json").read_text()), code


def _rows(f: Path) -> dict:
    """{row number: values} of an upload's first sheet (or CSV lines)."""
    if f.suffix == ".csv":
        with f.open(encoding="utf-8", newline="") as h:
            return {i: r for i, r in enumerate(csv.reader(h), 1)}
    ws = openpyxl.load_workbook(f).worksheets[0]
    return {r[0].row: [c.value for c in r] for r in ws.iter_rows()}


# ------------------------------------------------------------------------------------------- the set and its key
def test_the_key_covers_every_seeded_case_and_each_scenario(tmp_path):
    key = E.build(tmp_path / "snap")
    seeded = {e["case"] for e in key["files"][SEEDED]["lines"].values()}
    assert {c[0] for c in SV.CASES} <= seeded                       # one set: the acceptance test's cases, end to end
    outcomes = {f: v["outcome"] for f, v in key["files"].items()}
    assert sum(o == "skipped" for o in outcomes.values()) >= 2
    assert all(v["reason_contains"] for v in key["files"].values() if v["outcome"] == "skipped")
    assert all(v["lines"] for v in key["files"].values() if v["outcome"] == "reviewed")
    assert [(g["issue"], g["currently"]) for g in key["known_gaps"]] == [(67, "not recognized")]
    assert all(f in key["files"] for g in key["known_gaps"] for f in g["files"])
    assert {t: v["outcome"] for t, v in key["trackers"].items()} == {TRACKER: "checked",
                                                                     NOT_APPLICABLE: "not applicable"}


def test_lines_are_scored_one_by_one_except_a_duplicate(tmp_path):
    """Review finding: scoring a case's lines together let a flag sit on the wrong (correct) line. Only the duplicate's
    two lines share a group; the second line of "Base rate missing" is a control of its own."""
    key = E.build(tmp_path / "snap")
    groups = {}
    for row, e in key["files"][SEEDED]["lines"].items():
        groups.setdefault(e["group"], []).append(int(row))
    assert [g for g, rows in groups.items() if len(rows) > 1] == ["Duplicate line"]
    second = [e for e in key["files"][SEEDED]["lines"].values() if e["group"] == "Base rate missing #2"]
    assert second and second[0]["must"] == [] and second[0]["may"] == []


def test_the_key_points_at_the_lines_it_describes(tmp_path):
    """Every row the key lists holds an invented worker, and every invented worker's row is listed: the key and the
    files can't drift apart (a title row or a header shift would show here)."""
    snap = tmp_path / "snap"
    key = E.build(snap)
    for rel, exp in key["files"].items():
        project, name = rel.split("/", 1)
        rows = _rows(snap / "Projects" / project / "Uploads" / name)
        workers = {i for i, r in rows.items() if any(str(v).startswith(E.SURNAME + ",") for v in r if v)}
        if exp["outcome"] == "reviewed":
            assert workers == {int(r) for r in exp["lines"]}, rel
        assert "classification" in [str(v).strip().lower() for v in rows[exp["header_row"]] if v], rel


def test_only_invented_people_in_the_set(tmp_path):
    """HANDOFF section 6: the evaluation set is invented. In every upload (xlsx and CSV), every Employee Name is a
    'Qzvorak, Evaltest nnn' and every employee ID a 7319nnn."""
    snap = tmp_path / "snap"
    E.build(snap)
    uploads = [f for f in (snap / "Projects").rglob("Uploads/*") if f.suffix in (".xlsx", ".csv")
               and "Wage rates" not in f.name]
    assert len(uploads) == 14
    for f in uploads:
        rows = list(_rows(f).values())
        h = next(i for i, r in enumerate(rows) if "Employee Name" in r)
        name, eid = rows[h].index("Employee Name"), rows[h].index("Employee ID (LCPtracker)")
        for r in rows[h + 1:]:
            if r == rows[h] or all(x in (None, "") for x in r) or r[0] == "Totals":
                continue                        # Eval K's second header; Eval I's empty and Totals footer rows
            assert str(r[name]).startswith(f"{E.SURNAME}, {E.GIVEN} "), f.name
            assert str(r[eid]).startswith(str(E.ID_BASE // 1000)) and len(str(r[eid])) == 7, f.name


def test_the_never_learnt_name_is_in_every_place_and_in_no_name_column(tmp_path):
    """#53's case (Eval L): the invented name is in each of the seven places, and in no column the tokenizer learns
    names from (it is never learnt: only the fail-closed settings keep it off the AWS side)."""
    snap = tmp_path / "snap"
    E.build(snap)
    p = snap / "Projects" / "Eval L - Names never learnt"
    assert any(E.NEVER_LEARNT in f.name for f in (p / "Uploads").iterdir())
    assert any(E.NEVER_LEARNT in d.name for d in (p / "Files").iterdir() if d.is_dir())
    wb = openpyxl.load_workbook(p / "Uploads" / "Week 41.xlsx")
    assert any(E.NEVER_LEARNT in t for t in wb.sheetnames) and wb.properties.creator == E.NEVER_LEARNT
    rows = [[c.value for c in r] for r in wb.worksheets[0].iter_rows()]
    name = rows[0].index("Employee Name")
    assert "Foreman" in rows[0] and not any(E.NEVER_LEARNT in str(r[name]) for r in rows[1:])
    assert R.WORKER.search(E.NEVER_LEARNT)                     # the privacy check searches for it too


def test_two_uploads_with_one_stem_are_refused(tmp_path):
    """Their reviewed workbooks would have the same name."""
    b = E.Builder(tmp_path)
    p = b.project("Eval X")
    b.expect(p, "Week 1.xlsx", "reviewed")
    with pytest.raises(ValueError):
        b.expect(p, "Week 1.csv", "reviewed")


# ---------------------------------------------------------------------------------------------------- scoring
def _copy(run, tmp_path, events=False):
    """A copy of the run to plant a fault in. Without ``events``, the structured run log is taken out (the copy's
    config.yaml says it was off: "not measured"): a fault planted in a workbook after the run would also make the log
    disagree, counting twice. (test_the_run_log_is_checked_against_the_run plants faults in the log itself.)"""
    import yaml
    keep, key, _, _ = run
    dst = tmp_path / "run"
    shutil.copytree(keep / "run", dst)
    if not events:
        for f in (dst / "results" / "run_logs").glob("*__events.jsonl"):
            f.unlink()
        conf = dst / "results" / "config.yaml"
        cfg = yaml.safe_load(conf.read_text())
        cfg["run_log"] = {"events": False}
        conf.write_text(yaml.safe_dump(cfg, sort_keys=False))
    return dst, json.loads(json.dumps(key))


def _reviewed(run_dir: Path, rel: str) -> Path:
    project, name = rel.split("/", 1)
    return R.reviewed_for(run_dir / "results" / "Projects" / project, name, " - REVIEWED")


def _sheet(wb):
    return [w for w in wb.worksheets if w.title not in R.AGENT_SHEETS][0]


def _row_of(key, rel, case):
    return next(int(r) for r, e in key["files"][rel]["lines"].items() if e["case"] == case)


def _summary(run_dir: Path) -> Path:
    return next((run_dir / "results" / "run_logs").glob("*__run_summary.json"))


def test_this_checkout_matches_the_key(run):
    _, key, s, code = run
    t = s["totals"]
    assert code == 0, R.report(s)
    assert s["health"] == [] and s["worker_names_on_aws_side"] == []
    assert t["seeded_cases_detected"] == t["seeded_cases"] > 30
    assert t["control_cases_clean"] == t["control_cases"] > 30
    assert t["false_flags"] == 0 and t["flag_precision"] == 1.0
    assert t["files_as_expected"] == t["files"] - 1 and t["files_known_gap"] == 1
    assert [g["status"] for g in s["known_gaps"]] == ["still open"]                  # #67
    assert t["trackers_as_expected"] == t["trackers"] == 2
    assert t["tracker_findings_detected"] == t["tracker_findings_seeded"] >= 12
    assert t["tracker_controls_clean"] == t["tracker_controls"] and t["tracker_not_due_lines"] == 4
    assert t["tracker_false_marks"] == 0 and t["tracker_lines_known_gap"] == 0
    assert s["lni_asked_wrongly"] == [] and s["lni_requests"] > 10
    assert s["run_log"]["measured"] and s["run_log"]["problems"] == []
    assert s["run_log"]["events"]["run_start"] == s["run_log"]["events"]["run_end"] == 1
    assert s["run_log"]["events"]["file"] == t["files"]          # every payroll upload, the unrecognized one too


def test_a_missed_flag_fails_the_run(run, tmp_path):
    d, key = _copy(run, tmp_path)
    row = _row_of(key, SEEDED, "Overtime rate below 1.5x base")
    f = _reviewed(d, SEEDED)
    wb = openpyxl.load_workbook(f)
    for c in _sheet(wb)[row]:
        c.comment = None
    wb.save(f)
    s = R.score(d / "results", key)
    assert s["failures"] == 1
    (bad,) = [c for c in s["cases"] if c["missed"] and not c["known_gap"]]
    assert bad["missed"] == ["R03"] and bad["rows"] == [row]
    assert s["per_rule"]["R03"]["detected"] == s["per_rule"]["R03"]["seeded"] - 1
    assert "R03" in R.report(s) and "FAIL" in R.report(s)


def test_a_flag_on_a_clean_line_is_a_false_flag(run, tmp_path):
    d, key = _copy(run, tmp_path)
    rel = "Eval B - Clean weeks/Week ending 09-06.xlsx"
    row = int(next(iter(key["files"][rel]["lines"])))
    f = _reviewed(d, rel)
    wb = openpyxl.load_workbook(f)
    _sheet(wb).cell(row, 3).comment = Comment("[SUSPECTED] S09 Duplicate payroll line: planted", "test")
    _sheet(wb).cell(1, 1).comment = Comment("[RULE BROKEN] R01 planted on the header", "test")
    wb.save(f)
    s = R.score(d / "results", key)
    assert s["totals"]["false_flags"] == 2 and s["failures"] == 2
    assert s["per_rule"]["S09"]["false_flags"] == 1 and s["per_rule"]["R01"]["false_flags"] == 1
    assert any(c["case"] == R.UNLISTED and c["rows"] == [1] for c in s["cases"])


def test_a_flag_moved_to_the_cases_correct_line_is_caught(run, tmp_path):
    """Review finding 3: "Base rate missing" has a seeded line and a correct line; S08 moved from the first to the
    second (with a false R08 there) must fail: the first misses S08, the second is a control with flags."""
    d, key = _copy(run, tmp_path)
    first = _row_of(key, SEEDED, "Base rate missing")
    f = _reviewed(d, SEEDED)
    wb = openpyxl.load_workbook(f)
    ws = _sheet(wb)
    for c in ws[first]:
        c.comment = None
    ws.cell(first + 1, 30).comment = Comment("[SUSPECTED] S08 Missing or negative value: planted\n"
                                             "[RULE BROKEN] R08 Base / total prevailing rate not met: planted", "t")
    wb.save(f)
    s = R.score(d / "results", key)
    missed = [c for c in s["cases"] if c["missed"] and not c["known_gap"]]
    extra = [c for c in s["cases"] if c["extra"]]
    assert [c["rows"] for c in missed] == [[first]] and missed[0]["missed"] == ["S08"]
    assert [c["rows"] for c in extra] == [[first + 1]] and extra[0]["extra"] == ["R08", "S08"]


def test_a_county_only_line_checked_only_as_state_is_caught(run, tmp_path):
    """Review finding 4: Eval E tells the passes apart. Notes on the county-only laborer that say "Checked as state"
    (not union) fail the words check; a run that raises nothing on it misses its R07/R08."""
    d, key = _copy(run, tmp_path)
    rel = "Eval E - Jurisdiction unclear/Week 35.xlsx"
    row = _row_of(key, rel, E.county_only_cases()[0][0])
    f = _reviewed(d, rel)
    wb = openpyxl.load_workbook(f)
    ws = _sheet(wb)
    cells = [c for c in ws[row] if c.comment is not None]
    assert cells and all("Checked as union" in c.comment.text for c in cells)
    for c in cells:
        c.comment = Comment(c.comment.text.replace("Checked as union", "Checked as state"), "t")
    wb.save(f)
    s = R.score(d / "results", key)
    (bad,) = [c for c in s["cases"] if c["notes"]]
    assert bad["rows"] == [row] and any("Checked as union" in n for n in bad["notes"])
    assert any("Checked as state" in n for n in bad["notes"]) and s["failures"] == 1
    for c in cells:
        c.comment = None
    wb.save(f)
    s = R.score(d / "results", key)
    (bad,) = [c for c in s["cases"] if c["missed"] and not c["known_gap"]]
    assert bad["missed"] == ["R07", "R08"]


def test_a_skip_for_another_reason_is_a_difference(run, tmp_path):
    d, key = _copy(run, tmp_path)
    key["files"]["Eval D - No wage rates/Week 36.xlsx"]["reason_contains"] = "a reason the agent never gives"
    s = R.score(d / "results", key)
    (bad,) = [f for f in s["files"] if not f["ok"] and not f["known_gap"]]
    assert bad["file"].startswith("Eval D") and bad["got"] == "skipped"
    assert "should say" in R.report(s)


def test_a_skip_reason_missing_from_the_admin_email_is_a_difference(run, tmp_path):
    import email
    from email import policy
    d, key = _copy(run, tmp_path)
    eml = next((d / "results" / "run_logs").glob("*__email_preview_admin.eml"))
    assert "no wage rates file in the project" in R.eml_text(eml)
    msg = email.message_from_bytes(eml.read_bytes(), policy=policy.default)
    for part in msg.walk():
        if part.get_content_maintype() == "text" and not part.is_multipart():
            part.set_content(part.get_content().replace("no wage rates file in the project", "(removed)"),
                             subtype=part.get_content_subtype())
    eml.write_bytes(msg.as_bytes())
    s = R.score(d / "results", key)
    (bad,) = [f for f in s["files"] if not f["ok"] and not f["known_gap"]]
    assert bad["file"].startswith("Eval D") and not bad["reason_in_admin_email"]
    assert "administrator's email too" in R.report(s)


def test_a_file_reviewed_when_it_should_be_skipped_is_one_difference(run, tmp_path):
    d, key = _copy(run, tmp_path)
    key["files"][SEEDED]["outcome"] = "skipped"
    key["files"][SEEDED]["reason_contains"] = ""
    s = R.score(d / "results", key)
    assert s["failures"] == 1 and not any(c["missed"] or c["extra"] for c in s["cases"]
                                          if c["file"] == SEEDED)


def test_a_crashed_task_or_missing_email_fails_the_run(run, tmp_path):
    """Review finding 1: a run that crashes partway, or writes no email, can't pass."""
    d, key = _copy(run, tmp_path)
    f = _summary(d)
    summary = json.loads(f.read_text())
    project = next(iter(summary["details"]))
    summary["details"][project]["intent"].update(crashed=True, errors=["planted"])
    summary["projects"][project].append("ERROR (intent): planted crash")
    f.write_text(json.dumps(summary))
    s = R.score(d / "results", key)
    assert len(s["health"]) == 2 and s["failures"] == 2 and "crashed" in R.report(s)
    summary["email"] = "not sent: planted"
    f.write_text(json.dumps(summary))
    next((d / "results" / "run_logs").glob("*__email_preview_staff.eml")).unlink()
    s = R.score(d / "results", key)
    assert any("emails were not written" in p for p in s["health"])
    assert any("no staff email preview" in p for p in s["health"])
    f.unlink()
    assert R.score(d / "results", key)["health"] == ["no run summary was written (the agent didn't finish)"]


def test_a_known_gap_is_reported_not_counted_and_its_changes_are_noticed(run, tmp_path):
    d, key = _copy(run, tmp_path)
    s = R.score(d / "results", key)
    gap_cases = [c for c in s["cases"] if c["known_gap"]]
    assert gap_cases and all(c["missed"] for c in gap_cases if c["must"]) and s["failures"] == 0
    gap = next(g for g in key["known_gaps"] if g["issue"] == 67)
    # the gap's file ending another way (here: gone from Uploads, as a held-back file would be) is a difference
    project, name = gap["files"][0].split("/", 1)
    (d / "results" / "Projects" / project / "Uploads" / name).unlink()
    s = R.score(d / "results", key)
    assert s["failures"] == 1 and s["known_gaps"][0]["status"].startswith("changed")
    # pretend #67 is fixed: the key's gap file now expects what the run did, and has no lines
    d, key = _copy(run, tmp_path / "again")
    for f in gap["files"]:
        key["files"][f] = {"outcome": "not recognized", "header_row": 9, "lines": {}, "reason_contains": ""}
    s = R.score(d / "results", key)
    assert s["known_gaps"][0]["status"].startswith("passes now")


def test_a_worker_name_left_anywhere_in_the_run_fails_it(run, tmp_path):
    """The privacy check: the tokenizer must replace every invented worker's name and ID before the agent sees the
    files. Review finding 2: also inside emails (base64 parts and attachments), gzip, UTF-16, folder names, the
    steps' printed output."""
    _, key, s, _ = run
    assert s["totals"]["worker_names_on_aws_side"] == 0
    d, key = _copy(run, tmp_path)
    f = _reviewed(d, "Eval B - Clean weeks/Week ending 08-30.xlsx")
    wb = openpyxl.load_workbook(f)
    _sheet(wb).cell(2, 5).value = f"{E.SURNAME}, {E.GIVEN} 999"   # an untokenized name, as a broken tokenizer leaves it
    wb.save(f)
    logs = d / "results" / "run_logs"
    (logs / "extra.txt").write_text(f"line for {E.GIVEN} 042\n")
    attach = base64.b64encode(f"Employee,ID\nsomeone,{E.ID_BASE + 7}\n".encode()).decode()
    (logs / "planted.eml").write_text(
        "MIME-Version: 1.0\nContent-Type: multipart/mixed; boundary=XX\n\n--XX\nContent-Type: text/plain\n\nhello\n"
        f"--XX\nContent-Type: text/csv\nContent-Transfer-Encoding: base64\nContent-Disposition: attachment; "
        f"filename=a.csv\n\n{attach}\n--XX--\n")
    (logs / "packed.gz").write_bytes(gzip.compress(E.SURNAME.encode()))
    (logs / "wide.txt").write_bytes(E.SURNAME.encode("utf-16-le"))
    (d / "results" / f"{E.SURNAME} folder").mkdir()
    (d / "step_output" / "agent.txt").write_text(f"Traceback ... {E.SURNAME}, {E.GIVEN} 001\n")
    s = R.score(d / "results", key)
    found = " ".join(s["worker_names_on_aws_side"])
    for part in ("Week ending 08-30", "extra.txt", "planted.eml", "packed.gz", "wide.txt", "<name> folder (name)",
                 "step_output/agent.txt"):
        assert part in found, part
    assert s["failures"] == len(s["worker_names_on_aws_side"]) == 7
    assert "still holding an invented worker's name or ID: 7" in R.report(s)
    assert f"{E.GIVEN} 999" not in R.report(s) and E.SURNAME not in R.report(s)   # files named, never the names


def test_a_decimal_is_not_taken_for_an_employee_id():
    """Verification round: "0.7319004" (a number with seven decimals) is not the invented ID 7319004."""
    assert not R._hits(b"<v>0.7319004</v>") and R._hits(b"<v>7319004</v>")


def test_an_unreadable_file_counts_as_a_find(tmp_path):
    """Fail closed: a file the check can't open is reported, not passed over."""
    (tmp_path / "broken.xlsx").write_bytes(b"PK\x03\x04 not really a zip")
    (tmp_path / "fine.txt").write_text("nothing here\n")
    found = R.worker_names_found(tmp_path)
    assert found == ["broken.xlsx (could not be read: BadZipFile)"]


def test_the_reports_numbers_are_the_jsons(run):
    _, _, s, _ = run
    text = R.report(s)
    t = s["totals"]
    assert f"{t['seeded_cases_detected']} of {t['seeded_cases']}" in text
    assert "header-below-row-7 (#67): still open" in text
    assert E.SURNAME not in text                                # the report names cases and rows, never workers


# ------------------------------------------------------------------------------- the intent/affidavit check
AGENT_NOTE = "Compliance Validation Agent"


def _note(text):
    return Comment(f"{AGENT_NOTE}:\n- {text}", AGENT_NOTE)


def test_the_tracker_cases_cover_every_kind_of_finding_and_the_key_matches_the_file(tmp_path):
    """Every finding the check can make (eval_lni.KINDS) is seeded, with correct and not-yet-due lines beside them
    (review finding 4: all of them); the key's rows are the tracker's rows; contractors are invented companies."""
    snap = tmp_path / "snap"
    key = E.build(snap)
    exp = key["trackers"][TRACKER]
    seeded = {m.split(":", 1)[1] for e in exp["lines"].values() for m in e["must"]}
    assert seeded == set(L.KINDS)
    ws = openpyxl.load_workbook(snap / "Projects" / TRACKER / "Tracker" / L.TRACKER_FILE).worksheets[0]
    contractors = {r[0].row: r[1].value for r in ws.iter_rows(min_row=2)}
    assert set(contractors) == {int(r) for r in exp["lines"]}
    assert [c["line"][1] for c in L.CASES] == [contractors[r] for r in sorted(contractors)]
    assert sum(1 for e in exp["lines"].values() if e["not_due"]) == 4 and exp["lni_only"] == ["Foxtrennic Traffic Control"]
    assert exp["summary"]["hiring"] == 1 and exp["summary"]["lni_only"] == 1
    staff = [int(r) for r, e in exp["lines"].items() if e["staff_note"]]
    assert staff and ws.cell(staff[0], L.COLUMN_OF["intent_id"]).comment.text == L.STAFF_NOTE
    assert not R.WORKER.search(json.dumps(key["trackers"]))
    assert all(L.kind_of(t) == k for k, t in [("hiring_mismatch", "Hiring contractor doesn't match L&I intent 8800007"),
                                              ("affidavit_missing", "Missing affidavit found in L&I's records: 8810017"),
                                              ("unknown", "something new")])


def _ask(f, op, i, url=L.LNI_URL, body=None):
    """A request through the stand-in as the proxy, as the agent's live client sends it (urllib: conftest blocks
    requests)."""
    import urllib.error
    import urllib.request
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({"http": f.proxy, "https": f.proxy}))
    data = json.dumps(body if body is not None else {"ServiceName": "PWIAPublicService",
                                                      "UrlData": f"api/Public/{op}/{i}", "RestHttpMethod": "GET"})
    req = urllib.request.Request(url, data=data.encode(), headers={"Content-Type": "application/json"})
    try:
        with opener.open(req, timeout=10) as r:
            return r.status, json.loads(r.read())
    except urllib.error.HTTPError as e:
        return e.code, None


def test_the_stand_in_answers_only_what_the_live_client_sends_and_records_everything():
    """The answers parse with agent/lni_live.py's own readers; an unknown ID is null (as L&I's service answers).
    Review finding 3: any other shape, address or method is refused and recorded; finding 11: it is the run's only
    proxy, so a request to anything else is refused and recorded too."""
    from agent import lni_live
    with L.FakeLookup() as f:
        row = lni_live.intent_row(_ask(f, "GetIntent", 8800002)[1])
        assert row["intent_id"] == 8800002 and row["companyname"] == "Brivantel Paving LLC"
        assert row["_affidavit_ids"] == ["8810002", "8810014"] and row["application_received_date"] == "2026-05-04"
        aff = lni_live.affidavit_row(_ask(f, "GetAffidavit", 8810016)[1])
        assert aff["intent_id"] == 8800015 and aff["status"] == "Approved" and aff["formreceiveddate"] == "2026-09-15"
        assert _ask(f, "GetIntent", 8800099) == (200, None)
        tree = _ask(f, "GetProjectStructureTree", L.PROJECT)[1]
        assert {"8800001", "8800030", "8800031"} <= lni_live._ids(tree, "IntentId", set())
        n = len(f.requests)
        assert _ask(f, "GetIntent", 1, body={"ServiceName": "WRONG", "UrlData": "api/Public/GetIntent/1",
                                             "RestHttpMethod": "GET"})[0] == 400
        assert _ask(f, "GetIntent", 1, body={"ServiceName": "PWIAPublicService", "UrlData": "api/Public/GetIntent/1",
                                             "RestHttpMethod": "GET", "extra": "x"})[0] == 400
        assert _ask(f, "GetIntent", 1, body=[1, 2])[0] == 400
        assert _ask(f, "DownloadAllProjectContractors", 1)[0] == 400
        assert _ask(f, "GetIntent", 1, url=L.LNI_URL.replace("?op=UiGatewayRestfulOperation", ""))[0] == 400
        assert _ask(f, "GetIntent", 1, url="http://example.invalid/anything")[0] == 403
        new = f.requests[n:]
        assert len(new) == 6 and not any(q["answered"] for q in new)
        assert all("headers" in q and "body" in q for q in f.requests if "error" not in q)


def test_the_stand_ins_address_is_the_live_clients_production_address():
    """Verification round: if L&I's address in agent/lni_live.py changes, the stand-in must follow."""
    from agent import lni_live
    assert L.LNI_URL == lni_live.DEFAULT_URL.replace("https://", "http://", 1)


def _tracker(run_dir: Path, project=TRACKER) -> Path:
    return run_dir / "results" / "Projects" / project / "Tracker" / L.TRACKER_FILE


def _tracker_row(key, case, project=TRACKER):
    return next(int(r) for r, e in key["trackers"][project]["lines"].items() if e["case"] == case)


def _bad(s):
    return {ln["row"]: ln for ln in s["tracker_lines"] if ln["missed"] or ln["extra"] or ln["words"]}


def test_a_missed_tracker_mark_and_a_mark_on_a_correct_line_are_caught(run, tmp_path):
    d, key = _copy(run, tmp_path)
    f = _tracker(d)
    wb = openpyxl.load_workbook(f)
    ws = wb.worksheets[0]
    missed = _tracker_row(key, "Affidavit at L&I, not in the tracker")
    for c in ws[missed]:
        c.comment = None
    control = _tracker_row(key, "Subcontractor finished, affidavit on file (control)")
    ws.cell(control, 5).comment = _note("Intent ID not found in L&I records for this project")
    due = _tracker_row(key, "Not started, no intent yet (not due: counted, not marked)")
    ws.cell(due, 5).comment = _note("No intent for this contractor in L&I's records for this project")
    ws.cell(1, 2).comment = Comment(f"{AGENT_NOTE}:\nsomething new, without the dash", AGENT_NOTE)   # finding 6
    ws.cell(control, 2).comment = Comment("a staff member's own note", "Staff")            # not the agent's: ignored
    wb["Project Info"]["A1"].comment = _note("planted on another tab")                    # finding 6
    wb.save(f)
    s = R.score(d / "results", key)
    bad = _bad(s)
    assert bad[missed]["missed"] == ["affidavit_id:affidavit_missing (note)"]
    assert bad[control]["extra"] == ["intent_id:intent_not_found"]
    assert bad[due]["extra"] == ["intent_id:no_intent_at_lni"]
    assert bad[1]["case"] == R.UNLISTED and bad[1]["extra"] == ["contractor:unknown"]
    (g,) = [t for t in s["trackers"] if t["project"] == TRACKER]
    assert any("Project Info!A1" in p for p in g["problems"])
    assert "affidavit_missing" in R.report(s) and "FAIL" in R.report(s)


def test_the_review_log_the_red_fill_and_the_numbers_are_checked(run, tmp_path):
    """Review findings 1, 2, 5: the Review Log's rows (the complete list; a cell that already had a note gets its
    message there only), the red fill, and the record numbers staff copy."""
    d, key = _copy(run, tmp_path)
    f = _tracker(d)
    wb = openpyxl.load_workbook(f)
    ws, log = wb.worksheets[0], wb["Review Log"]
    staff = _tracker_row(key, "Intent ID not at L&I, the cell already has a staff note (Review Log only)")
    hiring = _tracker_row(key, "Hiring contractor isn't the one on the intent")
    missing = _tracker_row(key, "Affidavit at L&I, not in the tracker")
    for r in list(log.iter_rows(min_row=2)):
        if str(r[2].value or "") in (f"E{staff}", f"C{hiring}"):
            for c in r:
                c.value = None
    ws.cell(missing, 7).comment = _note("Missing affidavit found in L&I's records: 8810002 (intent 8800002); add it "
                                        "to the tracker (L&I: 8810002)")
    ws.cell(missing, 7).fill = openpyxl.styles.PatternFill()
    wb.save(f)
    s = R.score(d / "results", key)
    bad = _bad(s)
    assert bad[staff]["missed"] == ["intent_id:intent_not_found"]                      # in the Review Log only
    assert bad[hiring]["missed"] == ["hiring:hiring_mismatch"]                         # its note alone isn't enough
    assert bad[missing]["words"] == ["affidavit_id:affidavit_missing: note gives ['8810002', '8800002'], not "
                                     "['8810017', '8800017']"]
    # verification round: the right number only in the "(L&I: ...)" part is not enough
    assert R.record_numbers("Missing affidavit found in L&I's records: 8810002 (intent 8800017); add it to the "
                            "tracker (L&I: 8810017)") == ["8810002", "8800017"]
    assert R.record_numbers("Intent belongs to a different L&I project (L&I: 9900202)") == ["9900202"]
    (g,) = [t for t in s["trackers"] if t["project"] == TRACKER]
    assert any("without the red fill" in p and f"row {missing}" in p for p in g["problems"])
    assert any("red cells without a finding" in p for p in g["problems"])          # the two log rows emptied


def test_the_run_lines_review_log_not_due_counts_and_outcome_are_checked(run, tmp_path):
    """Review finding 7 too: a warning, the counts, and the not-applicable reason."""
    d, key = _copy(run, tmp_path)
    f = _summary(d)
    summary = json.loads(f.read_text())
    lines = summary["projects"][TRACKER]
    lines = [l.replace("1 line(s) not started", "1 line(s) not started").replace("2 line(s) not started",
                                                                                "3 line(s) not started") for l in lines]
    first = next(i for i, l in enumerate(lines) if l.startswith("intent:"))
    lines[first] += "; WARNING planted"
    summary["projects"][TRACKER] = [l for l in lines if "hiring contractor(s)" not in l]
    summary["projects"][NOT_APPLICABLE] = ["intent: not applicable (tracker has no intent/affidavit columns)"]
    f.write_text(json.dumps(summary))
    k2 = json.loads(json.dumps(key))
    k2["trackers"][TRACKER]["lni_only"] = ["Foxtrennic Traffic Control", "Havrolde Demolition"]
    s = R.score(d / "results", k2)
    problems = {t["project"]: " | ".join(t["problems"]) for t in s["trackers"] if not t["ok"]}
    for words in ("not due yet", "Havrolde", "a warning on a clean check", "0 hiring, expected 1"):
        assert words in problems[TRACKER], words
    assert "doesn't say \"Project Info says no L&I filing required\"" in problems[NOT_APPLICABLE]
    summary["projects"][NOT_APPLICABLE] = ["intent: 0 invalid/missing value(s) marked in Compliance Tracker.xlsx"]
    f.write_text(json.dumps(summary))
    s = R.score(d / "results", key)
    assert any('ended "checked"' in p for t in s["trackers"] for p in t["problems"])


def test_a_tracker_not_checked_is_one_difference_not_one_per_line(run, tmp_path):
    """Review finding 14: a checkout whose check doesn't run (an older one, L&I unreachable) shows as the outcome,
    not as every line missed; the tracker gap is "not measured"."""
    d, key = _copy(run, tmp_path)
    _tracker_gap(key, "Intent ID is not a number", ["intent_id:intent_not_found"])
    f = _summary(d)
    summary = json.loads(f.read_text())
    summary["projects"][TRACKER] = ["intent: NOT CHECKED - L&I's live lookup could not be reached, even after retries"]
    f.write_text(json.dumps(summary))
    wb = openpyxl.Workbook()
    wb.active.title = L.TRACKER_SHEET
    wb.active.append(L.HEADER)
    wb.save(_tracker(d))                                                   # as it was: no marks, no Review Log
    s = R.score(d / "results", key)
    assert not any(ln["missed"] for ln in s["tracker_lines"])
    (g,) = [t for t in s["trackers"] if t["project"] == TRACKER]
    assert not g["ok"] and 'ended "not checked"' in g["problems"][0]
    assert s["known_gaps"][-1]["status"].startswith("not measured")


def test_asking_l_and_i_about_a_tracker_that_isnt_checked_or_anything_else_is_caught(run, tmp_path):
    d, key = _copy(run, tmp_path)
    req = json.loads((d / "lni_requests.json").read_text())
    assert req and all(q["answered"] for q in req)                        # this checkout asks L&I only, as it should
    assert not any(q["id"] in key["trackers"][NOT_APPLICABLE]["never_ask"] for q in req)
    req += [{"method": "POST", "url": L.LNI_URL, "op": "GetIntent", "id": "8800090", "answered": True},
            {"method": "CONNECT", "url": "sam.gov:443", "op": None, "id": None, "answered": False}]
    (d / "lni_requests.json").write_text(json.dumps(req))
    s = R.score(d / "results", key)
    assert s["lni_asked_wrongly"] == ["GetIntent 8800090 (a tracker that isn't checked)", "refused: CONNECT sam.gov:443"]
    assert s["failures"] == 2
    (d / "lni_requests.json").unlink()                                     # review finding 10: not silently passed
    s = R.score(d / "results", key)
    assert s["failures"] == 1 and "wasn't recorded" in s["lni_asked_wrongly"][0]


def _tracker_gap(key, case, lines_currently):
    """Pins a tracker line as a known gap (as eval_lni.GAPS does), today ending with ``lines_currently``."""
    row = _tracker_row(key, case)
    key["trackers"][TRACKER]["lines"][str(row)]["known_gap"] = "planted"
    key["known_gaps"].append({"id": "planted", "issue": 175, "files": [], "trackers": [TRACKER],
                              "currently": "; ".join(lines_currently), "lines_currently": list(lines_currently),
                              "why": "test"})
    return row


def test_the_tracker_gap_passes_when_fixed_and_is_noticed_when_it_changes(run, tmp_path):
    """#175 is fixed (intent.strict_ids, on in the harness), so the gap is planted: the line pinned as ending "not
    found", as it did before the fix."""
    d, key = _copy(run, tmp_path)
    row = _tracker_gap(key, "Intent ID is not a number", ["intent_id:intent_not_found"])
    s = R.score(d / "results", key)                     # the run is right: the gap passes now
    assert s["known_gaps"][-1]["status"].startswith("passes now") and s["failures"] == 0
    f = _tracker(d)
    wb = openpyxl.load_workbook(f)
    for r in wb["Review Log"].iter_rows(min_row=2):
        if str(r[2].value or "") == f"E{row}":
            r[7].value = "Intent ID not found in L&I records for this project"                # as before #175
    wb.worksheets[0].cell(row, 5).comment = _note("Intent ID not found in L&I records for this project")
    wb.save(f)
    s = R.score(d / "results", key)
    assert s["known_gaps"][-1]["status"] == "still open" and s["failures"] == 0
    for r in wb["Review Log"].iter_rows(min_row=2):                       # another result: the right words, but
        if str(r[2].value or "") == f"E{row}":                            # only in the Review Log, not the note
            r[7].value = "Intent ID is not a valid number"
    wb.worksheets[0].cell(row, 5).comment = None
    wb.save(f)
    s = R.score(d / "results", key)
    assert s["known_gaps"][-1]["status"].startswith("changed") and s["failures"] >= 1


def test_a_gap_on_payroll_lines_tolerates_only_todays_flags(run, tmp_path):
    """A known gap on lines of a reviewed payroll (rows_currently): a line raising exactly today's rules is
    tolerated (reported, not a false flag); any other result is a difference; the right result passes the gap."""
    rel = "Eval I - Unreadable values/Week 38.xlsx"
    d, key = _copy(run, tmp_path)
    row = _row_of(key, rel, "Base rate typed as words")
    key["known_gaps"].append({"id": "planted", "issue": 65, "files": [rel], "currently": "reviewed", "why": "test",
                              "rows_currently": {rel: {str(row): ["R07", "R08"]}}})
    s = R.score(d / "results", key)
    assert s["known_gaps"][-1]["status"].startswith("passes now") and s["failures"] == 0
    f = _reviewed(d, rel)
    wb = openpyxl.load_workbook(f)
    for c in _sheet(wb)[row]:
        c.comment = None
    _sheet(wb).cell(row, 34).comment = Comment("[RULE BROKEN] R08 Base / total prevailing rate not met: planted", "t")
    _sheet(wb).cell(row, 38).comment = Comment("[RULE BROKEN] R07 Fringe requirement not met: planted", "t")
    wb.save(f)
    s = R.score(d / "results", key)                     # today's flags: tolerated
    assert s["known_gaps"][-1]["status"] == "still open" and s["failures"] == 0 and s["totals"]["false_flags"] == 0
    _sheet(wb).cell(row, 38).comment = None                                             # another result: R08 alone
    wb.save(f)
    s = R.score(d / "results", key)
    assert s["known_gaps"][-1]["status"].startswith("changed") and s["failures"] == 1
    assert s["totals"]["false_flags"] == 1
    for c in _sheet(wb)[row]:                                                           # another: nothing at all
        c.comment = None
    wb.save(f)
    s = R.score(d / "results", key)                     # a miss only: still a difference on a gap's line
    assert s["known_gaps"][-1]["status"].startswith("changed") and s["failures"] == 1


# ----------------------------------------------------------------------------------------- the structured run log
def _events(run_dir: Path) -> Path:
    return next((run_dir / "results" / "run_logs").glob("*__events.jsonl"))


def _lines(f: Path) -> list:
    return [json.loads(x) for x in f.read_text().splitlines()]


def _write(f: Path, lines) -> None:
    f.write_text("".join(json.dumps(d, sort_keys=True) + "\n" for d in lines))


def test_the_run_log_is_checked_against_the_run(run, tmp_path):
    d, key = _copy(run, tmp_path, events=True)
    assert R.score(d / "results", key)["failures"] == 0
    f = _events(d)
    good = _lines(f)

    def problems(lines):
        _write(f, lines)
        s = R.score(d / "results", key)
        assert s["failures"] == len(s["run_log"]["problems"]), R.report(s)
        return s["run_log"]["problems"]

    assert problems(good[:-1]) == ["the last line isn't the only run_end"]           # a run that never ended
    reviewed = next(i for i, x in enumerate(good) if x["event"] == "file" and x["outcome"] == "reviewed"
                    and x["project"] == SEEDED.split("/")[0])
    wrong = json.loads(json.dumps(good))
    wrong[reviewed]["flags"]["R03"] += 1                                             # a count the notes don't show
    assert problems(wrong) == [f"{SEEDED.split('/')[0]}: flag counts differ from the reviewed workbook's notes (R03)"]
    del wrong[reviewed]                                                              # a file line missing
    (p,) = [x for x in problems(wrong) if "file lines say" in x]
    assert p.startswith(SEEDED.split("/")[0]) and len(problems(wrong)) == 3          # (+ its flags, run_end count)
    extra = json.loads(json.dumps(good))
    next(x for x in extra if x["event"] == "task")["note"] = "x"                    # a field outside the schema
    p, missing = problems(extra)                     # a line failing the schema isn't used: its task is missing too
    assert "fails the schema: fields note" in p and missing.startswith("task lines:")
    skipped = json.loads(json.dumps(good))
    for x in skipped:
        if x.get("reason") == "no_wage_rates":
            x["reason"] = "wage_type_unknown"
    assert problems(skipped) == ["Eval D - No wage rates: no skipped file line with reason no_wage_rates"]
    task = json.loads(json.dumps(good))
    task = [x for x in task if not (x["event"] == "task" and x["task"] == "intent" and x["project"] == TRACKER)]
    n = sum(1 for x in good if x["event"] == "task")
    assert problems(task) == [f"task lines: {n - 1} for {n} tasks run"]
    status = json.loads(json.dumps(good))                                            # review 2: a task's status
    next(x for x in status if x["event"] == "task" and x["outcome"] == "skipped")["outcome"] = "ok"
    assert problems(status) == [f"task lines: {n} for {n} tasks run (a task's project, name or outcome differs)"]
    ended = json.loads(json.dumps(good))
    ended[-1]["outcome"] = "crashed"                                                # review 2: run_end's outcome
    assert problems(ended) == ["run_end's outcome or counts differ from the run summary"]
    ended[-1]["outcome"], ended[-1]["warnings"] = good[-1]["outcome"], good[-1]["warnings"] + 1
    assert problems(ended) == ["run_end's outcome or counts differ from the run summary"]
    begun = json.loads(json.dumps(good))
    begun[0]["tasks"] = ["payroll"]
    assert problems(begun) == ["run_start's tasks or project count differ from the run summary"]
    numbers = json.loads(json.dumps(good))
    next(x for x in numbers if x["event"] == "file" and x["file_n"] == 3)["file_n"] = 2
    (p,) = problems(numbers)
    assert p.endswith("file numbers aren't 1 to 3")
    swapped = json.loads(json.dumps(good))                                           # review 2: file by file
    b1, b2 = [x for x in swapped if x["event"] == "file" and x["project"] == "Eval C - Messy layouts"
              and x["outcome"] == "reviewed"]
    b1["flags"]["R01"] += 1
    b2["flags"]["R01"] -= 1
    assert problems(swapped) == ["Eval C - Messy layouts: flag counts differ from the reviewed workbook's notes "
                                 "(file by file)"]
    nobody = json.loads(json.dumps(good))                                            # review 1: no crash
    del next(x for x in nobody if x["event"] == "file")["project"]
    nobody.insert(2, {**nobody[1], "project": "../run_logs"})                       # review 11: not a project folder
    found = problems(nobody)
    assert any("names no project of the run" in x for x in found) and any("fails the schema: fields project" in x
                                                                          for x in found)
    for field, value in (("flags", [1]), ("flags", "R01"), ("project", ["P"]), ("run_id", [1]), ("event", ["task"])):
        typed = json.loads(json.dumps(good))                                         # verification N1: no crash
        next(x for x in typed if x["event"] == "file")[field] = value
        assert problems(typed)
    gone = json.loads(json.dumps(good))
    gone[-1]["dropped"] = 1
    assert problems(gone) == ["run_end: 1 field(s) dropped, 0 failed write(s)"]
    f.unlink()                                       # the file missing although the run's config.yaml turned it on
    assert R.score(d / "results", key)["run_log"]["problems"] == ["0 events files in run_logs (expected 1)"]
    assert R.score(d / "results", key, events=None)["run_log"]["measured"] is False  # a checkout without it


if __name__ == "__main__":                                      # ci.yml falls back to running a test file directly
    sys.exit(pytest.main([__file__, "-q", "-p", "no:cacheprovider"]))
