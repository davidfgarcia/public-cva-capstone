"""
Checks for the five-folder project layout: migration of the previous layout, recognition of uploads, and full runs.
Every name, number and company here is invented, except the published wage samples in model_checks/.

    python tests/test_layout.py
    pytest tests/
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
sys.path.insert(0, str(AGENT))

from agent import uploads                                     # noqa: E402
from agent.config import load_config                          # noqa: E402
from agent.layout import Paths, ensure_layout, migrate_project   # noqa: E402
from agent.runner import run                                  # noqa: E402

SAMPLES = AGENT / "model_checks" / "wage_files"
PAYROLL_HEADER = ["Project Name", "Contractor Name", "Jurisdiction", "Week End Date", "Employee Name",
                  "Employee ID (LCPtracker)", "Craft", "Classification", "ST Hours", "OT Hours", "DT Hours",
                  "Total Hours", "Base Hourly", "Overtime Hourly", "Doubletime Hourly", "Cash Fringe Rate",
                  "3rd Party Fringe Rate", "Grand Total Wages"]


def payroll_xlsx(path: Path, jurisdiction="Federal - Pierce County"):   # (since Nov 7 a line must say its wage type)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(PAYROLL_HEADER)
    ws.append(["Test Project", "Acme Paving", jurisdiction, "2026-09-06", "Rivera, Ana", 10234, "Boilermakers",
               "Journey Level", 40, 0, 0, 40, 60.0, 90.0, 120.0, 0, 18.89, 2400])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def tracker_xlsx(path: Path, status="open"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compliance Tracker"
    ws.append(["Status", "Contractor", "Hiring Contractor", "Intent ID", "Affidavit ID"])
    ws.append([status, "Acme Paving", "Prime", 1234567, None])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def touch(path: Path, text="x"):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix == ".xlsx":
        wb = openpyxl.Workbook()
        wb.active.append([text])
        wb.save(path)
    else:
        path.write_text(text)
    return path


def old_layout_project(root: Path, name="Corvath - XHA Juniper Park"):
    """A project in the previous layout, shaped like the real ones (file names invented or generic)."""
    p = root / "Projects" / name
    t = p / "tasks"
    payroll_xlsx(t / "payroll" / "inbox" / "XHA 9.1.26-9.30.26.xlsx")
    payroll_xlsx(t / "payroll" / "archive" / "2026-10-22-16-38-33__inbox" / "Older payroll.xlsx")
    touch(t / "payroll" / "reviewed" / "Older payroll - REVIEWED.xlsx")
    touch(t / "payroll" / "logs" / "2026-10-22-16-38-33__log.txt", "timestamp_utc,level\n")
    shutil.copy(SAMPLES / "federal_WA20250093.txt", t / "payroll" / "wages" / "federal" / "inbox" / "WA20250093.txt") \
        if (t / "payroll" / "wages" / "federal" / "inbox").mkdir(parents=True) is None else None
    touch(t / "payroll" / "wages" / "federal" / "current" / "current_wage_rates.xlsx")
    touch(t / "payroll" / "wages" / "federal" / "archive" / "2026-10-22-16-38-33__inbox" / "Corvath_Federal WA20250091.pdf")
    touch(t / "payroll" / "wages" / "federal" / "archive" / "2026-10-22-16-38-33__current" / "current_wage_rates.xlsx")
    touch(t / "payroll" / "wages" / "logs" / "2026-10-22-16-38-33__log.txt", "timestamp_utc,level\n")
    for k in ("state", "union"):
        for sub in ("inbox", "current", "archive"):
            (t / "payroll" / "wages" / k / sub).mkdir(parents=True, exist_ok=True)
    touch(t / "intent" / "lni_data" / "intents.csv")
    touch(t / "intent" / "lni_data" / "lni_snapshot.json", "{}")
    touch(t / "intent" / "archive" / "2026-10-22-20-42-53__lni_data" / "lni_snapshot.json", "{}")
    touch(t / "intent" / "logs" / "2026-10-22-20-42-53__log.txt", "timestamp_utc,level\n")
    tracker_xlsx(p / "tracker" / "XHA 50002 Compliance Tracker.xlsx")
    touch(p / "tracker" / "archive" / "2026-10-22-20-42-53__tracker" / "XHA 50002 Compliance Tracker.xlsx")
    touch(p / "unused" / "Corvath - Federal Only.zip")
    (p / ".agent_state.json").write_text(json.dumps({"timestamps": "utc", "payroll_last": {
        "file": "Older payroll.xlsx", "archived_input": "tasks/payroll/archive/2026-10-22-16-38-33__inbox/Older payroll.xlsx"}}))
    return p


def five_folder_project(root: Path, name="Atrovan - Test"):
    """A project in the five-folder layout (Upload, Reviewed payroll, Tracker, Wage rates, Archive)."""
    p = root / "Projects" / name
    payroll_xlsx(p / "Upload" / "Week 40.xlsx")
    touch(p / "Upload" / "notes.docx")
    touch(p / "Reviewed payroll" / "2026-10-20 Week 39 - REVIEWED.xlsx")
    tracker_xlsx(p / "Tracker" / "Tracker.xlsx")
    touch(p / "Wage rates" / "State.xlsx")
    a = p / "Archive" / "2026-10-20-10-00-00"
    payroll_xlsx(a / "Uploaded files" / "Week 39.xlsx")
    tracker_xlsx(a / "Uploaded files" / "Tracker.xlsx")
    shutil.copy(SAMPLES / "federal_WA20250093.txt", (a / "Uploaded files").mkdir(parents=True, exist_ok=True) or a / "Uploaded files" / "WA20250093.txt")
    touch(a / "Replaced" / "Reviewed payroll" / "2026-10-19 Week 39 - REVIEWED.xlsx")
    touch(a / "Replaced" / "Tracker" / "Tracker.xlsx")
    touch(a / "Replaced" / "Wage rates" / "State.xlsx")
    touch(a / "Replaced" / "Upload" / "dup.pdf")
    touch(a / "Removed" / "wrong.pdf")
    touch(a / "L&I data" / "intents.csv")
    touch(a / "Logs" / "payroll__log.txt")
    touch(p / "Archive" / "Unused" / "old.zip")
    (p / ".agent_state.json").write_text(json.dumps({"payroll_last": {
        "file": "Week 39.xlsx", "archived_input": "Archive/2026-10-20-10-00-00/Uploaded files/Week 39.xlsx"}}))
    return p


def test_migration_from_the_five_folder_layout():
    with tempfile.TemporaryDirectory() as d:
        p = five_folder_project(Path(d))
        before = len(tree(p))
        migrate_project(p)
        after = tree(p)
        assert len(after) == before, "a file was lost or duplicated"
        assert sorted(x.name for x in p.iterdir() if not x.name.startswith(".")) == ["Files", "Payroll", "Tracker", "Uploads"]
        ts = "2026-10-20-10-00-00"
        expect = ["Uploads/Week 40.xlsx", "Uploads/notes.docx", "Payroll/2026-10-20 Week 39 - REVIEWED.xlsx",
                  "Tracker/Tracker.xlsx", "Files/Wage rates/State.xlsx",
                  f"Payroll/Archived/{ts}/Week 39.xlsx", f"Payroll/Archived/{ts}/2026-10-19 Week 39 - REVIEWED.xlsx",
                  f"Tracker/Archived/{ts}/Tracker (as uploaded).xlsx", f"Tracker/Archived/{ts}/Tracker (previous).xlsx",
                  f"Files/Archived/{ts}/WA20250093.txt", f"Files/Archived/{ts}/State (previous).xlsx",
                  f"Files/Archived/{ts}/dup (replaced upload).pdf", f"Files/Archived/{ts}/wrong (removed).pdf",
                  f"Files/L&I data/{ts}/intents.csv", f"Files/Logs/{ts}/payroll__log.txt", "Files/Unused/old.zip"]
        missing = [e for e in expect if e not in after]
        assert not missing, (missing, after)
        st = json.loads((p / ".agent_state.json").read_text())
        assert st["payroll_last"]["archived_input"] == f"Payroll/Archived/{ts}/Week 39.xlsx"
        assert ensure_layout(p) == [] and tree(p) == after


def test_new_folders_are_not_mistaken_for_old_ones_on_case_insensitive_drives():
    """On a Mac, Windows or Google Drive, 'payroll' finds 'Payroll'. Emulated with lower-case symlinks."""
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "Projects" / "Alpha"
        ensure_layout(p)
        touch(p / "Payroll" / "2026-10-20 Week 39 - REVIEWED.xlsx")
        touch(p / "Payroll" / "Archived" / "2026-10-20-10-00-00" / "Week 39.xlsx")
        for low, real in (("payroll", "Payroll"), ("tracker", "Tracker"), ("files", "Files")):
            (p / low).symlink_to(p / real, target_is_directory=True)
        names = set(os.listdir(p))
        before = tree(p)
        # os.listdir shows the symlinks, but a real case-insensitive folder lists only one name: hide them
        import agent.layout as L
        real_listdir = L.os.listdir
        L.os.listdir = lambda x: [n for n in real_listdir(x) if not (Path(x) == p and n in ("payroll", "tracker", "files"))]
        try:
            assert ensure_layout(p) == []
        finally:
            L.os.listdir = real_listdir
        assert names == set(os.listdir(p)) and tree(p) == before


def test_an_interrupted_migration_resumes():
    with tempfile.TemporaryDirectory() as d:
        p = old_layout_project(Path(d))
        (p / "tracker").rename(p / "_tracker_being_moved")         # as left by a run that stopped right there
        before = len(tree(p))
        migrate_project(p)
        after = tree(p)
        assert len(after) == before and "Tracker/XHA 50002 Compliance Tracker.xlsx" in after, after
        assert not (p / "_tracker_being_moved").exists()


def test_repeated_name_clashes_never_overwrite():
    from agent.layout import _merge_dir
    with tempfile.TemporaryDirectory() as d:
        d = Path(d)
        touch(d / "dst" / "a.txt", "KEPT")
        touch(d / "dst" / "a__migrated.txt", "EARLIER MIGRATED")
        touch(d / "src" / "a.txt", "NEW")
        _merge_dir(d / "src", d / "dst")
        texts = sorted(f.read_text() for f in (d / "dst").iterdir())
        assert texts == ["EARLIER MIGRATED", "KEPT", "NEW"], texts


def tree(p: Path):
    return sorted(str(f.relative_to(p)) for f in p.rglob("*") if f.is_file() and f.name != ".DS_Store")


def test_migration_from_the_previous_layout():
    with tempfile.TemporaryDirectory() as d:
        p = old_layout_project(Path(d))
        before = len(tree(p))
        migrate_project(p)
        after = tree(p)
        assert len(after) == before, "a file was lost or duplicated"
        assert sorted(x.name for x in p.iterdir() if not x.name.startswith(".")) == \
            ["Files", "Payroll", "Tracker", "Uploads"]
        expect = [
            "Uploads/XHA 9.1.26-9.30.26.xlsx", "Uploads/WA20250093.txt",
            "Payroll/Older payroll - REVIEWED.xlsx",
            "Tracker/XHA 50002 Compliance Tracker.xlsx",
            "Files/Wage rates/Federal.xlsx",
            "Payroll/Archived/2026-10-22-16-38-33/Older payroll.xlsx",
            "Files/Archived/2026-10-22-16-38-33/Corvath_Federal WA20250091.pdf",
            "Files/Archived/2026-10-22-16-38-33/Federal (previous).xlsx",
            "Files/Logs/2026-10-22-16-38-33/payroll__log.txt",
            "Files/Logs/2026-10-22-16-38-33/wages__log.txt",
            "Files/Logs/2026-10-22-20-42-53/intent__log.txt",
            "Files/L&I data/2026-10-22-20-42-53/lni_snapshot.json",
            "Tracker/Archived/2026-10-22-20-42-53/XHA 50002 Compliance Tracker (previous).xlsx",
            "Files/Unused/Corvath - Federal Only.zip",
            ".agent/lni_data/intents.csv", ".agent/lni_data/lni_snapshot.json",
        ]
        missing = [e for e in expect if e not in after]
        assert not missing, (missing, after)
        st = json.loads((p / ".agent_state.json").read_text())
        assert st["payroll_last"]["archived_input"] == "Payroll/Archived/2026-10-22-16-38-33/Older payroll.xlsx"
        assert ensure_layout(p) == [] and tree(p) == after        # running it again changes nothing


def test_uploads_are_recognized_by_content():
    with tempfile.TemporaryDirectory() as d:
        up = Path(d) / "Upload"
        assert uploads.classify(payroll_xlsx(up / "anything.xlsx"))[0] == "payroll"
        assert uploads.classify(tracker_xlsx(up / "named oddly.xlsx"))[0] == "tracker"
        shutil.copy(SAMPLES / "state_LNI_Pierce_sample.csv", up / "rates.csv")
        assert uploads.classify(up / "rates.csv")[0] == "wages"
        shutil.copy(SAMPLES / "federal_WA20250093.txt", up / "decision.txt")
        assert uploads.classify(up / "decision.txt")[0] == "wages"
        assert uploads.classify(touch(up / "notes.docx"))[0] is None
        assert uploads.classify(touch(up / "photo.txt", "hello"))[0] is None


def test_a_full_run_in_the_new_layout():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        shutil.copy(AGENT / "config.yaml", root / "config.yaml")
        p = root / "Projects" / "Test Project"
        ensure_layout(p)
        up = Paths(p).upload
        payroll_xlsx(up / "Week 36.xlsx")
        shutil.copy(SAMPLES / "state_LNI_Pierce_sample.csv", up / "Journey level wages_Pierce County_6.26.25 LNI.csv")
        shutil.copy(SAMPLES / "federal_WA20250093.txt", up / "WA20250093.txt")
        tracker_xlsx(up / "New tracker.xlsx")
        touch(up / "random notes.docx")
        tracker_xlsx(Paths(p).tracker / "Old tracker.xlsx", status="old")
        cfg = load_config(AGENT / "config.yaml", root)
        rep = run("payroll", ["Test Project"], cfg=cfg, notify=False)
        lines = rep["projects"]["Test Project"]
        ts = rep["run_utc"]
        files = tree(p)
        # A payroll run leaves an uploaded tracker alone: only the intent check uses (and installs) it
        assert "Uploads/New tracker.xlsx" in files and "Tracker/Old tracker.xlsx" in files
        lines_t = uploads.install_tracker_upload(p)
        files_t = tree(p)
        assert any(l.startswith("tracker: New tracker.xlsx is now the current tracker") for l in lines_t), lines_t
        assert "Tracker/New tracker.xlsx" in files_t and "Tracker/Old tracker.xlsx" not in files_t
        assert any(f.startswith("Tracker/Archived/") and f.endswith("/Old tracker (replaced).xlsx") for f in files_t)
        assert any(f.startswith("Tracker/Archived/") and f.endswith("/New tracker (as uploaded).xlsx") for f in files_t)
        rates = [f for f in files if f.startswith("Files/Wage rates/")]              # named for the run, Pacific
        assert [f.split("/")[-1].split(" ")[0] for f in rates] == ["Federal", "State"], rates
        assert all(re.fullmatch(r"Files/Wage rates/(Federal|State) \d{4}(-\d\d){5} P[SD]T\.xlsx", f) for f in rates), rates
        reviewed = [f for f in files if f.startswith("Payroll/") and "/Archived/" not in f]
        assert reviewed == [f"Payroll/{ts[:10]} Week 36 - REVIEWED.xlsx"], reviewed
        assert f"Payroll/Archived/{ts}/Week 36.xlsx" in files
        assert f"Files/Archived/{ts}/WA20250093.txt" in files
        assert any(f.startswith(f"Files/Logs/{ts}/") for f in files)
        assert sorted(f for f in files if f.startswith("Uploads/")) == ["Uploads/New tracker.xlsx", "Uploads/random notes.docx"]
        assert any("random notes.docx not recognized" in l for l in lines)
        # A second payroll upload of the same week, after new wage rates: the earlier review moves to the archive
        payroll_xlsx(up / "Week 36.xlsx")
        rep2 = run("payroll", ["Test Project"], cfg=cfg, notify=False)
        files = tree(p)
        if rep2["run_utc"][:10] == ts[:10]:
            assert [f for f in files if f.startswith("Payroll/") and "/Archived/" not in f] == reviewed     # same day: rewritten
        else:
            assert f"Payroll/Archived/{rep2['run_utc']}/{reviewed[0].split('/')[-1]}" in files


def test_a_run_never_migrates_a_project_in_the_old_layout():
    """A run stops for a project still in an old layout (error in the email, its tasks skipped) and changes nothing."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        shutil.copy(AGENT / "config.yaml", root / "config.yaml")
        p = five_folder_project(root)
        before = {str(f.relative_to(p)) for f in p.rglob("*")}
        cfg = load_config(AGENT / "config.yaml", root)
        rep = run("payroll", [p.name], cfg=cfg, notify=False)
        lines = rep["projects"][p.name]
        assert lines[0].startswith("ERROR (layout): this project is still in the old folder layout"), lines
        assert "payroll: SKIPPED - the project is still in the old folder layout" in lines, lines
        assert rep["outcome"] != "SUCCESS" and rep["level"] in ("Failure", "Critical"), (rep["outcome"], rep["level"])
        assert {str(f.relative_to(p)) for f in p.rglob("*")} == before
        # empty old folders alone don't count as the old layout
        q = root / "Projects" / "Gamma"
        (q / "Upload").mkdir(parents=True)
        (q / "Archive" / "Unused").mkdir(parents=True)
        assert ensure_layout(q) == [] and (q / "Uploads").is_dir()


def test_migrate_layout_script_dry_run_then_apply():
    import io
    import contextlib
    import migrate_layout as M
    with tempfile.TemporaryDirectory() as d:
        root = Path(d) / "compliance"
        five_folder_project(root)
        old_layout_project(root, "Beta - School")
        new = root / "Projects" / "Gamma"
        ensure_layout(new)
        touch(new / "Payroll" / "r.xlsx")
        touch(root / "Projects" / "Atrovan - Test" / "Archive" / "2026-10-20-10-00-00" / ".DS_Store")   # left behind
        touch(root / "Projects" / "Atrovan - Test" / "Upload" / "sub" / "x.csv", "in a subfolder")
        touch(root / "Projects" / "Atrovan - Test" / "Archive" / "loose.txt", "loose in Archive")
        snap = lambda: {str(f.relative_to(root)): f.read_bytes() for f in root.rglob("*") if f.is_file()}
        before = snap()
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            assert M.main([str(root / "Projects")]) == 0
        text = out.getvalue()
        assert "DRY RUN" in text and "CHECK: every file is still there, unchanged" in text, text
        assert "Upload/Week 40.xlsx  ->  Projects/Atrovan - Test/Uploads/Week 40.xlsx" in text, text
        assert "Removed/wrong.pdf  ->  Projects/Atrovan - Test/Files/Archived/2026-10-20-10-00-00/wrong (removed).pdf" in text, text
        assert snap() == before, "the dry run changed the real folder"
        out = io.StringIO()
        # as on the Mac's Drive folder from Claude's shell: deleting (even an empty folder) isn't allowed
        def refuse(*a, **k):
            raise PermissionError("Operation not permitted")
        real = Path.rmdir, Path.unlink
        Path.rmdir = Path.unlink = refuse
        try:
            with contextlib.redirect_stdout(out):
                assert M.main([str(root / "Projects"), "--apply"]) == 0
        finally:
            Path.rmdir, Path.unlink = real
        assert "APPLIED" in out.getvalue() and "CHECK: every file is still there" in out.getvalue(), out.getvalue()
        for name in ("Atrovan - Test", "Beta - School", "Gamma"):
            assert M.old_layout_dirs(root / "Projects" / name, with_files_only=False) == [], name
            assert (root / "Projects" / name / "Uploads").is_dir()
        hold = next((root / "_sync_backup").glob("*__layout_migration"))
        assert (hold / "report.txt").is_file() and (hold / "before.tsv").is_file()
        assert (hold / "old folders" / "Atrovan - Test" / "Archive").is_dir()     # set aside, not deleted
        assert (root / "Projects" / "Gamma" / "Payroll" / "r.xlsx").is_file()
        at = root / "Projects" / "Atrovan - Test"
        assert (at / "Files" / "Unused" / "Upload" / "sub" / "x.csv").read_text() == "in a subfolder"
        found = [f for f in at.rglob("loose.txt")]
        assert found and "Archive" not in found[0].relative_to(at).parts[:1], found
        assert json.loads((at / ".agent_state.json").read_text())["payroll_last"]["archived_input"] == \
            "Payroll/Archived/2026-10-20-10-00-00/Week 39.xlsx"
        assert not [f for f in (hold / "old folders").rglob("*") if f.is_file() and f.name != ".DS_Store"]
        # every file's content is still somewhere (state file updated by design)
        after = list(snap().values())
        for rel, data in before.items():
            if not rel.endswith((".agent_state.json", ".DS_Store")):
                assert data in after, rel
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            assert M.main([str(root / "Projects"), "--apply"]) == 0
        assert "Nothing to do." in out.getvalue()


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"\nAll {len(tests)} tests passed.")
