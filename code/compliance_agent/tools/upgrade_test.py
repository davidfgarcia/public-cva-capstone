"""The upgrade test: a drive that has already run, then a change that renames or reshapes what AWS holds, then the
next runs on the same state. The harness (run_evals.py) starts every run from nothing, so it can't see a change that
breaks what an earlier run left behind: on Nov 15, 2026 the tokenizer's opaque_file_names (#53) renamed the archived
payrolls on AWS, and a Full rerun found "no payroll uploaded" because the agent looks the last one up by the name it
saved (fixed in #195). This test would have failed on that.

    python tools/upgrade_test.py                   # this checkout; ~2 minutes; invented data only
    python tools/upgrade_test.py --keep /tmp/up    # keep the runs to look at

Steps, on the harness's invented projects (eval_cases.py), offline as the harness is:
  1. the run as before: the tokenizer with #53's three settings OFF, then the agent (every task)
  2. the change: the AWS-side tree renamed as the tokenizer's first pass with the settings ON renames it
     (pii_tokenizer/names.py aws_rel, each path below the drive's top; the agent's state and dot-folders stay)
  3. a Full rerun (force) on that state, then an Incremental run
and checks, for every project whose payroll run 1 reviewed: run 3 reviewed it again ("re-reviewed"); for every
project whose tracker run 1 checked: run 3 checked it again; no task failed. Exit 0 = pass, 1 = a check failed,
3 = a run failed.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import eval_cases  # noqa: E402
import eval_lni  # noqa: E402
import run_evals  # noqa: E402
import snapshot_run  # noqa: E402

OFF = {"fail_closed_tables": False, "opaque_sheet_names": False, "opaque_file_names": False}

# The change of step 2, run in the tokenizer's checkout: every path below the top renamed as the mirror would
_RENAME = r'''
import dataclasses, json, os, sys, pathlib
sys.path.insert(0, ".")
from pii_tokenizer import config
from pii_tokenizer.files import Context
from pii_tokenizer.names import aws_rel
from pii_tokenizer.tokens import Vault
root = pathlib.Path(sys.argv[1])
cfg = config.load()
cfg = dataclasses.replace(cfg, vnext={**cfg.vnext, "opaque_file_names": True})
ctx = Context.build(cfg, bytes(range(32)), Vault())
moved = 0
for p in sorted((q for q in root.rglob("*") if q.is_file()), key=lambda q: -len(q.parts)):
    rel = p.relative_to(root).as_posix()
    if any(s.startswith(".") for s in rel.split("/")) or rel.startswith(("run_logs/", "config.yaml")):
        continue
    new = aws_rel(rel, ctx.scrubber, ctx.config)
    if new != rel:
        (root / new).parent.mkdir(parents=True, exist_ok=True)
        os.replace(p, root / new)
        moved += 1
for d in sorted((q for q in root.rglob("*") if q.is_dir()), key=lambda q: -len(q.parts)):
    if not any(d.iterdir()):
        d.rmdir()                                    # the old folders, emptied by the renames
print(json.dumps({"renamed": moved}))
'''

_RUN = r'''
import json, sys, pathlib
sys.path.insert(0, ".")
from agent.config import load_config
from agent.runner import run
root = pathlib.Path(sys.argv[1])
cfg = load_config(root / "config.yaml", root)
rep = run("all", None, force=sys.argv[2] == "force", cfg=cfg, notify=True)
print("UPGRADE-RUN " + json.dumps({"outcome": rep.get("outcome"), "projects": rep.get("projects") or {}}))
'''


def staff_names(snap: Path) -> int:
    """Every upload renamed the way staff files are named ("Zorvex_PayrollDetailReport Week 37.xlsx": a contractor's
    word and a glued report name, none of them known words), so that the change of step 2 renames what the agent
    relies on, as it did on the real drive. (The harness's own names are made of known words: they don't move.)"""
    n = 0
    for f in sorted((snap / "Projects").glob("*/Uploads/*")):
        if f.is_file() and f.suffix.lower() in (".xlsx", ".xlsm", ".csv", ".pdf", ".txt"):
            f.rename(f.with_name(f"Zorvex_PayrollDetailReport {f.name}"))
            n += 1
    return n


def _agent(code: Path, res: Path, env: dict, py: str, mode: str, out: Path) -> dict:
    with open(out / f"agent_{mode}.txt", "wb") as f:
        done = subprocess.run([py, "-c", _RUN, str(res), mode], cwd=code / "compliance_agent", env=env,
                              stdout=f, stderr=subprocess.STDOUT)
    text = (out / f"agent_{mode}.txt").read_text(errors="replace")
    if done.returncode:
        raise SystemExit(f"the {mode} run failed (exit {done.returncode})")
    line = next(ln for ln in reversed(text.splitlines()) if ln.startswith("UPGRADE-RUN "))
    return json.loads(line[len("UPGRADE-RUN "):])


def _lines(rep: dict, project: str) -> list:
    out = []
    for ln in rep["projects"].get(project) or []:
        out.append(ln if isinstance(ln, str) else json.dumps(ln))
    return out


def check(run1: dict, run3: dict) -> list:
    """The upgrade's problems: what run 1 did that the Full rerun after the change didn't do again."""
    problems = []
    for project in sorted(run1["projects"]):
        before, after = _lines(run1, project), _lines(run3, project)
        if any(ln.startswith("payroll: reviewed") for ln in before) and \
                not any(ln.startswith("payroll: re-reviewed") for ln in after):
            problems.append(f"{project}: payroll reviewed in run 1, not reviewed again by the Full rerun: "
                            + "; ".join(ln for ln in after if ln.startswith("payroll:"))[:300])
        if any(ln.startswith("intent:") and "marked in" in ln for ln in before) and \
                not any(ln.startswith("intent:") and ("marked in" in ln or "no invalid" in ln or "all " in ln)
                        for ln in after):
            problems.append(f"{project}: tracker checked in run 1, not checked by the Full rerun: "
                            + "; ".join(ln for ln in after if ln.startswith("intent:"))[:300])
        if any("FAILED" in ln for ln in after):
            problems.append(f"{project}: a task failed after the change")
    return problems


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--code", type=Path, default=HERE.parents[1], help="the checkout's code/ folder")
    ap.add_argument("--keep", type=Path, help="keep everything in this (empty) folder")
    ap.add_argument("--python", default=sys.executable)
    a = ap.parse_args(argv)
    code = a.code.resolve()
    with tempfile.TemporaryDirectory(prefix="cva-upgrade-") as tmp:
        base = (a.keep or Path(tmp)).resolve()
        snap, out = base / "snapshot", base / "run"
        run_evals.make(snap)
        staff_names(snap)
        lookup = eval_lni.FakeLookup().start()
        try:
            # 1. the run as before the change (#53's settings off)
            snapshot_run.run_checkout(code, snap, out, a.python, a.python, capture=True,
                                      config_extra={**run_evals.lni_settings()}, proxy=lookup.proxy,
                                      tokenizer_vnext=OFF)
            res = out / "results"
            env = snapshot_run.offline_env(out, lookup.proxy)
            summary = run_evals.run_summary(res)
            run1 = {"projects": summary.get("projects") or {}}
            # 2. the change: the tokenizer's first pass with opaque_file_names renames what AWS holds
            done = subprocess.run([a.python, "-c", _RENAME, str(res)], cwd=code / "tokenizer", capture_output=True,
                                  text=True)
            if done.returncode:
                print(done.stdout + done.stderr, file=sys.stderr)
                return 3
            renamed = json.loads(done.stdout.strip().splitlines()[-1])["renamed"]
            # 3. a Full rerun, then an Incremental run, on that state
            run3 = _agent(code, res, env, a.python, "force", out)
            run4 = _agent(code, res, env, a.python, "incremental", out)
        except SystemExit as e:
            print(f"a run failed: {e}", file=sys.stderr)
            for f in sorted(out.glob("agent_*.txt")) + sorted((out / "step_output").glob("*.txt")):
                print(f"--- {f.name}\n" + "\n".join(f.read_text(errors="replace").splitlines()[-30:]), file=sys.stderr)
            return 3
        finally:
            lookup.stop()
        problems = check(run1, run3)
        if any("FAILED" in ln for p in run4["projects"] for ln in _lines(run4, p)):
            problems.append("the Incremental run after the change had a failed task")
        print(f"## CVA upgrade test: {'PASS' if not problems else 'FAIL'}")
        print(f"- run 1 (#53's settings off): {len(run1['projects'])} project(s); the change renamed {renamed} "
              f"file(s) on the AWS side; then a Full rerun and an Incremental run on that state")
        for p in problems:
            print(f"- {p}")
        if not renamed:
            print("- the change renamed nothing: the test proves nothing (check eval_cases' file names)")
            return 1
        return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
