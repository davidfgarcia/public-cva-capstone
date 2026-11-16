"""Run one checkout of the CVA on an input snapshot, for tools/compare_runs.py (vNext phase 1; docs/vNext.md section 12).

**Invented data only.** The snapshot is built by ``--make`` from the tests' invented projects (tests/test_layout.py:
invented workers, contractors and numbers, plus the published wage samples in model_checks/). Never point this at a
staff folder or a copy of one (HANDOFF section 6).

Procedure, comparing ``main`` with ``vnext`` (or any two checkouts):

    git worktree add /tmp/cva-main origin/main
    git worktree add /tmp/cva-vnext origin/vnext
    T=code/compliance_agent/tools
    python $T/snapshot_run.py --make /tmp/snap                                     # the input, built once
    python $T/snapshot_run.py --code /tmp/cva-main/code  --snapshot /tmp/snap --out /tmp/run-main
    python $T/snapshot_run.py --code /tmp/cva-vnext/code --snapshot /tmp/snap --out /tmp/run-vnext
    python $T/compare_runs.py /tmp/run-main /tmp/run-vnext

Each run writes ``<out>/tokenized/`` (the checkout's tokenizer applied to every file of the snapshot with a fixed test
key, in two passes so names learnt in one file are scrubbed in the others: a stand-in for the Drive mirror) and
``<out>/results/`` (the checkout's agent run on a copy of ``tokenized/``: every task, every project, the outcome
emails saved as .eml previews instead of being sent). Python: ``--agent-python`` (3.9 like the server, default: this
interpreter) and ``--tokenizer-python`` (3.12 like CI, default: this interpreter). No network is needed: the L&I and
SAM.gov lookups fail the same way in both runs: both steps run with every proxy pointed at a closed port, no cloud
credentials or API keys, and an agent config with the AI review, the local model, SAM.gov downloads and error
reporting off. The tokenizer step writes ``tokenized/tokenize_report.json`` (files written, left out by the file
rules, held back); it stops the run if it wrote nothing.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent

_TOKENIZE = r'''
import dataclasses, inspect, json, os, sys, pathlib
sys.path.insert(0, ".")
from pii_tokenizer import config, screen
from pii_tokenizer.files import Context, tokenize_file
from pii_tokenizer.tokens import Vault
src, dst = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
cfg = config.load()
# vNext settings the caller turns on (tools/run_evals.py: #53's), those this checkout knows
want = json.loads(os.environ.get("CVA_EVAL_TOKENIZER_VNEXT") or "{}")
cfg = dataclasses.replace(cfg, vnext={**cfg.vnext, **{k: v for k, v in want.items()
                                                     if k in getattr(config, "VNEXT_SETTINGS", ())}})
ctx = Context.build(cfg, bytes(range(32)), Vault())                    # the tests' fixed key: invented data only
try:                                                                     # the AWS side's names, as the mirror makes
    from pii_tokenizer.names import aws_rel                              # them (#53); older checkouts: as written
except ImportError:
    aws_rel = None
files = sorted(f for f in (src / "Projects").rglob("*") if f.is_file())
counts = {"files": len(files), "written": 0, "left_out": 0, "held_back": 0}
renamed = {}         # the AWS-side name of a file the tokenizer renamed (#53), by a hash of its staff path: no name
import hashlib
for final in (False, True):                                              # pass 1 learns names, pass 2 writes
    for f in files:
        rel = f.relative_to(src)
        how = ctx.config.handling(f.name)                                # as mirror._tokenize
        try:
            if how == "scan_pdf":
                data = f.read_bytes()
                screen.screen_pdf(rel.as_posix(), data, ctx.scrubber)
                name, out = f.name, data
            else:
                # the file's path, as the mirror passes it (#111), when the checkout's tokenizer takes it
                kw = {"rel": rel.as_posix()} if "rel" in inspect.signature(tokenize_file).parameters else {}
                name, out, _ = tokenize_file(f.name, f.read_bytes() if how != "skip" else b"", ctx, **kw)
                if out is not None and screen.residual(name, out, ctx.scrubber):
                    raise screen.Held("residual")
        except Exception:                                                # held back, as the mirror would
            counts["held_back"] += final
            continue
        if not final:
            continue
        if out is None:
            counts["left_out"] += 1                                      # the file rules leave it out (.json, .zip)
            continue
        out_rel = (rel.parent / name).as_posix()
        aws = aws_rel(out_rel, ctx.scrubber, ctx.config) if aws_rel else out_rel
        if aws != out_rel:
            renamed[hashlib.sha256(rel.as_posix().encode()).hexdigest()] = aws
        target = dst / aws
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(out)
        counts["written"] += 1
(dst / "tokenize_report.json").write_text(json.dumps(counts, indent=1))   # compared too
(dst.parent / "aws_names.json").write_text(json.dumps(renamed, indent=1, sort_keys=True))   # for run_evals.py
print("tokenized: " + ", ".join(f"{k} {v}" for k, v in counts.items()))
sys.exit(0 if counts["written"] else 4)                                  # nothing written: a broken tokenizer
'''

_RUN = r'''
import sys, pathlib
sys.path.insert(0, ".")
from agent.config import load_config
from agent.runner import run
root = pathlib.Path(sys.argv[1])
cfg = load_config(root / "config.yaml", root)
rep = run("all", None, cfg=cfg, notify=True)
print("run:", rep.get("outcome"), "-", len(rep.get("projects") or {}), "project(s)")
'''


def make_snapshot(dest: Path) -> None:
    """An invented input: two projects (one in the current layout, one in the previous), payroll, tracker, wage file."""
    sys.path.insert(0, str(AGENT))
    sys.path.insert(0, str(AGENT / "tests"))
    import test_layout as T                                              # noqa: E402 - the tests' invented builders
    from agent.layout import Paths, ensure_layout                        # noqa: E402
    if dest.exists() and any(dest.iterdir()):
        raise SystemExit("the snapshot folder must be empty")
    dest.mkdir(parents=True, exist_ok=True)
    shutil.copy(AGENT / "config.yaml", dest / "config.yaml")
    a = dest / "Projects" / "Alpha - Bridge"
    ensure_layout(a)
    T.payroll_xlsx(Paths(a).upload / "Week 36.xlsx")
    shutil.copy(T.SAMPLES / "federal_WA20250093.txt", Paths(a).upload / "WA20250093.txt")
    T.tracker_xlsx(Paths(a).tracker / "Tracker.xlsx")
    T.old_layout_project(dest, "Beta - School")


# The agent's settings for a snapshot run: no AI review, no local model, no SAM.gov download, no error reporting
OFFLINE = {("ai_review", "enabled"): False, ("llm", "enabled"): False, ("wages", "sam_gov", "enabled"): False,
           ("error_reporting", "sentry_dsn"): ""}
DEAD_PROXY = "http://127.0.0.1:9"                                        # nothing listens: every request fails at once
DROP_ENV = ("AWS_", "SENTRY_", "ANTHROPIC_", "OPENAI_", "OLLAMA_", "GOOGLE_", "COMPLIANCE_", "SES_")


def offline_config(src: Path, dst: Path, extra: dict = None) -> None:
    """The snapshot's config.yaml with OFFLINE applied, then ``extra`` ({(key, key...): value}: tools/run_evals.py
    points the intent check at its local stand-in for L&I's live lookup)."""
    import yaml
    cfg = yaml.safe_load(src.read_text()) or {}
    for keys, value in {**OFFLINE, **(extra or {})}.items():
        d = cfg
        for k in keys[:-1]:
            if not isinstance(d.get(k), dict):
                d[k] = {}
            d = d[k]
        d[keys[-1]] = value
    dst.write_text(yaml.safe_dump(cfg, sort_keys=False))


def offline_env(out: Path, proxy: str = None) -> dict:
    """The environment for both steps: no cloud credentials, no API keys, every network request refused, emails
    saved as previews. ``proxy``: the proxy every request goes to instead of the closed port (tools/run_evals.py:
    tools/eval_lni.py's stand-in, which answers L&I's live lookup only, refuses everything else and records it all)."""
    env = {k: v for k, v in os.environ.items() if not k.upper().startswith(DROP_ENV)}
    for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        env[k] = proxy or DEAD_PROXY
    env.pop("NO_PROXY", None)                                            # nothing goes around the proxy
    env.pop("no_proxy", None)
    env.update({"SENTRY_DSN": "off", "COMPLIANCE_NOTIFY": "dry-run",
                "COMPLIANCE_NOTIFY_CONFIG": str(out / "no-notify-config.yaml"),
                "COMPLIANCE_NOTIFY_TO": "staff@cva-test.invalid", "COMPLIANCE_NOTIFY_ADMIN_TO": "admin@cva-test.invalid",
                "AWS_ACCESS_KEY_ID": "none", "AWS_SECRET_ACCESS_KEY": "none", "AWS_EC2_METADATA_DISABLED": "true",
                "AWS_CONFIG_FILE": os.devnull, "AWS_SHARED_CREDENTIALS_FILE": os.devnull, "HOME": str(out / "home")})
    return env


def run_checkout(code: Path, snapshot: Path, out: Path, agent_py: str, tok_py: str, capture: bool = False,
                 config_extra: dict = None, proxy: str = None, tokenizer_vnext: dict = None) -> None:
    """``capture``: each step's printed output goes to ``<out>/step_output/<step>.txt`` instead of the terminal
    (tools/run_evals.py: it searches that output too for worker names). ``config_extra`` and ``proxy``: see
    offline_config and offline_env (neither is used for compare_runs). ``tokenizer_vnext``: tokenizer vNext settings
    to turn on for the run (tools/run_evals.py: #53's), as {setting: true}."""
    if out.exists() and any(out.iterdir()):
        raise SystemExit("the output folder must be empty")
    tok, res = out / "tokenized", out / "results"
    tok.mkdir(parents=True)
    (out / "home").mkdir()
    offline_config(snapshot / "config.yaml", tok / "config.yaml", config_extra)   # the agent's settings: no worker data
    env = offline_env(out, proxy)
    if tokenizer_vnext:
        env["CVA_EVAL_TOKENIZER_VNEXT"] = json.dumps(tokenizer_vnext)
    stamp = "%Y-%m-%d-%H-%M-%S"
    start = dt.datetime.now(dt.timezone.utc).strftime(stamp)
    for step, cmd, cwd in (("tokenizer", [tok_py, "-c", _TOKENIZE, str(snapshot), str(tok)], code / "tokenizer"),
                           ("agent", None, None)):
        if step == "agent":
            shutil.copytree(tok, res)
            cmd, cwd = [agent_py, "-c", _RUN, str(res)], code / "compliance_agent"
        if capture:
            (out / "step_output").mkdir(exist_ok=True)
            with open(out / "step_output" / f"{step}.txt", "wb") as f:
                done = subprocess.run(cmd, cwd=cwd, env=env, stdout=f, stderr=subprocess.STDOUT)
        else:
            done = subprocess.run(cmd, cwd=cwd, env=env)
        if done.returncode:
            raise SystemExit(f"the {step} step failed (exit {done.returncode})")
    # This run's time window, for compare_runs (it masks run times inside it, and compares every other date)
    (out / "run_window.json").write_text(json.dumps(
        {"start_utc": start, "end_utc": dt.datetime.now(dt.timezone.utc).strftime(stamp)}, indent=1))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Run a CVA checkout on an invented input snapshot.")
    ap.add_argument("--make", type=Path, help="build the invented input snapshot in this (empty) folder and stop")
    ap.add_argument("--code", type=Path, help="the checkout's code/ folder")
    ap.add_argument("--snapshot", type=Path)
    ap.add_argument("--out", type=Path)
    ap.add_argument("--agent-python", default=sys.executable)
    ap.add_argument("--tokenizer-python", default=sys.executable)
    a = ap.parse_args(argv)
    if a.make:
        make_snapshot(a.make)
        return 0
    if not (a.code and a.snapshot and a.out):
        ap.error("--code, --snapshot and --out are needed (or --make)")
    run_checkout(a.code.resolve(), a.snapshot.resolve(), a.out.resolve(), a.agent_python, a.tokenizer_python)
    return 0


if __name__ == "__main__":
    sys.exit(main())
