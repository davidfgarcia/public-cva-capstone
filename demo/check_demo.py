"""Checks the demo end to end, the way the video uses it, and scores it against the answer key.

1. adds the large project's last weekly payroll (``To upload/``) through the UI's upload API, as the page does;
2. runs both tasks (payroll review, intent/affidavit check) on both projects through the UI's run API (one run);
3. scores what staff would see against ``demo_data/answer_key.json`` with the evaluation harness's own scorer
   (code/compliance_agent/tools/run_evals.py ``score``): every seeded payroll line and tracker line, the skipped
   upload's reason, the emails, the structured run log, the L&I requests;
4. searches every file the run left (reviewed workbooks, the marked trackers, Audit and Review Log sheets, the .eml
   previews, run logs, the events file, the stand-in's request log) for anything real: the deny list below, plus any
   employee ID outside the demo's 7320nnn range.

    python demo/check_demo.py                    # a fresh demo in a temporary folder, then the checks
    python demo/check_demo.py --demo /tmp/d --keep

Exit code 0 when everything matches and nothing real was found.
"""
from __future__ import annotations

import argparse
import email
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
AGENT = REPO / "code" / "compliance_agent"
sys.path.insert(0, str(AGENT / "tools"))
sys.path.insert(0, str(HERE))

# A deny list (one word or phrase per line, "#" for comments): words that must never appear in a demo output, such as
# a real firm's, project's or person's names. It is kept OUTSIDE this repository on purpose (a list of real names in a
# public repository would itself be a leak): pass it with --deny or CVA_DEMO_DENY.
API = "http://127.0.0.1:{port}"


def call(port, path, data=None, raw=None, method=None):
    req = urllib.request.Request(API.format(port=port) + path, method=method or ("POST" if data or raw else "GET"),
                                 data=raw if raw is not None else (json.dumps(data).encode() if data else None),
                                 headers={"X-Compliance-App": "1", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read() or b"{}")


def wait_ui(port, seconds=60):
    t = time.time()
    while time.time() - t < seconds:
        try:
            return call(port, "/api/state")
        except Exception:
            time.sleep(0.5)
    raise SystemExit("the UI didn't start")


def run_all(port, projects) -> dict:
    res = call(port, "/api/run", {"plan": [{"project": p, "tasks": ["payroll", "intent"]} for p in projects],
                                  "where": "local", "force": False})
    job = res.get("job") or res.get("id")
    t = time.time()
    while time.time() - t < 900:
        q = call(port, "/api/queue")
        if not q["running"] and not q["waiting"] and q["finished"]:
            return q["finished"][0]
        time.sleep(2)
    raise SystemExit(f"the run (job {job}) didn't finish in 15 minutes")


def texts(path: Path):
    """Every piece of text in a file the run left: workbook parts (zip members and their names), .eml parts,
    plain text. A file that can't be read is reported (fail closed)."""
    suf = path.suffix.lower()
    try:
        if suf in (".xlsx", ".xlsm", ".zip", ".docx"):
            with zipfile.ZipFile(path) as z:
                for n in z.namelist():
                    yield n
                    yield z.read(n).decode("utf-8", "replace")
            return
        if suf == ".eml":
            msg = email.message_from_bytes(path.read_bytes())
            for k, v in msg.items():
                yield f"{k}: {v}"
            for part in msg.walk():
                if part.get_filename():
                    yield part.get_filename()
                payload = part.get_payload(decode=True)
                if payload:
                    if part.get_content_maintype() == "text":
                        yield payload.decode(part.get_content_charset() or "utf-8", "replace")
                    elif (part.get_filename() or "").lower().endswith((".xlsx", ".zip")):
                        tmp = Path(tempfile.mkdtemp()) / "a.xlsx"
                        tmp.write_bytes(payload)
                        yield from texts(tmp)
            return
        if suf in (".png", ".jpg", ".pdf", ".gz"):
            yield path.name
            return
        yield path.read_bytes().decode("utf-8", "replace")
    except Exception as e:                                      # noqa: BLE001
        yield f"UNREADABLE {type(e).__name__}"


def privacy(demo: Path, deny) -> list:
    """[(file, what)] for anything real found in the demo folder."""
    pat = re.compile("|".join(r"(?<![A-Za-z0-9])" + re.escape(w) + r"(?![A-Za-z0-9])" for w in deny), re.I) if deny else None
    found = []
    workers = json.loads((demo / "demo_data" / "workers.json").read_text())
    ids = {str(w["id"]) for w in workers}
    for f in sorted(p for p in demo.rglob("*") if p.is_file()):
        rel = f.relative_to(demo).as_posix()
        if rel.startswith((".home/ai_review",)):
            continue
        for t in [rel, *texts(f)]:
            if t.startswith("UNREADABLE"):
                found.append((rel, t))
                continue
            if pat and (m := pat.search(t)):
                found.append((rel, f"deny-listed word {m[0]!r}"))
            for m in re.finditer(r"(?<!\d)73\d{5}(?!\d)", t):
                if m[0].startswith("7320") and m[0] not in ids:
                    found.append((rel, f"employee ID {m[0]} not in the demo's roster"))
    return sorted(set(found))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Run and score the CVA demo.")
    ap.add_argument("--demo", type=Path, help="a demo folder to use (default: a fresh one in a temporary folder)")
    ap.add_argument("--port", type=int, default=8790)
    ap.add_argument("--lni-port", type=int, default=8791)
    ap.add_argument("--ai", choices=("evidence", "claude"), default="evidence")
    ap.add_argument("--keep", action="store_true", help="keep the temporary demo folder")
    ap.add_argument("--deny", type=Path, default=os.environ.get("CVA_DEMO_DENY"),
                    help="a file of words that must not appear in any output (kept outside the repository)")
    a = ap.parse_args(argv)
    demo = (a.demo or Path(tempfile.mkdtemp(prefix="cva-demo-")) / "CVA_demo").expanduser().resolve()
    ui = subprocess.Popen([sys.executable, str(HERE / "run_demo.py"), "--demo", str(demo), "--reset",
                           "--no-browser", "--port", str(a.port), "--lni-port", str(a.lni_port), "--ai", a.ai],
                          stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    try:
        state = wait_ui(a.port)
        projects = [p["name"] for p in state["projects"]]
        for f in sorted((demo / "To upload").glob("*")):
            large = next(p for p in projects if p.startswith("Norvanta"))
            from urllib.parse import quote
            call(a.port, f"/api/upload?project={quote(large)}&name={quote(f.name)}", raw=f.read_bytes())
        job = run_all(a.port, projects)
        print(f"run finished: {job.get('outcome')} ({len(job.get('results') or [])} tasks)")
    finally:
        ui.terminate()
        try:
            ui.wait(10)
        except subprocess.TimeoutExpired:
            ui.kill()
    import run_evals as RE
    # The harness runs with the AI review off; here it is on and adds a "Review decisions" sheet to each reviewed
    # payroll, which the scorer must skip like the agent's other sheets
    RE.AGENT_SHEETS = set(RE.AGENT_SHEETS) | {"Review decisions"}
    key = json.loads((demo / "demo_data" / "answer_key.json").read_text())
    requests = json.loads((demo / "demo_data" / "lni_requests.json").read_text()) \
        if (demo / "demo_data" / "lni_requests.json").is_file() else []
    out = RE.score(demo, key, run=demo, requests=requests)
    report = RE.report(out, {"checkout": "demo", "seconds": 0})
    print(report)
    deny = [w.strip() for w in Path(a.deny).expanduser().read_text().splitlines()
            if w.strip() and not w.startswith("#")] if a.deny else []
    leaks = privacy(demo, deny)
    print(f"\nPrivacy search ({f'{len(deny)} deny-listed words and ' if deny else 'no deny list given (--deny): '}employee IDs): "
          + ("nothing real found" if not leaks else f"{len(leaks)} finding(s)"))
    for rel, what in leaks[:50]:
        print(f"  - {rel}: {what}")
    if not a.keep and not a.demo:
        shutil.rmtree(demo.parent, ignore_errors=True)
    else:
        print(f"demo folder kept: {demo}")
    return 0 if out.get("failures", 1) == 0 and not leaks else 1


if __name__ == "__main__":
    sys.exit(main())
