"""Shows the demo's files as the cloud side (AWS) would see them: the CVA's real tokenizer, run offline.

    python demo/show_tokenized.py                      # tokenizes ~/Downloads/CVA_demo, opens the comparison page
    python demo/show_tokenized.py --demo /tmp/d --no-browser

In production the tokenizer runs in the firm's own Google Cloud project: it copies each file to a tokenized drive for
AWS, with workers' names, IDs and anything else the agent does not read replaced by keyed, reversible tokens, and
puts the names back into the results on the way home. Here the same code runs on this computer, with the settings
the firm uses (code/tokenizer/tokenizer_config.yaml), on the demo's invented files, in two passes as the mirror
does (the first learns names, the second writes).

**Demo key only:** the tokens are made with the tests' fixed key (32 bytes 0..31), never a real one, so they are
the same on every computer and every run. Nothing is uploaded anywhere.

It writes:
    <demo>/AWS view/Projects/...      each file as AWS would hold it (file and folder names tokenized too)
    <demo>/AWS view/compare.html      one payroll's first lines, as staff see them and as AWS sees them, and every
                                      file's name on both sides
"""
from __future__ import annotations

import argparse
import html
import shutil
import sys
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOKENIZER = HERE.parent / "code" / "tokenizer"
sys.path.insert(0, str(TOKENIZER))

DEFAULT_DEMO = Path.home() / "Downloads" / "CVA_demo"
SHOW = ["Contractor Name", "Employee Name", "Employee ID (LCPtracker)", "Craft", "Classification", "ST Hours",
        "Base Hourly", "3rd Party Fringe Rate"]
DEMO_KEY = bytes(range(32))


def tokenize(demo: Path, out: Path) -> list[tuple[str, str, str]]:
    """[(staff path, AWS path, outcome)] for every file under <demo>/Projects."""
    from pii_tokenizer import config, screen
    from pii_tokenizer.files import Context, tokenize_file
    from pii_tokenizer.names import aws_rel
    from pii_tokenizer.tokens import Vault
    cfg = config.load(TOKENIZER / "tokenizer_config.yaml")
    ctx = Context.build(cfg, DEMO_KEY, Vault())
    files = sorted(f for f in (demo / "Projects").rglob("*") if f.is_file() and not f.name.startswith("."))
    rows = []
    for final in (False, True):                       # pass 1 learns names, pass 2 writes (as the mirror does)
        for f in files:
            rel = f.relative_to(demo)
            how = ctx.config.handling(f.name)
            try:
                if how == "scan_pdf":
                    data = f.read_bytes()
                    screen.screen_pdf(rel.as_posix(), data, ctx.scrubber)
                    name, data_out = f.name, data
                else:
                    name, data_out, _ = tokenize_file(f.name, f.read_bytes() if how != "skip" else b"", ctx,
                                                      rel=rel.as_posix())
                    if data_out is not None and screen.residual(name, data_out, ctx.scrubber):
                        raise screen.Held("residual")
            except Exception as e:                      # noqa: BLE001 - held back, as the mirror would
                if final:
                    rows.append((rel.as_posix(), "", f"held back ({type(e).__name__})"))
                continue
            if not final:
                continue
            if data_out is None:
                rows.append((rel.as_posix(), "", "not copied (file rules)"))
                continue
            aws = aws_rel((rel.parent / name).as_posix(), ctx.scrubber, ctx.config)
            target = out / aws
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data_out)
            rows.append((rel.as_posix(), aws, "copied, tokenized"))
    return rows


def first_lines(path: Path, n: int = 10):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h or "") for h in rows[0]]
    cols = [head.index(c) for c in SHOW if c in head]
    return [head[i] for i in cols], [[r[i] for i in cols] for r in rows[1:n + 1]]


def table(head, rows) -> str:
    h = "".join(f"<th>{html.escape(str(x))}</th>" for x in head)
    b = "".join("<tr>" + "".join(f"<td>{html.escape('' if v is None else str(v))}</td>" for v in r) + "</tr>"
                for r in rows)
    return f"<table><tr>{h}</tr>{b}</table>"


def page(demo: Path, out: Path, rows) -> Path:
    payroll = next(((s, a) for s, a, _ in rows if a and "PayrollDetailReport" in s and s.endswith(".xlsx")), None)
    parts = ["<h1>The demo's files, as staff see them and as AWS sees them</h1>",
             "<p>Same CVA tokenizer and settings as in production, run on this computer on invented data, with a "
             "fixed demo key. Names, IDs and every column the agent doesn't read become reversible tokens; the "
             "hours, rates and trade the checks need stay readable.</p>"]
    if payroll:
        sh, sr = first_lines(demo / payroll[0])
        ah, ar = first_lines(out / payroll[1])
        parts += [f"<h2>Staff side: {html.escape(payroll[0].split('/')[-1])}</h2>", table(sh, sr),
                  f"<h2>AWS side: {html.escape(payroll[1].split('/')[-1])}</h2>", table(ah, ar)]
    parts += ["<h2>Every file, both sides</h2>",
              table(["Staff side (Google)", "AWS side", "Outcome"], [(s, a or "-", o) for s, a, o in rows])]
    css = ("body{font:14px -apple-system,Segoe UI,Helvetica,Arial,sans-serif;margin:28px;color:#1d2522}"
           "h1{font-size:22px}h2{font-size:16px;margin-top:26px}table{border-collapse:collapse;margin-top:8px}"
           "td,th{border:1px solid #d6dcd9;padding:4px 8px;text-align:left;white-space:nowrap}"
           "th{background:#eef3f0}p{max-width:900px;color:#3c4743}")
    f = out / "compare.html"
    f.write_text(f"<!doctype html><meta charset='utf-8'><title>CVA demo: AWS view</title><style>{css}</style>"
                 + "".join(parts))
    return f


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Show the demo's files as AWS would see them.")
    ap.add_argument("--demo", type=Path, default=DEFAULT_DEMO)
    ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args(argv)
    demo = a.demo.expanduser().resolve()
    if not (demo / "demo_data" / "answer_key.json").is_file():
        raise SystemExit(f"{demo} isn't a demo folder: run demo/run_demo.py first")
    out = demo / "AWS view"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir()
    rows = tokenize(demo, out)
    f = page(demo, out, rows)
    print(f"{sum(1 for r in rows if r[1])} files tokenized into {out}; comparison: {f}")
    if not a.no_browser:
        webbrowser.open(f.as_uri())
    return 0


if __name__ == "__main__":
    sys.exit(main())
