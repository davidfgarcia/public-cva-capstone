"""AI review of orange flags (agent/ai_review.py, agent/knowledge.py): mode selection, patterns, the knowledge base,
Claude suggestions through a FAKE client (no network, no key), reuse of answers, the call limit, shadow vs shown, and
staff decisions read back from the "Review decisions" sheet. Every regulation text, name and company here is
invented.

    python tests/test_ai_review.py
"""
import datetime as dt
import json
import os
import shutil
import sys
import tempfile
import types
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent
sys.path.insert(0, str(AGENT))
sys.path.insert(0, str(HERE))

import test_layout as T                                          # noqa: E402
from agent import ai_review as A                                 # noqa: E402
from agent import knowledge as K                                 # noqa: E402
from agent.config import load_config                             # noqa: E402
from agent.layout import Paths, ensure_layout                    # noqa: E402
from agent.payroll import Issue, RULE_TITLES                     # noqa: E402

RCW_HTML = """<html><head><title>x</title></head><body><div>Chapter 39.12 RCW</div>
<div><h3>RCW 39.12.010</h3><span>Definitions.</span><p>"Prevailing rate of wage" means the rate of hourly wage, usual
benefits and overtime paid in the locality to the majority of workers in the same trade.</p></div>
<div><h3>RCW 39.12.020</h3><p>Overtime. The hourly wages to be paid to laborers, workers, or mechanics upon all public
works shall be not less than the prevailing rate of wage for an hourly work day of eight hours.</p></div>
<div><h3>RCW 39.12.030</h3><p>Specifications to state minimum hourly rate. The specifications for every contract shall
contain a provision stating the hourly minimum rate of wage.</p></div></body></html>"""

ECFR_XML = """<DIV5 N="5" TYPE="PART"><HEAD>PART 5</HEAD>
<DIV8 N="5.2" TYPE="SECTION"><HEAD>§ 5.2 Definitions.</HEAD><P>Laborer or mechanic includes at least those workers
whose duties are manual or physical in nature.</P></DIV8>
<DIV8 N="5.5" TYPE="SECTION"><HEAD>§ 5.5 Contract provisions.</HEAD><P>Overtime requirements. No contractor shall
require any laborer to work more than forty hours in any workweek unless paid at one and one-half times the basic rate
for all hours worked in excess of forty hours.</P></DIV8>
<DIV8 N="5.6" TYPE="SECTION"><HEAD>§ 5.6 Enforcement.</HEAD><P>The agency shall cause investigations to be made as may
be necessary to assure compliance with the labor standards clauses.</P></DIV8></DIV5>"""


class Resp:
    def __init__(self, text=None, js=None):
        self.text, self._js = text, js

    def json(self):
        return self._js


def fake_getter(calls):
    def get(url, params=None, headers=None, timeout=60, what=""):
        calls.append(url)
        if "titles.json" in url:
            return Resp(js={"titles": [{"number": 29, "up_to_date_as_of": "2026-10-20"}]})
        if "ecfr.gov" in url:
            return Resp(text=ECFR_XML)
        if "RCW" in url:
            return Resp(text=RCW_HTML)
        return Resp(text=RCW_HTML.replace("RCW 39.12.0", "WAC 296-127-0"))
    return get


class FakeClaude:
    """messages.create returns a reply starting 'FLAG:' (or 'NO FLAG:') citing the first search result."""

    def __init__(self, verdict="FLAG", fail_after=None):
        self.n, self.verdict, self.fail_after = 0, verdict, fail_after
        self.messages = self

    def create(self, model, max_tokens, messages, thinking=None, extra_body=None):
        self.n += 1
        self.last = {"max_tokens": max_tokens, "thinking": thinking, "extra_body": extra_body,
                     "prompt": messages[0]["content"][-1]["text"]}
        if self.fail_after is not None and self.n > self.fail_after:
            raise ConnectionError("network down")
        blocks = [b for b in messages[0]["content"] if b.get("type") == "search_result"]
        cite = types.SimpleNamespace(title=blocks[0]["title"], source=blocks[0]["source"], cited_text="...")
        verdict = self.verdict(thinking) if callable(self.verdict) else self.verdict
        text = types.SimpleNamespace(type="text", text=f"{verdict}: overtime rules apply to these hours.",
                                     citations=[cite])
        content = [text]
        if thinking:                     # a thinking block comes first, as the API returns it
            content.insert(0, types.SimpleNamespace(type="thinking", thinking="source, date, rate, verdict"))
        usage = types.SimpleNamespace(input_tokens=3000, output_tokens=1500 if thinking else 300,
                                      output_tokens_details=types.SimpleNamespace(thinking_tokens=1200 if thinking else 0))
        return types.SimpleNamespace(content=content, usage=usage)


TOKENS = ["EMP-0123456789", "EMP-abcdef0123"]     # employee values as the tokenizer leaves them


def new_run():
    """Each agent run has its own call budget, time budget and knowledge index."""
    A._RUN = None


def issue(rule, row, craft="Laborers", cls="General Laborer", decision="WA L&I 2025", detail="45 ST hours"):
    return Issue("orange", rule, row, "classification", detail,
                 {"kind": "state", "craft": craft, "classification": cls, "wage_decision": decision,
                  "contractor": "Acme Paving", "week": dt.date(2026, 9, 6), "employee": "Rivera, Ana"})


def project(root: Path, name="Alpha"):
    shutil.copy(AGENT / "config.yaml", root / "config.yaml")
    p = root / "Projects" / name
    ensure_layout(p)
    cfg = load_config(AGENT / "config.yaml", root)
    cfg["ai_review"] = {**(cfg.get("ai_review") or {}), "cache_dir": str(root / "_cache")}
    return p, cfg


class Env:
    """Set and restore environment variables, and pretend the anthropic package is installed."""

    def __init__(self, **kv):
        self.kv, self.old = kv, {}

    def __enter__(self):
        for k, v in self.kv.items():
            self.old[k] = os.environ.get(k)
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        self.had = sys.modules.get("anthropic")
        sys.modules["anthropic"] = types.SimpleNamespace(Anthropic=lambda **k: FakeClaude())
        return self

    def __exit__(self, *a):
        for k, v in self.old.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        if self.had is None:
            sys.modules.pop("anthropic", None)
        else:
            sys.modules["anthropic"] = self.had


def test_mode_is_chosen_without_any_setting():
    cfg = {"ai_review": {}}
    with Env(ANTHROPIC_API_KEY=None, SOURCE_REMOTE=None):
        s = A.settings(cfg)
        assert s["on"] and s["mode"] == "local" and s["shadow"] is True, s
    with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE=None):
        assert A.settings(cfg)["mode"] == "local"                 # desktop / local runs never call Claude by default
    with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
        assert A.settings(cfg)["mode"] == "claude"                # the AWS pipeline with a key
    with Env(ANTHROPIC_API_KEY=None, SOURCE_REMOTE="gdrive:compliance"):
        s = A.settings(cfg)
        assert s["mode"] == "local" and "no ANTHROPIC_API_KEY" in s["why"], s
    assert not A.settings({"ai_review": {"enabled": False}})["on"]


def test_patterns_group_lines():
    iss = [issue("S05", r) for r in range(2, 12)] + [issue("S05", 20, cls="Flagger")] + \
          [issue("S03", 30)] + [Issue("red", "R02", 40, "classification", "x", {})]
    pats = A.group_patterns(iss, "Alpha", {"S05", "S11"})
    assert [(p["rule"], p["classification"], p["lines"]) for p in pats] == \
        [("S05", "General Laborer", 10), ("S05", "Flagger", 1)], pats
    assert pats[0]["id"] == A.pattern_id("Alpha", "S05", "state", "Laborers", "General Laborer", "WA L&I 2025")


def test_regulations_are_downloaded_split_and_kept():
    with tempfile.TemporaryDirectory() as d:
        calls = []
        chunks, notes = K.regulation_chunks(Path(d), getter=fake_getter(calls))
        assert not notes, notes
        cites = {c.citation for c in chunks}
        assert {"§ 5.5", "RCW 39.12.020", "WAC 296-127-020"} <= cites, cites
        assert all(c.source_type == "regulation" for c in chunks)
        fed = next(c for c in chunks if c.citation == "§ 5.5")
        assert fed.wage_source == "federal" and "forty hours" in fed.text
        # fresh copies: nothing downloaded again
        n = len(calls)
        K.regulation_chunks(Path(d), getter=fake_getter(calls))
        assert len(calls) == n
        # stale copy and the site down: the saved copy is used, tried at most once a day
        for f in (Path(d) / "regulations").glob("*.json"):
            js = json.loads(f.read_text())
            js["fetched"] = "2020-01-01"
            f.write_text(json.dumps(js))

        def down(*a, **k):
            raise ConnectionError("blocked")
        chunks2, notes2 = K.regulation_chunks(Path(d), getter=down)
        assert len(chunks2) == len(chunks) and all("was used" in x for x in notes2), notes2
        chunks3, notes3 = K.regulation_chunks(Path(d), getter=fake_getter(calls))
        assert len(calls) == n and all("tried earlier today" in x for x in notes3), notes3


def test_staff_documents_by_folder():
    with tempfile.TemporaryDirectory() as d:
        kd = Path(d)
        (kd / "contracts" / "Alpha" / "union").mkdir(parents=True)
        (kd / "contracts" / "Alpha" / "union" / "CBA.txt").write_text(
            "ARTICLE 7 - OVERTIME\nWork over eight hours in a day is paid at time and one-half.\n\n"
            "ARTICLE 8 - HOLIDAYS\nWork on a recognized holiday is paid at double time.")
        (kd / "guidance").mkdir()
        (kd / "guidance" / "notes.md").write_text("Some guidance about apprentices and ratios on the job site.\n")
        (kd / "letter.docx").write_text("x")
        chunks, notes = K.document_chunks(kd, ["Alpha"])
        cba = [c for c in chunks if c.title.startswith("CBA")]
        assert len(cba) == 2 and all(c.source_type == "contract" and c.project_id == "Alpha" and c.wage_source == "union"
                                     for c in cba), cba
        g = [c for c in chunks if c.title.startswith("notes")]
        assert {c.wage_source for c in g} == {"federal", "state", "union"} and g[0].source_type == "guidance"
        assert any("letter.docx" in n for n in notes), notes


def _setup_kb(root):
    kd = root / "reference" / "knowledge" / "regulations" / "state"
    kd.mkdir(parents=True)
    (kd / "overtime.txt").write_text(
        "Section 1 Overtime. Hours over forty in a workweek and over eight in a day are paid at one and one-half "
        "times the hourly rate for laborers on public works.\n\nSection 2 Classifications. Workers are paid the rate "
        "of the classification of the work actually performed, such as general laborer or flagger.")


def test_claude_suggestions_are_reused_limited_and_shadowed():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", r, detail="45 straight-time hours in the week and no overtime") for r in range(2, 6)] + \
              [issue("S01", 9, cls="Flagger", detail="9 straight-time hours on Monday (over 8)")]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            fake = FakeClaude()
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake)
            assert res["mode"] == "claude" and res["shadow"] and fake.n == 2, (res["lines"], fake.n)
            assert all(x["suggestion"]["raise_flag"] for x in res["patterns"])
            assert "2 new Claude call(s), 0 reused" in res["lines"][0] and "shadow" in res["lines"][0], res["lines"]
            # shadow: the cells' notes and the sheet don't show it
            A.annotate_issues(res)
            assert "AI suggestion" not in iss[0].detail
            wb = openpyxl.Workbook()
            A.decisions_sheet(wb, res)
            ws = wb[K.DECISIONS_SHEET]
            rows = list(ws.iter_rows(min_row=5, values_only=True))
            assert len(rows) == 2 and rows[0][10] is None and rows[0][11] is None and rows[0][12], rows
            # the same evidence next run: answers reused, no call
            fake2 = FakeClaude()
            new_run()
            res2 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake2)
            assert fake2.n == 0 and "0 new Claude call(s), 2 reused" in res2["lines"][0], res2["lines"]
            # a limit of 1 call, then an outage: one suggestion, a warning, the rest without
            st = p / ".agent" / A.STATE_FILE
            st.unlink()
            cfg["ai_review"]["max_calls_per_run"] = 1
            new_run()
            res3 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude())
            assert any("1 pattern(s) got no suggestion" in l and "max_calls_per_run" in l for l in res3["lines"]), res3["lines"]
            st.unlink()
            cfg["ai_review"]["max_calls_per_run"] = 200
            new_run()
            res4 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude(fail_after=0))
            assert any("a Claude call failed" in l for l in res4["lines"]), res4["lines"]
            # an answer without a clear verdict is not used or kept; a bolded one is
            st.unlink()
            new_run()
            res6 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS,
                            client=FakeClaude("Based on the sources, FLAG"))
            assert not any(x["suggestion"] for x in res6["patterns"]) and "without a clear verdict" in res6["lines"][0]
            assert not json.loads(st.read_text())["suggestions"], "an unclear answer was kept"
            new_run()
            res7 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("**FLAG**"))
            assert all(x["suggestion"]["raise_flag"] for x in res7["patterns"]), res7["lines"]
            st.unlink()
            # shown (shadow off): the sheet and the cells' notes carry it
            cfg["ai_review"]["shadow"] = False
            new_run()
            res5 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("NO FLAG"))
            A.annotate_issues(res5)
            assert "AI suggestion: likely fine" in iss[0].detail
            wb = openpyxl.Workbook()
            A.decisions_sheet(wb, res5)
            r0 = list(wb[K.DECISIONS_SHEET].iter_rows(min_row=5, values_only=True))[0]
            assert r0[10] == "Likely fine" and "Sources:" in r0[11], r0


def test_the_run_log_gets_the_reviews_calls_and_tokens_summed_from_the_replies():
    """docs/run_log.md: the ai_review line's tokens are the replies' own usage, added up (judge_cell's
    usage.input / usage.output), so a renamed key can't leave them at 0 unnoticed."""
    from unittest import mock
    from agent import events
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", r, detail="45 straight-time hours in the week and no overtime") for r in range(2, 6)] + \
              [issue("S01", 9, cls="Flagger", detail="9 straight-time hours on Monday (over 8)")]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"), \
                mock.patch.object(events, "emit") as emit:
            fake = FakeClaude()
            new_run()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake)
        (call,) = [c for c in emit.call_args_list if c.args[0] == "ai_review"]
        f = call.kwargs
        assert fake.n == 2 and f["ai.calls"] == 2 and f["ai.patterns"] == 2 and f["ai.mode"] == "claude"
        assert f["gen_ai.usage.input_tokens"] == 2 * 3000 and f["gen_ai.usage.output_tokens"] == 2 * 300
        assert f["gen_ai.request.model"] == cfg["ai_review"].get("model", A.DEFAULTS["model"]) and f["ai.stopped"] is None


def test_untokenized_payroll_is_never_sent():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", 2, detail="45 straight-time hours in the week")]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            fake = FakeClaude()
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=["Rivera, Ana"], client=fake)
            assert fake.n == 0 and res["mode"] == "local" and "isn't tokenized" in res["lines"][0], res["lines"]
            cfg["ai_review"]["mode"] = "claude"                  # set explicitly: still refused without allow_untokenized
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=[], client=fake)
            assert fake.n == 0 and res["mode"] == "local"


def test_names_never_reach_the_query():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", 2, detail="Rivera, Ana: 45 straight-time hours in the week")]
        cfg["ai_review"]["allow_untokenized"] = True      # a staff computer with Claude explicitly allowed
        seen = []

        class Spy(FakeClaude):
            def create(self, model, max_tokens, messages):
                seen.append(json.dumps(messages))
                return super().create(model, max_tokens, messages)
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            new_run()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=["Rivera, Ana"], client=Spy())
        assert seen and "Rivera" not in seen[0] and "[EMPLOYEE]" in seen[0]


def test_staff_decisions_are_read_back_carried_over_and_compared():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", r, detail="45 straight-time hours in the week and no overtime") for r in range(2, 6)]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("FLAG"))
            wb = openpyxl.Workbook()
            A.decisions_sheet(wb, res)
            ws = wb[K.DECISIONS_SHEET]
            ws.cell(5, 9, "Keep")
            ws.cell(5, 10, "Paid straight time over 40: back pay owed")
            out = Paths(p).reviewed / "2026-10-20 Week 36 - REVIEWED.xlsx"
            wb.save(out)
            dec = K.read_decisions(root / "Projects")
            pid = res["patterns"][0]["id"]
            assert dec[pid]["decision"] == "keep" and "back pay" in dec[pid]["note"], dec
            # next review: the decision is carried into the new sheet, used as evidence, and compared
            new_run()
            res2 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("FLAG"))
            assert res2["patterns"][0]["decision"]["decision"] == "keep"
            # the decision doesn't change the cache key (no paid call), and agreement waits for enough decisions
            assert "1 reused" in res2["lines"][0] and "agreement with staff shown from 10 decisions (1 so far)" \
                in res2["lines"][0], res2["lines"]
            cfg["ai_review"]["min_decisions_for_agreement"] = 1
            new_run()
            res2 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("FLAG"))
            assert "staff decisions agree with it on 1 of 1" in res2["lines"][0], res2["lines"]
            # a later answer (here: after the evidence changed) never replaces the first one in the comparison
            st = p / ".agent" / A.STATE_FILE
            js = json.loads(st.read_text())
            js["suggestions"][pid]["cache_key"] = "changed"
            st.write_text(json.dumps(js))
            new_run()
            res3 = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("NO FLAG"))
            assert res3["patterns"][0]["suggestion"]["raise_flag"] is False
            assert "staff decisions agree with it on 1 of 1" in res3["lines"][0], res3["lines"]
            wb2 = openpyxl.Workbook()
            A.decisions_sheet(wb2, res2)
            r0 = list(wb2[K.DECISIONS_SHEET].iter_rows(min_row=5, values_only=True))[0]
            assert r0[8] == "Keep" and "back pay" in r0[9], r0
            chunks = K.decision_chunks(dec)
            assert chunks[0].source_type == "history" and chunks[0].valid_from is None


def test_a_payroll_run_in_local_mode_adds_the_sheet_and_calls_nothing():
    from agent.runner import run
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        cfg["projects"] = {"Alpha": {"intent_applicable": False}}
        shutil.copy(T.SAMPLES / "state_LNI_Pierce_sample.csv", Paths(p).upload / "Journey level wages_Pierce County_6.26.25 LNI.csv")
        run("wages", ["Alpha"], cfg=cfg, notify=False)
        pay = T.payroll_xlsx(Paths(p).upload / "Week 36.xlsx", jurisdiction="State - Pierce")
        wb = openpyxl.load_workbook(pay)
        ws = wb.active
        v = [c.value for c in ws[2]]
        v[8], v[11] = 45, 45                      # 45 straight-time hours in the week, no overtime: S05
        ws.append(v)
        wb.save(pay)
        with Env(ANTHROPIC_API_KEY=None, SOURCE_REMOTE=None):
            rep = run("payroll", ["Alpha"], cfg=cfg, notify=False)
        lines = rep["projects"]["Alpha"]
        ai = [l for l in lines if "ai review" in l]
        assert ai and ai[0].startswith("payroll:   ai review (local; auto: not on the AWS pipeline): 1 pattern(s) of S05"), lines
        out = next(Paths(p).reviewed.glob("*REVIEWED.xlsx"))
        wb = openpyxl.load_workbook(out)
        assert K.DECISIONS_SHEET in wb.sheetnames, wb.sheetnames
        assert not (p / ".agent" / A.STATE_FILE).exists()      # local mode keeps no suggestions


def test_call_limit_and_a_failure_hold_for_the_whole_run():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        q = root / "Projects" / "Beta"
        ensure_layout(q)
        _setup_kb(root)
        cfg["ai_review"].update({"download_regulations": False, "max_calls_per_run": 1})
        iss = [issue("S05", 2, detail="45 straight-time hours in the week")]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            new_run()
            f1, f2 = FakeClaude(), FakeClaude()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=f1)
            r2 = A.review(q, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=f2)
            assert f1.n == 1 and f2.n == 0 and any("max_calls_per_run" in l for l in r2["lines"]), r2["lines"]
            cfg["ai_review"]["max_calls_per_run"] = 200
            new_run()
            bad, ok = FakeClaude(fail_after=0), FakeClaude()
            (p / ".agent" / A.STATE_FILE).unlink()
            (q / ".agent" / A.STATE_FILE).unlink()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=bad)
            r3 = A.review(q, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=ok)
            assert ok.n == 0 and any("a Claude call failed" in l for l in r3["lines"]), r3["lines"]


def test_parsing_is_careful():
    cite = r"RCW\s+39\.12\.\d{3}[a-z]?"
    txt = ("[PDF] RCW 39.12.010 Definitions.\nThe prevailing rate of wage means the usual rate in the locality.\n"
           "See RCW 39.12.020 for overtime, which applies widely and at some length in the statute body.\n"
           "RCW 39.12.020\nOvertime: hours over eight in a day are paid at the overtime rate on public works.\n")
    secs = dict(K.split_sections(txt, cite))
    assert set(secs) == {"RCW 39.12.010", "RCW 39.12.020"}, secs
    assert secs["RCW 39.12.020"].startswith("RCW 39.12.020\nOvertime"), secs
    long = "§ 5.5 Contract provisions.\n" + "\n".join(f"({i}) " + "Overtime words. " * 80 for i in range(40))
    parts = K._pieces(long)
    assert len(parts) > 5 and all(len(x) <= 4400 and x.startswith("§ 5.5") for x in parts)
    assert "(39)" in parts[-1]                                   # nothing cut off at the end
    with tempfile.TemporaryDirectory() as d:
        kd = Path(d)
        (kd / "memo.txt").write_text("A long memo about apprentices.\n\n" + "Ratio rules apply. " * 200 +
                                     "\n\nSection 9 Contact\nCall the office.")
        chunks, _ = K.document_chunks(kd)
        assert sum(len(c.text) for c in chunks if c.wage_source == "state") > 3000, "the memo was cut"
        cache = Path(d) / "c"
        (cache / "regulations").mkdir(parents=True)
        (cache / "regulations" / "CFR-29-5.json").write_text('{"fetched": "2026-')   # a damaged copy
        chunks, notes = K.regulation_chunks(cache, getter=fake_getter([]))
        assert any(c.citation == "§ 5.5" for c in chunks), notes


def test_the_real_ecfr_xml_parses():
    """An excerpt of the real eCFR XML for 29 CFR Part 5 (saved 2026-10-25): one chunk per section."""
    secs = K.ecfr_sections((HERE / "fixtures" / "ecfr_29_cfr_part5_excerpt.xml").read_text())
    assert [c for c, _ in secs] == ["§ 5.1", "§ 5.2"], secs         # "§§ 5.3-5.4 [Reserved]" is left out
    assert all(len(b) > 100 for _, b in secs)


def test_the_real_wa_legislature_pages_parse():
    """Excerpts of the real app.leg.wa.gov pages (saved 2026-10-25, scripts and attributes removed): the section
    headings are found, the table of contents is not taken for sections, WAC citations get their 'WAC ' prefix."""
    fx = HERE / "fixtures"
    pages = {"RCW": (fx / "rcw_39_12_page_excerpt.html").read_text(), "WAC": (fx / "wac_296_127_page_excerpt.html").read_text()}

    def get(url, **k):
        return types.SimpleNamespace(text=pages["RCW" if "RCW" in url else "WAC"])
    with tempfile.TemporaryDirectory() as d:
        chunks, notes = K.regulation_chunks(Path(d), getter=get, regs=K.REGULATIONS[1:])
    assert not notes, notes
    cites = {c.citation for c in chunks}
    assert {"RCW 39.12.010", "RCW 39.12.017", "RCW 39.12.020", "WAC 296-127-010", "WAC 296-127-011"} <= cites, cites
    when = next(c for c in chunks if c.citation == "WAC 296-127-011" and "bids by prime contractors" in c.text)
    assert "remain in effect for the duration of the contract" in when.text
    res = next(c for c in chunks if c.citation == "RCW 39.12.017")
    assert "residential" in res.text.lower()
    assert all(len(c.text) <= 4400 for c in chunks)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)


def test_extended_thinking_is_asked_for_and_changes_the_cache_key():
    """thinking_effort set: the call asks for adaptive thinking at that effort with room for the answer, the prompt lists the
    reasoning steps, the visible answer is read as before, and answers made without thinking aren't reused."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", 2, detail="45 straight-time hours in the week and no overtime")]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            plain = FakeClaude()
            new_run()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=plain)
            assert plain.last["thinking"] is None and plain.last["max_tokens"] == 4000
            assert "Before answering" not in plain.last["prompt"]
            cfg["ai_review"]["thinking_effort"] = "medium"
            think = FakeClaude()
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=think)
            assert think.n == 1, "an answer made without thinking must not be reused"
            assert think.last["thinking"] == {"type": "adaptive"}             # current models: adaptive only
            assert think.last["extra_body"] == {"output_config": {"effort": "medium"}}
            assert think.last["max_tokens"] == 16000 and "which date governs" in think.last["prompt"]
            sg = res["patterns"][0]["suggestion"]
            assert sg["raise_flag"] is True and sg["citations"] and "source, date" not in sg["explanation"]


def test_the_thinking_experiment_compares_and_reports():
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", r, cls=c, detail="45 straight-time hours in the week and no overtime")
               for r, c in ((2, "General Laborer"), (3, "Flagger"), (4, "Pipe Layer"))]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance"):
            new_run()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=FakeClaude("FLAG"))
            # thinking says NO FLAG: every verdict changes; the suggestions in use are not touched
            fake = FakeClaude(lambda thinking: "NO FLAG" if thinking else "FLAG")
            with Env(CVA_AI_EXPERIMENT="2"):
                new_run()
                res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake)
            assert fake.n == 4, fake.n                                # 2 patterns x (without + with), 0 normal calls
            line = next(l for l in res["lines"] if "experiment" in l)
            assert "2 pattern(s); verdict changed on 2 of 2 (0 to FLAG, 2 to NO FLAG)" in line, line
            assert "same verdict as the suggestion in use on 2 of 2" in line and "300 -> 1500 output tokens" in line
            assert all(x["suggestion"]["raise_flag"] for x in res["patterns"])     # suggestions unchanged
            runs = json.loads((p / ".agent" / A.EXPERIMENT_FILE).read_text())
            assert runs[0]["summary"]["changed_to_no_flag"] == 2 and len(runs[0]["patterns"]) == 2
            # without the option: no experiment calls
            fake2 = FakeClaude("FLAG")
            new_run()
            res = A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake2)
            assert fake2.n == 0 and not any("experiment" in l for l in res["lines"])


def test_evidence_gaps_are_counted_per_rule_and_marked_in_the_sheet():
    """Reviewer feedback: make the evidence gap clear. A pattern with no evidence gets no suggestion (as before),
    and its row in the Review decisions sheet says so, in shadow mode too; the email counts the gaps per rule."""
    import openpyxl
    pats = [{"rule": "S01", "no_evidence": True}, {"rule": "S01", "no_evidence": False},
            {"rule": "S11", "no_evidence": True}, {"rule": "S11", "no_evidence": None}]
    assert A.evidence_gaps(pats) == "S01 1 of 2 (50%), S11 1 of 1 (100%)"
    base = {"kind": "state", "craft": "Laborers", "classification": "General Laborer", "jurisdiction": "",
            "wage_decision": "WA L&I", "contractor": "Acme", "lines": 3, "example": "9 ST hours", "decision": None,
            "suggestion": None, "sources": ""}
    res = {"shadow": True, "mode": "claude",
           "patterns": [{**base, "rule": "S01", "id": "p1", "no_evidence": True},
                        {**base, "rule": "S05", "id": "p2", "no_evidence": False, "sources": "RCW 39.12.020"}]}
    wb = openpyxl.Workbook()
    A.decisions_sheet(wb, res)
    ws = wb["Review decisions"]
    rows = {r[12]: r for r in ws.iter_rows(min_row=5, values_only=True)}
    assert rows["p1"][11] == A.NO_EVIDENCE_TEXT and rows["p1"][10] is None        # the gap, and no suggestion
    assert rows["p2"][11] is None                                                  # shadow: sources stay hidden


def test_thinking_calls_get_a_longer_timeout_and_the_experiment_limit_holds_for_the_run():
    """Review of PR #35: a thinking call can take well over the client's 30 s, so it gets its own limit; and
    --ai-experiment N counts across every payroll of the run, not per payroll."""
    class Timed(FakeClaude):
        def with_options(self, timeout):
            self.timeout = timeout
            return self
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        p, cfg = project(root)
        q = root / "Projects" / "Beta"
        ensure_layout(q)
        _setup_kb(root)
        cfg["ai_review"]["download_regulations"] = False
        iss = [issue("S05", r_, cls=c, detail="45 straight-time hours in the week and no overtime")
               for r_, c in ((2, "General Laborer"), (3, "Flagger"))]
        with Env(ANTHROPIC_API_KEY="test-key", SOURCE_REMOTE="gdrive:compliance", CVA_AI_EXPERIMENT="2"):
            fake = Timed()
            new_run()
            A.review(p, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake)
            n1 = fake.n
            A.review(q, cfg, iss, kinds={"state"}, titles=RULE_TITLES, names=TOKENS, client=fake)
            # first payroll: 2 normal calls + 2 patterns x 2 experiment calls; second: 2 normal calls, no experiment
            assert n1 == 6 and fake.n == 8, (n1, fake.n)
            assert fake.timeout == 180
