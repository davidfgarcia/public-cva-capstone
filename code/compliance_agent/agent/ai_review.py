"""AI review of orange flags (retrieval and ranking: agent/compliance_rag.py; knowledge: agent/knowledge.py).

The payroll review's judgment-type orange flags (ai_review.flags; by default S01, S05, S11, S12, S13, S14: daily
and weekly overtime, classification matches, overtime pay, which determination applies) are grouped into patterns:
one per rule + wage type + craft/classification (+ the Jurisdiction text for S13/S14), however many lines share it.
For each pattern the agent finds the evidence (regulations, staff documents, staff decisions; filtered, searched two
ways, scored, then maximum relevance / minimum redundancy) and, in Claude mode, asks Claude for a suggestion with
citations: keep the flag or likely fine. A suggestion never turns a cell red or clears it; staff decide, in the
reviewed payroll's "Review decisions" sheet, and those decisions become evidence for later runs.

Modes (ai_review.mode, default auto):
  local   evidence only, all on this machine: nothing is sent anywhere
  claude  evidence + Claude's suggestion. Only when ANTHROPIC_API_KEY is set AND the payroll is proven tokenized
          (every employee value is a token such as EMP-7f3a91c2d0); auto also requires the AWS pipeline (SOURCE_REMOTE).
          run_pipeline.sh reads the key from AWS Secrets Manager (compliance-agent/anthropic-api-key) when it exists.
Shadow (ai_review.shadow, default true): suggestions are kept in the project's .agent/ai_review.json and only
counted in the email, not shown to staff; false shows them.
Costs and time stay bounded, for the whole run (all projects, all payroll files): one call per pattern, never per
line; a pattern's answer is reused while the law, documents and model are the same (up to ai_review.cache_days);
at most ai_review.max_calls_per_run calls and ai_review.max_seconds seconds of calls per run; the first failed call
stops further calls for the run. The payroll review always goes on."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import re
import time
from collections import Counter, defaultdict
from pathlib import Path

from .util import err_text, log


def _what(e, n) -> str:
    """An exception for a run line: its type, and off AWS its text (util.err_text, #56), without repeating the type."""
    t, text = type(e).__name__, err_text(e)[:n]
    return t if text == t else f"{t}: {text}"


try:
    from . import compliance_rag as R
except ImportError:                   # pragma: no cover
    import compliance_rag as R

DEFAULTS = {"enabled": "auto", "mode": "auto", "shadow": True, "flags": ["S01", "S05", "S11", "S12", "S13", "S14"],
            "max_calls_per_run": 200, "max_seconds": 300, "cache_days": 90, "model": "claude-sonnet-5-5",
            "embedder": "auto", "reranker": "auto", "knowledge_dir": "reference/knowledge",
            "cache_dir": "~/.compliance_agent/ai_review", "refresh_days": 30, "download_regulations": True,
            "allow_untokenized": False, "min_decisions_for_agreement": 10,
            "thinking_effort": "",         # low|medium|high|max: adaptive thinking (chain of thought) on every call
            "experiment_seconds": 1200}    # --ai-experiment: time allowed for its calls (on top of max_seconds)
# A pattern for which nothing was found in the regulations, documents or past decisions: said in the Review
# decisions sheet even in shadow mode (it's an evidence gap, not a suggestion), so staff know to decide unaided.
NO_EVIDENCE_TEXT = ("No supporting evidence found in the regulations, staff documents or past decisions: decide "
                    "from the payroll and the law (no AI suggestion is made without evidence).")
EXPERIMENT_FILE = "ai_thinking_experiment.json"    # in the project's .agent/ folder: --ai-experiment results
PROMPT_VERSION = "2"
STATE_FILE = "ai_review.json"          # in the project's .agent/ folder
TOKEN_RE = re.compile(r"^[A-Z]{2,5}-[0-9a-f]{10}$")
VERDICT_RE = re.compile(r"^\W*(NO FLAG|FLAG)\b", re.I)
JURISDICTION_RULES = {"S13", "S14"}    # patterns keyed by the Jurisdiction text as well


def settings(cfg) -> dict:
    """The resolved settings: s['on'], s['mode'] ('local' | 'claude'), s['why'] (how the mode was chosen).
    'claude' here still needs each payroll to be proven tokenized (see review())."""
    s = {**DEFAULTS, **(cfg.get("ai_review") or {})}
    s["on"] = s["enabled"] not in (False, "false", "off", "no")
    key = bool(os.environ.get("ANTHROPIC_API_KEY"))
    on_aws = bool(os.environ.get("SOURCE_REMOTE"))
    try:
        import anthropic  # noqa: F401
        have_lib = True
    except ImportError:
        have_lib = False
    mode = str(s["mode"]).lower()
    if mode == "claude":
        s["mode"] = "claude" if key and have_lib else "local"
        s["why"] = "set to claude" + ("" if key and have_lib else
                                      " but " + ("no ANTHROPIC_API_KEY" if not key else "the anthropic package isn't installed")
                                      + ": local")
    elif mode == "local":
        s["mode"], s["why"] = "local", "set to local"
    else:
        s["mode"] = "claude" if key and on_aws and have_lib else "local"
        s["why"] = ("auto: AWS pipeline with an API key" if s["mode"] == "claude" else
                    "auto: " + ("not on the AWS pipeline" if not on_aws else
                                "no ANTHROPIC_API_KEY" if not key else "the anthropic package isn't installed"))
    return s


def tokenized(names) -> bool:
    """True when there are employee values and every one is a token: the file came through the tokenizer."""
    vals = [str(n).strip() for n in names if str(n or "").strip()]
    return bool(vals) and all(TOKEN_RE.match(v) for v in vals)


# ------------------------------------------------------------------------------------------ run-wide state
class _Run:
    """One per agent run (all projects and payroll files): the Claude client, calls made, the time budget, a
    failure that stops further calls, and the knowledge index (built once)."""

    def __init__(self, key):
        self.key, self.calls, self.spent, self.stopped, self.client = key, 0, 0.0, None, None
        self.index = None          # (settings key, pipeline, decisions, notes)
        self.exp_done, self.exp_spent = 0, 0.0      # --ai-experiment: patterns compared and seconds, whole run


_RUN = None


def _run_state():
    global _RUN
    from .util import run_timestamp
    key = run_timestamp()
    if _RUN is None or _RUN.key != key:
        _RUN = _Run(key)
    return _RUN


# ------------------------------------------------------------------------------------------------ the index
def _models(s):
    emb = str(s["embedder"])
    if emb == "voyage" and s["mode"] != "claude":
        emb = "auto"                                   # local mode sends nothing anywhere, Voyage included
    if emb == "auto":
        try:
            import sentence_transformers  # noqa: F401
            emb = "local"
        except ImportError:
            emb = "hashing"
    rr = str(s["reranker"])
    if rr == "auto" or (rr == "claude" and s["mode"] != "claude"):
        rr = "cross-encoder" if emb == "local" else "overlap"
    embedder = R.make_embedder(emb)
    return embedder, R.make_reranker(rr, embedder), f"{emb}/{rr}"


def build_index(cfg, s, getter=None):
    """(pipeline or None, decisions, notes): every knowledge source, indexed once per run."""
    run = _run_state()
    skey = (s["mode"], s["embedder"], s["reranker"])
    if run.index is not None and run.index[0] == skey:
        return run.index[1:]
    from . import knowledge as K
    root = Path(cfg["root"])
    notes, chunks = [], []
    if s["download_regulations"]:
        c, n = K.regulation_chunks(Path(os.path.expanduser(s["cache_dir"])), int(s["refresh_days"]), getter)
        chunks += c
        notes += n
    pp = Path(cfg["projects_path"])
    projects = [p.name for p in pp.iterdir() if p.is_dir()] if pp.is_dir() else []
    kd = Path(s["knowledge_dir"])
    c, n = K.document_chunks(kd if kd.is_absolute() else root / kd, projects)
    chunks += c
    notes += n
    decisions = K.read_decisions(pp) if pp.is_dir() else {}
    # what Claude sees of a staff decision: the pattern and the outcome, never the free-text note (it may name people)
    chunks += K.decision_chunks(decisions, with_notes=s["mode"] != "claude")
    pipe = None
    if chunks:
        embedder, reranker, _ = _models(s)
        pipe = R.RetrievalPipeline(chunks, embedder, reranker)
    run.index = (skey, pipe, decisions, notes)
    return pipe, decisions, notes


# --------------------------------------------------------------------------------------------- the patterns
def _norm(x) -> str:
    return " ".join(str(x or "").split()).lower()


def pattern_key(rule, kind, craft, cls, jurisdiction=""):
    return (rule, _norm(kind) or "state", _norm(craft), _norm(cls),
            _norm(jurisdiction) if rule in JURISDICTION_RULES else "")


def pattern_id(project, rule, kind, craft, cls, jurisdiction="") -> str:
    raw = "|".join((_norm(project),) + pattern_key(rule, kind, craft, cls, jurisdiction))
    return "P-" + hashlib.sha1(raw.encode()).hexdigest()[:12]


def group_patterns(issues, project, flags):
    """[pattern dict] for the orange issues whose rule is in flags, most lines first."""
    groups = defaultdict(list)
    for i in issues:
        if i.severity == "red" or i.rule not in flags:
            continue
        c = i.ctx or {}
        groups[pattern_key(i.rule, c.get("kind"), c.get("craft"), c.get("classification"),
                           c.get("jurisdiction"))].append(i)
    out = []
    for key, lst in groups.items():
        c0 = lst[0].ctx or {}
        weeks = [i.ctx.get("week") for i in lst if isinstance(i.ctx.get("week"), dt.date)]
        contractor = Counter(str(i.ctx.get("contractor") or "") for i in lst).most_common(1)[0][0]
        decs = Counter(str(i.ctx.get("wage_decision") or "") for i in lst if i.ctx.get("wage_decision"))
        out.append({"id": pattern_id(project, key[0], key[1], c0.get("craft"), c0.get("classification"),
                                     c0.get("jurisdiction")),
                    "rule": key[0], "kind": key[1], "craft": str(c0.get("craft") or ""),
                    "classification": str(c0.get("classification") or ""),
                    "jurisdiction": str(c0.get("jurisdiction") or ""),
                    "wage_decision": decs.most_common(1)[0][0] if decs else "", "contractor": contractor,
                    "lines": len({i.row for i in lst}), "issues": lst, "example": lst[0].detail,
                    "week_end": max(weeks) if weeks else dt.date.today()})
    return sorted(out, key=lambda p: -p["lines"])


# The legal question behind each check, in the words the regulations and contracts use: the searches match on it
LEGAL_QUESTIONS = {
    "S01": "Daily overtime: more than eight straight-time hours in a day; whether hours over eight in a day must be "
           "paid at the overtime rate",
    "S05": "Weekly overtime: more than forty straight-time hours in a workweek with no overtime paid; overtime at one "
           "and one-half times the rate for hours over forty",
    "S11": "Classification: the payroll classification differs from the wage determination's classification; which "
           "classification and rate apply to the work performed",
    "S12": "Overtime pay below the published overtime rate for the classification; the required overtime rate "
           "including fringe benefits",
    "S13": "The wage determination or county named for the work has no rate for this classification; which wage "
           "determination and county rates apply",
    "S14": "Which wage determination applies when the contract or payroll names none: the rates in effect on the bid "
           "due date or the award date",
}


def _query(p, project, kinds, titles, names):
    q = LEGAL_QUESTIONS.get(p["rule"], titles.get(p["rule"], p["rule"]))
    desc = R.redact(f"{q}. Payroll example: {p['example']} ({p['lines']} line(s) like this)", names)
    return R.FlagQuery(project_id=project, applicable_sources=set(kinds) | {p["kind"]}, week_end=p["week_end"],
                       classification=f"{p['craft']} / {p['classification']}", contractor=p["contractor"],
                       issue_type=titles.get(p["rule"], p["rule"]).lower().replace(" ", "_"), description=desc)


def _cache_key(p, result, s):
    """Stable while the pattern, the law and documents found, the model and the prompt are the same. Not the
    line counts or amounts (they change weekly), nor staff decisions (they'd make Claude echo staff)."""
    ev = "|".join(sorted(f"{c.chunk.chunk_id}:{hashlib.sha1(c.chunk.text.encode()).hexdigest()[:10]}"
                         for c in result.passages if c.chunk.source_type != "history"))
    version = PROMPT_VERSION + (f"+thinking-{s['thinking_effort']}" if s.get("thinking_effort") else "")
    return hashlib.sha256(f"{version}|{s['model']}|{p['id']}|{ev}".encode()).hexdigest()[:20]


def _sources(result, n=4):
    return "; ".join(c.chunk.title for c in result.passages[:n])


def _without_own_case(result, pid):
    """Drop the pattern's own staff decision from its evidence: the suggestion must stand on its own."""
    own = f"CASE-{pid}"
    result.passages = [c for c in result.passages if c.chunk.chunk_id != own]
    result.no_evidence = not result.passages
    return result


def evidence_gaps(pats) -> str:
    """Per rule, the patterns with no evidence found: 'S01 30 of 80 (38%), S11 19 of 45 (42%)'. A shadow-mode
    measure: a high share points at missing evidence (documents, past decisions), not at the model."""
    out = []
    for r in sorted({p["rule"] for p in pats if p.get("no_evidence") is not None}):
        n = [p for p in pats if p["rule"] == r and p.get("no_evidence") is not None]
        k = sum(1 for p in n if p["no_evidence"])
        out.append(f"{r} {k} of {len(n)} ({100 * k / len(n):.0f}%)")
    return ", ".join(out)


# ----------------------------------------------------------------------------------------- the experiment
class _Experiment:
    """--ai-experiment N (run.py; CVA_AI_EXPERIMENT): for up to N patterns of this run, two fresh calls on the same
    evidence, one without and one with adaptive thinking (at thinking_effort, else "high"),
    compared with each other and with the suggestion in use (usually from an earlier run). Results go to
    .agent/ai_thinking_experiment.json and one line of the administrator's email. Suggestions are not changed."""

    def __init__(self, s, run):
        self.run = run                       # N and the time limit hold for the whole run, not per payroll
        self.n = int(os.environ.get("CVA_AI_EXPERIMENT") or 0)
        self.effort = str(s.get("thinking_effort") or "high")
        self.limit = float(s.get("experiment_seconds") or 1200)
        self.model = s["model"]
        self.rows, self.stopped = [], ""

    def wanted(self):
        return (self.n > 0 and self.run.exp_done < self.n and self.run.exp_spent < self.limit and not self.stopped)

    def add(self, p, result, client):
        out = {"pattern": p["id"], "rule": p["rule"], "in_use": (p.get("suggestion") or {}).get("raise_flag")}
        for name, effort in (("plain", ""), ("thinking", self.effort)):
            try:
                j = R.judge_cell(result, model=self.model, client=client, thinking_effort=effort)
            except Exception as e:                                 # noqa: BLE001
                self.stopped = _what(e, 160)
                log.warning("AI review experiment stopped: %s", self.stopped)
                return
            m = VERDICT_RE.match(j.get("explanation") or "")
            self.run.exp_spent += j.get("seconds") or 0
            out[name] = {"verdict": (m.group(1).upper() == "FLAG") if m else None, "seconds": j.get("seconds"),
                         "usage": j.get("usage"), "citations": len(j.get("citations") or []),
                         "explanation": (j.get("explanation") or "")[:800]}
        self.rows.append(out)
        self.run.exp_done += 1

    def summary(self):
        rows = [r for r in self.rows if "thinking" in r]
        if not rows:
            return None, None
        def avg(k, f):
            v = [f(r[k]) for r in rows if f(r[k]) is not None]
            return sum(v) / len(v) if v else 0
        decided = [r for r in rows if r["plain"]["verdict"] is not None and r["thinking"]["verdict"] is not None]
        changed = [r for r in decided if r["plain"]["verdict"] != r["thinking"]["verdict"]]
        to_flag = sum(1 for r in changed if r["thinking"]["verdict"])
        same_as_before = [r for r in rows if r["in_use"] is not None and r["plain"]["verdict"] is not None]
        stable = sum(1 for r in same_as_before if r["in_use"] == r["plain"]["verdict"])
        stats = {"patterns": len(rows), "decided_both": len(decided), "changed": len(changed),
                 "changed_to_flag": to_flag, "changed_to_no_flag": len(changed) - to_flag,
                 "unclear_plain": sum(1 for r in rows if r["plain"]["verdict"] is None),
                 "unclear_thinking": sum(1 for r in rows if r["thinking"]["verdict"] is None),
                 "plain_repeats_suggestion_in_use": [stable, len(same_as_before)],
                 "citations_plain": sum(1 for r in rows if r["plain"]["citations"]),
                 "citations_thinking": sum(1 for r in rows if r["thinking"]["citations"]),
                 "seconds_plain": round(avg("plain", lambda x: x["seconds"]), 1),
                 "seconds_thinking": round(avg("thinking", lambda x: x["seconds"]), 1),
                 "output_tokens_plain": round(avg("plain", lambda x: (x["usage"] or {}).get("output"))),
                 "output_tokens_thinking": round(avg("thinking", lambda x: (x["usage"] or {}).get("output"))),
                 "input_tokens": round(avg("plain", lambda x: (x["usage"] or {}).get("input"))),
                 "thinking_tokens": round(avg("thinking", lambda x: (x["usage"] or {}).get("thinking"))),
                 # a check that the calls without thinking really didn't think
                 "thinking_tokens_plain": round(avg("plain", lambda x: (x["usage"] or {}).get("thinking"))),
                 "thinking_effort": self.effort, "model": self.model, "stopped": self.stopped}
        line = (f"payroll:   ai review experiment (adaptive thinking, effort {self.effort}): {len(rows)} pattern(s); "
                f"verdict changed on {len(changed)} of {len(decided)} ({to_flag} to FLAG, {len(changed) - to_flag} "
                f"to NO FLAG); without thinking, the same verdict as the suggestion in use on {stable} of "
                f"{len(same_as_before)}; citations kept on {stats['citations_thinking']} of {len(rows)} "
                f"(without: {stats['citations_plain']}); {stats['seconds_plain']} s -> {stats['seconds_thinking']} s "
                f"and {stats['output_tokens_plain']} -> {stats['output_tokens_thinking']} output tokens per call "
                f"({stats['thinking_tokens_plain']} -> {stats['thinking_tokens']} of them thinking)"
                + (f"; stopped early: {self.stopped}" if self.stopped else ""))
        return line, stats


# ----------------------------------------------------------------------------------------------- the review
def review(project_dir: Path, cfg, issues, *, kinds=(), titles=None, names=(), client=None, getter=None):
    """Runs the AI review for one reviewed payroll. Returns a dict:
    patterns   [pattern + 'suggestion' (or None) + 'decision' (staff, if any) + 'sources']
    lines      result lines for the email ('payroll:   ai review: ...', 'WARNING ai review: ...')
    audit      notes for the reviewed payroll's Audit sheet
    shadow     True when suggestions are not to be shown
    Never raises: a problem becomes a warning line and the review goes on without suggestions."""
    s = settings(cfg)
    out = {"patterns": [], "lines": [], "audit": [], "shadow": bool(s["shadow"]), "mode": s["mode"]}
    if not s["on"]:
        return out
    project = Path(project_dir).name
    pats = group_patterns(issues, project, set(s["flags"]))
    if not pats:
        return out
    if s["mode"] == "claude" and not s["allow_untokenized"] and not tokenized(names):
        s["mode"], s["why"] = "local", "the payroll isn't tokenized (employee values aren't tokens): nothing sent"
        out["mode"] = "local"
    run = _run_state()
    titles = titles or {}
    need_retrieval = s["mode"] == "claude" or not s["shadow"]      # local + shadow: nothing would be shown
    pipe, decisions, notes = None, {}, []
    try:
        if need_retrieval:
            pipe, decisions, notes = build_index(cfg, s, getter)
        else:
            from . import knowledge as K
            pp = Path(cfg["projects_path"])
            decisions = K.read_decisions(pp) if pp.is_dir() else {}
    except Exception as e:                                         # noqa: BLE001
        log.warning("AI review: the knowledge base could not be built: %s", err_text(e), exc_info=True)
        out["lines"].append(f"WARNING ai review: no evidence this run, the knowledge base could not be built "
                            f"({err_text(e)[:160]})")
    for n in notes:
        log.info("AI review knowledge: %s", n)
    if s["mode"] == "claude":           # a missing regulation matters when Claude is asked to rely on the evidence
        out["lines"] += [f"WARNING ai review: {n}" for n in notes if "not downloaded" in n and "was used" not in n]
    st_file = Path(project_dir) / ".agent" / STATE_FILE
    try:
        state = json.loads(st_file.read_text()) if st_file.is_file() else {}
    except (ValueError, OSError):
        state = {}
    sugg = state.setdefault("suggestions", {})
    calls = cached = no_ev = skipped = unclear = 0
    tokens_in = tokens_out = 0
    spent0 = run.spent
    if s["mode"] == "claude" and client is None and run.client is None and not run.stopped:
        try:
            import anthropic
            run.client = anthropic.Anthropic(max_retries=1, timeout=30)
        except Exception as e:                                     # noqa: BLE001
            run.stopped = f"the Claude client could not be created ({err_text(e)[:120]})"
    client = client or run.client
    today = dt.date.today()
    exp = _Experiment(s, run) if s["mode"] == "claude" else None
    for p in pats:
        p["decision"] = decisions.get(p["id"])
        p["suggestion"], p["sources"], p["no_evidence"] = None, "", None     # None: not searched
        if pipe is None:
            continue
        result = _without_own_case(pipe.retrieve(_query(p, project, kinds, titles, names)), p["id"])
        p["sources"] = _sources(result)
        p["no_evidence"] = bool(result.no_evidence)
        if result.no_evidence:
            no_ev += 1
            continue
        if s["mode"] != "claude":
            continue
        ck = _cache_key(p, result, s)
        prev = sugg.get(p["id"])
        if prev and prev.get("cache_key") == ck and \
                (today - dt.date.fromisoformat(prev.get("at", "2000-01-01"))).days <= int(s["cache_days"]):
            p["suggestion"] = prev
            cached += 1
            if exp and exp.wanted():
                exp.add(p, result, client)
            continue
        if run.stopped or run.calls >= int(s["max_calls_per_run"]) or run.spent >= float(s["max_seconds"]):
            skipped += 1
            continue
        t0 = time.monotonic()
        try:
            j = R.judge_cell(result, model=s["model"], client=client, thinking_effort=str(s["thinking_effort"] or ""))
        except Exception as e:                                     # noqa: BLE001 - outage, quota, bad key
            run.stopped = f"a Claude call failed ({_what(e, 120)})"
            log.warning("AI review: %s; no more calls this run", run.stopped)
            skipped += 1
            continue
        finally:
            run.spent += time.monotonic() - t0
        run.calls += 1
        calls += 1
        usage = j.get("usage") if isinstance(j.get("usage"), dict) else {}
        tokens_in += usage.get("input") if isinstance(usage.get("input"), int) else 0
        tokens_out += usage.get("output") if isinstance(usage.get("output"), int) else 0
        m = VERDICT_RE.match(j.get("explanation") or "")
        if not m:                                                  # no clear verdict: no suggestion, not kept
            unclear += 1
            continue
        verdict = m.group(1).upper() == "FLAG"
        first = (prev or {}).get("first")
        rec = {"raise_flag": verdict, "explanation": j["explanation"][:2000],
               "citations": [c.get("title") for c in j.get("citations") or [] if c.get("title")][:8],
               "cache_key": ck, "model": s["model"], "at": today.isoformat(),
               "rule": p["rule"], "craft_class": f"{p['craft']} / {p['classification']}",
               "evidence": [c.chunk.chunk_id for c in result.passages],
               # the first suggestion made before staff decided: what the agreement figure compares
               "first": first if first is not None else (None if p["decision"] else verdict)}
        sugg[p["id"]] = rec
        p["suggestion"] = rec
        if exp and exp.wanted():
            exp.add(p, result, client)
    if exp and exp.rows:
        line, stats = exp.summary()
        if line:
            out["lines"].append(line)
            log.info("AI review experiment", extra={"details": stats})
            try:
                ef = Path(project_dir) / ".agent" / EXPERIMENT_FILE
                runs = json.loads(ef.read_text()) if ef.is_file() else []
                runs.append({"at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                             "summary": stats, "patterns": exp.rows})
                ef.parent.mkdir(parents=True, exist_ok=True)
                ef.write_text(json.dumps(runs, indent=1, default=str))
            except (OSError, ValueError) as e:
                log.warning("AI review experiment not saved: %s", err_text(e))
    if s["mode"] == "claude":
        try:
            st_file.parent.mkdir(parents=True, exist_ok=True)
            tmp = st_file.with_suffix(".tmp")
            tmp.write_text(json.dumps(state, indent=1, default=str))
            tmp.replace(st_file)
        except OSError as e:
            log.warning("AI review state not saved: %s", err_text(e))
    # agreement: this project's staff decisions against the suggestion made before them
    both = []
    for pid, d in decisions.items():
        sg = sugg.get(pid)
        if d["project"] == project and sg and sg.get("first") is not None:
            both.append((sg["first"], d))
    agree = sum(1 for first, d in both if first == (d["decision"] == "keep"))
    out["patterns"] = pats
    judged = sum(1 for p in pats if p["suggestion"])
    rules = ", ".join(sorted(set(p["rule"] for p in pats)))
    if s["mode"] == "claude":
        line = (f"payroll:   ai review ({'shadow: not shown to staff' if s['shadow'] else 'shown in Review decisions'}): "
                f"{len(pats)} pattern(s) of {rules}; {judged} with a suggestion "
                f"({calls} new Claude call(s), {cached} reused)")
        if no_ev:
            line += f"; {no_ev} with no evidence found ({evidence_gaps(pats)})"
        if unclear:
            line += f"; {unclear} answer(s) without a clear verdict (not used)"
        mind = int(s["min_decisions_for_agreement"])
        if both:
            line += (f"; staff decisions agree with it on {agree} of {len(both)}" if len(both) >= mind else
                     f"; agreement with staff shown from {mind} decisions ({len(both)} so far)")
        out["lines"].append(line)
        if skipped:
            why = run.stopped or (f"the limit of {s['max_calls_per_run']} calls per run (ai_review.max_calls_per_run)"
                                  if run.calls >= int(s["max_calls_per_run"]) else
                                  f"the {s['max_seconds']} s of calls allowed per run (ai_review.max_seconds)")
            out["lines"].append(f"WARNING ai review: {skipped} pattern(s) got no suggestion this run: {why}")
    else:
        line = f"payroll:   ai review (local; {s['why']}): {len(pats)} pattern(s) of {rules} listed for staff decisions"
        if need_retrieval:
            line += f", related sources found for {(len(pats) - no_ev) if pipe is not None else 0}"
        out["lines"].append(line)
    _event(s, run, pats, calls, cached, skipped, no_ev, unclear, tokens_in, tokens_out, run.spent - spent0)
    out["audit"].append(("AI review", f"{len(pats)} pattern(s) of orange flags ({rules}) listed in the Review "
                         f"decisions sheet for staff decisions (Keep / Dismiss). Mode: {s['mode']} ({s['why']})."
                         + (f" Knowledge: {'; '.join(notes[:3])}" if notes else "")))
    return out


def _event(s, run, pats, calls, cached, skipped, no_ev, unclear, tokens_in, tokens_out, seconds):
    """The structured run log's ai_review line (docs/run_log.md): counts, the model and tokens, never text."""
    from . import events
    claude = s["mode"] == "claude"
    stopped = None
    if skipped:
        why = str(run.stopped or "")
        stopped = ("client_failed" if why.startswith("the Claude client") else "call_failed" if why else
                   "call_limit" if run.calls >= int(s["max_calls_per_run"]) else "time_limit")
    events.emit("ai_review", **{"ai.mode": "claude" if claude else "local", "ai.patterns": len(pats),
                                "ai.calls": calls, "ai.cached": cached, "ai.skipped": skipped,
                                "ai.no_evidence": no_ev, "ai.unclear": unclear, "ai.stopped": stopped,
                                "gen_ai.request.model": str(s["model"]) if claude else None,
                                "gen_ai.usage.input_tokens": tokens_in if claude else None,
                                "gen_ai.usage.output_tokens": tokens_out if claude else None,
                                "duration_ms": int(max(seconds, 0) * 1000)})


# --------------------------------------------------------------------------------------- reviewed payroll
def annotate_issues(res):
    """Visible mode: add the suggestion to each flagged cell's note."""
    if res["shadow"]:
        return
    for p in res["patterns"]:
        sg = p.get("suggestion")
        if not sg:
            continue
        verdict = "keep the flag" if sg["raise_flag"] else "likely fine"
        for i in p["issues"]:
            i.detail += f"\nAI suggestion: {verdict} (see the Review decisions sheet)"


def _clean(v):
    if not isinstance(v, str):
        return v
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    return ILLEGAL_CHARACTERS_RE.sub("", v)


def decisions_sheet(wb, res):
    """The "Review decisions" sheet: one row per pattern, with staff's Decision and Note columns (carried over
    from earlier reviews of the same pattern) and, when not in shadow mode, the AI suggestion and its sources."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from .knowledge import DECISION_HEADERS, DECISIONS_SHEET
    pats = res.get("patterns") or []
    if not pats:
        return
    if DECISIONS_SHEET in wb.sheetnames:
        del wb[DECISIONS_SHEET]
    ws = wb.create_sheet(DECISIONS_SHEET, 1 if len(wb.sheetnames) > 1 else 0)
    ws["A1"] = "Orange-flag patterns: decide each one (Keep = a real problem, Dismiss = fine) and add a note."
    ws["A1"].font = Font(bold=True)
    ws["A2"] = "Your decisions are evidence for the review of later payrolls, in every project."
    ws.append([])
    ws.append(DECISION_HEADERS)
    for c in ws[4]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDE7E2")
    show = not res["shadow"] and res["mode"] == "claude"
    for n, p in enumerate(pats, 1):
        d, sg = p.get("decision") or {}, p.get("suggestion")
        ai = ("Keep the flag" if sg["raise_flag"] else "Likely fine") if (show and sg) else ""
        why = (sg["explanation"] + ("\nSources: " + "; ".join(sg["citations"]) if sg.get("citations") else "")) \
            if (show and sg) else (f"Related sources: {p['sources']}" if not res["shadow"] and p.get("sources") else "")
        if p.get("no_evidence"):
            why = NO_EVIDENCE_TEXT          # shown in shadow mode too: an evidence gap, not a suggestion
        ws.append([_clean(x) for x in (
            n, p["rule"], p["kind"], f"{p['craft']} / {p['classification']}",
            p["jurisdiction"] if p["rule"] in JURISDICTION_RULES else p["wage_decision"], p["contractor"], p["lines"],
            p["example"][:500], (d.get("decision") or "").title() or None, d.get("note") or None, ai or None,
            why[:2000] or None, p["id"])])
    dv = DataValidation(type="list", formula1='"Keep,Dismiss"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"I5:I{4 + len(pats)}")
    for col, w in zip("ABCDEFGHIJKLM", (4, 8, 9, 40, 22, 22, 7, 60, 14, 40, 14, 60, 16)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=5):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A5"
