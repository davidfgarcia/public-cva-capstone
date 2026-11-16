"""
demo.py
=======

A runnable walk-through of compliance_rag.py on a small, INVENTED data set.

    python demo.py                                  offline stand-ins (no installs, no keys)
    python demo.py --embedder local --reranker cross-encoder
                                                    real open-source models on this machine
    python demo.py --embedder voyage --reranker claude
                                                    hosted models (VOYAGE_API_KEY, ANTHROPIC_API_KEY)
    python demo.py --no-judge                       skip the Claude judgment even if a key is set

IMPORTANT: every regulation, contract clause, contractor and case below is
sample text written for this demonstration. None of it is the wording of a
real statute, agreement or decision, and none of it should be relied on.

The demo shows, for one flagged payroll cell:
  1. which sources Stage 1 removes, and why;
  2. what relevance-only ranking would send to the model (lots of repeats);
  3. what the full pipeline sends (maximum relevance, minimum redundancy);
  4. evaluation metrics for both, plus a stage-by-stage ablation;
  5. the Claude "search_result" blocks and the retrieval manifest;
  6. optionally, a real Claude judgment, if ANTHROPIC_API_KEY is set.
"""

import argparse
from dataclasses import replace
import json
import os
import sys
from datetime import date, timedelta

import os as _os, sys as _sys
# the library now lives with the agent that uses it: code/compliance_agent/agent/compliance_rag.py
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "compliance_agent", "agent"))
from compliance_rag import (
    DEFAULT_CONFIG, EMBEDDERS, RERANKERS, FlagQuery, LabeledQuery, RetrievalPipeline, case_to_chunk,
    chunk_agreement, chunk_regulation, evaluate, make_embedder, make_reranker,
    passes_filter, to_search_result_blocks,
)


# -----------------------------------------------------------------------------
# 1. SAMPLE KNOWLEDGE SOURCES (invented text, see the warning above)
# -----------------------------------------------------------------------------

FEDERAL_RULES_CURRENT = """
Sample federal labor standards (illustrative text only)

§ 1.1 Overtime pay. Hours worked over 40 in a workweek are paid at one and one-half
times the basic hourly rate. When a worker performs work in more than one
classification in the same week, overtime is paid at the rate of the classification
in which the overtime hours were actually worked. If payroll records do not show
which classification the overtime hours belong to, overtime is paid at the higher rate.

§ 1.2 Payroll records. Certified payrolls must show, for each worker, the
classification and the hours worked in each classification on each day.

§ 1.3 Apprentices. Apprentices may be paid less than the journeyworker rate only if
registered in an approved program and only within the approved ratio of apprentices
to journeyworkers on the job site.

§ 1.4 Fringe benefits. Fringe benefits may be paid into an approved plan or as cash
in addition to the basic hourly rate.
"""

# An OLDER version of the overtime rule, no longer in force. Stage 1 must
# exclude it: it says the opposite of the current rule.
FEDERAL_RULES_SUPERSEDED = """
Sample federal labor standards, prior edition (illustrative text only)

§ 1.1 Overtime pay. When a worker performs work in more than one classification in
the same week, overtime may be paid at the lower of the applicable rates.
"""

# A STATE rule. The demo project is federal + union only, so Stage 1 must
# exclude it even though it is about overtime too.
STATE_RULES = """
Sample state prevailing wage rules (illustrative text only)

§ 2.1 Overtime on state public works. Overtime for split-classification work on
state projects is paid at the rate for the classification worked.
"""

FEDERAL_GUIDANCE = """
Sample determination notes (illustrative text only)

§ 9.1 Footnote to Laborer Group 3. Laborer Group 3 covers general site laborers.
Workers who also operate equipment must be paid the operator rate for those hours.
"""

UNION_AGREEMENT_P07 = """
Sample collective bargaining agreement, Project P07 (illustrative text only)

ARTICLE 7 - OVERTIME. Work over eight hours in a day or forty hours in a week is paid
at time and one-half. Sunday work is paid at double time.

ARTICLE 9 - SPLIT CLASSIFICATIONS. A worker assigned to two classifications on the
same day is paid the higher of the two rates for all hours worked that day,
including overtime.
"""

# An agreement for a DIFFERENT project. Contract terms bind only their own
# project, so Stage 1 must exclude it.
UNION_AGREEMENT_P02 = """
Sample collective bargaining agreement, Project P02 (illustrative text only)

ARTICLE 9 - SPLIT CLASSIFICATIONS. Split-classification overtime is paid at the lower
rate when agreed in writing in advance.
"""


def build_chunks():
    """Chunk every sample source and every sample past case."""
    chunks = []
    chunks += chunk_regulation(FEDERAL_RULES_CURRENT, prefix="FED", title="Sample federal standards",
                               source_type="regulation", wage_source="federal",
                               valid_from=date(2025, 1, 1))
    chunks += chunk_regulation(FEDERAL_RULES_SUPERSEDED, prefix="FED-OLD",
                               title="Sample federal standards (prior edition)",
                               source_type="regulation", wage_source="federal",
                               valid_to=date(2024, 12, 31))
    chunks += chunk_regulation(STATE_RULES, prefix="STATE", title="Sample state rules",
                               source_type="regulation", wage_source="state")
    chunks += chunk_regulation(FEDERAL_GUIDANCE, prefix="FED-NOTE", title="Sample determination notes",
                               source_type="guidance", wage_source="federal")
    chunks += chunk_agreement(UNION_AGREEMENT_P07, prefix="CBA-P07", title="Sample CBA, Project P07",
                              project_id="P07", valid_from=date(2025, 7, 1))
    chunks += chunk_agreement(UNION_AGREEMENT_P02, prefix="CBA-P02", title="Sample CBA, Project P02",
                              project_id="P02", valid_from=date(2025, 7, 1))
    chunks += [case_to_chunk(c, person_names=c.get("employees", ())) for c in sample_cases()]
    return chunks


def sample_cases():
    """
    Invented past decisions. The heart of the redundancy problem: Contractor A
    triggered the same split-overtime pattern almost every week for months.
    """
    cases = []
    start = date(2026, 3, 7)
    # 12 near-identical weekly cases: same contractor, trade and issue.
    for i in range(12):
        corrected = i in (2, 6, 10)
        cases.append(dict(
            case_id=f"A{i + 1:02d}", project_id="P07", wage_source="federal",
            contractor="Contractor A", classification="Laborer Group 3",
            issue_type="overtime_split", decided_on=start + timedelta(weeks=i * 2),
            outcome="corrected" if corrected else "dismissed",
            summary=("Contractor A paid overtime for a Laborer Group 3 worker who also "
                     "operated equipment, at the lower laborer rate."),
            note=("Overtime hours were not recorded by classification; back pay at the higher rate ordered."
                  if corrected else
                  "Payroll showed the overtime hours were worked as a laborer; rate correct."),
            # Employee names appear in the raw note and must be redacted.
            employees=["Jordan Smith"],
        ))
    # The same pattern recorded with a slightly different label, so the exact
    # key does not match. Near-identical text should still collapse it.
    cases.append(dict(
        case_id="A13", project_id="P07", wage_source="federal",
        contractor="Contractor A", classification="Laborer Grp 3",
        issue_type="overtime_split", decided_on=date(2026, 9, 26), outcome="dismissed",
        summary=("Contractor A paid overtime for a Laborer Group 3 worker who also "
                 "operated equipment, at the lower laborer rate."),
        note="Payroll showed the overtime hours were worked as a laborer; rate correct.",
    ))
    # A DIFFERENT contractor with the same issue: distinct evidence, keep it.
    cases.append(dict(
        case_id="B01", project_id="P03", wage_source="federal",
        contractor="Contractor B", classification="Laborer Group 3",
        issue_type="overtime_split", decided_on=date(2026, 5, 2), outcome="corrected",
        summary=("Contractor B paid all overtime at the laborer rate for a worker split "
                 "between laborer and equipment operator, with no daily breakdown."),
        note="No classification breakdown on the payroll; overtime repaid at the operator rate.",
    ))
    # Same contractor, DIFFERENT issue: not a duplicate of the A-pattern.
    cases.append(dict(
        case_id="A20", project_id="P07", wage_source="federal",
        contractor="Contractor A", classification="Laborer Group 3",
        issue_type="apprentice_ratio", decided_on=date(2026, 8, 15), outcome="corrected",
        summary="Contractor A exceeded the apprentice-to-journeyworker ratio on site.",
        note="Two apprentices repaid at the journeyworker rate.",
    ))
    # A decision made AFTER the week under review: cannot be used as precedent.
    cases.append(dict(
        case_id="A99", project_id="P07", wage_source="federal",
        contractor="Contractor A", classification="Laborer Group 3",
        issue_type="overtime_split", decided_on=date(2026, 10, 17), outcome="corrected",
        summary=("Contractor A paid overtime for a Laborer Group 3 worker who also "
                 "operated equipment, at the lower laborer rate."),
        note="Back pay ordered.",
    ))
    return cases


# -----------------------------------------------------------------------------
# 2. THE FLAGGED CELL
# -----------------------------------------------------------------------------

QUERY = FlagQuery(
    project_id="P07",
    applicable_sources={"federal", "union"},
    week_end=date(2026, 10, 10),
    classification="Laborer Group 3",
    contractor="Contractor A",
    issue_type="overtime_split",
    description=("Worker logged 14 overtime hours while splitting the week between laborer "
                 "and equipment operator work; all overtime was paid at the lower laborer rate"),
)


# -----------------------------------------------------------------------------
# 3. HELPERS FOR PRINTING
# -----------------------------------------------------------------------------

def why_removed(chunk, q):
    """Explain a Stage 1 removal in plain words."""
    if chunk.wage_source not in q.applicable_sources:
        return f"wage source '{chunk.wage_source}' does not apply to project {q.project_id}"
    if not chunk.is_valid_on(q.week_end):
        if chunk.valid_to and q.week_end > chunk.valid_to:
            return f"superseded (in force until {chunk.valid_to})"
        return f"not yet in force on {q.week_end} (from {chunk.valid_from})"
    if chunk.source_type == "contract":
        return f"contract for project {chunk.project_id}, not {q.project_id}"
    return "?"


def show(title, result):
    print(f"\n{title}")
    print("-" * len(title))
    if result.no_evidence:
        print("  No passage cleared the relevance floor: no orange flag.")
        return
    for i, c in enumerate(result.passages, 1):
        extra = f"   [{c.cluster_summary}]" if c.cluster_summary else ""
        print(f"  {i}. {c.chunk.chunk_id:<14} {c.chunk.source_type:<10} "
              f"score={c.score:.3f}{extra}")


# -----------------------------------------------------------------------------
# 4. RUN IT
# -----------------------------------------------------------------------------

def parse_args():
    """Command-line switches that choose the real models, with no code edits."""
    ap = argparse.ArgumentParser(description="Retrieval and ranking demo on invented sample data.")
    ap.add_argument("--embedder", choices=EMBEDDERS, default="hashing",
                    help="hashing = offline stand-in (default); local = open-source model on this "
                         "machine; voyage = Voyage AI (needs VOYAGE_API_KEY)")
    ap.add_argument("--reranker", choices=RERANKERS, default="overlap",
                    help="overlap = offline stand-in (default); cross-encoder = local reranking "
                         "model; claude = Claude scores relevance (needs ANTHROPIC_API_KEY)")
    ap.add_argument("--embedder-model", help="override the embedding model name")
    ap.add_argument("--reranker-model", help="override the reranking model name (for claude, a model id)")
    ap.add_argument("--judge-model", default="claude-sonnet-5-5",
                    help="Claude model for the final judgment (default: claude-sonnet-5-5)")
    ap.add_argument("--relevance-floor", type=float,
                    help="drop passages whose relevance is below this (0 to 1; default 0.3). Different "
                         "rerankers score on different scales, so tune this when you switch rerankers")
    ap.add_argument("--no-judge", action="store_true", help="skip the Claude judgment")
    return ap.parse_args()


def main():
    args = parse_args()
    try:
        embedder = make_embedder(args.embedder, args.embedder_model)
        reranker = make_reranker(args.reranker, embedder, args.reranker_model)
    except RuntimeError as e:
        sys.exit(f"Can't start: {e}")
    floor = DEFAULT_CONFIG.relevance_floor if args.relevance_floor is None else args.relevance_floor
    print(f"Embedder: {args.embedder}   Reranker: {args.reranker}   Relevance floor: {floor}")

    chunks = build_chunks()
    cfg = DEFAULT_CONFIG if args.relevance_floor is None else replace(
        DEFAULT_CONFIG, relevance_floor=args.relevance_floor)
    pipeline = RetrievalPipeline(chunks, embedder=embedder, reranker=reranker, config=cfg)
    print(f"Indexed {len(chunks)} chunks. Index version {pipeline.index_version}.")

    # --- Stage 1: what gets filtered out, and why ---------------------------
    print("\nStage 1 removals")
    print("----------------")
    for c in chunks:
        if not passes_filter(c, QUERY):
            print(f"  {c.chunk_id:<14} {why_removed(c, QUERY)}")

    # --- Relevance only vs the full pipeline ---------------------------------
    relevance_only = pipeline.retrieve(QUERY, stages={"filter", "fusion", "score"})
    full = pipeline.retrieve(QUERY)
    show("Relevance-only ranking (no redundancy handling)", relevance_only)
    show("Full pipeline (maximum relevance, minimum redundancy)", full)

    # --- Evaluation and ablation ----------------------------------------------
    # The reviewer-marked passages that justify a decision on this cell.
    labeled = [
        LabeledQuery(QUERY, relevant_ids={"FED-1-1", "FED-1-2", "CBA-P07-article-9",
                                          "FED-NOTE-9-1", "CASE-B01"}),
    ]
    print("\nEvaluation (one labeled query; a real test set would hold many)")
    print("-----------------------------------------------------------------")
    runs = {
        "Full pipeline": RetrievalPipeline.ALL_STAGES,
        "Without MMR": {"filter", "fusion", "score", "collapse"},
        "Without collapse": {"filter", "fusion", "score", "mmr"},
        "Without Stage 4": {"filter", "fusion", "score"},
        "Without filter": {"fusion", "score", "collapse", "mmr"},
        "Semantic search only": {"filter", "score", "collapse", "mmr"},
    }
    print(f"  {'Run':<22}{'recall@8':>9}{'nDCG@8':>8}{'dupes':>7}{'invalid':>9}")
    for name, stages in runs.items():
        m = evaluate(pipeline, labeled, stages=stages)
        print(f"  {name:<22}{m['recall@k']:>9.2f}{m['ndcg@k']:>8.2f}"
              f"{m['duplicate_rate']:>7.2f}{int(m['validity_errors']):>9}")

    # --- What Claude would receive --------------------------------------------
    blocks = to_search_result_blocks(full)
    print("\nFirst search_result block that would be sent to Claude")
    print("------------------------------------------------------")
    print(json.dumps(blocks[0], indent=2))

    print("\nRetrieval manifest (stored with the weekly run)")
    print("-----------------------------------------------")
    m = full.manifest()
    print(json.dumps({k: m[k] for k in ("project_id", "week_end", "index_version", "no_evidence")}, indent=2))
    print(f"  ... plus {len(m['passages'])} passages and the full config")

    # --- Optional: a real Claude judgment --------------------------------------
    if args.no_judge:
        print("\n(Claude judgment skipped: --no-judge.)")
    elif os.environ.get("ANTHROPIC_API_KEY"):
        from compliance_rag import judge_cell
        try:
            verdict = judge_cell(full, model=args.judge_model)
        except Exception as e:          # e.g. low credit balance, bad key, unknown model
            print(f"\nClaude's judgment failed: {type(e).__name__}: {e}")
            print("Check the key, the model name and the Console credit balance.")
            return
        print("\nClaude's judgment")
        print("-----------------")
        print(verdict["explanation"])
        # Claude may cite two text blocks of the same source (a past case and its
        # pattern summary); list each source once.
        for title in dict.fromkeys(ci["title"] for ci in verdict["citations"]):
            print(f"  cited: {title}")
    if args.reranker == "claude":
        print(f"\nClaude reranking requests this run: {reranker.calls} "
              "(each passage is scored once and reused).")
    else:
        print("\n(No ANTHROPIC_API_KEY in this Terminal window, so Claude was not called and no "
              "credits were used. Set it and `pip install anthropic` to see Claude's judgment.)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # A hosted model (Voyage or Claude) failing mid-run: say what failed in
        # one line instead of a long traceback. Keys are never printed.
        sys.exit(f"Demo stopped: {type(e).__name__}: {e}")
