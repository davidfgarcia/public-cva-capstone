"""
test_compliance_rag.py
======================

Checks that each ranking rule does what the design document says.

    python test_compliance_rag.py        (no test framework needed)
    pytest test_compliance_rag.py        (also works, if pytest is installed)
"""

from dataclasses import replace
from datetime import date, timedelta
from types import SimpleNamespace

import demo
import os as _os, sys as _sys
# the library now lives with the agent that uses it: code/compliance_agent/agent/compliance_rag.py
_sys.path.insert(0, _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "compliance_agent", "agent"))
from compliance_rag import (
    DEFAULT_CONFIG, Candidate, Chunk, ClaudeReranker, RetrievalPipeline,
    age_weight, case_to_chunk, chunk_agreement, chunk_regulation,
    collapse_near_duplicates, mmr_select, passes_filter, reciprocal_rank_fusion,
    redact, to_search_result_blocks,
)


def _chunk(cid, source_type="regulation", **kw):
    """A minimal chunk for tests."""
    base = dict(chunk_id=cid, text=f"text of {cid}", title=cid,
                source_type=source_type, wage_source="federal")
    base.update(kw)
    return Chunk(**base)


# --- Stage 1 -----------------------------------------------------------------

def test_filter_removes_superseded_out_of_scope_and_future():
    q = demo.QUERY
    assert not passes_filter(_chunk("old", valid_to=date(2024, 12, 31)), q)        # superseded
    assert not passes_filter(_chunk("state", wage_source="state"), q)              # wrong source
    assert not passes_filter(_chunk("cba", "contract", wage_source="union",
                                    project_id="P02"), q)                         # other project
    assert not passes_filter(_chunk("later", "history",
                                    valid_from=date(2026, 10, 17)), q)            # decided later
    assert passes_filter(_chunk("ok", valid_from=date(2025, 1, 1)), q)


# --- Stage 2 -----------------------------------------------------------------

def test_rrf_rewards_agreement_between_lists():
    fused = reciprocal_rank_fusion([["a", "b"], ["b", "c"]], k=60)
    # "b" appears in both lists, so it beats "a" (rank 1, but in one list only),
    # which in turn beats "c" (rank 2 in one list only).
    assert fused["b"] > fused["a"] > fused["c"]
    assert abs(fused["a"] - 1 / 61) < 1e-12           # 1 / (60 + rank 1)


# --- Stage 3 -----------------------------------------------------------------

def test_age_weight_halves_every_half_life():
    c = _chunk("h", "history", decided_on=date(2026, 1, 1))
    w = age_weight(c, date(2026, 1, 1) + timedelta(weeks=26), 26)
    assert abs(w - 0.5) < 1e-9                       # exactly one half-life later
    assert age_weight(_chunk("law"), date(2030, 1, 1), 26) == 1.0  # law never decays


def test_no_evidence_when_nothing_clears_the_floor():
    p = RetrievalPipeline(demo.build_chunks())
    res = p.retrieve(demo.QUERY, config=replace(DEFAULT_CONFIG, relevance_floor=0.99))
    assert res.no_evidence and res.passages == []


# --- Stage 4 -----------------------------------------------------------------

def _cand(chunk, score, vec):
    return Candidate(chunk=chunk, score=score, relevance=score), vec


def test_collapse_merges_same_pattern_and_keeps_most_recent():
    key = dict(contractor="X", classification="L3", issue_type="ot")
    a, va = _cand(_chunk("a", "history", decided_on=date(2026, 3, 1), outcome="dismissed", **key), 0.9, [1, 0])
    b, vb = _cand(_chunk("b", "history", decided_on=date(2026, 6, 1), outcome="corrected", **key), 0.5, [0, 1])
    out = collapse_near_duplicates([a, b], {"a": va, "b": vb}, DEFAULT_CONFIG)
    assert len(out) == 1
    survivor = out[0]
    assert survivor.chunk.chunk_id == "b"             # most recent case survives
    assert survivor.score == 0.9                      # with the cluster's best score
    assert "2 similar cases" in survivor.cluster_summary


def test_mmr_quotas():
    cfg = replace(DEFAULT_CONFIG, final_k=4, min_regulations=1, max_history_clusters=1)
    cands, vecs = [], {}
    for i in range(3):                                # three distinct, high-scoring cases
        c, v = _cand(_chunk(f"h{i}", "history"), 0.9, [1 if j == i else 0 for j in range(5)])
        cands.append(c); vecs[c.chunk.chunk_id] = v
    reg, vr = _cand(_chunk("reg"), 0.1, [0, 0, 0, 1, 0])  # weak, but the only regulation
    cands.append(reg); vecs["reg"] = vr
    picked = [c.chunk.chunk_id for c in mmr_select(cands, vecs, cfg)]
    assert "reg" in picked                            # regulation quota met
    assert sum(p.startswith("h") for p in picked) == 1  # history cap respected


def test_mmr_prefers_novel_passage_over_repeat():
    cfg = replace(DEFAULT_CONFIG, final_k=2, min_regulations=0)
    a, va = _cand(_chunk("a"), 0.90, [1, 0])
    a2, va2 = _cand(_chunk("a2"), 0.89, [1, 0])          # near copy of a
    b, vb = _cand(_chunk("b"), 0.60, [0, 1])             # less relevant, but new
    picked = [c.chunk.chunk_id for c in mmr_select([a, a2, b], {"a": va, "a2": va2, "b": vb}, cfg)]
    assert picked == ["a", "b"]


# --- Chunking, privacy, Claude formats ----------------------------------------

def test_chunking_keeps_sections_and_articles_whole():
    regs = chunk_regulation("§ 1.1 First rule.\n§ 1.2 Second rule.", prefix="F", title="T",
                            source_type="regulation", wage_source="federal")
    assert [c.citation for c in regs] == ["§ 1.1", "§ 1.2"]
    arts = chunk_agreement("ARTICLE 1 - A\ntext\nARTICLE 2 - B\ntext", prefix="C",
                           title="CBA", project_id="P1")
    assert len(arts) == 2 and all(c.wage_source == "union" for c in arts)


def test_redaction():
    out = redact("Jordan Smith (123-45-6789, js@x.com, 206-555-0100) was underpaid",
                 ["Jordan Smith"])
    assert "Jordan" not in out and "123-45-6789" not in out
    assert "js@x.com" not in out and "555-0100" not in out
    c = case_to_chunk(dict(case_id="1", project_id="P", wage_source="federal", contractor="X",
                           classification="L", issue_type="ot", outcome="dismissed",
                           decided_on=date(2026, 1, 1), summary="Ana Ruiz overtime", note=""),
                      person_names=["Ana Ruiz"])
    assert "Ana" not in c.text and "X" == c.contractor


def test_search_result_blocks_have_citations_enabled():
    p = RetrievalPipeline(demo.build_chunks())
    blocks = to_search_result_blocks(p.retrieve(demo.QUERY))
    assert blocks and all(b["type"] == "search_result" and b["citations"]["enabled"] for b in blocks)


def test_claude_reranker_skips_thinking_blocks():
    reply = SimpleNamespace(content=[
        SimpleNamespace(type="thinking", thinking="let me weigh these..."),
        SimpleNamespace(type="text", text="```json\n[0.9, 0.2]\n```")])
    fake = SimpleNamespace(messages=SimpleNamespace(create=lambda **kw: reply))
    assert ClaudeReranker(client=fake).score("q", [_chunk("a"), _chunk("b")]) == [0.9, 0.2]


def test_claude_reranker_scores_each_passage_once():
    asked = []
    def create(**kw):
        import re
        n = len(re.findall(r"\n\[\d+\] ", kw["messages"][0]["content"]))
        asked.append(n)
        return SimpleNamespace(content=[SimpleNamespace(type="text", text=str([0.5] * n))])
    rr = ClaudeReranker(client=SimpleNamespace(messages=SimpleNamespace(create=create)))
    a, b, c = _chunk("a"), _chunk("b"), _chunk("c")
    rr.score("q", [a, b]); rr.score("q", [a, b]); rr.score("q", [a, b, c])
    assert rr.calls == 2 and asked == [2, 1]          # second call only for the new passage


def test_claude_reranker_falls_back_safely_on_bad_output():
    fake = SimpleNamespace(messages=SimpleNamespace(
        create=lambda **kw: SimpleNamespace(content=[SimpleNamespace(text="not json")])))
    scores = ClaudeReranker(client=fake).score("q", [_chunk("a"), _chunk("b")])
    assert scores == [0.0, 0.0]          # a bad reply can only remove evidence


def test_full_demo_has_no_repeats_and_no_invalid_passages():
    p = RetrievalPipeline(demo.build_chunks())
    res = p.retrieve(demo.QUERY)
    keys = [c.chunk.pattern_key() for c in res.passages if c.chunk.pattern_key()]
    assert len(keys) == len(set(keys))
    assert all(passes_filter(c.chunk, demo.QUERY) for c in res.passages)


# --- Production components (tested with stand-ins; no downloads or keys) -----

def test_cross_encoder_scores_are_squashed_to_0_1():
    from compliance_rag import CrossEncoderReranker
    fake = SimpleNamespace(predict=lambda pairs: [5.0, -5.0, 0.0][:len(pairs)])
    s = CrossEncoderReranker(encoder=fake).score("q", [_chunk("a"), _chunk("b"), _chunk("c")])
    assert s[0] > 0.99 and s[1] < 0.01 and abs(s[2] - 0.5) < 1e-9


def test_voyage_embedder_batches_and_normalizes():
    from compliance_rag import VoyageEmbedder
    calls = []
    def embed(texts, model):
        calls.append(len(texts))
        return SimpleNamespace(embeddings=[[3.0, 4.0] for _ in texts])
    vecs = VoyageEmbedder(client=SimpleNamespace(embed=embed)).embed(["t"] * 300)
    assert calls == [128, 128, 44]                    # three requests, not one
    assert all(abs(v[0] - 0.6) < 1e-9 for v in vecs)  # unit length


def test_pipeline_runs_with_local_style_components():
    from compliance_rag import CrossEncoderReranker, SentenceTransformerEmbedder, HashingEmbedder
    h = HashingEmbedder()
    fake_st = SimpleNamespace(encode=lambda texts: h.embed(texts))
    emb = SentenceTransformerEmbedder(encoder=fake_st)
    rr = CrossEncoderReranker(encoder=SimpleNamespace(predict=lambda pairs: [2.0] * len(pairs)))
    res = RetrievalPipeline(demo.build_chunks(), embedder=emb, reranker=rr).retrieve(demo.QUERY)
    assert res.passages and all(passes_filter(c.chunk, demo.QUERY) for c in res.passages)


def test_factories_explain_missing_keys():
    import os
    from compliance_rag import make_embedder, make_reranker, HashingEmbedder
    saved = {k: os.environ.pop(k, None) for k in ("VOYAGE_API_KEY", "ANTHROPIC_API_KEY")}
    try:
        for fn in (lambda: make_embedder("voyage"),
                   lambda: make_reranker("claude", HashingEmbedder())):
            try:
                fn(); assert False, "should have raised"
            except RuntimeError as e:
                assert "API_KEY" in str(e)
    finally:
        for k, v in saved.items():
            if v is not None:
                os.environ[k] = v


if __name__ == "__main__":
    tests = [v for k, v in dict(globals()).items() if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"\nAll {len(tests)} tests passed.")
