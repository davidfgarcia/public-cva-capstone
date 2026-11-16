# Compliance RAG: retrieval and ranking for the orange-flag review

Python implementation of the design in *Retrieval and Ranking Design for
AI-Assisted Compliance Review*.

## Files

| File | What it is |
| --- | --- |
| `../compliance_agent/agent/compliance_rag.py` | The library (it moved next to the agent, which uses it for the AI review of orange flags; see `agent/ai_review.py`): chunking, redaction, keyword and semantic search, the four ranking stages, Claude citation blocks, evaluation metrics |
| `demo.py` | A walk-through on invented sample data, comparing relevance-only ranking with the full pipeline |
| `test_compliance_rag.py` | Sixteen checks: one per ranking rule, plus the production components (tested with stand-ins) |

## Run it

Python 3.10 or later, standard library only:

```
python demo.py
python test_compliance_rag.py
```

## The four stages, and where they live

| Stage | Function |
| --- | --- |
| 1. Filter by wage source, validity date and project | `passes_filter`, `stage1_filter` |
| 2. Keyword + semantic search, merged by reciprocal rank fusion | `BM25Index`, `VectorIndex`, `reciprocal_rank_fusion` |
| 3. Relevance x authority x age, with a relevance floor | `stage3_score`, `age_weight` |
| 4. Maximum relevance, minimum redundancy | `collapse_near_duplicates`, `mmr_select` |

All weights and thresholds are in `RankingConfig`.

## Moving from demo to production

The demo runs offline with two stand-ins. Real models are chosen with
command-line switches; no code changes are needed.

| Run | Command | Needs |
| --- | --- | --- |
| Offline stand-ins (default) | `python3 demo.py` | Nothing |
| Local open-source models | `python3 demo.py --embedder local --reranker cross-encoder` | `pip3 install sentence-transformers` (first run downloads about 180 MB of models, then works offline; no data leaves the machine) |
| Claude as reranker | `python3 demo.py --reranker claude` | `pip3 install anthropic`, `ANTHROPIC_API_KEY`, Console credits |
| Voyage embeddings | `python3 demo.py --embedder voyage` | `pip3 install voyageai`, `VOYAGE_API_KEY` (separate Voyage AI account) |

Switches can be combined, for example `--embedder local --reranker claude`.
Other options: `--relevance-floor` sets the relevance cut-off (tune it per reranker), `--embedder-model`, `--reranker-model` and `--judge-model`
override model names, and `--no-judge` skips the final Claude call.

The final Claude judgment (`judge_cell`) runs whenever `ANTHROPIC_API_KEY` is
set, unless `--no-judge` is given. If a Claude call fails (for example a low
credit balance or an unknown model name), the demo says so in one line.

Different rerankers spread their scores differently, so re-check
`relevance_floor` in `RankingConfig` against the evaluation results when you
switch rerankers.

## Known limitation of the stand-ins

The demo's keyword-overlap reranker misses one relevant passage (the payroll
record-keeping rule), giving recall of 0.80 instead of 1.00. A real reranking
model, which understands that the rule explains how split hours must be
recorded, would be expected to recover it. This is also why the evaluation
functions exist: to measure exactly this kind of gap.

## Sample data

All regulations, agreements, contractors and cases in `demo.py` are invented
for illustration. None of it is the wording of a real statute, agreement or
decision.
