"""
compliance_rag.py
=================

Retrieval and ranking logic for the orange-flag review described in
"Retrieval and Ranking Design for AI-Assisted Compliance Review".

WHAT THIS MODULE DOES
---------------------
When the payroll check finds a cell that passes every hard-coded (red) rule
but still looks suspicious, the AI review needs evidence: the regulation,
contract clause or past reviewer decision that bears on that cell. This module
finds that evidence and ranks it in four stages:

    Stage 1  FILTER     Drop sources that do not apply to this project or
                        were not in force on the payroll's week-end date.
    Stage 2  SEARCH     Find candidates two ways (keyword and semantic) and
                        merge the two lists with reciprocal rank fusion (RRF).
    Stage 3  SCORE      Rerank for relevance, then weight by legal authority
                        and (for past decisions only) by age. Drop anything
                        below a relevance floor.
    Stage 4  MRMR       Maximum relevance, minimum redundancy:
                        (a) collapse near-duplicate past cases into one entry
                            with a count, then
                        (b) select the final passages with maximal marginal
                            relevance (MMR), with quotas.

The selected passages are then handed to Claude as "search result" content
blocks with citations enabled, so every orange flag cites its sources.

WHAT IS AND IS NOT INCLUDED
---------------------------
* Everything runs with the Python standard library only, so the module and
  the demo work offline.
* Two pieces have simple built-in stand-ins that should be swapped for real
  models in production (see the classes for how):
    - HashingEmbedder  -> stands in for a real embedding model
                          (e.g. Voyage AI, or a local open-source model).
    - OverlapReranker  -> stands in for a reranking model, or for Claude
                          used as a reranker (ClaudeReranker below).
* Calls to Claude need the `anthropic` package and an ANTHROPIC_API_KEY.
  They are only made if you call them; nothing here calls an API on import.

All weights and thresholds are STARTING VALUES, collected in RankingConfig,
to be tuned against the evaluation set (see the evaluation functions).
"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field, asdict, replace
from datetime import date
from typing import Iterable, Protocol, Sequence


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

@dataclass(frozen=True)
class RankingConfig:
    """
    Every tunable number in the ranking, in one place.

    Keeping them together means an evaluation run can try different values
    (see `evaluate`) without touching the ranking code, and each archived
    review can record exactly which settings produced it.
    """

    # --- Stage 2: candidate search -------------------------------------------
    # How many results each search method contributes before fusion.
    candidates_per_method: int = 50
    # The constant in reciprocal rank fusion. 60 is the value used in the
    # original RRF paper; larger values flatten the gap between ranks.
    rrf_k: int = 60
    # How many fused candidates go on to the (more expensive) reranker.
    rerank_top_n: int = 50

    # --- Stage 3: scoring -----------------------------------------------------
    # Authority weight by source type. Law outranks guidance and contract
    # terms, which outrank a past reviewer's decision.
    authority_weights: tuple = (
        ("regulation", 1.00),
        ("guidance", 0.85),
        ("contract", 0.85),
        ("history", 0.70),
    )
    # Past decisions lose half their weight every this many weeks, so recent
    # practice counts more than old practice. Law and contracts do not decay:
    # whether they apply is decided by their validity dates in Stage 1.
    history_half_life_weeks: float = 26.0
    # Candidates whose relevance (0..1) is below this are dropped outright.
    relevance_floor: float = 0.30

    # --- Stage 4: maximum relevance, minimum redundancy -----------------------
    # Two past cases whose embeddings are at least this similar are treated
    # as the same pattern and collapsed into one entry.
    near_duplicate_similarity: float = 0.95
    # MMR trade-off. 1.0 = rank on relevance alone; lower = penalize
    # redundancy more strongly.
    mmr_lambda: float = 0.70
    # How many passages the AI judgment receives.
    final_k: int = 8
    # Quotas: at least this many statute/regulation passages (if any cleared
    # the floor), and at most this many past-decision clusters, so precedent
    # supports the law instead of replacing it.
    min_regulations: int = 1
    max_history_clusters: int = 3

    def authority(self, source_type: str) -> float:
        """Look up the authority weight for a source type."""
        return dict(self.authority_weights).get(source_type, 0.5)


DEFAULT_CONFIG = RankingConfig()

# The four kinds of source the index holds (see the design document).
SOURCE_TYPES = ("regulation", "guidance", "contract", "history")
# The three wage sources a project can be subject to.
WAGE_SOURCES = ("federal", "state", "union")


# =============================================================================
# 2. DATA MODEL
# =============================================================================

@dataclass
class Chunk:
    """
    One retrievable passage, plus the metadata the pipeline filters and
    ranks on.

    A chunk is always one natural unit of its source: one regulation section,
    one contract article, one determination footnote, or one resolved case.
    Cutting a regulation mid-section can drop the condition that makes it
    apply, so chunking never splits those units (see the chunk_* functions).
    """

    chunk_id: str                 # Stable identifier, e.g. "FED-REG-5-5-a"
    text: str                     # The passage itself (already redacted)
    title: str                    # Human-readable name shown in citations
    source_type: str              # "regulation" | "guidance" | "contract" | "history"
    wage_source: str              # "federal" | "state" | "union"

    # Validity window: the dates the rule or rate was in force.
    # None means "open-ended" on that side.
    valid_from: date | None = None
    valid_to: date | None = None

    # Scope, where the source is specific to one project/contractor/trade.
    project_id: str | None = None
    contractor: str | None = None
    classification: str | None = None

    # Past decisions only.
    issue_type: str | None = None     # e.g. "overtime_split"
    outcome: str | None = None        # "corrected" | "dismissed" | "escalated"
    decided_on: date | None = None

    # Optional citation label, e.g. a section number, shown to the reviewer.
    citation: str = ""

    def is_valid_on(self, day: date) -> bool:
        """True if this chunk was in force on `day`."""
        if self.valid_from and day < self.valid_from:
            return False
        if self.valid_to and day > self.valid_to:
            return False
        return True

    def pattern_key(self) -> tuple | None:
        """
        The key that identifies a repeated pattern in past decisions:
        same contractor, same classification, same kind of issue.
        Only past decisions have one.
        """
        if self.source_type != "history":
            return None
        return (self.contractor, self.classification, self.issue_type)


@dataclass
class FlagQuery:
    """
    The suspicious payroll cell, turned into a retrieval query.

    The structured fields drive Stage 1 filtering; `text()` turns the whole
    thing into a sentence for the keyword and semantic searches.
    """

    project_id: str
    applicable_sources: set           # e.g. {"federal", "union"}
    week_end: date                    # The payroll's week-end date
    classification: str
    contractor: str
    issue_type: str
    description: str                  # What looked wrong, in plain words

    def text(self) -> str:
        """The query as one string, for search and reranking."""
        return (
            f"{self.description}. Classification: {self.classification}. "
            f"Contractor: {self.contractor}. Issue: {self.issue_type.replace('_', ' ')}."
        )


@dataclass
class Candidate:
    """
    A chunk travelling through the pipeline, carrying the scores each stage
    adds. Keeping every intermediate score makes the ranking explainable:
    a reviewer (or a grader) can see why each passage was chosen.
    """

    chunk: Chunk
    rrf: float = 0.0                  # Stage 2 fused rank score
    relevance: float = 0.0            # Stage 3 reranker score, 0..1
    authority_weight: float = 1.0     # Stage 3
    age_weight: float = 1.0           # Stage 3
    score: float = 0.0                # Stage 3 final: relevance x authority x age
    mmr: float | None = None          # Stage 4 score at the moment it was picked
    # Stage 4a: when near-duplicate past cases are collapsed, the survivor
    # keeps the others here and gets a one-line summary with counts.
    members: list = field(default_factory=list)
    cluster_summary: str | None = None

    @property
    def cluster_size(self) -> int:
        return 1 + len(self.members)


@dataclass
class RetrievalResult:
    """What the pipeline returns for one flagged cell."""

    query: FlagQuery
    passages: list                    # Final Candidates, in rank order
    no_evidence: bool                 # True if nothing cleared the floor
    index_version: str                # Fingerprint of the index used
    config: RankingConfig
    trace: dict = field(default_factory=dict)   # Per-stage details (optional)

    def manifest(self) -> dict:
        """
        A compact, JSON-safe record of this retrieval, to store alongside the
        weekly run's manifest. It lets anyone later reconstruct exactly which
        evidence an orange flag was based on (the reproducibility safeguard).
        """
        return {
            "project_id": self.query.project_id,
            "week_end": self.query.week_end.isoformat(),
            "index_version": self.index_version,
            "no_evidence": self.no_evidence,
            "config": asdict(self.config),
            "passages": [
                {
                    "chunk_id": c.chunk.chunk_id,
                    "source_type": c.chunk.source_type,
                    "relevance": round(c.relevance, 4),
                    "score": round(c.score, 4),
                    "mmr": None if c.mmr is None else round(c.mmr, 4),
                    "cluster_size": c.cluster_size,
                }
                for c in self.passages
            ],
        }


# =============================================================================
# 3. PRIVACY: REDACTION BEFORE INDEXING
# =============================================================================

# Patterns for personal details that must never enter the index. The history
# keeps the pattern and the decision, not the person.
_SSN = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")
_EMAIL = re.compile(r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b")
_PHONE = re.compile(r"\b(?:\+?1[-. ]?)?\(?\d{3}\)?[-. ]?\d{3}[-. ]?\d{4}\b")


def redact(text: str, person_names: Iterable[str] = ()) -> str:
    """
    Remove personal details from a past case before it is indexed.

    `person_names` should be the employee names from the payroll file the
    case came from (the code cannot reliably spot names on its own).
    Contractor (company) names are kept: they are part of the pattern.
    """
    out = _SSN.sub("[ID]", text)
    out = _EMAIL.sub("[EMAIL]", out)
    out = _PHONE.sub("[PHONE]", out)
    # Longest names first, so "Ana Maria Lopez" is replaced before "Ana".
    for name in sorted({n for n in person_names if n}, key=len, reverse=True):
        out = re.sub(re.escape(name), "[EMPLOYEE]", out, flags=re.IGNORECASE)
    return out


# =============================================================================
# 4. CHUNKING: ONE NATURAL UNIT PER CHUNK
# =============================================================================

# A regulation section header such as "§ 5.5(a)(1)" or "Section 12".
_SECTION_HEADER = re.compile(r"^\s*(§\s*[\d.]+(?:\([a-z0-9]+\))*|Section\s+[\w.]+)", re.MULTILINE)
# A contract article header such as "ARTICLE 7 - OVERTIME".
_ARTICLE_HEADER = re.compile(r"^\s*(ARTICLE\s+[\w.]+)", re.MULTILINE | re.IGNORECASE)


def _split_on_headers(text: str, header: re.Pattern) -> list:
    """
    Split `text` at each header match, keeping the header with its body.
    Returns (header, body) pairs. Text before the first header is dropped,
    because it is usually a table of contents or preamble.
    """
    matches = list(header.finditer(text))
    parts = []
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        parts.append((m.group(1).strip(), text[m.start():end].strip()))
    return parts


def _slug(s: str) -> str:
    """Turn a header like '§ 5.5(a)(1)' into an id-safe '5-5-a-1'."""
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def chunk_regulation(text: str, *, prefix: str, title: str, source_type: str,
                     wage_source: str, valid_from: date | None = None,
                     valid_to: date | None = None) -> list:
    """
    Split a statute, regulation or guidance document into one chunk per
    section, so a chunk never separates a rule from its conditions.
    """
    chunks = []
    for header, body in _split_on_headers(text, _SECTION_HEADER):
        chunks.append(Chunk(
            chunk_id=f"{prefix}-{_slug(header)}",
            text=body,
            title=f"{title}, {header}",
            citation=header,
            source_type=source_type,
            wage_source=wage_source,
            valid_from=valid_from,
            valid_to=valid_to,
        ))
    return chunks


def chunk_agreement(text: str, *, prefix: str, title: str, project_id: str,
                    valid_from: date | None = None,
                    valid_to: date | None = None) -> list:
    """
    Split a collective bargaining agreement or contract into one chunk per
    article. Agreements are project-specific and always "union" sources.
    """
    chunks = []
    for header, body in _split_on_headers(text, _ARTICLE_HEADER):
        chunks.append(Chunk(
            chunk_id=f"{prefix}-{_slug(header)}",
            text=body,
            title=f"{title}, {header}",
            citation=header,
            source_type="contract",
            wage_source="union",
            project_id=project_id,
            valid_from=valid_from,
            valid_to=valid_to,
        ))
    return chunks


def case_to_chunk(case: dict, person_names: Iterable[str] = ()) -> Chunk:
    """
    Turn one resolved orange flag (from the preserved historical records)
    into a chunk. One case = one chunk.

    Expected keys: case_id, project_id, wage_source, contractor,
    classification, issue_type, outcome, decided_on (date), summary, note.
    """
    text = redact(
        f"{case['summary']} Reviewer decision: {case['outcome']}. "
        f"Reviewer note: {case.get('note', '')}",
        person_names,
    )
    return Chunk(
        chunk_id=f"CASE-{case['case_id']}",
        text=text,
        title=f"Past case {case['case_id']} ({case['outcome']}, {case['decided_on']:%b %d, %Y})",
        source_type="history",
        wage_source=case["wage_source"],
        project_id=case["project_id"],
        contractor=case["contractor"],
        classification=case["classification"],
        issue_type=case["issue_type"],
        outcome=case["outcome"],
        decided_on=case["decided_on"],
        # A decision is evidence from the day it was made onward.
        valid_from=case["decided_on"],
    )


# =============================================================================
# 5. TEXT HELPERS
# =============================================================================

# Small stopword list: enough to keep BM25 focused on meaningful terms.
_STOPWORDS = set("""
a an and are as at be by for from has have in is it its of on or that the
this to was were will with which who whom than then there their they not
""".split())

# Tokens are words, or section-style numbers such as "5.5" or "296-127".
_TOKEN = re.compile(r"[a-z0-9]+(?:[.\-][a-z0-9]+)*")


def _stem(token: str) -> str:
    """
    A deliberately tiny stemmer, so "splitting", "split" and "splits" match,
    and "classifications" matches "classification". Numbers and section
    numbers are left alone. A production system would use a proper stemmer
    (e.g. Snowball) or rely on the embedding model for word forms.
    """
    if token[0].isdigit() or len(token) <= 4:
        return token
    for suffix in ("ations", "ation", "ings", "ing", "ies", "ed", "es", "s"):
        if token.endswith(suffix) and len(token) - len(suffix) >= 3:
            stem = token[: -len(suffix)]
            # "splitting" -> "splitt" -> "split": drop a doubled final letter.
            if len(stem) > 3 and stem[-1] == stem[-2]:
                stem = stem[:-1]
            return stem
    return token


def tokenize(text: str) -> list:
    """Lowercase, split into word/number tokens, drop stopwords, stem."""
    return [_stem(t) for t in _TOKEN.findall(text.lower()) if t not in _STOPWORDS]


def cosine(a: Sequence[float], b: Sequence[float]) -> float:
    """Cosine similarity of two vectors (already L2-normalized -> dot product)."""
    return sum(x * y for x, y in zip(a, b))


# =============================================================================
# 6. STAGE 2a: KEYWORD SEARCH (BM25)
# =============================================================================

class BM25Index:
    """
    Classic BM25 keyword search, written out so the scoring is visible.

    Keyword search is the half of Stage 2 that catches EXACT terms:
    classification names ("Laborer Group 3"), section numbers ("5.5"),
    contractor names. Semantic search can blur these; keyword search does not.

    In production this could be a library (rank_bm25) or a search engine
    (OpenSearch); the behaviour is the same.
    """

    def __init__(self, k1: float = 1.5, b: float = 0.75):
        # k1 controls how quickly repeated terms stop adding score;
        # b controls how much long chunks are penalized. Standard defaults.
        self.k1, self.b = k1, b
        self.docs: dict = {}         # chunk_id -> Counter of tokens
        self.lengths: dict = {}      # chunk_id -> token count
        self.df: Counter = Counter() # token -> number of chunks containing it
        self.avg_len = 0.0

    def fit(self, chunks: Iterable[Chunk]) -> "BM25Index":
        for c in chunks:
            toks = tokenize(c.title + " " + c.text)
            self.docs[c.chunk_id] = Counter(toks)
            self.lengths[c.chunk_id] = len(toks)
            self.df.update(set(toks))
        self.avg_len = sum(self.lengths.values()) / max(len(self.lengths), 1)
        return self

    def _idf(self, term: str) -> float:
        # Rare terms are worth more than common ones. The +1 inside the log
        # keeps the value positive even for terms in most chunks.
        n, df = len(self.docs), self.df.get(term, 0)
        return math.log(1 + (n - df + 0.5) / (df + 0.5))

    def search(self, query: str, allowed_ids: set, top_n: int) -> list:
        """Return the top_n chunk ids among `allowed_ids`, best first."""
        q_terms = tokenize(query)
        scores = {}
        for cid in allowed_ids:
            tf, length = self.docs[cid], self.lengths[cid]
            s = 0.0
            for term in q_terms:
                f = tf.get(term, 0)
                if f:
                    norm = f * (self.k1 + 1) / (
                        f + self.k1 * (1 - self.b + self.b * length / self.avg_len))
                    s += self._idf(term) * norm
            if s > 0:
                scores[cid] = s
        return sorted(scores, key=scores.get, reverse=True)[:top_n]


# =============================================================================
# 7. STAGE 2b: SEMANTIC SEARCH (EMBEDDINGS)
# =============================================================================

class Embedder(Protocol):
    """Anything that turns texts into L2-normalized vectors."""

    def embed(self, texts: Sequence[str]) -> list: ...


class HashingEmbedder:
    """
    OFFLINE STAND-IN for a real embedding model.

    It hashes words and word pairs into a fixed-size vector, which captures
    shared vocabulary but not true meaning. It exists so the pipeline and
    demo run with no downloads or API keys.

    Replace it in production with one of:
      * VoyageEmbedder below (Anthropic's docs point to Voyage AI, since
        Anthropic does not offer its own embedding model), or
      * SentenceTransformerEmbedder below, if the firm decides payroll data
        must stay on its own machines (the local-hosting decision).
    """

    def __init__(self, dim: int = 1024):
        self.dim = dim

    def _bucket(self, feature: str) -> int:
        # A stable hash (Python's built-in hash() changes between runs).
        return int(hashlib.md5(feature.encode()).hexdigest(), 16) % self.dim

    def embed(self, texts: Sequence[str]) -> list:
        vectors = []
        for t in texts:
            v = [0.0] * self.dim
            toks = tokenize(t)
            for tok in toks:
                v[self._bucket(tok)] += 1.0
            for a, b in zip(toks, toks[1:]):          # word pairs add context
                v[self._bucket(a + "_" + b)] += 0.5
            norm = math.sqrt(sum(x * x for x in v)) or 1.0
            vectors.append([x / norm for x in v])
        return vectors


class VoyageEmbedder:
    """
    Production embedder using Voyage AI (requires `pip install voyageai`
    and a VOYAGE_API_KEY). Imported lazily so the module works without it.
    """

    # Voyage limits how many texts one request may carry, so long lists are
    # sent in batches of this size.
    BATCH = 128

    def __init__(self, model: str = "voyage-3", client=None):
        if client is None:
            import voyageai                  # noqa: imported only when used
            client = voyageai.Client()       # reads VOYAGE_API_KEY from the environment
        self.client, self.model = client, model

    def embed(self, texts: Sequence[str]) -> list:
        texts = list(texts)
        out = []
        for i in range(0, len(texts), self.BATCH):
            res = self.client.embed(texts[i:i + self.BATCH], model=self.model)
            out.extend(_normalize(v) for v in res.embeddings)
        return out


class SentenceTransformerEmbedder:
    """
    Local embedder (requires `pip install sentence-transformers`). Nothing
    leaves the machine, which satisfies a "run locally" hosting decision.
    The first run downloads the model (about 90 MB) and caches it; later runs
    work offline.
    """

    def __init__(self, model: str = "all-MiniLM-L6-v2", encoder=None):
        if encoder is None:
            from sentence_transformers import SentenceTransformer   # noqa
            encoder = SentenceTransformer(model)
        self.model = encoder

    def embed(self, texts: Sequence[str]) -> list:
        return [_normalize(list(v)) for v in self.model.encode(list(texts))]


def _normalize(v: Sequence[float]) -> list:
    n = math.sqrt(sum(x * x for x in v)) or 1.0
    return [x / n for x in v]


class VectorIndex:
    """
    Minimal in-memory vector store: holds one normalized vector per chunk
    and returns the most similar ones. In production this would be a vector
    database (pgvector, OpenSearch, or a local store such as FAISS or
    LanceDB); the interface stays the same.
    """

    def __init__(self, embedder: Embedder):
        self.embedder = embedder
        # Two vectors per chunk, for two different jobs:
        #   vectors       title + text, used for SEARCH (the title helps match)
        #   text_vectors  text only, used to measure REDUNDANCY in Stage 4.
        # Titles carry case ids and dates ("Past case A12, Aug 15") that differ
        # even when two cases say the same thing, so they must not count when
        # deciding whether two passages repeat each other.
        self.vectors: dict = {}
        self.text_vectors: dict = {}

    def fit(self, chunks: Sequence[Chunk]) -> "VectorIndex":
        vecs = self.embedder.embed([c.title + " " + c.text for c in chunks])
        self.vectors = {c.chunk_id: v for c, v in zip(chunks, vecs)}
        tvecs = self.embedder.embed([c.text for c in chunks])
        self.text_vectors = {c.chunk_id: v for c, v in zip(chunks, tvecs)}
        return self

    def search(self, query: str, allowed_ids: set, top_n: int) -> list:
        q = self.embedder.embed([query])[0]
        scored = {cid: cosine(q, self.vectors[cid]) for cid in allowed_ids}
        return sorted(scored, key=scored.get, reverse=True)[:top_n]


# =============================================================================
# 8. STAGE 1: FILTER BEFORE RANKING
# =============================================================================

def passes_filter(chunk: Chunk, query: FlagQuery) -> bool:
    """
    Hard rules that remove a chunk before any scoring.

    Superseded or out-of-scope text is EXCLUDED, not down-ranked: a low rank
    still leaves it within reach of the model, and an outdated rule presented
    as current is worse than no rule at all.
    """
    # 1. Only wage sources that apply to this project
    #    (e.g. a state-only project never sees union agreements).
    if chunk.wage_source not in query.applicable_sources:
        return False
    # 2. Only text in force on the payroll's week-end date.
    #    For past decisions, valid_from is the decision date, so a decision
    #    made after the week being reviewed cannot be used as precedent.
    if not chunk.is_valid_on(query.week_end):
        return False
    # 3. Contract terms bind only their own project.
    if chunk.source_type == "contract" and chunk.project_id != query.project_id:
        return False
    # Past decisions from OTHER projects are allowed: the same contractor
    # often makes the same mistake on several jobs, and that is useful
    # evidence. The wage-source rule above still applies to them.
    return True


def stage1_filter(chunks: Iterable[Chunk], query: FlagQuery) -> list:
    return [c for c in chunks if passes_filter(c, query)]


# =============================================================================
# 9. STAGE 2c: RECIPROCAL RANK FUSION
# =============================================================================

def reciprocal_rank_fusion(ranked_lists: Sequence[Sequence[str]], k: int = 60) -> dict:
    """
    Merge several best-first lists of ids into one score per id:

        RRF(d) = sum over lists of  1 / (k + rank_in_list(d))

    A chunk that ranks well in EITHER list scores well, and one that ranks
    well in BOTH scores best. Only ranks are used, so the keyword and semantic
    scores (which are on different scales) never have to be compared directly.
    """
    fused = defaultdict(float)
    for ranked in ranked_lists:
        for rank, cid in enumerate(ranked, start=1):
            fused[cid] += 1.0 / (k + rank)
    return dict(fused)


# =============================================================================
# 10. STAGE 3: RERANKING AND WEIGHTING
# =============================================================================

class Reranker(Protocol):
    """Anything that scores each chunk's relevance to the query, 0..1."""

    def score(self, query: str, chunks: Sequence[Chunk]) -> list: ...


class OverlapReranker:
    """
    OFFLINE STAND-IN for a reranking model.

    Blends embedding similarity with the share of query terms the chunk
    contains, then clips to 0..1. Good enough to demonstrate the pipeline;
    replace with a cross-encoder model or with ClaudeReranker in production.
    """

    def __init__(self, embedder: Embedder):
        self.embedder = embedder

    def score(self, query: str, chunks: Sequence[Chunk]) -> list:
        q_vec = self.embedder.embed([query])[0]
        c_vecs = self.embedder.embed([c.title + " " + c.text for c in chunks])
        q_terms = set(tokenize(query))
        out = []
        for c, v in zip(chunks, c_vecs):
            overlap = len(q_terms & set(tokenize(c.title + " " + c.text))) / max(len(q_terms), 1)
            out.append(max(0.0, min(1.0, 0.5 * cosine(q_vec, v) + 0.5 * overlap)))
        return out


class ClaudeReranker:
    """
    Uses Claude to score relevance (requires `pip install anthropic` and an
    ANTHROPIC_API_KEY). One call scores a whole batch of candidates.

    Claude returns a JSON list of numbers; anything unparseable falls back to
    0.0 for that chunk, so a bad response can only remove evidence, never
    invent it.

    Scores are cached per (query, chunk): Claude's scores can vary slightly
    between calls, so scoring each passage once keeps every ranking in a run
    consistent, makes evaluation runs comparable, and avoids paying twice.
    """

    def __init__(self, model: str = "claude-sonnet-5-5", client=None):
        if client is None:
            import anthropic                 # noqa: imported only when used
            client = anthropic.Anthropic()
        self.client, self.model = client, model
        self._cache: dict = {}               # (query, chunk_id) -> score
        self.calls = 0                       # Claude requests made (for cost visibility)

    def score(self, query: str, chunks: Sequence[Chunk]) -> list:
        todo = [c for c in chunks if (query, c.chunk_id) not in self._cache]
        if todo:
            for c, r in zip(todo, self._score_uncached(query, todo)):
                self._cache[(query, c.chunk_id)] = r
        return [self._cache[(query, c.chunk_id)] for c in chunks]

    def _score_uncached(self, query: str, chunks: Sequence[Chunk]) -> list:
        self.calls += 1
        numbered = "\n\n".join(f"[{i}] {c.title}\n{c.text}" for i, c in enumerate(chunks))
        prompt = (
            "You are ranking evidence for a prevailing-wage compliance reviewer.\n"
            f"Flagged payroll cell: {query}\n\n"
            f"Passages:\n{numbered}\n\n"
            "For each passage, give a relevance score from 0 to 1: how directly it "
            "helps decide whether this cell is wrong. Reply with only a JSON list "
            f"of {len(chunks)} numbers, in passage order."
        )
        msg = self.client.messages.create(
            # Generous limit: newer models may "think" before answering, and that
            # thinking counts toward max_tokens. Too low a limit leaves no answer.
            model=self.model, max_tokens=4000,
            messages=[{"role": "user", "content": prompt}],
        )
        return self._parse_scores(_reply_text(msg), len(chunks))

    @staticmethod
    def _parse_scores(text: str, n: int) -> list:
        """Read the JSON list of scores from Claude's reply, tolerating a code
        fence or a sentence around it. Missing or bad values become 0.0."""
        m = re.search(r"\[[^\[\]]*\]", text or "")
        try:
            scores = json.loads(m.group(0)) if m else []
            out = [max(0.0, min(1.0, float(x))) for x in scores][:n]
        except (ValueError, TypeError):
            out = []
        return out + [0.0] * (n - len(out))


def _reply_text(msg) -> str:
    """
    All the text in a Claude reply. A reply is a list of content blocks, and
    the answer is not always the first one: a model may return a "thinking"
    block before its text. Only blocks of type "text" are the answer.
    """
    return "".join(getattr(b, "text", "") for b in getattr(msg, "content", [])
                   if getattr(b, "type", "text") == "text")


class CrossEncoderReranker:
    """
    Local reranking model (requires `pip install sentence-transformers`).

    A cross-encoder reads the query and one passage TOGETHER and scores how
    well they match, which is more accurate than comparing two separate
    embeddings. It runs on the machine, so no data leaves it. The default
    model is small (about 90 MB, downloaded once and cached).

    The model outputs an unbounded score; a sigmoid squashes it into 0..1 so
    it can be compared with the relevance floor like every other reranker.
    Different rerankers spread their scores differently, so re-check
    `relevance_floor` with the evaluation set when you switch.
    """

    def __init__(self, model: str = "cross-encoder/ms-marco-MiniLM-L-6-v2", encoder=None):
        if encoder is None:
            from sentence_transformers import CrossEncoder   # noqa
            encoder = CrossEncoder(model)
        self.model = encoder

    def score(self, query: str, chunks: Sequence[Chunk]) -> list:
        raw = self.model.predict([(query, c.title + " " + c.text) for c in chunks])
        return [1.0 / (1.0 + math.exp(-float(x))) for x in raw]


# -----------------------------------------------------------------------------
# Choosing components by name (used by demo.py's command-line switches)
# -----------------------------------------------------------------------------

EMBEDDERS = ("hashing", "voyage", "local")
RERANKERS = ("overlap", "claude", "cross-encoder")


def make_embedder(name: str = "hashing", model: str | None = None) -> Embedder:
    """
    Build an embedder by name:
      hashing  offline stand-in, no installs or keys (the default)
      voyage   Voyage AI: `pip install voyageai` and VOYAGE_API_KEY
      local    open-source model on this machine: `pip install sentence-transformers`
    Missing packages or keys raise a RuntimeError that says exactly what to do.
    """
    if name == "hashing":
        return HashingEmbedder()
    if name == "voyage":
        if not os.environ.get("VOYAGE_API_KEY"):
            raise RuntimeError("The voyage embedder needs a VOYAGE_API_KEY (create one at voyageai.com).")
        try:
            return VoyageEmbedder(**({"model": model} if model else {}))
        except ImportError:
            raise RuntimeError("Install the Voyage library first: pip3 install voyageai") from None
    if name == "local":
        try:
            return SentenceTransformerEmbedder(**({"model": model} if model else {}))
        except ImportError:
            raise RuntimeError("Install the local models first: pip3 install sentence-transformers") from None
    raise ValueError(f"Unknown embedder '{name}'; choose one of {EMBEDDERS}")


def make_reranker(name: str, embedder: Embedder, model: str | None = None) -> Reranker:
    """
    Build a reranker by name:
      overlap        offline stand-in (the default)
      claude         Claude scores relevance: `pip install anthropic` and ANTHROPIC_API_KEY
      cross-encoder  local reranking model: `pip install sentence-transformers`
    """
    if name == "overlap":
        return OverlapReranker(embedder)
    if name == "claude":
        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise RuntimeError("The claude reranker needs ANTHROPIC_API_KEY set in this Terminal window.")
        try:
            return ClaudeReranker(**({"model": model} if model else {}))
        except ImportError:
            raise RuntimeError("Install the Anthropic library first: pip3 install anthropic") from None
    if name == "cross-encoder":
        try:
            return CrossEncoderReranker(**({"model": model} if model else {}))
        except ImportError:
            raise RuntimeError("Install the local models first: pip3 install sentence-transformers") from None
    raise ValueError(f"Unknown reranker '{name}'; choose one of {RERANKERS}")


def age_weight(chunk: Chunk, as_of: date, half_life_weeks: float) -> float:
    """
    Exponential decay for past decisions only:

        w_age = 0.5 ** (weeks since decision / half-life)

    A decision made one half-life ago counts half as much as one made today.
    Law and contracts return 1.0: their applicability is already settled by
    their validity dates in Stage 1.
    """
    if chunk.source_type != "history" or chunk.decided_on is None:
        return 1.0
    weeks = max((as_of - chunk.decided_on).days, 0) / 7.0
    return 0.5 ** (weeks / half_life_weeks)


def stage3_score(candidates: list, query: FlagQuery, reranker: Reranker,
                 cfg: RankingConfig) -> list:
    """
    Score each candidate:

        s(d) = r(d) x w_auth(d) x w_age(d)

    and drop anything whose relevance r is below the floor. The floor is
    applied to r, not to s, so a highly relevant past case is not lost just
    because history carries a lower authority weight.
    """
    if not candidates:
        return []
    relevances = reranker.score(query.text(), [c.chunk for c in candidates])
    kept = []
    for cand, r in zip(candidates, relevances):
        cand.relevance = r
        cand.authority_weight = cfg.authority(cand.chunk.source_type)
        cand.age_weight = age_weight(cand.chunk, query.week_end, cfg.history_half_life_weeks)
        cand.score = r * cand.authority_weight * cand.age_weight
        if r >= cfg.relevance_floor:
            kept.append(cand)
    return sorted(kept, key=lambda c: c.score, reverse=True)


# =============================================================================
# 11. STAGE 4a: COLLAPSE NEAR-DUPLICATE PAST CASES
# =============================================================================

def _summarize_cluster(cands: Sequence[Candidate]) -> str:
    """
    One line that turns a pile of repeated cases into evidence, e.g.
    "12 similar cases since Mar 2026: 9 dismissed, 3 corrected".
    The count is itself useful to the model, and it costs one slot, not twelve.
    """
    outcomes = Counter(c.chunk.outcome for c in cands)
    first = min(c.chunk.decided_on for c in cands if c.chunk.decided_on)
    parts = ", ".join(f"{n} {o}" for o, n in outcomes.most_common())
    return f"{len(cands)} similar cases since {first:%b %Y}: {parts}"


def collapse_near_duplicates(candidates: list, vectors: dict,
                             cfg: RankingConfig) -> list:
    """
    Merge past cases that describe the same pattern into one entry.

    Two past cases are "the same pattern" if EITHER
      * they share contractor, classification and issue type (exact key), or
      * their embeddings are at least `near_duplicate_similarity` alike
        (catches the same pattern recorded with slightly different labels).

    The survivor of each cluster is its MOST RECENT case (current practice
    matters most), with the best score in the cluster, and a summary line.
    Laws, guidance and contracts are never collapsed: each one is distinct.
    """
    history = [c for c in candidates if c.chunk.source_type == "history"]
    others = [c for c in candidates if c.chunk.source_type != "history"]

    def same_pattern(a: Candidate, b: Candidate) -> bool:
        if a.chunk.pattern_key() == b.chunk.pattern_key():
            return True
        return cosine(vectors[a.chunk.chunk_id],
                      vectors[b.chunk.chunk_id]) >= cfg.near_duplicate_similarity

    # Greedy clustering: walk the cases best-first; each joins the first
    # existing cluster in which it matches ANY member, or starts a new one.
    # Matching any member (not just the first) matters: a case relabelled
    # "Laborer Grp 3" joins by text similarity, and then later cases with the
    # standard label can still join through the other members' exact key.
    clusters: list = []
    for cand in sorted(history, key=lambda c: c.score, reverse=True):
        for cluster in clusters:
            if any(same_pattern(cand, member) for member in cluster):
                cluster.append(cand)
                break
        else:
            clusters.append([cand])

    collapsed = []
    for cluster in clusters:
        survivor = max(cluster, key=lambda c: c.chunk.decided_on or date.min)
        survivor.score = max(c.score for c in cluster)
        survivor.relevance = max(c.relevance for c in cluster)
        survivor.members = [c for c in cluster if c is not survivor]
        if survivor.members:
            survivor.cluster_summary = _summarize_cluster(cluster)
        collapsed.append(survivor)

    return sorted(others + collapsed, key=lambda c: c.score, reverse=True)


# =============================================================================
# 12. STAGE 4b: MMR SELECTION WITH QUOTAS
# =============================================================================

def mmr_select(candidates: list, vectors: dict, cfg: RankingConfig) -> list:
    """
    Maximum relevance, minimum redundancy, via maximal marginal relevance.

    Passages are picked one at a time. At each step, every remaining
    candidate d is scored

        MMR(d) = lambda * s(d)  -  (1 - lambda) * max_{d' already picked} sim(d, d')

    so it is rewarded for its own score and penalized for resembling what is
    already selected. The best MMR is picked, and the loop repeats until
    `final_k` passages are chosen.

    Two quotas shape the selection:
      * min_regulations: if any statute/regulation cleared the floor, at
        least that many are included (the law must be in front of the model);
      * max_history_clusters: at most that many past-decision entries, so
        precedent supports the law instead of replacing it.
    """
    remaining = list(candidates)
    selected: list = []
    regs_available = sum(c.chunk.source_type == "regulation" for c in remaining)
    regs_needed = min(cfg.min_regulations, regs_available)

    while remaining and len(selected) < cfg.final_k:
        n_regs = sum(c.chunk.source_type == "regulation" for c in selected)
        n_hist = sum(c.chunk.source_type == "history" for c in selected)
        slots_left = cfg.final_k - len(selected)

        # Which candidates are eligible for this slot?
        eligible = remaining
        if regs_needed - n_regs >= slots_left:
            # The remaining slots are all needed to meet the regulation quota.
            eligible = [c for c in remaining if c.chunk.source_type == "regulation"]
        if n_hist >= cfg.max_history_clusters:
            eligible = [c for c in eligible if c.chunk.source_type != "history"]
        if not eligible:
            break

        def mmr_score(c: Candidate) -> float:
            if not selected:
                return cfg.mmr_lambda * c.score
            redundancy = max(cosine(vectors[c.chunk.chunk_id], vectors[s.chunk.chunk_id])
                             for s in selected)
            return cfg.mmr_lambda * c.score - (1 - cfg.mmr_lambda) * redundancy

        best = max(eligible, key=mmr_score)
        best.mmr = mmr_score(best)
        selected.append(best)
        remaining.remove(best)

    return selected


# =============================================================================
# 13. THE PIPELINE
# =============================================================================

class RetrievalPipeline:
    """
    Wires the four stages together.

        pipeline = RetrievalPipeline(chunks)            # offline stand-ins
        result = pipeline.retrieve(flag_query)
        result.passages        -> the evidence, best first
        result.no_evidence     -> True means: do NOT raise an orange flag

    Pass `stages` to switch stages off for evaluation (an "ablation"), e.g.
    stages={"filter", "fusion", "score"} runs without Stage 4.
    """

    ALL_STAGES = frozenset({"filter", "fusion", "score", "collapse", "mmr"})

    def __init__(self, chunks: Sequence[Chunk], embedder: Embedder | None = None,
                 reranker: Reranker | None = None, config: RankingConfig = DEFAULT_CONFIG):
        self.chunks = {c.chunk_id: c for c in chunks}
        self.embedder = embedder or HashingEmbedder()
        self.reranker = reranker or OverlapReranker(self.embedder)
        self.config = config
        self.bm25 = BM25Index().fit(chunks)
        self.vindex = VectorIndex(self.embedder).fit(list(chunks))
        self.index_version = self._fingerprint()

    def _fingerprint(self) -> str:
        """A short hash of every chunk's id and text: the index version."""
        h = hashlib.sha256()
        for cid in sorted(self.chunks):
            h.update(cid.encode())
            h.update(self.chunks[cid].text.encode())
        return h.hexdigest()[:16]

    def retrieve(self, query: FlagQuery, *, config: RankingConfig | None = None,
                 stages: Iterable[str] = ALL_STAGES, trace: bool = False) -> RetrievalResult:
        cfg = config or self.config
        stages = set(stages)
        log: dict = {}

        # --- Stage 1: filter -------------------------------------------------
        pool = (stage1_filter(self.chunks.values(), query) if "filter" in stages
                else list(self.chunks.values()))
        allowed = {c.chunk_id for c in pool}
        if trace:
            log["stage1_kept"] = sorted(allowed)
            log["stage1_removed"] = sorted(set(self.chunks) - allowed)

        # --- Stage 2: keyword + semantic search, fused ------------------------
        q_text = query.text()
        kw = self.bm25.search(q_text, allowed, cfg.candidates_per_method)
        sem = self.vindex.search(q_text, allowed, cfg.candidates_per_method)
        if "fusion" in stages:
            fused = reciprocal_rank_fusion([kw, sem], k=cfg.rrf_k)
        else:
            fused = reciprocal_rank_fusion([sem], k=cfg.rrf_k)   # semantic only
        top_ids = sorted(fused, key=fused.get, reverse=True)[:cfg.rerank_top_n]
        candidates = [Candidate(chunk=self.chunks[cid], rrf=fused[cid]) for cid in top_ids]
        if trace:
            log["stage2_keyword"], log["stage2_semantic"] = kw[:10], sem[:10]

        # --- Stage 3: rerank, weight, floor ------------------------------------
        if "score" in stages:
            candidates = stage3_score(candidates, query, self.reranker, cfg)
        else:
            for c in candidates:                    # relevance = fused rank only
                c.relevance = c.score = c.rrf
        if trace:
            log["stage3"] = [(c.chunk.chunk_id, round(c.relevance, 3), round(c.score, 3))
                             for c in candidates]

        if not candidates:
            # Nothing cleared the floor: report it rather than flag on a guess.
            return RetrievalResult(query, [], True, self.index_version, cfg, log)

        # --- Stage 4: maximum relevance, minimum redundancy --------------------
        vecs = self.vindex.text_vectors      # text-only vectors: see VectorIndex
        if "collapse" in stages:
            candidates = collapse_near_duplicates(candidates, vecs, cfg)
        if "mmr" in stages:
            final = mmr_select(candidates, vecs, cfg)
        else:
            final = candidates[:cfg.final_k]       # plain top-k by score
        if trace:
            log["stage4"] = [(c.chunk.chunk_id, c.cluster_size) for c in final]

        return RetrievalResult(query, final, False, self.index_version, cfg, log)


# =============================================================================
# 14. HANDING THE EVIDENCE TO CLAUDE, WITH CITATIONS
# =============================================================================

def to_search_result_blocks(result: RetrievalResult) -> list:
    """
    Convert the final passages into Claude "search_result" content blocks.

    Claude does not search; this code does, and passes the results in this
    format. With citations enabled, Claude's answer cites the blocks it used,
    which gives every orange flag its "sources cited for the reviewer".
    Collapsed clusters carry their summary line as a second text block, so
    Claude sees "12 similar cases: 9 dismissed, 3 corrected" with the case.
    """
    blocks = []
    for cand in result.passages:
        c = cand.chunk
        content = [{"type": "text", "text": c.text}]
        if cand.cluster_summary:
            content.append({"type": "text", "text": f"Pattern history: {cand.cluster_summary}."})
        blocks.append({
            "type": "search_result",
            "source": f"kb://{c.source_type}/{c.chunk_id}",
            "title": c.title,
            "content": content,
            "citations": {"enabled": True},
        })
    return blocks


# The order a reviewer works through an orange flag; with extended thinking the model is told to follow it.
REASONING_STEPS = ("which wage source applies to this line (state, federal or union), which date governs it (the "
                   "bid due date or the award date), which rate, scope or overtime rule was in effect then, and "
                   "only then whether the cell meets it")


def judge_cell(result: RetrievalResult, *, model: str = "claude-sonnet-5-5",
               client=None, thinking_effort: str = "") -> dict:
    """
    Ask Claude for the orange-flag judgment on one cell, using only the
    retrieved evidence. Requires `pip install anthropic` and an API key.

    thinking_effort ("low", "medium", "high" or "max") turns on adaptive
    thinking (chain of thought kept out of the visible answer): the model
    decides how much to reason, guided by the effort level (the API's
    output_config.effort), and is told the steps to work through
    (REASONING_STEPS). Current models take adaptive thinking only, not a
    fixed thinking budget. The visible answer keeps
    the same form ("FLAG: ..." / "NO FLAG: ..."); the reasoning is not kept.

    Returns {"raise_flag": bool, "explanation": str, "citations": [...],
    "usage": {"input": n, "output": n}, "seconds": s}.
    If retrieval found no evidence, no call is made and no flag is raised.
    """
    if result.no_evidence:
        return {"raise_flag": False, "explanation": "No supporting evidence found.",
                "citations": []}
    if client is None:
        import anthropic                     # noqa: imported only when used
        client = anthropic.Anthropic()

    instruction = (
        "You are assisting a prevailing-wage compliance reviewer. A payroll cell "
        "passed every hard-coded rule but looked suspicious:\n"
        f"{result.query.text()}\nProject: {result.query.project_id} "
        f"(applicable wage sources: {', '.join(sorted(result.query.applicable_sources))}). "
        f"Week ending {result.query.week_end:%B %d, %Y}.\n\n"
        "Using ONLY the search results above, decide whether this cell deserves a "
        "reviewer's attention. Start your answer with 'FLAG:' or 'NO FLAG:', then "
        "explain in two or three sentences, citing the sources you relied on. "
        "Past decisions are context, not rules: the law and contract terms come first."
    )
    extra = {}
    if thinking_effort:
        instruction += (f" Before answering, work out {REASONING_STEPS}; weigh both readings (the flag holds, the "
                        "flag doesn't) before you decide.")
        # adaptive thinking: max_tokens caps thinking and answer together, so it gets ample room. The effort level
        # goes in output_config, sent as extra_body so older SDK versions pass it through.
        extra = {"thinking": {"type": "adaptive"}, "max_tokens": 16000,
                 "extra_body": {"output_config": {"effort": str(thinking_effort)}}}
    import time as _time
    t0 = _time.monotonic()
    # a thinking call takes longer than the client's usual 30 s limit (agent/ai_review.py): its own, longer one
    caller = client.with_options(timeout=180) if thinking_effort and hasattr(client, "with_options") else client
    msg = caller.messages.create(
        model=model, max_tokens=extra.pop("max_tokens", 4000),    # room for thinking plus the answer
        messages=[{"role": "user",
                   "content": to_search_result_blocks(result) + [{"type": "text", "text": instruction}]}],
        **extra,
    )
    seconds = _time.monotonic() - t0
    text = _reply_text(msg)
    cites = []
    for b in msg.content:
        for ci in getattr(b, "citations", None) or []:
            cites.append({"source": getattr(ci, "source", None),
                          "title": getattr(ci, "title", None),
                          "cited_text": getattr(ci, "cited_text", None)})
    u = getattr(msg, "usage", None)
    details = getattr(u, "output_tokens_details", None)
    return {"raise_flag": text.strip().upper().startswith("FLAG"),
            "explanation": text.strip(), "citations": cites, "seconds": round(seconds, 2),
            "usage": {"input": getattr(u, "input_tokens", None), "output": getattr(u, "output_tokens", None),
                      "thinking": getattr(details, "thinking_tokens", None) if details is not None else None}}


# =============================================================================
# 15. EVALUATION
# =============================================================================
# The test set is built from the firm's own history: past borderline cells,
# each with the chunk ids a reviewer marked as actually justifying the
# decision. These functions measure retrieval and ranking separately from the
# model's final judgment, so a weak flag can be traced to bad evidence or to
# bad reasoning.

@dataclass
class LabeledQuery:
    query: FlagQuery
    relevant_ids: set        # chunk ids a reviewer marked as justifying


def recall_at_k(selected_ids: Sequence[str], relevant: set, k: int) -> float:
    """Share of the justifying passages that reached the top k."""
    if not relevant:
        return 1.0
    return len(set(selected_ids[:k]) & relevant) / len(relevant)


def ndcg_at_k(selected_ids: Sequence[str], relevant: set, k: int) -> float:
    """
    Normalized discounted cumulative gain: rewards putting the justifying
    passages FIRST. 1.0 means every relevant passage is ranked above every
    irrelevant one.
    """
    dcg = sum(1.0 / math.log2(i + 2) for i, cid in enumerate(selected_ids[:k]) if cid in relevant)
    ideal = sum(1.0 / math.log2(i + 2) for i in range(min(len(relevant), k)))
    return dcg / ideal if ideal else 1.0


def duplicate_rate(result: RetrievalResult, vectors: dict,
                   threshold: float = 0.95) -> float:
    """
    Share of the final passages that repeat one already selected: same past-
    case pattern, or near-identical text. Target: under 10%.
    """
    seen_keys, seen_vecs, dupes = set(), [], 0
    for c in result.passages:
        key = c.chunk.pattern_key()
        v = vectors[c.chunk.chunk_id]
        if (key and key in seen_keys) or any(cosine(v, s) >= threshold for s in seen_vecs):
            dupes += 1
        if key:
            seen_keys.add(key)
        seen_vecs.append(v)
    return dupes / len(result.passages) if result.passages else 0.0


def validity_errors(result: RetrievalResult) -> int:
    """Superseded or out-of-scope passages in the final set. Target: zero."""
    return sum(not passes_filter(c.chunk, result.query) for c in result.passages)


def evaluate(pipeline: RetrievalPipeline, labeled: Sequence[LabeledQuery], *,
             config: RankingConfig | None = None,
             stages: Iterable[str] = RetrievalPipeline.ALL_STAGES) -> dict:
    """
    Run every labeled query and average the metrics. Call it once per
    configuration or stage set to compare them; switching stages off one at a
    time shows what each contributes.
    """
    cfg = config or pipeline.config
    totals = defaultdict(float)
    for lq in labeled:
        res = pipeline.retrieve(lq.query, config=cfg, stages=stages)
        ids = [c.chunk.chunk_id for c in res.passages]
        totals["recall@k"] += recall_at_k(ids, lq.relevant_ids, cfg.final_k)
        totals["ndcg@k"] += ndcg_at_k(ids, lq.relevant_ids, cfg.final_k)
        totals["duplicate_rate"] += duplicate_rate(res, pipeline.vindex.text_vectors,
                                                   cfg.near_duplicate_similarity)
        totals["validity_errors"] += validity_errors(res)
    n = max(len(labeled), 1)
    return {m: (v if m == "validity_errors" else round(v / n, 3)) for m, v in totals.items()}
