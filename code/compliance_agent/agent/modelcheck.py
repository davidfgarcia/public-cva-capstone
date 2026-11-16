"""Model check: run the local model on a fixed set of test files with known answers and report whether its output
still matches. Use it after installing or updating Ollama, before switching to another model, or when the run log
warns that the model's files changed.

    python run.py --check-model                    # the model in config.yaml
    python run.py --check-model --model qwen3:14b  # a candidate, without changing config.yaml

Test set (folder `model_checks/` next to run.py, or llm.check_dir in config.yaml):
  wage_files/<file>                      a wage file to read with the model (any format)
  wage_files/<file>.expected.csv         the rows it should produce: craft, classification, basic, fringe, total
  candidates/<wage file>                 wage files whose classifications are offered to the model
  classifications.csv                    candidates_file, kind, craft, classification, expected
                                         (expected = "CRAFT | Classification" as listed, or "none")
  results/<ts>__model_check.csv          written by each check (<ts> = yyyy-MM-dd-HH-mm-ss, UTC)
"""
import csv
import time
from difflib import SequenceMatcher
from pathlib import Path

from .config import AGENT_DIR
from .llm import LocalLLM
from .util import log, norm, run_timestamp, to_float
from .wages import sniff_and_parse

TOL = 0.011


def _check_dir(cfg):
    d = Path((cfg.get("llm") or {}).get("check_dir") or "model_checks")
    return d if d.is_absolute() else AGENT_DIR / d


def _key(craft, cls):
    return norm(f"{craft or ''} {cls or ''}"), norm(cls or "")


def _score_wage_file(llm, f: Path, expected_csv: Path, min_score):
    expected = list(csv.DictReader(open(expected_csv, newline="", encoding="utf-8-sig")))
    t0 = time.time()
    got = llm.extract_wage_rows(f)
    secs = time.time() - t0
    by_full, by_cls = {}, {}
    for r in got:
        k_full, k_cls = _key(r.get("craft"), r.get("classification"))
        by_full.setdefault(k_full, r)
        by_cls.setdefault(k_cls, r)
    ok, wrong, missing, details = 0, 0, 0, []
    for e in expected:
        k_full, k_cls = _key(e.get("craft"), e.get("classification"))
        r = by_full.get(k_full) or by_cls.get(k_cls)
        if r is None:
            best = max(by_full, key=lambda k: SequenceMatcher(None, k, k_full).ratio(), default=None)
            if best and SequenceMatcher(None, best, k_full).ratio() >= 0.9:
                r = by_full[best]
        if r is None:
            missing += 1
            details.append(f"missing: {e.get('craft')} / {e.get('classification')}")
            continue
        diffs = []
        for f_ in ("basic", "fringe", "total"):
            want, have = to_float(e.get(f_)), to_float(r.get(f_))
            if want is not None and (have is None or abs(want - have) > TOL):
                diffs.append(f"{f_} {have} (expected {want})")
        if diffs:
            wrong += 1
            details.append(f"wrong: {e.get('classification')}: " + "; ".join(diffs))
        else:
            ok += 1
    extra = max(0, len(got) - ok - wrong)
    score = ok / len(expected) if expected else 0.0
    passed = score >= min_score and extra <= max(1, int(0.05 * len(expected)))
    return {"test": f"wage file: {f.name}", "passed": passed,
            "result": f"{ok}/{len(expected)} rows right, {wrong} wrong, {missing} missing, {extra} extra",
            "seconds": round(secs, 1), "details": " | ".join(details[:15])}


def _score_classifications(llm, d: Path):
    out = []
    path = d / "classifications.csv"
    if not path.exists():
        return out
    cand_cache = {}
    for row in csv.DictReader(open(path, newline="", encoding="utf-8-sig")):
        src = d / "candidates" / row["candidates_file"]
        if src not in cand_cache:
            rows = sniff_and_parse(src, row.get("kind") or "federal")      # built-in parsers only (no model)
            cand_cache[src] = sorted({f"{r['craft']} | {r['classification']}" for r in rows})
        t0 = time.time()
        got = llm.suggest_classification(row["craft"], row["classification"], cand_cache[src])
        secs = time.time() - t0
        want = (row.get("expected") or "").strip()
        passed = (got or "none").strip().lower() == want.lower()
        out.append({"test": f"classification: {row['craft']} / {row['classification']}", "passed": passed,
                    "result": f"suggested {got or 'none'}; expected {want}", "seconds": round(secs, 1), "details": ""})
    return out


def run_check(cfg, model=None):
    llm_cfg = dict(cfg.get("llm") or {})
    llm_cfg["enabled"] = True
    if model:
        llm_cfg["model"] = model
    cfg = {**cfg, "llm": llm_cfg}
    llm = LocalLLM(cfg)
    d = _check_dir(cfg)
    if not d.exists():
        raise SystemExit(f"No model test set at {d}")
    if not llm.available():
        raise SystemExit(f"Model {llm.model} is not available at {llm.endpoint} (is Ollama running and the model pulled?)")
    min_score = float(llm_cfg.get("check_min_score", 0.95))
    results = []
    for f in sorted((d / "wage_files").glob("*")):
        if f.name.endswith(".expected.csv") or f.name.startswith("."):
            continue
        exp = f.with_name(f.name + ".expected.csv")
        if not exp.exists():
            log.warning("No expected rows for %s; skipped", f.name)
            continue
        log.info("Checking %s ...", f.name)
        results.append(_score_wage_file(llm, f, exp, min_score))
    log.info("Checking classification suggestions ...")
    results += _score_classifications(llm, d)
    ts = run_timestamp(cfg.get("timezone", "UTC"))
    res_dir = d / "results"
    res_dir.mkdir(exist_ok=True)
    out = res_dir / f"{ts}__model_check.csv"
    with open(out, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["run_utc", "model", "model_id", "test", "passed", "result", "seconds", "details"])
        for r in results:
            w.writerow([ts, llm.model, llm.model_id, r["test"], "PASS" if r["passed"] else "FAIL",
                        r["result"], r["seconds"], r["details"]])
    passed = sum(r["passed"] for r in results)
    return {"model": llm.model, "model_id": llm.model_id, "passed": passed, "total": len(results),
            "all_passed": passed == len(results) and bool(results), "results": results, "report": str(out)}
