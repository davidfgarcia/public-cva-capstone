"""Optional local model (Ollama or any OpenAI-compatible server, e.g. vLLM / llama.cpp server).
Nothing here is required for a run: if the endpoint is down or disabled, callers get None/[]."""
import json
import re
from pathlib import Path

import requests

from .util import err_text, log

WAGE_SCHEMA_HINT = (
    'Return ONLY a JSON array. Each element: {"craft": str, "classification": str, '
    '"wage_decision": str, "issue_date": "YYYY-MM-DD" or null, "basic": number or null, '
    '"fringe": number or null, "total": number or null, "ot_total": number or null, '
    '"county": str or null}. One element per wage classification. Do not invent values.')


class LocalLLM:
    def __init__(self, cfg):
        c = cfg.get("llm") or {}
        self.enabled = c.get("enabled", "auto")
        self.provider = c.get("provider", "ollama")
        self.endpoint = c.get("endpoint", "http://localhost:11434").rstrip("/")
        self.model = c.get("model", "qwen3:8b")
        self.timeout = c.get("timeout_seconds", 300)
        self.think = c.get("think", False)     # reasoning models (e.g. qwen3): skip the "thinking" text
        self._ok = None
        self.model_id = None
        root = cfg.get("root")
        # remembers which model files were used, so a silent change of a tag shows up in the logs
        self.state_file = (Path(root) / "run_logs" / "model_ids.json") if root else None

    def available(self):
        if self.enabled is False or str(self.enabled).lower() == "false":
            return False
        if self._ok is None:
            self._ok = False
            try:
                url = f"{self.endpoint}/api/tags" if self.provider == "ollama" else f"{self.endpoint}/v1/models"
                r = requests.get(url, timeout=3)
                if r.ok:
                    self._ok = self._identify(r.json())
            except Exception:
                pass
            if self._ok is False and self.model_id is None:
                log.info("Local model not reachable at %s; continuing without it", self.endpoint)
        return self._ok

    def _identify(self, listing):
        """Find the configured model on the server, log its ID, and warn if it changed since the last run."""
        if self.provider == "ollama":
            models = {m.get("name"): m for m in listing.get("models", [])}
            m = models.get(self.model) or models.get(f"{self.model}:latest")
            if not m:
                log.warning("Model %s is not installed on the model server; continuing without it", self.model,
                            extra={"details": {"endpoint": self.endpoint, "installed": ", ".join(sorted(models)) or "none",
                                               "fix": f"ollama pull {self.model}"}})
                self.model_id = ""
                return False
            self.model_id = (m.get("digest") or "")[:12]
            info = {"endpoint": self.endpoint, "model": self.model, "model_id": self.model_id,
                    "size_gb": round((m.get("size") or 0) / 1e9, 1),
                    "quantization": (m.get("details") or {}).get("quantization_level", "")}
        else:
            ids = [d.get("id") for d in listing.get("data", [])]
            if self.model not in ids:
                log.warning("Model %s is not served at %s; continuing without it", self.model, self.endpoint,
                            extra={"details": {"served": ", ".join(i for i in ids if i)}})
                self.model_id = ""
                return False
            self.model_id = self.model
            info = {"endpoint": self.endpoint, "model": self.model}
        log.info("Local model available: %s (ID %s)", self.model, self.model_id or "n/a", extra={"details": info})
        self._compare_with_last(info)
        return True

    def _compare_with_last(self, info):
        if not self.state_file:
            return
        try:
            last = json.loads(self.state_file.read_text())
        except Exception:
            last = {}
        prev = (last.get(self.model) or {}).get("model_id")
        if prev and prev != self.model_id:
            log.warning("The files of model %s changed since the last run (ID %s, before %s). Results may differ; "
                        "run the model check (python run.py --check-model)", self.model, self.model_id, prev,
                        extra={"details": {"model": self.model, "previous_id": prev, "current_id": self.model_id}})
        if prev != self.model_id:
            from .util import utc_stamp
            last[self.model] = {"model_id": self.model_id, "first_seen_utc": utc_stamp(),
                                **{k: v for k, v in info.items() if k not in ("model",)}}
            try:
                self.state_file.parent.mkdir(parents=True, exist_ok=True)
                self.state_file.write_text(json.dumps(last, indent=2))
            except Exception:
                pass

    def complete(self, prompt, json_mode=True):
        if not self.available():
            return None
        try:
            if self.provider == "ollama":
                body = {"model": self.model, "prompt": prompt, "stream": False,
                        "options": {"temperature": 0}, "think": self.think}
                if json_mode:
                    body["format"] = "json"
                r = requests.post(f"{self.endpoint}/api/generate", json=body, timeout=self.timeout)
                r.raise_for_status()
                return r.json().get("response")
            body = {"model": self.model, "temperature": 0,
                    "messages": [{"role": "user", "content": prompt}]}
            r = requests.post(f"{self.endpoint}/v1/chat/completions", json=body, timeout=self.timeout)
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"]
        except Exception as e:
            log.warning("Local model call failed: %s", err_text(e))
            return None

    @staticmethod
    def _json(text):
        if not text:
            return None
        try:
            return json.loads(text)
        except Exception:
            m = re.search(r"(\[.*\]|\{.*\})", text, re.S)
            if m:
                try:
                    return json.loads(m.group(1))
                except Exception:
                    return None
        return None

    def _file_text(self, path: Path):
        ext = path.suffix.lower()
        if ext == ".pdf":
            from .pdftext import pdf_text
            return pdf_text(path)
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
            return "\n".join("\t".join("" if c is None else str(c) for c in r)
                             for r in ws.iter_rows(values_only=True))
        return path.read_text(errors="replace")

    def extract_wage_rows(self, path: Path, chunk_chars=12000):
        from .util import to_date, to_float
        text = self._file_text(path)
        out = []
        for i in range(0, len(text), chunk_chars):
            chunk = text[i:i + chunk_chars]
            resp = self._json(self.complete(
                f"Extract the prevailing/union wage rates from this document excerpt.\n{WAGE_SCHEMA_HINT}\n"
                f"If there are none, return [].\n\n---\n{chunk}\n---"))
            if isinstance(resp, dict):
                resp = next((v for v in resp.values() if isinstance(v, list)), [])
            for r in resp or []:
                if not isinstance(r, dict) or not r.get("classification"):
                    continue
                out.append(dict(
                    craft=str(r.get("craft") or r["classification"]),
                    classification=str(r["classification"]),
                    wage_decision=str(r.get("wage_decision") or path.stem),
                    issue_date=to_date(r.get("issue_date")), basic=to_float(r.get("basic")),
                    fringe=to_float(r.get("fringe")), total=to_float(r.get("total")),
                    ot_total=to_float(r.get("ot_total")), county=r.get("county"),
                    notes="extracted by local model - verify"))
        return out

    def suggest_classification(self, craft, classification, candidates):
        if not candidates:
            return None
        resp = self._json(self.complete(
            "A certified payroll line lists this trade:\n"
            f"  craft: {craft}\n  classification: {classification}\n"
            "Which ONE of the following wage-determination classifications is the same work? "
            'Answer as JSON {"match": "<exact line from the list>" or null, "confidence": 0-1}.\n\n'
            + "\n".join(candidates)))
        if isinstance(resp, dict) and resp.get("match") in candidates and (resp.get("confidence") or 0) >= 0.6:
            return resp["match"]
        return None
