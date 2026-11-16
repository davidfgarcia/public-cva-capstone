"""Federal Davis-Bacon wage determinations downloaded from SAM.gov.

A project lists the determinations that apply to it, each with the modification locked in at contract award:

    projects:
      "Atrovan - ...":
        federal_wage_determinations:
          - {decision: WA20240001, mod: 4}       # or mod: latest

Before the wage merge, each one not yet in the project's federal rates is downloaded as text (the same format staff
download by hand) into the project's Uploads/ folder as "<decision> mod <n> (SAM.gov).txt"; the normal merge then
reads it, records it and archives it. If SAM.gov can't be reached (after retries for temporary failures), the run
goes on with the federal rates already in the project, and the email says the download failed."""
from __future__ import annotations

import re
from pathlib import Path

from .util import log

DEFAULT_BASE = "https://sam.gov/api/prod/wdol/v1/wd"
DECISION_RE = re.compile(r"^[A-Z]{2}\d{8}$")


def _settings(cfg):
    s = (cfg.get("wages") or {}).get("sam_gov") or {}
    return bool(s.get("enabled", False)), (s.get("base") or DEFAULT_BASE).rstrip("/"), int(s.get("timeout_seconds", 60))


def wanted(cfg, project_name):
    """[(decision, mod or 'latest')] from projects.<name>.federal_wage_determinations, plus problems found."""
    from .config import project_cfg
    out, bad = [], []
    for e in project_cfg(cfg, project_name).get("federal_wage_determinations") or []:
        if isinstance(e, str):
            e = {"decision": e}
        d = str((e or {}).get("decision") or "").strip().upper()
        mod = (e or {}).get("mod")
        if not DECISION_RE.match(d):
            bad.append(f"'{d or e}' is not a wage decision number like WA20240001")
            continue
        if mod is None:
            bad.append(f"{d}: no 'mod' given; set the modification in force at contract award (or mod: latest)")
            continue
        if str(mod).lower() == "latest":
            out.append((d, "latest"))
        elif str(mod).isdigit():
            out.append((d, int(mod)))
        else:
            bad.append(f"{d}: mod '{mod}' should be a number or 'latest'")
    return out, bad


def have(rates_rows, decision, mod):
    """True when the project's federal rates already hold that decision at that modification, as a rate that says
    its wage type (one merged from a file without a Jurisdiction column isn't used, so it doesn't count: the official
    determination is downloaded then)."""
    return any(str(r.get("wage_decision") or "").upper() == decision and (r.get("mod") or 0) == mod
               and r.get("jurisdiction") for r in rates_rows)


def latest_mod(decision, base, timeout, getter):
    r = getter(f"{base}/{decision}/history", timeout=timeout, what=f"SAM.gov history of {decision}")
    revs = [w.get("revisionNumber") for w in (r.json().get("_embedded") or {}).get("wageDetermination", [])]
    revs = [x for x in revs if isinstance(x, int)]
    if not revs:
        raise ValueError(f"SAM.gov lists no modifications of {decision}")
    return max(revs)


def document(decision, mod, base, timeout, getter) -> str:
    r = getter(f"{base}/{decision}/{mod}", timeout=timeout, what=f"SAM.gov {decision} mod {mod}")
    text = str(r.json().get("document") or "").strip()
    if len(text) >= 2 and text[0] == text[-1] == '"':
        text = text[1:-1]
    if not re.search(rf"General Decision Number:\s*{decision}\b", text):
        raise ValueError(f"SAM.gov's answer for {decision} mod {mod} is not that wage determination")
    mods = re.findall(r"^\s*(\d+)\s+\d{2}/\d{2}/\d{4}\s*$", text, re.M)
    if mods and int(mods[-1]) != mod:
        raise ValueError(f"SAM.gov returned modification {mods[-1]} when {mod} was asked for")
    return text


def fetch_for_project(project_dir: Path, cfg, getter=None) -> list[str]:
    """Download the project's federal wage determinations that aren't in its rates yet. Returns summary lines."""
    enabled, base, timeout = _settings(cfg)
    want, bad = wanted(cfg, Path(project_dir).name)
    lines = [f"WARNING federal: projects.{Path(project_dir).name}.federal_wage_determinations: {b}" for b in bad]
    if not want or not enabled:
        return lines
    from .layout import Paths
    from .wages import read_table
    from . import net
    getter = getter or net.get
    paths = Paths(project_dir)
    rows = read_table(paths.rate_file("federal"))
    for decision, mod in want:
        try:
            m = latest_mod(decision, base, timeout, getter) if mod == "latest" else mod
            if have(rows, decision, m):
                log.debug("Federal wage determination already in use", extra={"details": {"decision": decision, "mod": m}})
                continue
            name = f"{decision} mod {m} (SAM.gov).txt"
            if (paths.upload / name).exists():
                continue
            text = document(decision, m, base, timeout, getter)
            paths.upload.mkdir(parents=True, exist_ok=True)
            (paths.upload / name).write_text(text)
            log.info("Downloaded a federal wage determination from SAM.gov",
                     extra={"details": {"decision": decision, "mod": m, "file": name}})
            lines.append(f"federal: downloaded {decision} mod {m} from SAM.gov")
        except Exception as e:                         # network, SAM.gov changes: the run goes on without it
            log.warning("Could not download %s from SAM.gov: %s", decision, e,
                        extra={"details": {"decision": decision, "mod": mod}})
            lines.append(f"WARNING federal: could not download {decision} (mod {mod}) from SAM.gov ({e}); the "
                         "federal rates already in the project were used. Download the .txt from "
                         f"sam.gov/wage-determination/{decision} into Uploads if this keeps happening")
    return lines
