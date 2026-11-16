"""Reference data extracted from the reference rules workbook, so the agent never needs the workbook itself.

The reference rules workbook (WageWorkbookNew.xlsx) holds proprietary information and stays on the computer it lives on: it is
never read by scheduled runs, never synced to Google Drive, AWS or S3 (deploy/rclone-filter.txt excludes
private_reference/), and never bundled with the agent.

What the agent needed from it:
  * the review rules (the formulas of its Data sheet): built into agent/payroll.py as R01-R08 and S01; the mapping of
    each formula to its rule is in docs/description.md, section 4.6;
  * the union wage table (sheet "WSCCA Wages"): extracted with
        python run.py --extract-union-rates <path to WageWorkbookNew.xlsx>
    into <root>/reference/union_wages.xlsx (rates only: craft, classification, rates, dates, Vac/PTO). Projects
    without their own union wage data use that table. Re-run the command when the workbook's union table changes;
    the previous table is kept in reference/archive/<ts>__union_wages.xlsx.
Nothing else from the workbook (its Data sheet holds past payroll lines with employee names) is copied.
"""
import hashlib
import logging
import shutil
from pathlib import Path

import openpyxl

from .util import run_timestamp, utc_stamp
from .wages import parse_union_table, read_table, write_table

log = logging.getLogger("compliance.reference")
UNION_SHEET = "WSCCA Wages"


def union_reference_path(cfg):
    return cfg["root"] / cfg["wages"].get("union_reference", "reference/union_wages.xlsx")


def extract_union_rates(workbook, cfg, sheet=UNION_SHEET):
    """Reads the union table from the reference rules workbook (read-only) and writes reference/union_wages.xlsx.
    Returns (output path, number of rows, archived previous file or None)."""
    workbook = Path(workbook)
    digest = hashlib.sha256(workbook.read_bytes()).hexdigest()[:12]
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet}' not found in {workbook.name} (sheets: {', '.join(wb.sheetnames)})")
    rows = parse_union_table(wb[sheet], "union", f"reference rules workbook / {sheet}")
    wb.close()
    if not rows:
        raise ValueError(f"No rates found in '{sheet}' of {workbook.name}")
    out = union_reference_path(cfg)
    out.parent.mkdir(parents=True, exist_ok=True)
    archived = None
    if out.exists():
        arch = out.parent / "archive"
        arch.mkdir(exist_ok=True)
        archived = arch / f"{run_timestamp()}__{out.name}"
        shutil.copy2(out, archived)
    write_table(out, rows, [[utc_stamp(), f"{workbook.name} / {sheet} (sha256 {digest})", len(rows)]])
    check = read_table(out)
    if len(check) != len(rows):
        raise RuntimeError(f"Wrote {len(rows)} rows but read back {len(check)}")
    log.info("Union rates extracted", extra={"details": {"rows": len(rows), "out": Path(out).name, "workbook_sha256": digest}})
    return out, len(rows), archived


def load_union_reference(cfg):
    """Union rates for projects without their own union wage data ([] if the table is missing)."""
    p = union_reference_path(cfg)
    if not p.exists():
        log.warning("Union reference table %s not found; union lines without project union data get no rate. "
                    "Create it with: python run.py --extract-union-rates <reference rules workbook>", Path(p).name)
        return []
    return read_table(p)
