"""cva_filekinds: the file-kind registry of the Compliance Validation Agent (docs/vNext.md section 6).

One definition per kind of file: how it is recognised, the columns its readers use and their types, and what of each
column may cross to AWS. The tokenizer and the agent both import it, so the two can no longer drift apart.

The one source is code/filekinds/cva_filekinds/. ``python code/filekinds/sync.py`` copies it into
code/tokenizer/pii_tokenizer/cva_filekinds/ and code/compliance_agent/agent/cva_filekinds/; a test in each component
fails if its copy differs. Edit the source, never a copy. Every change to a kind is a privacy change: it gets an
independent review and bumps the kind's ``version`` (and REGISTRY_VERSION).

Kinds so far: the wage table and the agent's merged wage table (vNext PR 2); the tracker, Project Info and Project Data
(PR 3); payroll exports (PR 4). Not a workbook kind: the administrator's settings file (settings_file, issue #151).
"""
from .core import (AS_IS, CROSSES, DROP, TEXT_TOKEN, TOKEN, TYPES, Column, FileKind, Recognition, columns_read,
                   crosses_to_aws, norm_header, recognise_workbook, sheet_index)
from .merged_wage_table import MERGED_WAGE_TABLE
from .payroll_export import PAYROLL_EXPORT
from .project_data import PROJECT_DATA
from .project_info import PROJECT_INFO
from .tracker import TRACKER
from .wage_table import WAGE_TABLE

REGISTRY_VERSION = "9"
# Order matters: the first kind that recognises a workbook wins. The merged table's row 1 also has a Classification
# and a Basic column, and the tokenizer has always decided it as the agent's own output first (doubt.sheet_rules).
# A payroll export comes before the wage table: the tokenizer never takes a sheet with a week or hours column for a wage
# table (wage_table.PAYROLLISH), and every payroll header has the week-end date. A tracker workbook also has a Project
# Info tab, so the tracker comes before it.
KINDS = (MERGED_WAGE_TABLE, PAYROLL_EXPORT, WAGE_TABLE, TRACKER, PROJECT_INFO, PROJECT_DATA)
BY_NAME = {k.kind: k for k in KINDS}

__all__ = ["AS_IS", "CROSSES", "DROP", "TEXT_TOKEN", "TOKEN", "TYPES", "Column", "FileKind", "Recognition",
           "columns_read", "crosses_to_aws", "norm_header", "recognise_workbook", "sheet_index", "KINDS", "BY_NAME",
           "REGISTRY_VERSION", "MERGED_WAGE_TABLE", "PAYROLL_EXPORT", "WAGE_TABLE", "TRACKER", "PROJECT_INFO",
           "PROJECT_DATA"]
