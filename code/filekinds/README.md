# cva_filekinds: the file-kind registry

One definition per kind of file the Compliance Validation Agent handles: how it is recognised, the columns its readers
use and their types, and what of each column may cross to AWS (`docs/vNext.md` section 6;
`docs/filekinds_inventory.md`).

- **The one source** is `cva_filekinds/` here. Standard library only, Python 3.9.
- **Copies:** each component's deploy builds only its own folder, so `python code/filekinds/sync.py` copies the source
  into `code/tokenizer/pii_tokenizer/cva_filekinds/` and `code/compliance_agent/agent/cva_filekinds/` (each `.py` gets
  a "generated copy" first line). Commit the copies with the source.
- **Checks:** `python code/filekinds/sync.py --check` (in CI), and a test in each component
  (`tests/test_filekinds.py`) fail if a copy differs from the source, byte for byte, subfolders included (hidden files such as `.DS_Store` ignored). The Mac's
  deploy scripts don't run the check: deploy from a commit CI has passed.
- **Every change to a kind is a privacy change:** independent review; bump the kind's `version` and `REGISTRY_VERSION`.
- **Tests:** `python -m pytest -q code/filekinds/tests` (3.9 and 3.12 in CI).

Until phase 3, the registry describes; the tokenizer's decision is still `pii_tokenizer/doubt.py sheet_rules`
(`recognise_workbook` and `crosses_to_aws` are not wired in).
