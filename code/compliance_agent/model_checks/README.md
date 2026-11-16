# Model test set

Used by `python run.py --check-model` (see docs/installation.md, section 8.4). Each check sends these files to the
local model and compares its answers with the expected ones below. Everything stays on the machine running the check.

* `wage_files/<file>` + `wage_files/<file>.expected.csv`: a wage file and the rows it should produce
  (craft, classification, basic, fringe, total; blank = not checked). The starter files are a real federal wage
  determination (WA20250093) and a 40-row sample of an L&I state wage CSV; their expected rows come from the agent's
  built-in parsers, which read these formats exactly.
* `candidates/<file>` + `classifications.csv`: payroll classifications without a wage rate, the wage file whose
  classifications are offered, and the right answer (`CRAFT | Classification` exactly as listed, or `none`).
  The starter cases come from the Selkirra payroll.
* `results/<ts>__model_check.csv`: one report per check (model, model ID, each test, PASS/FAIL, timing).

To add a case: drop a file in `wage_files/` with its `.expected.csv`, or add a row to `classifications.csv`.
Best additions are real files the built-in parsers could not read, with rows you checked by hand.
