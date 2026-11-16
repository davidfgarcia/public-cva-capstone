# Running the CVA demo

The demo runs the Compliance Validation Agent's own web UI on your computer, on **invented data only**: two invented
construction projects, their payroll exports, compliance trackers and wage tables, and invented L&I records served by
a local stand-in. Nothing real is read, and nothing leaves your computer (unless you choose the AI review's Claude
mode, which sends tokenized patterns to the Claude API).

## 1. Set up (once)

You need Python 3.12 or newer and git. On a Mac (Terminal):

```bash
git clone https://github.com/davidfgarcia/public-cva-capstone
cd public-cva-capstone
python3 -m venv .venv && source .venv/bin/activate
pip install -r code/compliance_agent/requirements.txt -r code/tokenizer/requirements.txt pytest
```

## 2. Start it

```bash
bash demo/run_demo.sh                 # builds ~/Downloads/CVA_demo the first time, then opens the UI
```

The page opens at http://127.0.0.1:8775/. Stop it with Ctrl+C in the terminal (or the page's **Quit** button).

Options (all passed through to `demo/run_demo.py`):

| Command | What it does |
|---|---|
| `bash demo/run_demo.sh --reset` | Throws the demo folder away and builds it again: a clean state for another take |
| `bash demo/run_demo.sh --ai claude` | The AI review also asks Claude for a suggestion. Needs `export ANTHROPIC_API_KEY=...` in the same terminal first (never stored); costs a few cents |
| `bash demo/run_demo.sh --demo /some/folder` | Uses another folder for the demo data |
| `python demo/show_tokenized.py` | Runs the real tokenizer offline on the demo files and opens a page comparing a payroll as staff see it and as AWS sees it |
| `python demo/check_demo.py` | Runs everything end to end on a fresh copy and scores it against the answer key (no browser needed) |

The default AI review mode is **evidence only**: the review retrieves the passages that bear on each orange pattern
and lists them, with no model call. In the real system the review runs in shadow mode (suggestions hidden from staff);
the demo turns shadow mode off so the suggestions can be seen.

## 3. What is in the demo

| Project | What it has |
|---|---|
| **Quillmere Construction - Brightfen Library Annex** | A prime, 3 subcontractors at work and 1 not started yet, and 2 second-tier subs; 2 weekly payrolls (30 lines) with 3 seeded findings; a tracker with every kind of intent/affidavit finding |
| **Norvanta Builders - Kestrel Point Housing Phase 2** | A prime, 10 subcontractors and 2 second-tier subs; 3 weekly payrolls in Uploads plus a 4th in `To upload/` (259 lines in all) with 17 seeded findings across R01–R08 and S01–S15; a union trade; a resubmitted file with two exports pasted together, which is skipped with a plain reason; a tracker with every kind of finding and a contractor L&I has that the tracker lacks |

Every name is invented: company names are built on coined words (checked against Washington L&I's public contractor
licence data on Nov 16, 2026: no match), workers have employee IDs in the range 7320001–7320999, intent IDs are
88200nn, affidavit IDs 88300nn and L&I project numbers 99005nn. `demo_data/answer_key.json` lists every seeded line
and what the CVA must show for it.

## 4. The click path

1. **Add files** tab: the two projects and their unprocessed files (payroll exports, wage tables). Tick
   *Norvanta Builders - Kestrel Point Housing Phase 2*, then drop `~/Downloads/CVA_demo/To upload/Kestrel Point
   PayrollDetailReport WE 11-15-2026.xlsx` on **Drop files here**.
2. **Run tasks** tab: tick **All** under *Payroll review* and under *Intent/affidavit check*, keep **Incremental**,
   click **Run 4 tasks**. It finishes in a few seconds: *Finished · 4 to review*, with each task's findings.
3. **Results** tab: the four results, with Run and Findings for each.
4. Open the outputs in Finder (`~/Downloads/CVA_demo/Projects/<project>/`):
   - `Payroll/<date> ... - REVIEWED.xlsx`: red and orange cells with notes, the **Comments** sheet listing every
     issue, the **Review decisions** sheet with the AI review's evidence for each orange pattern;
   - `Tracker/... Compliance Tracker.xlsx`: red cells with notes and the **Review Log** sheet;
   - `~/Downloads/CVA_demo/run_logs/*__email_preview_staff.eml` and `..._admin.eml`: the two outcome emails
     (double-click to open them in Mail; they were never sent).
5. Optional: `python demo/show_tokenized.py` for the privacy boundary: the same payroll with names and IDs as
   tokens, and the file names on the AWS side.

To record again from the start: Ctrl+C, then `bash demo/run_demo.sh --reset`.

## 5. What keeps the demo apart from a real set-up

- The UI runs with `HOME` set to `<demo folder>/.home`, so it reads and writes its own settings there, never
  `~/.compliance_agent/settings.json`.
- The only network route is the local L&I stand-in (`demo/lni_standin.py`), set as the proxy for every request: it
  answers L&I's lookup with invented records and refuses everything else (every request is recorded in
  `demo_data/lni_requests.json`). In `--ai claude` mode, and only then, the Claude API host goes around it.
- No AWS, Google, Sentry or email credentials are passed to the UI; the outcome emails are saved as `.eml` previews.
- `--reset` deletes only a folder that the generator made (it carries a `.cva-demo` marker file).
