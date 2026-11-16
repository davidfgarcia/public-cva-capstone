#!/usr/bin/env python3
"""Compliance agent command line.

  python run.py                                  # all tasks, all projects (what the schedule runs)
  python run.py --task payroll                   # wage-data merge + payroll review
  python run.py --task intent                    # intent/affidavit validation
  python run.py --task wages                     # only merge new wage determination files
  python run.py --task payroll,intent            # several tasks
  python run.py --project Brovane --project QRX  # one or more projects (name or its beginning)
  python run.py --force                          # re-review the last payroll even if nothing changed
  python run.py --root /data/compliance          # folder that holds projects/ and reference/
  python run.py --ui                             # open the desktop UI
  python run.py --check-model                    # test the local model against model_checks/ (known answers)
  python run.py --check-model --model qwen3:14b  # test a candidate model without changing config.yaml
  python run.py --test-email                     # send a test email to the recipients in notifications.config
  python run.py --no-email                       # run without sending the outcome email
  python run.py --extract-union-rates <WageWorkbookNew.xlsx>   # refresh reference/union_wages.xlsx from the
                                                 # reference rules workbook (run on the computer that holds it)

Every run emails its outcome (success, partial failure or failure) to the recipients in notifications.config.
"""
import argparse
import json
import logging
import os
import sys


def _error_reporting(component, config=None, root=None):
    """Sentry, privacy first (agent/errors.py): DSN from SENTRY_DSN or error_reporting in config.yaml."""
    from agent.config import AGENT_DIR
    from agent.errors import dsn_from_config, environment_from_config, release_name, setup
    cfg = {}
    try:
        from agent.config import load_config
        cfg = load_config(config, root)
    except Exception:                         # a broken config is reported by the run itself
        pass
    try:
        return setup(component, dsn_from_config(cfg), environment_from_config(cfg), release_name(AGENT_DIR))
    except Exception:                         # error reporting must never stop a run
        return False


def console_summary(report: dict) -> dict:
    """What a run prints at its end. On AWS this output goes to the pipeline log and CloudWatch, so it holds no file
    name or path (#56): a summary line can name a payroll, and a payroll can be named after a worker. It gives the
    outcome and, per project, how many summary lines, errors and warnings there were; the lines themselves are in
    the run summary (run_logs/<run>__run_summary.json) and the outcome email."""
    keep = ("run_utc", "finished_utc", "duration_s", "tasks", "force", "trigger", "outcome", "model", "email")
    out = {k: report[k] for k in keep if k in report}
    out["projects"] = {p: {"lines": len(v), "errors": sum(1 for l in v if str(l).startswith("ERROR")),
                           "warnings": sum(1 for l in v if str(l).startswith("WARNING"))}
                       for p, v in (report.get("projects") or {}).items()}
    if report.get("run_utc"):
        out["full_report"] = f"run_logs/{report['run_utc']}__run_summary.json"
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--task", default="all", help="all | wages | payroll | intent  (comma-separate several)")
    ap.add_argument("--project", action="append", help="repeat for several projects")
    ap.add_argument("--plan", action="append", metavar="TASKS:PROJECT",
                    help="each project its own tasks, in one run, e.g. --plan payroll,intent:Brovane --plan intent:QRX "
                         "(repeat; --task and --project are then not used)")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--ai-experiment", type=int, default=0, metavar="N",
                    help="AI review: compare answers without and with extended thinking on up to N patterns")
    ap.add_argument("--config")
    ap.add_argument("--root")
    ap.add_argument("--ui", action="store_true", help="start the desktop UI instead of running")
    ap.add_argument("--check-model", action="store_true", help="test the local model on model_checks/ and exit")
    ap.add_argument("--model", help="with --check-model: the model to test (default: llm.model in config.yaml)")
    ap.add_argument("--trigger", help="how the run was started, shown in the outcome email (default: manual)")
    ap.add_argument("--no-email", action="store_true", help="don't send the outcome email")
    ap.add_argument("--test-email", action="store_true", help="send a test email and exit")
    ap.add_argument("--extract-union-rates", metavar="WORKBOOK",
                    help="copy the union wage table from the reference rules workbook to reference/union_wages.xlsx and exit")
    # used by deploy/run_pipeline.sh to send the outcome email after the results are uploaded
    ap.add_argument("--notify-pipeline", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--tokenizer-failed", default="", help=argparse.SUPPRESS)
    ap.add_argument("--run-projects", default="", help=argparse.SUPPRESS)    # the pipeline's projects, one per line
    ap.add_argument("--empty-projects", default="", help=argparse.SUPPRESS)  # of those, the ones with nothing to pull
    ap.add_argument("--unpulled-projects", default="", help=argparse.SUPPRESS)  # full run: folders with nothing to pull
    ap.add_argument("--started", help=argparse.SUPPRESS)
    ap.add_argument("--failed-step", default="", help=argparse.SUPPRESS)
    ap.add_argument("--pipeline-log", default="", help=argparse.SUPPRESS)
    ap.add_argument("--run-args", default="", help=argparse.SUPPRESS)
    ap.add_argument("--report-step", default="", help=argparse.SUPPRESS)   # Sentry: "Pipeline step failed: <step>"
    ap.add_argument("-q", "--quiet", action="store_true")
    a = ap.parse_args(argv)
    if a.ui:
        from ui.server import main as ui_main
        return ui_main(root=a.root, config=a.config)
    logging.basicConfig(level=logging.WARNING if a.quiet else logging.INFO)
    from agent.util import utc_log_formatter
    for h in logging.getLogger().handlers:   # on AWS (CVA_COMPONENT) the console goes to CloudWatch: no traceback
        h.setFormatter(utc_log_formatter(path_free=bool(os.environ.get("CVA_COMPONENT"))))
    if os.environ.get("CVA_COMPONENT"):       # set on AWS only (/etc/compliance-agent/env, the Lambda handler)
        _error_reporting(os.environ["CVA_COMPONENT"], a.config, a.root)
    if a.extract_union_rates:
        from agent.config import load_config
        from agent.reference import extract_union_rates
        out, n, archived = extract_union_rates(a.extract_union_rates, load_config(a.config, a.root))
        print(f"{n} union rates written to {out}" + (f" (previous table kept as {archived})" if archived else ""))
        print("The workbook itself was only read; keep it off Google Drive and AWS.")
        return 0
    if a.test_email:
        from agent.notify import send_test
        print(send_test(a.root or "."))
        return 0
    if a.report_step:
        from agent.errors import flush, report_step_failure
        report_step_failure(a.report_step)
        flush()
        return 0
    if a.notify_pipeline:
        if a.failed_step and a.failed_step != "tokenize":     # the tokenizer not done is reported by email, not a bug
            from agent.errors import report_step_failure
            report_step_failure(a.failed_step)    # not for "lock" or "projects", which aren't bugs
        from agent.notify import notify_pipeline
        result = notify_pipeline(a.root, a.started, a.failed_step, args=a.run_args, pipeline_log=a.pipeline_log,
                                 trigger=a.trigger, tokenizer_failed=a.tokenizer_failed,
                                 run_projects=a.run_projects, empty_projects=a.empty_projects,
                                 unpulled_projects=a.unpulled_projects)
        print("Outcome email:", result)
        return 1 if str(result).startswith("not sent:") else 0    # 1 (failed, not "send_on" skipped): the short alert
    if a.check_model:
        from agent.config import load_config
        from agent.modelcheck import run_check
        r = run_check(load_config(a.config, a.root), a.model)
        print(f"\nModel {r['model']} (ID {r['model_id']}): {r['passed']}/{r['total']} checks passed")
        for x in r["results"]:
            print(f"  {'PASS' if x['passed'] else 'FAIL'}  {x['test']}: {x['result']}  ({x['seconds']} s)")
            if x["details"] and not x["passed"]:
                print(f"        {x['details']}")
        print(f"Report: {r['report']}")
        return 0 if r["all_passed"] else 1
    from agent.runner import parse_plan, parse_tasks, run
    try:
        parse_tasks(a.task)
        plan = parse_plan(a.plan)
    except ValueError as e:
        ap.error(str(e))
    if plan and (a.project or a.task != "all"):
        ap.error("--plan gives each project its tasks: use it without --task and --project")
    if a.ai_experiment:
        os.environ["CVA_AI_EXPERIMENT"] = str(a.ai_experiment)     # read by agent/ai_review.py
    report = run(a.task, a.project, a.force, a.config, a.root, trigger=a.trigger, notify=not a.no_email, plan=plan)
    print(json.dumps(console_summary(report), indent=2, default=str))
    return 0 if report.get("outcome") == "SUCCESS" else 1


if __name__ == "__main__":
    sys.exit(main())
