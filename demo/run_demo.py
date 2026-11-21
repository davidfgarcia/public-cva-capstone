"""Starts the CVA demo: invented data, a local stand-in for L&I, the CVA's own web UI on this computer.

    python demo/run_demo.py                          # build the demo folder if needed, then start (evidence-only AI)
    python demo/run_demo.py --ai claude              # the AI review also asks Claude (needs ANTHROPIC_API_KEY set)
    python demo/run_demo.py --reset                  # throw the demo folder away and build it again (re-record)
    python demo/run_demo.py --demo ~/Downloads/CVA_demo --port 8775 --no-browser

What it guarantees (DEMO.md has the details):
* **Invented data only**: demo/make_demo_data.py writes every file.
* **Kept apart from a real set-up**: the UI runs with HOME set to <demo>/.home, so it reads and writes its own
  settings there (never ~/.compliance_agent/settings.json), with the demo's config.yaml and notifications.config.
* **Nothing real is reached**: the only proxy is the local L&I stand-in (demo/lni_standin.py), which answers L&I's
  lookup with invented records and refuses every other request; no AWS, Google, Sentry or SES credentials are
  passed on (every AWS_*, GOOGLE_*, SENTRY_* ... variable is dropped). In ``--ai claude`` mode, and only then, the
  Claude API host goes around the proxy, with the key from this terminal's ANTHROPIC_API_KEY (never stored).
* **No email is sent**: the outcome emails are saved as .eml previews in <demo>/run_logs/.
"""
from __future__ import annotations

import argparse
import os
import signal
import socket
import subprocess
import sys
import time
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
AGENT = REPO / "code" / "compliance_agent"
sys.path.insert(0, str(HERE))

import make_demo_data as M                                    # noqa: E402
from lni_standin import DemoLookup                            # noqa: E402

DEFAULT_DEMO = Path.home() / "Downloads" / "CVA_demo"
DROP_ENV = ("AWS_", "SENTRY_", "ANTHROPIC_", "OPENAI_", "OLLAMA_", "GOOGLE_", "GCLOUD_", "CLOUDSDK_", "COMPLIANCE_",
            "SES_", "SMTP_", "SOURCE_REMOTE", "MIRROR_REMOTE", "CVA_", "VOYAGE_", "SOCRATA_")
CLAUDE_HOST = "api.anthropic.com"


def set_ai_mode(demo: Path, mode: str) -> None:
    """The demo config's AI review: 'evidence' = local mode (retrieval and citations, nothing sent anywhere);
    'claude' = Claude's suggestion too. Shadow mode is off in the demo so the Review decisions sheet shows it.
    The demo's payrolls are invented and not tokenized, so Claude mode allows untokenized payrolls here only."""
    f = demo / "config.yaml"
    head = "".join(line for line in f.read_text().splitlines(keepends=True) if line.startswith("#"))
    cfg = yaml.safe_load(f.read_text())
    ai = cfg.setdefault("ai_review", {})
    ai.update({"enabled": True, "shadow": False, "mode": "claude" if mode == "claude" else "local",
               "allow_untokenized": mode == "claude", "download_regulations": False})
    f.write_text(head + yaml.safe_dump(cfg, sort_keys=False))


def demo_env(demo: Path, proxy: str, ai: str) -> dict:
    env = {k: v for k, v in os.environ.items() if not k.upper().startswith(DROP_ENV)}
    for k in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        env[k] = proxy
    env.pop("NO_PROXY", None)
    env.pop("no_proxy", None)
    if ai == "claude":
        env["ANTHROPIC_API_KEY"] = os.environ["ANTHROPIC_API_KEY"]
        env["NO_PROXY"] = env["no_proxy"] = CLAUDE_HOST
    home = demo / ".home"
    home.mkdir(exist_ok=True)
    env.update({"HOME": str(home), "USERPROFILE": str(home),
                "COMPLIANCE_NOTIFY": "dry-run", "COMPLIANCE_NOTIFY_CONFIG": str(demo / "notifications.config"),
                "SENTRY_DSN": "off", "AWS_EC2_METADATA_DISABLED": "true",
                "AWS_CONFIG_FILE": os.devnull, "AWS_SHARED_CREDENTIALS_FILE": os.devnull,
                "PYTHONDONTWRITEBYTECODE": "1"})
    return env


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Start the CVA demo on invented data.")
    ap.add_argument("--demo", type=Path, default=DEFAULT_DEMO, help=f"the demo folder (default {DEFAULT_DEMO})")
    ap.add_argument("--ai", choices=("evidence", "claude"), default="evidence",
                    help="AI review: evidence only (default, offline) or with Claude's suggestion")
    ap.add_argument("--reset", action="store_true", help="rebuild the demo folder from scratch first")
    ap.add_argument("--port", type=int, default=8775, help="the web UI's port (default 8775: not the real CVA's 8765, so a browser never mixes the two)")
    ap.add_argument("--lni-port", type=int, default=8771, help="the L&I stand-in's port (default 8771)")
    ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args(argv)
    demo = a.demo.expanduser().resolve()
    if a.ai == "claude" and not os.environ.get("ANTHROPIC_API_KEY"):
        raise SystemExit("--ai claude needs ANTHROPIC_API_KEY set in this terminal (it is never written anywhere)")
    if a.reset:
        M.reset(demo)
    if not (demo / "demo_data" / "answer_key.json").is_file():
        if demo.exists() and any(demo.iterdir()):
            raise SystemExit(f"{demo} isn't a demo folder and isn't empty: choose another --demo")
        M.build(demo)
        print(f"Built the demo data in {demo}")
    for port in (a.port, a.lni_port):               # the UI would quietly move to another port: say so instead
        with socket.socket() as sk:
            if sk.connect_ex(("127.0.0.1", port)) == 0:
                raise SystemExit(f"port {port} is in use (another demo still running?): stop it, or use --port")
    set_ai_mode(demo, a.ai)
    lookup = DemoLookup(demo / "demo_data" / "lni_records.json", demo / "demo_data" / "lni_requests.json",
                        a.lni_port).start()
    env = demo_env(demo, lookup.proxy, a.ai)
    if a.no_browser:
        env["COMPLIANCE_UI_NO_BROWSER"] = "1"
    env["COMPLIANCE_UI_PORT"] = str(a.port)
    print(f"L&I stand-in on {lookup.proxy} (invented records). AI review: {a.ai}. Emails: previews in "
          f"{demo / 'run_logs'}. Stop with Ctrl+C.", flush=True)
    proc = subprocess.Popen([sys.executable, "run.py", "--ui", "--root", str(demo), "--config",
                             str(demo / "config.yaml")], cwd=AGENT, env=env)

    def stop(*_):
        signal.signal(signal.SIGTERM, signal.SIG_IGN)      # once: the clean-up below must not be interrupted
        raise KeyboardInterrupt

    signal.signal(signal.SIGTERM, stop)             # check_demo.py stops the demo this way: the UI goes with it
    try:
        return proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        try:
            proc.wait(10)
        except subprocess.TimeoutExpired:
            proc.kill()
        return 0
    finally:
        signal.signal(signal.SIGINT, signal.SIG_IGN)
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
        lookup.stop()


if __name__ == "__main__":
    sys.exit(main())
