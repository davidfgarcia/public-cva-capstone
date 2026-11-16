"""
Google-side tokenizer: entry point.

On Google Cloud (one Cloud Run job per folder pair; started by the AWS job that uses the folder, right before
it reads and right after it writes, and once a night by Cloud Scheduler as a backstop):

    python main.py run        keep the raw folder and its tokenized copy in step (see pii_tokenizer/mirror.py);
                              the folder pair comes from TOKENIZER_MIRROR, or --mirror, else all pairs
    python main.py check      check access to the key, the bucket and every folder; changes nothing

On a computer, for trying it on a file (needs the key in TOKEN_KEY, or --key-file):

    python main.py tokenize   payroll.xlsx payroll.tokenized.xlsx --vault vault.json.enc
    python main.py detokenize results.xlsx results.readable.xlsx --vault vault.json.enc
    python main.py new-key    print a new random key (pipe it straight into Secret Manager, never into a chat)

Exit code 0 = done, 1 = failed (Cloud Run marks the execution as failed and the alert policy emails you).
"""
from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import re
import secrets
import sys
import time
from pathlib import Path

from pii_tokenizer import config as config_mod
from pii_tokenizer import gcp
from pii_tokenizer.files import Context, detokenize_file, tokenize_file
from pii_tokenizer.mirror import dumps_state, run_all
from pii_tokenizer.tokens import Keyring, TokenError, Vault, load_key

VAULT_OBJ = "vault.json.enc"
LOCK_STALE_SECONDS = 35 * 60         # longer than the job's 30-minute time limit
LOCK_WAIT_SECONDS = 12 * 60          # a run started by AWS waits for one already going, rather than skipping

log = logging.getLogger("pii_tokenizer")


def _local_key(args) -> bytes:
    if getattr(args, "key_file", None):
        return load_key(Path(args.key_file).read_bytes())
    if os.environ.get("TOKEN_KEY"):
        return load_key(os.environ["TOKEN_KEY"])
    raise SystemExit("Set TOKEN_KEY (base64 or hex) or pass --key-file.")


# ------------------------------------------------------------------------------------------------
def cmd_run(cfg, mirror: str | None) -> int:
    """One run for one folder pair (TOKENIZER_MIRROR, set on each Cloud Run job), or for all of them.

    Each folder pair has its own job, state file and lock, so the AWS job for one folder never waits on the
    other. The vault is shared: a save that races with another job's save is merged, never overwritten."""
    if not cfg.state_bucket or not cfg.key_secret:
        raise SystemExit("STATE_BUCKET and KEY_SECRET must be set.")
    names = [m.name for m in cfg.mirrors]
    if mirror and mirror not in names:
        raise SystemExit(f"TOKENIZER_MIRROR={mirror!r} is not one of the mirrors in the config: {', '.join(names)}")
    label = mirror or "all"
    state_obj, lock_obj = f"state/{label}.json", f"locks/{label}.lock"
    auth = gcp.Auth()
    bucket = gcp.Bucket(cfg.state_bucket, auth)
    if not _take_lock(bucket, lock_obj):
        raise RuntimeError(f"another run for '{label}' has held the lock for over {LOCK_WAIT_SECONDS // 60} minutes")
    try:
        key = load_key(gcp.access_secret(cfg.key_secret, auth))
        kr = Keyring(key)
        vblob, vgen = bucket.get(VAULT_OBJ)
        vault = Vault.from_bytes(vblob, kr)
        sblob, sgen = bucket.get(state_obj)
        state = json.loads(sblob) if sblob else {}
        ctx = Context.build(cfg, key, vault)
        gens = {"vault": vgen, "state": sgen}

        def save_vault():
            for _ in range(6):
                try:
                    gens["vault"] = bucket.put(VAULT_OBJ, vault.to_bytes(kr), if_generation=gens["vault"])
                    vault.dirty = False
                    return
                except gcp.GoogleError as e:
                    if e.status != 412:
                        raise
                # The other folder's job saved the vault meanwhile: take in its new tokens and try again.
                blob, gens["vault"] = bucket.get(VAULT_OBJ)
                vault.merge(Vault.from_bytes(blob, kr))
            raise RuntimeError("could not save the token vault (it kept changing)")

        last = {"t": 0.0}

        def save_state():
            wait = 1.1 - (time.time() - last["t"])     # Cloud Storage allows about one write a second per object
            if wait > 0:
                time.sleep(wait)
            gens["state"] = bucket.put(state_obj, dumps_state(state), if_generation=gens["state"])
            last["t"] = time.time()

        lines = run_all(gcp.Drive(auth), ctx, state, save_vault, save_state, only=mirror)
        if vault.dirty:
            save_vault()
        for line in report_counts(lines):     # counts only: the lines name staff files (#21)
            print(line)
        print(f"Done ({label}): {len(lines)} changes reported, vault holds {len(vault.entries)} tokens.")
        return 0
    finally:
        bucket.delete(lock_obj)


def report_counts(lines: list[str]) -> list[str]:
    """The run's report for the job's log (Cloud Logging): how many changes of each kind, per folder pair, and no
    file path or reason (#21). The report lines name staff files and folders, which can hold a worker's name; the
    log follows the rule of the per-file progress lines (mirror._progress: Drive ID, type and size only). The files
    held back, with why, are listed for staff in the raw folder's Held back from AWS.csv.

    A line is "[<pair>] <kind>: <file>: <detail>" or "[<pair>] warning: <text>" (mirror.RunReport.lines); only
    the pair and the kind are kept. A line of another shape is counted as "other"."""
    counts: dict[tuple[str, str], int] = {}
    for line in lines:
        m = re.match(r"\[([^\]\n]{1,80})\] ([^:\n]{1,60}):", line)
        key = (m.group(1), m.group(2).strip()) if m else ("", "other")
        counts[key] = counts.get(key, 0) + 1
    return [f"[{pair}] {kind}: {n}" if pair else f"{kind}: {n}" for (pair, kind), n in counts.items()]


def _take_lock(bucket, lock_obj, sleep=time.sleep) -> bool:
    deadline = time.time() + LOCK_WAIT_SECONDS
    while True:
        body = json.dumps({"started": time.time(), "execution": os.environ.get("CLOUD_RUN_EXECUTION", "")}).encode()
        try:
            bucket.put(lock_obj, body, if_generation=0)
            return True
        except gcp.GoogleError as e:
            if e.status != 412:
                raise
        data, gen = bucket.get(lock_obj)
        if data is None:
            continue                                # released between our two calls
        if time.time() - json.loads(data).get("started", 0) > LOCK_STALE_SECONDS:
            try:                                    # left behind by a crashed run
                bucket.put(lock_obj, body, if_generation=gen)
                return True
            except gcp.GoogleError:
                pass
        if time.time() > deadline:
            return False
        sleep(15)


def cmd_check(cfg) -> int:
    auth, ok = gcp.Auth(), True

    def step(label, fn):
        nonlocal ok
        try:
            print(f"OK    {label}: {fn()}")
        except Exception as e:  # noqa: BLE001 - report every problem, not just the first
            ok = False
            print(f"FAIL  {label}: {what(e)}")           # the job's log: never the error's text (#158)

    step("key secret", lambda: f"{len(load_key(gcp.access_secret(cfg.key_secret, auth))) * 8}-bit key")
    bucket = gcp.Bucket(cfg.state_bucket, auth)
    step("state bucket", lambda: "vault present" if bucket.get(VAULT_OBJ)[0] else "reachable (no vault yet)")
    drive = gcp.Drive(auth)
    for m in cfg.mirrors:
        step(f"{m.name}: raw folder", lambda m=m: drive.get(m.raw_folder_id)["name"])
        step(f"{m.name}: tokenized folder", lambda m=m: drive.get(m.tokenized_folder_id)["name"])
    return 0 if ok else 1


def cmd_file(cfg, args, direction) -> int:
    key = _local_key(args)
    kr = Keyring(key)
    vpath = Path(args.vault)
    vault = Vault.from_bytes(vpath.read_bytes() if vpath.exists() else None, kr)
    data = Path(args.input).read_bytes()
    if direction == "tokenize":
        ctx = Context.build(cfg, key, vault)
        name, out, detail = tokenize_file(Path(args.input).name, data, ctx)
        if out is None:
            print(detail)
            return 1
        if vault.dirty:
            vpath.write_bytes(vault.to_bytes(kr))
    else:
        out, rep = detokenize_file(Path(args.input).name, data, vault)
        detail = rep.summary()
    Path(args.output).write_bytes(out)
    print(f"{args.output}: {detail}")
    return 0


def _error_reporting(config_path=None) -> bool:
    """Sentry, privacy first (pii_tokenizer/errors.py: error type and code location only, never data).
    DSN from SENTRY_DSN or error_reporting.sentry_dsn in tokenizer_config.yaml."""
    from pii_tokenizer import errors
    import yaml
    try:
        path = Path(config_path or os.environ.get("TOKENIZER_CONFIG") or config_mod.DEFAULT_PATH)
        raw = yaml.safe_load(path.read_text()) or {}
    except Exception:                          # the config is loaded (and checked) again right after
        raw = {}
    return errors.setup("tokenizer", errors.dsn_from_config(raw), errors.environment_from_config(raw),
                        errors.release_name())


class _PathFreeFormatter(logging.Formatter):
    """A traceback keeps the code locations and the error's type, never its text, which can name a staff file
    (`download '<name>'`): the job's console goes to Cloud Logging (#158)."""

    def format(self, record):
        saved = record.exc_text                 # logging caches the traceback text on the record: not this one's
        record.exc_text = None
        try:
            return super().format(record)
        finally:
            record.exc_text = saved

    def formatException(self, ei):
        import traceback
        frames = "".join(traceback.format_tb(ei[2])) if ei and ei[2] else ""
        return f"Traceback (most recent call last):\n{frames}{ei[0].__name__ if ei and ei[0] else 'error'} (text left out)"

    def formatStack(self, stack_info):
        return ""


def what(e) -> str:
    """An error as the job's log may show it: its type and HTTP status, never its text (#158)."""
    status = getattr(e, "status", None)
    return type(e).__name__ + (f" (HTTP {status})" if isinstance(status, int) else "")


def setup_logging():
    """The job's console (Cloud Logging): tracebacks without the error's text (#158)."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    for h in logging.getLogger().handlers:
        h.setFormatter(_PathFreeFormatter("%(levelname)s %(message)s"))


def main(argv=None) -> int:
    setup_logging()
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", help="tokenizer_config.yaml to use (default: the one next to this file)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p_run = sub.add_parser("run")
    p_run.add_argument("--mirror", help="only this folder pair (default: TOKENIZER_MIRROR, else all)")
    sub.add_parser("check")
    sub.add_parser("new-key")
    for name in ("tokenize", "detokenize"):
        p = sub.add_parser(name)
        p.add_argument("input")
        p.add_argument("output")
        p.add_argument("--vault", default="vault.json.enc", help="local encrypted vault file")
        p.add_argument("--key-file")
    args = ap.parse_args(argv)
    if args.cmd == "new-key":
        sys.stdout.write(base64.b64encode(secrets.token_bytes(32)).decode())
        return 0
    if args.cmd in ("run", "check") and os.environ.get("CLOUD_RUN_JOB"):   # set by Cloud Run on its jobs only
        _error_reporting(args.config)
    cfg = config_mod.load(args.config)
    try:
        if args.cmd == "run":
            return cmd_run(cfg, args.mirror or os.environ.get("TOKENIZER_MIRROR") or None)
        if args.cmd == "check":
            return cmd_check(cfg)
        return cmd_file(cfg, args, args.cmd)
    except Exception as e:                          # noqa: BLE001 - any error: never its text in the job's log
        if args.cmd in ("run", "check"):            # the job's log (Cloud Logging): type and status only (#158)
            print(f"FAILED: {what(e)}", file=sys.stderr)
        else:                                       # a file command run by hand on a staff computer: the whole text
            print(f"FAILED: {e}", file=sys.stderr)
        log.error("tokenizer %s failed", args.cmd, exc_info=True)   # Sentry gets the error type and code location
        return 1


if __name__ == "__main__":
    sys.exit(main())
