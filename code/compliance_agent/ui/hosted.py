"""Hosted mode of the app: the same pages, served from Google Cloud Run behind Identity-Aware Proxy (IAP) at
https://cva.example (web/deploy_web.sh sets it up).

Differences from the desktop app, all decided here:
  * Who: every request must carry IAP's signed header, checked against Google's keys and this app's audience.
    IAP itself only lets in the accounts it was told to (the firm's Google Workspace), so the app never sees
    anyone else; the check here makes sure a request cannot reach the app any other way.
  * What it shows: the firm's shared drive, read directly with the app's own Google service account (no key
    file: the token comes from Cloud Run's metadata server). The readable data never leaves Google, as today.
    The drive is listed in one pass and mirrored into a local cache: file names, sizes and dates as empty
    placeholders, and the content only of what the pages read (Uploads, the agent's state file, run summaries).
  * What it changes: uploads, removals and new projects are written to the shared drive; and, for administrators
    only (Managers of the shared drive, asked of Drive itself), the Settings page's file (issue #151,
    agent/cva_filekinds/settings_file.py). Nothing else.
  * Runs: on AWS only, through the CVA-RunTask Session Manager document, which takes one checked request.
    AWS is reached with short-lived credentials for the role that trusts this service account
    (AssumeRoleWithWebIdentity with a Google-signed ID token): no AWS key is stored anywhere.

Settings (environment variables, set by web/deploy_web.sh):
  CVA_HOSTED=1, CVA_DRIVE_ID (the shared drive), CVA_ROOT_FOLDER_ID (the compliance folder staff use; the drive
  ID if it is the drive's top level), CVA_IAP_AUDIENCE, CVA_PUBLIC_ORIGIN (https://cva.example),
  CVA_AWS_ROLE_ARN, CVA_AWS_REGION, CVA_INSTANCE_NAME (default compliance-agent), CVA_CACHE_DIR (default /tmp/cva).
"""
from __future__ import annotations

import base64
import datetime as dt
import json
import os
import re
import shutil
import sys
import threading
import time
import urllib.parse
import urllib.request
from pathlib import Path

import logging

log = logging.getLogger("compliance")

HOSTED = os.environ.get("CVA_HOSTED") == "1"
CACHE = Path(os.environ.get("CVA_CACHE_DIR", "/tmp/cva"))
IAP_CERTS = "https://www.gstatic.com/iap/verify/public_key"
IAP_ISSUER = "https://cloud.google.com/iap"
METADATA_ID_TOKEN = ("http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/identity"
                     "?format=full&audience=")
AWS_AUDIENCE = "cva-aws"           # the audience of the ID token sent to AWS (the role trusts the account's ID)
FOLDER_MIME = "application/vnd.google-apps.folder"
# what the pages read the content of; everything else is mirrored as an empty placeholder of the right size
# Files whose contents the mirror holds (all others are empty placeholders of the right size and date): uploads, the
# agent's state, run summaries, each project's tracker (read only to tell whether the intent/affidavit check
# applies, ui/server.py intent_not_applicable), and the tokenizer's "Tracker sources.csv" (its Project status column
# greys out Closed and Not started projects: #151 step 3)
CONTENT_RE = re.compile(r"^(Projects/[^/]+/(Uploads/[^/]+|Tracker/[^/]+\.(?i:xlsx|xlsm)|\.agent_state\.json)"
                        r"|run_logs/[^/]+__run_summary\.json|CVA settings/cva_settings\.json"
                        r"|Tracker sources\.csv)$")      # each project's status (#151 step 3; ui/server.py)
ADMIN_ROLE = "organizer"              # Drive's name for a shared drive's Manager: the administrators (#151)
ADMIN_CACHE_SECONDS = 300
ADMIN_FAILURE_SECONDS = 60            # Drive couldn't say: closed, and not asked again for a minute (#151 review F7)
SKIP_RE = re.compile(r"(^|/)(Archived|_sync_backup|\.agent)(/|$)|(^|/)~\$")
REFRESH_SECONDS = 30
LIMITED_SECONDS = 60                  # after Google said "slow down" through every retry: fail at once meanwhile
RETRY_FAILED_SECONDS = 600            # a file that couldn't be downloaded (a Sheet too large to export) is retried
                                      # after this, or as soon as it changes


def _gcp():
    """The tokenizer's small Google client (standard library only), copied into the image as ui/_gcp.py."""
    try:
        from ui import _gcp as g                      # in the image
    except ImportError:                               # from source: code/tokenizer/pii_tokenizer/gcp.py
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tokenizer"))
        from pii_tokenizer import gcp as g
    return g


def setting(name, default=""):
    return os.environ.get(name, default).strip()


# ------------------------------------------------------------------------------------------------ who
class Unauthorized(Exception):
    pass


def verify_iap(headers, verifier=None) -> str:
    """The signed-in user's email, from IAP's signed header; Unauthorized if missing or not valid for this app."""
    token = headers.get("X-Goog-IAP-JWT-Assertion") or ""
    audience = setting("CVA_IAP_AUDIENCE")
    if not token or not audience:
        raise Unauthorized("no IAP assertion")
    if verifier is None:
        verifier = _verify_with_google_keys
    try:
        claims = verifier(token, audience)
    except Exception as e:                            # bad signature, expired, wrong audience
        raise Unauthorized(f"invalid IAP assertion: {e}") from None
    if claims.get("iss") != IAP_ISSUER or not claims.get("email"):
        raise Unauthorized("IAP assertion from an unexpected issuer")
    return str(claims["email"]).lower()


_KEYS = {"certs": None, "until": 0.0}


def _verify_with_google_keys(token, audience):
    """Checks the signature (ES256) against IAP's public keys, kept for an hour, and the audience and expiry."""
    from google.auth import jwt
    if not _KEYS["certs"] or time.time() > _KEYS["until"]:
        with urllib.request.urlopen(IAP_CERTS, timeout=10) as resp:
            _KEYS["certs"], _KEYS["until"] = json.loads(resp.read()), time.time() + 3600
    try:
        return jwt.decode(token, certs=_KEYS["certs"], audience=audience)
    except ValueError:
        _KEYS["until"] = 0.0                          # a rotated key: fetch the keys again next time
        raise


def origin_ok(headers) -> bool:
    """Changes must come from the app's own pages (same origin), like the desktop app's local check."""
    want = setting("CVA_PUBLIC_ORIGIN").rstrip("/").lower()
    origin = (headers.get("Origin") or "").rstrip("/").lower()
    return not origin or (bool(want) and origin == want)


# ------------------------------------------------------------------------------------------------ drive
def _rfc3339(s: str) -> float:
    return dt.datetime.fromisoformat(s.replace("Z", "+00:00")).timestamp()


class DriveView:
    """A local mirror of the shared drive's compliance folder, refreshed at most every REFRESH_SECONDS."""

    def __init__(self, drive=None, drive_id=None, root_id=None, cache=CACHE):
        self.drive_id = drive_id or setting("CVA_DRIVE_ID")
        self.root_id = root_id or setting("CVA_ROOT_FOLDER_ID") or self.drive_id
        self.cache = Path(cache)
        self._drive = drive
        self.lock = threading.Lock()
        self.last, self.ids, self.folders, self.md5 = 0.0, {}, {}, {}
        self.failed: dict[str, tuple] = {}    # path -> (fingerprint, retry after): a file that can't be downloaded
        self.limited_until = 0.0              # Google said "slow down" lately: don't queue another minute of retries
        self.settings_fid = None              # the settings file this app made or saved last (Drive's search can lag)
        self.current, self._gen = None, 0

    @property
    def drive(self):
        if self._drive is None:
            g = _gcp()
            self._drive = g.Drive(g.Auth())
        return self._drive

    @property
    def root(self) -> Path:
        """The current mirror (a versioned folder: a refresh never changes the files a request is reading)."""
        return self.current or (self.cache / "none")

    def _list_all(self) -> list[dict]:
        """Every item of the shared drive in one listing (a few calls), instead of one call per folder."""
        g = _gcp()
        out, token = [], None
        while True:
            p = {"corpora": "drive", "driveId": self.drive_id, "includeItemsFromAllDrives": "true",
                 "supportsAllDrives": "true", "q": "trashed = false", "pageSize": "1000",
                 "fields": "nextPageToken,files(id,name,mimeType,size,modifiedTime,md5Checksum,parents)"}
            if token:
                p["pageToken"] = token
            data = json.loads(self.drive.call("GET", f"{g.DRIVE}/files", p, what="list the shared drive"))
            out += data.get("files", [])
            token = data.get("nextPageToken")
            if not token:
                return out

    @staticmethod
    def _safe_name(name) -> bool:
        """A Drive name usable as one path part: not empty, '.', '..', and no separator (a Drive name may hold '/')."""
        return (isinstance(name, str) and name not in ("", ".", "..") and not any(c in name for c in "/\\\0")
                and len(name) <= 255)

    def refresh(self, force=False) -> Path:
        with self.lock:
            if not force and time.time() - self.last < REFRESH_SECONDS and self.current:
                return self.current
            g = _gcp()
            if time.time() < self.limited_until:      # review F2: requests waiting on the lock fail at once
                raise g.GoogleError("Could not list the shared drive: Google is limiting requests", 403,
                                    rate_limited=True)
            try:
                items = self._list_all()
            except g.GoogleError as e:
                if e.rate_limited:
                    self.limited_until = time.time() + LIMITED_SECONDS
                raise
            by_id = {i["id"]: i for i in items}
            paths: dict[str, str | None] = {self.root_id: ""}

            def rel(i, depth=0):
                if i["id"] in paths:
                    return paths[i["id"]]
                parent = (i.get("parents") or [None])[0]
                if depth > 40 or parent is None or (parent not in by_id and parent != self.root_id) \
                        or not self._safe_name(i.get("name")):
                    paths[i["id"]] = None
                    return None
                base = paths.get(parent) if parent == self.root_id else rel(by_id[parent], depth + 1)
                paths[i["id"]] = None if base is None else (f"{base}/{i['name']}" if base else i["name"])
                return paths[i["id"]]

            all_folders, folders, files, dupes = {"": self.root_id}, [""], {}, []
            for i in items:
                r = rel(i)
                if not r:
                    continue
                mime = i.get("mimeType", "")
                if mime == FOLDER_MIME:
                    all_folders.setdefault(r, i["id"])   # every folder, for writes (archives found, not duplicated)
                    if not SKIP_RE.search(r):
                        folders.append(r)
                    continue
                if mime.startswith("application/vnd.google-apps."):
                    if mime != g.GSHEET_MIME:
                        continue                          # Google Docs, Slides, shortcuts: nothing the agent reads
                    if not r.lower().endswith(".xlsx"):
                        r += ".xlsx"                      # a Google Sheet is read as the workbook it exports to
                if SKIP_RE.search(r):
                    continue
                if r in files:
                    dupes.append(r)                       # two files of the same name in one folder: first one wins
                    continue
                files[r] = i
            if dupes:                                     # never a path: it can hold a name
                log.warning("Shared drive: %d file name(s) used twice in one folder; the app shows one of each",
                            len(dupes))
            self._gen += 1
            new = self.cache / f"v{self._gen}"
            if new.exists():
                shutil.rmtree(new)
            for r in folders:
                (new / r).mkdir(parents=True, exist_ok=True)
            md5, problems, failed = {}, [], {}
            old_root = self.current
            limited = False                               # Google said "slow down": no more downloads this time
            for r, i in files.items():
                dest = new / r
                if not dest.resolve().is_relative_to(new.resolve()):
                    continue                              # never write outside the mirror
                dest.parent.mkdir(parents=True, exist_ok=True)
                mtime = _rfc3339(i["modifiedTime"]) if i.get("modifiedTime") else time.time()
                data, kept = None, False
                if CONTENT_RE.match(r):
                    old = old_root / r if old_root else None
                    # unchanged since the last refresh: its MD5, or for a Google Sheet (no MD5) its last change
                    fp = i.get("md5Checksum") or (f"modified {i['modifiedTime']}" if i.get("modifiedTime") else None)
                    same = fp and self.md5.get(r) == fp
                    gave_up = self.failed.get(r)
                    if same and old and old.is_file():
                        try:
                            os.link(old, dest)            # the same bytes, not a second copy in memory
                        except OSError:
                            shutil.copyfile(old, dest)
                        kept = True
                    elif gave_up and gave_up[0] == fp and time.time() < gave_up[1]:
                        problems.append(r)                # failed lately and unchanged: not tried again yet
                        failed[r] = gave_up
                    elif limited:
                        problems.append(r)                # tried again on the next refresh
                    else:
                        try:                              # never waits inside the lock (review F1)
                            data = self.drive.download(i, rate_wait=False) or b""
                            md5[r] = fp
                        except Exception as e:            # one unreadable file must not break the whole app
                            problems.append(r)
                            if getattr(e, "rate_limited", False):
                                limited = True            # not the file's fault: retried on the next refresh
                            else:
                                failed[r] = (fp, time.time() + RETRY_FAILED_SECONDS)
                    if same:
                        md5[r] = fp
                if kept:
                    pass
                elif data is not None:
                    dest.write_bytes(data)
                else:
                    with open(dest, "wb") as fh:          # placeholder: right size and date, no content
                        fh.truncate(int(i.get("size") or 0))
                os.utime(dest, (mtime, mtime))
            if problems:                                  # never a path or an error text: both can hold names
                log.warning("Shared drive: %d file(s) could not be read; the app shows them without contents",
                            len(problems))
            self.current = new
            for old in self.cache.glob("v*"):            # keep this mirror and the one before it
                if old.is_dir() and old not in (new, old_root):
                    shutil.rmtree(old, ignore_errors=True)
            self.folders, self.md5, self.ids = all_folders, md5, {r: i["id"] for r, i in files.items()}
            self.failed = failed
            self.last = time.time()
            return new

    # ---------------------------------------------------------------------------------------- writes
    def _folder(self, rel_dir: str) -> str:
        """The Drive ID of rel_dir under the compliance folder, creating missing folders."""
        if rel_dir in self.folders:
            return self.folders[rel_dir]
        parent = rel_dir.rsplit("/", 1)[0] if "/" in rel_dir else ""
        fid = self.drive.mkdir(self._folder(parent), rel_dir.rsplit("/", 1)[-1])
        self.folders[rel_dir] = fid
        return fid

    def _archive(self, project: str, rel_file: str, label: str, ts: str) -> str:
        """Move a file of the project into Files/Archived/<ts>/ as '<name> (<label>).<ext>'; returns its new path."""
        from agent.layout import ARCHIVED, FILES_DIR, labelled
        dest_dir = f"Projects/{project}/{FILES_DIR}/{ARCHIVED}/{ts}"
        name = labelled(rel_file.rsplit("/", 1)[-1], label)
        parent_old = self.folders[rel_file.rsplit("/", 1)[0]]
        self.drive.move(self.ids[rel_file], self._folder(dest_dir), parent_old, name=name)
        return f"{FILES_DIR}/{ARCHIVED}/{ts}/{name}"

    def upload(self, project: str, name: str, data: bytes, ts: str) -> str:
        from agent.layout import UPLOAD_DIR
        self.refresh(force=True)
        g = _gcp()
        rel = f"Projects/{project}/{UPLOAD_DIR}/{name}"
        note = ""
        with self.lock:
            if rel in self.ids:
                self._archive(project, rel, "replaced upload", ts)
                note = "an unprocessed file with the same name was already in Uploads; it was moved to the archive"
            mime = g.MIME_BY_EXT.get(Path(name).suffix.lower(), "application/octet-stream")
            self.drive.create(self._folder(f"Projects/{project}/{UPLOAD_DIR}"), name, data, mime)
            self.last = 0.0
        return note

    def remove(self, project: str, name: str, ts: str) -> str:
        from agent.layout import UPLOAD_DIR
        self.refresh(force=True)
        rel = f"Projects/{project}/{UPLOAD_DIR}/{name}"
        with self.lock:
            if rel not in self.ids:
                raise FileNotFoundError(name)
            dest = self._archive(project, rel, "removed", ts)
            self.last = 0.0
        return dest

    # ---------------------------------------------------------------------------------------- administrators (#151)
    def is_admin(self, email: str) -> bool:
        """Whether this person is a Manager of the shared drive (asked of Drive, kept 5 minutes). A Manager reached
        through a group doesn't count: the check is by person. On any error: not an administrator."""
        email = (email or "").lower()
        now = time.time()
        if not email:
            return False
        self._roles = getattr(self, "_roles", {})
        hit = self._roles.get(email)
        if hit and now - hit[1] < ADMIN_CACHE_SECONDS:
            return hit[0]
        if now - getattr(self, "_roles_failed", 0.0) < ADMIN_FAILURE_SECONDS:
            return False
        g = _gcp()
        try:
            roles, token = {}, None
            while True:
                p = {"supportsAllDrives": "true", "pageSize": "100",
                     "fields": "nextPageToken,permissions(role,type,emailAddress)"}
                if token:
                    p["pageToken"] = token
                data = json.loads(self.drive.call("GET", f"{g.DRIVE}/files/{urllib.parse.quote(self.drive_id)}"
                                                  "/permissions", p, what="read the shared drive's members"))
                for m in data.get("permissions", []):
                    if m.get("type") == "user" and m.get("emailAddress"):
                        roles[m["emailAddress"].lower()] = m.get("role")
                token = data.get("nextPageToken")
                if not token:
                    break
        except Exception:
            self._roles_failed = now
            log.warning("Couldn't read the shared drive's members; the Settings page is closed for now")
            return False
        self._roles = {e: (r == ADMIN_ROLE, now) for e, r in roles.items()}
        self._roles.setdefault(email, (False, now))          # not a member as a person: asked again in 5 minutes
        return self._roles[email][0]

    def _settings_file(self) -> dict | None:
        """The Settings page's file as Drive has it now: one small query in its folder, not a new listing of the
        whole drive (that took seconds: David, Nov 10). The folder's ID comes from the cached listing."""
        from agent.cva_filekinds import settings_file as SF
        folder_rel, name = SF.REL.rsplit("/", 1)
        self.refresh()                                   # cached for REFRESH_SECONDS: only the folder ID is used
        folder = self.folders.get(folder_rel)
        if folder is None:
            return None
        g = _gcp()
        q = f"'{folder}' in parents and name = '{name}' and trashed = false"
        data = json.loads(self.drive.call("GET", f"{g.DRIVE}/files", {
            "q": q, "corpora": "drive", "driveId": self.drive_id, "supportsAllDrives": "true",
            "includeItemsFromAllDrives": "true", "fields": "files(id,name,mimeType,modifiedTime)",
            "orderBy": "modifiedTime desc"}, what="find the settings file"))
        files = data.get("files") or []
        if not files and self.settings_fid:    # Drive's search can miss a file made seconds ago: ask for it by ID
            try:                               # (kept in memory: a restart in those seconds still loses it)
                f = self.drive.get(self.settings_fid)
            except g.GoogleError as e:
                if e.status != 404:            # busy, quota, 5xx: say so; never "no file", which would make a 2nd
                    raise
                f = None                       # deleted since
            if f and not f.get("trashed") and folder in (f.get("parents") or []) and f.get("name") == name:
                return f
            self.settings_fid = None
        return files[0] if files else None

    def read_settings(self) -> tuple[str | None, bytes | None]:
        """(its Drive ID, its bytes) as they are on the drive now; (None, None) if there is none yet."""
        f = self._settings_file()
        return (f["id"], self.drive.download(f)) if f else (None, None)

    def write_settings(self, data: bytes, fid: str | None) -> None:
        """Nov 12, 2026: there were two cva_settings.json. Drive's search can miss a file made seconds before, so a
        second save right after the first-ever one found none and made another. The app now keeps the ID of the
        file it made or saved, and reads and saves that one when the search finds nothing."""
        from agent.cva_filekinds import settings_file as SF
        folder, name = SF.REL.rsplit("/", 1)
        with self.lock:
            fid = fid or self.settings_fid      # two first-ever saves in one second: the later one wins (no 2nd file)
            if fid:
                self.drive.update(fid, data, "application/json")
            else:
                made = self.drive.create(self._folder(folder), name, data, "application/json") or {}
                fid = made.get("id")
            self.settings_fid = fid
            self.last = 0.0

    def new_project(self, name: str) -> None:
        from agent.layout import FILES_DIR, PAYROLL_DIR, TRACKER_DIR, UPLOAD_DIR
        self.refresh(force=True)
        with self.lock:
            if f"Projects/{name}" in self.folders:
                raise FileExistsError(name)
            for sub in (UPLOAD_DIR, PAYROLL_DIR, TRACKER_DIR, FILES_DIR):
                self._folder(f"Projects/{name}/{sub}")
            self.last = 0.0


VIEW = DriveView() if HOSTED else None


# ------------------------------------------------------------------------------------------------ runs on AWS
def _id_token(audience=AWS_AUDIENCE, urlopen=urllib.request.urlopen) -> str:
    req = urllib.request.Request(METADATA_ID_TOKEN + urllib.parse.quote(audience),
                                 headers={"Metadata-Flavor": "Google"})
    with urlopen(req, timeout=10) as r:
        return r.read().decode()


def aws_session(sts=None, token_fn=_id_token):
    """boto3 session with short-lived credentials of CVA_AWS_ROLE_ARN, signed in with this service account."""
    import boto3
    region = setting("CVA_AWS_REGION", "us-west-2")
    sts = sts or boto3.client("sts", region_name=region)
    r = sts.assume_role_with_web_identity(RoleArn=setting("CVA_AWS_ROLE_ARN"), RoleSessionName="cva-web-app",
                                          WebIdentityToken=token_fn(), DurationSeconds=900)
    c = r["Credentials"]
    return boto3.Session(aws_access_key_id=c["AccessKeyId"], aws_secret_access_key=c["SecretAccessKey"],
                         aws_session_token=c["SessionToken"], region_name=region)


def _server_checks():
    """deploy/run_request.py: the checks the server applies, applied here first so a refusal is shown at once."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("run_request", Path(__file__).resolve().parents[1] / "deploy" / "run_request.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run_request(tasks, projects, force, user, plan=None) -> str:
    """The base64 request the CVA-RunTask document takes (checked here, and again on the server). plan:
    [{"project", "tasks"}], each project its own tasks (one run, one email); tasks and projects are then not used."""
    trigger = re.sub(r"[^A-Za-z0-9 .,@()_:/-]", "", f"manual (web app, {user})")[:160]
    body = ({"plan": [{"project": i["project"], "tasks": list(i["tasks"])} for i in plan]} if plan else
            {"tasks": list(tasks), "projects": list(projects)})
    body.update({"force": bool(force), "trigger": trigger})
    req = base64.b64encode(json.dumps(body).encode()).decode()
    chk = _server_checks()
    try:
        chk.parse(req)
    except chk.BadRequest as e:
        raise ValueError(f"The run can't be sent: {e}.") from None
    return req


def project_name_ok(name: str) -> bool:
    """Project names the server accepts in a run request (letters, digits, spaces and . , & ( ) ' _ -)."""
    return bool(_server_checks().PROJECT_RE.fullmatch(name)) and ".." not in name


def start_on_aws(tasks, projects, force, user, session=None, wait_seconds=15, plan=None) -> str:
    """Start one run on the Compliance Validation Agent's server; returns the Session Manager command ID. plan: see
    run_request."""
    s = session or aws_session()
    name = setting("CVA_INSTANCE_NAME", "compliance-agent")
    res = s.client("ec2").describe_instances(Filters=[{"Name": "tag:Name", "Values": [name]},
                                                      {"Name": "instance-state-name", "Values": ["running"]}])
    ids = [i["InstanceId"] for r in res.get("Reservations", []) for i in r.get("Instances", [])]
    if not ids:
        raise RuntimeError("The Compliance Validation Agent is not running right now. Please tell the administrator.")
    ssm = s.client("ssm")
    r = ssm.send_command(InstanceIds=ids[:1], DocumentName="CVA-RunTask",
                         Parameters={"request": [run_request(tasks, projects, force, user, plan)]},
                         Comment="started from the web app")
    cid = r["Command"]["CommandId"]
    # a refused request or a run already in progress ends within seconds: report it instead of "started"
    for _ in range(wait_seconds // 3):
        time.sleep(3)
        try:
            inv = ssm.get_command_invocation(CommandId=cid, InstanceId=ids[0])
        except Exception:                             # not registered yet
            continue
        if inv.get("Status") in ("Failed", "Cancelled", "TimedOut"):
            # its output stays in AWS Systems Manager (Run Command, this command ID): it can name files (#158)
            log.warning("CVA-RunTask command %s ended %s", cid, inv.get("Status"))
            raise RuntimeError("The Compliance Validation Agent did not run it. Please try again; if it happens again, tell the "
                               "administrator, who can see why.")
        if inv.get("Status") == "Success":
            if "Another run is in progress" in (inv.get("StandardOutputContent") or ""):
                raise RuntimeError("Another run is in progress: try again when it has finished.")
            break
    return cid
