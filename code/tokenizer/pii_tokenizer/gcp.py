"""
Small clients for the three Google services the job uses: Google Drive, Cloud Storage and Secret Manager.

Standard library only (urllib): little code to audit.
On Cloud Run the access token comes from the metadata server, as the job's own service account, so there
is no key file anywhere. For a local test run, set GOOGLE_OAUTH_ACCESS_TOKEN instead.
"""
from __future__ import annotations

import json
import os
import random
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

DRIVE = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD = "https://www.googleapis.com/upload/drive/v3"
GCS = "https://storage.googleapis.com/storage/v1"
GCS_UPLOAD = "https://storage.googleapis.com/upload/storage/v1"
SECRETS = "https://secretmanager.googleapis.com/v1"
SCOPES = "https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/cloud-platform"
METADATA_TOKEN = ("http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token"
                  "?scopes=" + urllib.parse.quote(SCOPES, safe=","))

FOLDER_MIME = "application/vnd.google-apps.folder"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_BY_EXT = {".xlsx": XLSX_MIME, ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
               ".csv": "text/csv", ".txt": "text/plain", ".json": "application/json", ".md": "text/markdown",
               ".pdf": "application/pdf"}


class GoogleError(RuntimeError):
    def __init__(self, msg, status=None, rate_limited=False, quota=False):
        super().__init__(msg)
        self.status = status
        self.rate_limited = rate_limited      # Google said "slow down" and kept saying it through every retry
        self.quota = quota                    # a daily quota is used up: not retried, it doesn't clear in a minute


# Google's "slow down" answers: 429, and 403 with one of these reasons (Drive's "User rate limit exceeded"; David,
# Nov 12, 2026: a Run tasks click failed on it). A refused request did nothing, so retrying it is safe even for a
# request that creates something. Any other 403 is final: no access, and the daily quotas (dailyLimitExceeded,
# quotaExceeded), which don't clear in a minute (review F3).
RATE_LIMIT_REASONS = {"rateLimitExceeded", "userRateLimitExceeded"}
QUOTA_REASONS = {"dailyLimitExceeded", "quotaExceeded"}
RATE_LIMIT_WAITS = (2, 4, 8, 16, 32)          # seconds, plus up to 1 s at random: about a minute in all


def _reasons(body: bytes) -> set:
    try:
        err = json.loads(body.decode("utf-8", "replace")).get("error") or {}
    except (ValueError, AttributeError):
        return set()
    if not isinstance(err, dict):
        return set()
    return {e.get("reason") for e in err.get("errors") or [] if isinstance(e, dict)}


def is_rate_limit(code, body: bytes) -> bool:
    return code == 429 or (code == 403 and bool(_reasons(body) & RATE_LIMIT_REASONS))


def is_quota(code, body: bytes) -> bool:
    return code == 403 and bool(_reasons(body) & QUOTA_REASONS)


class Auth:
    def __init__(self, urlopen=urllib.request.urlopen):
        self._urlopen = urlopen
        self._token, self._expires = None, 0.0

    def token(self) -> str:
        if os.environ.get("GOOGLE_OAUTH_ACCESS_TOKEN"):
            return os.environ["GOOGLE_OAUTH_ACCESS_TOKEN"]
        if self._token and time.time() < self._expires - 60:
            return self._token
        req = urllib.request.Request(METADATA_TOKEN, headers={"Metadata-Flavor": "Google"})
        with self._urlopen(req, timeout=10) as r:
            data = json.loads(r.read())
        self._token, self._expires = data["access_token"], time.time() + int(data.get("expires_in", 300))
        return self._token


class _Client:
    def __init__(self, auth: Auth, urlopen=urllib.request.urlopen):
        self.auth, self._urlopen = auth, urlopen

    def call(self, method, url, params=None, body=None, headers=None, what="call Google", ok404=False,
             retry=True, rate_wait=True):
        """retry=False for requests that create something: a retry after a lost reply could create it twice.
        rate_wait=False: a "slow down" fails at once (rate_limited), for a caller that must not wait (the web app's
        refresh, which holds a lock: review F1)."""
        if params:
            url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
        h = {"Authorization": f"Bearer {self.auth.token()}", **(headers or {})}
        attempt = slowed = 0
        while True:
            req = urllib.request.Request(url, data=body, method=method, headers=h)
            try:
                with self._urlopen(req, timeout=120) as r:
                    return r.read()
            except urllib.error.HTTPError as e:
                if e.code == 404 and ok404:
                    return None
                try:
                    raw = e.read() or b""
                except Exception:                 # noqa: BLE001 - a body cut off: decided on the code alone (F4)
                    raw = b""
                if is_rate_limit(e.code, raw):
                    if rate_wait and slowed < len(RATE_LIMIT_WAITS):
                        time.sleep(RATE_LIMIT_WAITS[slowed] + random.random())
                        slowed += 1
                        continue
                    raise GoogleError(f"Could not {what}: Google is limiting requests (HTTP {e.code}, still after "
                                      f"{slowed} tries)", e.code, rate_limited=True) from None
                if attempt < 4 and retry and e.code in (500, 502, 503, 504):
                    time.sleep(2 ** attempt)
                    attempt += 1
                    continue
                detail = raw[:300].decode("utf-8", "replace")
                raise GoogleError(f"Could not {what}: HTTP {e.code} {detail}", e.code,
                                  quota=is_quota(e.code, raw)) from None
            except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
                if retry and attempt < 4:
                    time.sleep(2 ** attempt)
                    attempt += 1
                    continue
                raise GoogleError(f"Could not {what}: {getattr(e, 'reason', e)}") from None


# ------------------------------------------------------------------------------------------------
class Drive(_Client):
    FIELDS = "id,name,mimeType,md5Checksum,modifiedTime,parents,size,trashed"

    def get(self, file_id):
        raw = self.call("GET", f"{DRIVE}/files/{urllib.parse.quote(file_id)}",
                        {"fields": self.FIELDS, "supportsAllDrives": "true"}, what=f"open Drive item {file_id}")
        return json.loads(raw)

    def children(self, folder_id):
        if "'" in folder_id or "\\" in folder_id:
            raise GoogleError("Invalid Drive folder ID.")
        out, token = [], None
        while True:
            p = {"q": f"'{folder_id}' in parents and trashed = false", "fields": f"nextPageToken,files({self.FIELDS})",
                 "pageSize": "1000", "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"}
            if token:
                p["pageToken"] = token
            data = json.loads(self.call("GET", f"{DRIVE}/files", p, what="list a Drive folder"))
            out += data.get("files", [])
            token = data.get("nextPageToken")
            if not token:
                return out

    def tree(self, root_id):
        """Every file under root_id: (files {id: meta with 'rel'}, folders {rel dir: id}). Root rel dir is ''."""
        root = self.get(root_id)
        if root.get("mimeType") != FOLDER_MIME:
            raise GoogleError(f"Drive item {root_id} is not a folder (or is not shared with the tokenizer).")
        files, folders, todo = {}, {"": root_id}, [("", root_id)]
        while todo:
            rel_dir, fid = todo.pop()
            for f in self.children(fid):
                rel = f"{rel_dir}/{f['name']}" if rel_dir else f["name"]
                if f["mimeType"] == FOLDER_MIME:
                    folders[rel] = f["id"]
                    todo.append((rel, f["id"]))
                else:
                    files[f["id"]] = {**f, "rel": rel}
        return files, folders

    def download(self, f, rate_wait=True) -> bytes:
        fid = urllib.parse.quote(f["id"])
        if f.get("mimeType") == GSHEET_MIME:
            return self.call("GET", f"{DRIVE}/files/{fid}/export", {"mimeType": XLSX_MIME},
                             what=f"export Google Sheet '{f.get('name')}'", rate_wait=rate_wait)
        return self.call("GET", f"{DRIVE}/files/{fid}", {"alt": "media", "supportsAllDrives": "true"},
                         what=f"download '{f.get('name')}'", rate_wait=rate_wait)

    def create(self, parent_id, name, data: bytes, mime: str) -> dict:
        boundary = uuid.uuid4().hex
        meta = json.dumps({"name": name, "parents": [parent_id]}).encode()
        body = (f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n".encode() + meta +
                f"\r\n--{boundary}\r\nContent-Type: {mime}\r\n\r\n".encode() + data + f"\r\n--{boundary}--".encode())
        raw = self.call("POST", f"{DRIVE_UPLOAD}/files",
                        {"uploadType": "multipart", "supportsAllDrives": "true", "fields": self.FIELDS}, body,
                        {"Content-Type": f"multipart/related; boundary={boundary}"}, what=f"upload '{name}'",
                        retry=False)
        return json.loads(raw)

    def update(self, file_id, data: bytes, mime: str, name: str | None = None) -> dict:
        raw = self.call("PATCH", f"{DRIVE_UPLOAD}/files/{urllib.parse.quote(file_id)}",
                        {"uploadType": "media", "supportsAllDrives": "true", "fields": self.FIELDS}, data,
                        {"Content-Type": mime}, what=f"update Drive file {file_id}")
        meta = json.loads(raw)
        if name and meta.get("name") != name:
            meta = self.patch(file_id, {"name": name})
        return meta

    def patch(self, file_id, body: dict, params: dict | None = None) -> dict:
        raw = self.call("PATCH", f"{DRIVE}/files/{urllib.parse.quote(file_id)}",
                        {"supportsAllDrives": "true", "fields": self.FIELDS, **(params or {})},
                        json.dumps(body).encode(), {"Content-Type": "application/json"},
                        what=f"change Drive file {file_id}")
        return json.loads(raw)

    def mkdir(self, parent_id, name) -> str:
        raw = self.call("POST", f"{DRIVE}/files", {"supportsAllDrives": "true", "fields": "id"},
                        json.dumps({"name": name, "parents": [parent_id], "mimeType": FOLDER_MIME}).encode(),
                        {"Content-Type": "application/json"}, what=f"create folder '{name}'", retry=False)
        return json.loads(raw)["id"]

    def move(self, file_id, new_parent, old_parent, name=None) -> dict:
        return self.patch(file_id, {"name": name} if name else {},
                          {"addParents": new_parent, "removeParents": old_parent})

    def anyone_can_open(self, file_id) -> bool:
        """True if the file is shared with "anyone with the link" (read-only check)."""
        raw = self.call("GET", f"{DRIVE}/files/{urllib.parse.quote(file_id)}/permissions",
                        {"supportsAllDrives": "true", "fields": "permissions(type)"}, what=f"read sharing of {file_id}")
        return any(p.get("type") == "anyone" for p in json.loads(raw).get("permissions", []))

    def rename(self, file_id, name) -> dict:
        return self.patch(file_id, {"name": name})

    def trash(self, file_id) -> None:
        self.patch(file_id, {"trashed": True})


# ------------------------------------------------------------------------------------------------
class Bucket(_Client):
    """One private Cloud Storage bucket holding vault.json.enc, state.json and the run lock."""

    def __init__(self, name: str, auth: Auth, urlopen=urllib.request.urlopen):
        super().__init__(auth, urlopen)
        self.name = name.removeprefix("gs://").strip("/")

    def get(self, obj) -> tuple[bytes | None, int]:
        q = urllib.parse.quote(obj, safe="")
        meta = self.call("GET", f"{GCS}/b/{self.name}/o/{q}", what=f"read gs://{self.name}/{obj}", ok404=True)
        if meta is None:
            return None, 0
        gen = int(json.loads(meta)["generation"])
        data = self.call("GET", f"{GCS}/b/{self.name}/o/{q}", {"alt": "media", "generation": gen},
                         what=f"read gs://{self.name}/{obj}")
        return data, gen

    def put(self, obj, data: bytes, if_generation: int | None = None) -> int:
        p = {"uploadType": "media", "name": obj}
        if if_generation is not None:
            p["ifGenerationMatch"] = str(if_generation)
        raw = self.call("POST", f"{GCS_UPLOAD}/b/{self.name}/o", p, data,
                        {"Content-Type": "application/octet-stream"}, what=f"write gs://{self.name}/{obj}")
        return int(json.loads(raw)["generation"])

    def delete(self, obj) -> None:
        self.call("DELETE", f"{GCS}/b/{self.name}/o/{urllib.parse.quote(obj, safe='')}",
                  what=f"delete gs://{self.name}/{obj}", ok404=True)


def access_secret(name: str, auth: Auth, urlopen=urllib.request.urlopen) -> bytes:
    """name: projects/<project>/secrets/<secret>/versions/<version or latest>."""
    import base64
    if not name.startswith("projects/"):
        raise GoogleError("KEY_SECRET must look like projects/<project>/secrets/<name>/versions/latest")
    raw = _Client(auth, urlopen).call("GET", f"{SECRETS}/{name}:access", what=f"read secret {name}")
    return base64.b64decode(json.loads(raw)["payload"]["data"])
