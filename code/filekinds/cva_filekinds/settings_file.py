"""The administrator's settings file (issue #151): what it may hold, and what of it crosses to AWS.

WHERE AND WHO
    `CVA settings/cva_settings.json` at the top of the compliance shared drive. Only the web app's Settings page
    writes it, for an administrator (a Manager of the shared drive). Every member of the drive can read it: it holds
    no worker data (choices, numbers, tracker status words and a few email addresses).

WHAT IT HOLDS (the whole list: anything else is refused)
    intent.lni_source      "live" (L&I website) or "data.wa.gov" (L&I public data); see agent/lni_live.py
    intent.lni_max_age_days  0-14: how old L&I's public data may be (only with data.wa.gov)
    ai_review.enabled / wages.sam_gov.enabled / wages.union_fallback   "on" or "off" (Nov 12, 2026, #151 step 2)
    intent.*_statuses      tracker words: up to 20 words or short phrases (up to 3 words each), letters, spaces and
                           hyphens only; [] (stored when the page's entry is "(none)") is an empty list on purpose
    email.send_on          "all" or "problems": the staff email (#151 review F2)
    email.admin.send_on    "all" or "problems": the technical email (David, Nov 12, 2026: a quick way to stop
                           emails piling up while something is broken; the notices and the short alert always go)
    email.subject_prefix   up to 40 letters, spaces and []-: (no digits and no dots: no ID, number or web address
                           can be stored there and reach every subject; #151 review F1)
    email.staff.to/cc      operational emails (to staff): addresses, one per line on the page
    email.admin.to/cc      technical emails (to the administrator)
    Addresses must be on the firm's domain, or one of ALLOWED_ADDRESSES (changed by a reviewed PR only).
    A setting that isn't in the file keeps the value it has without the page (config.yaml, notifications.config and
    the server's environment).

THE FILE (written by the page)
    {"version": 1, "rev": 3, "values": {...}, "changed_at": "2026-11-10T17:40:00Z",
     "changed_by": "<administrator's email>", "checksum": "<sha256 of the values>",
     "history": [{"at", "by", "setting", "old", "new"}, ...]}
    rev counts the saves: the page refuses a save made from an older view of the file.
    The checksum tells a file edited by hand (in Drive) from one the page wrote; it is not a secret, so it catches
    mistakes, not someone set on hiding a change.

WHAT CROSSES TO AWS (the tokenizer, pii_tokenizer/mirror.py, calls for_aws)
    The values (checked again here: the known settings only, each of its type), changed_at, changed_by and whether the
    file was edited outside the page. Never the history. A file that fails a check is held back whole.
"""
from __future__ import annotations

import hashlib
import json
import re
import time

REL = "CVA settings/cva_settings.json"          # on the shared drive (relative to the compliance folder's top)
SERVER_REL = ".agent/cva_settings.json"         # the server's copy, outside what the server pushes back
VERSION = 1
FIRM_DOMAINS = ("firm.example", "firm-alt.example")   # both the firm's (David, Nov 11, 2026)
ALLOWED_ADDRESSES = ("admin@firm.example",)        # David's own, for the technical emails (Nov 10, 2026)
MAX_ADDRESSES = 10
MAX_HISTORY = 200
EMAIL = re.compile(r"^[a-z0-9._%+-]{1,64}@[a-z0-9-]+(\.[a-z0-9-]+)+$", re.ASCII)
MAX_WORDS, MAX_WORD = 20, 30
MAX_WORDS_IN_ENTRY = 3                          # a status is a word or a short phrase, not a sentence (review F5)
NONE_ENTRY = "(none)"                           # on the page: this list empty on purpose (review F6)
# Tracker words: letters (any language), spaces, hyphens; no digits, so no ID or number can be stored as a "word"
WORD = re.compile(r"[^\W\d_](?:[^\W\d_]| |-){0,29}")
TEXT = re.compile(r"[A-Za-z \[\]\-:]*", re.ASCII)      # no digits, no dots (#151 review F1)
STAMP = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z", re.ASCII)
# On record (#151 review F9): an address's local part is free text, so someone editing the file by hand could put
# words in it (up to 10 addresses x 4 lists). Only firm addresses (or ALLOWED_ADDRESSES) pass, and a hand edit is
# flagged in the administrator's email.

SETTINGS = {
    "intent.lni_source": {
        "label": "L&I source", "kind": "choice",
        "choices": {"live": "L&I website (live)", "data.wa.gov": "L&I public data (data.wa.gov)"},
        "help": "Where the intent/affidavit check reads L&I's records."},
    "intent.lni_max_age_days": {
        "label": "L&I public data: how old it may be (days)", "kind": "number", "min": 0, "max": 14,
        "help": "Only when the L&I source is the public data: older than this, the intent check is skipped and the "
                "emails say why. 0: skip only when the data can't be reached."},
    "ai_review.enabled": {"label": "AI review of the orange checks", "kind": "choice",
                          "choices": {"on": "On", "off": "Off"},
                          "help": "Off: no AI suggestions at all (they are only counted in your email for now)."},
    "wages.sam_gov.enabled": {"label": "Download federal wage determinations", "kind": "choice",
                              "choices": {"on": "On", "off": "Off"},
                              "help": "Off: a federal determination listed for a project isn't downloaded when the "
                                      "project doesn't have it."},
    "wages.union_fallback": {"label": "Union rates from the reference table", "kind": "choice",
                             "choices": {"on": "On", "off": "Off"},
                             "help": "Off: a project without union rates of its own reports \"no rate found\" for "
                                     "union lines."},
    "intent.not_started_statuses": {"label": "Tracker status: not started", "kind": "words",
                                    "help": "Statuses meaning the contractor hasn't started work (a status that "
                                            "contains one of these words). No names."},
    "intent.work_started_statuses": {"label": "Tracker status: work started", "kind": "words",
                                     "help": "Statuses meaning work has started, so an intent is due (the whole "
                                             "status, exactly). No names."},
    "intent.completed_statuses": {"label": "Tracker status: completed", "kind": "words",
                                  "help": "Statuses meaning the contractor's work is done (the whole status, "
                                          "exactly). Such a line with no affidavit anywhere yet is only counted in "
                                          "the tracker's Audit sheet, not marked: filing can take long. No names."},
    "intent.skip_statuses": {"label": "Tracker status: lines not checked", "kind": "words",
                             "help": "Lines with one of these statuses (the whole status, exactly) aren't "
                                     "checked at all. No names."},
    "intent.lni_ignore_statuses": {"label": "L&I statuses to ignore", "kind": "words",
                                   "help": "L&I intents whose status has a word starting with one of these aren't "
                                           "reported as missing from the tracker. No names."},
    "email.send_on": {"label": "Send the staff email", "kind": "choice",
                      "choices": {"all": "After every run", "problems": "Only when something went wrong"},
                      "help": "\"Something went wrong\": the run didn't start, failed or partly failed (a task "
                              "couldn't do its job), or a payroll couldn't be fully checked. New red or orange marks "
                              "alone don't count."},
    "email.admin.send_on": {"label": "Send the technical email", "kind": "choice",
                            "choices": {"all": "After every run", "problems": "Only when something went wrong"},
                            "help": "As for the staff email. The notices about recipient changes, and the short "
                                    "alert when an email can't be sent, always go."},
    "email.subject_prefix": {"label": "Email subject: start", "kind": "text", "max": 40,
                             "help": "Letters, spaces and []-: only (no digits, no dots). No names."},
    "email.staff.to": {"label": "Operational emails (staff): To", "kind": "emails",
                       "help": "Who gets the staff email after each run."},
    "email.staff.cc": {"label": "Operational emails (staff): Cc", "kind": "emails", "help": ""},
    "email.admin.to": {"label": "Technical emails (administrator): To", "kind": "emails",
                       "help": "Who gets the full email with every detail, warnings and errors."},
    "email.admin.cc": {"label": "Technical emails (administrator): Cc", "kind": "emails", "help": ""},
}


class Invalid(ValueError):
    """A value or a file the settings may not hold; the message is plain words for the page (no value repeated)."""


def address_allowed(a: str) -> bool:
    return a in ALLOWED_ADDRESSES or a.rsplit("@", 1)[-1] in FIRM_DOMAINS


def clean_value(key, value, policy=True):
    """The value as stored (addresses lower-cased, without repeats); Invalid if it isn't one the setting may take.
    policy=False checks the shape only (used by tests; the page, the tokenizer and the agent all apply the policy)."""
    spec = SETTINGS.get(key)
    if spec is None:
        raise Invalid("not a setting the page manages")
    if spec["kind"] == "choice":
        if not isinstance(value, str) or value not in spec["choices"]:
            raise Invalid(f"{spec['label']}: not one of the choices")
        return value
    if spec["kind"] == "number":
        if isinstance(value, str) and re.fullmatch(r"[0-9]{1,3}", value.strip(), re.ASCII):
            value = int(value.strip())
        if type(value) is not int or not spec["min"] <= value <= spec["max"]:
            raise Invalid(f"{spec['label']}: a whole number from {spec['min']} to {spec['max']}")
        return value
    if spec["kind"] == "text":
        if not isinstance(value, str) or not TEXT.fullmatch(value) or len(value) > spec["max"]:
            raise Invalid(f"{spec['label']}: up to {spec['max']} letters, spaces and []-: only (no digits, no dots)")
        return value.strip()
    if spec["kind"] == "words":
        if isinstance(value, str):
            value = value.split("\n")
        if not isinstance(value, list) or not all(isinstance(w, str) for w in value):
            raise Invalid(f"{spec['label']}: a list of words")
        if [w.strip().lower() for w in value if w.strip()] == [NONE_ENTRY]:
            return []                                 # empty on purpose (review F6)
        out = []
        for w in (" ".join(w.split()).lower() for w in value):
            if not w:
                continue
            if not WORD.fullmatch(w) or len(w.split()) > MAX_WORDS_IN_ENTRY:
                raise Invalid(f"{spec['label']}: one entry isn't a word or short phrase (up to {MAX_WORDS_IN_ENTRY} "
                              f"words: letters, spaces and hyphens, up to {MAX_WORD} characters)")
            if w not in out:
                out.append(w)
        if len(out) > MAX_WORDS:
            raise Invalid(f"{spec['label']}: at most {MAX_WORDS} entries")
        return out
    if isinstance(value, str):
        value = re.split(r"[\s,;]+", value)
    if not isinstance(value, list) or not all(isinstance(a, str) for a in value):
        raise Invalid(f"{spec['label']}: a list of email addresses")
    out = []
    for a in (a.strip().lower() for a in value):
        if not a:
            continue
        if not EMAIL.match(a) or len(a) > 254:
            raise Invalid(f"{spec['label']}: one of the entries isn't an email address")
        if policy and not address_allowed(a):
            raise Invalid(f"{spec['label']}: {a} is not on one of the firm's domains ({', '.join(FIRM_DOMAINS)})")
        if a not in out:
            out.append(a)
    if len(out) > MAX_ADDRESSES:
        raise Invalid(f"{spec['label']}: at most {MAX_ADDRESSES} addresses")
    return out


def clean_values(values, policy=True) -> dict:
    if not isinstance(values, dict):
        raise Invalid("the settings aren't a list of values")
    return {k: clean_value(k, v, policy) for k, v in values.items()}


def checksum(values: dict) -> str:
    return hashlib.sha256(json.dumps(values, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _stamp(at: str) -> str:
    """The time as written by the page, or "" (a real date and time, ASCII digits only: #151 review F9)."""
    if not STAMP.fullmatch(at):
        return ""
    try:
        time.strptime(at, "%Y-%m-%dT%H:%M:%SZ")
    except ValueError:
        return ""
    return at


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def parse(data: bytes) -> dict:
    """The file's document; Invalid if it isn't one (any problem: never another error)."""
    try:
        doc = json.loads(data.decode("utf-8"))
    except Exception:
        raise Invalid("the settings file can't be read") from None
    if not isinstance(doc, dict) or type(doc.get("version")) is not int or doc["version"] != VERSION:
        raise Invalid("the settings file isn't one the page wrote (version)")
    if not isinstance(doc.get("values", {}), dict) or not isinstance(doc.get("history", []), list):
        raise Invalid("the settings file isn't one the page wrote (layout)")
    return doc


def page_values(doc: dict | None) -> tuple[dict, list]:
    """(the values the page can show, [labels of the ones it can't use]): for the page, which must open even when a
    hand edit left a bad value (#151 review F12)."""
    good, bad = {}, []
    for k, v in ((doc or {}).get("values") or {}).items():
        try:
            good[k] = clean_value(k, v)
        except Invalid:
            bad.append(SETTINGS.get(k, {}).get("label", "a setting the page doesn't know"))
    return good, bad


def history(doc: dict | None) -> list:
    return [h for h in ((doc or {}).get("history") or []) if isinstance(h, dict)]


def edited_outside_page(doc: dict) -> bool:
    return doc.get("checksum") != checksum(doc.get("values") or {})


def updated(doc: dict | None, changes: dict, by: str, at: str | None = None) -> dict:
    """A new document: `changes` ({setting: value, or None to go back to the value without the page}) checked with
    the address policy and applied; one history entry per setting that changed. Invalid if a value isn't allowed."""
    doc = doc or {"version": VERSION, "values": {}, "history": []}
    at = at or _now()
    values = dict(doc.get("values") or {})
    hist = history(doc)
    # a value already in the file that can't be used (a hand edit) is removed by any save, never kept and confirmed
    # (#151 review F4); the page lists such values before saving
    for key in list(values):
        try:
            clean_value(key, values[key])
        except Invalid:
            values.pop(key)
            hist.append({"at": at, "by": by, "setting": key if key in SETTINGS else "(unknown)",
                         "old": "(couldn't be used)", "new": None})
    for key, value in changes.items():
        new = None if value is None else clean_value(key, value)
        old = values.get(key)
        if new == old:
            continue
        if new is None:
            values.pop(key, None)
        else:
            values[key] = new
        hist.append({"at": at, "by": by, "setting": key, "old": old, "new": new})
    clean_values(values)                              # (all checked above: never fails here)
    for email in ("staff", "admin"):                  # the same address on To and Cc of one email (David, Nov 11)
        if f"email.{email}.to" not in changes and f"email.{email}.cc" not in changes:
            continue                                  # only the lists being saved (review L3)
        both = set(values.get(f"email.{email}.to") or []) & set(values.get(f"email.{email}.cc") or [])
        if both:
            raise Invalid(f"{SETTINGS[f'email.{email}.cc']['label']}: {', '.join(sorted(both))} is already on To")
    rev = doc.get("rev") if type(doc.get("rev")) is int else 0
    return {"version": VERSION, "rev": rev + 1, "values": values, "changed_at": at, "changed_by": by,
            "checksum": checksum(values), "history": hist[-MAX_HISTORY:]}


def for_aws(data: bytes) -> bytes:
    """What the tokenizer copies to AWS: the checked values, when and by whom, and whether the file was edited outside
    the page. Invalid (the file is held back) if anything in the values isn't a known setting of its type."""
    doc = parse(data)
    values = clean_values(doc.get("values") or {})       # the address policy too: a hand-edited file can't add one
    by = str(doc.get("changed_by") or "").lower()
    at = str(doc.get("changed_at") or "")
    out = {"version": VERSION, "values": values,
           "changed_at": _stamp(at),
           "changed_by": by if EMAIL.match(by) and address_allowed(by) else "",
           "edited_outside_page": edited_outside_page(doc)}
    return json.dumps(out, indent=1, sort_keys=True).encode()


def read(data: bytes, server_copy: bool = False) -> tuple[dict, dict]:
    """(values, info) from the page's file, or (server_copy) the server's copy rebuilt by the tokenizer, for the agent.
    The values are checked as the page would; info: changed_at, changed_by, edited_outside_page."""
    doc = parse(data)
    values = clean_values(doc.get("values") or {})
    # the server's copy (rebuilt by the tokenizer: no checksum) says it; the page's own file is checked here
    # (#151 review F10: a key added by hand can't hide a hand edit)
    if server_copy:
        edited = doc.get("edited_outside_page") is not False
    else:
        edited = edited_outside_page(doc)
    at, by = str(doc.get("changed_at") or ""), str(doc.get("changed_by") or "").lower()
    return values, {"changed_at": _stamp(at),
                    "changed_by": by if EMAIL.match(by) and address_allowed(by) else "",
                    "edited_outside_page": bool(edited)}
