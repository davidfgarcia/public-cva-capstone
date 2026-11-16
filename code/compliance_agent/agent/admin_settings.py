"""The Settings page's choices (issue #151), as the agent uses them.

The web app's Settings page (administrators only: Managers of the compliance shared drive) writes
`CVA settings/cva_settings.json` on the drive; the tokenizer copies its checked values to AWS (never its history), and
each run on the server fetches that copy into `.agent/cva_settings.json` (deploy/run_pipeline.sh). What the file may
hold, and the checks, are in cva_filekinds/settings_file.py (shared with the tokenizer and the web app).

What the agent does with it:
  * intent.lni_source and the other config.yaml settings in CONFIG_KEYS replace config.yaml's (load_config).
  * email.send_on, email.admin.send_on and email.subject_prefix replace notifications.config's.
  * The four recipient lists replace notifications.config's and the server's environment (COMPLIANCE_NOTIFY_TO,
    COMPLIANCE_NOTIFY_CC, COMPLIANCE_NOTIFY_ADMIN_TO): once a list is set on the page, the page is the only place it
    is set. A list the page hasn't set keeps its value from before. If the server's environment still sets a list the
    page now sets, the administrator's email says so (remove NOTIFY_TO / NOTIFY_ADMIN_TO from AWS_CONFIG).
  * The administrator's email names the page's settings in effect, when and by whom they were last changed, and warns
    when the file was edited outside the page or couldn't be used (then the values from before the page apply).
  * When the page changes a recipient list, the next run's email goes to the new list, and a short notice goes to the
    old and the new addresses of that audience (staff lists, administrator lists: one notice each; the web app can't
    send email). Only the server announces; what it last sent to, and which version of the file, are kept on the
    server only, in .agent/settings_seen.json (never synced: deploy/rclone-filter.txt). The notice is remembered as
    sent only once every notice went out.
No value is ever logged: only which settings are set.
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path

from .cva_filekinds import settings_file as SF

log = logging.getLogger("compliance")
SEEN = Path(".agent") / "settings_seen.json"
NOTIFY_KEYS = {"email.staff.to": "to", "email.staff.cc": "cc", "email.admin.to": "admin_to",
               "email.admin.cc": "admin_cc"}
LIST_NAMES = {"to": "Operational emails (staff): To", "cc": "Operational emails (staff): Cc",
              "admin_to": "Technical emails (administrator): To", "admin_cc": "Technical emails (administrator): Cc"}


def on_server() -> bool:
    """A run of the AWS pipeline (it sets SOURCE_REMOTE): the only one that announces recipient changes."""
    return bool(os.environ.get("SOURCE_REMOTE"))


def load(root) -> tuple[dict, dict | None, str | None]:
    """(values, info, problem) from the server's copy, or else the drive's own file (the desktop app). No file:
    ({}, None, None). A file that can't be used: ({}, None, <plain words>), and the values from before apply."""
    if not root:
        return {}, None, None
    for rel in (SF.SERVER_REL, SF.REL):
        p = Path(root) / rel
        if p.is_file():
            try:
                values, info = SF.read(p.read_bytes(), server_copy=rel == SF.SERVER_REL)
                return values, info, None
            except Exception:                        # any problem with the file: never a crash (#151 review F2)
                log.warning("The Settings page's file couldn't be used; the settings from before the page apply")
                return {}, None, "the Settings page's file couldn't be used, so the settings from before the page apply"
    seen = _read_seen(root)
    if seen and seen.get("page_present"):             # there was one last time (#151 review F5, N6)
        return {}, None, ("the Settings page's file wasn't there this run (removed, or held back by the tokenizer), "
                          "so the settings from before the page apply")
    return {}, None, None


# The page's settings that replace config.yaml's: {setting: (section path, key, how the value is stored there)}
CONFIG_KEYS = {
    "intent.lni_source": (("intent",), "lni_source", str),
    "intent.lni_max_age_days": (("intent",), "lni_max_age_days", int),
    "ai_review.enabled": (("ai_review",), "enabled", lambda v: "auto" if v == "on" else False),
    "wages.sam_gov.enabled": (("wages", "sam_gov"), "enabled", lambda v: v == "on"),
    "wages.union_fallback": (("wages",), "union_fallback", lambda v: v == "on"),
    "intent.not_started_statuses": (("intent",), "not_started_statuses", list),
    "intent.work_started_statuses": (("intent",), "work_started_statuses", list),
    "intent.completed_statuses": (("intent",), "completed_statuses", list),
    "intent.skip_statuses": (("intent",), "skip_statuses", list),
    "intent.lni_ignore_statuses": (("intent",), "lni_ignore_statuses", list),
}
NOTIFY_OPTIONS = {"email.send_on": "send_on", "email.admin.send_on": "admin_send_on",
                  "email.subject_prefix": "subject_prefix"}


def _page_form(key, v):
    """config.yaml's value of a page setting, in the page's form (the inverse of CONFIG_KEYS), or None."""
    kind = SF.SETTINGS[key]["kind"]
    if key == "ai_review.enabled":                    # exactly as agent/ai_review.py settings() reads it (N8)
        return "off" if v in (False, "false", "off", "no") else "on"
    if kind == "choice" and isinstance(v, bool):
        return "on" if v else "off"
    if kind == "choice":
        return v if isinstance(v, str) and v in SF.SETTINGS[key]["choices"] else None   # review N3
    if kind == "number":
        return v if type(v) is int else None
    if kind == "words":
        return [str(w).lower() for w in v] if isinstance(v, list) else None
    return None


def server_values(config_path=None) -> dict:
    """{setting: the value it has without the page}, from config.yaml as deployed (never with the page's file
    applied), in the page's form: for the Settings page to show next to each field (#151 review F3: a list typed
    on the page replaces the server's whole list). config.yaml's settings only: the email options and lists are in
    the server's notifications settings, which the web app can't see. Nothing (or a value missing) on any problem."""
    import yaml
    from .config import AGENT_DIR
    try:
        with open(config_path or os.environ.get("COMPLIANCE_CONFIG") or AGENT_DIR / "config.yaml") as f:
            cfg = yaml.safe_load(f) or {}
    except Exception:                                   # noqa: BLE001 (the page opens without them)
        return {}
    out = {}
    for key, (path, field, _conv) in CONFIG_KEYS.items():
        section = cfg
        for part in path:
            section = section.get(part) if isinstance(section, dict) else None
        try:
            v = _page_form(key, section[field]) if isinstance(section, dict) and field in section else None
        except Exception:                               # noqa: BLE001 (an odd value: shown as not known)
            v = None
        if v is not None:
            out[key] = v
    return out


def apply_config(cfg) -> None:
    values, _, _ = load(cfg.get("root"))
    for key, (path, field, conv) in CONFIG_KEYS.items():
        if key in values:
            section = cfg
            for part in path:
                section = section.setdefault(part, {})
            section[field] = conv(values[key])


def apply_notify(ncfg, root) -> dict:
    """The page's recipient lists in place of the ones from before; ncfg["page"] says what was used."""
    values, info, problem = load(root)
    ncfg.setdefault("admin_cc", [])
    for key, field in NOTIFY_KEYS.items():
        if key in values:
            ncfg[field] = list(values[key])
    for key, field in NOTIFY_OPTIONS.items():
        if key in values:
            ncfg[field] = values[key]
    ncfg["page"] = {"values": values, "info": info, "problem": problem}
    return ncfg


def summary_lines(ncfg) -> list[str]:
    """For the administrator's email: the page's settings in effect, and anything to look at."""
    page = ncfg.get("page") or {}
    values, info, problem = page.get("values") or {}, page.get("info"), page.get("problem")
    refused = [f"WARNING: the email service refused {', '.join(ncfg['refused'])} (not verified: Amazon SES is "
               "still in test mode); the email went to the others. Verify the address in SES (us-east-2, Verified "
               "identities), or remove it from its list (on the Settings page, or in the server's settings)."] if ncfg.get("refused") else []
    if problem:
        return refused + [f"Settings page: {problem}."]
    if not values:
        return refused
    when = (info or {}).get("changed_at", "").replace("T", " ").replace("Z", " UTC")
    by = (info or {}).get("changed_by") or "unknown"
    out = [f"Settings page (last changed {when or 'at an unknown time'} by {by}):"]
    for key, spec in SF.SETTINGS.items():
        if key not in values:
            continue
        v = values[key]
        shown = (spec["choices"].get(v, v) if spec["kind"] == "choice" else
                 (", ".join(v) or ("nobody" if spec["kind"] == "emails" else "(none)")) if isinstance(v, list) else
                 (repr(v) if v == "" else str(v)))
        if key == "intent.lni_max_age_days" and \
                (values.get("intent.lni_source") or server_values().get("intent.lni_source")) == "live":
            shown += " (not used: the L&I source is the website)"          # review F8
        out.append(f"  {spec['label']}: {shown}")
    out = refused + out
    if (info or {}).get("edited_outside_page"):
        out.append("  WARNING: the settings file was changed outside the Settings page (in Drive). Its values are "
                   "used; open the Settings page and save to confirm them.")
    return out


def recipient_notices(ncfg, root) -> list[tuple[str, str, str, list]]:
    """[(audience, subject, text, recipients)]: one notice per audience (staff lists, administrator lists) whose
    lists the page changed since the last emails. Only on the server (one record of what was announced), only for
    lists the page sets, and only when the page's file has changed since then (#151 review F5, F13, F14). The first
    time: nothing (only remembered)."""
    page = ncfg.get("page") or {}
    info, values = page.get("info"), page.get("values") or {}
    seen = _read_seen(root)
    if not on_server() or not info or seen is None or info.get("changed_at") == seen.get("page_changed_at"):
        return []
    before = seen.get("lists") or {}
    out = []
    for audience, fields in (("staff", ("to", "cc")), ("administrator", ("admin_to", "admin_cc"))):
        page_set = set(values) | set(seen.get("page_keys") or [])      # set now, or set last time (review N4)
        changes = [(LIST_NAMES[f], list(before.get(f) or []), list(ncfg.get(f) or [])) for f in fields
                   if _key(f) in page_set and sorted(before.get(f) or []) != sorted(ncfg.get(f) or [])]
        if changes:
            subject, text, everyone = change_notice(changes, info)
            out.append((audience, subject, text, everyone))
    return out


def remember_recipients(ncfg, root) -> None:
    """What the emails went to, and which version of the page's file: on the server only."""
    if not root or not on_server():
        return
    info = (ncfg.get("page") or {}).get("info")
    seen = _read_seen(root) or {}
    page = ncfg.get("page") or {}
    record = {"lists": _lists(ncfg), "page_present": bool(info),
              "page_keys": sorted(page.get("values") or {}) if info else seen.get("page_keys", []),
              "page_changed_at": info.get("changed_at") if info else seen.get("page_changed_at", "")}
    p = Path(root) / SEEN
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(record, indent=1, sort_keys=True))
    except OSError:
        log.warning("Couldn't remember the email recipients; a change may be announced again")


def change_notice(changes, info) -> tuple[str, str, list]:
    """(subject, text, recipients): to everyone on the old and the new lists that changed."""
    by = (info or {}).get("changed_by") or "an administrator"
    lines = [f"The Compliance Validation Agent's email recipients were changed on its Settings page by {by}.", ""]
    everyone = []
    for name, old, new in changes:
        added, removed = [a for a in new if a not in old], [a for a in old if a not in new]
        lines.append(f"{name}:")
        if added:
            lines.append(f"  added: {', '.join(added)}")
        if removed:
            lines.append(f"  removed: {', '.join(removed)}")
        lines.append(f"  now: {', '.join(new) or 'nobody'}")
        everyone += [a for a in dict.fromkeys(old + new) if a not in everyone]
    lines += ["", "If this wasn't expected, tell the administrator."]
    return "Email recipients changed", "\n".join(lines), everyone


def _key(field):
    return next(k for k, f in NOTIFY_KEYS.items() if f == field)


def _lists(ncfg) -> dict:
    return {f: list(ncfg.get(f) or []) for f in LIST_NAMES}


def _read_seen(root):
    if not root:
        return None
    try:
        data = json.loads((Path(root) / SEEN).read_text())
        return data if isinstance(data, dict) else None
    except (OSError, ValueError):
        return None
