"""Changes worth telling the people who read the outcome emails about.

ENTRIES: for the administrator's email (a new reader, script or tool, or a change in what the agent does); each
appears in its "What's new" part for SHOW_DAYS days after its date. STAFF: for the staff email, written for people
who know payroll rules but not the system (no settings, servers or file formats); each appears once, in the next
staff email after it ships, then never again. Add entries whenever such a change ships (newest first)."""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

SHOW_DAYS = 14

ENTRIES = []          # the firm's change history is not in the public copy


STAFF: list[tuple[str, str, str]] = []

STAFF_SENT = Path(".agent") / "whats_new_sent.json"     # STAFF entries already in a staff email (never synced)


def _key(entry) -> str:
    return f"{entry[0]}|{entry[1]}"


def staff_unsent(root) -> list[tuple[str, str, str]]:
    """STAFF entries not yet in a staff email. With no record yet (a new server or computer), none: what shipped
    before is not news there."""
    try:
        seen = set(json.loads((Path(root) / STAFF_SENT).read_text()))
    except (OSError, ValueError, TypeError):
        return []
    return [e for e in STAFF if _key(e) not in seen]


def staff_remember(root) -> None:
    """After a staff email went out: every current STAFF entry counts as told."""
    path = Path(root) / STAFF_SENT
    try:
        seen = set(json.loads(path.read_text()))
    except (OSError, ValueError, TypeError):
        seen = set()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(sorted(seen | {_key(e) for e in STAFF})))
    except OSError:
        pass


def recent(today: dt.date | None = None, days: int = SHOW_DAYS) -> list[tuple[str, str, str]]:
    today = today or dt.datetime.now(dt.timezone.utc).date()
    out = []
    for d, title, text in ENTRIES:
        age = (today - dt.date.fromisoformat(d)).days
        if 0 <= age <= days:
            out.append((d, title, text))
    return out
