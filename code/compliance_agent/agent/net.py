"""HTTP GET with retries for temporary failures only (dropped connection, timeout, 429 Too Many Requests, 5xx).
A permanent error (e.g. 400 Bad Request, 404 Not Found) is raised at once: asking again would get the same answer."""
from __future__ import annotations

import time

import requests

from .util import log

WAITS = (2, 5, 15)                    # seconds before the 2nd, 3rd and 4th attempt
TRANSIENT_STATUS = {429, 500, 502, 503, 504}


def post(url, json=None, headers=None, timeout=60, waits=WAITS, sleep=time.sleep, what="request"):
    """requests.post with the same retries as get() (L&I's live lookup answers POSTs only; agent/lni_live.py)."""
    return get(url, headers=headers, timeout=timeout, waits=waits, sleep=sleep, what=what, _json=json, _method="POST")


def get(url, params=None, headers=None, timeout=60, waits=WAITS, sleep=time.sleep, what="request", _json=None,
        _method="GET"):
    """requests.get with up to len(waits) retries for temporary failures. Returns the Response (status 2xx);
    raises the last error otherwise."""
    attempt = 0
    while True:
        try:
            if _method == "POST":
                r = requests.post(url, json=_json, headers=headers, timeout=timeout)
            else:
                r = requests.get(url, params=params, headers=headers, timeout=timeout)
            r.raise_for_status()
            return r
        except (requests.ConnectionError, requests.Timeout, requests.HTTPError) as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            transient = not isinstance(e, requests.HTTPError) or status in TRANSIENT_STATUS
            if not transient or attempt >= len(waits):
                raise
            wait = waits[attempt]
            if status == 429:                 # honour the server's Retry-After when it gives one (capped)
                try:
                    wait = min(60, max(wait, int(e.response.headers.get("Retry-After", wait))))
                except (TypeError, ValueError):
                    pass
            log.info("%s failed (%s); retrying in %d s", what, status or type(e).__name__, wait,
                     extra={"details": {"url": url, "attempt": attempt + 1}})
            sleep(wait)
            attempt += 1
