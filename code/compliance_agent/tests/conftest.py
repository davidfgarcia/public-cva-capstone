"""Shared test setup: no test reaches a real service.

Since Nov 10, 2026 config.yaml reads L&I's live lookup (intent.lni_source: live), which answers POSTs. In this process
a real POST fails at once (tests give the client a fake `post`, or patch requests.post themselves). Tests that start
the agent as a separate process (the pipeline, run.py) inherit a proxy that refuses every connection, so they can't
reach L&I, data.wa.gov or anything else either (found on Nov 10: those runs were calling L&I's website, also on CI)."""
import os

import pytest
import requests

NO_NETWORK = "http://127.0.0.1:9"       # nothing listens there: every connection is refused at once


@pytest.fixture(autouse=True, scope="session")
def _no_network_for_child_processes():
    keys = ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy", "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy")
    old = {k: os.environ.get(k) for k in keys}
    os.environ.update({"HTTPS_PROXY": NO_NETWORK, "HTTP_PROXY": NO_NETWORK, "https_proxy": NO_NETWORK,
                       "http_proxy": NO_NETWORK, "ALL_PROXY": NO_NETWORK, "all_proxy": NO_NETWORK,
                       "NO_PROXY": "127.0.0.1,localhost", "no_proxy": "127.0.0.1,localhost"})
    yield
    for k, v in old.items():
        if v is None:
            os.environ.pop(k, None)
        else:
            os.environ[k] = v


@pytest.fixture(autouse=True)
def _no_real_posts(monkeypatch):
    def refuse(*a, **k):
        raise requests.ConnectionError("tests never POST to a real service")
    monkeypatch.setattr(requests, "post", refuse)
