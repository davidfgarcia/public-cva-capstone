"""A local stand-in for L&I's live lookup, serving the demo's invented records (demo_data/lni_records.json).

It is the evaluation harness's stand-in (code/compliance_agent/tools/eval_lni.py FakeLookup) with the demo's records:
the same answers in the same shapes, on 127.0.0.1 only. The demo runs the CVA with this stand-in as its **only
proxy** (HTTP_PROXY / HTTPS_PROXY), so:

* the intent/affidavit check's requests to L&I (the production address in plain http, demo config.yaml) reach this
  stand-in, never L&I;
* every other request (https CONNECT, any other host) is refused here, so the demo can't reach anything real;
* every request is recorded, in full, in demo_data/lni_requests.json (rewritten after each one).

    python demo/lni_standin.py <demo folder> [--port 8771]       # run_demo.sh starts it
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
AGENT = HERE.parent / "code" / "compliance_agent"
sys.path.insert(0, str(AGENT / "tools"))

import eval_lni as LNI                                       # noqa: E402


class DemoLookup(LNI.FakeLookup):
    def __init__(self, records: Path, log: Path = None, port: int = 0):
        super().__init__()
        data = json.loads(records.read_text())
        self.intents = {int(k): v for k, v in data["intents"].items()}
        self.affidavits = {int(k): v for k, v in data["affidavits"].items()}
        self.trees = {int(k): v for k, v in data["trees"].items()}
        self.log, self.port = log, port

    def _record(self, rec):
        super()._record(rec)
        if self.log:
            with self._lock:
                self.log.write_text(json.dumps(self.requests, indent=1))

    def start(self):
        if not self.port:
            return super().start()
        from http.server import ThreadingHTTPServer
        # the parent builds its handler inside start(); bind it to the fixed port by starting it, then moving it
        super().start()
        handler = self._server.RequestHandlerClass
        self.stop()
        import threading
        self._server = ThreadingHTTPServer(("127.0.0.1", self.port), handler)
        threading.Thread(target=self._server.serve_forever, daemon=True).start()
        self.proxy = f"http://127.0.0.1:{self.port}"
        return self


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Serve the demo's invented L&I records on 127.0.0.1.")
    ap.add_argument("demo", type=Path)
    ap.add_argument("--port", type=int, default=8771)
    a = ap.parse_args(argv)
    data = a.demo.expanduser().resolve() / "demo_data"
    lookup = DemoLookup(data / "lni_records.json", data / "lni_requests.json", a.port).start()
    print(f"L&I stand-in (invented records) on {lookup.proxy}; requests recorded in {data / 'lni_requests.json'}",
          flush=True)
    try:
        while True:
            time.sleep(3600)
    except KeyboardInterrupt:
        lookup.stop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
