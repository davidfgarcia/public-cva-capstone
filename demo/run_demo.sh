#!/usr/bin/env bash
# Starts the CVA demo on invented data (see DEMO.md). Any option goes to demo/run_demo.py:
#   bash demo/run_demo.sh [--reset] [--ai evidence|claude] [--demo FOLDER] [--port 8775]
set -euo pipefail
cd "$(dirname "$0")/.."
PY="${PYTHON:-python3}"
if [ -x .venv/bin/python ]; then PY=.venv/bin/python; fi
exec "$PY" demo/run_demo.py "$@"
