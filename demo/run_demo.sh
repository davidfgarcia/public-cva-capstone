#!/usr/bin/env bash
# Starts the CVA demo on invented data (see DEMO.md). Any option goes to demo/run_demo.py:
#   bash demo/run_demo.sh [--reset] [--ai evidence|claude] [--demo FOLDER] [--port 8775]
# The first time, it makes a private Python environment in .venv/ and installs the requirements there.
set -euo pipefail
cd "$(dirname "$0")/.."
if [ ! -x .venv/bin/python ]; then
  PY="${PYTHON:-python3}"
  if ! "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
    echo "The demo needs Python 3.10 or newer (3.12 recommended); '$PY' is $("$PY" --version 2>&1)." >&2
    echo "Install a newer Python, or run: PYTHON=/path/to/python3.12 bash demo/run_demo.sh" >&2
    exit 1
  fi
  echo "First run: setting up .venv/ and installing the requirements (a minute or two)..."
  "$PY" -m venv .venv
  .venv/bin/python -m pip install --quiet --upgrade pip
  .venv/bin/python -m pip install --quiet -r code/compliance_agent/requirements.txt \
      -r code/tokenizer/requirements.txt pytest
fi
exec .venv/bin/python demo/run_demo.py "$@"
