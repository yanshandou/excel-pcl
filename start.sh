#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1

fail() {
    echo "Setup failed. Install Python 3.10+ and check your network."
    read -rp "Press Enter to exit..."
    exit 1
}

if command -v py >/dev/null 2>&1; then
    py -3 -m venv .venv || fail
elif command -v python3 >/dev/null 2>&1; then
    python3 -m venv .venv || fail
else
    python -m venv .venv || fail
fi

if [ -x .venv/Scripts/python ] || [ -x .venv/Scripts/python.exe ]; then
    VENV_PY=.venv/Scripts/python
else
    VENV_PY=.venv/bin/python
fi

"$VENV_PY" -m pip install -r requirements.txt || fail
"$VENV_PY" app.py
