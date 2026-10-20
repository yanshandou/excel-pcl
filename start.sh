#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1

fail() {
    echo "Setup failed. Install Python 3.10+ and check your network."
    read -rp "Press Enter to exit..."
    exit 1
}

find_venv_python() {
    if [ -f .venv/Scripts/python.exe ]; then
        VENV_PY=.venv/Scripts/python.exe
    elif [ -x .venv/bin/python ]; then
        VENV_PY=.venv/bin/python
    else
        VENV_PY=
    fi
}

find_venv_python
if [ -z "$VENV_PY" ]; then
    if command -v py >/dev/null 2>&1; then
        py -3 -m venv .venv || fail
    elif command -v python3 >/dev/null 2>&1; then
        python3 -m venv .venv || fail
    else
        python -m venv .venv || fail
    fi
    find_venv_python
    [ -n "$VENV_PY" ] || fail
fi

"$VENV_PY" -m pip install -r requirements.txt || fail
"$VENV_PY" app.py "$@"
