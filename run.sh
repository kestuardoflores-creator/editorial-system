#!/usr/bin/env bash
set -e
BASE="$(cd "$(dirname "$0")" && pwd)"
VENV="$BASE/.venv"
[ -d "$VENV" ] || python3 -m venv "$VENV"
"$VENV/bin/pip" install -q -r "$BASE/app/assembler/requirements.txt"
exec "$VENV/bin/python" "$BASE/app/launcher.py"
