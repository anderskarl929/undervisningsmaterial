#!/usr/bin/env bash
# Wrapper för classroom.py.
# Använder .venv/ om den finns (krävs för PDF/DOCX/ODT-konvertering),
# annars systemets python3 (Google Docs-only-mode).
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="$HERE/.venv/bin/python"
[ -x "$PY" ] || PY="python3"
exec "$PY" "$HERE/classroom.py" "$@"
