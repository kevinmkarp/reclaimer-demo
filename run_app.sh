#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BUNDLED_PY="/Users/kevinkarp/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [[ -x "$BUNDLED_PY" ]]; then
  PYTHON="$BUNDLED_PY"
else
  PYTHON="${PYTHON:-python3}"
fi

if [[ ! -d "$APP_DIR/.deps/streamlit" || ! -d "$APP_DIR/.deps/openai" ]]; then
  "$PYTHON" -m pip install --target "$APP_DIR/.deps" -r "$APP_DIR/requirements.txt"
fi

mkdir -p "$APP_DIR/.home"

HOME="$APP_DIR/.home" \
PYTHONPATH="$APP_DIR/.deps${PYTHONPATH:+:$PYTHONPATH}" \
  "$PYTHON" -m streamlit run "$APP_DIR/app.py" \
  --global.developmentMode false \
  --browser.gatherUsageStats false \
  --server.port "${PORT:-8501}" \
  --server.headless true
