#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
RECLAIMER_DISABLE_LIVE_WEB=1 exec "$APP_DIR/run_app.sh"
