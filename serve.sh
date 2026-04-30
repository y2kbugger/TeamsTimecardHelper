#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

HOST="${HOST:-localhost}"
PORT="${PORT:-8999}"

exec python3 -m http.server "$PORT" --bind "$HOST"