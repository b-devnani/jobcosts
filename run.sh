#!/usr/bin/env bash
# Start the Job Cost Projection web tool.
#
#   ADMIN_PASSWORD=mysecret ./run.sh           # custom admin password
#   PORT=9000 ./run.sh                         # custom port
set -euo pipefail

cd "$(dirname "$0")"

PORT="${PORT:-8000}"
HOST="${HOST:-127.0.0.1}"

if ! python3 -c "import fastapi" >/dev/null 2>&1; then
  echo "Installing dependencies…"
  pip3 install -r requirements.txt
fi

echo "Job Cost Projection tool -> http://${HOST}:${PORT}"
echo "Admin password: ${ADMIN_PASSWORD:-admin (default — set ADMIN_PASSWORD to change)}"
exec python3 -m uvicorn backend.app:app --host "$HOST" --port "$PORT"
