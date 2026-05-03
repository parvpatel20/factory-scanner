#!/usr/bin/env bash
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$DIR"

# Ensure .env exists
if [[ ! -f .env ]]; then
  echo "ERROR: .env file not found."
  echo "       Create one by copying .env.example and filling in GROQ_API_KEY."
  exit 1
fi

# Install / update dependencies silently
echo "==> Checking dependencies..."
pip3 install -r requirements.txt -q --disable-pip-version-check

# Use gunicorn in production; fallback to Flask dev server if gunicorn unavailable
PORT="${PORT:-5050}"
WORKERS="${WEB_CONCURRENCY:-2}"

if command -v gunicorn &>/dev/null; then
  echo "==> Starting Factory Scanner on http://0.0.0.0:${PORT} (gunicorn, ${WORKERS} workers)"
  exec gunicorn \
    --bind "0.0.0.0:${PORT}" \
    --workers "${WORKERS}" \
    --timeout 150 \
    --access-logfile - \
    --error-logfile - \
    --log-level info \
    server:app
else
  echo "==> gunicorn not found — starting with Flask dev server on http://0.0.0.0:${PORT}"
  exec python3 server.py
fi
