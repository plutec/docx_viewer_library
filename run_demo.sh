#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if ! command -v python >/dev/null 2>&1; then
  echo "Python is not available in PATH."
  exit 1
fi

if ! python -c "import fastapi, jinja2, multipart, uvicorn" >/dev/null 2>&1; then
  echo "Missing dependencies. Activate your environment and install requirements-demo.txt."
  exit 1
fi

cd "${ROOT_DIR}"
exec uvicorn demo.app:app --reload --host 127.0.0.1 --port 8000
