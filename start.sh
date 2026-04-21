#!/bin/sh
set -e
if command -v python3 >/dev/null 2>&1; then
  exec python3 serve_static.py
fi
exec python serve_static.py
