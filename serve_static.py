#!/usr/bin/env python3
"""Sirve la carpeta del proyecto en 0.0.0.0:PORT (Railway, local)."""
import os
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PORT = int(os.environ.get("PORT", "8080"))


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)


def main():
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print(f"Serving {ROOT} at http://0.0.0.0:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
