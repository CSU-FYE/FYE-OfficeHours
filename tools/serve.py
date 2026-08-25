#!/usr/bin/env python3
"""Local preview server.

Same as `python3 -m http.server`, except it tells the browser not to cache.

The stock server sends no Cache-Control header at all, so browsers fall back to
a heuristic and will happily serve a stale assets/app.js from memory after you
have edited it — you reload, nothing changes, and you go looking for a bug that
is not there. GitHub Pages sends a real max-age, so this only affects previews.

    python3 tools/serve.py [port]
"""

import http.server
import socketserver
import sys
from pathlib import Path

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8777


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        super().end_headers()

    def log_message(self, fmt, *args):
        if "GET" in fmt % args and " 200 " not in fmt % args:
            super().log_message(fmt, *args)


class Server(socketserver.TCPServer):
    allow_reuse_address = True


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    import os

    os.chdir(root)
    with Server(("", PORT), NoCacheHandler) as httpd:
        print(f"Office hours preview: http://localhost:{PORT}  (ctrl-c to stop)")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass
