"""Local dev server: static files that are never cached, plus a side channel for agents.

POST /__agent/report  body: a text report; written to .agent/<name>.txt (name from ?name=, default "report").
POST /__agent/log     body: one line; appended to .agent/browser.log (console errors, uncaught exceptions).
"""
import http.server
import os
import pathlib
import sys
import urllib.parse

ROOT = pathlib.Path(__file__).resolve().parent
AGENT_DIR = ROOT / ".agent"
MAX_BODY_BYTES = 1 << 20


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_POST(self):
        url = urllib.parse.urlsplit(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY_BYTES:
            self.send_error(413)
            return
        body = self.rfile.read(length).decode("utf-8", "replace")
        AGENT_DIR.mkdir(exist_ok=True)
        if url.path == "/__agent/report":
            name = urllib.parse.parse_qs(url.query).get("name", ["report"])[0]
            if not name.replace("-", "").isalnum():
                self.send_error(400)
                return
            (AGENT_DIR / f"{name}.txt").write_text(body)
        elif url.path == "/__agent/log":
            with open(AGENT_DIR / "browser.log", "a") as log:
                log.write(body.rstrip("\n") + "\n")
        else:
            self.send_error(404)
            return
        self.send_response(204)
        self.end_headers()

    def log_message(self, format, *args):
        if os.environ.get("QUIET"):
            return
        super().log_message(format, *args)


def main():
    host = os.environ.get("HOST", "localhost")
    port = int(os.environ.get("PORT", "8999"))
    server = http.server.ThreadingHTTPServer((host, port), Handler)
    print(f"Serving {ROOT} at http://{host}:{port}/", file=sys.stderr, flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
