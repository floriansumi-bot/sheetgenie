"""Local dev server for verifying the frontend without the Vercel CLI or a key.

  * Serves the static site from public/.
  * POST /api/improve -> a canned {improvedPrompt, notes, spec} (mock; the real
    endpoint needs an API key).
  * POST /api/generate -> the REAL api/generate.py rendering logic (no key needed),
    so the download path is exercised genuinely end-to-end.

Run:  .venv\\Scripts\\python.exe tools\\devserver.py   (PORT env optional, default 8000)
"""

import importlib.util
import json
import mimetypes
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PUBLIC = os.path.join(ROOT, "public")

_EXTRA_TYPES = {
    ".js": "text/javascript",
    ".mjs": "text/javascript",
    ".svg": "image/svg+xml",
    ".webmanifest": "application/manifest+json",
    ".json": "application/json",
}


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(ROOT, rel))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


gen = _load("generate_mod", os.path.join("api", "generate.py"))

CANNED = {
    "improvedPrompt": "A monthly budget tracker with a Category column, Budgeted "
                      "and Actual currency columns, and a Variance formula column "
                      "(Budgeted minus Actual), plus a bar chart comparing budgeted "
                      "vs actual by category.",
    "notes": "I built you a budget tracker with a variance formula and a bar chart "
             "comparing budgeted vs actual.",
    "spec": {
        "title": "Monthly Budget Tracker",
        "sheets": [{
            "name": "Budget",
            "columns": [
                {"header": "Category", "type": "text"},
                {"header": "Budgeted", "type": "currency"},
                {"header": "Actual", "type": "currency"},
                {"header": "Variance", "type": "formula", "formula": "=B{row}-C{row}"},
            ],
            "rows": [
                ["Rent", 1500, 1500, None],
                ["Groceries", 400, 462.30, None],
                ["Transport", 120, 98.5, None],
                ["Utilities", 220, 205, None],
                ["Entertainment", 150, 187.4, None],
            ],
            "freezeHeader": True,
            "autoFilter": True,
            "charts": [{
                "type": "bar", "title": "Budgeted vs Actual",
                "categoriesColumn": 1, "valueColumns": [2, 3],
            }],
        }],
    },
}


class H(BaseHTTPRequestHandler):
    def _json(self, status, obj):
        b = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_POST(self):
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b""
        path = self.path.split("?")[0].rstrip("/")

        if path == "/api/improve":
            self._json(200, CANNED)
            return

        if path == "/api/generate":
            try:
                payload = json.loads(raw or b"{}")
                spec = payload.get("spec")
                gen._validate_spec(spec)
                data = gen._build_workbook(spec)
            except gen.SpecError as e:
                self._json(400, {"error": str(e)})
                return
            except Exception as e:  # noqa: BLE001
                self._json(500, {"error": "render failed: %s" % e})
                return
            name = gen._slugify_filename(payload.get("filename"), spec)
            self.send_response(200)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Disposition", 'attachment; filename="%s"' % name)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self._json(404, {"error": "not found"})

    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", ""):
            path = "/index.html"
        fp = os.path.normpath(os.path.join(PUBLIC, path.lstrip("/")))
        if not fp.startswith(PUBLIC) or not os.path.isfile(fp):
            fp = os.path.join(PUBLIC, "index.html")
        ext = os.path.splitext(fp)[1].lower()
        ctype = _EXTRA_TYPES.get(ext) or mimetypes.guess_type(fp)[0] or "application/octet-stream"
        with open(fp, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *a):
        return


if __name__ == "__main__":
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", "8000"))
    print("SheetGenie dev server on http://127.0.0.1:%d" % port)
    ThreadingHTTPServer(("127.0.0.1", port), H).serve_forever()
