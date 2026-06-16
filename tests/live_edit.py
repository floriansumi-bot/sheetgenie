"""LIVE test of the EDIT path: give the real model an existing spec + an
instruction and confirm it returns the modified spec (new column added, original
data preserved). Reads the key from .env (never printed). Run:
  .venv\\Scripts\\python.exe tests\\live_edit.py
"""
import importlib.util
import io
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(name, rel):
    spec = importlib.util.spec_from_file_location(name, os.path.join(ROOT, rel))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _load_dotenv():
    path = os.path.join(ROOT, ".env")
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def drive(imp, body):
    raw = json.dumps(body).encode("utf-8")
    cap = {"status": None}

    class T(imp.handler):
        def __init__(self):
            self.rfile = io.BytesIO(raw)
            self.wfile = io.BytesIO()
            self.headers = {"Content-Length": str(len(raw))}

        def send_response(self, code):
            cap["status"] = code

        def send_header(self, k, v):
            pass

        def end_headers(self):
            pass

    h = T()
    h.do_POST()
    try:
        return cap["status"], json.loads(h.wfile.getvalue().decode("utf-8"))
    except Exception:
        return cap["status"], None


def main():
    _load_dotenv()
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY not set (put it in .env).")
        return 2
    imp = _load("improve_mod", os.path.join("api", "improve.py"))

    base = {"title": "Monthly Budget", "sheets": [{
        "name": "Budget",
        "columns": [
            {"header": "Category", "type": "text"},
            {"header": "Budgeted", "type": "currency"},
            {"header": "Actual", "type": "currency"},
        ],
        "rows": [["Rent", 1500, 1500], ["Groceries", 400, 462.30]],
    }]}

    body = {
        "prompt": "Add a 'Notes' text column at the end, and add a 'Variance' formula "
                  "column that is Budgeted minus Actual. Keep the existing rows.",
        "baseSpec": base,
    }
    print("Editing base spec (cols: Category, Budgeted, Actual) ...")
    s, r = drive(imp, body)
    r = r or {}
    print("HTTP", s, "| status:", r.get("status"))
    spec = r.get("spec") or {}
    sheet0 = (spec.get("sheets") or [{}])[0]
    cols = [c.get("header") for c in sheet0.get("columns", [])]
    types = [c.get("type") for c in sheet0.get("columns", [])]
    print("result columns:", cols)
    print("result rows[0]:", (sheet0.get("rows") or [None])[0])

    has_notes = any(str(c).lower() == "notes" for c in cols)
    has_formula = "formula" in types
    preserved = "Category" in cols and "Budgeted" in cols and "Actual" in cols
    n_rows = len(sheet0.get("rows") or [])
    ok = s == 200 and r.get("status") == "ready" and has_notes and has_formula and preserved and n_rows >= 2
    print("checks: notes=%s formula=%s original_cols=%s rows>=2=%s"
          % (has_notes, has_formula, preserved, n_rows >= 2))
    print("\nRESULT:", "PASS" if ok else "CHECK OUTPUT")
    return 0


if __name__ == "__main__":
    sys.exit(main())
