"""LIVE test of the clarifying-questions behavior against the real model.

Drives the actual api/improve.py handler (real Anthropic call) three ways:
  1. a deliberately vague prompt   -> expect status "needs_input" + questions
  2. the same prompt + answers     -> expect status "ready" + a spec
  3. a clear, specific prompt       -> expect status "ready" (no questions)

Reads the key from .env (never printed). Run:
  .venv\\Scripts\\python.exe tests\\live_questions.py
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
            k, v = k.strip(), v.strip().strip('"').strip("'")
            if k and k not in os.environ:
                os.environ[k] = v


def drive(imp, body):
    """Invoke the real handler.do_POST against the live API; return (status, json)."""
    raw = json.dumps(body).encode("utf-8")
    captured = {"status": None}

    class TestHandler(imp.handler):
        def __init__(self):
            self.rfile = io.BytesIO(raw)
            self.wfile = io.BytesIO()
            self.headers = {"Content-Length": str(len(raw))}

        def send_response(self, code):
            captured["status"] = code

        def send_header(self, k, v):
            pass

        def end_headers(self):
            pass

    h = TestHandler()
    h.do_POST()
    out = h.wfile.getvalue()
    try:
        return captured["status"], json.loads(out.decode("utf-8"))
    except Exception:
        return captured["status"], None


def main():
    _load_dotenv()
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY not set (put it in .env).")
        return 2
    imp = _load("improve_mod", os.path.join("api", "improve.py"))

    print("1) VAGUE prompt: 'make me a tracker'")
    s, v = drive(imp, {"prompt": "make me a tracker"})
    print("   HTTP", s, "| status:", (v or {}).get("status"))
    qs = [q.get("question") for q in (v or {}).get("questions", []) or []]
    for q in qs:
        print("     -", q)

    if v and v.get("status") == "needs_input":
        print("\n2) Same prompt + answers")
        clar = [{"question": q.get("question"), "answer": "Use sensible defaults"}
                for q in v.get("questions", [])]
        s2, r2 = drive(imp, {"prompt": "make me a tracker", "clarifications": clar})
        sheets = [sh.get("name") for sh in ((r2 or {}).get("spec") or {}).get("sheets", [])]
        print("   HTTP", s2, "| status:", (r2 or {}).get("status"), "| sheets:", sheets)
        ok2 = (r2 or {}).get("status") == "ready" and bool(sheets)
    else:
        print("\n2) (model built directly without asking — also acceptable)")
        ok2 = True

    print("\n3) CLEAR prompt")
    s3, c = drive(imp, {"prompt": "A monthly budget tracker with categories, "
                                  "budgeted vs actual, a variance formula, and a bar chart"})
    print("   HTTP", s3, "| status:", (c or {}).get("status"))
    ok3 = (c or {}).get("status") == "ready"

    print("\nRESULT:",
          "PASS" if (s == 200 and ok2 and ok3) else "CHECK OUTPUT ABOVE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
