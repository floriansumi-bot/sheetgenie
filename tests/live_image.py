"""LIVE test: photo -> spreadsheet. Generates a receipt image, sends it through the
real api/improve.py handler (vision), and checks the model extracted the data.

Reads the key from .env (never printed). Run:
  .venv\\Scripts\\python.exe tests\\live_image.py
"""
import base64
import importlib.util
import io
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

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


def _font(size):
    for path in (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\segoeui.ttf"):
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def make_receipt_png():
    img = Image.new("RGB", (760, 430), "white")
    d = ImageDraw.Draw(img)
    d.text((40, 28), "CITY CAFE  —  RECEIPT", fill="black", font=_font(40))
    rows = [("Item", "Price"), ("Cappuccino", "4.50"), ("Sandwich", "8.20"),
            ("Orange Juice", "3.00"), ("Muffin", "2.80"), ("TOTAL", "18.50")]
    y = 120
    for item, price in rows:
        d.text((40, y), item, fill="black", font=_font(32))
        d.text((560, y), price, fill="black", font=_font(32))
        y += 50
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return img, buf.getvalue()


def drive(imp, body):
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
    try:
        return captured["status"], json.loads(h.wfile.getvalue().decode("utf-8"))
    except Exception:
        return captured["status"], None


def main():
    _load_dotenv()
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY not set (put it in .env).")
        return 2
    imp = _load("improve_mod", os.path.join("api", "improve.py"))
    gen = _load("generate_mod", os.path.join("api", "generate.py"))

    img, png = make_receipt_png()
    img.save(os.path.join(ROOT, "live_image_input.local.png"))
    b64 = base64.b64encode(png).decode("ascii")
    print("receipt PNG: %d bytes (%d base64 chars)" % (len(png), len(b64)))

    body = {
        "prompt": "Turn this cafe receipt into a spreadsheet listing each item and its price.",
        "files": [{"type": "image", "media_type": "image/png", "data": b64, "name": "receipt.png"}],
    }
    s, r = drive(imp, body)
    r = r or {}
    print("HTTP", s, "| status:", r.get("status"))

    spec = r.get("spec") or {}
    for sh in spec.get("sheets", []):
        print("Sheet:", sh.get("name"), "| columns:", [c.get("header") for c in sh.get("columns", [])])
        for row in (sh.get("rows") or [])[:12]:
            print("   ", row)

    if spec:
        gen._validate_spec(spec)
        data = gen._build_workbook(spec)
        out = os.path.join(ROOT, "live_image_output.local.xlsx")
        with open(out, "wb") as f:
            f.write(data)
        print("wrote", out, len(data), "bytes")

    blob = json.dumps(spec).lower()
    markers = ["cappuccino", "sandwich", "juice", "muffin", "4.5", "8.2", "2.8"]
    hits = [m for m in markers if m in blob]
    print("extracted markers found:", hits)
    print("\nRESULT:", "PASS" if (s == 200 and r.get("status") == "ready" and len(hits) >= 4) else "CHECK OUTPUT")
    return 0


if __name__ == "__main__":
    sys.exit(main())
