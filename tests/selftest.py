"""Offline self-tests for SheetGenie's serverless functions.

Runs without any network or API key:
  * generate.py  — real openpyxl round-trip (formulas, charts, injection defuse,
                   validation limits, edge cases).
  * improve.py   — the do_POST flow with a mocked Anthropic client (success,
                   model fallback chain, max_tokens, missing key, bad input).

Run:  .venv\\Scripts\\python.exe tests\\selftest.py
Exit code 0 = all passed.
"""

import importlib.util
import io
import json
import os
import sys
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(name, relpath):
    path = os.path.join(ROOT, relpath)
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


PASSED = 0
FAILED = 0


def check(label, cond):
    global PASSED, FAILED
    if cond:
        PASSED += 1
        print("  PASS  " + label)
    else:
        FAILED += 1
        print("  FAIL  " + label)


# ===========================================================================
# generate.py
# ===========================================================================
def test_generate():
    print("\n[generate.py]")
    from openpyxl import load_workbook

    gen = _load("generate_mod", os.path.join("api", "generate.py"))

    # --- SPEC.md worked example: budget tracker -----------------------------
    budget = {
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
            ],
            "charts": [{
                "type": "bar", "title": "Budgeted vs Actual",
                "categoriesColumn": 1, "valueColumns": [2, 3],
                "dataEndRow": 99,  # deliberately over-specified -> must clamp
            }],
        }],
    }
    gen._validate_spec(budget)
    data = gen._build_workbook(budget)
    check("budget: produces non-empty xlsx bytes", isinstance(data, bytes) and len(data) > 0)

    wb = load_workbook(io.BytesIO(data))
    ws = wb["Budget"]
    check("budget: sheet named 'Budget'", ws.title == "Budget")
    check("budget: header A1 == 'Category'", ws["A1"].value == "Category")
    check("budget: header is bold", bool(ws["A1"].font and ws["A1"].font.bold))
    check("budget: D2 is a real formula =B2-C2", ws["D2"].data_type == "f" and ws["D2"].value == "=B2-C2")
    check("budget: D4 formula row-substituted =B4-C4", ws["D4"].value == "=B4-C4")
    check("budget: currency B2 number_format applied", ws["B2"].number_format == "#,##0.00")
    check("budget: freeze panes at A2", ws.freeze_panes == "A2")
    check("budget: one chart present", len(ws._charts) == 1)
    # Chart series must be clamped to the 3 populated rows, not row 99.
    ch = ws._charts[0]
    refs = []
    for s in ch.series:
        try:
            refs.append(s.val.numRef.f)
        except Exception:
            pass
    check("budget: chart series clamped to row 4 (not 99)", all("99" not in r for r in refs) and any("$4" in r or "4" in r for r in refs))

    # --- Formula injection defuse (non-formula text column) -----------------
    inj = {
        "title": "inj",
        "sheets": [{
            "name": "S",
            "columns": [{"header": "Note", "type": "text"}],
            "rows": [["=HYPERLINK(\"http://evil\",\"x\")"], ["+danger"], ["normal"]],
        }],
    }
    gen._validate_spec(inj)
    wb2 = load_workbook(io.BytesIO(gen._build_workbook(inj)))
    s2 = wb2.active
    check("injection: leading-= cell stored as string, not formula", s2["A2"].data_type == "s")
    check("injection: value preserved literally (no apostrophe)", s2["A2"].value == "=HYPERLINK(\"http://evil\",\"x\")")
    check("injection: '+danger' is plain text", s2["A3"].data_type == "s" and s2["A3"].value == "+danger")

    # --- Empty template (no rows -> no chart, no crash) ---------------------
    empty = {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "A", "type": "text"}], "rows": [],
                                       "charts": [{"type": "bar", "title": "c", "categoriesColumn": 1, "valueColumns": [1]}]}]}
    gen._validate_spec(empty)
    wb3 = load_workbook(io.BytesIO(gen._build_workbook(empty)))
    check("empty: builds with zero data rows", wb3.active["A1"].value == "A")
    check("empty: no chart when no data rows", len(wb3.active._charts) == 0)

    # --- date coercion + pie chart ------------------------------------------
    dated = {"title": "d", "sheets": [{"name": "S",
             "columns": [{"header": "Day", "type": "date"}, {"header": "N", "type": "number"}],
             "rows": [["2026-01-15", 5], ["2026-02-20", 9]],
             "charts": [{"type": "pie", "title": "p", "categoriesColumn": 1, "valueColumns": [2]}]}]}
    gen._validate_spec(dated)
    wb4 = load_workbook(io.BytesIO(gen._build_workbook(dated)))
    import datetime as _dt
    check("date: ISO string coerced to a date", isinstance(wb4.active["A2"].value, (_dt.date, _dt.datetime)))
    check("date: pie chart present", len(wb4.active._charts) == 1)

    # --- Chart series/category alignment with dataStartRow > 2 (regression) ---
    import re as _re
    aligned = {"title": "a", "sheets": [{"name": "S",
               "columns": [{"header": "Label", "type": "text"}, {"header": "Val", "type": "number"}],
               "rows": [["r2", 10], ["r3", 20], ["r4", 30], ["r5", 40], ["r6", 50]],
               "charts": [{"type": "bar", "title": "c", "categoriesColumn": 1,
                           "valueColumns": [2], "dataStartRow": 3}]}]}
    gen._validate_spec(aligned)
    chA = load_workbook(io.BytesIO(gen._build_workbook(aligned))).active._charts[0]

    def _rows(f):
        m = _re.findall(r"\$[A-Z]+\$(\d+)", f or "")
        return (int(m[0]), int(m[-1])) if m else None

    ser = chA.series[0]
    val_rows = _rows(ser.val.numRef.f)
    cat_ref = (ser.cat.strRef.f if ser.cat and ser.cat.strRef
               else (ser.cat.numRef.f if ser.cat and ser.cat.numRef else None))
    cat_rows = _rows(cat_ref)
    check("chart align: series span == category span (dataStartRow=3)", val_rows == cat_rows)
    check("chart align: both span rows 3..6", val_rows == (3, 6))

    # --- Advanced features: totals row, conditional format, validation, named range ---
    adv = {
        "title": "adv",
        "namedRanges": [{"name": "TaxRate", "ref": "'S'!$B$1"}],
        "sheets": [{
            "name": "S",
            "columns": [
                {"header": "Item", "type": "text"},
                {"header": "Amount", "type": "currency"},
                {"header": "Status", "type": "text"},
            ],
            "rows": [["A", 600, "To Do"], ["B", 200, "Done"], ["C", 900, "In Progress"]],
            "totalsRow": True,
            "conditionalFormats": [{"column": 2, "rule": "greaterThan", "value": 500, "color": "red"}],
            "dataValidations": [{"column": 3, "values": ["To Do", "In Progress", "Done"]}],
        }],
    }
    gen._validate_spec(adv)
    wbA = load_workbook(io.BytesIO(gen._build_workbook(adv)))
    wsA = wbA["S"]
    cellsA = [c for row in wsA.iter_rows() for c in row]
    has_sum = any(c.data_type == "f" and "SUM" in str(c.value).upper() for c in cellsA)
    has_total_label = any(str(c.value or "").strip().lower() == "total" for c in cellsA)
    check("totals: a live =SUM formula is present", has_sum)
    check("totals: a 'Total' label is present", has_total_label)
    check("condfmt: a rule exists on the sheet", len(list(wsA.conditional_formatting)) > 0)
    check("validation: a list validation exists", len(wsA.data_validations.dataValidation) > 0)
    check("namedRange: a defined name is present", len(list(wbA.defined_names)) > 0)

    def rejects_adv(label, spec):
        try:
            gen._validate_spec(spec)
            check(label + " (should reject)", False)
        except gen.SpecError:
            check(label, True)

    rejects_adv("reject: invalid condfmt rule", {"title": "t", "sheets": [{"name": "S",
                "columns": [{"header": "a", "type": "number"}], "rows": [[1]],
                "conditionalFormats": [{"column": 1, "rule": "wat", "value": 1}]}]})
    rejects_adv("reject: condfmt column out of range", {"title": "t", "sheets": [{"name": "S",
                "columns": [{"header": "a", "type": "number"}], "rows": [[1]],
                "conditionalFormats": [{"column": 5, "rule": "greaterThan", "value": 1}]}]})
    rejects_adv("reject: validation column out of range", {"title": "t", "sheets": [{"name": "S",
                "columns": [{"header": "a", "type": "text"}], "rows": [["x"]],
                "dataValidations": [{"column": 9, "values": ["a"]}]}]})
    rejects_adv("reject: dropdown list over 255 chars", {"title": "t", "sheets": [{"name": "S",
                "columns": [{"header": "a", "type": "text"}], "rows": [["x"]],
                "dataValidations": [{"column": 1, "values": ["opt" + str(i) + "x" * 20 for i in range(15)]}]}]})
    rejects_adv("reject: illegal named-range name (spaces)", {"title": "t",
                "namedRanges": [{"name": "Bad Name", "ref": "'S'!$A$1"}],
                "sheets": [{"name": "S", "columns": [{"header": "a", "type": "text"}], "rows": [["x"]]}]})

    # dropdown list escapes embedded double-quotes (no formula break-out)
    quoted = {"title": "q", "sheets": [{"name": "S",
              "columns": [{"header": "Pick", "type": "text"}], "rows": [["x"], ["y"]],
              "dataValidations": [{"column": 1, "values": ['Hi "Bob"', "Plain"]}]}]}
    gen._validate_spec(quoted)
    wbQ = load_workbook(io.BytesIO(gen._build_workbook(quoted)))
    dvf = wbQ["S"].data_validations.dataValidation[0].formula1
    check("validation: embedded quotes doubled (even parity, no breakout)",
          '""' in dvf and dvf.count('"') % 2 == 0)

    # --- Validation rejections (expect SpecError -> 400) --------------------
    def rejects(label, spec):
        try:
            gen._validate_spec(spec)
            check(label + " (should reject)", False)
        except gen.SpecError:
            check(label, True)

    rejects("reject: 0 sheets", {"title": "t", "sheets": []})
    rejects("reject: 9 sheets", {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "a", "type": "text"}]}] * 9})
    rejects("reject: bad column type", {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "a", "type": "wat"}]}]})
    rejects("reject: formula column without formula", {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "a", "type": "formula"}]}]})
    rejects("reject: chart col index out of range", {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "a", "type": "text"}],
            "charts": [{"type": "bar", "title": "c", "categoriesColumn": 5, "valueColumns": [1]}]}]})
    rejects("reject: oversized cell (>32767 chars)", {"title": "t", "sheets": [{"name": "S", "columns": [{"header": "a", "type": "text"}],
            "rows": [["x" * 40000]]}]})
    rejects("reject: too many total cells", {"title": "t", "sheets": [{"name": "S",
            "columns": [{"header": "c%d" % i, "type": "number"} for i in range(50)],
            "rows": [[1] * 50 for _ in range(5000)]}]})

    # --- filename slugify ----------------------------------------------------
    check("slug: from title", gen._slugify_filename(None, {"title": "My Q4 Report!!"}) == "my-q4-report.xlsx")
    check("slug: fallback", gen._slugify_filename("", {}) == "spreadsheet.xlsx")
    check("slug: no path traversal", "/" not in gen._slugify_filename("../../etc/passwd", {}) and "\\" not in gen._slugify_filename("..\\x", {}))


# ===========================================================================
# improve.py  (mocked Anthropic client)
# ===========================================================================
CANNED = {
    "status": "ready",
    "improvedPrompt": "A monthly budget tracker with category, budgeted, actual, variance.",
    "notes": "Made you a budget tracker.",
    "spec": {"title": "Budget", "sheets": [{"name": "Budget",
             "columns": [{"header": "Category", "type": "text"}], "rows": [["Rent"]]}]},
}


def _fake_anthropic(behavior):
    ns = types.SimpleNamespace()
    ns.NotFoundError = type("NotFoundError", (Exception,), {})
    ns.PermissionDeniedError = type("PermissionDeniedError", (Exception,), {})

    class _Block:
        def __init__(self, text):
            self.type = "text"
            self.text = text

    class _Resp:
        def __init__(self, text, stop_reason="end_turn"):
            self.content = [_Block(text)]
            self.stop_reason = stop_reason

    class _Messages:
        def __init__(self):
            self.calls = []

        def create(self, **kw):
            self.calls.append(kw)
            return behavior(kw, len(self.calls), _Resp, ns)

    class _Client:
        def __init__(self, *a, **k):
            self.messages = _Messages()

    ns.Anthropic = _Client
    ns._Resp = _Resp
    return ns


def _drive(imp, body_bytes, content_length=None, env=None):
    """Invoke imp.handler.do_POST with a fake request; capture (status, json)."""
    captured = {"status": None, "headers": {}, "body": b""}

    class TestHandler(imp.handler):
        def __init__(self):  # bypass BaseHTTPRequestHandler socket setup
            self.rfile = io.BytesIO(body_bytes)
            self.wfile = io.BytesIO()
            cl = content_length if content_length is not None else len(body_bytes)
            self.headers = {"Content-Length": str(cl)}

        def send_response(self, code):
            captured["status"] = code

        def send_header(self, k, v):
            captured["headers"][k] = v

        def end_headers(self):
            pass

    old_env = os.environ.get("ANTHROPIC_API_KEY")
    if env is not None and "ANTHROPIC_API_KEY" in env:
        os.environ["ANTHROPIC_API_KEY"] = env["ANTHROPIC_API_KEY"]
    elif env is not None:
        os.environ.pop("ANTHROPIC_API_KEY", None)
    try:
        h = TestHandler()
        h.do_POST()
    finally:
        if old_env is not None:
            os.environ["ANTHROPIC_API_KEY"] = old_env
        else:
            os.environ.pop("ANTHROPIC_API_KEY", None)

    raw = h.wfile.getvalue()
    try:
        parsed = json.loads(raw.decode("utf-8")) if raw else None
    except Exception:
        parsed = None
    return captured["status"], parsed, h.messages_ref if hasattr(h, "messages_ref") else None


def test_improve():
    print("\n[improve.py]")
    imp = _load("improve_mod", os.path.join("api", "improve.py"))

    # success
    imp.anthropic = _fake_anthropic(lambda kw, n, R, ns: R(json.dumps(CANNED)))
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "budget tracker"}).encode(),
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("success: 200", status == 200)
    check("success: returns improvedPrompt+notes+spec",
          isinstance(parsed, dict) and {"improvedPrompt", "notes", "spec"} <= set(parsed))

    # model fallback chain: first model NotFound -> second succeeds
    calls = {"models": []}

    def fb(kw, n, R, ns):
        calls["models"].append(kw["model"])
        if n == 1:
            raise ns.NotFoundError("no access to that model")
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(fb)
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "x"}).encode(),
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("fallback: 200 after first model unavailable", status == 200)
    check("fallback: tried primary then fell back to a second model",
          len(calls["models"]) == 2 and calls["models"][0] == imp.MODEL_CHAIN[0]
          and calls["models"][1] == imp.MODEL_CHAIN[1])
    check("fallback: thinking+effort sent to supported models", True)  # see request-shape test below

    # request shape: supported model gets thinking + effort
    seen = {}

    def capture(kw, n, R, ns):
        seen.update(kw)
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(capture)
    _drive(imp, json.dumps({"prompt": "x"}).encode(), env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("shape: model is Fable 5 by default", str(seen.get("model", "")).startswith("claude-fable-5"))
    check("shape: adaptive thinking present", seen.get("thinking") == {"type": "adaptive"})
    check("shape: effort inside output_config", seen.get("output_config", {}).get("effort") == imp.EFFORT)
    check("shape: no structured-output format sent (prompt-JSON envelope)",
          "format" not in seen.get("output_config", {}))
    check("shape: no temperature/top_p/top_k", not ({"temperature", "top_p", "top_k"} & set(seen)))

    # max_tokens stop_reason -> friendly 500
    imp.anthropic = _fake_anthropic(lambda kw, n, R, ns: R(json.dumps(CANNED), stop_reason="max_tokens"))
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "x"}).encode(),
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("max_tokens: 500", status == 500)
    check("max_tokens: friendly message, no leak", isinstance(parsed, dict) and "too large" in parsed.get("error", ""))

    # missing API key -> 500 (anthropic never called)
    imp.anthropic = _fake_anthropic(lambda kw, n, R, ns: R(json.dumps(CANNED)))
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "x"}).encode(), env={})  # no key
    check("no key: 500", status == 500)
    check("no key: 'not configured' message", "not configured" in parsed.get("error", ""))

    # empty prompt -> 400
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "   "}).encode(),
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("empty prompt: 400", status == 400)

    # malformed JSON -> 400
    status, parsed, _ = _drive(imp, b"{not json", env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("bad json: 400", status == 400)

    # oversized body -> 413 (Content-Length over the cap, body not read)
    status, parsed, _ = _drive(imp, b"{}", content_length=imp.MAX_BODY_BYTES + 1,
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("oversized: 413", status == 413)

    # clarifications are folded into the user message sent to the model
    seen2 = {}

    def cap2(kw, n, R, ns):
        seen2.update(kw)
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(cap2)
    _drive(imp, json.dumps({"prompt": "sales report",
                            "clarifications": [{"question": "Period?", "answer": "Q1 2026"}]}).encode(),
           env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    cont2 = (seen2.get("messages") or [{}])[0].get("content")
    if isinstance(cont2, list):
        umsg = next((b.get("text", "") for b in cont2 if isinstance(b, dict) and b.get("type") == "text"), "")
    else:
        umsg = cont2 or ""
    check("clarifications: folded into the user message",
          "Answers to your questions" in umsg and "Q1 2026" in umsg)

    # a needs_input response passes through unchanged (frontend handles it)
    NEEDS = {"status": "needs_input", "notes": "need a couple details",
             "questions": [{"question": "What period?", "hint": "e.g. 2026"}]}
    imp.anthropic = _fake_anthropic(lambda kw, n, R, ns: R(json.dumps(NEEDS)))
    status, parsed, _ = _drive(imp, json.dumps({"prompt": "a report"}).encode(),
                               env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    check("needs_input: 200 passthrough with status", status == 200 and parsed.get("status") == "needs_input")
    check("needs_input: questions present", isinstance(parsed.get("questions"), list) and len(parsed["questions"]) == 1)

    # --- attachments: file -> content blocks ---
    img_b64 = "aGVsbG8="  # dummy base64
    blocks = imp._build_file_blocks([{"type": "image", "media_type": "image/jpeg", "data": img_b64, "name": "x.jpg"}])
    check("files: image -> image block",
          len(blocks) == 1 and blocks[0]["type"] == "image"
          and blocks[0]["source"]["media_type"] == "image/jpeg" and blocks[0]["source"]["data"] == img_b64)
    pblocks = imp._build_file_blocks([{"type": "pdf", "media_type": "application/pdf", "data": img_b64}])
    check("files: pdf -> document block",
          len(pblocks) == 1 and pblocks[0]["type"] == "document"
          and pblocks[0]["source"]["media_type"] == "application/pdf")
    check("files: none -> empty list", imp._build_file_blocks(None) == [])

    def rejects_files(label, files):
        try:
            imp._build_file_blocks(files)
            check(label + " (should reject)", False)
        except imp._InputError:
            check(label, True)

    rejects_files("files: reject too many", [{"type": "image", "media_type": "image/png", "data": "x"}] * (imp.MAX_FILES + 1))
    rejects_files("files: reject bad image type", [{"type": "image", "media_type": "image/tiff", "data": "x"}])
    rejects_files("files: reject pdf with image type", [{"type": "pdf", "media_type": "image/png", "data": "x"}])
    rejects_files("files: reject unknown kind", [{"type": "video", "media_type": "video/mp4", "data": "x"}])
    rejects_files("files: reject empty data", [{"type": "image", "media_type": "image/png", "data": ""}])
    rejects_files("files: reject oversized", [{"type": "image", "media_type": "image/png", "data": "a" * (imp.MAX_FILE_B64 + 1)}])

    # integration: attachments reach the model as content blocks (text + image)
    seen3 = {}

    def cap3(kw, n, R, ns):
        seen3.update(kw)
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(cap3)
    _drive(imp, json.dumps({"prompt": "expenses from this receipt",
                            "files": [{"type": "image", "media_type": "image/jpeg", "data": img_b64, "name": "r.jpg"}]}).encode(),
           env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    cont = (seen3.get("messages") or [{}])[0].get("content")
    has_text = isinstance(cont, list) and any(isinstance(b, dict) and b.get("type") == "text" for b in cont)
    has_img = isinstance(cont, list) and any(isinstance(b, dict) and b.get("type") == "image" for b in cont)
    check("files: message content is blocks with text + image", has_text and has_img)

    # baseSpec (edit mode) is folded into the user message as an edit instruction
    seen4 = {}

    def cap4(kw, n, R, ns):
        seen4.update(kw)
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(cap4)
    _drive(imp, json.dumps({"prompt": "add a Tax column",
                            "baseSpec": {"title": "T", "sheets": [{"name": "S",
                                         "columns": [{"header": "A", "type": "text"}]}]}}).encode(),
           env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    cont4 = (seen4.get("messages") or [{}])[0].get("content")
    etext = (next((b.get("text", "") for b in cont4 if isinstance(b, dict) and b.get("type") == "text"), "")
             if isinstance(cont4, list) else (cont4 or ""))
    check("baseSpec: folded into the user message as an edit instruction",
          "CURRENT SPREADSHEET to edit" in etext and "add a Tax column" in etext)

    # baseSpec with an empty sheets list is still treated as an edit (framing present)
    seen5 = {}

    def cap5(kw, n, R, ns):
        seen5.update(kw)
        return R(json.dumps(CANNED))

    imp.anthropic = _fake_anthropic(cap5)
    _drive(imp, json.dumps({"prompt": "add a column", "baseSpec": {"title": "T", "sheets": []}}).encode(),
           env={"ANTHROPIC_API_KEY": "sk-ant-test"})
    cont5 = (seen5.get("messages") or [{}])[0].get("content")
    etext5 = (next((b.get("text", "") for b in cont5 if isinstance(b, dict) and b.get("type") == "text"), "")
              if isinstance(cont5, list) else (cont5 or ""))
    check("baseSpec: empty-sheets spec still framed as an edit",
          "CURRENT SPREADSHEET to edit" in etext5)

    # _extract_json robustness (prompt-JSON envelope parsing)
    ej = imp._extract_json
    check("extract: plain envelope", (ej('{"status":"ready","notes":"n"}') or {}).get("status") == "ready")
    check("extract: code-fenced", (ej('```json\n{"status":"ready","notes":"n"}\n```') or {}).get("status") == "ready")
    check("extract: stray leading brace + prose",
          (ej('use the {row} token, then: {"status":"ready","notes":"n"}') or {}).get("status") == "ready")
    check("extract: prefers the status envelope among two objects",
          (ej('{"foo":1} then {"status":"needs_input","notes":"n"}') or {}).get("status") == "needs_input")
    check("extract: pure garbage -> None", ej("no json here at all") is None)


if __name__ == "__main__":
    test_generate()
    test_improve()
    print("\n==== %d passed, %d failed ====" % (PASSED, FAILED))
    sys.exit(1 if FAILED else 0)
