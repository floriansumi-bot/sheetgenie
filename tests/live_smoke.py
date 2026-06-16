"""LIVE end-to-end smoke test against the real Anthropic API.

Confirms the actual model (Fable 5, or the fallback chain) returns a schema-valid
spec, then renders it to a real .xlsx. Your key is read from the environment and
NEVER printed.

Setup (key never appears in chat):
  PowerShell (this session):   $env:ANTHROPIC_API_KEY = "sk-ant-..."
  Persist for new terminals:   setx ANTHROPIC_API_KEY "sk-ant-..."   (reopen shell)
Then run:
  .venv\\Scripts\\python.exe tests\\live_smoke.py ["your prompt"]
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
    """Load KEY=VALUE pairs from a local .env (gitignored) into the environment."""
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


def main():
    _load_dotenv()
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY is not set. Put it in a local .env file "
              "(ANTHROPIC_API_KEY=sk-ant-...), or set it in the environment.")
        return 2

    imp = _load("improve_mod", os.path.join("api", "improve.py"))
    gen = _load("generate_mod", os.path.join("api", "generate.py"))
    import anthropic

    prompt = sys.argv[1] if len(sys.argv) > 1 else (
        "A monthly budget tracker with categories, budgeted vs actual columns, "
        "a variance formula, and a bar chart comparing budgeted vs actual.")

    client = anthropic.Anthropic()
    resp, used = None, None
    for model in imp.MODEL_CHAIN:
        oc = {"format": {"type": "json_schema", "schema": imp.RESPONSE_SCHEMA}}
        kw = {}
        if imp._supports_thinking(model):
            kw["thinking"] = {"type": "adaptive"}
            oc["effort"] = imp.EFFORT
        try:
            print("Calling %s ..." % model)
            resp = client.messages.create(
                model=model, max_tokens=imp.MAX_TOKENS, system=imp.SYSTEM_PROMPT,
                messages=[{"role": "user", "content": "Request:\n" + prompt}],
                output_config=oc, **kw)
            used = model
            break
        except (anthropic.NotFoundError, anthropic.PermissionDeniedError) as e:
            print("  %s unavailable (%s) -> falling back" % (model, type(e).__name__))
            continue

    if resp is None:
        print("No model in the chain was accessible with this key.")
        return 1

    if resp.stop_reason == "max_tokens":
        print("Hit max_tokens — raise MAX_TOKENS or simplify the prompt.")
        return 1

    text = next(b.text for b in resp.content if b.type == "text")
    data = json.loads(text)
    print("\nMODEL USED:      ", used)
    print("improvedPrompt:  ", data["improvedPrompt"][:240])
    print("notes:           ", data["notes"])

    spec = data["spec"]
    gen._validate_spec(spec)
    xlsx = gen._build_workbook(spec)
    out = os.path.join(ROOT, "live_smoke_output.local.xlsx")
    with open(out, "wb") as f:
        f.write(xlsx)
    sheets = [s.get("name") for s in spec.get("sheets", [])]
    print("\nWrote %s (%d bytes); sheets: %s" % (out, len(xlsx), sheets))

    u = getattr(resp, "usage", None)
    if u:
        print("tokens: input=%s output=%s" % (getattr(u, "input_tokens", "?"), getattr(u, "output_tokens", "?")))
    print("\nLIVE SMOKE PASS — open the .xlsx to see your spreadsheet.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
