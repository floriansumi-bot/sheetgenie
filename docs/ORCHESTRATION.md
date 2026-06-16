# SheetGenie — Agent Orchestration

How this app was built by a team of AI agents, and the methodology behind it.
This is both documentation and the playbook for re-running the build.

## Why multi-agent (the 2026 picture)
Production multi-agent systems in 2026 converge on a small set of control-flow
topologies — **supervisor/hierarchical, pipeline, fan-out, debate, and dynamic
handoff**. The trade-off is coordination overhead: lateral peer-to-peer chatter
adds ~58% token overhead, while a centralized supervisor adds ~285%. The lesson:
**use a supervisor for direction and verification, but let workers run in parallel
on disjoint, well-specified slices** so they don't need to talk to each other —
they coordinate through a shared written contract instead.

Sources:
- [6 Multi-Agent Orchestration Patterns for Production (2026) — beam.ai](https://beam.ai/agentic-insights/multi-agent-orchestration-patterns-production)
- [Multi-Agent Orchestration: 5 Patterns That Work in 2026 — digitalapplied.com](https://www.digitalapplied.com/blog/multi-agent-orchestration-5-patterns-that-work)
- [Multi-Agent Systems Explained: 2026 Patterns — decodethefuture.org](https://decodethefuture.org/en/multi-agent-systems-explained/)

## The pattern we use: Supervisor + Fan-out + Adversarial verify
```
            ┌────────────────── SUPERVISOR (orchestrator) ──────────────────┐
            │  writes the shared contract (SPEC.md), then dispatches work    │
            └───────────────────────────┬────────────────────────────────────┘
                                         │
        ┌────────────────┬───────────────┴───────────────┬──────────────────┐
        ▼                ▼                               ▼                  ▼
  Frontend agent   Improve-API agent            Generate-API agent     (each writes a
  public/*         api/improve.py               api/generate.py         DISJOINT file set)
        └────────────────┴───────────────┬───────────────┴──────────────────┘
                                         │  results land on disk
                                         ▼
                              QA / adversarial verify agent
                       (reads everything, checks against SPEC.md,
                        reports integration & correctness defects)
                                         │
                                         ▼
                          SUPERVISOR applies fixes, runs the app,
                          verifies end-to-end, then ships
```

### How the agents "communicate"
They do **not** message each other directly (that's where token overhead and
drift come from). Instead they share state the way the 2026 pipeline pattern
prescribes: **through a single written contract** — [`SPEC.md`](SPEC.md) — plus the
files on disk. Every agent reads `SPEC.md` first, builds its slice to that exact
HTTP + `SpreadsheetSpec` contract, and the supervisor integrates. The contract is
the message bus.

### The agent team
| Agent | Slice (disjoint files) | Charter |
|-------|------------------------|---------|
| **Frontend** | `public/*` | PWA shell, voice + typed capture, improve→preview→generate→download UX, responsive + installable |
| **Improve-API** | `api/improve.py` | Anthropic call with structured output → valid `SpreadsheetSpec` + improved prompt; sanitized errors |
| **Generate-API** | `api/generate.py` | Deterministic openpyxl renderer: columns, formats, formulas, freeze/filter, bar/line/pie charts |
| **QA / verifier** | (read-only) | Adversarially check each slice against `SPEC.md`; find integration gaps & correctness bugs |
| **Supervisor** | integration, config, docs | Owns the contract, fixes defects, runs the app, verifies end-to-end, deploys |

### The loops
1. **Build loop (fan-out):** dispatch the three builders in parallel; each runs its
   own internal read→write→self-check loop until its slice satisfies `SPEC.md`.
2. **Verify loop (adversarial):** the QA agent tries to break the contract — mismatched
   request/response shapes, unhandled spec edge cases, missing CORS, key leakage. Findings
   feed back to the supervisor.
3. **Integration loop (supervisor):** install deps, run the functions locally, drive the
   real UI in a browser preview, fix, re-run — until a prompt produces a correct `.xlsx`.

### Skillsets pulled in
- **`claude-api`** skill — current Anthropic model IDs, pricing, and the structured-output
  / tool-use patterns used by `api/improve.py`.
- **`xlsx`** domain knowledge (openpyxl) — encoded directly into `api/generate.py`.
- Live web research — orchestration patterns, Vercel Python runtime, Web Speech support
  (cited above and in the build log).

## Re-running the build
The supervisor re-dispatches the three builders against `SPEC.md`, then the QA pass,
then verifies locally. Because the contract is fixed and the file sets are disjoint,
the build is reproducible and parallel-safe.
