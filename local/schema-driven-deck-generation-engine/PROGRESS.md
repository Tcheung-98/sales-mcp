# Pitch deck — goal + progress

> **Living file.** Update when a ticket lands or the associate flow changes.  
> **End-state SoT (do not fork):** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)  
> **Agents:** `.cursor/rules/pitch-deck-end-scope.mdc` is the always-on snapshot of this + the SoT.  
> **Last updated:** 2026-08-24 (I3 — associate confirm/swap → lock mix)

---

## Goal

Associates pitch from a **human-locked mix**, then get a FortuneAI PPTX.

```text
Form (Discovery) → requirements / spec / PRD
  → SalesGPT Ideation brain proposes offerings
  → Associate selects the options they want          ← I3; gate
  → build_deck (C1 spine + A5 product clones + C2 fills)
```

This is already in the Workflow spec (Discovery → Ideation → Creation). It is **not** “model picks slides while generating.”

**Humans vet offerings (products).** They do not pick Hunter `Slide #`. After I3 lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #`; C1 pastes that page under the funded divider. C2 never changes those clones.

---

## Who owns which step

| Step | Owner | Ticket |
|---|---|---|
| Form / Discovery fields | Schema + Prodie/UI | C3 done; wire PI-2350 |
| Propose offerings | Logic Guide engine + GTM/inventory | I1, I2 |
| Associate select/swap | MCP `confirm_mix`; seller UX later | I3 **done** (engine); Prodie UI PI-2350 |
| Product slide identity | GTM Product Tags exact map | A5 done |
| Deck body (dividers + clones) | `assemble_skeleton` | C1 landed |
| Intro/narrative/investment/thanks fills | Placeholder pipeline | C2 **done** ([PI-2757](https://fortune.atlassian.net/browse/PI-2757), PR [#26](https://github.com/Tcheung-98/sales-mcp/pull/26)) |
| Seller actually calls `build_deck` | Prodie | PI-2350 |

C2 tests may still stub `confirmed_products`. The engineer path is now `propose_mix` → `confirm_mix` → `build_deck`.

---

## Progress (2026-08-24)

| ID | Ticket | Status | Notes |
|---|---|---|---|
| A1–A3 | PI-2517–2519 | Done | Clone/delete/insert + `build_deck` |
| C3 | PI-2758 | Done | `DiscoverySchema` + `DeckSchema` |
| A5 | PI-2541 | Done | Exact Deck Path / Slide #; merged #23 |
| C1 | PI-2756 | **Done** | FortuneAI spine, unfunded dividers dropped, A5 inserts; merged #24 |
| I1 | PI-2759 | **Done** | GTM + inventory + pricing loaders; [#27](https://github.com/Tcheung-98/sales-mcp/pull/27) |
| I2 | PI-2760 | **Done** | Logic Guide V1 mix engine; [#28](https://github.com/Tcheung-98/sales-mcp/pull/28) |
| I3 | PI-2761 | **Done** (MCP) | `propose_mix` + `confirm_mix` lock `confirmed_products`; this PR |
| C2 | PI-2757 | **Done** | Deterministic + bounded AI fills in `build()`; [#26](https://github.com/Tcheung-98/sales-mcp/pull/26) |
| Wire | PI-2350 | Later | Prodie form + confirm UI + `build_deck` |
| Stylist | PI-2754 | Shelved | Not MVP |

**Creation rail** after C2: engineer can `build_deck` with a stubbed mix and get a seller-readable PPTX (deterministic + bounded AI narrative copy). Live Claude + manual FortuneAI PPTX review recommended before prod deploy.

**Associate rail** after I3: engineer can `propose_mix(discovery)` → seller edits → `confirm_mix(...)` → `build_deck(deck_schema)` without hand-building `confirmed_products`. Prodie/SalesGPT UI is still PI-2350.

---

## C1 vs C2 (keep this straight)

- **C1** — Load FortuneAI (the arc is the file). Drop unfunded category dividers. Paste A5 clones under funded chapters. Leave intro / all audience variants / all program variants / investment / thanks as stock.
- **C2** — Fill stock slots (deterministic + bounded Claude); pick one Audience page and one Program Overview page; fill investment/thanks. Do not rewrite product clones. Do not pick the mix.

**Associate rail** still needs Prodie (PI-2350) before the form-to-deck loop is real. MCP propose → confirm → build is wired.

---

## I3 working notes (this cycle)

- `ingestion/confirm_mix.py` — `ConfirmMixRequest`, `ProposedProduct` → `Product`, fail-loud GTM / unavailable / empty-mix checks.
- MCP: `propose_mix` (S3 GTM + inventory → Logic Guide) and `confirm_mix` (tier pick + drop/swap/add → `DeckSchema`).
- Escalation intakes return `{status: "escalation"}` — do not call `build_deck`.
- Tests: `PYTHONPATH=. uv run pytest tests/test_confirm_mix.py tests/test_ideation_to_deck_integration.py tests/test_server.py -q`
- Optional live smoke: `PYTHONPATH=. uv run python tests/smoke_ideation_to_deck.py --mock-ai` (needs `S3_SNAPSHOT_BUCKET`).

**I2 follow-ups (not this ticket):** richer Logic Guide funding rules, re-run engine on swap, session persistence.

---

## C2 working notes (this cycle)

- Plan: review-gated chunks; C2 landed in [#25](https://github.com/Tcheung-98/sales-mcp/pull/25) (deterministic) + [#26](https://github.com/Tcheung-98/sales-mcp/pull/26) (bounded AI + docs).
- **`ingestion/placeholder_fills.py` + `ingestion/placeholder_ai.py`** — `build()` fills after `assemble_skeleton`; `assemble_skeleton` stays Anthropic-free.
- Smoke: `PYTHONPATH=. uv run python tests/smoke_build_live.py --mock-ai` (real S3 template/GTM; mock Claude).

---

## File index (this folder)

| File | Use |
|---|---|
| `END-SCOPE-SOT.md` | Canonical end state + Workflow/Logic Guide distillation |
| `I1-DATA-SOURCES.md` | PI-2759 access path, sheet contract, sync/ownership |
| `PROGRESS.md` | This file — goal + ticket status |
| `C2-PLACEHOLDER-INVENTORY.md` | C2 Chunk 0 token/shape locks |
| `PI-2757-TECH-DEBT.md` | C2 leftover gaps; live-template probe before more fill rewrites |
| `CLEANUP-TODO.md` | Titan/RAG delete later; not C2 |
| `NEXT-STEPS.md` | Historical Phase A/B + retention; defer to SoT on conflict |
| `TICKETS.md` / `JIRA.md` / `README.md` | Historical A/B writeups; defer to SoT |
| `B2-PLAN.md` | Shelved stylist |
