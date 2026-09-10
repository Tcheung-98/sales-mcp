# Pitch deck — goal + progress

> **Living file.** Update when a ticket lands or the upstream handoff changes.  
> **End-state SoT (do not fork):** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)  
> **Agents:** `.cursor/rules/pitch-deck-end-scope.mdc` is the always-on snapshot of this + the SoT.  
> **Last updated:** 2026-09-10 (post-cleanup handoff)

---

## Goal

**sales-mcp is deck generation only.** Upstream owns Discovery + product lock; this repo assembles the PPTX.

```text
Pitch Deck Builder / Sales HQ / caller
  → complete DeckSchema (Discovery + confirmed_products[])
  → build_deck (validate + C1 spine + A5 clones + C2 fills + QA rail)
  → PPTX
```

**The caller owns the mix.** Each `confirmed_products[]` entry must include `name`, `category`, `price`, `cadence`.  
**sales-mcp `build_deck` assembles.** It does not rank, propose, or choose products.

**MCP surface (runtime):** `build_deck` (primary) · `confirm_mix` (optional legacy hydrate).

---

## Progress

| ID | Ticket | Status | Notes |
|---|---|---|---|
| A1–A3 | PI-2517–2519 | Done | Clone/delete/insert + `build_deck` |
| C3 | PI-2758 | Done | `DiscoverySchema` + `DeckSchema` |
| A5 | PI-2541 | Done | Exact Deck Path / Slide # |
| C1 | PI-2756 | Done | FortuneAI spine + conditional dividers |
| C2 | PI-2757 | Done | Deterministic + bounded AI fills |
| I1 | PI-2759 | Done | GTM + inventory + pricing loaders |
| I3 | PI-2761 | Optional legacy | `confirm_mix` when caller omits full products |
| QA rail | B2–B4 | Done | Always-on in `build_deck` — see `DECK-QA-ARCHITECTURE.md` |
| RAG / Logic Guide cleanup | — | **Done (2026-09-10)** | Removed Titan research tools, dead generator path, `logic_guide/` |
| Upstream wire | PI-2350 | **Out of repo** | Pitch Deck Builder sends locked DeckSchema |

---

## C1 vs C2

- **C1** — Load FortuneAI. Drop unfunded category dividers. Paste A5 clones under funded chapters.
- **C2** — Fill stock slots; pick one Audience and one Program Overview variant; fill investment/thanks. Do not rewrite product clones.

---

## Repo layout (handoff)

| Path | Use |
|---|---|
| `server.py` | MCP tools: `build_deck`, `confirm_mix` |
| `ingestion/generator.py` | C1 assemble + build orchestration |
| `ingestion/placeholder_fills.py` | C2 deterministic + AI fills |
| `ingestion/gtm_product_map.py` | A5 exact product slide map |
| `ingestion/schema.py` | Discovery + Deck Pydantic models |
| `docs/END-SCOPE-SOT.md` | Canonical end state |
| `docs/I1-DATA-SOURCES.md` | GTM + inventory S3 contract |
| `tests/gtm_fixtures.py` | Representative workbook fixtures for tests |
