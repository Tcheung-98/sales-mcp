# Pitch deck — goal + progress

> **Living file.** Update when a ticket lands or the upstream handoff changes.  
> **End-state SoT (do not fork):** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)  
> **Agents:** `.cursor/rules/pitch-deck-end-scope.mdc` is the always-on snapshot of this + the SoT.  
> **Last updated:** 2026-09-09 (sales-mcp = deckgen only; locked DeckSchema from upstream)

---

## Goal

**sales-mcp is deck generation only.** Upstream owns Discovery + product lock; this repo assembles the PPTX.

```text
Pitch Deck Builder / Sales HQ / caller
  → complete DeckSchema (Discovery + confirmed_products[])
  → build_deck (validate + C1 spine + A5 clones + C2 fills [+ optional QA rail])
  → PPTX
```

**The caller owns the mix.** Each `confirmed_products[]` entry must include `name`, `category`, `price`, `cadence`.  
**sales-mcp `build_deck` assembles.** It does not rank, propose, or choose products.

**Humans vet offerings (products).** They do not pick Hunter `Slide #`. After upstream lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #`; C1 pastes that page under the funded divider. C2 never changes those clones.

**Prodie is not the ideation brain on this MCP.** Historical Prodie menu/checkbox design lives in [`docs/PRODIE-IDEATION-SPEC.md`](../../docs/PRODIE-IDEATION-SPEC.md) — not the runtime contract.

---

## Who owns which step

| Step | Owner | Ticket |
|---|---|---|
| Discovery + product lock | **Upstream** (Pitch Deck Builder / Sales HQ) | PI-2350 (out of repo) |
| Primary Creation call | `build_deck(full DeckSchema)` | C1/C2 done |
| Optional name/price hydrate | `confirm_mix` (legacy) | I3 / PI-2761 |
| Product slide identity | GTM Product Tags exact map | A5 done |
| Deck body (dividers + clones) | `assemble_skeleton` | C1 landed |
| Intro/narrative/investment/thanks fills | Placeholder pipeline | C2 **done** |
| Logic Guide engine in repo | Reference/tests only | I2 demoted |

There is no `propose_mix` MCP tool. `ingestion/logic_guide/` is reference/test code, not associate runtime.

C2 tests pass `confirmed_products` directly as if upstream already locked the mix.

---

## Progress (2026-09-09)

| ID | Ticket | Status | Notes |
|---|---|---|---|
| A1–A3 | PI-2517–2519 | Done | Clone/delete/insert + `build_deck` |
| C3 | PI-2758 | Done | `DiscoverySchema` + `DeckSchema` |
| A5 | PI-2541 | Done | Exact Deck Path / Slide #; merged #23 |
| C1 | PI-2756 | **Done** | FortuneAI spine, unfunded dividers dropped, A5 inserts; merged #24 |
| I1 | PI-2759 | Done | GTM + inventory + pricing sources |
| I2 | PI-2760 | **Reference only** | LogicGuideEngine — tests/fixtures, not runtime |
| I3 | PI-2761 | Optional legacy | `confirm_mix` when caller omits full products |
| C2 | PI-2757 | **Done** | Deterministic + bounded AI fills in `build()` |
| Upstream wire | PI-2350 | **Out of repo** | Pitch Deck Builder sends locked DeckSchema |
| Stylist | PI-2754 | Shelved → QA rail | Headless Cursor QA (see DECK-QA-ARCHITECTURE.md) |

**Creation rail:** `build_deck` with a fully hydrated `DeckSchema` produces a seller-readable PPTX. Live Claude + manual FortuneAI PPTX review recommended before prod deploy.

**Upstream rail:** Pitch Deck Builder / Sales HQ must send the complete payload — not rely on sales-mcp to ideate.

---

## C1 vs C2 (keep this straight)

- **C1** — Load FortuneAI. Drop unfunded category dividers. Paste A5 clones under funded chapters. Leave intro / audience / program / investment / thanks as stock.
- **C2** — Fill stock slots (deterministic + bounded Claude); pick one Audience page and one Program Overview page; fill investment/thanks. Do not rewrite product clones. Do not pick the mix.

---

## File index (this folder)

| File | Use |
|---|---|
| `END-SCOPE-SOT.md` | Canonical end state + Workflow/Logic Guide distillation |
| `../../docs/PRODIE-IDEATION-SPEC.md` | **Historical** Prodie ideation design — not sales-mcp runtime contract |
| `PROGRESS.md` | This file — goal + ticket status |
| `C2-PLACEHOLDER-INVENTORY.md` | C2 Chunk 0 token/shape locks |
| `PI-2757-TECH-DEBT.md` | C2 leftover gaps |
| `CLEANUP-TODO.md` | Titan/RAG delete later |
| `NEXT-STEPS.md` | Historical Phase A/B + retention; defer to SoT on conflict |
| `TICKETS.md` / `JIRA.md` / `README.md` | Historical A/B writeups; defer to SoT |
| `B2-PLAN.md` | Shelved stylist |
