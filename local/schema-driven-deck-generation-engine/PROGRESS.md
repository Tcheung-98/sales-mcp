# Pitch deck — goal + progress

> **Living file.** Update when a ticket lands or the associate flow changes.  
> **End-state SoT (do not fork):** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)  
> **Agents:** `.cursor/rules/pitch-deck-end-scope.mdc` is the always-on snapshot of this + the SoT.  
> **Last updated:** 2026-08-25 (sequence: Discovery → later Prodie menu → later unbuilt select UI → `build_deck`)

---

## Goal

Associates **later** pick products from a **Prodie menu**, then deckgen builds a FortuneAI PPTX. They do not pick during Discovery. Select UI is **not built**. See [`SEQUENCE.md`](SEQUENCE.md).

```text
Form (Discovery) ± SalesGPT conversation
  → later: Prodie returns relevant products (Logic Guide V1 + GTM + inventory)
  → later: Associate checkboxes in SalesGPT UI (not yet implemented)   ← I3; gate
  → then: pass locked spec → build_deck (C1 spine + A5 clones + C2 fills)
```

**Prodie later returns relevant products.** It does not build the deck and does not lock a mix at Discovery time.  
**The associate’s later checks are the mix.**  
**sales-mcp `build_deck` assembles.** It does not rank products.

**Humans vet offerings (products).** They do not pick Hunter `Slide #`. After I3 lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #`; C1 pastes that page under the funded divider. C2 never changes those clones.

---

## Who owns which step

| Step | Owner | Ticket |
|---|---|---|
| Form / Discovery fields | Schema + Prodie/UI | C3 done; wire PI-2350 |
| Propose relevant offerings | **Prodie** (later; Logic Guide V1 + GTM/inventory) | PI-2350 / P3 |
| Associate select/swap | SalesGPT checkboxes (**not implemented**) | I3 / PI-2350 / P4 |
| Optional validate locked mix | `confirm_mix` | I3 MCP (not ranking) |
| Product slide identity | GTM Product Tags exact map | A5 done |
| Deck body (dividers + clones) | `assemble_skeleton` | C1 landed |
| Intro/narrative/investment/thanks fills | Placeholder pipeline | C2 **done** |
| Call `build_deck` with locked spec | Prodie handoff | PI-2350 |

In-repo `propose_mix` / `LogicGuideEngine` = **optional reference**, not associate MVP. Guide Media Mix auto-fund is **not** required.

C2 tests stub I3: pass `confirmed_products` as if the associate already chose.

---

## Progress (2026-08-25)

| ID | Ticket | Status | Notes |
|---|---|---|---|
| A1–A3 | PI-2517–2519 | Done | Clone/delete/insert + `build_deck` |
| C3 | PI-2758 | Done | `DiscoverySchema` + `DeckSchema` |
| A5 | PI-2541 | Done | Exact Deck Path / Slide #; merged #23 |
| C1 | PI-2756 | **Done** | FortuneAI spine, unfunded dividers dropped, A5 inserts; merged #24 |
| I1 | PI-2759 | Done | GTM + inventory + pricing sources |
| I2 | PI-2760 | **Demoted** | MCP mix engine is not associate MVP |
| I3 | PI-2761 | MCP lock done | Select UI still unbuilt (PI-2350 / P4) |
| C2 | PI-2757 | **Done** | Deterministic + bounded AI fills in `build()` |
| Wire | PI-2350 | **MVP remaining** | Discovery (P1) + catalog (P2) in flight; then menu + checkboxes + pass spec |
| Stylist | PI-2754 | Shelved | Not MVP |

**Creation rail** after C2: engineer can `build_deck` with a stubbed mix and get a seller-readable PPTX. Live Claude + manual FortuneAI PPTX review recommended before prod deploy.

**Associate rail** still needs later Prodie propose + later checkbox confirm (unbuilt UI) + `build_deck` handoff.

---

## C1 vs C2 (keep this straight)

- **C1** — Load FortuneAI. Drop unfunded category dividers. Paste A5 clones under funded chapters. Leave intro / audience / program / investment / thanks as stock.
- **C2** — Fill stock slots (deterministic + bounded Claude); pick one Audience page and one Program Overview page; fill investment/thanks. Do not rewrite product clones. Do not pick the mix.

---

## File index (this folder)

| File | Use |
|---|---|
| `SEQUENCE.md` | Sequential Discovery → menu → later select UI → `build_deck` |
| `END-SCOPE-SOT.md` | Canonical end state + Workflow/Logic Guide distillation |
| `PRODIE-IDEATION-SPEC.md` | Prodie propose + select + pass spec to deckgen |
| `PROGRESS.md` | This file — goal + ticket status |
| `C2-PLACEHOLDER-INVENTORY.md` | C2 Chunk 0 token/shape locks |
| `PI-2757-TECH-DEBT.md` | C2 leftover gaps |
| `CLEANUP-TODO.md` | Titan/RAG delete later |
| `NEXT-STEPS.md` | Historical Phase A/B + retention; defer to SoT on conflict |
| `TICKETS.md` / `JIRA.md` / `README.md` | Historical A/B writeups; defer to SoT |
| `B2-PLAN.md` | Shelved stylist |
