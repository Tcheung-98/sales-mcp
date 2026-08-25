# Pitch Deck end-scope SoT

> **Captured:** 2026-08-17 · **Associate UX + progress restated:** 2026-08-18  
> **Split restated:** 2026-08-25 — Prodie **proposes relevant products** for associate selection; **deckgen** (`sales-mcp` `build_deck`) assembles the PPTX. sales-mcp does not choose the mix.  
> **Purpose:** Shared source of truth for project *end* scope after product delivered  
> Pitch Deck Workflow + Fortune Logic Guide V1. Prefer this over older Phase A/B RAG docs when they conflict.  
> **Product sources (read-only):**  
> - `~/Downloads/Pitch Deck Workflow.pdf` (docx twin referenced by GTM)  
> - `~/Downloads/Fortune_Logic_Guide_V1.pdf`  
> **Companion assets (owned outside this repo / by data+GTM):**  
> - `FortuneAI_DeckTemplate.pptx` — Creation spine (structure + stock/AI slides)  
> - `Fortune_AITool_GTM_Database` (+ Product Tags, GTM Tags, Audience Data, Product Category, …) — SharePoint / S3 sync  
> - `Fortune Inventory & Reservation Calendar 2026 Final` — availability + Pricing + Benchmarks  

Related: [`PROGRESS.md`](PROGRESS.md) · [`PRODIE-IDEATION-SPEC.md`](PRODIE-IDEATION-SPEC.md) (Prodie propose + select MVP) · [`CLEANUP-TODO.md`](CLEANUP-TODO.md) · [`NEXT-STEPS.md`](NEXT-STEPS.md) · Jira Creation [PI-2516](https://fortune.atlassian.net/browse/PI-2516) · Ideation [PI-2755](https://fortune.atlassian.net/browse/PI-2755) · Shelved stylist [PI-2754](https://fortune.atlassian.net/browse/PI-2754)

---

## One-liner end state

```text
Discovery (seller intake in Prodie)
  → Propose (Prodie: Logic Guide V1 + GTM + inventory → relevant, priced, available products)
  → Select (associate checkboxes — this is the mix)
  → Deckgen (sales-mcp build_deck: FortuneAI_DeckTemplate + exact GTM Deck Path/Slide # clones
       + deterministic placeholders)
  → PPTX out
```

**Associate experience (Prodie) — MVP:**

```text
Form (Discovery) ± SalesGPT conversation
  → Prodie shows relevant products (names + prices + flight/availability + short why)
  → Associate selects / swaps via checkboxes (running total / timelines)  (I3)
  → Only then: pass locked spec to sales-mcp build_deck  (C1 spine + A5 clones + C2 fills)
```

**Prodie does not build the deck.** It selects relevant products and collects the associate’s checks.  
**sales-mcp does not choose products.** It consumes the locked spec and clones slides.

There is **no `propose_mix` MCP tool**. In-repo `LogicGuideEngine` modules are isolated
reference/test code, not associate runtime. Guide **Media Mix Logic** (auto-fund a package)
is **not** required MVP; the associate’s selection is the mix.

Humans vet **products (offerings)**, not Hunter slide numbers. After lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #` and C1 pastes that page. C1/C2 must not invent, rank, or swap offerings. Prodie must not pick **slides** — only **product names** (plus category when names collide).

**Not the end state:** RAG/Titan similarity for product pages · Cursor vision stylist · industry Category_Presentation_* matrix as primary spine · Claude inventing product slide copy · generating a deck before the associate confirms the mix · Prodie auto-locking a funded Media Mix without checkboxes.

---

## What product provided (three pillars)

### 1. Pitch Deck Workflow — process + Creation recipe

Three mandatory stages:

| Stage | Job | Gate |
|---|---|---|
| **Discovery** | Capture seller inputs | Almost all fields required (only Platform/Product Specifics optional) |
| **Ideation** | Prodie proposes relevant placements using Logic Guide V1 + GTM + inventory; MCP does not rank | Seller must **select** mix **before** Creation |
| **Creation** | Deckgen assembles the pitch deck from the locked spec | Only after mix lock |

**Discovery inputs (required unless noted):** company name · industry (fixed enum) · budget (up to 3 tiers via “+”) · flight dates · campaign goal · targeting details · KPIs (enum) · KPI details · campaign narrative · preferred platforms/products (enum) · platform/product specifics *(optional)* · additional RFP details · *(Creation also needs client logo)*.

**Conference flag:** if Preferred Platforms includes Conference Sponsorship/Media → flag seller to reach GTM; do not silently invent conference slides.

**Creation deck anatomy (5 sections):**

| # | Section | Slide count | Method |
|---|---|---|---|
| 1 | Intro | 1 | Template slide 1 + AI title + logo + Month/Year |
| 2 | Narrative | 5 | Why Fortune (stock) · History of Trust (client name token) · Opportunity (AI) · Audience (2–6 card variant) · Program Overview (2–4 box variant) |
| 3 | Product Pitch | Variable | Category dividers (13–17) **only if funded** + **exact** product slide clones |
| 4 | Investment Summary | 1 | Template slide 18; repeating category blocks; per-product bullets; flag budget mismatch — never silent fix |
| 5 | Thank You | 1 | Template slide 19 stock + date + logo |

**Product page rule (hard):** after seller confirms mix → for each funded product look up GTM DB by exact name/category → read `Deck Path` + `Slide #` → copy that slide wholesale. **No AI rewrite. No similar-slide guess.** Missing row → stop and flag.

**Divider ↔ platform mapping (fixed order):**  
High-Impact Media (13) ← Digital Ads/Programmatic · Editorial Alignment (14) ← Newsletters · Premium Video (15) ← Vodcasts · Print (16) ← Print · Branded Content (17) ← Branded Content.

### 2. Fortune Logic Guide V1 — relevance policy (executed by Prodie)

The Guide tells Prodie **which products belong on the menu**. Prodie applies it **after** Discovery and **before** selection, using SharePoint / synced GTM + inventory as evidence. sales-mcp may expose **catalog, price, and availability** so Prodie does not invent rates; it must not choose the locked mix.

Policy the Guide still specifies for **relevance + availability** (do not silently drop):

1. **Availability** — drop sold/held vs flight dates via Inventory Calendar when the product is on the Products tab. Products not on that tab (e.g. non-takeover digital, branded content) have no inventory gate. Do not offer sold/held as default-checked pitch lines.
2. **Product Category Rules** — per-category candidate sets (Digital/Newsletters/Vodcasts/Branded Content/Print). Trigger phrases, defaults (Crown/Scroller, Full Page @$35K when Print selected), genuine-match tag rules, Branded Content video/written tracks. Independent of auto-funding a package.
3. **Explicit V1 exclusions (revisit later):** requirements-based branded content (spotlights, syndicated film, research support, events, lead gen, etc.), Digital Ad Creative, on-location video variants.

Guide **Media Mix Logic** (price → mandatory minimums → branded-content priority → Print → backfill → cross-category rules → lower-tier trim) is **reference**. Associate MVP does **not** require Prodie to lock that package. Checkboxes are the mix. The in-repo `LogicGuideEngine` “most expensive newsletter/vodcast wins” pass is **not** the picker.

### 3. Data / inventory SoT (data team + SharePoint)

| Asset | Role |
|---|---|
| `Fortune_AITool_GTM_Database` | GTM Tags (candidate pool) · Product Tags (`Deck Path` / `Slide #`) · Audience Data · Product Category names · product descriptions/tags |
| Inventory & Reservation Calendar 2026 Final | Availability during flight · Pricing + Benchmarks |
| `FortuneAI_DeckTemplate.pptx` | Intro/narrative/investment/thanks + dividers + Print slide 20 |
| Hunter product deck binaries | Targets named by `Deck Path` (synced to S3 `product-decks/` for runtime) |

This repo **consumes** synced copies (S3 keys / Graph fetch). It does **not** own inventory editing or SharePoint permissions.

---

## What this implies (architecture)

```text
Prodie  (form ± SalesGPT chat + Logic Guide V1 + GTM/inventory)
  │
  ├─ Discovery → DiscoverySchema
  │
  ├─ Propose ──reads──► Logic Guide V1 + GTM DB + Inventory/Pricing
  │     └─ relevant offerings → associate checkboxes → confirmed_products[]
  │
  └─ pass locked spec ──► sales-mcp build_deck
        (optional confirm_mix = validate names / availability / prices only)
        spine = FortuneAI_DeckTemplate (not Category_Presentation_*)
        product pages = exact Deck Path + Slide #   (A5 + C1)
        placeholders = deterministic + bounded AI fills   (C2)
```

| Implication | Detail |
|---|---|
| **Propose is Prodie’s job** | Form + conversation → relevant products → associate select (I3) → deckgen. MCP must not inject a mix. |
| **Build is deckgen’s job** | Clone + template + placeholder fill. Bounded AI only for named Workflow slots. Never rewrite product clones. |
| **RAG is not Creation-critical** | Titan/embeddings may linger for research tools; not the product-page selector. Prodie must not search Hunter decks for “similar slides.” |
| **Stylist (Phase B) is v2** | Shelved ([PI-2754](https://fortune.atlassian.net/browse/PI-2754)). Do not block MVP on Cursor vision. |
| **Schema vocab must converge** | Workflow platforms vs Logic Guide categories vs legacy `_VALID_CATEGORIES` (`Digital Media`/`Newsletter`/…) need aliases until unified. |
| **Data freshness is an ops dependency** | Wrong/outdated GTM xlsx or missing `product-decks/` binaries = loud failure, not silent quality loss. |

---

## Progress to this pivot (as of 2026-08-19)

Living ticket table: [`PROGRESS.md`](PROGRESS.md). Snapshot:

### Done / landed

| Piece | Evidence |
|---|---|
| Skeleton assembly path | A1 [PI-2517](https://fortune.atlassian.net/browse/PI-2517) Done · A2 [PI-2518](https://fortune.atlassian.net/browse/PI-2518) Done · A3 [PI-2519](https://fortune.atlassian.net/browse/PI-2519) Done |
| Discovery contract on schema | C3 [PI-2758](https://fortune.atlassian.net/browse/PI-2758) Done — `DiscoverySchema` + `DeckSchema` with Workflow fields, budget tiers, KPIs, platforms |
| Exact product clone map (code) | A5 [PI-2541](https://fortune.atlassian.net/browse/PI-2541) **Done** — merged [#23](https://github.com/Tcheung-98/sales-mcp/pull/23) |
| FortuneAI spine + conditional dividers | C1 [PI-2756](https://fortune.atlassian.net/browse/PI-2756) **Done** — merged [#24](https://github.com/Tcheung-98/sales-mcp/pull/24) |
| Placeholder pipeline (deterministic + bounded AI) | C2 [PI-2757](https://fortune.atlassian.net/browse/PI-2757) **Done** — [#25](https://github.com/Tcheung-98/sales-mcp/pull/25) + [#26](https://github.com/Tcheung-98/sales-mcp/pull/26) |
| Epic rename toward MVP | [PI-2516](https://fortune.atlassian.net/browse/PI-2516) = “Deterministic deck Creation (FortuneAI / Workflow MVP)” |
| Ideation epic filed | [PI-2755](https://fortune.atlassian.net/browse/PI-2755) + children I1–I3 |
| Stylist shelved | [PI-2754](https://fortune.atlassian.net/browse/PI-2754) On Hold |

### In flight / remaining

| Piece | Status |
|---|---|
| I1 GTM + inventory + pricing sources | [PI-2759](https://fortune.atlassian.net/browse/PI-2759) **Done** |
| I2 Logic Guide modules | [PI-2760](https://fortune.atlassian.net/browse/PI-2760) **Isolated** — no proposal MCP tool; not associate runtime |
| I3 seller confirm/swap | [PI-2761](https://fortune.atlassian.net/browse/PI-2761) MCP lock exists; Prodie checkbox UI still PI-2350 |
| Wire | [PI-2350](https://fortune.atlassian.net/browse/PI-2350) (+ [PI-2373](https://fortune.atlassian.net/browse/PI-2373)) **MVP remaining:** Prodie propose + select + pass spec to `build_deck` |

### Still legacy / demote (do not invest)

- ~~Industry `Category_Presentation_*` + franchise keyword `select_template()`~~ — retired at C1
- Titan `SlideEmbedder` / snapshot cosine as Creation matcher — remove after A5 merge (see CLEANUP-TODO)
- Claude blank-deck / Opus QA paths — dormant; delete when safe
- Phase B review packages / `cursor_sdk` — shelved

### Rough distance to Workflow end state

| Stage | Fit | Gap |
|---|---|---|
| Discovery | **Partial** | Schema enums + tiers landed in sales-mcp; Prodie still on old 6-field intake |
| Propose + select | **Not wired** | Guide-informed product list + checkboxes in Prodie (PI-2350) |
| Creation spine + fills | **Strong (C1 + C2)** | Live Claude + FortuneAI PPTX QA before prod |
| Product pages | **Strong** | Exact map path landed; coverage holes are GTM data issues |
| End-to-end seller flow | **Not yet** | Needs Prodie propose + select + `build_deck` handoff |

**Pivot verdict (2026-08-25):** Creation stays deterministic clone+fill. **Prodie MVP is a relevant-product menu + checkboxes.** Remaining work is (1) Prodie Discovery + propose + select, (2) pass locked spec to `build_deck`, (3) slim RAG/stylist-era weight.

---

## Ownership split

| Concern | Owner |
|---|---|
| Workflow / Logic Guide V1 / template structure | Product + GTM |
| GTM DB, inventory calendar, pricing, Hunter decks | Data / GTM (SharePoint SoT → sync to runtime) |
| Which products appear on the menu | **Prodie** (Logic Guide V1 + form/chat + GTM/inventory) |
| Which products go in the deck | **Associate checkboxes** |
| Creation assembly | **sales-mcp** `build_deck` |
| Optional name/availability validation | `confirm_mix` (not ranking) |
| Finished deck retention for humans | SharePoint (see retention notes in NEXT-STEPS) |

---

## Working rules for agents

1. Treat **Workflow + Logic Guide V1 + GTM sheets** as product SoT; older Phase B / RAG epic docs are historical unless explicitly revived.
2. Do not invent product slides or silently substitute near-matches. Prodie proposes **product names**; A5 maps slides.
3. Do not call `build_deck` before checkbox lock. The MCP exposes no product-proposal tool. Do not have Prodie assemble PPTX.
4. Prefer loud failures on missing Deck Path / inventory / price mismatch.
5. After A5 merges, cleanup Titan in a **separate** PR — don’t couple map swap with RAG deletion.
6. Conference Sponsorship → escalate to GTM; Branded Content V1 exclusions stay out until product revisits.
7. Do not put mix ranking (most-expensive-wins, Titan similarity) in Creation or as a hidden lock ahead of checkboxes.
