# Pitch Deck end-scope SoT

> **Captured:** 2026-08-17 · **Associate UX + progress restated:** 2026-08-18  
> **Purpose:** Shared source of truth for project *end* scope after product delivered  
> Pitch Deck Workflow + Fortune Logic Guide V1. Prefer this over older Phase A/B epic framing when they conflict.  
> **Product sources (read-only):**  
> - `~/Downloads/Pitch Deck Workflow.pdf` (docx twin referenced by GTM)  
> - `~/Downloads/Fortune_Logic_Guide_V1.pdf`  
> **Companion assets (owned outside this repo / by data+GTM):**  
> - `FortuneAI_DeckTemplate.pptx` — Creation spine (structure + stock/AI slides)  
> - `Fortune_AITool_GTM_Database` (+ Product Tags, GTM Tags, Audience Data, Product Category, …) — SharePoint / S3 sync  
> - `Fortune Inventory & Reservation Calendar 2026 Final` — availability + Pricing + Benchmarks  

Related: [`PROGRESS.md`](PROGRESS.md) (living goal + ticket status) · [`CLEANUP-TODO.md`](CLEANUP-TODO.md) · [`NEXT-STEPS.md`](NEXT-STEPS.md) · Jira Creation [PI-2516](https://fortune.atlassian.net/browse/PI-2516) · Ideation [PI-2755](https://fortune.atlassian.net/browse/PI-2755) · Shelved stylist [PI-2754](https://fortune.atlassian.net/browse/PI-2754)

---

## One-liner end state

```text
Discovery (seller intake)
  → Ideation (Logic Guide + GTM DB + inventory/pricing → proposed mix → seller confirm)
  → Creation (FortuneAI_DeckTemplate + exact GTM Deck Path/Slide # clones + deterministic placeholders)
  → PPTX out (SharePoint for humans; S3 optional machine URL)
```

**Associate experience (SalesGPT / Prodie) — this is in spec, not a new idea:**

```text
Form (Discovery) → requirements / spec / PRD on DiscoverySchema
  → SalesGPT Ideation brain proposes offerings (Logic Guide + GTM + inventory/pricing)
  → Associate selects / swaps the specific options they will pitch  (I3)
  → Only then: Creation / build_deck  (C1 spine + A5 clones + C2 fills)
```

Humans vet **products (offerings)**, not Hunter slide numbers. After lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #` and C1 pastes that page. C1/C2 must not invent or swap offerings.

**Not the end state:** RAG/Titan similarity for product pages · Cursor vision stylist · industry Category_Presentation_* matrix as primary spine · Claude inventing product slide copy · generating a deck before the associate confirms the mix.

---

## What product provided (three pillars)

### 1. Pitch Deck Workflow — process + Creation recipe

Three mandatory stages:

| Stage | Job | Gate |
|---|---|---|
| **Discovery** | Capture seller inputs | Almost all fields required (only Platform/Product Specifics optional) |
| **Ideation** | Choose placements that fit budget, flight, inventory, Logic Guide | Seller must confirm/swap mix **before** Creation |
| **Creation** | Assemble the pitch deck | Only after mix lock |

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

### 2. Fortune Logic Guide V1 — Ideation brain

Runs **after** Discovery, **before** Creation:

1. **Availability first** — drop sold/held vs flight dates via Inventory Calendar. Products not on the Products tab (e.g. non-takeover digital, branded content) have no inventory gate.
2. **Product Category Rules** — per-category candidate sets (Digital/Newsletters/Vodcasts/Branded Content/Print). Independent of budget. Trigger phrases, defaults (Crown/Scroller, Full Page @$35K when Print selected), genuine-match tag rules, Branded Content video/written tracks + caps.
3. **Media Mix Logic** — price candidates → fund mandatory minimums in order → Branded Content priority → Print → remaining Digital conditionals → backfill → cross-category conflict rules → format collision (include both).
4. **Multi-tier** — build largest tier first; lower tiers are **strict subsets** via trim (not re-run category rules); backfill Display to approach lower target; “don’t zero a category” does **not** apply to lower tiers.

**Explicit V1 exclusions (revisit later):** requirements-based branded content (spotlights, syndicated film, research support, events, lead gen, etc.), Digital Ad Creative, on-location video variants.

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
Prodie / seller UI  (associate form + confirm)
  │
  ├─ Discovery → DiscoverySchema (intake form = requirements/spec)
  │
  ├─ Ideation engine (Logic Guide / SalesGPT brain) ──reads──► GTM DB + Inventory/Pricing
  │     └─ proposed offerings → associate select/swap → confirmed_products[]
  │
  └─ Creation (sales-mcp)  — only after I3 lock
        build_deck(DeckSchema, template_url?)
          spine = FortuneAI_DeckTemplate (not Category_Presentation_*)
          product pages = exact Deck Path + Slide #   (A5 + C1)
          placeholders = deterministic + bounded AI fills   (C2)
```

| Implication | Detail |
|---|---|
| **Ideation is a first-class product surface** | Not a Prodie-only chat habit. Form → proposed offerings → associate select (I3) → Creation. Needs I1 data + I2 engine + I3 confirm. |
| **Creation is mostly deterministic** | Clone + template select + placeholder fill. Bounded AI only for named Workflow slots (title, Opportunity body, audience titles, program overview blurbs). |
| **RAG is not Creation-critical** | Titan/embeddings may linger for research tools; not the product-page selector. |
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

### In flight

| Piece | Status |
|---|---|
| I1 GTM + inventory + pricing sources | [PI-2759](https://fortune.atlassian.net/browse/PI-2759) **In Progress** |

### Not started (MVP remaining)

| ID | Ticket | Owns |
|---|---|---|
| I2 | [PI-2760](https://fortune.atlassian.net/browse/PI-2760) | Logic Guide engine (category rules + media mix + tiers) |
| I3 | [PI-2761](https://fortune.atlassian.net/browse/PI-2761) | Seller confirm/swap → lock `confirmed_products` |
| Wire | [PI-2350](https://fortune.atlassian.net/browse/PI-2350) (+ [PI-2373](https://fortune.atlassian.net/browse/PI-2373)) | Prodie access / tool wiring |

### Still legacy / demote (do not invest)

- ~~Industry `Category_Presentation_*` + franchise keyword `select_template()`~~ — retired at C1 (`prepare_deck` removed; FortuneAI only)
- Titan `SlideEmbedder` / snapshot cosine as Creation matcher — remove after A5 merge (see CLEANUP-TODO)
- Claude blank-deck / Opus QA paths — dormant; delete when safe
- Phase B review packages / `cursor_sdk` — shelved

### Rough distance to Workflow end state

| Stage | Fit | Gap |
|---|---|---|
| Discovery | **Strong** | Schema enums + tiers landed; Prodie/UI intake + logo plumbing still wiring |
| Ideation | **Weak → building** | Spec clear; engine + data + confirm loop unfinished |
| Creation spine + fills | **Strong (C1 + C2)** | FortuneAI + placeholders landed; live Claude manual QA before prod |
| Product pages | **Strong** | Exact map path landed; coverage holes are GTM data issues |
| End-to-end seller flow | **Not yet** | Needs I2+I3+Prodie |

**Pivot verdict:** Directionally aligned — no second architecture rewrite. Remaining work is (1) implement Ideation as specified, (2) retarget Creation to FortuneAI_DeckTemplate + Workflow fill rules, (3) slim RAG/stylist-era weight after A5.

---

## Ownership split

| Concern | Owner |
|---|---|
| Workflow / Logic Guide / template structure | Product + GTM |
| GTM DB, inventory calendar, pricing, Hunter decks | Data / GTM (SharePoint SoT → sync to runtime) |
| Ideation engine + Creation assembly MCP | This repo (`sales-mcp`) + Prodie orchestration |
| Seller UX confirm/swap | Prodie / AE surface (I3) |
| Finished deck retention for humans | SharePoint (see retention notes in NEXT-STEPS) |

---

## Working rules for agents

1. Treat **Workflow + Logic Guide + GTM sheets** as product SoT; older Phase B / RAG epic docs are historical unless explicitly revived.
2. Do not invent product slides or silently substitute near-matches.
3. Do not fund Creation before a seller-confirmed mix (I3 gate). Form → proposed offerings → associate select → then `build_deck`.
4. Prefer loud failures on missing Deck Path / inventory / price mismatch.
5. After A5 merges, cleanup Titan in a **separate** PR — don’t couple map swap with RAG deletion.
6. Conference Sponsorship → escalate to GTM; Branded Content V1 exclusions stay out until product revisits.
|
