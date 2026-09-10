# Pitch Deck end-scope SoT

> **Captured:** 2026-08-17 · **Scope corrected:** 2026-09-09 — **sales-mcp is deck generation only.** Upstream sends a complete locked `DeckSchema`; this repo validates, assembles (C1), fills (C2), and uploads PPTX.  
> **Purpose:** Shared source of truth for project *end* scope after product delivered  
> Pitch Deck Workflow + Fortune Logic Guide V1. Prefer this over older Phase A/B RAG docs when they conflict.  
> **Product sources (read-only):**  
> - `~/Downloads/Pitch Deck Workflow.pdf` (docx twin referenced by GTM)  
> - `~/Downloads/Fortune_Logic_Guide_V1.pdf`  
> **Companion assets (owned outside this repo / by data+GTM):**  
> - `FortuneAI_DeckTemplate.pptx` — Creation spine (structure + stock/AI slides)  
> - `Fortune_AITool_GTM_Database` (+ Product Tags, GTM Tags, Audience Data, Product Category, …) — SharePoint / S3 sync  
> - `Fortune Inventory & Reservation Calendar 2026 Final` — availability + Pricing + Benchmarks  

Related: [`PROGRESS.md`](PROGRESS.md) · [`PRODIE-IDEATION-SPEC.md`](PRODIE-IDEATION-SPEC.md) (**historical** Prodie ideation — not sales-mcp runtime contract) · Jira Creation [PI-2516](https://fortune.atlassian.net/browse/PI-2516)

---

## One-liner end state

```text
Upstream (Pitch Deck Builder / Sales HQ / caller)
  → complete DeckSchema (Discovery + confirmed_products[] with name/category/price/cadence)
  → sales-mcp build_deck (validate + FortuneAI_DeckTemplate + exact GTM Deck Path/Slide # clones
       + deterministic placeholders [+ optional QA rail])
  → PPTX out
```

**sales-mcp contract:**

```text
Locked DeckSchema payload
  → build_deck validates Pydantic + business rules
  → C1 assemble_skeleton (FortuneAI spine + conditional dividers + A5 exact clones)
  → C2 apply_placeholders (deterministic + bounded Claude on named stock slots)
  → [optional] deck QA rail (B2–B4)
  → presigned PPTX URL
```

**sales-mcp does not choose products.** It consumes the locked spec and clones slides.  
**Prodie is not the ideation brain on this MCP.** If Prodie appears, it forwards an already-locked payload — it does not propose a menu or run checkboxes through sales-mcp.

There is **no `propose_mix` MCP tool**. **`confirm_mix`** is optional legacy when the caller sends only name+category and needs price/cadence hydration from GTM.

Humans vet **products (offerings)**, not Hunter slide numbers. After upstream lock, A5 maps each confirmed name/category to `Deck Path` + `Slide #` and C1 pastes that page. C1/C2 must not invent, rank, or swap offerings.

**Not the end state:** sales-mcp ideating or proposing products · industry Category_Presentation_* matrix as primary spine · Claude rewriting product slide copy · Prodie auto-locking a funded Media Mix through this MCP.

---

## What product provided (three pillars)

### 1. Pitch Deck Workflow — process + Creation recipe

Three mandatory stages (ownership split across systems):

| Stage | Job | Owner (MVP) |
|---|---|---|
| **Discovery** | Capture seller inputs | Upstream (Pitch Deck Builder / Sales HQ) |
| **Ideation / mix lock** | Seller confirms which products to pitch | Upstream — **not** sales-mcp |
| **Creation** | Deckgen assembles the pitch deck from the locked spec | **sales-mcp** `build_deck` |

**Discovery inputs (required unless noted):** company name · industry (fixed enum) · budget (up to 3 tiers via “+”) · flight dates · campaign goal · targeting details · KPIs (enum) · KPI details · campaign narrative · preferred platforms/products (enum) · platform/product specifics *(optional)* · additional RFP details · *(Creation also needs client logo)*.

**Confirmed products (Creation):** each entry needs `name`, `category`, `price`, `cadence` — sent in `confirmed_products[]` on the primary path.

**Conference flag:** if Preferred Platforms includes Conference Sponsorship/Media → flag seller to reach GTM; do not silently invent conference slides.

**Creation deck anatomy (5 sections):**

| # | Section | Slide count | Method |
|---|---|---|---|
| 1 | Intro | 1 | Template slide 1 + AI title + logo + Month/Year |
| 2 | Narrative | 5 | Why Fortune (stock) · History of Trust (client name token) · Opportunity (AI) · Audience (2–6 card variant) · Program Overview (2–4 box variant) |
| 3 | Product Pitch | Variable | Category dividers (13–17) **only if funded** + **exact** product slide clones |
| 4 | Investment Summary | 1 | Template slide 18; repeating category blocks; per-product bullets; flag budget mismatch — never silent fix |
| 5 | Thank You | 1 | Template slide 19 stock + date + logo |

**Product page rule (hard):** after upstream confirms mix → for each funded product look up GTM DB by exact name/category → read `Deck Path` + `Slide #` → copy that slide wholesale. **No AI rewrite. No similar-slide guess.** Missing row → stop and flag.

**Divider ↔ platform mapping (fixed order):**  
High-Impact Media (13) ← Digital Ads/Programmatic · Editorial Alignment (14) ← Newsletters · Premium Video (15) ← Vodcasts · Print (16) ← Print · Branded Content (17) ← Branded Content.

### 2. Fortune Logic Guide V1 — reference policy (not sales-mcp runtime)

The Guide describes **which products are relevant** for a brief. Upstream systems (Pitch Deck Builder, Sales HQ) may use it when building the product menu — **sales-mcp does not execute Logic Guide ideation at runtime.**

Policy the Guide still specifies (useful for upstream + tests):

1. **Availability** — drop sold/held vs flight dates via Inventory Calendar when the product is on the Products tab. Products not on that tab (e.g. non-takeover digital, branded content) have no inventory gate.
2. **Product Category Rules** — per-category candidate sets (Digital/Newsletters/Vodcasts/Branded Content/Print). Trigger phrases, defaults, genuine-match tag rules, Branded Content video/written tracks.
3. **Explicit V1 exclusions:** requirements-based branded content, Digital Ad Creative, on-location video variants, Conferences/Lists auto-pitch.

Guide **Media Mix Logic** (auto-fund a package) is **reference only**. The locked `confirmed_products[]` is the mix.

### 3. Data / inventory SoT (data team + SharePoint)

| Asset | Role |
|---|---|
| `Fortune_AITool_GTM_Database` | GTM Tags · Product Tags (`Deck Path` / `Slide #`) · Audience Data · Product Category names |
| Inventory & Reservation Calendar 2026 Final | Availability during flight · Pricing + Benchmarks |
| `FortuneAI_DeckTemplate.pptx` | Intro/narrative/investment/thanks + dividers + Print slide 20 |
| Hunter product deck binaries | Targets named by `Deck Path` (synced to S3 `product-decks/` for runtime) |

This repo **consumes** synced S3 snapshots. It does **not** own inventory editing or SharePoint permissions.

---

## What this implies (architecture)

```text
Upstream (Pitch Deck Builder / Sales HQ / MCP client)
  │
  └─ locked DeckSchema ──► sales-mcp build_deck
        (optional confirm_mix = legacy validate/hydrate when products incomplete)
        spine = FortuneAI_DeckTemplate (not Category_Presentation_*)
        product pages = exact Deck Path + Slide #   (A5 + C1)
        placeholders = deterministic + bounded AI fills   (C2)
        optional QA rail = B2 review package → B3 deterministic → B4 headless Cursor
```

| Implication | Detail |
|---|---|
| **Mix lock is upstream** | Discovery + `confirmed_products[]` arrive together on the primary path. MCP must not inject a mix. |
| **Build is deckgen’s job** | Clone + template + placeholder fill. Bounded AI only for named Workflow slots. Never rewrite product clones. |
| **QA rail** | Headless Cursor QA on every `build_deck`. See `docs/DECK-QA-ARCHITECTURE.md`. |
| **Schema vocab must converge** | Workflow platforms vs Logic Guide categories vs legacy `_VALID_CATEGORIES` need aliases until unified. |
| **Data freshness is an ops dependency** | Wrong/outdated GTM xlsx or missing `product-decks/` binaries = loud failure, not silent quality loss. |

---

## Progress snapshot (see PROGRESS.md for living table)

### Done / landed

| Piece | Evidence |
|---|---|
| Skeleton assembly path | A1–A3 Done |
| Discovery contract on schema | C3 Done — `DiscoverySchema` + `DeckSchema` |
| Exact product clone map | A5 Done |
| FortuneAI spine + conditional dividers | C1 Done |
| Placeholder pipeline | C2 Done |
| GTM + inventory sources | I1 Done |
| Optional legacy validate/hydrate | I3 `confirm_mix` exists |
| Deck QA rail | B2–B4 always-on in `build_deck` |
| RAG / Titan / Logic Guide cleanup | Done 2026-09-10 |

### Remaining (mostly out of repo)

| Piece | Status |
|---|---|
| Upstream Discovery + product lock UI | PI-2350 Pitch Deck Builder / Sales HQ |
| End-to-end seller flow | Needs upstream to send locked DeckSchema to `build_deck` |

**Pivot verdict (2026-09-09):** Creation stays deterministic clone+fill in sales-mcp. **Remaining associate UX is upstream** (PI-2350). sales-mcp primary contract is `build_deck(full DeckSchema)`.

---

## Ownership split

| Concern | Owner |
|---|---|
| Workflow / Logic Guide V1 / template structure | Product + GTM |
| GTM DB, inventory calendar, pricing, Hunter decks | Data / GTM (SharePoint SoT → sync to runtime) |
| Discovery + product lock (menu, checkboxes, mix) | **Upstream** (Pitch Deck Builder / Sales HQ) — **not sales-mcp** |
| Creation assembly | **sales-mcp** `build_deck` |
| Optional name/availability/price hydrate | `confirm_mix` (legacy — prefer full payload) |
| Finished deck retention for humans | SharePoint (see retention notes in NEXT-STEPS) |

---

## Working rules for agents

1. Treat **Workflow + GTM sheets** as product SoT for Creation; older Phase B / RAG epic docs are historical unless explicitly revived.
2. Do not invent product slides or silently substitute near-matches. Upstream locks **product names**; A5 maps slides.
3. Primary path: `build_deck(full DeckSchema)`. The MCP exposes no product-proposal tool. Do not have sales-mcp ideate or assemble menus.
4. Prefer loud failures on missing Deck Path / inventory / price mismatch.
5. Conference Sponsorship → escalate to GTM; Branded Content V1 exclusions stay out until product revisits.
