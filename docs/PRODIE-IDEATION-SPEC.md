# Prodie Ideation — requirements / spec (MVP)

> **Status:** Intended MVP (2026-08-25, split restated)  
> **SoT:** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)  
> **Deckgen:** `sales-mcp` `build_deck`  
> **Policy source:** Fortune Logic Guide V1 + GTM Product Tags + Inventory Calendar. Prodie uses these to **propose relevant, available products**. It does **not** build the PPTX.

This spec is the contract for Prodie. Matching lives in Prodie. Assembly lives in deckgen.

---

## 1. Goal

An associate describes a pitch (form ± SalesGPT). Prodie **selects relevant products** (Logic Guide V1 + GTM tags + inventory) and **shows them for selection**. The associate checks what to pitch. Prodie then **passes the locked spec** to the deckgen engine (`build_deck`). Deckgen clones FortuneAI + GTM `Deck Path` / `Slide #`.

**Success:** The PPTX contains exactly the products the associate checked — not a mix invented inside sales-mcp or `build_deck`.

Prodie does **not** run Media Mix spend-allocation as a required MVP step. Humans pick the mix. Prodie’s job is a relevant, priced, availability-aware **menu**.

---

## 2. System split

| Actor | Owns | Must not |
|---|---|---|
| **Associate** | Discovery inputs; final mix via checkboxes | Hunter `Deck Path` / `Slide #` |
| **Prodie** | Discovery capture; propose relevant products; selection UI; pass locked spec to deckgen | Assemble PPTX; pick slides; invent SKUs/rates; rewrite product-slide copy |
| **sales-mcp (deckgen)** | Validate schema if needed; `build_deck` clone + fills | Choose or rank which products to pitch |

There is **no `propose_mix` MCP tool**. In-repo `LogicGuideEngine` modules are isolated
reference/test code only; they are not loaded by the associate runtime.

---

## 3. Associate UX

### 3.1 Discovery (form)

Required Workflow fields (map 1:1 to `DiscoverySchema`):

| Form | Schema |
|---|---|
| Company name | `company_name` |
| Industry (enum) | `industry` — Technology, Professional Services, Healthcare, Financial Services, Energy, Lifestyle, Luxury |
| Budget (1–3 tiers, “+”) | `budgets[].amount` + optional `label` |
| Flight dates | `flight_dates.start` / `.end` |
| Campaign goal | `campaign_goal` |
| Targeting details | `targeting_details` |
| KPIs (enum, multi) | `kpis` — Brand Lift, Viewability, Awareness, Engagement, Lead Generation |
| KPI details | `kpi_details` |
| Campaign narrative | `campaign_narrative` |
| Preferred platforms/products (enum, multi) | `preferred_platforms_products` |
| Additional RFP / PRD details | `additional_rfp_details` |
| Client logo | `client_logo` (HTTPS URL or SharePoint path sales-mcp can fetch) |
| Platform/product specifics | `platform_or_product_specifics` (**only optional field**) |

**Preferred platforms enum:** Branded Content · Digital Ads/Programmatic · Newsletters · Vodcasts · Print · Lists & Rankings Sponsorship · Conference Sponsorship/Media.

**Conversation:** If the associate used SalesGPT chat first, Prodie may **pre-fill** the form from the thread. The submitted form is still the Discovery contract.

**Budget ≥ $750k** (max tier): surface GTM escalation; do not auto-build.

### 3.2 Propose (Prodie)

On submit, Prodie:

1. Reads Fortune Logic Guide V1 (policy for *which products are relevant*).
2. Reads GTM Database Product Tags (name, category, **GTM TAGS**; Deck Path is **not** shown) and Inventory Calendar (availability + Pricing). Prefer the same synced facts deckgen uses (S3 / catalog tools) so names and prices match Creation.
3. Applies Guide **category / relevance** rules to Discovery **and** conversation context. Drops sold/held vs flight when the product is on the Inventory Products tab (no inventory row → no availability gate).
4. Returns a **proposed product list** (and optional alternates), **not** a PPTX.

Show each line as:

- Product name  
- Category (human: Newsletters, Digital, etc.)  
- Rate (from inventory Pricing; verbatim text **and** numeric)  
- Flight fit: available / sold-held / no inventory gate  
- **Why proposed** (short): Guide rule and/or tags/phrases/platforms — so the associate can disagree  

Do **not** show Deck Path, Slide #, or file names.

If Conference Sponsorship/Media or Lists & Rankings is selected: **stop proposing a mix**, tell the associate to reach GTM, do not call `build_deck`.

Do **not** require Prodie to fund a Media Mix (mandatory minimums, backfill, cross-category spend). That Guide section may inform ranking/order of the list; the associate’s checkboxes are the mix.

### 3.3 Select (checkboxes) — I3

- Default-check Prodie’s recommended lines; associate can uncheck, check alternates, or add from a **search of GTM product names** (exact catalog, not slide search).
- Running **mix total** vs selected budget tier; warn on over/under (deckgen will fail loud on Investment mismatch).
- Timelines: selected products’ availability vs `flight_dates`.
- Primary action: **Lock mix & generate deck**.
- No generate until at least one product is checked (unless escalation path).
- Do not default-check sold-out / held lines.

### 3.4 Handoff to deckgen

After lock, Prodie passes **Discovery + `confirmed_products[]`** into sales-mcp:

- Preferred: `build_deck(deck_schema)` once the schema includes the locked products.  
- Optional: `confirm_mix` first if MCP still validates names / availability / prices — validation only, not ranking.

Surface MCP errors verbatim (missing GTM row, sold-out, bad logo, etc.). Do not silently drop a product.

---

## 4. How Prodie picks *relevant* products (Logic Guide V1)

The Guide is **policy for the menu**, not a requirement to auto-build a funded package.

**Order:**

1. **Availability** — drop sold/held vs flight when the product is on the Inventory Products tab. Products with no inventory row (many non-takeover digital, branded content) have **no** availability gate.
2. **Category pools** — only open pools for selected `preferred_platforms_products` (Digital Ads/Programmatic, Newsletters, Vodcasts, Branded Content, Print). Do not pitch Conferences/Lists.
3. **Category relevance** (independent of “fund a mix”):
   - **Newsletters / Vodcasts:** match Discovery + chat to Product Tags **GTM TAGS** (genuine match — meaningful overlap, not a single accidental substring). Prefer **fit to brief** over highest rate card. Multi-select is allowed; do **not** collapse to “most expensive wins.”
   - **Digital:** Guide defaults (Crown/Scroller unless seller excluded display); conditionals (livestream → in-banner; video creative → pre-roll; LinkedIn; social; key dates → takeovers). Display backfill SKUs are not relevance-defaults.
   - **Branded Content:** video vs written tracks from the brief; V1 exclusions stay out (spotlights, syndicated film, research support, events, lead gen, on-location variants, Digital Ad Creative).
   - **Print:** Full Page when Print is selected (Guide default ~$35K); do not invent other print SKUs until GTM has rows.

**Evidence Prodie must cite per line:** Guide section (or “default for selected platform”), tag/phrase hits.

**GTM TAGS** are a controlled vocabulary. Map **intent** (form + chat) onto those tags. Do not invent new tag values.

Media Mix Logic (mandatory minimums, branded-content priority, trim-to-tier) is **not** the associate MVP picker. If used at all, it may sort or default-check; it must not hide products the associate should still be able to check, and it must not replace checkbox lock.

---

## 5. Data Prodie may use

| Source | Use |
|---|---|
| Fortune Logic Guide V1 | Relevance policy |
| Pitch Deck Workflow | Process + Creation expectations (do not assemble the deck in Prodie) |
| `Fortune_AITool_GTM_Database` Product Tags | Name, category, GTM TAGS; existence check |
| Same workbook Audience Data | Optional targeting context; Creation still matches `targeting_details` for Reach/Index |
| Inventory Calendar | Availability + Pricing + Benchmarks |
| SalesGPT thread | Pre-fill Discovery; inform tag/intent matching |
| sales-mcp catalog tools (when available) | Same names/prices Creation will use |

**Do not use:** Titan/RAG slide search, `Category_Presentation_*` as the proposal spine, Hunter PPTX internals, inventing rates.

---

## 6. Spec passed to deckgen

Prodie calls `confirm_mix(discovery, selected_products)` after checkbox lock.

Each selected product contains:

- `name` — exact GTM Product Name
- `category` — optional GTM or schema category; required only when a name exists in multiple categories

Prodie does **not** send price, cadence, score, tier, swap, or add/drop instructions.
sales-mcp resolves price and cadence from the authoritative inventory workbook, validates
availability against Discovery flight dates, maps the GTM category to Creation, and returns:

- `deck_schema` — valid `DeckSchema` containing canonical `confirmed_products`
- `confirmed_products` — exact canonical lines
- `mix_total`
- `warnings` — budget mismatch notes (no silent price adjustment)

Unknown, ambiguous, duplicate, or unavailable names fail loud. Conference/Lists and ≥$750k
return escalation. Pass the returned `deck_schema` unchanged into `build_deck`.

Creation will: FortuneAI spine → drop unfunded dividers → clone exact slides → C2 fills (bounded Claude on **spine** slots only). Product pages are **not** rewritten.

---

## 7. Non-goals (MVP)

- Prodie picking `Slide #` or browsing product-deck PPTX  
- Prodie assembling or styling the PPTX  
- Prodie auto-funding a Media Mix as the locked package  
- Any MCP product-ranking/proposal tool
- Most-expensive-wins as the newsletter picker  
- AI copy on cloned product slides  
- Cursor stylist (PI-2754)  
- Auto-pitch Conferences / Lists  
- Inventing Print beyond GTM rows  
- Generating a deck before checkbox lock  

---

## 8. Acceptance criteria

- [ ] Associate can submit a complete Discovery form (chat pre-fill allowed).  
- [ ] Conference/Lists → escalation, no deck.  
- [ ] Proposed lines include name, category, price, availability, and a Guide/tag rationale.  
- [ ] Checkboxes update mix total and flight warnings live.  
- [ ] Locked mix is the only set of product pages in the PPTX (spot-check name on clones).  
- [ ] Missing GTM row / hold / price failure is shown; no silent substitute.  
- [ ] MCP tool list contains no product-ranking/proposal tool.  
- [ ] Production path does not have Prodie build the PPTX.  
- [ ] Logo URL works for C2 intro/thanks.  
- [ ] More than one newsletter can be locked if the associate checks them.  

---

## 9. Tickets

| Work | Ticket |
|---|---|
| Prodie form + relevant-product list + checkboxes + pass spec to `build_deck` | [PI-2350](https://fortune.atlassian.net/browse/PI-2350) (MVP remaining) |
| MCP mix lock / validation | [PI-2761](https://fortune.atlassian.net/browse/PI-2761) — flat `selected_products` lock implemented |
| Optional MCP catalog for prices/availability | I1 data already on S3; expose if Prodie should not scrape SharePoint for $ |
| In-repo Logic Guide engine | [PI-2760](https://fortune.atlassian.net/browse/PI-2760) — reference only |

---

## 10. Open points (don’t block the split)

1. `build_deck` directly vs `confirm_mix` then `build_deck`.  
2. Whether Prodie reads Logic Guide from SharePoint each time or a cached extract.  
3. How strictly Prodie must implement Guide category tables vs “Guide-informed LLM + cite evidence.” MVP allows LLM relevance **if** every line is cited and the associate can uncheck.  
4. Print “Full Page” Deck Path in GTM (`FortuneAI_DeckTemplate`) is a **data** bug — Prodie should not paper over it; deckgen must fail loud.  
