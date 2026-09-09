# Prodie Ideation — historical requirements / spec

> ### ⚠️ Not the sales-mcp runtime contract (2026-09-09)
>
> This document describes an **earlier design** where Prodie proposed a product menu and associates
> confirmed via checkboxes before calling deckgen. **That is not how sales-mcp works today.**
>
> **Current contract:** upstream (**Pitch Deck Builder**, **Sales HQ**, or any caller) sends a
> **complete locked `DeckSchema`** (Discovery + `confirmed_products[]` with name, category, price,
> cadence) directly to **`build_deck`**. sales-mcp validates, assembles (C1), fills (C2), and
> uploads PPTX. It does not ideate or propose products.
>
> **`confirm_mix`** remains as optional legacy validation/hydration when the caller omits full
> product fields — not the primary path.
>
> **Canonical runtime SoT:** [`local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md`](../local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md)  
> **Living progress:** [`local/schema-driven-deck-generation-engine/PROGRESS.md`](../local/schema-driven-deck-generation-engine/PROGRESS.md)

> **Original status:** Intended MVP (2026-08-25, split restated) — **superseded for sales-mcp scope**  
> **Deckgen:** `sales-mcp` `build_deck`  
> **Policy source (historical):** Fortune Logic Guide V1 + GTM Product Tags + Inventory Calendar — intended for upstream ideation systems, not sales-mcp runtime.

The sections below are retained as **reference** for upstream work (PI-2350) and Logic Guide policy. They do **not** describe obligations of the sales-mcp MCP server.

---

## 1. Goal (historical Prodie UX)

An associate describes a pitch (form ± SalesGPT). A **upstream ideation UI** (originally spec'd for Prodie) would select relevant products (Logic Guide V1 + GTM tags + inventory) and show them for selection. The associate checks what to pitch. The locked spec would then reach deckgen (`build_deck`).

**Success (still true for deckgen):** The PPTX contains exactly the products in `confirmed_products[]` — not a mix invented inside sales-mcp or `build_deck`.

---

## 2. System split (historical vs current)

| Actor | Historical design | **Current sales-mcp role** |
|---|---|---|
| **Associate / upstream** | Discovery inputs; final mix via checkboxes | Sends complete `DeckSchema` to `build_deck` |
| **Prodie (historical)** | Discovery capture; propose products; selection UI | **Not ideation brain on MCP** — transport/client only if it forwards locked payload |
| **sales-mcp (deckgen)** | Validate schema if needed; `build_deck` clone + fills | **Deck generation only** — validate, C1, C2, upload |

There is **no `propose_mix` MCP tool**. In-repo `LogicGuideEngine` modules are isolated reference/test code only.

---

## 3. Associate UX (historical — upstream / PI-2350)

Sections 3.1–3.4 below describe the **Pitch Deck Builder / Sales HQ** experience that was originally assigned to Prodie. Implement there — not in sales-mcp.

### 3.1 Discovery (form)

Required Workflow fields (map 1:1 to `DiscoverySchema`) — unchanged; see [`README.md`](../README.md) field table.

### 3.2 Propose (upstream ideation — not sales-mcp)

On submit, an upstream system would apply Logic Guide V1 + GTM + inventory to return a proposed product list. sales-mcp does **not** expose this.

### 3.3 Select (checkboxes — upstream)

Associate confirms mix upstream. Locked `confirmed_products[]` arrives at `build_deck`.

### 3.4 Handoff to deckgen (current primary path)

**Preferred:** `build_deck(deck_schema)` once the schema includes locked products with `name`, `category`, `price`, `cadence`.

**Optional legacy:** `confirm_mix(discovery, selected_products)` when the caller sends only `[{name, category?}, ...]` — validation + price/cadence hydration only.

Surface MCP errors verbatim (missing GTM row, sold-out, bad logo, etc.). Do not silently drop a product.

---

## 4. How Logic Guide V1 informs relevance (upstream reference)

The Guide is **policy for product menus**, executed by upstream systems — not by sales-mcp at runtime.

Retained detail: availability gates, category pools, GTM TAGS matching, Branded Content tracks, Print defaults, Conference/Lists escalation, Media Mix as reference-only.

See original sections in git history for full Guide tables. Agents implementing **upstream** ideation should read Fortune Logic Guide V1 directly.

---

## 5. Data sources (shared — used by build_deck and confirm_mix)

| Source | sales-mcp use |
|---|---|
| `Fortune_AITool_GTM_Database` Product Tags | A5 exact clone map |
| Same workbook Audience Data | C2 audience Reach/Index |
| Inventory Calendar | Optional `confirm_mix` availability/pricing hydrate |
| Fortune Logic Guide V1 | Upstream ideation reference only |

**Do not use in Creation path:** Titan/RAG slide search, `Category_Presentation_*` as spine, inventing rates.

---

## 6. Spec passed to deckgen (current)

**Primary:** full `DeckSchema` with `confirmed_products[]` — each entry includes `name`, `category`, `price`, `cadence`.

**Legacy:** `confirm_mix(discovery, selected_products)` → `deck_schema` when caller omits prices/cadences.

Creation: FortuneAI spine → drop unfunded dividers → clone exact slides → C2 fills (bounded Claude on **spine** slots only). Product pages are **not** rewritten.

---

## 7. Non-goals (sales-mcp)

- sales-mcp ideating or proposing products  
- sales-mcp picking `Slide #` or browsing product-deck PPTX  
- Any MCP product-ranking/proposal tool  
- AI copy on cloned product slides  
- Auto-pitch Conferences / Lists inside deckgen  

---

## 8. Acceptance criteria (split by owner)

**sales-mcp (deckgen) — done / in repo:**

- [x] `build_deck(full DeckSchema)` assembles FortuneAI + exact clones + C2 fills  
- [x] Missing GTM row / hold / price failure fails loud  
- [x] MCP tool list contains no product-ranking/proposal tool  
- [x] Logo URL works for C2 intro/thanks  

**Upstream (PI-2350) — out of repo:**

- [ ] Associate can submit Discovery and lock a product mix  
- [ ] Conference/Lists → escalation before calling `build_deck`  
- [ ] Locked mix is the only set of product pages in the PPTX  

---

## 9. Tickets

| Work | Ticket |
|---|---|
| Pitch Deck Builder / upstream product lock + `build_deck` handoff | [PI-2350](https://fortune.atlassian.net/browse/PI-2350) |
| Optional MCP mix hydrate / validation | [PI-2761](https://fortune.atlassian.net/browse/PI-2761) — legacy `confirm_mix` |
| In-repo Logic Guide engine | [PI-2760](https://fortune.atlassian.net/browse/PI-2760) — reference/tests only |

---

## 10. Resolved (2026-09-09)

1. **Primary path:** `build_deck(full DeckSchema)` — not `confirm_mix` first.  
2. Logic Guide execution belongs **upstream**, not sales-mcp MCP runtime.  
3. Prodie through this MCP is **transport only**, not ideation.
