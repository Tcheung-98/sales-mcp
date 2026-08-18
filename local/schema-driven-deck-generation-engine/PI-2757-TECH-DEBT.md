# PI-2757 tech debt — placeholder pipeline

> **Ticket:** [PI-2757](https://fortune.atlassian.net/browse/PI-2757) / C2  
> **PR:** https://github.com/Tcheung-98/sales-mcp/pull/25  
> **Captured:** 2026-08-18  
> **Status:** Feature incomplete (Chunk 4 wired; Chunk 5 AI tokens still leftover). Do **not** treat pytest-green as “works on FortuneAI.”  
> **Related locks:** [`C2-PLACEHOLDER-INVENTORY.md`](C2-PLACEHOLDER-INVENTORY.md) · [`PROGRESS.md`](PROGRESS.md)

Do not spend a cycle rewriting fill logic until a live-template probe disagrees. Fail-loud is correct.

---

## Not debt (working as designed)

- Fill stock slots, then `delete_unused_variants`. Indices 0–11 stay valid because C1 only mutates from divider index 12.
- Investment / thanks via `slides[-2]` / `[-1]` after A5 clones.
- A5 product clones are not rewritten.
- Budget mismatch vs max tier (or label `/total/i`) — Chunk 0 lock.
- `[TITLE]`, Opportunity `[HEADER]`/`[BODY]`, `[AUDIENCE TITLE]`, program one-liners left for Chunk 5.

---

## Do not code yet (need a live probe)

These look like bugs from CI. They are unknowns about the live file, or already locked as follow-ons.

| Item | Why wait | Probe |
|---|---|---|
| Token / index mismatch | Inventory inspected S3 FortuneAI (2026-08-17); **fills have not been run on that PPTX**, only `fortuneai_placeholder_fixture`. Wrong tokens or variant counts → fail loud or fill the wrong slide. | Dump text on live slides 1–19; run `build()` against S3/SharePoint FortuneAI with injected `audience_data` + `logo_bytes` + a stub mix; open the PPTX. |
| `clone_shape_below` is `p:sp` only | Inventory says investment is one text box. A group/picture will fail loud (`missing shape id`). Don’t invent group clone until the live box disagrees. | Clone 2+ funded categories on the live investment slide. |
| Reach/Index cell types | Inventory: Reach is compact str (`1.1M`); Index is int. `data_only` + uncalculated formulas or numeric cells with Excel formats would dump raw numbers or skip rows. | Open GTM `Audience Data` and confirm cell types, not just the grid display. |
| Audience substring match | Chunk 0: case-insensitive known-name match; prose ignored; `<2` hits fail loud. Short names (`IT`, `CEO`) can false-positive. Apostrophe folding is in. Word boundaries are a product call after seeing real `targeting_details`. | Run `match_targeting` against a few real Discovery strings + the 52-row sheet. |
| C1 divider order vs live PPTX | Already flagged in inventory. Live 13–17 is Branded / Editorial / Premium Video / High-Impact / Print. **C1 follow-up, not a C2 fill.** Do not silently retitle dividers in C2. | After C1 caption-match (or GTM reorder), re-check High-Impact / Print / Branded mixes. |

---

## Known gaps (ticket-scoped, later chunks)

### 1. SharePoint / Graph logo fetch

README: `client_logo` is “URL or SharePoint path.” Inventory: HTTPS in CI; SharePoint `sites/` path via Graph if not http(s).

`fetch_logo_bytes` is HTTPS-only. Tests inject `logo_bytes`, so CI stays green. A real `build_deck` with a library path will raise.

**When:** Prodie wire (PI-2350) or a C2 follow-up once intake actually sends SharePoint paths. Same bytes on intro + thanks.

### 2. Chunk 5 — AI tokens

Still leftover after Chunk 4: intro `[TITLE]`, Opportunity `[HEADER]`/`[BODY]`, `[AUDIENCE TITLE]`, Program `Product description.` (except 1-category stock second box). Guide: 3–6 word titles; ~85 word body; no em dashes on AI fields.

### 3. Associate-facing audience failures

`targeting_details` that never names a GTM `Audience Segment` (or names only one) fail at `select_audience_variant`. That will be the first seller-visible error. Don’t soften to invent segments. Product may want a clearer message / form hint, not a matcher rewrite.

### 4. Warnings vs Prodie

`>6` audience and `>4` program return `warnings[]` on `build()`. MCP/Prodie must surface them. If the client drops unknown keys, the seller never sees “prioritize top 6.”

### 5. `tests/c2_fixture.py`

Local duplicate of `tests/fortuneai_placeholder_fixture.py` (ticket-chunk name). Do not commit. Delete when convenient.

---

## How to test at this phase (I2 / I3 / Prodie not required)

Engineer harness, not the form loop:

1. Load real `FortuneAI_DeckTemplate` from S3/SharePoint (`local/.../scratch/` has a copy).
2. Inject `audience_data=` and `logo_bytes=` (skip live GTM xlsx + HTTPS logo).
3. Pass `confirmed_products` that already exist on the GTM map (Editorial / CEO Daily is a valid path; High-Impact/Print/Branded still hit the C1 order mismatch).
4. `DeckGenerator.build(...)` and open the PPTX: intro date+logo, History client name, one audience page with verbatim Reach/Index, one program page, investment dollars + cloned category boxes, thank-you, product clone titles unchanged, unused audience/program pages gone, leftover `[` only on Chunk 5 AI slots.

---

## Hygiene if pushing mid-chunk

- Include `tests/test_placeholder_fills.py` (covers `apply_placeholders`; generator happy-path is not enough).
- Do not commit `tests/c2_fixture.py` or local scratch PPTX/xlsx.
