# I1 — Ideation data sources (access + refresh)

> **Ticket:** [PI-2759](https://fortune.atlassian.net/browse/PI-2759) · **Epic:** [PI-2755](https://fortune.atlassian.net/browse/PI-2755)  
> **Status:** Chunk C landed (inventory by flight); Chunk D pending  
> **End-state SoT:** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)

Ideation (Logic Guide / SalesGPT brain) reads three SharePoint-owned assets. Creation already consumes part of the GTM workbook; I1 makes the **full Ideation surface** machine-readable and documents how data stays fresh.

```text
SharePoint (human SoT)  →  sync  →  S3 snapshot  →  sales-mcp loaders  →  I2 engine
```

This repo **consumes** snapshots. It does **not** edit inventory, pricing, or GTM rows.

---

## Assets

| Asset | SharePoint SoT | Runtime access (sales-mcp) | Ideation use |
|---|---|---|---|
| `Fortune_AITool_GTM_Database.xlsx` | GTM / data team *(path below)* | S3 `GTM_DATABASE_KEY` | GTM Tags, Product Category, product catalog; also Product Tags + Audience Data (Creation) |
| `Fortune Inventory & Reservation Calendar 2026 Final.xlsx` | DPS — Sales Library on Dream Team site | S3 `INVENTORY_CALENDAR_KEY` *(Chunk C)* | Products tab → availability by flight; Pricing + Benchmarks → funding math |
| Hunter product decks | SharePoint GTM library | S3 `PRODUCT_DECKS_PREFIX` | Creation only (A5 clones); not an I1 loader |

**Not Google Tag Manager.** “GTM” here = Fortune’s internal product/GTM mapping workbook.

---

## SharePoint locations

### Inventory & Reservation Calendar 2026 Final

| Field | Value |
|---|---|
| Site | `fortunemg.sharepoint.com/sites/dream_team` |
| Library | Sales Library |
| File | `Fortune Inventory & Reservation Calendar 2026 Final.xlsx` |
| Direct link | [Open in SharePoint](https://fortunemg.sharepoint.com/sites/dream_team/_layouts/15/Doc.aspx?sourcedoc=%7B4BBE7984-98EF-4C07-ACCE-662E390A7207%7D&file=Fortune%20Inventory%20%26%20Reservation%20Calendar%202026%20Final.xlsx) |

**Prodie today:** SalesHQ MCP tools read this file live via Microsoft Graph (service account). There is **no** automated S3 snapshot for sales-mcp yet — Chunk C will assume manual or scripted upload to `INVENTORY_CALENDAR_KEY` until an ops pipeline exists.

### Fortune_AITool_GTM_Database

| Field | Value |
|---|---|
| Site | `fortunemg.sharepoint.com/sites/dream_team` *(expected; confirm with GTM)* |
| File | `Fortune_AITool_GTM_Database.xlsx` |
| Indexed SharePoint URL | **Not confirmed** in SalesHQ whitelist or Confluence as of 2026-08-24 |

**Creation today:** A5 + C2 load **Product Tags** and **Audience Data** from the S3 object at `GTM_DATABASE_KEY` (default `templates/Fortune_AITool_GTM_Database.xlsx`). Confirm the live SharePoint path with GTM/data team and keep S3 in sync before deploy.

---

## Workbook tabs (I1 contract)

### GTM database (`GTM_DATABASE_KEY`)

| Sheet | Consumer today | I1 / I2 use |
|---|---|---|
| **Product Tags** | A5 Creation (`gtm_product_map.py`) | Deck Path + Slide #; **GTM TAGS** column = Logic Guide candidate pool (Chunk B) |
| **Audience Data** | C2 Creation (`audience_data.py`) | Reach / Index (unchanged) |
| **Product Category** | Chunk B (`gtm_ideation_catalog.py`) | Category display titles + descriptions |

Additional tabs may exist; I1 loaders only depend on columns documented in Chunks B–D.

### Inventory calendar (`INVENTORY_CALENDAR_KEY`)

| Sheet | I1 / I2 use |
|---|---|
| **Products** | Master list of **inventory-gated** placements (`Product / Placement`, type, cadence, launch, weekday flags). Not on this tab → **no inventory gate**. |
| **Inventory** | Dated grid: `Status` = Available / Held / Sold / Holiday per product × date. Cross-reference seller `flight_dates` here. |
| **Pricing + Benchmarks** | Rates for Logic Guide funding (Chunk D) |

**Chunk C loaders:** `ingestion/inventory_calendar.py` — `InventoryProductRegistry`, `InventoryAvailability`, `InventoryCalendar.check_inventory_gate()`.

SOV tabs (*Lists Availability*, *Conference Media Availability*) are not loaded in Chunk C; daily/grid rows on **Inventory** cover most Ideation candidates.

| Products / Inventory columns (live workbook) |
|---|
| Products: `Product / Placement`, `Product Type`, `Cadence`, `Launch`, `Mon`–`Sun` |
| Inventory: `Date`, `Product / Placement`, `Product Type`, `Status`, … |

---

## Runtime access path (decision)

| Path | When to use | Notes |
|---|---|---|
| **S3 snapshot** *(default for sales-mcp)* | Lightsail MCP, local dev, CI | Same pattern as `GTM_DATABASE_KEY` + `FORTUNEAI_TEMPLATE_KEY`. Fail loud if object missing or stale. |
| **Microsoft Graph** | Prodie SalesHQ tools today (inventory) | Requires `GRAPH_*` + `SHAREPOINT_SITE_ID`. Not wired into Ideation loaders in this repo yet. |
| **Live SharePoint URL in request** | `build_deck` template only | FortuneAI template URL pattern; **not** used for Ideation xlsx. |

**I1 default:** sync SharePoint → `S3_SNAPSHOT_BUCKET` under `templates/` keys; loaders read bytes via boto3 (see `load_gtm_product_map_from_s3`).

---

## Environment variables

| Variable | Default S3 key | Purpose |
|---|---|---|
| `S3_SNAPSHOT_BUCKET` | *(required)* | Bucket for templates + snapshots |
| `GTM_DATABASE_KEY` | `templates/Fortune_AITool_GTM_Database.xlsx` | GTM workbook |
| `INVENTORY_CALENDAR_KEY` | `templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx` | Inventory + pricing workbook |
| `PRODUCT_DECKS_PREFIX` | `product-decks/` | Hunter PPTX binaries (Creation) |

Canonical defaults and sheet names: `ingestion/ideation_data_keys.py`.

Legacy note: `RATE_SHEET_KEY` (CSV rate card) was used by the old RAG deck generator prompt. Logic Guide I1 pricing comes from the **Inventory calendar Pricing tab**, not `RATE_SHEET_KEY`.

---

## Sync + refresh (ops)

| Asset | Owner *(confirm)* | Cadence *(confirm)* | sales-mcp sync |
|---|---|---|---|
| GTM database | GTM / data team | On product or slide-map changes | Manual or scripted upload to `GTM_DATABASE_KEY`; verify Product Tags + Audience Data after sync |
| Inventory calendar | DPS | **Daily** (reservations) | Manual or scripted upload to `INVENTORY_CALENDAR_KEY` until automated |
| Product decks | GTM / Hunter library | When Product Tags `Deck Path` changes | Existing corpus / `product-decks/` sync (Creation) |

**Before prod Ideation:** confirm GTM SharePoint path, who runs sync, and a stale-data policy (loud failure vs TTL warning — MVP prefers **loud failure** per END-SCOPE-SOT).

**Local dev:** copy current xlsx files into the dev bucket or use AWS credentials pointed at `fortune-sales-mcp-dev-artifacts`.

```bash
# Example manual sync (adjust paths and profile)
aws s3 cp Fortune_AITool_GTM_Database.xlsx \
  s3://fortune-sales-mcp-dev-artifacts/templates/Fortune_AITool_GTM_Database.xlsx

aws s3 cp "Fortune Inventory & Reservation Calendar 2026 Final.xlsx" \
  s3://fortune-sales-mcp-dev-artifacts/templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx
```

---

## Failure modes (MVP)

| Condition | Behavior |
|---|---|
| Missing S3 object | Loader raises; Ideation / build must not silently substitute |
| Missing sheet or column | Loader raises with sheet/column name |
| Product on Products tab but held/sold in flight | I2 drops candidate (availability first) |
| Product **not** on Products tab | No inventory constraint |
| Missing price for fundable product | I2 / loader fails loud |

---

## PI-2759 acceptance criteria map

| AC | Chunk |
|---|---|
| Documented access path | **A** (this file + README + `.env.example`) |
| Load Tags / products / audience / deck path | **B** *(Product Category + Product Tags / GTM TAGS via `gtm_ideation_catalog.py`)*; deck path + audience already in Creation |
| Load inventory by flight | **C** (`inventory_calendar.py`) |
| Load pricing for funding steps | D |
| Refresh/ownership note | **A** (Sync + refresh section above) |

---

## Open confirmations (GTM / DPS)

1. Exact SharePoint folder URL for `Fortune_AITool_GTM_Database.xlsx`
2. Named owner + notification when GTM workbook or inventory calendar changes
3. Whether sync stays manual upload or moves to scheduled Graph → S3 job

Update this doc when those are confirmed.
