# I1 — GTM + inventory data sources (access + refresh)

> **Ticket:** [PI-2759](https://fortune.atlassian.net/browse/PI-2759)  
> **Status:** I1 complete  
> **End-state SoT:** [`END-SCOPE-SOT.md`](END-SCOPE-SOT.md)

`build_deck` and optional `confirm_mix` read three SharePoint-owned assets via S3 snapshots.

```text
SharePoint (human SoT)  →  sync  →  S3 snapshot  →  sales-mcp loaders  →  build_deck / confirm_mix
```

This repo **consumes** snapshots. It does **not** edit inventory, pricing, or GTM rows.

---

## Assets

| Asset | Runtime access (sales-mcp) | Use |
|---|---|---|
| `Fortune_AITool_GTM_Database.xlsx` | S3 `GTM_DATABASE_KEY` | Product Tags (A5 clones), Audience Data (C2), Product Category + GTM TAGS (`confirm_mix`) |
| `Fortune Inventory & Reservation Calendar 2026 Final.xlsx` | S3 `INVENTORY_CALENDAR_KEY` | Products tab → availability; Pricing + Benchmarks → rates |
| Hunter product decks | S3 `PRODUCT_DECKS_PREFIX` | Creation only (A5 clones) |

**Not Google Tag Manager.** “GTM” here = Fortune’s internal product/GTM mapping workbook.

---

## Workbook tabs

### GTM database (`GTM_DATABASE_KEY`)

| Sheet | Loader | Use |
|---|---|---|
| **Product Tags** | `gtm_product_map.py` | Deck Path + Slide # (A5) |
| **Audience Data** | `audience_data.py` | Reach / Index (C2) |
| **Product Category** | `gtm_ideation_catalog.py` | Category titles (`confirm_mix`) |

### Inventory calendar (`INVENTORY_CALENDAR_KEY`)

| Sheet | Loader | Use |
|---|---|---|
| **Products** | `inventory_calendar.py` | Inventory-gated placements |
| **Inventory** | `inventory_calendar.py` | Available / Held / Sold vs flight dates |
| **Pricing + Benchmarks** | `inventory_pricing.py` | Authoritative rates for `confirm_mix` |

Combined entry: `ingestion/inventory_workbook.py` → `InventoryWorkbook`.

---

## Environment variables

| Variable | Default S3 key | Purpose |
|---|---|---|
| `S3_SNAPSHOT_BUCKET` | *(required)* | Bucket for templates + product decks |
| `GTM_DATABASE_KEY` | `templates/Fortune_AITool_GTM_Database.xlsx` | GTM workbook |
| `INVENTORY_CALENDAR_KEY` | `templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx` | Inventory + pricing workbook |
| `PRODUCT_DECKS_PREFIX` | `product-decks/` | Hunter PPTX binaries |
| `FORTUNEAI_TEMPLATE_KEY` | `templates/FortuneAI_DeckTemplate.pptx` | Creation spine |

Canonical defaults: `ingestion/ideation_data_keys.py`.

---

## Sync + refresh (ops)

| Asset | sales-mcp sync |
|---|---|
| GTM database | Upload to `GTM_DATABASE_KEY` when Product Tags or Audience Data change |
| Inventory calendar | Upload to `INVENTORY_CALENDAR_KEY` (reservations change frequently) |
| Product decks | Sync to `product-decks/` when Product Tags `Deck Path` changes |

**MVP policy:** missing or stale S3 objects → loud failure, not silent substitute.

```bash
aws s3 cp Fortune_AITool_GTM_Database.xlsx \
  s3://fortune-sales-mcp-dev-artifacts/templates/Fortune_AITool_GTM_Database.xlsx

aws s3 cp "Fortune Inventory & Reservation Calendar 2026 Final.xlsx" \
  s3://fortune-sales-mcp-dev-artifacts/templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx
```

---

## Failure modes

| Condition | Behavior |
|---|---|
| Missing S3 object | Loader raises; build must not silently substitute |
| Missing sheet or column | Loader raises with sheet/column name |
| Product held/sold in flight | `confirm_mix` drops or fails loud |
| Missing price for selected product | `confirm_mix` / build fails loud |
