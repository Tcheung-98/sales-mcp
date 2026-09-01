# sales-mcp

Internal Python MCP server that powers Prodie's pitch deck **Creation** for Fortune sales associates.
Associates fill a Discovery form (optionally informed by a SalesGPT conversation). **Prodie** proposes
relevant, priced, available products; the associate confirms via checkboxes. This server then
generates a Fortune-branded PPTX (FortuneAI spine + exact Hunter product-slide clones).
It does **not** choose the mix. Prodie does **not** assemble the PPTX.

End-scope SoT (sales-mcp checkout): `local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md`.  
**Prodie Ideation spec:** [`docs/PRODIE-IDEATION-SPEC.md`](docs/PRODIE-IDEATION-SPEC.md) — Prodie menu + select; this server `build_deck`.

---

## Requirements

- Python 3.12
- [uv](https://github.com/astral-sh/uv)
- Docker (for building/deploying)
- AWS CLI (for deploying)

---

## Environment setup

```bash
cp .env.example .env
```

| Variable | Description |
|---|---|
| `GRAPH_TENANT_ID` | Azure AD tenant ID |
| `GRAPH_CLIENT_ID` | Azure AD app registration client ID |
| `GRAPH_CLIENT_SECRET` | Azure AD app registration client secret |
| `SHAREPOINT_SITE_ID` | SharePoint site ID (`hostname,site-guid,web-guid`) |
| `S3_SNAPSHOT_BUCKET` | S3 bucket (`fortune-sales-mcp-dev-artifacts` for dev) |
| `ANTHROPIC_API_KEY` | Anthropic API key (local dev only — prod uses Secrets Manager) |
| `MCP_SHARED_SECRET` | Bearer token for Cowork → MCP auth |
| `RULEBOOK_KEY` | S3 key for Fortune GTM skill doc (default: `templates/rulebook.docx`) |
| `GTM_DATABASE_KEY` | S3 key for `Fortune_AITool_GTM_Database.xlsx` (default: `templates/Fortune_AITool_GTM_Database.xlsx`) |
| `INVENTORY_CALENDAR_KEY` | S3 key for inventory + pricing workbook (default: `templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx`) |
| `PRODUCT_DECKS_PREFIX` | S3 prefix for Hunter product decks referenced by Deck Path (default: `product-decks/`) |
| `FORTUNEAI_TEMPLATE_KEY` | S3 key for Creation spine (default: `templates/FortuneAI_DeckTemplate.pptx`) |
| `TEMPLATE_URL_ALLOWED_HOSTS` | Optional extra hosts for `build_deck` template URLs (comma-separated) |

---

## Local dev

```bash
# Install deps
uv sync

# Run with hot-reload
uv run uvicorn server:app --reload --port 8000

# Test tools interactively
npx @modelcontextprotocol/inspector http://localhost:8000/mcp
```

Set `Authorization: Bearer <MCP_SHARED_SECRET>` in the inspector's auth settings.

---

## Ingestion

Walks the SharePoint GTM library, parses all `.pptx` decks, uploads PPTX files to S3 corpus,
and regenerates the embeddings snapshot.

```bash
uv run python scratch/run_ingest.py
```

Output:
- `s3://bucket/corpus/{deck_id}.pptx` — source PPTX files for cloning
- `s3://bucket/snapshots/{run_ts}/` — embeddings + slide metadata

---

## Tests

```bash
uv run pytest
uv run ruff check .
```

---

## Deploy

```bash
# 1. Build for linux/amd64 (required for Lightsail)
docker buildx build --platform linux/amd64 --load -t fortune-sales-ai-mcp:dev .

# 2. Push to Lightsail (bump label each release)
aws lightsail push-container-image \
  --service-name fortune-sales-mcp \
  --label v1-X \
  --image fortune-sales-ai-mcp:dev

# 3. Update containers.json with the new image digest printed above

# 4. Deploy
aws lightsail create-container-service-deployment \
  --service-name fortune-sales-mcp \
  --containers file://containers.json \
  --public-endpoint file://public-endpoint.json
```

`containers.json` contains `MCP_SHARED_SECRET` — never commit it. Use `containers.json.template`
as reference. Generate a secret with `openssl rand -hex 32`.

---

## Testing the live server

```bash
# Health check (no auth)
curl https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/health

# Initialize session
curl -s -D - -X POST https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/mcp \
  -H "Authorization: Bearer <MCP_SHARED_SECRET>" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"0.1"}}}'

# List tools (replace SESSION_ID from mcp-session-id response header)
curl -s -X POST https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/mcp \
  -H "Authorization: Bearer <MCP_SHARED_SECRET>" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "mcp-session-id: SESSION_ID" \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
```

---

## Architecture decisions

**Corpus-clone approach** — slides are cloned from real Fortune closed-won decks, not generated
programmatically. Visual quality (typography, shapes, imagery) is preserved from the source.
Client-specific text is replaced post-clone via placeholder targeting.

**Schema-driven generation** — deck generation requires a fully hydrated `DeckSchema`
(Discovery intake + confirmed products). `DiscoverySchema` covers Workflow Discovery fields.
`DeckSchema` extends it with non-empty `confirmed_products` for Creation. **Prodie** proposes the
mix (Logic Guide V1 + GTM/inventory) and enforces sufficiency; this server validates independently
via Pydantic and clones exact GTM `Deck Path` / `Slide #` pages into FortuneAI_DeckTemplate.

**Discovery ↔ Creation handoff (PI-2758)** — field map for Prodie / Sales HQ forms:

| Workflow / form field | Schema field | Required |
| --- | --- | --- |
| Company name | `company_name` (alias `client_name`) | yes |
| Industry | `industry` | yes |
| Budget (up to 3 tiers) | `budgets[].amount` (+ optional `label`) | yes (1–3) |
| Flight dates | `flight_dates.start` / `.end` | yes |
| Campaign goal | `campaign_goal` | yes |
| Targeting details | `targeting_details` | yes |
| KPIs | `kpis` | yes |
| KPI details | `kpi_details` | yes |
| Campaign narrative | `campaign_narrative` | yes |
| Preferred platforms/products | `preferred_platforms_products` | yes |
| Additional RFP details | `additional_rfp_details` | yes |
| Client logo | `client_logo` (URL or SharePoint path) | yes |
| Platform/product specifics | `platform_or_product_specifics` | no |
| Confirmed mix (Ideation out) | `confirmed_products` | Creation only |

Industry enum (Workflow): Technology, Professional Services, Healthcare, Financial Services,
Energy, Lifestyle, Luxury. Legacy `Tech` normalizes to `Technology`. Legacy
`budget_quarterly` still shims to a single budget tier. Escalation uses the max tier amount
(≥ $750k → GTM).

**Propose + select (MVP) vs this server** — Prodie shows a relevant-product menu (Logic Guide V1 +
GTM + inventory). Associates lock via checkboxes; then pass the spec to `build_deck`
(`confirm_mix` validates names/prices/availability first).
There is **no `propose_mix` MCP tool**. The in-repo Logic Guide modules remain isolated
reference/test code and are not part of the associate runtime.

**Creation lock (I3 / PI-2761)** — `confirm_mix(discovery, selected_products)` accepts the
associate's complete checkbox list as `[{name, category?}, ...]`. It does not accept prices,
tiers, swaps, scores, or ranking instructions. The server resolves exact GTM identity plus
authoritative price/cadence and flight availability, then returns `deck_schema` with
`confirmed_products` for `build_deck`. Conference / Lists platforms return
`status: escalation` (do not build). Unavailable products cannot be confirmed in MVP.

**FortuneAI assembly + C2 fills (C1 / PI-2756 + C2 / PI-2757)** — `build_deck(schema, template_url?)`
validates the handoff and assembles from **FortuneAI_DeckTemplate** (not industry
`Category_Presentation_*`). Optional `template_url` must name FortuneAI_DeckTemplate
(SharePoint download URL); when omitted, the template loads from S3
(`FORTUNEAI_TEMPLATE_KEY`). **C1** (`assemble_skeleton`, Anthropic-free): keeps intro /
narrative / investment / thank-you stock layout; inserts category dividers **only when ≥1
funded product** maps to that section (fixed order: High-Impact Media → Editorial Alignment →
Premium Video → Print → Branded Content); product pages under each divider are **exact** GTM
Product Tags clones (`Deck Path` + `Slide #`). **C2** (`apply_placeholders` after assembly):
fills date/logo/history/audience metrics/program types/investment/thanks, bounded Claude for
named narrative slots, drops unused audience/program variant pages. Events / Conference products
fail loud (GTM escalate). Missing or ambiguous map rows fail loud (no Titan substitute). No
stylist (PI-2754 shelved).

**Per-slide fill method (FortuneAI stock spine, pre-product insert):**

| Slide role | Method | Source |
|---|---|---|
| Intro | AI + data + logo | Claude `[TITLE]`; generate-time Month/Year `[DATE]`; HTTPS `client_logo` |
| Why Fortune | Stock | Unchanged |
| History of Trust | Stock + swap | `[client name]` → `company_name` |
| Opportunity | AI | Claude `[HEADER]` + `[BODY]` |
| Audience (one variant kept) | AI + data | Claude `[AUDIENCE TITLE]`; Reach/Index from Audience Data |
| Program Overview (one variant kept) | AI + data | Divider names as `PRODUCT TYPE`; Claude program blurbs (1-category: second box is stock Fortune sentence) |
| Category dividers | Stock (C1) | Conditional insert only |
| Product pages | Master Deck Pull (A5) | Exact GTM `Deck Path` + `Slide #` clone |
| Investment | Data pull + math | Mix sum `[BUDGET]`; per-category bullets; budget mismatch fails loud |
| Thank You | Stock + data + logo | Same date/logo as intro |

**GTM workbook (A5 + C2)** — sync the Hunter workbook and decks it names into S3 before build:

- Object: `GTM_DATABASE_KEY` (default `templates/Fortune_AITool_GTM_Database.xlsx`)
- **Product Tags** sheet → exact product slide map (A5)
- **Audience Data** sheet → segment / Reach / Index for audience cards (C2); matched via
  `targeting_details`; never invent metrics

Product Tags lookup and Audience Data load are separate passes over the same xlsx.

**Ideation data sources (I1 / PI-2759)** — Prodie reads GTM DB +
inventory calendar + pricing from S3 snapshots (SharePoint is human SoT). Access path, sheet
contract, sync/ownership, and env defaults: [`local/schema-driven-deck-generation-engine/I1-DATA-SOURCES.md`](local/schema-driven-deck-generation-engine/I1-DATA-SOURCES.md).
Canonical keys: `ingestion/ideation_data_keys.py`. **Chunk B:** `ingestion/gtm_ideation_catalog.py`
loads Product Category + Product Tags (`GTM TAGS` column) from the same xlsx; **Chunk C:**
`ingestion/inventory_calendar.py` loads Products + Inventory tabs for flight availability;
**Chunk D:** `ingestion/inventory_pricing.py` + `inventory_workbook.py` for rates.
Creation already uses Product Tags + Audience Data from `GTM_DATABASE_KEY`.

**GTM product slide map (A5 / PI-2541)** — Product pages are deterministic:

- Lookup: exact `Product Name` + `Product Category` (schema aliases: `Newsletter`→`Newsletters`,
  `Digital Media`→`Digital Ads/Programmatic`)
- Binaries: `product-decks/{Deck Path}` (adds `.pptx` when the sheet omits the extension)

Known Product Tags coverage gaps (flag for GTM; do not invent substitutes):

- Print is sparse (only Full Page); Deck Path is `FortuneAI_DeckTemplate` (no `.pptx` in sheet)
- Many Digital Ads section/sub-section takeovers share Slide #9 on High Impact Media
- Duplicate Branded Content rows (same name/path/slide, different GTM TAGS) — deduped as one
- `Term Sheet` / `Next To Lead` appear in both Newsletters and Vodcasts — category required

**Titan / RAG (legacy)** — `search_decks`, `filter_decks_by_tags`, and corpus embeddings remain
for research and ingest. They are **not** the associate Creation path (exact GTM map only).

---

## MCP tools

| Tool | Purpose |
|---|---|
| `build_deck` | FortuneAI assembly + C2 fills → presigned PPTX URL |
| `confirm_mix` | Validate locked product list → `deck_schema` for `build_deck` |
| `search_decks` | Semantic slide search (research; not associate picker) |
| `filter_decks_by_tags` | Tag filter on corpus metadata |
| `get_slide_content` | Fetch slide text for a deck id |

---

## Repo layout

- `server.py` — FastMCP app + tool handlers
- `ingestion/generator.py` — FortuneAI assembly (`assemble_skeleton`, `build`)
- `ingestion/placeholder_fills.py` — C2 deterministic + AI placeholder fills
- `ingestion/gtm_product_map.py` — A5 exact product slide map
- `ingestion/schema.py` — Discovery + Deck Pydantic models
- `tests/` — unit tests (no live S3/Anthropic in default suite)
