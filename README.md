# sales-mcp

Internal Python MCP server for Fortune pitch deck **Creation** (deck generation only).

The upstream caller — **Pitch Deck Builder**, **Sales HQ**, or any MCP client — sends a **complete
`DeckSchema` payload**: Discovery fields plus `confirmed_products[]` (each with `name`, `category`,
`price`, `cadence`). This server validates that payload, assembles the FortuneAI spine and exact
GTM product clones (C1), fills stock placeholders (C2), optionally runs the deck QA rail, and
uploads a PPTX. It does **not** ideate, propose a product menu, or choose the mix.

End-scope SoT (sales-mcp checkout): `docs/END-SCOPE-SOT.md`.  
Historical Prodie ideation design (not the runtime contract): [`docs/PRODIE-IDEATION-SPEC.md`](docs/PRODIE-IDEATION-SPEC.md).

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
| `S3_SNAPSHOT_BUCKET` | S3 bucket (`fortune-sales-mcp-dev-artifacts` for dev) |
| `ANTHROPIC_API_KEY` | Anthropic API key (local dev only — prod uses Secrets Manager) |
| `MCP_SHARED_SECRET` | Bearer token for Cowork → MCP auth |
| `GTM_DATABASE_KEY` | S3 key for `Fortune_AITool_GTM_Database.xlsx` (default: `templates/Fortune_AITool_GTM_Database.xlsx`) |
| `INVENTORY_CALENDAR_KEY` | S3 key for inventory + pricing workbook (default: `templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx`) |
| `PRODUCT_DECKS_PREFIX` | S3 prefix for Hunter product decks referenced by Deck Path (default: `product-decks/`) |
| `FORTUNEAI_TEMPLATE_KEY` | S3 key for Creation spine (default: `templates/FortuneAI_DeckTemplate.pptx`) |
| `TEMPLATE_URL_ALLOWED_HOSTS` | Optional extra hosts for `build_deck` template URLs (comma-separated) |
| `DECK_QA_DISABLED` | Bypass the always-on deck QA rail (`1`/`true`/`yes`) — local dev / emergency ops only, never in prod |
| `DECK_QA_SKIP_VISION` | Run B2+B3 but skip the B4 vision pass (`1`/`true`/`yes`) — local dev only |
| `DECK_QA_TIMEOUT_S` | Wall-clock budget for the QA rail (default: `600`); overrun fails loud |
| `CURSOR_API_KEY` | Cursor SDK key for the B4 vision pass — required in production |

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

`containers.json` contains `MCP_SHARED_SECRET` and `CURSOR_API_KEY` — never commit it. Use
`containers.json.template` as reference. Generate MCP auth with `openssl rand -hex 32`.

Pushes to `main` / `dev` deploy via `.github/workflows/deploy.yml`. Before the first QA-enabled
prod deploy, set GitHub Actions secrets `PROD_CURSOR_API_KEY` and `DEV_CURSOR_API_KEY` (repo
Settings → Secrets). The workflow injects them as `CURSOR_API_KEY` on Lightsail plus
`DECK_QA_TIMEOUT_S=600`.

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

**Corpus-clone approach** — product slides are cloned from exact GTM `Deck Path` + `Slide #`
(Hunter product decks on S3), not generated programmatically. Stock FortuneAI slides keep
template typography; client-specific text is filled via the C2 placeholder pipeline.

**Schema-driven generation** — the primary path is `build_deck(full DeckSchema)`. The caller owns
Discovery intake and product lock upstream; this server consumes a fully hydrated `DeckSchema`
(Discovery fields + non-empty `confirmed_products`). `DiscoverySchema` covers Workflow Discovery
fields; `DeckSchema` extends it for Creation. Pydantic validation runs here; product pages clone
exact GTM `Deck Path` / `Slide #` into FortuneAI_DeckTemplate (A5 + C1).

**Discovery ↔ Creation handoff (PI-2758)** — field map for upstream forms (Pitch Deck Builder / Sales HQ):

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
| Confirmed mix (locked upstream) | `confirmed_products` | Creation only — each entry needs `name`, `category`, `price`, `cadence` |

Industry enum (Workflow): Technology, Professional Services, Healthcare, Financial Services,
Energy, Lifestyle, Luxury. Legacy `Tech` normalizes to `Technology`. Legacy
`budget_quarterly` still shims to a single budget tier. Escalation uses the max tier amount
(≥ $750k → GTM).

**Primary path vs legacy helpers** — call `build_deck(deck_schema)` with the locked payload.
There is **no `propose_mix` MCP tool**. **`confirm_mix`** (I3 / PI-2761) is optional
legacy: when the caller sends only `[{name, category?}, ...]`, it validates GTM identity,
hydrates authoritative price/cadence from inventory, checks flight availability, and returns
`deck_schema` for `build_deck`. Prefer sending the full `confirmed_products[]` directly.
Conference / Lists platforms return `status: escalation` (do not build).

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
fail loud (GTM escalate). Missing or ambiguous map rows fail loud (no Titan substitute).

**Headless deck QA gate (B2–B4)** — **always on**: every `build_deck` runs the post-C2
deck through a review package (draft + slide PNGs + manifest), deterministic checks, and a
headless Cursor vision pass with at most **one** fix loop; product clones are flag-only. A
deterministic failure, an explicit QA failure, or a `DECK_QA_TIMEOUT_S` overrun (default 600s)
returns `status: error` with `qa_report` — an unreviewed deck is never delivered. Review
packages land under the ephemeral `review-packages/{uuid}/` S3 prefix (expire them with a
7–30 day lifecycle rule). `DECK_QA_DISABLED` (whole rail) and `DECK_QA_SKIP_VISION` (B4 only)
are local-dev/emergency bypasses — never set in production.
Contract: [`docs/DECK-QA-ARCHITECTURE.md`](docs/DECK-QA-ARCHITECTURE.md).

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

**GTM + inventory data (I1 / PI-2759)** — `build_deck` and optional `confirm_mix` read GTM DB +
inventory calendar + pricing from S3 snapshots (SharePoint is human SoT). Access path, sheet
contract, sync/ownership, and env defaults: [`docs/I1-DATA-SOURCES.md`](docs/I1-DATA-SOURCES.md).
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

---

## MCP tools

| Tool | Purpose |
|---|---|
| `build_deck` | Primary Creation path: validate `DeckSchema`, C1 assemble + C2 fills → presigned PPTX URL |
| `confirm_mix` | Optional legacy: validate name+category list, hydrate prices → `deck_schema` for `build_deck` |

---

## Repo layout

- `server.py` — FastMCP app (`build_deck`, `confirm_mix`)
- `ingestion/generator.py` — FortuneAI assembly (`assemble_skeleton`, `build`)
- `ingestion/placeholder_fills.py` — C2 deterministic + AI placeholder fills
- `ingestion/gtm_product_map.py` — A5 exact product slide map
- `ingestion/schema.py` — Discovery + Deck Pydantic models
- `docs/END-SCOPE-SOT.md` — Canonical end-state contract
- `docs/PROGRESS.md` — Living goal + ticket status
- `docs/I1-DATA-SOURCES.md` — GTM + inventory S3 sync contract
- `tests/` — unit tests (no live S3/Anthropic in default suite)
