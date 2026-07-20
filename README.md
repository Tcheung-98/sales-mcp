# sales-mcp

Internal Python MCP server that powers Prodie's pitch deck generation for Fortune sales associates.
Associates have a conversation with Prodie about a client, confirm the product selection and budget,
and Prodie uses this server to generate a Fortune-branded PPTX from the real closed-won corpus.

Live endpoint: `https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com`

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

**Schema-driven generation** — deck generation requires a fully hydrated `DeckSchema` (client,
industry, budget, confirmed products). Prodie enforces sufficiency during conversation; the server
validates independently via Pydantic and selects the SharePoint template filename from industry,
franchise keywords, and product names.

**Schema validation + skeleton assembly** — `prepare_deck(schema)` validates the handoff payload
and returns the template filename for AE review. `build_deck(schema, template_url)` fetches the
SharePoint template, swaps product placeholders with corpus clones, and returns a presigned PPTX
URL. Clone lives in `DeckGenerator.assemble_skeleton` (Anthropic-free); apply helpers live in
`ingestion/pptx_tools.py` for the future Cursor stylist. Skeleton only for now — no LLM copy or
stylist pass yet.

**In-memory vector search** — embeddings loaded into numpy at startup, cosine similarity at query
time. No FAISS index; the corpus (2,404 slides) is small enough that in-memory search is fast and
removes a dependency.

**Streamable HTTP transport** — SSE dropped because Lightsail's load balancer rewrites the `Host`
header, triggering the MCP SDK's DNS-rebinding protection with 421 errors. DNS-rebinding protection
disabled at the SDK level via `TransportSecuritySettings`; the LB is the trust boundary.

**Anthropic API (not Bedrock Claude)** — generation uses the Anthropic API directly. Bedrock is
used only for embeddings (Titan Text v2). API key stored in AWS Secrets Manager; local dev uses
`ANTHROPIC_API_KEY` env var.

**Python + FastMCP** — official MCP Python SDK (`mcp[cli]`). Do not use the standalone `fastmcp`
PyPI package; the SDK ships `mcp.server.fastmcp.FastMCP` directly.
