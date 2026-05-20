# sales-mcp

Internal Python MCP server exposing Fortune's closed-won sales decks to Claude Cowork via MCP. Deployed as a Docker container on AWS Lightsail.

Live endpoint: `https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com`

## Requirements

- Python 3.12
- [uv](https://github.com/astral-sh/uv)
- Docker (for building/deploying)
- AWS CLI (for deploying)

## Environment setup

Copy `.env.example` to `.env` and fill in the values:

```bash
cp .env.example .env
```

| Variable | Description |
|---|---|
| `GRAPH_TENANT_ID` | Azure AD tenant ID |
| `GRAPH_CLIENT_ID` | Azure AD app registration client ID |
| `GRAPH_CLIENT_SECRET` | Azure AD app registration client secret |
| `SHAREPOINT_SITE_ID` | SharePoint site ID (format: `hostname,site-guid,web-guid`) |
| `S3_SNAPSHOT_BUCKET` | S3 bucket for Parquet snapshots (`fortune-sales-mcp-dev-artifacts` for dev, `fortune-sales-mcp-artifacts` for prod) |
| `S3_SNAPSHOT_PREFIX` | S3 prefix for deck snapshots (default: `snapshots`) |
| `S3_FAILED_PREFIX` | S3 prefix for failed-deck records (default: `failed`) |

The Azure app registration requires `Sites.Read.All` and `Files.Read.All` application permissions (not delegated) granted by an admin.

## Ingestion

The ingestion pipeline walks the `GTM Current` SharePoint library, parses all `.pptx` decks, and writes Parquet snapshots to S3.

```bash
# Run a full ingest (requires .env populated)
uv run python -c "
from dotenv import load_dotenv; load_dotenv(dotenv_path='.env')
from ingestion.ingestor import run_ingest
slides, failed = run_ingest()
print(f'slides: {len(slides)} | failed: {len(failed)}')
"
```

Output lands at:
- `s3://{S3_SNAPSHOT_BUCKET}/snapshots/{run_ts}/decks.parquet` — one row per slide
- `s3://{S3_SNAPSHOT_BUCKET}/failed/{run_ts}/failed.parquet` — corrupt/unreadable decks (omitted if empty)

Run tests:

```bash
uv run pytest
```

## Local dev

```bash
# Install deps
uv sync

# Run with hot-reload
uv run uvicorn server:app --reload --port 8000
```

To test tools interactively, point MCP Inspector at the local server:

```bash
npx @modelcontextprotocol/inspector http://localhost:8000/mcp
```

The inspector runs in your browser and lets you call tools manually without writing curl commands. Set the `Authorization` header to `Bearer <MCP_SHARED_SECRET>` in the inspector's auth settings.

## Deploy

```bash
# 1. Build for linux/amd64 (required for Lightsail)
docker buildx build --platform linux/amd64 --load -t fortune-sales-ai-mcp:dev .

# 2. Push to Lightsail (bump label each release)
aws lightsail push-container-image \
  --service-name fortune-sales-mcp \
  --label v1-11 \
  --image fortune-sales-ai-mcp:dev

# 3. Update containers.json with the new image digest printed above

# 4. Deploy
aws lightsail create-container-service-deployment \
  --service-name fortune-sales-mcp \
  --containers file://containers.json \
  --public-endpoint file://public-endpoint.json
```

> **Note:** `containers.json` contains `MCP_SHARED_SECRET` and must not be committed. Use `containers.json.template` as a reference and keep your local `containers.json` gitignored.
>
> Generate a secret with: `openssl rand -hex 32`

## Testing the live server

The MCP streamable HTTP protocol requires an `initialize` handshake before any other call.

**Step 1 — initialize and grab the session ID from the response headers:**

```bash
curl -s -D - -X POST https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/mcp \
  -H "Authorization: Bearer <MCP_SHARED_SECRET>" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"0.1"}}}'
```

**Step 2 — list tools (replace `SESSION_ID` with the value from the `mcp-session-id` header above):**

```bash
curl -s -X POST https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/mcp \
  -H "Authorization: Bearer <MCP_SHARED_SECRET>" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "mcp-session-id: SESSION_ID" \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
```

**Health check (no auth required):**

```bash
curl https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com/health
```

## Architecture decisions

**Python + FastMCP** — MCP's official Python SDK (`mcp[cli]`). The standalone `fastmcp` PyPI package is not used; the SDK ships `mcp.server.fastmcp.FastMCP` directly.

**Streamable HTTP transport (not SSE)** — SSE was dropped because Lightsail's load balancer rewrites the `Host` header, causing the MCP SDK's DNS-rebinding protection to reject every request with 421. Streamable HTTP sidesteps this; DNS-rebinding protection is disabled at the SDK level via `TransportSecuritySettings` since the LB is already the trust boundary.

**AWS Lightsail container (not Lambda + Amplify)** — the MCP server is a long-running stateful process that holds SSE connections open. Lambda's execution model (short-lived, request/response) is a poor fit. Lightsail gives a persistent container with a public HTTPS endpoint and no cold-start latency.

**v0 scope** — RAG retrieval and rate-card lookup only. LLM inference happens in Cowork via its native pptx skill; the server does not make Claude API calls. S3 is the storage layer for the deck corpus; no DynamoDB in v0.
