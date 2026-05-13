# sales-mcp

Internal Python MCP server exposing Fortune's closed-won sales decks to Claude Cowork via MCP. Deployed as a Docker container on AWS Lightsail.

Live endpoint: `https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com`

## Requirements

- Python 3.12
- [uv](https://github.com/astral-sh/uv)
- Docker (for building/deploying)
- AWS CLI (for deploying)

## Local dev

```bash
# Install deps
uv sync

# Run with hot-reload
uv run uvicorn server:app --reload --port 8000
```

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
