# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is
Internal Python FastMCP server exposing Fortune's closed-won sales decks to Claude Cowork via MCP. It is a retrieval/RAG backend — Claude (via Cowork's native pptx skill) does the generation. Deployed as a Docker container on AWS Lightsail Container Service.

## Dev commands

```bash
# Install deps
uv sync

# Run locally (hot-reload)
uv run uvicorn server:app --reload --port 8000

# Test with MCP Inspector (pointed at local server)
npx @modelcontextprotocol/inspector

# Build for linux/amd64 (required for Lightsail)
docker buildx build --platform linux/amd64 --load -t fortune-sales-ai-mcp:dev .

# Smoke-test the container locally
docker run --rm -p 8000:8000 fortune-sales-ai-mcp:dev
curl http://localhost:8000/health
```

## Deploy to AWS Lightsail

```bash
# 1. Push image to Lightsail (bump label each release, e.g. v1.8)
aws lightsail push-container-image \
  --service-name fortune-sales-mcp \
  --label v1.X \
  --image fortune-sales-ai-mcp:dev

# 2. Update containers.json with the new ":fortune-sales-mcp.vX.Y" digest printed above

# 3. Deploy
aws lightsail create-container-service-deployment \
  --service-name fortune-sales-mcp \
  --containers file://containers.json \
  --public-endpoint file://public-endpoint.json
```

Live endpoint: `https://fortune-sales-mcp.tj3ek8xjdg9br.us-east-1.cs.amazonlightsail.com`

## Architecture

- **`server.py`** — single entry point. Defines `FastMCP` instance, registers `@mcp.tool()` handlers, appends `/health` route, and exports `app = mcp.streamable_http_app()` for uvicorn.
- **Transport**: Streamable HTTP (not SSE). SSE was dropped because of 421 "Invalid Host" errors behind Lightsail's load balancer; `TrustedHostMiddleware(allowed_hosts=["*"])` is inserted at the outermost middleware position as a workaround.
- **Auth**: shared-secret bearer token (Cowork → MCP server). Implementation pending.
- **Storage**: S3 for deck corpus and generated artifacts (not yet wired in server.py).
- **SDK**: official `mcp[cli]` package (`mcp.server.fastmcp.FastMCP`). Do NOT use the standalone `fastmcp` PyPI package.

## Key decisions / constraints
- Python 3.12, deps managed with `uv` (not pip directly).
- Docker image must target `linux/amd64` — Lightsail nodes are x86.
- `deploy.json` is the full deployment spec; `containers.json` and `public-endpoint.json` are the split files used by the CLI commands above.
- v0 scope: RAG retrieval + rate-card lookup only. Out of scope: email/brief generation, Salesforce integration, PII scrubbing, server-side .pptx generation, predictive pricing.
