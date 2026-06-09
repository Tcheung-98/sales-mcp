# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Internal Python FastMCP server that powers Prodie's pitch deck generation for Fortune sales
associates. The server handles corpus retrieval, schema validation, generation orchestration,
and PPTX assembly. Prodie (Claude Cowork) handles the conversation with the AE and calls
this server's MCP tools when ready to build.

Deployed as a Docker container on AWS Lightsail Container Service.

## Dev commands

```bash
# Install deps
uv sync

# Run locally (hot-reload)
uv run uvicorn server:app --reload --port 8000

# Test with MCP Inspector (pointed at local server)
npx @modelcontextprotocol/inspector http://localhost:8000/mcp

# Lint (must pass before every PR)
uv run ruff check .

# Tests
uv run pytest

# Build for linux/amd64 (required for Lightsail)
docker buildx build --platform linux/amd64 --load -t fortune-sales-ai-mcp:dev .

# Smoke-test the container locally
docker run --rm -p 8000:8000 fortune-sales-ai-mcp:dev
curl http://localhost:8000/health
```

## Deploy to AWS Lightsail

```bash
# 1. Push image to Lightsail (bump label each release)
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

- **`server.py`** — FastMCP instance, `@mcp.tool()` handlers, `/health` route, auth middleware.
- **`ingestion/schema.py`** — Pydantic `DeckSchema` + `Product`. Validates incoming schema, enforces $750K escalation rule, constructs deterministic arc from `confirmed_products`.
- **`ingestion/generator.py`** — `DeckGenerator`: calls Claude to fill arc slots from corpus candidates, builds PPTX via corpus-clone engine, uploads to S3.
- **`ingestion/retriever.py`** — `SlideRetriever`: loads embeddings snapshot at startup, cosine similarity search, tag-based filtering.
- **`ingestion/ingestor.py`** — offline ingestion pipeline: SharePoint → parse → S3 corpus upload → embed → snapshot.

## Generation flow

```
outline_deck(schema) / build_deck(schema)
    ↓
ingestion/schema.py  — validate, build deterministic arc from confirmed_products
    ↓
ingestion/retriever.py — per-slot corpus search (k=8 per slot, targeted queries)
    ↓
ingestion/generator.py._call_claude() — Claude picks best slide per slot, rewrites text
    ↓
ingestion/generator.py._build_pptx() — clone slides from S3 corpus, apply replacements
    ↓
S3 upload → presigned URL returned
```

## Key decisions / constraints

- Python 3.12, deps managed with `uv` (not pip directly).
- Docker image must target `linux/amd64` — Lightsail nodes are x86.
- `containers.json` contains `MCP_SHARED_SECRET` — never commit it.
- SDK: official `mcp[cli]` package (`mcp.server.fastmcp.FastMCP`). Do NOT use the standalone `fastmcp` PyPI package.
- Anthropic API key: stored in AWS Secrets Manager (`fortune-sales-mcp/claude-api-key`). Local dev uses `ANTHROPIC_API_KEY` env var instead.
- `deploy.json` is the full deployment spec; `containers.json` and `public-endpoint.json` are the split files used by the CLI commands above.
- `uv run ruff check .` must pass before any PR. No exceptions.
- Body text replacement only targets placeholder shapes (shapes with a valid `placeholder_format.idx`). Non-placeholder shapes (chart annotations, floating text boxes) are never overwritten.
- Hardcoded placeholder indices (`idx=0` title, `idx=19` eyebrow, `idx=12/10` cover) are Fortune blank template assumptions. If the template changes, these may need updating.
