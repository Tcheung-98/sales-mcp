# sales-mcp

Python MCP server for sales workflows using [FastMCP](https://github.com/modelcontextprotocol/python-sdk) (`mcp` package).

## Requirements

- **Python 3.10+** (3.12 is used in local development)

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -U pip
pip install -e .
```

## Run (stdio, local MCP clients)

```bash
python server.py
```

Dependencies and the pinned MCP SDK version are declared in `pyproject.toml`.
