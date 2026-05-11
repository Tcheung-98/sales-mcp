import sys

from mcp.server.fastmcp import FastMCP

# just naming it Sales MCP for now
mcp = FastMCP(
    "Sales MCP",
    instructions="You are a sales assistant. You are given a deck of slides and you need to help the user answer questions about the deck.",
    # debug=True enables noisy asyncio logs (e.g. "Using selector: KqueueSelector") and is not required for stdio.
    debug=False,
    log_level="INFO",
)

# Stub data so you have something to query
DECKS = [
    {"id": "deck1", "title": "Tech Q1 Pitch", "industry": "tech"},
    {"id": "deck2", "title": "Healthcare Q2", "industry": "healthcare"},
]


@mcp.tool()
def search_historical_decks(industry: str) -> str:
    """Find historical decks by industry."""
    # return [d for d in DECKS if d["industry"] == industry]
    return "who knows who cares go away"

@mcp.tool()
def hello(name: str) -> str:
    """Sanity check tool."""
    return f"Hello, {name}!"

if __name__ == "__main__":
    # MCP stdio: only JSON-RPC on stdout. Human status goes to stderr.
    print(
        "sales-mcp: stdio server ready; waiting for MCP client on stdin.",
        file=sys.stderr,
        flush=True,
    )
    mcp.run(transport="stdio")