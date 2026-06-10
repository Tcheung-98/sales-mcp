import os

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

from ingestion.retriever import SlideRetriever

_EXPECTED_TOKEN = os.environ.get("MCP_SHARED_SECRET")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.url.path == "/health":
            return await call_next(request)
        if not _EXPECTED_TOKEN:
            return JSONResponse({"error": "server misconfigured"}, status_code=503)
        auth = request.headers.get("authorization", "")
        if not auth.startswith("Bearer ") or auth[7:] != _EXPECTED_TOKEN:
            return JSONResponse({"error": "unauthorized"}, status_code=401)
        return await call_next(request)


mcp = FastMCP(
    "Sales MCP",
    transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
)

_retriever: SlideRetriever | None = None


def _get_retriever() -> SlideRetriever:
    global _retriever
    if _retriever is None:
        _retriever = SlideRetriever()
    return _retriever


@mcp.tool()
def search_decks(query: str, k: int = 5) -> list[dict]:
    """Search Fortune sales decks by semantic similarity. Returns the k most relevant slides."""
    return _get_retriever().search(query, k=k)


@mcp.tool()
def filter_decks_by_tags(
    industry: str | None = None,
    sub_industry: str | None = None,
    product_line: str | None = None,
    deal_size: str | None = None,
    client_name: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    deck_type: str | None = None,
    limit: int = 20,
) -> list[dict]:
    """
    Filter Fortune sales decks by structured metadata tags. Use this when the account
    executive specifies known criteria such as a specific industry, deck type, or date
    range — for example "show me all Tech pitches from Q1" or "find Finance proposals".
    All supplied filters apply with AND semantics; omit any filter to match any value.
    Returns deck-level metadata (deck_id, title, tags, slide_count, source_path), not
    slide content — follow up with get_slide_content to read the actual slides.
    Use search_decks instead for open-ended natural-language queries.
    """
    return _get_retriever().filter_decks_by_tags(
        industry=industry,
        sub_industry=sub_industry,
        product_line=product_line,
        deal_size=deal_size,
        client_name=client_name,
        date_from=date_from,
        date_to=date_to,
        deck_type=deck_type,
        limit=limit,
    )


@mcp.tool()
def get_slide_content(deck_id: str, slide_numbers: list[int] | None = None) -> list[dict]:
    """
    Retrieve full slide content (title, body text, layout) for a specific deck.
    Call this after search_decks or filter_decks_by_tags returns a deck_id you want
    to read in detail. Pass slide_numbers to fetch specific slides; omit to get every
    slide in the deck. Returns an empty list if the deck_id is not found or the
    requested slide numbers don't exist — never raises an error.
    """
    return _get_retriever().get_slide_content(deck_id, slide_numbers)


async def health(request):
    return JSONResponse({"status": "ok"})


app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.add_middleware(AuthMiddleware)

if __name__ == "__main__":
    mcp.run()
