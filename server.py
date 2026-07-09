import os

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from pydantic import ValidationError
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.middleware.cors import CORSMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

from ingestion.generator import DeckGenerator
from ingestion.retriever import SlideRetriever
from ingestion.schema import DeckSchema, select_template

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
_generator: DeckGenerator | None = None


def _get_retriever() -> SlideRetriever:
    global _retriever
    if _retriever is None:
        _retriever = SlideRetriever()
    return _retriever


def _get_generator() -> DeckGenerator:
    global _generator
    if _generator is None:
        _generator = DeckGenerator()
    return _generator


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


@mcp.tool()
def prepare_deck(schema: dict) -> dict:
    """
    Validate a deck schema and return the template that would be used for AE review.
    Call this before build_deck. Returns status 'ok' with template_filename if valid,
    'incomplete' with missing fields if not, or 'escalation' if the budget meets the
    GTM threshold. Prodie should fetch the template_filename from the Fortune Sales
    Automation SharePoint folder and pass its URL to build_deck.
    """
    try:
        parsed = DeckSchema.model_validate(schema)
    except ValidationError as exc:
        missing = []
        errors = []
        for e in exc.errors():
            loc = ".".join(str(p) for p in e["loc"])
            if e["type"] == "missing":
                missing.append(loc)
            else:
                errors.append(f"{loc}: {e['msg']}")
        if any("GTM team" in err for err in errors):
            return {"status": "escalation", "message": errors[0]}
        return {"status": "incomplete", "missing": missing, "errors": errors}
    return {"status": "ok", "template_filename": select_template(parsed)}


@mcp.tool()
def build_deck(schema: dict, template_url: str) -> dict:
    """
    Build a Fortune pitch deck from a confirmed schema and an approved template.
    template_url: pre-authenticated download URL for the .pptx template file,
    resolved by Prodie from the Fortune Sales Automation SharePoint folder.
    Validates the schema, populates the template with client-specific copy via
    Claude, uploads to S3, and returns a presigned download URL valid for 24h.
    """
    try:
        parsed = DeckSchema.model_validate(schema)
    except ValidationError as exc:
        errors = [f"{'.'.join(str(p) for p in e['loc'])}: {e['msg']}" for e in exc.errors()]
        return {"status": "error", "message": "; ".join(errors)}
    try:
        return _get_generator().build(parsed, template_url, _get_retriever())
    except ValueError as exc:
        return {"status": "error", "message": str(exc)}


async def health(request):
    return JSONResponse({"status": "ok"})


app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.add_middleware(AuthMiddleware)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

if __name__ == "__main__":
    mcp.run()
