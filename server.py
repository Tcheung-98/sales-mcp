import os

import boto3
from botocore.exceptions import BotoCoreError, ClientError
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

from ingestion import faq as faq_module
from ingestion.generator import DeckGenerator
from ingestion.retriever import SlideRetriever

_EXPECTED_TOKEN = os.environ.get("MCP_SHARED_SECRET")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.url.path in ("/health", "/secrets-check"):
            return await call_next(request)
        # Fail closed if the secret was never injected into the environment.
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
_s3 = None

_S3_BUCKET = os.environ.get("S3_SNAPSHOT_BUCKET")
_RATE_SHEET_KEY = os.environ.get("RATE_SHEET_KEY")
_RULEBOOK_KEY = os.environ.get("RULEBOOK_KEY", "templates/rulebook.docx")


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
def search_sales_knowledge(query: str, k: int = 5) -> dict:
    """
    Search Fortune's sales knowledge base. Returns pricing/rate card data, sales
    guidelines, and relevant slide content for a given query. Use for any question
    about Fortune products, pricing, packaging, or sales approach.
    """
    global _s3
    if _s3 is None:
        _s3 = boto3.client("s3")

    return faq_module.search_sales_knowledge(
        query=query,
        retriever=_get_retriever(),
        s3=_s3,
        bucket=_S3_BUCKET,
        rate_sheet_key=_RATE_SHEET_KEY,
        rulebook_key=_RULEBOOK_KEY,
        k=k,
    )


def search_decks(query: str, k: int = 5) -> list[dict]:
    """Search Fortune sales decks by semantic similarity. Returns the k most relevant slides."""
    return _get_retriever().search(query, k=k)

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

def get_slide_content(deck_id: str, slide_numbers: list[int] | None = None) -> list[dict]:
    """
    Retrieve full slide content (title, body text, layout) for a specific deck.
    Call this after search_decks or filter_decks_by_tags returns a deck_id you want
    to read in detail. Pass slide_numbers to fetch specific slides; omit to get every
    slide in the deck. Returns an empty list if the deck_id is not found or the
    requested slide numbers don't exist — never raises an error.
    """
    return _get_retriever().get_slide_content(deck_id, slide_numbers)

def generate_deck(brief: str, k: int = 10) -> dict:
    """
    Generate a Fortune-branded PowerPoint deck from a plain-text brief. Retrieves the
    k most relevant slides from the corpus as grounding context, then calls Claude to
    author the slide content following Fortune's style. Returns a presigned S3 download
    URL (valid 24 hours), the S3 URI, slide count, and the original brief.
    Use this when an account executive provides a brief and wants a ready-to-use deck.
    """
    context_slides = _get_retriever().search(brief, k=k)
    return _get_generator().generate(brief, context_slides)


def hello(name: str) -> str:
    """Sanity check tool."""
    return f"Hello, {name}!"

async def health(request):
    return JSONResponse({"status": "ok"})

# Temporary to test IAM deployment and permissions. Will delete after confirming the secret check.
async def secrets_check(request):
    secret_name = "fortune-sales-mcp/claude-api-key"
    try:
        client = boto3.client("secretsmanager", region_name="us-east-1")
        client.get_secret_value(SecretId=secret_name)
        return JSONResponse({"status": "ok", "secret": secret_name})
    except ClientError as e:
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)
    except BotoCoreError as e:
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)

app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.routes.append(Route("/secrets-check", secrets_check))
app.add_middleware(AuthMiddleware)

if __name__ == "__main__":
    mcp.run()