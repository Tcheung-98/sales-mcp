import os

import boto3
import requests
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from pydantic import ValidationError
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

from ingestion.confirm_mix import confirm_mix_from_dict
from ingestion.generator import DeckGenerator
from ingestion.gtm_ideation_catalog import (
    GtmIdeationCatalog,
    load_gtm_ideation_catalog_from_s3,
)
from ingestion.inventory_workbook import (
    InventoryWorkbook,
    load_inventory_workbook_from_s3,
)
from ingestion.retriever import SlideRetriever
from ingestion.schema import BUDGET_ESCALATION_ERROR, DeckSchema

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
_ideation_gtm: GtmIdeationCatalog | None = None
_ideation_inventory: InventoryWorkbook | None = None


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


def _get_product_catalogs() -> tuple[GtmIdeationCatalog, InventoryWorkbook]:
    """Load authoritative GTM, pricing, and availability data for validation."""
    global _ideation_gtm, _ideation_inventory
    if _ideation_gtm is None or _ideation_inventory is None:
        s3 = boto3.client("s3")
        bucket = os.environ["S3_SNAPSHOT_BUCKET"]
        _ideation_gtm = load_gtm_ideation_catalog_from_s3(s3, bucket)
        _ideation_inventory = load_inventory_workbook_from_s3(s3, bucket)
    assert _ideation_gtm is not None and _ideation_inventory is not None
    return _ideation_gtm, _ideation_inventory


@mcp.tool()
def confirm_mix(
    discovery: dict,
    selected_products: list[dict],
) -> dict:
    """
    Validate and lock the associate's final Prodie checkbox selection.

    This tool does not propose, rank, score, fund, swap, or add products. Pass the
    complete final list as [{name, category?}, ...]. Product names must exactly
    match GTM Product Tags; category disambiguates duplicate names. The server
    resolves authoritative price/cadence, validates flight availability, and
    returns deck_schema ready for build_deck.
    """
    gtm, inventory = _get_product_catalogs()
    payload = {
        "discovery": discovery,
        "selected_products": selected_products,
    }
    return confirm_mix_from_dict(payload, gtm=gtm, inventory=inventory)


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
def build_deck(schema: dict, template_url: str | None = None) -> dict:
    """
    Assemble a Fortune pitch deck from a confirmed schema (C1 spine + C2 fills).

    Does not choose products — pass a seller-locked mix (Prodie + I3). Always uses
    FortuneAI_DeckTemplate as the Creation spine (intro / narrative /
    conditional category dividers / investment / thank you). Product pages are
    exact GTM Product Tags clones (Deck Path + Slide #).

    template_url: optional pre-authenticated HTTPS download URL for
    FortuneAI_DeckTemplate.pptx (SharePoint). When omitted, the template is loaded
    from S3 (FORTUNEAI_TEMPLATE_KEY, default templates/FortuneAI_DeckTemplate.pptx).

    After assembly, C2 fills run: date, logo, history client name, audience
    Reach/Index (Audience Data sheet), program category labels, investment blocks,
    thank-you date/logo, bounded Claude copy for intro title, Opportunity
    header/body, audience title, and program one-liners. Unused audience/program
    variant pages are dropped. No stylist (PI-2754 shelved).

    Uploads to S3 and returns a presigned download URL (24h) plus optional
    warnings[] (e.g. >6 audience segments truncated to the 6-card page).

    Fail loud (status: error): missing/ambiguous GTM Product Tags row, unmapped
    category, unreadable client_logo (HTTPS required), <2 audience segment matches,
    matched segment missing from Audience Data, stated budget vs mix total mismatch,
    investment category box clone failure, AI validation failure after retry.
    Missing/ambiguous map rows never fall back to Titan/RAG substitute.
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
        if any(e["type"] == BUDGET_ESCALATION_ERROR for e in exc.errors()):
            return {"status": "escalation", "message": errors[0]}
        return {"status": "incomplete", "missing": missing, "errors": errors}
    try:
        return _get_generator().build(parsed, template_url)
    except ValueError as exc:
        return {"status": "error", "message": str(exc)}
    except requests.RequestException as exc:
        return {"status": "error", "message": f"Failed to fetch template: {exc}"}


async def health(request):
    return JSONResponse({"status": "ok"})


app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.add_middleware(AuthMiddleware)

if __name__ == "__main__":
    mcp.run()
