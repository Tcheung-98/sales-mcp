"""Validate a Prodie-selected product list and lock it for Creation.

This module does not propose, rank, fund, swap, or otherwise choose products.
Prodie sends the associate's final checkbox selection; sales-mcp resolves each
exact GTM product against authoritative pricing and inventory, then emits the
``DeckSchema`` consumed by ``build_deck``.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field, ValidationError

from ingestion.gtm_ideation_catalog import GtmIdeationCatalog
from ingestion.inventory_workbook import InventoryWorkbook
from ingestion.placeholder_fills import format_usd, mix_total, stated_total_budget
from ingestion.schema import DeckSchema, DiscoverySchema, Product

_GTM_TO_SCHEMA_CATEGORY = {
    "Digital Ads/Programmatic": "Digital Media",
    "Newsletters": "Newsletter",
    "Vodcasts": "Vodcasts",
    "Branded Content": "Branded Content",
    "Print": "Print",
}
_ESCALATION_PLATFORMS = {
    "Conference Sponsorship/Media",
    "Lists & Rankings Sponsorship",
}


class ConfirmMixError(ValueError):
    """Seller confirmation rejected — fail loud."""


class SelectedProduct(BaseModel):
    """One product checked by the associate in Prodie."""

    name: str
    category: str | None = None


class ConfirmMixRequest(BaseModel):
    """Creation lock supplied directly by Prodie after checkbox confirmation."""

    discovery: DiscoverySchema
    selected_products: list[SelectedProduct] = Field(min_length=1)


def derive_cadence(name: str, pricing_text: str) -> str:
    """Derive Creation cadence from authoritative product/pricing text."""
    text = pricing_text.casefold()
    name = name.casefold()
    if "/day" in text or "daily" in name:
        return "weekly"
    if "/week" in text:
        return "weekly"
    if "/month" in text:
        return "monthly"
    if "/year" in text or "annual" in text:
        return "annual"
    return "quarterly"


def _resolve_selected_product(
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
    selected: SelectedProduct,
    discovery: DiscoverySchema,
) -> Product:
    """Resolve one exact GTM selection to authoritative Creation fields."""
    try:
        row = gtm.products.lookup(selected.name, selected.category)
        pricing = inventory.pricing.lookup(row.product_name)
        amount = inventory.pricing.primary_amount(row.product_name)
    except ValueError as exc:
        raise ConfirmMixError(str(exc)) from exc

    if not inventory.calendar.is_available_for_flight(
        row.product_name,
        discovery.flight_dates.start,
        discovery.flight_dates.end,
    ):
        raise ConfirmMixError(
            f"Cannot confirm unavailable product {row.product_name!r} for flight "
            f"{discovery.flight_dates.start}–{discovery.flight_dates.end}"
        )

    schema_category = _GTM_TO_SCHEMA_CATEGORY.get(row.category)
    if schema_category is None:
        raise ConfirmMixError(
            f"Product {row.product_name!r} has unsupported GTM category "
            f"{row.category!r} for Creation"
        )
    return Product(
        name=row.product_name,
        category=schema_category,
        price=amount,
        cadence=derive_cadence(row.product_name, pricing.pricing),
    )


def confirm_mix(
    request: ConfirmMixRequest,
    *,
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
) -> tuple[DeckSchema, list[str]]:
    """Validate Prodie's final checkbox selection and lock it for Creation."""
    escalations = sorted(
        set(request.discovery.preferred_platforms_products) & _ESCALATION_PLATFORMS
    )
    if escalations:
        raise ConfirmMixError(
            "GTM escalation required for selected platform(s): "
            + ", ".join(escalations)
        )

    products: list[Product] = []
    seen: set[tuple[str, str]] = set()
    for selected in request.selected_products:
        product = _resolve_selected_product(gtm, inventory, selected, request.discovery)
        key = (product.name, product.category)
        if key in seen:
            raise ConfirmMixError(
                f"Duplicate selected product {product.name!r} "
                f"in category {product.category!r}"
            )
        seen.add(key)
        products.append(product)

    deck = DeckSchema(
        **request.discovery.model_dump(),
        confirmed_products=products,
    )
    return deck, budget_warnings(deck)


def budget_warnings(deck: DeckSchema) -> list[str]:
    """Surface locked mix vs seller-stated budget; never mutate either."""
    warnings: list[str] = []
    mix = mix_total(deck)
    stated = stated_total_budget(deck)
    if abs(stated - mix) > 0.005:
        warnings.append(
            f"Stated total budget {format_usd(stated)} differs from mix total "
            f"{format_usd(mix)} — build_deck will fail until aligned"
        )
    return warnings


def confirm_mix_from_dict(
    payload: dict[str, Any],
    *,
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
) -> dict[str, Any]:
    """MCP-friendly wrapper — returns status dict."""
    try:
        request = ConfirmMixRequest.model_validate(payload)
    except ValidationError as exc:
        if any(error["type"] == "budget_escalation" for error in exc.errors()):
            return {"status": "escalation", "message": exc.errors()[0]["msg"]}
        return {"status": "incomplete", "errors": [str(exc)]}

    try:
        deck, warnings = confirm_mix(request, gtm=gtm, inventory=inventory)
    except ConfirmMixError as exc:
        if str(exc).startswith("GTM escalation required"):
            return {"status": "escalation", "message": str(exc)}
        return {"status": "error", "message": str(exc)}

    return {
        "status": "ok",
        "deck_schema": deck.model_dump(mode="json"),
        "warnings": warnings,
        "mix_total": mix_total(deck),
        "confirmed_products": [
            p.model_dump() for p in deck.confirmed_products
        ],
    }
