"""Unit tests for associate confirm/swap → DeckSchema (PI-2761 / I3)."""

from __future__ import annotations

import pytest

from ingestion.confirm_mix import (
    ConfirmMixError,
    ConfirmMixRequest,
    confirm_mix,
    confirm_mix_from_dict,
    derive_cadence,
    deserialize_ideation,
    proposed_to_product,
    serialize_ideation,
)
from ingestion.gtm_ideation_catalog import GtmIdeationCatalog
from ingestion.inventory_workbook import InventoryWorkbook
from ingestion.logic_guide.engine import LogicGuideEngine
from ingestion.logic_guide.models import IdeationResult, ProposedProduct
from ingestion.schema import DiscoverySchema, Product
from tests.logic_guide_fixtures import (
    base_discovery_fields,
    build_representative_engine,
    build_workbook_bytes,
)


@pytest.fixture
def engine() -> LogicGuideEngine:
    return build_representative_engine()


@pytest.fixture
def catalogs() -> tuple[GtmIdeationCatalog, InventoryWorkbook]:
    data = build_workbook_bytes()
    return (
        GtmIdeationCatalog.from_xlsx_bytes(data),
        InventoryWorkbook.from_xlsx_bytes(data),
    )


def _discovery(**overrides) -> DiscoverySchema:
    return DiscoverySchema.model_validate(base_discovery_fields(**overrides))


def _propose(engine: LogicGuideEngine, **overrides) -> IdeationResult:
    return engine.propose(_discovery(**overrides))


def test_serialize_roundtrip(engine: LogicGuideEngine):
    result = _propose(engine)
    restored = deserialize_ideation(serialize_ideation(result))
    assert len(restored.tiers) == len(result.tiers)
    assert restored.tiers[0].products[0].name == result.tiers[0].products[0].name


def test_proposed_to_product_cadence():
    proposed = ProposedProduct(
        name="CEO Daily",
        category="Newsletter",
        gtm_category="Newsletters",
        price=50_000.0,
        pricing_text="$5,000/day",
    )
    product = proposed_to_product(proposed)
    assert product.cadence == "weekly"
    assert product.category == "Newsletter"
    assert product.price == 50_000.0


def test_derive_cadence_quarterly_default():
    proposed = ProposedProduct(
        name="Crown Unit",
        category="Digital Media",
        gtm_category="Digital Ads/Programmatic",
        price=25_000.0,
        pricing_text="$25,000",
        cadence="quarterly",
    )
    assert derive_cadence(proposed) == "quarterly"


def test_confirm_tier_by_index(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(
        budgets=[
            {"amount": 500_000.0, "label": "High"},
            {"amount": 120_000.0, "label": "Low"},
        ]
    )
    ideation = engine.propose(discovery)
    assert len(ideation.tiers) == 2

    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=1,
    )
    deck, warnings, tier = confirm_mix(request, gtm=gtm, inventory=inventory)
    assert tier.label == "Low"
    assert deck.confirmed_products
    assert all(isinstance(p, Product) for p in deck.confirmed_products)


def test_confirm_tier_by_budget_target(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(budgets=[{"amount": 250_000.0, "label": "Primary"}])
    ideation = engine.propose(discovery)
    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        budget_target=250_000.0,
    )
    deck, _, tier = confirm_mix(request, gtm=gtm, inventory=inventory)
    assert tier.budget_target == 250_000.0
    assert len(deck.confirmed_products) >= 1


def test_drop_product(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery()
    ideation = engine.propose(discovery)
    tier = ideation.tiers[0]
    drop_name = tier.products[0].name
    remaining = len(tier.products) - 1

    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=0,
        drop_products=[drop_name],
    )
    deck, _, _ = confirm_mix(request, gtm=gtm, inventory=inventory)
    assert len(deck.confirmed_products) == remaining
    assert drop_name not in {p.name for p in deck.confirmed_products}


def test_swap_within_pool(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(
        targeting_details="CEO and c-suite leadership dealmakers venture capital",
    )
    ideation = engine.propose(discovery)
    tier = ideation.tiers[0]
    newsletter = next(p for p in tier.products if p.gtm_category == "Newsletters")

    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=0,
        swaps=[
            {
                "from": newsletter.name,
                "from_category": newsletter.gtm_category,
                "to": "Term Sheet",
                "to_category": "Newsletters",
            }
        ],
    )
    deck, _, _ = confirm_mix(request, gtm=gtm, inventory=inventory)
    names = {p.name for p in deck.confirmed_products}
    assert "Term Sheet" in names


def test_drop_all_raises(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(preferred_platforms_products=["Print"])
    ideation = engine.propose(discovery)
    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=0,
        drop_products=["Full Page"],
    )
    with pytest.raises(ConfirmMixError, match="Empty mix"):
        confirm_mix(request, gtm=gtm, inventory=inventory)


def test_unavailable_product_blocked(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    from datetime import date

    sold = build_workbook_bytes(
        inventory_slots=[
            (
                date(2026, 9, 1),
                "Tue",
                36,
                "Sep",
                "Newsletter",
                "CEO Daily",
                "Sold",
            ),
        ]
    )
    blocked_engine = LogicGuideEngine(
        GtmIdeationCatalog.from_xlsx_bytes(sold),
        InventoryWorkbook.from_xlsx_bytes(sold),
    )
    discovery = _discovery(
        flight_dates={"start": "2026-09-01", "end": "2026-09-01"},
        targeting_details="CEO and c-suite leadership",
    )
    ideation = blocked_engine.propose(discovery)
    assert "CEO Daily" in ideation.unavailable_products

    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=0,
    )
    # Tier should not include CEO Daily; if it does, confirm_mix blocks unavailable
    for p in ideation.tiers[0].products:
        if p.name in ideation.unavailable_products:
            sold_gtm = GtmIdeationCatalog.from_xlsx_bytes(sold)
            sold_inv = InventoryWorkbook.from_xlsx_bytes(sold)
            with pytest.raises(ConfirmMixError, match="unavailable"):
                confirm_mix(request, gtm=sold_gtm, inventory=sold_inv)
            return


def test_escalation_short_circuit(catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(
        preferred_platforms_products=["Conference Sponsorship/Media", "Newsletters"]
    )
    ideation = IdeationResult(
        escalations=["Conference requires GTM escalation"],
    )
    result = confirm_mix_from_dict(
        {
            "discovery": discovery.model_dump(mode="json"),
            "ideation": serialize_ideation(ideation),
            "tier_index": 0,
        },
        gtm=gtm,
        inventory=inventory,
    )
    assert result["status"] == "escalation"
    assert result["escalations"]


def test_conference_propose_escalation(engine: LogicGuideEngine):
    result = _propose(
        engine,
        preferred_platforms_products=["Conference Sponsorship/Media", "Newsletters"],
    )
    assert result.requires_gtm_escalation


def test_budget_warning_passthrough(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery(budgets=[{"amount": 30_000.0, "label": "Primary"}])
    ideation = engine.propose(discovery)
    request = ConfirmMixRequest(
        discovery=discovery,
        ideation=serialize_ideation(ideation),
        tier_index=0,
    )
    deck, warnings, _ = confirm_mix(request, gtm=gtm, inventory=inventory)
    assert deck.confirmed_products
    assert any("differs" in w.casefold() or "ideation note" in w.casefold() for w in warnings)


def test_confirm_mix_from_dict_ok(engine: LogicGuideEngine, catalogs):
    gtm, inventory = catalogs
    discovery = _discovery()
    ideation = engine.propose(discovery)
    result = confirm_mix_from_dict(
        {
            "discovery": discovery.model_dump(mode="json"),
            "ideation": serialize_ideation(ideation),
            "tier_index": 0,
        },
        gtm=gtm,
        inventory=inventory,
    )
    assert result["status"] == "ok"
    assert result["deck_schema"]["confirmed_products"]
    assert result["confirmed_products"]
