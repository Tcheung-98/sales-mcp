"""Tests for Prodie-selected product lock → DeckSchema."""

from __future__ import annotations

from datetime import date

import pytest

from ingestion.confirm_mix import (
    ConfirmMixError,
    ConfirmMixRequest,
    confirm_mix,
    confirm_mix_from_dict,
    derive_cadence,
)
from ingestion.gtm_ideation_catalog import GtmIdeationCatalog
from ingestion.inventory_workbook import InventoryWorkbook
from tests.logic_guide_fixtures import base_discovery_fields, build_workbook_bytes


@pytest.fixture
def catalogs() -> tuple[GtmIdeationCatalog, InventoryWorkbook]:
    data = build_workbook_bytes()
    return (
        GtmIdeationCatalog.from_xlsx_bytes(data),
        InventoryWorkbook.from_xlsx_bytes(data),
    )


def _payload(*products: dict, **discovery_overrides) -> dict:
    return {
        "discovery": base_discovery_fields(**discovery_overrides),
        "selected_products": list(products),
    }


def test_direct_selection_resolves_authoritative_product(catalogs):
    gtm, inventory = catalogs
    request = ConfirmMixRequest.model_validate(
        _payload({"name": "CEO Daily", "category": "Newsletters"})
    )

    deck, warnings = confirm_mix(request, gtm=gtm, inventory=inventory)

    assert len(deck.confirmed_products) == 1
    product = deck.confirmed_products[0]
    assert product.name == "CEO Daily"
    assert product.category == "Newsletter"
    assert product.price == 5_000
    assert product.cadence == "weekly"
    assert warnings


def test_client_price_is_not_accepted_or_trusted(catalogs):
    gtm, inventory = catalogs
    payload = _payload(
        {
            "name": "Crown Unit",
            "category": "Digital Ads/Programmatic",
            "price": 1,
            "cadence": "annual",
        },
        budgets=[{"amount": 25_000}],
    )

    result = confirm_mix_from_dict(payload, gtm=gtm, inventory=inventory)

    assert result["status"] == "ok"
    assert result["confirmed_products"][0]["price"] == 25_000
    assert result["confirmed_products"][0]["cadence"] == "quarterly"


def test_category_disambiguates_duplicate_product_name(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "Term Sheet", "category": "Vodcasts"},
            budgets=[{"amount": 6_000}],
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "ok"
    assert result["confirmed_products"][0]["category"] == "Vodcasts"


def test_ambiguous_name_requires_category(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload({"name": "Term Sheet"}),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "error"
    assert "multiple categories" in result["message"]


def test_unknown_product_fails_loud(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload({"name": "Invented Product"}),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "error"
    assert "exact Product Name match required" in result["message"]


def test_duplicate_selection_fails_loud(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            {"name": "CEO Daily", "category": "Newsletters"},
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "error"
    assert "Duplicate selected product" in result["message"]


def test_empty_selection_is_incomplete(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "incomplete"
    assert "selected_products" in result["errors"][0]


def test_sold_product_is_rejected(catalogs):
    gtm, _inventory = catalogs
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
    inventory = InventoryWorkbook.from_xlsx_bytes(sold)

    result = confirm_mix_from_dict(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            flight_dates={"start": "2026-09-01", "end": "2026-09-01"},
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "error"
    assert "unavailable product" in result["message"]


def test_non_inventory_gated_product_is_allowed(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "Crown Unit", "category": "Digital Ads/Programmatic"},
            budgets=[{"amount": 25_000}],
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "ok"


@pytest.mark.parametrize(
    "platform",
    ["Conference Sponsorship/Media", "Lists & Rankings Sponsorship"],
)
def test_gtm_platforms_escalate(platform, catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            preferred_platforms_products=[platform, "Newsletters"],
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "escalation"
    assert platform in result["message"]


def test_budget_escalation_is_preserved(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            budgets=[{"amount": 750_000}],
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "escalation"


def test_mix_total_and_budget_warning(catalogs):
    gtm, inventory = catalogs

    result = confirm_mix_from_dict(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            {"name": "Crown Unit", "category": "Digital Ads/Programmatic"},
            budgets=[{"amount": 30_000}],
        ),
        gtm=gtm,
        inventory=inventory,
    )

    assert result["status"] == "ok"
    assert result["mix_total"] == 30_000
    assert result["warnings"] == []


@pytest.mark.parametrize(
    ("name", "pricing", "expected"),
    [
        ("CEO Daily", "$5,000/day", "weekly"),
        ("Product", "$10,000/week", "weekly"),
        ("Product", "$10,000/month", "monthly"),
        ("Product", "$10,000 annual", "annual"),
        ("Product", "$10,000", "quarterly"),
    ],
)
def test_derive_cadence(name, pricing, expected):
    assert derive_cadence(name, pricing) == expected


def test_confirm_mix_raises_for_escalation_platform(catalogs):
    gtm, inventory = catalogs
    request = ConfirmMixRequest.model_validate(
        _payload(
            {"name": "CEO Daily", "category": "Newsletters"},
            preferred_platforms_products=["Conference Sponsorship/Media"],
        )
    )

    with pytest.raises(ConfirmMixError, match="GTM escalation"):
        confirm_mix(request, gtm=gtm, inventory=inventory)
