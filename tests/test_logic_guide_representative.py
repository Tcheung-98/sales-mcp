"""Representative Logic Guide scenarios (PI-2760 merge blockers)."""

from __future__ import annotations

from datetime import date

import pytest

from ingestion.gtm_ideation_catalog import GtmIdeationCatalog
from ingestion.inventory_workbook import InventoryWorkbook
from ingestion.logic_guide.candidates import gather_candidates
from ingestion.logic_guide.platforms import selected_gtm_categories
from ingestion.schema import DiscoverySchema
from tests.logic_guide_fixtures import (
    REPRESENTATIVE_INVENTORY_SLOTS,
    base_discovery_fields,
    build_representative_engine,
    build_workbook_bytes,
)


def _discovery(**overrides) -> DiscoverySchema:
    return DiscoverySchema.model_validate(base_discovery_fields(**overrides))


@pytest.fixture
def engine():
    return build_representative_engine()


def test_multi_category_representative_mix(engine):
    discovery = _discovery(
        preferred_platforms_products=[
            "Newsletters",
            "Digital Ads/Programmatic",
            "Branded Content",
        ],
        budgets=[{"amount": 500_000.0, "label": "Primary"}],
        campaign_narrative=(
            "CEO leadership audience with documentary video film "
            "and written thought leadership article"
        ),
    )
    result = engine.propose(discovery)
    assert not result.requires_gtm_escalation
    assert result.tiers
    cats = {p.gtm_category for p in result.tiers[0].products}
    assert "Newsletters" in cats
    assert "Digital Ads/Programmatic" in cats
    assert "Branded Content" in cats


def test_tight_budget_surfaces_note_not_silent_fix(engine):
    discovery = _discovery(
        budgets=[{"amount": 30_000.0, "label": "Primary"}],
        targeting_details="CEO and c-suite leadership dealmakers venture capital",
    )
    result = engine.propose(discovery)
    tier = result.tiers[0]
    assert tier.total > tier.budget_target
    assert any("cannot fund mandatory" in note.casefold() for note in result.notes)


def test_linkedin_social_anchor(engine):
    discovery = _discovery(
        preferred_platforms_products=["Digital Ads/Programmatic"],
        targeting_details="LinkedIn paid social campaign for enterprise B2B buyers",
    )
    result = engine.propose(discovery)
    names = {p.name for p in result.tiers[0].products}
    assert "LinkedIn BrandLink" in names
    assert "Paid Social Video" not in names


def test_print_only_full_page(engine):
    discovery = _discovery(preferred_platforms_products=["Print"])
    result = engine.propose(discovery)
    tier = result.tiers[0]
    assert len(tier.products) == 1
    product = tier.products[0]
    assert product.name == "Full Page"
    assert product.price == 35_000.0
    assert product.gtm_category == "Print"


def test_bc_video_written_cap(engine):
    discovery = _discovery(
        preferred_platforms_products=["Branded Content"],
        campaign_narrative=(
            "Documentary-style video film plus written article thought leadership profile"
        ),
    )
    result = engine.propose(discovery)
    tier = result.tiers[0]
    video = [p for p in tier.products if "video" in p.name.casefold()]
    written = [p for p in tier.products if p not in video]
    assert len(video) == 1
    assert len(written) == 1
    assert video[0].name == "Documentary-Style Video"
    assert written[0].name == "Long-Form Q&A Article"


def test_format_collision_term_sheet_in_newsletter_and_vodcast(engine):
    discovery = _discovery(
        preferred_platforms_products=["Newsletters", "Vodcasts"],
        targeting_details="dealmakers venture capital M&A finance audience",
    )
    gtm_categories = selected_gtm_categories(discovery.preferred_platforms_products)
    raw = gather_candidates(discovery, engine._gtm, gtm_categories)
    term_rows = [r for r in raw if r.product_name == "Term Sheet"]
    assert len(term_rows) == 2
    assert {r.category for r in term_rows} == {"Newsletters", "Vodcasts"}

    result = engine.propose(discovery)
    tier = result.tiers[0]
    term_funded = [p for p in tier.products if p.name == "Term Sheet"]
    assert len(term_funded) == 2
    assert {p.gtm_category for p in term_funded} == {"Newsletters", "Vodcasts"}


def test_inventory_blocked_product_dropped():
    sold_slots = [
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
    data = build_workbook_bytes(inventory_slots=sold_slots)
    from ingestion.logic_guide.engine import LogicGuideEngine

    engine = LogicGuideEngine(
        GtmIdeationCatalog.from_xlsx_bytes(data),
        InventoryWorkbook.from_xlsx_bytes(data),
    )
    discovery = _discovery(
        flight_dates={"start": "2026-09-01", "end": "2026-09-01"},
        targeting_details="CEO and c-suite leadership",
    )
    result = engine.propose(discovery)
    assert "CEO Daily" in result.unavailable_products
    funded_names = {p.name for p in result.tiers[0].products}
    assert "CEO Daily" not in funded_names


def test_inventory_available_product_kept(engine):
    discovery = _discovery(
        flight_dates={"start": "2026-09-01", "end": "2026-09-01"},
        targeting_details="CEO and c-suite leadership",
    )
    result = engine.propose(discovery)
    assert "CEO Daily" not in result.unavailable_products
    newsletter_funded = [
        p for p in result.tiers[0].products if p.gtm_category == "Newsletters"
    ]
    assert newsletter_funded
    assert newsletter_funded[0].name in {"CEO Daily", "Term Sheet"}


def test_default_inventory_fixture_has_ceo_daily_available_slot():
    """Sanity: representative slots include an Available CEO Daily row."""
    assert any(
        row[5] == "CEO Daily" and row[6] == "Available"
        for row in REPRESENTATIVE_INVENTORY_SLOTS
    )
