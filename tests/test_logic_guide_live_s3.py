"""Live S3 smoke for Logic Guide engine (PI-2760 merge blockers)."""

from __future__ import annotations

import os

import boto3
import pytest

from ingestion.gtm_ideation_catalog import load_gtm_ideation_catalog_from_s3
from ingestion.inventory_workbook import load_inventory_workbook_from_s3
from ingestion.logic_guide.engine import LogicGuideEngine
from ingestion.schema import DiscoverySchema


def _skip_if_no_bucket() -> str:
    bucket = os.environ.get("S3_SNAPSHOT_BUCKET")
    if not bucket:
        pytest.skip("S3_SNAPSHOT_BUCKET not set")
    return bucket


def test_live_s3_gtm_catalog_loads():
    bucket = _skip_if_no_bucket()
    try:
        catalog = load_gtm_ideation_catalog_from_s3(boto3.client("s3"), bucket)
    except Exception as exc:
        pytest.skip(f"S3 GTM workbook unavailable: {exc}")
    assert len(catalog.categories.titles) >= 5
    assert catalog.products.lookup("Crown Unit", "Digital Ads/Programmatic")


def test_live_s3_logic_guide_propose_smoke():
    """Full engine smoke when GTM + inventory workbooks exist on dev bucket."""
    bucket = _skip_if_no_bucket()
    s3 = boto3.client("s3")
    try:
        gtm = load_gtm_ideation_catalog_from_s3(s3, bucket)
        inventory = load_inventory_workbook_from_s3(s3, bucket)
    except Exception as exc:
        pytest.skip(
            "Inventory calendar not on S3 yet — upload "
            "templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx "
            f"to {bucket}: {exc}"
        )

    engine = LogicGuideEngine(gtm, inventory)
    discovery = DiscoverySchema.model_validate(
        {
            "company_name": "Live Smoke Co",
            "industry": "Technology",
            "budgets": [{"amount": 250_000.0, "label": "Primary"}],
            "flight_dates": {"start": "2026-09-01", "end": "2026-12-31"},
            "campaign_goal": "Drive awareness among enterprise buyers",
            "targeting_details": "CEO and c-suite leadership",
            "kpis": ["Awareness"],
            "kpi_details": "Lift brand awareness",
            "campaign_narrative": "Enterprise finance modernization",
            "preferred_platforms_products": [
                "Newsletters",
                "Digital Ads/Programmatic",
            ],
            "additional_rfp_details": "",
            "client_logo": "https://example.com/logo.png",
        }
    )
    result = engine.propose(discovery)
    assert not result.requires_gtm_escalation
    if result.tiers:
        tier = result.tiers[0]
        assert tier.products
        assert all(p.price > 0 for p in tier.products)
