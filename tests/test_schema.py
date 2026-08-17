import pytest
from pydantic import ValidationError

from ingestion.schema import DeckSchema, DiscoverySchema


def _discovery_fields(**overrides) -> dict:
    base = {
        "company_name": "Acme Corp",
        "industry": "Technology",
        "budgets": [{"amount": 50_000.0, "label": "Primary"}],
        "flight_dates": {"start": "2026-09-01", "end": "2026-12-31"},
        "campaign_goal": "Drive consideration among enterprise buyers",
        "targeting_details": "US enterprise tech decision-makers",
        "kpis": ["Awareness", "Engagement"],
        "kpi_details": "Lift brand awareness 10%; engagement rate above benchmark",
        "campaign_narrative": "Acme helps mid-market CFOs modernize finance ops",
        "preferred_platforms_products": ["Newsletters", "Branded Content"],
        "additional_rfp_details": "Prefer Q4 flight; avoid holiday blackout weeks",
        "client_logo": "https://example.com/acme-logo.png",
    }
    base.update(overrides)
    return base


def _valid_schema(**overrides) -> dict:
    base = {
        **_discovery_fields(),
        "confirmed_products": [
            {
                "name": "CIO Intelligence Newsletter",
                "cadence": "monthly",
                "price": 35_000.0,
                "category": "Newsletter",
            }
        ],
    }
    base.update(overrides)
    return base


def test_valid_schema_passes():
    schema = DeckSchema.model_validate(_valid_schema())
    assert schema.company_name == "Acme Corp"
    assert schema.client_name == "Acme Corp"
    assert schema.industry == "Technology"
    assert len(schema.budgets) == 1
    assert schema.budgets[0].amount == 50_000.0


def test_discovery_schema_without_mix():
    discovery = DiscoverySchema.model_validate(_discovery_fields())
    assert discovery.company_name == "Acme Corp"
    assert discovery.platform_or_product_specifics is None


def test_client_name_alias_accepted():
    data = _valid_schema()
    data.pop("company_name")
    data["client_name"] = "Alias Corp"
    schema = DeckSchema.model_validate(data)
    assert schema.company_name == "Alias Corp"
    assert schema.client_name == "Alias Corp"


def test_budget_quarterly_shim():
    data = _valid_schema()
    data.pop("budgets")
    data["budget_quarterly"] = 40_000.0
    schema = DeckSchema.model_validate(data)
    assert len(schema.budgets) == 1
    assert schema.budgets[0].amount == 40_000.0


def test_tech_industry_alias_normalizes():
    schema = DeckSchema.model_validate(_valid_schema(industry="Tech"))
    assert schema.industry == "Technology"


def test_multi_tier_budgets():
    schema = DeckSchema.model_validate(
        _valid_schema(
            budgets=[
                {"amount": 25_000.0, "label": "Good"},
                {"amount": 50_000.0, "label": "Better"},
                {"amount": 100_000.0, "label": "Best"},
            ]
        )
    )
    assert len(schema.budgets) == 3


def test_missing_required_field_raises():
    data = _valid_schema()
    del data["company_name"]
    with pytest.raises(ValidationError):
        DeckSchema.model_validate(data)


def test_budget_zero_raises():
    with pytest.raises(ValidationError, match="greater than zero"):
        DeckSchema.model_validate(
            _valid_schema(budgets=[{"amount": 0}])
        )


def test_budget_escalation_raises():
    with pytest.raises(ValidationError, match="GTM team"):
        DeckSchema.model_validate(
            _valid_schema(budgets=[{"amount": 750_000}])
        )


def test_budget_escalation_on_max_tier():
    with pytest.raises(ValidationError, match="GTM team"):
        DeckSchema.model_validate(
            _valid_schema(
                budgets=[
                    {"amount": 50_000.0},
                    {"amount": 800_000.0},
                ]
            )
        )


def test_too_many_budget_tiers_raises():
    with pytest.raises(ValidationError, match="at most 3"):
        DeckSchema.model_validate(
            _valid_schema(
                budgets=[
                    {"amount": 10_000.0},
                    {"amount": 20_000.0},
                    {"amount": 30_000.0},
                    {"amount": 40_000.0},
                ]
            )
        )


def test_confirmed_products_empty_raises():
    with pytest.raises(ValidationError):
        DeckSchema.model_validate(_valid_schema(confirmed_products=[]))


def test_invalid_kpi_raises():
    with pytest.raises(ValidationError, match="kpis"):
        DeckSchema.model_validate(_valid_schema(kpis=["Not A KPI"]))


def test_flight_dates_end_before_start_raises():
    with pytest.raises(ValidationError, match="on or after"):
        DeckSchema.model_validate(
            _valid_schema(
                flight_dates={"start": "2026-12-01", "end": "2026-01-01"}
            )
        )


def test_vodcasts_category_accepted():
    schema = DeckSchema.model_validate(
        _valid_schema(
            confirmed_products=[
                {
                    "name": "Fortune Tech Vodcast",
                    "cadence": "monthly",
                    "price": 25_000.0,
                    "category": "Vodcasts",
                }
            ]
        )
    )
    assert schema.confirmed_products[0].category == "Vodcasts"


def test_economic_development_industry_rejected():
    with pytest.raises(ValidationError, match="industry"):
        DiscoverySchema.model_validate(
            _discovery_fields(industry="Economic Development")
        )
