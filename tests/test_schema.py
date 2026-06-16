import pytest
from pydantic import ValidationError

from ingestion.schema import DeckSchema, select_template


def _valid_schema(**overrides) -> dict:
    base = {
        "client_name": "Acme Corp",
        "industry": "Tech",
        "budget_quarterly": 50_000.0,
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
    assert schema.client_name == "Acme Corp"
    assert schema.industry == "Tech"


def test_missing_required_field_raises():
    data = _valid_schema()
    del data["client_name"]
    with pytest.raises(ValidationError):
        DeckSchema.model_validate(data)


def test_budget_zero_raises():
    with pytest.raises(ValidationError, match="greater than zero"):
        DeckSchema.model_validate(_valid_schema(budget_quarterly=0))


def test_budget_escalation_raises():
    with pytest.raises(ValidationError, match="GTM team"):
        DeckSchema.model_validate(_valid_schema(budget_quarterly=750_000))


def test_confirmed_products_empty_raises():
    with pytest.raises(ValidationError):
        DeckSchema.model_validate(_valid_schema(confirmed_products=[]))


def test_select_template_industry_match():
    schema = DeckSchema.model_validate(_valid_schema(industry="Tech"))
    assert select_template(schema) == "Category_Presentation_Technology.pptx.url"


def test_select_template_franchise_takes_precedence():
    schema = DeckSchema.model_validate(
        _valid_schema(
            industry="Tech",
            confirmed_products=[
                {
                    "name": "Fortune CFO Newsletter",
                    "cadence": "monthly",
                    "price": 35_000.0,
                    "category": "Newsletter",
                }
            ],
        )
    )
    assert select_template(schema) == "Franchise_Presentation_Fortune_CFO.pptx.url"


def test_select_template_general_fallback():
    schema = DeckSchema.model_validate(
        _valid_schema(
            industry="Tech",
            confirmed_products=[
                {
                    "name": "Standard Display Ad",
                    "cadence": "monthly",
                    "price": 10_000.0,
                    "category": "Digital Media",
                }
            ],
        )
    )
    # No franchise keyword match — falls back to industry template
    assert select_template(schema) == "Category_Presentation_Technology.pptx.url"
