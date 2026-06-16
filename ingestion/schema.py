from __future__ import annotations

from pydantic import BaseModel, field_validator

# Cadence and category are validated against fixed lists to catch Prodie typos at handoff.
# Update these if Fortune adds new cadences or product categories to the rate card.
_VALID_CADENCES = {"annual", "quarterly", "monthly", "weekly"}
_VALID_CATEGORIES = {"Digital Media", "Newsletter", "Branded Content", "Events", "Print"}
_VALID_INDUSTRIES = {
    "Tech",
    "Financial Services",
    "Professional Services",
    "Energy",
    "Healthcare",
    "Luxury",
    "Economic Development",
}
_ESCALATION_THRESHOLD = 750_000

# Template filenames in the Fortune Sales Automation SharePoint folder.
# Industry templates are the default; franchise templates take precedence when a product name
# matches a known franchise keyword. Fall back to general if no match found.
# v1: replace these dicts with a config JSON loaded from S3 so GTM can update without a deploy.
_INDUSTRY_TEMPLATES: dict[str, str] = {
    "Tech": "Category_Presentation_Technology.pptx.url",
    "Financial Services": "Category_Presentation_Financial.pptx.url",
    "Healthcare": "Category_Presentation_Healthcare.pptx.url",
    "Luxury": "Category_Presentation_Luxury.pptx.url",
    "Energy": "Category_Presentation_Energy.pptx.url",
    "Professional Services": "Category_Presentation_Professional_Services.pptx.url",
    "Economic Development": "Category_Presentation_Economic_Development.pptx.url",
}

_FRANCHISE_KEYWORDS: dict[str, str] = {
    "Fortune 500": "Franchise_Presentation_Fortune_500.pptx.url",
    "Fortune Daily": "Franchise_Presentation_Fortune_Daily.pptx.url",
    "Fortune CFO": "Franchise_Presentation_Fortune_CFO.pptx.url",
    "Fortune CIO": "Franchise_Presentation_Fortune_CIO.pptx.url",
    "Crypto": "Franchise_Presentation_Crypto.pptx.url",
}

_GENERAL_TEMPLATE = "General_Presentation_Fortune_Overall.pptx.url"


class Product(BaseModel):
    name: str
    cadence: str
    price: float
    category: str

    @field_validator("cadence")
    @classmethod
    def validate_cadence(cls, v: str) -> str:
        if v not in _VALID_CADENCES:
            raise ValueError(
                f"cadence must be one of {sorted(_VALID_CADENCES)}, got {v!r}"
            )
        return v

    @field_validator("category")
    @classmethod
    def validate_category(cls, v: str) -> str:
        if v not in _VALID_CATEGORIES:
            raise ValueError(
                f"category must be one of {sorted(_VALID_CATEGORIES)}, got {v!r}"
            )
        return v


class DeckSchema(BaseModel):
    client_name: str
    industry: str
    budget_quarterly: float
    confirmed_products: list[Product]
    buyer_persona: str | None = None
    objective: str | None = None
    target_audience: str | None = None
    upsell: Product | None = None
    next_steps_contact: str | None = None
    tone_notes: str | None = None

    @field_validator("industry")
    @classmethod
    def validate_industry(cls, v: str) -> str:
        if v not in _VALID_INDUSTRIES:
            raise ValueError(
                f"industry must be one of {sorted(_VALID_INDUSTRIES)}, got {v!r}"
            )
        return v

    @field_validator("confirmed_products")
    @classmethod
    def validate_confirmed_products(cls, v: list[Product]) -> list[Product]:
        if not v:
            raise ValueError("confirmed_products must not be empty")
        return v

    @field_validator("budget_quarterly")
    @classmethod
    def validate_budget(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("budget_quarterly must be greater than zero")
        if v >= _ESCALATION_THRESHOLD:
            raise ValueError(
                f"Budget of ${v:,.0f} meets or exceeds the ${_ESCALATION_THRESHOLD:,.0f} "
                "quarterly threshold. Route this opportunity to the GTM team — do not "
                "generate a deck."
            )
        return v


def select_template(schema: DeckSchema) -> str:
    """Return the SharePoint template filename for a given schema.

    Franchise templates take precedence over industry templates when a confirmed
    product name contains a known franchise keyword. Falls back to the general
    template if no industry or franchise match is found.
    """
    all_product_names = " ".join(p.name for p in schema.confirmed_products)
    if schema.upsell:
        all_product_names += f" {schema.upsell.name}"

    for keyword, filename in _FRANCHISE_KEYWORDS.items():
        if keyword.lower() in all_product_names.lower():
            return filename

    if filename := _INDUSTRY_TEMPLATES.get(schema.industry):
        return filename

    return _GENERAL_TEMPLATE
