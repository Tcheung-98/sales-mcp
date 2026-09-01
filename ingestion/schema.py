from __future__ import annotations

from datetime import date
from typing import Any

from pydantic import (
    AliasChoices,
    BaseModel,
    ConfigDict,
    Field,
    field_validator,
    model_validator,
)
from pydantic_core import PydanticCustomError

# Cadence and category are validated against fixed lists to catch Prodie typos at handoff.
# Update these if Fortune adds new cadences or product categories to the rate card.
_VALID_CADENCES = {"annual", "quarterly", "monthly", "weekly"}
# Product mix categories (Creation). Vodcasts funds Premium Video divider (C1).
# Events remains valid on Product for Ideation/escalation but has no FortuneAI divider.
_VALID_CATEGORIES = {
    "Digital Media",
    "Newsletter",
    "Branded Content",
    "Events",
    "Print",
    "Vodcasts",
}

# Workflow Discovery industries (Pitch Deck Workflow Guide). Tech → Technology on ingest.
_VALID_INDUSTRIES = {
    "Technology",
    "Professional Services",
    "Healthcare",
    "Financial Services",
    "Energy",
    "Lifestyle",
    "Luxury",
}
_INDUSTRY_ALIASES = {
    "Tech": "Technology",
}

_VALID_KPIS = {
    "Brand Lift",
    "Viewability",
    "Awareness",
    "Engagement",
    "Lead Generation",
}

_VALID_PREFERRED_PLATFORMS = {
    "Branded Content",
    "Digital Ads/Programmatic",
    "Newsletters",
    "Vodcasts",
    "Lists & Rankings Sponsorship",
    "Print",
    "Conference Sponsorship/Media",
}

_ESCALATION_THRESHOLD = 750_000
BUDGET_ESCALATION_ERROR = "budget_escalation"
_MAX_BUDGET_TIERS = 3


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


class BudgetTier(BaseModel):
    """One of up to three Discovery budget tiers."""

    amount: float
    label: str | None = None

    @field_validator("amount")
    @classmethod
    def validate_amount(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("budget tier amount must be greater than zero")
        return v


class FlightDates(BaseModel):
    start: date
    end: date

    @model_validator(mode="after")
    def validate_range(self) -> FlightDates:
        if self.end < self.start:
            raise ValueError("flight_dates.end must be on or after flight_dates.start")
        return self


class DiscoverySchema(BaseModel):
    """Pitch Deck Workflow Discovery intake (no confirmed mix yet).

    Platform/product specifics are the only optional Workflow field.
    """

    model_config = ConfigDict(populate_by_name=True)

    company_name: str = Field(
        validation_alias=AliasChoices("company_name", "client_name"),
    )
    industry: str
    budgets: list[BudgetTier]
    flight_dates: FlightDates
    campaign_goal: str
    targeting_details: str
    kpis: list[str]
    kpi_details: str
    campaign_narrative: str
    preferred_platforms_products: list[str]
    additional_rfp_details: str
    client_logo: str
    platform_or_product_specifics: str | None = None

    # Soft / legacy Creation hints (optional; not Workflow Discovery required fields).
    buyer_persona: str | None = None
    objective: str | None = None
    target_audience: str | None = None
    upsell: Product | None = None
    next_steps_contact: str | None = None
    tone_notes: str | None = None

    @property
    def client_name(self) -> str:
        """Back-compat for assemble/build paths that still read client_name."""
        return self.company_name

    @model_validator(mode="before")
    @classmethod
    def coerce_legacy_budget(cls, data: Any) -> Any:
        if not isinstance(data, dict):
            return data
        if "budgets" not in data and "budget_quarterly" in data:
            data = {**data, "budgets": [{"amount": data["budget_quarterly"]}]}
        return data

    @field_validator("industry")
    @classmethod
    def validate_industry(cls, v: str) -> str:
        v = _INDUSTRY_ALIASES.get(v, v)
        if v not in _VALID_INDUSTRIES:
            raise ValueError(
                f"industry must be one of {sorted(_VALID_INDUSTRIES)}, got {v!r}"
            )
        return v

    @field_validator("budgets")
    @classmethod
    def validate_budgets(cls, v: list[BudgetTier]) -> list[BudgetTier]:
        if not v:
            raise ValueError("budgets must contain at least one tier")
        if len(v) > _MAX_BUDGET_TIERS:
            raise ValueError(
                f"budgets supports at most {_MAX_BUDGET_TIERS} tiers, got {len(v)}"
            )
        max_amount = max(tier.amount for tier in v)
        if max_amount >= _ESCALATION_THRESHOLD:
            raise PydanticCustomError(
                BUDGET_ESCALATION_ERROR,
                "Budget of ${amount:,.0f} meets or exceeds the ${threshold:,.0f} "
                "quarterly threshold. Route this opportunity to the GTM team — do not "
                "generate a deck.",
                {"amount": max_amount, "threshold": _ESCALATION_THRESHOLD},
            )
        return v

    @field_validator("kpis")
    @classmethod
    def validate_kpis(cls, v: list[str]) -> list[str]:
        if not v:
            raise ValueError("kpis must not be empty")
        invalid = sorted({kpi for kpi in v if kpi not in _VALID_KPIS})
        if invalid:
            raise ValueError(
                f"kpis must be chosen from {sorted(_VALID_KPIS)}; "
                f"invalid: {invalid}"
            )
        return v

    @field_validator("preferred_platforms_products")
    @classmethod
    def validate_preferred_platforms(cls, v: list[str]) -> list[str]:
        if not v:
            raise ValueError("preferred_platforms_products must not be empty")
        invalid = sorted({p for p in v if p not in _VALID_PREFERRED_PLATFORMS})
        if invalid:
            raise ValueError(
                f"preferred_platforms_products must be chosen from "
                f"{sorted(_VALID_PREFERRED_PLATFORMS)}; invalid: {invalid}"
            )
        return v


class DeckSchema(DiscoverySchema):
    """Creation handoff: Discovery intake + seller-confirmed product mix."""

    confirmed_products: list[Product]

    @field_validator("confirmed_products")
    @classmethod
    def validate_confirmed_products(cls, v: list[Product]) -> list[Product]:
        if not v:
            raise ValueError("confirmed_products must not be empty")
        return v
