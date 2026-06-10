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


def build_arc(schema: DeckSchema) -> list[dict]:
    slots = []
    i = 0

    # Cover: always clones blank template slide 0, no corpus search needed
    slots.append({"slot": i, "role": "cover", "query": ""})
    i += 1

    opener_query = f"{schema.industry} audience reach market opportunity Fortune"
    if schema.buyer_persona:
        opener_query = f"{schema.buyer_persona} {opener_query}"
    slots.append({"slot": i, "role": "opener", "query": opener_query})
    i += 1

    for product in schema.confirmed_products:
        slots.append({
            "slot": i,
            "role": "product",
            "query": f"{product.name} {product.cadence} Fortune pitch",
        })
        i += 1

    if schema.upsell:
        slots.append({
            "slot": i,
            "role": "upsell",
            "query": f"{schema.upsell.name} upsell add-on Fortune",
        })
        i += 1

    slots.append({
        "slot": i,
        "role": "investment",
        "query": "pricing investment bundled proposal rate",
    })
    i += 1

    slots.append({
        "slot": i,
        "role": "next_steps",
        "query": "next steps timeline partnership proposal",
    })

    return slots
