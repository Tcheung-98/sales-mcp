"""Media mix funding logic (Logic Guide V1 — simplified MVP pass)."""

from __future__ import annotations

from dataclasses import dataclass

from ingestion.logic_guide.models import ProposedProduct, TierProposal
from ingestion.logic_guide.platforms import GTM_TO_SCHEMA_CATEGORY
from ingestion.logic_guide.triggers import _SOCIAL_PHRASES, seller_text, text_contains_any
from ingestion.schema import BudgetTier, DiscoverySchema

_DIGITAL_CATEGORY = "Digital Ads/Programmatic"
_NEWSLETTERS_CATEGORY = "Newsletters"
_VODCASTS_CATEGORY = "Vodcasts"
_BC_CATEGORY = "Branded Content"
_PRINT_CATEGORY = "Print"

_DISPLAY_BACKFILL = (
    "Run of Fortune Display",
    "Contextually/Geotargeted Display",
    "Audience Targeted Display",
)


@dataclass(frozen=True)
class PricedCandidate:
    product: ProposedProduct


def _pick_cheapest(pool: list[PricedCandidate]) -> PricedCandidate | None:
    if not pool:
        return None
    return min(pool, key=lambda c: c.product.price)


def _pick_most_expensive(pool: list[PricedCandidate]) -> PricedCandidate | None:
    if not pool:
        return None
    return max(pool, key=lambda c: c.product.price)


def _by_gtm(pool: list[PricedCandidate], category: str) -> list[PricedCandidate]:
    return [c for c in pool if c.product.gtm_category == category]


def _find_named(pool: list[PricedCandidate], name: str) -> PricedCandidate | None:
    for item in pool:
        if item.product.name == name:
            return item
    return None


def _digital_anchor(
    discovery: DiscoverySchema, pool: list[PricedCandidate]
) -> PricedCandidate | None:
    text = seller_text(discovery)
    digital = _by_gtm(pool, _DIGITAL_CATEGORY)
    if text_contains_any(text, ("linkedin",)):
        linkedin = _find_named(digital, "LinkedIn BrandLink")
        if linkedin:
            return linkedin
    if text_contains_any(text, _SOCIAL_PHRASES):
        social = _find_named(digital, "Paid Social Video")
        if social:
            return social
    return _find_named(digital, "Crown Unit") or _pick_cheapest(digital)


def fund_tier(
    discovery: DiscoverySchema,
    budget: BudgetTier,
    priced: list[PricedCandidate],
    selected_gtm_categories: list[str],
) -> TierProposal:
    """Fund one tier using mandatory-minimum ordering + simple backfill."""
    remaining = budget.amount
    funded: list[PricedCandidate] = []
    pool = list(priced)

    def add_mandatory(item: PricedCandidate | None) -> bool:
        nonlocal remaining
        if item is None or item in funded:
            return False
        funded.append(item)
        remaining -= item.product.price
        return True

    def add_optional(item: PricedCandidate | None) -> bool:
        nonlocal remaining
        if item is None or item in funded:
            return False
        if item.product.price > remaining:
            return False
        funded.append(item)
        remaining -= item.product.price
        return True

    if _DIGITAL_CATEGORY in selected_gtm_categories:
        add_mandatory(_digital_anchor(discovery, pool))

    if _NEWSLETTERS_CATEGORY in selected_gtm_categories:
        news = _by_gtm(pool, _NEWSLETTERS_CATEGORY)
        add_mandatory(_pick_most_expensive(news) or _pick_cheapest(news))

    if _VODCASTS_CATEGORY in selected_gtm_categories:
        vod = _by_gtm(pool, _VODCASTS_CATEGORY)
        add_mandatory(_pick_most_expensive(vod) or _pick_cheapest(vod))

    if _BC_CATEGORY in selected_gtm_categories:
        bc = _by_gtm(pool, _BC_CATEGORY)
        _BC_VIDEO_NAMES = frozenset(
            {
                "Executive Q&A (Remote)",
                "Documentary-Style Video",
                "Hosted Interview Video (Remote)",
            }
        )
        video_pool = [
            c
            for c in bc
            if c.product.name in _BC_VIDEO_NAMES
            or "video" in c.product.name.casefold()
        ]
        written_pool = [c for c in bc if c not in video_pool]
        add_mandatory(_pick_most_expensive(video_pool))
        add_mandatory(_pick_most_expensive(written_pool))

    if _PRINT_CATEGORY in selected_gtm_categories:
        add_mandatory(_find_named(pool, "Full Page"))

    # Simple digital backfill: Scroller if ≥$25k left, else Crown top-up, else Display.
    if _DIGITAL_CATEGORY in selected_gtm_categories and remaining > 0:
        digital = _by_gtm(pool, _DIGITAL_CATEGORY)
        if remaining >= 25_000:
            add_optional(_find_named(digital, "Scroller Unit"))
        if remaining > 0:
            crown = _find_named(digital, "Crown Unit")
            if crown and crown not in funded and crown.product.price <= remaining:
                add_optional(crown)
        if remaining > 0:
            for name in _DISPLAY_BACKFILL:
                add_optional(_find_named(digital, name))

    products = tuple(c.product for c in funded)
    total = sum(p.price for p in products)
    return TierProposal(
        budget_target=budget.amount,
        label=budget.label,
        products=products,
        total=total,
    )


def build_lower_tier(upper: TierProposal, budget: BudgetTier) -> TierProposal:
    """Trim upper tier to fit lower budget (Logic Guide subset rule)."""
    products = list(upper.products)
    if not products:
        return TierProposal(
            budget_target=budget.amount,
            label=budget.label,
            products=(),
            total=0.0,
        )

    while products and sum(p.price for p in products) > budget.amount:
        cheapest = min(products, key=lambda p: p.price)
        if cheapest.price <= 0:
            break
        products.remove(cheapest)
        if sum(p.price for p in products) <= budget.amount:
            break
        if not any(p.price < cheapest.price for p in products):
            expensive = max(products, key=lambda p: p.price)
            products.remove(expensive)

    total = sum(p.price for p in products)
    return TierProposal(
        budget_target=budget.amount,
        label=budget.label,
        products=tuple(products),
        total=total,
    )


def to_proposed_product(
    name: str,
    gtm_category: str,
    price: float,
    pricing_text: str,
) -> ProposedProduct:
    schema_cat = GTM_TO_SCHEMA_CATEGORY.get(gtm_category, gtm_category)
    return ProposedProduct(
        name=name,
        category=schema_cat,
        gtm_category=gtm_category,
        price=price,
        pricing_text=pricing_text,
    )
