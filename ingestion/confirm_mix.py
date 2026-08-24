"""Associate confirm/swap → lock mix (PI-2761 / I3).

Pure functions bridging Logic Guide ``IdeationResult`` to Creation ``DeckSchema``.
"""

from __future__ import annotations

from typing import Any

from pydantic import AliasChoices, BaseModel, ConfigDict, Field, model_validator

from ingestion.gtm_ideation_catalog import GtmIdeationCatalog
from ingestion.inventory_workbook import InventoryWorkbook
from ingestion.logic_guide.media_mix import to_proposed_product
from ingestion.logic_guide.models import IdeationResult, ProposedProduct, TierProposal
from ingestion.placeholder_fills import format_usd, mix_total, stated_total_budget
from ingestion.schema import DeckSchema, DiscoverySchema, Product


class ConfirmMixError(ValueError):
    """Seller confirmation rejected — fail loud."""


class ProductSwap(BaseModel):
    """Swap one funded product for another from the proposal pool."""

    model_config = ConfigDict(populate_by_name=True)

    from_name: str = Field(validation_alias=AliasChoices("from_name", "from"))
    to_name: str = Field(validation_alias=AliasChoices("to_name", "to"))
    from_category: str | None = None
    to_category: str | None = None


class ManualAdd(BaseModel):
    """Add a product from the GTM catalog (priced via inventory workbook)."""

    name: str
    category: str | None = None


class ConfirmMixRequest(BaseModel):
    """Seller lock request after ``propose_mix``."""

    discovery: DiscoverySchema
    ideation: dict[str, Any]
    tier_index: int | None = None
    budget_target: float | None = None
    tier_label: str | None = None
    drop_products: list[str] = Field(default_factory=list)
    swaps: list[ProductSwap] = Field(default_factory=list)
    add_products: list[ManualAdd] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_tier_selector(self) -> ConfirmMixRequest:
        selectors = [
            self.tier_index is not None,
            self.budget_target is not None,
            self.tier_label is not None,
        ]
        if sum(selectors) != 1:
            raise ValueError(
                "Exactly one of tier_index, budget_target, or tier_label is required"
            )
        return self


def serialize_ideation(result: IdeationResult) -> dict[str, Any]:
    """JSON-serializable ``IdeationResult`` for MCP handoff."""
    return {
        "tiers": [
            {
                "budget_target": tier.budget_target,
                "label": tier.label,
                "total": tier.total,
                "products": [
                    {
                        "name": p.name,
                        "category": p.category,
                        "gtm_category": p.gtm_category,
                        "price": p.price,
                        "pricing_text": p.pricing_text,
                        "cadence": p.cadence,
                    }
                    for p in tier.products
                ],
            }
            for tier in result.tiers
        ],
        "escalations": list(result.escalations),
        "unavailable_products": list(result.unavailable_products),
        "notes": list(result.notes),
    }


def deserialize_ideation(data: dict[str, Any]) -> IdeationResult:
    """Rebuild ``IdeationResult`` from ``serialize_ideation`` output."""
    tiers: list[TierProposal] = []
    for raw_tier in data.get("tiers") or []:
        products = tuple(
            ProposedProduct(
                name=p["name"],
                category=p["category"],
                gtm_category=p["gtm_category"],
                price=float(p["price"]),
                pricing_text=p.get("pricing_text", ""),
                cadence=p.get("cadence", "quarterly"),
            )
            for p in raw_tier.get("products") or []
        )
        tiers.append(
            TierProposal(
                budget_target=float(raw_tier["budget_target"]),
                label=raw_tier.get("label"),
                products=products,
                total=float(raw_tier.get("total", sum(p.price for p in products))),
            )
        )
    return IdeationResult(
        tiers=tiers,
        escalations=list(data.get("escalations") or []),
        unavailable_products=list(data.get("unavailable_products") or []),
        notes=list(data.get("notes") or []),
    )


def derive_cadence(proposed: ProposedProduct) -> str:
    """Map pricing text / product type to schema cadence (MVP rules)."""
    text = (proposed.pricing_text or "").casefold()
    name = proposed.name.casefold()
    if "/day" in text or "daily" in name:
        return "weekly"
    if "/week" in text:
        return "weekly"
    if "/month" in text:
        return "monthly"
    if proposed.cadence in {"annual", "quarterly", "monthly", "weekly"}:
        return proposed.cadence
    return "quarterly"


def proposed_to_product(proposed: ProposedProduct) -> Product:
    """Map Logic Guide proposal row to Creation ``Product``."""
    return Product(
        name=proposed.name,
        cadence=derive_cadence(proposed),
        price=proposed.price,
        category=proposed.category,
    )


def _product_key(proposed: ProposedProduct) -> tuple[str, str]:
    return (proposed.name, proposed.gtm_category)


def proposal_pool(ideation: IdeationResult) -> dict[tuple[str, str], ProposedProduct]:
    """All proposed products across tiers, keyed by (name, gtm_category)."""
    pool: dict[tuple[str, str], ProposedProduct] = {}
    for tier in ideation.tiers:
        for product in tier.products:
            pool[_product_key(product)] = product
    return pool


def select_tier(
    ideation: IdeationResult,
    *,
    tier_index: int | None = None,
    budget_target: float | None = None,
    tier_label: str | None = None,
) -> TierProposal:
    """Pick one tier from ideation output."""
    if not ideation.tiers:
        raise ConfirmMixError("No funded tiers to confirm")

    if tier_index is not None:
        if tier_index < 0 or tier_index >= len(ideation.tiers):
            raise ConfirmMixError(
                f"tier_index {tier_index} out of range (0–{len(ideation.tiers) - 1})"
            )
        return ideation.tiers[tier_index]

    if budget_target is not None:
        matches = [
            t for t in ideation.tiers if abs(t.budget_target - budget_target) < 0.005
        ]
        if not matches:
            available = [t.budget_target for t in ideation.tiers]
            raise ConfirmMixError(
                f"No tier with budget_target {budget_target!r} "
                f"(available: {available})"
            )
        if len(matches) > 1:
            raise ConfirmMixError(
                f"Ambiguous budget_target {budget_target!r}: "
                f"{len(matches)} matching tiers"
            )
        return matches[0]

    assert tier_label is not None
    label_cf = tier_label.casefold()
    matches = [
        t
        for t in ideation.tiers
        if t.label and t.label.casefold() == label_cf
    ]
    if not matches:
        labels = [t.label for t in ideation.tiers if t.label]
        raise ConfirmMixError(
            f"No tier with label {tier_label!r} (available labels: {labels})"
        )
    if len(matches) > 1:
        raise ConfirmMixError(
            f"Ambiguous tier_label {tier_label!r}: {len(matches)} matching tiers"
        )
    return matches[0]


def _find_in_tier(
    products: list[ProposedProduct],
    name: str,
    category: str | None,
) -> ProposedProduct:
    matches = [p for p in products if p.name == name]
    if category:
        matches = [p for p in matches if p.gtm_category == category]
    if not matches:
        hint = f" in category {category!r}" if category else ""
        raise ConfirmMixError(f"Product {name!r}{hint} not in selected tier")
    if len(matches) > 1:
        cats = sorted({p.gtm_category for p in matches})
        raise ConfirmMixError(
            f"Ambiguous product {name!r} in tier — pass from_category "
            f"(matches: {cats})"
        )
    return matches[0]


def _lookup_pool_product(
    pool: dict[tuple[str, str], ProposedProduct],
    name: str,
    category: str | None,
) -> ProposedProduct:
    if category:
        key = (name, category)
        if key not in pool:
            raise ConfirmMixError(
                f"Swap/add target {name!r} in category {category!r} "
                "not in proposal pool (MVP: pick from proposal only for swaps)"
            )
        return pool[key]
    matches = [p for p in pool.values() if p.name == name]
    if not matches:
        raise ConfirmMixError(
            f"Product {name!r} not in proposal pool "
            "(MVP: swaps must target a proposed product)"
        )
    if len(matches) > 1:
        cats = sorted({p.gtm_category for p in matches})
        raise ConfirmMixError(
            f"Ambiguous proposal pool match for {name!r} — pass to_category "
            f"(available: {cats})"
        )
    return matches[0]


def _manual_proposed_product(
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
    name: str,
    category: str | None,
) -> ProposedProduct:
    row = gtm.products.lookup(name, category)
    pricing_row = inventory.pricing.lookup(name)
    amount = inventory.pricing.primary_amount(name)
    return to_proposed_product(
        row.product_name,
        row.category,
        amount,
        pricing_row.pricing,
    )


def _validate_gtm_row(
    gtm: GtmIdeationCatalog,
    proposed: ProposedProduct,
) -> None:
    gtm.products.lookup(proposed.name, proposed.gtm_category)


def apply_confirm_edits(
    tier: TierProposal,
    ideation: IdeationResult,
    request: ConfirmMixRequest,
    *,
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
) -> list[ProposedProduct]:
    """Apply drops, swaps, and manual adds to the selected tier."""
    pool = proposal_pool(ideation)
    products = list(tier.products)
    unavailable = set(ideation.unavailable_products)

    for drop in request.drop_products:
        before = len(products)
        products = [p for p in products if p.name != drop]
        if len(products) == before:
            raise ConfirmMixError(f"drop_products: {drop!r} not in selected tier")

    for swap in request.swaps:
        _find_in_tier(products, swap.from_name, swap.from_category)
        replacement = _lookup_pool_product(pool, swap.to_name, swap.to_category)
        products = [
            replacement if p.name == swap.from_name and (
                swap.from_category is None or p.gtm_category == swap.from_category
            )
            else p
            for p in products
        ]
        # Deduplicate same (name, gtm_category) after swap
        seen: set[tuple[str, str]] = set()
        deduped: list[ProposedProduct] = []
        for p in products:
            key = _product_key(p)
            if key in seen:
                continue
            seen.add(key)
            deduped.append(p)
        products = deduped

    for add in request.add_products:
        if add.name in unavailable:
            raise ConfirmMixError(
                f"Cannot add unavailable product {add.name!r} "
                "(explicit override not supported in MVP)"
            )
        proposed = _manual_proposed_product(gtm, inventory, add.name, add.category)
        if _product_key(proposed) in {_product_key(p) for p in products}:
            raise ConfirmMixError(
                f"add_products: {add.name!r} already in confirmed mix"
            )
        products.append(proposed)

    if not products:
        raise ConfirmMixError("Empty mix after drops — at least one product required")

    for proposed in products:
        if proposed.name in unavailable:
            raise ConfirmMixError(
                f"Cannot confirm unavailable product {proposed.name!r} "
                "(explicit override not supported in MVP)"
            )
        _validate_gtm_row(gtm, proposed)

    return products


def budget_warnings(deck: DeckSchema, selected_tier: TierProposal) -> list[str]:
    """Surface mix vs selected tier / stated budget mismatches."""
    warnings: list[str] = []
    mix = mix_total(deck)
    if abs(mix - selected_tier.budget_target) > 0.005:
        warnings.append(
            f"Mix total {format_usd(mix)} differs from selected tier budget "
            f"{format_usd(selected_tier.budget_target)}"
        )
    stated = stated_total_budget(deck)
    if abs(stated - mix) > 0.005:
        warnings.append(
            f"Stated total budget {format_usd(stated)} differs from mix total "
            f"{format_usd(mix)} — build_deck will fail until aligned"
        )
    return warnings


def confirm_mix(
    request: ConfirmMixRequest,
    *,
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
) -> tuple[DeckSchema, list[str], TierProposal]:
    """Lock seller edits → ``DeckSchema`` ready for ``build_deck``."""
    ideation = deserialize_ideation(request.ideation)
    if ideation.requires_gtm_escalation:
        raise ConfirmMixError(
            "Ideation requires GTM escalation — cannot confirm mix: "
            + "; ".join(ideation.escalations)
        )

    tier = select_tier(
        ideation,
        tier_index=request.tier_index,
        budget_target=request.budget_target,
        tier_label=request.tier_label,
    )
    confirmed = apply_confirm_edits(
        tier, ideation, request, gtm=gtm, inventory=inventory
    )
    deck = DeckSchema(
        **request.discovery.model_dump(),
        confirmed_products=[proposed_to_product(p) for p in confirmed],
    )
    warnings = budget_warnings(deck, tier)
    for note in ideation.notes:
        warnings.append(f"Ideation note: {note}")
    return deck, warnings, tier


def confirm_mix_from_dict(
    payload: dict[str, Any],
    *,
    gtm: GtmIdeationCatalog,
    inventory: InventoryWorkbook,
) -> dict[str, Any]:
    """MCP-friendly wrapper — returns status dict."""
    try:
        request = ConfirmMixRequest.model_validate(payload)
    except Exception as exc:
        return {"status": "incomplete", "errors": [str(exc)]}

    ideation = deserialize_ideation(request.ideation)
    if ideation.requires_gtm_escalation:
        return {
            "status": "escalation",
            "escalations": list(ideation.escalations),
            "message": "GTM escalation — associate cannot lock mix for auto-generation",
        }

    try:
        deck, warnings, tier = confirm_mix(
            request, gtm=gtm, inventory=inventory
        )
    except ConfirmMixError as exc:
        return {"status": "error", "message": str(exc)}

    return {
        "status": "ok",
        "deck_schema": deck.model_dump(mode="json"),
        "warnings": warnings,
        "selected_tier": {
            "budget_target": tier.budget_target,
            "label": tier.label,
            "total": tier.total,
            "product_names": [p.name for p in tier.products],
        },
        "confirmed_products": [
            p.model_dump() for p in deck.confirmed_products
        ],
    }
