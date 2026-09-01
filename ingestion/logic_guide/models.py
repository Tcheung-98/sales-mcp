"""Output models for Logic Guide ideation."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ProposedProduct:
    """One funded product in a tier proposal."""

    name: str
    category: str  # schema category (Digital Media, Newsletter, …)
    gtm_category: str
    price: float
    pricing_text: str
    cadence: str = "quarterly"


@dataclass(frozen=True)
class TierProposal:
    """Funded mix for one budget tier."""

    budget_target: float
    label: str | None
    products: tuple[ProposedProduct, ...]
    total: float


@dataclass
class IdeationResult:
    """Logic Guide output for one Discovery intake."""

    tiers: list[TierProposal] = field(default_factory=list)
    escalations: list[str] = field(default_factory=list)
    unavailable_products: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def requires_gtm_escalation(self) -> bool:
        return bool(self.escalations)
