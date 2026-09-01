"""Logic Guide V1 orchestrator."""

from __future__ import annotations

import logging

from ingestion.gtm_ideation_catalog import GtmIdeationCatalog, GtmProductCandidate
from ingestion.inventory_workbook import InventoryWorkbook
from ingestion.logic_guide.candidates import gather_candidates
from ingestion.logic_guide.media_mix import (
    PricedCandidate,
    build_lower_tier,
    fund_tier,
    to_proposed_product,
)
from ingestion.logic_guide.models import IdeationResult
from ingestion.logic_guide.platforms import (
    escalation_reasons,
    selected_gtm_categories,
)
from ingestion.schema import DiscoverySchema

logger = logging.getLogger(__name__)


class LogicGuideEngine:
    """Fortune Logic Guide V1 — Discovery → proposed mix by tier."""

    def __init__(
        self,
        gtm: GtmIdeationCatalog,
        inventory: InventoryWorkbook,
    ) -> None:
        self._gtm = gtm
        self._inventory = inventory

    def propose(self, discovery: DiscoverySchema) -> IdeationResult:
        result = IdeationResult()

        escalations = escalation_reasons(discovery.preferred_platforms_products)
        if escalations:
            result.escalations.extend(escalations)
            return result

        gtm_categories = selected_gtm_categories(
            discovery.preferred_platforms_products
        )
        if not gtm_categories:
            result.notes.append("No fundable platforms selected after escalation filter.")
            return result

        candidates = gather_candidates(discovery, self._gtm, gtm_categories)
        available, unavailable = self._filter_availability(discovery, candidates)
        result.unavailable_products.extend(unavailable)

        priced = self._price_candidates(available, result)
        if not priced:
            result.notes.append("No priced, available candidates for selected categories.")
            return result

        budgets = sorted(discovery.budgets, key=lambda b: b.amount, reverse=True)
        largest = fund_tier(discovery, budgets[0], priced, gtm_categories)
        if largest.total > largest.budget_target:
            result.notes.append(
                f"Largest tier ${largest.budget_target:,.0f} cannot fund mandatory "
                f"minimums (mix total ${largest.total:,.0f})."
            )
        result.tiers.append(largest)

        for tier_budget in budgets[1:]:
            trimmed = build_lower_tier(largest, tier_budget)
            result.tiers.append(trimmed)

        return result

    def _filter_availability(
        self,
        discovery: DiscoverySchema,
        candidates: list[GtmProductCandidate],
    ) -> tuple[list[GtmProductCandidate], list[str]]:
        start = discovery.flight_dates.start
        end = discovery.flight_dates.end
        available: list[GtmProductCandidate] = []
        unavailable: list[str] = []
        for row in candidates:
            if self._inventory.calendar.is_available_for_flight(
                row.product_name, start, end
            ):
                available.append(row)
            else:
                unavailable.append(row.product_name)
                logger.info(
                    "dropping unavailable candidate %r for flight %s–%s",
                    row.product_name,
                    start,
                    end,
                )
        return available, unavailable

    def _price_candidates(
        self,
        candidates: list[GtmProductCandidate],
        result: IdeationResult,
    ) -> list[PricedCandidate]:
        priced: list[PricedCandidate] = []
        for row in candidates:
            try:
                pricing_row = self._inventory.pricing.lookup(row.product_name)
                amount = self._inventory.pricing.primary_amount(row.product_name)
            except ValueError as exc:
                result.notes.append(str(exc))
                continue
            priced.append(
                PricedCandidate(
                    product=to_proposed_product(
                        row.product_name,
                        row.category,
                        amount,
                        pricing_row.pricing,
                    )
                )
            )
        return priced
