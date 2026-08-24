"""Per-category candidate gathering (Logic Guide Product Category Rules)."""

from __future__ import annotations

from ingestion.gtm_ideation_catalog import GtmIdeationCatalog, GtmProductCandidate
from ingestion.logic_guide.triggers import (
    _BC_VIDEO_PHRASES,
    _BC_WRITTEN_PHRASES,
    _BROAD_AWARENESS_PHRASES,
    _EXCLUDE_DIGITAL_ANCHORS,
    _KEY_DATE_PHRASES,
    _LIVESTREAM_PHRASES,
    _NEWSLETTER_PHRASES,
    _SOCIAL_PHRASES,
    _VIDEO_CREATIVE_PHRASES,
    product_matches_tags,
    seller_text,
    text_contains_any,
)
from ingestion.schema import DiscoverySchema

_DIGITAL_CATEGORY = "Digital Ads/Programmatic"
_NEWSLETTERS_CATEGORY = "Newsletters"
_VODCASTS_CATEGORY = "Vodcasts"
_BC_CATEGORY = "Branded Content"
_PRINT_CATEGORY = "Print"

# Backfill-only digital products — never category-rule candidates.
_DIGITAL_BACKFILL_ONLY = frozenset(
    {
        "Run of Fortune Display",
        "Contextually/Geotargeted Display",
        "Audience Targeted Display",
    }
)

_DIGITAL_CONDITIONAL_BY_NAME = {
    "In-Banner Streaming Video": _LIVESTREAM_PHRASES,
    "Run-of-Fortune Video Pre-Roll": _VIDEO_CREATIVE_PHRASES,
    "LinkedIn BrandLink": ("linkedin",),
    "Paid Social Video": _SOCIAL_PHRASES,
}


def _by_category(
    catalog: GtmIdeationCatalog, category: str
) -> list[GtmProductCandidate]:
    return catalog.products.products_in_category(category)


def _find_by_name(
    products: list[GtmProductCandidate], name: str
) -> GtmProductCandidate | None:
    for row in products:
        if row.product_name == name:
            return row
    return None


def gather_digital_candidates(
    discovery: DiscoverySchema, catalog: GtmIdeationCatalog
) -> list[GtmProductCandidate]:
    text = seller_text(discovery)
    pool = _by_category(catalog, _DIGITAL_CATEGORY)
    picked: list[GtmProductCandidate] = []

    if not text_contains_any(text, _EXCLUDE_DIGITAL_ANCHORS):
        for anchor in ("Crown Unit", "Scroller Unit"):
            row = _find_by_name(pool, anchor)
            if row:
                picked.append(row)

    for product_name, phrases in _DIGITAL_CONDITIONAL_BY_NAME.items():
        if product_name == "LinkedIn BrandLink" and "linkedin" not in text:
            continue
        if product_name == "Paid Social Video" and not text_contains_any(
            text, _SOCIAL_PHRASES
        ):
            continue
        if product_name not in ("LinkedIn BrandLink", "Paid Social Video"):
            if not text_contains_any(text, phrases):
                continue
        row = _find_by_name(pool, product_name)
        if row:
            picked.append(row)

    if text_contains_any(text, _KEY_DATE_PHRASES):
        for name in (
            "Crown Unit: Homepage Takeover + First impression Takeover",
            "Homepage + First Impression Takeover with Crown",
        ):
            row = _find_by_name(pool, name)
            if row:
                picked.append(row)
                break

    for row in pool:
        if row.product_name in _DIGITAL_BACKFILL_ONLY:
            continue
        if "Takeover" in row.product_name and product_matches_tags(text, row.gtm_tags):
            if row not in picked:
                picked.append(row)

    return _dedupe_candidates(picked)


def gather_newsletter_candidates(
    discovery: DiscoverySchema, catalog: GtmIdeationCatalog
) -> list[GtmProductCandidate]:
    text = seller_text(discovery)
    pool = _by_category(catalog, _NEWSLETTERS_CATEGORY)
    picked: list[GtmProductCandidate] = []

    for row in pool:
        if row.product_name == "Run of Fortune Newsletters":
            if text_contains_any(text, _BROAD_AWARENESS_PHRASES) and text_contains_any(
                text, _NEWSLETTER_PHRASES
            ):
                picked.append(row)
            continue
        if product_matches_tags(text, row.gtm_tags):
            picked.append(row)

    return _dedupe_candidates(picked)


def gather_vodcast_candidates(
    discovery: DiscoverySchema, catalog: GtmIdeationCatalog
) -> list[GtmProductCandidate]:
    text = seller_text(discovery)
    pool = _by_category(catalog, _VODCASTS_CATEGORY)
    return _dedupe_candidates(
        [row for row in pool if product_matches_tags(text, row.gtm_tags)]
    )


def gather_branded_content_candidates(
    discovery: DiscoverySchema, catalog: GtmIdeationCatalog
) -> list[GtmProductCandidate]:
    text = seller_text(discovery)
    pool = _by_category(catalog, _BC_CATEGORY)
    video_hit = text_contains_any(text, tuple(p for p in _BC_VIDEO_PHRASES))
    written_hit = text_contains_any(text, tuple(p for p in _BC_WRITTEN_PHRASES))

    if not video_hit and not written_hit:
        row = _find_by_name(pool, "Long-Form Article")
        return [row] if row else []

    picked: list[GtmProductCandidate] = []
    if video_hit:
        for name in (
            "Executive Q&A (Remote)",
            "Documentary-Style Video",
            "Hosted Interview Video (Remote)",
        ):
            row = _find_by_name(pool, name)
            if row:
                picked.append(row)
    if written_hit:
        for name in (
            "Long-Form Q&A Article",
            "Long-Form Article",
            "Enhanced Long-Form Article (2x)",
            "Enhanced Long-Form Article (4x)",
        ):
            row = _find_by_name(pool, name)
            if row and row not in picked:
                picked.append(row)

    if not picked:
        row = _find_by_name(pool, "Long-Form Article")
        if row:
            picked.append(row)
    return _dedupe_candidates(picked)


def gather_print_candidates(
    catalog: GtmIdeationCatalog,
) -> list[GtmProductCandidate]:
    pool = _by_category(catalog, _PRINT_CATEGORY)
    row = _find_by_name(pool, "Full Page")
    return [row] if row else []


def gather_candidates(
    discovery: DiscoverySchema,
    catalog: GtmIdeationCatalog,
    gtm_categories: list[str],
) -> list[GtmProductCandidate]:
    """Collect candidates across selected Logic Guide categories."""
    picked: list[GtmProductCandidate] = []
    if _DIGITAL_CATEGORY in gtm_categories:
        picked.extend(gather_digital_candidates(discovery, catalog))
    if _NEWSLETTERS_CATEGORY in gtm_categories:
        picked.extend(gather_newsletter_candidates(discovery, catalog))
    if _VODCASTS_CATEGORY in gtm_categories:
        picked.extend(gather_vodcast_candidates(discovery, catalog))
    if _BC_CATEGORY in gtm_categories:
        picked.extend(gather_branded_content_candidates(discovery, catalog))
    if _PRINT_CATEGORY in gtm_categories:
        picked.extend(gather_print_candidates(catalog))
    return _dedupe_candidates(picked)


def _dedupe_candidates(
    rows: list[GtmProductCandidate],
) -> list[GtmProductCandidate]:
    seen: set[tuple[str, str]] = set()
    out: list[GtmProductCandidate] = []
    for row in rows:
        key = (row.product_name, row.category)
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out
