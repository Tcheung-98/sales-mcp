"""Workflow platform strings → Logic Guide / schema categories."""

from __future__ import annotations

# Platforms that must escalate to GTM — never auto-pitch (Logic Guide + Workflow).
GTM_ESCALATION_PLATFORMS: frozenset[str] = frozenset(
    {
        "Conference Sponsorship/Media",
        "Lists & Rankings Sponsorship",
    }
)

# Discovery preferred_platforms_products → GTM Product Tags category column.
PLATFORM_TO_GTM_CATEGORY: dict[str, str] = {
    "Digital Ads/Programmatic": "Digital Ads/Programmatic",
    "Newsletters": "Newsletters",
    "Vodcasts": "Vodcasts",
    "Branded Content": "Branded Content",
    "Print": "Print",
}

# GTM category → DeckSchema Product.category enum value.
GTM_TO_SCHEMA_CATEGORY: dict[str, str] = {
    "Digital Ads/Programmatic": "Digital Media",
    "Newsletters": "Newsletter",
    "Vodcasts": "Vodcasts",
    "Branded Content": "Branded Content",
    "Print": "Print",
}


def selected_gtm_categories(preferred_platforms: list[str]) -> list[str]:
    """Map seller-selected platforms to GTM categories (non-escalation only)."""
    categories: list[str] = []
    for platform in preferred_platforms:
        if platform in GTM_ESCALATION_PLATFORMS:
            continue
        gtm = PLATFORM_TO_GTM_CATEGORY.get(platform)
        if gtm and gtm not in categories:
            categories.append(gtm)
    return categories


def escalation_reasons(preferred_platforms: list[str]) -> list[str]:
    reasons: list[str] = []
    for platform in preferred_platforms:
        if platform in GTM_ESCALATION_PLATFORMS:
            reasons.append(
                f"Preferred platform {platform!r} requires GTM escalation — "
                "do not auto-generate a pitch mix."
            )
    return reasons
