"""Deterministic trigger-phrase and GTM tag matching for Logic Guide."""

from __future__ import annotations

from ingestion.schema import DiscoverySchema

_MIN_TAG_LEN = 4


def seller_text(discovery: DiscoverySchema) -> str:
    """Combined seller prose used for tag / trigger matching."""
    parts = [
        discovery.targeting_details,
        discovery.campaign_narrative,
        discovery.campaign_goal,
        discovery.kpi_details,
        discovery.additional_rfp_details,
        discovery.platform_or_product_specifics or "",
    ]
    return " ".join(p for p in parts if p).casefold()


def text_contains_any(text: str, phrases: tuple[str, ...]) -> bool:
    folded = text.casefold()
    return any(phrase.casefold() in folded for phrase in phrases)


def genuine_tag_match(text: str, tag: str) -> bool:
    """Case-insensitive substring match with short-tag guard (Logic Guide genuine-match)."""
    needle = tag.strip().casefold()
    if not needle:
        return False
    if len(needle) < _MIN_TAG_LEN and " " not in needle:
        return False
    return needle in text.casefold()


def product_matches_tags(text: str, tags: tuple[str, ...]) -> bool:
    return any(genuine_tag_match(text, tag) for tag in tags)


_EXCLUDE_DIGITAL_ANCHORS = (
    "no display",
    "exclude high-impact",
    "skip crown/scroller",
    "skip crown",
    "skip scroller",
    "no digital ads",
)

_LIVESTREAM_PHRASES = (
    "livestream",
    "live stream",
    "live video",
    "streaming event",
    "watch live",
)

_VIDEO_CREATIVE_PHRASES = (
    "video pre-roll",
    "pre-roll",
    "use our video",
    "existing video asset",
    "video creative",
)

_SOCIAL_PHRASES = (
    "social",
    "social media",
    "social channels",
    "social extension",
    "paid social",
    "organic social",
    "instagram",
)

_KEY_DATE_PHRASES = (
    "launch date",
    "key date",
    "coincide with",
    "timed to",
    "anniversary",
    "product launch",
)

_BROAD_AWARENESS_PHRASES = (
    "broad reach",
    "wide awareness",
    "mass audience",
)

_NEWSLETTER_PHRASES = ("newsletter", "newsletters", "email")

_BC_VIDEO_PHRASES = (
    "video",
    "film",
    "footage",
    "documentary",
    "interview video",
    "b-roll",
)

_BC_WRITTEN_PHRASES = (
    "article",
    "written piece",
    "op-ed",
    "byline",
    "thought leadership piece",
)

_PROFILE_PHRASES = (
    "executive",
    "ceo",
    "cfo",
    "profile",
    "named leader",
    "interview with",
)

_MISSION_NARRATIVE_PHRASES = (
    "mission",
    "company story",
    "objective",
    "brand narrative",
)
