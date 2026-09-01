"""FortuneAI_DeckTemplate product-pitch category dividers (Workflow slides 13–17).

Dividers are included only when ≥1 funded product maps into that bucket.
Fixed order matches Pitch Deck Workflow / END-SCOPE-SOT.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from ingestion.gtm_product_map import normalize_category

# Narrative spine: slides 1–12 (indices 0–11). Product dividers live at the
# physical indices below (not sequential — matches FortuneAI_DeckTemplate on S3).
# Pitch output order is still Workflow 13→17 via CATEGORY_DIVIDERS.
FORTUNEAI_DIVIDER_COUNT = 5
FORTUNEAI_INVESTMENT_SLIDE_INDEX = 17
FORTUNEAI_THANKS_SLIDE_INDEX = 18
FORTUNEAI_MIN_SLIDES = FORTUNEAI_THANKS_SLIDE_INDEX + 1  # through thank you

# category index 0..4 → physical divider slide index in FortuneAI_DeckTemplate
FORTUNEAI_DIVIDER_SLIDE_INDEX: tuple[int, ...] = (
    15,  # High-Impact Media
    13,  # Editorial Alignment
    14,  # Premium Video
    16,  # Print
    12,  # Branded Content
)

# Legacy alias — first physical divider index in the template file (not pitch order).
FORTUNEAI_FIRST_DIVIDER_INDEX = min(FORTUNEAI_DIVIDER_SLIDE_INDEX)

FORTUNEAI_TEMPLATE_BASENAME = "FortuneAI_DeckTemplate.pptx"
_DEFAULT_FORTUNEAI_TEMPLATE_KEY = "templates/FortuneAI_DeckTemplate.pptx"


@dataclass(frozen=True)
class CategoryDivider:
    """One Workflow product-pitch section divider."""

    name: str
    # GTM Product Tags / preferred-platform category strings after normalize_category.
    gtm_categories: frozenset[str]


# Fixed order 13→17.
CATEGORY_DIVIDERS: tuple[CategoryDivider, ...] = (
    CategoryDivider(
        "High-Impact Media",
        frozenset({"Digital Ads/Programmatic"}),
    ),
    CategoryDivider(
        "Editorial Alignment",
        frozenset({"Newsletters"}),
    ),
    CategoryDivider(
        "Premium Video",
        frozenset({"Vodcasts"}),
    ),
    CategoryDivider(
        "Print",
        frozenset({"Print"}),
    ),
    CategoryDivider(
        "Branded Content",
        frozenset({"Branded Content"}),
    ),
)

# Categories that must never land in Creation product pitch (Workflow escalate).
_CREATION_BLOCKED_GTM_CATEGORIES = frozenset({"Events"})


def divider_index_for_category(category: str) -> int:
    """Return 0-based index into CATEGORY_DIVIDERS for a schema/GTM category.

    Raises ValueError when the category has no FortuneAI divider (e.g. Events)
    or is unknown.
    """
    gtm = normalize_category(category)
    if gtm in _CREATION_BLOCKED_GTM_CATEGORIES:
        raise ValueError(
            f"Product category {category!r} (GTM {gtm!r}) cannot be placed in "
            "FortuneAI Creation — escalate Conference/Events to GTM"
        )
    for i, divider in enumerate(CATEGORY_DIVIDERS):
        if gtm in divider.gtm_categories:
            return i
    raise ValueError(
        f"Product category {category!r} (GTM {gtm!r}) has no FortuneAI "
        "category divider mapping"
    )


def fortuneai_template_key() -> str:
    """S3 object key for the FortuneAI spine template."""
    return os.environ.get(
        "FORTUNEAI_TEMPLATE_KEY", _DEFAULT_FORTUNEAI_TEMPLATE_KEY
    )


def is_fortuneai_template_url(template_url: str) -> bool:
    """True when a SharePoint/download URL names FortuneAI_DeckTemplate."""
    path = template_url.split("?", 1)[0].rstrip("/")
    name = path.rsplit("/", 1)[-1]
    # SharePoint sometimes serves "FortuneAI_DeckTemplate.pptx.url" link files.
    return "fortuneai_decktemplate" in name.lower()
