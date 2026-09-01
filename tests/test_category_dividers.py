"""Tests for FortuneAI category divider mapping."""

import pytest

from ingestion.category_dividers import (
    CATEGORY_DIVIDERS,
    FORTUNEAI_DIVIDER_SLIDE_INDEX,
    divider_index_for_category,
    is_fortuneai_template_url,
)


def test_divider_slide_index_matches_template_layout():
    """Physical divider positions in FortuneAI_DeckTemplate (not sequential 12–16)."""
    names = [CATEGORY_DIVIDERS[i].name for i in range(len(FORTUNEAI_DIVIDER_SLIDE_INDEX))]
    assert names == [
        "High-Impact Media",
        "Editorial Alignment",
        "Premium Video",
        "Print",
        "Branded Content",
    ]
    assert FORTUNEAI_DIVIDER_SLIDE_INDEX == (15, 13, 14, 16, 12)


def test_divider_order_matches_workflow():
    assert [d.name for d in CATEGORY_DIVIDERS] == [
        "High-Impact Media",
        "Editorial Alignment",
        "Premium Video",
        "Print",
        "Branded Content",
    ]


@pytest.mark.parametrize(
    ("category", "expected_name"),
    [
        ("Digital Media", "High-Impact Media"),
        ("Digital Ads/Programmatic", "High-Impact Media"),
        ("Newsletter", "Editorial Alignment"),
        ("Newsletters", "Editorial Alignment"),
        ("Vodcasts", "Premium Video"),
        ("Print", "Print"),
        ("Branded Content", "Branded Content"),
    ],
)
def test_divider_index_for_known_categories(category, expected_name):
    idx = divider_index_for_category(category)
    assert CATEGORY_DIVIDERS[idx].name == expected_name


def test_events_category_fails_loud():
    with pytest.raises(ValueError, match="escalate"):
        divider_index_for_category("Events")


def test_unknown_category_fails_loud():
    with pytest.raises(ValueError, match="no FortuneAI"):
        divider_index_for_category("Lists & Rankings Sponsorship")


def test_is_fortuneai_template_url():
    assert is_fortuneai_template_url(
        "https://fortune.sharepoint.com/sites/x/FortuneAI_DeckTemplate.pptx"
    )
    assert is_fortuneai_template_url(
        "https://fortune.sharepoint.com/FortuneAI_DeckTemplate.pptx.url?download=1"
    )
    assert not is_fortuneai_template_url(
        "https://fortune.sharepoint.com/Category_Presentation_Technology.pptx"
    )
