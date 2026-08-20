"""C2 deterministic fills (unwired from assemble_skeleton; called from build)."""

from __future__ import annotations

from datetime import date
from unittest.mock import MagicMock, patch

import pytest

from ingestion.placeholder_fills import (
    EM_DASH,
    PROGRAM_STOCK_BLURB,
    apply_placeholders,
    fetch_logo_bytes,
    format_presentation_date,
    format_usd,
    stated_total_budget,
)
from ingestion.schema import DeckSchema, Product
from tests.fortuneai_placeholder_fixture import (
    MINIMAL_PNG,
    SAMPLE_OPPORTUNITY_BODY,
    SAMPLE_PROGRAM_BLURB,
    WHY_FORTUNE_STOCK,
    build_fortuneai_fixture_prs,
    mock_placeholder_ai,
    sample_audience_data,
)


def _slide_blob(slide) -> str:
    parts: list[str] = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            parts.append(para.text)
    return "\n".join(parts)


def _schema(**overrides) -> DeckSchema:
    defaults = dict(
        company_name="Acme Corp",
        industry="Technology",
        budgets=[{"amount": 50_000}],
        flight_dates={"start": "2026-09-01", "end": "2026-12-31"},
        campaign_goal="Drive consideration among enterprise buyers",
        targeting_details=(
            "US enterprise tech decision-makers, Chief Executive Officer, "
            "C-suite, Chief Financial Officer"
        ),
        kpis=["Awareness", "Engagement"],
        kpi_details="Lift brand awareness 10%; engagement rate above benchmark",
        campaign_narrative="Acme helps mid-market CFOs modernize finance ops",
        preferred_platforms_products=["Newsletters", "Branded Content"],
        additional_rfp_details="Prefer Q4 flight; avoid holiday blackout weeks",
        client_logo="https://example.com/acme-logo.png",
        confirmed_products=[
            Product(
                name="CEO Daily",
                cadence="weekly",
                price=50_000,
                category="Newsletter",
            )
        ],
    )
    defaults.update(overrides)
    return DeckSchema(**defaults)


def _apply(prs=None, schema=None, **kwargs):
    return apply_placeholders(
        prs or build_fortuneai_fixture_prs(),
        schema or _schema(),
        audience=sample_audience_data(),
        logo_bytes=MINIMAL_PNG,
        as_of=date(2026, 8, 18),
        **kwargs,
    )


def test_format_presentation_date_month_year():
    assert format_presentation_date(date(2026, 8, 18)) == "August 2026"


def test_stated_total_budget_prefers_total_label():
    schema = _schema(
        budgets=[
            {"amount": 10_000, "label": "Digital"},
            {"amount": 50_000, "label": "Total budget"},
        ],
        confirmed_products=[
            Product(name="CEO Daily", cadence="weekly", price=50_000, category="Newsletter")
        ],
    )
    assert stated_total_budget(schema) == 50_000


def test_stated_total_budget_ignores_subtotal_label():
    schema = _schema(
        budgets=[
            {"amount": 10_000, "label": "Subtotal"},
            {"amount": 50_000, "label": "Newsletter"},
        ],
        confirmed_products=[
            Product(name="CEO Daily", cadence="weekly", price=50_000, category="Newsletter")
        ],
    )
    assert stated_total_budget(schema) == 50_000


def test_apply_without_ai_leaves_ai_tokens_and_why_fortune():
    prs = build_fortuneai_fixture_prs()
    warnings = _apply(prs)
    assert warnings == []
    intro = _slide_blob(prs.slides[0])
    assert "August 2026" in intro
    assert "[DATE]" not in intro
    assert "[TITLE]" in intro
    assert WHY_FORTUNE_STOCK in _slide_blob(prs.slides[1])
    history = _slide_blob(prs.slides[2])
    assert "Acme Corp" in history
    assert "[client name]" not in history
    opp = _slide_blob(prs.slides[3])
    assert "[HEADER]" in opp
    assert "[BODY]" in opp
    thanks = _slide_blob(prs.slides[-1])
    assert "August 2026" in thanks
    assert "Thank you!" in thanks
    assert "[DATE]" not in thanks


def test_apply_with_ai_fills_named_slots_and_leaves_why_fortune():
    prs = build_fortuneai_fixture_prs()
    ai = mock_placeholder_ai()
    _apply(prs, ai=ai)
    intro = _slide_blob(prs.slides[0])
    assert "ACME CORP ENTERPRISE PARTNERSHIP" in intro
    assert "[TITLE]" not in intro
    assert WHY_FORTUNE_STOCK in _slide_blob(prs.slides[1])
    opp = _slide_blob(prs.slides[3])
    assert "Lead With Confidence Today" in opp
    assert SAMPLE_OPPORTUNITY_BODY in opp
    assert "[HEADER]" not in opp
    assert "[BODY]" not in opp
    audience_pages = [s for s in prs.slides if "Reach enterprise leaders" in _slide_blob(s)]
    assert len(audience_pages) == 1
    assert "[AUDIENCE TITLE]" not in _slide_blob(audience_pages[0])
    program_pages = [s for s in prs.slides if "PROGRAM OVERVIEW" in _slide_blob(s)]
    assert SAMPLE_PROGRAM_BLURB in _slide_blob(program_pages[0])
    ai.intro_title.assert_called_once()
    ai.opportunity_body.assert_called_once()


def test_apply_keeps_only_3_card_audience_and_pulls_metrics_verbatim():
    prs = build_fortuneai_fixture_prs()
    _apply(prs, ai=mock_placeholder_ai())
    audience_pages = [
        s for s in prs.slides if "Chief Executive Officer" in _slide_blob(s)
    ]
    assert len(audience_pages) == 1
    blob = _slide_blob(audience_pages[0])
    assert "[AUDIENCE TITLE]" not in blob
    assert "Chief Executive Officer" in blob
    assert "C-suite" in blob
    assert "Chief Financial Officer" in blob
    assert "1.1M" in blob
    assert "3.6M" in blob
    assert "422K" in blob
    assert "154" in blob
    assert "[AUDIENCE SEGMENT]" not in blob
    assert "[REACH]" not in blob
    assert "[INDEX]" not in blob
    assert "Acme Corp" in blob


def test_apply_1_category_program_second_box_is_stock():
    prs = build_fortuneai_fixture_prs()
    _apply(prs, ai=mock_placeholder_ai())
    program_pages = [s for s in prs.slides if "PROGRAM OVERVIEW" in _slide_blob(s)]
    assert len(program_pages) == 1
    blob = _slide_blob(program_pages[0])
    assert "Editorial Alignment" in blob
    assert "Fortune" in blob
    assert blob.count("Product description.") == 0
    assert SAMPLE_PROGRAM_BLURB in blob
    assert PROGRAM_STOCK_BLURB in blob
    assert "PRODUCT TYPE" not in blob


def test_apply_investment_two_categories_divider_order():
    prs = build_fortuneai_fixture_prs()
    schema = _schema(
        budgets=[{"amount": 30_000}],
        targeting_details="Chief Executive Officer, C-suite",
        confirmed_products=[
            Product(name="Print Mag", cadence="annual", price=10_000, category="Print"),
            Product(
                name="Display", cadence="monthly", price=20_000, category="Digital Media"
            ),
        ],
    )
    _apply(prs, schema, ai=mock_placeholder_ai())
    blob = _slide_blob(prs.slides[-2])
    assert format_usd(30_000) in blob
    assert "[BUDGET]" not in blob
    assert blob.find("High-Impact Media") < blob.find("Print")
    assert f"Display {EM_DASH} $20,000" in blob
    assert f"Print Mag {EM_DASH} $10,000" in blob
    assert "[PRODUCT CATEGORY]" not in blob


def test_apply_budget_mismatch_raises():
    prs = build_fortuneai_fixture_prs()
    schema = _schema(budgets=[{"amount": 100_000}])
    with pytest.raises(ValueError, match="does not match mix total"):
        _apply(prs, schema, ai=mock_placeholder_ai())


def test_apply_lt_2_audience_raises():
    prs = build_fortuneai_fixture_prs()
    schema = _schema(targeting_details="US enterprise tech decision-makers")
    with pytest.raises(ValueError, match="at least 2"):
        _apply(prs, schema, ai=mock_placeholder_ai())


def test_apply_gt_6_audience_warns_and_keeps_6_card():
    prs = build_fortuneai_fixture_prs()
    schema = _schema(
        targeting_details=(
            "Chief Executive Officer, C-suite, Chief Financial Officer, "
            "Chief Information Officer, Chief Technology Officer, "
            "Chief Data Officer, Active Investor, Wealthy (HNW)"
        )
    )
    warnings = _apply(prs, schema, ai=mock_placeholder_ai())
    assert warnings
    assert "top 6" in warnings[0]
    audience_pages = [
        s for s in prs.slides if "Chief Technology Officer" in _slide_blob(s)
    ]
    assert len(audience_pages) == 1
    blob = _slide_blob(audience_pages[0])
    # Truncated by Index desc: CTO 255, CDO 219, CIO 207, C-suite 172, CEO 154, CFO 146
    assert "Chief Technology Officer" in blob
    assert "Active Investor" not in blob
    assert "[AUDIENCE SEGMENT]" not in blob


def test_fetch_logo_bytes_https(monkeypatch):
    resp = MagicMock()
    resp.content = MINIMAL_PNG
    resp.raise_for_status = MagicMock()
    with patch("ingestion.placeholder_fills.requests.get", return_value=resp) as mock_get:
        assert fetch_logo_bytes("https://example.com/logo.png") == MINIMAL_PNG
    mock_get.assert_called_once()


def test_fetch_logo_bytes_rejects_non_https():
    with pytest.raises(ValueError, match="HTTPS"):
        fetch_logo_bytes("http://example.com/logo.png")
    with pytest.raises(ValueError, match="HTTPS"):
        fetch_logo_bytes("/sites/foo/logo.png")
    with pytest.raises(ValueError, match="missing"):
        fetch_logo_bytes("  ")
