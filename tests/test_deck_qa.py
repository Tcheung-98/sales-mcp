"""B3 deterministic deck QA (docs/DECK-QA-ARCHITECTURE.md §6).

Offline: the CI FortuneAI fixture driven through C2 fills. No S3, no Cursor,
no LibreOffice. The fixture keeps all five stock dividers and has no A5 clones,
so it is 13 slides rather than 8 + P — the manifest is built from the deck.
"""

from __future__ import annotations

from datetime import date

import pytest

from ingestion.category_dividers import FORTUNEAI_TEMPLATE_BASENAME
from ingestion.deck_qa import (
    PITCH_START_INDEX,
    DeckQaError,
    QaReport,
    run_deterministic_qa,
)
from ingestion.manifest import ReviewManifest, SlideManifestEntry
from ingestion.placeholder_fills import TITLE_TOKEN, apply_placeholders
from ingestion.pptx_tools import delete_slide, replace_token
from ingestion.schema import DeckSchema, Product
from tests.fortuneai_placeholder_fixture import (
    MINIMAL_PNG,
    build_fortuneai_fixture_prs,
    mock_placeholder_ai,
    sample_audience_data,
)

_INTRO_TITLE = "ACME CORP ENTERPRISE PARTNERSHIP"
_EXPECTED_CHECKS = [
    "slide_count",
    "leftover_tokens",
    "investment_budget",
    "product_presence",
    "divider_order",
    "tail_slides",
    "pptx_integrity",
    "manifest_consistency",
]


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
            "Chief Executive Officer, C-suite, Chief Financial Officer"
        ),
        kpis=["Awareness", "Engagement"],
        kpi_details="Lift brand awareness 10%; engagement rate above benchmark",
        campaign_narrative="Acme helps mid-market CFOs modernize finance ops",
        preferred_platforms_products=["Newsletters"],
        additional_rfp_details="Prefer Q4 flight",
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


def _filled_prs(schema: DeckSchema):
    prs = build_fortuneai_fixture_prs()
    apply_placeholders(
        prs,
        schema,
        audience=sample_audience_data(),
        logo_bytes=MINIMAL_PNG,
        as_of=date(2026, 8, 18),
        ai=mock_placeholder_ai(),
    )
    return prs


def _manifest(prs, schema: DeckSchema) -> ReviewManifest:
    """§5 role map: 0 cover, 1–5 narrative, everything after that `other`."""
    slides = [
        SlideManifestEntry(
            slide_index=index,
            role="cover" if index == 0 else "narrative" if index <= 5 else "other",
        )
        for index in range(len(prs.slides))
    ]
    return ReviewManifest(
        client_name=schema.company_name,
        template_key=FORTUNEAI_TEMPLATE_BASENAME,
        slide_count=len(prs.slides),
        slides=slides,
    )


def _failed_names(report: QaReport) -> list[str]:
    return [check.name for check in report.failures]


def test_deterministic_qa_passes_on_filled_fixture():
    schema = _schema()
    prs = _filled_prs(schema)
    report = run_deterministic_qa(prs, schema, _manifest(prs, schema))
    assert _failed_names(report) == []
    assert report.passed
    assert [check.name for check in report.checks] == _EXPECTED_CHECKS
    assert "passed" in report.summary()


def test_missing_funded_divider_fails():
    schema = _schema()
    prs = _filled_prs(schema)
    editorial = next(
        index
        for index, slide in enumerate(prs.slides)
        if index >= PITCH_START_INDEX and "Editorial Alignment" in _slide_blob(slide)
    )
    delete_slide(prs, editorial)
    report = run_deterministic_qa(prs, schema, _manifest(prs, schema))
    assert _failed_names(report) == ["divider_order"]
    assert "Editorial Alignment" in report.failures[0].message


def test_leftover_title_token_fails():
    schema = _schema()
    prs = _filled_prs(schema)
    assert replace_token(prs.slides[0], _INTRO_TITLE, TITLE_TOKEN) == 1
    report = run_deterministic_qa(prs, schema, _manifest(prs, schema))
    assert _failed_names(report) == ["leftover_tokens"]
    assert not report.passed
    message = report.failures[0].message
    assert TITLE_TOKEN in message
    assert "slide 0" in message


def test_budget_mismatch_fails():
    filled = _schema()
    prs = _filled_prs(filled)
    # C2 raises on a mismatch, so the deck is filled from a matching schema and
    # QA is handed the mismatched one — the standalone --package-dir case.
    mismatched = _schema(budgets=[{"amount": 100_000}])
    report = run_deterministic_qa(prs, mismatched, _manifest(prs, mismatched))
    assert _failed_names(report) == ["investment_budget"]
    assert "$100,000" in report.failures[0].message
    assert "$50,000" in report.failures[0].message


def test_slide_count_mismatch_fails():
    schema = _schema()
    prs = _filled_prs(schema)
    manifest = _manifest(prs, schema).model_copy(update={"slide_count": 10})
    report = run_deterministic_qa(prs, schema, manifest)
    assert _failed_names(report) == ["slide_count"]


def test_manifest_index_out_of_range_fails():
    schema = _schema()
    prs = _filled_prs(schema)
    manifest = _manifest(prs, schema)
    manifest.slides.append(SlideManifestEntry(slide_index=len(prs.slides), role="other"))
    report = run_deterministic_qa(prs, schema, manifest)
    assert _failed_names(report) == ["manifest_consistency"]


def test_deck_qa_error_carries_report_and_is_value_error():
    schema = _schema()
    prs = _filled_prs(schema)
    report = run_deterministic_qa(prs, schema, _manifest(prs, schema))
    with pytest.raises(ValueError) as excinfo:
        raise DeckQaError(report.summary(), report=report)
    assert isinstance(excinfo.value, DeckQaError)
    assert excinfo.value.report is report
