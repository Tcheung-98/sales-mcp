"""Bounded Claude placeholder fills (C2 Chunk 5)."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from ingestion.placeholder_ai import (
    BODY_WORD_MAX,
    BODY_WORD_MIN,
    PROGRAM_BLURB_WORD_MAX,
    PROGRAM_BLURB_WORD_MIN,
    TITLE_WORD_MAX,
    TITLE_WORD_MIN,
    PlaceholderAI,
    validate_opportunity_body,
    validate_program_blurb,
    validate_short_title,
)
from ingestion.schema import DeckSchema, Product
from tests.fortuneai_placeholder_fixture import SAMPLE_OPPORTUNITY_BODY


def _schema(**overrides) -> DeckSchema:
    defaults = dict(
        company_name="Acme Corp",
        industry="Technology",
        budgets=[{"amount": 50_000}],
        flight_dates={"start": "2026-09-01", "end": "2026-12-31"},
        campaign_goal="Drive consideration among enterprise buyers",
        targeting_details="Chief Executive Officer, C-suite",
        kpis=["Awareness", "Engagement"],
        kpi_details="Lift brand awareness 10%; engagement rate above benchmark",
        campaign_narrative="Acme helps mid-market CFOs modernize finance ops",
        preferred_platforms_products=["Newsletters", "Branded Content"],
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


def _valid_body() -> str:
    return SAMPLE_OPPORTUNITY_BODY


def _recording_caller(responses: list[str]):
    calls: list[tuple[str, str]] = []

    def caller(*, system: str, user: str, max_tokens: int = 512) -> str:
        calls.append((system, user))
        if not responses:
            raise RuntimeError("no canned responses left")
        return responses.pop(0)

    return caller, calls


def test_validate_short_title_all_caps():
    validate_short_title("ACME CORP PARTNERSHIP", all_caps=True, slot="Intro title")


def test_validate_short_title_rejects_em_dash():
    with pytest.raises(ValueError, match="em dash"):
        validate_short_title("Lead — Now", slot="Header")


def test_validate_opportunity_body_requires_fortune_closer():
    words = SAMPLE_OPPORTUNITY_BODY.split()
    no_fortune_closer = " ".join(words[:-14]) + " Acme can win with credibility now."
    with pytest.raises(ValueError, match="Fortune"):
        validate_opportunity_body(no_fortune_closer)


def test_validate_program_blurb_word_count():
    with pytest.raises(ValueError, match=f"{PROGRAM_BLURB_WORD_MIN}"):
        validate_program_blurb("Too short blurb here.")


def test_intro_title_prompt_includes_all_caps_and_word_count():
    caller, calls = _recording_caller(["ACME CORP ENTERPRISE DEAL"])
    ai = PlaceholderAI(caller)
    title = ai.intro_title(_schema())
    assert title == "ACME CORP ENTERPRISE DEAL"
    system, user = calls[0]
    assert "ALL CAPS" in system
    assert f"{TITLE_WORD_MIN}" in system and f"{TITLE_WORD_MAX}" in system
    assert "no em dash" in system.lower() or "em dash" in system.lower()
    assert "Acme Corp" in user


def test_opportunity_body_prompt_includes_fortune_closer_and_word_count():
    caller, calls = _recording_caller([_valid_body()])
    ai = PlaceholderAI(caller)
    body = ai.opportunity_body(_schema())
    assert body.endswith("business.")
    system, user = calls[0]
    assert "Fortune" in system
    assert f"{BODY_WORD_MIN}" in system and f"{BODY_WORD_MAX}" in system
    assert "em dash" in system.lower()
    assert "starts with Fortune" in system or "STARTS with the word Fortune" in system
    assert "RFP" in user or "RFP" in system


def test_audience_title_prompt_includes_sentence_case():
    caller, calls = _recording_caller(["Reach enterprise decision-makers"])
    ai = PlaceholderAI(caller)
    title = ai.audience_title(_schema(), ["Chief Executive Officer", "C-suite"])
    assert title == "Reach enterprise decision-makers"
    system, _ = calls[0]
    assert "sentence case" in system.lower()


def test_program_blurb_prompt_includes_word_count():
    blurb = (
        "Fortune newsletters connect Acme with decision-makers through trusted "
        "weekly editorial environments and strategic insight daily."
    )
    caller, calls = _recording_caller([blurb])
    ai = PlaceholderAI(caller)
    result = ai.program_blurb(_schema(), "Editorial Alignment")
    assert result == blurb
    system, user = calls[0]
    assert f"{PROGRAM_BLURB_WORD_MIN}" in system and f"{PROGRAM_BLURB_WORD_MAX}" in system
    assert "Editorial Alignment" in user


def test_call_retries_once_then_fails():
    bad = "BAD"
    caller, calls = _recording_caller([bad, bad])
    ai = PlaceholderAI(caller)
    with pytest.raises(ValueError, match="failed after retry"):
        ai.intro_title(_schema())
    assert len(calls) == 2
    assert "failed validation" in calls[1][1]


def test_call_succeeds_on_second_attempt():
    caller, calls = _recording_caller(["bad", "ACME CORP ENTERPRISE WIN"])
    ai = PlaceholderAI(caller)
    assert ai.intro_title(_schema()) == "ACME CORP ENTERPRISE WIN"
    assert len(calls) == 2


def test_from_anthropic_wraps_client():
    mock_client = MagicMock()
    mock_block = MagicMock()
    mock_block.text = "SHAPE THE FUTURE NOW"
    mock_client.messages.create.return_value.content = [mock_block]
    ai = PlaceholderAI.from_anthropic(mock_client, model="claude-test")
    title = ai.opportunity_header(_schema())
    assert title == "SHAPE THE FUTURE NOW"
    mock_client.messages.create.assert_called_once()
