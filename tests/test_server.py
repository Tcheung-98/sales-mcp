import asyncio
from unittest.mock import MagicMock

from server import (
    build_deck,
    confirm_mix,
    filter_decks_by_tags,
    get_slide_content,
    mcp,
    search_decks,
)


def _valid_schema(**overrides) -> dict:
    base = {
        "company_name": "Acme Corp",
        "industry": "Technology",
        "budgets": [{"amount": 50_000.0}],
        "flight_dates": {"start": "2026-09-01", "end": "2026-12-31"},
        "campaign_goal": "Drive consideration among enterprise buyers",
        "targeting_details": "US enterprise tech decision-makers",
        "kpis": ["Awareness", "Engagement"],
        "kpi_details": "Lift brand awareness 10%; engagement rate above benchmark",
        "campaign_narrative": "Acme helps mid-market CFOs modernize finance ops",
        "preferred_platforms_products": ["Newsletters", "Branded Content"],
        "additional_rfp_details": "Prefer Q4 flight; avoid holiday blackout weeks",
        "client_logo": "https://example.com/acme-logo.png",
        "confirmed_products": [
            {
                "name": "CIO Intelligence Newsletter",
                "cadence": "monthly",
                "price": 35_000.0,
                "category": "Newsletter",
            }
        ],
    }
    base.update(overrides)
    return base


def test_mcp_exposes_no_product_proposal_tool():
    names = {tool.name for tool in asyncio.run(mcp.list_tools())}
    assert "propose_mix" not in names
    assert {"confirm_mix", "build_deck"} <= names


def test_search_decks_delegates_to_retriever(mocker):
    fake_results = [{"deck_id": "d1", "slide_number": 1, "score": 0.9}]
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = fake_results
    mocker.patch("server._get_retriever", return_value=mock_retriever)
    results = search_decks(query="Fortune Amazon", k=3)
    mock_retriever.search.assert_called_once_with("Fortune Amazon", k=3)
    assert results == fake_results


def test_get_slide_content_delegates_to_retriever(mocker):
    fake_slides = [{
        "deck_id": "d1", "slide_number": 1, "title": "Intro",
        "body_text": ["Hello"], "layout_name": "Title", "source_path": "/GTM/d1.pptx",
    }]
    mock_retriever = MagicMock()
    mock_retriever.get_slide_content.return_value = fake_slides
    mocker.patch("server._get_retriever", return_value=mock_retriever)
    results = get_slide_content(deck_id="d1", slide_numbers=[1])
    mock_retriever.get_slide_content.assert_called_once_with("d1", [1])
    assert results == fake_slides


def test_get_slide_content_no_slide_numbers(mocker):
    mock_retriever = MagicMock()
    mock_retriever.get_slide_content.return_value = []
    mocker.patch("server._get_retriever", return_value=mock_retriever)
    get_slide_content(deck_id="d1")
    mock_retriever.get_slide_content.assert_called_once_with("d1", None)


def test_build_deck_delegates_to_generator(mocker):
    fake_result = {
        "download_url": "https://example.com/test.pptx",
        "slide_count": 25,
        "client_name": "Acme Corp",
        "template_key": "FortuneAI_DeckTemplate.pptx",
    }
    mock_generator = MagicMock()
    mock_generator.build.return_value = fake_result
    mocker.patch("server._get_generator", return_value=mock_generator)
    result = build_deck(
        schema=_valid_schema(),
        template_url="https://fortune.sharepoint.com/FortuneAI_DeckTemplate.pptx",
    )
    mock_generator.build.assert_called_once()
    assert result == fake_result


def test_build_deck_allows_omitted_template_url(mocker):
    fake_result = {
        "download_url": "https://example.com/test.pptx",
        "slide_count": 16,
        "client_name": "Acme Corp",
        "template_key": "FortuneAI_DeckTemplate.pptx",
    }
    mock_generator = MagicMock()
    mock_generator.build.return_value = fake_result
    mocker.patch("server._get_generator", return_value=mock_generator)
    result = build_deck(schema=_valid_schema())
    mock_generator.build.assert_called_once()
    assert mock_generator.build.call_args.args[1] is None
    assert result == fake_result


def test_build_deck_escalation_budget():
    result = build_deck(
        schema=_valid_schema(budgets=[{"amount": 750_000}]),
        template_url="https://fortune.sharepoint.com/FortuneAI_DeckTemplate.pptx",
    )
    assert result["status"] == "escalation"
    assert "GTM" in result["message"]


def test_build_deck_incomplete_schema():
    result = build_deck(schema={"company_name": "Acme Corp"})
    assert result["status"] == "incomplete"
    assert "industry" in result["missing"]
    assert "budgets" in result["missing"]


def test_build_deck_assembly_error(mocker):
    mock_generator = MagicMock()
    mock_generator.build.side_effect = ValueError("host not allowed")
    mocker.patch("server._get_generator", return_value=mock_generator)
    result = build_deck(
        schema=_valid_schema(),
        template_url="https://evil.example.com/FortuneAI_DeckTemplate.pptx",
    )
    assert result["status"] == "error"
    assert "host not allowed" in result["message"]


def test_filter_decks_by_tags_delegates_to_retriever(mocker):
    fake_results = [{
        "deck_id": "tech1", "title": "Tech Pitch", "tags": {},
        "slide_count": 2, "source_path": "/Technology/tech1.pptx",
    }]
    mock_retriever = MagicMock()
    mock_retriever.filter_decks_by_tags.return_value = fake_results
    mocker.patch("server._get_retriever", return_value=mock_retriever)
    results = filter_decks_by_tags(industry="Technology", deck_type="Pitch", limit=5)
    mock_retriever.filter_decks_by_tags.assert_called_once_with(
        industry="Technology", sub_industry=None, product_line=None, deal_size=None,
        client_name=None, date_from=None, date_to=None, deck_type="Pitch", limit=5,
    )
    assert results == fake_results


def test_confirm_mix_delegates_flat_selection(mocker):
    fake = {
        "status": "ok",
        "deck_schema": _valid_schema(),
        "warnings": [],
        "confirmed_products": _valid_schema()["confirmed_products"],
        "mix_total": 35_000,
    }
    mocker.patch("server._get_product_catalogs", return_value=(MagicMock(), MagicMock()))
    wrapper = mocker.patch("server.confirm_mix_from_dict", return_value=fake)
    result = confirm_mix(
        discovery=_valid_schema(),
        selected_products=[
            {"name": "CIO Intelligence Newsletter", "category": "Newsletters"}
        ],
    )
    assert result["status"] == "ok"
    assert result["deck_schema"]["company_name"] == "Acme Corp"
    payload = wrapper.call_args.args[0]
    assert set(payload) == {"discovery", "selected_products"}
    assert payload["selected_products"][0]["name"] == "CIO Intelligence Newsletter"
