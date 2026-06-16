from unittest.mock import MagicMock

from server import build_deck, filter_decks_by_tags, get_slide_content, outline_deck, search_decks


def _valid_schema(**overrides) -> dict:
    base = {
        "client_name": "Acme Corp",
        "industry": "Tech",
        "budget_quarterly": 50_000.0,
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


def test_outline_deck_valid_returns_template_filename():
    result = outline_deck(schema=_valid_schema())
    assert result["status"] == "ok"
    assert "template_filename" in result
    assert result["template_filename"].endswith(".pptx.url")


def test_outline_deck_missing_fields_returns_incomplete():
    result = outline_deck(schema={"client_name": "Acme Corp"})
    assert result["status"] == "incomplete"
    assert "industry" in result["missing"]
    assert "budget_quarterly" in result["missing"]


def test_outline_deck_escalation_budget():
    result = outline_deck(schema=_valid_schema(budget_quarterly=750_000))
    assert result["status"] == "escalation"
    assert "GTM" in result["message"]


def test_build_deck_delegates_to_generator(mocker):
    fake_result = {
        "s3_uri": "s3://bucket/generated/test.pptx",
        "download_url": "https://example.com/test.pptx",
        "slide_count": 25,
        "client_name": "Acme Corp",
    }
    mock_generator = MagicMock()
    mock_generator.build.return_value = fake_result
    mocker.patch("server._get_generator", return_value=mock_generator)
    result = build_deck(
        schema=_valid_schema(),
        template_url="https://sharepoint.example.com/template.pptx",
    )
    mock_generator.build.assert_called_once()
    assert result == fake_result


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
