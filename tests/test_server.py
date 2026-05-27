from unittest.mock import MagicMock

from server import filter_decks_by_tags, get_slide_content, hello, search_decks


def test_hello():
    assert hello(name="world") == "Hello, world!"


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
