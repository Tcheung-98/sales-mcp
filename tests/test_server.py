from server import hello, search_historical_decks


def test_hello():
    assert hello(name="world") == "Hello, world!"


def test_search_historical_decks_match():
    results = search_historical_decks(industry="tech")
    assert len(results) == 1
    assert results[0]["id"] == "deck1"


def test_search_historical_decks_no_match():
    assert search_historical_decks(industry="finance") == []
