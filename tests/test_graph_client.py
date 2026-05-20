import os
import pytest
from unittest.mock import MagicMock
from ingestion.graph_client import GraphClient


# --- fixture ---
# We patch msal at construction time so no real Azure credentials are needed.
# We also patch _token() to return a dummy string so _get() doesn't try to auth.
# Every test gets a clean client with no real network access.
@pytest.fixture
def client(mocker, monkeypatch):
    monkeypatch.setenv("SHAREPOINT_SITE_ID", "test-site-id")
    mocker.patch("msal.ConfidentialClientApplication")
    c = GraphClient(tenant_id="t", client_id="c", client_secret="s")
    mocker.patch.object(c, "_token", return_value="dummy-token")
    return c


# --- _get: rate limit retry ---
# Graph API returns 429 when rate limited. _get must wait (Retry-After header)
# and retry rather than failing immediately. We patch time.sleep so the test
# doesn't actually wait, and mock the session to return 429 once then 200.
def test_get_retries_on_429(client, mocker):
    mocker.patch("time.sleep")

    mock_429 = MagicMock()
    mock_429.status_code = 429
    mock_429.headers = {"Retry-After": "0"}

    mock_200 = MagicMock()
    mock_200.status_code = 200
    mock_200.json.return_value = {"value": ["item"]}
    mock_200.raise_for_status = MagicMock()

    client._session.get = MagicMock(side_effect=[mock_429, mock_200])

    result = client._get("https://example.com")
    assert result == {"value": ["item"]}
    assert client._session.get.call_count == 2


# --- _get: exhausted retries ---
# If every attempt comes back as 429, _get should give up after max_retries
# and raise a RuntimeError rather than looping forever.
def test_get_raises_after_max_retries(client, mocker):
    mocker.patch("time.sleep")

    mock_429 = MagicMock()
    mock_429.status_code = 429
    mock_429.headers = {}

    client._session.get = MagicMock(return_value=mock_429)

    with pytest.raises(RuntimeError, match="failed after"):
        client._get("https://example.com", max_retries=3)

    assert client._session.get.call_count == 3


# --- _paginate: follows nextLink ---
# Graph API paginates at 100 items. Each page includes @odata.nextLink when
# there are more results. _paginate must follow that link and yield all items
# across all pages as a flat sequence.
def test_paginate_yields_all_items_across_pages(client, mocker):
    page1 = {"value": ["a", "b"], "@odata.nextLink": "https://example.com/page2"}
    page2 = {"value": ["c", "d"]}

    mocker.patch.object(client, "_get", side_effect=[page1, page2])

    results = list(client._paginate("https://example.com/page1"))
    assert results == ["a", "b", "c", "d"]


# --- list_decks: root-level files are skipped ---
# main_template.pptx lives at the library root and is not a closed-won deck.
# Any file at root level must be excluded — only GTM Current is the entry point.
def test_list_decks_skips_root_level_files(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [
                {"id": "f1", "name": "main_template.pptx", "file": {}},
            ]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert results == []


# --- list_decks: non-GTM folders are ignored ---
# Legacy folders (ACCOUNT MANAGEMENT, GTM Purge, etc.) at root are skipped.
# Only "GTM Current" is the authoritative source of truth for industry tagging.
def test_list_decks_ignores_non_gtm_current_folders(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [
                {"id": "folder-acct", "name": "ACCOUNT MANAGEMENT", "folder": {}},
                {"id": "folder-gtm", "name": "GTM Current (Q4 '24 - Forward)", "folder": {}},
            ]}
        if "folder-gtm" in url:
            return {"value": []}
        # ACCOUNT MANAGEMENT should never be walked
        if "folder-acct" in url:
            return {"value": [{"id": "d1", "name": "deck.pptx", "file": {}}]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert results == []


# --- list_decks: industry derived from GTM Current subfolder name ---
# A .pptx inside GTM Current/Technology/ must have _industry = "Technology".
def test_list_decks_industry_from_folder_name(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [{"id": "folder-gtm", "name": "GTM Current (Q4 '24 - Forward)", "folder": {}}]}
        if "folder-gtm" in url:
            return {"value": [{"id": "folder-tech", "name": "Technology", "folder": {}}]}
        if "folder-tech" in url:
            return {"value": [{"id": "d1", "name": "pitch.pptx", "file": {}, "webUrl": "/Technology/pitch.pptx"}]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert len(results) == 1
    assert results[0]["_industry"] == "Technology"
    assert results[0]["_sub_industry"] == ""


# --- list_decks: sub_industry from second-level folder inside GTM Current ---
# A .pptx inside GTM Current/Technology/Tech_Client Pitches/ must have
# _industry = "Technology" and _sub_industry = "Tech_Client Pitches".
def test_list_decks_sub_industry_from_nested_folder(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [{"id": "folder-gtm", "name": "GTM Current (Q4 '24 - Forward)", "folder": {}}]}
        if "folder-gtm" in url:
            return {"value": [{"id": "folder-tech", "name": "Technology", "folder": {}}]}
        if "folder-tech" in url:
            return {"value": [{"id": "folder-client", "name": "Tech_Client Pitches", "folder": {}}]}
        if "folder-client" in url:
            return {"value": [{"id": "d1", "name": "pitch.pptx", "file": {}, "webUrl": "/Technology/Tech_Client Pitches/pitch.pptx"}]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert len(results) == 1
    assert results[0]["_industry"] == "Technology"
    assert results[0]["_sub_industry"] == "Tech_Client Pitches"


# --- list_decks: does not recurse beyond two levels inside an industry folder ---
# The walk stops at sub-industry level. A third-level folder is ignored entirely.
def test_list_decks_does_not_recurse_beyond_two_levels(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [{"id": "folder-gtm", "name": "GTM Current (Q4 '24 - Forward)", "folder": {}}]}
        if "folder-gtm" in url:
            return {"value": [{"id": "folder-tech", "name": "Technology", "folder": {}}]}
        if "folder-tech" in url:
            return {"value": [{"id": "folder-client", "name": "Tech_Client Pitches", "folder": {}}]}
        if "folder-client" in url:
            return {"value": [{"id": "folder-deep", "name": "SubClient", "folder": {}}]}
        # this level should never be reached
        if "folder-deep" in url:
            return {"value": [{"id": "d1", "name": "deep.pptx", "file": {}}]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert results == []


# --- extract_tags: General Deck Shells flagged as template ---
# Slides in the "General Deck Shells" sub-folder are formatting templates,
# not content. deck_type="template" lets the retrieval layer filter them out.
def test_extract_tags_general_deck_shells_flagged_as_template(client):
    item = {"_industry": "(GP) General Presentation", "_sub_industry": "General Deck Shells"}
    tags = client.extract_tags(item)
    assert tags.deck_type == "template"


# --- extract_tags: non-template sub-industries have empty deck_type ---
def test_extract_tags_content_sub_industry_has_empty_deck_type(client):
    item = {"_industry": "Technology", "_sub_industry": "Tech_Client Pitches"}
    tags = client.extract_tags(item)
    assert tags.deck_type == ""


# --- list_decks: non-pptx files are ignored ---
# .pdf, .docx, and other file types inside industry folders are skipped.
def test_list_decks_ignores_non_pptx_files(client, mocker):
    def fake_get(url, params=None, max_retries=5):
        if "root/children" in url:
            return {"value": [{"id": "folder-gtm", "name": "GTM Current (Q4 '24 - Forward)", "folder": {}}]}
        if "folder-gtm" in url:
            return {"value": [{"id": "folder-tech", "name": "Technology", "folder": {}}]}
        if "folder-tech" in url:
            return {"value": [
                {"id": "d1", "name": "brief.pdf", "file": {}},
                {"id": "d2", "name": "notes.docx", "file": {}},
                {"id": "d3", "name": "pitch.pptx", "file": {}, "webUrl": "/Technology/pitch.pptx"},
            ]}
        return {"value": []}

    mocker.patch.object(client, "_get", side_effect=fake_get)

    results = client.list_decks()
    assert len(results) == 1
    assert results[0]["name"] == "pitch.pptx"
