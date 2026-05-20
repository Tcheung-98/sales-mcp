import re
import pytest
from unittest.mock import MagicMock, patch

from ingestion.ingestor import run_ingest
from ingestion.models import SlideRow, Tags


SAMPLE_TAGS = Tags(industry="TechMedia", sub_industry="B2B")

DECK_ITEM_1 = {"id": "deck-1", "name": "pitch.pptx", "webUrl": "/TechMedia/pitch.pptx", "_industry": "TechMedia", "_sub_industry": "B2B"}
DECK_ITEM_2 = {"id": "deck-2", "name": "brief.pptx", "webUrl": "/TechMedia/brief.pptx", "_industry": "TechMedia", "_sub_industry": "B2B"}


def make_slide(deck_id="deck-1", slide_number=1) -> SlideRow:
    return SlideRow(
        deck_id=deck_id,
        source_path="/TechMedia/pitch.pptx",
        content_hash="abc123",
        ingested_at="2026-05-20T00:00:00Z",
        slide_number=slide_number,
        layout_name="Title Slide",
        title="Test Slide",
        body_text=["Body text"],
        speaker_notes="",
        tags=SAMPLE_TAGS,
    )


@pytest.fixture
def mock_client():
    c = MagicMock()
    c.list_decks.return_value = [DECK_ITEM_1, DECK_ITEM_2]
    c.download_deck.return_value = b"fake-pptx-bytes"
    c.extract_tags.return_value = SAMPLE_TAGS
    return c


@pytest.fixture
def mock_writer():
    w = MagicMock()
    w.write_decks.return_value = "s3://bucket/snapshots/ts/decks.parquet"
    w.write_failed.return_value = None
    return w


# --- happy path: slide rows from all decks are collected and returned ---
def test_happy_path_returns_all_slide_rows(mock_client, mock_writer):
    slides_1 = [make_slide("deck-1", 1), make_slide("deck-1", 2)]
    slides_2 = [make_slide("deck-2", 1)]

    with patch("ingestion.ingestor.parse_pptx", side_effect=[slides_1, slides_2]):
        slides, failed = run_ingest(client=mock_client, writer=mock_writer)

    assert len(slides) == 3
    assert len(failed) == 0


# --- failed deck: FailedRecord has correct deck_id and error message ---
def test_failed_deck_adds_failed_record(mock_client, mock_writer):
    mock_client.download_deck.side_effect = [Exception("network error"), b"ok-bytes"]

    with patch("ingestion.ingestor.parse_pptx", return_value=[make_slide("deck-2", 1)]):
        _, failed = run_ingest(client=mock_client, writer=mock_writer)

    assert len(failed) == 1
    assert failed[0].deck_id == "deck-1"
    assert "network error" in failed[0].error


# --- failed deck does not abort the run: other decks still processed ---
def test_failed_deck_does_not_crash_run(mock_client, mock_writer):
    mock_client.download_deck.side_effect = [RuntimeError("boom"), b"ok-bytes"]

    with patch("ingestion.ingestor.parse_pptx", return_value=[make_slide("deck-2", 1)]):
        slides, failed = run_ingest(client=mock_client, writer=mock_writer)

    assert len(slides) == 1
    assert len(failed) == 1


# --- run_ts matches ISO 8601 UTC format (YYYY-MM-DDTHH:MM:SSZ) ---
def test_run_ts_format(mock_client, mock_writer):
    mock_client.list_decks.return_value = []

    run_ingest(client=mock_client, writer=mock_writer)

    _, kwargs = mock_writer.write_decks.call_args
    assert re.match(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$", kwargs["run_ts"])


# --- source_path falls back to item["name"] when webUrl is absent ---
def test_source_path_falls_back_to_name(mock_client, mock_writer):
    mock_client.list_decks.return_value = [
        {"id": "deck-1", "name": "fallback.pptx", "_industry": "TechMedia", "_sub_industry": ""}
    ]
    mock_client.download_deck.side_effect = Exception("forced")

    _, failed = run_ingest(client=mock_client, writer=mock_writer)

    assert len(failed) == 1
    assert failed[0].source_path == "fallback.pptx"


# --- write_decks is called once with all collected slide rows ---
def test_write_decks_called_with_all_rows(mock_client, mock_writer):
    mock_client.list_decks.return_value = [DECK_ITEM_1]
    rows = [make_slide("deck-1", 1), make_slide("deck-1", 2)]

    with patch("ingestion.ingestor.parse_pptx", return_value=rows):
        run_ingest(client=mock_client, writer=mock_writer)

    mock_writer.write_decks.assert_called_once()
    positional_rows = mock_writer.write_decks.call_args[0][0]
    assert len(positional_rows) == 2


# --- write_failed is always called, even when no decks failed ---
def test_write_failed_always_called(mock_client, mock_writer):
    mock_client.list_decks.return_value = []

    run_ingest(client=mock_client, writer=mock_writer)

    mock_writer.write_failed.assert_called_once()


# --- parse error (corrupt pptx) is captured as a FailedRecord, not a crash ---
def test_parse_error_captured_as_failed_record(mock_client, mock_writer):
    mock_client.list_decks.return_value = [DECK_ITEM_1]

    with patch("ingestion.ingestor.parse_pptx", side_effect=ValueError("corrupt pptx")):
        slides, failed = run_ingest(client=mock_client, writer=mock_writer)

    assert len(slides) == 0
    assert len(failed) == 1
    assert failed[0].deck_id == "deck-1"
    assert "corrupt pptx" in failed[0].error
