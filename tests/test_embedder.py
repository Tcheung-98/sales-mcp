import json
from unittest.mock import MagicMock, patch

import numpy as np
import pytest
from botocore.exceptions import ClientError

from ingestion.embedder import _DIMENSIONS, _MAX_RETRIES, SlideEmbedder
from ingestion.models import SlideRow, Tags

# --- fixtures ---

def make_slide(
    deck_id="deck-1",
    slide_number=1,
    title="Title",
    body_text=None,
    deck_type="",
) -> SlideRow:
    return SlideRow(
        deck_id=deck_id,
        source_path="/FORTUNE/deck.pptx",
        content_hash="abc123",
        ingested_at="2026-05-21T00:00:00Z",
        slide_number=slide_number,
        body_text=body_text or ["Body line"],
        speaker_notes="",
        tags=Tags(industry="FORTUNE", deck_type=deck_type),
    )


def _fake_invoke_response(embedding: list[float]) -> dict:
    import io
    body_bytes = json.dumps({"embedding": embedding}).encode()
    return {"body": io.BytesIO(body_bytes)}


@pytest.fixture
def embedder(monkeypatch):
    monkeypatch.setenv("AWS_REGION", "us-east-1")
    monkeypatch.setenv("BEDROCK_EMBEDDING_MODEL_ID", "amazon.titan-embed-text-v2:0")
    with patch("boto3.client"):
        e = SlideEmbedder()
    e._client = MagicMock()
    e._client.invoke_model.side_effect = lambda **_: _fake_invoke_response([0.1] * _DIMENSIONS)
    return e


def _throttle_error() -> ClientError:
    return ClientError(
        {"Error": {"Code": "ThrottlingException", "Message": "Rate exceeded"}},
        "InvokeModel",
    )


# --- test_embed_slides_returns_correct_shape ---
# 3 non-template slides → shape (3, 1024)
def test_embed_slides_returns_correct_shape(embedder):
    rows = [make_slide(slide_number=i) for i in range(1, 4)]
    vectors, meta = embedder.embed_slides(rows)
    assert vectors.shape == (3, _DIMENSIONS)
    assert vectors.dtype == np.float32


# --- test_embed_slides_skips_template_slides ---
# Template slides must never be sent to Bedrock.
def test_embed_slides_skips_template_slides(embedder):
    rows = [
        make_slide(slide_number=1, deck_type="template"),
        make_slide(slide_number=2, deck_type=""),
    ]
    vectors, meta = embedder.embed_slides(rows)
    assert vectors.shape == (1, _DIMENSIONS)
    assert embedder._client.invoke_model.call_count == 1


# --- test_embed_slides_handles_empty_title ---
# None title should fall back to body_text; Bedrock must still be called.
def test_embed_slides_handles_empty_title(embedder):
    row = make_slide(title=None, body_text=["Just body text"])
    row = row.model_copy(update={"title": None})
    vectors, meta = embedder.embed_slides([row])
    assert vectors.shape == (1, _DIMENSIONS)
    call_body = json.loads(embedder._client.invoke_model.call_args.kwargs["body"])
    assert call_body["inputText"] == "Just body text"


# --- test_embed_slides_meta_aligns_with_vectors ---
# Meta list length must exactly equal the number of embedding rows.
def test_embed_slides_meta_aligns_with_vectors(embedder):
    rows = [make_slide(slide_number=i) for i in range(1, 6)]
    vectors, meta = embedder.embed_slides(rows)
    assert len(meta) == len(vectors)
    assert meta[0]["deck_id"] == "deck-1"
    assert meta[0]["slide_number"] == 1


# --- test_embed_slides_empty_input_returns_empty ---
def test_embed_slides_empty_input_returns_empty(embedder):
    vectors, meta = embedder.embed_slides([])
    assert vectors.shape == (0, _DIMENSIONS)
    assert meta == []
    embedder._client.invoke_model.assert_not_called()


# --- test_embed_text_retries_on_throttle ---
# First call raises ThrottlingException; second succeeds.
def test_embed_text_retries_on_throttle(embedder, mocker):
    sleep_mock = mocker.patch("time.sleep")
    mocker.patch("random.uniform", return_value=0.25)
    embedder._client.invoke_model.side_effect = [
        _throttle_error(),
        _fake_invoke_response([0.5] * _DIMENSIONS),
    ]
    result = embedder._embed_text("hello")
    assert len(result) == _DIMENSIONS
    assert embedder._client.invoke_model.call_count == 2
    sleep_mock.assert_called_once_with(1.25)


# --- test_embed_text_raises_after_max_retries ---
# Exhausting all retries must raise RuntimeError (not ClientError).
def test_embed_text_raises_after_max_retries(embedder, mocker):
    mocker.patch("time.sleep")
    embedder._client.invoke_model.side_effect = _throttle_error()
    with pytest.raises(RuntimeError, match="throttled"):
        embedder._embed_text("hello")
    assert embedder._client.invoke_model.call_count == _MAX_RETRIES + 1


# --- test_embed_text_raises_on_unexpected_dimension ---
# Wrong embedding size should fail fast before writing vectors.
def test_embed_text_raises_on_unexpected_dimension(embedder):
    embedder._client.invoke_model.side_effect = lambda **_: _fake_invoke_response([0.1, 0.2])
    with pytest.raises(RuntimeError, match="Unexpected embedding size"):
        embedder._embed_text("hello")


# --- test_embed_slides_filters_empty_text ---
# Slide with no title and empty body_text must be skipped entirely.
def test_embed_slides_filters_empty_text(embedder):
    empty_row = make_slide(title=None, body_text=[])
    empty_row = empty_row.model_copy(update={"title": None, "body_text": []})
    normal_row = make_slide(slide_number=2)
    vectors, meta = embedder.embed_slides([empty_row, normal_row])
    assert vectors.shape == (1, _DIMENSIONS)
    assert embedder._client.invoke_model.call_count == 1
