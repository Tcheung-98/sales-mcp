import io
import json
from unittest.mock import MagicMock

import numpy as np
import pyarrow as pa
import pyarrow.parquet as pq
import pytest

from ingestion.retriever import SlideRetriever


RUN_TS = "2026-05-27T10:00:00Z"

# Three orthogonal unit vectors — dot product with a query vec gives a clean,
# predictable score per slide without floating-point ambiguity.
_VECTORS = np.eye(3, 1024, dtype=np.float32)

_META = [
    {"deck_id": "d1", "slide_number": 1, "source_path": "/GTM/d1.pptx"},
    {"deck_id": "d2", "slide_number": 1, "source_path": "/GTM/d2.pptx"},
    {"deck_id": "d3", "slide_number": 1, "source_path": "/GTM/d3.pptx"},
]

_DECKS = [
    {
        "deck_id": "d1", "slide_number": 1, "source_path": "/GTM/d1.pptx",
        "content_hash": "a", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Deck One", "body_text": json.dumps(["Alpha", "Bravo"]),
        "tags": "{}", "tag_sources": None,
    },
    {
        "deck_id": "d2", "slide_number": 1, "source_path": "/GTM/d2.pptx",
        "content_hash": "b", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Deck Two", "body_text": json.dumps(["Charlie"]),
        "tags": "{}", "tag_sources": None,
    },
    {
        "deck_id": "d3", "slide_number": 1, "source_path": "/GTM/d3.pptx",
        "content_hash": "c", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Deck Three", "body_text": json.dumps(["Delta"]),
        "tags": "{}", "tag_sources": None,
    },
]


def _s3_mock(vectors=_VECTORS, meta=_META, decks=_DECKS, run_ts=RUN_TS):
    n = len(vectors)
    npy_buf = io.BytesIO()
    np.save(npy_buf, vectors)

    meta_buf = io.BytesIO()
    pq.write_table(pa.Table.from_pylist(meta[:n]), meta_buf)

    decks_buf = io.BytesIO()
    pq.write_table(pa.Table.from_pylist(decks[:n]), decks_buf)

    manifest = json.dumps({"run_ts": run_ts}).encode()

    def get_object(Bucket, Key):
        if "latest.json" in Key:
            return {"Body": io.BytesIO(manifest)}
        if Key.endswith(".npy"):
            return {"Body": io.BytesIO(npy_buf.getvalue())}
        if "embeddings_meta" in Key:
            return {"Body": io.BytesIO(meta_buf.getvalue())}
        if "decks.parquet" in Key:
            return {"Body": io.BytesIO(decks_buf.getvalue())}
        raise ValueError(f"Unexpected S3 key: {Key}")

    mock = MagicMock()
    mock.get_object.side_effect = get_object
    return mock


def _embedder_mock(query_vec=None):
    if query_vec is None:
        query_vec = np.zeros(1024, dtype=np.float32)
        query_vec[0] = 1.0
    mock = MagicMock()
    mock._embed_text.return_value = query_vec.tolist()
    return mock


@pytest.fixture
def retriever(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    mocker.patch("ingestion.retriever.boto3.client", return_value=_s3_mock())
    return SlideRetriever(bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock())


# --- k=3 → exactly 3 results ---
def test_search_returns_k_results(retriever):
    assert len(retriever.search("query", k=3)) == 3


# --- results are sorted highest score first ---
def test_search_results_sorted_by_score(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    q = np.zeros(1024, dtype=np.float32)
    q[0], q[1], q[2] = 0.9, 0.5, 0.1  # d1 > d2 > d3
    mocker.patch("ingestion.retriever.boto3.client", return_value=_s3_mock())
    r = SlideRetriever(bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock(q))
    scores = [res["score"] for res in r.search("query", k=3)]
    assert scores == sorted(scores, reverse=True)


# --- every result has all required fields ---
def test_search_result_fields(retriever):
    result = retriever.search("query", k=1)[0]
    for field in ["deck_id", "slide_number", "source_path", "title", "body_text", "score"]:
        assert field in result, f"missing field: {field}"


# --- zero vectors → empty list, embedder never called ---
def test_search_empty_corpus_returns_empty(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    mocker.patch("ingestion.retriever.boto3.client", return_value=_s3_mock(vectors=np.empty((0, 1024), dtype=np.float32)))
    embedder = _embedder_mock()
    r = SlideRetriever(bucket="test-bucket", snapshot_prefix="snapshots", embedder=embedder)
    assert r.search("anything") == []
    embedder._embed_text.assert_not_called()


# --- k larger than corpus → capped at corpus size, no error ---
def test_search_k_capped_at_corpus_size(retriever):
    assert len(retriever.search("query", k=10)) == 3


# --- run_ts from latest.json is used to construct all three data S3 keys ---
def test_load_latest_reads_correct_run_ts(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    s3 = _s3_mock(run_ts="2026-01-15T08:00:00Z")
    mocker.patch("ingestion.retriever.boto3.client", return_value=s3)
    SlideRetriever(bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock())
    keys = [c.kwargs["Key"] for c in s3.get_object.call_args_list]
    data_keys = [k for k in keys if "latest.json" not in k]
    assert all("2026-01-15T08:00:00Z" in k for k in data_keys)


# --- body_text JSON string in Parquet is returned as list[str] ---
def test_body_text_deserialized_from_json(retriever):
    body_text = retriever.search("query", k=1)[0]["body_text"]
    assert isinstance(body_text, list)
    assert all(isinstance(s, str) for s in body_text)


# --- score is a Python float, not numpy float32 ---
def test_search_score_is_float(retriever):
    assert type(retriever.search("query", k=1)[0]["score"]) is float
