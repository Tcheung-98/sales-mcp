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
    return SlideRetriever(
        bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock()
    )


# --- k=3 → exactly 3 results ---
def test_search_returns_k_results(retriever):
    assert len(retriever.search("query", k=3)) == 3


# --- results are sorted highest score first ---
def test_search_results_sorted_by_score(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    q = np.zeros(1024, dtype=np.float32)
    q[0], q[1], q[2] = 0.9, 0.5, 0.1  # d1 > d2 > d3
    mocker.patch("ingestion.retriever.boto3.client", return_value=_s3_mock())
    r = SlideRetriever(
        bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock(q)
    )
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
    empty = np.empty((0, 1024), dtype=np.float32)
    mocker.patch("ingestion.retriever.boto3.client", return_value=_s3_mock(vectors=empty))
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


# ---------------------------------------------------------------------------
# get_slide_content tests
# ---------------------------------------------------------------------------

_MULTI_DECKS = [
    {
        "deck_id": "d1", "slide_number": 1, "source_path": "/GTM/d1.pptx",
        "content_hash": "a", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Slide One", "body_text": json.dumps(["Alpha"]),
        "tags": "{}", "tag_sources": None,
    },
    {
        "deck_id": "d1", "slide_number": 2, "source_path": "/GTM/d1.pptx",
        "content_hash": "b", "ingested_at": RUN_TS, "layout_name": "Content",
        "title": "Slide Two", "body_text": json.dumps(["Bravo", "Charlie"]),
        "tags": "{}", "tag_sources": None,
    },
    {
        "deck_id": "d1", "slide_number": 3, "source_path": "/GTM/d1.pptx",
        "content_hash": "c", "ingested_at": RUN_TS, "layout_name": "Content",
        "title": None, "body_text": json.dumps(["Delta"]),
        "tags": "{}", "tag_sources": None,
    },
]

_MULTI_VECTORS = np.eye(3, 1024, dtype=np.float32)
_MULTI_META = [
    {"deck_id": "d1", "slide_number": 1, "source_path": "/GTM/d1.pptx"},
    {"deck_id": "d1", "slide_number": 2, "source_path": "/GTM/d1.pptx"},
    {"deck_id": "d1", "slide_number": 3, "source_path": "/GTM/d1.pptx"},
]


@pytest.fixture
def multi_retriever(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    mocker.patch(
        "ingestion.retriever.boto3.client",
        return_value=_s3_mock(
            vectors=_MULTI_VECTORS, meta=_MULTI_META, decks=_MULTI_DECKS
        ),
    )
    return SlideRetriever(
        bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock()
    )


# --- no slide_numbers → all slides returned, sorted by slide_number ---
def test_get_slide_content_returns_all_slides(multi_retriever):
    results = multi_retriever.get_slide_content("d1")
    assert len(results) == 3
    assert [r["slide_number"] for r in results] == [1, 2, 3]


# --- specific slide_numbers → only those slides returned ---
def test_get_slide_content_filters_by_slide_numbers(multi_retriever):
    results = multi_retriever.get_slide_content("d1", slide_numbers=[1, 3])
    assert len(results) == 2
    assert {r["slide_number"] for r in results} == {1, 3}


# --- unknown deck_id → empty list, no error ---
def test_get_slide_content_missing_deck_id_returns_empty(retriever):
    assert retriever.get_slide_content("nonexistent") == []


# --- slide numbers that don't exist are silently skipped ---
def test_get_slide_content_missing_slide_numbers_skipped(multi_retriever):
    results = multi_retriever.get_slide_content("d1", slide_numbers=[2, 99])
    assert len(results) == 1
    assert results[0]["slide_number"] == 2


# --- all required output fields are present ---
def test_get_slide_content_result_fields(multi_retriever):
    result = multi_retriever.get_slide_content("d1", slide_numbers=[1])[0]
    for field in ["deck_id", "slide_number", "title", "body_text", "layout_name", "source_path"]:
        assert field in result, f"missing field: {field}"


# --- body_text is list[str], not JSON string ---
def test_get_slide_content_body_text_is_list(multi_retriever):
    result = multi_retriever.get_slide_content("d1", slide_numbers=[2])[0]
    assert isinstance(result["body_text"], list)
    assert result["body_text"] == ["Bravo", "Charlie"]


# --- None title is preserved (not coerced to empty string) ---
def test_get_slide_content_none_title_preserved(multi_retriever):
    result = multi_retriever.get_slide_content("d1", slide_numbers=[3])[0]
    assert result["title"] is None


# --- slide_numbers=[] → empty list, no error ---
def test_get_slide_content_empty_slide_numbers_returns_empty(multi_retriever):
    assert multi_retriever.get_slide_content("d1", slide_numbers=[]) == []


# ---------------------------------------------------------------------------
# filter_decks_by_tags tests
# ---------------------------------------------------------------------------

def _tags(industry="", sub_industry="", product_line="", deal_size="",
          client_name="", date="", deck_type="", status=""):
    return json.dumps({
        "industry": industry, "sub_industry": sub_industry,
        "product_line": product_line, "deal_size": deal_size,
        "client_name": client_name, "date": date,
        "deck_type": deck_type, "status": status,
    })


_FILTER_DECKS = [
    # tech1: 2 slides, Technology / Enterprise SaaS / Pitch / 2026-01-15
    {
        "deck_id": "tech1", "slide_number": 1, "source_path": "/Technology/tech1.pptx",
        "content_hash": "a", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Tech Pitch One", "body_text": json.dumps(["A"]),
        "tags": _tags(
            industry="Technology", sub_industry="Enterprise SaaS",
            deck_type="Pitch", date="2026-01-15",
        ),
        "tag_sources": None,
    },
    {
        "deck_id": "tech1", "slide_number": 2, "source_path": "/Technology/tech1.pptx",
        "content_hash": "b", "ingested_at": RUN_TS, "layout_name": "Content",
        "title": "Slide 2", "body_text": json.dumps(["B"]),
        "tags": _tags(
            industry="Technology", sub_industry="Enterprise SaaS",
            deck_type="Pitch", date="2026-01-15",
        ),
        "tag_sources": None,
    },
    # fin1: 1 slide, Finance / Proposal / 2026-03-10
    {
        "deck_id": "fin1", "slide_number": 1, "source_path": "/Finance/fin1.pptx",
        "content_hash": "c", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Finance Proposal", "body_text": json.dumps(["C"]),
        "tags": _tags(industry="Finance", deck_type="Proposal", date="2026-03-10"),
        "tag_sources": None,
    },
    # tech2: 1 slide, Technology / Media / Pitch / 2025-11-20
    {
        "deck_id": "tech2", "slide_number": 1, "source_path": "/Technology/tech2.pptx",
        "content_hash": "d", "ingested_at": RUN_TS, "layout_name": "Title",
        "title": "Tech Pitch Two", "body_text": json.dumps(["D"]),
        "tags": _tags(
            industry="Technology", sub_industry="Media",
            deck_type="Pitch", date="2025-11-20",
        ),
        "tag_sources": None,
    },
]

_FILTER_VECTORS = np.eye(4, 1024, dtype=np.float32)
_FILTER_META = [
    {"deck_id": "tech1", "slide_number": 1, "source_path": "/Technology/tech1.pptx"},
    {"deck_id": "tech1", "slide_number": 2, "source_path": "/Technology/tech1.pptx"},
    {"deck_id": "fin1",  "slide_number": 1, "source_path": "/Finance/fin1.pptx"},
    {"deck_id": "tech2", "slide_number": 1, "source_path": "/Technology/tech2.pptx"},
]


@pytest.fixture
def filter_retriever(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    mocker.patch(
        "ingestion.retriever.boto3.client",
        return_value=_s3_mock(
            vectors=_FILTER_VECTORS, meta=_FILTER_META, decks=_FILTER_DECKS
        ),
    )
    return SlideRetriever(
        bucket="test-bucket", snapshot_prefix="snapshots", embedder=_embedder_mock()
    )


# --- no filters → all 3 unique decks returned ---
def test_filter_no_filters_returns_all_decks(filter_retriever):
    results = filter_retriever.filter_decks_by_tags()
    assert {r["deck_id"] for r in results} == {"tech1", "fin1", "tech2"}


# --- industry filter → only matching decks ---
def test_filter_by_industry(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(industry="Technology")
    assert {r["deck_id"] for r in results} == {"tech1", "tech2"}


# --- AND semantics: industry + deck_type must both match ---
def test_filter_and_semantics(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(industry="Finance", deck_type="Pitch")
    assert results == []


# --- no corpus match → empty list, no error ---
def test_filter_no_match_returns_empty(filter_retriever):
    assert filter_retriever.filter_decks_by_tags(industry="Healthcare") == []


# --- filter is case-insensitive ---
def test_filter_case_insensitive(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(industry="technology")
    assert {r["deck_id"] for r in results} == {"tech1", "tech2"}


# --- limit caps the number of results ---
def test_filter_limit(filter_retriever):
    assert len(filter_retriever.filter_decks_by_tags(limit=1)) == 1


# --- date_from excludes decks older than the threshold ---
def test_filter_date_from(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(date_from="2026-01-01")
    ids = {r["deck_id"] for r in results}
    assert "tech1" in ids and "fin1" in ids
    assert "tech2" not in ids  # 2025-11-20 < 2026-01-01


# --- date_to excludes decks newer than the threshold ---
def test_filter_date_to(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(date_to="2026-02-01")
    ids = {r["deck_id"] for r in results}
    assert "tech1" in ids and "tech2" in ids
    assert "fin1" not in ids  # 2026-03-10 > 2026-02-01


# --- required output fields present ---
def test_filter_result_fields(filter_retriever):
    result = filter_retriever.filter_decks_by_tags(industry="Finance")[0]
    for field in ["deck_id", "title", "tags", "slide_count", "source_path"]:
        assert field in result, f"missing field: {field}"


# --- slide_count reflects actual number of slides per deck ---
def test_filter_slide_count_correct(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(
        industry="Technology", sub_industry="Enterprise SaaS"
    )
    assert results[0]["deck_id"] == "tech1"
    assert results[0]["slide_count"] == 2


# --- title comes from slide 1 ---
def test_filter_deck_title_from_slide_1(filter_retriever):
    results = filter_retriever.filter_decks_by_tags(industry="Finance")
    assert results[0]["title"] == "Finance Proposal"


# --- filtering on a sparse field (empty string) excludes decks that don't have it set ---
def test_filter_sparse_field_excludes_unset(filter_retriever):
    # product_line is "" on all decks in fixture; filtering by a value should return nothing
    assert filter_retriever.filter_decks_by_tags(product_line="Enterprise") == []
