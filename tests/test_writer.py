import io
import json
from unittest.mock import MagicMock

import pyarrow.parquet as pq
import pytest

from ingestion.models import FailedRecord, SlideRow, Tags
from ingestion.writer import S3ParquetWriter

# --- fixtures ---

SAMPLE_TAGS = Tags(industry="FORTUNE", sub_industry="CASE STUDIES")

SAMPLE_SLIDE = SlideRow(
    deck_id="deck-001",
    source_path="/FORTUNE/CASE STUDIES/acme.pptx",
    content_hash="abc123",
    ingested_at="2026-05-19T10:00:00",
    slide_number=1,
    layout_name="Title Slide",
    title="Acme Corp Pitch",
    body_text=["Line one", "Line two"],
    speaker_notes="Some notes",
    tags=SAMPLE_TAGS,
    tag_sources=None,
)

SAMPLE_FAILED = FailedRecord(
    deck_id="deck-bad",
    source_path="/FORTUNE/bad.pptx",
    error="ZeroDivisionError: division by zero",
    failed_at="2026-05-19T10:01:00",
    layer="bronze",
)


@pytest.fixture
def writer(mocker, monkeypatch):
    monkeypatch.setenv("S3_SNAPSHOT_BUCKET", "test-bucket")
    mocker.patch("boto3.client")
    w = S3ParquetWriter(bucket="test-bucket", snapshot_prefix="snapshots", failed_prefix="failed")
    w._s3 = MagicMock()
    return w


# --- write_decks: uploads to correct key ---
# The S3 key must be snapshots/{run_ts}/decks.parquet so the ingestor can
# rely on a deterministic path for a given run timestamp.
def test_write_decks_uses_correct_s3_key(writer):
    writer.write_decks([SAMPLE_SLIDE], run_ts="2026-05-19T10:00:00Z")
    key = writer._s3.put_object.call_args.kwargs["Key"]
    assert key == "snapshots/2026-05-19T10:00:00Z/decks.parquet"
    assert writer._s3.put_object.call_args.kwargs["Bucket"] == "test-bucket"


# --- write_decks: returns s3 URI ---
def test_write_decks_returns_s3_uri(writer):
    uri = writer.write_decks([SAMPLE_SLIDE], run_ts="2026-05-19T10:00:00Z")
    assert uri == "s3://test-bucket/snapshots/2026-05-19T10:00:00Z/decks.parquet"


# --- write_decks: Parquet schema is correct ---
# Downstream FAISS pipeline depends on these exact column names.
def test_write_decks_parquet_schema(writer):
    captured = {}

    def fake_put(**kwargs):
        captured["body"] = kwargs["Body"]

    writer._s3.put_object = MagicMock(
        side_effect=lambda **kw: captured.update({"body": kw["Body"]})
    )
    writer.write_decks([SAMPLE_SLIDE], run_ts="2026-05-19T10:00:00Z")

    table = pq.read_table(io.BytesIO(captured["body"]))
    col_names = table.schema.names
    for col in ["deck_id", "source_path", "content_hash", "ingested_at",
                "slide_number", "layout_name", "title", "body_text", "tags", "tag_sources"]:
        assert col in col_names, f"missing column: {col}"


# --- write_decks: tags and body_text are JSON strings ---
# pyarrow doesn't support nested structs mixed with nullables cleanly,
# so we serialize to JSON strings for Parquet compatibility.
def test_write_decks_serializes_nested_fields_as_json(writer):
    captured = {}
    writer._s3.put_object = MagicMock(
        side_effect=lambda **kw: captured.update({"body": kw["Body"]})
    )
    writer.write_decks([SAMPLE_SLIDE], run_ts="2026-05-19T10:00:00Z")

    table = pq.read_table(io.BytesIO(captured["body"]))
    row = {col: table.column(col)[0].as_py() for col in table.schema.names}

    tags_dict = json.loads(row["tags"])
    assert tags_dict["industry"] == "FORTUNE"
    assert tags_dict["sub_industry"] == "CASE STUDIES"

    body = json.loads(row["body_text"])
    assert body == ["Line one", "Line two"]

    assert row["tag_sources"] is None


# --- write_failed: uploads to correct key ---
def test_write_failed_uses_correct_s3_key(writer):
    writer.write_failed([SAMPLE_FAILED], run_ts="2026-05-19T10:00:00Z")
    key = writer._s3.put_object.call_args.kwargs["Key"]
    assert key == "failed/2026-05-19T10:00:00Z/failed.parquet"


# --- write_failed: returns None when list is empty ---
# No write should happen; returning None lets the ingestor skip the log line.
def test_write_failed_returns_none_when_empty(writer):
    result = writer.write_failed([], run_ts="2026-05-19T10:00:00Z")
    assert result is None
    writer._s3.put_object.assert_not_called()


# --- write_failed: returns s3 URI when records present ---
def test_write_failed_returns_s3_uri(writer):
    uri = writer.write_failed([SAMPLE_FAILED], run_ts="2026-05-19T10:00:00Z")
    assert uri == "s3://test-bucket/failed/2026-05-19T10:00:00Z/failed.parquet"


# --- write_embeddings: uploads to correct S3 keys ---
def test_write_embeddings_uses_correct_s3_keys(writer):
    import numpy as np
    vectors = np.zeros((3, 1024), dtype=np.float32)
    meta = [{"deck_id": "d1", "slide_number": 1, "source_path": "/p.pptx"}] * 3
    writer.write_embeddings(vectors, meta, run_ts="2026-05-21T00:00:00Z")

    calls = writer._s3.put_object.call_args_list
    keys = [c.kwargs["Key"] for c in calls]
    assert "snapshots/2026-05-21T00:00:00Z/embeddings.npy" in keys
    assert "snapshots/2026-05-21T00:00:00Z/embeddings_meta.parquet" in keys


# --- write_embeddings: returns both S3 URIs ---
def test_write_embeddings_returns_uris(writer):
    import numpy as np
    vectors = np.zeros((2, 1024), dtype=np.float32)
    meta = [{"deck_id": "d1", "slide_number": i, "source_path": "/p.pptx"} for i in range(2)]
    npy_uri, meta_uri = writer.write_embeddings(vectors, meta, run_ts="2026-05-21T00:00:00Z")
    assert npy_uri == "s3://test-bucket/snapshots/2026-05-21T00:00:00Z/embeddings.npy"
    assert meta_uri == "s3://test-bucket/snapshots/2026-05-21T00:00:00Z/embeddings_meta.parquet"


# --- write_embeddings: meta parquet has correct columns ---
def test_write_embeddings_meta_parquet_schema(writer):
    import io

    import numpy as np
    import pyarrow.parquet as pq

    vectors = np.zeros((1, 1024), dtype=np.float32)
    meta = [{"deck_id": "d1", "slide_number": 1, "source_path": "/p.pptx"}]

    captured = {}
    orig_put = writer._s3.put_object

    def capture_put(**kwargs):
        if kwargs["Key"].endswith(".parquet"):
            captured["body"] = kwargs["Body"]
        return orig_put(**kwargs)

    writer._s3.put_object = MagicMock(side_effect=capture_put)
    writer.write_embeddings(vectors, meta, run_ts="2026-05-21T00:00:00Z")

    table = pq.read_table(io.BytesIO(captured["body"]))
    assert set(table.schema.names) >= {"deck_id", "slide_number", "source_path"}


# --- write_decks: multiple rows all written ---
def test_write_decks_multiple_rows(writer):
    slide2 = SAMPLE_SLIDE.model_copy(update={"slide_number": 2, "title": "Slide Two"})
    captured = {}
    writer._s3.put_object = MagicMock(
        side_effect=lambda **kw: captured.update({"body": kw["Body"]})
    )
    writer.write_decks([SAMPLE_SLIDE, slide2], run_ts="2026-05-19T10:00:00Z")

    table = pq.read_table(io.BytesIO(captured["body"]))
    assert len(table) == 2
