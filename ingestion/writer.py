import io
import json
import logging
import os

import boto3
import pyarrow as pa
import pyarrow.parquet as pq

from ingestion.models import FailedRecord, SlideRow

logger = logging.getLogger(__name__)

_DECKS_SCHEMA = pa.schema([
    ("deck_id", pa.string()),
    ("source_path", pa.string()),
    ("content_hash", pa.string()),
    ("ingested_at", pa.string()),
    ("slide_number", pa.int32()),
    ("layout_name", pa.string()),
    ("title", pa.string()),
    ("body_text", pa.string()),   # JSON-encoded list[str]
    ("tags", pa.string()),        # JSON-encoded Tags dict
    ("tag_sources", pa.string()), # JSON-encoded TagSources dict, or null
])

_FAILED_SCHEMA = pa.schema([
    ("deck_id", pa.string()),
    ("source_path", pa.string()),
    ("error", pa.string()),
    ("failed_at", pa.string()),
    ("layer", pa.string()),
])


class S3ParquetWriter:
    def __init__(
        self,
        bucket: str | None = None,
        snapshot_prefix: str | None = None,
        failed_prefix: str | None = None,
    ):
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._snapshot_prefix = snapshot_prefix or os.environ.get("S3_SNAPSHOT_PREFIX", "snapshots")
        self._failed_prefix = failed_prefix or os.environ.get("S3_FAILED_PREFIX", "failed")
        self._s3 = boto3.client("s3")

    def _upload(self, table: pa.Table, key: str) -> str:
        buf = io.BytesIO()
        pq.write_table(table, buf)
        buf.seek(0)
        self._s3.put_object(Bucket=self._bucket, Key=key, Body=buf.getvalue())
        s3_uri = f"s3://{self._bucket}/{key}"
        logger.info("wrote %d rows to %s", len(table), s3_uri)
        return s3_uri

    def write_decks(self, rows: list[SlideRow], run_ts: str) -> str:
        records = [
            {
                "deck_id": r.deck_id,
                "source_path": r.source_path,
                "content_hash": r.content_hash,
                "ingested_at": r.ingested_at,
                "slide_number": r.slide_number,
                "layout_name": r.layout_name,
                "title": r.title,
                "body_text": json.dumps(r.body_text),
                "tags": r.tags.model_dump_json(),
                "tag_sources": r.tag_sources.model_dump_json() if r.tag_sources else None,
            }
            for r in rows
        ]
        table = pa.Table.from_pylist(records, schema=_DECKS_SCHEMA)
        key = f"{self._snapshot_prefix}/{run_ts}/decks.parquet"
        return self._upload(table, key)

    def write_failed(self, records: list[FailedRecord], run_ts: str) -> str | None:
        if not records:
            return None
        rows = [
            {
                "deck_id": r.deck_id,
                "source_path": r.source_path,
                "error": r.error,
                "failed_at": r.failed_at,
                "layer": r.layer,
            }
            for r in records
        ]
        table = pa.Table.from_pylist(rows, schema=_FAILED_SCHEMA)
        key = f"{self._failed_prefix}/{run_ts}/failed.parquet"
        return self._upload(table, key)
