import io
import json
import logging
import os

import boto3
import numpy as np
import pyarrow.parquet as pq

from ingestion.embedder import SlideEmbedder

logger = logging.getLogger(__name__)


def _tag_match(value: str, filter_val: str | None) -> bool:
    if not filter_val:
        return True
    return value.lower() == filter_val.lower()


class SlideRetriever:
    def __init__(
        self,
        bucket: str | None = None,
        snapshot_prefix: str | None = None,
        embedder: SlideEmbedder | None = None,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._snapshot_prefix = snapshot_prefix or os.environ.get("S3_SNAPSHOT_PREFIX", "snapshots")
        self._s3 = boto3.client("s3")
        self._embedder = embedder or SlideEmbedder()
        self._vectors: np.ndarray = np.empty((0, 1024), dtype=np.float32)
        self._meta: list[dict] = []
        self._slides: dict[tuple[str, int], dict] = {}
        self._load_latest()

    def _load_latest(self) -> None:
        manifest_key = f"{self._snapshot_prefix}/latest.json"
        resp = self._s3.get_object(Bucket=self._bucket, Key=manifest_key)
        run_ts = json.loads(resp["Body"].read())["run_ts"]

        npy_key = f"{self._snapshot_prefix}/{run_ts}/embeddings.npy"
        resp = self._s3.get_object(Bucket=self._bucket, Key=npy_key)
        self._vectors = np.load(io.BytesIO(resp["Body"].read()))

        meta_key = f"{self._snapshot_prefix}/{run_ts}/embeddings_meta.parquet"
        resp = self._s3.get_object(Bucket=self._bucket, Key=meta_key)
        self._meta = pq.read_table(io.BytesIO(resp["Body"].read())).to_pylist()

        decks_key = f"{self._snapshot_prefix}/{run_ts}/decks.parquet"
        resp = self._s3.get_object(Bucket=self._bucket, Key=decks_key)
        self._slides = {}
        for row in pq.read_table(io.BytesIO(resp["Body"].read())).to_pylist():
            row["body_text"] = json.loads(row["body_text"])
            self._slides[(row["deck_id"], row["slide_number"])] = row

        logger.info("loaded %d vectors from snapshot %s", len(self._vectors), run_ts)

    def search(self, query_text: str, k: int = 5) -> list[dict]:
        if len(self._vectors) == 0:
            return []

        query_vec = np.array(self._embedder._embed_text(query_text), dtype=np.float32)
        scores = self._vectors @ query_vec
        k = min(k, len(scores))
        top_k_idx = np.argsort(scores)[-k:][::-1]

        results = []
        for idx in top_k_idx:
            meta = self._meta[idx]
            slide = self._slides.get((meta["deck_id"], meta["slide_number"]), {})
            results.append({
                "deck_id": meta["deck_id"],
                "slide_number": meta["slide_number"],
                "source_path": meta["source_path"],
                "title": slide.get("title"),
                "body_text": slide.get("body_text", []),
                "score": float(scores[idx]),
            })

        return results

    def get_slide_content(
        self, deck_id: str, slide_numbers: list[int] | None = None
    ) -> list[dict]:
        if slide_numbers is None:
            rows = sorted(
                (v for (did, _), v in self._slides.items() if did == deck_id),
                key=lambda r: r["slide_number"],
            )
        else:
            rows = [
                self._slides[(deck_id, sn)]
                for sn in slide_numbers
                if (deck_id, sn) in self._slides
            ]

        return [
            {
                "deck_id": row["deck_id"],
                "slide_number": row["slide_number"],
                "title": row.get("title"),
                "body_text": row.get("body_text", []),
                "layout_name": row.get("layout_name"),
                "source_path": row["source_path"],
            }
            for row in rows
        ]

    def filter_decks_by_tags(
        self,
        industry: str | None = None,
        sub_industry: str | None = None,
        product_line: str | None = None,
        deal_size: str | None = None,
        client_name: str | None = None,
        date_from: str | None = None,
        date_to: str | None = None,
        deck_type: str | None = None,
        limit: int = 20,
    ) -> list[dict]:
        # Build deck-level view by aggregating slide rows (tags are deck-level, same on every slide)
        decks: dict[str, dict] = {}
        for (deck_id, slide_number), row in self._slides.items():
            if deck_id not in decks:
                raw = row.get("tags", "{}")
                tags = json.loads(raw) if isinstance(raw, str) else raw
                decks[deck_id] = {
                    "deck_id": deck_id,
                    "source_path": row["source_path"],
                    "tags": tags,
                    "slide_count": 0,
                    "title": None,
                }
            decks[deck_id]["slide_count"] += 1
            if slide_number == 1:
                decks[deck_id]["title"] = row.get("title")

        results = []
        for deck in decks.values():
            tags = deck["tags"]
            if not _tag_match(tags.get("industry", ""), industry):
                continue
            if not _tag_match(tags.get("sub_industry", ""), sub_industry):
                continue
            if not _tag_match(tags.get("product_line", ""), product_line):
                continue
            if not _tag_match(tags.get("deal_size", ""), deal_size):
                continue
            if not _tag_match(tags.get("client_name", ""), client_name):
                continue
            if not _tag_match(tags.get("deck_type", ""), deck_type):
                continue
            date = tags.get("date", "")
            if date_from and (not date or date < date_from):
                continue
            if date_to and (not date or date > date_to):
                continue
            results.append({
                "deck_id": deck["deck_id"],
                "title": deck["title"],
                "tags": deck["tags"],
                "slide_count": deck["slide_count"],
                "source_path": deck["source_path"],
            })

        return results[:limit]
