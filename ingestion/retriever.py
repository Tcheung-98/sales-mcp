import io
import json
import logging
import os

import boto3
import numpy as np
import pyarrow.parquet as pq

from ingestion.embedder import SlideEmbedder

logger = logging.getLogger(__name__)


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
