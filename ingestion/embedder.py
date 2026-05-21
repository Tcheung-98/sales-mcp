import json
import logging
import os
import random
import time

import boto3
import numpy as np
from botocore.exceptions import ClientError

from ingestion.models import SlideRow

logger = logging.getLogger(__name__)

_DEFAULT_MODEL_ID = "amazon.titan-embed-text-v2:0"
_DIMENSIONS = 1024
_MAX_RETRIES = 5


class SlideEmbedder:
    def __init__(self, region: str | None = None, model_id: str | None = None) -> None:
        self._region = region or os.environ.get("AWS_REGION", "us-east-1")
        self._model_id = model_id or os.environ.get(
            "BEDROCK_EMBEDDING_MODEL_ID", _DEFAULT_MODEL_ID
        )
        self._client = boto3.client("bedrock-runtime", region_name=self._region)

    def _embed_text(self, text: str) -> list[float]:
        payload = json.dumps({"inputText": text, "dimensions": _DIMENSIONS, "normalize": True})
        last_exc: Exception | None = None
        for attempt in range(_MAX_RETRIES + 1):
            try:
                response = self._client.invoke_model(
                    modelId=self._model_id,
                    body=payload,
                    contentType="application/json",
                    accept="application/json",
                )
                embedding = json.loads(response["body"].read())["embedding"]
                if len(embedding) != _DIMENSIONS:
                    raise RuntimeError(
                        f"Unexpected embedding size: got {len(embedding)}, expected {_DIMENSIONS}"
                    )
                return embedding
            except ClientError as exc:
                if exc.response["Error"]["Code"] != "ThrottlingException":
                    raise
                last_exc = exc
                if attempt < _MAX_RETRIES:
                    time.sleep((2 ** attempt) + random.uniform(0, 0.25))
        raise RuntimeError(
            f"Bedrock throttled after {_MAX_RETRIES} retries"
        ) from last_exc

    def embed_slides(
        self, rows: list[SlideRow]
    ) -> tuple[np.ndarray, list[dict]]:
        meta: list[dict] = []
        vectors: list[list[float]] = []

        for row in rows:
            if row.tags.deck_type == "template":
                continue
            text = f"{row.title or ''} {' '.join(row.body_text)}".strip()
            if not text:
                continue
            vec = self._embed_text(text)
            vectors.append(vec)
            meta.append({
                "deck_id": row.deck_id,
                "slide_number": row.slide_number,
                "source_path": row.source_path,
            })

        if not vectors:
            return np.empty((0, _DIMENSIONS), dtype=np.float32), []

        return np.array(vectors, dtype=np.float32), meta
