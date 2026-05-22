import logging
from datetime import datetime, timezone

from ingestion.embedder import SlideEmbedder
from ingestion.graph_client import GraphClient
from ingestion.models import FailedRecord, SlideRow
from ingestion.parser import parse_pptx
from ingestion.writer import S3ParquetWriter

logger = logging.getLogger(__name__)


def run_ingest(
    client: GraphClient | None = None,
    writer: S3ParquetWriter | None = None,
    embedder: SlideEmbedder | None = None,
) -> tuple[list[SlideRow], list[FailedRecord]]:
    client = client or GraphClient()
    writer = writer or S3ParquetWriter()
    embedder = embedder or SlideEmbedder()

    run_ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    decks = client.list_decks()
    logger.info("found %d decks to ingest", len(decks))

    slides: list[SlideRow] = []
    failed: list[FailedRecord] = []

    for item in decks:
        deck_id = item["id"]
        source_path = item.get("webUrl", item["name"])
        try:
            pptx_bytes = client.download_deck(deck_id)
            tags = client.extract_tags(item)
            rows = parse_pptx(
                pptx_bytes, deck_id=deck_id, source_path=source_path,
                ingested_at=run_ts, tags=tags,
            )
            slides.extend(rows)
        except Exception as exc:
            logger.warning("deck %s failed: %s", deck_id, exc)
            failed.append(FailedRecord(
                deck_id=deck_id,
                source_path=source_path,
                error=str(exc),
                failed_at=run_ts,
                layer="bronze",
            ))

    writer.write_decks(slides, run_ts=run_ts)
    writer.write_failed(failed, run_ts=run_ts)
    vectors, meta = embedder.embed_slides(slides)
    writer.write_embeddings(vectors, meta, run_ts=run_ts)
    logger.info("ingest complete — %d slides, %d failed decks", len(slides), len(failed))

    return slides, failed
