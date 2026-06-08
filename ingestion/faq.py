import csv
import io
import logging

from docx import Document

from ingestion.retriever import SlideRetriever

logger = logging.getLogger(__name__)


def search_sales_knowledge(
    query: str,
    retriever: SlideRetriever,
    s3,
    bucket: str,
    rate_sheet_key: str | None,
    rulebook_key: str,
    k: int = 5,
) -> dict:
    """
    Search Fortune's sales knowledge base. Returns pricing/rate card data, sales
    guidelines, and relevant slide content for a given query. Use for any question
    about Fortune products, pricing, packaging, or sales approach.
    """

    # --- Rate card ---
    # Fetch the rate card CSV from S3 and format each product as a single readable
    # line of pipe-separated fields. This is the authoritative source for pricing —
    # slide RAG alone will not return accurate numbers.
    # Blank rows (empty Inventory column) are skipped — these are section headers
    # or spacer rows in the spreadsheet.
    rate_card = ""
    if rate_sheet_key:
        resp = s3.get_object(Bucket=bucket, Key=rate_sheet_key)
        reader = csv.DictReader(io.StringIO(resp["Body"].read().decode("utf-8")))
        products = [row for row in reader if row.get("Inventory", "").strip()]
        lines = []
        for p in products:
            parts = [p.get("Inventory", "").strip(), p.get("Product Category", "").strip()]
            if p.get("Cadence", "").strip():
                parts.append(p["Cadence"].strip())
            if p.get("Vertical", "").strip():
                parts.append(f"Verticals: {p['Vertical'].strip()}")
            if p.get("Audience Alignment", "").strip():
                parts.append(f"Audience: {p['Audience Alignment'].strip()}")
            if p.get("Contextual Alignment", "").strip():
                parts.append(f"Context: {p['Contextual Alignment'].strip()}")
            pricing = []
            for label, col in [
                ("Daily", "Daily_Cost"),
                ("Weekly", "Weekly_Cost"),
                ("Monthly", "Monthly_Cost"),
                ("Quarterly", "Quarterly_Cost"),
                ("Half-Year", "Half_Year_Cost"),
                ("Annual", "Annual_Cost"),
                ("CPM", "CPM_Rate"),
                ("Min", "Product_Minimum"),
                ("Flat Fee", "Flat_Fee"),
            ]:
                # Some cells already include a $ sign (e.g. range values like
                # "$60,000-$175,000") — don't double-prefix those.
                val = p.get(col, "").strip()
                if val:
                    prefix = "" if val.startswith("$") else "$"
                    pricing.append(f"{label}: {prefix}{val}")
            if pricing:
                parts.append(" | ".join(pricing))
            if p.get("Notes", "").strip():
                parts.append(f"Notes: {p['Notes'].strip()}")
            lines.append(" | ".join(filter(None, parts)))
        rate_card = "\n".join(lines)
        logger.info("loaded rate card: %d products", len(products))

    # --- Rulebook / guidelines ---
    # Fetch the sales rulebook docx from S3 and extract all non-empty paragraphs.
    # This is the authoritative source for sales approach, tone, and strategy.
    resp = s3.get_object(Bucket=bucket, Key=rulebook_key)
    doc = Document(io.BytesIO(resp["Body"].read()))
    guidelines = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    logger.info("loaded rulebook from s3://%s/%s", bucket, rulebook_key)

    # --- Slide RAG ---
    # Semantic search over the embedded slide corpus. Returns the k most relevant
    # slides as supporting examples and proof points. source_path is a raw SharePoint
    # URL — excluded here since it is not useful in a chat response.
    slides = [
        {field: val for field, val in slide.items() if field != "source_path"}
        for slide in retriever.search(query, k=k)
    ]

    return {
        "rate_card": rate_card,
        "guidelines": guidelines,
        "slides": slides,
    }
