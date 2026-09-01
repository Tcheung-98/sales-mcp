"""Combined inventory calendar loaders (Products, Inventory, Pricing)."""

from __future__ import annotations

from dataclasses import dataclass

from ingestion.ideation_data_keys import inventory_calendar_s3_key
from ingestion.inventory_calendar import InventoryCalendar, load_inventory_calendar_from_s3
from ingestion.inventory_pricing import InventoryPricing, load_inventory_pricing_from_s3


@dataclass(frozen=True)
class InventoryWorkbook:
    """Full Ideation inventory xlsx: availability + pricing."""

    calendar: InventoryCalendar
    pricing: InventoryPricing

    @classmethod
    def from_xlsx_bytes(cls, data: bytes) -> InventoryWorkbook:
        return cls(
            calendar=InventoryCalendar.from_xlsx_bytes(data),
            pricing=InventoryPricing.from_xlsx_bytes(data),
        )


def load_inventory_workbook_from_s3(
    s3_client, bucket: str, key: str | None = None
) -> InventoryWorkbook:
    """Load Products, Inventory, and Pricing + Benchmarks from one S3 object."""
    s3_key = inventory_calendar_s3_key(key)
    resp = s3_client.get_object(Bucket=bucket, Key=s3_key)
    return InventoryWorkbook.from_xlsx_bytes(resp["Body"].read())


__all__ = [
    "InventoryWorkbook",
    "load_inventory_workbook_from_s3",
    "load_inventory_calendar_from_s3",
    "load_inventory_pricing_from_s3",
]
