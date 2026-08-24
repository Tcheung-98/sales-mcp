"""S3 keys and workbook tab names for Ideation data sources (PI-2759 / I1).

Loaders in Chunks B–D import from here so env defaults stay in one place.
Creation loaders may adopt these constants in a follow-up; today only
``gtm_product_map`` and ``audience_data`` duplicate ``GTM_DATABASE_KEY``.
"""

from __future__ import annotations

import os

# --- S3 object keys (under S3_SNAPSHOT_BUCKET) ---

DEFAULT_GTM_DATABASE_KEY = "templates/Fortune_AITool_GTM_Database.xlsx"
DEFAULT_INVENTORY_CALENDAR_KEY = (
    "templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx"
)

# --- GTM database sheets ---

GTM_SHEET_PRODUCT_TAGS = "Product Tags"
GTM_SHEET_AUDIENCE_DATA = "Audience Data"
GTM_SHEET_PRODUCT_CATEGORY = "Product Category"

# GTM tag strings for Logic Guide live on Product Tags (not a separate sheet).
GTM_COL_PRODUCT_CATEGORY = "Product Category"
GTM_COL_PRODUCT_NAME = "Product Name"
GTM_COL_GTM_TAGS = "GTM TAGS"
GTM_COL_CATEGORY_TITLE = "Title"
GTM_COL_CATEGORY_DESCRIPTION = "Description"

# --- Inventory calendar sheets ---

INVENTORY_SHEET_PRODUCTS = "Products"
INVENTORY_SHEET_PRICING = "Pricing"
INVENTORY_SHEET_BENCHMARKS = "Benchmarks"


def gtm_database_s3_key(key: str | None = None) -> str:
    """Resolved S3 key for the GTM workbook."""
    return key or os.environ.get("GTM_DATABASE_KEY", DEFAULT_GTM_DATABASE_KEY)


def inventory_calendar_s3_key(key: str | None = None) -> str:
    """Resolved S3 key for the inventory + pricing workbook."""
    return key or os.environ.get(
        "INVENTORY_CALENDAR_KEY", DEFAULT_INVENTORY_CALENDAR_KEY
    )
