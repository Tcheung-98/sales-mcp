"""Env defaults for I1 Ideation data sources."""

from ingestion.ideation_data_keys import (
    DEFAULT_GTM_DATABASE_KEY,
    DEFAULT_INVENTORY_CALENDAR_KEY,
    GTM_SHEET_GTM_TAGS,
    INVENTORY_SHEET_PRICING,
    INVENTORY_SHEET_PRODUCTS,
    gtm_database_s3_key,
    inventory_calendar_s3_key,
)


def test_default_s3_keys():
    assert DEFAULT_GTM_DATABASE_KEY == "templates/Fortune_AITool_GTM_Database.xlsx"
    assert (
        DEFAULT_INVENTORY_CALENDAR_KEY
        == "templates/Fortune_Inventory_Reservation_Calendar_2026_Final.xlsx"
    )


def test_sheet_name_constants():
    assert GTM_SHEET_GTM_TAGS == "GTM Tags"
    assert INVENTORY_SHEET_PRODUCTS == "Products"
    assert INVENTORY_SHEET_PRICING == "Pricing"


def test_resolved_keys_from_env(monkeypatch):
    monkeypatch.setenv("GTM_DATABASE_KEY", "custom/gtm.xlsx")
    monkeypatch.setenv("INVENTORY_CALENDAR_KEY", "custom/inv.xlsx")
    assert gtm_database_s3_key() == "custom/gtm.xlsx"
    assert inventory_calendar_s3_key() == "custom/inv.xlsx"


def test_explicit_key_overrides_env(monkeypatch):
    monkeypatch.setenv("GTM_DATABASE_KEY", "env/gtm.xlsx")
    assert gtm_database_s3_key("explicit/gtm.xlsx") == "explicit/gtm.xlsx"


def test_defaults_when_env_unset(monkeypatch):
    monkeypatch.delenv("GTM_DATABASE_KEY", raising=False)
    monkeypatch.delenv("INVENTORY_CALENDAR_KEY", raising=False)
    assert gtm_database_s3_key() == DEFAULT_GTM_DATABASE_KEY
    assert inventory_calendar_s3_key() == DEFAULT_INVENTORY_CALENDAR_KEY
