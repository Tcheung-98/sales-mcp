"""Tests for review-package manifest stub (PI-2518 / A2)."""

from ingestion.manifest import (
    ReviewManifest,
    manifest_validation_errors,
    validate_manifest,
)


def _valid_manifest(**overrides) -> dict:
    data = {
        "schema_version": "1",
        "client_name": "Acme Corp",
        "template_key": "Category_Presentation_Technology.pptx",
        "slide_count": 2,
        "slides": [
            {"slide_index": 0, "role": "narrative"},
            {
                "slide_index": 1,
                "role": "product",
                "product_name": "Fortune 500 List",
                "source_path": "corpus/deck.pptx",
                "source_slide_number": 3,
            },
        ],
    }
    data.update(overrides)
    return data


def test_validate_manifest_accepts_valid():
    m = validate_manifest(_valid_manifest())
    assert isinstance(m, ReviewManifest)
    assert m.client_name == "Acme Corp"
    assert m.slide_count == 2
    assert m.slides[1].role == "product"


def test_validate_manifest_defaults_schema_version():
    data = _valid_manifest()
    del data["schema_version"]
    m = validate_manifest(data)
    assert m.schema_version == "1"


def test_manifest_validation_errors_empty_when_valid():
    assert manifest_validation_errors(_valid_manifest()) == []


def test_manifest_validation_errors_on_missing_fields():
    errors = manifest_validation_errors({"client_name": "Acme"})
    assert errors
    assert any("slide_count" in e or "slides" in e or "template_key" in e for e in errors)


def test_manifest_rejects_bad_role():
    data = _valid_manifest()
    data["slides"][0]["role"] = "stylist"
    errors = manifest_validation_errors(data)
    assert any("role" in e for e in errors)
