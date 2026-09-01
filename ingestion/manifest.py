"""Review-package manifest stub for Cursor stylist handoff (Phase B).

``assemble_skeleton`` returns ``AssembledSkeleton`` with product-clone provenance;
B2 will write draft.pptx + PNGs + a full ``ReviewManifest`` alongside this schema.
"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, ValidationError

MANIFEST_SCHEMA_VERSION: Literal["1"] = "1"


class SlideManifestEntry(BaseModel):
    slide_index: int = Field(ge=0, description="0-based index in the draft PPTX")
    role: Literal["cover", "narrative", "product", "other"]
    product_name: str | None = None
    source_path: str | None = None
    source_slide_number: int | None = Field(
        default=None, ge=1, description="1-based corpus slide when role=product"
    )


class ReviewManifest(BaseModel):
    """Stub contract for the Cursor review package (B2+)."""

    schema_version: Literal["1"] = MANIFEST_SCHEMA_VERSION
    client_name: str
    template_key: str
    slide_count: int = Field(ge=1)
    slides: list[SlideManifestEntry]


def validate_manifest(data: dict) -> ReviewManifest:
    """Parse and validate a manifest dict. Raises pydantic.ValidationError."""
    return ReviewManifest.model_validate(data)


def manifest_validation_errors(data: dict) -> list[str]:
    """Return human-readable errors, or [] if valid."""
    try:
        validate_manifest(data)
    except ValidationError as exc:
        return [
            f"{'.'.join(str(p) for p in e['loc'])}: {e['msg']}" for e in exc.errors()
        ]
    return []
