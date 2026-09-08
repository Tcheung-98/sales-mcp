"""Review-package manifest for the deck QA rail (docs/DECK-QA-ARCHITECTURE.md §5).

``assemble_skeleton`` returns a bare ``Presentation``; ``ingestion.review_package``
derives product-clone provenance from ``plan_pitch_sequence`` and writes
draft.pptx + PNGs + this manifest alongside the serialized DeckSchema.
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, ValidationError, model_validator

MANIFEST_SCHEMA_VERSION: Literal["1"] = "1"


class SlideManifestEntry(BaseModel):
    slide_index: int = Field(ge=0, description="0-based index in the draft PPTX")
    role: Literal["cover", "narrative", "product", "other"]
    slide_kind: str | None = Field(
        default=None,
        description="Discriminator inside role=other: divider / investment / thank_you",
    )
    editable: bool = Field(
        default=True,
        description="False on A5 product clones — QA may flag them but never edit them",
    )
    product_name: str | None = None
    source_path: str | None = None
    source_slide_number: int | None = Field(
        default=None, ge=1, description="1-based corpus slide when role=product"
    )

    @model_validator(mode="before")
    @classmethod
    def default_editable_from_role(cls, data: Any) -> Any:
        if isinstance(data, dict) and "editable" not in data:
            return {**data, "editable": data.get("role") != "product"}
        return data


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
