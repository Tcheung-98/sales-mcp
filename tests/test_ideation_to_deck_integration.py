"""Integration: propose → confirm → assemble + mock AI fills (no live S3)."""

from __future__ import annotations

from unittest.mock import MagicMock

from ingestion.confirm_mix import confirm_mix_from_dict, serialize_ideation
from ingestion.placeholder_fills import apply_placeholders
from ingestion.schema import DeckSchema, DiscoverySchema
from tests.fortuneai_placeholder_fixture import (
    MINIMAL_PNG,
    fortuneai_fixture_bytes,
    mock_placeholder_ai,
    sample_audience_data,
)
from tests.logic_guide_fixtures import base_discovery_fields, build_representative_engine
from tests.test_generator import _blank_bytes, _build_generator, _product_map_for


def _newsletter_discovery() -> dict:
    return base_discovery_fields(
        preferred_platforms_products=["Newsletters"],
        budgets=[{"amount": 250_000.0, "label": "Primary"}],
        targeting_details=(
            "Chief Executive Officer, C-suite, Chief Financial Officer"
        ),
    )


def test_propose_confirm_assemble_no_leftover_tokens():
    engine = build_representative_engine()
    discovery = DiscoverySchema.model_validate(_newsletter_discovery())
    ideation = engine.propose(discovery)
    assert not ideation.requires_gtm_escalation
    assert ideation.tiers

    confirm = confirm_mix_from_dict(
        {
            "discovery": _newsletter_discovery(),
            "ideation": serialize_ideation(ideation),
            "tier_index": 0,
        },
        gtm=engine._gtm,
        inventory=engine._inventory,
    )
    assert confirm["status"] == "ok"
    mix = sum(p["price"] for p in confirm["confirmed_products"])
    deck_data = confirm["deck_schema"]
    deck_data["budgets"] = [{"amount": mix, "label": "Primary"}]

    schema = DeckSchema.model_validate(deck_data)
    generator = _build_generator()
    generator._gtm_product_map = _product_map_for(*schema.confirmed_products)
    template_bytes = fortuneai_fixture_bytes()
    product_deck_bytes = _blank_bytes(slide_count=6)

    def _s3_get_object(*, Bucket, Key, **kwargs):
        data = template_bytes if "FortuneAI" in Key else product_deck_bytes
        return {"Body": MagicMock(read=MagicMock(return_value=data))}

    generator._s3.get_object.side_effect = _s3_get_object
    prs = generator.assemble_skeleton(schema, template_url=None)

    warnings = apply_placeholders(
        prs,
        schema,
        audience=sample_audience_data(),
        logo_bytes=MINIMAL_PNG,
        ai=mock_placeholder_ai(),
    )

    blob = "\n".join(
        para.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
        for para in shape.text_frame.paragraphs
    )
    for token in ("[TITLE]", "[HEADER]", "[BODY]", "[DATE]", "[LOGO]", "[client name]"):
        assert token not in blob, f"leftover {token!r}"
    assert schema.confirmed_products[0].name in blob
    assert warnings is not None
