"""B2 review package: draft.pptx + slide PNGs + manifest (docs/DECK-QA-ARCHITECTURE.md §5).

Built from the post-C2 deck, i.e. after ``apply_placeholders`` has deleted the
unused audience/program variants, so the deck is ``8 + P`` slides where
``P = funded dividers + confirmed products``. Roles come from that index map, not
from slide content: indices 1–5 are the surviving narrative spine and everything
from 6 to ``5 + P`` is a divider or an A5 product clone.

``plan_pitch_sequence`` is the single source of pitch-section order; PR-F
refactors ``assemble_skeleton`` onto it so the ordering has one implementation.
"""

from __future__ import annotations

import io
import json
import logging
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from pptx.oxml.ns import qn

from ingestion.category_dividers import (
    CATEGORY_DIVIDERS,
    FORTUNEAI_DIVIDER_COUNT,
    FORTUNEAI_TEMPLATE_BASENAME,
    divider_index_for_category,
)
from ingestion.gtm_product_map import GtmProductMap, ProductSlideRef
from ingestion.manifest import ReviewManifest, SlideManifestEntry
from ingestion.render_slides import render_slides
from ingestion.schema import DeckSchema

logger = logging.getLogger(__name__)

PitchItem = tuple[Literal["divider"], int] | tuple[Literal["product"], ProductSlideRef]

DRAFT_NAME = "draft.pptx"
MANIFEST_NAME = "manifest.json"
DECK_SCHEMA_NAME = "deck_schema.json"
SLIDES_DIRNAME = "slides"
DETERMINISTIC_REPORT_NAME = "qa_deterministic.json"
CURSOR_REPORT_NAME = "qa_cursor.json"

RENDER_DPI = 150

# Post-C2 geometry (§5): 0 cover, 1–5 narrative, 6..5+P pitch, then investment + thanks.
COVER_INDEX = 0
PITCH_START_INDEX = 6
_TAIL_SLIDE_COUNT = 2
_FIXED_SLIDE_COUNT = PITCH_START_INDEX + _TAIL_SLIDE_COUNT  # the "8" in 8 + P


def plan_pitch_sequence(
    schema: DeckSchema, gtm_map: GtmProductMap
) -> list[PitchItem]:
    """Funded dividers + A5 clones in Workflow pitch order.

    Drives the ordering loop in ``DeckGenerator.assemble_skeleton``, which calls
    this rather than keeping a second copy. Pure: no S3, no template load. Raises
    ValueError with every placement failure at once (missing GTM row, Events
    escalation).
    """
    groups: list[list[ProductSlideRef]] = [[] for _ in CATEGORY_DIVIDERS]
    failures: list[str] = []
    for product in schema.confirmed_products:
        try:
            divider_i = divider_index_for_category(product.category)
            ref = gtm_map.lookup(product.name, product.category)
        except ValueError as exc:
            failures.append(str(exc))
            continue
        groups[divider_i].append(ref)
    if failures:
        raise ValueError(
            "FortuneAI product placement failed: " + "; ".join(failures)
        )

    plan: list[PitchItem] = []
    for i in range(FORTUNEAI_DIVIDER_COUNT):
        if not groups[i]:
            continue
        plan.append(("divider", i))
        plan.extend(("product", ref) for ref in groups[i])
    return plan


def build_manifest(
    schema: DeckSchema,
    plan: list[PitchItem],
    *,
    slide_count: int,
    template_key: str = FORTUNEAI_TEMPLATE_BASENAME,
) -> ReviewManifest:
    """Assign §5 roles by index. Product entries carry A5 provenance and editable=false."""
    expected = _FIXED_SLIDE_COUNT + len(plan)
    if slide_count != expected:
        raise ValueError(
            f"post-C2 deck must have {expected} slides "
            f"({_FIXED_SLIDE_COUNT} stock + {len(plan)} pitch); got {slide_count}"
        )

    slides = [SlideManifestEntry(slide_index=COVER_INDEX, role="cover")]
    slides += [
        SlideManifestEntry(slide_index=i, role="narrative")
        for i in range(COVER_INDEX + 1, PITCH_START_INDEX)
    ]
    for offset, (kind, payload) in enumerate(plan):
        index = PITCH_START_INDEX + offset
        if kind == "divider":
            slides.append(
                SlideManifestEntry(
                    slide_index=index, role="other", slide_kind="divider"
                )
            )
        else:
            slides.append(
                SlideManifestEntry(
                    slide_index=index,
                    role="product",
                    product_name=payload.product_name,
                    # Raw GTM Deck Path, not the product-decks/ S3 key.
                    source_path=payload.deck_path,
                    source_slide_number=payload.slide_number,
                )
            )
    slides.append(
        SlideManifestEntry(
            slide_index=slide_count - 2, role="other", slide_kind="investment"
        )
    )
    slides.append(
        SlideManifestEntry(
            slide_index=slide_count - 1, role="other", slide_kind="thank_you"
        )
    )

    return ReviewManifest(
        client_name=schema.client_name,
        template_key=template_key,
        slide_count=slide_count,
        slides=slides,
    )


def _serialize(prs) -> bytes:
    """Save the deck, renumbering slide parts first.

    A5 clones arrive carrying their source partnames, so an un-renumbered save
    emits two zip entries under one ``ppt/slides/slideNN.xml``. ``build()`` fixes
    that with ``_renumber_slide_parts`` — but only after packaging (§8), which
    would leave draft.pptx itself malformed for B3/B4.
    """
    prs.part.rename_slide_parts(
        [sld_id.get(qn("r:id")) for sld_id in prs.slides._sldIdLst]
    )
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _report_json(report: object) -> str:
    """Serialize a B3/B4 report. Accepts JSON text, a dict, or an object with to_json()."""
    payload = report.to_json() if hasattr(report, "to_json") else report
    if isinstance(payload, (str, bytes)):
        return payload.decode() if isinstance(payload, bytes) else payload
    return json.dumps(payload, indent=2, default=str)


@dataclass(frozen=True)
class ReviewPackage:
    """On-disk review package handed to B3 (deterministic QA) and B4 (Cursor)."""

    root: Path
    draft_path: Path
    manifest_path: Path
    slides_dir: Path
    schema_path: Path
    manifest: ReviewManifest
    pptx_bytes: bytes
    slide_pngs: tuple[Path, ...] = ()

    def write_deterministic_report(self, report: object) -> Path:
        """Write qa_deterministic.json (B3)."""
        return self._write_report(DETERMINISTIC_REPORT_NAME, report)

    def write_cursor_report(self, report: object) -> Path:
        """Write qa_cursor.json (B4)."""
        return self._write_report(CURSOR_REPORT_NAME, report)

    def _write_report(self, name: str, report: object) -> Path:
        path = self.root / name
        path.write_text(_report_json(report), encoding="utf-8")
        return path


def build_review_package(
    prs,
    schema: DeckSchema,
    *,
    plan: list[PitchItem] | None = None,
    gtm_map: GtmProductMap | None = None,
    output_dir: Path | str | None = None,
    template_key: str = FORTUNEAI_TEMPLATE_BASENAME,
    dpi: int = RENDER_DPI,
) -> ReviewPackage:
    """Write draft.pptx, deck_schema.json, manifest.json and slide PNGs for QA.

    ``prs`` is the post-C2 presentation. Pass ``plan`` when the caller already has
    one (``assemble_skeleton``), otherwise supply ``gtm_map`` and it is computed
    with :func:`plan_pitch_sequence`. ``output_dir`` defaults to a fresh temp
    directory that is left on disk for the caller.
    """
    if plan is None:
        if gtm_map is None:
            raise ValueError(
                "build_review_package needs either plan= or gtm_map= to order the "
                "pitch section"
            )
        plan = plan_pitch_sequence(schema, gtm_map)

    manifest = build_manifest(
        schema, plan, slide_count=len(prs.slides), template_key=template_key
    )

    root = (
        Path(tempfile.mkdtemp(prefix="review-package-"))
        if output_dir is None
        else Path(output_dir)
    )
    slides_dir = root / SLIDES_DIRNAME
    slides_dir.mkdir(parents=True, exist_ok=True)

    pptx_bytes = _serialize(prs)
    draft_path = root / DRAFT_NAME
    draft_path.write_bytes(pptx_bytes)

    manifest_path = root / MANIFEST_NAME
    manifest_path.write_text(
        json.dumps(manifest.model_dump(mode="json", exclude_none=True), indent=2),
        encoding="utf-8",
    )
    schema_path = root / DECK_SCHEMA_NAME
    schema_path.write_text(
        json.dumps(schema.model_dump(mode="json"), indent=2), encoding="utf-8"
    )

    # render_slides already names outputs slide-NNN.png; None renders every slide.
    pngs = render_slides(draft_path, None, output_dir=slides_dir, dpi=dpi)

    logger.info(
        "review package: %d slides, %d PNG(s) in %s", manifest.slide_count, len(pngs), root
    )
    return ReviewPackage(
        root=root,
        draft_path=draft_path,
        manifest_path=manifest_path,
        slides_dir=slides_dir,
        schema_path=schema_path,
        manifest=manifest,
        pptx_bytes=pptx_bytes,
        slide_pngs=tuple(pngs),
    )


def _main(argv: list[str] | None = None) -> int:
    """CLI: ``python -m ingestion.review_package`` — package an existing post-C2 deck."""
    import argparse

    from pptx import Presentation

    parser = argparse.ArgumentParser(
        description="Build a B2 review package (draft.pptx + PNGs + manifest.json)."
    )
    parser.add_argument("pptx", type=Path, help="Post-C2 draft .pptx")
    parser.add_argument(
        "--schema", type=Path, required=True, help="Serialized DeckSchema JSON"
    )
    parser.add_argument(
        "--gtm-xlsx",
        type=Path,
        required=True,
        help="Fortune_AITool_GTM_Database.xlsx (Product Tags) for the pitch plan",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        default=None,
        help="Package directory (default: a new temp dir)",
    )
    parser.add_argument("--dpi", type=int, default=RENDER_DPI)
    args = parser.parse_args(argv)

    schema = DeckSchema.model_validate_json(args.schema.read_text(encoding="utf-8"))
    gtm_map = GtmProductMap.from_xlsx_bytes(args.gtm_xlsx.read_bytes())
    package = build_review_package(
        Presentation(str(args.pptx)),
        schema,
        gtm_map=gtm_map,
        output_dir=args.output_dir,
        dpi=args.dpi,
    )
    print(package.root)
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
