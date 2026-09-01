"""C2 audience / program variant selection.

Stock FortuneAI indices 0–11 stay valid through C1 (C1 only mutates from the
first divider at index 12). Identify variants by those indices, not layout names
(Chunk 0: Audience/Program layouts are not unique). ``build()`` calls
``delete_unused_variants`` after deterministic fills.
"""

from __future__ import annotations

from dataclasses import dataclass

from ingestion.pptx_tools import delete_slide

# 0-based slide index → card/box count (Workflow slides 5–9 and 10–12).
AUDIENCE_VARIANTS: tuple[tuple[int, int], ...] = (
    (4, 2),
    (5, 3),
    (6, 4),
    (7, 5),
    (8, 6),
)
PROGRAM_VARIANTS: tuple[tuple[int, int], ...] = (
    (9, 2),
    (10, 3),
    (11, 4),
)

AUDIENCE_INDICES: tuple[int, ...] = tuple(i for i, _ in AUDIENCE_VARIANTS)
PROGRAM_INDICES: tuple[int, ...] = tuple(i for i, _ in PROGRAM_VARIANTS)

_MIN_AUDIENCE_SEGMENTS = 2
_MAX_AUDIENCE_SEGMENTS = 6
_MIN_PROGRAM_CATEGORIES = 1
_MAX_PROGRAM_BOXES = 4


@dataclass(frozen=True)
class VariantChoice:
    """Which stock FortuneAI variant page to keep."""

    keep_index: int
    size: int
    warning: str | None = None


def select_audience_variant(segment_count: int) -> VariantChoice:
    """Pick the 2–6 card Audience page. ``<2`` fails; ``>6`` keeps 6-card + warning."""
    if segment_count < _MIN_AUDIENCE_SEGMENTS:
        raise ValueError(
            f"Audience slide needs at least {_MIN_AUDIENCE_SEGMENTS} segments; "
            f"got {segment_count}. Flag the seller to add targeting segments."
        )
    if segment_count > _MAX_AUDIENCE_SEGMENTS:
        return VariantChoice(
            keep_index=8,
            size=_MAX_AUDIENCE_SEGMENTS,
            warning=(
                f"{segment_count} audience segments exceeds "
                f"{_MAX_AUDIENCE_SEGMENTS}; using the 6-card page. Ask the "
                "seller to prioritize their top 6."
            ),
        )
    keep_index = 4 + (segment_count - 2)
    return VariantChoice(keep_index=keep_index, size=segment_count)


def select_program_variant(funded_category_count: int) -> VariantChoice:
    """Pick the 2–4 box Program Overview page.

    1 funded category → 2-box page (second box is a Fortune-wide capability).
    ``>4`` → 4-box page + warning (template has no 5-box layout).
    """
    if funded_category_count < _MIN_PROGRAM_CATEGORIES:
        raise ValueError(
            f"Program Overview needs at least {_MIN_PROGRAM_CATEGORIES} funded "
            f"category; got {funded_category_count}."
        )
    if funded_category_count == 1:
        return VariantChoice(keep_index=9, size=2)
    if funded_category_count > _MAX_PROGRAM_BOXES:
        return VariantChoice(
            keep_index=11,
            size=_MAX_PROGRAM_BOXES,
            warning=(
                f"{funded_category_count} funded categories exceeds "
                f"{_MAX_PROGRAM_BOXES} program boxes; using the 4-box page."
            ),
        )
    keep_index = 9 + (funded_category_count - 2)
    return VariantChoice(keep_index=keep_index, size=funded_category_count)


def unused_variant_indices(*, audience_keep: int, program_keep: int) -> list[int]:
    if audience_keep not in AUDIENCE_INDICES:
        raise ValueError(
            f"audience_keep {audience_keep} is not an Audience variant "
            f"index {list(AUDIENCE_INDICES)}"
        )
    if program_keep not in PROGRAM_INDICES:
        raise ValueError(
            f"program_keep {program_keep} is not a Program Overview variant "
            f"index {list(PROGRAM_INDICES)}"
        )
    unused = [i for i in AUDIENCE_INDICES if i != audience_keep]
    unused.extend(i for i in PROGRAM_INDICES if i != program_keep)
    return unused


def delete_unused_variants(
    prs, *, audience_keep: int, program_keep: int
) -> None:
    """Drop unused Audience and Program pages, back-to-front so indices stay valid."""
    unused = unused_variant_indices(
        audience_keep=audience_keep, program_keep=program_keep
    )
    for idx in sorted(unused, reverse=True):
        delete_slide(prs, idx)
