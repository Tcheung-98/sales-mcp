"""FortuneAI placeholder fills (C2).

Wired from DeckGenerator.build after assemble_skeleton: deterministic fills
first, then bounded Claude slots when ``ai`` is provided. Never rewrites A5
product clones or Why Fortune stock copy.
"""

from __future__ import annotations

import logging
import re
from datetime import date
from typing import TYPE_CHECKING
from urllib.parse import urlparse

import requests

from ingestion.audience_data import AudienceData, AudienceRow, rank_segments_by_index
from ingestion.category_dividers import CATEGORY_DIVIDERS, divider_index_for_category
from ingestion.placeholders import (
    delete_unused_variants,
    select_audience_variant,
    select_program_variant,
)
from ingestion.pptx_tools import (
    APOS,
    CLIENT_NAME_POSSESSIVE_TOKEN,
    CLIENT_NAME_TOKEN,
    clone_shape_below,
    insert_logo,
    iter_shapes,
    replace_first_token,
    replace_token,
)
from ingestion.schema import DeckSchema, Product

if TYPE_CHECKING:
    from ingestion.placeholder_ai import PlaceholderAI

logger = logging.getLogger(__name__)

TITLE_TOKEN = "[TITLE]"
HEADER_TOKEN = "[HEADER]"
BODY_TOKEN = "[BODY]"
AUDIENCE_TITLE_TOKEN = "[AUDIENCE TITLE]"
DATE_TOKEN = "[DATE]"
AUDIENCE_SEGMENT_TOKEN = "[AUDIENCE SEGMENT]"
REACH_TOKEN = "[REACH]"
INDEX_TOKEN = "[INDEX]"
BUDGET_TOKEN = "[BUDGET]"
PRODUCT_CATEGORY_TOKEN = "[PRODUCT CATEGORY]"
PRODUCT_TYPE_LITERAL = "PRODUCT TYPE"
PRODUCT_DESCRIPTION_LITERAL = "Product description."
CLIENT_NAME_UPPER_TOKEN = "[CLIENT NAME]"
EM_DASH = "\u2014"
_TOTAL_LABEL = re.compile(r"\btotal\b", re.I)
_INVESTMENT_CLONE_GAP_EMU = 200_000

# 1-category Program Overview: second box is stock, not a second AI blurb.
PROGRAM_STOCK_TYPE = "Fortune"
PROGRAM_STOCK_BLURB = (
    "Fortune reaches leaders across editorial, live events, and premium brand environments."
)


def format_usd(amount: float) -> str:
    if float(amount).is_integer():
        return f"${amount:,.0f}"
    return f"${amount:,.2f}"


def format_presentation_date(as_of: date | None = None) -> str:
    """Workflow generate-time Month Year (not flight start, not intake date)."""
    return (as_of or date.today()).strftime("%B %Y")


def mix_total(schema: DeckSchema) -> float:
    return sum(product.price for product in schema.confirmed_products)


def stated_total_budget(schema: DeckSchema) -> float:
    """Seller stated total: a tier labeled with the word total, else max of budgets[].amount."""
    labeled = [
        tier for tier in schema.budgets if tier.label and _TOTAL_LABEL.search(tier.label)
    ]
    if labeled:
        return max(tier.amount for tier in labeled)
    return max(tier.amount for tier in schema.budgets)


def fetch_logo_bytes(client_logo: str) -> bytes:
    """Download the intake logo. HTTPS only; fail loud on empty/unreadable URLs."""
    url = (client_logo or "").strip()
    if not url:
        raise ValueError("client_logo is missing")
    parsed = urlparse(url)
    if parsed.scheme != "https":
        raise ValueError(
            "client_logo must be an HTTPS URL "
            f"(got {parsed.scheme or 'empty'!r}). "
            "SharePoint library paths are not fetched in this build."
        )
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as exc:
        raise ValueError(f"Failed to download client_logo: {exc}") from exc
    if not resp.content:
        raise ValueError("client_logo download was empty")
    return resp.content


def funded_divider_buckets(schema: DeckSchema) -> list[tuple[str, list[Product]]]:
    """Funded C1 dividers in Workflow order, each with its confirmed products."""
    buckets: list[list[Product]] = [[] for _ in CATEGORY_DIVIDERS]
    for product in schema.confirmed_products:
        buckets[divider_index_for_category(product.category)].append(product)
    return [
        (CATEGORY_DIVIDERS[i].name, products)
        for i, products in enumerate(buckets)
        if products
    ]


def apply_placeholders(
    prs,
    schema: DeckSchema,
    *,
    audience: AudienceData,
    logo_bytes: bytes,
    as_of: date | None = None,
    ai: PlaceholderAI | None = None,
) -> list[str]:
    """Fill stock FortuneAI slots on an assembled skeleton. Returns warnings."""
    warnings: list[str] = []
    date_text = format_presentation_date(as_of)
    company = schema.company_name
    possessive = f"{company}{APOS}s"
    buckets = funded_divider_buckets(schema)
    divider_names = [name for name, _ in buckets]

    matched = audience.match_targeting(schema.targeting_details)
    audience_choice = select_audience_variant(len(matched))
    rows = list(matched)
    if audience_choice.warning:
        warnings.append(audience_choice.warning)
        rows = rank_segments_by_index(matched)[: audience_choice.size]

    program_choice = select_program_variant(len(divider_names))
    program_names = divider_names
    if program_choice.warning:
        warnings.append(program_choice.warning)
        program_names = divider_names[: program_choice.size]

    _assert_budget_matches_mix(schema)

    intro = prs.slides[0]
    _require_replace(intro, DATE_TOKEN, date_text, what="Intro")
    insert_logo(intro, logo_bytes)

    history = prs.slides[2]
    _fill_history(history, company)

    audience_slide = prs.slides[audience_choice.keep_index]
    _fill_audience_slide(audience_slide, possessive, rows)

    program_slide = prs.slides[program_choice.keep_index]
    _fill_program_slide(program_slide, program_names)

    investment = prs.slides[-2]
    _fill_investment_slide(investment, schema, buckets)

    thanks = prs.slides[-1]
    _require_replace(thanks, DATE_TOKEN, date_text, what="Thank You")
    insert_logo(thanks, logo_bytes)

    if ai is not None:
        _fill_ai_placeholders(
            prs,
            schema,
            ai,
            audience_idx=audience_choice.keep_index,
            program_idx=program_choice.keep_index,
            audience_segments=[row.segment for row in rows],
            program_categories=program_names,
            funded_category_count=len(divider_names),
        )

    delete_unused_variants(
        prs,
        audience_keep=audience_choice.keep_index,
        program_keep=program_choice.keep_index,
    )
    logger.info("applied placeholders; %d warning(s)", len(warnings))
    return warnings


def _fill_ai_placeholders(
    prs,
    schema: DeckSchema,
    ai: PlaceholderAI,
    *,
    audience_idx: int,
    program_idx: int,
    audience_segments: list[str],
    program_categories: list[str],
    funded_category_count: int,
) -> None:
    """Bounded Claude fills on role-tagged spine slides only."""
    intro = prs.slides[0]
    _require_replace(
        intro, TITLE_TOKEN, ai.intro_title(schema), what="Intro"
    )

    opp = prs.slides[3]
    _require_replace(
        opp, HEADER_TOKEN, ai.opportunity_header(schema), what="Opportunity"
    )
    _require_replace(
        opp, BODY_TOKEN, ai.opportunity_body(schema), what="Opportunity"
    )

    audience = prs.slides[audience_idx]
    _require_replace(
        audience,
        AUDIENCE_TITLE_TOKEN,
        ai.audience_title(schema, audience_segments),
        what="Audience",
    )

    program = prs.slides[program_idx]
    _fill_program_ai_blurbs(
        program,
        schema,
        ai,
        program_categories,
        funded_category_count=funded_category_count,
    )


def _fill_program_ai_blurbs(
    slide,
    schema: DeckSchema,
    ai: PlaceholderAI,
    categories: list[str],
    *,
    funded_category_count: int,
) -> None:
    desc_shapes = [
        shape
        for shape in iter_shapes(slide)
        if PRODUCT_DESCRIPTION_LITERAL in _shape_text(shape)
    ]
    if not desc_shapes:
        raise ValueError("Program Overview has no Product description. boxes")

    if funded_category_count == 1:
        blurb = ai.program_blurb(schema, categories[0])
        _set_paragraph_text(desc_shapes[0].text_frame.paragraphs[0], blurb)
        return

    if len(desc_shapes) < len(categories):
        raise ValueError(
            "Program Overview needs "
            f"{len(categories)} Product description. boxes; "
            f"found {len(desc_shapes)}"
        )
    for shape, category in zip(desc_shapes[: len(categories)], categories, strict=True):
        blurb = ai.program_blurb(schema, category)
        _set_paragraph_text(shape.text_frame.paragraphs[0], blurb)


def _assert_budget_matches_mix(schema: DeckSchema) -> None:
    stated = stated_total_budget(schema)
    mix = mix_total(schema)
    if abs(stated - mix) > 0.005:
        raise ValueError(
            f"Stated total budget {format_usd(stated)} does not match mix total "
            f"{format_usd(mix)}. Prices were not changed."
        )


def _require_replace(slide, token: str, text: str, *, what: str) -> None:
    if replace_token(slide, token, text) < 1:
        raise ValueError(f"{what} slide is missing token {token!r}")


def _require_first(slide, token: str, text: str, *, what: str) -> None:
    if replace_first_token(slide, token, text) < 1:
        raise ValueError(f"{what}: token {token!r} not found")


def _fill_history(slide, company: str) -> None:
    hits = replace_token(slide, CLIENT_NAME_TOKEN, company)
    hits += replace_token(slide, CLIENT_NAME_UPPER_TOKEN, company)
    if hits < 1:
        raise ValueError("History of Trust slide is missing [client name]")


def _fill_audience_slide(
    slide,
    possessive: str,
    rows: list[AudienceRow],
) -> None:
    if replace_token(slide, CLIENT_NAME_POSSESSIVE_TOKEN, possessive) < 1:
        raise ValueError("Audience slide is missing [CLIENT NAME'S]")
    for row in rows:
        _require_first(
            slide, AUDIENCE_SEGMENT_TOKEN, row.segment, what="Audience card"
        )
        _require_first(slide, REACH_TOKEN, row.reach, what="Audience card")
        _require_first(slide, INDEX_TOKEN, row.index, what="Audience card")


def _fill_program_slide(slide, divider_names: list[str]) -> None:
    if not divider_names:
        raise ValueError("Program Overview has no funded categories")
    _require_first(
        slide, PRODUCT_TYPE_LITERAL, divider_names[0], what="Program Overview"
    )
    if len(divider_names) == 1:
        _require_first(
            slide, PRODUCT_TYPE_LITERAL, PROGRAM_STOCK_TYPE, what="Program Overview"
        )
        desc_shapes = [
            shape
            for shape in iter_shapes(slide)
            if PRODUCT_DESCRIPTION_LITERAL in _shape_text(shape)
        ]
        if len(desc_shapes) < 2:
            raise ValueError(
                "Program Overview 1-category page needs a second "
                "Product description. box for the stock Fortune sentence"
            )
        _set_paragraph_text(
            desc_shapes[1].text_frame.paragraphs[0], PROGRAM_STOCK_BLURB
        )
        return
    for name in divider_names[1:]:
        _require_first(slide, PRODUCT_TYPE_LITERAL, name, what="Program Overview")


def _shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text or "" for p in shape.text_frame.paragraphs)


def _find_category_box(slide):
    for shape in iter_shapes(slide):
        if PRODUCT_CATEGORY_TOKEN in _shape_text(shape):
            return shape
    return None


def _set_paragraph_text(para, text: str) -> None:
    runs = para.runs
    if runs:
        runs[0].text = text
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        para.text = text


def _fill_category_box(shape, category: str, products: list[Product]) -> None:
    lines = [f"{product.name} {EM_DASH} {format_usd(product.price)}" for product in products]
    tf = shape.text_frame
    paras = list(tf.paragraphs)
    if not paras:
        raise ValueError("Investment category box has no paragraphs")
    _set_paragraph_text(paras[0], category)
    for i, line in enumerate(lines):
        idx = i + 1
        if idx < len(paras):
            _set_paragraph_text(paras[idx], line)
        else:
            para = tf.add_paragraph()
            para.text = line
    leftover_start = 1 + len(lines)
    # Re-read; add_paragraph may have grown the list.
    paras = list(tf.paragraphs)
    for para in paras[leftover_start:]:
        _set_paragraph_text(para, "")


def _fill_investment_slide(
    slide,
    schema: DeckSchema,
    buckets: list[tuple[str, list[Product]]],
) -> None:
    _require_replace(slide, BUDGET_TOKEN, format_usd(mix_total(schema)), what="Investment")
    template = _find_category_box(slide)
    if template is None:
        raise ValueError("Investment slide is missing [PRODUCT CATEGORY]")
    boxes = [template]
    height = int(template.height)
    try:
        for i in range(1, len(buckets)):
            boxes.append(
                clone_shape_below(
                    slide, template, i * (height + _INVESTMENT_CLONE_GAP_EMU)
                )
            )
    except ValueError as exc:
        raise ValueError(f"Investment category box clone failed: {exc}") from exc
    if len(boxes) != len(buckets):
        raise ValueError(
            "Investment category box clone failed: "
            f"needed {len(buckets)} boxes, got {len(boxes)}"
        )
    for box, (category, products) in zip(boxes, buckets, strict=True):
        _fill_category_box(box, category, products)
