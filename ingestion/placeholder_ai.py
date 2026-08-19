"""Bounded Claude fills for FortuneAI named slots (C2 Chunk 5).

One small call per slot — not the corpus-clone JSON arc. Validates word counts,
no em dashes, and Fortune closing on Opportunity body; retries once then fails loud.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    import anthropic

    from ingestion.schema import DeckSchema

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "claude-sonnet-4-6"

EM_DASH = "\u2014"
EM_DASH_ASCII = "--"

TITLE_WORD_MIN = 3
TITLE_WORD_MAX = 6
BODY_WORD_MIN = 75
BODY_WORD_MAX = 95
PROGRAM_BLURB_WORD_MIN = 10
PROGRAM_BLURB_WORD_MAX = 15

_SHARED_RULES = """\
RULES (strict):
- Titles and headers: 3–6 words only.
- No em dashes (— or --) anywhere in your answer.
- Return ONLY the copy for this slot — no quotes, labels, or markdown.
"""

_INTRO_RULES = """\
INTRO TITLE:
- 3–6 words.
- ALL CAPS.
- Client-facing deck title; do not mention RFP mechanics.
"""

_HEADER_RULES = """\
OPPORTUNITY HEADER:
- 3–6 words.
- Sentence case or title case (not ALL CAPS).
- Industry tension or ambition for this client.
"""

_OPPORTUNITY_BODY_RULES = """\
OPPORTUNITY BODY:
- About 85 words (75–95 acceptable).
- Structure: (1) industry tension, (2–3) this client's chance to lead,
  (4) closing sentence that STARTS with the word Fortune.
- Do not quote or paraphrase RFP boilerplate.
- No em dashes.
"""

_AUDIENCE_TITLE_RULES = """\
AUDIENCE TITLE:
- 3–6 words.
- Sentence case (not ALL CAPS).
- Summarize the target audience theme for this client.
"""

_PROGRAM_BLURB_RULES = """\
PROGRAM ONE-LINER:
- 10–15 words.
- One sentence describing how Fortune delivers this product category for the client.
- No em dashes.
"""


class ClaudeCaller(Protocol):
    def __call__(self, *, system: str, user: str, max_tokens: int = 512) -> str: ...


def _word_count(text: str) -> int:
    return len(text.split())


def _has_em_dash(text: str) -> bool:
    return EM_DASH in text or EM_DASH_ASCII in text


def _last_sentence(text: str) -> str:
    parts = re.split(r"(?<=[.!?])\s+", text.strip())
    return parts[-1] if parts else text.strip()


def validate_short_title(
    text: str,
    *,
    all_caps: bool = False,
    sentence_case: bool = False,
    slot: str,
) -> None:
    wc = _word_count(text)
    if wc < TITLE_WORD_MIN or wc > TITLE_WORD_MAX:
        raise ValueError(
            f"{slot} must be {TITLE_WORD_MIN}–{TITLE_WORD_MAX} words; got {wc}"
        )
    if _has_em_dash(text):
        raise ValueError(f"{slot} must not contain em dashes")
    if all_caps and text != text.upper():
        raise ValueError(f"{slot} must be ALL CAPS")
    if sentence_case and text == text.upper():
        raise ValueError(f"{slot} must be sentence case, not ALL CAPS")


def validate_opportunity_body(text: str) -> None:
    wc = _word_count(text)
    if wc < BODY_WORD_MIN or wc > BODY_WORD_MAX:
        raise ValueError(
            f"Opportunity body must be {BODY_WORD_MIN}–{BODY_WORD_MAX} words; got {wc}"
        )
    if _has_em_dash(text):
        raise ValueError("Opportunity body must not contain em dashes")
    closer = _last_sentence(text)
    if not closer.startswith("Fortune"):
        raise ValueError(
            "Opportunity body closing sentence must start with Fortune"
        )


def validate_program_blurb(text: str) -> None:
    wc = _word_count(text)
    if wc < PROGRAM_BLURB_WORD_MIN or wc > PROGRAM_BLURB_WORD_MAX:
        raise ValueError(
            f"Program blurb must be {PROGRAM_BLURB_WORD_MIN}–{PROGRAM_BLURB_WORD_MAX} "
            f"words; got {wc}"
        )
    if _has_em_dash(text):
        raise ValueError("Program blurb must not contain em dashes")


def _brief_context(schema: DeckSchema) -> str:
    products = ", ".join(
        f"{p.name} ({p.category}, {p.cadence})" for p in schema.confirmed_products
    )
    kpis = ", ".join(schema.kpis)
    return (
        f"Company: {schema.company_name}\n"
        f"Industry: {schema.industry}\n"
        f"Campaign goal: {schema.campaign_goal}\n"
        f"Campaign narrative: {schema.campaign_narrative}\n"
        f"Targeting: {schema.targeting_details}\n"
        f"KPIs: {kpis}\n"
        f"KPI details: {schema.kpi_details}\n"
        f"Confirmed products: {products}\n"
    )


def _call_validated(
    caller: ClaudeCaller,
    *,
    system: str,
    user: str,
    validate,
    slot: str,
    max_tokens: int = 512,
) -> str:
    last_error: ValueError | None = None
    prompt = user
    for attempt in range(2):
        raw = caller(system=system, user=prompt, max_tokens=max_tokens)
        text = raw.strip()
        try:
            validate(text)
            return text
        except ValueError as exc:
            last_error = exc
            logger.warning("%s validation failed (attempt %d): %s", slot, attempt + 1, exc)
            prompt = (
                f"{user}\n\nYour previous answer failed validation: {exc}. "
                "Follow the rules exactly and try again."
            )
    raise ValueError(f"{slot} AI fill failed after retry: {last_error}")


class PlaceholderAI:
    """Bounded Claude fills for FortuneAI named slots."""

    def __init__(self, caller: ClaudeCaller) -> None:
        self._caller = caller

    @classmethod
    def from_anthropic(
        cls,
        client: anthropic.Anthropic,
        *,
        model: str = DEFAULT_MODEL,
    ) -> PlaceholderAI:
        def caller(*, system: str, user: str, max_tokens: int = 512) -> str:
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            )
            block = response.content[0]
            if not hasattr(block, "text"):
                raise ValueError(f"Unexpected Claude response block: {type(block)}")
            return block.text.strip()

        return cls(caller)

    def intro_title(self, schema: DeckSchema) -> str:
        system = f"You write Fortune pitch deck intro titles.\n\n{_SHARED_RULES}\n\n{_INTRO_RULES}"
        user = f"{_brief_context(schema)}\nWrite the intro [TITLE] (3–6 words, ALL CAPS)."
        return _call_validated(
            self._caller,
            system=system,
            user=user,
            validate=lambda t: validate_short_title(
                t, all_caps=True, slot="Intro title"
            ),
            slot="Intro title",
            max_tokens=64,
        )

    def opportunity_header(self, schema: DeckSchema) -> str:
        system = (
            f"You write Fortune pitch deck opportunity headers.\n\n"
            f"{_SHARED_RULES}\n\n{_HEADER_RULES}"
        )
        user = (
            f"{_brief_context(schema)}\nWrite the Opportunity [HEADER] (3–6 words)."
        )
        return _call_validated(
            self._caller,
            system=system,
            user=user,
            validate=lambda t: validate_short_title(t, slot="Opportunity header"),
            slot="Opportunity header",
            max_tokens=64,
        )

    def opportunity_body(self, schema: DeckSchema) -> str:
        system = (
            f"You write Fortune pitch deck opportunity body copy.\n\n"
            f"{_SHARED_RULES}\n\n{_OPPORTUNITY_BODY_RULES}"
        )
        user = (
            f"{_brief_context(schema)}\n"
            "Write the Opportunity [BODY] (~85 words). "
            "End with a sentence that starts with Fortune."
        )
        return _call_validated(
            self._caller,
            system=system,
            user=user,
            validate=validate_opportunity_body,
            slot="Opportunity body",
            max_tokens=512,
        )

    def audience_title(self, schema: DeckSchema, segments: list[str]) -> str:
        segment_line = ", ".join(segments)
        system = (
            f"You write Fortune pitch deck audience titles.\n\n"
            f"{_SHARED_RULES}\n\n{_AUDIENCE_TITLE_RULES}"
        )
        user = (
            f"{_brief_context(schema)}\n"
            f"Matched audience segments: {segment_line}\n"
            "Write the [AUDIENCE TITLE] (3–6 words, sentence case)."
        )
        return _call_validated(
            self._caller,
            system=system,
            user=user,
            validate=lambda t: validate_short_title(
                t, sentence_case=True, slot="Audience title"
            ),
            slot="Audience title",
            max_tokens=64,
        )

    def program_blurb(self, schema: DeckSchema, category_name: str) -> str:
        system = (
            f"You write Fortune pitch deck program one-liners.\n\n"
            f"{_SHARED_RULES}\n\n{_PROGRAM_BLURB_RULES}"
        )
        user = (
            f"{_brief_context(schema)}\n"
            f"Product category on this program box: {category_name}\n"
            "Write one Product description. blurb (10–15 words)."
        )
        return _call_validated(
            self._caller,
            system=system,
            user=user,
            validate=validate_program_blurb,
            slot=f"Program blurb ({category_name})",
            max_tokens=128,
        )
