"""Token replace and logo insert against a FortuneAI-shaped spine.

Primitives are unwired from assemble_skeleton; tests call them directly.
"""

from __future__ import annotations

import io

import pytest
from pptx import Presentation
from pptx.util import Inches

from ingestion.pptx_tools import (
    APOS,
    CLIENT_NAME_POSSESSIVE_TOKEN,
    CLIENT_NAME_TOKEN,
    LOGO_TOKEN,
    insert_logo,
    replace_token,
)
from tests.fortuneai_placeholder_fixture import (
    HISTORY_BODY,
    MINIMAL_PNG,
    WHY_FORTUNE_STOCK,
    build_fortuneai_fixture_prs,
    fortuneai_fixture_bytes,
)


def _slide_texts(slide) -> list[str]:
    texts: list[str] = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            texts.append(para.text)
    return texts


def _embedded_images(slide) -> list:
    images = []
    for shape in slide.shapes:
        try:
            images.append(shape.image)
        except AttributeError:
            continue
    return images


def test_fixture_has_19_slides_and_named_tokens():
    prs = build_fortuneai_fixture_prs()
    assert len(prs.slides) == 19
    intro = " ".join(_slide_texts(prs.slides[0]))
    assert "[TITLE]" in intro
    assert "[DATE]" in intro
    history = " ".join(_slide_texts(prs.slides[2]))
    assert CLIENT_NAME_TOKEN in history
    assert WHY_FORTUNE_STOCK in " ".join(_slide_texts(prs.slides[1]))
    thanks = " ".join(_slide_texts(prs.slides[18]))
    assert "[DATE]" in thanks
    assert "Thank you!" in thanks


def test_replace_token_swaps_history_client_name():
    prs = build_fortuneai_fixture_prs()
    history = prs.slides[2]
    hits = replace_token(history, CLIENT_NAME_TOKEN, "Acme Corp")
    blob = " ".join(_slide_texts(history))
    assert hits >= 1
    assert "Acme Corp" in blob
    assert CLIENT_NAME_TOKEN not in blob
    assert f"Acme Corp{APOS}s opportunity" in blob


def test_replace_token_leaves_why_fortune_unchanged():
    prs = build_fortuneai_fixture_prs()
    why = prs.slides[1]
    before = _slide_texts(why)
    hits = replace_token(why, CLIENT_NAME_TOKEN, "Acme Corp")
    assert hits == 0
    assert _slide_texts(why) == before
    assert WHY_FORTUNE_STOCK in " ".join(before)


def test_replace_token_leaves_unmatched_text_alone():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box.text_frame.paragraphs[0].add_run().text = "Fortune reaches 42 million people"
    hits = replace_token(slide, CLIENT_NAME_TOKEN, "Acme Corp")
    assert hits == 0
    assert "Fortune reaches 42 million people" in " ".join(_slide_texts(slide))


def test_replace_token_audience_possessive_curly_apostrophe():
    prs = build_fortuneai_fixture_prs()
    audience = prs.slides[4]  # 2-card page
    hits = replace_token(audience, CLIENT_NAME_POSSESSIVE_TOKEN, f"Acme Corp{APOS}s")
    blob = " ".join(_slide_texts(audience))
    assert hits >= 1
    assert CLIENT_NAME_POSSESSIVE_TOKEN not in blob
    assert f"OVERDELIVERS Acme Corp{APOS}s TARGET" in blob


def test_replace_token_empty_token_raises():
    prs = build_fortuneai_fixture_prs()
    with pytest.raises(ValueError, match="non-empty"):
        replace_token(prs.slides[0], "", "x")


def test_replace_token_literal_product_type():
    prs = build_fortuneai_fixture_prs()
    program = prs.slides[9]  # 2-box program
    hits = replace_token(program, "PRODUCT TYPE", "Editorial Alignment")
    blob = " ".join(_slide_texts(program))
    assert hits == 2
    assert "PRODUCT TYPE" not in blob
    assert blob.count("Editorial Alignment") == 2


def test_insert_logo_lands_on_intro_and_thanks():
    prs = build_fortuneai_fixture_prs()
    insert_logo(prs.slides[0], MINIMAL_PNG)
    insert_logo(prs.slides[18], MINIMAL_PNG)
    intro_imgs = _embedded_images(prs.slides[0])
    thanks_imgs = _embedded_images(prs.slides[18])
    assert len(intro_imgs) == 1
    assert len(thanks_imgs) == 1
    assert intro_imgs[0].blob.startswith(b"\x89PNG")
    assert thanks_imgs[0].blob.startswith(b"\x89PNG")
    assert LOGO_TOKEN not in " ".join(_slide_texts(prs.slides[0]))
    assert LOGO_TOKEN not in " ".join(_slide_texts(prs.slides[18]))


def test_insert_logo_missing_placeholder_fails_loud():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank, no picture ph
    with pytest.raises(ValueError, match="logo placeholder not found"):
        insert_logo(slide, MINIMAL_PNG)


def test_insert_logo_garbage_bytes_fail_loud():
    prs = build_fortuneai_fixture_prs()
    with pytest.raises(ValueError, match="unreadable"):
        insert_logo(prs.slides[0], b"not-an-image")
    with pytest.raises(ValueError, match="empty"):
        insert_logo(prs.slides[0], b"")


def test_fortuneai_fixture_bytes_roundtrip():
    data = fortuneai_fixture_bytes()
    prs = Presentation(io.BytesIO(data))
    assert len(prs.slides) == 19
    assert CLIENT_NAME_TOKEN in HISTORY_BODY
