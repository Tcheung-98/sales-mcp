import io
from unittest.mock import MagicMock, patch

from pptx import Presentation

from ingestion.generator import DeckGenerator


def _blank_bytes(slide_count: int = 1) -> bytes:
    """Minimal blank template with enough slides for _clone_slide(source, 0, ...)."""
    prs = Presentation()
    for _ in range(slide_count):
        prs.slides.add_slide(prs.slide_layouts[1])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _build_generator() -> DeckGenerator:
    with patch("boto3.client", return_value=MagicMock()):
        generator = DeckGenerator(bucket="test-bucket")
    generator._s3 = MagicMock()
    generator._blank_bytes = _blank_bytes(slide_count=6)
    return generator


def test_build_pptx_correct_slide_count():
    generator = _build_generator()
    slides = [
        {"slide_type": "title", "title": "Client Growth Plan", "subtitle": "Q3 2026"},
        {"slide_type": "product", "title": "Fortune.com", "bullets": ["Reaches 18M uniques"]},
    ]
    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    assert len(prs.slides) == 2


def test_build_pptx_title_slide_sets_headline():
    generator = _build_generator()
    slides = [{"slide_type": "title", "title": "Client Growth Plan", "subtitle": "Q3 2026"}]
    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide = prs.slides[0]
    all_text = {
        ph.placeholder_format.idx: ph.text_frame.text
        for ph in slide.placeholders
        if ph.has_text_frame
    }
    assert "Client Growth Plan" in all_text.values()
    assert "Q3 2026" in all_text.values()


def test_build_pptx_content_slide_populates_title_and_bullets():
    generator = _build_generator()
    bullets = [
        "Finance decision-makers at scale",
        "Contextual programs around C-suite leadership",
        "Measured outcomes in premium business environments",
    ]
    slides = [{"slide_type": "product", "title": "Recommended Products", "bullets": bullets}]
    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide = prs.slides[0]

    assert slide.shapes.title.text == "Recommended Products"

    body = next(
        ph for ph in slide.placeholders
        if ph.has_text_frame and ph.placeholder_format.idx == 1
    )
    body_text = [p.text for p in body.text_frame.paragraphs if p.text.strip()]
    assert body_text == bullets



def test_build_pptx_multiple_content_types():
    generator = _build_generator()
    slides = [
        {"slide_type": "title", "title": "T", "subtitle": "S"},
        {"slide_type": "product", "title": "P", "bullets": ["B"]},
        {"slide_type": "proof", "title": "Proof", "bullets": ["P"]},
        {"slide_type": "investment", "title": "Inv", "bullets": ["I"]},
        {"slide_type": "next_steps", "title": "Next", "bullets": ["N"]},
    ]
    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    assert len(prs.slides) == 5
