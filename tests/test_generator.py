import io
from unittest.mock import MagicMock, patch

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from ingestion.generator import DeckGenerator


def _seed_bytes(slide_count: int = 1) -> bytes:
    prs = Presentation()
    for idx in range(slide_count):
        layout = prs.slide_layouts[0] if idx == 0 else prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        if slide.shapes.title:
            slide.shapes.title.text = f"Seed {idx}"
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _build_generator() -> DeckGenerator:
    with patch("boto3.client", return_value=MagicMock()):
        generator = DeckGenerator(bucket="test-bucket", seed_key="seed.pptx")
    generator._s3 = MagicMock()
    generator._seed_bytes = _seed_bytes(slide_count=2)
    return generator


def test_build_pptx_replaces_seed_slides_and_sets_title_fields():
    generator = _build_generator()
    slides = [
        {"slide_type": "title", "title": "Client Growth Plan", "subtitle": "Q3 2026"},
        {"slide_type": "content", "title": "Why Fortune", "bullets": ["Trusted authority"]},
    ]

    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))

    assert len(prs.slides) == 2

    title_slide = prs.slides[0]
    text_placeholders = {
        ph.placeholder_format.idx: ph.text_frame.text
        for ph in title_slide.placeholders
        if ph.has_text_frame
    }
    assert text_placeholders.get(0) == "Client Growth Plan"
    assert text_placeholders.get(1) == "Q3 2026"


def test_build_pptx_formats_content_slide_bullets_and_branding():
    generator = _build_generator()
    bullets = [
        "Finance decision-makers at scale",
        "Contextual programs around C-suite leadership",
        "Measured outcomes in premium business environments",
    ]
    slides = [{"slide_type": "content", "title": "Recommended Products", "bullets": bullets}]

    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide = prs.slides[0]

    assert slide.background.fill.fore_color.rgb == RGBColor(0x10, 0x18, 0x5F)
    assert slide.shapes.title.text == "Recommended Products"

    body = next(
        ph for ph in slide.placeholders if ph.has_text_frame and ph.placeholder_format.idx == 1
    )
    body_text = [p.text for p in body.text_frame.paragraphs if p.text.strip()]
    assert body_text == bullets

    for paragraph in body.text_frame.paragraphs:
        if not paragraph.text.strip():
            continue
        assert paragraph.font.color.rgb == RGBColor(0xFF, 0xFF, 0xFF)
        p_pr = paragraph._p.get_or_add_pPr()
        bullet = p_pr.find(qn("a:buChar"))
        assert bullet is not None
        assert bullet.get("char") == "—"


def test_build_pptx_prunes_non_content_placeholders():
    generator = _build_generator()
    slides = [{"slide_type": "content", "title": "Next Steps", "bullets": ["Finalize scope"]}]

    pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide = prs.slides[0]

    placeholder_indexes = {ph.placeholder_format.idx for ph in slide.placeholders}
    assert placeholder_indexes <= {0, 1}
