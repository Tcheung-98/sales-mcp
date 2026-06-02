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


def test_load_pptx_caches_by_key():
    generator = _build_generator()
    generator._s3.get_object.return_value = {"Body": io.BytesIO(_blank_bytes())}

    prs_a1 = generator._load_pptx("corpus/deck_a.pptx")
    prs_a2 = generator._load_pptx("corpus/deck_a.pptx")

    assert prs_a1 is prs_a2
    generator._s3.get_object.assert_called_once()


def _clone_slide_data(source_path: str, slide_number: int = 1, **replacements) -> dict:
    return {
        "action": "clone",
        "source_path": source_path,
        "slide_number": slide_number,
        "replacements": replacements,
    }


def test_build_pptx_correct_slide_count():
    generator = _build_generator()
    source_prs = Presentation(io.BytesIO(_blank_bytes()))
    with patch.object(generator, "_load_pptx", return_value=source_prs):
        slides = [
            {"action": "cover", "title": "Client Growth Plan", "subtitle": "Q3 2026"},
            _clone_slide_data("corpus/deck.pptx"),
        ]
        pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    assert len(prs.slides) == 2


def test_build_pptx_title_slide_sets_headline():
    generator = _build_generator()
    slides = [{"action": "cover", "title": "Client Growth Plan", "subtitle": "Q3 2026"}]
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


def test_build_pptx_clone_action_calls_load_pptx():
    generator = _build_generator()
    source_prs = Presentation(io.BytesIO(_blank_bytes()))
    with patch.object(generator, "_load_pptx", return_value=source_prs) as mock_load:
        slides = [
            {"action": "cover", "title": "T", "subtitle": "S"},
            _clone_slide_data("corpus/product.pptx", slide_number=1),
        ]
        pptx_bytes = generator._build_pptx(slides)
    mock_load.assert_called_once_with("corpus/product.pptx")
    assert len(Presentation(io.BytesIO(pptx_bytes)).slides) == 2


def test_build_pptx_multiple_clone_actions():
    generator = _build_generator()
    source_prs = Presentation(io.BytesIO(_blank_bytes()))
    with patch.object(generator, "_load_pptx", return_value=source_prs):
        slides = [
            {"action": "cover", "title": "T", "subtitle": "S"},
            _clone_slide_data("corpus/deck.pptx"),
            _clone_slide_data("corpus/deck.pptx"),
            _clone_slide_data("corpus/deck.pptx"),
            _clone_slide_data("corpus/deck.pptx"),
        ]
        pptx_bytes = generator._build_pptx(slides)
    prs = Presentation(io.BytesIO(pptx_bytes))
    assert len(prs.slides) == 5
