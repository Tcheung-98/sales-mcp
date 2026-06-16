import io
from unittest.mock import MagicMock, patch

from docx import Document
from pptx import Presentation
from pptx.util import Inches

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
    return generator


def _slide_with_textbox(text: str):
    """Minimal slide with a text box containing the given text as a single run."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    txBox.text_frame.paragraphs[0].add_run().text = text
    return slide


def test_apply_replacements_swaps_client_name():
    slide = _slide_with_textbox("your brand reaches millions")
    DeckGenerator._apply_replacements(slide, {"client_name": "Acme Corp"})
    texts = [
        run.text
        for sh in slide.shapes if sh.has_text_frame
        for para in sh.text_frame.paragraphs
        for run in para.runs
    ]
    assert any("Acme Corp" in t for t in texts)
    assert not any("your brand" in t.lower() for t in texts)


def test_apply_replacements_case_insensitive():
    slide = _slide_with_textbox("YOUR COMPANY is the leader")
    DeckGenerator._apply_replacements(slide, {"client_name": "Acme Corp"})
    texts = [
        run.text
        for sh in slide.shapes if sh.has_text_frame
        for para in sh.text_frame.paragraphs
        for run in para.runs
    ]
    assert any("Acme Corp" in t for t in texts)


def test_apply_replacements_title_overwrites_placeholder():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title, Content
    slide.shapes.title.text = "Old Title"
    DeckGenerator._apply_replacements(slide, {"title": "New Title"})
    assert slide.shapes.title.text == "New Title"


def test_apply_replacements_leaves_unmatched_text_alone():
    slide = _slide_with_textbox("Fortune reaches 42 million people")
    DeckGenerator._apply_replacements(slide, {"client_name": "Acme Corp"})
    texts = [
        run.text
        for sh in slide.shapes if sh.has_text_frame
        for para in sh.text_frame.paragraphs
        for run in para.runs
    ]
    assert any("Fortune reaches 42 million people" in t for t in texts)


def test_apply_replacements_noop_on_empty_dict():
    slide = _slide_with_textbox("your brand")
    DeckGenerator._apply_replacements(slide, {})
    texts = [
        run.text
        for sh in slide.shapes if sh.has_text_frame
        for para in sh.text_frame.paragraphs
        for run in para.runs
    ]
    assert any("your brand" in t for t in texts)


def _rulebook_bytes(text: str = "Rule 1: always outline first.\nRule 2: $750K escalates.") -> bytes:
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def test_load_rulebook_extracts_text():
    generator = _build_generator()
    generator._s3.get_object.return_value = {"Body": io.BytesIO(_rulebook_bytes())}

    text = generator._load_rulebook()

    assert "always outline first" in text
    assert "$750K escalates" in text
    generator._s3.get_object.assert_called_once()


def test_load_rulebook_caches():
    generator = _build_generator()
    generator._s3.get_object.return_value = {"Body": io.BytesIO(_rulebook_bytes())}

    generator._load_rulebook()
    generator._load_rulebook()

    generator._s3.get_object.assert_called_once()


def test_load_pptx_caches_by_key():
    generator = _build_generator()
    generator._s3.get_object.return_value = {"Body": io.BytesIO(_blank_bytes())}

    prs_a1 = generator._load_pptx("corpus/deck_a.pptx")
    prs_a2 = generator._load_pptx("corpus/deck_a.pptx")

    assert prs_a1 is prs_a2
    generator._s3.get_object.assert_called_once()




