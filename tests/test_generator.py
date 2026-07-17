import io
from unittest.mock import MagicMock, patch

import anthropic
import pytest
from docx import Document
from pptx import Presentation
from pptx.util import Inches

from ingestion import generator as generator_module
from ingestion.generator import DeckGenerator
from ingestion.schema import DeckSchema, Product


def _blank_bytes(slide_count: int = 1) -> bytes:
    """Minimal blank template with enough slides for _clone_slide(source, 0, ...)."""
    prs = Presentation()
    for _ in range(slide_count):
        prs.slides.add_slide(prs.slide_layouts[1])
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _fortune_template_bytes() -> bytes:
    """Template with landmark + product slides matching patched layout constants."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[1])  # landmark
    prs.slides.add_slide(prs.slide_layouts[5])  # product slide
    prs.slides.add_slide(prs.slide_layouts[5])  # product slide
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


@pytest.fixture
def fortune_template_layouts(monkeypatch):
    monkeypatch.setattr(
        generator_module, "_PRODUCT_SECTION_LANDMARK", "Title and Content"
    )
    monkeypatch.setattr(generator_module, "_PRODUCT_LAYOUT", "Title Only")


def _presentation_with_n_slides(n: int) -> Presentation:
    prs = Presentation()
    for _ in range(n):
        prs.slides.add_slide(prs.slide_layouts[1])
    return prs


def _build_generator() -> DeckGenerator:
    with patch("boto3.client", return_value=MagicMock()):
        generator = DeckGenerator(bucket="test-bucket")
    generator._s3 = MagicMock()
    generator._blank_bytes = _blank_bytes(slide_count=6)
    return generator


def _build_schema(**overrides) -> DeckSchema:
    defaults = dict(
        client_name="Acme Corp",
        industry="Tech",
        budget_quarterly=100_000,
        confirmed_products=[
            Product(
                name="Fortune 500 List",
                cadence="annual",
                price=50_000,
                category="Newsletter",
            )
        ],
    )
    defaults.update(overrides)
    return DeckSchema(**defaults)


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


def test_call_claude_context_includes_full_coordinates():
    generator = _build_generator()
    generator._rulebook_text = "Rule: always lead with value."
    generator._api_key = "test-key"

    arc = [
        {"slot": 0, "role": "opener", "query": "market opportunity audience"},
        {"slot": 1, "role": "product", "query": "audience reach"},
    ]
    context_by_slot = {
        0: [
            {
                "source_path": "corpus/Fortune_GP_2026.pptx",
                "slide_number": 7,
                "title": "Market Opportunity",
                "body_text": ["$10B TAM", "Fortune reach: 42M"],
            }
        ],
        1: [
            {
                "source_path": "corpus/Fortune_500_2025.pptx",
                "slide_number": 3,
                "title": "Audience",
                "body_text": [],
            }
        ],
    }

    captured: dict = {}

    def fake_create(**kwargs):
        captured["user_msg"] = kwargs["messages"][0]["content"]
        mock_block = MagicMock(spec=anthropic.types.TextBlock)
        mock_block.text = '[{"action": "cover", "title": "T", "subtitle": "S"}]'
        mock_response = MagicMock()
        mock_response.usage.input_tokens = 100
        mock_response.usage.output_tokens = 50
        mock_response.content = [mock_block]
        return mock_response

    with patch("anthropic.Anthropic") as MockAnthropic:
        MockAnthropic.return_value.messages.create.side_effect = fake_create
        generator._call_claude("Test brief", arc, context_by_slot)

    user_msg = captured["user_msg"]
    assert "corpus/Fortune_GP_2026.pptx" in user_msg
    assert "slide: 7" in user_msg
    assert "corpus/Fortune_500_2025.pptx" in user_msg
    assert "slide: 3" in user_msg


def test_delete_slide_removes_slide():
    prs = _presentation_with_n_slides(3)
    DeckGenerator._delete_slide(prs, 1)
    assert len(prs.slides) == 2


def test_delete_slide_correct_slide_removed():
    prs = _presentation_with_n_slides(3)
    titles = ["A", "B", "C"]
    for slide, title in zip(prs.slides, titles):
        slide.shapes.title.text = title
    DeckGenerator._delete_slide(prs, 1)
    remaining = [s.shapes.title.text for s in prs.slides]
    assert remaining == ["A", "C"]


def test_insert_slide_at_moves_to_position():
    prs = _presentation_with_n_slides(3)
    titles = ["A", "B", "C"]
    for slide, title in zip(prs.slides, titles):
        slide.shapes.title.text = title
    source_prs = _presentation_with_n_slides(1)
    source_prs.slides[0].shapes.title.text = "NEW"
    DeckGenerator._clone_slide(source_prs, 0, prs)
    DeckGenerator._insert_slide_at(prs, 1)
    assert len(prs.slides) == 4
    assert prs.slides[1].shapes.title.text == "NEW"
    assert prs.slides[0].shapes.title.text == "A"
    assert prs.slides[2].shapes.title.text == "B"


def test_validate_template_url_allows_sharepoint():
    DeckGenerator._validate_template_url(
        "https://fortune.sharepoint.com/sites/sales/template.pptx"
    )


def test_validate_template_url_rejects_http():
    with pytest.raises(ValueError, match="HTTPS"):
        DeckGenerator._validate_template_url(
            "http://fortune.sharepoint.com/template.pptx"
        )


def test_validate_template_url_rejects_unknown_host():
    with pytest.raises(ValueError, match="host not allowed"):
        DeckGenerator._validate_template_url("https://evil.example.com/template.pptx")


def test_validate_template_url_allows_extra_host(monkeypatch):
    monkeypatch.setenv("TEMPLATE_URL_ALLOWED_HOSTS", "cdn.example.com")
    DeckGenerator._validate_template_url("https://cdn.example.com/template.pptx")


def test_build_happy_path_returns_payload(fortune_template_layouts):
    generator = _build_generator()
    generator._s3.put_object.return_value = {}
    generator._s3.generate_presigned_url.return_value = "https://s3.example.com/deck.pptx"

    schema = _build_schema()
    source_prs = _presentation_with_n_slides(1)
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.85, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with (
        patch("requests.get") as mock_get,
        patch.object(generator, "_load_pptx", return_value=source_prs),
    ):
        mock_get.return_value.content = _fortune_template_bytes()
        mock_get.return_value.raise_for_status = MagicMock()
        result = generator.build(
            schema,
            "https://fortune.sharepoint.com/template.pptx",
            mock_retriever,
        )

    assert result["download_url"] == "https://s3.example.com/deck.pptx"
    assert result["client_name"] == "Acme Corp"
    assert result["template_key"] == "template.pptx"
    assert "slide_count" in result
    generator._s3.put_object.assert_called_once()
    put_kwargs = generator._s3.put_object.call_args.kwargs
    assert put_kwargs["Key"].startswith("generated/")
    assert put_kwargs["Key"].endswith(".pptx")


def test_build_weak_match_raises(fortune_template_layouts):
    generator = _build_generator()
    schema = _build_schema()
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.3, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with patch("requests.get") as mock_get:
        mock_get.return_value.content = _fortune_template_bytes()
        mock_get.return_value.raise_for_status = MagicMock()
        with pytest.raises(ValueError, match="No good corpus match"):
            generator.build(
                schema,
                "https://fortune.sharepoint.com/template.pptx",
                mock_retriever,
            )


def test_build_missing_landmark_raises():
    generator = _build_generator()
    schema = _build_schema()
    mock_retriever = MagicMock()

    with patch("requests.get") as mock_get:
        mock_get.return_value.content = _blank_bytes(3)
        mock_get.return_value.raise_for_status = MagicMock()
        with pytest.raises(ValueError, match="missing product section landmark"):
            generator.build(
                schema,
                "https://fortune.sharepoint.com/template.pptx",
                mock_retriever,
            )


def test_build_deletes_and_inserts_correct_slide_count(fortune_template_layouts):
    generator = _build_generator()
    generator._s3.put_object.return_value = {}
    generator._s3.generate_presigned_url.return_value = "https://s3.example.com/deck.pptx"

    schema = _build_schema(
        confirmed_products=[
            Product(
                name="Product A", cadence="annual", price=10_000, category="Newsletter"
            ),
            Product(
                name="Product B",
                cadence="monthly",
                price=5_000,
                category="Digital Media",
            ),
        ]
    )
    source_prs = _presentation_with_n_slides(1)
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.9, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with (
        patch("requests.get") as mock_get,
        patch.object(generator, "_load_pptx", return_value=source_prs),
    ):
        mock_get.return_value.content = _fortune_template_bytes()
        mock_get.return_value.raise_for_status = MagicMock()
        result = generator.build(
            schema,
            "https://fortune.sharepoint.com/template.pptx",
            mock_retriever,
        )

    # 1 landmark + 2 cloned product slides = 3
    assert result["slide_count"] == 3
