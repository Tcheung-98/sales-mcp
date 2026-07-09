import io
from unittest.mock import MagicMock, patch

import pytest
from docx import Document
from pptx import Presentation
from pptx.util import Inches

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


def _presentation_with_n_slides(n: int) -> Presentation:
    prs = Presentation()
    for _ in range(n):
        prs.slides.add_slide(prs.slide_layouts[1])
    return prs


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
    # Clone slide 0 (A) — appends as slide 3, then move it to position 1
    source_prs = _presentation_with_n_slides(1)
    source_prs.slides[0].shapes.title.text = "NEW"
    DeckGenerator._clone_slide(source_prs, 0, prs)
    DeckGenerator._insert_slide_at(prs, 1)
    assert len(prs.slides) == 4
    assert prs.slides[1].shapes.title.text == "NEW"
    assert prs.slides[0].shapes.title.text == "A"
    assert prs.slides[2].shapes.title.text == "B"


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


def _fake_slides() -> list[dict]:
    return [
        {"slide_index": 0, "layout_name": "COVER blue option", "title": "Fortune x Acme", "body_text": []},  # noqa: E501
        {
            "slide_index": 1,
            "layout_name": "3_LINE_Curve with image from left",
            "title": "Why Fortune",
            "body_text": ["Fortune reaches 40M executives"],
        },
        {
            "slide_index": 2,
            "layout_name": "11_Title Only",
            "title": "Acme Scroller",
            "body_text": ["100% SOV", "Premium placement"],
        },
    ]


def test_call_claude_review_returns_arc_string():
    generator = _build_generator()
    generator._api_key = "test-key"

    mock_response = MagicMock()
    mock_response.content = [MagicMock(text="The arc for Acme is...")]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        result = generator._call_claude_review("Client: Acme Corp", _fake_slides())

    assert isinstance(result, str)
    assert "Acme" in result
    mock_anthropic.return_value.messages.create.assert_called_once()


def test_call_claude_review_includes_all_slides_in_prompt():
    generator = _build_generator()
    generator._api_key = "test-key"

    mock_response = MagicMock()
    mock_response.content = [MagicMock(text="arc context")]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        generator._call_claude_review("brief", _fake_slides())

    call_kwargs = mock_anthropic.return_value.messages.create.call_args
    prompt = call_kwargs.kwargs["messages"][0]["content"]
    assert "COVER blue option" in prompt
    assert "Why Fortune" in prompt
    assert "Acme Scroller" in prompt


def test_call_claude_write_returns_slide_list():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "Rule 1: be persuasive."

    expected = [
        {"slide_index": 0, "title": "Acme x Fortune", "eyebrow": "", "body": [], "client_name": "Acme Corp"},  # noqa: E501
        {
            "slide_index": 1,
            "title": "Your Audience",
            "eyebrow": "WHY FORTUNE",
            "body": ["40M execs"],
            "client_name": "Acme Corp",
        },
    ]
    mock_block = MagicMock()
    mock_block.type = "tool_use"
    mock_block.name = "write_deck_copy"
    mock_block.input = {"slides": expected}
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        result = generator._call_claude_write("brief", "arc context", _fake_slides())

    assert result == expected


def test_call_claude_write_raises_if_no_tool_use():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "rules"

    mock_block = MagicMock()
    mock_block.type = "text"
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        with pytest.raises(ValueError, match="did not return deck copy"):
            generator._call_claude_write("brief", "arc", _fake_slides())


def _build_schema(**overrides) -> DeckSchema:
    defaults = dict(
        client_name="Acme Corp",
        industry="Tech",
        budget_quarterly=100_000,
        confirmed_products=[
            Product(name="Fortune 500 List", cadence="annual", price=50_000, category="Newsletter")
        ],
    )
    defaults.update(overrides)
    return DeckSchema(**defaults)


def _template_bytes(slide_count: int = 3) -> bytes:
    prs = _presentation_with_n_slides(slide_count)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


_QA_APPROVED = {"approved": True, "failing_slides": []}


def test_overflow_flags_clean():
    reps = [{"slide_index": 0, "title": "Short title", "body": ["Short bullet"]}]
    assert DeckGenerator._overflow_flags(reps) == []


def test_overflow_flags_title_too_long():
    reps = [{"slide_index": 1, "title": "A" * 81, "body": []}]
    result = DeckGenerator._overflow_flags(reps)
    assert len(result) == 1
    assert result[0]["slide_index"] == 1
    assert any("title" in issue for issue in result[0]["issues"])


def test_overflow_flags_body_line_too_long():
    reps = [{"slide_index": 2, "title": "OK", "body": ["B" * 121]}]
    result = DeckGenerator._overflow_flags(reps)
    assert len(result) == 1
    assert result[0]["slide_index"] == 2
    assert any("body line" in issue for issue in result[0]["issues"])


def test_call_claude_qa_approved():
    generator = _build_generator()
    generator._api_key = "test-key"

    mock_block = MagicMock()
    mock_block.type = "tool_use"
    mock_block.name = "review_deck_copy"
    mock_block.input = {"approved": True, "failing_slides": []}
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        result = generator._call_claude_qa("brief", _fake_slides(), [
            {"slide_index": 0, "title": "T", "body": [], "client_name": "Acme"}
        ])

    assert result["approved"] is True
    assert result["failing_slides"] == []


def test_call_claude_qa_returns_failures():
    generator = _build_generator()
    generator._api_key = "test-key"

    mock_block = MagicMock()
    mock_block.type = "tool_use"
    mock_block.name = "review_deck_copy"
    mock_block.input = {
        "approved": False,
        "failing_slides": [{"slide_index": 1, "issues": ["wrong tone for CISO buyer"]}],
    }
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        result = generator._call_claude_qa("brief", _fake_slides(), [])

    assert result["approved"] is False
    assert result["failing_slides"][0]["slide_index"] == 1


def test_call_claude_qa_raises_if_no_tool_use():
    generator = _build_generator()
    generator._api_key = "test-key"

    mock_block = MagicMock()
    mock_block.type = "text"
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        with pytest.raises(ValueError, match="Sonnet QA did not return review output"):
            generator._call_claude_qa("brief", _fake_slides(), [])


def test_call_claude_write_includes_issues_in_prompt():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "rules"

    mock_block = MagicMock()
    mock_block.type = "tool_use"
    mock_block.name = "write_deck_copy"
    mock_block.input = {"slides": []}
    mock_response = MagicMock()
    mock_response.content = [mock_block]

    with patch("anthropic.Anthropic") as mock_anthropic:
        mock_anthropic.return_value.messages.create.return_value = mock_response
        generator._call_claude_write(
            "brief", "arc", _fake_slides(), issues={1: ["wrong tone for CISO"]}
        )

    call_kwargs = mock_anthropic.return_value.messages.create.call_args
    prompt = call_kwargs.kwargs["messages"][0]["content"]
    assert "ISSUES TO FIX" in prompt
    assert "wrong tone for CISO" in prompt


def test_build_happy_path_returns_payload():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "rules"
    generator._s3.put_object.return_value = {}
    generator._s3.generate_presigned_url.return_value = "https://s3.example.com/deck.pptx"

    schema = _build_schema()
    source_prs = _presentation_with_n_slides(1)
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.85, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with patch("requests.get") as mock_get, \
         patch.object(generator, "_load_pptx", return_value=source_prs), \
         patch.object(generator, "_call_claude_review", return_value="arc context"), \
         patch.object(generator, "_call_claude_write", return_value=[
             {"slide_index": 0, "title": "T", "body": [], "client_name": "Acme Corp"},
             {"slide_index": 1, "title": "T", "body": [], "client_name": "Acme Corp"},
             {"slide_index": 2, "title": "T", "body": [], "client_name": "Acme Corp"},
             {"slide_index": 3, "title": "T", "body": [], "client_name": "Acme Corp"},
         ]), \
         patch.object(generator, "_call_claude_qa", return_value=_QA_APPROVED):
        mock_get.return_value.content = _template_bytes(3)
        mock_get.return_value.raise_for_status = MagicMock()
        result = generator.build(schema, "https://example.com/template.pptx", mock_retriever)

    assert result["download_url"] == "https://s3.example.com/deck.pptx"
    assert result["client_name"] == "Acme Corp"
    assert result["template_key"] == "template.pptx"
    assert "slide_count" in result
    generator._s3.put_object.assert_called_once()


def test_build_weak_match_raises():
    generator = _build_generator()
    schema = _build_schema()
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.3, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with patch("requests.get") as mock_get:
        mock_get.return_value.content = _template_bytes()
        mock_get.return_value.raise_for_status = MagicMock()
        with pytest.raises(ValueError, match="No good corpus match"):
            generator.build(schema, "https://example.com/template.pptx", mock_retriever)


def test_build_qa_revises_failing_slides():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "rules"
    generator._s3.put_object.return_value = {}
    generator._s3.generate_presigned_url.return_value = "https://s3.example.com/deck.pptx"

    schema = _build_schema()
    source_prs = _presentation_with_n_slides(1)
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.9, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    qa_responses = iter([
        {"approved": False, "failing_slides": [{"slide_index": 1, "issues": ["weak narrative"]}]},
        _QA_APPROVED,
    ])
    _rep = {"title": "T", "body": [], "client_name": "Acme Corp"}
    write_responses = iter([
        [{**_rep, "slide_index": i} for i in range(4)],
        [{"slide_index": 1, "title": "Revised", "body": ["Better copy"], "client_name": "Acme Corp"}],  # noqa: E501
    ])

    with patch("requests.get") as mock_get, \
         patch.object(generator, "_load_pptx", return_value=source_prs), \
         patch.object(generator, "_call_claude_review", return_value="arc"), \
         patch.object(generator, "_call_claude_write", side_effect=write_responses) as mock_write, \
         patch.object(generator, "_call_claude_qa", side_effect=qa_responses) as mock_qa:
        mock_get.return_value.content = _template_bytes(3)
        mock_get.return_value.raise_for_status = MagicMock()
        result = generator.build(schema, "https://example.com/template.pptx", mock_retriever)

    assert result["download_url"] == "https://s3.example.com/deck.pptx"
    assert mock_write.call_count == 2  # initial + 1 revision
    assert mock_qa.call_count == 2     # failed + approved


def test_build_deletes_and_inserts_correct_slide_count():
    generator = _build_generator()
    generator._api_key = "test-key"
    generator._rulebook_text = "rules"
    generator._s3.put_object.return_value = {}
    generator._s3.generate_presigned_url.return_value = "https://s3.example.com/deck.pptx"

    schema = _build_schema(
        confirmed_products=[
            Product(name="Product A", cadence="annual", price=10_000, category="Newsletter"),
            Product(name="Product B", cadence="monthly", price=5_000, category="Digital Media"),
        ]
    )
    source_prs = _presentation_with_n_slides(1)
    mock_retriever = MagicMock()
    mock_retriever.search.return_value = [
        {"score": 0.9, "source_path": "corpus/deck.pptx", "slide_number": 1}
    ]

    with patch("requests.get") as mock_get, \
         patch.object(generator, "_load_pptx", return_value=source_prs), \
         patch.object(generator, "_call_claude_review", return_value="arc"), \
         patch.object(generator, "_call_claude_write", return_value=[
             {"slide_index": i, "title": "T", "body": [], "client_name": "Acme Corp"}
             for i in range(5)
         ]), \
         patch.object(generator, "_call_claude_qa", return_value=_QA_APPROVED):
        mock_get.return_value.content = _template_bytes(3)
        mock_get.return_value.raise_for_status = MagicMock()
        result = generator.build(schema, "https://example.com/template.pptx", mock_retriever)

    # 3 template slides + 2 product clones = 5
    assert result["slide_count"] == 5

