"""Guardrailed PPTX helpers for skeleton assembly and (later) Cursor stylist apply.

Clone/delete stay on DeckGenerator; this module owns text-apply helpers so agent
scripts can call them without importing the full generator.

C2 (PI-2757) token/logo primitives live here. They are unwired from assemble_skeleton
until a later chunk; tests call them directly.
"""

from __future__ import annotations

import io
import re

from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

# FortuneAI History of Trust / Audience possessive tokens use U+2019.
APOS = "\u2019"
CLIENT_NAME_TOKEN = "[client name]"
CLIENT_NAME_POSSESSIVE_TOKEN = f"[CLIENT NAME{APOS}S]"
LOGO_TOKEN = "[LOGO]"

_CLIENT_NAME_PLACEHOLDERS = re.compile(
    r"your brand|your company|client name|your organization",
    re.IGNORECASE,
)

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"
_JPEG_MAGIC = b"\xff\xd8"
_GIF_MAGIC = b"GIF8"


def set_ph_text(ph, text: str) -> None:
    """Replace placeholder text while preserving run-level formatting."""
    tf = ph.text_frame
    first_para = tf.paragraphs[0]
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)
    runs = first_para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        first_para.text = text


def apply_replacements(slide, replacements: dict) -> None:
    """Apply title/eyebrow/body/client_name replacements on a slide."""
    if not replacements:
        return

    title = replacements.get("title")
    eyebrow = replacements.get("eyebrow")
    client_name = replacements.get("client_name")
    body = replacements.get("body")

    ph_map = {
        ph.placeholder_format.idx: ph
        for ph in slide.placeholders
        if ph.has_text_frame
    }

    if title and (ph := ph_map.get(0)):
        set_ph_text(ph, title)
    if eyebrow and (ph := ph_map.get(19)):
        set_ph_text(ph, eyebrow)

    # body: rewrite paragraphs one-for-one across non-title/non-eyebrow shapes.
    # body_text[] in the retriever is scraped in slide order, so traverse the same way.
    if body:
        _HANDLED_IDX = {0, 19}
        body_shapes = []
        for s in slide.shapes:
            if not s.has_text_frame:
                continue
            try:
                idx = s.placeholder_format.idx
            except ValueError:
                continue
            if idx in _HANDLED_IDX:
                continue
            body_shapes.append(s)
        all_paras = [p for s in body_shapes for p in s.text_frame.paragraphs]
        for i, new_text in enumerate(body):
            if i >= len(all_paras):
                break
            para = all_paras[i]
            runs = para.runs
            if runs:
                runs[0].text = new_text
                for r in runs[1:]:
                    r._r.getparent().remove(r._r)
            else:
                para.text = new_text

    if client_name:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if _CLIENT_NAME_PLACEHOLDERS.search(run.text):
                        run.text = _CLIENT_NAME_PLACEHOLDERS.sub(
                            client_name, run.text
                        )


def iter_shapes(slide):
    """Yield every shape on a slide, including those nested in groups."""

    def _walk(container):
        for shape in container.shapes:
            yield shape
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from _walk(shape)

    yield from _walk(slide)


def _shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    return "\n".join(p.text or "" for p in shape.text_frame.paragraphs)


def replace_token(slide, token: str, text: str) -> int:
    """Replace exact ``token`` substrings in all text frames. Returns hit count.

    Prefers in-run replace so formatting stays. If the token is split across
    runs in a paragraph, the paragraph is collapsed onto the first run.
    """
    if not token:
        raise ValueError("token must be non-empty")

    hits = 0
    for shape in iter_shapes(slide):
        if not getattr(shape, "has_text_frame", False):
            continue
        for para in shape.text_frame.paragraphs:
            run_hit = False
            for run in para.runs:
                if token not in run.text:
                    continue
                n = run.text.count(token)
                run.text = run.text.replace(token, text)
                hits += n
                run_hit = True
            if run_hit:
                continue
            full = para.text or ""
            if token not in full:
                continue
            n = full.count(token)
            new = full.replace(token, text)
            if para.runs:
                para.runs[0].text = new
                for run in para.runs[1:]:
                    run.text = ""
            else:
                para.text = new
            hits += n
    return hits


def _is_picture_placeholder(shape) -> bool:
    try:
        return shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE
    except (ValueError, AttributeError):
        return False


def _find_logo_placeholder(slide):
    pictures = [s for s in iter_shapes(slide) if _is_picture_placeholder(s)]
    labeled = [s for s in pictures if LOGO_TOKEN in _shape_text(s)]
    if len(labeled) == 1:
        return labeled[0]
    if len(labeled) > 1:
        raise ValueError("slide has more than one [LOGO] picture placeholder")
    idx11 = []
    for shape in pictures:
        try:
            if shape.placeholder_format.idx == 11:
                idx11.append(shape)
        except (ValueError, AttributeError):
            continue
    if len(idx11) == 1:
        return idx11[0]
    if len(pictures) == 1:
        return pictures[0]
    if pictures:
        raise ValueError(
            "slide has multiple picture placeholders and none are [LOGO] "
            "or placeholder idx 11"
        )
    return None


def _validate_logo_bytes(image_bytes: bytes) -> None:
    if not image_bytes:
        raise ValueError("logo image is empty")
    if image_bytes.startswith(_PNG_MAGIC):
        return
    if image_bytes.startswith(_JPEG_MAGIC):
        return
    if image_bytes.startswith(_GIF_MAGIC):
        return
    raise ValueError("logo image is unreadable (expected PNG, JPEG, or GIF)")


def insert_logo(slide, image_bytes: bytes) -> None:
    """Insert ``image_bytes`` into the slide's [LOGO] picture placeholder.

    FortuneAI intro/thanks use a PICTURE placeholder (idx 11) whose text is
    ``[LOGO]``. Fixtures may use any single picture placeholder. Fails loud if
    the placeholder is missing or the bytes are not a PNG/JPEG/GIF.
    """
    _validate_logo_bytes(image_bytes)
    placeholder = _find_logo_placeholder(slide)
    if placeholder is None:
        raise ValueError("logo placeholder not found ([LOGO] picture placeholder)")
    try:
        placeholder.insert_picture(io.BytesIO(image_bytes))
    except Exception as exc:
        raise ValueError(f"logo image is unreadable: {exc}") from exc


def delete_slide(prs, slide_idx: int) -> None:
    """Remove a slide by 0-based index (sldIdLst + relationship)."""
    n = len(prs.slides)
    if slide_idx < 0 or slide_idx >= n:
        raise ValueError(f"slide index {slide_idx} out of range (0..{n - 1})")
    sld_id_lst = prs.slides._sldIdLst
    sld_id = sld_id_lst[slide_idx]
    r_id = sld_id.get(qn("r:id"))
    prs.part.drop_rel(r_id)
    sld_id_lst.remove(sld_id)
