"""Guardrailed PPTX helpers for skeleton assembly and (later) Cursor stylist apply.

Clone/delete stay on DeckGenerator; this module owns text-apply helpers so agent
scripts can call them without importing the full generator.
"""

from __future__ import annotations

import re

_CLIENT_NAME_PLACEHOLDERS = re.compile(
    r"your brand|your company|client name|your organization",
    re.IGNORECASE,
)


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
