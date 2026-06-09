import csv
import io
import json
import logging
import os
import re

import anthropic
import boto3
from docx import Document
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType
from pypdf import PdfReader

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"
_COST_PER_M_INPUT = 3.00
_COST_PER_M_OUTPUT = 15.00

# Corpus slides use these strings as stand-ins for the client name.
# All are matched case-insensitively so "Your Brand" and "YOUR BRAND" both swap.
_CLIENT_NAME_PLACEHOLDERS = re.compile(
    r"your brand|your company|client name|your organization",
    re.IGNORECASE,
)


_SYSTEM_PROMPT_TEMPLATE = """\
You are a senior AE at Fortune Media Group building a custom pitch deck for a specific client.

=== FORTUNE SALES SKILL — RULEBOOK ===
{rulebook_text}
=== END RULEBOOK ===

OUTPUT FORMAT
Return ONLY a JSON array — no prose, no markdown fences. One element per arc slot, in slot order.

Slot 1 (cover) is always:
  {{"action": "cover", "title": "...", "subtitle": "..."}}

All other slots clone a corpus slide:
  {{
    "action": "clone",
    "source_path": "...",
    "slide_number": 7,
    "reasoning": "One sentence: why this slide fits this slot and client.",
    "replacements": {{
      "title": "...",
      "client_name": "Acme Corp",
      "body": ["P1 text", "P2 text", ...]
    }}
  }}

FIELD RULES
- For each non-cover slot, pick exactly one slide from THAT SLOT'S candidates
- Do not use candidates listed under a different slot
- source_path and slide_number must come exactly from the candidates shown for that slot
- reasoning: required on every clone — explain why this slide fits this slot's role
- replacements.client_name: always set to the client name from the brief
- replacements.title: override if needed; omit to keep original
- replacements.eyebrow: override if needed; omit to keep original
- replacements.body: rewrite paragraphs to be client-specific. Match the P1/P2/... structure \
shown. Preserve structural headers; rewrite descriptive copy. 18 words max per bullet. \
OMIT body entirely for: (1) data-heavy slides — audience metrics, bar charts, structured \
tables, When/Where/Who grids, pricing tables; (2) event/conference slides (Brainstorm, \
live events, summits) — their value is in the visual design and event details table, not \
rewritten prose. For investment slides, limit body to 5 bullets maximum — one per product \
line, no sub-bullets, no notes. Only provide body for slides with a clear narrative text \
block and no data visualization or event details table.

SELECTION RULES
- Output exactly one slide per slot, in slot order
- Follow the arc, escalation, and voice rules in the rulebook above
- Use the rate card for accurate product names and pricing\
"""


def _format_slide(s: dict) -> str:
    lines = [
        f"[source: {s.get('source_path', '')} | slide: {s.get('slide_number', '?')}]",
        f"Title: {s.get('title') or '(none)'}",
    ]
    for i, b in enumerate(s.get("body_text") or [], start=1):
        lines.append(f"P{i}: {b}")
    return "\n".join(lines)


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        blank_key: str | None = None,
        rulebook_key: str | None = None,
        rate_sheet_key: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._blank_key = blank_key or os.environ.get(
            "PPTX_BLANK_DECK_KEY", "templates/blank.pptx"
        )
        self._rulebook_key = rulebook_key or os.environ.get(
            "RULEBOOK_KEY", "templates/rulebook.docx"
        )
        self._rate_sheet_key = rate_sheet_key or os.environ.get("RATE_SHEET_KEY")
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._blank_bytes: bytes | None = None
        self._pptx_cache: dict[str, PresentationType] = {}
        self._rulebook_text: str | None = None
        self._api_key: str | None = None
        self._rate_sheet: str | None = None

    @staticmethod
    def _clone_slide(
        source_prs: PresentationType, slide_idx: int, target_prs: PresentationType
    ):
        source_slide = source_prs.slides[slide_idx]

        # Add a placeholder slide — gives us a proper slide part + sldIdLst entry
        new_slide = target_prs.slides.add_slide(target_prs.slide_layouts[0])

        # Register slide-level image parts from source into target; build rId map
        rId_map: dict[str, str] = {}
        for rId, rel in source_slide.part.rels.items():
            if "image" in rel.reltype:
                image_part = source_slide.part.related_part(rId)
                new_rId = new_slide.part.relate_to(image_part, rel.reltype)
                rId_map[rId] = new_rId

        # Serialize source cSld, rewire rIds, re-parse with pptx element classes
        # (parse_xml, not etree.fromstring, so spTree and other pptx attrs are available)
        cSld_xml = etree.tostring(source_slide._element.cSld)
        for old_rId, new_rId in rId_map.items():
            cSld_xml = cSld_xml.replace(
                f'r:embed="{old_rId}"'.encode(), f'r:embed="{new_rId}"'.encode()
            )
            cSld_xml = cSld_xml.replace(
                f'r:link="{old_rId}"'.encode(), f'r:link="{new_rId}"'.encode()
            )
        fixed_cSld = parse_xml(cSld_xml)

        # Swap target slide's cSld with the fixed clone
        tgt_sld = new_slide._element
        tgt_sld.replace(tgt_sld.cSld, fixed_cSld)

        # shapes is a lazyproperty cached during add_slide; invalidate so next access
        # gets a fresh SlideShapes pointing at the new spTree, not the discarded one
        new_slide.__dict__.pop("shapes", None)

        # Rewire layout relationship to match source slide's actual layout
        src_layout_name = source_slide.slide_layout.name
        target_layout = next(
            (lay for lay in target_prs.slide_layouts if lay.name == src_layout_name),
            target_prs.slide_layouts[0],
        )
        for rId, rel in list(new_slide.part.rels.items()):
            if rel.reltype == RT.SLIDE_LAYOUT:
                new_slide.part.drop_rel(rId)
                break
        new_slide.part.relate_to(target_layout.part, RT.SLIDE_LAYOUT)

        return new_slide

    def _load_pptx(self, s3_key: str) -> PresentationType:
        if s3_key not in self._pptx_cache:
            resp = self._s3.get_object(Bucket=self._bucket, Key=s3_key)
            self._pptx_cache[s3_key] = Presentation(io.BytesIO(resp["Body"].read()))
            logger.info("loaded pptx from s3://%s/%s", self._bucket, s3_key)
        return self._pptx_cache[s3_key]

    def _load_rulebook(self) -> str:
        if self._rulebook_text is None:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._rulebook_key)
            body = resp["Body"].read()
            if self._rulebook_key.endswith(".pdf"):
                pdf = PdfReader(io.BytesIO(body))
                self._rulebook_text = "\n".join(
                    page.extract_text() for page in pdf.pages if page.extract_text()
                )
            else:
                doc = Document(io.BytesIO(body))
                self._rulebook_text = "\n".join(
                    p.text for p in doc.paragraphs if p.text.strip()
                )
            logger.info("loaded rulebook from s3://%s/%s", self._bucket, self._rulebook_key)
        return self._rulebook_text

    def _get_api_key(self) -> str:
        if self._api_key is None:
            # Local dev: ANTHROPIC_API_KEY in .env bypasses Secrets Manager
            if env_key := os.environ.get("ANTHROPIC_API_KEY"):
                self._api_key = env_key
                return env_key
            sm = boto3.client("secretsmanager", region_name="us-east-1")
            resp = sm.get_secret_value(SecretId=self._secret_name)
            raw = resp["SecretString"]
            # Handle JSON-wrapped secrets e.g. {"api_key": "sk-ant-..."}
            try:
                parsed = json.loads(raw)
                self._api_key = next(iter(parsed.values())) if isinstance(parsed, dict) else raw
            except (json.JSONDecodeError, StopIteration):
                self._api_key = raw
        return self._api_key

    def _call_claude(
        self, brief: str, arc: list[dict], context_by_slot: dict[int, list[dict]]
    ) -> list[dict]:
        context_parts = []
        for slot in arc:
            if not slot.get("query"):
                continue
            candidates = context_by_slot.get(slot["slot"], [])
            if not candidates:
                continue
            context_parts.append(
                f"=== Slot {slot['slot']}: {slot['role'].upper()} ===\n"
                + "\n\n".join(_format_slide(s) for s in candidates)
            )
        context_text = "\n\n".join(context_parts)

        arc_summary = "\n".join(
            f"  Slot {s['slot']}: {s['role']}" for s in arc
        )

        if self._rate_sheet is None and self._rate_sheet_key:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._rate_sheet_key)
            reader = csv.DictReader(io.StringIO(resp["Body"].read().decode("cp1252")))
            products = [row for row in reader if row.get("Inventory", "").strip()]
            lines = []
            for p in products:
                parts = [p.get("Inventory", "").strip(), p.get("Product Category", "").strip()]
                if p.get("Cadence", "").strip():
                    parts.append(p["Cadence"].strip())
                if p.get("Vertical", "").strip():
                    parts.append(f"Verticals: {p['Vertical'].strip()}")
                if p.get("Audience Alignment", "").strip():
                    parts.append(f"Audience: {p['Audience Alignment'].strip()}")
                if p.get("Contextual Alignment", "").strip():
                    parts.append(f"Context: {p['Contextual Alignment'].strip()}")
                pricing = []
                for label, col in [
                    ("Daily", "Daily_Cost"), ("Weekly", "Weekly_Cost"),
                    ("Monthly", "Monthly_Cost"), ("Quarterly", "Quarterly_Cost"),
                    ("Half-Year", "Half_Year_Cost"), ("Annual", "Annual_Cost"),
                    ("CPM", "CPM_Rate"), ("Min", "Product_Minimum"), ("Flat Fee", "Flat_Fee"),
                ]:
                    val = p.get(col, "").strip()
                    if val:
                        prefix = "" if val.startswith("$") else "$"
                        pricing.append(f"{label}: {prefix}{val}")
                if pricing:
                    parts.append(" | ".join(pricing))
                if p.get("Notes", "").strip():
                    parts.append(f"Notes: {p['Notes'].strip()}")
                lines.append(" | ".join(filter(None, parts)))
            self._rate_sheet = "\n".join(lines)
            logger.info("loaded rate sheet: %d products", len(products))
        rate_section = (
            f"\n\nFortune product & rate card (use for accurate product names and pricing):\n"
            f"{self._rate_sheet}"
            if self._rate_sheet else ""
        )

        user_msg = (
            f"Brief:\n{brief}"
            f"\n\nDeck arc to fill:\n{arc_summary}"
            f"{rate_section}"
            f"\n\nCandidate slides organized by slot:\n{context_text}"
        )

        system_prompt = _SYSTEM_PROMPT_TEMPLATE.format(
            rulebook_text=self._load_rulebook()
        )

        client = anthropic.Anthropic(api_key=self._get_api_key())
        response = client.messages.create(
            model=self._model,
            max_tokens=4096,
            system=system_prompt,
            messages=[{"role": "user", "content": user_msg}],
        )

        logger.info(
            "claude usage — input: %d tokens, output: %d tokens, cost: $%.4f",
            response.usage.input_tokens,
            response.usage.output_tokens,
            (response.usage.input_tokens / 1_000_000 * _COST_PER_M_INPUT)
            + (response.usage.output_tokens / 1_000_000 * _COST_PER_M_OUTPUT),
        )
        block = response.content[0]
        if not isinstance(block, anthropic.types.TextBlock):
            raise ValueError(f"Unexpected response block type: {type(block)}")
        raw = block.text.strip()
        # Strip markdown fences if Claude ignored the prompt instruction
        if raw.startswith("```"):
            raw = raw.split("\n", 1)[-1]
            raw = raw.rsplit("```", 1)[0].strip()
        logger.debug("claude raw response: %s", raw[:200])
        slides = json.loads(raw)
        if not isinstance(slides, list):
            raise ValueError(f"Claude returned non-list JSON: {type(slides)}")
        return slides

    @staticmethod
    def _dedup_shape_ids(slide, output_prs) -> None:
        """Renumber cNvPr ids on a freshly-cloned slide to avoid collisions across slides."""
        used: set[int] = set()
        for s in output_prs.slides:
            if s._element is slide._element:
                continue
            for el in s._element.iter(qn("p:cNvPr")):
                try:
                    used.add(int(el.get("id", 0)))
                except (ValueError, TypeError):
                    pass
        next_id = max(used, default=0) + 1
        for el in slide._element.iter(qn("p:cNvPr")):
            try:
                sid = int(el.get("id", 0))
            except (ValueError, TypeError):
                sid = 0
            if sid == 0 or sid in used:
                el.set("id", str(next_id))
                next_id += 1
            else:
                used.add(sid)

    @staticmethod
    def _set_ph_text(ph, text: str) -> None:
        """Replace placeholder text while preserving run-level formatting (font, color, size)."""
        tf = ph.text_frame
        first_para = tf.paragraphs[0]
        # Remove extra paragraphs bottom-up
        for para in tf.paragraphs[1:]:
            para._p.getparent().remove(para._p)
        runs = first_para.runs
        if runs:
            # Set text on first run, preserving its rPr; drop extra runs
            runs[0].text = text
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
        else:
            first_para.text = text

    @staticmethod
    def _apply_replacements(slide, replacements: dict) -> None:
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

        # title / eyebrow: overwrite their placeholder slots preserving formatting
        if title and (ph := ph_map.get(0)):
            DeckGenerator._set_ph_text(ph, title)
        if eyebrow and (ph := ph_map.get(19)):
            DeckGenerator._set_ph_text(ph, eyebrow)

        # body: rewrite paragraphs one-for-one across all non-title/non-eyebrow shapes.
        # body_text[] in the retriever is built by scraping every shape's text in slide order,
        # so we traverse the same shapes here to ensure Claude's replacements land in the
        # right places — regardless of which placeholder indices Fortune uses (idx 12, 17, etc.)
        if body:
            _HANDLED_IDX = {0, 19}
            body_shapes = []
            for s in slide.shapes:
                if not s.has_text_frame:
                    continue
                try:
                    idx = s.placeholder_format.idx
                except ValueError:
                    continue  # non-placeholder shape (chart annotation, floating text box) — skip
                if idx in _HANDLED_IDX:
                    continue
                body_shapes.append(s)
            # Flatten all paragraphs from those shapes into one ordered list
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

        # client_name: find-and-replace within runs to preserve formatting
        if client_name:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if _CLIENT_NAME_PLACEHOLDERS.search(run.text):
                            run.text = _CLIENT_NAME_PLACEHOLDERS.sub(client_name, run.text)

    def _build_pptx(self, slides: list[dict]) -> bytes:
        if self._blank_bytes is None:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._blank_key)
            self._blank_bytes = resp["Body"].read()
            logger.info("loaded blank template from s3://%s/%s", self._bucket, self._blank_key)
        blank_prs = Presentation(io.BytesIO(self._blank_bytes))  # cover slide source
        output_prs = Presentation(io.BytesIO(self._blank_bytes))  # output deck

        # Drop relationship AND sldId element for each seed slide
        for sld_id in list(output_prs.slides._sldIdLst):
            r_id = sld_id.get(qn("r:id"))
            output_prs.part.drop_rel(r_id)
            output_prs.slides._sldIdLst.remove(sld_id)

        for slide_data in slides:
            action = slide_data.get("action", "clone")
            if action == "cover":
                slide = self._clone_slide(blank_prs, 0, output_prs)
                DeckGenerator._dedup_shape_ids(slide, output_prs)
                ph_map = {
                    ph.placeholder_format.idx: ph
                    for ph in slide.placeholders
                    if ph.has_text_frame
                }
                for title_idx in (12, 0):
                    if ph := ph_map.get(title_idx):
                        DeckGenerator._set_ph_text(ph, slide_data.get("title", ""))
                        break
                for sub_idx in (10, 1):
                    if ph := ph_map.get(sub_idx):
                        DeckGenerator._set_ph_text(ph, slide_data.get("subtitle", ""))
                        break
            elif action == "clone":
                source_prs = self._load_pptx(slide_data["source_path"])
                slide = self._clone_slide(source_prs, slide_data["slide_number"] - 1, output_prs)
                DeckGenerator._dedup_shape_ids(slide, output_prs)
                self._apply_replacements(slide, slide_data.get("replacements", {}))
            else:
                raise ValueError(f"Unknown slide action: {action!r}")

        buf = io.BytesIO()
        output_prs.save(buf)
        return buf.getvalue()

