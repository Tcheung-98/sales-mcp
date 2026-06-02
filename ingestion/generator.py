import io
import json
import logging
import os
import uuid

import anthropic
import boto3
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"
_GENERATED_PREFIX = "generated"

_SYSTEM_PROMPT = """\
You are a senior AE at Fortune Media Group writing a custom pitch deck for a specific client.

Return ONLY a JSON array — no prose, no markdown fences. Each element is one slide object
with a "slide_type" field. Use exactly these types and fields:

  {"slide_type": "title",       "title": "...", "subtitle": "..."}
  {"slide_type": "product",     "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "proof",       "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "investment",  "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "next_steps",  "title": "...", "eyebrow": "...", "bullets": ["..."]}

Field rules:
- "eyebrow": short ALL-CAPS label above the title (e.g. "RECOMMENDED PRODUCTS", \
"PERFORMANCE DATA", "INVESTMENT", "NEXT STEPS"). Omit only if nothing fits.
- "bullets": plain bullet fragments. Max 18 words each. Max 3 per slide.

You will receive: (1) a client brief, (2) Fortune's product & rate card, \
(3) reference slides from past Fortune decks.

How to use each input:
- Brief: defines the client, their industry, their goal, their target audience — \
every slide must speak directly to THIS client's situation
- Rate card: pick 3-5 products that genuinely fit this client's vertical and audience; \
use real product names, real audience descriptors, and real pricing — but price is the \
last thing you lead with, not the first
- Reference slides: pull in any specific performance stats, audience indices, reach numbers, \
or proof points that are relevant to this client — these make bullets credible

Writing rules — write like a pitch, not a spec sheet:
- Each bullet: product or proof point + why it fits THIS client. Max 18 words. No exceptions.
- Lead with the value to the client, not the product name or the price
- 3 bullets per slide maximum. Cut the weakest one if you have 4.
- Fragments only. No full sentences.
- No filler: ban "leverage", "synergy", "best-in-class", "cutting-edge", "seamless", \
"robust", "drive engagement"
- Pricing appears only on the investment slide, never elsewhere

Deck structure (slide_type, in order):
1. title — client-specific headline naming their goal, never "Fortune Media Pitch"
2. product × 1-2 — each product by what it does and who it reaches, tied to the client's \
target audience; eyebrow "RECOMMENDED PRODUCTS"
3. proof — performance stats or audience data from reference slides; eyebrow "PERFORMANCE \
DATA"; omit this slide entirely if no relevant data exists in the reference slides
4. investment — clean pricing tiers using real figures from the rate card; eyebrow \
"INVESTMENT"; pricing appears ONLY here
5. next_steps — 3 concrete actions with implied timeline; eyebrow "NEXT STEPS"\
"""


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        blank_key: str | None = None,
        rate_sheet_key: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._blank_key = blank_key or os.environ.get(
            "PPTX_BLANK_DECK_KEY", "templates/blank.pptx"
        )
        self._rate_sheet_key = rate_sheet_key or os.environ.get("RATE_SHEET_KEY")
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._blank_bytes: bytes | None = None
        self._api_key: str | None = None
        self._rate_sheet: str | None = None

    @staticmethod
    def _clear_seed_slides(prs: PresentationType) -> None:
        # Proper deletion: drop relationship AND sldId element.
        for sld_id in list(prs.slides._sldIdLst):
            r_id = sld_id.get(qn("r:id"))
            prs.part.drop_rel(r_id)
            prs.slides._sldIdLst.remove(sld_id)

    @staticmethod
    def _set_title(slide, text: str) -> None:
        if slide.shapes.title:
            slide.shapes.title.text = text

    @staticmethod
    def _set_body(slide, bullets: list[str]) -> None:
        # Clear all non-title text frames.
        for ph in slide.placeholders:
            if ph.placeholder_format.idx != 0 and ph.has_text_frame:
                ph.text_frame.clear()
        # idx 1 is the content body in all Fortune "Title and Content" layouts.
        ph_map = {
            ph.placeholder_format.idx: ph
            for ph in slide.placeholders
            if ph.has_text_frame
        }
        body_ph = ph_map.get(1) or next(
            (ph for idx, ph in ph_map.items() if idx != 0), None
        )
        if body_ph is None:
            return
        tf = body_ph.text_frame
        for i, bullet in enumerate(bullets):
            if i == 0:
                tf.paragraphs[0].text = bullet
            else:
                tf.add_paragraph().text = bullet

    @staticmethod
    def _prune_placeholders(slide, keep: frozenset = frozenset({0, 1})) -> None:
        for ph in list(slide.placeholders):
            if ph.placeholder_format.idx not in keep:
                parent = ph._element.getparent()
                if parent is not None:
                    parent.remove(ph._element)

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

    @staticmethod
    def _populate_title_slide(slide, slide_data: dict) -> None:
        # COVER blue option has no idx 0: idx 12 = headline, idx 10 = subtitle.
        ph_map = {
            ph.placeholder_format.idx: ph
            for ph in slide.placeholders
            if ph.has_text_frame
        }
        for title_idx in (12, 0):
            ph = ph_map.get(title_idx)
            if ph:
                ph.text_frame.clear()
                ph.text_frame.paragraphs[0].text = slide_data.get("title", "")
                break
        for sub_idx in (10, 1):
            ph = ph_map.get(sub_idx)
            if ph:
                ph.text_frame.clear()
                ph.text_frame.paragraphs[0].text = slide_data.get("subtitle", "")
                break

    def _populate_content_slide(self, slide, slide_data: dict) -> None:
        self._set_title(slide, slide_data.get("title", ""))
        self._set_body(slide, slide_data.get("bullets", []))
        eyebrow = slide_data.get("eyebrow")
        if eyebrow:
            ph_map = {
                ph.placeholder_format.idx: ph
                for ph in slide.placeholders
                if ph.has_text_frame
            }
            eyebrow_ph = ph_map.get(20)
            if eyebrow_ph:
                eyebrow_ph.text_frame.clear()
                eyebrow_ph.text_frame.paragraphs[0].text = eyebrow
        keep = frozenset({0, 1, 20}) if eyebrow else frozenset({0, 1})
        self._prune_placeholders(slide, keep)

    def _load_blank(self) -> bytes:
        if self._blank_bytes is None:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._blank_key)
            self._blank_bytes = resp["Body"].read()
            logger.info("loaded blank template from s3://%s/%s", self._bucket, self._blank_key)
        return self._blank_bytes

    def _load_rate_sheet(self) -> str | None:
        if self._rate_sheet is None and self._rate_sheet_key:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._rate_sheet_key)
            products = json.loads(resp["Body"].read())
            lines = []
            for p in products:
                parts = [p.get("Inventory", ""), p.get("Product Category", "")]
                if p.get("Cadence"):
                    parts.append(p["Cadence"])
                if p.get("Vertical"):
                    parts.append(f"Verticals: {p['Vertical']}")
                if p.get("Audience Alignment"):
                    parts.append(f"Audience: {p['Audience Alignment']}")
                if p.get("Contextual Alignment"):
                    parts.append(f"Context: {p['Contextual Alignment']}")
                pricing = []
                for label, key in [
                    ("Daily", "Daily_Cost"), ("Weekly", "Weekly_Cost"),
                    ("Monthly", "Monthly_Cost"), ("Quarterly", "Quarterly_Cost"),
                    ("Annual", "Annual_Cost"), ("CPM", "CPM_Rate"),
                    ("Min", "Product_Minimum"), ("Flat Fee", "Flat_Fee"),
                ]:
                    if p.get(key):
                        pricing.append(f"{label}: ${p[key]}")
                if pricing:
                    parts.append(" | ".join(pricing))
                if p.get("Notes"):
                    parts.append(f"Notes: {p['Notes']}")
                lines.append(" | ".join(filter(None, parts)))
            self._rate_sheet = "\n".join(lines)
            logger.info("loaded rate sheet: %d products", len(products))
        return self._rate_sheet

    def _get_api_key(self) -> str:
        if self._api_key is None:
            # Local dev: ANTHROPIC_API_KEY in .env bypasses Secrets Manager
            if env_key := os.environ.get("ANTHROPIC_API_KEY"):
                self._api_key = env_key
                return self._api_key
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

    def _call_claude(self, brief: str, context_slides: list[dict]) -> list[dict]:
        context_text = "\n\n".join(
            f"[Slide from {s.get('source_path', '').split('/')[-1]}]\n"
            f"Title: {s.get('title') or '(none)'}\n"
            + "\n".join(f"- {b}" for b in (s.get("body_text") or []))
            for s in context_slides
        )

        rate_sheet = self._load_rate_sheet()
        rate_section = (
            f"\n\nFortune product & rate card (use for accurate product names and pricing):\n"
            f"{rate_sheet}"
            if rate_sheet else ""
        )

        user_msg = (
            f"Brief:\n{brief}"
            f"{rate_section}"
            f"\n\nReference slides from Fortune corpus:\n{context_text}"
        )

        client = anthropic.Anthropic(api_key=self._get_api_key())
        response = client.messages.create(
            model=self._model,
            max_tokens=4096,
            system=_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )

        input_tokens = response.usage.input_tokens
        output_tokens = response.usage.output_tokens
        cost_usd = (input_tokens / 1_000_000 * 3.00) + (output_tokens / 1_000_000 * 15.00)
        logger.info(
            "claude usage — input: %d tokens, output: %d tokens, cost: $%.4f",
            input_tokens,
            output_tokens,
            cost_usd,
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

    def _build_pptx(self, slides: list[dict]) -> bytes:
        blank = self._load_blank()
        source_prs = Presentation(io.BytesIO(blank))   # read-only clone source
        output_prs = Presentation(io.BytesIO(blank))   # output (inherits Fortune master/theme)
        self._clear_seed_slides(output_prs)

        layout_map = {lay.name: lay for lay in output_prs.slide_layouts}
        content_layout = layout_map.get("8_Title and Content", output_prs.slide_layouts[2])

        for slide_data in slides:
            slide_type = slide_data.get("slide_type", "product")
            if slide_type == "title":
                slide = self._clone_slide(source_prs, 0, output_prs)
                self._populate_title_slide(slide, slide_data)
            else:
                # product, proof, investment, next_steps
                slide = output_prs.slides.add_slide(content_layout)
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = RGBColor(0x10, 0x18, 0x5F)
                self._populate_content_slide(slide, slide_data)

        buf = io.BytesIO()
        output_prs.save(buf)
        return buf.getvalue()

    def _upload(self, pptx_bytes: bytes) -> dict:
        key = f"{_GENERATED_PREFIX}/{uuid.uuid4()}.pptx"
        self._s3.put_object(
            Bucket=self._bucket,
            Key=key,
            Body=pptx_bytes,
            ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        url = self._s3.generate_presigned_url(
            "get_object",
            Params={"Bucket": self._bucket, "Key": key},
            ExpiresIn=86400,
        )
        return {"s3_uri": f"s3://{self._bucket}/{key}", "download_url": url}

    def generate(self, brief: str, context_slides: list[dict]) -> dict:
        logger.info("generating deck for brief: %.80s", brief)
        slides = self._call_claude(brief, context_slides)
        logger.info("claude returned %d slides", len(slides))
        pptx_bytes = self._build_pptx(slides)
        result = self._upload(pptx_bytes)
        return {**result, "slide_count": len(slides), "brief": brief}
