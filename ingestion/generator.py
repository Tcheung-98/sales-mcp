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
  {"slide_type": "why_fortune", "title": "...", "body": "..."}
  {"slide_type": "product",     "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "proof",       "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "investment",  "title": "...", "eyebrow": "...", "bullets": ["..."]}
  {"slide_type": "next_steps",  "title": "...", "eyebrow": "...", "bullets": ["..."]}

Field rules:
- "eyebrow": short ALL-CAPS label above the title (e.g. "RECOMMENDED PRODUCTS", \
"PERFORMANCE DATA", "INVESTMENT", "NEXT STEPS"). Omit only if nothing fits.
- "body" (why_fortune only): 1-3 sentences of descriptive prose. Not bullets, not fragments.
- "bullets": em-dash items. Max 18 words each. Max 3 per slide. Fragments only.

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
2. why_fortune — Fortune's authority and reach framed around THIS client's sector and goal; \
body is 1-3 sentences of prose
3. product × 1-2 — each product by what it does and who it reaches, tied to the client's \
target audience; eyebrow "RECOMMENDED PRODUCTS"
4. proof — performance stats or audience data from reference slides; eyebrow "PERFORMANCE \
DATA"; omit this slide entirely if no relevant data exists in the reference slides
5. investment — clean pricing tiers using real figures from the rate card; eyebrow \
"INVESTMENT"; pricing appears ONLY here
6. next_steps — 3 concrete actions with implied timeline; eyebrow "NEXT STEPS"\
"""


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        seed_key: str | None = None,
        rate_sheet_key: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._seed_key = seed_key or os.environ["PPTX_SEED_DECK_KEY"]
        self._rate_sheet_key = rate_sheet_key or os.environ.get("RATE_SHEET_KEY")
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._seed_bytes: bytes | None = None
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
    def _pick_layout(prs: PresentationType, layout_map: dict, *names: str):
        for name in names:
            if name in layout_map:
                return layout_map[name]
        return prs.slide_layouts[2]

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
    def _apply_brand_bg(slide) -> None:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0x10, 0x18, 0x5F)

    @staticmethod
    def _set_text_white(slide) -> None:
        white = RGBColor(0xFF, 0xFF, 0xFF)
        for ph in slide.placeholders:
            if not ph.has_text_frame:
                continue
            for para in ph.text_frame.paragraphs:
                para.font.color.rgb = white
                for run in para.runs:
                    run.font.color.rgb = white

    @staticmethod
    def _add_bullets(slide) -> None:
        for ph in slide.placeholders:
            if ph.placeholder_format.idx != 1 or not ph.has_text_frame:
                continue
            for para in ph.text_frame.paragraphs:
                if not para.text.strip():
                    continue
                p_pr = para._p.get_or_add_pPr()
                for tag in (qn("a:buNone"), qn("a:buChar"), qn("a:buAutoNum")):
                    existing = p_pr.find(tag)
                    if existing is not None:
                        p_pr.remove(existing)
                bu_char = etree.SubElement(p_pr, qn("a:buChar"))
                bu_char.set("char", "—")

    @staticmethod
    def _prune_placeholders(slide, keep: frozenset = frozenset({0, 1})) -> None:
        sp_tree = slide.shapes._spTree
        for ph in list(slide.placeholders):
            if ph.placeholder_format.idx not in keep:
                sp_tree.remove(ph._element)

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

    @staticmethod
    def _populate_why_fortune_slide(slide, slide_data: dict) -> None:
        # 7_Title Only: idx 17 = eyebrow (0.35in), idx 0 = title (0.84in), idx 1 = body (1.29in)
        if slide.shapes.title:
            slide.shapes.title.text = slide_data.get("title", "")
        ph_map = {ph.placeholder_format.idx: ph for ph in slide.placeholders if ph.has_text_frame}
        body_ph = ph_map.get(1)
        if body_ph:
            body_ph.text_frame.clear()
            body_ph.text_frame.paragraphs[0].text = slide_data.get("body", "")
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0x10, 0x18, 0x5F)
        white = RGBColor(0xFF, 0xFF, 0xFF)
        for ph in slide.placeholders:
            if ph.has_text_frame:
                for para in ph.text_frame.paragraphs:
                    para.font.color.rgb = white
                    for run in para.runs:
                        run.font.color.rgb = white
        DeckGenerator._prune_placeholders(slide, frozenset({0, 1}))

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
        self._apply_brand_bg(slide)
        self._set_text_white(slide)
        self._add_bullets(slide)
        keep = frozenset({0, 1, 20}) if eyebrow else frozenset({0, 1})
        self._prune_placeholders(slide, keep)

    def _load_seed(self) -> bytes:
        if self._seed_bytes is None:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._seed_key)
            self._seed_bytes = resp["Body"].read()
            logger.info("loaded seed deck from s3://%s/%s", self._bucket, self._seed_key)
        return self._seed_bytes

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
        seed = self._load_seed()
        prs = Presentation(io.BytesIO(seed))

        layout_map = {layout.name: layout for layout in prs.slide_layouts}
        self._clear_seed_slides(prs)

        for slide_data in slides:
            slide_type = slide_data.get("slide_type", "product")
            if slide_type == "title":
                layout = self._pick_layout(prs, layout_map, "COVER blue option", "Title Slide")
                slide = prs.slides.add_slide(layout)
                self._populate_title_slide(slide, slide_data)
            elif slide_type == "why_fortune":
                layout = self._pick_layout(prs, layout_map, "7_Title Only", "Title Only")
                slide = prs.slides.add_slide(layout)
                self._populate_why_fortune_slide(slide, slide_data)
            else:
                # product, proof, investment, next_steps
                layout = self._pick_layout(
                    prs, layout_map,
                    "8_Title and Content", "2_Title and Content", "Title and Content",
                )
                slide = prs.slides.add_slide(layout)
                self._populate_content_slide(slide, slide_data)

        buf = io.BytesIO()
        prs.save(buf)
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
