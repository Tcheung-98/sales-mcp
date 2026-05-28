import io
import json
import logging
import os
import uuid

import anthropic
import boto3
from pptx import Presentation
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"
_GENERATED_PREFIX = "generated"

_SYSTEM_PROMPT = """\
You are a senior AE at Fortune Media Group building a custom pitch deck.

Return ONLY a JSON array — no prose, no markdown fences. Each element is one slide object:
- Title slide (first): {"slide_type": "title", "title": "...", "subtitle": "..."}
- Content slides:      {"slide_type": "content", "title": "...", "bullets": ["...", "..."]}

You will receive: (1) a client brief, (2) Fortune's product & rate card, \
(3) reference slides from past decks.

How to use each input:
- Brief: defines the client's vertical, audience, and goal — everything must serve THIS client
- Rate card: your primary source for product names, pricing, cadence, and audience fit — \
pick 3-6 products that genuinely match the brief's vertical and target audience, \
cite real product names and real pricing figures
- Reference slides: style and structure inspiration only — do not copy their generic content

Deck structure (in order):
1. Title slide — client-specific headline, not "Fortune Media Pitch"
2. Why Fortune — 1 slide, specific reach numbers or editorial authority relevant to their sector
3. Recommended products — 1-2 slides naming specific Fortune products from the rate card \
with cadence and pricing
4. Audience match — 1 slide showing Fortune's audience aligns with their target (use \
Audience Alignment and Contextual Alignment from the rate card)
5. Investment — 1 slide with actual pricing tiers from the rate card for the recommended mix
6. Next steps — 1 slide, specific and actionable

Writing rules:
- Bullets are fragments. Start with a number, a product name, or a noun — never a verb phrase.
- 2-4 bullets per slide maximum.
- No filler words: ban "leverage", "synergy", "best-in-class", "cutting-edge", "seamless"
- Every bullet must contain a specific fact, name, or number. Cut anything vague.\
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

        logger.info(
            "claude usage — input: %d tokens, output: %d tokens",
            response.usage.input_tokens,
            response.usage.output_tokens,
        )
        raw = response.content[0].text.strip()
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

        # Proper slide deletion: drop both the relationship AND the sldId element
        for sld_id in list(prs.slides._sldIdLst):
            rId = sld_id.get(qn("r:id"))
            prs.part.drop_rel(rId)
            prs.slides._sldIdLst.remove(sld_id)

        layout_map = {layout.name: layout for layout in prs.slide_layouts}

        def _pick_layout(*names: str):
            for name in names:
                if name in layout_map:
                    return layout_map[name]
            return prs.slide_layouts[2]

        def _set_title(slide, text: str) -> None:
            if slide.shapes.title:
                slide.shapes.title.text = text

        def _set_body(slide, bullets: list[str]) -> None:
            # Clear ALL non-title placeholders first so labels like "INSERT EYEBROW" don't show
            body_ph = None
            for ph in slide.placeholders:
                if ph.placeholder_format.idx == 0:
                    continue
                if ph.has_text_frame:
                    ph.text_frame.clear()
                    # Body placeholder (idx=1) gets the bullets; others stay blank
                    if ph.placeholder_format.idx == 1:
                        body_ph = ph
            if body_ph is None:
                # Fallback: use the first non-title text placeholder we can find
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx != 0 and ph.has_text_frame:
                        body_ph = ph
                        break
            if body_ph is None:
                return
            tf = body_ph.text_frame
            for i, bullet in enumerate(bullets):
                if i == 0:
                    tf.paragraphs[0].text = bullet
                else:
                    p = tf.add_paragraph()
                    p.text = bullet

        for slide_data in slides:
            slide_type = slide_data.get("slide_type", "content")

            if slide_type == "title":
                layout = _pick_layout("Title Slide", "COVER blue option")
                slide = prs.slides.add_slide(layout)
                _set_title(slide, slide_data.get("title", ""))
                _set_body(slide, [slide_data.get("subtitle", "")])
            else:
                layout = _pick_layout(
                    "TITLE+CONTENT_1-LINE", "1_TITLE+CONTENT_1-LINE",
                    "2_Title and Content", "Title and Content",
                )
                slide = prs.slides.add_slide(layout)
                _set_title(slide, slide_data.get("title", ""))
                _set_body(slide, slide_data.get("bullets", []))

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
