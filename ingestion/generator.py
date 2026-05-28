import io
import json
import logging
import os
import uuid

import anthropic
import boto3
from pptx import Presentation

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"
_GENERATED_PREFIX = "generated"

_SYSTEM_PROMPT = """\
You are a sales deck writer for Fortune Media Group. Write like a senior AE, not a consultant.

Return ONLY a JSON array — no prose, no markdown fences. Each element is one slide object:
- Title slide (first): {"slide_type": "title", "title": "...", "subtitle": "..."}
- Content slides:      {"slide_type": "content", "title": "...", "bullets": ["...", "..."]}

Rules:
- 6-10 slides maximum. Cut anything that doesn't directly serve the brief.
- Bullets are fragments, not sentences. Start with a number or a noun, never a verb phrase.
- 2-4 bullets per slide. If you need more, split the slide.
- No filler: ban "leverage", "synergy", "best-in-class", "cutting-edge", "robust", "seamless"
- No throat-clearing: no "overview", "introduction", or "in conclusion" slides
- Be specific — pull real details from the reference slides. If you have no specific detail, \
omit the point entirely rather than generalizing.
- Every slide must earn its place. If you can't say what it adds, cut it.\
"""


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        seed_key: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._seed_key = seed_key or os.environ["PPTX_SEED_DECK_KEY"]
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._seed_bytes: bytes | None = None
        self._api_key: str | None = None

    def _load_seed(self) -> bytes:
        if self._seed_bytes is None:
            resp = self._s3.get_object(Bucket=self._bucket, Key=self._seed_key)
            self._seed_bytes = resp["Body"].read()
            logger.info("loaded seed deck from s3://%s/%s", self._bucket, self._seed_key)
        return self._seed_bytes

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

        user_msg = f"Brief:\n{brief}\n\nReference slides from Fortune corpus:\n{context_text}"

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

        # Remove all existing slides while preserving slide masters/layouts
        xml_slides = prs.slides._sldIdLst
        for sld in list(xml_slides):
            xml_slides.remove(sld)

        layout_map = {layout.name: layout for layout in prs.slide_layouts}

        def _pick_layout(preferred: str, fallback_idx: int):
            return layout_map.get(preferred) or prs.slide_layouts[fallback_idx]

        for slide_data in slides:
            slide_type = slide_data.get("slide_type", "content")

            if slide_type == "title":
                layout = _pick_layout("Title Slide", 0)
                slide = prs.slides.add_slide(layout)
                for ph in slide.placeholders:
                    idx = ph.placeholder_format.idx
                    if idx == 0:
                        ph.text = slide_data.get("title", "")
                    elif idx == 1:
                        ph.text = slide_data.get("subtitle", "")
            else:
                layout = _pick_layout("Title and Content", 1)
                slide = prs.slides.add_slide(layout)
                for ph in slide.placeholders:
                    idx = ph.placeholder_format.idx
                    if idx == 0:
                        ph.text = slide_data.get("title", "")
                    elif idx == 1:
                        tf = ph.text_frame
                        tf.clear()
                        bullets = slide_data.get("bullets", [])
                        for i, bullet in enumerate(bullets):
                            if i == 0:
                                tf.paragraphs[0].text = bullet
                                tf.paragraphs[0].level = 0
                            else:
                                p = tf.add_paragraph()
                                p.text = bullet
                                p.level = 0

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
