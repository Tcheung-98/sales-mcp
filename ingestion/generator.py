import io
import json
import logging
import os
import re
import uuid

import anthropic
import boto3
from docx import Document
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"
_GENERATED_PREFIX = "generated"
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
Return ONLY a JSON array — no prose, no markdown fences. Each element is one slide.

The first slide is always the cover:
  {{"action": "cover", "title": "...", "subtitle": "..."}}

All other slides clone a real Fortune corpus slide by coordinate:
  {{"action": "clone", "source_path": "...", "slide_number": 7,
   "replacements": {{"title": "...", "client_name": "Acme Corp"}}}}

FIELD RULES
- source_path and slide_number must come exactly from the candidate slides listed in the \
user message — do not invent coordinates
- replacements.client_name: always set to the client name from the brief; the engine will \
replace "your brand" / "your company" / "your organization" in the cloned slide's text
- replacements.title: set only if you want to override the corpus slide's existing title; \
omit to keep the original
- replacements.eyebrow: set only if you want to override the corpus slide's eyebrow label; \
omit to keep the original

SELECTION RULES
- Select 6–12 slides total including the cover
- Prefer variety — draw from multiple source decks where candidates allow
- Follow the arc, escalation, and voice rules in the rulebook above
- The user message contains: (1) a client brief, (2) Fortune's rate card, \
(3) candidate corpus slides with their coordinates — use all three\
"""


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
            doc = Document(io.BytesIO(resp["Body"].read()))
            self._rulebook_text = "\n".join(
                p.text for p in doc.paragraphs if p.text.strip()
            )
            logger.info("loaded rulebook from s3://%s/%s", self._bucket, self._rulebook_key)
        return self._rulebook_text

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

    def _call_claude(self, brief: str, context_slides: list[dict]) -> list[dict]:
        context_text = "\n\n".join(
            f"[source: {s.get('source_path', '')} | slide: {s.get('slide_number', '?')}]\n"
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
    def _apply_replacements(slide, replacements: dict) -> None:
        if not replacements:
            return

        title = replacements.get("title")
        eyebrow = replacements.get("eyebrow")
        client_name = replacements.get("client_name")

        # title / eyebrow: overwrite their placeholder slots
        if title or eyebrow:
            ph_map = {
                ph.placeholder_format.idx: ph
                for ph in slide.placeholders
                if ph.has_text_frame
            }
            if title and (ph := ph_map.get(0)):
                ph.text_frame.clear()
                ph.text_frame.paragraphs[0].text = title
            if eyebrow and (ph := ph_map.get(19)):
                ph.text_frame.clear()
                ph.text_frame.paragraphs[0].text = eyebrow

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
        blank = self._load_blank()
        blank_prs = Presentation(io.BytesIO(blank))   # cover slide source
        output_prs = Presentation(io.BytesIO(blank))  # output (inherits Fortune master/theme)

        # Drop relationship AND sldId element for each seed slide
        for sld_id in list(output_prs.slides._sldIdLst):
            r_id = sld_id.get(qn("r:id"))
            output_prs.part.drop_rel(r_id)
            output_prs.slides._sldIdLst.remove(sld_id)

        for slide_data in slides:
            action = slide_data.get("action", "clone")
            if action == "cover":
                slide = self._clone_slide(blank_prs, 0, output_prs)
                # COVER blue option has no idx 0: idx 12 = headline, idx 10 = subtitle.
                ph_map = {
                    ph.placeholder_format.idx: ph
                    for ph in slide.placeholders
                    if ph.has_text_frame
                }
                for title_idx in (12, 0):
                    if ph := ph_map.get(title_idx):
                        ph.text_frame.clear()
                        ph.text_frame.paragraphs[0].text = slide_data.get("title", "")
                        break
                for sub_idx in (10, 1):
                    if ph := ph_map.get(sub_idx):
                        ph.text_frame.clear()
                        ph.text_frame.paragraphs[0].text = slide_data.get("subtitle", "")
                        break
            elif action == "clone":
                source_prs = self._load_pptx(slide_data["source_path"])
                slide = self._clone_slide(source_prs, slide_data["slide_number"] - 1, output_prs)
                self._apply_replacements(slide, slide_data.get("replacements", {}))
            else:
                raise ValueError(f"Unknown slide action: {action!r}")

        buf = io.BytesIO()
        output_prs.save(buf)
        return buf.getvalue()

    def generate(self, brief: str, context_slides: list[dict]) -> dict:
        logger.info("generating deck for brief: %.80s", brief)
        slides = self._call_claude(brief, context_slides)
        logger.info("claude returned %d slides", len(slides))
        pptx_bytes = self._build_pptx(slides)
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
        return {
            "s3_uri": f"s3://{self._bucket}/{key}",
            "download_url": url,
            "slide_count": len(slides),
            "brief": brief,
        }
