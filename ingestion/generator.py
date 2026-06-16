import io
import json
import logging
import os
import re

import boto3
from docx import Document
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType
from pypdf import PdfReader

from ingestion.schema import DeckSchema

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


_VOICE_RULES = """\
You are a senior AE at Fortune Media Group writing a custom pitch deck for a specific client. \
Your job is to persuade, not to describe. Every word you write should make the client feel that \
Fortune is the obvious partner for their business goals.

=== FORTUNE SALES SKILL — RULEBOOK ===
{rulebook_text}
=== END RULEBOOK ===

VOICE RULES — apply to every word you write
- Write like a pitch, not a spec sheet. Lead with what the client gains, not what the product does.
- Each bullet: one client benefit + one proof point. 18 words max. Fragments are fine.
- 3 bullets per slide maximum. Cut the weakest if you have 4.
- Translate specs into outcomes. "100% SOV" becomes "CrowdStrike owns every impression." \
"42% unique open rate" becomes "42 of every 100 readers opens this. Above every benchmark."
- Pricing lives on the investment slide only. Never mention $ or cost on any other slide.
- Address the client directly and specifically. "CrowdStrike's buyers" not "your target audience." \
"The CISOs Fortune reaches" not "decision makers."
- Never use: em dashes, "leverage", "synergy", "best-in-class", "cutting-edge", "seamless", \
"robust", "drive engagement", "unlock", "elevate".\
"""


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        rulebook_key: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._rulebook_key = rulebook_key or os.environ.get(
            "RULEBOOK_KEY", "templates/rulebook.docx"
        )
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._pptx_cache: dict[str, PresentationType] = {}
        self._rulebook_text: str | None = None
        self._api_key: str | None = None

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
    def _delete_slide(prs: PresentationType, slide_idx: int) -> None:
        sld_id_lst = prs.slides._sldIdLst
        sld_id = sld_id_lst[slide_idx]
        r_id = sld_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)

    @staticmethod
    def _insert_slide_at(prs: PresentationType, position: int) -> None:
        # _clone_slide always appends; move the last sldId to the target position
        sld_id_lst = prs.slides._sldIdLst
        sld_id = sld_id_lst[-1]
        sld_id_lst.remove(sld_id)
        sld_id_lst.insert(position, sld_id)

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

    def build(self, schema: DeckSchema, template_url: str) -> dict:
        """Build a deck by populating the approved template with client-specific copy.

        template_url: pre-authenticated URL to the actual .pptx file, resolved and
        provided by Prodie from the Fortune Sales Automation SharePoint folder.
        Returns a presigned S3 URL valid for 24h.
        """
        raise NotImplementedError("template-populate build not yet implemented")

