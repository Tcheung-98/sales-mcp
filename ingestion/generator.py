import io
import json
import logging
import os
import re
from uuid import uuid4

import anthropic
import boto3
import requests
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
_DEFAULT_MODEL = "claude-opus-4-8"
_COST_PER_M_INPUT = 3.00
_COST_PER_M_OUTPUT = 15.00

# Corpus slides use these strings as stand-ins for the client name.
# All are matched case-insensitively so "Your Brand" and "YOUR BRAND" both swap.
_CLIENT_NAME_PLACEHOLDERS = re.compile(
    r"your brand|your company|client name|your organization",
    re.IGNORECASE,
)


_WEAK_MATCH_THRESHOLD = 0.5
_PRODUCT_SECTION_LANDMARK = "2_DARK BLUE/ BRIGHT BLUE CAPSULE"
_PRODUCT_LAYOUT = "11_Title Only"

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


_WRITE_TOOL: dict = {
    "name": "write_deck_copy",
    "description": "Output replacement copy for every slide in the deck.",
    "input_schema": {
        "type": "object",
        "properties": {
            "slides": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "slide_index": {"type": "integer"},
                        "title": {"type": "string"},
                        "eyebrow": {"type": "string"},
                        "body": {"type": "array", "items": {"type": "string"}, "maxItems": 3},
                        "client_name": {"type": "string"},
                    },
                    "required": ["slide_index", "title", "body", "client_name"],
                },
            }
        },
        "required": ["slides"],
    },
}

_QA_TOOL: dict = {
    "name": "review_deck_copy",
    "description": "Report which slides have copy issues. Only flag slides with genuine problems.",
    "input_schema": {
        "type": "object",
        "properties": {
            "approved": {"type": "boolean"},
            "failing_slides": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "slide_index": {"type": "integer"},
                        "issues": {"type": "array", "items": {"type": "string"}},
                    },
                    "required": ["slide_index", "issues"],
                },
            },
        },
        "required": ["approved", "failing_slides"],
    },
}

_REVIEW_MODEL = "claude-sonnet-4-6"
_MAX_REVISIONS = 3
_TITLE_MAX_CHARS = 80
_BODY_LINE_MAX_CHARS = 120


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

    def _call_claude_review(self, brief: str, slides: list[dict]) -> str:
        lines = [
            "You are reviewing a Fortune Media pitch deck before writing copy.",
            "",
            "CLIENT BRIEF:",
            brief,
            "",
            f"DECK STRUCTURE ({len(slides)} slides):",
        ]
        for s in slides:
            lines.append(f"\nSlide {s['slide_index'] + 1} [{s['layout_name']}]")
            if s.get("title"):
                lines.append(f"  Title: {s['title']}")
            for i, b in enumerate(s.get("body_text") or [], start=1):
                lines.append(f"  Body {i}: {b}")
        lines += [
            "",
            "Read the full deck and brief. Write a concise arc brief (200-300 words) covering:",
            "- The core story arc for this specific client",
            "- What each slide section needs to accomplish",
            "- Tone and emphasis for this buyer and industry",
            "",
            "This arc brief will guide copy writing for every slide.",
        ]
        client = anthropic.Anthropic(api_key=self._get_api_key())
        response = client.messages.create(
            model=self._model,
            max_tokens=1024,
            messages=[{"role": "user", "content": "\n".join(lines)}],
        )
        return response.content[0].text

    def _call_claude_write(
        self,
        brief: str,
        arc_context: str,
        slides: list[dict],
        issues: dict[int, list[str]] | None = None,
    ) -> list[dict]:
        rulebook = self._load_rulebook()
        system = _VOICE_RULES.format(rulebook_text=rulebook)
        lines = [
            "CLIENT BRIEF:",
            brief,
            "",
            "NARRATIVE ARC:",
            arc_context,
            "",
            f"SLIDES TO REWRITE ({len(slides)} slides):",
        ]
        for s in slides:
            lines.append(f"\nSlide {s['slide_index'] + 1} [{s['layout_name']}]")
            if s.get("title"):
                lines.append(f"  Current title: {s['title']}")
            for i, b in enumerate(s.get("body_text") or [], start=1):
                lines.append(f"  Current body {i}: {b}")
            if issues and (slide_issues := issues.get(s["slide_index"])):
                lines.append(f"  ISSUES TO FIX: {'; '.join(slide_issues)}")
        lines += [
            "",
            "Rewrite every slide using the write_deck_copy tool.",
        ]
        client = anthropic.Anthropic(api_key=self._get_api_key())
        response = client.messages.create(
            model=self._model,
            max_tokens=4096,
            system=system,
            messages=[{"role": "user", "content": "\n".join(lines)}],
            tools=[_WRITE_TOOL],
            tool_choice={"type": "tool", "name": "write_deck_copy"},
        )
        for block in response.content:
            if block.type == "tool_use" and block.name == "write_deck_copy":
                return block.input["slides"]
        raise ValueError("Claude did not return deck copy")

    def _call_claude_qa(
        self, brief: str, slides_content: list[dict], replacements: list[dict]
    ) -> dict:
        rep_map = {r["slide_index"]: r for r in replacements}
        lines = [
            "You are a senior editor reviewing AI-written copy for a Fortune Media pitch deck.",
            "",
            "CLIENT BRIEF:",
            brief,
            "",
            f"GENERATED DECK COPY ({len(slides_content)} slides):",
        ]
        for s in slides_content:
            idx = s["slide_index"]
            rep = rep_map.get(idx, {})
            lines.append(f"\nSlide {idx + 1} [{s['layout_name']}]")
            lines.append(f"  Title: {rep.get('title', '')}")
            if rep.get("eyebrow"):
                lines.append(f"  Eyebrow: {rep['eyebrow']}")
            for b in rep.get("body") or []:
                lines.append(f"  • {b}")
        lines += [
            "",
            "Flag only slides with genuine problems: narrative gaps, wrong tone for the buyer,",
            "missing critical product information, or copy that contradicts the brief.",
            "Set approved=true if the deck is ready to send. Use the review_deck_copy tool.",
        ]
        client = anthropic.Anthropic(api_key=self._get_api_key())
        response = client.messages.create(
            model=_REVIEW_MODEL,
            max_tokens=1024,
            messages=[{"role": "user", "content": "\n".join(lines)}],
            tools=[_QA_TOOL],
            tool_choice={"type": "tool", "name": "review_deck_copy"},
        )
        for block in response.content:
            if block.type == "tool_use" and block.name == "review_deck_copy":
                return block.input
        raise ValueError("Sonnet QA did not return review output")

    @staticmethod
    def _overflow_flags(replacements: list[dict]) -> list[dict]:
        failures = []
        for r in replacements:
            slide_issues = []
            title = r.get("title") or ""
            if len(title) > _TITLE_MAX_CHARS:
                slide_issues.append(
                    f"title is {len(title)} chars (max {_TITLE_MAX_CHARS})"
                )
            for line in r.get("body") or []:
                if len(line) > _BODY_LINE_MAX_CHARS:
                    slide_issues.append(
                        f"body line too long ({len(line)} chars, max {_BODY_LINE_MAX_CHARS}):"
                        f" {line[:40]}…"
                    )
            if slide_issues:
                failures.append({"slide_index": r["slide_index"], "issues": slide_issues})
        return failures

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

    def build(self, schema: DeckSchema, template_url: str, retriever) -> dict:
        """Fetch template, swap product slides from corpus, write copy via Claude, upload to S3.

        template_url: pre-authenticated URL to the .pptx template, resolved by Prodie.
        retriever: SlideRetriever instance for corpus search.
        Returns presigned S3 URL valid for 24h.
        """
        # Fetch template
        resp = requests.get(template_url, timeout=30)
        resp.raise_for_status()
        prs = Presentation(io.BytesIO(resp.content))

        # Detect product slides: 11_Title Only layout after the capsule landmark
        landmark_idx = next(
            (i for i, s in enumerate(prs.slides)
             if s.slide_layout.name == _PRODUCT_SECTION_LANDMARK),
            None,
        )
        product_indices = [
            i for i, s in enumerate(prs.slides)
            if landmark_idx is not None
            and i > landmark_idx
            and s.slide_layout.name == _PRODUCT_LAYOUT
        ]
        insertion_index = product_indices[0] if product_indices else len(prs.slides)

        # Delete existing product slides back-to-front to preserve indices
        for idx in sorted(product_indices, reverse=True):
            self._delete_slide(prs, idx)

        # Corpus search — collect best candidate per product, fail fast on weak matches
        failures = []
        best_candidates = []
        for product in schema.confirmed_products:
            candidates = retriever.search(product.name, k=8)
            best_score = max((c["score"] for c in candidates), default=0.0)
            if best_score < _WEAK_MATCH_THRESHOLD:
                failures.append({"product": product.name, "best_score": best_score})
            else:
                best_candidates.append(max(candidates, key=lambda c: c["score"]))

        if failures:
            raise ValueError(
                "No good corpus match for: "
                + ", ".join(f"{f['product']} (score {f['best_score']:.2f})" for f in failures)
            )

        # Clone product slides from corpus into deck at the insertion point
        for i, best in enumerate(best_candidates):
            source_prs = self._load_pptx(best["source_path"])
            new_slide = self._clone_slide(source_prs, best["slide_number"] - 1, prs)
            self._insert_slide_at(prs, insertion_index + i)
            self._dedup_shape_ids(new_slide, prs)

        # Extract slide contents for Claude
        slides_content = []
        for i, slide in enumerate(prs.slides):
            title = ""
            body_text = []
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                try:
                    ph_idx = shape.placeholder_format.idx
                except ValueError:
                    continue
                text = shape.text_frame.text.strip()
                if not text:
                    continue
                if ph_idx == 0:
                    title = text
                else:
                    body_text.append(text)
            slides_content.append({
                "slide_index": i,
                "layout_name": slide.slide_layout.name,
                "title": title,
                "body_text": body_text,
            })

        # Build client brief string
        brief_lines = [
            f"Client: {schema.client_name}",
            f"Industry: {schema.industry}",
            f"Quarterly budget: ${schema.budget_quarterly:,.0f}",
        ]
        if schema.buyer_persona:
            brief_lines.append(f"Buyer persona: {schema.buyer_persona}")
        if schema.objective:
            brief_lines.append(f"Objective: {schema.objective}")
        if schema.target_audience:
            brief_lines.append(f"Target audience: {schema.target_audience}")
        if schema.tone_notes:
            brief_lines.append(f"Tone: {schema.tone_notes}")
        brief_lines.append(f"\nConfirmed products ({len(schema.confirmed_products)}):")
        for p in schema.confirmed_products:
            brief_lines.append(f"  - {p.name} | {p.category} | {p.cadence} | ${p.price:,.0f}")
        if schema.upsell:
            u = schema.upsell
            brief_lines.append(f"\nUpsell: {u.name} | {u.category} | ${u.price:,.0f}")
        if schema.next_steps_contact:
            brief_lines.append(f"Next steps contact: {schema.next_steps_contact}")
        brief = "\n".join(brief_lines)

        # Claude: arc review → Opus writes all copy → Sonnet QA loop → apply
        arc_context = self._call_claude_review(brief, slides_content)
        replacements = self._call_claude_write(brief, arc_context, slides_content)
        replacement_map = {r["slide_index"]: r for r in replacements}

        for revision in range(_MAX_REVISIONS):
            issue_map: dict[int, list[str]] = {}
            for f in self._overflow_flags(list(replacement_map.values())):
                issue_map.setdefault(f["slide_index"], []).extend(f["issues"])
            qa = self._call_claude_qa(brief, slides_content, list(replacement_map.values()))
            for f in qa["failing_slides"]:
                issue_map.setdefault(f["slide_index"], []).extend(f["issues"])
            if not issue_map:
                break
            logger.info("QA pass %d: revising %d slide(s)", revision + 1, len(issue_map))
            failing_slides = [s for s in slides_content if s["slide_index"] in issue_map]
            revised = self._call_claude_write(
                brief, arc_context, failing_slides, issues=issue_map
            )
            for r in revised:
                replacement_map[r["slide_index"]] = r

        for i, slide in enumerate(prs.slides):
            if i in replacement_map:
                self._apply_replacements(slide, replacement_map[i])

        # Upload and return presigned URL
        buf = io.BytesIO()
        prs.save(buf)
        key = f"generated/{uuid4()}.pptx"
        self._s3.put_object(Bucket=self._bucket, Key=key, Body=buf.getvalue())
        url = self._s3.generate_presigned_url(
            "get_object",
            Params={"Bucket": self._bucket, "Key": key},
            ExpiresIn=86400,
        )
        logger.info("deck uploaded to s3://%s/%s", self._bucket, key)
        return {
            "download_url": url,
            "slide_count": len(prs.slides),
            "client_name": schema.client_name,
            "template_key": template_url.split("?")[0].rstrip("/").split("/")[-1],
        }

