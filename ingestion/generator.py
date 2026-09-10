import io
import json
import logging
import os
import re
import shutil
import time
from urllib.parse import urlparse
from uuid import uuid4

import anthropic
import boto3
import requests
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.presentation import Presentation as PresentationType

from ingestion.audience_data import AudienceData, load_audience_data_from_s3
from ingestion.category_dividers import (
    FORTUNEAI_DIVIDER_SLIDE_INDEX,
    FORTUNEAI_MIN_SLIDES,
    FORTUNEAI_TEMPLATE_BASENAME,
    fortuneai_template_key,
    is_fortuneai_template_url,
)
from ingestion.deck_qa import DeckQaError, run_deterministic_qa
from ingestion.deck_qa_agent import (
    CursorQaReport,
    QaIssue,
    resolve_timeout_s,
    run_headless_cursor_qa,
)
from ingestion.gtm_product_map import (
    GtmProductMap,
    ProductSlideRef,
    load_gtm_product_map_from_s3,
    product_deck_s3_key,
)
from ingestion.placeholder_ai import PlaceholderAI
from ingestion.placeholder_fills import apply_placeholders, fetch_logo_bytes
from ingestion.pptx_tools import delete_slide, sync_sections
from ingestion.render_slides import RenderSlidesError
from ingestion.review_package import (
    ReviewPackage,
    build_review_package,
    plan_pitch_sequence,
)
from ingestion.schema import DeckSchema

logger = logging.getLogger(__name__)

_SECRET_NAME = "fortune-sales-mcp/claude-api-key"
_DEFAULT_MODEL = "claude-sonnet-4-6"

_TEMPLATE_URL_HOST_SUFFIXES = (".sharepoint.com", ".sharepoint.us", ".microsoft.com")

# Headless QA gate (docs/DECK-QA-ARCHITECTURE.md §8). QA is always on: every
# build_deck call runs B2→B3→B4. The two env vars below are bypasses, not
# feature flags — local dev speed and emergency ops only, never set in prod.
DECK_QA_DISABLED_ENV = "DECK_QA_DISABLED"
DECK_QA_SKIP_VISION_ENV = "DECK_QA_SKIP_VISION"
_DECK_QA_TRUTHY = {"1", "true", "yes"}
REVIEW_PACKAGE_PREFIX = "review-packages"
# Payload warning when the rail is bypassed — the caller must see the deck is un-QA'd.
DECK_QA_BYPASSED_WARNING = (
    f"deck QA was bypassed via {DECK_QA_DISABLED_ENV}; "
    "this deck shipped without its quality gate"
)

_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _template_key_from_url(template_url: str) -> str:
    return template_url.split("?")[0].rstrip("/").split("/")[-1]


def _env_truthy(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in _DECK_QA_TRUTHY


def deck_qa_disabled() -> bool:
    """Emergency/local-dev bypass of the whole QA rail (§8). Never set in prod."""
    return _env_truthy(DECK_QA_DISABLED_ENV)


def deck_qa_vision_skipped() -> bool:
    """Local-dev bypass of B4 only: B2+B3 still run, no Cursor key needed (§8)."""
    return _env_truthy(DECK_QA_SKIP_VISION_ENV)


class DeckGenerator:
    def __init__(
        self,
        bucket: str | None = None,
        secret_name: str = _SECRET_NAME,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        self._bucket = bucket or os.environ["S3_SNAPSHOT_BUCKET"]
        self._secret_name = secret_name
        self._model = model
        self._s3 = boto3.client("s3")
        self._pptx_cache: dict[str, PresentationType] = {}
        self._api_key: str | None = None
        self._gtm_product_map: GtmProductMap | None = None
        self._audience_data: AudienceData | None = None

    @staticmethod
    def _import_ctx(target_prs: PresentationType) -> dict:
        """Per-target bookkeeping for parts pulled in from other packages.

        ``next_partname`` only sees parts already reachable from the package
        root, so track allocations ourselves; ``layouts`` keeps repeat clones of
        one source layout from importing a fresh master each time.
        """
        ctx = getattr(target_prs, "_clone_import_ctx", None)
        if ctx is None:
            ctx = {"package": target_prs.part.package, "taken": set(), "layouts": {}}
            setattr(target_prs, "_clone_import_ctx", ctx)
        ctx["taken"] |= {part.partname for part in ctx["package"].iter_parts()}
        return ctx

    @staticmethod
    def _alloc_partname(ctx: dict, source_partname) -> PackURI:
        stem = source_partname.filename.rsplit(".", 1)[0]
        match = re.match(r"[A-Za-z_]*", stem)
        prefix = (match.group() if match is not None else "") or "part"
        n = 1
        while True:
            candidate = PackURI(
                f"{source_partname.baseURI}/{prefix}{n}.{source_partname.ext}"
            )
            if candidate not in ctx["taken"]:
                ctx["taken"].add(candidate)
                return candidate
            n += 1

    @staticmethod
    def _importable_part(target_prs: PresentationType, source_part):
        """Return a binary part safe to relate into ``target_prs``.

        Product decks and FortuneAI share media partnames (``image63.png``,
        ``hdphoto40.wdp``, …). Relating the source part directly would emit two
        zip entries under one partname, so copy the blob to a free partname
        whenever the name is already taken.
        """
        ctx = DeckGenerator._import_ctx(target_prs)
        if source_part.partname not in ctx["taken"]:
            ctx["taken"].add(source_part.partname)
            return source_part
        return Part(
            DeckGenerator._alloc_partname(ctx, source_part.partname),
            source_part.content_type,
            ctx["package"],
            source_part.blob,
        )

    @staticmethod
    def _copy_xml_part(ctx: dict, source_part) -> XmlPart:
        # Keep the source's own part class (SlideLayoutPart, SlideMasterPart, …)
        # so the copy still answers python-pptx's object-model traversals. Parts
        # python-pptx has no class for — themes — load blob-backed, so read them
        # as generic XML instead.
        source_cls = type(source_part)
        part_cls = source_cls if issubclass(source_cls, XmlPart) else XmlPart
        return part_cls(
            DeckGenerator._alloc_partname(ctx, source_part.partname),
            source_part.content_type,
            ctx["package"],
            parse_xml(source_part.blob),
        )

    @staticmethod
    def _existing_layout_ids(target_prs: PresentationType) -> list[int]:
        """All ``sldLayoutId/@id`` values already in the package (must be unique)."""
        ids: list[int] = []
        for master in target_prs.slide_masters:
            layout_id_lst = master._element.find(qn("p:sldLayoutIdLst"))
            if layout_id_lst is None:
                continue
            for layout_id_el in layout_id_lst.findall(qn("p:sldLayoutId")):
                raw = layout_id_el.get("id")
                if raw is None:
                    continue
                try:
                    ids.append(int(raw))
                except ValueError:
                    pass
        return ids

    @staticmethod
    def _import_slide_layout(target_prs: PresentationType, source_layout):
        """Copy ``source_layout`` — plus its master and theme — into ``target_prs``.

        Cloned product slides keep placeholders, and placeholders inherit font,
        size, colour and bullet formatting from their layout/master/theme. Point
        them at a FortuneAI layout instead and the slide silently restyles, so
        bring the real chain along. The copied master is pruned to just this
        layout, which keeps the import to roughly the layout's own weight.
        """
        ctx = DeckGenerator._import_ctx(target_prs)
        cache_key = id(source_layout.part)
        if cache_key in ctx["layouts"]:
            return ctx["layouts"][cache_key]

        src_layout_part = source_layout.part
        src_master_part = source_layout.slide_master.part

        master_part = DeckGenerator._copy_xml_part(ctx, src_master_part)
        master_map: dict[str, str] = {}
        for rId, rel in src_master_part.rels.items():
            # Layout list is rebuilt below so the copy carries only what we need.
            if rel.reltype == RT.SLIDE_LAYOUT:
                continue
            if rel.is_external:
                master_map[rId] = master_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            elif rel.reltype == RT.THEME:
                master_map[rId] = master_part.relate_to(
                    DeckGenerator._import_theme(target_prs, ctx, rel.target_part),
                    rel.reltype,
                )
            else:
                master_map[rId] = master_part.relate_to(
                    DeckGenerator._importable_part(target_prs, rel.target_part),
                    rel.reltype,
                )

        layout_part = DeckGenerator._copy_xml_part(ctx, src_layout_part)
        layout_map: dict[str, str] = {}
        for rId, rel in src_layout_part.rels.items():
            if rel.reltype == RT.SLIDE_MASTER:
                layout_map[rId] = layout_part.relate_to(master_part, rel.reltype)
            elif rel.is_external:
                layout_map[rId] = layout_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                layout_map[rId] = layout_part.relate_to(
                    DeckGenerator._importable_part(target_prs, rel.target_part),
                    rel.reltype,
                )
        DeckGenerator._remap_rids(layout_part._element, layout_map)

        DeckGenerator._remap_rids(master_part._element, master_map)
        layout_id_lst = master_part._element.find(qn("p:sldLayoutIdLst"))
        if layout_id_lst is None:
            raise ValueError(
                f"slide master {src_master_part.partname} has no sldLayoutIdLst"
            )
        for child in list(layout_id_lst):
            layout_id_lst.remove(child)
        layout_ids = DeckGenerator._existing_layout_ids(target_prs)
        new_layout_id = max(layout_ids, default=2147483648) + 1
        layout_id_lst.append(
            parse_xml(
                f'<p:sldLayoutId xmlns:p="{_PML_NS}" xmlns:r="{_RELS_NS}" '
                f'id="{new_layout_id}" '
                f'r:id="{master_part.relate_to(layout_part, RT.SLIDE_LAYOUT)}"/>'
            )
        )

        master_id_lst = target_prs.part._element.find(qn("p:sldMasterIdLst"))
        existing = [int(e.get("id")) for e in master_id_lst if e.get("id")]
        master_id_lst.append(
            parse_xml(
                f'<p:sldMasterId xmlns:p="{_PML_NS}" xmlns:r="{_RELS_NS}" '
                f'id="{max(existing, default=2147483648) + 1}" '
                f'r:id="{target_prs.part.relate_to(master_part, RT.SLIDE_MASTER)}"/>'
            )
        )

        ctx["layouts"][cache_key] = layout_part
        return layout_part

    @staticmethod
    def _import_theme(target_prs: PresentationType, ctx: dict, source_theme_part):
        theme_part = DeckGenerator._copy_xml_part(ctx, source_theme_part)
        theme_map: dict[str, str] = {}
        for rId, rel in source_theme_part.rels.items():
            if rel.is_external:
                theme_map[rId] = theme_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                theme_map[rId] = theme_part.relate_to(
                    DeckGenerator._importable_part(target_prs, rel.target_part),
                    rel.reltype,
                )
        DeckGenerator._remap_rids(theme_part._element, theme_map)
        return theme_part

    @staticmethod
    def _remap_rids(element, rId_map: dict[str, str]) -> None:
        """Rewrite every relationship-namespace attribute (r:embed, r:id, r:link, …).

        Single pass over attributes: sequential string replacement would corrupt
        the mapping whenever a new rId collides with an old one still unwritten.
        """
        prefix = f"{{{_RELS_NS}}}"
        for el in element.iter():
            for attr in list(el.keys()):
                if not attr.startswith(prefix):
                    continue
                new_rId = rId_map.get(el.get(attr))
                if new_rId is not None:
                    el.set(attr, new_rId)

    @staticmethod
    def _clone_slide(
        source_prs: PresentationType, slide_idx: int, target_prs: PresentationType
    ):
        source_slide = source_prs.slides[slide_idx]

        # Add a placeholder slide — gives us a proper slide part + sldIdLst entry
        new_slide = target_prs.slides.add_slide(target_prs.slide_layouts[0])

        # Carry every source relationship except the layout (rewired below) and
        # notes (deliberately dropped). Copying only images leaves hyperlinks,
        # hdphoto sidecars and media as dangling r:ids, which makes PowerPoint
        # declare the deck corrupt and "repair" it on open.
        rId_map: dict[str, str] = {}
        for rId, rel in source_slide.part.rels.items():
            if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
                continue
            if rel.is_external:
                new_rId = new_slide.part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                new_rId = new_slide.part.relate_to(
                    DeckGenerator._importable_part(
                        target_prs, source_slide.part.related_part(rId)
                    ),
                    rel.reltype,
                )
            rId_map[rId] = new_rId

        # Serialize source cSld and re-parse with pptx element classes (parse_xml,
        # not etree.fromstring, so spTree and other pptx attrs are available)
        fixed_cSld = parse_xml(etree.tostring(source_slide._element.cSld))
        DeckGenerator._remap_rids(fixed_cSld, rId_map)

        # Swap target slide's cSld with the fixed clone
        tgt_sld = new_slide._element
        tgt_sld.replace(tgt_sld.cSld, fixed_cSld)

        # shapes is a lazyproperty cached during add_slide; invalidate so next access
        # gets a fresh SlideShapes pointing at the new spTree, not the discarded one
        new_slide.__dict__.pop("shapes", None)

        # Point the clone at the source slide's own layout, imported into this
        # package. Matching a FortuneAI layout by name instead would leave the
        # slide's placeholders inheriting the wrong fonts, bullets and colours.
        layout_part = DeckGenerator._import_slide_layout(
            target_prs, source_slide.slide_layout
        )
        for rId, rel in list(new_slide.part.rels.items()):
            if rel.reltype == RT.SLIDE_LAYOUT:
                new_slide.part.drop_rel(rId)
                break
        new_slide.part.relate_to(layout_part, RT.SLIDE_LAYOUT)

        return new_slide

    @staticmethod
    def _delete_slide(prs: PresentationType, slide_idx: int) -> None:
        delete_slide(prs, slide_idx)

    @staticmethod
    def _renumber_slide_parts(prs: PresentationType) -> None:
        """Assign contiguous slide partnames before save (avoids OPC duplicate warnings)."""
        r_ids = [sld_id.get(qn("r:id")) for sld_id in prs.slides._sldIdLst]
        prs.part.rename_slide_parts(r_ids)

    @staticmethod
    def _insert_slide_at(prs: PresentationType, position: int) -> None:
        # _clone_slide always appends; move the last sldId to the target position
        sld_id_lst = prs.slides._sldIdLst
        sld_id = sld_id_lst[-1]
        sld_id_lst.remove(sld_id)
        sld_id_lst.insert(position, sld_id)

    @staticmethod
    def _validate_template_url(template_url: str) -> None:
        parsed = urlparse(template_url)
        if parsed.scheme != "https":
            raise ValueError("template_url must use HTTPS")
        host = (parsed.hostname or "").lower()
        if not host:
            raise ValueError("template_url must include a host")
        extra_hosts = {
            h.strip().lower()
            for h in os.environ.get("TEMPLATE_URL_ALLOWED_HOSTS", "").split(",")
            if h.strip()
        }
        if host in extra_hosts or any(
            host.endswith(suffix) for suffix in _TEMPLATE_URL_HOST_SUFFIXES
        ):
            return
        raise ValueError(f"template_url host not allowed: {host}")

    def _load_pptx(self, s3_key: str) -> PresentationType:
        if s3_key not in self._pptx_cache:
            resp = self._s3.get_object(Bucket=self._bucket, Key=s3_key)
            self._pptx_cache[s3_key] = Presentation(io.BytesIO(resp["Body"].read()))
            logger.info("loaded pptx from s3://%s/%s", self._bucket, s3_key)
        return self._pptx_cache[s3_key]

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

        def _sid(el) -> int:
            try:
                return int(el.get("id", 0))
            except (ValueError, TypeError):
                return 0

        own = list(slide._element.iter(qn("p:cNvPr")))
        # Reserve ids further down this slide too, otherwise a replacement id can
        # collide with a shape we have not visited yet and duplicate it.
        reserved = used | {_sid(el) for el in own}
        next_id = max(reserved, default=0) + 1
        taken: set[int] = set()
        for el in own:
            sid = _sid(el)
            if sid == 0 or sid in used or sid in taken:
                while next_id in reserved:
                    next_id += 1
                el.set("id", str(next_id))
                reserved.add(next_id)
                taken.add(next_id)
                next_id += 1
            else:
                taken.add(sid)

    def _get_gtm_product_map(self) -> GtmProductMap:
        if self._gtm_product_map is None:
            self._gtm_product_map = load_gtm_product_map_from_s3(self._s3, self._bucket)
        return self._gtm_product_map

    def _get_audience_data(self) -> AudienceData:
        if self._audience_data is None:
            self._audience_data = load_audience_data_from_s3(self._s3, self._bucket)
        return self._audience_data

    def _load_fortuneai_template(
        self, template_url: str | None
    ) -> tuple[PresentationType, str]:
        """Load FortuneAI_DeckTemplate from SharePoint URL or S3.

        Returns (presentation, template_key used in build payload).
        Always returns a fresh Presentation — assembly mutates the deck.
        """
        if template_url:
            self._validate_template_url(template_url)
            if not is_fortuneai_template_url(template_url):
                raise ValueError(
                    "template_url must point to FortuneAI_DeckTemplate "
                    f"(got {template_url.split('?', 1)[0]!r})"
                )
            resp = requests.get(template_url, timeout=30)
            resp.raise_for_status()
            return Presentation(io.BytesIO(resp.content)), FORTUNEAI_TEMPLATE_BASENAME

        key = fortuneai_template_key()
        try:
            resp = self._s3.get_object(Bucket=self._bucket, Key=key)
            data = resp["Body"].read()
        except Exception as exc:
            raise ValueError(
                f"Failed to load FortuneAI template from "
                f"s3://{self._bucket}/{key}: {exc}"
            ) from exc
        logger.info("loaded FortuneAI template from s3://%s/%s", self._bucket, key)
        return Presentation(io.BytesIO(data)), FORTUNEAI_TEMPLATE_BASENAME

    def _clone_product_ref(
        self, ref: ProductSlideRef, target_prs: PresentationType
    ):
        s3_key = product_deck_s3_key(ref.deck_path)
        try:
            source_prs = self._load_pptx(s3_key)
        except Exception as exc:
            raise ValueError(
                f"Failed to load Deck Path {ref.deck_path!r} for product "
                f"{ref.product_name!r} (s3://{self._bucket}/{s3_key}): {exc}"
            ) from exc
        if ref.slide_number > len(source_prs.slides):
            raise ValueError(
                f"Slide # {ref.slide_number} out of range for product "
                f"{ref.product_name!r} in {ref.deck_path!r} "
                f"({len(source_prs.slides)} slides)"
            )
        new_slide = self._clone_slide(source_prs, ref.slide_number - 1, target_prs)
        self._dedup_shape_ids(new_slide, target_prs)
        return new_slide

    def assemble_skeleton(
        self,
        schema: DeckSchema,
        template_url: str | None = None,
        product_map: GtmProductMap | None = None,
    ) -> PresentationType:
        """Assemble FortuneAI spine + conditional dividers + exact GTM product clones.

        Keeps intro/narrative/investment/thank-you stock (C2 fills in ``build``). Inserts
        category dividers only when ≥1 funded product maps to that section. Product
        pages are wholesale A5 clones — no AI edits.
        """
        gtm_map = product_map if product_map is not None else self._get_gtm_product_map()
        # Resolve placement before mutating the template so map misses fail loud early.
        # plan_pitch_sequence is the one implementation of pitch order (§5); the
        # review-package manifest is built from the same plan.
        plan = plan_pitch_sequence(schema, gtm_map)
        prs, _template_key = self._load_fortuneai_template(template_url)

        if len(prs.slides) < FORTUNEAI_MIN_SLIDES:
            raise ValueError(
                f"FortuneAI_DeckTemplate must have at least {FORTUNEAI_MIN_SLIDES} "
                f"slides (intro/narrative + dividers + investment + thank you); "
                f"got {len(prs.slides)}"
            )

        # Trim optional tail after thank-you (e.g. stock slide 20).
        while len(prs.slides) > FORTUNEAI_MIN_SLIDES:
            self._delete_slide(prs, len(prs.slides) - 1)

        # Divider slides in the template file are not in Workflow pitch order.
        # Clone dividers from a pristine copy, drop all stock dividers, then
        # insert funded sections in Workflow order (13→17) before investment.
        divider_src_prs, _ = self._load_fortuneai_template(template_url)

        for idx in sorted(FORTUNEAI_DIVIDER_SLIDE_INDEX, reverse=True):
            self._delete_slide(prs, idx)

        insert_at = len(prs.slides) - 2  # before investment + thank you
        for item in reversed(plan):
            match item:
                case ("divider", category_index):
                    src_idx = FORTUNEAI_DIVIDER_SLIDE_INDEX[category_index]
                    self._clone_slide(divider_src_prs, src_idx, prs)
                case ("product", ref):
                    self._clone_product_ref(ref, prs)
            self._insert_slide_at(prs, insert_at)

        sync_sections(prs)
        return prs

    def _upload_review_package(self, package: ReviewPackage, prefix: str) -> None:
        """Copy the review package to ``s3://{bucket}/review-packages/{uuid}/``.

        Ops note: review packages are ephemeral (Hard Rule 8). Expire them with a
        7–30 day lifecycle rule on the ``review-packages/`` prefix in the bucket —
        the final PPTX is the deliverable, this is inspection material. An upload
        failure must not mask the QA verdict, so it is logged and swallowed.
        """
        try:
            for path in sorted(package.root.rglob("*")):
                if not path.is_file():
                    continue
                self._s3.put_object(
                    Bucket=self._bucket,
                    Key=prefix + path.relative_to(package.root).as_posix(),
                    Body=path.read_bytes(),
                )
        except Exception:
            logger.warning(
                "failed to upload review package to s3://%s/%s",
                self._bucket,
                prefix,
                exc_info=True,
            )
            return
        logger.info("review package uploaded to s3://%s/%s", self._bucket, prefix)

    def _run_deck_qa(
        self,
        prs: PresentationType,
        schema: DeckSchema,
        gtm_map: GtmProductMap,
    ) -> tuple[PresentationType, dict]:
        """Run B2 + B3 + B4 under one ``DECK_QA_TIMEOUT_S`` budget (§8).

        Returns the deck to deliver — re-loaded from disk when B4 saved fixes,
        otherwise the one passed in — plus the ``qa`` block for the build payload.
        Raises DeckQaError on a B3 failure, a B4 ``passed: false``, or a timeout:
        a deck whose quality gate did not complete is not a deliverable. The
        review package is uploaded to S3 either way, so a timed-out run can be
        inspected after the fact.
        """
        review_id = uuid4()
        prefix = f"{REVIEW_PACKAGE_PREFIX}/{review_id}/"
        budget = resolve_timeout_s()
        deadline = time.monotonic() + budget

        # Packaging renumbers slide parts before saving draft.pptx, so B3/B4 get a
        # well-formed deck and build()'s own _renumber_slide_parts is a no-op (§8).
        try:
            package = build_review_package(
                prs, schema, plan=plan_pitch_sequence(schema, gtm_map)
            )
        except RenderSlidesError as exc:
            # LibreOffice/poppler failure is infrastructure, but the gate did not
            # complete, so nothing ships (§8). RenderSlidesError is a RuntimeError;
            # unwrapped it would escape server.py's handlers as an unstructured crash.
            report = CursorQaReport(
                passed=False,
                issues=[
                    QaIssue(
                        None,
                        "error",
                        f"review package could not be built (slide render): {exc}",
                    )
                ],
            )
            raise DeckQaError(
                f"Deck QA could not render slides for review (review {review_id}): "
                f"{exc}",
                report=report,
            ) from exc
        try:
            det_report = run_deterministic_qa(prs, schema, package.manifest)
            package.write_deterministic_report(det_report)
            if not det_report.passed:
                raise DeckQaError(det_report.summary(), report=det_report)

            if deck_qa_vision_skipped():
                logger.warning(
                    "B4 vision QA skipped via %s (review %s) — local dev only",
                    DECK_QA_SKIP_VISION_ENV,
                    review_id,
                )
                return prs, {
                    "deterministic_passed": True,
                    "cursor_passed": None,
                    "vision_skipped": True,
                    "review_package_key": prefix,
                }

            remaining = deadline - time.monotonic()
            if remaining <= 0:
                # B2 + B3 spent the whole budget; B4 never started.
                report = CursorQaReport(
                    passed=False,
                    issues=[
                        QaIssue(
                            None,
                            "error",
                            f"deck QA spent its {budget:g}s budget on packaging and "
                            "deterministic checks; the vision pass never started",
                        )
                    ],
                )
                package.write_cursor_report(report)
                raise DeckQaError(
                    f"Headless deck QA did not finish within {budget:g}s "
                    f"(review {review_id}); the deck was not delivered — retry, or "
                    "raise DECK_QA_TIMEOUT_S if renders are legitimately slow",
                    report=report,
                )

            try:
                agent_report = run_headless_cursor_qa(package, timeout_s=remaining)
            except ValueError as exc:
                report = CursorQaReport(
                    passed=False,
                    issues=[QaIssue(None, "error", str(exc))],
                )
                package.write_cursor_report(report)
                raise DeckQaError(
                    f"Headless deck QA could not start (review {review_id}): {exc}",
                    report=report,
                ) from exc
            # The runner cancels itself at its budget and reports the timeout as an
            # error issue, so a timed-out B4 fails here like any other non-pass:
            # the quality gate did not complete, so nothing ships (§8).
            if not agent_report.passed:
                raise DeckQaError(agent_report.summary(), report=agent_report)
            if agent_report.fixes_applied:
                # B4 edits draft.pptx on disk while we hold our own
                # Presentation; without this re-load the fixes never reach
                # the uploaded deck (§8). An edit the agent failed to report
                # arrives here too — the runner reconciles fixes_applied
                # against the draft's digest.
                prs = Presentation(str(package.draft_path))

            return prs, {
                "deterministic_passed": True,
                "cursor_passed": True,
                "review_package_key": prefix,
            }
        finally:
            self._upload_review_package(package, prefix)
            shutil.rmtree(package.root, ignore_errors=True)

    def build(
        self,
        schema: DeckSchema,
        template_url: str | None = None,
        product_map: GtmProductMap | None = None,
        audience_data: AudienceData | None = None,
        logo_bytes: bytes | None = None,
    ) -> dict:
        """Assemble FortuneAI PPTX, fill placeholders, QA, upload.

        template_url: optional pre-authenticated SharePoint download URL for
        FortuneAI_DeckTemplate. When omitted, loads from S3 (FORTUNEAI_TEMPLATE_KEY).
        Product pages use exact GTM Deck Path / Slide #. C2 fills date, logo,
        history, audience Reach/Index, program types, investment, and bounded
        Claude copy for intro, Opportunity, audience title, and program blurbs.

        The deck then goes through the QA rail (§8) by default: review package →
        deterministic checks → headless Cursor vision pass, with a ``qa`` block in
        the payload. A timeout fails loud — an unreviewed deck is not delivered.
        DECK_QA_DISABLED bypasses the rail (emergency/local only, with a warning);
        DECK_QA_SKIP_VISION runs B2+B3 but skips B4 (local dev).
        """
        prs = self.assemble_skeleton(schema, template_url, product_map=product_map)
        audience = (
            audience_data if audience_data is not None else self._get_audience_data()
        )
        logo = (
            logo_bytes if logo_bytes is not None else fetch_logo_bytes(schema.client_logo)
        )
        ai = PlaceholderAI.from_anthropic(
            anthropic.Anthropic(api_key=self._get_api_key()),
            model=self._model,
        )
        warnings = apply_placeholders(
            prs, schema, audience=audience, logo_bytes=logo, ai=ai
        )
        qa_block = None
        if deck_qa_disabled():
            logger.warning(DECK_QA_BYPASSED_WARNING)
            warnings.append(DECK_QA_BYPASSED_WARNING)
        else:
            gtm_map = (
                product_map if product_map is not None else self._get_gtm_product_map()
            )
            prs, qa_block = self._run_deck_qa(prs, schema, gtm_map)
        self._renumber_slide_parts(prs)
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
        payload = {
            "download_url": url,
            "slide_count": len(prs.slides),
            "client_name": schema.client_name,
            "template_key": FORTUNEAI_TEMPLATE_BASENAME,
            "warnings": warnings,
        }
        if qa_block is not None:
            payload["qa"] = qa_block
        return payload

