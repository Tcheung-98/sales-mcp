# Deck QA architecture

> **Status:** MVP-required (rescopes shelved PI-2754 vision stylist).  
> **Scope:** `sales-mcp` deck generation only. Upstream product lock lives in Pitch Deck Builder (PI-2350).  
> **Related:** [`END-SCOPE-SOT.md`](../local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md) · [`pitch-deck-end-scope.mdc`](../.cursor/rules/pitch-deck-end-scope.mdc)

---

## 1. Problem

`build_deck` produces a **structurally correct** FortuneAI PPTX after C1 assembly and C2 placeholder fills, but output is often **not associate-sendable** without a final polish pass on wording and layout. The deck QA rail is the quality gate between C2 completion and delivery.

---

## 2. Pipeline

```text
Locked DeckSchema (complete payload from upstream)
  → C1 assemble_skeleton (FortuneAI + dividers + A5 clones)
  → C2 apply_placeholders (deterministic + bounded Claude on named stock slots)
  → B2 review package (draft.pptx + PNGs + manifest.json)
  → B3 deterministic QA (fail fast, no vision)
  → B4 vision QA (rules + one fix loop)
  → final.pptx upload + presigned URL
```

The QA rail runs inside `sales-mcp` when `DECK_QA_ENABLED` is set. Any MCP client that calls `build_deck` with a locked schema can use it; no specific upstream client is required.

---

## 3. Components

| Piece | Module | Role |
|-------|--------|------|
| C1 assembly | `ingestion/generator.py` → `assemble_skeleton` | FortuneAI spine + Workflow-order dividers + exact GTM clones |
| C2 fills | `ingestion/placeholder_fills.py`, `placeholder_ai.py` | Stock-slide tokens and bounded narrative copy; never edits product clones |
| A5 map | `ingestion/gtm_product_map.py` | Exact `Deck Path` + `Slide #` lookup |
| B2 package | `ingestion/review_package.py` | `draft.pptx`, PNGs, `manifest.json`, `deck_schema.json` |
| B3 checks | `ingestion/deck_qa.py` | Deterministic structural checks before vision |
| B4 vision | `ingestion/deck_qa_agent.py` | Vision review + one fix pass on editable slides |
| Slide render | `ingestion/render_slides.py` | LibreOffice + `pdftoppm` (150 DPI default) |
| Manifest types | `ingestion/manifest.py` | `ReviewManifest`, `SlideManifestEntry` |
| MCP entry | `server.py` → `build_deck` | Validates schema, runs pipeline, returns presigned URL |

Variant deletion (`ingestion/placeholders.py` → `delete_unused_variants`) runs inside C2 and determines the post-C2 slide geometry documented in §5.

---

## 4. Hard rules

1. **No product clone edits** — slides with `role: product` in the manifest may be flagged but never modified.
2. **No Titan/RAG slide pick** — do not use `search_decks` to substitute product slides.
3. **No mix changes** — QA cannot add, remove, or reorder products or dividers.
4. **No product ideation** — sales-mcp does not propose or rank products (`propose_mix` does not exist).
5. **Fail loud** — missing GTM row, budget mismatch, fewer than two audience segments: error before or during build, not silent fix.
6. **Workflow divider pitch order** — High-Impact → Editorial → Premium Video → Print → Branded Content. Physical template indices are not sequential; see `FORTUNEAI_DIVIDER_SLIDE_INDEX` in `ingestion/category_dividers.py`.
7. **Max one vision fix loop** per deck (B4). A second failure returns `status: error` with `qa_report`.
8. **Review packages are ephemeral** — S3 prefix with short lifecycle (7–30 days); the final PPTX is the deliverable.

### Editable slides in B4

| `role` | May fix | Method |
|--------|---------|--------|
| `cover` | Yes | Bounded `PlaceholderAI` slots; `set_ph_text` / `replace_token` |
| `narrative` | Yes | Same; includes Opportunity, audience title, program blurbs |
| `other` | Yes | Dividers, investment, thank-you — text replace only |
| `product` | **No** | Flag only in `qa_report.issues[]` |

Use `replace_token`, `replace_first_token`, or `set_ph_text` on FortuneAI stock slides. Do **not** use `pptx_tools.apply_replacements` — it keys off placeholder `idx` 0 and 19 and was written for Titan-retrieved corpus slides.

---

## 5. Review package (B2)

`ingestion/review_package.py` → `build_review_package()` writes:

```text
review-packages/{review_id}/
  draft.pptx
  manifest.json
  deck_schema.json
  slides/
    slide-000.png
    slide-001.png
    ...
  qa_deterministic.json    # written by B3
  qa_cursor.json           # written by B4
```

### Manifest contract

Each `SlideManifestEntry` in `manifest.json` includes:

| Field | Purpose |
|-------|---------|
| `slide_index` | 0-based index in the post-C2 deck |
| `role` | `cover` · `narrative` · `product` · `other` |
| `slide_kind` | For `other`: `divider` · `investment` · `thank_you` |
| `editable` | `false` on A5 product clones |
| `product_name`, `source_path`, `source_slide_number` | A5 provenance when `role: product` |

Example (partial):

```json
{
  "schema_version": "1",
  "client_name": "Acme Corp",
  "template_key": "FortuneAI_DeckTemplate.pptx",
  "slide_count": 10,
  "slides": [
    {"slide_index": 0, "role": "cover", "editable": true},
    {"slide_index": 7, "role": "product", "editable": false,
     "product_name": "CEO Daily",
     "source_path": "Fortune_Newsletters_2026.pptx",
     "source_slide_number": 3}
  ]
}
```

### Slide geometry (post-C2)

The manifest reflects the deck **after** C2 deletes unused audience/program variants.

```text
FortuneAI_DeckTemplate.pptx     19 slides (template)
after C1 assemble_skeleton      14 + P slides
after C2 apply_placeholders      8 + P slides
```

`P = funded_divider_count + len(confirmed_products)`.

| Final index | Slide | `role` | `editable` |
|-------------|-------|--------|------------|
| `0` | Intro / cover | `cover` | `true` |
| `1`–`5` | Narrative spine (5 surviving pages) | `narrative` | `true` |
| `6` … `5 + P` | Dividers + A5 clones (Workflow order) | `other` / `product` | `true` / **`false`** |
| `6 + P` | Investment | `other` | `true` |
| `7 + P` | Thank you | `other` | `true` |

**Important:** indices `1–11` in the *template* are not all `narrative` in the final deck. After C2, only indices `1–5` are narrative; index `6+` are dividers and product clones. Marking product clones as `narrative` would set `editable: true` and violate Hard Rule 1.

Single-product newsletter example (`P = 2`, 10 slides total): `tests/smoke_build_live.py` and `tests/test_generator.py`.

### Pitch sequence and provenance

`plan_pitch_sequence(schema, gtm_map)` in `ingestion/review_package.py` is the single source of pitch-section order. `assemble_skeleton` calls the same planner so assembly and manifest never disagree.

`source_path` in the manifest is the raw GTM `Deck Path` (e.g. `Fortune_Newsletters_2026.pptx`), not the S3 key. `product_deck_s3_key()` adds the `product-decks/` prefix at load time.

`build_review_package` renumbers slide parts before saving `draft.pptx` so A5 clones do not produce duplicate `ppt/slides/slideNN.xml` zip entries.

---

## 6. Deterministic QA (B3)

Module: `ingestion/deck_qa.py` → `run_deterministic_qa(prs, schema, manifest)`.

| Check | Fail condition |
|-------|----------------|
| Slide count | `len(prs.slides) != manifest.slide_count` |
| Leftover tokens | Any unfilled token from the C2 constant list survives in slide text |
| Investment budget | Stated tier ≠ sum of `confirmed_products.price` |
| Product presence | Each confirmed product name appears in deck text |
| Divider order | Funded dividers in Workflow order |
| Tail slides | Last two slides are investment + thank-you |
| PPTX integrity | Save and re-open without exception |
| Manifest consistency | Every `slide_index` in range |

Import token literals from `ingestion/placeholder_fills` and `ingestion/pptx_tools` rather than duplicating strings. Scan all slides including `role: product` clones.

B4 does not run when B3 fails.

---

## 7. Vision QA (B4)

Module: `ingestion/deck_qa_agent.py` → `run_headless_cursor_qa(package, *, timeout_s)`.

CLI shim: `scripts/run_deck_qa.py`.

### Inputs

- Review package directory (local or S3)
- `CURSOR_API_KEY` when `DECK_QA_ENABLED` is truthy
- `.cursor/skills/deck-qa/SKILL.md` — vision checklist and fix rules

### Responsibilities

1. Read `manifest.json` and `qa_deterministic.json`
2. Vision-review PNGs for **editable** slides only
3. Evaluate leftover placeholders, alignment, readability, off-brand issues, weak copy on allowed slots
4. Pass, or apply **one** fix pass using the methods in §4
5. Write `qa_cursor.json`:

```json
{
  "passed": true,
  "loop_count": 1,
  "issues": [{"slide_index": 4, "severity": "warning", "message": "..."}],
  "fixes_applied": ["opportunity_body"]
}
```

B4 is not exposed as an associate-facing MCP tool.

---

## 8. `build_deck` integration

When `DECK_QA_ENABLED` is `1`, `true`, or `yes`:

```text
apply_placeholders (C2)
  → build_review_package (B2)
  → run_deterministic_qa (B3) — raises on failure
  → run_headless_cursor_qa (B4)
  → re-load draft.pptx if B4 applied fixes (fixes are written on disk)
  → renumber, save, upload
```

### Response shape

```json
{
  "download_url": "...",
  "slide_count": 16,
  "warnings": [],
  "qa": {
    "deterministic_passed": true,
    "cursor_passed": true,
    "timed_out": false,
    "review_package_key": "review-packages/{uuid}/"
  }
}
```

On QA failure, `build_deck` returns `status: error` with `qa_report`. Define `DeckQaError` to carry the report through `server.py`.

### Environment

| Variable | Default | Meaning |
|----------|---------|---------|
| `DECK_QA_ENABLED` | `false` | When unset, `build()` skips the rail entirely |
| `DECK_QA_TIMEOUT_S` | `600` | Wall-clock budget for B2+B3+B4 combined |
| `CURSOR_API_KEY` | — | Required when QA is enabled |

### Timeout policy

`build_deck` stays **synchronous**. With QA enabled, a single call may run LibreOffice conversion, PNG rasterization, and a vision pass — plausibly several minutes.

On timeout: **fail soft** — upload the deck as-is, set `qa.timed_out: true` and `qa.cursor_passed: false`, append a warning to `warnings[]`. A timeout is infrastructure, not a quality verdict.

B3 failures and an explicit B4 `passed: false` still **raise** (`DeckQaError`).

---

## 9. Out of scope

- Upstream Pitch Deck Builder / product lock UI (PI-2350)
- Rewriting A5 product slides
- Three-tier investment column layout
- `search_decks` / Titan in the Creation path
- SharePoint upload of review packages
- More than one vision fix loop

---

## 10. Tests

```bash
# Manifest and render
PYTHONPATH=. uv run pytest tests/test_manifest.py tests/test_render_slides.py -q

# Assembly
PYTHONPATH=. uv run pytest tests/test_generator.py -k assemble -q

# Review package and deterministic QA
PYTHONPATH=. uv run pytest tests/test_review_package.py tests/test_deck_qa.py -q

# Golden geometry
PYTHONPATH=. uv run pytest tests/test_bamboohr_golden.py -q

# Live smoke (S3 + mock AI)
PYTHONPATH=. uv run python tests/smoke_build_live.py --mock-ai
```

### Golden test guidance

Prefer offline fixtures using product names from `tests/logic_guide_fixtures.py` (`Crown Unit`, `CEO Daily`, `Term Sheet`, `Long-Form Article`, `Full Page`) with a fixture `GtmProductMap`. The assertion that matters is `slide_count == 8 + P` per §5, not specific dollar figures.

If testing against live S3/GTM, confirm product names and prices against `GTM_DATABASE_KEY` and the inventory Pricing sheet before asserting end-to-end behavior.
