# Deck QA architecture (headless Cursor styling gate)

> **Status:** MVP-required (rescopes shelved PI-2754).  
> **Audience:** Subagents, reviewers, and implementers. **This file is SoT** for the QA rail — do not invent APIs or file paths not listed here.  
> **Repo:** `sales-mcp` only. Prodie wire (PI-2350) is a separate track.  
> **End-scope (Creation):** `local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md` (tracked) + `.cursor/rules/pitch-deck-end-scope.mdc`

> ### ⚠️ Base branch: `fix/fortuneai-deck-assembly`, not `main`
>
> Everything in this document describes the deck geometry on **`fix/fortuneai-deck-assembly`** (open PR #31). `origin/main` does **not** have `FORTUNEAI_DIVIDER_SLIDE_INDEX` — it still has only `FORTUNEAI_FIRST_DIVIDER_INDEX = 12` and a different divider-insert path, so the slide indices in §5 and the ordering rule in §4.6 are wrong against `main`.
>
> **All Wave 1–3 branches must be cut from `fix/fortuneai-deck-assembly`** (or from `main` only after PR #31 merges). Do not "rebase onto main before opening PR" until #31 is in.

---

## 1. Problem statement

`build_deck` today produces a **structurally correct** FortuneAI PPTX (C1 assembly + C2 fills + A5 product clones) but output is **not associate-sendable** without styling/wording polish. **Headless Cursor QA is a non-negotiable MVP gate** between C2 completion and delivery.

---

## 2. MVP definition (updated)

```text
Locked DeckSchema (from Prodie + confirm_mix / I3)
  → C1 assemble_skeleton (FortuneAI + dividers + A5 clones)
  → C2 apply_placeholders (deterministic + bounded Claude on named stock slots)
  → B2 review package (draft.pptx + PNGs + manifest.json)
  → B3 deterministic QA (fail fast, no vision)
  → B4 headless Cursor agent (vision + rules; max 1 fix loop)
  → final.pptx upload + presigned URL
```

**Prodie does not run Cursor.** `sales-mcp` (or a co-located worker it calls) runs the QA rail.

---

## 3. What already exists (do not reimplement)

| Asset | Location | Notes |
|-------|----------|-------|
| FortuneAI assembly | `ingestion/generator.py` → `assemble_skeleton`, `build` | Workflow divider order via `FORTUNEAI_DIVIDER_SLIDE_INDEX` in `ingestion/category_dividers.py` |
| C2 fills | `ingestion/placeholder_fills.py`, `ingestion/placeholder_ai.py` | Does **not** edit product clones |
| A5 map | `ingestion/gtm_product_map.py` | Exact `Deck Path` + `Slide #` only |
| Slide → PNG | `ingestion/render_slides.py` | LibreOffice + `pdftoppm`; CLI `python -m ingestion.render_slides` |
| Manifest schema (stub) | `ingestion/manifest.py` | `ReviewManifest`, `SlideManifestEntry`; roles: `cover`, `narrative`, `product`, `other` only |
| Variant deletion | `ingestion/placeholders.py` → `delete_unused_variants` | Called **inside** `apply_placeholders`; drops 4 audience + 2 program pages. Drives §5 geometry |
| Budget vs mix guard | `placeholder_fills._assert_budget_matches_mix` | Already raises in C2 — see §6 note |
| MCP entry | `server.py` → `build_deck` | Returns presigned URL; docstring still says stylist shelved — update in integration PR |
| Docker LO | `Dockerfile` | `libreoffice-impress`, `poppler-utils`, fonts already installed |
| Clone fidelity tests | `tests/test_clone_fidelity.py` | Mechanical regressions |
| Live smoke pattern | `tests/smoke_build_live.py` | Acme + CEO Daily; `--mock-ai` |
| BambooHR fixture needs | Prior smoke | ≥2 audience segments in `targeting_details`; mix total must equal stated Tier 1 budget |

**Does not exist yet (planned work):**

- `ingestion/review_package.py` — build package from `Presentation` + schema
- `ingestion/deck_qa.py` — deterministic checks
- `scripts/` — the directory itself does not exist; `scripts/run_deck_qa.py` is the headless Cursor SDK runner
- `.cursor/skills/` — does not exist; `.cursor/skills/deck-qa/SKILL.md` is the vision-QA instruction set
- Wiring in `DeckGenerator.build()` or post-build hook

**Exists on a branch, not yet merged — read before starting PR-A:**

`origin/PI-2522-review-package-builder` (commit `48845b4`, "feat(PI-2522): return `AssembledSkeleton` with product-clone provenance") already changes `assemble_skeleton` to return provenance, and is what the `ingestion/manifest.py` module docstring is describing. On `fix/fortuneai-deck-assembly` that change is **not present**: `assemble_skeleton` returns a bare `Presentation` and the docstring claim is stale.

This matters because product-clone provenance is the only reliable source for `product_name` / `source_path` / `source_slide_number` in the manifest, and it lives in `generator.py`, which §13 reserves for PR-F. See §5 "Where provenance comes from" for the resolution.

---

## 4. Hard rules (never violate)

1. **No product clone edits** — slides with `role: product` in manifest: no text/layout/image changes in QA or fixes.
2. **No Titan/RAG slide pick** — no `search_decks` for substitute product slides.
3. **No mix changes** — QA cannot add/remove/reorder products or dividers.
4. **No `propose_mix` MCP tool** — not part of associate runtime.
5. **Fail loud** — missing GTM row, budget mismatch, `<2` audience segments: error before or during build, not silent fix.
6. **Workflow divider pitch order** — High-Impact → Editorial → Premium Video → Print → Branded Content (category index 0–4). Physical template indices are **not** sequential; see `FORTUNEAI_DIVIDER_SLIDE_INDEX`.
7. **Max 1 Cursor fix loop** per deck (B4). Second failure → `status: error` with `qa_report`.
8. **Review packages are ephemeral** — S3 prefix with short lifecycle (7–30 days); final PPTX is the deliverable.

### Editable slide roles in B4 (vision QA + fixes)

| `role` | May fix styling/wording | Method |
|--------|-------------------------|--------|
| `cover` | Yes | Re-run bounded `PlaceholderAI` slots; `set_ph_text` / `replace_token` |
| `narrative` | Yes (stock FortuneAI only) | Same; includes Opportunity, audience title, program blurbs |
| `other` | Yes (dividers, investment, thank-you) | Deterministic fills + text replace; no layout invention |
| `product` | **No** | Flag only in `qa_report.issues[]` |

**Use `replace_token` / `replace_first_token` / `set_ph_text`, not `apply_replacements`.** `pptx_tools.apply_replacements` keys off placeholder `idx` 0 and 19 and rewrites body paragraphs positionally — it was written for Titan-retrieved corpus slides. FortuneAI stock slides carry named text tokens in ordinary text boxes, so `apply_replacements` will either no-op or clobber the wrong paragraphs.

---

## 5. Review package contract (B2)

Directory layout (local or S3 `review-packages/{uuid}/`):

```text
review-packages/{review_id}/
  draft.pptx          # Post-C2, pre-QA (or post-QA as final.pptx on success)
  manifest.json       # ReviewManifest v1
  deck_schema.json    # Serialized DeckSchema (for re-fill if needed)
  slides/
    slide-000.png     # 0-based index, 150 DPI default (render_slides)
    slide-001.png
    ...
  qa_deterministic.json   # B3 output (written before B4)
  qa_cursor.json          # B4 output (pass/fail, issues, fixes_applied)
```

### `manifest.json` (extends existing stub)

Use `ingestion/manifest.py` types. Each slide entry **must** include:

Worked example: the §12 BambooHR mix is 5 products across 3 funded categories, so `P = 3 + 5 = 8` and `slide_count = 8 + 8 = 16`. The pitch section occupies indices 6–13 in Workflow order (High-Impact → Editorial → Print/Branded Content), so index 9 is the first Editorial product and index 13 is the last pitch slide.

```json
{
  "schema_version": "1",
  "client_name": "BambooHR",
  "template_key": "FortuneAI_DeckTemplate.pptx",
  "slide_count": 16,
  "slides": [
    {
      "slide_index": 0,
      "role": "cover",
      "editable": true
    },
    {
      "slide_index": 8,
      "role": "other",
      "slide_kind": "divider",
      "editable": true
    },
    {
      "slide_index": 9,
      "role": "product",
      "editable": false,
      "product_name": "CEO Daily",
      "source_path": "Fortune_Newsletters_2026.pptx",
      "source_slide_number": 3
    }
  ]
}
```

**Implementation note:** Add `editable: bool` to `SlideManifestEntry` in the B2 PR (default derived from `role`). Do not add role `stylist` (rejected by `tests/test_manifest.py`).

### Slide geometry (derive roles from this, not from guesswork)

The manifest is built from the **post-C2, post-variant-deletion** deck. That is *not* the same index space as the template or as the post-C1 skeleton, and the difference is what decides `editable`.

```text
FortuneAI_DeckTemplate.pptx        19 slides (FORTUNEAI_MIN_SLIDES)
  0–11   narrative spine
  12–16  the five dividers (physical order, NOT pitch order)
  17     investment
  18     thank you

after assemble_skeleton (C1)       14 + P slides
  all five stock dividers deleted, funded sections re-cloned in Workflow
  order and inserted at index 12; P = (funded dividers) + (product clones)

after apply_placeholders (C2)      8 + P slides
  delete_unused_variants drops 4 of the 5 audience pages (template 4–8)
  and 2 of the 3 program pages (template 9–11) → 6 slides removed
```

**Final (manifest) index map — `P = funded_divider_count + len(confirmed_products)`:**

| Final index | Slide | `role` | `editable` |
|-------------|-------|--------|------------|
| `0` | Intro / cover | `cover` | `true` |
| `1` | Why Fortune (stock copy) | `narrative` | `true` |
| `2` | History of Trust | `narrative` | `true` |
| `3` | Opportunity | `narrative` | `true` |
| `4` | Audience (surviving variant) | `narrative` | `true` |
| `5` | Program Overview (surviving variant) | `narrative` | `true` |
| `6 … 5 + P` | Pitch section — divider or A5 clone, in Workflow order | `other` / `product` | `true` / **`false`** |
| `6 + P` (`slides[-2]`) | Investment | `other` | `true` |
| `7 + P` (`slides[-1]`) | Thank you | `other` | `true` |

Cross-checks that must hold, and that the existing suite already asserts: `tests/smoke_build_live.py` expects **10** slides for one newsletter product (P = 2), and `tests/test_generator.py::test_build_*` asserts `slide_count == 10` for the same mix. `test_assemble_single_newsletter_keeps_editorial_divider_only` asserts **16** for the pre-C2 skeleton (14 + 2).

Actual output of that P = 2 case, driven through `apply_placeholders` on the CI fixture (first text run per slide):

```text
post-C1: 16 slides    post-C2: 10 slides  (= 8 + P)
  0: ACME CORP ENTERPRISE PARTNERSHIP | September 2026      cover
  1: FORTUNE POWERS THE LEADING MINDS IN BUSINESS           narrative
  2: TRUST IS THE ULTIMATE COMPETITIVE ADVANTAGE            narrative
  3: Lead With Confidence Today | Enterprise buyers are…    narrative
  4: FORTUNE OVERDELIVERS Acme Corp's TARGET AUDIENCE       narrative
  5: PROGRAM OVERVIEW | Editorial Alignment                 narrative
  6: Editorial Alignment                                    other / divider
  7: CEO DAILY CLONE                                        product  ← editable: false
  8: $50,000 | Editorial Alignment                          other / investment
  9: Thank you! | September 2026                            other / thank_you
```

Under the superseded `1–11 → narrative` rule, indices 6–9 all become `narrative, editable: true` — including index 7, the A5 clone.

> **Do not** write `indices 1–11 → narrative`. That range is the *template* spine; after C2 only indices 1–5 survive, and 6+ are dividers and product clones. Marking them `narrative` sets `editable: true` on A5 clones and hands the QA agent permission to rewrite them — a direct violation of Hard Rule 1.

### Where provenance comes from

PR-A may not touch `ingestion/generator.py` (§13), so it cannot read `AssembledSkeleton` provenance. Resolve it this way, in order of preference:

1. **Preferred.** PR-A adds a pure planner to its own new module: `plan_pitch_sequence(schema, gtm_map) -> list[PitchItem]`, where `PitchItem` is either `("divider", category_index)` or `("product", ProductSlideRef)`. This reproduces `_group_products_by_divider` + the Workflow ordering loop in `assemble_skeleton`, and `ProductSlideRef` already carries `product_name`, `deck_path`, and `slide_number`. PR-F then refactors `assemble_skeleton` to call the same planner, so there is exactly one ordering implementation.
2. **Fallback.** `build_review_package` accepts an optional pre-built plan and uses it when the caller supplies one.

Do **not** duplicate the ordering loop inline inside `build_review_package` and leave `generator.py` with its own copy — that is a silent-desync trap the first time divider order changes.

`source_path` is the raw GTM `Deck Path` value (e.g. `Fortune_Newsletters_2026.pptx`), not the S3 key. `ingestion.gtm_product_map.product_deck_s3_key()` adds the `product-decks/` prefix at load time; the manifest records what GTM said.

### Telling `other` slides apart

All of divider / investment / thank-you collapse to `role: other`, which leaves the B4 agent unable to distinguish them. Roles are frozen at four values by `tests/test_manifest.py`, so add an optional free-text discriminator instead — `slide_kind: str | None` (`"divider"`, `"investment"`, `"thank_you"`) — rather than widening the `role` enum.

---

## 6. Deterministic QA (B3)

Module: `ingestion/deck_qa.py`

Function signature (planned):

```python
def run_deterministic_qa(
    prs,
    schema: DeckSchema,
    manifest: ReviewManifest,
) -> QaReport: ...
```

**Checks (minimum):**

| Check | Fail condition |
|-------|----------------|
| Slide count | `len(prs.slides) != manifest.slide_count` |
| Leftover tokens | Any token in the list below survives in slide text |
| Investment budget | Stated tier amount ≠ sum(`confirmed_products.price`) when single tier used |
| Product presence | Each `confirmed_products[].name` appears in deck text (case-insensitive) |
| Divider order | Funded dividers appear in Workflow order (title substring match on stock dividers) |
| Tail slides | Last two slides are investment + thank-you stock |
| PPTX integrity | Save to bytes and re-open without exception |
| Manifest consistency | Every manifest slide_index in range |

Output: `qa_deterministic.json` with `passed: bool`, `checks: [{name, passed, message}]`.

**Leftover-token list.** `tests/smoke_build_live.py::LEFTOVER_TOKENS` is the starting point but is not complete — it omits the audience-card and investment tokens, which the smoke checks separately or not at all. B3 must cover all of these, and should import the constants from `ingestion.placeholder_fills` / `ingestion.pptx_tools` rather than re-typing the literals:

| Constant | Literal | Defined in |
|----------|---------|------------|
| `TITLE_TOKEN` | `[TITLE]` | `placeholder_fills` |
| `HEADER_TOKEN` | `[HEADER]` | `placeholder_fills` |
| `BODY_TOKEN` | `[BODY]` | `placeholder_fills` |
| `AUDIENCE_TITLE_TOKEN` | `[AUDIENCE TITLE]` | `placeholder_fills` |
| `DATE_TOKEN` | `[DATE]` | `placeholder_fills` |
| `AUDIENCE_SEGMENT_TOKEN` | `[AUDIENCE SEGMENT]` | `placeholder_fills` |
| `REACH_TOKEN` | `[REACH]` | `placeholder_fills` |
| `INDEX_TOKEN` | `[INDEX]` | `placeholder_fills` |
| `BUDGET_TOKEN` | `[BUDGET]` | `placeholder_fills` |
| `PRODUCT_CATEGORY_TOKEN` | `[PRODUCT CATEGORY]` | `placeholder_fills` |
| `PRODUCT_TYPE_LITERAL` | `PRODUCT TYPE` | `placeholder_fills` |
| `PRODUCT_DESCRIPTION_LITERAL` | `Product description.` | `placeholder_fills` |
| `CLIENT_NAME_UPPER_TOKEN` | `[CLIENT NAME]` | `placeholder_fills` |
| `CLIENT_NAME_TOKEN` | `[client name]` | `pptx_tools` |
| `CLIENT_NAME_POSSESSIVE_TOKEN` | `[CLIENT NAME'S]` (curly apostrophe — use the constant) | `pptx_tools` |
| `LOGO_TOKEN` | `[LOGO]` | `pptx_tools` |
| — | `[PRODUCT + PRODUCT PRICE]` | investment slide; no constant yet, add one in PR-B |

Scan **all** slides for these, including `role: product` clones — B3 reports on product slides even though B4 may not fix them.

**Note on the budget check.** `placeholder_fills._assert_budget_matches_mix` already raises during C2, so in the `build_deck` path this check can never fail — C2 would have thrown first. Keep it anyway: B3 also runs standalone against a review package directory (`scripts/run_deck_qa.py --package-dir`), where nothing upstream has validated the schema. Reuse `mix_total()` / `stated_total_budget()` from `placeholder_fills`; do not reimplement the tier-selection rule (a tier whose label matches `total` wins, else `max(budgets[].amount)`).

**B4 must not run if B3 `passed: false`** (except optional `force` flag for dev).

---

## 7. Headless Cursor QA (B4)

Runner: `scripts/run_deck_qa.py` (invokes Cursor SDK — **not** an MCP tool exposed to Prodie).

### Inputs

- Path or S3 URI to review package directory
- `CURSOR_API_KEY` (or SDK auth per Cursor docs)
- Repo checkout with `.cursor/skills/deck-qa/SKILL.md`

### Agent responsibilities

1. Read `manifest.json` + `qa_deterministic.json`
2. Vision-review PNGs for **editable** slides only
3. Evaluate: leftover visual placeholders, broken alignment, unreadable text, off-brand obvious issues, weak copy on allowed slots
4. Either **pass** or apply **one** fix pass:
   - Re-invoke `PlaceholderAI` for failed narrative slots (same bounds as C2)
   - `pptx_tools.set_ph_text` / `replace_token` / `replace_first_token` on allowed slides (see §4 — not `apply_replacements`)
   - Re-render affected PNGs
5. Write `qa_cursor.json`:

```json
{
  "passed": true,
  "loop_count": 1,
  "issues": [{"slide_index": 4, "severity": "warning", "message": "..."}],
  "fixes_applied": ["opportunity_body", "program_blurb_Editorial Alignment"]
}
```

**Field semantics** (added while authoring the skill; the shape above did not pin these, and each one changes whether a deck ships):

- `severity` — `info` | `warning` | `error`.
- `passed` — `false` only when an `error` lands on an `editable: true` slide. Issues on `editable: false` product clones are **flag-only** per the §4 role table: report them at `warning` at most. B4 is forbidden from fixing them, so failing the deck on one blocks delivery on something no one in the pass can repair — that is a GTM data escalation, not a QA gate failure.
- `loop_count` — `0` when B4 reviewed and changed nothing, `1` when it ran the fix pass. Never higher (Hard Rule 7). Re-rendering and re-viewing PNGs to verify a fix is part of the same loop.
- `fixes_applied` — slot ids from §7 step 4, or a short stable label for a deterministic token fix. Must be non-empty **iff** `draft.pptx` changed on disk; §8's re-load keys off it.

### Outputs

- Updated `draft.pptx` (or `final.pptx`)
- `qa_cursor.json`
- On pass: upload final to S3 `generated/{uuid}.pptx` (same as today)

**Do not** use browser automation for QA. **Do not** expose `run_deck_qa` as associate-facing MCP tool in MVP.

---

## 8. Integration with `build_deck` (Wave 3)

`DeckGenerator.build()` flow after C2:

```python
# Pseudocode — exact shape decided in integration PR
warnings = apply_placeholders(prs, schema, audience=..., logo_bytes=..., ai=ai)

qa_block = None
if _deck_qa_enabled():
    # build_review_package renumbers slide parts before saving draft.pptx and
    # mutates prs in the process — see "three things" note 2 below.
    package = build_review_package(prs, schema, plan=plan)   # kwarg name: `plan`, see §5
    det_report = run_deterministic_qa(prs, schema, package.manifest)
    package.write_deterministic_report(det_report)
    if not det_report.passed:
        raise DeckQaError(det_report.summary(), report=det_report)

    agent_report = run_headless_cursor_qa(package, timeout_s=DECK_QA_TIMEOUT_S)
    if not agent_report.passed:
        raise DeckQaError(agent_report.summary(), report=agent_report)

    # B4 edits draft.pptx ON DISK. Re-load it or every fix is thrown away.
    if agent_report.fixes_applied:
        prs = Presentation(package.draft_path)

    qa_block = {...}

self._renumber_slide_parts(prs)
# save + upload
```

### Three things the integration PR must get right

**1. Reload the deck after B4.** B4 writes fixes into `draft.pptx` inside the package directory. `build()` holds a separate in-memory `Presentation`, saves *that*, and uploads it. Without the re-load, QA runs, reports `passed: true`, applies fixes — and ships the unfixed deck. Either re-open `package.draft_path` as shown, or make the runner hand back PPTX bytes; pick one and state it in the PR description.

**2. `draft.pptx` must be renumbered before it is saved.** A5 clones arrive carrying their source partnames, so saving an assembled deck without `rename_slide_parts` emits two zip entries under one `ppt/slides/slideNN.xml` (plus its `.rels`). Verified on the five-product mix by disabling the renumber: `UserWarning: Duplicate name: 'ppt/slides/slide19.xml'`. A three-clone toy case does **not** reproduce it — you need multiple dividers cloned from a second template copy alongside products from distinct source decks.

The pseudocode above renumbers only at the end, which would hand B3 and B4 a malformed draft. `build_review_package` therefore renumbers before serializing, which also mutates the caller's `Presentation` and makes `build()`'s later `_renumber_slide_parts` a harmless no-op. Keep both calls: packaging must not assume it is running inside `build()`.

Because packaging renumbers, `draft.pptx` and `final.pptx` *are* byte-identical when QA changes nothing. The "product slide unchanged after QA" item in the prompts-doc verification checklist should still compare *rendered PNGs* rather than hashes, but for the real reason: a B4 fix on any editable slide rewrites shared package parts, so a whole-file hash tells you nothing about whether a specific product slide moved.

**3. Failure semantics.** `server.py::build_deck` catches `ValueError` and returns `{"status": "error", "message": str(exc)}`, so a bare `raise ValueError` loses the report and contradicts Hard Rule 7 ("second failure → `status: error` with `qa_report`"). Define `DeckQaError(ValueError)` carrying `.report`, and widen the `server.py` handler to attach `qa_report` when present:

```python
except DeckQaError as exc:
    return {"status": "error", "message": str(exc), "qa_report": exc.report.to_json()}
except ValueError as exc:
    return {"status": "error", "message": str(exc)}
```

Subclassing `ValueError` keeps the existing handler working if PR-F ships before the `server.py` edit.

### Runner interface (pin this, don't leave it to the implementer)

`run_headless_cursor_qa(package, *, timeout_s)` is an **in-process helper**, imported by `generator.py`; the CLI is a thin `argparse` wrapper over the same helper. Do not shell out to `python scripts/run_deck_qa.py` from `build()` — `scripts/` is not in `packages` in `pyproject.toml`, so it is not importable or resolvable from an installed wheel or the Docker image without extra path work.

**Decided in PR-E: the helper lives in `ingestion/deck_qa_agent.py`** and `scripts/run_deck_qa.py` is a CLI shim that puts the repo root on `sys.path` and calls `ingestion.deck_qa_agent.main()`. PR-F imports `from ingestion.deck_qa_agent import run_headless_cursor_qa`. `scripts` was **not** added to `[tool.setuptools] packages`.

It returns a `CursorQaReport` (§7 shape: `passed`, `loop_count`, `issues`, `fixes_applied`), not the B3 `QaReport` — both expose `passed` / `summary()` / `to_json()`, so `DeckQaError(report=...)` works with either. The helper resolves the skill file from the repo checkout at `.cursor/skills/deck-qa/SKILL.md`, so that path must exist wherever the rail runs, including the Docker image.

Return shape addition (backward compatible):

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

On failure the `qa` block is not returned (the call raises); the report travels in `qa_report` per the failure semantics above.

### Environment

| Var | Default | Meaning |
|-----|---------|---------|
| `DECK_QA_ENABLED` | `false` | Unset or anything outside `{"1","true","yes"}` (case-insensitive) → skip the whole rail; `build()` behaves exactly as today |
| `DECK_QA_TIMEOUT_S` | `600` | Wall-clock budget for B2+B3+B4 combined |
| `CURSOR_API_KEY` | — | Required when `DECK_QA_ENABLED` is truthy |

### Resolved: `build_deck` stays synchronous and fails soft on timeout

`build_deck` is an MCP tool that Prodie calls and blocks on. Today it is one LibreOffice-free assembly plus a handful of Claude calls. With the rail on, a single call additionally does: a LibreOffice PPTX→PDF conversion, ~16 `pdftoppm` rasterizations, and a headless Cursor agent session with up to 16 attached images and a fix loop. That is plausibly **several minutes**, against MCP clients that commonly time out well before that.

**Decision (2026-09-08): option (a) — ship synchronous, fail soft on timeout.** Considered and rejected: **(b)** fail closed on timeout, which makes a slow runner indistinguishable from a bad deck and blocks the associate; **(c)** split into `build_deck` → `{"status": "qa_pending", "review_id": ...}` plus a second poll tool, which is the architecturally correct answer but changes the Prodie contract and therefore pulls PI-2350 into this track.

Option (a) is the only one that keeps `DECK_QA_ENABLED=false` and `=true` on the same call signature, which is what lets the flag be a true no-op rollback.

**What PR-F must implement:**

1. B2 + B3 + B4 run inside `build()` under a single `DECK_QA_TIMEOUT_S` budget (default `600`).
2. **On timeout, do not raise.** Upload the deck as it stands and return normally with `qa.cursor_passed: false`, `qa.timed_out: true`, and a human-readable entry appended to the existing `warnings[]` list. Associates already read `warnings[]`, so the degradation is visible without a new field they'd have to learn.
3. **A timeout is not a QA failure.** `DeckQaError` is for B3 failing or B4 returning `passed: false` — genuine quality verdicts that must fail loud per Hard Rule 5. A timeout is an infrastructure symptom and must not be laundered into either a pass or a quality failure.
4. **B3 failures still raise**, regardless of the timeout budget. B3 is fast, deterministic, and has no legitimate reason to time out; a deterministic failure means the deck is actually wrong.
5. Log timeouts at `WARNING` with the `review_id`, so the rate of soft-failures is measurable. If it is not near zero in practice, revisit **(c)** — a gate that times out often is not a gate.

This weakens the guarantee: under sustained timeouts a deck can ship un-QA'd. That is accepted for MVP because the alternative blocks delivery on runner latency, and because the review package is still written to S3 for after-the-fact inspection.

---

## 9. Parallel PR map (dependency order)

```text
Wave 1 (parallel — no shared files):
  PR-A  B2 review package + manifest `editable` field
  PR-B  B3 ingestion/deck_qa.py + tests
  PR-C  deck-qa skill + .cursor/rules + docs pointer to this file
  PR-D  Golden BambooHR test (tests/test_bamboohr_golden.py or extend smoke)

Wave 2 (after Wave 1 merges):
  PR-E  scripts/run_deck_qa.py (Cursor SDK headless runner)

Wave 3 (after PR-E):
  PR-F  Wire build_deck + server docstring + DECK_QA_ENABLED
```

**Serialization rule:** Only **one** open PR may touch `ingestion/generator.py` at a time (Wave 3).

---

## 10. Out of scope (do not build in these PRs)

- Prodie UI / checkboxes (PI-2350)
- Rewriting A5 product slides
- Three-tier investment column layout (human deck parity)
- Branded connector slides between products
- `search_decks` / Titan in Creation path
- SharePoint upload of review packages
- More than 1 Cursor fix loop

---

## 11. Test commands

```bash
# Unit tests (no S3)
PYTHONPATH=. uv run pytest tests/test_manifest.py tests/test_render_slides.py -q

# Assembly
PYTHONPATH=. uv run pytest tests/test_generator.py -k assemble -q

# Live-ish smoke (S3 + mock AI)
PYTHONPATH=. uv run python tests/smoke_build_live.py --mock-ai

# BambooHR golden (after PR-D)
PYTHONPATH=. uv run pytest tests/test_bamboohr_golden.py -q
```

---

## 12. BambooHR reference schema (golden test)

Use for PR-D and manual QA:

- `company_name`: BambooHR
- `budgets`: Tier 1 `$400,000` (mix must sum exactly)
- `flight_dates`: 2026-09-28 → 2026-12-31
- `targeting_details`: must match **≥2** Audience Data segments (e.g. `Chief Executive Officer, Chief Financial Officer, C-suite leadership`)
- `confirmed_products` (example Tier 1 mix = $400k):
  - CEO Daily, Newsletter, $75,000
  - CFO Daily, Newsletter, $60,000
  - Fortune Workplace Innovation, Newsletter, $85,000
  - Dynamic Content Hub, Branded Content, $95,000
  - Crown Unit, Digital Media, $85,000

> ### ⚠️ These product names and prices are unverified
>
> Only **CEO Daily** and **Crown Unit** appear anywhere in the checked-out repo (`tests/logic_guide_fixtures.py`). **CFO Daily**, **Fortune Workplace Innovation**, and **Dynamic Content Hub** appear in no fixture, so nothing confirms they exist as GTM `Product Tags` rows. `assemble_skeleton` calls `gtm_map.lookup(name, category)`, which **fails loud** on an unknown name — so a golden test built on these names dies at assembly against the real GTM workbook.
>
> The prices also contradict the rate card the repo does have: `REPRESENTATIVE_PRICING_ROWS` lists CEO Daily at `$5,000/day` and Crown Unit at `$25,000`, not `$75,000` / `$85,000`. That is fine for a hand-built `DeckSchema` (C2 only checks that the mix sums to the stated budget) but it is **not** a mix `confirm_mix` could ever produce, since `confirm_mix` derives price from the inventory workbook. Do not present this as a realistic end-to-end fixture.
>
> **PR-D must do one of these, and say which in the PR description:**
>
> 1. **Preferred — unit-level, offline.** Build the golden `DeckSchema` from names confirmed present in `tests/logic_guide_fixtures.py::REPRESENTATIVE_GTM_ROWS`, and drive assembly with a fixture `GtmProductMap` rather than live S3. Safe candidates spanning ≥3 dividers: `Crown Unit` (Digital Ads/Programmatic), `CEO Daily` + `Term Sheet` (Newsletters), `Long-Form Article` (Branded Content), `Full Page` (Print). Keep the $400,000 total; the point of the golden is divider order and slide geometry, not price realism.
> 2. **Verify first.** Confirm all five names against the live GTM workbook (`GTM_DATABASE_KEY`), correct the prices from the inventory Pricing sheet, and gate the test on `S3_SNAPSHOT_BUCKET`.
>
> Under either option the assertion that matters is `slide_count == 8 + P` with the §5 index map, not the specific dollar figures.

`targeting_details` is a free-text **string**, not a list — it is fuzzy-matched against Audience Data rows by `AudienceData.match_targeting`. It must resolve to ≥2 rows or `select_audience_variant` raises. `"Chief Executive Officer, Chief Financial Officer, C-suite leadership"` is only safe if all three segments exist in the workbook; the CI fixture has `C-suite`, not `C-suite leadership`.

---

## 13. File touch matrix (prevent conflicts)

| File / area | PR-A | PR-B | PR-C | PR-D | PR-E | PR-F |
|-------------|------|------|------|------|------|------|
| `ingestion/review_package.py` | create | — | — | — | — | import |
| `ingestion/manifest.py` | extend | — | — | — | — | — |
| `ingestion/deck_qa.py` | — | create | — | — | import | import |
| `.cursor/skills/deck-qa/` | — | — | create | — | read | — |
| `scripts/run_deck_qa.py` | — | — | — | — | create | call |
| `ingestion/generator.py` | — | — | — | — | — | **only** |
| `server.py` | — | — | — | — | — | **only** |
| `tests/test_bamboohr_golden.py` | — | — | — | create | — | — |
| `README.md` | — | — | — | — | — | **only** |
| `.env.example` | — | — | — | — | — | **only** |
| `pyproject.toml` | — | — | — | — | deps | — |
| `local/.../PROGRESS.md` | — | — | — | — | — | **only** |

**Corrections to the original matrix.** As first drafted it omitted `README.md` and `.env.example` while PR-E, PR-F *and* PR-G were each told to edit them, and omitted `pyproject.toml` which PR-E must touch for the SDK dependency. Resolution above: README and `.env.example` are PR-F-only. PR-E documents `CURSOR_API_KEY` in its PR description rather than editing `.env.example`; PR-G drops its README paragraph and edits only `.cursor/rules/pitch-deck-end-scope.mdc`.

**`local/` is tracked, not gitignored.** `git ls-files local` returns four files, including `END-SCOPE-SOT.md` and `PROGRESS.md`. Both PR-F and PR-G contain a conditional "if `local/` is gitignored…" — that branch is dead. `PROGRESS.md` is tracked and PR-F must update its status table.

Two other stale statements to clean up while nearby (either PR is fine, just not both): `ingestion/manifest.py`'s module docstring asserts `assemble_skeleton` returns `AssembledSkeleton`, which is false on the base branch, and `ingestion/render_slides.py`'s docstring still says "Cursor stylist".

---

## 14. Changelog

| Date | Change |
|------|--------|
| 2026-09-08 | Initial architecture — headless Cursor QA required for MVP |
| 2026-09-08 | §8 open decision resolved: `build_deck` stays synchronous and **fails soft** on timeout (`qa.timed_out`, warning appended, deck still delivered). A timeout is infrastructure, not a quality verdict — B3 failures and B4 `passed: false` still raise `DeckQaError` |
| 2026-09-08 | Wave 1 landed (PI-2522 B2, B3, deck-qa skill, BambooHR golden). §8 corrected: `draft.pptx` must be renumbered before saving or A5 clones emit duplicate `ppt/slides/slideNN.xml` zip entries; the prior "never byte-identical" rationale was wrong |
| 2026-09-08 | PR-C: §7 `qa_cursor.json` field semantics pinned (`severity` enum, `passed` rule for flag-only product-clone issues, `loop_count`, `fixes_applied` iff-changed) — the shape alone left the exit contract ambiguous for the skill |
| 2026-09-08 | PR-E: B4 runner landed as `ingestion/deck_qa_agent.py` + `scripts/run_deck_qa.py` shim (§8 import-path decision resolved, option a). `fixes_applied` is reconciled against the draft's on-disk digest, so an unreported agent edit still triggers §8's re-load |
| 2026-09-08 | Verified against the checkout and corrected before subagent deploy: base branch is `fix/fortuneai-deck-assembly` (PR #31), not `main`; §5 post-C2 index map rewritten (was `1–11 narrative`, actually `1–5`, which would have marked A5 clones editable); §5 provenance sourcing and `PI-2522` prerequisite documented; §6 leftover-token list completed from source constants; §4/§7 fix method changed from `apply_replacements` to `replace_token`; §8 in-memory-vs-on-disk reload bug, `DeckQaError` failure semantics, runner import path, and sync-call timeout risk called out; §12 golden products flagged as unverified; §13 matrix gaps closed |
