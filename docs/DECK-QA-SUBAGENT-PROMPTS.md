# Deck QA subagent deployment prompts

> Copy one prompt per subagent. **Every agent must read `docs/DECK-QA-ARCHITECTURE.md` first** and cite it in the PR.  
> **Repo:** `/Users/tcheung/Library/CloudStorage/OneDrive-FortuneMedia(USA)Corporation/Documents/GitHub/sales-mcp`  
> **Base branch:** `fix/fortuneai-deck-assembly` — **not `main`**. See below.  
> **Do not merge PRs** unless the user explicitly asks.  
> **PR title:** must match a ticket id when a Jira ticket exists (e.g. `PI-2522: B2 review package builder`).

## Base branch — read before launching anything

`origin/main` does **not** contain `FORTUNEAI_DIVIDER_SLIDE_INDEX`; it still has the older `FORTUNEAI_FIRST_DIVIDER_INDEX = 12` and a different divider-insertion path. The slide geometry the architecture doc describes exists only on **`fix/fortuneai-deck-assembly`** (open PR #31, "fix: FortuneAI deck assembly and product-pitch order").

Cut every Wave 1–3 branch from `fix/fortuneai-deck-assembly`. Do not tell agents to rebase onto `main` until #31 merges. If #31 merges mid-flight, rebase the open QA branches then.

## Ticket IDs

Two of these already have tickets; do not open new ones for them.

| PR | Ticket | Notes |
|----|--------|-------|
| PR-A | **PI-2522** | Existing ticket + branch `origin/PI-2522-review-package-builder` (review package builder) |
| PR-B–PR-G | need new tickets | Create before opening the PR; the title must match the ticket name |

PI-2521 (LibreOffice `render_slides`) and PI-2754 (Cursor stylist, being rescoped by this work) are already closed/shelved — reference them, don't reuse them.

---

## Shared preamble (prepend to every subagent)

```text
You are implementing one slice of the Deck QA MVP rail for sales-mcp.

MANDATORY READ FIRST: docs/DECK-QA-ARCHITECTURE.md (entire file). Treat it as SoT.
Do not invent file paths, APIs, MCP tools, or Workflow rules not in that doc or the repo.

BASE BRANCH: fix/fortuneai-deck-assembly. Branch from it, not from main.
main lacks FORTUNEAI_DIVIDER_SLIDE_INDEX and has different slide geometry.

SLIDE GEOMETRY (architecture doc §5) — the single easiest thing to get wrong.
The final deck is 8 + P slides, where P = funded_dividers + len(confirmed_products):
  0        cover
  1-5      narrative (Why Fortune, History, Opportunity, Audience, Program Overview)
  6..5+P   pitch section: dividers (role other) and A5 clones (role product, editable false)
  6+P      investment (other)
  7+P      thank you (other)
The template's 0-11 narrative range does NOT survive C2 — delete_unused_variants
drops 6 variant pages inside apply_placeholders. If you index as if 1-11 were
narrative you will mark product clones editable and break Hard Rule 1.
Sanity check against the suite: one newsletter product (P=2) => 10 slides.

Hard rules:
- Do NOT edit A5 product clone slides in QA code (role: product, editable: false).
- Do NOT add search_decks / Titan / propose_mix to the Creation path.
- Do NOT touch ingestion/generator.py or server.py unless your prompt explicitly says so.
- Do NOT merge PRs. Open PR and return URL.
- Run relevant tests before finishing; report pass/fail counts.
- Follow existing code style; minimal diff; no unrelated refactors.
- No Co-authored-by trailers for AI.

When done, return:
1. PR URL
2. Files changed (bullet list)
3. How to test (exact commands)
4. Anything blocked on other PRs
```

---

## Wave 1 — run these in parallel (4 subagents)

### Subagent PR-A — B2 review package builder

```text
[SHARED PREAMBLE]

Ticket: PI-2522 (existing — "review package builder")

Branch: feat/deck-qa-review-package, cut from fix/fortuneai-deck-assembly

READ FIRST, IN ADDITION TO THE ARCHITECTURE DOC:
- git show origin/PI-2522-review-package-builder:ingestion/generator.py
  That branch already makes assemble_skeleton return AssembledSkeleton with
  product-clone provenance. It is NOT on your base branch, and it touches
  generator.py, which is reserved for PR-F. Do not merge or cherry-pick it.
  Read it so your provenance model matches, then follow §5 "Where provenance
  comes from" and build a standalone planner instead.
- ingestion/manifest.py's module docstring claims assemble_skeleton already
  returns AssembledSkeleton. That is false on your base branch. Fix the
  docstring as part of this PR.

SCOPE:
Implement B2 review package builder per docs/DECK-QA-ARCHITECTURE.md §5.

DELIVERABLES:
1. New module `ingestion/review_package.py` with:
   - `plan_pitch_sequence(schema: DeckSchema, gtm_map: GtmProductMap) -> list[PitchItem]`
     Pure function. Mirrors _group_products_by_divider + the Workflow ordering
     loop in assemble_skeleton. PitchItem is ("divider", category_index) or
     ("product", ProductSlideRef). PR-F will refactor generator.py to call this,
     so keep it dependency-free and importable without S3.
   - `build_review_package(prs, schema: DeckSchema, *, plan: list[PitchItem] | None = None, output_dir: Path | None = None) -> ReviewPackage`
     When plan is None, compute it via plan_pitch_sequence.
     NOTE the kwarg is `plan`. Earlier drafts said `product_clones` and
     `product_provenance`; those names are dead.
   - `ReviewPackage` dataclass with explicit paths, not just bytes:
     `root: Path`, `draft_path: Path`, `manifest_path: Path`, `slides_dir: Path`,
     `manifest: ReviewManifest`, `pptx_bytes: bytes`
     plus `write_deterministic_report(report)` / `write_cursor_report(report)`
     helpers that PR-E and PR-F will call.
   - Writes manifest.json + deck_schema.json per §5 (include `editable` per slide)
   - Calls `ingestion.render_slides.render_slides` for all slide indices (150 DPI);
     it already names files slide-NNN.png, so pass output_dir and do not rename.
   - Assigns roles from the §5 index map. Derive product indices from the plan,
     not by scanning slide content.
   - source_path is the raw GTM `Deck Path` (e.g. "Fortune_Newsletters_2026.pptx"),
     NOT the product-decks/ S3 key.

2. Extend `ingestion/manifest.py`:
   - Add `editable: bool` to `SlideManifestEntry` (default from role: product → false, else true)
   - Add `slide_kind: str | None` to distinguish divider / investment / thank_you
     within role "other" (see §5). Do NOT widen the role enum —
     tests/test_manifest.py::test_manifest_rejects_bad_role locks it to four values.
   - Keep roles limited to: cover, narrative, product, other

3. Tests `tests/test_review_package.py`:
   - Use tests/fortuneai_placeholder_fixture (build_fortuneai_fixture_prs,
     mock_placeholder_ai, sample_audience_data, MINIMAL_PNG) — no live S3.
   - Assert manifest validates; product slides have editable=false; role map
     matches §5 for a known mix; slide_count == 8 + P.
   - render_slides needs LibreOffice + pdftoppm. Gate any PNG-count test on
     shutil.which("soffice") / SOFFICE_BIN and skip cleanly in CI, or mock
     render_slides. Do NOT make the default unit run depend on LibreOffice.

4. Optional CLI: `python -m ingestion.review_package --help` writing to a temp dir

OUT OF SCOPE:
- generator.py, server.py, Cursor SDK, deck_qa.py

ACCEPTANCE:
- pytest tests/test_review_package.py tests/test_manifest.py pass
- No new MCP tools

Open PR against fix/fortuneai-deck-assembly. Title: "PI-2522: B2 review package builder"
```

---

### Subagent PR-B — B3 deterministic QA

```text
[SHARED PREAMBLE]

Ticket label: PI-XXXX-B3-deterministic-qa

Branch: feat/deck-qa-deterministic

SCOPE:
Implement B3 deterministic QA per docs/DECK-QA-ARCHITECTURE.md §6.

DELIVERABLES:
1. New module `ingestion/deck_qa.py`:
   - `QaCheckResult`, `QaReport` dataclasses
   - `DeckQaError(ValueError)` carrying `.report` — PR-F needs it (see §8).
     Subclass ValueError so server.py's existing handler keeps working.
   - `run_deterministic_qa(prs, schema: DeckSchema, manifest: ReviewManifest) -> QaReport`
   - All checks listed in architecture §6 (leftover tokens, budget, product names, tail slides, integrity, manifest consistency)
   - `QaReport.to_json()` and `.summary()` for qa_deterministic.json shape
   - Leftover tokens: IMPORT the constants from ingestion.placeholder_fills and
     ingestion.pptx_tools per the §6 table. Do not retype the literals —
     CLIENT_NAME_POSSESSIVE_TOKEN uses a curly apostrophe (pptx_tools.APOS) and
     will not match if you type a straight quote. tests/smoke_build_live.py's
     LEFTOVER_TOKENS is an incomplete subset; §6 is the full list.
     Add a constant for "[PRODUCT + PRODUCT PRICE]" (investment slide) — it has
     none today. Put it in placeholder_fills next to BUDGET_TOKEN.
   - Budget check: reuse mix_total() / stated_total_budget() from
     placeholder_fills. Do not reimplement the tier rule. Note in a comment that
     C2's _assert_budget_matches_mix already raises on this in the build_deck
     path, so this check only bites for standalone package runs.

2. Tests `tests/test_deck_qa.py`:
   - Pass case using fortuneai fixture + filled placeholders
   - Fail cases: leftover [TITLE], budget mismatch (one test each)
   - No S3, no Cursor, no LibreOffice

OUT OF SCOPE:
- generator.py, server.py, review_package.py (import types from manifest.py only)
- Vision / PNG analysis

ACCEPTANCE:
- pytest tests/test_deck_qa.py pass
- ruff clean on new files

Open PR. Title: "PI-XXXX: B3 deterministic deck QA"
```

---

### Subagent PR-C — Deck QA Cursor skill + rule

```text
[SHARED PREAMBLE]

Ticket label: PI-XXXX-B4-deck-qa-skill

Branch: feat/deck-qa-cursor-skill

SCOPE:
Authoritative instructions for headless Cursor vision QA per docs/DECK-QA-ARCHITECTURE.md §7.

DELIVERABLES:
Note: `.cursor/skills/` does not exist yet — you are creating the directory.

1. `.cursor/skills/deck-qa/SKILL.md`:
   - When to use (after B3 passes)
   - Inputs: review package layout, manifest editable flags. State that the agent
     keys off `editable`, never off slide position, and that role "other" is
     disambiguated by `slide_kind` (divider / investment / thank_you).
   - Vision checklist (stock slides): placeholders visible, text overflow, alignment, brand tone
   - Fix allowed actions: PlaceholderAI re-call, then pptx_tools.set_ph_text /
     replace_token / replace_first_token.
     Do NOT list apply_replacements — it keys off placeholder idx 0/19 and
     rewrites body paragraphs positionally, which is the Titan corpus-slide
     path. On FortuneAI stock slides it no-ops or clobbers the wrong runs.
     See architecture §4.
   - Forbidden: product slides, reordering, new slides, corpus search
   - Exit: qa_cursor.json schema, max 1 loop
   - Link to DECK-QA-ARCHITECTURE.md

2. `.cursor/rules/deck-qa-gate.mdc` (alwaysApply: false, globs: ingestion/**, scripts/run_deck_qa.py):
   - Short pointer to skill + architecture doc
   - Hard rules one-liner

3. Update `docs/DECK-QA-ARCHITECTURE.md` §7 only if skill reveals a gap (minimal edit; note in PR)

OUT OF SCOPE:
- Python runner (PR-E)
- generator.py / server.py

ACCEPTANCE:
- Skill is self-contained; another agent can run headless QA without guessing rules
- No code changes outside docs + .cursor/

Open PR. Title: "PI-XXXX: Deck QA Cursor skill and rules"
```

---

### Subagent PR-D — BambooHR golden test

```text
[SHARED PREAMBLE]

Ticket label: PI-XXXX-bamboohr-golden

Branch: feat/deck-qa-bamboohr-golden

SCOPE:
Golden path test fixture per docs/DECK-QA-ARCHITECTURE.md §12.

DELIVERABLES:
WARNING — READ §12's callout BEFORE WRITING THE FIXTURE:
Three of the five listed products (CFO Daily, Fortune Workplace Innovation,
Dynamic Content Hub) appear in NO repo fixture and may not be real GTM
Product Tags rows. gtm_map.lookup fails loud on unknown names, so an assembly
test built on them dies against the real workbook. The listed prices also
contradict tests/logic_guide_fixtures.py::REPRESENTATIVE_PRICING_ROWS
(CEO Daily is $5,000/day, Crown Unit is $25,000).

Default to §12 option 1: build the golden from names confirmed present in
REPRESENTATIVE_GTM_ROWS — Crown Unit (Digital Media), CEO Daily + Term Sheet
(Newsletter), Long-Form Article (Branded Content), Full Page (Print) — keep the
$400,000 total, and drive assembly with a fixture GtmProductMap, not live S3.
Say explicitly in the PR description which option you took and why.

1. `tests/bamboohr_golden.py` — shared schema factory `bamboohr_tier1_schema() -> DeckSchema`:
   - Products and prices summing to exactly $400,000 (C2 raises otherwise)
   - Every Product needs name, category, price AND cadence — cadence must be one
     of annual/quarterly/monthly/weekly (schema.py _VALID_CADENCES). "daily" is
     not valid; CEO Daily resolves to "weekly" via confirm_mix.derive_cadence.
   - targeting_details is a STRING, fuzzy-matched by AudienceData.match_targeting.
     It must resolve to >=2 rows or select_audience_variant raises. Verify your
     string against the audience source you test with — the CI fixture has
     "C-suite", not "C-suite leadership".
   - budgets must stay under the $750,000 escalation threshold ($400k is fine).

2. `tests/test_bamboohr_golden.py`:
   - `@pytest.mark.integration` or env-gated: skip without S3_SNAPSHOT_BUCKET
     (that var is already in .env.example)
   - Test assembly order: divider Workflow order, products present, investment last-2
   - Assert final slide_count == 8 + P per architecture §5 (P = funded dividers +
     products). This is the assertion the golden exists for; the dollar figures
     are not the point.
   - Test deterministic QA passes on mock-filled deck (use mock_placeholder_ai from tests.fortuneai_placeholder_fixture)
   - Do NOT require Cursor SDK in CI

3. Document in test docstring: how to run locally

OUT OF SCOPE:
- generator.py changes unless required for test hooks (avoid)
- review_package.py, deck_qa.py (use imports if merged; if not merged, test assembly only and note dependency)

ACCEPTANCE:
- Unit parts pass without S3
- Integration test skips gracefully without AWS

Open PR. Title: "PI-XXXX: BambooHR golden deck test"
```

---

## Wave 2 — after Wave 1 merges (1 subagent)

### Subagent PR-E — Headless Cursor QA runner

```text
[SHARED PREAMBLE]

Ticket label: PI-XXXX-B4-headless-runner

Branch: feat/deck-qa-headless-runner

PREREQUISITE: PR-A, PR-B, PR-C merged to main (rebase this branch onto main).

SCOPE:
Headless Cursor QA runner per docs/DECK-QA-ARCHITECTURE.md §7.

DELIVERABLES:
FIRST DECISION — MAKE IT EXPLICITLY, PR-F DEPENDS ON IT (architecture §8):
`scripts/` does not exist and is NOT in pyproject.toml's
[tool.setuptools] packages = ["ingestion"], so scripts/run_deck_qa.py is not
importable from an installed wheel or the Docker image. PR-F needs to call
run_headless_cursor_qa IN-PROCESS, not by shelling out. Choose one and state it
at the top of the PR description:
  (a) put the helper in ingestion/deck_qa_agent.py and make
      scripts/run_deck_qa.py a thin argparse shim over it  <- recommended
  (b) add "scripts" to [tool.setuptools] packages with an __init__.py

1. `scripts/run_deck_qa.py` (+ ingestion/deck_qa_agent.py if option a):
   - Importable helper: `run_headless_cursor_qa(package, *, timeout_s) -> QaReport`
   - CLI: `--package-dir PATH` (local review package with draft.pptx, manifest.json, slides/)
   - Run B3 first; exit 1 if deterministic fails (unless `--force`)
   - Invoke Cursor SDK Agent (read sdk skill at ~/.cursor/skills-cursor/sdk/SKILL.md):
     - Load `.cursor/skills/deck-qa/SKILL.md` as agent instructions
     - Attach slide PNGs for editable indices only (token budget: cap at 12 images or architecture default)
     - Max 1 fix loop
   - Write qa_cursor.json via package.write_cursor_report()
   - Update draft.pptx in package dir on fixes, and set fixes_applied truthfully —
     PR-F keys its re-load of the deck off that field (§8). If you write fixes to
     disk without reporting them, build_deck ships the unfixed deck.
   - Honour DECK_QA_TIMEOUT_S (default 600).

2. `tests/test_run_deck_qa.py`:
   - Mock Cursor SDK / subprocess — do not call live API in CI
   - Test: B3 failure skips agent; success path writes qa_cursor.json
   - Test: fixes_applied non-empty implies draft.pptx mtime/bytes changed

3. `pyproject.toml` deps: add the Cursor SDK only if not present. Note that
   requires-python is >=3.10; if the SDK needs newer, say so in the PR and do
   NOT silently bump requires-python.

4. Do NOT edit README.md or .env.example — both are PR-F-only (§13). Document
   CURSOR_API_KEY and the manual dry-run in your PR description instead.

OUT OF SCOPE:
- build_deck wiring (PR-F)
- Changing product clone behavior

ENV VARS (document only):
- CURSOR_API_KEY
- DECK_QA_ENABLED (used in PR-F)

ACCEPTANCE:
- pytest tests/test_run_deck_qa.py pass
- Manual dry-run instructions in PR description

Open PR. Title: "PI-XXXX: Headless Cursor deck QA runner"
```

---

## Wave 3 — after Wave 2 merges (1 subagent)

### Subagent PR-F — build_deck integration

```text
[SHARED PREAMBLE]

Ticket label: PI-XXXX-build-deck-qa-gate

Branch: feat/deck-qa-build-integration

PREREQUISITE: PR-A, PR-B, PR-E merged to main.

SCOPE:
Wire QA rail into build path per docs/DECK-QA-ARCHITECTURE.md §8.

DELIVERABLES:
1. `ingestion/generator.py` — `DeckGenerator.build()`:
   - After apply_placeholders + before _renumber_slide_parts, gated on
     DECK_QA_ENABLED (unset/false => byte-identical behaviour to today):
       - build_review_package(prs, schema, plan=plan)
       - run_deterministic_qa(...)  -> raise DeckQaError on failure
       - run_headless_cursor_qa(package, timeout_s=DECK_QA_TIMEOUT_S) in-process
   - *** MUST-FIX FROM THE ARCHITECTURE REVIEW (§8) ***
     B4 edits draft.pptx ON DISK; build() holds a separate in-memory
     Presentation and saves THAT. Without re-loading, QA reports success,
     applies fixes, and uploads the unfixed deck. After a passing agent run with
     fixes_applied, re-open package.draft_path (or take bytes back from the
     runner) before _renumber_slide_parts.
   - Refactor assemble_skeleton to use review_package.plan_pitch_sequence so the
     divider/product ordering has ONE implementation (§5). This is the only PR
     allowed to touch generator.py.
   - Timeout behaviour: implement §8 option (a) — on timeout, return the un-QA'd
     deck with qa.cursor_passed=false plus a warning. Do not hang the MCP call.
   - Upload review package to S3 `review-packages/{uuid}/` (short TTL prefix — no lifecycle rule required in code, document ops note)
   - Extend return dict with `qa` block per architecture §8

2. `server.py`:
   - Update `build_deck` docstring: headless Cursor QA when DECK_QA_ENABLED;
     remove "No stylist (PI-2754 shelved)" wording (line ~169)
   - Add a `except DeckQaError` arm BEFORE the existing `except ValueError`, to
     attach qa_report to the error dict (§8). DeckQaError subclasses ValueError,
     so ordering matters.

3. `.env.example` — DECK_QA_ENABLED=false, DECK_QA_TIMEOUT_S=600, CURSOR_API_KEY=

4. `tests/test_generator.py` — test build with DECK_QA_ENABLED unset (default, no
   behavior change) AND with it set to "false"/"0". Existing tests assert
   slide_count == 10; they must still pass untouched.

5. `README.md` — the Architecture section currently says "No stylist (PI-2754
   shelved)" (line ~199). Replace with the QA gate + link DECK-QA-ARCHITECTURE.md.

6. `local/schema-driven-deck-generation-engine/PROGRESS.md` IS TRACKED
   (git ls-files local returns 4 files). Update its progress table. The
   "if local/ is gitignored" conditional in the earlier draft was wrong —
   ignore it.

OUT OF SCOPE:
- Prodie
- New MCP tools for QA

ACCEPTANCE:
- Full pytest suite pass (ignore smoke live if documented)
- DECK_QA_ENABLED=false preserves current behavior

Open PR. Title: "PI-XXXX: Wire headless deck QA into build_deck"
```

---

## Optional doc-only subagent (parallel with Wave 1)

### Subagent PR-G — SoT / PROGRESS rescope

```text
[SHARED PREAMBLE]

Branch: docs/deck-qa-mvp-rescope

SCOPE:
Docs only. Rescope MVP to require headless Cursor QA.

DELIVERABLES:
`local/` is NOT gitignored — git ls-files local returns 4 tracked files
(END-SCOPE-SOT.md, PROGRESS.md, I1-DATA-SOURCES.md, PI-2757-TECH-DEBT.md).
The earlier "if local/ is gitignored" conditional was wrong; ignore it.

1. Edit ONLY:
   - `.cursor/rules/pitch-deck-end-scope.mdc` — move "Cursor stylist MVP (PI-2754)"
     out of "Not end state" and into the MVP pipeline as the headless QA gate;
     keep every product-clone hard rule verbatim
   - `local/schema-driven-deck-generation-engine/END-SCOPE-SOT.md` — same rescope
   - `docs/PRODIE-IDEATION-SPEC.md` (line ~177) also lists "Cursor stylist
     (PI-2754)" as out of scope; update for consistency

2. Do NOT touch:
   - `README.md` and `.env.example` — PR-F owns both (§13). Earlier draft told
     you to edit README; that conflicts with PR-F.
   - `local/.../PROGRESS.md` — PR-F owns it
   - `.cursor/sales-mcp-cursor-backup/` — untracked local backup copies, leave alone

3. Do NOT claim PI ticket numbers are closed; use "in progress"

OUT OF SCOPE:
- Code changes

Open PR. Title: "PI-XXXX: Rescope MVP for headless deck QA"
```

---

## Deployment order (for you, the coordinator)

```text
0. Confirm PR #31 (fix/fortuneai-deck-assembly) state; branch everything from it
1. Launch PR-A, PR-B, PR-C, PR-D (and optionally PR-G) in parallel
2. Review + merge Wave 1
3. Launch PR-E
4. Review + merge
5. Launch PR-F
6. End-to-end: DECK_QA_ENABLED=true + BambooHR golden + manual open final.pptx
```

---

## Verification checklist (human)

After all PRs land:

- [ ] `DECK_QA_ENABLED=false` — smoke_build_live still passes (10 slides, Acme + CEO Daily)
- [ ] `DECK_QA_ENABLED=true` — BambooHR schema builds, qa.deterministic_passed + qa.cursor_passed true
- [ ] Product slide unchanged after QA — **pixel-diff the rendered PNG, not the file hash.**
      `_renumber_slide_parts` runs after packaging, so draft.pptx and final.pptx
      are never byte-identical even when QA changes nothing (§8).
- [ ] **A QA fix actually reaches the delivered deck.** Force one narrative fix,
      then confirm the downloaded final.pptx contains it — this is the §8
      re-load bug, and a green qa_cursor.json does not prove it.
- [ ] Manifest spot-check: `editable: false` on every `role: product` index, and
      slide_count == 8 + P
- [ ] review package on S3 under `review-packages/` (not SharePoint)
- [ ] build_deck response includes `qa` block; a forced B3 failure returns
      `status: error` WITH `qa_report`
- [ ] End-to-end latency measured against the MCP client timeout (§8 open decision)
