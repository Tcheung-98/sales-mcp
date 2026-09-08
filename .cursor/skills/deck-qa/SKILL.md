---
name: deck-qa
description: >-
  Vision QA pass (B4) on a generated Fortune pitch deck review package. Use
  after deterministic QA (B3) passes, when given a review package directory
  containing draft.pptx, manifest.json, slides/slide-NNN.png and
  qa_deterministic.json. Reviews rendered slides for leftover placeholders,
  overflow, alignment and weak copy, applies at most one bounded fix pass, and
  writes qa_cursor.json.
---

# Deck QA (B4 headless vision pass)

You are the styling gate between C2 placeholder fills and delivery of a Fortune
pitch deck. Architecture source of truth: `docs/DECK-QA-ARCHITECTURE.md` (§4
hard rules, §5 review package contract, §7 this pass). Read it if anything here
is ambiguous, but this file is meant to be enough on its own.

Your output is a `qa_cursor.json` verdict, plus at most **one** round of fixes
written into `draft.pptx` inside the package directory.

## When to use

Run this pass only when **all** of the following hold:

- You have a review package directory (local path or a synced copy of
  `review-packages/{review_id}/`).
- `qa_deterministic.json` exists and has `"passed": true`. If it is `false`,
  **stop** — B3 failures are caller-fixable errors, not styling problems. The
  only exception is an explicit `--force` dev run, which you should note in
  `issues[]`.
- You have not already run a fix loop on this package (`loop_count` cap of 1,
  see Exit contract).

Do not run this pass to "improve" a deck on request outside the QA rail, and do
not run it as an MCP tool — `run_deck_qa` is deliberately not exposed to Prodie
or to associates.

## Inputs

```text
review-packages/{review_id}/
  draft.pptx              # post-C2 deck; the file you edit if you fix anything
  manifest.json           # ReviewManifest v1 — the authority on what you may touch
  deck_schema.json        # serialized DeckSchema; context for copy, and needed to re-run PlaceholderAI
  slides/
    slide-000.png         # 0-based index, matches slide_index in the manifest, 150 DPI
    slide-001.png
    ...
  qa_deterministic.json   # B3 output; read before you start
  qa_cursor.json          # your output; you create it
```

### The manifest is the authority — never slide position

Every decision about whether you may touch a slide comes from `manifest.json`.

- **Key off `editable`.** `editable: true` means you may restyle and reword it.
  `editable: false` means hands off, flag only.
- **Never infer permission from slide position.** The deck is `8 + P` slides
  where `P = funded dividers + confirmed products`, so the pitch section's
  extent moves with the mix. An earlier draft of the spec mapped "indices 1–11"
  to narrative, which would have marked A5 product clones editable. Do not
  reconstruct the index map; read the flag.
- **`role: other` is ambiguous by itself.** Dividers, the investment slide and
  the thank-you slide all carry `role: other`. Use the `slide_kind`
  discriminator to tell them apart: `"divider"`, `"investment"`, `"thank_you"`.

| `role` | `slide_kind` | What it is | You may edit |
|--------|--------------|------------|--------------|
| `cover` | — | Intro slide: title, date, client logo | Yes |
| `narrative` | — | Stock FortuneAI spine: Why Fortune, History of Trust, Opportunity, Audience, Program Overview | Yes |
| `other` | `divider` | Category divider (High-Impact, Editorial, Premium Video, Print, Branded Content) | Yes, text only |
| `other` | `investment` | Investment / pricing slide | Yes, text only |
| `other` | `thank_you` | Closing slide | Yes, text only |
| `product` | — | A5 product clone lifted verbatim from a GTM source deck | **No — flag only** |

A `product` entry also carries `product_name`, `source_path` (the raw GTM
`Deck Path`, e.g. `Fortune_Newsletters_2026.pptx`) and `source_slide_number`.
Those are provenance for your issue messages, not an invitation to open the
source deck.

## What to review

Load `slides/slide-NNN.png` for every manifest entry and view it. Review
**editable** slides for fixable problems; still *look at* `editable: false`
product clones so you can report on them, but never fix them.

Checklist for stock (editable) slides:

1. **Visible leftover placeholders.** Any bracketed token still rendered on the
   slide — `[TITLE]`, `[HEADER]`, `[BODY]`, `[AUDIENCE TITLE]`, `[DATE]`,
   `[AUDIENCE SEGMENT]`, `[REACH]`, `[INDEX]`, `[BUDGET]`,
   `[PRODUCT CATEGORY]`, `[CLIENT NAME]`, `[client name]`, `[CLIENT NAME’S]`
   (curly apostrophe — `pptx_tools.CLIENT_NAME_POSSESSIVE_TOKEN`), `[LOGO]`,
   `[PRODUCT + PRODUCT PRICE]` — or stock literals that were meant to
   be replaced: `PRODUCT TYPE`, `Product description.`. B3 already scans the
   XML for these, so a hit here usually means the token is inside an image or
   grouped shape B3 could not see. Report it either way.
2. **Text overflow / truncation.** Copy running past its text box, clipped
   descenders, text colliding with the next shape, an auto-shrunk paragraph
   that has become unreadably small.
3. **Broken alignment.** Shapes visibly knocked off the template grid, a body
   block no longer left-aligned with its header, ragged card rows on the
   audience slide, a divider caption off-center.
4. **Unreadable text.** Low contrast against the background image, text under a
   logo or gradient, font too small to read at presentation size.
5. **Obvious off-brand issues.** Wrong-looking colours against the FortuneAI
   template, a stretched or distorted client logo, mixed fonts, an em dash in
   generated copy (C2 forbids them).
6. **Weak copy on allowed slots.** Only for the named AI slots listed under
   Fixes. Generic filler, copy that does not name the client's actual
   opportunity, a program blurb that describes nothing specific, an intro title
   that restates the RFP mechanics. Weak-but-valid copy is a `warning`, not a
   failure — do not churn copy that is merely not to your taste.

Log everything you find in `issues[]`, including things you cannot fix.

## Allowed fixes

You get **one** fix pass. Fix only what is on an `editable: true` slide.

### Step 1 — regenerate narrative copy with `PlaceholderAI`

For copy problems on a named AI slot, re-invoke the same bounded generator C2
used, from `ingestion/placeholder_ai.py`, with the schema from
`deck_schema.json`. Do not hand-write replacement copy for these slots and do
not relax the bounds — the validators re-run and will reject it.

| Slot id (use in `fixes_applied`) | Method | Bounds enforced by C2 |
|----------------------------------|--------|-----------------------|
| `intro_title` | `PlaceholderAI.intro_title(schema)` | 3–6 words, ALL CAPS, no em dash, no client name |
| `opportunity_header` | `PlaceholderAI.opportunity_header(schema)` | 3–6 words, sentence case, no em dash |
| `opportunity_body` | `PlaceholderAI.opportunity_body(schema)` | 75–95 words, no em dash, closing sentence starts with "Fortune" |
| `audience_title` | `PlaceholderAI.audience_title(schema, segments)` | 3–6 words, sentence case, no job titles |
| `program_blurb_<Category>` | `PlaceholderAI.program_blurb(schema, category_name)` | 10–15 words, no em dash |

Each of these validates and retries once internally, then raises. If a slot
raises after its retry, do not fall back to hand-written copy — record the
failure as an `error` issue and let the pass fail.

### Step 2 — write the text into the deck

Use `ingestion.pptx_tools`:

- `set_ph_text(ph, text)` — replace a placeholder's text, preserving run
  formatting. Use when you have the placeholder shape itself.
- `replace_token(slide, token, text)` — replace every occurrence of an exact
  token string across the slide's text frames. Returns a hit count; a return of
  `0` means you targeted the wrong slide.
- `replace_first_token(slide, token, text)` — same, first occurrence only.
  Use for repeated tokens filled row-by-row, such as the audience cards.

Then save `draft.pptx` in place and re-render only the slides you changed:

```bash
python -m ingestion.render_slides <package>/draft.pptx \
  --indices 3,5 --output-dir <package>/slides
```

Re-view the regenerated PNGs to confirm the fix landed and did not create new
overflow. That verification is part of the same single loop, not a second one.

### Do not use `apply_replacements`

`pptx_tools.apply_replacements` looks like the obvious fix helper. It is not,
and it will silently damage the deck.

It keys off placeholder `idx` 0 (title) and `idx` 19 (eyebrow) and rewrites body
paragraphs by position. That contract came from Titan-retrieved corpus slides,
where the layout guarantees those indices. FortuneAI stock slides carry named
text **tokens** in ordinary text boxes, not indexed placeholders — so
`apply_replacements` either finds nothing and no-ops, or matches an unrelated
placeholder and clobbers the wrong runs. Use `replace_token` /
`replace_first_token` / `set_ph_text` instead. See §4 of the architecture doc.

## Forbidden

Violating any of these is a hard failure of the pass, not a judgement call.

1. **Never edit a slide with `editable: false`.** These are A5 product clones,
   reproduced verbatim from the GTM source deck at an exact `Deck Path` and
   `Slide #`. No text, layout, image, or formatting change — not even fixing an
   obvious typo. Flag it in `issues[]` with the `slide_index` and
   `product_name` and move on.
2. **Never reorder slides.** Divider order is the Workflow pitch order
   (High-Impact → Editorial → Premium Video → Print → Branded Content) and is
   fixed upstream in C1.
3. **Never add or delete slides.** The deck is `8 + P` slides when you receive
   it and must be `8 + P` slides when you finish.
4. **Never change the product mix.** No adding, removing, substituting or
   repricing products. The mix is the associate's locked selection.
5. **No corpus lookups.** No `search_decks`, no Titan/RAG retrieval, no opening
   product source decks to find a "better" slide. If a product slide is bad,
   that is a GTM data issue for a human.
6. **No browser automation.** Everything you need is in the package directory.
7. **No re-running C1 assembly or full `apply_placeholders`.** You are patching
   a finished deck, not rebuilding it.

## Exit contract

Write `qa_cursor.json` into the package directory. Always write it — including
when you fail.

```json
{
  "passed": true,
  "loop_count": 1,
  "issues": [
    {"slide_index": 4, "severity": "warning", "message": "Audience card 3 reach value sits 4pt below its label baseline."}
  ],
  "fixes_applied": ["opportunity_body", "program_blurb_Editorial Alignment"]
}
```

- `loop_count` — `0` if you reviewed and changed nothing, `1` if you ran the fix
  pass. Never more than `1`. If problems remain after one fix pass, set
  `passed: false` and stop; the caller raises `DeckQaError` rather than letting
  you loop again.
- `severity` — one of `info`, `warning`, `error`. Only an `error` on an
  `editable: true` slide sets `passed: false`. Issues on `editable: false`
  product clones are **flag-only**: report them at `warning` at most, and never
  let them fail the deck, since you are forbidden from fixing them and a hard
  failure would block delivery on something no one in this pass can repair.
- `fixes_applied` — the slot ids from the fix table (or a short stable label for
  a deterministic token fix, e.g. `investment_budget_token`). Empty list when
  you changed nothing.

### `fixes_applied` must be truthful

This is the one field that can silently ship a broken deck. You edit
`draft.pptx` on disk, but the caller holds its own in-memory `Presentation` and
**only re-loads from disk when `fixes_applied` is non-empty**
(`docs/DECK-QA-ARCHITECTURE.md` §8). So:

- If you wrote any change to `draft.pptx`, `fixes_applied` must be non-empty, or
  your fix is discarded and the unfixed deck is delivered with a green report.
- If you did not change the file, `fixes_applied` must be empty, or the caller
  re-loads a stale draft and throws away nothing — but also masks whether QA
  actually did anything.

Report what you did, exactly.

## Reference

- `docs/DECK-QA-ARCHITECTURE.md` — SoT. §4 hard rules and fix-method rationale,
  §5 review package + manifest contract and slide geometry, §6 deterministic QA
  checks, §7 this pass, §8 how the build path consumes your report.
- `.cursor/rules/pitch-deck-end-scope.mdc` — product-level end scope and the
  product-clone hard rules.
