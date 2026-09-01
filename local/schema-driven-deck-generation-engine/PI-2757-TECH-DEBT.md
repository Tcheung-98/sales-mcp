# PI-2757 tech debt — placeholder pipeline

> **Ticket:** [PI-2757](https://fortune.atlassian.net/browse/PI-2757) / C2  
> **PR:** https://github.com/Tcheung-98/sales-mcp/pull/26  
> **Captured:** 2026-08-19  
> **Status:** C2 feature complete (deterministic + bounded AI). CI + mock-AI smoke on real S3 FortuneAI pass. Live Claude manual QA recommended before prod.  
> **Related locks:** [`C2-PLACEHOLDER-INVENTORY.md`](C2-PLACEHOLDER-INVENTORY.md) · [`PROGRESS.md`](PROGRESS.md)

---

## Verified (2026-08-19)

- `tests/smoke_build_live.py --mock-ai` — real S3 FortuneAI + GTM Audience Data; mocked Claude; CEO Daily mix → 10 slides, no leftover tokens, CEO DAILY clone intact.
- Full pytest suite green (229 passed).

---

## Remaining gaps (follow-on, not C2 blockers)

| Item | Notes |
|---|---|
| SharePoint / Graph logo fetch | C2 uses HTTPS `client_logo` only; SharePoint library paths fail loud. |
| C1 divider order vs live PPTX | Live slides 13–17 order may not match Workflow captions; C1 follow-up, not C2. |
| Live Claude end-to-end | Run `tests/smoke_build_live.py` without `--mock-ai` and open the PPTX before prod deploy. |
| Prodie rewire | PI-2350 — seller-facing `build_deck` wiring separate from this ticket. |

Do not spend a cycle rewriting fill logic until a live-template probe disagrees. Fail-loud is correct.
