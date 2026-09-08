#!/usr/bin/env python3
"""CLI shim for the B4 headless Cursor QA runner (docs/DECK-QA-ARCHITECTURE.md §7).

The implementation lives in ``ingestion.deck_qa_agent`` because ``scripts/`` is not
in ``[tool.setuptools] packages`` and therefore is not importable from an installed
wheel or the Docker image; PR-F imports ``run_headless_cursor_qa`` in-process (§8).

    PYTHONPATH=. uv run python scripts/run_deck_qa.py --package-dir review-packages/x
"""

from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from ingestion.deck_qa_agent import main  # noqa: E402 — needs the repo root on sys.path

if __name__ == "__main__":
    raise SystemExit(main())
