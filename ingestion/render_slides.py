"""Render selected PPTX slides to PNGs for vision QA (Phase B1).

Pipeline: LibreOffice headless (PPTX → PDF) → poppler ``pdftoppm`` (PDF → PNG).
Intended for Cursor stylist / review-package scripts (B2+), not as an MCP tool.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from collections.abc import Sequence
from pathlib import Path

logger = logging.getLogger(__name__)

_DEFAULT_DPI = 150
_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


class RenderSlidesError(RuntimeError):
    """Raised when LibreOffice / poppler conversion fails."""


def _resolve_soffice(explicit: str | None = None) -> str:
    if explicit:
        path = Path(explicit)
        if not path.is_file():
            raise RenderSlidesError(f"soffice not found at {explicit!r}")
        return str(path)

    env = os.environ.get("SOFFICE_BIN")
    if env:
        return _resolve_soffice(env)

    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found

    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac.is_file():
        return str(mac)

    raise RenderSlidesError(
        "LibreOffice not found (soffice/libreoffice). "
        "Install it in the image/host or set SOFFICE_BIN."
    )


def _resolve_pdftoppm() -> str:
    found = shutil.which("pdftoppm")
    if not found:
        raise RenderSlidesError(
            "pdftoppm not found (poppler-utils). "
            "Install poppler-utils in the image/host."
        )
    return found


def _pptx_slide_count(pptx_path: Path) -> int:
    from pptx import Presentation

    return len(Presentation(str(pptx_path)).slides)


def _write_pptx(pptx: str | Path | bytes, dest: Path) -> Path:
    if isinstance(pptx, (bytes, bytearray)):
        dest.write_bytes(pptx)
        return dest
    src = Path(pptx)
    if not src.is_file():
        raise FileNotFoundError(f"PPTX not found: {src}")
    shutil.copy2(src, dest)
    return dest


def _run(cmd: list[str], *, timeout: float) -> None:
    logger.debug("running: %s", " ".join(cmd))
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise RenderSlidesError(f"command timed out: {' '.join(cmd)}") from exc
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        raise RenderSlidesError(
            f"command failed ({proc.returncode}): {' '.join(cmd)}"
            + (f"\n{err}" if err else "")
        )


def _convert_to_pdf(
    pptx_path: Path,
    out_dir: Path,
    *,
    soffice_bin: str,
    timeout: float,
) -> Path:
    # LibreOffice writes <stem>.pdf into --outdir.
    _run(
        [
            soffice_bin,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(pptx_path),
        ],
        timeout=timeout,
    )
    pdf_path = out_dir / f"{pptx_path.stem}.pdf"
    if not pdf_path.is_file():
        # Some builds rename oddly; take the only PDF in out_dir.
        pdfs = sorted(out_dir.glob("*.pdf"))
        if len(pdfs) != 1:
            raise RenderSlidesError(
                f"LibreOffice did not produce a PDF in {out_dir} (found {pdfs})"
            )
        pdf_path = pdfs[0]
    return pdf_path


def _pdf_page_to_png(
    pdf_path: Path,
    page_1based: int,
    dest_png: Path,
    *,
    pdftoppm: str,
    dpi: int,
    timeout: float,
) -> None:
    # -singlefile → dest_png.png without a page suffix.
    prefix = dest_png.with_suffix("")
    _run(
        [
            pdftoppm,
            "-png",
            "-r",
            str(dpi),
            "-f",
            str(page_1based),
            "-l",
            str(page_1based),
            "-singlefile",
            str(pdf_path),
            str(prefix),
        ],
        timeout=timeout,
    )
    if not dest_png.is_file():
        raise RenderSlidesError(f"pdftoppm did not write {dest_png}")
    if dest_png.read_bytes()[:8] != _PNG_MAGIC:
        raise RenderSlidesError(f"output is not a PNG: {dest_png}")


def render_slides(
    pptx: str | Path | bytes,
    slide_indices: Sequence[int] | None = None,
    *,
    output_dir: str | Path | None = None,
    soffice_bin: str | None = None,
    dpi: int = _DEFAULT_DPI,
    timeout: float = 180.0,
) -> list[Path]:
    """Convert selected slides of a PPTX to PNG files.

    Parameters
    ----------
    pptx:
        Path to a ``.pptx`` file, or raw PPTX bytes.
    slide_indices:
        0-based slide indices to render. ``None`` renders every slide.
        Results are returned in the same order as this sequence.
    output_dir:
        Directory for PNG outputs. Created if missing. When omitted, a
        temporary directory is created and left on disk for the caller
        (paths remain valid after return).
    soffice_bin:
        Optional path to ``soffice`` / ``libreoffice``. Otherwise uses
        ``SOFFICE_BIN`` or ``PATH``.
    dpi:
        Rasterization DPI for ``pdftoppm`` (default 150).
    timeout:
        Per-subprocess timeout in seconds.

    Returns
    -------
    list[pathlib.Path]
        PNG paths named ``slide-000.png``, ``slide-001.png``, … matching
        each requested index.

    Raises
    ------
    FileNotFoundError
        PPTX path does not exist.
    ValueError
        Empty / out-of-range / duplicate-unfriendly index list issues.
    RenderSlidesError
        Missing tools or conversion failure.
    """
    if dpi < 36 or dpi > 600:
        raise ValueError(f"dpi out of range: {dpi}")

    soffice = _resolve_soffice(soffice_bin)
    pdftoppm = _resolve_pdftoppm()

    if output_dir is None:
        out = Path(tempfile.mkdtemp(prefix="render_slides_"))
    else:
        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)

    work = Path(tempfile.mkdtemp(prefix="render_slides_work_"))
    try:
        pptx_path = _write_pptx(pptx, work / "deck.pptx")
        slide_count = _pptx_slide_count(pptx_path)
        if slide_count < 1:
            raise ValueError("PPTX has no slides")

        if slide_indices is None:
            indices = list(range(slide_count))
        else:
            indices = list(slide_indices)
            if not indices:
                raise ValueError("slide_indices is empty")
            bad = [i for i in indices if i < 0 or i >= slide_count]
            if bad:
                raise ValueError(
                    f"slide_indices out of range for {slide_count}-slide deck: {bad}"
                )

        pdf_path = _convert_to_pdf(
            pptx_path, work, soffice_bin=soffice, timeout=timeout
        )

        results: list[Path] = []
        for idx in indices:
            dest = out / f"slide-{idx:03d}.png"
            _pdf_page_to_png(
                pdf_path,
                idx + 1,
                dest,
                pdftoppm=pdftoppm,
                dpi=dpi,
                timeout=timeout,
            )
            results.append(dest)

        logger.info(
            "render_slides: wrote %d PNG(s) to %s (indices=%s)",
            len(results),
            out,
            indices,
        )
        return results
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _main(argv: list[str] | None = None) -> int:
    """CLI for Cursor agent scripts: ``python -m ingestion.render_slides``."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Render selected PPTX slides to PNGs (LibreOffice + pdftoppm)."
    )
    parser.add_argument("pptx", type=Path, help="Path to .pptx")
    parser.add_argument(
        "--indices",
        "-i",
        help="Comma-separated 0-based slide indices (default: all)",
    )
    parser.add_argument(
        "--output-dir",
        "-o",
        type=Path,
        required=True,
        help="Directory for slide-NNN.png outputs",
    )
    parser.add_argument("--dpi", type=int, default=_DEFAULT_DPI)
    parser.add_argument("--soffice", default=None, help="Path to soffice binary")
    args = parser.parse_args(argv)

    indices: list[int] | None
    if args.indices is None or args.indices.strip() == "":
        indices = None
    else:
        indices = [int(x.strip()) for x in args.indices.split(",") if x.strip()]

    paths = render_slides(
        args.pptx,
        indices,
        output_dir=args.output_dir,
        soffice_bin=args.soffice,
        dpi=args.dpi,
    )
    for p in paths:
        print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
