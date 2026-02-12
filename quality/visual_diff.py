"""Visual diff engine — render PDF & DOCX to images and compute SSIM.

Provides pixel-level comparison between the original PDF and converted
DOCX by rendering both to PNG images and computing the Structural
Similarity Index (SSIM).  Optionally generates a diff overlay image
highlighting the areas of difference.

Requirements
------------
- PyMuPDF  (for PDF → PNG rendering)
- scikit-image  (for SSIM computation)
- Pillow  (for image manipulation)
- LibreOffice  (for DOCX → PDF → PNG pipeline)
"""

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageChops, ImageDraw

logger = logging.getLogger(__name__)

# Default rendering DPI.
_DEFAULT_DPI = 150


def find_libreoffice() -> str | None:
    """Auto-detect the LibreOffice ``soffice`` executable path.

    Checks common installation directories on Windows, macOS, and Linux.
    Returns ``None`` if LibreOffice is not found.
    """
    candidates = [
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Linux
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/snap/bin/libreoffice",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    # Fall back to PATH lookup.
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    return exe


class VisualDiff:
    """Compare a PDF and its converted DOCX visually.

    Parameters
    ----------
    dpi : int
        Resolution for rendering pages (default 150).
    libreoffice_path : str | None
        Explicit path to ``soffice``.  ``"auto"`` (default) triggers
        auto-detection.
    """

    def __init__(
        self,
        dpi: int = _DEFAULT_DPI,
        libreoffice_path: str | None = "auto",
    ) -> None:
        self.dpi = dpi
        if libreoffice_path == "auto":
            self._lo_path = find_libreoffice()
        else:
            self._lo_path = libreoffice_path

        if self._lo_path:
            logger.info("LibreOffice found: %s", self._lo_path)
        else:
            logger.warning(
                "LibreOffice not found.  DOCX → PNG rendering will not work.  "
                "Install LibreOffice or set libreoffice_path in config."
            )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def compare(
        self,
        pdf_path: str,
        docx_path: str,
        output_dir: str | None = None,
    ) -> dict[str, Any]:
        """Run full visual comparison between *pdf_path* and *docx_path*.

        Parameters
        ----------
        pdf_path : str
            Path to the original PDF.
        docx_path : str
            Path to the converted DOCX.
        output_dir : str | None
            Directory for rendered images and diff overlays.
            A temporary directory is used if not specified.

        Returns
        -------
        dict
            ``page_scores``: list of per-page SSIM scores (0.0–1.0).
            ``overall_score``: mean SSIM across all pages.
            ``quality_level``: ``"green"`` (≥ 95%), ``"yellow"`` (≥ 85%),
            or ``"red"`` (< 85%).
            ``diff_images``: list of paths to diff overlay PNGs.
            ``pdf_images``: list of paths to rendered PDF page PNGs.
            ``docx_images``: list of paths to rendered DOCX page PNGs.
        """
        use_temp = output_dir is None
        if use_temp:
            output_dir = tempfile.mkdtemp(prefix="vdiff_")

        os.makedirs(output_dir, exist_ok=True)

        # 1. Render PDF pages.
        pdf_images = self.render_pdf(pdf_path, output_dir)
        logger.info("Rendered %d PDF page(s) at %d DPI.", len(pdf_images), self.dpi)

        # 2. Render DOCX pages.
        docx_images = self.render_docx(docx_path, output_dir)
        logger.info("Rendered %d DOCX page(s).", len(docx_images))

        # 3. Compute SSIM per page (only for pages both documents have).
        page_scores: list[float] = []
        diff_images: list[str] = []

        num_compare = min(len(pdf_images), len(docx_images))
        for i in range(num_compare):
            score, diff_path = self._compare_page(
                pdf_images[i], docx_images[i], output_dir, i
            )
            page_scores.append(score)
            diff_images.append(diff_path)

        # Note if page counts differ.
        pdf_pages = len(pdf_images)
        docx_pages = len(docx_images)
        if pdf_pages != docx_pages:
            logger.warning(
                "Page count mismatch: PDF has %d page(s), DOCX has %d page(s).",
                pdf_pages,
                docx_pages,
            )

        overall = sum(page_scores) / len(page_scores) if page_scores else 0.0

        # Penalize if DOCX has extra pages (content overflow).
        if docx_pages > pdf_pages and pdf_pages > 0:
            overflow_penalty = 0.05 * (docx_pages - pdf_pages)
            overall = max(0.0, overall - overflow_penalty)
            logger.info(
                "Overflow penalty: %.1f%% (DOCX has %d extra page(s)).",
                overflow_penalty * 100,
                docx_pages - pdf_pages,
            )

        if overall >= 0.95:
            level = "green"
        elif overall >= 0.85:
            level = "yellow"
        else:
            level = "red"

        logger.info(
            "Visual diff complete — overall SSIM: %.3f (%s)", overall, level
        )

        return {
            "page_scores": page_scores,
            "overall_score": overall,
            "quality_level": level,
            "diff_images": diff_images,
            "pdf_images": pdf_images,
            "docx_images": docx_images,
            "pdf_page_count": pdf_pages,
            "docx_page_count": docx_pages,
        }

    # ------------------------------------------------------------------
    # PDF rendering
    # ------------------------------------------------------------------

    def render_pdf(self, pdf_path: str, output_dir: str) -> list[str]:
        """Render each page of *pdf_path* to a PNG at ``self.dpi``."""
        doc = fitz.open(pdf_path)
        images: list[str] = []
        zoom = self.dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        for i in range(doc.page_count):
            page = doc[i]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out_path = os.path.join(output_dir, f"pdf_page_{i}.png")
            pix.save(out_path)
            images.append(out_path)

        doc.close()
        return images

    # ------------------------------------------------------------------
    # DOCX rendering (via LibreOffice)
    # ------------------------------------------------------------------

    def render_docx(self, docx_path: str, output_dir: str) -> list[str]:
        """Convert DOCX → PDF via LibreOffice, then render pages to PNG.

        Falls back to an empty list if LibreOffice is not available.
        """
        if not self._lo_path:
            logger.error("Cannot render DOCX: LibreOffice not found.")
            return []

        # Convert DOCX → PDF using LibreOffice headless.
        tmp_dir = tempfile.mkdtemp(prefix="lo_convert_")
        try:
            cmd = [
                self._lo_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", tmp_dir,
                os.path.abspath(docx_path),
            ]
            logger.debug("Running: %s", " ".join(cmd))
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode != 0:
                logger.error(
                    "LibreOffice conversion failed (rc=%d): %s",
                    result.returncode,
                    result.stderr,
                )
                return []

            # Find the generated PDF.
            base = Path(docx_path).stem + ".pdf"
            pdf_out = os.path.join(tmp_dir, base)
            if not os.path.isfile(pdf_out):
                # Fallback: look for any PDF in tmpdir.
                pdfs = list(Path(tmp_dir).glob("*.pdf"))
                if pdfs:
                    pdf_out = str(pdfs[0])
                else:
                    logger.error("No PDF output from LibreOffice.")
                    return []

            # Render the intermediate PDF to PNG in its own temp dir
            # to avoid overwriting the original PDF page renders.
            docx_render_dir = tempfile.mkdtemp(prefix="docx_render_")
            try:
                raw_images = self.render_pdf(pdf_out, docx_render_dir)

                # Move rendered files to the output dir with docx_ prefix.
                renamed: list[str] = []
                for i, img_path in enumerate(raw_images):
                    new_name = os.path.join(output_dir, f"docx_page_{i}.png")
                    shutil.copy2(img_path, new_name)
                    renamed.append(new_name)
            finally:
                shutil.rmtree(docx_render_dir, ignore_errors=True)

            return renamed

        except subprocess.TimeoutExpired:
            logger.error("LibreOffice conversion timed out.")
            return []
        except Exception:
            logger.exception("DOCX rendering failed.")
            return []
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    # ------------------------------------------------------------------
    # Per-page comparison
    # ------------------------------------------------------------------

    def _compare_page(
        self,
        pdf_img_path: str,
        docx_img_path: str,
        output_dir: str,
        page_num: int,
    ) -> tuple[float, str]:
        """Compare two page images and return (SSIM, diff_image_path)."""
        try:
            from skimage.metrics import structural_similarity as ssim

            img_a = Image.open(pdf_img_path).convert("RGB")
            img_b = Image.open(docx_img_path).convert("RGB")

            # Resize to matching dimensions (use the larger canvas).
            target_w = max(img_a.width, img_b.width)
            target_h = max(img_a.height, img_b.height)

            canvas_a = Image.new("RGB", (target_w, target_h), (255, 255, 255))
            canvas_a.paste(img_a, (0, 0))
            canvas_b = Image.new("RGB", (target_w, target_h), (255, 255, 255))
            canvas_b.paste(img_b, (0, 0))

            arr_a = np.array(canvas_a)
            arr_b = np.array(canvas_b)

            # Compute SSIM with a difference map.
            score, diff_map = ssim(
                arr_a, arr_b, full=True, channel_axis=2,
                data_range=255,
            )

            # Build the diff overlay: highlight regions where diff < 0.9.
            diff_gray = np.mean(diff_map, axis=2)
            mask = diff_gray < 0.9
            overlay = canvas_a.copy()
            draw = ImageDraw.Draw(overlay)

            # Mark difference regions with semi-transparent red.
            diff_img = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
            diff_draw = ImageDraw.Draw(diff_img)

            # Find difference regions and draw red boxes.
            ys, xs = np.where(mask)
            if len(xs) > 0:
                # Group nearby pixels into bounding boxes using a simple grid.
                grid = 20
                boxes: set[tuple[int, int, int, int]] = set()
                for y, x in zip(ys, xs):
                    bx = (x // grid) * grid
                    by = (y // grid) * grid
                    boxes.add((bx, by, bx + grid, by + grid))
                for box in boxes:
                    diff_draw.rectangle(box, fill=(255, 0, 0, 80))

            overlay = Image.alpha_composite(
                canvas_a.convert("RGBA"), diff_img
            )

            diff_path = os.path.join(output_dir, f"diff_page_{page_num}.png")
            overlay.save(diff_path)

            return float(score), diff_path

        except ImportError:
            logger.error("scikit-image not installed.  Cannot compute SSIM.")
            return 0.0, ""
        except Exception:
            logger.exception("SSIM comparison failed for page %d.", page_num)
            return 0.0, ""
