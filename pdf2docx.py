#!/usr/bin/env python3
"""PDF to Word Conversion Tool — Main Orchestrator.

Converts a PDF file to a high-fidelity Word document (.docx), preserving
text formatting, images, tables, headings, lists, and layout as closely
as possible.

Usage
-----
    python pdf2docx.py input.pdf output.docx
    python pdf2docx.py --ocr scanned.pdf output.docx
    python pdf2docx.py --verbose --validate input.pdf output.docx
    python pdf2docx.py --batch C:\\pdfs\\ --output-dir C:\\output\\
    python pdf2docx.py --password "secret" encrypted.pdf output.docx
"""

import argparse
import glob
import json
import logging
import os
import sys
import tempfile
from typing import Any, Callable

import fitz  # PyMuPDF

from extractors.text_extractor import TextExtractor
from extractors.image_extractor import ImageExtractor
from extractors.table_extractor import TableExtractor
from extractors.metadata_extractor import MetadataExtractor
from analyzers.semantic_analyzer import SemanticAnalyzer
from builders.docx_builder import DocxBuilder
from utils.pdf_info import PDFInfo
from utils.ocr_handler import OCRHandler
from utils.progress import ProgressTracker
from utils.validator import OutputValidator
from quality.visual_diff import VisualDiff
from quality.ai_comparator import AIComparator
from quality.correction_engine import CorrectionEngine
from quality.ai_layout_analyzer import AILayoutAnalyzer

logger = logging.getLogger("pdf2docx")

# Path to the default config file shipped alongside this script.
_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config", "settings.json")


# ------------------------------------------------------------------ #
#  Helpers                                                            #
# ------------------------------------------------------------------ #

def _load_config(config_path: str | None = None) -> dict[str, Any]:
    """Load configuration from a JSON file, falling back to defaults."""
    path = config_path or _CONFIG_PATH
    defaults: dict[str, Any] = {
        "page_size": "auto",
        "ocr_enabled": False,
        "ocr_language": "eng",
        "preserve_fonts": True,
        "image_quality": "original",
        "table_detection": "auto",
        "heading_detection": "both",
        "skip_watermarks": False,
        "verbose": False,
        "fallback_font": "Arial",
        "max_image_dpi": 300,
        "skip_ocr_if_no_tesseract": True,
    }
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                user = json.load(fh)
            defaults.update(user)
        except Exception:
            logger.warning("Could not load config from '%s'; using defaults.", path)
    return defaults


def _normalize_blocks(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Unpack the ``bbox`` tuple into separate ``x0, y0, x1, y1`` keys.

    The extractors store ``bbox`` as ``(x0, y0, x1, y1)``, but the
    analyzers expect individual keys for sorting and comparison.
    """
    for b in blocks:
        bbox = b.get("bbox")
        if bbox and len(bbox) == 4:
            b["x0"], b["y0"], b["x1"], b["y1"] = bbox
    return blocks


def _filter_text_in_tables(
    text_blocks: list[dict[str, Any]],
    tables: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Remove text blocks whose centre point falls inside a table bbox.

    This prevents table cell text from being duplicated as standalone
    body paragraphs.
    """
    if not tables:
        return text_blocks

    filtered: list[dict[str, Any]] = []
    for block in text_blocks:
        bx0 = block.get("x0", 0)
        by0 = block.get("y0", 0)
        bx1 = block.get("x1", 0)
        by1 = block.get("y1", 0)
        if bx1 == 0 and by1 == 0:
            # Fallback: try bbox tuple.
            bbox = block.get("bbox")
            if bbox and len(bbox) == 4:
                bx0, by0, bx1, by1 = bbox
            else:
                filtered.append(block)
                continue
        cx = (bx0 + bx1) / 2
        cy = (by0 + by1) / 2
        block_page = block.get("page_num", -1)
        in_table = False
        for tbl in tables:
            # Only filter text on the same page as the table.
            if tbl.get("page_num", -2) != block_page:
                continue
            tx0 = tbl.get("x0", 0)
            ty0 = tbl.get("y0", 0)
            tx1 = tbl.get("x1", 0)
            ty1 = tbl.get("y1", 0)
            if tx1 == 0 and ty1 == 0:
                tbl_bbox = tbl.get("bbox")
                if tbl_bbox and len(tbl_bbox) == 4:
                    tx0, ty0, tx1, ty1 = tbl_bbox
                else:
                    continue
            if tx0 <= cx <= tx1 and ty0 <= cy <= ty1:
                in_table = True
                break
        if not in_table:
            filtered.append(block)
    return filtered


def _update_metadata_margins(
    metadata: dict[str, Any],
    text_blocks: list[dict[str, Any]],
    images: list[dict[str, Any]],
    tables: list[dict[str, Any]],
) -> None:
    """Compute per-page content margins and store in *metadata*.

    Scans all extracted items for their bounding boxes to determine
    how close the content gets to the page edges.  The computed margins
    are stored under ``metadata["pages"][n]["margins"]`` as a dict with
    ``top``, ``bottom``, ``left``, ``right`` in PDF points.  A minimum
    of 14 pt (~0.2 in) is enforced.
    """
    _MIN_MARGIN = 14.0
    pages = metadata.get("pages", [])
    if not pages:
        return

    # Collect bounding boxes per page.
    per_page: dict[int, list[tuple[float, float, float, float]]] = {}

    for block in text_blocks:
        pn = block.get("page_num", 0)
        x0, y0, x1, y1 = _get_bbox(block)
        if x1 > x0 and y1 > y0:
            per_page.setdefault(pn, []).append((x0, y0, x1, y1))

    for img in images:
        pn = img.get("page_num", 0)
        x0, y0, x1, y1 = _get_bbox(img)
        if x1 > x0 and y1 > y0:
            per_page.setdefault(pn, []).append((x0, y0, x1, y1))

    for tbl in tables:
        pn = tbl.get("page_num", 0)
        x0, y0, x1, y1 = _get_bbox(tbl)
        if x1 > x0 and y1 > y0:
            per_page.setdefault(pn, []).append((x0, y0, x1, y1))

    for pn, bboxes in per_page.items():
        if pn >= len(pages):
            continue
        pg = pages[pn]
        pw = float(pg.get("width", 612))
        ph = float(pg.get("height", 792))

        min_x = min(b[0] for b in bboxes)
        min_y = min(b[1] for b in bboxes)
        max_x = max(b[2] for b in bboxes)
        max_y = max(b[3] for b in bboxes)

        pg["margins"] = {
            "top": max(_MIN_MARGIN, min_y),
            "bottom": max(_MIN_MARGIN, ph - max_y),
            "left": max(_MIN_MARGIN, min_x),
            "right": max(_MIN_MARGIN, pw - max_x),
        }


def _get_bbox(item: dict[str, Any]) -> tuple[float, float, float, float]:
    """Extract bounding box from an extracted item dict."""
    bbox = item.get("bbox")
    if bbox and len(bbox) >= 4:
        return float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])
    return (
        float(item.get("x0", 0)),
        float(item.get("y0", 0)),
        float(item.get("x1", 0)),
        float(item.get("y1", 0)),
    )


# ------------------------------------------------------------------ #
#  Core pipeline                                                      #
# ------------------------------------------------------------------ #

def convert_pdf(
    input_path: str,
    output_path: str,
    config: dict[str, Any],
    password: str | None = None,
    validate: bool = False,
    visual_validate: bool = False,
    ai_compare: bool = False,
    ai_strategy: str = "A",
    progress_callback: Callable[[str, int, int, str], None] | None = None,
) -> str:
    """Run the full conversion pipeline for a single PDF.

    Parameters
    ----------
    input_path : str
        Path to the source PDF.
    output_path : str
        Destination path for the ``.docx`` file.
    config : dict
        Application configuration.
    password : str | None
        Optional password for encrypted PDFs.
    validate : bool
        If ``True``, run validation on the output file.
    visual_validate : bool
        If ``True``, render both documents and compute SSIM scores.
    ai_compare : bool
        If ``True``, use Gemini AI to detect visual differences and
        auto-correct them (requires ``GEMINI_API_KEY``).
    ai_strategy : str
        ``"A"`` for post-build correction loop (default),
        ``"B"`` for pre-build AI-guided layout analysis.
    progress_callback : callable | None
        Optional ``(stage, current, total, message) -> None`` callback
        invoked at each pipeline stage so callers (e.g. Web UI) can
        report progress.

    Returns
    -------
    str
        The absolute path of the generated ``.docx``.
    """
    def _progress(stage: str, current: int, total: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(stage, current, total, message)

    # Auto-flushing print so output appears immediately in web server threads.
    def _print(*args: Any, **kwargs: Any) -> None:
        kwargs.setdefault("flush", True)
        print(*args, **kwargs)

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    logger.info("Converting '%s' → '%s'", input_path, output_path)

    if not os.path.isfile(input_path):
        logger.error("Input file not found: '%s'", input_path)
        sys.exit(1)

    # ── Stage 0: Quick analysis ──────────────────────────────────────
    pdf_info_analyzer = PDFInfo()
    info = pdf_info_analyzer.analyze(input_path)
    logger.info(
        "PDF info — pages: %d, text-based: %s, has images: %s, has tables: %s",
        info.get("page_count", 0),
        info.get("is_text_based"),
        info.get("has_images"),
        info.get("has_tables"),
    )

    # ── Open PDF document ────────────────────────────────────────────
    try:
        doc = fitz.open(input_path)
    except Exception as exc:
        logger.error("Cannot open PDF '%s': %s", input_path, exc)
        sys.exit(1)

    if doc.is_encrypted:
        if password:
            if not doc.authenticate(password):
                logger.error("Incorrect password for '%s'.", input_path)
                doc.close()
                sys.exit(1)
        else:
            logger.error("PDF is encrypted. Supply password with --password.")
            doc.close()
            sys.exit(1)

    page_count = doc.page_count
    progress = ProgressTracker(total=page_count, description="Extracting pages")

    # ── Stage 1: Extract metadata ────────────────────────────────────
    _progress("analyzing", 0, page_count, "Extracting metadata…")
    meta_extractor = MetadataExtractor()
    metadata = meta_extractor.extract(doc)

    # ── Stage 2: Content extraction (per-page) ───────────────────────
    text_extractor = TextExtractor()
    image_extractor = ImageExtractor(doc)
    table_extractor = TableExtractor()
    ocr_handler = OCRHandler()

    all_text_blocks: list[dict[str, Any]] = []
    all_images: list[dict[str, Any]] = []
    all_tables: list[dict[str, Any]] = []
    all_links: list[dict[str, Any]] = []

    # Create a temporary directory for extracted images.
    temp_dir = tempfile.mkdtemp(prefix="pdf2docx_imgs_")

    ocr_enabled = config.get("ocr_enabled", False)
    ocr_language = config.get("ocr_language", "eng")

    for page_num in range(page_count):
        page = doc[page_num]
        progress.set_description(f"Page {page_num + 1}/{page_count}")

        # ── Text ──
        _progress("extracting", page_num + 1, page_count,
                  f"Extracting page {page_num + 1}/{page_count}")
        text_blocks = text_extractor.extract(page, page_num)

        # If page has barely any text and OCR is enabled, try OCR.
        text_len = sum(len(b.get("text", "")) for b in text_blocks)
        if text_len < 10 and ocr_enabled:
            if ocr_handler.is_available():
                logger.info("Page %d: sparse text, running OCR…", page_num + 1)
                ocr_blocks = ocr_handler.ocr_page(page, language=ocr_language)
                text_blocks.extend(ocr_blocks)
            else:
                logger.warning(
                    "Page %d: needs OCR but Tesseract is not installed.", page_num + 1
                )

        # ── Images ──
        images = image_extractor.extract(page, page_num, temp_dir)
        all_images.extend(images)

        # ── Tables ──
        tables = table_extractor.extract(page, page_num)

        # ── Hyperlinks ──
        links = text_extractor.extract_links(page, page_num)
        all_links.extend(links)

        # ── Filter out text that falls inside table regions ──
        text_blocks = _filter_text_in_tables(text_blocks, tables)

        all_text_blocks.extend(text_blocks)
        all_tables.extend(tables)

        progress.update()

    progress.close()
    doc.close()

    # ── Normalize extracted data ─────────────────────────────────────
    all_text_blocks = _normalize_blocks(all_text_blocks)
    all_images = _normalize_blocks(all_images)
    all_tables = _normalize_blocks(all_tables)

    logger.info(
        "Extraction complete — %d text spans, %d images, %d tables, %d links",
        len(all_text_blocks),
        len(all_images),
        len(all_tables),
        len(all_links),
    )

    # ── Compute content-based margins ────────────────────────────────
    _update_metadata_margins(metadata, all_text_blocks, all_images, all_tables)

    # ── Stage 3: Semantic analysis ───────────────────────────────────
    _progress("analyzing", page_count, page_count, "Semantic analysis…")
    logger.info("Running semantic analysis…")
    analyzer = SemanticAnalyzer(config)
    extracted_data: dict[str, Any] = {
        "text_blocks": all_text_blocks,
        "images": all_images,
        "tables": all_tables,
        "links": all_links,
        "metadata": metadata,
    }
    structure_map = analyzer.analyze(extracted_data)
    logger.info("Structure map: %d elements", len(structure_map))

    # ── Stage 4: Build Word document ─────────────────────────────────
    formatting_overrides: dict[str, Any] = {}

    # Strategy B: AI-guided pre-build analysis
    if ai_compare and ai_strategy.upper() == "B":
        ai_config = config.get("ai_comparison", {})
        api_key = ai_config.get("api_key") or os.environ.get("GEMINI_API_KEY", "")
        model = ai_config.get("model", "gemini-2.0-flash")

        if api_key:
            _progress("ai_layout", page_count, page_count,
                      "AI layout analysis (Strategy B)…")
            logger.info("Running AI layout analysis (Strategy B)…")
            try:
                layout_analyzer = AILayoutAnalyzer(
                    api_key=api_key, model=model,
                )
                formatting_overrides = layout_analyzer.analyze_layout(
                    input_path, structure_map,
                    progress_callback=lambda cur, total, msg: _progress(
                        "ai_layout", cur, total, msg
                    ),
                )
                n_text = len(formatting_overrides.get("text_overrides", {}))
                n_table = len(formatting_overrides.get("table_overrides", []))
                logger.info(
                    "AI layout analysis complete — %d text overrides, "
                    "%d table overrides",
                    n_text, n_table,
                )
                _print(f"\n  AI layout: {n_text} text overrides, "
                       f"{n_table} table overrides")
            except Exception as exc:
                logger.warning("AI layout analysis failed: %s — "
                               "proceeding without overrides.", exc)
        else:
            logger.warning("Strategy B selected but no API key — "
                           "building without AI overrides.")

    _progress("building", page_count, page_count, "Building Word document…")
    logger.info("Building Word document…")
    builder = DocxBuilder(config)
    saved_path = builder.build(structure_map, metadata, output_path,
                               formatting_overrides=formatting_overrides or None)
    logger.info("Word document saved → '%s'", saved_path)

    # ── Stage 5: Validation (optional) ───────────────────────────────
    _progress("validating", page_count, page_count, "Validating output…")
    if validate:
        logger.info("Validating output…")
        validator = OutputValidator()
        report = validator.quality_score(saved_path, info)
        score = report.get("quality_score", "?")
        level = report.get("quality_level", "?")
        _print(f"\n  Quality: {level} ({score}/100)")
        metrics = report.get("metrics", {})
        if metrics:
            logger.info(
                "Metrics: %d paragraphs, %d tables, %d images, %d headings",
                metrics.get("paragraphs", 0),
                metrics.get("tables", 0),
                metrics.get("images", 0),
                metrics.get("headings", 0),
            )
        if not report["valid"]:
            logger.warning("Validation issues:")
            for issue in report.get("issues", []):
                logger.warning("  • %s", issue)
        for warning in report.get("warnings", []):
            logger.warning("  ⚠ %s", warning)

    # ── Stage 6-9: Visual diff + AI comparison (optional) ─────────────
    if visual_validate or ai_compare:
        _progress("visual_diff", page_count, page_count, "Running visual diff…")
        vd_config = config.get("visual_diff", {})
        dpi = vd_config.get("dpi", 150)
        lo_path = vd_config.get("libreoffice_path", "auto")

        vdiff = VisualDiff(dpi=dpi, libreoffice_path=lo_path)

        # Output dir for diff images alongside the DOCX.
        diff_dir = os.path.splitext(saved_path)[0] + "_visual_diff"
        vd_result = vdiff.compare(input_path, saved_path, diff_dir)

        overall_ssim = vd_result.get("overall_score", 0)
        vd_level = vd_result.get("quality_level", "red")
        page_scores = vd_result.get("page_scores", [])
        pdf_pages = vd_result.get("pdf_page_count", 0)
        docx_pages = vd_result.get("docx_page_count", 0)

        _print(f"\n  Visual SSIM: {overall_ssim:.1%} ({vd_level})")
        if pdf_pages != docx_pages:
            _print(f"  Page count: PDF={pdf_pages}, DOCX={docx_pages}"
                  f" (overflow: {docx_pages - pdf_pages} extra pages)")
        for i, s in enumerate(page_scores):
            tag = "[OK]" if s >= 0.95 else "[!!]" if s < 0.85 else "[--]"
            _print(f"    Page {i + 1}: {s:.1%} {tag}")

        # AI comparison + correction loop (Strategy A only).
        # Strategy B already applied overrides pre-build; just report SSIM.
        if ai_compare and overall_ssim < 0.95 and ai_strategy.upper() == "A":
            ai_config = config.get("ai_comparison", {})
            api_key = ai_config.get("api_key") or os.environ.get("GEMINI_API_KEY", "")
            model = ai_config.get("model", "gemini-2.0-flash")
            max_rounds = ai_config.get("max_rounds", 3)

            if api_key:
                comparator = AIComparator(
                    api_key=api_key,
                    model=model,
                )

                for round_num in range(1, max_rounds + 1):
                    _progress(
                        "ai_compare", round_num, max_rounds,
                        f"AI comparison round {round_num}/{max_rounds}…",
                    )
                    _print(f"\n  AI comparison round {round_num}/{max_rounds}…")

                    all_diffs = comparator.compare_pages(
                        vd_result["pdf_images"],
                        vd_result["docx_images"],
                    )

                    # Flatten diffs.
                    flat_diffs = [d for page_diffs in all_diffs for d in page_diffs]

                    if not flat_diffs:
                        _print("    No differences detected by AI.")
                        break

                    _print(f"    {len(flat_diffs)} difference(s) found.")
                    for d in flat_diffs[:5]:
                        _print(f"      [{d.get('severity', '?')}] "
                               f"{d.get('area', '?')}: {d.get('issue', '?')}")
                    if len(flat_diffs) > 5:
                        _print(f"      … and {len(flat_diffs) - 5} more")

                    # Apply corrections.
                    _progress(
                        "auto_fix", round_num, max_rounds,
                        f"Applying fixes (round {round_num})…",
                    )
                    engine = CorrectionEngine(saved_path)
                    fixes = engine.apply_fixes(flat_diffs)
                    _print(f"    Applied {fixes} fix(es).")

                    if fixes == 0:
                        _print("    No actionable fixes — stopping loop.")
                        break

                    # Re-render and re-score.
                    vd_result = vdiff.compare(input_path, saved_path, diff_dir)
                    overall_ssim = vd_result.get("overall_score", 0)
                    vd_level = vd_result.get("quality_level", "red")
                    _print(f"    Updated SSIM: {overall_ssim:.1%} ({vd_level})")

                    if overall_ssim >= 0.95:
                        _print("    Target SSIM reached!")
                        break
            else:
                _print("  [SKIP] No GEMINI_API_KEY set — skipping AI comparison.")

    # ── Clean up temp images ─────────────────────────────────────────
    try:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    _progress("complete", page_count, page_count, "Conversion complete")
    _print(f"\n[OK] Conversion complete: {saved_path}")
    return saved_path


# ------------------------------------------------------------------ #
#  Batch conversion                                                   #
# ------------------------------------------------------------------ #

def convert_batch(
    input_dir: str,
    output_dir: str,
    config: dict[str, Any],
    password: str | None = None,
    validate: bool = False,
) -> None:
    """Convert all PDF files in *input_dir* to ``.docx``."""
    pattern = os.path.join(input_dir, "*.pdf")
    pdf_files = sorted(glob.glob(pattern))
    if not pdf_files:
        logger.error("No PDF files found in '%s'.", input_dir)
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    total = len(pdf_files)
    print(f"Batch converting {total} PDF(s)…")

    for i, pdf_path in enumerate(pdf_files, 1):
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_path = os.path.join(output_dir, f"{base}.docx")
        print(f"\n[{i}/{total}] {os.path.basename(pdf_path)}")
        try:
            convert_pdf(pdf_path, out_path, config, password, validate)
        except SystemExit:
            logger.warning("Skipping '%s' due to error.", pdf_path)
        except Exception:
            logger.exception("Unexpected error converting '%s'.", pdf_path)

    print(f"\nBatch complete — {total} file(s) processed → {output_dir}")


# ------------------------------------------------------------------ #
#  CLI entry point                                                    #
# ------------------------------------------------------------------ #

def main() -> None:
    """Parse arguments and run the conversion."""
    parser = argparse.ArgumentParser(
        prog="pdf2docx",
        description="Convert PDF documents to high-fidelity Word (.docx) files.",
    )

    # Positional arguments (optional when --batch is used).
    parser.add_argument("input", nargs="?", help="Input PDF file path.")
    parser.add_argument("output", nargs="?", help="Output .docx file path.")

    # Optional flags.
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Enable OCR for scanned / image-based pages (requires Tesseract).",
    )
    parser.add_argument(
        "--password",
        type=str,
        default=None,
        help="Password for encrypted PDFs.",
    )
    parser.add_argument(
        "--config",
        type=str,
        default=None,
        help="Path to a custom configuration JSON file.",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose (DEBUG) logging.",
    )
    parser.add_argument(
        "--validate",
        action="store_true",
        help="Run validation checks on the output document.",
    )
    parser.add_argument(
        "--batch",
        type=str,
        default=None,
        help="Directory containing PDF files for batch conversion.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Output directory for batch conversion results.",
    )
    parser.add_argument(
        "--skip-watermarks",
        action="store_true",
        help="Attempt to skip watermark / background images.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output files without prompting.",
    )
    parser.add_argument(
        "--visual-validate",
        action="store_true",
        help="Render PDF and DOCX, then compute SSIM visual similarity scores.",
    )
    parser.add_argument(
        "--ai-compare",
        action="store_true",
        help="Use AI (Gemini) to detect and fix visual differences (requires GEMINI_API_KEY).",
    )
    parser.add_argument(
        "--ai-strategy",
        type=str,
        choices=["A", "B"],
        default="A",
        help="AI strategy: A = post-build correction loop, B = pre-build AI-guided layout (default: A).",
    )

    args = parser.parse_args()

    # ── Logging ──────────────────────────────────────────────────────
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
        stream=sys.stderr,
    )

    # ── Config ───────────────────────────────────────────────────────
    config = _load_config(args.config)
    if args.ocr:
        config["ocr_enabled"] = True
    if args.skip_watermarks:
        config["skip_watermarks"] = True
    if args.verbose:
        config["verbose"] = True

    # ── Batch mode ───────────────────────────────────────────────────
    if args.batch:
        out_dir = args.output_dir or os.path.join(args.batch, "docx_output")
        convert_batch(args.batch, out_dir, config, args.password, args.validate)
        return

    # ── Single-file mode ─────────────────────────────────────────────
    if not args.input:
        parser.error("Please provide an input PDF file (or use --batch for batch mode).")

    input_path = args.input
    output_path = args.output
    if not output_path:
        # Default: same name with .docx extension, in the same directory.
        base = os.path.splitext(input_path)[0]
        output_path = base + ".docx"

    # Check overwrite.
    if os.path.exists(output_path) and not args.force:
        resp = input(f"Output file '{output_path}' exists. Overwrite? [y/N] ").strip().lower()
        if resp not in ("y", "yes"):
            print("Aborted.")
            sys.exit(0)

    convert_pdf(input_path, output_path, config, args.password, args.validate,
                args.visual_validate, args.ai_compare,
                ai_strategy=args.ai_strategy)


if __name__ == "__main__":
    main()
