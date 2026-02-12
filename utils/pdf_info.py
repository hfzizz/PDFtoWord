"""PDF information analysis module.

Provides quick analysis of PDF characteristics without full extraction.
"""

import logging
import os
from typing import Any, Dict

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

TEXT_THRESHOLD = 10  # Minimum characters to consider a page as having text
TABLE_CHECK_PAGES = 3  # Number of pages to check for tables


class PDFInfo:
    """Analyzes PDF files and returns structural information."""

    def analyze(self, pdf_path: str) -> Dict[str, Any]:
        """Analyze a PDF file and return its characteristics.

        Args:
            pdf_path: Path to the PDF file.

        Returns:
            Dictionary containing PDF characteristics:
                - path: Absolute file path.
                - file_size_mb: File size in megabytes.
                - is_text_based: Whether >50% of pages have extractable text.
                - is_scanned: Opposite of is_text_based.
                - is_hybrid: Some pages have text, some don't.
                - is_encrypted: Whether the PDF is encrypted.
                - page_count: Total number of pages.
                - needs_ocr: Whether OCR is needed (scanned or hybrid).
                - has_images: Whether the PDF contains images.
                - has_tables: Whether tables were detected on the first few pages.
        """
        pdf_path = os.path.abspath(pdf_path)
        file_size_mb = os.path.getsize(pdf_path) / (1024 * 1024)

        result: Dict[str, Any] = {
            "path": pdf_path,
            "file_size_mb": round(file_size_mb, 2),
            "is_text_based": False,
            "is_scanned": False,
            "is_hybrid": False,
            "is_encrypted": False,
            "page_count": 0,
            "needs_ocr": False,
            "has_images": False,
            "has_tables": False,
        }

        try:
            doc = fitz.open(pdf_path)
        except Exception as exc:
            logger.error("Failed to open PDF '%s': %s", pdf_path, exc)
            raise

        try:
            if doc.is_encrypted:
                result["is_encrypted"] = True
                logger.warning("PDF is encrypted: %s", pdf_path)
                result["page_count"] = doc.page_count
                return result

            page_count = doc.page_count
            result["page_count"] = page_count

            if page_count == 0:
                return result

            # --- Text detection ---
            pages_with_text = 0
            pages_without_text = 0

            for page in doc:
                text = page.get_text("text").strip()
                if len(text) > TEXT_THRESHOLD:
                    pages_with_text += 1
                else:
                    pages_without_text += 1

                # --- Image detection (stop once found) ---
                if not result["has_images"]:
                    if page.get_images(full=False):
                        result["has_images"] = True

            text_ratio = pages_with_text / page_count
            result["is_text_based"] = text_ratio > 0.5
            result["is_scanned"] = not result["is_text_based"]
            result["is_hybrid"] = 0 < pages_with_text < page_count
            result["needs_ocr"] = result["is_scanned"] or result["is_hybrid"]

            # --- Table detection on first few pages ---
            result["has_tables"] = self._check_tables(doc)

        finally:
            doc.close()

        logger.info(
            "Analyzed PDF '%s': %d pages, text_based=%s, needs_ocr=%s",
            pdf_path,
            result["page_count"],
            result["is_text_based"],
            result["needs_ocr"],
        )
        return result

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _check_tables(doc: fitz.Document) -> bool:
        """Check the first few pages for tables using find_tables.

        Args:
            doc: An open PyMuPDF Document.

        Returns:
            True if tables were found on any of the checked pages.
        """
        pages_to_check = min(TABLE_CHECK_PAGES, doc.page_count)
        for idx in range(pages_to_check):
            try:
                page = doc[idx]
                tables = page.find_tables()
                if tables and len(tables.tables) > 0:
                    return True
            except Exception as exc:  # pragma: no cover
                logger.debug(
                    "Table detection failed on page %d: %s", idx, exc
                )
        return False
