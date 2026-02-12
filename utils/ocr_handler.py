"""OCR handler module.

Provides OCR capabilities for scanned PDF pages using Tesseract.
"""

import io
import logging
from collections import defaultdict
from typing import Any, Dict, List, Optional

import fitz  # PyMuPDF

logger = logging.getLogger(__name__)

OCR_DPI = 300
DEFAULT_FONT = "Arial"
DEFAULT_FONT_SIZE = 12
DEFAULT_COLOR = (0, 0, 0)


class OCRHandler:
    """Handles OCR processing of PDF pages via Tesseract."""

    def is_available(self) -> bool:
        """Check whether Tesseract OCR is installed and reachable.

        Returns:
            True if pytesseract can communicate with an installed
            Tesseract binary, False otherwise.
        """
        try:
            import pytesseract  # noqa: F811

            pytesseract.get_tesseract_version()
            return True
        except ImportError:
            logger.warning("pytesseract is not installed.")
            return False
        except Exception as exc:
            logger.warning("Tesseract not available: %s", exc)
            return False

    def ocr_page(
        self,
        page: fitz.Page,
        language: str = "eng",
    ) -> List[Dict[str, Any]]:
        """Run OCR on a single PDF page and return text blocks.

        The page is rendered to a high-resolution pixmap, converted to a
        PIL Image, and processed through Tesseract.  Words are grouped
        into lines following Tesseract's block/line numbering.

        Args:
            page: A PyMuPDF ``fitz.Page`` object.
            language: Tesseract language code (default ``"eng"``).

        Returns:
            A list of text-block dicts, each containing:
                - ``text``: The recognized line text.
                - ``bbox``: Tuple ``(x0, y0, x1, y1)`` in page coordinates.
                - ``font``: Default font name (``"Arial"``).
                - ``size``: Default font size (``12``).
                - ``color``: Default color ``(0, 0, 0)``.
                - ``ocr``: ``True`` flag indicating OCR origin.
        """
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            logger.warning(
                "pytesseract or Pillow is not installed; "
                "cannot perform OCR. Returning empty list."
            )
            return []

        try:
            # Render page at high DPI
            pixmap = page.get_pixmap(dpi=OCR_DPI)
            img = Image.open(io.BytesIO(pixmap.tobytes("png")))

            # Run Tesseract
            data: Dict[str, List[Any]] = pytesseract.image_to_data(
                img,
                lang=language,
                output_type=pytesseract.Output.DICT,
            )

            # Scale factors: pixmap pixels → PDF points
            page_rect = page.rect
            scale_x = page_rect.width / pixmap.width
            scale_y = page_rect.height / pixmap.height

            # Group words into lines keyed by (block_num, line_num)
            lines: Dict[tuple, List[Dict[str, Any]]] = defaultdict(list)
            n_words = len(data["text"])

            for i in range(n_words):
                word = data["text"][i].strip()
                if not word:
                    continue

                conf = int(data["conf"][i])
                if conf < 0:
                    continue

                key = (data["block_num"][i], data["line_num"][i])
                lines[key].append(
                    {
                        "word": word,
                        "left": data["left"][i],
                        "top": data["top"][i],
                        "width": data["width"][i],
                        "height": data["height"][i],
                    }
                )

            # Build text blocks from grouped lines
            blocks: List[Dict[str, Any]] = []
            for _key, words in lines.items():
                if not words:
                    continue

                text = " ".join(w["word"] for w in words)
                x0 = min(w["left"] for w in words)
                y0 = min(w["top"] for w in words)
                x1 = max(w["left"] + w["width"] for w in words)
                y1 = max(w["top"] + w["height"] for w in words)

                # Convert pixel coordinates to page points
                bbox = (
                    round(x0 * scale_x, 2),
                    round(y0 * scale_y, 2),
                    round(x1 * scale_x, 2),
                    round(y1 * scale_y, 2),
                )

                blocks.append(
                    {
                        "text": text,
                        "bbox": bbox,
                        "font": DEFAULT_FONT,
                        "size": DEFAULT_FONT_SIZE,
                        "color": DEFAULT_COLOR,
                        "ocr": True,
                        "page_num": page.number,
                        "bold": False,
                        "italic": False,
                        "flags": 0,
                    }
                )

            logger.info(
                "OCR produced %d text blocks for page %d",
                len(blocks),
                page.number,
            )
            return blocks

        except Exception as exc:
            logger.error("OCR failed on page %d: %s", page.number, exc)
            return []
