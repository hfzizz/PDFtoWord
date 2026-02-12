"""Text extraction from PDF pages at the span level using PyMuPDF."""

import logging
from typing import Any

import fitz

logger = logging.getLogger(__name__)


class TextExtractor:
    """Extracts text blocks from a PDF page at the span level.

    Uses ``page.get_text("dict")`` to walk blocks → lines → spans and
    returns a flat list of span-level dicts with positional and style
    metadata.
    """

    # Font-flag bit masks (per the PDF spec / PyMuPDF constants)
    _SUPERSCRIPT = 1 << 0
    _ITALIC = 1 << 1
    _SERIF = 1 << 2
    _MONOSPACE = 1 << 3
    _BOLD = 1 << 4

    @staticmethod
    def _int_color_to_rgb(color_int: int) -> tuple[int, int, int]:
        """Convert a PyMuPDF integer colour value to an (r, g, b) tuple.

        PyMuPDF encodes span colours as a single integer where the bytes
        represent ``0x00RRGGBB``.
        """
        r = (color_int >> 16) & 0xFF
        g = (color_int >> 8) & 0xFF
        b = color_int & 0xFF
        return (r, g, b)

    @staticmethod
    def _derotation_matrix(page: fitz.Page) -> fitz.Matrix | None:
        """Return a matrix that maps rotated coordinates back to the
        unrotated page space, or *None* if the page is not rotated.
        """
        rotation = page.rotation
        if rotation == 0:
            return None
        return page.derotation_matrix

    def extract(self, page: fitz.Page, page_num: int = 0) -> list[dict[str, Any]]:
        """Extract span-level text blocks from *page*.

        Parameters
        ----------
        page:
            A ``fitz.Page`` object.
        page_num:
            The page number (0-based) used to tag each result dict.

        Returns
        -------
        list[dict[str, Any]]
            Each dict contains the keys ``text``, ``bbox``, ``font``,
            ``size``, ``color``, ``bold``, ``italic``, ``flags`` and
            ``page_num``.  Returns an empty list when extraction fails.
        """
        try:
            text_dict: dict[str, Any] = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            derotation = self._derotation_matrix(page)
            results: list[dict[str, Any]] = []

            for block in text_dict.get("blocks", []):
                # Skip image blocks (type == 1)
                if block.get("type") != 0:
                    continue

                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        size: float = span.get("size", 0.0)
                        if size < 0.5:
                            continue

                        text: str = span.get("text", "")
                        if not text:
                            continue

                        bbox = span.get("bbox", (0, 0, 0, 0))
                        # Zero-width check
                        if abs(bbox[2] - bbox[0]) < 0.01 and abs(bbox[3] - bbox[1]) < 0.01:
                            continue

                        # Apply derotation when the page is rotated
                        if derotation is not None:
                            rect = fitz.Rect(bbox) * derotation
                            bbox = tuple(rect)

                        flags: int = span.get("flags", 0)
                        color_int: int = span.get("color", 0)

                        results.append(
                            {
                                "text": text,
                                "bbox": bbox,
                                "font": span.get("font", ""),
                                "size": float(size),
                                "color": self._int_color_to_rgb(color_int),
                                "bold": bool(flags & self._BOLD),
                                "italic": bool(flags & self._ITALIC),
                                "flags": flags,
                                "page_num": page_num,
                                "underline": False,
                                "strikethrough": False,
                            }
                        )

            # Detect underline/strikethrough by checking for horizontal
            # line drawings that overlap text spans.
            self._detect_underline_strikethrough(page, results)

            return results

        except Exception:
            logger.exception("Failed to extract text from page %s", page_num)
            return []

    @staticmethod
    def _detect_underline_strikethrough(
        page: fitz.Page, spans: list[dict[str, Any]]
    ) -> None:
        """Check for horizontal line drawings that indicate underline/strikethrough.

        Modifies *spans* in place, setting ``underline`` and/or
        ``strikethrough`` to ``True`` when detected.
        """
        if not spans:
            return
        try:
            drawings = page.get_drawings()
        except Exception:
            return
        # Collect horizontal lines.
        h_lines: list[tuple[float, float, float, float]] = []  # x0, x1, y, width
        for d in drawings:
            if d.get("fill") is not None:
                continue  # filled = background, not line
            for item in d.get("items", []):
                if item[0] != "l":
                    continue
                p1, p2 = fitz.Point(item[1]), fitz.Point(item[2])
                if abs(p1.y - p2.y) > 2:
                    continue  # not horizontal
                line_len = abs(p2.x - p1.x)
                if line_len < 5:
                    continue
                h_lines.append((min(p1.x, p2.x), max(p1.x, p2.x),
                                (p1.y + p2.y) / 2, line_len))

        if not h_lines:
            return

        for span in spans:
            bbox = span.get("bbox", (0, 0, 0, 0))
            sx0, sy0, sx1, sy1 = bbox
            span_mid_y = (sy0 + sy1) / 2
            span_height = sy1 - sy0
            if span_height <= 0:
                continue
            for lx0, lx1, ly, ll in h_lines:
                # Line must roughly overlap the span horizontally.
                if lx1 < sx0 + 2 or lx0 > sx1 - 2:
                    continue
                if ly < sy0 - 2 or ly > sy1 + 2:
                    continue
                # Underline: line near bottom of text.
                if ly > span_mid_y:
                    span["underline"] = True
                else:
                    span["strikethrough"] = True

    def extract_links(self, page: fitz.Page, page_num: int = 0) -> list[dict[str, Any]]:
        """Extract hyperlinks from *page*.

        Parameters
        ----------
        page:
            A ``fitz.Page`` object.
        page_num:
            The page number (0-based) used to tag each result dict.

        Returns
        -------
        list[dict[str, Any]]
            Each dict contains ``uri``, ``bbox`` and ``page_num``.
            External links have a normal URL as ``uri``; internal
            (goto) links use ``"#page_N"`` format.  Returns an empty
            list when extraction fails.
        """
        try:
            links: list[dict[str, Any]] = []
            for link in page.get_links():
                kind = link.get("kind")
                if kind == fitz.LINK_URI:
                    uri = link.get("uri", "")
                    if not uri:
                        continue
                    rect = link.get("from", fitz.Rect())
                    bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                    links.append({"uri": uri, "bbox": bbox, "page_num": page_num})
                elif kind == fitz.LINK_GOTO:
                    target_page = link.get("page", 0)
                    uri = f"#page_{target_page}"
                    rect = link.get("from", fitz.Rect())
                    bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                    links.append({"uri": uri, "bbox": bbox, "page_num": page_num})
                # Skip other link types (LINK_NAMED, LINK_LAUNCH, etc.)
            return links
        except Exception:
            logger.exception("Failed to extract links from page %s", page_num)
            return []
