"""Semantic analyzer module — main entry point for content analysis.

Combines layout, font, and content analysis to produce an ordered list of
semantic document elements (headings, paragraphs, lists, tables, images,
page breaks) that the Word builder can consume directly.
"""

import logging
import re
from typing import Any

from .font_analyzer import FontAnalyzer
from .layout_analyzer import LayoutAnalyzer

logger = logging.getLogger(__name__)

# Bullet characters that signal an unordered list item.
_BULLET_CHARS = set("•●○■□-*▪")

# Regex for an ordered-list prefix (digits, single letter, or roman numerals
# followed by '.' or ')').
_ORDERED_LIST_RE = re.compile(
    r"^(\d+|[a-zA-Z]|[ivxlcdmIVXLCDM]+)[.)]\s"
)

# Default left-margin tolerance in points for alignment detection.
_ALIGNMENT_TOLERANCE = 10.0


class SemanticAnalyzer:
    """Orchestrates layout, font, and content analysis into a semantic structure.

    The analyzer receives raw extracted data (text blocks, images, tables,
    metadata) and returns an ordered list of element dicts ready for the
    Word document builder.

    Args:
        config: Configuration dict (typically loaded from ``settings.json``).
    """

    def __init__(self, config: dict[str, Any] | None = None) -> None:
        self.config: dict[str, Any] = config or {}
        self._layout_analyzer = LayoutAnalyzer()
        self._font_analyzer = FontAnalyzer()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def analyze(self, extracted_data: dict[str, Any]) -> list[dict[str, Any]]:
        """Produce an ordered list of semantic document elements.

        Args:
            extracted_data: Dict containing:
                - ``text_blocks``: list of text block dicts (all pages,
                  sorted by page → y → x).
                - ``images``: list of image dicts.
                - ``tables``: list of table dicts.
                - ``metadata``: dict with at least ``pages`` (list of
                  ``{"width": float, "height": float, …}``).

        Returns:
            Ordered list of element dicts.  See module docstring for the
            element schemas.
        """
        text_blocks: list[dict[str, Any]] = extracted_data.get("text_blocks", [])
        images: list[dict[str, Any]] = extracted_data.get("images", [])
        tables: list[dict[str, Any]] = extracted_data.get("tables", [])
        links: list[dict[str, Any]] = extracted_data.get("links", [])
        metadata: dict[str, Any] = extracted_data.get("metadata", {})

        pages_meta: list[dict[str, Any]] = metadata.get("pages", [])

        # --- Font analysis (global across all pages) -----------------------
        font_result = self._font_analyzer.analyze(text_blocks)
        body_font = font_result["body_font"]
        heading_fonts = font_result["heading_fonts"]
        font_map = font_result["font_map"]
        body_size: float = body_font["size"]

        logger.debug("Font analysis done. Body: %s", body_font)

        # Build a fast lookup: (font, size) → heading level (or None).
        heading_lookup: dict[tuple[str, float], int] = {
            (h["name"], h["size"]): h["level"] for h in heading_fonts
        }

        # --- Detect and filter headers / footers --------------------------
        page_count = len(pages_meta) if pages_meta else max(
            (b.get("page_num", 0) for b in text_blocks), default=0
        ) + 1
        header_texts, footer_texts, text_blocks = self._detect_headers_footers(
            text_blocks, pages_meta, page_count,
        )
        if header_texts:
            logger.debug("Detected header text(s): %s", header_texts)
        if footer_texts:
            logger.debug("Detected footer text(s): %s", footer_texts)

        # --- Group text blocks by page ------------------------------------
        blocks_by_page: dict[int, list[dict[str, Any]]] = {}
        for block in text_blocks:
            pg = block.get("page_num", 0)
            blocks_by_page.setdefault(pg, []).append(block)

        # --- Collect all page numbers we know about ------------------------
        all_page_nums: set[int] = set(blocks_by_page.keys())
        for img in images:
            all_page_nums.add(img.get("page_num", 0))
        for tbl in tables:
            all_page_nums.add(tbl.get("page_num", 0))

        # --- Process each page ---------------------------------------------
        elements: list[dict[str, Any]] = []
        prev_page: int | None = None

        for page_num in sorted(all_page_nums):
            # Insert page break between pages.
            if prev_page is not None and page_num != prev_page:
                # Determine orientation for the upcoming page.
                _pg_meta = pages_meta[page_num] if page_num < len(pages_meta) else {}
                _orientation = "landscape" if _pg_meta.get("is_landscape") else "portrait"
                _page_break: dict[str, Any] = {
                    "type": "page_break",
                    "page_num": page_num,
                    "orientation": _orientation,
                }
                # Attach per-page margins so the builder can create a
                # properly sized section for each page.
                _pg_margins = _pg_meta.get("margins")
                if _pg_margins:
                    _page_break["margins"] = _pg_margins
                elements.append(_page_break)
            prev_page = page_num

            # Page dimensions for layout / alignment.
            page_width, page_height = self._page_dimensions(pages_meta, page_num)

            # Layout analysis for this page.
            page_blocks = blocks_by_page.get(page_num, [])
            layout = self._layout_analyzer.analyze(page_blocks, page_width, page_height)

            num_columns = layout.get("num_columns", 1)

            # Validate multi-column: each column must have enough blocks
            # to be considered a real column (avoid false positives from
            # centered headings or stray text at different x-positions).
            if num_columns > 1:
                cols_blocks = layout.get("text_blocks_by_column", {})
                min_blocks_per_col = 3
                real_cols = sum(
                    1 for blks in cols_blocks.values()
                    if len(blks) >= min_blocks_per_col
                )
                if real_cols < 2:
                    num_columns = 1  # Fall back to single-column.

            # -- Convert text blocks → paragraphs / headings / list items ---
            if num_columns > 1:
                # Multi-column: process each column left-to-right so that
                # left-column text precedes right-column text in the output.
                logger.debug(
                    "Page %d: multi-column layout detected (%d columns)",
                    page_num,
                    num_columns,
                )
                text_elements: list[dict[str, Any]] = []
                cols_blocks = layout.get("text_blocks_by_column", {})
                for col_idx in sorted(cols_blocks.keys()):
                    col_blocks = cols_blocks[col_idx]
                    col_body_x0 = self._estimate_body_x0(col_blocks)
                    col_elements = self._process_text_blocks(
                        col_blocks,
                        page_num,
                        body_size,
                        col_body_x0,
                        heading_lookup,
                        font_map,
                        page_width,
                    )
                    text_elements.extend(col_elements)
            else:
                # Single-column: original behaviour.
                body_x0 = self._estimate_body_x0(page_blocks)
                text_elements = self._process_text_blocks(
                    page_blocks,
                    page_num,
                    body_size,
                    body_x0,
                    heading_lookup,
                    font_map,
                    page_width,
                )

            # -- Gather images & tables for this page -----------------------
            page_images = [
                self._make_image_element(img)
                for img in images
                if img.get("page_num", 0) == page_num
            ]
            page_tables = [
                self._make_table_element(tbl)
                for tbl in tables
                if tbl.get("page_num", 0) == page_num
            ]

            # -- Merge & sort by y-position ---------------------------------
            page_elements = text_elements + page_images + page_tables
            page_elements.sort(key=lambda e: (e.get("_y0", 0), e.get("_x0", 0)))

            # Strip internal sort keys before appending.
            for elem in page_elements:
                elem.pop("_y0", None)
                elem.pop("_x0", None)

            elements.extend(page_elements)

        # --- Attach hyperlinks to paragraph/heading elements ---------------
        self._attach_links(elements, links)

        # --- Insert header / footer elements at the start ------------------
        if footer_texts:
            elements.insert(0, {"type": "footer", "text": "\n".join(footer_texts)})
        if header_texts:
            elements.insert(0, {"type": "header", "text": "\n".join(header_texts)})

        logger.info("Semantic analysis complete: %d elements produced.", len(elements))
        return elements

    # ------------------------------------------------------------------
    # Text processing
    # ------------------------------------------------------------------

    def _process_text_blocks(
        self,
        blocks: list[dict[str, Any]],
        page_num: int,
        body_size: float,
        body_x0: float,
        heading_lookup: dict[tuple[str, float], int],
        font_map: dict[str, str],
        page_width: float,
    ) -> list[dict[str, Any]]:
        """Convert raw text blocks into semantic elements.

        Consecutive text spans with matching font/size/style on adjacent
        lines are grouped into single paragraphs.
        """
        if not blocks:
            return []

        sorted_blocks = sorted(blocks, key=lambda b: (b.get("y0", 0), b.get("x0", 0)))

        elements: list[dict[str, Any]] = []
        group: list[dict[str, Any]] = [sorted_blocks[0]]
        prev_y1: float = 0.0

        for block in sorted_blocks[1:]:
            if self._should_merge(group[-1], block, body_size, page_width):
                group.append(block)
            else:
                elements.append(
                    self._finalize_group(
                        group, page_num, body_size, body_x0,
                        heading_lookup, font_map, page_width,
                        prev_y1,
                    )
                )
                prev_y1 = group[-1].get("y1", group[-1].get("y0", 0))
                group = [block]

        # Finalize last group.
        elements.append(
            self._finalize_group(
                group, page_num, body_size, body_x0,
                heading_lookup, font_map, page_width,
                prev_y1,
            )
        )
        return elements

    @staticmethod
    def _is_list_start(text: str) -> bool:
        """Return ``True`` if *text* looks like a list item prefix."""
        stripped = text.lstrip()
        if not stripped:
            return False
        if stripped[0] in _BULLET_CHARS:
            return True
        if _ORDERED_LIST_RE.match(stripped):
            return True
        return False

    @staticmethod
    def _should_merge(
        prev: dict[str, Any],
        curr: dict[str, Any],
        body_size: float,
        page_width: float = 612.0,
    ) -> bool:
        """Decide whether two consecutive text blocks should be merged.

        Blocks are merged when they share the same font, size, and style,
        the vertical gap between them is less than 1.5 × the font size,
        **and** the previous line appears to be a wrapped continuation
        (reaching close to the right margin) rather than a deliberate
        short line (address block, signature block, etc.).
        """
        # Never merge list items — each should be a separate element.
        curr_text = curr.get("text", "")
        if SemanticAnalyzer._is_list_start(curr_text):
            return False

        if prev.get("font") != curr.get("font"):
            return False
        if prev.get("size") != curr.get("size"):
            return False
        if prev.get("bold") != curr.get("bold"):
            return False
        if prev.get("italic") != curr.get("italic"):
            return False

        size = float(curr.get("size", body_size) or body_size)
        y_gap = curr.get("y0", 0) - prev.get("y1", prev.get("y0", 0))
        if y_gap >= 1.5 * size:
            return False

        # Short-line check: if the previous line does not reach close
        # to the right margin, it ended intentionally (line break),
        # not because the text wrapped.  In that case do NOT merge.
        prev_x0 = prev.get("x0", 0)
        prev_x1 = prev.get("x1", 0)
        line_width = prev_x1 - prev_x0
        # Estimate the content area (page width minus typical margins).
        content_width = max(page_width - 144, page_width * 0.6)
        if content_width > 0 and line_width < content_width * 0.6:
            return False

        return True

    def _finalize_group(
        self,
        group: list[dict[str, Any]],
        page_num: int,
        body_size: float,
        body_x0: float,
        heading_lookup: dict[tuple[str, float], int],
        font_map: dict[str, str],
        page_width: float,
        prev_y1: float = 0.0,
    ) -> dict[str, Any]:
        """Convert a group of merged text blocks into a single element dict."""
        # Representative block for formatting (use the first block).
        rep = group[0]
        raw_font = rep.get("font", "")
        clean_font = FontAnalyzer._strip_subset_prefix(raw_font)  # noqa: SLF001
        size = float(rep.get("size", body_size))
        is_bold = bool(rep.get("bold", False))
        is_italic = bool(rep.get("italic", False))
        color = rep.get("color", (0, 0, 0))
        fallback_font = font_map.get(clean_font, self.config.get("fallback_font", "Arial"))

        # Combine text from all blocks.
        texts: list[str] = []
        for i, blk in enumerate(group):
            txt = blk.get("text", "")
            if i > 0:
                y_gap = blk.get("y0", 0) - group[i - 1].get(
                    "y1", group[i - 1].get("y0", 0)
                )
                sep = "\n" if y_gap > size else " "
                texts.append(sep)
            texts.append(txt)
        merged_text = "".join(texts).strip()

        # Build per-block runs so the builder can apply mixed formatting.
        runs: list[dict[str, Any]] = []
        for i, blk in enumerate(group):
            txt = blk.get("text", "")
            if i > 0:
                y_gap = blk.get("y0", 0) - group[i - 1].get(
                    "y1", group[i - 1].get("y0", 0)
                )
                if y_gap > size:
                    runs.append({"text": "\n"})
                else:
                    txt = " " + txt
            blk_raw_font = blk.get("font", "")
            blk_clean_font = FontAnalyzer._strip_subset_prefix(blk_raw_font)  # noqa: SLF001
            runs.append({
                "text": txt,
                "font": font_map.get(blk_clean_font, self.config.get("fallback_font", "Arial")),
                "size": float(blk.get("size", body_size)),
                "bold": bool(blk.get("bold", False)),
                "italic": bool(blk.get("italic", False)),
                "color": blk.get("color", (0, 0, 0)),
                "underline": bool(blk.get("underline", False)),
                "strikethrough": bool(blk.get("strikethrough", False)),
                "superscript": bool(blk.get("superscript", False)),
                "highlight_color": blk.get("highlight_color"),
                "x0": blk.get("x0", 0),
                "y0": blk.get("y0", 0),
                "x1": blk.get("x1", 0),
                "y1": blk.get("y1", 0),
            })

        # Positional info (for sorting).
        y0 = group[0].get("y0", 0)
        x0 = group[0].get("x0", 0)

        # Alignment.
        alignment = self._detect_alignment(group, body_x0, page_width)

        # Spacing before: vertical gap between this element and previous.
        spacing_before = min(max(y0 - prev_y1, 0), 72.0)

        # Left indentation relative to the body left margin.
        indent_left = max(x0 - body_x0, 0)
        if indent_left < 5.0:
            indent_left = 0.0

        # --- Line spacing (ratio relative to font size) -------------------
        # Inspired by pdf2docx: compute spacing between consecutive line
        # baselines within the group and express as a ratio (1.0 = single,
        # 1.5 = 1.5 spacing, 2.0 = double).
        line_spacing: float | None = None
        if len(group) >= 2:
            gaps: list[float] = []
            for i in range(1, len(group)):
                g_y0 = group[i].get("y0", 0)
                p_y0 = group[i - 1].get("y0", 0)
                gap = g_y0 - p_y0
                if gap > 0:
                    gaps.append(gap)
            if gaps and size > 0:
                avg_gap = sum(gaps) / len(gaps)
                ratio = round(avg_gap / size, 1)
                # Only store meaningful ratios (0.8 – 3.0).
                if 0.8 <= ratio <= 3.0 and abs(ratio - 1.0) > 0.15:
                    line_spacing = ratio

        # --- First-line indent detection -----------------------------------
        # If the first line starts noticeably to the right of subsequent
        # lines, treat the difference as a first-line indent.
        first_line_indent: float = 0.0
        if len(group) >= 2:
            first_x0 = group[0].get("x0", 0)
            rest_x0s = [b.get("x0", 0) for b in group[1:]]
            avg_rest_x0 = sum(rest_x0s) / len(rest_x0s)
            diff = first_x0 - avg_rest_x0
            if diff > 8.0:  # at least ~8 pt indent
                first_line_indent = diff

        # Check for heading.
        heading_level = heading_lookup.get((clean_font, size))
        if heading_level is not None:
            fmt_heading: dict[str, Any] = {
                "font": fallback_font,
                "size": size,
                "bold": is_bold,
                "italic": is_italic,
                "color": color,
                "alignment": alignment,
                "spacing_before": spacing_before,
                "indent_left": indent_left,
            }
            if line_spacing is not None:
                fmt_heading["line_spacing"] = line_spacing
            return {
                "type": "heading",
                "level": heading_level,
                "text": merged_text,
                "page_num": page_num,
                "formatting": fmt_heading,
                "_y0": y0,
                "_x0": x0,
            }

        # Check for list item.
        list_info = self._detect_list(merged_text, x0, body_x0)
        if list_info is not None:
            return {
                "type": "list_item",
                "text": list_info["text"],
                "page_num": page_num,
                "level": list_info["level"],
                "bullet_type": list_info["bullet_type"],
                "formatting": {
                    "font": fallback_font,
                    "size": size,
                    "bold": is_bold,
                    "italic": is_italic,
                    "color": color,
                    "alignment": alignment,
                    "spacing_before": spacing_before,
                    "indent_left": indent_left,
                    **({
                        "line_spacing": line_spacing,
                    } if line_spacing is not None else {}),
                },
                "_y0": y0,
                "_x0": x0,
            }

        # Default: paragraph.
        fmt_para: dict[str, Any] = {
            "font": fallback_font,
            "size": size,
            "bold": is_bold,
            "italic": is_italic,
            "color": color,
            "alignment": alignment,
            "spacing_before": spacing_before,
            "indent_left": indent_left,
        }
        if line_spacing is not None:
            fmt_para["line_spacing"] = line_spacing
        if first_line_indent > 0:
            fmt_para["first_line_indent"] = first_line_indent
        return {
            "type": "paragraph",
            "text": merged_text,
            "runs": runs,
            "page_num": page_num,
            "formatting": fmt_para,
            "_y0": y0,
            "_x0": x0,
        }

    # ------------------------------------------------------------------
    # Detection helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _detect_list(
        text: str, x0: float, body_x0: float
    ) -> dict[str, Any] | None:
        """Detect whether *text* is a list item and return its metadata.

        Returns ``None`` if the text is not a list item, otherwise a dict
        with ``text`` (stripped of bullet/number prefix), ``level``, and
        ``bullet_type``.
        """
        stripped = text.lstrip()
        if not stripped:
            return None

        indent_pts = max(x0 - body_x0, 0)
        level = max(int(indent_pts / 20), 0)

        # Bullet detection.
        if stripped[0] in _BULLET_CHARS:
            clean_text = stripped[1:].lstrip()
            return {"text": clean_text, "level": level, "bullet_type": "bullet"}

        # Ordered list detection.
        m = _ORDERED_LIST_RE.match(stripped)
        if m:
            clean_text = stripped[m.end():].lstrip()
            return {"text": clean_text, "level": level, "bullet_type": "number"}

        return None

    @staticmethod
    def _detect_headers_footers(
        text_blocks: list[dict[str, Any]],
        pages_meta: list[dict[str, Any]],
        page_count: int,
    ) -> tuple[list[str], list[str], list[dict[str, Any]]]:
        """Detect repeated header/footer text across pages.

        Text that appears in the top ~10 % of the page on >= 50 % of
        pages is treated as a header.  Text in the bottom ~10 % is a
        footer.  Detected blocks are removed from the returned list.

        Returns:
            ``(header_texts, footer_texts, filtered_blocks)``
        """
        if page_count < 3 or not text_blocks:
            return [], [], text_blocks

        threshold = 0.08  # top/bottom 8 % zone
        min_page_ratio = 0.50  # must appear on >= 50 % of pages
        min_page_count = 3  # must appear on at least 3 pages

        # Build a mapping page_num → page_height.
        def _page_height(page_num: int) -> float:
            if pages_meta and 0 <= page_num < len(pages_meta):
                return float(pages_meta[page_num].get("height", 792))
            return 792.0  # US-Letter fallback

        # Classify each block as top-zone, bottom-zone, or body.
        # Key: normalised text → {"zone": "top"|"bottom", "pages": set}
        from collections import defaultdict

        zone_counts: dict[str, dict[str, Any]] = defaultdict(
            lambda: {"top_pages": set(), "bottom_pages": set()}
        )

        for block in text_blocks:
            page_num = block.get("page_num", 0)
            ph = _page_height(page_num)
            y0 = block.get("y0", 0)
            y1 = block.get("y1", y0)

            norm_text = block.get("text", "").strip()
            if not norm_text:
                continue

            if y1 <= ph * threshold:
                zone_counts[norm_text]["top_pages"].add(page_num)
            elif y0 >= ph * (1 - threshold):
                zone_counts[norm_text]["bottom_pages"].add(page_num)

        header_set: set[str] = set()
        footer_set: set[str] = set()

        for text, info in zone_counts.items():
            top_count = len(info["top_pages"])
            bottom_count = len(info["bottom_pages"])
            if top_count >= max(page_count * min_page_ratio, min_page_count):
                header_set.add(text)
            if bottom_count >= max(page_count * min_page_ratio, min_page_count):
                footer_set.add(text)

        if not header_set and not footer_set:
            return [], [], text_blocks

        # Filter blocks and collect ordered unique texts.
        filtered: list[dict[str, Any]] = []
        seen_headers: list[str] = []
        seen_footers: list[str] = []
        header_seen_set: set[str] = set()
        footer_seen_set: set[str] = set()

        for block in text_blocks:
            norm_text = block.get("text", "").strip()
            if norm_text in header_set:
                if norm_text not in header_seen_set:
                    seen_headers.append(norm_text)
                    header_seen_set.add(norm_text)
                continue
            if norm_text in footer_set:
                if norm_text not in footer_seen_set:
                    seen_footers.append(norm_text)
                    footer_seen_set.add(norm_text)
                continue
            filtered.append(block)

        return seen_headers, seen_footers, filtered

    @staticmethod
    def _detect_alignment(
        blocks: list[dict[str, Any]],
        body_x0: float,
        page_width: float,
    ) -> str:
        """Determine text alignment from block positions.

        Returns one of ``"left"``, ``"center"``, ``"right"``, or
        ``"justify"`` (currently defaults to ``"left"`` when ambiguous).
        """
        if not blocks or page_width <= 0:
            return "left"

        tolerance = _ALIGNMENT_TOLERANCE
        avg_x0 = sum(b.get("x0", 0) for b in blocks) / len(blocks)
        avg_x1 = sum(b.get("x1", 0) for b in blocks) / len(blocks)

        text_center = (avg_x0 + avg_x1) / 2
        page_center = page_width / 2

        # Center-aligned: midpoint close to page center.
        if abs(text_center - page_center) < tolerance:
            return "center"

        # Right-aligned: x1 close to page width (right margin) but x0 varies.
        right_margin = page_width * 0.95  # Allow 5 % margin from edge.
        if avg_x1 >= right_margin and avg_x0 > body_x0 + tolerance:
            return "right"

        # Default: left-aligned.
        return "left"

    # ------------------------------------------------------------------
    # Element constructors
    # ------------------------------------------------------------------

    @staticmethod
    def _make_image_element(img: dict[str, Any]) -> dict[str, Any]:
        """Convert an image dict from the extractor into a semantic element."""
        return {
            "type": "image",
            "path": img.get("path", ""),
            "width": float(img.get("width", 0)),
            "height": float(img.get("height", 0)),
            "page_num": img.get("page_num", 0),
            "_y0": img.get("y0", 0),
            "_x0": img.get("x0", 0),
        }

    @staticmethod
    def _make_table_element(tbl: dict[str, Any]) -> dict[str, Any]:
        """Convert a table dict from the extractor into a semantic element."""
        rows = tbl.get("rows", [])
        num_rows = tbl.get("num_rows", len(rows))
        num_cols = tbl.get("num_cols", len(rows[0]) if rows else 0)
        result: dict[str, Any] = {
            "type": "table",
            "rows": rows,
            "num_rows": num_rows,
            "num_cols": num_cols,
            "page_num": tbl.get("page_num", 0),
            "_y0": tbl.get("y0", 0),
            "_x0": tbl.get("x0", 0),
        }
        # Pass through enhanced table data if present.
        if "col_widths" in tbl:
            result["col_widths"] = tbl["col_widths"]
        if "header_row" in tbl:
            result["header_row"] = tbl["header_row"]
        if "cell_styles" in tbl:
            result["cell_styles"] = tbl["cell_styles"]
        if "table_height" in tbl:
            result["table_height"] = tbl["table_height"]
        if "row_heights" in tbl:
            result["row_heights"] = tbl["row_heights"]
        return result

    # ------------------------------------------------------------------
    # Utility helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _page_dimensions(
        pages_meta: list[dict[str, Any]], page_num: int
    ) -> tuple[float, float]:
        """Return ``(width, height)`` for a given page number.

        Falls back to sensible A4 defaults when metadata is unavailable.
        """
        if pages_meta and 0 <= page_num < len(pages_meta):
            pm = pages_meta[page_num]
            return float(pm.get("width", 612)), float(pm.get("height", 792))
        # Default to US Letter (612 × 792 pt).
        return 612.0, 792.0

    @staticmethod
    def _estimate_body_x0(blocks: list[dict[str, Any]]) -> float:
        """Estimate the typical left margin (x0) of body text.

        Uses the most common x0 value among the provided text blocks.
        """
        if not blocks:
            return 0.0
        x0_counter: dict[float, int] = {}
        for b in blocks:
            x = round(b.get("x0", 0), 1)
            text_len = max(len(b.get("text", "")), 1)
            x0_counter[x] = x0_counter.get(x, 0) + text_len
        return max(x0_counter, key=x0_counter.get)  # type: ignore[arg-type]

    # ------------------------------------------------------------------
    # Hyperlink attachment
    # ------------------------------------------------------------------

    @staticmethod
    def _attach_links(
        elements: list[dict[str, Any]],
        links: list[dict[str, Any]],
    ) -> None:
        """Attach hyperlinks to paragraph/heading elements based on overlap.

        For each link, find text elements on the same page whose text
        contains a substring that spatially overlaps with the link rect.
        Adds a ``"links"`` list to matched elements with ``text``,
        ``uri``, ``start``, and ``end`` keys.
        """
        if not links:
            return

        # Group links by page.
        links_by_page: dict[int, list[dict[str, Any]]] = {}
        for link in links:
            pg = link.get("page_num", 0)
            links_by_page.setdefault(pg, []).append(link)

        for elem in elements:
            if elem.get("type") not in ("paragraph", "heading"):
                continue
            pg = elem.get("page_num", -1)
            page_links = links_by_page.get(pg, [])
            if not page_links:
                continue

            text = elem.get("text", "")
            runs = elem.get("runs", [])
            if not text or not runs:
                continue

            matched: list[dict[str, Any]] = []
            for link in page_links:
                link_bbox = link.get("bbox", (0, 0, 0, 0))
                lx0, ly0, lx1, ly1 = link_bbox

                # Find runs whose bbox overlaps the link rect.
                link_text_parts: list[str] = []
                for run in runs:
                    rx0 = run.get("x0", 0)
                    ry0 = run.get("y0", 0)
                    rx1 = run.get("x1", 0)
                    ry1 = run.get("y1", 0)
                    # Check overlap.
                    if rx1 > lx0 and rx0 < lx1 and ry1 > ly0 and ry0 < ly1:
                        link_text_parts.append(run.get("text", ""))

                if link_text_parts:
                    link_text = " ".join(link_text_parts).strip()
                    # Find position in the full text.
                    start = text.find(link_text)
                    if start >= 0:
                        matched.append({
                            "text": link_text,
                            "uri": link.get("uri", ""),
                            "start": start,
                            "end": start + len(link_text),
                        })

            if matched:
                elem["links"] = matched
