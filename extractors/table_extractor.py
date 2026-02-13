"""Table extraction from PDF pages using PyMuPDF's built-in table detection."""

import logging
from typing import Any

import fitz

logger = logging.getLogger(__name__)

# PyMuPDF span flag bits
_BOLD_FLAG = 1 << 4
_ITALIC_FLAG = 1 << 1


class TableExtractor:
    """Extracts tables from a PDF page using ``page.find_tables()``.

    Relies on PyMuPDF ≥ 1.23.0 which ships the ``find_tables`` method.
    If the method is unavailable or extraction fails the extractor
    degrades gracefully and returns an empty list.
    """

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def extract(self, page: fitz.Page, page_num: int) -> list[dict[str, Any]]:
        """Detect and extract tables on *page*.

        Parameters
        ----------
        page:
            A ``fitz.Page`` object.
        page_num:
            0-based page index added to each result dict.

        Returns
        -------
        list[dict[str, Any]]
            Each dict contains ``bbox``, ``rows``, ``num_rows``,
            ``num_cols``, ``col_widths``, ``header_row`` and
            ``page_num``.  Returns an empty list when
            ``find_tables`` is not available or extraction fails.
        """
        try:
            if not hasattr(page, "find_tables"):
                logger.warning(
                    "page.find_tables() not available – upgrade PyMuPDF to ≥ 1.23.0"
                )
                return []

            tables = page.find_tables()
            results: list[dict[str, Any]] = []

            for table in tables:
                try:
                    bbox = tuple(table.bbox)  # fitz.Rect → (x0, y0, x1, y1)
                    raw_rows: list[list[str | None]] = table.extract()

                    # Replace None (merged / empty cells) with empty strings
                    rows: list[list[str]] = [
                        [cell if cell is not None else "" for cell in row]
                        for row in raw_rows
                    ]

                    # Collapse columns that are entirely (or nearly) empty.
                    # PyMuPDF's grid can produce many sub-columns for
                    # merged cells; this reduces them to logical columns.
                    rows = self._collapse_empty_columns(rows)

                    # Remove rows that are entirely empty (whitespace-only).
                    rows = [
                        row for row in rows
                        if any(cell.strip() for cell in row)
                    ]

                    num_rows = len(rows)
                    num_cols = max((len(r) for r in rows), default=0)

                    # --- column widths (proportional) ----------------------
                    # After column collapsing, the original cell geometry
                    # no longer maps to the logical columns.  Compute
                    # proportional widths from the bbox for collapsed tables.
                    original_num_cols = max(
                        (len(r) for r in raw_rows), default=0
                    )
                    if num_cols != original_num_cols:
                        # Collapsed: infer widths from text positions on page.
                        col_widths = self._compute_col_widths_from_text(
                            page, bbox, num_cols, rows,
                        )
                    else:
                        col_widths = self._compute_col_widths(table, num_cols)

                    # --- header row detection ------------------------------
                    header_row = self._detect_header_row(page, table, rows)

                    collapsed = (num_cols != original_num_cols)

                    # --- per-row heights from cell geometry ----------------
                    table_height = bbox[3] - bbox[1]
                    row_heights = self._compute_row_heights(table, num_rows, table_height)

                    results.append(
                        {
                            "bbox": bbox,
                            "rows": rows,
                            "num_rows": num_rows,
                            "num_cols": num_cols,
                            "col_widths": col_widths,
                            "header_row": header_row,
                            "page_num": page_num,
                            "table_height": table_height,
                            "row_heights": row_heights,
                            "cell_styles": self._extract_cell_styles(
                                page, table, rows, num_rows, num_cols,
                                col_widths=col_widths,
                                collapsed=collapsed,
                            ),
                        }
                    )
                except Exception:
                    logger.warning(
                        "Failed to extract data from a table on page %s – skipping",
                        page_num,
                        exc_info=True,
                    )
                    continue

            return results

        except Exception:
            logger.exception("Table extraction failed on page %s", page_num)
            return []

    # ------------------------------------------------------------------
    # Per-cell style extraction
    # ------------------------------------------------------------------

    @staticmethod
    def _extract_cell_styles(
        page: fitz.Page,
        table: Any,
        rows: list[list[str]],
        num_rows: int,
        num_cols: int,
        *,
        col_widths: list[float] | None = None,
        collapsed: bool = False,
    ) -> list[list[dict[str, Any]]]:
        """Extract per-cell formatting: background color, font styles, alignment.

        Returns a 2-D list matching ``rows`` dimensions.  Each entry is a
        dict with optional keys:
            - ``bg_color``: ``(r, g, b)`` tuple (0–255) or ``None``
            - ``font``: str — dominant font name in this cell
            - ``size``: float — dominant font size
            - ``bold``: bool
            - ``italic``: bool
            - ``color``: ``(r, g, b)`` text colour
            - ``alignment``: ``"left"`` | ``"center"`` | ``"right"``

        When *collapsed* is True, ``runs`` are omitted from cell styles
        because the approximate cell rectangles may not match the true
        cell boundaries, leading to garbled text content.
        """
        raw_cells = getattr(table, "cells", None)
        raw_num_cols = 0
        if raw_cells:
            # Determine original grid column count from the raw cell list.
            raw_extract = table.extract()
            raw_num_cols = max((len(r) for r in raw_extract), default=0)

        # Build the table bounding box for fallback cell rects.
        bbox = table.bbox
        tbl_rect = fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3])

        # Pre-process page drawings for efficient border detection.
        h_lines, v_lines = TableExtractor._collect_page_lines(page)

        styles: list[list[dict[str, Any]]] = []

        for r_idx in range(num_rows):
            row_styles: list[dict[str, Any]] = []
            for c_idx in range(num_cols):
                cell_style: dict[str, Any] = {}

                # Try to get the cell rectangle from the table geometry.
                cell_rect: fitz.Rect | None = None
                if raw_cells and raw_num_cols > 0 and num_cols == raw_num_cols:
                    # Grid was NOT collapsed — direct cell lookup.
                    flat_idx = r_idx * raw_num_cols + c_idx
                    if flat_idx < len(raw_cells):
                        cr = raw_cells[flat_idx]
                        cell_rect = fitz.Rect(cr[0], cr[1], cr[2], cr[3])
                else:
                    # Grid was collapsed — approximate cell rect using
                    # proportional column widths and equal row heights.
                    row_h = tbl_rect.height / num_rows if num_rows > 0 else tbl_rect.height
                    if col_widths and len(col_widths) == num_cols:
                        x0 = tbl_rect.x0 + sum(col_widths[:c_idx]) * tbl_rect.width
                        col_w = col_widths[c_idx] * tbl_rect.width
                    else:
                        col_w = tbl_rect.width / num_cols if num_cols > 0 else tbl_rect.width
                        x0 = tbl_rect.x0 + c_idx * col_w
                    y0 = tbl_rect.y0 + r_idx * row_h
                    cell_rect = fitz.Rect(x0, y0, x0 + col_w, y0 + row_h)

                if cell_rect is not None:
                    # --- Background colour detection ---------------------------
                    bg = TableExtractor._detect_cell_background(page, cell_rect)
                    if bg is not None:
                        cell_style["bg_color"] = bg

                    # --- Text formatting inside cell ---------------------------
                    text_info = TableExtractor._extract_cell_text_formatting(
                        page, cell_rect
                    )
                    if text_info:
                        cell_style.update(text_info)

                    # --- Cell border detection ---------------------------------
                    borders = TableExtractor._match_border(
                        cell_rect, h_lines, v_lines
                    )
                    if borders:
                        cell_style["borders"] = borders

                    # For collapsed tables the approximate cell rects may
                    # span multiple visual cells, producing garbled run
                    # text.  Keep only the dominant formatting metadata.
                    if collapsed:
                        cell_style.pop("runs", None)

                row_styles.append(cell_style)
            styles.append(row_styles)

        return styles

    @staticmethod
    def _detect_cell_background(
        page: fitz.Page, rect: fitz.Rect
    ) -> tuple[int, int, int] | None:
        """Detect the background fill colour of a cell area.

        Inspects drawing commands (paths) that fill the cell region.
        Returns an ``(r, g, b)`` tuple (0–255) or ``None`` if no fill
        is found or the fill is white.
        """
        try:
            drawings = page.get_drawings()
            best_fill: tuple[int, int, int] | None = None
            best_area = 0.0

            for d in drawings:
                fill = d.get("fill")
                if fill is None:
                    continue
                # fill is (r, g, b) with floats 0.0–1.0
                d_rect = d.get("rect")
                if d_rect is None:
                    continue
                dr = fitz.Rect(d_rect)
                # Check if this drawing substantially overlaps the cell.
                overlap = dr & rect  # intersection
                if overlap.is_empty:
                    continue
                overlap_area = overlap.width * overlap.height
                cell_area = rect.width * rect.height
                if cell_area > 0 and overlap_area / cell_area > 0.5:
                    if overlap_area > best_area:
                        best_area = overlap_area
                        r = int(fill[0] * 255)
                        g = int(fill[1] * 255)
                        b = int(fill[2] * 255)
                        # Skip pure white — that's just paper.
                        if (r, g, b) != (255, 255, 255):
                            best_fill = (r, g, b)
            return best_fill
        except Exception:
            return None

    @staticmethod
    def _extract_cell_text_formatting(
        page: fitz.Page, rect: fitz.Rect
    ) -> dict[str, Any]:
        """Extract dominant text formatting from spans inside *rect*."""
        info: dict[str, Any] = {}
        try:
            td = page.get_text("dict", clip=rect, flags=0)
            spans: list[dict[str, Any]] = []
            for block in td.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        if span.get("text", "").strip():
                            spans.append(span)
            if not spans:
                return info

            # Dominant span = longest text
            dom = max(spans, key=lambda s: len(s.get("text", "")))
            flags = dom.get("flags", 0)
            font_name = dom.get("font", "")

            info["font"] = font_name
            info["size"] = float(dom.get("size", 0))
            info["bold"] = bool(flags & _BOLD_FLAG) or "bold" in font_name.lower()
            info["italic"] = bool(flags & _ITALIC_FLAG) or "italic" in font_name.lower()

            # Text colour
            color_int = dom.get("color", 0)
            info["color"] = (
                (color_int >> 16) & 0xFF,
                (color_int >> 8) & 0xFF,
                color_int & 0xFF,
            )

            # --- Underline / strikethrough detection -----------------------
            # PyMuPDF does not expose underline in span flags.  We detect
            # them by looking for thin horizontal line drawings that sit
            # at the bottom (underline) or middle (strikethrough) of the
            # text line.
            info["underline"] = False
            info["strikethrough"] = False
            try:
                drawings = page.get_drawings()
                for d in drawings:
                    if d.get("fill") is not None:
                        continue  # filled rects are backgrounds, not lines
                    for item in d.get("items", []):
                        if item[0] != "l":  # 'l' = line
                            continue
                        p1, p2 = fitz.Point(item[1]), fitz.Point(item[2])
                        # Must be roughly horizontal & within the cell
                        if abs(p1.y - p2.y) > 2:
                            continue
                        if p1.x < rect.x0 - 2 or p2.x > rect.x1 + 2:
                            continue
                        mid_y = (p1.y + p2.y) / 2
                        if rect.y0 <= mid_y <= rect.y1:
                            line_len = abs(p2.x - p1.x)
                            if line_len < 5:
                                continue
                            # Underline: line near bottom of text
                            text_bottom = rect.y1
                            text_top = rect.y0
                            text_mid = (text_top + text_bottom) / 2
                            if mid_y > text_mid:
                                info["underline"] = True
                            else:
                                info["strikethrough"] = True
            except Exception:
                pass

            # Alignment: compare span x-positions to cell rect
            all_x0 = [s["bbox"][0] for s in spans if "bbox" in s]
            all_x1 = [s["bbox"][2] for s in spans if "bbox" in s]
            if all_x0 and all_x1:
                avg_x0 = sum(all_x0) / len(all_x0)
                avg_x1 = sum(all_x1) / len(all_x1)
                cell_center = (rect.x0 + rect.x1) / 2
                text_center = (avg_x0 + avg_x1) / 2
                left_gap = avg_x0 - rect.x0
                right_gap = rect.x1 - avg_x1

                if abs(text_center - cell_center) < 5:
                    info["alignment"] = "center"
                elif right_gap < 5 and left_gap > 10:
                    info["alignment"] = "right"
                else:
                    info["alignment"] = "left"

            # Build per-span runs for mixed formatting within cell
            if len(spans) > 1:
                runs: list[dict[str, Any]] = []
                for s in spans:
                    sf = s.get("flags", 0)
                    sfont = s.get("font", "")
                    sc = s.get("color", 0)
                    runs.append({
                        "text": s.get("text", ""),
                        "font": sfont,
                        "size": float(s.get("size", 0)),
                        "bold": bool(sf & _BOLD_FLAG) or "bold" in sfont.lower(),
                        "italic": bool(sf & _ITALIC_FLAG) or "italic" in sfont.lower(),
                        "color": ((sc >> 16) & 0xFF, (sc >> 8) & 0xFF, sc & 0xFF),
                    })
                info["runs"] = runs

        except Exception:
            logger.debug("Cell text extraction failed", exc_info=True)
        return info

    # ------------------------------------------------------------------
    # Border detection helpers  (inspired by pdf2docx library)
    # ------------------------------------------------------------------

    @staticmethod
    def _collect_page_lines(
        page: fitz.Page,
    ) -> tuple[list[tuple[float, float, float, float, tuple[int, int, int]]],
               list[tuple[float, float, float, float, tuple[int, int, int]]]]:
        """Pre-process page drawings into horizontal and vertical segments.

        Extracts line and rectangle-edge strokes from the page's drawing
        commands once so that border matching per cell is cheap.

        Returns ``(h_lines, v_lines)`` where each entry is
        ``(start, end, pos, width, (r, g, b))``:
        - *h_lines*: ``(x0, x1, y, stroke_width, rgb)``
        - *v_lines*: ``(y0, y1, x, stroke_width, rgb)``
        """
        h_lines: list[tuple[float, float, float, float, tuple[int, int, int]]] = []
        v_lines: list[tuple[float, float, float, float, tuple[int, int, int]]] = []
        tolerance = 2.0

        try:
            drawings = page.get_drawings()
        except Exception:
            return h_lines, v_lines

        for d in drawings:
            stroke_color = d.get("color")
            if stroke_color is None:
                continue
            stroke_width = d.get("width", 1.0)
            rgb = (
                int(stroke_color[0] * 255),
                int(stroke_color[1] * 255),
                int(stroke_color[2] * 255),
            )

            for item in d.get("items", []):
                kind = item[0]
                if kind == "l":  # line segment
                    p1, p2 = fitz.Point(item[1]), fitz.Point(item[2])
                    if abs(p1.y - p2.y) <= tolerance:
                        h_lines.append((
                            min(p1.x, p2.x), max(p1.x, p2.x),
                            (p1.y + p2.y) / 2, stroke_width, rgb,
                        ))
                    elif abs(p1.x - p2.x) <= tolerance:
                        v_lines.append((
                            min(p1.y, p2.y), max(p1.y, p2.y),
                            (p1.x + p2.x) / 2, stroke_width, rgb,
                        ))
                elif kind == "re":  # rectangle → 4 edges
                    rect = fitz.Rect(item[1])
                    if rect.width < 0.5 or rect.height < 0.5:
                        continue
                    h_lines.append((rect.x0, rect.x1, rect.y0, stroke_width, rgb))
                    h_lines.append((rect.x0, rect.x1, rect.y1, stroke_width, rgb))
                    v_lines.append((rect.y0, rect.y1, rect.x0, stroke_width, rgb))
                    v_lines.append((rect.y0, rect.y1, rect.x1, stroke_width, rgb))

        return h_lines, v_lines

    @staticmethod
    def _match_border(
        cell_rect: fitz.Rect,
        h_lines: list[tuple[float, float, float, float, tuple[int, int, int]]],
        v_lines: list[tuple[float, float, float, float, tuple[int, int, int]]],
    ) -> dict[str, dict[str, Any]]:
        """Match pre-processed line segments to cell edges.

        For each side of *cell_rect*, finds the closest line that is
        aligned with and overlaps that edge.

        Returns a dict with optional keys ``"top"``, ``"bottom"``,
        ``"left"``, ``"right"``, each mapping to
        ``{"width": float, "color": (r, g, b)}``.
        """
        borders: dict[str, dict[str, Any]] = {}
        tol = 3.0
        cx0, cy0, cx1, cy1 = (
            cell_rect.x0, cell_rect.y0, cell_rect.x1, cell_rect.y1,
        )
        cell_w = cx1 - cx0
        cell_h = cy1 - cy0

        # --- horizontal edges (top / bottom) -------------------------------
        for edge_y, side in ((cy0, "top"), (cy1, "bottom")):
            best_dist = tol + 1
            best_info = None
            for x0, x1, y, w, c in h_lines:
                dist = abs(y - edge_y)
                if dist <= tol and dist < best_dist:
                    overlap = min(x1, cx1) - max(x0, cx0)
                    if cell_w > 0 and overlap / cell_w > 0.3:
                        best_dist = dist
                        best_info = {"width": w, "color": c}
            if best_info:
                borders[side] = best_info

        # --- vertical edges (left / right) ---------------------------------
        for edge_x, side in ((cx0, "left"), (cx1, "right")):
            best_dist = tol + 1
            best_info = None
            for y0, y1, x, w, c in v_lines:
                dist = abs(x - edge_x)
                if dist <= tol and dist < best_dist:
                    overlap = min(y1, cy1) - max(y0, cy0)
                    if cell_h > 0 and overlap / cell_h > 0.3:
                        best_dist = dist
                        best_info = {"width": w, "color": c}
            if best_info:
                borders[side] = best_info

        return borders

    # ------------------------------------------------------------------
    # Empty-column collapsing
    # ------------------------------------------------------------------

    @staticmethod
    def _collapse_empty_columns(
        rows: list[list[str]],
    ) -> list[list[str]]:
        """Remove columns where all cells are empty.

        PyMuPDF's ``find_tables()`` creates a fine-grained grid that
        splits merged cells into many sub-columns.  Most of those
        sub-columns are entirely empty (or contain ``None`` which was
        already mapped to ``""``).  Removing them yields the logical
        column structure the user expects.

        Additionally, when two adjacent columns are "complementary"
        (one has mostly headers, the other has mostly data values,
        and they never overlap in the same row), they are merged.
        """
        if not rows:
            return rows

        num_cols = max(len(r) for r in rows)
        if num_cols <= 1:
            return rows

        # 1. Identify columns that are entirely empty.
        keep: list[bool] = []
        for col_idx in range(num_cols):
            has_content = False
            for row in rows:
                if col_idx < len(row) and row[col_idx].strip():
                    has_content = True
                    break
            keep.append(has_content)

        # 2. Build filtered rows retaining only non-empty columns.
        kept_indices = [i for i, k in enumerate(keep) if k]
        if not kept_indices:
            return rows  # All empty — return as-is.

        filtered_rows: list[list[str]] = []
        for row in rows:
            filtered_rows.append([
                row[i] if i < len(row) else "" for i in kept_indices
            ])

        # 3. Merge adjacent complementary columns: two columns where
        #    the cells are never both non-empty in the same row.
        merged = True
        while merged:
            merged = False
            new_cols = len(filtered_rows[0]) if filtered_rows else 0
            if new_cols <= 1:
                break
            for ci in range(new_cols - 1):
                # Check if columns ci and ci+1 are complementary.
                overlap = False
                for row in filtered_rows:
                    a = row[ci].strip() if ci < len(row) else ""
                    b = row[ci + 1].strip() if ci + 1 < len(row) else ""
                    if a and b:
                        overlap = True
                        break
                if not overlap:
                    # Merge ci+1 into ci.
                    for row in filtered_rows:
                        a = row[ci].strip() if ci < len(row) else ""
                        b = row[ci + 1].strip() if ci + 1 < len(row) else ""
                        row[ci] = (a + " " + b).strip()
                    # Remove column ci+1.
                    for row in filtered_rows:
                        if ci + 1 < len(row):
                            del row[ci + 1]
                    merged = True
                    break  # Restart scan after structural change.

        return filtered_rows

    # ------------------------------------------------------------------
    # Column-width computation
    # ------------------------------------------------------------------

    @staticmethod
    def _compute_row_heights(
        table: Any,
        num_rows: int,
        table_height: float,
    ) -> list[float]:
        """Compute per-row heights from cell geometry.

        Returns fractional heights summing to ~1.0.  Falls back to
        equal heights if cell geometry is unavailable.
        """
        try:
            cells = table.cells
            if not cells:
                return [1.0 / num_rows] * num_rows

            # Gather unique y-boundaries from cell rects.
            y_set: set[float] = set()
            for cell in cells:
                if cell is not None:
                    y_set.add(round(cell[1], 1))  # y0
                    y_set.add(round(cell[3], 1))  # y1

            y_sorted = sorted(y_set)
            if len(y_sorted) < 2:
                return [1.0 / num_rows] * num_rows

            # Raw row heights from consecutive Y boundaries.
            raw_heights: list[float] = []
            for i in range(len(y_sorted) - 1):
                raw_heights.append(y_sorted[i + 1] - y_sorted[i])

            if len(raw_heights) != num_rows:
                # Mismatch — row count changed due to collapsing /
                # empty-row removal.  Fall back to uniform.
                return [1.0 / num_rows] * num_rows

            total = sum(raw_heights) or 1.0
            return [h / total for h in raw_heights]

        except Exception:
            return [1.0 / num_rows] * num_rows

    @staticmethod
    def _compute_col_widths_from_text(
        page: fitz.Page,
        table_bbox: tuple[float, float, float, float],
        num_cols: int,
        rows: list[list[str]],
    ) -> list[float]:
        """Infer proportional column widths from text positions on the page.

        For collapsed tables, PyMuPix cell geometry is unreliable
        (transposed or merged coordinates).  Instead, we extract all
        text spans that fall within the table bounding box and cluster
        them into columns by detecting x-position gaps.

        Falls back to equal widths when the analysis cannot determine
        column boundaries.
        """
        if num_cols <= 0:
            return []
        equal: list[float] = [1.0 / num_cols] * num_cols

        try:
            tx0, ty0, tx1, ty1 = table_bbox
            table_width = tx1 - tx0
            if table_width <= 0:
                return equal

            # Collect all text spans within the table region.
            blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
            spans: list[dict] = []
            for block in blocks:
                if block.get("type") != 0:
                    continue
                for line in block["lines"]:
                    for span in line["spans"]:
                        sb = span["bbox"]
                        text = span["text"].strip()
                        if not text:
                            continue
                        # Must overlap with table bbox.
                        if sb[2] < tx0 or sb[0] > tx1 or sb[3] < ty0 or sb[1] > ty1:
                            continue
                        spans.append({"x0": sb[0], "x1": sb[2], "text": text})

            if not spans:
                return equal

            # Sort spans by x-midpoint and find gaps to determine column
            # boundaries.  A gap is defined as a break in x-coverage that
            # exceeds a threshold.
            spans.sort(key=lambda s: (s["x0"] + s["x1"]) / 2)

            # Build a list of x-intervals (merge overlapping spans).
            intervals: list[list[float]] = []
            for sp in spans:
                if intervals and sp["x0"] <= intervals[-1][1] + 2.0:
                    # Extend interval in both directions.
                    intervals[-1][0] = min(intervals[-1][0], sp["x0"])
                    intervals[-1][1] = max(intervals[-1][1], sp["x1"])
                else:
                    intervals.append([sp["x0"], sp["x1"]])

            if len(intervals) < 2:
                return equal

            # Find the (num_cols - 1) largest gaps between intervals.
            gaps: list[tuple[float, float, float]] = []  # (gap_size, gap_start, gap_end)
            for i in range(len(intervals) - 1):
                gap_start = intervals[i][1]
                gap_end = intervals[i + 1][0]
                gap_size = gap_end - gap_start
                if gap_size > 0:
                    gaps.append((gap_size, gap_start, gap_end))

            gaps.sort(key=lambda g: g[0], reverse=True)

            if len(gaps) < num_cols - 1:
                return equal

            # Use the top (num_cols - 1) gaps as column boundaries.
            boundaries = sorted(
                [(g[1] + g[2]) / 2 for g in gaps[: num_cols - 1]]
            )

            # Build edges: table_left, boundary1, boundary2, ..., table_right
            edges = [tx0] + boundaries + [tx1]
            widths = [(edges[i + 1] - edges[i]) / table_width for i in range(num_cols)]

            # Sanity check: reject if any column is negative or unreasonably
            # small (< 1%).
            if any(w < 0.01 for w in widths):
                return equal

            # Normalise.
            total = sum(widths)
            if total > 0:
                widths = [w / total for w in widths]
            return widths

        except Exception:
            logger.debug("Text-based column width inference failed", exc_info=True)
            return equal

    @staticmethod
    def _compute_col_widths(
        table: Any, num_cols: int
    ) -> list[float]:
        """Return proportional column widths (fractions summing to ~1.0).

        Uses ``table.cells`` – a flat list of ``(x0, y0, x1, y1)``
        tuples in row-major order — to determine the unique column
        boundaries from the first row, then derives each column's
        proportional width relative to the full table width.

        Falls back to equal widths (``1/num_cols``) when cell geometry
        is unavailable.
        """
        if num_cols <= 0:
            return []

        cells = getattr(table, "cells", None)
        if not cells:
            return [1.0 / num_cols] * num_cols

        # The first row contains the first ``num_cols`` cells.
        first_row_cells = cells[:num_cols]

        # Collect the unique left (x0) and rightmost (x1) edges.
        x_edges: list[float] = []
        for cell_rect in first_row_cells:
            x0 = cell_rect[0]
            if x0 not in x_edges:
                x_edges.append(x0)
        # Append the right edge of the last cell in the first row.
        last_x1 = first_row_cells[-1][2]
        x_edges.append(last_x1)

        x_edges.sort()

        # Sanity check — we expect exactly ``num_cols + 1`` edges.
        if len(x_edges) < 2:
            return [1.0 / num_cols] * num_cols

        total_width = x_edges[-1] - x_edges[0]
        if total_width <= 0:
            return [1.0 / num_cols] * num_cols

        widths: list[float] = []
        for i in range(len(x_edges) - 1):
            widths.append((x_edges[i + 1] - x_edges[i]) / total_width)

        # If merged cells produced fewer edges than expected, pad with
        # equal splits for the remaining columns.
        while len(widths) < num_cols:
            widths.append(1.0 / num_cols)

        # Trim in case of extra edges (shouldn't happen, but be safe).
        widths = widths[:num_cols]

        # Normalise so the fractions sum to exactly 1.0.
        total = sum(widths)
        if total > 0:
            widths = [w / total for w in widths]

        return widths

    # ------------------------------------------------------------------
    # Header-row detection
    # ------------------------------------------------------------------

    @staticmethod
    def _detect_header_row(
        page: fitz.Page,
        table: Any,
        rows: list[list[str]],
    ) -> bool:
        """Heuristic: decide whether the first row is a header.

        Strategy
        --------
        1. If the table has fewer than 2 rows it cannot have a distinct
           header — return ``False``.
        2. Sample the text spans that fall inside the first row's
           bounding box.  If more than half of them are bold (flag
           ``2**4 = 16`` in PyMuPDF span flags), treat the row as a
           header.
        3. Fall back to ``False`` if the check cannot be performed.
        """
        if len(rows) < 2:
            return False

        try:
            cells = getattr(table, "cells", None)
            num_cols = max((len(r) for r in rows), default=0)
            if not cells or num_cols == 0:
                return False

            # Build a rect covering the first row from its cell geometry.
            first_row_cells = cells[:num_cols]
            row_x0 = min(c[0] for c in first_row_cells)
            row_y0 = min(c[1] for c in first_row_cells)
            row_x1 = max(c[2] for c in first_row_cells)
            row_y1 = max(c[3] for c in first_row_cells)
            header_rect = fitz.Rect(row_x0, row_y0, row_x1, row_y1)

            # Extract text spans that intersect the header row rect.
            text_dict = page.get_text("dict", clip=header_rect, flags=0)
            bold_count = 0
            total_count = 0
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        total_count += 1
                        # PyMuPDF span flags: bit 4 (value 16) → bold
                        if span.get("flags", 0) & 16:
                            bold_count += 1

            # Also check font name: many PDFs use a separate bold
            # font family (e.g. "Helvetica-Bold") instead of setting
            # the bold flag bit.
            bold_font_count = 0
            for block in text_dict.get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        fname = span.get("font", "").lower()
                        if "bold" in fname or "heavy" in fname or "black" in fname:
                            bold_font_count += 1

            all_bold = bold_count + bold_font_count
            if total_count > 0 and all_bold / total_count > 0.5:
                return True

        except Exception:
            logger.debug(
                "Header detection heuristic failed – defaulting to False",
                exc_info=True,
            )

        return False
