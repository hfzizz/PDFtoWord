"""Layout analyzer module for detecting column structure and page sections.

Analyzes text block positions to determine the number of columns,
column boundaries, and section/page breaks in a PDF page.
"""

import logging
from typing import Any

logger = logging.getLogger(__name__)


class LayoutAnalyzer:
    """Analyzes page layout to detect columns, sections, and structural breaks.

    Uses spatial clustering of text block positions to determine the column
    structure of a page and assigns each text block to its detected column.
    """

    # Fraction of page width that constitutes a gap between columns.
    COLUMN_GAP_THRESHOLD = 0.20
    # Fraction of page height that constitutes a section/page break.
    VERTICAL_GAP_THRESHOLD = 0.30

    def analyze(
        self,
        text_blocks: list[dict[str, Any]],
        page_width: float,
        page_height: float,
    ) -> dict[str, Any]:
        """Analyze the layout of text blocks on a page.

        Args:
            text_blocks: List of text block dicts from TextExtractor.
                Each dict is expected to have at least ``x0``, ``y0``,
                ``x1``, ``y1`` keys with float coordinates.
            page_width: Width of the page in points.
            page_height: Height of the page in points.

        Returns:
            A dict containing:
                - ``num_columns`` (int): Number of detected columns (1–3).
                - ``columns`` (list[tuple[float, float]]): Column boundary
                  tuples ``(x_start, x_end)`` for each column.
                - ``text_blocks_by_column`` (dict[int, list[dict]]): Mapping
                  of column index to the text blocks belonging to that column.
                - ``section_breaks`` (list[dict]): Detected vertical gaps that
                  suggest section or page breaks, each with ``y_start``,
                  ``y_end``, and ``gap_size`` keys.
        """
        if not text_blocks or page_width <= 0 or page_height <= 0:
            logger.debug("No text blocks or invalid page dimensions; defaulting to single column.")
            return self._single_column_result(text_blocks)

        # --- Step 1: Collect and sort unique left-edge x0 values -----------
        x0_values = sorted({block["x0"] for block in text_blocks if "x0" in block})

        if not x0_values:
            return self._single_column_result(text_blocks)

        # --- Step 2: Cluster x0 values by gap threshold --------------------
        gap_threshold = self.COLUMN_GAP_THRESHOLD * page_width
        clusters: list[list[float]] = [[x0_values[0]]]

        for x0 in x0_values[1:]:
            if x0 - clusters[-1][-1] > gap_threshold:
                clusters.append([x0])
            else:
                clusters[-1].append(x0)

        # Limit to a maximum of 3 columns.
        if len(clusters) > 3:
            logger.debug(
                "Detected %d clusters; merging down to 3 columns.", len(clusters)
            )
            clusters = self._merge_clusters_to_max(clusters, max_columns=3)

        num_columns = len(clusters)

        # --- Step 3: Determine column boundaries ---------------------------
        columns = self._compute_column_boundaries(clusters, text_blocks, page_width)

        # --- Step 4: Assign text blocks to columns -------------------------
        text_blocks_by_column: dict[int, list[dict[str, Any]]] = {
            i: [] for i in range(num_columns)
        }

        for block in text_blocks:
            col_idx = self._assign_column(block, columns)
            text_blocks_by_column[col_idx].append(block)

        # --- Step 5: Detect section breaks (large vertical gaps) -----------
        section_breaks = self._detect_section_breaks(text_blocks, page_height)

        logger.debug(
            "Layout analysis complete: %d column(s), %d section break(s).",
            num_columns,
            len(section_breaks),
        )

        return {
            "num_columns": num_columns,
            "columns": columns,
            "text_blocks_by_column": text_blocks_by_column,
            "section_breaks": section_breaks,
        }

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _single_column_result(
        text_blocks: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Return a default single-column result."""
        return {
            "num_columns": 1,
            "columns": [(0.0, 0.0)],
            "text_blocks_by_column": {0: list(text_blocks) if text_blocks else []},
            "section_breaks": [],
        }

    @staticmethod
    def _merge_clusters_to_max(
        clusters: list[list[float]], max_columns: int
    ) -> list[list[float]]:
        """Merge the closest adjacent clusters until we have at most *max_columns*."""
        while len(clusters) > max_columns:
            # Find the pair of adjacent clusters with the smallest gap.
            min_gap = float("inf")
            merge_idx = 0
            for i in range(len(clusters) - 1):
                gap = clusters[i + 1][0] - clusters[i][-1]
                if gap < min_gap:
                    min_gap = gap
                    merge_idx = i
            clusters[merge_idx] = clusters[merge_idx] + clusters[merge_idx + 1]
            del clusters[merge_idx + 1]
        return clusters

    @staticmethod
    def _compute_column_boundaries(
        clusters: list[list[float]],
        text_blocks: list[dict[str, Any]],
        page_width: float,
    ) -> list[tuple[float, float]]:
        """Compute ``(x_start, x_end)`` boundaries for each column cluster.

        ``x_start`` is the minimum x0 in the cluster.  ``x_end`` is the
        maximum x1 of any text block whose x0 falls within the cluster range.
        """
        columns: list[tuple[float, float]] = []
        for cluster in clusters:
            x_start = min(cluster)
            x_end_candidates = [
                block.get("x1", x_start)
                for block in text_blocks
                if block.get("x0", -1) >= x_start
                and block.get("x0", float("inf")) <= max(cluster)
            ]
            x_end = max(x_end_candidates) if x_end_candidates else min(page_width, x_start + 100)
            columns.append((x_start, x_end))
        return columns

    def _assign_column(
        self,
        block: dict[str, Any],
        columns: list[tuple[float, float]],
    ) -> int:
        """Assign a text block to the most appropriate column index."""
        x0 = block.get("x0", 0.0)
        best_col = 0
        best_dist = float("inf")
        for idx, (col_start, col_end) in enumerate(columns):
            # If inside the column range, distance is 0.
            if col_start <= x0 <= col_end:
                return idx
            dist = min(abs(x0 - col_start), abs(x0 - col_end))
            if dist < best_dist:
                best_dist = dist
                best_col = idx
        return best_col

    def _detect_section_breaks(
        self,
        text_blocks: list[dict[str, Any]],
        page_height: float,
    ) -> list[dict[str, float]]:
        """Detect large vertical gaps that suggest section or page breaks.

        A gap larger than ``VERTICAL_GAP_THRESHOLD * page_height`` between
        consecutive text blocks (sorted by y-position) is reported.
        """
        if len(text_blocks) < 2:
            return []

        gap_threshold = self.VERTICAL_GAP_THRESHOLD * page_height

        sorted_blocks = sorted(text_blocks, key=lambda b: (b.get("y0", 0), b.get("x0", 0)))

        breaks: list[dict[str, float]] = []
        for i in range(len(sorted_blocks) - 1):
            y_end_current = sorted_blocks[i].get("y1", sorted_blocks[i].get("y0", 0))
            y_start_next = sorted_blocks[i + 1].get("y0", 0)
            gap = y_start_next - y_end_current
            if gap > gap_threshold:
                breaks.append({
                    "y_start": y_end_current,
                    "y_end": y_start_next,
                    "gap_size": gap,
                })

        return breaks
