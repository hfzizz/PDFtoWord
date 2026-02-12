"""Font analyzer module for identifying fonts, heading levels, and fallback mapping.

Examines text block font metadata to determine the body font, heading fonts
with levels, and a mapping from PDF font names to Word-compatible fallbacks.
"""

import logging
import re
from collections import Counter
from typing import Any

logger = logging.getLogger(__name__)

# Subset prefix pattern (e.g. "ABCDEF+TimesNewRoman" → "TimesNewRoman").
_SUBSET_PREFIX_RE = re.compile(r"^[A-Z]{6}\+")

# --- Known font family sets ------------------------------------------------
_SERIF_FONTS = {"times", "georgia", "garamond", "cambria", "palatino"}
_MONO_FONTS = {"courier", "consolas", "monaco"}
_SANS_FONTS = {"helvetica", "arial", "calibri", "verdana", "tahoma"}


class FontAnalyzer:
    """Analyzes font usage across text blocks to classify body vs. heading fonts.

    Produces a font fallback map from PDF-embedded font names to standard
    Word-compatible font families, and determines heading levels based on
    relative font sizes.
    """

    # Heading size thresholds relative to the body font size.
    HEADING1_RATIO = 1.4
    HEADING2_RATIO = 1.2
    HEADING3_RATIO = 1.05

    def analyze(self, text_blocks: list[dict[str, Any]]) -> dict[str, Any]:
        """Analyze text blocks to determine body font, heading fonts, and font map.

        Args:
            text_blocks: List of text block dicts.  Each dict should contain
                at least ``font`` (str) and ``size`` (float) keys.  ``bold``
                (bool) is optional and used for Heading 3 detection.

        Returns:
            A dict with:
                - ``body_font``: ``{"name": str, "size": float}``
                - ``heading_fonts``: list of
                  ``{"name": str, "size": float, "level": int}``
                - ``font_map``: dict mapping original font names to fallback
                  Word font names.
        """
        if not text_blocks:
            logger.debug("No text blocks provided; returning empty font analysis.")
            return {
                "body_font": {"name": "Arial", "size": 12.0},
                "heading_fonts": [],
                "font_map": {},
            }

        # --- Step 1: Normalize font names and count (font, size) combos ----
        font_size_counter: Counter[tuple[str, float]] = Counter()
        all_font_names: set[str] = set()

        for block in text_blocks:
            raw_font = block.get("font", "")
            size = float(block.get("size", 0))
            clean_font = self._strip_subset_prefix(raw_font)
            all_font_names.add(clean_font)

            # Weight by text length so short labels don't outweigh body text.
            text_len = max(len(block.get("text", "")), 1)
            font_size_counter[(clean_font, size)] += text_len

        # --- Step 2: Determine body font (most common combo) ---------------
        most_common = font_size_counter.most_common(1)[0]
        body_font_name, body_size = most_common[0]
        body_font: dict[str, Any] = {"name": body_font_name, "size": body_size}

        logger.debug("Body font detected: %s @ %.1fpt", body_font_name, body_size)

        # --- Step 3: Detect heading fonts ----------------------------------
        heading_fonts = self._detect_heading_fonts(
            text_blocks, font_size_counter, body_size
        )

        # --- Step 4: Build font fallback map --------------------------------
        font_map = {name: self._map_to_fallback(name) for name in all_font_names}

        return {
            "body_font": body_font,
            "heading_fonts": heading_fonts,
            "font_map": font_map,
        }

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _strip_subset_prefix(font_name: str) -> str:
        """Remove a 6-letter uppercase subset prefix (e.g. ``ABCDEF+``)."""
        return _SUBSET_PREFIX_RE.sub("", font_name)

    @staticmethod
    def _map_to_fallback(font_name: str) -> str:
        """Map a PDF font name to a Word-compatible fallback font family.

        Rules (evaluated in order):
        1. Contains "serif" (case-insensitive) or matches a known serif → Times New Roman
        2. Contains "mono" or matches a known monospace → Courier New
        3. Contains "sans" or matches a known sans-serif → Arial
        4. Default → Arial
        """
        lower = font_name.lower().replace("-", "").replace(" ", "")

        # Check serif (but not sans-serif — "sans" checked later).
        if "serif" in lower and "sans" not in lower:
            return "Times New Roman"
        for family in _SERIF_FONTS:
            if family in lower:
                return "Times New Roman"

        # Check monospace.
        if "mono" in lower:
            return "Courier New"
        for family in _MONO_FONTS:
            if family in lower:
                return "Courier New"

        # Check sans-serif.
        if "sans" in lower:
            return "Arial"
        for family in _SANS_FONTS:
            if family in lower:
                return "Arial"

        # Default fallback.
        return "Arial"

    def _detect_heading_fonts(
        self,
        text_blocks: list[dict[str, Any]],
        font_size_counter: Counter[tuple[str, float]],
        body_size: float,
    ) -> list[dict[str, Any]]:
        """Identify heading fonts and assign heading levels.

        Heading levels are assigned based on size ratio to body:
            - >= 1.4× body → Heading 1
            - >= 1.2× body → Heading 2
            - >= 1.05× body AND bold → Heading 3
        """
        if body_size <= 0:
            return []

        # Gather unique (font, size) combos that are larger than body.
        seen: set[tuple[str, float]] = set()
        headings: list[dict[str, Any]] = []

        # Build a quick lookup: (font, size) → has_bold
        bold_lookup: dict[tuple[str, float], bool] = {}
        for block in text_blocks:
            clean = self._strip_subset_prefix(block.get("font", ""))
            size = float(block.get("size", 0))
            is_bold = bool(block.get("bold", False))
            key = (clean, size)
            if is_bold:
                bold_lookup[key] = True
            else:
                bold_lookup.setdefault(key, False)

        for (font, size) in font_size_counter:
            if size <= body_size:
                continue
            key = (font, size)
            if key in seen:
                continue
            seen.add(key)

            ratio = size / body_size
            is_bold = bold_lookup.get(key, False)

            if ratio >= self.HEADING1_RATIO:
                level = 1
            elif ratio >= self.HEADING2_RATIO:
                level = 2
            elif ratio >= self.HEADING3_RATIO and is_bold:
                level = 3
            else:
                # Not large enough (or not bold for level 3) — skip.
                continue

            headings.append({"name": font, "size": size, "level": level})
            logger.debug(
                "Heading level %d: %s @ %.1fpt (ratio %.2f)",
                level, font, size, ratio,
            )

        # Sort headings by level then size descending for deterministic output.
        headings.sort(key=lambda h: (h["level"], -h["size"]))
        return headings
