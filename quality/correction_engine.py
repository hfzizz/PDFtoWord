"""Correction engine — applies fixes based on AI-detected differences.

Takes a list of structured differences (from ``AIComparator``) and
modifies the DOCX document to address them.  This is the auto-fix
component of the compare → fix → re-render loop.
"""

import logging
from typing import Any

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

logger = logging.getLogger(__name__)

# Constants
MAX_ISSUE_LOG_LENGTH = 50


class CorrectionEngine:
    """Apply corrections to a DOCX based on detected differences.

    Parameters
    ----------
    docx_path : str
        Path to the DOCX file to modify.
    """

    def __init__(self, docx_path: str) -> None:
        self.docx_path = docx_path
        self._doc = Document(docx_path)
        self._fixes_applied = 0

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def apply_fixes(
        self,
        differences: list[dict[str, Any]],
    ) -> int:
        """Apply corrections for the given list of differences.

        Parameters
        ----------
        differences : list[dict]
            Flat list of difference dicts across all pages.
            Each must have ``type``, ``issue``, ``area``, ``severity``.

        Returns
        -------
        int
            Number of fixes successfully applied.
        """
        if not differences:
            logger.debug("No differences to fix.")
            return 0
        
        self._fixes_applied = 0

        # Validate input
        valid_diffs = []
        for i, diff in enumerate(differences):
            if not isinstance(diff, dict):
                logger.warning("Skipping non-dict difference at index %d", i)
                continue
            if "type" not in diff:
                logger.warning("Skipping difference without 'type' field: %s", diff.get("issue", "unknown"))
                continue
            valid_diffs.append(diff)
        
        if not valid_diffs:
            logger.warning("No valid differences to fix after validation.")
            return 0

        logger.info("Applying corrections for %d difference(s).", len(valid_diffs))

        # Group by type for batch processing.
        for diff in valid_diffs:
            diff_type = diff.get("type", "")
            try:
                handler = self._handlers.get(diff_type)
                if handler:
                    before = self._fixes_applied
                    handler(self, diff)
                    after = self._fixes_applied
                    if after > before:
                        logger.debug(
                            "Applied fix for '%s': %s",
                            diff_type,
                            diff.get("issue", "")[:MAX_ISSUE_LOG_LENGTH],
                        )
                else:
                    logger.debug(
                        "No handler for difference type '%s': %s",
                        diff_type,
                        diff.get("issue", "")[:MAX_ISSUE_LOG_LENGTH],
                    )
            except Exception as e:
                logger.warning(
                    "Failed to apply fix for '%s': %s (error: %s)",
                    diff_type,
                    diff.get("issue", "")[:MAX_ISSUE_LOG_LENGTH],
                    str(e),
                )

        if self._fixes_applied > 0:
            try:
                self._doc.save(self.docx_path)
                logger.info(
                    "Correction engine: %d fix(es) applied and saved.",
                    self._fixes_applied,
                )
            except Exception as e:
                logger.error("Failed to save corrected document: %s", e)
                raise

        return self._fixes_applied

    # ------------------------------------------------------------------
    # Fix handlers
    # ------------------------------------------------------------------

    def _fix_font_size(self, diff: dict[str, Any]) -> None:
        """Attempt to adjust font sizes based on AI feedback."""
        issue = diff.get("issue", "").lower()
        # Parse direction: "smaller" or "larger"
        direction = 0
        if "smaller" in issue or "too small" in issue:
            direction = 1  # increase
        elif "larger" in issue or "too large" in issue or "bigger" in issue:
            direction = -1  # decrease

        if direction == 0:
            return

        area = diff.get("area", "").lower()
        page_num = diff.get("page_num", 0)

        # Try to find matching paragraphs.
        for para in self._doc.paragraphs:
            for run in para.runs:
                if run.font.size is not None:
                    current_pt = run.font.size.pt
                    adjustment = max(0.5, current_pt * 0.05)
                    new_size = current_pt + (direction * adjustment)
                    run.font.size = Pt(new_size)
                    self._fixes_applied += 1
                    return  # Apply to first match only.

    def _fix_alignment(self, diff: dict[str, Any]) -> None:
        """Fix paragraph alignment issues."""
        issue = diff.get("issue", "").lower()
        target = None
        if "center" in issue:
            target = WD_ALIGN_PARAGRAPH.CENTER
        elif "right" in issue:
            target = WD_ALIGN_PARAGRAPH.RIGHT
        elif "left" in issue:
            target = WD_ALIGN_PARAGRAPH.LEFT
        elif "justify" in issue:
            target = WD_ALIGN_PARAGRAPH.JUSTIFY

        if target is None:
            return

        area = diff.get("area", "").lower()
        # Apply to paragraphs that seem to match the area description.
        for para in self._doc.paragraphs:
            if area and area in para.text.lower():
                para.alignment = target
                self._fixes_applied += 1
                return

    def _fix_spacing(self, diff: dict[str, Any]) -> None:
        """Adjust paragraph spacing."""
        issue = diff.get("issue", "").lower()
        for para in self._doc.paragraphs:
            pf = para.paragraph_format
            if "too much" in issue or "extra" in issue or "large" in issue:
                if pf.space_before and pf.space_before.pt > 2:
                    pf.space_before = Pt(pf.space_before.pt * 0.7)
                    self._fixes_applied += 1
                    return
                if pf.space_after and pf.space_after.pt > 2:
                    pf.space_after = Pt(pf.space_after.pt * 0.7)
                    self._fixes_applied += 1
                    return
            elif "too little" in issue or "missing" in issue or "small" in issue:
                current = pf.space_before.pt if pf.space_before else 0
                pf.space_before = Pt(current + 2)
                self._fixes_applied += 1
                return

    def _fix_border(self, diff: dict[str, Any]) -> None:
        """Add missing table borders."""
        for table in self._doc.tables:
            # Ensure all cells have borders.
            for row in table.rows:
                for cell in row.cells:
                    tc_pr = cell._element.get_or_add_tcPr()
                    borders = tc_pr.find(qn("w:tcBorders"))
                    if borders is None:
                        borders = etree.SubElement(tc_pr, qn("w:tcBorders"))
                        for side in ("top", "left", "bottom", "right"):
                            el = etree.SubElement(borders, qn(f"w:{side}"))
                            el.set(qn("w:val"), "single")
                            el.set(qn("w:sz"), "4")
                            el.set(qn("w:space"), "0")
                            el.set(qn("w:color"), "000000")
                        self._fixes_applied += 1
            return  # Apply to first table only.

    def _fix_shading(self, diff: dict[str, Any]) -> None:
        """Fix cell shading issues."""
        issue = diff.get("issue", "").lower()
        if "missing" in issue:
            # Can't determine the correct color without more info.
            logger.debug("Shading fix requires color info — skipping.")
        self._fixes_applied += 0  # Placeholder — needs color info.

    def _fix_font_color(self, diff: dict[str, Any]) -> None:
        """Fix font color differences."""
        # Without specific color info from the AI, we can't auto-fix.
        logger.debug("Font color fix — needs specific color value.")

    def _fix_bold(self, diff: dict[str, Any]) -> None:
        """Toggle bold formatting."""
        issue = diff.get("issue", "").lower()
        should_be_bold = "should be bold" in issue or "missing bold" in issue
        area = diff.get("area", "").lower()

        for para in self._doc.paragraphs:
            if area and area in para.text.lower():
                for run in para.runs:
                    run.font.bold = should_be_bold
                    self._fixes_applied += 1
                return

    def _fix_italic(self, diff: dict[str, Any]) -> None:
        """Toggle italic formatting."""
        issue = diff.get("issue", "").lower()
        should_be_italic = "should be italic" in issue or "missing italic" in issue
        area = diff.get("area", "").lower()

        for para in self._doc.paragraphs:
            if area and area in para.text.lower():
                for run in para.runs:
                    run.font.italic = should_be_italic
                    self._fixes_applied += 1
                return

    # Map from difference type to handler method.
    _handlers = {
        "font_size": _fix_font_size,
        "font_family": lambda self, d: None,  # Complex — skip for now.
        "font_color": _fix_font_color,
        "bold": _fix_bold,
        "italic": _fix_italic,
        "underline": lambda self, d: None,
        "alignment": _fix_alignment,
        "spacing": _fix_spacing,
        "border": _fix_border,
        "shading": _fix_shading,
        "image": lambda self, d: None,  # Needs re-extraction.
        "layout": lambda self, d: None,
        "missing_content": lambda self, d: None,
        "extra_content": lambda self, d: None,
    }
