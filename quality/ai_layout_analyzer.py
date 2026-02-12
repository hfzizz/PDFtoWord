"""AI-powered layout analysis using Google Gemini.

Sends rendered PDF page images to the Gemini API *before* DOCX
construction so that the builder can apply formatting overrides
(font sizes, colors, alignment, table styles, etc.) that more
faithfully reproduce the original PDF layout.
"""

import json
import logging
import os
from typing import Any

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Prompt sent alongside each page image.
# ---------------------------------------------------------------------------
_LAYOUT_PROMPT = """\
You are a document layout analysis expert.  I am showing you a single page
from a PDF document.  Your task is to identify every visible text element
and every table on this page and report their formatting details so that
the page can be faithfully reconstructed in a Word document.

Return ONLY valid JSON with three top-level keys:

1. "text_elements" — an array of objects, one per distinct text block or
   paragraph you can see.  Each object must have:
   - "text_snippet": the first ~40 characters of the text (verbatim).
   - "font_size_pt": estimated font size in points (number).
   - "font_color": hex colour string, e.g. "#000000".
   - "bold": true or false.
   - "italic": true or false.
   - "alignment": exactly one of "left", "center", "right", "justify".
   - "background_color": hex colour string or null if no highlight/shading.

2. "table_styles" — an array of objects, one per table on the page.
   Each object must have:
   - "table_index": 0-based index of the table on this page.
   - "header_bg_color": hex colour of the header row background, or null.
   - "border_style": exactly one of "thin", "medium", "thick", "none".
   - "cell_padding": exactly one of "tight", "normal", "wide".

3. "page_style" — a single object with:
   - "dominant_font": the most-used font family name (e.g. "Arial").
   - "dominant_size": the most-used font size in points (number).

Guidelines:
- Font sizes must be in points.  Colours must be hex (e.g. "#4472C4").
- Focus on accuracy over completeness — only report what you are
  confident about.  It is better to omit an uncertain element than to
  guess incorrectly.
- If there are no tables, return an empty array for "table_styles".
- Do NOT wrap the JSON in markdown code fences.
"""


class AILayoutAnalyzer:
    """Analyze PDF page layout using Google Gemini vision.

    The returned overrides dict can be fed directly into ``DocxBuilder``
    to improve conversion fidelity.

    Parameters
    ----------
    api_key : str | None
        Google API key.  Falls back to ``GEMINI_API_KEY`` env var.
    model : str
        Gemini model name (default ``gemini-2.0-flash``).
    max_tokens : int
        Maximum tokens per API call.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str = "gemini-2.0-flash",
        max_tokens: int = 8192,
    ) -> None:
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY", "")
        self.model_name = model
        self.max_tokens = max_tokens
        self._client: Any = None
        self._total_tokens = 0

    @property
    def is_available(self) -> bool:
        """Check whether the Gemini API is configured and usable."""
        return bool(self.api_key)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def analyze_layout(
        self,
        pdf_path: str,
        structure_map: list[dict],
        progress_callback: Any = None,
    ) -> dict:
        """Analyze all pages of a PDF and return formatting overrides.

        Parameters
        ----------
        pdf_path : str
            Path to the source PDF file.
        structure_map : list[dict]
            The extracted structure map (not mutated — used for context
            only if needed in future enhancements).
        progress_callback : callable | None
            Optional ``(current, total, message)`` callback for progress
            reporting.

        Returns
        -------
        dict
            ``{"text_overrides": {...}, "table_overrides": [...],
            "page_styles": {...}}``
        """
        empty_result: dict[str, Any] = {
            "text_overrides": {},
            "table_overrides": [],
            "page_styles": {},
        }

        if not self.is_available:
            logger.warning(
                "Gemini API key not set — skipping AI layout analysis."
            )
            return empty_result

        if not os.path.isfile(pdf_path):
            logger.error("PDF file not found: %s", pdf_path)
            return empty_result

        try:
            self._init_client()
        except Exception as e:
            logger.error("Failed to initialize Gemini client: %s", e)
            return empty_result

        # Render pages to PIL images via PyMuPDF.
        try:
            page_images = self._render_pages(pdf_path)
        except Exception as e:
            logger.error("Failed to render PDF pages: %s", e)
            return empty_result

        if not page_images:
            logger.warning("No pages rendered from %s", pdf_path)
            return empty_result

        combined: dict[str, Any] = {
            "text_overrides": {},
            "table_overrides": [],
            "page_styles": {},
        }

        num_pages = len(page_images)
        for page_num, img in enumerate(page_images):
            logger.info(
                "AI layout analysis — page %d / %d", page_num + 1, num_pages
            )
            if progress_callback:
                try:
                    progress_callback(
                        page_num,
                        num_pages,
                        f"AI analyzing page {page_num + 1}/{num_pages}",
                    )
                except Exception:
                    pass

            page_result = self._analyze_single_page(img, page_num)
            self._merge_page_result(combined, page_result, page_num)

        logger.info(
            "AI layout analysis complete — %d page(s), %d total token(s) used.",
            num_pages,
            self._total_tokens,
        )
        return combined

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------

    def _init_client(self) -> None:
        """Lazily initialise the Gemini client."""
        if self._client is not None:
            return
        try:
            from google import genai

            self._client = genai.Client(api_key=self.api_key)
        except ImportError:
            logger.error(
                "google-genai package not installed.  "
                "Run: pip install google-genai"
            )
            raise
        except Exception:
            logger.exception("Failed to initialise Gemini client.")
            raise

    @staticmethod
    def _render_pages(pdf_path: str) -> list:
        """Render every page of *pdf_path* at 150 DPI to PIL Images."""
        import fitz  # PyMuPDF
        from PIL import Image
        import io

        images: list = []
        doc = fitz.open(pdf_path)
        try:
            zoom = 150 / 72  # 150 DPI / default 72 DPI
            matrix = fitz.Matrix(zoom, zoom)
            for page in doc:
                pix = page.get_pixmap(matrix=matrix)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                images.append(img)
        finally:
            doc.close()
        return images

    def _analyze_single_page(
        self, image: Any, page_num: int
    ) -> dict[str, Any]:
        """Send a single page image to Gemini and return parsed result."""
        try:
            from google.genai import types

            response = self._client.models.generate_content(
                model=self.model_name,
                contents=[_LAYOUT_PROMPT, image],
                config=types.GenerateContentConfig(
                    max_output_tokens=self.max_tokens,
                    temperature=0.1,
                ),
            )
        except Exception as e:
            logger.error(
                "Gemini API call failed for page %d: %s", page_num, e
            )
            return {}

        # Track token usage.
        if hasattr(response, "usage_metadata"):
            usage = response.usage_metadata
            tokens = getattr(usage, "total_token_count", 0)
            self._total_tokens += tokens
            logger.debug("Page %d used %d tokens.", page_num, tokens)

        text = response.text or ""
        if not text:
            logger.warning(
                "Empty response from Gemini for page %d.", page_num
            )
            return {}

        return self._parse_response(text, page_num)

    # ------------------------------------------------------------------
    # Response parsing
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_response(text: str, page_num: int) -> dict[str, Any]:
        """Parse the Gemini response text into a structured dict.

        Returns a dict with ``text_elements``, ``table_styles``, and
        ``page_style`` keys (all validated).  Returns an empty dict on
        parse failure.
        """
        cleaned = text.strip()
        # Strip markdown code fences if present.
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")
            lines = [l for l in lines if not l.strip().startswith("```")]
            cleaned = "\n".join(lines)

        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError:
            logger.warning(
                "Could not parse AI response as JSON for page %d:\n%s",
                page_num,
                text[:500],
            )
            return {}

        if not isinstance(parsed, dict):
            logger.warning(
                "Unexpected response type for page %d: %s",
                page_num,
                type(parsed).__name__,
            )
            return {}

        result: dict[str, Any] = {
            "text_elements": [],
            "table_styles": [],
            "page_style": {},
        }

        # --- text_elements ---
        raw_elements = parsed.get("text_elements")
        if isinstance(raw_elements, list):
            for item in raw_elements:
                if not isinstance(item, dict):
                    continue
                snippet = item.get("text_snippet")
                if not snippet or not isinstance(snippet, str):
                    continue
                result["text_elements"].append(
                    {
                        "text_snippet": snippet,
                        "font_size_pt": _to_float(item.get("font_size_pt")),
                        "font_color": _to_color(item.get("font_color")),
                        "bold": bool(item.get("bold", False)),
                        "italic": bool(item.get("italic", False)),
                        "alignment": _to_alignment(item.get("alignment")),
                        "background_color": _to_color(
                            item.get("background_color")
                        ),
                    }
                )

        # --- table_styles ---
        raw_tables = parsed.get("table_styles")
        if isinstance(raw_tables, list):
            for item in raw_tables:
                if not isinstance(item, dict):
                    continue
                result["table_styles"].append(
                    {
                        "table_index": int(item.get("table_index", 0)),
                        "header_bg_color": _to_color(
                            item.get("header_bg_color")
                        ),
                        "border_style": _to_border(item.get("border_style")),
                        "cell_padding": _to_padding(
                            item.get("cell_padding")
                        ),
                    }
                )

        # --- page_style ---
        raw_style = parsed.get("page_style")
        if isinstance(raw_style, dict):
            result["page_style"] = {
                "dominant_font": str(
                    raw_style.get("dominant_font", "")
                )
                or None,
                "dominant_size": _to_float(
                    raw_style.get("dominant_size")
                ),
            }

        return result

    # ------------------------------------------------------------------
    # Merging per-page results into the combined dict
    # ------------------------------------------------------------------

    @staticmethod
    def _merge_page_result(
        combined: dict[str, Any],
        page_result: dict[str, Any],
        page_num: int,
    ) -> None:
        """Merge a single-page parse result into *combined* in-place."""
        if not page_result:
            return

        # Text overrides — keyed by normalised snippet.
        for elem in page_result.get("text_elements", []):
            key = elem["text_snippet"].strip().lower()
            if not key:
                continue
            combined["text_overrides"][key] = {
                "font_size_pt": elem.get("font_size_pt"),
                "font_color": elem.get("font_color"),
                "bold": elem.get("bold", False),
                "italic": elem.get("italic", False),
                "alignment": elem.get("alignment"),
                "background_color": elem.get("background_color"),
            }

        # Table overrides — append with page number.
        for tbl in page_result.get("table_styles", []):
            combined["table_overrides"].append(
                {
                    "page_num": page_num,
                    "table_index": tbl.get("table_index", 0),
                    "header_bg_color": tbl.get("header_bg_color"),
                    "border_style": tbl.get("border_style"),
                }
            )

        # Page styles.
        page_style = page_result.get("page_style")
        if page_style:
            combined["page_styles"][page_num] = page_style


# ---------------------------------------------------------------------------
# Small validation helpers
# ---------------------------------------------------------------------------

def _to_float(value: Any) -> float | None:
    """Coerce *value* to ``float`` or return ``None``."""
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _to_color(value: Any) -> str | None:
    """Return a sanitised hex colour string or ``None``."""
    if not value or not isinstance(value, str):
        return None
    value = value.strip()
    if value.startswith("#") and len(value) in (4, 7, 9):
        return value.upper()
    return None


def _to_alignment(value: Any) -> str | None:
    """Validate alignment value."""
    valid = {"left", "center", "right", "justify"}
    if isinstance(value, str) and value.strip().lower() in valid:
        return value.strip().lower()
    return None


def _to_border(value: Any) -> str:
    """Validate border style value."""
    valid = {"thin", "medium", "thick", "none"}
    if isinstance(value, str) and value.strip().lower() in valid:
        return value.strip().lower()
    return "thin"


def _to_padding(value: Any) -> str:
    """Validate cell padding value."""
    valid = {"tight", "normal", "wide"}
    if isinstance(value, str) and value.strip().lower() in valid:
        return value.strip().lower()
    return "normal"
