"""AI vision comparison using Google Gemini.

Sends rendered PDF and DOCX page images to the Gemini API for
intelligent visual comparison.  Returns structured differences that
the correction engine can act on.
"""

import json
import logging
import os
from typing import Any

logger = logging.getLogger(__name__)

# Default comparison prompt sent alongside the page images.
_COMPARISON_PROMPT = """\
You are a document layout comparison expert.  I am showing you two images
of the same document page.

**Image 1** is the ORIGINAL PDF page.
**Image 2** is a CONVERTED Word (DOCX) page.

Your task: identify every visual difference between them.

Return ONLY a JSON array.  Each element must be an object with these keys:
- "area": short description of the location (e.g. "header", "row 3 col 2", "bottom paragraph")
- "issue": what is different (e.g. "missing cell border", "font size smaller", "image shifted right")
- "type": exactly one of: "font_size", "font_family", "font_color", "bold", "italic",
  "underline", "alignment", "spacing", "border", "shading", "image", "layout", "missing_content", "extra_content"
- "severity": "high", "medium", or "low"

If the two images look identical, return an empty array: []

Focus on: fonts, colors, spacing, text alignment, borders, cell shading,
images (position & size), and overall layout.  Ignore very minor
anti-aliasing or sub-pixel rendering artifacts.
"""


class AIComparator:
    """Compare PDF and DOCX page images using Google Gemini vision.

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

    def compare_pages(
        self,
        pdf_images: list[str],
        docx_images: list[str],
    ) -> list[list[dict[str, Any]]]:
        """Compare each pair of rendered page images.

        Parameters
        ----------
        pdf_images : list[str]
            Paths to rendered PDF page PNGs.
        docx_images : list[str]
            Paths to rendered DOCX page PNGs.

        Returns
        -------
        list[list[dict]]
            Outer list = pages.  Inner list = differences found on
            that page.  Each difference is a dict with ``area``,
            ``issue``, ``type``, and ``severity`` keys.
        """
        # Input validation
        if not pdf_images or not docx_images:
            logger.warning("Empty image lists provided — no comparison possible.")
            return []
        
        if not self.is_available:
            logger.warning("Gemini API key not set — skipping AI comparison.")
            return [[] for _ in pdf_images]

        # Validate image files exist
        import os
        for i, img_path in enumerate(pdf_images):
            if not os.path.isfile(img_path):
                logger.error("PDF image not found: %s (page %d)", img_path, i)
                return [[] for _ in pdf_images]
        
        for i, img_path in enumerate(docx_images):
            if not os.path.isfile(img_path):
                logger.error("DOCX image not found: %s (page %d)", img_path, i)
                return [[] for _ in pdf_images]

        try:
            self._init_client()
        except Exception as e:
            logger.error("Failed to initialize Gemini client: %s", e)
            return [[] for _ in pdf_images]

        results: list[list[dict[str, Any]]] = []
        num_pages = min(len(pdf_images), len(docx_images))

        for i in range(num_pages):
            logger.info("AI comparison — page %d / %d", i + 1, num_pages)
            diffs = self._compare_single_page(pdf_images[i], docx_images[i], i)
            results.append(diffs)

        logger.info(
            "AI comparison complete — %d page(s), %d total token(s) used.",
            num_pages,
            self._total_tokens,
        )
        return results

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
                "google-generativeai package not installed.  "
                "Run: pip install google-generativeai"
            )
            raise
        except Exception:
            logger.exception("Failed to initialise Gemini client.")
            raise

    def _compare_single_page(
        self,
        pdf_img_path: str,
        docx_img_path: str,
        page_num: int,
    ) -> list[dict[str, Any]]:
        """Send one pair of page images to Gemini and parse the response."""
        try:
            from google.genai import types
            from PIL import Image

            # Validate images can be opened
            try:
                img_pdf = Image.open(pdf_img_path)
            except Exception as e:
                logger.error("Failed to open PDF image %s: %s", pdf_img_path, e)
                return []
            
            try:
                img_docx = Image.open(docx_img_path)
            except Exception as e:
                logger.error("Failed to open DOCX image %s: %s", docx_img_path, e)
                return []

            # Make API call with timeout protection
            try:
                response = self._client.models.generate_content(
                    model=self.model_name,
                    contents=[
                        _COMPARISON_PROMPT,
                        img_pdf,
                        img_docx,
                    ],
                    config=types.GenerateContentConfig(
                        max_output_tokens=self.max_tokens,
                        temperature=0.1,
                    ),
                )
            except Exception as e:
                logger.error("Gemini API call failed for page %d: %s", page_num, e)
                return []

            # Track token usage.
            if hasattr(response, "usage_metadata"):
                usage = response.usage_metadata
                tokens = getattr(usage, "total_token_count", 0)
                self._total_tokens += tokens
                logger.debug("Page %d used %d tokens.", page_num, tokens)

            # Parse JSON from the response text.
            text = response.text or ""
            if not text:
                logger.warning("Empty response from Gemini for page %d.", page_num)
                return []
            
            return self._parse_response(text, page_num)

        except ImportError as e:
            logger.error(
                "Required package not available: %s. "
                "Run: pip install google-generativeai Pillow",
                e
            )
            return []
        except Exception as e:
            logger.exception("AI comparison failed for page %d: %s", page_num, e)
            return []

    @staticmethod
    def _parse_response(
        text: str, page_num: int
    ) -> list[dict[str, Any]]:
        """Parse the Gemini response text into a list of difference dicts."""
        # Strip markdown code fences if present.
        cleaned = text.strip()
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")
            # Remove first and last lines (fences).
            lines = [l for l in lines if not l.strip().startswith("```")]
            cleaned = "\n".join(lines)

        try:
            parsed = json.loads(cleaned)
            if isinstance(parsed, list):
                # Validate each entry.
                valid: list[dict[str, Any]] = []
                for item in parsed:
                    if isinstance(item, dict) and "issue" in item:
                        item.setdefault("area", "unknown")
                        item.setdefault("type", "layout")
                        item.setdefault("severity", "medium")
                        item["page_num"] = page_num
                        valid.append(item)
                return valid
        except json.JSONDecodeError:
            logger.warning(
                "Could not parse AI response as JSON for page %d:\n%s",
                page_num,
                text[:500],
            )
        return []
