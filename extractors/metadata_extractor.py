"""Document-level metadata extraction using PyMuPDF."""

import logging
from typing import Any

import fitz

logger = logging.getLogger(__name__)


class MetadataExtractor:
    """Extracts document metadata, per-page dimensions, TOC and encryption
    status from a ``fitz.Document``.
    """

    @staticmethod
    def _is_landscape(width: float, height: float, rotation: int) -> bool:
        """Determine whether the effective page orientation is landscape.

        After a 90° or 270° rotation the visual width and height are
        swapped, so we account for that before comparing.
        """
        if rotation in (90, 270):
            width, height = height, width
        return width > height

    def extract(self, doc: fitz.Document) -> dict[str, Any]:
        """Return a comprehensive metadata dict for *doc*.

        Parameters
        ----------
        doc:
            An open ``fitz.Document``.

        Returns
        -------
        dict[str, Any]
            Keys: ``page_count``, ``metadata``, ``pages``,
            ``pdf_version``, ``is_encrypted``, ``has_toc``, ``toc``.
        """
        try:
            metadata: dict[str, Any] = doc.metadata or {}
            page_count: int = doc.page_count

            # Table of contents
            try:
                toc: list[list[Any]] = doc.get_toc()
            except Exception:
                logger.debug("Could not retrieve TOC")
                toc = []

            pages: list[dict[str, Any]] = []
            for page_num in range(page_count):
                try:
                    page: fitz.Page = doc.load_page(page_num)
                    rect = page.rect
                    width = float(rect.width)
                    height = float(rect.height)
                    rotation = page.rotation

                    pages.append(
                        {
                            "page_num": page_num,
                            "width": width,
                            "height": height,
                            "rotation": rotation,
                            "is_landscape": self._is_landscape(width, height, rotation),
                        }
                    )
                except Exception:
                    logger.warning(
                        "Failed to read page %s metadata – skipping", page_num,
                        exc_info=True,
                    )
                    continue

            pdf_version: str = metadata.get("format", "")

            return {
                "page_count": page_count,
                "metadata": metadata,
                "pages": pages,
                "pdf_version": pdf_version,
                "is_encrypted": doc.is_encrypted,
                "has_toc": len(toc) > 0,
                "toc": toc,
            }

        except Exception:
            logger.exception("Metadata extraction failed")
            return {
                "page_count": 0,
                "metadata": {},
                "pages": [],
                "pdf_version": "",
                "is_encrypted": False,
                "has_toc": False,
                "toc": [],
            }
