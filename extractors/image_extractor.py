"""Image extraction from PDF pages using PyMuPDF."""

import logging
import os
from typing import Any

import fitz
from PIL import Image
import io

logger = logging.getLogger(__name__)


class ImageExtractor:
    """Extracts embedded images from PDF pages.

    Requires a ``fitz.Document`` reference because
    ``doc.extract_image(xref)`` operates at the document level.
    """

    def __init__(self, doc: fitz.Document) -> None:
        """
        Parameters
        ----------
        doc:
            The open ``fitz.Document`` that owns the pages being processed.
        """
        self._doc = doc

    @staticmethod
    def _ensure_rgb(image_bytes: bytes, ext: str) -> tuple[bytes, str]:
        """Convert CMYK or other non-RGB images to RGB PNG using Pillow.

        Returns the (possibly converted) image bytes and the final file
        extension.
        """
        try:
            img = Image.open(io.BytesIO(image_bytes))
            if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
                img = img.convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                return buf.getvalue(), "png"
            return image_bytes, ext
        except Exception:
            # If Pillow cannot process the image, return the original bytes.
            return image_bytes, ext

    def extract(
        self,
        page: fitz.Page,
        page_num: int,
        temp_dir: str,
    ) -> list[dict[str, Any]]:
        """Extract images embedded in *page* and save them to *temp_dir*.

        Parameters
        ----------
        page:
            A ``fitz.Page`` object.
        page_num:
            0-based page index, used for naming saved files.
        temp_dir:
            Directory where extracted image files are written.

        Returns
        -------
        list[dict[str, Any]]
            Each dict contains ``path``, ``bbox``, ``width``, ``height``,
            ``page_num`` and ``ext``.
        """
        results: list[dict[str, Any]] = []
        os.makedirs(temp_dir, exist_ok=True)

        try:
            image_list = page.get_images(full=True)
        except Exception:
            logger.exception("Failed to list images on page %s", page_num)
            return results

        for idx, img_info in enumerate(image_list):
            xref: int = img_info[0]
            try:
                base_image: dict[str, Any] = self._doc.extract_image(xref)
                if not base_image or "image" not in base_image:
                    logger.warning(
                        "Empty image data for xref %s on page %s – skipping",
                        xref,
                        page_num,
                    )
                    continue

                raw_bytes: bytes = base_image["image"]
                ext: str = base_image.get("ext", "png")
                width: int = base_image.get("width", 0)
                height: int = base_image.get("height", 0)

                # Ensure the colour space is RGB-compatible
                raw_bytes, ext = self._ensure_rgb(raw_bytes, ext)

                # Determine bounding box on the page
                bbox: tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)
                try:
                    rects = page.get_image_rects(xref)
                    if rects:
                        rect = rects[0]
                        bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                except Exception:
                    logger.debug(
                        "Could not determine rect for xref %s on page %s",
                        xref,
                        page_num,
                    )

                filename = f"img_page{page_num}_index{idx}.{ext}"
                save_path = os.path.join(temp_dir, filename)

                with open(save_path, "wb") as f:
                    f.write(raw_bytes)

                results.append(
                    {
                        "path": save_path,
                        "bbox": bbox,
                        # Use display dimensions (points) from the page
                        # rect when available; fall back to raw pixel dims.
                        "width": bbox[2] - bbox[0] if bbox != (0.0, 0.0, 0.0, 0.0) else width,
                        "height": bbox[3] - bbox[1] if bbox != (0.0, 0.0, 0.0, 0.0) else height,
                        "page_num": page_num,
                        "ext": ext,
                    }
                )

            except Exception:
                logger.warning(
                    "Failed to extract image xref %s on page %s – skipping",
                    xref,
                    page_num,
                    exc_info=True,
                )
                continue

        return results
