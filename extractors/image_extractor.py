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
    def _flatten_to_rgb(img: Image.Image) -> Image.Image:
        """Flatten any image to RGB, compositing alpha onto white.

        Handles RGBA, LA, PA, P (with transparency), and CMYK modes.
        Returns an RGB ``Image`` ready for saving.
        """
        # Palette images: convert first (may gain an alpha channel).
        if img.mode == "P":
            img = img.convert("RGBA")

        # Composite images with alpha onto a white background.
        if img.mode in ("RGBA", "LA", "PA"):
            if img.mode != "RGBA":
                img = img.convert("RGBA")
            background = Image.new("RGBA", img.size, (255, 255, 255, 255))
            background.paste(img, mask=img.split()[-1])
            return background.convert("RGB")

        # CMYK / other exotic modes → RGB.
        if img.mode != "RGB":
            return img.convert("RGB")

        return img

    @classmethod
    def _ensure_rgb(cls, image_bytes: bytes, ext: str) -> tuple[bytes, str]:
        """Convert image bytes to RGB PNG, compositing transparency onto white.

        Returns the (possibly converted) image bytes and the final file
        extension (always ``"png"``).
        """
        try:
            img = Image.open(io.BytesIO(image_bytes))
            img = cls._flatten_to_rgb(img)
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            return buf.getvalue(), "png"
        except Exception:
            # If Pillow cannot process the image, return the original bytes.
            return image_bytes, ext

    def _apply_smask(self, xref: int, smask_xref: int,
                     raw_bytes: bytes) -> bytes:
        """Apply a PDF Soft Mask (SMask) as an alpha channel.

        Parameters
        ----------
        xref : int
            Xref of the base image.
        smask_xref : int
            Xref of the soft-mask image.
        raw_bytes : bytes
            Raw bytes of the base image (without mask applied).

        Returns
        -------
        bytes
            PNG bytes of the base image with the SMask applied as alpha
            and composited onto a white background.
        """
        try:
            base_img = Image.open(io.BytesIO(raw_bytes))
            if base_img.mode not in ("RGB", "RGBA", "L"):
                base_img = base_img.convert("RGB")
            if base_img.mode == "L":
                base_img = base_img.convert("RGB")

            # Read the soft mask from the PDF.
            mask_data = self._doc.extract_image(smask_xref)
            if mask_data and "image" in mask_data:
                mask_img = Image.open(io.BytesIO(mask_data["image"]))
                # Ensure mask is grayscale and same size as base.
                mask_img = mask_img.convert("L")
                if mask_img.size != base_img.size:
                    mask_img = mask_img.resize(base_img.size, Image.LANCZOS)

                # Apply mask as alpha channel.
                if base_img.mode == "RGBA":
                    # Replace existing alpha with SMask.
                    r, g, b, _ = base_img.split()
                    base_img = Image.merge("RGBA", (r, g, b, mask_img))
                else:
                    base_img.putalpha(mask_img)

                # Composite onto white.
                base_img = self._flatten_to_rgb(base_img)

            buf = io.BytesIO()
            base_img.save(buf, format="PNG")
            return buf.getvalue()
        except Exception:
            logger.debug(
                "Failed to apply SMask (xref=%s, smask=%s) — using original.",
                xref, smask_xref, exc_info=True,
            )
            return raw_bytes

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
            smask_xref: int = img_info[1] if len(img_info) > 1 else 0
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

                # Apply soft mask (SMask) if present — this is how PDFs
                # store transparency for images like logos.
                if smask_xref and smask_xref > 0:
                    logger.debug(
                        "Applying SMask (xref=%s) to image xref=%s on page %s",
                        smask_xref, xref, page_num,
                    )
                    raw_bytes = self._apply_smask(xref, smask_xref, raw_bytes)
                    ext = "png"

                # Ensure the colour space is RGB and flatten any remaining
                # transparency onto a white background.
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
