"""Style editing routes (prompt-driven, content-safe)."""

import logging
import os

from flask import Blueprint, jsonify, request

from utils.pdf_info import PDFInfo
from utils.validator import OutputValidator
from web.services.converter import ConverterService
from web.services.style_editor import StyleEditor

logger = logging.getLogger(__name__)

style_bp = Blueprint("style", __name__)

_converter: ConverterService | None = None


def init_style(converter: ConverterService) -> None:
    global _converter
    _converter = converter


@style_bp.route("/api/style/<job_id>/apply", methods=["POST"])
def apply_style(job_id: str):
    assert _converter is not None
    job = _converter.get_job(job_id)
    if job is None:
        return jsonify({"error": "Job not found."}), 404
    if job.status != "complete":
        return jsonify({"error": "Conversion not finished yet."}), 409
    if not os.path.isfile(job.output_path):
        return jsonify({"error": "Output file missing."}), 404

    payload = request.get_json(silent=True) or {}
    prompt = str(payload.get("prompt", "")).strip()
    if not prompt:
        return jsonify({"error": "Prompt is required."}), 400

    api_key = str(payload.get("gemini_api_key", "")).strip() or None
    model = str(payload.get("model", "gemini-2.0-flash")).strip() or "gemini-2.0-flash"

    editor = StyleEditor()
    try:
        result = editor.apply_prompt(
            job.output_path,
            prompt=prompt,
            api_key=api_key,
            model=model,
        )
    except Exception as exc:
        logger.exception("Style apply failed for job %s", job_id)
        return jsonify({"error": str(exc)}), 500

    # Recompute quality report after restyling.
    try:
        info = PDFInfo().analyze(job.input_path)
        job.quality_report = OutputValidator().quality_score(job.output_path, info)
    except Exception as exc:
        logger.warning("Quality update after style edit failed: %s", exc)

    return jsonify(
        {
            "ok": True,
            "changed": result.changed,
            "summary": result.summary,
            "rules": result.rules,
            "download_url": f"/api/download/{job_id}",
            "quality_report": job.quality_report,
        }
    )
