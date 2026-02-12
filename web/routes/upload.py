"""Upload & conversion routes."""

import json
import logging
from flask import Blueprint, request, jsonify

from web.services.converter import ConverterService
from web.services.file_manager import (
    allowed_file,
    save_upload,
    output_path_for,
    MAX_UPLOAD_SIZE,
)

logger = logging.getLogger(__name__)

upload_bp = Blueprint("upload", __name__)

# The ConverterService instance is injected by the app factory.
_converter: ConverterService | None = None


def init_upload(converter: ConverterService) -> None:
    """Wire the shared ConverterService into this blueprint."""
    global _converter
    _converter = converter


@upload_bp.route("/api/upload", methods=["POST"])
def upload_and_convert():
    """Accept a PDF upload, create a job, and start conversion.

    Expects ``multipart/form-data`` with:
      - ``file``: the PDF file
      - ``settings`` (optional JSON): ``{"ocr": bool, "password": str, …}``

    Returns JSON: ``{"job_id": "…", "filename": "…"}``.
    """
    assert _converter is not None, "ConverterService not initialised"

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"error": "Empty file."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only PDF files are supported."}), 400

    # Optional settings JSON.
    raw_settings = request.form.get("settings", "{}")
    try:
        settings = json.loads(raw_settings)
    except (json.JSONDecodeError, TypeError):
        settings = {}

    # Create job (gets an id).
    original = file.filename
    job = _converter.create_job(
        input_path="",           # will be set after save
        output_path="",
        original_filename=original,
        settings=settings,
    )

    # Save the uploaded file.
    input_path = save_upload(file, job.id)
    output_path = output_path_for(job.id, original)
    job.input_path = input_path
    job.output_path = output_path

    # Start background conversion.
    _converter.start_job(job.id)

    return jsonify({"job_id": job.id, "filename": original}), 202


@upload_bp.route("/api/batch-upload", methods=["POST"])
def batch_upload():
    """Accept multiple PDF uploads at once.

    Returns JSON list of ``{"job_id", "filename"}`` objects.
    """
    assert _converter is not None

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded."}), 400

    raw_settings = request.form.get("settings", "{}")
    try:
        settings = json.loads(raw_settings)
    except (json.JSONDecodeError, TypeError):
        settings = {}

    results = []
    for file in files:
        if not file or not file.filename or not allowed_file(file.filename):
            continue
        original = file.filename
        job = _converter.create_job("", "", original, settings)
        input_path = save_upload(file, job.id)
        output_path = output_path_for(job.id, original)
        job.input_path = input_path
        job.output_path = output_path
        _converter.start_job(job.id)
        results.append({"job_id": job.id, "filename": original})

    if not results:
        return jsonify({"error": "No valid PDF files found."}), 400

    return jsonify(results), 202
