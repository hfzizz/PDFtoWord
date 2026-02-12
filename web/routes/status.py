"""Status / progress routes — SSE + JSON polling."""

import json
import time
import logging
from flask import Blueprint, Response, jsonify, stream_with_context

from web.services.converter import ConverterService

logger = logging.getLogger(__name__)

status_bp = Blueprint("status", __name__)

_converter: ConverterService | None = None


def init_status(converter: ConverterService) -> None:
    global _converter
    _converter = converter


@status_bp.route("/api/status/<job_id>")
def job_status_json(job_id: str):
    """Return current job state as JSON (for polling)."""
    assert _converter is not None
    job = _converter.get_job(job_id)
    if job is None:
        return jsonify({"error": "Job not found."}), 404
    return jsonify(job.to_dict())


@status_bp.route("/api/status/<job_id>/stream")
def job_status_sse(job_id: str):
    """Stream job progress as Server-Sent Events."""
    assert _converter is not None

    def generate():
        last_sent = ""
        while True:
            job = _converter.get_job(job_id)
            if job is None:
                yield "data: {\"error\": \"Job not found.\"}\n\n"
                break

            payload = json.dumps(job.to_dict())
            # Only send if state changed.
            if payload != last_sent:
                yield f"data: {payload}\n\n"
                last_sent = payload

            if job.status in ("complete", "failed"):
                break

            time.sleep(0.5)

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@status_bp.route("/api/history")
def history():
    """Return all jobs (most recent first)."""
    assert _converter is not None
    return jsonify(_converter.list_jobs())


@status_bp.route("/api/job/<job_id>", methods=["DELETE"])
def delete_job(job_id: str):
    """Delete a job and its associated files."""
    assert _converter is not None
    if _converter.delete_job(job_id):
        return jsonify({"deleted": True})
    return jsonify({"error": "Job not found."}), 404
