"""Converter service — wraps pdf2docx.convert_pdf() for the web UI.

Runs conversions in background threads and exposes per-job progress
that the SSE endpoint can stream to the browser.
"""

import logging
import os
import sys
import threading
import time
import uuid
from datetime import datetime, timezone
from typing import Any

# Ensure the project root is importable.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from pdf2docx import convert_pdf, _load_config          # noqa: E402
from utils.pdf_info import PDFInfo                       # noqa: E402
from utils.validator import OutputValidator              # noqa: E402

logger = logging.getLogger(__name__)


class ConversionJob:
    """Tracks state for a single PDF → DOCX conversion."""

    def __init__(
        self,
        job_id: str,
        input_path: str,
        output_path: str,
        original_filename: str,
        settings: dict[str, Any] | None = None,
    ) -> None:
        self.id = job_id
        self.input_path = input_path
        self.output_path = output_path
        self.original_filename = original_filename
        self.settings: dict[str, Any] = settings or {}

        # Progress state
        self.status: str = "queued"        # queued → running → complete | failed
        self.stage: str = ""               # analyzing, extracting, building, …
        self.progress: int = 0             # 0-100
        self.page_count: int = 0
        self.current_page: int = 0
        self.message: str = ""
        self.error: str | None = None

        # Results
        self.quality_report: dict[str, Any] | None = None
        self.created_at: str = datetime.now(timezone.utc).isoformat()
        self.completed_at: str | None = None

    # ── Serialisation ────────────────────────────────────────────────
    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "original_filename": self.original_filename,
            "status": self.status,
            "stage": self.stage,
            "progress": self.progress,
            "page_count": self.page_count,
            "current_page": self.current_page,
            "message": self.message,
            "error": self.error,
            "quality_report": self.quality_report,
            "created_at": self.created_at,
            "completed_at": self.completed_at,
        }


class ConverterService:
    """Manages conversion jobs — creation, execution, progress."""

    def __init__(self) -> None:
        self._jobs: dict[str, ConversionJob] = {}
        self._lock = threading.Lock()

    # ── Public API ───────────────────────────────────────────────────

    def create_job(
        self,
        input_path: str,
        output_path: str,
        original_filename: str,
        settings: dict[str, Any] | None = None,
    ) -> ConversionJob:
        """Create a new (queued) conversion job."""
        job_id = uuid.uuid4().hex[:12]
        job = ConversionJob(job_id, input_path, output_path,
                            original_filename, settings)
        with self._lock:
            self._jobs[job_id] = job

        # Pre-analyse to get page count (fast).
        try:
            pdf_info = PDFInfo()
            info = pdf_info.analyze(input_path)
            job.page_count = info.get("page_count", 0)
        except Exception:
            pass

        return job

    def start_job(self, job_id: str) -> None:
        """Launch background thread for the given job."""
        job = self.get_job(job_id)
        if job is None:
            raise ValueError(f"Unknown job: {job_id}")
        thread = threading.Thread(
            target=self._run_conversion, args=(job,), daemon=True
        )
        thread.start()

    def get_job(self, job_id: str) -> ConversionJob | None:
        with self._lock:
            return self._jobs.get(job_id)

    def list_jobs(self) -> list[dict[str, Any]]:
        with self._lock:
            return [j.to_dict() for j in reversed(list(self._jobs.values()))]

    def delete_job(self, job_id: str) -> bool:
        with self._lock:
            job = self._jobs.pop(job_id, None)
        if job is None:
            return False
        # Clean up files.
        for p in (job.input_path, job.output_path):
            try:
                if p and os.path.isfile(p):
                    os.remove(p)
            except OSError:
                pass
        return True

    # ── Internal ─────────────────────────────────────────────────────

    def _run_conversion(self, job: ConversionJob) -> None:
        """Execute the conversion pipeline (runs in a worker thread)."""
        job.status = "running"
        job.message = "Starting conversion…"

        # Build config from job settings.
        config = _load_config()
        settings = job.settings or {}
        if settings.get("ocr"):
            config["ocr_enabled"] = True
        if settings.get("skip_watermarks"):
            config["skip_watermarks"] = True
        config["verbose"] = False

        password = settings.get("password") or None
        
        # AI comparison settings
        ai_compare = settings.get("ai_compare", False)
        gemini_api_key = settings.get("gemini_api_key") or None
        
        # Store API key in config if provided (for this conversion only)
        if ai_compare and gemini_api_key:
            if "ai_comparison" not in config:
                config["ai_comparison"] = {}
            config["ai_comparison"]["api_key"] = gemini_api_key

        def _progress_callback(
            stage: str, current: int, total: int, message: str
        ) -> None:
            job.stage = stage
            job.current_page = current
            if total > 0:
                job.page_count = total
            job.message = message
            # Map to 0-100 percentage.
            if stage == "extracting" and total:
                job.progress = int(current / total * 60)   # 0-60 %
            elif stage == "analyzing":
                job.progress = 65
            elif stage == "building":
                job.progress = 70
            elif stage == "validating":
                job.progress = 75
            elif stage == "visual_diff":
                job.progress = 80
            elif stage == "ai_compare":
                # AI comparison rounds: 80-95%
                base = 80
                if total > 0:
                    job.progress = base + int((current / total) * 15)
                else:
                    job.progress = base
            elif stage == "auto_fix":
                job.progress = 95
            elif stage == "complete":
                job.progress = 100

        try:
            convert_pdf(
                job.input_path,
                job.output_path,
                config,
                password=password,
                validate=False,
                visual_validate=ai_compare,  # Visual diff needed for AI compare
                ai_compare=ai_compare,
                progress_callback=_progress_callback,
            )

            # Always run quality scoring for the web report.
            try:
                pdf_info = PDFInfo()
                info = pdf_info.analyze(job.input_path)
                validator = OutputValidator()
                job.quality_report = validator.quality_score(
                    job.output_path, info
                )
            except Exception as ve:
                logger.warning("Validation failed: %s", ve)

            job.status = "complete"
            job.progress = 100
            job.message = "Conversion complete"
            job.completed_at = datetime.now(timezone.utc).isoformat()

        except SystemExit:
            job.status = "failed"
            job.error = "Conversion failed (PDF may be encrypted or corrupt)."
            job.message = job.error

        except Exception as exc:
            logger.exception("Conversion job %s failed", job.id)
            job.status = "failed"
            job.error = str(exc)
            job.message = f"Error: {exc}"
