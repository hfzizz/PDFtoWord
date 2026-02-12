"""File manager — handles uploads, temp directories, and cleanup."""

import os
import shutil
import tempfile
from typing import Optional

# Persistent uploads directory (inside the project).
_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(_BASE, "uploads")
OUTPUT_DIR = os.path.join(_BASE, "outputs")

# Ensure directories exist.
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Maximum upload size in bytes (50 MB).
MAX_UPLOAD_SIZE = 50 * 1024 * 1024

ALLOWED_EXTENSIONS = {".pdf"}


def allowed_file(filename: str) -> bool:
    """Check if the filename has a valid extension."""
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS


def save_upload(file_storage, job_id: str) -> str:
    """Save an uploaded file and return its absolute path.

    ``file_storage`` is a Werkzeug ``FileStorage`` object.
    """
    original = file_storage.filename or "upload.pdf"
    safe_name = f"{job_id}_{_safe_filename(original)}"
    dest = os.path.join(UPLOAD_DIR, safe_name)
    file_storage.save(dest)
    return dest


def output_path_for(job_id: str, original_filename: str) -> str:
    """Return the output path for a converted DOCX."""
    base = os.path.splitext(_safe_filename(original_filename))[0]
    return os.path.join(OUTPUT_DIR, f"{job_id}_{base}.docx")


def cleanup_job(job_id: str) -> None:
    """Remove upload and output files for a given job_id."""
    for directory in (UPLOAD_DIR, OUTPUT_DIR):
        for fname in os.listdir(directory):
            if fname.startswith(job_id):
                try:
                    os.remove(os.path.join(directory, fname))
                except OSError:
                    pass


def _safe_filename(filename: str) -> str:
    """Sanitise a filename to avoid directory traversal."""
    # Keep only the basename and replace potentially dangerous chars.
    name = os.path.basename(filename)
    name = name.replace(" ", "_")
    # Remove anything that isn't alphanumeric, dash, underscore, or dot.
    return "".join(c for c in name if c.isalnum() or c in ("_", "-", "."))
