"""Flask application factory for the PDF-to-Word Web UI."""

import logging
import os
import sys

# Ensure project root is on sys.path so we can import pdf2docx, etc.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from flask import Flask, render_template  # noqa: E402

from web.services.converter import ConverterService       # noqa: E402
from web.routes.upload import upload_bp, init_upload       # noqa: E402
from web.routes.status import status_bp, init_status       # noqa: E402
from web.routes.download import download_bp, init_download # noqa: E402
from web.routes.style import style_bp, init_style          # noqa: E402


def create_app() -> Flask:
    """Create and configure the Flask application."""
    app = Flask(
        __name__,
        template_folder=os.path.join(os.path.dirname(__file__), "templates"),
        static_folder=os.path.join(os.path.dirname(__file__), "static"),
    )

    app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

    # Shared converter service.
    converter = ConverterService()

    # Wire the service into each blueprint.
    init_upload(converter)
    init_status(converter)
    init_download(converter)
    init_style(converter)

    # Register blueprints.
    app.register_blueprint(upload_bp)
    app.register_blueprint(status_bp)
    app.register_blueprint(download_bp)
    app.register_blueprint(style_bp)

    # ── Front-end routes ─────────────────────────────────────────────
    @app.route("/")
    def index():
        return render_template("index.html")

    return app


# ── Run directly: python -m web.app ──────────────────────────────────
if __name__ == "__main__":
    # Force unbuffered stdout so background-thread print() calls
    # appear immediately in the terminal.
    import functools
    sys.stdout.reconfigure(line_buffering=True)  # type: ignore[attr-defined]

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
        force=True,
    )
    application = create_app()
    print("\n  PDF-to-Word Web UI")
    print("  http://localhost:5000\n")
    application.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
