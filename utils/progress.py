"""Progress tracking module.

Wraps tqdm for progress-bar display with a simple print-based fallback.
"""

import logging
import sys
from typing import Optional

logger = logging.getLogger(__name__)

try:
    from tqdm import tqdm as _tqdm

    _HAS_TQDM = True
except ImportError:
    _HAS_TQDM = False


class ProgressTracker:
    """Display and track progress for long-running operations.

    Uses tqdm when available; otherwise falls back to simple percentage
    printing at every 10% milestone.

    Can be used as a context manager::

        with ProgressTracker(total=100, description="Processing") as p:
            for item in items:
                process(item)
                p.update()
    """

    def __init__(self, total: int, description: str = "Processing") -> None:
        """Initialise the progress tracker.

        Args:
            total: Total number of steps.
            description: Human-readable description shown alongside
                the progress bar.
        """
        self.total = total
        self.description = description
        self._current = 0
        self._closed = False

        if _HAS_TQDM:
            self._bar: Optional[_tqdm] = _tqdm(
                total=total,
                desc=description,
                unit="step",
                file=sys.stderr,
            )
        else:
            self._bar = None
            self._last_printed_pct = -1
            logger.debug(
                "tqdm not available; using print-based progress fallback."
            )

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def update(self, n: int = 1) -> None:
        """Advance the progress bar by *n* steps.

        Args:
            n: Number of steps to advance (default ``1``).
        """
        if self._closed:
            return

        self._current += n

        if self._bar is not None:
            self._bar.update(n)
        else:
            self._print_progress()

    def set_description(self, desc: str) -> None:
        """Update the progress description text.

        Args:
            desc: New description string.
        """
        self.description = desc
        if self._bar is not None:
            self._bar.set_description(desc)

    def close(self) -> None:
        """Close the progress bar and release resources."""
        if self._closed:
            return
        self._closed = True

        if self._bar is not None:
            self._bar.close()
        else:
            # Ensure 100% is printed on close
            if self._last_printed_pct < 100:
                print(f"{self.description}: 100%", file=sys.stderr)

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self) -> "ProgressTracker":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:  # noqa: ANN001
        self.close()

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _print_progress(self) -> None:
        """Print progress percentage at 10% intervals (fallback mode)."""
        if self.total <= 0:
            return

        pct = int(self._current / self.total * 100)
        # Print at every 10% boundary that hasn't been printed yet
        milestone = pct // 10 * 10
        if milestone > self._last_printed_pct:
            self._last_printed_pct = milestone
            print(f"{self.description}: {milestone}%", file=sys.stderr)
