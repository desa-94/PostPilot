from __future__ import annotations

import logging
import shutil
from pathlib import Path

_log = logging.getLogger("postpilot.fsutil")


def move_to_review(
    pdf_path: str | Path,
    review_dir: str | Path,
    reason: str,
    *,
    logger: logging.Logger | None = None,
) -> str:
    """Copy a letter PDF into the review folder, never overwriting.

    If a file of the same name already exists, a numeric suffix
    (``name__1.pdf``, ``name__2.pdf`` ...) is appended. Returns the final
    destination path as a string.

    Shared by the router (classification fallback) and the stamp splitter
    (fail-safe) so the move/dedupe logic lives in exactly one place.
    """
    pdf_path = Path(pdf_path)
    review_dir = Path(review_dir)
    review_dir.mkdir(parents=True, exist_ok=True)

    dest = review_dir / pdf_path.name
    counter = 1
    while dest.exists():
        dest = review_dir / f"{pdf_path.stem}__{counter}{pdf_path.suffix}"
        counter += 1

    shutil.copy2(pdf_path, dest)
    (logger or _log).warning("-> review/: %s (%s)", pdf_path.name, reason)
    return str(dest)
