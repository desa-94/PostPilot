"""PostPilot - local rule-based mail forwarding (no LLM, no cloud)."""

import logging

__version__ = "0.1.0"

for _noisy in ("pdfminer", "pdfplumber", "PIL"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)
