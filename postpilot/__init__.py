"""PostPilot — regelbasierte, lokale Briefweiterleitung (kein LLM, keine Cloud)."""

import logging

__version__ = "0.1.0"

# Laute Fremd-Logger daempfen. pdfminer (unter pdfplumber) schreibt sonst
# pro PDF-Token eine DEBUG-Zeile -> zehntausende Zeilen und minutenlange
# Laeufe, sobald irgendwo DEBUG-Logging (-v) aktiv ist. Wir wollen NUR unsere
# eigenen postpilot.*-DEBUG-Meldungen sehen, nicht die der Bibliotheken.
for _noisy in ("pdfminer", "pdfplumber", "PIL"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)
