"""
pdf_extract.py — Textextraktion aus gescannten Brief-PDFs.

Strategie
---------
Die HP-LaserJet-Scans enthalten bereits einen Textlayer (300 DPI, OCR an).
Dieser Layer ist fehlerhaft ("Finanzamt Internatjonq", "Flnanzamt"), aber fuer
die regelbasierte Klassifizierung ausreichend — die eigentliche Fehlertoleranz
leistet spaeter rapidfuzz im rule_engine. Daher:

  1. Textlayer pro Seite lesen (pdfplumber; primaer).
  2. Ist der Textlayer einer Seite praktisch leer -> optionaler OCR-Fallback
     (pytesseract). Standardmaessig ABGESCHALTET (kein Tesseract noetig).

Extrahierte Felder (kalibriert an TAX.pdf, 25 Briefe)
-----------------------------------------------------
  absender   : "Finanzamt <X>"            Trefferquote 25/25
  datum      : Dokumentdatum              25/25  (bevorzugt die "Datum"-Zeile,
               Stempeldatum wird ignoriert)
  empfaenger : "Amazon Deutschland ..."   sofern im Text vorhanden (~17/25)
  volltext   : kompletter Text            fuer die keyword-Regeln

Alle Felder sind best-effort: fehlende Werte => leerer String, nie Exception.

Abhaengigkeiten: pdfplumber. Optional: pytesseract + Pillow (nur OCR-Fallback).
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pdfplumber
except ImportError as exc:  # pragma: no cover
    raise ImportError("pdfplumber wird benoetigt: pip install pdfplumber") from exc

log = logging.getLogger("postpilot.pdf_extract")


# --------------------------------------------------------------------------- #
# Muster (kalibriert) — zentral, damit im rule_engine nicht dupliziert
# --------------------------------------------------------------------------- #
_MONTHS = (
    "Januar|Februar|M[aä]rz|April|Mai|Juni|Juli|August|"
    "September|Oktober|November|Dezember"
)
_MONTH_NUM = {
    "januar": 1,
    "februar": 2,
    "marz": 3,
    "maerz": 3,
    "april": 4,
    "mai": 5,
    "juni": 6,
    "juli": 7,
    "august": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "dezember": 12,
}

# Absender: "Finanzamt <Name>" — Name = Buchstaben/Umlaute/Bindestrich, ein Zeilenrest.
_RE_ABSENDER = re.compile(
    r"Finanzamt\s+([A-Za-zÄÖÜäöüß.\-]+(?:\s+[A-Za-zÄÖÜäöüß.\-]+)?)", re.IGNORECASE
)
# Empfaenger (Amazon-Entitaeten im Brief).
_RE_EMPFAENGER = re.compile(
    r"(Amazon\s+Deutschland\s+Services[A-Za-zÄÖÜäöüß. ]*)", re.IGNORECASE
)

# Datum: bevorzugt die "Datum"-Zeile des Briefkopfs; Stempeldatum bleibt aussen vor.
_RE_DATUM_LABEL_NUM = re.compile(
    r"Datum[^0-9]{0,15}(\d{1,2})\.(\d{1,2})\.(\d{4})", re.IGNORECASE
)
_RE_DATUM_LABEL_TXT = re.compile(
    r"Datum[^0-9]{0,15}(\d{1,2})\.?\s*(" + _MONTHS + r")\s+(\d{4})", re.IGNORECASE
)
_RE_DATUM_TXT = re.compile(
    r"\b(\d{1,2})\.?\s*(" + _MONTHS + r")\s+(\d{4})\b", re.IGNORECASE
)
_RE_DATUM_NUM = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b")


def _norm_date(d, m, y) -> str:
    """Normalisiert auf TT.MM.JJJJ. Monat kann Zahl oder (OCR-)Name sein."""
    if isinstance(m, str) and not m.isdigit():
        key = m.lower().replace("ä", "a").replace("ö", "o").replace("ü", "u")
        m = _MONTH_NUM.get(key, 0)
    try:
        return f"{int(d):02d}.{int(m):02d}.{int(y):04d}"
    except (ValueError, TypeError):
        return ""


# --------------------------------------------------------------------------- #
# Ergebnis-Objekt
# --------------------------------------------------------------------------- #
@dataclass
class ExtractResult:
    source: str
    absender: str = ""
    empfaenger: str = ""
    datum: str = ""
    datum_quelle: str = ""  # Herkunft des Datums (Nachvollziehbarkeit)
    volltext: str = ""
    page_count: int = 0
    used_ocr: bool = False  # True, wenn OCR-Fallback gegriffen hat
    empty_pages: list[int] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "datei": self.source,
            "absender": self.absender,
            "empfaenger": self.empfaenger,
            "datum": self.datum,
            "datum_quelle": self.datum_quelle,
            "seiten": self.page_count,
            "ocr": self.used_ocr,
        }


# --------------------------------------------------------------------------- #
# Textlayer lesen (+ optionaler OCR-Fallback)
# --------------------------------------------------------------------------- #
def _ocr_page(page, dpi: int = 300, lang: str = "deu") -> str:
    """OCR-Fallback fuer eine leere Seite. Nur aktiv, wenn pytesseract vorhanden."""
    try:
        import pytesseract  # lokaler Import: Tool laeuft auch ohne das Paket
    except ImportError:
        log.warning("OCR-Fallback angefordert, aber pytesseract fehlt — uebersprungen.")
        return ""
    try:
        img = page.to_image(resolution=dpi).original
        return pytesseract.image_to_string(img, lang=lang) or ""
    except Exception as exc:  # pragma: no cover
        log.warning("OCR-Fallback fehlgeschlagen: %s", exc)
        return ""


def read_pages(
    pdf_path: str | Path,
    *,
    ocr_fallback: bool = False,
    ocr_dpi: int = 300,
    ocr_lang: str = "deu",
    min_chars: int = 15,
) -> tuple[list[str], list[int], bool]:
    """
    Liest den Textlayer aller Seiten.

    Args:
        ocr_fallback: Wenn True und eine Seite < min_chars Zeichen liefert,
                      wird OCR versucht (benoetigt pytesseract). Default False.
        min_chars:    Schwelle, ab der eine Seite als "leer" gilt.

    Returns: (seiten_texte, leere_seiten_1indexiert, used_ocr)
    """
    pdf_path = Path(pdf_path)
    texts: list[str] = []
    empty: list[int] = []
    used_ocr = False
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            txt = (page.extract_text() or "").strip()
            if len(txt) < min_chars:
                empty.append(i + 1)
                if ocr_fallback:
                    ocr_txt = _ocr_page(page, ocr_dpi, ocr_lang).strip()
                    if ocr_txt:
                        txt = ocr_txt
                        used_ocr = True
            texts.append(txt)
    return texts, empty, used_ocr


# --------------------------------------------------------------------------- #
# Feld-Extraktion
# --------------------------------------------------------------------------- #
def extract_absender(text: str) -> str:
    m = _RE_ABSENDER.search(text)
    if not m:
        return ""
    name = re.sub(r"\s+", " ", m.group(1)).strip(" .,-")
    return f"Finanzamt {name}" if name else "Finanzamt"


def extract_empfaenger(text: str) -> str:
    m = _RE_EMPFAENGER.search(text)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


def extract_datum(text: str) -> tuple[str, str]:
    """
    Dokumentdatum als TT.MM.JJJJ + Herkunft.

    Prioritaet:
      1. "Datum"-Zeile im Briefkopf (numerisch, dann ausgeschrieben)
      2. erstes ausgeschriebenes Datum im Fliesstext
      3. erstes numerisches Datum im Fliesstext
    Das visuelle Stempeldatum liegt im farbigen Stempel und ist im Textlayer
    unzuverlaessig — es gewinnt daher praktisch nie gegen die "Datum"-Zeile.
    """
    m = _RE_DATUM_LABEL_NUM.search(text)
    if m:
        return _norm_date(*m.groups()), "Datum-Zeile"
    m = _RE_DATUM_LABEL_TXT.search(text)
    if m:
        return _norm_date(m.group(1), m.group(2), m.group(3)), "Datum-Zeile(txt)"
    m = _RE_DATUM_TXT.search(text)
    if m:
        return _norm_date(m.group(1), m.group(2), m.group(3)), "Fliesstext(txt)"
    m = _RE_DATUM_NUM.search(text)
    if m:
        return _norm_date(*m.groups()), "Fliesstext(num)"
    return "", "keins"


# --------------------------------------------------------------------------- #
# Haupteinstieg
# --------------------------------------------------------------------------- #
def extract(
    pdf_path: str | Path,
    *,
    ocr_fallback: bool = False,
    ocr_dpi: int = 300,
    ocr_lang: str = "deu",
) -> ExtractResult:
    """
    Extrahiert alle Felder aus einem (Einzel-)Brief-PDF.

    Fehlertolerant: fehlende Felder bleiben leer, es wird nie geworfen.
    Der 'volltext' dient dem rule_engine fuer die keyword-Regeln.
    """
    pdf_path = Path(pdf_path)
    texts, empty, used_ocr = read_pages(
        pdf_path, ocr_fallback=ocr_fallback, ocr_dpi=ocr_dpi, ocr_lang=ocr_lang
    )
    volltext = "\n".join(texts)
    datum, datum_quelle = extract_datum(volltext)
    result = ExtractResult(
        source=str(pdf_path),
        absender=extract_absender(volltext),
        empfaenger=extract_empfaenger(volltext),
        datum=datum,
        datum_quelle=datum_quelle,
        volltext=volltext,
        page_count=len(texts),
        used_ocr=used_ocr,
        empty_pages=empty,
    )
    log.info(
        "Extrakt '%s': absender=%r datum=%r(%s) seiten=%d ocr=%s leer=%s",
        pdf_path.name,
        result.absender,
        result.datum,
        result.datum_quelle,
        result.page_count,
        used_ocr,
        empty or "-",
    )
    return result


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Textfelder aus einem Brief-PDF extrahieren."
    )
    parser.add_argument("pdf", help="Pfad zum PDF")
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="OCR-Fallback aktivieren (pytesseract noetig)",
    )
    parser.add_argument("--full", action="store_true", help="Volltext mit ausgeben")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )
    res = extract(args.pdf, ocr_fallback=args.ocr)
    out = res.as_dict()
    if args.full:
        out["volltext"] = res.volltext
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    _main()
