from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

try:
    import pdfplumber
except ImportError as exc:  # pragma: no cover
    raise ImportError("pdfplumber required: pip install pdfplumber") from exc

log = logging.getLogger("postpilot.pdf_extract")


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

_RE_SENDER = re.compile(
    r"Finanzamt\s+([A-Za-zÄÖÜäöüß.\-]+(?:\s+[A-Za-zÄÖÜäöüß.\-]+)?)", re.IGNORECASE
)
_RE_RECIPIENT = re.compile(
    r"(Amazon\s+Deutschland\s+Services[A-Za-zÄÖÜäöüß. ]*)", re.IGNORECASE
)

_RE_DATE_LABEL_NUM = re.compile(
    r"Datum[^0-9]{0,15}(\d{1,2})\.(\d{1,2})\.(\d{4})", re.IGNORECASE
)
_RE_DATE_LABEL_TXT = re.compile(
    r"Datum[^0-9]{0,15}(\d{1,2})\.?\s*(" + _MONTHS + r")\s+(\d{4})", re.IGNORECASE
)
_RE_DATE_TXT = re.compile(
    r"\b(\d{1,2})\.?\s*(" + _MONTHS + r")\s+(\d{4})\b", re.IGNORECASE
)
_RE_DATE_NUM = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b")


def _norm_date(d, m, y) -> str:
    if isinstance(m, str) and not m.isdigit():
        key = m.lower().replace("ä", "a").replace("ö", "o").replace("ü", "u")
        m = _MONTH_NUM.get(key, 0)
    try:
        return f"{int(d):02d}.{int(m):02d}.{int(y):04d}"
    except (ValueError, TypeError):
        return ""


@dataclass
class ExtractResult:
    source: str
    sender: str = ""
    recipient: str = ""
    date: str = ""
    date_source: str = ""
    full_text: str = ""
    page_count: int = 0
    used_ocr: bool = False
    empty_pages: list[int] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "file": self.source,
            "sender": self.sender,
            "recipient": self.recipient,
            "date": self.date,
            "date_source": self.date_source,
            "pages": self.page_count,
            "ocr": self.used_ocr,
        }


def _ocr_page(page, dpi: int = 300, lang: str = "deu") -> str:
    try:
        import pytesseract
    except ImportError:
        log.warning("OCR fallback requested but pytesseract missing - skipped.")
        return ""
    try:
        img = page.to_image(resolution=dpi).original
        return pytesseract.image_to_string(img, lang=lang) or ""
    except Exception as exc:  # pragma: no cover
        log.warning("OCR fallback failed: %s", exc)
        return ""


def read_pages(
    pdf_path: str | Path,
    *,
    ocr_fallback: bool = False,
    ocr_dpi: int = 300,
    ocr_lang: str = "deu",
    min_chars: int = 15,
) -> tuple[list[str], list[int], bool]:
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


def extract_sender(text: str) -> str:
    m = _RE_SENDER.search(text)
    if not m:
        return ""
    name = re.sub(r"\s+", " ", m.group(1)).strip(" .,-")
    return f"Finanzamt {name}" if name else "Finanzamt"


def extract_recipient(text: str) -> str:
    m = _RE_RECIPIENT.search(text)
    return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""


def extract_date(text: str) -> tuple[str, str]:
    m = _RE_DATE_LABEL_NUM.search(text)
    if m:
        return _norm_date(*m.groups()), "header-line"
    m = _RE_DATE_LABEL_TXT.search(text)
    if m:
        return _norm_date(m.group(1), m.group(2), m.group(3)), "header-line(txt)"
    m = _RE_DATE_TXT.search(text)
    if m:
        return _norm_date(m.group(1), m.group(2), m.group(3)), "body(txt)"
    m = _RE_DATE_NUM.search(text)
    if m:
        return _norm_date(*m.groups()), "body(num)"
    return "", "none"


def extract(
    pdf_path: str | Path,
    *,
    ocr_fallback: bool = False,
    ocr_dpi: int = 300,
    ocr_lang: str = "deu",
) -> ExtractResult:
    pdf_path = Path(pdf_path)
    texts, empty, used_ocr = read_pages(
        pdf_path, ocr_fallback=ocr_fallback, ocr_dpi=ocr_dpi, ocr_lang=ocr_lang
    )
    full_text = "\n".join(texts)
    date, date_source = extract_date(full_text)
    result = ExtractResult(
        source=str(pdf_path),
        sender=extract_sender(full_text),
        recipient=extract_recipient(full_text),
        date=date,
        date_source=date_source,
        full_text=full_text,
        page_count=len(texts),
        used_ocr=used_ocr,
        empty_pages=empty,
    )
    log.info(
        "extract '%s': sender=%r date=%r(%s) pages=%d ocr=%s empty=%s",
        pdf_path.name,
        result.sender,
        result.date,
        result.date_source,
        result.page_count,
        used_ocr,
        empty or "-",
    )
    return result


def _main() -> None:
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Extract text fields from a letter PDF."
    )
    parser.add_argument("pdf", help="Path to the PDF")
    parser.add_argument(
        "--ocr", action="store_true", help="Enable OCR fallback (needs pytesseract)"
    )
    parser.add_argument(
        "--full", action="store_true", help="Include full text in output"
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )
    res = extract(args.pdf, ocr_fallback=args.ocr)
    out = res.as_dict()
    if args.full:
        out["full_text"] = res.full_text
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    _main()
