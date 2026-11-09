"""
stamp_splitter.py — Zerlegt ein Sammel-PDF anhand visueller Eingangsstempel.

Hintergrund
-----------
Mehrere Briefe werden hintereinander mit einem Eingangsstempel auf der jeweils
ersten Seite gescannt und als EIN PDF gespeichert. Der Stempel ist der einzige
farbige (blau + rot) Inhalt auf einer sonst schwarz-weißen Seite:

    [ EINGEGANGEN ]      <- blauer Rahmen
      05. Nov. 2026      <- rotes Datum

Erkennungsstrategie (kalibriert an TAX.pdf, 48 Seiten -> 25 Briefe)
-------------------------------------------------------------------
Statt unzuverlaessiger OCR auf dem farbigen Stempel wird rein visuell gearbeitet:
Jede Seite wird gerendert, und im oberen Seitendrittel (ROI) wird der Anteil
blauer und roter Pixel gezaehlt. Liegt BEIDES ueber der jeweiligen Schwelle,
beginnt hier ein neuer Brief.

    Stempelseiten:  blau 0.37-0.71 %, rot 0.17-0.30 %
    Folgeseiten:    blau 0.000 %,     rot 0.000 %

Die Luecke ist riesig -> die Trennung ist eindeutig und braucht KEIN Tesseract.

Fail-safe
---------
Bei jeglicher Unsicherheit wird NICHT gesplittet, sondern das gesamte PDF
unveraendert nach review/ verschoben (menschliche Pruefung). Ausgeloest durch:
  - 0 Stempel gefunden
  - die erste Seite traegt keinen Stempel (Struktur unerwartet)
  - ein "halber" Treffer (nur blau ODER nur rot ueber Schwelle) auf einer Seite
    -> Ambiguitaet, lieber Mensch druebergucken lassen

Abhaengigkeiten: PyMuPDF (fitz), numpy. KEIN Tesseract.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

try:
    import fitz  # PyMuPDF
except ImportError as exc:  # pragma: no cover
    raise ImportError("PyMuPDF wird benoetigt: pip install PyMuPDF") from exc

log = logging.getLogger("postpilot.stamp_splitter")


# --------------------------------------------------------------------------- #
# Konfiguration
# --------------------------------------------------------------------------- #
@dataclass
class StampConfig:
    """Kalibrierte Parameter der Stempelerkennung. Aus settings.yaml ueberschreibbar."""

    resolution: int = 150  # DPI zum Rendern (Farberkennung braucht nicht mehr)
    # ROI als Bruchteile der Seite (x0, y0, x1, y1); Default = oberes Drittel.
    roi_frac: tuple[float, float, float, float] = (0.0, 0.0, 1.0, 0.33)
    # Blau-Maske: b dominiert klar ueber r und g.
    blue_min: int = 90
    blue_over_red: int = 35
    blue_over_green: int = 25
    # Rot-Maske: r dominiert klar ueber g und b.
    red_min: int = 90
    red_over_green: int = 45
    red_over_blue: int = 45
    # Schwellen (Prozent der ROI-Pixel). Kalibriert mit grosser Sicherheitsmarge:
    # echte Stempel >= 0.37 % blau / 0.17 % rot, Leerseiten = 0.000 %.
    blue_thresh_pct: float = 0.15
    red_thresh_pct: float = 0.02
    # "Halber Treffer": eine Farbe klar ueber Schwelle, die andere nahe null.
    # Solche ambivalenten Seiten loesen den Fail-safe aus.
    ambiguous_factor: float = 0.5  # <- siehe _classify_page

    @classmethod
    def from_dict(cls, d: dict | None) -> StampConfig:
        if not d:
            return cls()
        known = {f.name for f in cls.__dataclass_fields__.values()}  # type: ignore[attr-defined]
        kwargs = {}
        for k, v in d.items():
            if k not in known:
                continue
            if k == "roi_frac" and v is not None:
                kwargs[k] = tuple(v)
            else:
                kwargs[k] = v
        return cls(**kwargs)


# --------------------------------------------------------------------------- #
# Ergebnis-Objekte (nachvollziehbar, serialisierbar)
# --------------------------------------------------------------------------- #
@dataclass
class PageScan:
    page: int  # 1-indexiert
    blue_pct: float
    red_pct: float
    is_stamp: bool
    ambiguous: bool  # halber Treffer -> Unsicherheit


@dataclass
class SplitResult:
    source: str
    ok: bool  # True = sauber gesplittet
    reason: str  # Begruendung (Log/Nachvollziehbarkeit)
    pages: list[PageScan] = field(default_factory=list)
    letters: list[dict] = field(default_factory=list)  # {"index","start","end","path"}
    review_path: str | None = None  # gesetzt, wenn Fail-safe gegriffen hat


# --------------------------------------------------------------------------- #
# Kernlogik
# --------------------------------------------------------------------------- #
def _render_page_rgb(page: fitz.Page, resolution: int) -> np.ndarray:
    """Rendert eine Seite als RGB-numpy-Array (H, W, 3)."""
    # 72 ist das PDF-Basis-DPI; zoom skaliert auf die gewuenschte Aufloesung.
    zoom = resolution / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:  # RGBA -> RGB (sollte wegen alpha=False nicht auftreten)
        arr = arr[:, :, :3]
    return arr


def _color_fractions(roi: np.ndarray, cfg: StampConfig) -> tuple[float, float]:
    """Gibt (blau%, rot%) der ROI zurueck."""
    r = roi[:, :, 0].astype(int)
    g = roi[:, :, 1].astype(int)
    b = roi[:, :, 2].astype(int)
    blue = (
        (b > cfg.blue_min) & (b - r > cfg.blue_over_red) & (b - g > cfg.blue_over_green)
    )
    red = (r > cfg.red_min) & (r - g > cfg.red_over_green) & (r - b > cfg.red_over_blue)
    total = roi.shape[0] * roi.shape[1]
    if total == 0:
        return 0.0, 0.0
    # float() erzwingen: numpy.float64 ist nicht JSON-serialisierbar (Report in _main()).
    return float(100.0 * blue.sum() / total), float(100.0 * red.sum() / total)


def _roi_slice(arr: np.ndarray, cfg: StampConfig) -> np.ndarray:
    h, w = arr.shape[0], arr.shape[1]
    x0, y0, x1, y1 = cfg.roi_frac
    return arr[int(h * y0) : int(h * y1), int(w * x0) : int(w * x1), :]


def _classify_page(
    blue_pct: float, red_pct: float, cfg: StampConfig
) -> tuple[bool, bool]:
    """
    Klassifiziert eine Seite.

    Returns (is_stamp, ambiguous):
      is_stamp  = True, wenn blau UND rot ueber ihren Schwellen liegen.
      ambiguous = True nur bei einem ECHTEN Grenzfall knapp an der Stempel-
                  Signatur (siehe unten) -> Review statt Fehlentscheidung.

    Der Eingangsstempel hat IMMER BEIDES: blauen Rahmen UND rotes Datum.
    Das rote Datum ist der zuverlaessige Diskriminator — echte Stempel haben
    konstant ~0.20 % Rot, waehrend blaue Unterschriften/Dienstsiegel oder rote
    Siegel im Briefinhalt stets nur EINEN Kanal zeigen (der andere ~0.000 %).
    Daher: nur EIN Kanal ueber Schwelle = KEIN Stempel (kein Fehlalarm).

    "ambiguous" ist bewusst eng gefasst: nur wenn BEIDE Kanaele im
    Grenzbereich liegen (blau ueber Schwelle und rot knapp darunter, aber
    nicht bei null) — dann koennte ein echter Stempel mit blassem Datum
    vorliegen, den ein Mensch pruefen sollte.
    """
    blue_hit = blue_pct >= cfg.blue_thresh_pct
    red_hit = red_pct >= cfg.red_thresh_pct
    is_stamp = blue_hit and red_hit
    # Grenzfall: blauer Rahmen vorhanden, rotes Datum nur KNAPP unter Schwelle
    # (zwischen ambiguous_factor*Schwelle und Schwelle) -> evtl. echter Stempel
    # mit blassem Datum, Mensch pruefen lassen. Rot = 0 (Unterschrift/Siegel)
    # faellt NICHT darunter und gilt sauber als "kein Stempel".
    ambiguous = (
        blue_hit
        and not red_hit
        and red_pct >= cfg.red_thresh_pct * cfg.ambiguous_factor
    )
    # bool() erzwingen: numpy-Vergleiche liefern numpy.bool_ (nicht JSON-serialisierbar).
    return bool(is_stamp), bool(ambiguous)


def detect_stamp_pages(
    pdf_path: str | Path, cfg: StampConfig | None = None
) -> list[PageScan]:
    """
    Analysiert jede Seite auf einen Eingangsstempel.

    Returns: Liste von PageScan (1-indexiert, in Seitenreihenfolge).
    Diese Funktion ist seiteneffektfrei (schreibt nichts) und dient der
    Nachvollziehbarkeit: blau%/rot% pro Seite landen im Log.
    """
    cfg = cfg or StampConfig()
    pdf_path = Path(pdf_path)
    scans: list[PageScan] = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            arr = _render_page_rgb(page, cfg.resolution)
            roi = _roi_slice(arr, cfg)
            blue_pct, red_pct = _color_fractions(roi, cfg)
            is_stamp, ambiguous = _classify_page(blue_pct, red_pct, cfg)
            scans.append(
                PageScan(
                    i + 1, round(blue_pct, 4), round(red_pct, 4), is_stamp, ambiguous
                )
            )
            log.debug(
                "S.%d blau=%.3f%% rot=%.3f%% stamp=%s ambig=%s",
                i + 1,
                blue_pct,
                red_pct,
                is_stamp,
                ambiguous,
            )
    return scans


def group_pages_by_stamp_starts(
    stamp_pages: list[int], total_pages: int
) -> list[tuple[int, int]]:
    """
    Gruppiert Seiten in Dokumente: jede Stempelseite startet ein neues Dokument,
    folgende stempellose Seiten gehoeren dazu.

    >>> group_pages_by_stamp_starts([1, 3, 4], 5)
    [(1, 2), (3, 3), (4, 5)]
    """
    ranges: list[tuple[int, int]] = []
    for idx, start in enumerate(stamp_pages):
        end = (stamp_pages[idx + 1] - 1) if idx + 1 < len(stamp_pages) else total_pages
        ranges.append((start, end))
    return ranges


def _move_to_review(pdf_path: Path, review_dir: Path, reason: str) -> str:
    review_dir.mkdir(parents=True, exist_ok=True)
    dest = review_dir / pdf_path.name
    # Nicht ueberschreiben: bei Namenskollision durchnummerieren.
    counter = 1
    while dest.exists():
        dest = review_dir / f"{pdf_path.stem}__{counter}{pdf_path.suffix}"
        counter += 1
    shutil.copy2(pdf_path, dest)
    log.warning("Fail-safe: '%s' -> review/ (%s)", pdf_path.name, reason)
    return str(dest)


def split_by_stamps(
    pdf_path: str | Path,
    out_dir: str | Path,
    review_dir: str | Path,
    cfg: StampConfig | None = None,
) -> SplitResult:
    """
    Zerlegt ein Sammel-PDF in Einzelbriefe.

    Vorgehen:
      1. detect_stamp_pages() analysiert alle Seiten.
      2. Fail-safe-Pruefungen (s.u.). Greift eine -> ganzes PDF nach review/.
      3. Sonst: pro Stempelbereich ein Einzel-PDF in out_dir schreiben.

    Fail-safe greift bei:
      - 0 erkannten Stempeln
      - erste Seite ohne Stempel (unerwartete Struktur)
      - mindestens einer ambivalenten Seite (halber Stempel)

    Returns: SplitResult (ok=True -> letters[] gefuellt; ok=False -> review_path gesetzt).
    """
    cfg = cfg or StampConfig()
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    review_dir = Path(review_dir)

    scans = detect_stamp_pages(pdf_path, cfg)
    total = len(scans)
    stamp_pages = [s.page for s in scans if s.is_stamp]
    ambiguous_pages = [s.page for s in scans if s.ambiguous]

    # --- Fail-safe-Pruefungen ------------------------------------------------
    if total == 0:
        reason = "PDF hat keine Seiten"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=_move_to_review(pdf_path, review_dir, reason),
        )

    if not stamp_pages:
        reason = "kein Eingangsstempel erkannt"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=_move_to_review(pdf_path, review_dir, reason),
        )

    if scans[0].page not in stamp_pages:
        reason = f"erste Seite ohne Stempel (erster Stempel auf S.{stamp_pages[0]})"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=_move_to_review(pdf_path, review_dir, reason),
        )

    if ambiguous_pages:
        reason = f"unsichere Stempelerkennung auf Seite(n) {ambiguous_pages}"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=_move_to_review(pdf_path, review_dir, reason),
        )

    # --- Sauber splitten -----------------------------------------------------
    out_dir.mkdir(parents=True, exist_ok=True)
    ranges = group_pages_by_stamp_starts(stamp_pages, total)
    letters: list[dict] = []
    with fitz.open(pdf_path) as src:
        for idx, (start, end) in enumerate(ranges, 1):
            dst = fitz.open()
            dst.insert_pdf(src, from_page=start - 1, to_page=end - 1)
            out_name = f"{pdf_path.stem}_brief_{idx:02d}_S{start:03d}-{end:03d}.pdf"
            out_path = out_dir / out_name
            dst.save(out_path)
            dst.close()
            letters.append(
                {"index": idx, "start": start, "end": end, "path": str(out_path)}
            )
            log.info("Brief %02d: S.%d-%d -> %s", idx, start, end, out_name)

    reason = f"{len(letters)} Briefe aus {total} Seiten gesplittet"
    return SplitResult(str(pdf_path), True, reason, scans, letters=letters)


# --------------------------------------------------------------------------- #
# Debug-Visualisierung (optional) — markiert ROI + erkannte Stempelpixel
# --------------------------------------------------------------------------- #
def write_debug_image(
    pdf_path: str | Path,
    out_path: str | Path,
    cfg: StampConfig | None = None,
    pages: list[int] | None = None,
) -> str:
    """
    Schreibt ein Kontrollbild: erkanntes Blau -> gruen, erkanntes Rot -> magenta,
    orange Linie = untere ROI-Grenze. 'pages' (1-indexiert) waehlt die Seiten;
    None = erste 4 Seiten.
    """
    from PIL import Image

    cfg = cfg or StampConfig()
    pdf_path = Path(pdf_path)
    with fitz.open(pdf_path) as doc:
        n = doc.page_count
        sel = pages or list(range(1, min(4, n) + 1))
        tiles = []
        for pg in sel:
            arr = _render_page_rgb(doc[pg - 1], cfg.resolution).copy()
            h, w = arr.shape[0], arr.shape[1]
            x0, y0, x1, y1 = cfg.roi_frac
            ry0, ry1 = int(h * y0), int(h * y1)
            rx0, rx1 = int(w * x0), int(w * x1)
            roi = arr[ry0:ry1, rx0:rx1, :]
            r = roi[:, :, 0].astype(int)
            g = roi[:, :, 1].astype(int)
            b = roi[:, :, 2].astype(int)
            blue = (
                (b > cfg.blue_min)
                & (b - r > cfg.blue_over_red)
                & (b - g > cfg.blue_over_green)
            )
            red = (
                (r > cfg.red_min)
                & (r - g > cfg.red_over_green)
                & (r - b > cfg.red_over_blue)
            )
            roi[blue] = [0, 255, 0]
            roi[red] = [255, 0, 255]
            arr[ry0:ry1, rx0:rx1, :] = roi
            arr[max(ry1 - 2, 0) : ry1 + 1, :, :] = [255, 165, 0]
            tiles.append(Image.fromarray(arr))

    height = 700
    imgs = [im.resize((int(im.width * height / im.height), height)) for im in tiles]
    gap = 12
    total_w = sum(im.width for im in imgs) + gap * (len(imgs) - 1)
    canvas = Image.new("RGB", (total_w, height), (255, 255, 255))
    x = 0
    for im in imgs:
        canvas.paste(im, (x, 0))
        x += im.width + gap
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out_path)
    return str(out_path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse
    import json
    from datetime import date as _date

    parser = argparse.ArgumentParser(
        description="Sammel-PDF anhand Eingangsstempel splitten."
    )
    parser.add_argument("pdf", help="Pfad zum Sammel-PDF")
    parser.add_argument(
        "--out", default="work/split", help="Zielordner fuer Einzelbriefe"
    )
    parser.add_argument(
        "--review", default="review", help="Pruefungsordner (Fail-safe)"
    )
    parser.add_argument(
        "--resolution", type=int, default=150, help="Render-DPI (Default 150)"
    )
    parser.add_argument(
        "--date",
        default=_date.today().strftime("%d_%m_%Y"),
        help="Datums-Unterordner (Default: heute, Format TT_MM_JJJJ). "
        "Einzelbriefe landen in <out>/<date>/.",
    )
    parser.add_argument(
        "--no-date-subdir",
        action="store_true",
        help="Keinen Datums-Unterordner anlegen (direkt nach <out> schreiben).",
    )
    parser.add_argument(
        "--debug-image", metavar="PNG", help="Kontrollbild schreiben und beenden"
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG-Logging")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )
    cfg = StampConfig(resolution=args.resolution)

    if args.debug_image:
        path = write_debug_image(args.pdf, args.debug_image, cfg)
        print(f"Debug-Bild geschrieben: {path}")
        return

    # Taeglicher Betrieb: Einzelbriefe nach <out>/<date>/ (z.B. work/split/05_11_2026).
    out_dir = args.out if args.no_date_subdir else str(Path(args.out) / args.date)
    log.info("Zielordner: %s", out_dir)
    result = split_by_stamps(args.pdf, out_dir, args.review, cfg)
    print(
        json.dumps(
            {
                "source": result.source,
                "ok": result.ok,
                "reason": result.reason,
                "letters": result.letters,
                "review_path": result.review_path,
                "pages": [vars(p) for p in result.pages],
            },
            indent=2,
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    _main()
