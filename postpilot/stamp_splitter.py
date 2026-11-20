from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

try:
    # PyMuPDF >= 1.24 ships the `pymupdf` name; the old `fitz` alias emits a
    # FutureWarning on import. Prefer `pymupdf`, fall back to `fitz` for
    # older installs. Either way we keep the local name `fitz`.
    import pymupdf as fitz
except ImportError as exc:  # pragma: no cover
    try:
        import fitz
    except ImportError:
        raise ImportError("PyMuPDF required: pip install PyMuPDF") from exc

log = logging.getLogger("postpilot.stamp_splitter")

from postpilot._fsutil import move_to_review


@dataclass
class StampConfig:
    resolution: int = 150
    roi_frac: tuple[float, float, float, float] = (0.0, 0.0, 1.0, 0.33)
    blue_min: int = 90
    blue_over_red: int = 35
    blue_over_green: int = 25
    red_min: int = 90
    red_over_green: int = 45
    red_over_blue: int = 45
    # Blue frame of the "EINGEGANGEN" stamp. Kept at 0.10 (not higher):
    # when the blue frame overlaps the black letterhead, blue ink over
    # black pixels drops below b>blue_min, so the measured blue fraction
    # falls to ~0.12-0.15. The red date (red_thresh_pct) stays the reliable
    # discriminator; blue only guards against red-only seal artifacts.
    blue_thresh_pct: float = 0.10
    red_thresh_pct: float = 0.02
    ambiguous_factor: float = 0.5
    # --- Yellow "Förmliche Zustellung" envelope detection ---------------
    # Some incoming mail starts with a yellow Postzustellungsurkunde
    # envelope (often scanned sideways), whose green-framed stamp cannot be
    # found by the blue+red rule (no blue at all; the yellow paper floods
    # the red mask). Instead we treat a strongly yellow page as its own
    # document separator: high red AND high green AND clearly low blue,
    # measured over the WHOLE page (the envelope fills it, orientation does
    # not matter). A plain white letter page scores ~0% here, so the
    # threshold is deliberately high.
    yellow_min: int = 120
    yellow_over_blue: int = 40
    yellow_thresh_pct: float = 20.0
    detect_yellow: bool = True

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


@dataclass
class PageScan:
    page: int
    blue_pct: float
    red_pct: float
    yellow_pct: float
    is_stamp: bool
    ambiguous: bool
    is_envelope: bool = False


@dataclass
class SplitResult:
    source: str
    ok: bool
    reason: str
    pages: list[PageScan] = field(default_factory=list)
    letters: list[dict] = field(default_factory=list)
    review_path: str | None = None


def _render_page_rgb(page: fitz.Page, resolution: int) -> np.ndarray:
    zoom = resolution / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n == 4:
        arr = arr[:, :, :3]
    return arr


def _color_fractions(roi: np.ndarray, cfg: StampConfig) -> tuple[float, float]:
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
    return float(100.0 * blue.sum() / total), float(100.0 * red.sum() / total)


def _yellow_fraction(arr: np.ndarray, cfg: StampConfig) -> float:
    """Fraction of strongly yellow pixels over the WHOLE page.

    Yellow = high red AND high green AND clearly lower blue. Used to flag
    the yellow Postzustellungsurkunde envelope, which starts a new
    document. Measured full-page (not ROI) so a sideways scan still counts.
    """
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    yellow = (r > cfg.yellow_min) & (g > cfg.yellow_min) & (r - b > cfg.yellow_over_blue) & (g - b > cfg.yellow_over_blue)
    total = arr.shape[0] * arr.shape[1]
    return 0.0 if total == 0 else float(100.0 * yellow.sum() / total)


def _roi_slice(arr: np.ndarray, cfg: StampConfig) -> np.ndarray:
    h, w = arr.shape[0], arr.shape[1]
    x0, y0, x1, y1 = cfg.roi_frac
    return arr[int(h * y0) : int(h * y1), int(w * x0) : int(w * x1), :]


def _classify_page(
    blue_pct: float, red_pct: float, cfg: StampConfig
) -> tuple[bool, bool]:
    blue_hit = blue_pct >= cfg.blue_thresh_pct
    red_hit = red_pct >= cfg.red_thresh_pct
    is_stamp = blue_hit and red_hit
    ambiguous = (
        blue_hit
        and not red_hit
        and red_pct >= cfg.red_thresh_pct * cfg.ambiguous_factor
    )
    return bool(is_stamp), bool(ambiguous)


def detect_stamp_pages(
    pdf_path: str | Path, cfg: StampConfig | None = None
) -> list[PageScan]:
    cfg = cfg or StampConfig()
    pdf_path = Path(pdf_path)
    scans: list[PageScan] = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            arr = _render_page_rgb(page, cfg.resolution)
            roi = _roi_slice(arr, cfg)
            blue_pct, red_pct = _color_fractions(roi, cfg)
            is_stamp, ambiguous = _classify_page(blue_pct, red_pct, cfg)
            yellow_pct = _yellow_fraction(arr, cfg) if cfg.detect_yellow else 0.0
            is_envelope = cfg.detect_yellow and yellow_pct >= cfg.yellow_thresh_pct
            scans.append(
                PageScan(
                    i + 1, round(blue_pct, 4), round(red_pct, 4),
                    round(yellow_pct, 4), is_stamp, ambiguous, is_envelope
                )
            )
            log.debug(
                "p.%d blue=%.3f%% red=%.3f%% yellow=%.3f%% stamp=%s ambig=%s env=%s",
                i + 1,
                blue_pct,
                red_pct,
                yellow_pct,
                is_stamp,
                ambiguous,
                is_envelope,
            )
    return scans


def group_pages_by_stamp_starts(
    stamp_pages: list[int], total_pages: int
) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    for idx, start in enumerate(stamp_pages):
        end = (stamp_pages[idx + 1] - 1) if idx + 1 < len(stamp_pages) else total_pages
        ranges.append((start, end))
    return ranges


def split_by_stamps(
    pdf_path: str | Path,
    out_dir: str | Path,
    review_dir: str | Path,
    cfg: StampConfig | None = None,
) -> SplitResult:
    cfg = cfg or StampConfig()
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    review_dir = Path(review_dir)

    scans = detect_stamp_pages(pdf_path, cfg)
    total = len(scans)
    stamp_pages = [s.page for s in scans if s.is_stamp]
    envelope_pages = [s.page for s in scans if s.is_envelope]
    ambiguous_pages = [s.page for s in scans if s.ambiguous]

    # A new document starts at a blue+red stamp OR a yellow envelope page.
    start_pages = sorted(set(stamp_pages) | set(envelope_pages))

    if total == 0:
        reason = "PDF has no pages"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=move_to_review(pdf_path, review_dir, reason, logger=log),
        )

    if not start_pages:
        reason = "no incoming stamp or envelope detected"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=move_to_review(pdf_path, review_dir, reason, logger=log),
        )

    if scans[0].page not in start_pages:
        reason = f"first page without stamp/envelope (first on p.{start_pages[0]})"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=move_to_review(pdf_path, review_dir, reason, logger=log),
        )

    if ambiguous_pages:
        reason = f"uncertain stamp detection on page(s) {ambiguous_pages}"
        return SplitResult(
            str(pdf_path),
            False,
            reason,
            scans,
            review_path=move_to_review(pdf_path, review_dir, reason, logger=log),
        )

    out_dir.mkdir(parents=True, exist_ok=True)
    ranges = group_pages_by_stamp_starts(start_pages, total)
    letters: list[dict] = []
    with fitz.open(pdf_path) as src:
        for idx, (start, end) in enumerate(ranges, 1):
            dst = fitz.open()
            dst.insert_pdf(src, from_page=start - 1, to_page=end - 1)
            out_name = f"{pdf_path.stem}_letter_{idx:02d}_p{start:03d}-{end:03d}.pdf"
            out_path = out_dir / out_name
            dst.save(out_path)
            dst.close()
            letters.append(
                {"index": idx, "start": start, "end": end, "path": str(out_path)}
            )
            log.info("letter %02d: p.%d-%d -> %s", idx, start, end, out_name)

    reason = f"{len(letters)} letters split from {total} pages"
    return SplitResult(str(pdf_path), True, reason, scans, letters=letters)


def write_debug_image(
    pdf_path: str | Path,
    out_path: str | Path,
    cfg: StampConfig | None = None,
    pages: list[int] | None = None,
) -> str:
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


def _main() -> None:
    import argparse
    import json
    from datetime import date as _date

    parser = argparse.ArgumentParser(description="Split a batch PDF by incoming stamp.")
    parser.add_argument("pdf", help="Path to the batch PDF")
    parser.add_argument("--out", default="work/split", help="Output folder for letters")
    parser.add_argument("--review", default="review", help="Review folder (fail-safe)")
    parser.add_argument(
        "--resolution", type=int, default=150, help="Render DPI (default 150)"
    )
    parser.add_argument(
        "--date",
        default=_date.today().strftime("%d_%m_%Y"),
        help="Date subfolder (default: today, DD_MM_YYYY). Letters go to <out>/<date>/.",
    )
    parser.add_argument(
        "--no-date-subdir",
        action="store_true",
        help="Do not create a date subfolder (write directly into <out>).",
    )
    parser.add_argument(
        "--debug-image", metavar="PNG", help="Write a debug image and exit"
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG logging")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )
    cfg = StampConfig(resolution=args.resolution)

    if args.debug_image:
        path = write_debug_image(args.pdf, args.debug_image, cfg)
        print(f"debug image written: {path}")
        return

    out_dir = args.out if args.no_date_subdir else str(Path(args.out) / args.date)
    log.info("output folder: %s", out_dir)
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
