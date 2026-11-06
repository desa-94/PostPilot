"""
__main__.py — Orchestrierung: ``python -m postpilot`` macht alles in einem Lauf.

Ablauf (Standard, fuer den taeglichen TAX-Betrieb):
  1. TAX.pdf anhand der Eingangsstempel in Einzelbriefe splitten
     -> work/split/<TT_MM_JJJJ>/ (Datum = heute, ueberschreibbar via --date)
  2. Diese Einzelbriefe klassifizieren und weiterleiten (nur TAX)
     -> Dry-Run (Outlook-Entwuerfe bzw. .eml), echter Versand nur mit --send

Vereinfachungen gegenueber den Einzel-CLIs:
  - EIN Befehl statt zwei, kein manuelles Wiederholen des Datums-Ordners.
  - Config wird automatisch gefunden (departments.local.yaml, sonst .yaml).
  - Sinnvolle Defaults: Quelle TAX.pdf, nur Abteilung TAX, Backend auto.

Beispiele:
  python -m postpilot                      # heute splitten + Dry-Run
  python -m postpilot --send               # heute splitten + WIRKLICH senden
  python -m postpilot --date 06_11_2026    # fuer ein anderes Datum
  python -m postpilot --pdf Sammel.pdf     # andere Quelldatei
"""

from __future__ import annotations

import argparse
import logging
from datetime import date
from pathlib import Path

from postpilot import router
from postpilot.router import default_config_path
from postpilot.stamp_splitter import StampConfig, split_by_stamps

log = logging.getLogger("postpilot.main")


def run(
    pdf: str | Path = "TAX.pdf",
    *,
    out_base: str | Path = "work/split",
    review_dir: str | Path = "review",
    day: str | None = None,
    only_department: str | None = "TAX",
    send: bool = False,
    backend: str = "auto",
    resolution: int = 150,
) -> dict:
    """
    Fuehrt Split + Routing in einem Lauf aus.

    Returns: {"split": SplitResult-dict-artig, "routed": [report...]}.
    """
    day = day or date.today().strftime("%d_%m_%Y")
    out_dir = Path(out_base) / day

    # --- 1) Splitten -----------------------------------------------------
    log.info("Splitte %s -> %s", pdf, out_dir)
    split = split_by_stamps(pdf, out_dir, review_dir, StampConfig(resolution=resolution))
    if not split.ok:
        log.warning("Splitting nicht erfolgt (%s) — ganzes PDF in review/.", split.reason)
        return {"split_ok": False, "reason": split.reason,
                "review_path": split.review_path, "routed": []}
    log.info("%d Einzelbriefe erzeugt.", len(split.letters))

    # --- 2) Routen -------------------------------------------------------
    cfg_path = default_config_path()
    log.info("Route %s (Config: %s, Backend: %s, %s)",
             out_dir, cfg_path, backend, "VERSAND" if send else "DRY-RUN")
    report = router.route_folder(
        out_dir, cfg_path,
        review_dir=review_dir,
        dry_run=not send,
        only_department=only_department,
        backend=backend,
    )
    return {"split_ok": True, "letters": len(split.letters), "routed": report}


def _main() -> None:
    parser = argparse.ArgumentParser(
        prog="postpilot",
        description="PostPilot: Sammel-PDF splitten und TAX-Briefe weiterleiten (ein Lauf).",
    )
    parser.add_argument("--pdf", default="TAX.pdf", help="Quell-Sammel-PDF (Default: TAX.pdf)")
    parser.add_argument("--date", default=None, help="Datums-Ordner TT_MM_JJJJ (Default: heute)")
    parser.add_argument("--out", default="work/split", help="Basis-Ausgabeordner")
    parser.add_argument("--review", default="review", help="Pruefungsordner")
    parser.add_argument("--only", default="TAX", help="nur diese Abteilung (Default: TAX)")
    parser.add_argument(
        "--send", action="store_true",
        help="WIRKLICH senden. Ohne dieses Flag: nur Entwuerfe/Dry-Run.",
    )
    parser.add_argument(
        "--backend", choices=["auto", "outlook", "eml"], default="auto",
        help="Mail-Backend (Default: auto — Windows->outlook, sonst->eml).",
    )
    parser.add_argument("--resolution", type=int, default=150, help="Render-DPI (Default 150)")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )

    result = run(
        args.pdf, out_base=args.out, review_dir=args.review, day=args.date,
        only_department=args.only, send=args.send, backend=args.backend,
        resolution=args.resolution,
    )

    # Zusammenfassung
    if not result["split_ok"]:
        print(f"\n[ABBRUCH] Splitting nicht moeglich: {result['reason']}")
        print(f"          Ganzes PDF -> {result.get('review_path')}")
        return

    report = result["routed"]
    modus = "VERSAND" if args.send else "DRY-RUN (Entwuerfe)"
    n_mail = sum(1 for r in report if r.get("aktion") in ("entwurf", "versendet", "eml_geschrieben", "erzeugt"))
    n_review = sum(1 for r in report if r.get("aktion") == "review")
    n_skip = sum(1 for r in report if r.get("aktion") == "uebersprungen")
    print(f"\n[{modus}] {result['letters']} Briefe gesplittet | "
          f"{n_mail} Mail(s), {n_review} -> review, {n_skip} uebersprungen")
    if not args.send:
        print("Hinweis: Dry-Run. Zum echten Versand erneut mit --send aufrufen.")


if __name__ == "__main__":
    _main()
