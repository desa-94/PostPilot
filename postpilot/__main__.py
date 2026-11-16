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
  python -m postpilot --fetch              # PDFs aus Outlook-Ordner ScanToMe
                                           #   holen und alle verarbeiten
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


def scantome_settings() -> dict:
    """
    Liest den optionalen 'scantome'-Abschnitt aus der (auto-gefundenen) Config.

    Returns: {"folder": <str|None>, "only_unread": <bool>} — Defaults, wenn
    kein scantome-Abschnitt existiert.
    """
    import yaml

    try:
        data = yaml.safe_load(Path(default_config_path()).read_text(encoding="utf-8")) or {}
    except Exception:
        return {"folder": None, "only_unread": True}
    sec = data.get("scantome") or {}
    return {"folder": sec.get("folder"), "only_unread": bool(sec.get("only_unread", True))}


def fetch_from_scantome(
    *,
    folder_path: str = "ScanToMe",
    inbox_base: str | Path = "work/inbox",
    day: str | None = None,
    only_unread: bool = True,
) -> list[str]:
    """
    Holt PDF-Anhaenge aus dem Outlook-Ordner ScanToMe nach work/inbox/<datum>/.

    Returns: Liste der lokal gespeicherten PDF-Pfade (kann leer sein).
    Import von outlook_client erfolgt lazy (nur Windows + pywin32).
    """
    from postpilot import outlook_client

    day = day or date.today().strftime("%d_%m_%Y")
    inbox = Path(inbox_base) / day
    rows = outlook_client.fetch_pdf_attachments(
        inbox, folder_path=folder_path, only_unread=only_unread
    )
    paths = [r["pdf_path"] for r in rows]
    log.info("ScanToMe: %d PDF(s) geholt -> %s", len(paths), inbox)
    return paths


def _main() -> None:
    parser = argparse.ArgumentParser(
        prog="postpilot",
        description="PostPilot: Sammel-PDF splitten und TAX-Briefe weiterleiten (ein Lauf).",
    )
    parser.add_argument("--pdf", default="TAX.pdf", help="Quell-Sammel-PDF (Default: TAX.pdf)")
    parser.add_argument(
        "--fetch", action="store_true",
        help="PDF-Anhaenge aus dem Outlook-Ordner ScanToMe holen und alle "
             "verarbeiten (statt der lokalen --pdf-Datei).",
    )
    parser.add_argument("--folder", default=None, help="Outlook-Ordner fuer --fetch (Default: aus Config scantome.folder)")
    parser.add_argument("--all", action="store_true", help="bei --fetch auch gelesene Mails holen")
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

    # Quell-PDFs bestimmen: aus ScanToMe holen oder lokale --pdf-Datei.
    if args.fetch:
        st = scantome_settings()
        folder = args.folder or st.get("folder")
        if not folder:
            parser.error(
                "Kein ScanToMe-Ordner konfiguriert. Entweder --folder angeben "
                "oder scantome.folder in der Config setzen "
                "(py -m postpilot.outlook_client --list-folders zeigt die Ordner)."
            )
        only_unread = st.get("only_unread", True) and not args.all
        sources = fetch_from_scantome(
            folder_path=folder, day=args.date, only_unread=only_unread
        )
        if not sources:
            print("\n[ScanToMe] Keine neuen PDF-Anhaenge gefunden.")
            return
    else:
        sources = [args.pdf]

    # Jede Quelle verarbeiten; Ergebnisse aggregieren. Bei --fetch je Quelle
    # ein eigener Datums-Unterordner, damit mehrere Sammel-PDFs nicht kollidieren.
    results = []
    for idx, src in enumerate(sources, 1):
        day = args.date
        if args.fetch and len(sources) > 1:
            base_day = args.date or date.today().strftime("%d_%m_%Y")
            day = f"{base_day}/{idx:02d}"
        results.append((src, run(
            src, out_base=args.out, review_dir=args.review, day=day,
            only_department=args.only, send=args.send, backend=args.backend,
            resolution=args.resolution,
        )))

    modus = "VERSAND" if args.send else "DRY-RUN (Entwuerfe)"
    tot_letters = tot_mail = tot_review = tot_skip = 0
    for src, result in results:
        if not result["split_ok"]:
            print(f"[ABBRUCH] {Path(src).name}: {result['reason']} -> {result.get('review_path')}")
            continue
        report = result["routed"]
        tot_letters += result["letters"]
        tot_mail += sum(1 for r in report if r.get("aktion") in ("entwurf", "versendet", "eml_geschrieben", "erzeugt"))
        tot_review += sum(1 for r in report if r.get("aktion") == "review")
        tot_skip += sum(1 for r in report if r.get("aktion") == "uebersprungen")
    print(f"\n[{modus}] {len(results)} Quelle(n) | {tot_letters} Briefe gesplittet | "
          f"{tot_mail} Mail(s), {tot_review} -> review, {tot_skip} uebersprungen")
    if not args.send:
        print("Hinweis: Dry-Run. Zum echten Versand erneut mit --send aufrufen.")


if __name__ == "__main__":
    _main()
