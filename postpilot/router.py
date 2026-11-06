"""
router.py — Verbindet Extraktion, Klassifizierung und Versand.

Ablauf pro Einzelbrief-PDF:
  1. pdf_extract.extract()  -> Felder (absender, datum, empfaenger, volltext)
  2. rule_engine.classify() -> Abteilung + Score + treffer (Nachvollziehbarkeit)
  3. auto_versand == False  -> PDF nach review/ (kein Versand)
     auto_versand == True   -> Betreff/Body aus der Abteilungs-Vorlage rendern
                               und Mail erzeugen (Entwurf oder Versand).

batch_mode der Abteilung:
  TAX hat batch_mode: true -> jeder Brief wird EINZELN verschickt (eine Mail
  pro PDF). Das ist genau der heutige TAX-Workflow.

Versand-Sicherheit:
  dry_run=True (Default) -> Outlook-Entwuerfe zum Pruefen. Nichts geht raus.
  dry_run=False          -> echter Versand.

Template-Platzhalter: {absender}, {datum}, {empfaenger}, {abteilung}, {score}.
Fehlt ein Platzhalter im Brief, wird er durch "" ersetzt (kein Absturz).
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from postpilot import pdf_extract, rule_engine
from postpilot.rule_engine import Klassifikation, RuleConfig

log = logging.getLogger("postpilot.router")


def default_config_path() -> str:
    """
    Findet die zu nutzende Config automatisch:
      1. config/departments.local.yaml  (echte Adressen, falls vorhanden)
      2. config/departments.yaml        (Platzhalter-Vorlage)
    So muss im Alltag kein --config mehr getippt werden.
    """
    local = Path("config/departments.local.yaml")
    if local.is_file():
        return str(local)
    return "config/departments.yaml"


def _select_backend(backend: str):
    """
    Waehlt das Mail-Backend und gibt (send_message-Funktion, name) zurueck.

    backend:
      "auto"    -> Windows: outlook, sonst: eml (plattformunabhaengig)
      "outlook" -> erzwingt outlook_client (nur Windows + pywin32)
      "eml"     -> erzwingt eml_writer (.eml-Dateien, ueberall)

    Import erfolgt lazy, damit der Router auch ohne pywin32 importierbar
    bleibt (z.B. fuer reine Klassifizierungs-Tests auf Linux).
    """
    import sys

    choice = backend
    if choice == "auto":
        choice = "outlook" if sys.platform.startswith("win") else "eml"

    if choice == "outlook":
        from postpilot import outlook_client

        return outlook_client.send_message, "outlook"
    if choice == "eml":
        from postpilot import eml_writer

        return eml_writer.send_message, "eml"
    raise ValueError(f"Unbekanntes Backend: {backend!r} (erlaubt: auto|outlook|eml)")


class _SafeDict(dict):
    """Fehlender Platzhalter -> leerer String statt KeyError."""

    def __missing__(self, key):
        log.warning("Template-Platzhalter {%s} nicht gefunden -> leer", key)
        return ""


def _render(template: str, k: Klassifikation) -> str:
    values = _SafeDict(
        absender=k.absender,
        datum=k.datum,
        empfaenger=k.empfaenger,
        abteilung=k.abteilung,
        score=k.score,
    )
    return template.format_map(values)


def _move_to_review(pdf: Path, review_dir: Path, grund: str) -> str:
    review_dir.mkdir(parents=True, exist_ok=True)
    dest = review_dir / pdf.name
    counter = 1
    while dest.exists():
        dest = review_dir / f"{pdf.stem}__{counter}{pdf.suffix}"
        counter += 1
    shutil.copy2(pdf, dest)
    log.warning("-> review/: %s (%s)", pdf.name, grund)
    return str(dest)


def route_one(
    pdf: str | Path,
    cfg: RuleConfig,
    *,
    review_dir: str | Path = "review",
    dry_run: bool = True,
    ocr_fallback: bool = False,
    backend: str = "auto",
) -> dict:
    """
    Verarbeitet EIN Einzelbrief-PDF: extrahieren -> klassifizieren ->
    senden/pruefen. Gibt einen nachvollziehbaren Report-Eintrag zurueck.
    """
    pdf = Path(pdf)
    ext = pdf_extract.extract(pdf, ocr_fallback=ocr_fallback)
    ext_d = {
        "absender": ext.absender,
        "volltext": ext.volltext,
        "empfaenger": ext.empfaenger,
        "datum": ext.datum,
        "datei": ext.source,
    }
    k = rule_engine.classify(ext_d, cfg)

    entry: dict = {
        "datei": pdf.name,
        "abteilung": k.abteilung,
        "score": k.score,
        "absender": k.absender,
        "datum": k.datum,
        "treffer": [t.as_dict() for t in k.treffer],
        "auto_versand": k.auto_versand,
    }

    # Fail-safe: kein eindeutiger Gewinner -> review/
    if not k.auto_versand:
        entry["aktion"] = "review"
        entry["grund"] = k.grund
        entry["review_path"] = _move_to_review(pdf, Path(review_dir), k.grund)
        return entry

    dept = rule_engine.department_by_name(cfg, k.abteilung)
    if not dept or not dept.get("email"):
        entry["aktion"] = "review"
        entry["grund"] = f"keine E-Mail-Adresse fuer Abteilung {k.abteilung}"
        entry["review_path"] = _move_to_review(pdf, Path(review_dir), entry["grund"])
        return entry

    subject = _render(dept.get("subject_template", "Weiterleitung: {absender}"), k)
    body = _render(dept.get("body_template", ""), k)

    # Zieladresse: Subroute-Adresse hat Vorrang vor der Abteilungs-Default-Adresse.
    # (classify() hat bei unklarer Subroute bereits auto_versand=False gesetzt,
    #  d.h. hier kommen nur Briefe an, deren Subroute sicher erkannt wurde.)
    ziel = k.subroute_email or dept["email"]

    # Backend lazy auswaehlen: outlook (Windows) oder eml (plattformunabhaengig).
    # Import erst hier, damit der Router auch ohne pywin32 importierbar bleibt.
    sender, backend_name = _select_backend(backend)
    res = sender(
        to=ziel,
        subject=subject,
        body=body,
        attachments=[pdf],
        dry_run=dry_run,
    )
    entry["backend"] = backend_name
    entry["aktion"] = res.get("status", "erzeugt")
    entry["an"] = ziel
    entry["subroute"] = k.subroute
    entry["betreff"] = subject
    if res.get("eml_path"):
        entry["eml_path"] = res["eml_path"]
    return entry


def route_folder(
    folder: str | Path,
    config_path: str | Path = "config/departments.yaml",
    *,
    review_dir: str | Path = "review",
    dry_run: bool = True,
    only_department: str | None = None,
    ocr_fallback: bool = False,
    backend: str = "auto",
) -> list[dict]:
    """
    Verarbeitet alle PDFs in einem Ordner.

    only_department: wenn gesetzt (z.B. "TAX"), werden NUR Briefe dieser
      Abteilung versendet; alle anderen bleiben unangetastet liegen (nicht
      nach review/ verschoben). Perfekt fuer "heute nur TAX rausschicken".
    """
    cfg = RuleConfig.load(config_path)
    folder = Path(folder)
    pdfs = sorted(folder.glob("*.pdf"))
    report: list[dict] = []

    for pdf in pdfs:
        # Vorab klassifizieren, um only_department zu respektieren, ohne
        # fremde Briefe anzufassen.
        if only_department:
            ext = pdf_extract.extract(pdf, ocr_fallback=ocr_fallback)
            ext_d = {
                "absender": ext.absender,
                "volltext": ext.volltext,
                "empfaenger": ext.empfaenger,
                "datum": ext.datum,
                "datei": ext.source,
            }
            k = rule_engine.classify(ext_d, cfg)
            # Nur ECHTE Fremd-Abteilungen ueberspringen. TAX-Briefe mit
            # unklarer Subroute (abteilung == only_department, aber
            # auto_versand == False) fallen bewusst durch zu route_one,
            # das sie nach review/ verschiebt — statt sie liegen zu lassen.
            if k.abteilung != only_department:
                log.info(
                    "uebersprungen (nicht %s): %s [%s, auto=%s]",
                    only_department,
                    pdf.name,
                    k.abteilung,
                    k.auto_versand,
                )
                report.append(
                    {
                        "datei": pdf.name,
                        "aktion": "uebersprungen",
                        "abteilung": k.abteilung,
                        "auto_versand": k.auto_versand,
                    }
                )
                continue

        report.append(
            route_one(
                pdf,
                cfg,
                review_dir=review_dir,
                dry_run=dry_run,
                ocr_fallback=ocr_fallback,
                backend=backend,
            )
        )

    return report


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="Briefe klassifizieren und weiterleiten."
    )
    parser.add_argument("folder", help="Ordner mit Einzelbrief-PDFs (z.B. work/split)")
    parser.add_argument(
        "--config",
        default=None,
        help="Config-Datei. Default: auto (departments.local.yaml falls "
        "vorhanden, sonst departments.yaml).",
    )
    parser.add_argument("--review", default="review")
    parser.add_argument("--only", metavar="ABT", help="nur diese Abteilung (z.B. TAX)")
    parser.add_argument(
        "--send",
        action="store_true",
        help="WIRKLICH senden. Ohne dieses Flag: nur Entwuerfe (Dry-Run).",
    )
    parser.add_argument("--ocr", action="store_true", help="OCR-Fallback aktivieren")
    parser.add_argument(
        "--backend",
        choices=["auto", "outlook", "eml"],
        default="auto",
        help="Mail-Backend: auto (Windows->outlook, sonst->eml), "
        "outlook (nur Windows), eml (.eml-Dateien, plattformunabhaengig). "
        "Default: auto.",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )

    report = route_folder(
        args.folder,
        args.config or default_config_path(),
        review_dir=args.review,
        dry_run=not args.send,
        only_department=args.only,
        ocr_fallback=args.ocr,
        backend=args.backend,
    )

    # Zusammenfassung
    modus = "VERSAND" if args.send else "DRY-RUN (Entwuerfe)"
    n_mail = sum(1 for r in report if r.get("aktion") in ("entwurf", "versendet"))
    n_review = sum(1 for r in report if r.get("aktion") == "review")
    n_skip = sum(1 for r in report if r.get("aktion") == "uebersprungen")
    print(json.dumps(report, indent=2, ensure_ascii=False))
    print(f"\n[{modus}] {n_mail} Mail(s), {n_review} -> review, {n_skip} uebersprungen")


if __name__ == "__main__":
    _main()
