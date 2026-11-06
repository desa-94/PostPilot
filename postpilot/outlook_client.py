"""
outlook_client.py — Outlook-Anbindung via pywin32 (Windows).

Aktueller Umfang (fuer den TAX-Versand heute)
---------------------------------------------
send_message(): Erzeugt eine Outlook-Mail mit Anhang und entweder
  - dry_run=True  (Default): speichert sie als ENTWURF (landet im Ordner
    "Entwuerfe"/"Drafts") -> du kannst sie in Outlook pruefen, bevor etwas
    rausgeht. Nichts wird versendet.
  - dry_run=False: versendet sofort (.Send()).

Das Lesen des ScanToMe-Ordners (Anhaenge abholen) kommt in einem spaeteren
Schritt — heute liegen die gesplitteten PDFs bereits lokal vor.

Abhaengigkeit: pywin32 (win32com.client). Nur unter Windows mit installiertem
Outlook lauffaehig.
"""

from __future__ import annotations

import logging
from collections.abc import Sequence
from pathlib import Path

log = logging.getLogger("postpilot.outlook_client")

# Outlook-MailItem-Konstante
_OL_MAIL_ITEM = 0


def _get_outlook():
    """Verbindet sich mit der laufenden/instanziierten Outlook-Applikation."""
    try:
        import win32com.client  # lokaler Import: Modul laedt auch ohne pywin32
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pywin32 wird benoetigt: pip install pywin32 (nur Windows + Outlook)"
        ) from exc
    return win32com.client.Dispatch("Outlook.Application")


def send_message(
    to: str,
    subject: str,
    body: str,
    attachments: Sequence[str | Path] | None = None,
    *,
    dry_run: bool = True,
) -> dict:
    """
    Erzeugt eine Mail mit Anhaengen.

    dry_run=True  -> als Entwurf speichern (nichts geht raus). Default.
    dry_run=False -> versenden.

    Returns: dict mit Status (fuer Log/Report), wirft bei echten Fehlern.
    """
    attachments = attachments or []
    # Anhaenge vorab pruefen — fehlende Datei soll NICHT stillschweigend
    # zu einer Mail ohne Anhang fuehren.
    missing = [str(a) for a in attachments if not Path(a).is_file()]
    if missing:
        raise FileNotFoundError(f"Anhang fehlt: {missing}")

    outlook = _get_outlook()
    mail = outlook.CreateItem(_OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    for a in attachments:
        mail.Attachments.Add(str(Path(a).resolve()))

    action = "versendet" if not dry_run else "als Entwurf gespeichert"
    if dry_run:
        mail.Save()  # landet im Entwuerfe-Ordner
    else:
        mail.Send()

    log.info(
        "Mail %s -> %s | Betreff: %s | Anhaenge: %d",
        action,
        to,
        subject,
        len(attachments),
    )
    return {
        "to": to,
        "subject": subject,
        "attachments": [str(a) for a in attachments],
        "dry_run": dry_run,
        "status": action,
    }


# --------------------------------------------------------------------------- #
# Kleiner Selbsttest (nur manuell, erzeugt einen harmlosen Entwurf)
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Outlook-Testentwurf erzeugen.")
    parser.add_argument("--to", required=True)
    parser.add_argument("--subject", default="PostPilot Test")
    parser.add_argument("--body", default="Testentwurf von PostPilot.")
    parser.add_argument("--attach", nargs="*", default=[])
    parser.add_argument(
        "--send", action="store_true", help="WIRKLICH senden (sonst Entwurf)"
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO, format="%(levelname)s %(name)s: %(message)s"
    )
    res = send_message(
        args.to, args.subject, args.body, args.attach, dry_run=not args.send
    )
    print(res)


if __name__ == "__main__":
    _main()
