"""
eml_writer.py — Plattformunabhaengiges Mail-Backend (.eml-Dateien).

Zweck
-----
Alternative zu outlook_client fuer Umgebungen ohne Outlook/pywin32 (Linux,
CI, lokale Tests). Statt eine Outlook-Mail zu erzeugen, schreibt dieses
Backend eine standardkonforme .eml-Datei (RFC 822 / MIME) mit Anhang.

.eml-Dateien lassen sich in praktisch jedem Mailprogramm (Thunderbird,
Outlook, Apple Mail, Gmail-Import) oeffnen und als Entwurf/Vorlage versenden.
So kann die komplette Pipeline — Splitten, Extraktion, Klassifizierung,
Routing, Mail-Erzeugung — auf Linux entwickelt und getestet werden; nur der
tatsaechliche Outlook-Versand bleibt Windows vorbehalten.

Bewusste Grenze
---------------
Dieses Backend VERSENDET NICHT. Es erzeugt immer nur eine Datei — unabhaengig
vom dry_run-Flag (das Flag wird akzeptiert, damit die Signatur identisch zu
outlook_client.send_message ist, austauschbar als Backend). Kein SMTP, kein
Netzwerk — passt zur Datenschutz-Leitplanke (nichts verlaesst den Rechner).

Signatur ist absichtlich identisch zu outlook_client.send_message(), damit
der Router beide Backends austauschbar aufrufen kann.

Abhaengigkeiten: nur Python-Standardbibliothek (email, pathlib, mimetypes).
"""

from __future__ import annotations

import logging
import mimetypes
import re
from collections.abc import Sequence
from datetime import UTC, datetime
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

log = logging.getLogger("postpilot.eml_writer")

# Wohin die .eml-Dateien geschrieben werden (vom Router ueberschreibbar).
DEFAULT_OUTBOX = "work/outbox"


def _safe_filename(text: str, maxlen: int = 80) -> str:
    """Macht aus einem Betreff einen dateisystem-sicheren Namen."""
    text = re.sub(r"[^\w\-. ]", "_", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "_", text.strip())
    return text[:maxlen] or "mail"


def send_message(
    to: str,
    subject: str,
    body: str,
    attachments: Sequence[str | Path] | None = None,
    *,
    dry_run: bool = True,
    outbox: str | Path = DEFAULT_OUTBOX,
    sender: str = "postpilot@localhost",
) -> dict:
    """
    Schreibt eine Mail als .eml-Datei in den outbox-Ordner.

    Signatur identisch zu outlook_client.send_message(); 'dry_run' wird
    akzeptiert, hat hier aber keine Wirkung (es wird nie versendet, immer
    nur eine Datei erzeugt).

    Returns: dict mit Status (fuer Log/Report), wirft bei fehlenden Anhaengen.
    """
    attachments = attachments or []
    missing = [str(a) for a in attachments if not Path(a).is_file()]
    if missing:
        raise FileNotFoundError(f"Anhang fehlt: {missing}")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg.set_content(body)

    for a in attachments:
        a = Path(a)
        ctype, encoding = mimetypes.guess_type(str(a))
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            a.read_bytes(), maintype=maintype, subtype=subtype, filename=a.name
        )

    outbox = Path(outbox)
    outbox.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S")
    fname = f"{stamp}_{_safe_filename(subject)}.eml"
    dest = outbox / fname
    counter = 1
    while dest.exists():
        dest = outbox / f"{stamp}_{_safe_filename(subject)}__{counter}.eml"
        counter += 1
    dest.write_bytes(bytes(msg))

    log.info(
        "EML geschrieben -> %s | An: %s | Betreff: %s | Anhaenge: %d",
        dest,
        to,
        subject,
        len(attachments),
    )
    return {
        "to": to,
        "subject": subject,
        "attachments": [str(a) for a in attachments],
        "dry_run": dry_run,
        "status": "eml_geschrieben",
        "eml_path": str(dest),
    }


# --------------------------------------------------------------------------- #
# CLI (kleiner Selbsttest)
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Test-.eml erzeugen.")
    parser.add_argument("--to", required=True)
    parser.add_argument("--subject", default="PostPilot Test")
    parser.add_argument("--body", default="Test-EML von PostPilot.")
    parser.add_argument("--attach", nargs="*", default=[])
    parser.add_argument("--outbox", default=DEFAULT_OUTBOX)
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO, format="%(levelname)s %(name)s: %(message)s"
    )
    res = send_message(
        args.to, args.subject, args.body, args.attach, outbox=args.outbox
    )
    print(res)


if __name__ == "__main__":
    _main()
