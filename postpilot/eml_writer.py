"""Cross-platform mail backend (.eml files).

An alternative to outlook_client for environments without Outlook/pywin32
(Linux, CI, local tests). Instead of creating an Outlook mail it writes a
standard .eml file (RFC 822 / MIME) with the attachment, openable in any mail
client. This never sends and never touches the network - it only writes a file,
regardless of the dry_run flag (accepted to keep the signature identical to
outlook_client.send_message so the router can swap backends freely).

Dependencies: Python standard library only (email, pathlib, mimetypes).
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

DEFAULT_OUTBOX = "work/outbox"


def _safe_filename(text: str, maxlen: int = 80) -> str:
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
    attachments = attachments or []
    missing = [str(a) for a in attachments if not Path(a).is_file()]
    if missing:
        raise FileNotFoundError(f"attachment missing: {missing}")

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
        "eml written -> %s | to: %s | subject: %s | attachments: %d",
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
        "status": "eml_written",
        "eml_path": str(dest),
    }


def _main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Create a test .eml file.")
    parser.add_argument("--to", required=True)
    parser.add_argument("--subject", default="PostPilot Test")
    parser.add_argument("--body", default="PostPilot test eml.")
    parser.add_argument("--attach", nargs="*", default=[])
    parser.add_argument("--outbox", default=DEFAULT_OUTBOX)
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    res = send_message(args.to, args.subject, args.body, args.attach, outbox=args.outbox)
    print(res)


if __name__ == "__main__":
    _main()
