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


def _resolve_folder(namespace, folder_path: str):
    """
    Findet einen Outlook-Ordner anhand eines Pfades.

    folder_path:
      - "ScanToMe"                 -> sucht rekursiv im Standard-Posteingang
      - "Posteingang/ScanToMe"     -> expliziter Pfad ab Postfach-Wurzel
      - "Postfach/Posteingang/Sub" -> mit fuehrendem Postfachnamen

    Returns den Folder oder wirft, wenn nicht gefunden (mit Hinweis).
    """
    parts = [p for p in folder_path.replace("\\", "/").split("/") if p]

    # Fall 1: einfacher Name -> rekursiv ab Standard-Posteingang suchen.
    if len(parts) == 1:
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        if inbox.Name.lower() == parts[0].lower():
            return inbox
        found = _find_subfolder(inbox, parts[0])
        if found is not None:
            return found
        # Auch auf oberster Ebene des Standard-Stores suchen.
        try:
            root = inbox.Parent  # Store-Root
            found = _find_subfolder(root, parts[0])
            if found is not None:
                return found
        except Exception:
            pass
        raise RuntimeError(
            f"Outlook-Ordner '{folder_path}' nicht gefunden. Tipp: expliziten "
            f"Pfad angeben, z.B. 'Posteingang/{parts[0]}'."
        )

    # Fall 2: expliziter Pfad -> Schritt fuer Schritt absteigen.
    # Startpunkt: passt der erste Teil auf einen Store/Postfachnamen?
    current = None
    for store in namespace.Folders:
        if store.Name.lower() == parts[0].lower():
            current = store
            parts = parts[1:]
            break
    if current is None:
        # Kein Postfachname vorangestellt -> ab Standard-Posteingang-Parent.
        current = namespace.GetDefaultFolder(6).Parent

    for part in parts:
        nxt = _find_subfolder(current, part, recursive=False)
        if nxt is None:
            raise RuntimeError(
                f"Outlook-Ordner '{folder_path}' nicht gefunden (Teil '{part}' fehlt)."
            )
        current = nxt
    return current


def _find_subfolder(folder, name: str, recursive: bool = True):
    """Sucht einen Unterordner per Name (case-insensitive). None, wenn nicht da."""
    try:
        subfolders = folder.Folders
    except Exception:
        return None
    for sub in subfolders:
        if sub.Name.lower() == name.lower():
            return sub
    if recursive:
        for sub in subfolders:
            hit = _find_subfolder(sub, name, recursive=True)
            if hit is not None:
                return hit
    return None


def list_all_folders(max_depth: int = 4) -> list[str]:
    """
    Listet alle Outlook-Ordner (ueber alle Postfaecher/Stores) als Pfade auf.

    Diagnose-Helfer: zeigt die echten Ordnernamen, damit der richtige
    folder_path fuer fetch_pdf_attachments gefunden werden kann (z.B. wenn
    "ScanToMe" in einem freigegebenen/zweiten Postfach liegt).

    Returns: Liste von Pfaden wie "Postfachname/Posteingang/ScanToMe".
    """
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    paths: list[str] = []

    def walk(folder, prefix: str, depth: int):
        name = f"{prefix}/{folder.Name}" if prefix else folder.Name
        paths.append(name)
        if depth >= max_depth:
            return
        try:
            for sub in folder.Folders:
                walk(sub, name, depth + 1)
        except Exception:
            pass

    for store in namespace.Folders:
        walk(store, "", 0)
    return paths


def fetch_pdf_attachments(
    out_dir: str | Path,
    *,
    folder_path: str = "ScanToMe",
    only_unread: bool = True,
    mark_read: bool = True,
) -> list[dict]:

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    folder = _resolve_folder(namespace, folder_path)
    total = folder.Items.Count

    results: list[dict] = []
    # Bei only_unread serverseitig filtern (Items.Restrict) statt alle Mails
    # zu laden. Das ist bei grossen Ordnern (Tausende Mails) drastisch
    # schneller — Outlook liefert nur die ungelesenen zurueck.
    if only_unread:
        try:
            filtered = folder.Items.Restrict("[UnRead] = True")
            items = list(filtered)
            log.info("Lese Ordner '%s' (%d ungelesen von %d gesamt)",
                     folder.Name, len(items), total)
        except Exception as exc:
            log.warning("Restrict fehlgeschlagen (%s) — lese alle %d Elemente.", exc, total)
            items = list(folder.Items)
    else:
        items = list(folder.Items)
        log.info("Lese Ordner '%s' (%d Elemente)", folder.Name, total)
    for item in items:
        # Nur echte Mails (olMail hat MessageClass IPM.Note*)
        try:
            if only_unread and not item.UnRead:
                continue
        except Exception:
            continue  # Nicht-Mail-Elemente (Termine o.ae.) ueberspringen

        entry_id = getattr(item, "EntryID", "")
        subject = getattr(item, "Subject", "") or ""
        received = str(getattr(item, "ReceivedTime", "") or "")
        sender = getattr(item, "SenderName", "") or ""

        saved_any = False
        try:
            for att in item.Attachments:
                fname = att.FileName or ""
                if not fname.lower().endswith(".pdf"):
                    continue
                # Eindeutiger lokaler Name: EntryID-Prefix gegen Kollisionen.
                safe = f"{entry_id[:12]}_{fname}" if entry_id else fname
                dest = out_dir / safe
                counter = 1
                while dest.exists():
                    dest = out_dir / f"{entry_id[:12]}_{counter}_{fname}"
                    counter += 1
                att.SaveAsFile(str(dest.resolve()))
                results.append(
                    {
                        "pdf_path": str(dest),
                        "message_id": entry_id,
                        "subject": subject,
                        "received": received,
                        "sender": sender,
                    }
                )
                saved_any = True
                log.info("Anhang gespeichert: %s (Mail: %s)", dest.name, subject[:50])
        except Exception as exc:
            log.warning("Fehler beim Lesen einer Mail '%s': %s", subject[:50], exc)
            continue

        if saved_any and mark_read:
            try:
                item.UnRead = False
                item.Save()
            except Exception as exc:  # pragma: no cover
                log.warning("Konnte Mail nicht als gelesen markieren: %s", exc)

    log.info("ScanToMe: %d PDF-Anhang/Anhaenge gespeichert.", len(results))
    return results


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
    parser.add_argument(
        "--list-folders", action="store_true",
        help="Alle Outlook-Ordner auflisten (Diagnose) und beenden.",
    )
    parser.add_argument(
        "--fetch",
        metavar="OUTDIR",
        help="Statt Entwurf: PDF-Anhaenge aus ScanToMe nach OUTDIR holen.",
    )
    parser.add_argument(
        "--folder", default="ScanToMe", help="Outlook-Ordner (Default ScanToMe)"
    )
    parser.add_argument("--all", action="store_true", help="auch gelesene Mails holen")
    parser.add_argument("--to", help="Empfaenger (nur fuer Entwurf/Versand noetig)")
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
    if args.list_folders:
        for p in list_all_folders():
            print(p)
        return
    if args.fetch:
        rows = fetch_pdf_attachments(
            args.fetch, folder_path=args.folder, only_unread=not args.all
        )
        for r in rows:
            print(r["pdf_path"])
        print(f"\n{len(rows)} PDF-Anhang/Anhaenge gespeichert.")
        return
    if not args.to:
        parser.error("--to ist erforderlich (ausser bei --fetch)")
    res = send_message(
        args.to, args.subject, args.body, args.attach, dry_run=not args.send
    )
    print(res)


if __name__ == "__main__":
    _main()
