"""Outlook integration via pywin32 (Windows).

send_message() creates an Outlook mail with attachments: as a draft when
dry_run=True (default, nothing leaves the machine), or sent when dry_run=False.

fetch_pdf_attachments() reads new (unread) mails from a folder, saves their PDF
attachments locally and returns one record per attachment (path, message id,
subject). Mails are marked read only after their attachments were saved.

Requires pywin32 (win32com.client); Windows with Outlook only.
"""

from __future__ import annotations

import logging
from collections.abc import Sequence
from pathlib import Path

log = logging.getLogger("postpilot.outlook_client")

_OL_MAIL_ITEM = 0


def _get_outlook():
    try:
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "pywin32 required: pip install pywin32 (Windows + Outlook only)"
        ) from exc
    return win32com.client.Dispatch("Outlook.Application")


def _resolve_folder(namespace, folder_path: str):
    parts = [p for p in folder_path.replace("\\", "/").split("/") if p]

    if len(parts) == 1:
        inbox = namespace.GetDefaultFolder(6)
        if inbox.Name.lower() == parts[0].lower():
            return inbox
        found = _find_subfolder(inbox, parts[0])
        if found is not None:
            return found
        try:
            root = inbox.Parent
            found = _find_subfolder(root, parts[0])
            if found is not None:
                return found
        except Exception:
            pass
        raise RuntimeError(
            f"Outlook folder '{folder_path}' not found. Tip: pass an explicit "
            f"path, e.g. 'Inbox/{parts[0]}'."
        )

    current = None
    for store in namespace.Folders:
        if store.Name.lower() == parts[0].lower():
            current = store
            parts = parts[1:]
            break
    if current is None:
        current = namespace.GetDefaultFolder(6).Parent

    for part in parts:
        nxt = _find_subfolder(current, part, recursive=False)
        if nxt is None:
            raise RuntimeError(
                f"Outlook folder '{folder_path}' not found (part '{part}' missing)."
            )
        current = nxt
    return current


def _find_subfolder(folder, name: str, recursive: bool = True):
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
    if only_unread:
        try:
            filtered = folder.Items.Restrict("[UnRead] = True")
            items = list(filtered)
            log.info("reading '%s' (%d unread of %d total)", folder.Name, len(items), total)
        except Exception as exc:
            log.warning("Restrict failed (%s) - reading all %d items.", exc, total)
            items = list(folder.Items)
    else:
        items = list(folder.Items)
        log.info("reading '%s' (%d items)", folder.Name, total)

    for item in items:
        try:
            if only_unread and not item.UnRead:
                continue
        except Exception:
            continue

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
                log.info("saved attachment: %s (mail: %s)", dest.name, subject[:50])
        except Exception as exc:
            log.warning("error reading a mail '%s': %s", subject[:50], exc)
            continue

        if saved_any and mark_read:
            try:
                item.UnRead = False
                item.Save()
            except Exception as exc:  # pragma: no cover
                log.warning("could not mark mail as read: %s", exc)

    log.info("fetch: %d PDF attachment(s) saved.", len(results))
    return results


def send_message(
    to: str,
    subject: str,
    body: str,
    attachments: Sequence[str | Path] | None = None,
    *,
    dry_run: bool = True,
) -> dict:
    attachments = attachments or []
    missing = [str(a) for a in attachments if not Path(a).is_file()]
    if missing:
        raise FileNotFoundError(f"attachment missing: {missing}")

    outlook = _get_outlook()
    mail = outlook.CreateItem(_OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    for a in attachments:
        mail.Attachments.Add(str(Path(a).resolve()))

    status = "sent" if not dry_run else "draft"
    if dry_run:
        mail.Save()
    else:
        mail.Send()

    log.info(
        "mail %s -> %s | subject: %s | attachments: %d",
        status,
        to,
        subject,
        len(attachments),
    )
    return {
        "to": to,
        "subject": subject,
        "attachments": [str(a) for a in attachments],
        "dry_run": dry_run,
        "status": status,
    }


def _main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Outlook helper (draft / fetch / list).")
    parser.add_argument(
        "--list-folders", action="store_true",
        help="List all Outlook folders (diagnostic) and exit.",
    )
    parser.add_argument(
        "--fetch", metavar="OUTDIR",
        help="Instead of a draft: fetch PDF attachments from the folder to OUTDIR.",
    )
    parser.add_argument("--folder", default="ScanToMe", help="Outlook folder (default ScanToMe)")
    parser.add_argument("--all", action="store_true", help="also fetch read mails")
    parser.add_argument("--to", help="recipient (only needed for draft/send)")
    parser.add_argument("--subject", default="PostPilot Test")
    parser.add_argument("--body", default="PostPilot test draft.")
    parser.add_argument("--attach", nargs="*", default=[])
    parser.add_argument("--send", action="store_true", help="actually send (else draft)")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
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
        print(f"\n{len(rows)} PDF attachment(s) saved.")
        return
    if not args.to:
        parser.error("--to is required (except with --fetch)")
    res = send_message(args.to, args.subject, args.body, args.attach, dry_run=not args.send)
    print(res)


if __name__ == "__main__":
    _main()
