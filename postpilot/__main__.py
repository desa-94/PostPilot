from __future__ import annotations

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from postpilot import router
from postpilot.router import default_config_path
from postpilot.stamp_splitter import StampConfig, split_by_stamps

log = logging.getLogger("postpilot.main")


USAGE = """\
PostPilot - rule-based letter forwarding

Usage:
  py -m postpilot                 Start the interactive menu
  py -m postpilot [options]       Run non-interactively (see below)

Common commands:
  py -m postpilot --send          Split TAX.pdf and actually send
  py -m postpilot --fetch         Fetch PDFs from the Outlook ScanToMe folder, then process
  py -m postpilot --fetch --send  Fetch and send
  py -m postpilot --pdf FILE      Use a different source PDF
  py -m postpilot --date DD_MM_YYYY   Use a specific date folder

Options:
  --pdf FILE        Source batch PDF (default: TAX.pdf)
  --fetch           Pull PDFs from the Outlook folder instead of --pdf
  --folder NAME     Outlook folder for --fetch (default: config scantome.folder)
  --all             With --fetch, also include already-read mails
  --date DD_MM_YYYY Date subfolder under --out (default: today)
  --out DIR         Base output folder (default: work/split)
  --review DIR      Review folder for held-back letters (default: review)
  --only DEPT       Only forward this department (default: TAX)
  --send            Actually send (default: dry run, drafts only)
  --no-flag         Do not update the Outlook follow-up flag after processing
  --backend NAME    auto | outlook | eml (default: auto)
  --resolution DPI  Render DPI for splitting (default: 150)
  -v, --verbose     Debug logging

Related helpers:
  py -m postpilot.outlook_client --list-folders   Show all Outlook folders
  py -m postpilot.stamp_splitter FILE --debug-image out.png   Stamp detection preview
"""


def usage() -> None:
    print(USAGE)


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
    day = day or date.today().strftime("%d_%m_%Y")
    out_dir = Path(out_base) / day

    log.info("splitting %s -> %s", pdf, out_dir)
    split = split_by_stamps(
        pdf, out_dir, review_dir, StampConfig(resolution=resolution)
    )
    if not split.ok:
        log.warning(
            "split not performed (%s) - whole PDF moved to review/.", split.reason
        )
        return {
            "split_ok": False,
            "reason": split.reason,
            "review_path": split.review_path,
            "routed": [],
        }
    log.info("%d letters created.", len(split.letters))

    cfg_path = default_config_path()
    log.info(
        "routing %s (config: %s, backend: %s, %s)",
        out_dir,
        cfg_path,
        backend,
        "SEND" if send else "DRY-RUN",
    )
    report = router.route_folder(
        out_dir,
        cfg_path,
        review_dir=review_dir,
        dry_run=not send,
        only_department=only_department,
        backend=backend,
    )
    return {"split_ok": True, "letters": len(split.letters), "routed": report}


def run_is_clean(result: dict) -> bool:
    """Decide whether a source mail was processed cleanly.

    Clean means: the split succeeded, at least one letter was actually
    mailed (draft/sent/eml_written/created), and nothing was held back to
    review. Deliberately skipped letters (foreign department via --only)
    do not block the green check. Anything in review keeps the mail flagged
    red so it stays visible as "needs manual attention".
    """
    if not result.get("split_ok"):
        return False
    report = result.get("routed") or []
    if not report:
        return False
    mailed = {"draft", "sent", "eml_written", "created"}
    for r in report:
        action = r.get("action")
        if action in mailed or action == "skipped":
            continue
        return False  # review or anything unexpected -> keep red
    return any(r.get("action") in mailed for r in report)


def scantome_settings() -> dict:
    import yaml

    try:
        data = (
            yaml.safe_load(Path(default_config_path()).read_text(encoding="utf-8"))
            or {}
        )
    except Exception:
        return {"folder": None, "only_unread": True}
    sec = data.get("scantome") or {}
    return {
        "folder": sec.get("folder"),
        "only_unread": bool(sec.get("only_unread", True)),
    }


def fetch_from_scantome(
    *,
    folder_path: str = "ScanToMe",
    inbox_base: str | Path = "work/inbox",
    day: str | None = None,
    only_unread: bool = True,
) -> list[dict]:
    """Fetch PDF attachments and return full rows (incl. message_id)."""
    from postpilot import outlook_client

    day = day or date.today().strftime("%d_%m_%Y")
    inbox = Path(inbox_base) / day
    rows = outlook_client.fetch_pdf_attachments(
        inbox, folder_path=folder_path, only_unread=only_unread
    )
    log.info("fetch: %d PDF(s) -> %s", len(rows), inbox)
    return rows


def process_sources(
    sources: list[tuple[str, str | None]],
    *,
    out_base: str | Path = "work/split",
    review_dir: str | Path = "review",
    day: str | None = None,
    only_department: str | None = "TAX",
    send: bool = False,
    backend: str = "auto",
    resolution: int = 150,
    update_flag: bool = True,
) -> list[tuple[str, dict]]:
    """Process (src_path, entry_id) pairs and optionally set Outlook flags.

    After a source is processed, if it ran cleanly and carries an entry_id,
    the originating mail's follow-up flag is turned into a green check mark.
    Runs in both dry-run and send mode (a draft counts as "handled").
    """
    results: list[tuple[str, dict]] = []
    multi = len(sources) > 1
    base_day = day or date.today().strftime("%d_%m_%Y")

    for idx, (src, entry_id) in enumerate(sources, 1):
        sub_day = f"{base_day}/{idx:02d}" if multi else day
        if not Path(src).is_file():
            log.warning("source not found: %s", src)
            results.append(
                (src, {"split_ok": False, "reason": "file not found", "routed": []})
            )
            continue
        try:
            result = run(
                src,
                out_base=out_base,
                review_dir=review_dir,
                day=sub_day,
                only_department=only_department,
                send=send,
                backend=backend,
                resolution=resolution,
            )
        except Exception as exc:  # pragma: no cover - defensive
            log.warning("processing failed for %s: %s", src, exc)
            results.append(
                (src, {"split_ok": False, "reason": str(exc), "routed": []})
            )
            continue
        results.append((src, result))

        if update_flag and entry_id:
            if run_is_clean(result):
                try:
                    from postpilot import outlook_client

                    outlook_client.mark_message_complete(entry_id, complete=True)
                except Exception as exc:  # pragma: no cover
                    log.warning("flag update failed: %s", exc)
            else:
                log.info("mail kept flagged red (not clean): %s", entry_id[:12])

    return results


def print_summary(results: list[tuple[str, dict]], *, send: bool) -> None:
    mode = "SEND" if send else "DRY-RUN (drafts)"
    tot_letters = tot_mail = tot_review = tot_skip = 0
    for src, result in results:
        if not result["split_ok"]:
            print(
                f"[ABORT] {Path(src).name}: {result['reason']} "
                f"-> {result.get('review_path')}"
            )
            continue
        report = result["routed"]
        tot_letters += result["letters"]
        tot_mail += sum(
            1
            for r in report
            if r.get("action") in ("draft", "sent", "eml_written", "created")
        )
        tot_review += sum(1 for r in report if r.get("action") == "review")
        tot_skip += sum(1 for r in report if r.get("action") == "skipped")
    print(
        f"\n[{mode}] {len(results)} source(s) | {tot_letters} letters split | "
        f"{tot_mail} mail(s), {tot_review} -> review, {tot_skip} skipped"
    )
    if not send:
        print("Note: dry run. Drafts are saved; re-run with SEND to actually send.")


# --------------------------------------------------------------------------
# Interactive menu
# --------------------------------------------------------------------------


def _confirm_send() -> bool:
    ans = input("Really SEND for real? [y/N] > ").strip().lower()
    return ans in ("y", "yes", "j", "ja")


def _menu_fetch(state) -> list[dict]:
    folder = state.folder
    if not folder:
        print(
            "No ScanToMe folder configured. Set scantome.folder in the config "
            "(py -m postpilot.outlook_client --list-folders)."
        )
        return []
    try:
        rows = fetch_from_scantome(
            folder_path=folder, day=state.date, only_unread=state.only_unread
        )
    except Exception as exc:
        print(f"[fetch failed] {exc}")
        return []
    if not rows:
        print("[ScanToMe] No new PDF attachments found.")
    else:
        print(f"[ScanToMe] fetched {len(rows)} PDF(s).")
    return rows


def _menu_process(state, *, send: bool):
    if send and not _confirm_send():
        print("Cancelled - nothing sent.")
        return None
    if state.fetched:
        sources = [(r["pdf_path"], r.get("message_id")) for r in state.fetched]
    else:
        if not Path(state.pdf).is_file():
            print(
                f"No source to process: '{state.pdf}' not found and nothing "
                "fetched yet.\nUse option 1 (Fetch) first, or place the PDF in "
                "the project folder."
            )
            return None
        sources = [(state.pdf, None)]
    results = process_sources(
        sources,
        out_base=state.out,
        review_dir=state.review,
        day=state.date,
        only_department=state.only,
        send=send,
        backend=state.backend,
        resolution=state.resolution,
        update_flag=state.update_flag,
    )
    print_summary(results, send=send)
    return results


def interactive_menu() -> None:
    import types

    logging.basicConfig(
        level=logging.INFO, format="%(levelname)s %(name)s: %(message)s"
    )
    st = scantome_settings()
    state = types.SimpleNamespace(
        pdf="TAX.pdf",
        folder=st.get("folder"),
        only_unread=st.get("only_unread", True),
        date=None,
        out="work/split",
        review="review",
        only="TAX",
        send=False,
        backend="auto",
        resolution=150,
        update_flag=True,
        fetched=[],
        last_results=None,
    )

    def current_day():
        return state.date or date.today().strftime("%d_%m_%Y")

    while True:
        mode = "SEND (real)" if state.send else "DRY-RUN (drafts)"
        if state.fetched:
            n = len(state.fetched)
            pending = f"{n} fetched PDF(s) ready"
        elif Path(state.pdf).is_file():
            pending = f"source file: {state.pdf}"
        else:
            pending = "nothing loaded - use 1) Fetch"
        print()
        print("=== PostPilot ===")
        print(f" Date folder: {current_day()}   |   Mode: {mode}")
        print(f" Source: {pending}")
        print()
        print(" 1) Fetch PDFs from Outlook")
        print(" 2) Process + create drafts   (dry run)")
        print(" 3) Process + SEND            (real send)")
        print(" 4) Fetch + process in one go")
        print(" 5) Toggle mode (dry-run / send)")
        print(" 6) Change date folder")
        print(" 7) Show last run summary")
        print(" 0) Exit")
        try:
            choice = input("Choice > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nBye.")
            return

        if choice == "1":
            state.fetched = _menu_fetch(state)
        elif choice == "2":
            res = _menu_process(state, send=False)
            if res is not None:
                state.last_results = res
        elif choice == "3":
            res = _menu_process(state, send=True)
            if res is not None:
                state.last_results = res
        elif choice == "4":
            state.fetched = _menu_fetch(state)
            if state.fetched:
                res = _menu_process(state, send=state.send)
                if res is not None:
                    state.last_results = res
        elif choice == "5":
            state.send = not state.send
            print(f"Mode is now: {'SEND (real)' if state.send else 'DRY-RUN'}")
        elif choice == "6":
            val = input("Date folder (DD_MM_YYYY, empty = today) > ").strip()
            state.date = val or None
        elif choice == "7":
            if state.last_results is None:
                print("No run yet.")
            else:
                print_summary(state.last_results, send=state.send)
        elif choice == "0":
            print("Bye.")
            return
        else:
            print("Unknown choice.")


# --------------------------------------------------------------------------
# Non-interactive CLI
# --------------------------------------------------------------------------


def _main() -> None:
    if len(sys.argv) == 1:
        interactive_menu()
        return

    parser = argparse.ArgumentParser(
        prog="postpilot",
        add_help=True,
        description=USAGE,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--pdf", default="TAX.pdf")
    parser.add_argument("--fetch", action="store_true")
    parser.add_argument("--folder", default=None)
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--date", default=None)
    parser.add_argument("--out", default="work/split")
    parser.add_argument("--review", default="review")
    parser.add_argument("--only", default="TAX")
    parser.add_argument("--send", action="store_true")
    parser.add_argument("--no-flag", action="store_true")
    parser.add_argument("--backend", choices=["auto", "outlook", "eml"], default="auto")
    parser.add_argument("--resolution", type=int, default=150)
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )

    if args.fetch:
        st = scantome_settings()
        folder = args.folder or st.get("folder")
        if not folder:
            parser.error(
                "No ScanToMe folder configured. Pass --folder or set "
                "scantome.folder in the config "
                "(py -m postpilot.outlook_client --list-folders lists folders)."
            )
        only_unread = st.get("only_unread", True) and not args.all
        rows = fetch_from_scantome(
            folder_path=folder, day=args.date, only_unread=only_unread
        )
        if not rows:
            print("\n[ScanToMe] No new PDF attachments found.")
            return
        sources = [(r["pdf_path"], r.get("message_id")) for r in rows]
    else:
        sources = [(args.pdf, None)]

    results = process_sources(
        sources,
        out_base=args.out,
        review_dir=args.review,
        day=args.date,
        only_department=args.only,
        send=args.send,
        backend=args.backend,
        resolution=args.resolution,
        update_flag=not args.no_flag,
    )
    print_summary(results, send=args.send)


if __name__ == "__main__":
    _main()
