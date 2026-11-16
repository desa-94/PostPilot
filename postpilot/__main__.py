"""Orchestration: one command runs the whole pipeline.

  1. Split the source PDF into individual letters by their incoming stamp
     -> work/split/<DD_MM_YYYY>/
  2. Classify and forward those letters (TAX only by default)
     -> dry run (Outlook drafts / .eml) unless --send is given.

Run `py -m postpilot` with no arguments to print the usage overview.
"""

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
  py -m postpilot [options]

Common commands:
  py -m postpilot                 Split TAX.pdf (today's folder) and create drafts (dry run)
  py -m postpilot --send          Same, but actually send
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
    split = split_by_stamps(pdf, out_dir, review_dir, StampConfig(resolution=resolution))
    if not split.ok:
        log.warning("split not performed (%s) - whole PDF moved to review/.", split.reason)
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


def scantome_settings() -> dict:
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
    from postpilot import outlook_client

    day = day or date.today().strftime("%d_%m_%Y")
    inbox = Path(inbox_base) / day
    rows = outlook_client.fetch_pdf_attachments(
        inbox, folder_path=folder_path, only_unread=only_unread
    )
    paths = [r["pdf_path"] for r in rows]
    log.info("fetch: %d PDF(s) -> %s", len(paths), inbox)
    return paths


def _main() -> None:
    if len(sys.argv) == 1:
        usage()
        return

    parser = argparse.ArgumentParser(prog="postpilot", add_help=True, description=USAGE,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pdf", default="TAX.pdf")
    parser.add_argument("--fetch", action="store_true")
    parser.add_argument("--folder", default=None)
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--date", default=None)
    parser.add_argument("--out", default="work/split")
    parser.add_argument("--review", default="review")
    parser.add_argument("--only", default="TAX")
    parser.add_argument("--send", action="store_true")
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
        sources = fetch_from_scantome(
            folder_path=folder, day=args.date, only_unread=only_unread
        )
        if not sources:
            print("\n[ScanToMe] No new PDF attachments found.")
            return
    else:
        sources = [args.pdf]

    results = []
    for idx, src in enumerate(sources, 1):
        day = args.date
        if args.fetch and len(sources) > 1:
            base_day = args.date or date.today().strftime("%d_%m_%Y")
            day = f"{base_day}/{idx:02d}"
        results.append(
            (
                src,
                run(
                    src,
                    out_base=args.out,
                    review_dir=args.review,
                    day=day,
                    only_department=args.only,
                    send=args.send,
                    backend=args.backend,
                    resolution=args.resolution,
                ),
            )
        )

    mode = "SEND" if args.send else "DRY-RUN (drafts)"
    tot_letters = tot_mail = tot_review = tot_skip = 0
    for src, result in results:
        if not result["split_ok"]:
            print(f"[ABORT] {Path(src).name}: {result['reason']} -> {result.get('review_path')}")
            continue
        report = result["routed"]
        tot_letters += result["letters"]
        tot_mail += sum(1 for r in report if r.get("action") in ("draft", "sent", "eml_written", "created"))
        tot_review += sum(1 for r in report if r.get("action") == "review")
        tot_skip += sum(1 for r in report if r.get("action") == "skipped")
    print(
        f"\n[{mode}] {len(results)} source(s) | {tot_letters} letters split | "
        f"{tot_mail} mail(s), {tot_review} -> review, {tot_skip} skipped"
    )
    if not args.send:
        print("Note: dry run. Re-run with --send to actually send.")


if __name__ == "__main__":
    _main()
