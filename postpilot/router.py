"""Ties extraction, classification and sending together.

Per individual letter PDF:
  1. pdf_extract.extract()  -> fields (sender, date, recipient, full_text)
  2. rule_engine.classify() -> department + score + matches
  3. auto_send == False -> move PDF to review/ (no send)
     auto_send == True  -> render subject/body from the department template and
                           create the mail (draft or send).

Send safety: dry_run=True (default) creates Outlook drafts / .eml files only.
dry_run=False actually sends.

Template placeholders: {sender}, {date}, {recipient}, {department}, {score}.
A missing placeholder renders as "" rather than raising.
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

from postpilot import pdf_extract, rule_engine
from postpilot.rule_engine import Classification, RuleConfig

log = logging.getLogger("postpilot.router")


def default_config_path() -> str:
    local = Path("config/departments.local.yaml")
    if local.is_file():
        return str(local)
    return "config/departments.yaml"


def _select_backend(backend: str):
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
    raise ValueError(f"unknown backend: {backend!r} (allowed: auto|outlook|eml)")


class _SafeDict(dict):
    def __missing__(self, key):
        log.warning("template placeholder {%s} not found -> empty", key)
        return ""


def _render(template: str, k: Classification) -> str:
    values = _SafeDict(
        sender=k.sender,
        date=k.date,
        recipient=k.recipient,
        department=k.department,
        score=k.score,
    )
    return template.format_map(values)


def _move_to_review(pdf: Path, review_dir: Path, reason: str) -> str:
    review_dir.mkdir(parents=True, exist_ok=True)
    dest = review_dir / pdf.name
    counter = 1
    while dest.exists():
        dest = review_dir / f"{pdf.stem}__{counter}{pdf.suffix}"
        counter += 1
    shutil.copy2(pdf, dest)
    log.warning("-> review/: %s (%s)", pdf.name, reason)
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
    pdf = Path(pdf)
    ext = pdf_extract.extract(pdf, ocr_fallback=ocr_fallback)
    ext_d = {
        "sender": ext.sender,
        "full_text": ext.full_text,
        "recipient": ext.recipient,
        "date": ext.date,
        "file": ext.source,
    }
    k = rule_engine.classify(ext_d, cfg)

    entry: dict = {
        "file": pdf.name,
        "department": k.department,
        "score": k.score,
        "sender": k.sender,
        "date": k.date,
        "matches": [m.as_dict() for m in k.matches],
        "auto_send": k.auto_send,
    }

    if not k.auto_send:
        entry["action"] = "review"
        entry["reason"] = k.reason
        entry["review_path"] = _move_to_review(pdf, Path(review_dir), k.reason)
        return entry

    dept = rule_engine.department_by_name(cfg, k.department)
    if not dept or not dept.get("email"):
        entry["action"] = "review"
        entry["reason"] = f"no email address for department {k.department}"
        entry["review_path"] = _move_to_review(pdf, Path(review_dir), entry["reason"])
        return entry

    subject = _render(dept.get("subject_template", "Forwarding: {sender}"), k)
    body = _render(dept.get("body_template", ""), k)

    target = k.subroute_email or dept["email"]

    sender, backend_name = _select_backend(backend)
    res = sender(
        to=target,
        subject=subject,
        body=body,
        attachments=[pdf],
        dry_run=dry_run,
    )
    entry["backend"] = backend_name
    entry["action"] = res.get("status", "created")
    entry["to"] = target
    entry["subroute"] = k.subroute
    entry["subject"] = subject
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
    cfg = RuleConfig.load(config_path)
    folder = Path(folder)
    pdfs = sorted(folder.glob("*.pdf"))
    report: list[dict] = []

    for pdf in pdfs:
        if only_department:
            ext = pdf_extract.extract(pdf, ocr_fallback=ocr_fallback)
            ext_d = {
                "sender": ext.sender,
                "full_text": ext.full_text,
                "recipient": ext.recipient,
                "date": ext.date,
                "file": ext.source,
            }
            k = rule_engine.classify(ext_d, cfg)
            if k.department != only_department:
                log.info(
                    "skipped (not %s): %s [%s, auto=%s]",
                    only_department,
                    pdf.name,
                    k.department,
                    k.auto_send,
                )
                report.append(
                    {
                        "file": pdf.name,
                        "action": "skipped",
                        "department": k.department,
                        "auto_send": k.auto_send,
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


def _main() -> None:
    import argparse
    import json

    parser = argparse.ArgumentParser(description="Classify and forward letters.")
    parser.add_argument("folder", help="Folder with individual letter PDFs (e.g. work/split)")
    parser.add_argument(
        "--config",
        default=None,
        help="Config file. Default: auto (departments.local.yaml if present, "
        "else departments.yaml).",
    )
    parser.add_argument("--review", default="review")
    parser.add_argument("--only", metavar="DEPT", help="only this department (e.g. TAX)")
    parser.add_argument(
        "--send",
        action="store_true",
        help="Actually send. Without this flag: drafts only (dry run).",
    )
    parser.add_argument("--ocr", action="store_true", help="Enable OCR fallback")
    parser.add_argument(
        "--backend",
        choices=["auto", "outlook", "eml"],
        default="auto",
        help="Mail backend: auto (Windows->outlook, else->eml), "
        "outlook (Windows only), eml (.eml files, cross-platform). Default: auto.",
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

    mode = "SEND" if args.send else "DRY-RUN (drafts)"
    n_mail = sum(1 for r in report if r.get("action") in ("draft", "sent", "eml_written", "created"))
    n_review = sum(1 for r in report if r.get("action") == "review")
    n_skip = sum(1 for r in report if r.get("action") == "skipped")
    print(json.dumps(report, indent=2, ensure_ascii=False))
    print(f"\n[{mode}] {n_mail} mail(s), {n_review} -> review, {n_skip} skipped")


if __name__ == "__main__":
    _main()
