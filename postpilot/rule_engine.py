"""Rule-based classification (no LLM, no cloud).

Each letter is scored against the rules in config/departments.yaml. Sender
rules are a strong signal; keyword rules are weak and additive. The highest
scoring department wins. A letter is held back for review when the winning
score is below fallback.min_score or two departments tie.

Every rule hit is recorded in the matches list (type, pattern, score, matched
text) so a routing decision can always be explained.

Sender rules may use rapidfuzz to tolerate OCR errors ("Flnanzamt" ~
"Finanzamt"). Without rapidfuzz the engine falls back to substring matching.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

log = logging.getLogger("postpilot.rule_engine")

try:
    from rapidfuzz import fuzz

    _HAVE_RAPIDFUZZ = True
except ImportError:  # pragma: no cover
    _HAVE_RAPIDFUZZ = False
    log.info("rapidfuzz not available - fuzzy matching falls back to substring.")


@dataclass
class Match:
    type: str
    pattern: str
    score: int
    matched: str

    def as_dict(self) -> dict:
        return {
            "type": self.type,
            "pattern": self.pattern,
            "score": self.score,
            "matched": self.matched,
        }


@dataclass
class Classification:
    file: str
    sender: str
    recipient: str
    date: str
    department: str
    score: int
    matches: list[Match] = field(default_factory=list)
    all_scores: dict = field(default_factory=dict)
    auto_send: bool = True
    reason: str = ""
    subroute: str = ""
    subroute_email: str = ""
    subroute_matches: list[Match] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "file": self.file,
            "sender": self.sender,
            "recipient": self.recipient,
            "date": self.date,
            "department": self.department,
            "score": self.score,
            "matches": [m.as_dict() for m in self.matches],
            "all_scores": self.all_scores,
            "auto_send": self.auto_send,
            "reason": self.reason,
            "subroute": self.subroute,
            "subroute_email": self.subroute_email,
            "subroute_matches": [m.as_dict() for m in self.subroute_matches],
        }


@dataclass
class RuleConfig:
    departments: list[dict]
    fallback: dict
    fuzzy_threshold: int = 85

    @classmethod
    def load(cls, path: str | Path) -> RuleConfig:
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
        depts = data.get("departments", []) or []
        fb = data.get("fallback", {}) or {}
        fuzzy = int(data.get("fuzzy_threshold", 85))
        return cls(departments=depts, fallback=fb, fuzzy_threshold=fuzzy)


_REGEX_META = set(r"[]()\{}.*+?^$|")


def _has_regex_meta(pattern: str) -> bool:
    return any(c in _REGEX_META for c in pattern)


def _sender_match(
    pattern: str, text: str, threshold: int, fuzzy_term: str | None = None
) -> str | None:
    try:
        m = re.search(pattern, text, re.IGNORECASE)
    except re.error:
        m = None
        if pattern.lower() in text.lower():
            idx = text.lower().find(pattern.lower())
            return text[idx : idx + len(pattern)]
    if m:
        return m.group(0)

    if _HAVE_RAPIDFUZZ:
        if fuzzy_term:
            needle = fuzzy_term.strip()
        elif not _has_regex_meta(pattern):
            needle = pattern.strip()
        else:
            needle = ""
        if needle:
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                if len(line) < len(needle) * 0.6:
                    continue
                if fuzz.partial_ratio(needle.lower(), line.lower()) >= threshold:
                    return line[:60]
    return None


def _keyword_match(pattern: str, text: str) -> str | None:
    try:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return m.group(0)
    except re.error:
        if pattern.lower() in text.lower():
            idx = text.lower().find(pattern.lower())
            return text[idx : idx + len(pattern)]
    return None


def classify(extract: dict, cfg: RuleConfig) -> Classification:
    sender = extract.get("sender", "") or ""
    full_text = extract.get("full_text", "") or ""
    recipient = extract.get("recipient", "") or ""
    date = extract.get("date", "") or ""
    file = extract.get("file") or extract.get("source", "") or ""

    sender_haystack = f"{sender}\n{full_text}"

    scores: dict[str, int] = {}
    matches_by_dept: dict[str, list[Match]] = {}

    for dept in cfg.departments:
        name = dept.get("name", "?")
        rules = dept.get("rules", {}) or {}
        total = 0
        hits: list[Match] = []

        for rule in rules.get("senders", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            fuzzy_term = rule.get("fuzzy")
            matched = _sender_match(
                pat, sender_haystack, cfg.fuzzy_threshold, fuzzy_term
            )
            if matched:
                total += sc
                hits.append(Match("sender", pat, sc, matched.strip()))

        for rule in rules.get("keywords", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            matched = _keyword_match(pat, full_text)
            if matched:
                total += sc
                hits.append(Match("keyword", pat, sc, matched.strip()))

        scores[name] = total
        matches_by_dept[name] = hits

    fb = cfg.fallback
    fb_name = fb.get("department", "UNGEKLAERT")
    min_score = int(fb.get("min_score", 0))

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    base = Classification(
        file=file,
        sender=sender,
        recipient=recipient,
        date=date,
        department=fb_name,
        score=0,
        all_scores=scores,
    )

    if not ranked or ranked[0][1] == 0:
        base.auto_send = False
        base.reason = "no rule matched"
        return base

    top_name, top_score = ranked[0]
    base.score = top_score
    base.matches = matches_by_dept.get(top_name, [])

    if len(ranked) > 1 and ranked[1][1] == top_score:
        base.auto_send = False
        base.reason = f"tie {top_name} vs. {ranked[1][0]} (both {top_score})"
        return base

    if top_score < min_score:
        base.auto_send = False
        base.reason = f"score {top_score} < min_score {min_score}"
        return base

    base.department = top_name
    base.auto_send = True

    winner_dept = None
    for d in cfg.departments:
        if d.get("name") == top_name:
            winner_dept = d
            break
    if winner_dept and winner_dept.get("subroutes"):
        sr_name, sr_email, sr_hits, sr_reason = classify_subroute(
            winner_dept, full_text, sender, cfg.fuzzy_threshold
        )
        base.subroute_matches = sr_hits
        if not sr_name:
            base.auto_send = False
            base.reason = f"subroute unclear ({sr_reason})"
            return base
        base.subroute = sr_name
        base.subroute_email = sr_email
    return base


def classify_subroute(
    dept: dict, full_text: str, sender: str, fuzzy_threshold: int = 85
) -> tuple[str, str, list[Match], str]:
    subroutes = dept.get("subroutes") or []
    if not subroutes:
        return "", "", [], "no subroutes defined"

    min_score = int(dept.get("subroute_min_score", 0))
    haystack = f"{sender}\n{full_text}"

    scores: dict[str, int] = {}
    matches_by_sr: dict[str, list[Match]] = {}
    email_by_sr: dict[str, str] = {}

    for sr in subroutes:
        name = sr.get("name", "?")
        email_by_sr[name] = sr.get("email", "")
        total = 0
        hits: list[Match] = []
        for rule in sr.get("patterns", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            fuzzy_term = rule.get("fuzzy")
            matched = _sender_match(pat, haystack, fuzzy_threshold, fuzzy_term)
            if matched:
                total += sc
                hits.append(Match("subroute", pat, sc, matched.strip()))
        scores[name] = total
        matches_by_sr[name] = hits

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    if not ranked or ranked[0][1] == 0:
        return "", "", [], "no subroute signal found"

    top_name, top_score = ranked[0]
    if len(ranked) > 1 and ranked[1][1] == top_score:
        return (
            "",
            "",
            matches_by_sr.get(top_name, []),
            f"tie subroute {top_name} vs. {ranked[1][0]} (both {top_score})",
        )
    if top_score < min_score:
        return (
            "",
            "",
            matches_by_sr.get(top_name, []),
            f"subroute score {top_score} < min {min_score}",
        )

    return top_name, email_by_sr.get(top_name, ""), matches_by_sr.get(top_name, []), ""


def department_by_name(cfg: RuleConfig, name: str) -> dict | None:
    for d in cfg.departments:
        if d.get("name") == name:
            return d
    if name == cfg.fallback.get("department"):
        return {"name": name, "email": cfg.fallback.get("email", "")}
    return None


def _main() -> None:
    import argparse
    import json

    from postpilot import pdf_extract

    parser = argparse.ArgumentParser(description="Classify letter PDFs by rules.")
    parser.add_argument("pdfs", nargs="+", help="PDF files or folders")
    parser.add_argument("--config", default="config/departments.yaml")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s %(name)s: %(message)s",
    )
    cfg = RuleConfig.load(args.config)

    paths: list[Path] = []
    for p in args.pdfs:
        pp = Path(p)
        paths.extend(sorted(pp.glob("*.pdf")) if pp.is_dir() else [pp])

    results = []
    for pdf in paths:
        ext = pdf_extract.extract(pdf)
        ext_d = {
            "sender": ext.sender,
            "full_text": ext.full_text,
            "recipient": ext.recipient,
            "date": ext.date,
            "file": ext.source,
        }
        k = classify(ext_d, cfg)
        results.append(k.as_dict())

    print(json.dumps(results, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    _main()
