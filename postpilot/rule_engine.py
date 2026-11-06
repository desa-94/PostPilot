"""
rule_engine.py — Regelbasierte Klassifizierung (KEIN LLM, KEINE Cloud).

Prinzip
-------
Jeder Brief wird gegen die Regeln aus config/departments.yaml gescort:

  senders  : starkes Signal (Absender passt)  -> hohe Scores
  keywords : schwaches Signal (Begriff im Text) -> summiert sich

Pro Abteilung werden alle Treffer-Scores summiert. Hoechster Score gewinnt.
Fail-safe (-> Fallback-Abteilung UNGEKLAERT, Pruefungsordner), wenn:
  - der Gewinner-Score < fallback.min_score liegt, ODER
  - ein Gleichstand zwischen zwei Abteilungen besteht.

Nachvollziehbarkeit
-------------------
Jeder Treffer landet in der 'treffer'-Liste mit Regeltyp, Muster, Score und
dem konkret gematchten Textausschnitt -> "Brief ging nach TAX, weil Regel X griff".

Fuzzy-Matching (OCR-tolerant)
-----------------------------
Absender-Regeln (senders) nutzen rapidfuzz, um OCR-Fehler zu tolerieren
("Flnanzamt" ~ "Finanzamt"). Fehlt rapidfuzz, faellt die Engine automatisch
auf simples Substring-Matching zurueck (Tool bleibt lauffaehig).

Abhaengigkeiten: PyYAML. Optional: rapidfuzz.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

log = logging.getLogger("postpilot.rule_engine")

# rapidfuzz ist optional. Fehlt es, nutzen wir Substring-Matching.
try:
    from rapidfuzz import fuzz

    _HAVE_RAPIDFUZZ = True
except ImportError:  # pragma: no cover
    _HAVE_RAPIDFUZZ = False
    log.info(
        "rapidfuzz nicht verfuegbar — Fuzzy-Matching faellt auf Substring zurueck."
    )


# --------------------------------------------------------------------------- #
# Ergebnis-Objekte
# --------------------------------------------------------------------------- #
@dataclass
class Treffer:
    """Ein einzelner Regeltreffer — die Einheit der Nachvollziehbarkeit."""

    typ: str  # "sender" | "keyword"
    muster: str  # das Regex/Muster aus der Config
    score: int  # beigetragener Score
    gematcht: str  # konkret gematchter Textausschnitt (fuzzy: Fundstelle)

    def as_dict(self) -> dict:
        return {
            "typ": self.typ,
            "muster": self.muster,
            "score": self.score,
            "gematcht": self.gematcht,
        }


@dataclass
class Klassifikation:
    datei: str
    absender: str
    empfaenger: str
    datum: str
    abteilung: str
    score: int
    treffer: list[Treffer] = field(default_factory=list)
    alle_scores: dict = field(
        default_factory=dict
    )  # {Abteilung: Score} fuer Transparenz
    auto_versand: bool = True  # False -> Pruefungsordner
    grund: str = ""  # Begruendung bei kein Auto-Versand

    # Sub-Routing innerhalb der Abteilung (z.B. TAX -> GARNISHMENT/DETAINTAKE).
    subroute: str = ""  # Name der Subroute (leer = keine)
    subroute_email: str = ""  # Ziel-Adresse der Subroute (ueberschreibt dept-email)
    subroute_treffer: list[Treffer] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "datei": self.datei,
            "absender": self.absender,
            "empfaenger": self.empfaenger,
            "datum": self.datum,
            "abteilung": self.abteilung,
            "score": self.score,
            "treffer": [t.as_dict() for t in self.treffer],
            "alle_scores": self.alle_scores,
            "auto_versand": self.auto_versand,
            "grund": self.grund,
            "subroute": self.subroute,
            "subroute_email": self.subroute_email,
            "subroute_treffer": [t.as_dict() for t in self.subroute_treffer],
        }


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
@dataclass
class RuleConfig:
    departments: list[dict]
    fallback: dict
    fuzzy_threshold: int = 85  # rapidfuzz partial_ratio-Schwelle fuer senders

    @classmethod
    def load(cls, path: str | Path) -> RuleConfig:
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
        depts = data.get("departments", []) or []
        fb = data.get("fallback", {}) or {}
        fuzzy = int(data.get("fuzzy_threshold", 85))
        return cls(departments=depts, fallback=fb, fuzzy_threshold=fuzzy)


# --------------------------------------------------------------------------- #
# Matching-Helfer
# --------------------------------------------------------------------------- #
# Zeichen, die ein Muster als Regex (statt Literal) kennzeichnen.
_REGEX_META = set(r"[]()\{}.*+?^$|")


def _has_regex_meta(pattern: str) -> bool:
    return any(c in _REGEX_META for c in pattern)


def _sender_match(
    pattern: str, text: str, threshold: int, fuzzy_term: str | None = None
) -> str | None:
    """
    Prueft einen sender-Pattern gegen den Text.
    1. Erst exaktes Regex/Substring (case-insensitive) — schnell & praezise.
    2. Dann fuzzy (rapidfuzz partial_ratio ueber Zeilen), um OCR-Fehler zu fangen.

    Fuzzy-Begriff (needle):
      - explizites fuzzy_term aus der Config hat Vorrang;
      - sonst nur, wenn 'pattern' ein reines Literal ist (keine Regex-Zeichen),
        wird das Muster selbst als needle genutzt;
      - bei Regex-Mustern OHNE fuzzy_term wird Fuzzy UEBERSPRUNGEN, damit kein
        zerhacktes needle (z.B. "Garnishm[a-z]{0,3}" -> "Garnishmaz03") entsteht
        und Zufallstreffer erzeugt.
    Returns: gematchter Ausschnitt (str) oder None.
    """
    # 1) exaktes (Regex-)Match
    try:
        m = re.search(pattern, text, re.IGNORECASE)
    except re.error:
        m = None
        if pattern.lower() in text.lower():
            idx = text.lower().find(pattern.lower())
            return text[idx : idx + len(pattern)]
    if m:
        return m.group(0)

    # 2) Fuzzy-Fallback (nur wenn rapidfuzz da ist)
    if _HAVE_RAPIDFUZZ:
        # needle sauber bestimmen — NIE aus einem Regex ableiten.
        if fuzzy_term:
            needle = fuzzy_term.strip()
        elif not _has_regex_meta(pattern):
            needle = pattern.strip()
        else:
            needle = ""  # Regex ohne expliziten fuzzy_term -> kein Fuzzy
        if needle:
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                # Zeilen, die deutlich kuerzer als das needle sind, koennen
                # keine OCR-Variante davon sein (z.B. eine "g"-Zeile ist kein
                # "Garnishment"). Das verhindert Zufallstreffer von
                # partial_ratio gegen sehr kurze Fragmente.
                if len(line) < len(needle) * 0.6:
                    continue
                if fuzz.partial_ratio(needle.lower(), line.lower()) >= threshold:
                    return line[:60]
    return None


def _keyword_match(pattern: str, text: str) -> str | None:
    """keywords: exaktes Regex/Substring-Match (case-insensitive). Kein Fuzzy."""
    try:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return m.group(0)
    except re.error:
        if pattern.lower() in text.lower():
            idx = text.lower().find(pattern.lower())
            return text[idx : idx + len(pattern)]
    return None


# --------------------------------------------------------------------------- #
# Kernfunktion
# --------------------------------------------------------------------------- #
def classify(extract: dict, cfg: RuleConfig) -> Klassifikation:
    """
    Klassifiziert EINEN Brief.

    Args:
        extract: dict aus pdf_extract (braucht 'absender', 'volltext';
                 optional 'empfaenger', 'datum', 'datei'/'source').
        cfg:     geladene RuleConfig.

    Returns: Klassifikation (mit treffer-Liste und alle_scores fuer Transparenz).
    """
    absender = extract.get("absender", "") or ""
    volltext = extract.get("volltext", "") or ""
    empfaenger = extract.get("empfaenger", "") or ""
    datum = extract.get("datum", "") or ""
    datei = extract.get("datei") or extract.get("source", "") or ""

    # Absender-Regeln pruefen primaer gegen das Absenderfeld, aber auch gegen
    # den Volltext (falls die Kopfzeile schlecht extrahiert wurde, steht der
    # Name oft trotzdem im Text).
    sender_haystack = f"{absender}\n{volltext}"

    scores: dict[str, int] = {}
    treffer_je_abt: dict[str, list[Treffer]] = {}

    for dept in cfg.departments:
        name = dept.get("name", "?")
        rules = dept.get("rules", {}) or {}
        total = 0
        hits: list[Treffer] = []

        for rule in rules.get("senders", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            fuzzy_term = rule.get("fuzzy")
            matched = _sender_match(
                pat, sender_haystack, cfg.fuzzy_threshold, fuzzy_term
            )
            if matched:
                total += sc
                hits.append(Treffer("sender", pat, sc, matched.strip()))

        for rule in rules.get("keywords", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            matched = _keyword_match(pat, volltext)
            if matched:
                total += sc
                hits.append(Treffer("keyword", pat, sc, matched.strip()))

        scores[name] = total
        treffer_je_abt[name] = hits

    fb = cfg.fallback
    fb_name = fb.get("department", "UNGEKLAERT")
    fb_email = fb.get("email", "")
    min_score = int(fb.get("min_score", 0))

    # Gewinner bestimmen
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    base = Klassifikation(
        datei=datei,
        absender=absender,
        empfaenger=empfaenger,
        datum=datum,
        abteilung=fb_name,
        score=0,
        alle_scores=scores,
    )

    if not ranked or ranked[0][1] == 0:
        base.auto_versand = False
        base.grund = "kein Regeltreffer"
        return base

    top_name, top_score = ranked[0]
    base.score = top_score
    base.treffer = treffer_je_abt.get(top_name, [])

    # Fail-safe 1: Gleichstand
    if len(ranked) > 1 and ranked[1][1] == top_score:
        base.auto_versand = False
        base.grund = f"Gleichstand {top_name} vs. {ranked[1][0]} (je {top_score})"
        return base

    # Fail-safe 2: unter Mindest-Score
    if top_score < min_score:
        base.auto_versand = False
        base.grund = f"Score {top_score} < min_score {min_score}"
        return base

    # Auto-Versand an Gewinner-Abteilung
    base.abteilung = top_name
    base.auto_versand = True

    # Sub-Routing: wenn die Gewinner-Abteilung Subrouten definiert, muss eine
    # sicher greifen — sonst -> review/ (kein blinder Versand an eine Adresse).
    winner_dept = None
    for d in cfg.departments:
        if d.get("name") == top_name:
            winner_dept = d
            break
    if winner_dept and (winner_dept.get("subroutes")):
        sr_name, sr_email, sr_hits, sr_grund = classify_subroute(
            winner_dept, volltext, absender, cfg.fuzzy_threshold
        )
        base.subroute_treffer = sr_hits
        if not sr_name:
            base.auto_versand = False
            base.grund = f"Subroute unklar ({sr_grund})"
            return base
        base.subroute = sr_name
        base.subroute_email = sr_email
    return base


def classify_subroute(
    dept: dict, volltext: str, absender: str, fuzzy_threshold: int = 85
) -> tuple[str, str, list[Treffer], str]:
    """
    Bestimmt die Subroute INNERHALB einer Abteilung anhand der Adresszeile.

    Gleiche Scoring-Mechanik wie die Abteilungen: pro Subroute alle
    patterns scoren, hoechster Score gewinnt. Fail-safe (-> ""), wenn:
      - keine Subroute ueber subroute_min_score liegt, ODER
      - Gleichstand zwischen zwei Subrouten besteht.

    Returns: (subroute_name, subroute_email, treffer, grund)
      subroute_name == "" bedeutet: keine sichere Subroute -> review/.
    """
    subroutes = dept.get("subroutes") or []
    if not subroutes:
        return "", "", [], "keine Subrouten definiert"

    min_score = int(dept.get("subroute_min_score", 0))
    haystack = f"{absender}\n{volltext}"

    scores: dict[str, int] = {}
    treffer_je_sr: dict[str, list[Treffer]] = {}
    email_je_sr: dict[str, str] = {}

    for sr in subroutes:
        name = sr.get("name", "?")
        email_je_sr[name] = sr.get("email", "")
        total = 0
        hits: list[Treffer] = []
        for rule in sr.get("patterns", []) or []:
            pat, sc = rule.get("pattern", ""), int(rule.get("score", 0))
            fuzzy_term = rule.get("fuzzy")
            matched = _sender_match(pat, haystack, fuzzy_threshold, fuzzy_term)
            if matched:
                total += sc
                hits.append(Treffer("subroute", pat, sc, matched.strip()))
        scores[name] = total
        treffer_je_sr[name] = hits

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    if not ranked or ranked[0][1] == 0:
        return "", "", [], "keine Subroute-Signale gefunden"

    top_name, top_score = ranked[0]
    # Gleichstand -> unsicher
    if len(ranked) > 1 and ranked[1][1] == top_score:
        return (
            "",
            "",
            treffer_je_sr.get(top_name, []),
            f"Gleichstand Subroute {top_name} vs. {ranked[1][0]} (je {top_score})",
        )
    # unter Mindest-Score -> unsicher
    if top_score < min_score:
        return (
            "",
            "",
            treffer_je_sr.get(top_name, []),
            f"Subroute-Score {top_score} < min {min_score}",
        )

    return top_name, email_je_sr.get(top_name, ""), treffer_je_sr.get(top_name, []), ""


def department_by_name(cfg: RuleConfig, name: str) -> dict | None:
    """Liefert den Abteilungs-Dict (fuer Router: email, templates) oder None."""
    for d in cfg.departments:
        if d.get("name") == name:
            return d
    if name == cfg.fallback.get("department"):
        return {"name": name, "email": cfg.fallback.get("email", "")}
    return None


# --------------------------------------------------------------------------- #
# CLI — klassifiziert alle PDFs in einem Ordner (nutzt pdf_extract)
# --------------------------------------------------------------------------- #
def _main() -> None:
    import argparse
    import json

    from postpilot import pdf_extract

    parser = argparse.ArgumentParser(
        description="Brief-PDFs regelbasiert klassifizieren."
    )
    parser.add_argument("pdfs", nargs="+", help="PDF-Dateien oder Ordner")
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
            "absender": ext.absender,
            "volltext": ext.volltext,
            "empfaenger": ext.empfaenger,
            "datum": ext.datum,
            "datei": ext.source,
        }
        k = classify(ext_d, cfg)
        results.append(k.as_dict())

    print(json.dumps(results, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    _main()
