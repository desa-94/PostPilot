# PostPilot

Regelbasierte Briefweiterleitung für die Poststelle. Überwacht einen
Outlook-Ordner, analysiert gescannte PDF-Briefe und leitet sie an die
zuständige Abteilung weiter – deterministisch, lokal, nachvollziehbar.

**Kein LLM, keine externe API, keine Cloud.** Die Klassifizierung läuft
ausschließlich über Regeln (Regex, Stichwörter, Absender-Matching). Briefinhalte
verlassen nie den Rechner. Jede Zuordnung ist erklärbar: "Brief ging nach TAX,
weil Regel X griff."

## Funktionen

- **Stempel-Splitting**: Ein Sammel-PDF (mehrere Briefe hintereinander gescannt)
  wird anhand der visuellen Eingangsstempel wieder in Einzelbriefe zerlegt –
  per Farberkennung, ohne OCR.
- **Textextraktion**: Liest den vorhandenen Textlayer (HP LaserJet, 300 DPI);
  optionaler OCR-Fallback für Seiten ohne Textlayer.
- **Regelbasierte Klassifizierung**: Scoring pro Abteilung aus
  `config/departments.yaml`. Höchster Score gewinnt; bei Unsicherheit
  (zu niedriger Score oder Gleichstand) landet der Brief im Prüfungsordner.

## Setup

```bash
pip install -r requirements.txt
```

Windows vorausgesetzt (Outlook-Anbindung via `pywin32`). Tesseract-OCR ist
optional und standardmäßig deaktiviert.

## Nutzung

Sammel-PDF in Einzelbriefe splitten:

```bash
python -m postpilot.stamp_splitter sammel.pdf --out work/split --review review
```

Einzelbriefe klassifizieren:

```bash
python -m postpilot.rule_engine work/split --config config/departments.yaml
```

## Konfiguration

Abteilungen und Regeln werden in `config/departments.yaml` gepflegt – pro
Abteilung Empfänger-Adresse, Betreff-/Body-Vorlage und Scoring-Regeln
(`senders`, `keywords`). Der Prüfungsordner greift, wenn kein eindeutiger
Gewinner über dem Mindest-Score steht.

## Projektstruktur

```
postpilot/
├── stamp_splitter.py   # Sammel-PDF anhand Eingangsstempel splitten
├── pdf_extract.py      # Textlayer lesen (+ optionaler OCR-Fallback)
└── rule_engine.py      # regelbasiertes Scoring & Klassifizierung
config/
└── departments.yaml    # Abteilungen, Regeln, Fallback
```
