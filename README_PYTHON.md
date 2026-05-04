# Rechnungsbearbeitung - Python Version

PDF-Rechnungen automatisch umbenennen mit Daten aus dem PDF.

## Was macht das Tool?

- Liest alle PDFs im Ordner (wo das Script liegt)
- Extrahiert: **Datum**, **Lieferant**, **Gesamtpreis (inkl. MwSt)**
- Benennt um direkt im Ordner: `26-001_2026-04-21_Böttcher_202,31EUR.pdf`
- Erstellt: `Rechnungsuebersicht.xlsx` mit Monats- und Jahresauswertung

## Installation

### Schritt 1: Python installieren

Lade Python 3.8+ herunter: https://python.org/
Wichtig: **"Add Python to PATH"** während Installation ankreuzen!

### Schritt 2: Abhängigkeiten installieren

```bash
# In den Projektordner wechseln
cd rechnungsbearbeitung

# Abhängigkeiten installieren
pip install -r requirements.txt
```

### Schritt 3: Als .exe bauen (optional)

```bash
# PyInstaller installieren
pip install pyinstaller

# .exe erstellen (mit Konsole, damit man Output sieht)
pyinstaller --onefile --console rechnung_processor.py

# Oder für Windows mit Icon (optional):
pyinstaller --onefile --console --icon=rechnung.ico rechnung_processor.py
```

Die .exe findest du unter `dist/RechnungProcessor.exe`

## Nutzung

### Als Python-Script:
```bash
python rechnung_processor.py
```

### Als .exe:
1. `RechnungProcessor.exe` in Ordner mit PDFs kopieren
2. Doppelklick
3. Fertig!

**Wichtig:**
- Bereits verarbeitete PDFs (die mit "26-" beginnen) werden übersprungen
- Gescannte PDFs können nicht gelesen werden (nur digitale PDFs)

## Beispiel

**Vorher:**
```
Rechnung_2026405926.pdf
320262289791.pdf
RechnungProcessor.exe
```

**Nachher:**
```
26-001_2026-04-21_Böttcher_202,31EUR.pdf
26-002_2026-04-29_Steinke_22,32EUR.pdf
Rechnungsuebersicht.xlsx
RechnungProcessor.exe
```

## Fehlerbehebung

**"ModuleNotFoundError: No module named 'pdfplumber'"**
→ `pip install pdfplumber openpyxl` ausführen

**"Keine neuen PDF-Dateien gefunden"**
→ Alle PDFs haben bereits "26-" am Anfang oder es sind keine PDFs im Ordner.

**"Gescanntes PDF"**
→ OCR nicht verfügbar. Nur digitale PDFs möglich.

## Dateien

```
rechnungsbearbeitung/
├── rechnung_processor.py      # Haupt-Code (Python)
├── requirements.txt            # Abhängigkeiten
├── rechnung_processor.spec     # PyInstaller Config (wird automatisch erstellt)
├── dist/                       # Hier landet die fertige .exe
│   └── RechnungProcessor.exe
├── build/                      # Temporäre Build-Dateien
└── README_PYTHON.md            # Diese Datei
```

## Lizenz

MIT - Frei verwendbar.
