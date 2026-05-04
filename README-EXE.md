# Rechnungsbearbeitung - Standalone .exe

PDF-Rechnungen automatisch umbenennen mit Daten aus dem PDF.

## Was macht das Tool?

- Liest alle PDFs im Ordner (wo die .exe liegt)
- Extrahiert: **Datum**, **Lieferant**, **Gesamtpreis (inkl. MwSt)**
- Benennt um: `26-001_2026-04-21_Böttcher_202,31EUR.pdf`
- Erstellt: `Rechnungsuebersicht.xlsx` mit Monats- und Jahresauswertung

## Installation als .exe

### Schritt 1: Node.js installieren

Lade Node.js LTS herunter: https://nodejs.org/
Installiere mit Standard-Einstellungen.

### Schritt 2: Projekt vorbereiten

```bash
# In den Projektordner wechseln
cd rechnungsbearbeitung

# Abhängigkeiten installieren
npm install

# pkg global installieren (für .exe-Erstellung)
npm install -g pkg
```

### Schritt 3: .exe erstellen

```bash
# Für Windows 64-bit
pkg rechnung-processor.js --targets node18-win-x64 --output RechnungProcessor.exe

# Für Windows 32-bit
pkg rechnung-processor.js --targets node18-win-x86 --output RechnungProcessor.exe
```

Die .exe entsteht im gleichen Ordner.

## Nutzung

1. **RechnungProcessor.exe** in einen Ordner mit PDF-Rechnungen kopieren
2. Doppelklick auf die .exe
3. Fertig! PDFs werden umbenannt, Excel wird erstellt

**Wichtig:**
- Bereits verarbeitete PDFs (die mit "26-" beginnen) werden übersprungen
- Gescannte PDFs können nicht gelesen werden (nur digitale PDFs)

## Beispiel

**Vorher:**
```
Rechnung_2026405926.pdf
320262289791.pdf
```

**Nachher:**
```
26-001_2026-04-21_Böttcher_202,31EUR.pdf
26-002_2026-04-29_Steinke_22,32EUR.pdf
Rechnungsuebersicht.xlsx
```

## Fehlerbehebung

**"Keine neuen PDF-Dateien gefunden"**
→ Alle PDFs haben bereits "26-" am Anfang oder es sind keine PDFs im Ordner.

**"Gescanntes PDF"**
→ OCR nicht verfügbar. Nur digitale PDFs möglich.

**Excel lässt sich nicht öffnen**
→ Schließen Sie Excel vor dem nächsten Durchlauf.

## Ordner-Struktur

```
rechnungsbearbeitung/
├── rechnung-processor.js      # Haupt-Code
├── package.json               # Abhängigkeiten
├── node_modules/              # Installierte Pakete
├── RechnungProcessor.exe      # Die erstellte .exe
└── README.md                  # Diese Datei
```

## Technische Details

**Verwendete Bibliotheken:**
- `pdf-parse` - PDF-Text extrahieren
- `xlsx` - Excel-Dateien erstellen

**Systemvoraussetzungen:**
- Windows 10/11 (64-bit empfohlen)
- 100 MB freier Speicher

## Lizenz

MIT - Frei verwendbar.
