# Rechnungsbearbeitung

PDF-Rechnungen automatisch umbenennen mit Daten aus dem PDF.

## Was macht das Tool?

- Liest alle PDFs im Ordner (wo die .exe liegt)
- Extrahiert: **Datum**, **Lieferant**, **Gesamtpreis (inkl. MwSt)**
- Benennt um: `26-001_2026-04-21_Böttcher_202,31EUR.pdf`
- Erstellt: `Rechnungsuebersicht.xlsx` mit Monats- und Jahresauswertung

## Download

**[⬇️ RechnungProcessor.exe herunterladen](../../releases/latest)**

Oder unter [Actions](../../actions) → Letzter Run → Artifacts

## Nutzung

1. `RechnungProcessor.exe` in Ordner mit PDFs kopieren
2. Doppelklick
3. Fertig! PDFs werden umbenannt, Excel wird erstellt

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

## Entwicklung

### Als Python-Script ausführen
```bash
pip install -r requirements.txt
python rechnung_processor.py
```

### Selbst als .exe bauen
```bash
pip install pyinstaller
pyinstaller --onefile --console rechnung_processor.py
```

## Automatischer Build

Bei jedem Push zu `main` wird automatisch eine .exe gebaut.

Siehe [.github/workflows/build.yml](.github/workflows/build.yml)

## Lizenz

MIT - Frei verwendbar.
