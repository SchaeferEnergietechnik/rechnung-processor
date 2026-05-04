#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rechnungsbearbeitung - PDF-Rechnungen verarbeiten (Python Version)

Funktionen:
- Liest PDFs aus dem Ordner, in dem das Script liegt
- Extrahiert: Gesamtpreis (inkl. MwSt), Lieferant, Rechnungsdatum
- Benennt um: 26-xxx_YYYY-MM-DD_Lieferant_Gesamtpreis.pdf
- Schreibt Excel mit Auswertung

Usage: python rechnung_processor.py
    oder als .exe mit PyInstaller
"""

import os
import re
import sys
from pathlib import Path
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print("Fehler: pdfplumber nicht installiert.")
    print("Installieren mit: pip install pdfplumber openpyxl")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("Fehler: openpyxl nicht installiert.")
    print("Installieren mit: pip install openpyxl")
    sys.exit(1)

# Konfiguration
CONFIG = {
    "start_nummer": 1,
    "prefix": "26-",
    "excel_datei": "Rechnungsuebersicht.xlsx"
}

# Bekannte Lieferanten mit korrekten Umlauten
LIEFERANTEN_MAP = {
    'ttcher': 'Böttcher',
    'bottcher': 'Böttcher',
    'stadtwerke wittenberge': 'Stadtwerke Wittenberge',
    'stadtwerke': 'Stadtwerke',
    'wittenberge': 'Stadtwerke Wittenberge',
    're-invent': 'RE-INvent',
    'invent': 'RE-INvent',
    'voelkner': 'Voelkner',
    'steinke': 'Steinke'
}

# Stoppwörter
STOPP_WORTER = ['umsatzsteuer', 'ust-id', 'steuernummer', 'bankverbindung', 'bic', 'iban', 
    'öffnungszeiten', 'montag', 'dienstag', 'mittwoch', 'donnerstag', 'freitag', 
    'strasse', 'ferchlipp', 'lichterfelde', 'altmärkische', 'deutschland', 'g.e.s.', 'energietechnik']


def parse_betrag(betrag_str):
    """Betrag normalisieren (String zu Float)"""
    if not betrag_str:
        return 0.0
    # Deutsche Format: 1.234,56 → 1234.56
    clean = betrag_str.replace('.', '').replace(',', '.')
    try:
        return float(clean)
    except ValueError:
        return 0.0


def format_betrag(betrag):
    """Betrag formatieren für Dateinamen"""
    return f"{betrag:.2f}".replace('.', ',')


def finde_lieferant(text):
    """Extrahiert den Lieferanten aus dem Text (mit Umlaut-Fix)"""
    text_lower = text.lower()
    
    # Prüfe auf bekannte Lieferanten
    for key, value in LIEFERANTEN_MAP.items():
        if key.lower() in text_lower:
            return value
    
    # Suche nach Firmen mit GmbH/AG/KG
    firma_muster = re.compile(r'([A-ZÄÖÜ][a-zäöüßA-ZÄÖÜ\s]{2,40}(?:GmbH|AG|KG|OHG|e\.?K|UG))', re.IGNORECASE)
    match = firma_muster.search(text)
    if match:
        kandidat = match.group(1).strip()
        if not any(wort in text_lower for wort in STOPP_WORTER):
            return bereinige_lieferant(kandidat)
    
    # Erste Zeilen durchgehen
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    for zeile in zeilen[:15]:
        if re.search(r'\b(GmbH|AG|KG|OHG|e\.?K|UG)\b', zeile, re.IGNORECASE):
            bereinigt = bereinige_lieferant(zeile)
            if len(bereinigt) > 3:
                zeile_lower = zeile.lower()
                if not any(wort in zeile_lower for wort in STOPP_WORTER):
                    return bereinigt
    
    return 'Unbekannt'


def bereinige_lieferant(name):
    """Bereinigt den Lieferantennamen für Dateinamen"""
    # Fixe häufige OCR-Fehler bei Umlauten
    fixed = name
    fixed = re.sub(r'ttcher', 'Böttcher', fixed, flags=re.IGNORECASE)
    fixed = re.sub(r'bottcher', 'Böttcher', fixed, flags=re.IGNORECASE)
    fixed = re.sub(r'schaefer', 'Schäfer', fixed, flags=re.IGNORECASE)
    fixed = fixed.replace('ae', 'ä').replace('oe', 'ö').replace('ue', 'ü').replace('ss', 'ß')
    
    # Entferne ungültige Zeichen für Dateinamen
    fixed = re.sub(r'[\\/:*?"<>|]', '_', fixed)
    fixed = re.sub(r'\s+', '_', fixed)
    fixed = re.sub(r'_+', '_', fixed)
    fixed = fixed.strip('_')
    
    return fixed[:40]


def finde_gesamtpreis(text):
    """Findet die Endsumme/Gesamtsumme"""
    muster_prioritaet = [
        r'(?:rechnungsbetrag|zu zahlen|fällig|noch zu zahlend)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        r'(?:gesamtbetrag|endbetrag|summe)[^\d]{0,50}(?:brutto)?[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        r'(?:zahlbar|betrag)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        r'(?:gesamt|total)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})'
    ]
    
    for muster in muster_prioritaet:
        match = re.search(muster, text, re.IGNORECASE)
        if match:
            betrag = parse_betrag(match.group(1))
            if 0 < betrag < 50000:
                return {"betrag": betrag, "original": match.group(1)}
    
    # Fallback: Höchster realistischer Betrag
    betrag_muster = re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2})')
    betraege = []
    
    for match in betrag_muster.finditer(text):
        betrag = parse_betrag(match.group(1))
        if 0 < betrag < 10000:
            betraege.append(betrag)
    
    if betraege:
        max_betrag = max(betraege)
        return {"betrag": max_betrag, "original": format_betrag(max_betrag)}
    
    return None


def finde_datum(text):
    """Extrahiert das Rechnungsdatum"""
    # Zuerst: Suche nach explizitem "Rechnungsdatum"
    rechnungs_datum_muster = re.compile(r'rechnungsdatum[\s:]+(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})', re.IGNORECASE)
    match = rechnungs_datum_muster.search(text)
    if match:
        tag = int(match.group(1))
        monat = int(match.group(2))
        jahr = int(match.group(3))
        if 1 <= tag <= 31 and 1 <= monat <= 12 and 2020 <= jahr <= 2030:
            return {"tag": tag, "monat": monat, "jahr": jahr, "iso": f"{jahr}-{monat:02d}-{tag:02d}"}
    
    # Alternative: Suche nach "Datum" oder "Rechnung" + Datum
    datum_mit_kontext = re.compile(r'(?:datum|rechnung|invoice)[\s\S]{0,30}(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})', re.IGNORECASE)
    for match in datum_mit_kontext.finditer(text):
        tag = int(match.group(1))
        monat = int(match.group(2))
        jahr = int(match.group(3))
        if 1 <= tag <= 31 and 1 <= monat <= 12 and 2020 <= jahr <= 2030:
            context_before = text[max(0, match.start() - 50):match.start()].lower()
            if not re.search(r'verbrauch|zeitraum|von|bis|abrechnung', context_before, re.IGNORECASE):
                return {"tag": tag, "monat": monat, "jahr": jahr, "iso": f"{jahr}-{monat:02d}-{tag:02d}"}
    
    # Fallback: Alle Datumsangaben sammeln
    muster = re.compile(r'(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})')
    daten = []
    
    for match in muster.finditer(text):
        tag = int(match.group(1))
        monat = int(match.group(2))
        jahr = int(match.group(3))
        
        if 1 <= tag <= 31 and 1 <= monat <= 12 and 2020 <= jahr <= 2030:
            context = text[max(0, match.start() - 80):match.end() + 80].lower()
            ist_verbrauch = re.search(r'verbrauch|zeitraum|von.*bis|abrechnungszeitraum', context, re.IGNORECASE)
            ist_rechnung = re.search(r'rechnung|rechnungsdatum|fällig|zahlbar', context, re.IGNORECASE)
            
            gewicht = 3 if ist_rechnung else (0 if ist_verbrauch else 1)
            if gewicht > 0:
                daten.append({"tag": tag, "monat": monat, "jahr": jahr, "gewicht": gewicht})
    
    if not daten:
        return None
    
    # Sortiere nach Gewicht
    zaehler = {}
    for d in daten:
        key = f"{d['jahr']}-{d['monat']:02d}-{d['tag']:02d}"
        zaehler[key] = zaehler.get(key, 0) + d['gewicht']
    
    haeufigstes = max(zaehler.items(), key=lambda x: x[1])[0]
    jahr, monat, tag = map(int, haeufigstes.split('-'))
    return {"tag": tag, "monat": monat, "jahr": jahr, "iso": haeufigstes}


def verarbeite_pdf(datei_pfad):
    """Verarbeitet eine einzelne PDF-Datei"""
    try:
        text = ""
        seiten = 0
        
        with pdfplumber.open(datei_pfad) as pdf:
            seiten = len(pdf.pages)
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        hat_text = len(text.strip()) > 100
        
        if not hat_text:
            return {
                "datei": os.path.basename(datei_pfad),
                "pfad": datei_pfad,
                "gescannt": True,
                "extrahiert": False,
                "hinweis": "Gescanntes PDF - OCR nicht verfügbar"
            }
        
        # Extrahiere alle Daten
        lieferant = finde_lieferant(text)
        gesamtpreis = finde_gesamtpreis(text)
        datum = finde_datum(text)
        
        return {
            "datei": os.path.basename(datei_pfad),
            "pfad": datei_pfad,
            "lieferant": lieferant,
            "gesamtpreis": gesamtpreis["betrag"] if gesamtpreis else 0.0,
            "gesamtpreis_original": gesamtpreis["original"] if gesamtpreis else '-',
            "datum": datum,
            "seiten": seiten,
            "extrahiert": True
        }
        
    except Exception as e:
        return {
            "datei": os.path.basename(datei_pfad),
            "pfad": datei_pfad,
            "fehler": str(e),
            "extrahiert": False
        }


def generiere_dateiname(daten, nummer):
    """Generiert neuen Dateinamen mit Datum"""
    nummer_str = f"{CONFIG['prefix']}{nummer:03d}"
    datum_str = daten["datum"]["iso"] if daten["datum"] else "0000-00-00"
    lieferant_safe = daten["lieferant"]
    preis_str = format_betrag(daten["gesamtpreis"])
    
    return f"{nummer_str}_{datum_str}_{lieferant_safe}_{preis_str}EUR.pdf"


def erstelle_excel(ergebnisse, ziel_pfad):
    """Erstellt die Excel-Datei"""
    wb = Workbook()
    
    # Haupt-Tabelle
    ws_haupt = wb.active
    ws_haupt.title = "Rechnungen"
    
    # Header
    headers = ["Lfd. Nr.", "Rechnungsdatum", "Lieferant", "Gesamtpreis inkl. MwSt (€)", "Anzahl Seiten", "Ursprünglicher Dateiname"]
    for col, header in enumerate(headers, 1):
        cell = ws_haupt.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Daten
    erfolgreiche = [e for e in ergebnisse if e["extrahiert"]]
    for idx, e in enumerate(erfolgreiche, 1):
        row = idx + 1
        ws_haupt.cell(row=row, column=1, value=f"{CONFIG['prefix']}{idx:03d}")
        ws_haupt.cell(row=row, column=2, value=e["datum"]["iso"] if e["datum"] else "-")
        ws_haupt.cell(row=row, column=3, value=e["lieferant"])
        ws_haupt.cell(row=row, column=4, value=format_betrag(e["gesamtpreis"]))
        ws_haupt.cell(row=row, column=5, value=e["seiten"])
        ws_haupt.cell(row=row, column=6, value=e["datei"])
    
    # Spaltenbreiten anpassen
    for col in range(1, 7):
        ws_haupt.column_dimensions[chr(64 + col)].width = 25
    ws_haupt.column_dimensions['C'].width = 30
    
    # Monatsauswertung
    ws_monat = wb.create_sheet("Monatsauswertung")
    monats_daten = {}
    for e in erfolgreiche:
        if e["datum"]:
            key = f"{e['datum']['jahr']}-{e['datum']['monat']:02d}"
            if key not in monats_daten:
                monats_daten[key] = {
                    "jahr": e["datum"]["jahr"],
                    "monat": e["datum"]["monat"],
                    "anzahl": 0,
                    "summe": 0.0
                }
            monats_daten[key]["anzahl"] += 1
            monats_daten[key]["summe"] += e["gesamtpreis"]
    
    # Header Monat
    for col, header in enumerate(["Jahr", "Monat", "Monatsname", "Anzahl Rechnungen", "Gesamtsumme (€)"], 1):
        cell = ws_monat.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Daten Monat
    for idx, (key, m) in enumerate(sorted(monats_daten.items()), 1):
        monatsname = datetime(m["jahr"], m["monat"], 1).strftime("%B")
        ws_monat.cell(row=idx + 1, column=1, value=m["jahr"])
        ws_monat.cell(row=idx + 1, column=2, value=m["monat"])
        ws_monat.cell(row=idx + 1, column=3, value=monatsname)
        ws_monat.cell(row=idx + 1, column=4, value=m["anzahl"])
        ws_monat.cell(row=idx + 1, column=5, value=format_betrag(m["summe"]))
    
    # Jahresauswertung
    ws_jahr = wb.create_sheet("Jahresauswertung")
    jahres_daten = {}
    for e in erfolgreiche:
        jahr = e["datum"]["jahr"] if e["datum"] else datetime.now().year
        if jahr not in jahres_daten:
            jahres_daten[jahr] = {"jahr": jahr, "anzahl": 0, "summe": 0.0}
        jahres_daten[jahr]["anzahl"] += 1
        jahres_daten[jahr]["summe"] += e["gesamtpreis"]
    
    # Header Jahr
    for col, header in enumerate(["Jahr", "Anzahl Rechnungen", "Gesamtsumme (€)"], 1):
        cell = ws_jahr.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Daten Jahr
    for idx, (jahr, j) in enumerate(sorted(jahres_daten.items()), 1):
        ws_jahr.cell(row=idx + 1, column=1, value=j["jahr"])
        ws_jahr.cell(row=idx + 1, column=2, value=j["anzahl"])
        ws_jahr.cell(row=idx + 1, column=3, value=format_betrag(j["summe"]))
    
    wb.save(ziel_pfad)


def main():
    """Hauptfunktion"""
    # Ordner ist automatisch der Ordner, in dem das Script/.exe liegt
    quell_ordner = os.getcwd()
    
    print(f"📁 Verarbeite PDFs in: {quell_ordner}\n")
    
    # Finde alle PDFs (ohne bereits umbenannte)
    prefix_lower = CONFIG["prefix"].lower()
    dateien = [
        os.path.join(quell_ordner, f) for f in os.listdir(quell_ordner)
        if f.lower().endswith('.pdf') and not f.lower().startswith(prefix_lower)
    ]
    
    if not dateien:
        print("⚠️  Keine neuen PDF-Dateien gefunden.")
        print("   (Dateien mit '26-' am Anfang werden übersprungen)")
        return
    
    print(f"📄 {len(dateien)} neue PDF(s) gefunden. Verarbeite...\n")
    
    # Verarbeite alle PDFs
    ergebnisse = []
    for i, pfad in enumerate(dateien, 1):
        datei_name = os.path.basename(pfad)
        print(f"[{i}/{len(dateien)}] {datei_name}...")
        
        daten = verarbeite_pdf(pfad)
        ergebnisse.append(daten)
        
        if daten["extrahiert"]:
            neuer_name = generiere_dateiname(daten, i)
            daten["neuer_name"] = neuer_name
            
            ziel_pfad = os.path.join(quell_ordner, neuer_name)
            
            # Umbenennen
            try:
                os.rename(pfad, ziel_pfad)
                print(f"  ✅ {neuer_name}")
            except Exception as e:
                print(f"  ⚠️  Konnte nicht umbenennen: {e}")
            
            print(f"     Lieferant: {daten['lieferant']}")
            print(f"     Datum: {daten['datum']['iso'] if daten['datum'] else '-'}")
            print(f"     Gesamtpreis: {format_betrag(daten['gesamtpreis'])} €")
        elif daten.get("gescannt"):
            print(f"  ⚠️  Übersprungen: {daten['hinweis']}")
        else:
            print(f"  ❌ Fehler: {daten.get('fehler', 'Unbekannt')}")
        print()
    
    # Excel erstellen
    excel_pfad = os.path.join(quell_ordner, CONFIG["excel_datei"])
    erstelle_excel(ergebnisse, excel_pfad)
    
    # Zusammenfassung
    erfolgreich = len([e for e in ergebnisse if e["extrahiert"]])
    fehler = len([e for e in ergebnisse if not e["extrahiert"]])
    gesamt_summe = sum(e["gesamtpreis"] for e in ergebnisse if e["extrahiert"])
    
    print("\n📊 Zusammenfassung:")
    print(f"   Erfolgreich verarbeitet: {erfolgreich}")
    print(f"   Fehler: {fehler}")
    print(f"   Gesamtsumme: {format_betrag(gesamt_summe)} €")
    print(f"\n📂 Excel erstellt: {CONFIG['excel_datei']}")
    
    # Warte auf Tastendruck
    print("\n-------------------------------------------")
    print("Drücken Sie Enter zum Beenden...")
    input()


if __name__ == "__main__":
    main()
