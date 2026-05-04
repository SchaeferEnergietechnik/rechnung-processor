#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rechnungsbearbeitung - PDF-Rechnungen verarbeiten (Python Version)

Funktionen:
- Liest PDFs aus dem Ordner, in dem das Script liegt
- Extrahiert: Gesamtpreis (inkl. MwSt), Lieferant, Rechnungsdatum
- OCR-Unterstützung für gescannte PDFs (falls Tesseract installiert)
- Benennt um: 26-xxx_YYYY-MM-DD_Lieferant_Gesamtpreis.pdf
- Schreibt Excel mit Auswertung

Usage: python rechnung_processor.py
    oder als .exe mit PyInstaller
"""

import os
import re
import sys
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime
from io import BytesIO

try:
    import pdfplumber
except ImportError:
    print("Fehler: pdfplumber nicht installiert.")
    print("Installieren mit: pip install pdfplumber openpyxl")
    sys.exit(1)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
except ImportError:
    print("Fehler: openpyxl nicht installiert.")
    print("Installieren mit: pip install openpyxl")
    sys.exit(1)

# Optional: OCR-Unterstützung
try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_VERFUEGBAR = True
except ImportError:
    OCR_VERFUEGBAR = False

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
    """Findet den BRUTTO-Gesamtbetrag (inkl. MwSt)"""
    
    # 1. Suche explizit nach BRUTTO-Beträgen
    brutto_muster = [
        # "Gesamtbetrag brutto" oder "Rechnungsbetrag brutto"
        r'(?:gesamtbetrag|rechnungsbetrag|endbetrag|summe)[\s\w]{0,20}(?:brutto|inkl\.?\s*MwSt|inklusive\s*MwSt)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        # "brutto" gefolgt von Betrag in derselben Zeile
        r'brutto[:\s]+[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        # "zu zahlen" oder "fällig" mit Betrag
        r'(?:zu\s*zahlen|fällig|zahlbar|betrag)[\s\w]{0,30}(?:brutto)?[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
    ]
    
    for muster in brutto_muster:
        match = re.search(muster, text, re.IGNORECASE)
        if match:
            betrag = parse_betrag(match.group(1))
            if 0 < betrag < 50000:
                return {"betrag": betrag, "original": match.group(1), "typ": "brutto"}
    
    # 2. Suche nach "Netto" - dann nächsten höheren Betrag als Brutto
    netto_match = re.search(r'netto[:\s]+[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})', text, re.IGNORECASE)
    if netto_match:
        netto_betrag = parse_betrag(netto_match.group(1))
        # Suche alle Beträge nach dem Netto
        position_nach_netto = text[netto_match.end():]
        betraege_nach_netto = []
        for match in re.finditer(r'(\d{1,3}(?:\.\d{3})*,\d{2})', position_nach_netto[:500]):
            b = parse_betrag(match.group(1))
            if netto_betrag * 0.9 < b < netto_betrag * 1.3:  # Ungefähr gleiche Größenordnung
                betraege_nach_netto.append(b)
        if betraege_nach_netto:
            brutto = max(betraege_nach_netto)
            return {"betrag": brutto, "original": format_betrag(brutto), "typ": "berechnet_brutto"}
    
    # 3. Fallback: Suche nach Schlüsselwörtern ohne "brutto"
    fallback_muster = [
        r'(?:rechnungsbetrag|zu\s*zahlen|fällig|noch\s*zu\s*zahlend)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        r'(?:gesamtbetrag|endbetrag)[^\d]{0,50}(?:brutto)?[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
        r'(?:zahlbar|betrag)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})',
    ]
    
    for muster in fallback_muster:
        match = re.search(muster, text, re.IGNORECASE)
        if match:
            betrag = parse_betrag(match.group(1))
            if 0 < betrag < 50000:
                return {"betrag": betrag, "original": match.group(1), "typ": "fallback"}
    
    # 4. Letzter Fallback: Höchster realistischer Betrag
    betrag_muster = re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2})')
    betraege = []
    
    for match in betrag_muster.finditer(text):
        betrag = parse_betrag(match.group(1))
        if 0 < betrag < 10000:
            betraege.append(betrag)
    
    if betraege:
        max_betrag = max(betraege)
        return {"betrag": max_betrag, "original": format_betrag(max_betrag), "typ": "max"}
    
    return None


def finde_datum(text):
    """Extrahiert das Rechnungsdatum"""
    # Zuerst: Suche nach explizitem "Rechnungsdatum"
    rechnungs_datum_muster = re.compile(r'rechnungsdatum[\s:]+(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})', re.IGNORECASE)
    match = re.search(rechnungs_datum_muster, text)
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


def ocr_pdf(datei_pfad):
    """Versucht OCR auf gescanntem PDF durchzuführen"""
    if not OCR_VERFUEGBAR:
        return None
    
    try:
        # Konvertiere PDF zu Bildern
        bilder = convert_from_path(datei_pfad, dpi=300, first_page=1, last_page=3)
        
        text = ""
        for bild in bilder:
            # OCR auf jedes Bild
            seiten_text = pytesseract.image_to_string(bild, lang='deu')
            text += seiten_text + "\n"
        
        return text if len(text.strip()) > 50 else None
    except Exception as e:
        print(f"  OCR-Fehler: {e}")
        return None


def verarbeite_pdf(datei_pfad):
    """Verarbeitet eine einzelne PDF-Datei"""
    try:
        text = ""
        seiten = 0
        ist_gescannt = False
        
        # Versuche normale Textextraktion
        with pdfplumber.open(datei_pfad) as pdf:
            seiten = len(pdf.pages)
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        # Wenig Text gefunden? Versuche OCR
        if len(text.strip()) < 100:
            ist_gescannt = True
            print("  📄 Wenig Text gefunden, versuche OCR...")
            ocr_text = ocr_pdf(datei_pfad)
            if ocr_text:
                text = ocr_text
                ist_gescannt = False  # OCR hat funktioniert
                print("  ✅ OCR erfolgreich")
            else:
                if OCR_VERFUEGBAR:
                    return {
                        "datei": os.path.basename(datei_pfad),
                        "pfad": datei_pfad,
                        "gescannt": True,
                        "extrahiert": False,
                        "hinweis": "Gescanntes PDF - OCR fehlgeschlagen (Tesseract installiert?)"
                    }
                else:
                    return {
                        "datei": os.path.basename(datei_pfad),
                        "pfad": datei_pfad,
                        "gescannt": True,
                        "extrahiert": False,
                        "hinweis": "Gescanntes PDF - OCR nicht verfügbar (pip install pdf2image pytesseract)"
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
            "gesamtpreis_typ": gesamtpreis["typ"] if gesamtpreis else '-',
            "datum": datum,
            "seiten": seiten,
            "ocr_verwendet": ist_gescannt,
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


def erstelle_excel(ergebnisse, ziel_pfad, quell_ordner):
    """Aktualisiert oder erstellt die Excel-Datei mit fortlaufender Liste"""
    
    # Lade bestehende Excel wenn vorhanden
    bestehende_eintraege = []
    if os.path.exists(ziel_pfad):
        try:
            wb_alt = load_workbook(ziel_pfad)
            ws_alt = wb_alt["Rechnungen"]
            
            # Lese bestehende Einträge (ab Zeile 2, da Zeile 1 = Header)
            for row in ws_alt.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1] and row[2]:  # Lfd. Nr., Datum, Lieferant
                    bestehende_eintraege.append({
                        "lfd_nr": row[0],
                        "datum": row[1],
                        "lieferant": row[2],
                        "gesamtpreis": row[3],
                        "seiten": row[4],
                        "ocr": row[5],
                        "original": row[6]
                    })
            print(f"📊 {len(bestehende_eintraege)} bestehende Einträge geladen")
        except Exception as e:
            print(f"⚠️  Konnte bestehende Excel nicht laden: {e}")
    
    # Sammle neue erfolgreiche Einträge
    neue_erfolgreiche = [e for e in ergebnisse if e["extrahiert"]]
    
    # Finde höchste vorhandene Nummer
    max_nummer = 0
    for eintrag in bestehende_eintraege:
        try:
            nummer = int(str(eintrag["lfd_nr"]).replace(CONFIG["prefix"], ""))
            max_nummer = max(max_nummer, nummer)
        except:
            pass
    
    # Erstelle neue Einträge mit fortlaufender Nummerierung
    neue_eintraege = []
    for idx, e in enumerate(neue_erfolgreiche, max_nummer + 1):
        neue_eintraege.append({
            "lfd_nr": f"{CONFIG['prefix']}{idx:03d}",
            "datum": e["datum"]["iso"] if e["datum"] else "-",
            "lieferant": e["lieferant"],
            "gesamtpreis": format_betrag(e["gesamtpreis"]),
            "seiten": e["seiten"],
            "ocr": "Ja" if e.get("ocr_verwendet") else "Nein",
            "original": e["datei"]
        })
    
    # Alle Einträge kombinieren
    alle_eintraege = bestehende_eintraege + neue_eintraege
    
    # Sortiere nach Datum (neueste zuerst)
    def sort_key(e):
        try:
            return e["datum"] if e["datum"] != "-" else "0000-00-00"
        except:
            return "0000-00-00"
    
    alle_eintraege.sort(key=sort_key, reverse=True)
    
    # Erstelle neues Workbook
    wb = Workbook()
    
    # Haupt-Tabelle
    ws_haupt = wb.active
    ws_haupt.title = "Rechnungen"
    
    # Header
    headers = ["Lfd. Nr.", "Rechnungsdatum", "Lieferant", "Gesamtpreis inkl. MwSt (€)", "Anzahl Seiten", "OCR verwendet", "Ursprünglicher Dateiname"]
    for col, header in enumerate(headers, 1):
        cell = ws_haupt.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Alle Daten schreiben
    for idx, e in enumerate(alle_eintraege, 1):
        row = idx + 1
        ws_haupt.cell(row=row, column=1, value=e["lfd_nr"])
        ws_haupt.cell(row=row, column=2, value=e["datum"])
        ws_haupt.cell(row=row, column=3, value=e["lieferant"])
        ws_haupt.cell(row=row, column=4, value=e["gesamtpreis"])
        ws_haupt.cell(row=row, column=5, value=e["seiten"])
        ws_haupt.cell(row=row, column=6, value=e["ocr"])
        ws_haupt.cell(row=row, column=7, value=e["original"])
    
    # Spaltenbreiten anpassen
    ws_haupt.column_dimensions['A'].width = 12
    ws_haupt.column_dimensions['B'].width = 18
    ws_haupt.column_dimensions['C'].width = 30
    ws_haupt.column_dimensions['D'].width = 25
    ws_haupt.column_dimensions['E'].width = 15
    ws_haupt.column_dimensions['F'].width = 15
    ws_haupt.column_dimensions['G'].width = 40
    
    # Monatsauswertung (aus allen Einträgen)
    ws_monat = wb.create_sheet("Monatsauswertung")
    monats_daten = {}
    
    for e in alle_eintraege:
        if e["datum"] and e["datum"] != "-":
            try:
                datum_parts = e["datum"].split("-")
                jahr = int(datum_parts[0])
                monat = int(datum_parts[1])
                key = f"{jahr}-{monat:02d}"
                
                if key not in monats_daten:
                    monats_daten[key] = {
                        "jahr": jahr,
                        "monat": monat,
                        "anzahl": 0,
                        "summe": 0.0
                    }
                monats_daten[key]["anzahl"] += 1
                # Parse Betrag (Deutsch: 1.234,56 → 1234.56)
                betrag_str = str(e["gesamtpreis"]).replace(".", "").replace(",", ".")
                monats_daten[key]["summe"] += float(betrag_str)
            except:
                pass
    
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
    
    for e in alle_eintraege:
        if e["datum"] and e["datum"] != "-":
            try:
                jahr = int(e["datum"].split("-")[0])
                if jahr not in jahres_daten:
                    jahres_daten[jahr] = {"jahr": jahr, "anzahl": 0, "summe": 0.0}
                jahres_daten[jahr]["anzahl"] += 1
                betrag_str = str(e["gesamtpreis"]).replace(".", "").replace(",", ".")
                jahres_daten[jahr]["summe"] += float(betrag_str)
            except:
                pass
    
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
    
    return len(neue_eintraege), len(alle_eintraege)  # Neue, Gesamt


def main():
    """Hauptfunktion"""
    # Ordner ist automatisch der Ordner, in dem das Script/.exe liegt
    quell_ordner = os.getcwd()
    
    print(f"📁 Verarbeite PDFs in: {quell_ordner}\n")
    
    if OCR_VERFUEGBAR:
        print("✅ OCR-Unterstützung verfügbar (für gescannte PDFs)\n")
    else:
        print("ℹ️  OCR nicht verfügbar - nur digitale PDFs können verarbeitet werden")
        print("   (pip install pdf2image pytesseract für OCR-Support)\n")
    
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
            print(f"     Gesamtpreis: {format_betrag(daten['gesamtpreis'])} € (Brutto)")
            if daten.get("ocr_verwendet"):
                print(f"     ℹ️  OCR verwendet")
        elif daten.get("gescannt"):
            print(f"  ⚠️  Übersprungen: {daten['hinweis']}")
        else:
            print(f"  ❌ Fehler: {daten.get('fehler', 'Unbekannt')}")
        print()
    
    # Excel erstellen
    excel_pfad = os.path.join(quell_ordner, CONFIG["excel_datei"])
    neue_eintraege, gesamt_eintraege = erstelle_excel(ergebnisse, excel_pfad, quell_ordner)
    
    # Zusammenfassung
    erfolgreich = len([e for e in ergebnisse if e["extrahiert"]])
    fehler = len([e for e in ergebnisse if not e["extrahiert"]])
    gesamt_summe = sum(e["gesamtpreis"] for e in ergebnisse if e["extrahiert"])
    
    print("\n📊 Zusammenfassung:")
    print(f"   Neue Rechnungen: {erfolgreich}")
    print(f"   Fehler: {fehler}")
    print(f"   Gesamtsumme (neu): {format_betrag(gesamt_summe)} €")
    print(f"\n📂 Excel aktualisiert: {CONFIG['excel_datei']}")
    print(f"   {neue_eintraege} neue Einträge hinzugefügt")
    print(f"   {gesamt_eintraege} Einträge insgesamt")
    
    # Warte auf Tastendruck
    print("\n-------------------------------------------")
    print("Drücken Sie Enter zum Beenden...")
    input()


if __name__ == "__main__":
    main()
