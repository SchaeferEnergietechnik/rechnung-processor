#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Rechnungsbearbeitung - PDF-Rechnungen verarbeiten (Python Version)

Funktionen:
- Liest PDFs aus dem Ordner, in dem das Script liegt
- Extrahiert: Gesamtpreis (inkl. MwSt), Lieferant, Rechnungsdatum
- OCR-Unterstützung für gescannte PDFs (falls Tesseract installiert)
- Benennt um: 26-xxx_Lieferant_YYYY-MM-DD_Gesamtpreis.pdf
- Schreibt Excel mit Auswertung, deutschem Datum und Monatssummen

Usage: python rechnung_processor.py
    oder als .exe mit PyInstaller
"""

import os
import re
import sys
import tempfile
import subprocess
import json
from pathlib import Path
from datetime import datetime
from io import BytesIO
from collections import defaultdict

try:
    import requests
    REQUESTS_VERFUEGBAR = True
except ImportError:
    REQUESTS_VERFUEGBAR = False
    print("⚠️  requests nicht installiert. Online-OCR deaktiviert.")
    print("   Installieren mit: pip install requests")

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
    "excel_datei": "Rechnungsuebersicht.xlsx",
    "ocr_space_api_key": "K83767756488957"  # Free OCR API Key
}

# Bekannte Lieferanten mit korrekten Umlauten - SPEZIFISCHE zuerst!
LIEFERANTEN_MAP = {
    # Spezifische Treffer zuerst (längere/genaue Strings)
    're-invent retail': 'RE-INvent Retail',
    'voelkner': 'Voelkner',
    'steinke': 'Steinke',
    'ttcher': 'Böttcher',
    'bottcher': 'Böttcher',
    'stadtwerke wittenberge': 'Stadtwerke Wittenberge',
    'amazon': 'Amazon',
    # Allgemeine Treffer zuletzt
    'stadtwerke': 'Stadtwerke',
    'wittenberge': 'Stadtwerke Wittenberge',
    'invent': 'RE-INvent',
}

# Stoppwörter - ERWEITERT
STOPP_WORTER = ['umsatzsteuer', 'ust-id', 'steuernummer', 'bankverbindung', 'bic', 'iban', 
    'öffnungszeiten', 'montag', 'dienstag', 'mittwoch', 'donnerstag', 'freitag', 
    'strasse', 'ferchlipp', 'lichterfelde', 'altmärkische', 'deutschland', 'g.e.s.', 'energietechnik',
    'gesamtbetrag', 'rechnungsbetrag', 'endbetrag', 'brutto', 'netto', 'zzgl', 'zuzüglich',
    'abzüglich', 'abschlag', 'zahlung', 'betrag', 'summe', 'bruttowert', 'nettowert']


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
    """Extrahiert den Lieferanten aus dem Text - verbesserte Version v3"""
    text_lower = text.lower()
    lines = text.split('\n')
    
    # HARTE REGEL: Bekannte Lieferanten direkt erkennen (Case-sensitive Teile)
    if 'stadtwerke wittenberge gmbh' in text_lower:
        return 'Stadtwerke_Wittenberge'
    if 'böttcher ag' in text_lower:
        return 'Böttcher'
    if 'bottcher ag' in text_lower:
        return 'Böttcher'
    if 're-invent retail' in text_lower or 're-invent' in text_lower:
        return 'RE-INvent_Retail'
    if 'voelkner' in text_lower:
        return 'Voelkner'
    if 'steinke' in text_lower:
        return 'Steinke'
    if 'ttt-filmserivce' in text_lower or 'filmserivce' in text_lower:
        return 'TTT-Filmservice'
    
    # AMAZON SPEZIAL: Wenn "Amazon" im Text ist, immer Amazon als Lieferant
    # (egal wer "Verkauft von" ist)
    if 'amazon eu' in text_lower or 'amazon' in text_lower:
        # Prüfe ob es wirklich eine Amazon-Rechnung ist
        if 'bestellnummer' in text_lower or 'amazon.de' in text_lower or 'verkauft von' in text_lower:
            return 'Amazon'
    
    # Suche nach explizitem "Rechnung" + Firma
    if 're-invent retail' in text_lower or 're-invent' in text_lower:
        return 'RE-INvent_Retail'
    if 'voelkner' in text_lower:
        return 'Voelkner'
    if 'steinke' in text_lower:
        return 'Steinke'
    
    # Suche nach explizitem "Rechnung" + Firma
    rechnung_pattern = re.search(r'ReCHNUNG\s*\n?\s*([A-Z][A-Za-zäöüß\s&\.]+(?:GmbH|AG|KG|OHG|e\.?K|UG))', text, re.IGNORECASE)
    if rechnung_pattern:
        kandidat = rechnung_pattern.group(1).strip()
        if ist_gueltiger_lieferant(kandidat):
            return bereinige_lieferant(kandidat)
    
    # Suche nach Firmen die G.E.S. beliefert (Firma vor oder nach G.E.S.)
    # Muster: Firma GmbH ... G.E.S. Energietechnik
    ges_pattern = re.search(r'([A-Z][A-Za-zäöüß\s&\.]+(?:GmbH|AG|KG|OHG|e\.?K|UG))[^.]{0,200}G\.E\.S', text, re.IGNORECASE | re.DOTALL)
    if ges_pattern:
        kandidat = ges_pattern.group(1).strip()
        if ist_gueltiger_lieferant(kandidat):
            return bereinige_lieferant(kandidat)
    
    # Erste 10 Zeilen = meist Absender
    absender_section = '\n'.join(lines[:10])
    firmen_absender = re.findall(r'([A-Z][A-Za-zäöüß\s&\.]+(?:GmbH|AG|KG|OHG|e\.?K|UG))', absender_section)
    for firma in firmen_absender:
        firma_clean = firma.strip()
        if ist_gueltiger_lieferant(firma_clean):
            return bereinige_lieferant(firma_clean)
    
    # Fallback: Erste Zeilen durchgehen
    for zeile in lines[:15]:
        if re.search(r'\b(GmbH|AG|KG|OHG|e\.?K|UG)\b', zeile, re.IGNORECASE):
            bereinigt = bereinige_lieferant(zeile)
            if ist_gueltiger_lieferant(bereinigt):
                return bereinigt
    
    return 'Unbekannt'


def ist_gueltiger_lieferant(name):
    """Prüft ob ein Name ein gültiger Lieferant ist (kein Stoppwort)"""
    if len(name) < 5:
        return False
    name_lower = name.lower()
    # Prüfe Stoppwörter
    for stop in STOPP_WORTER:
        if stop in name_lower:
            return False
    # Prüfe dass es keine reinen Betragsbezeichnungen sind
    if re.match(r'^(gesamt|rechnungs|end|brutto|netto|summe)', name_lower):
        return False
    return True


def bereinige_lieferant(name):
    """Bereinigt den Lieferantennamen für Dateinamen"""
    # Entferne Leerzeichen am Anfang/Ende
    fixed = name.strip()
    
    # Fixe häufige OCR-Fehler bei Umlauten (ERST spezifische, DANN generische!)
    fixed = re.sub(r'ttcher', 'ttcher', fixed, flags=re.IGNORECASE)  # Placeholder
    fixed = re.sub(r'bottcher', 'Böttcher', fixed, flags=re.IGNORECASE)
    fixed = re.sub(r'schaefer', 'Schäfer', fixed, flags=re.IGNORECASE)
    
    # Generische Umlaut-Fixe NUR wenn nicht schon Umlaute vorhanden
    if 'ä' not in fixed and 'ö' not in fixed and 'ü' not in fixed:
        fixed = fixed.replace('ae', 'ä').replace('oe', 'ö').replace('ue', 'ü').replace('ss', 'ß')
    
    # Entferne ungültige Zeichen für Dateinamen
    fixed = re.sub(r'[\\/:*?"<>|]', '_', fixed)
    fixed = re.sub(r'\s+', '_', fixed)
    fixed = re.sub(r'_+', '_', fixed)
    fixed = fixed.strip('_')
    
    return fixed[:40]


def finde_gesamtpreis(text):
    """Findet den BRUTTO-Gesamtbetrag (inkl. MwSt) - verbesserte Version v3"""
    text_lower = text.lower()
    alle_kandidaten = []
    
    # 0. HÖCHSTE Priorität: "Summe:" oder "Gesamtpreis" am Ende der Rechnung
    # Dies ist oft der finale Betrag
    hoch_prio_patterns = [
        (r'(?:summe|gesamtpreis|zahlbetrag)[\s:]+(\d{1,3}(?:\.\d{3})*,\d{2})\s*eur', 110),  # "Summe: 24,75 EUR"
        (r'gesamtpreis\s+(\d{1,3}(?:\.\d{3})*,\d{2})', 110),  # Amazon: "Gesamtpreis 9,47"
        (r'(?:noch\s+zu\s+zahlender\s+betrag|rechnungsbetrag)[^\d]{0,50}(\d{1,3}(?:\.\d{3})*,\d{2})', 100),
        (r'(?:zu\s*zahlen|fällig)[^\d]{0,50}(\d{1,3}(?:\.\d{3})*,\d{2})', 90),
        (r'zahlbar[^\d]{0,100}(\d{1,3}(?:\.\d{3})*,\d{2})', 80),
    ]
    
    for pattern, prio in hoch_prio_patterns:
        matches = list(re.finditer(pattern, text, re.IGNORECASE))
        for m in matches:
            betrag = parse_betrag(m.group(1))
            if 0 < betrag < 50000:
                # Prüfe Kontext - vermeide "Versandkosten", "Aktionsrabatt" etc.
                context_start = max(0, m.start() - 30)
                context = text[context_start:m.start()].lower()
                if any(bad in context for bad in ['versandkosten', 'rabatt', 'discount', 'einzelpreis', 'zwischensumme']):
                    continue  # Überspringe diese Treffer
                alle_kandidaten.append((betrag, prio, m.group(1), "endbetrag"))
    
    # 1. "Summe Bruttowert" (Voelkner-Style)
    match = re.search(r'summe\s+bruttowert[:\s]*(\d{1,3}(?:\.\d{3})*,\d{2})', text, re.IGNORECASE)
    if match:
        betrag = parse_betrag(match.group(1))
        if 0 < betrag < 50000:
            alle_kandidaten.append((betrag, 95, match.group(1), "bruttowert"))
    
    # 2. "Gesamtbetrag brutto" / "Endbetrag"
    brutto_patterns = [
        r'(?:endbetrag|gesamtbetrag)\s+brutto[:\s]*(\d{1,3}(?:\.\d{3})*,\d{2})',
    ]
    for pattern in brutto_patterns:
        matches = list(re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE))
        for m in matches:
            betrag = parse_betrag(m.group(1))
            if 0 < betrag < 50000:
                alle_kandidaten.append((betrag, 85, m.group(1), "brutto_explicit"))
    
    # 3. Generische "Gesamtbetrag" - sammle ALLE und nimm den HÖCHSTEN
    gesamt_matches = list(re.finditer(r'gesamtbetrag[:\s]*(\d{1,3}(?:\.\d{3})*,\d{2})', text, re.IGNORECASE))
    if len(gesamt_matches) > 1:
        betraege = [(parse_betrag(m.group(1)), m.group(1)) for m in gesamt_matches if 0 < parse_betrag(m.group(1)) < 50000]
        if betraege:
            max_betrag, original = max(betraege, key=lambda x: x[0])
            alle_kandidaten.append((max_betrag, 70, original, "gesamtbetrag_max"))
    elif gesamt_matches:
        betrag = parse_betrag(gesamt_matches[0].group(1))
        if 0 < betrag < 50000:
            alle_kandidaten.append((betrag, 70, gesamt_matches[0].group(1), "gesamtbetrag"))
    
    # 4. Wenn "Netto" gefunden → berechne Brutto (Netto * 1.19)
    netto_matches = list(re.finditer(r'(?:netto|warenwert)[^\d]{0,50}(\d{1,3}(?:\.\d{3})*,\d{2})', text, re.IGNORECASE))
    if netto_matches:
        netto_betraege = [(m, parse_betrag(m.group(1))) for m in netto_matches if 0 < parse_betrag(m.group(1)) < 50000]
        if netto_betraege:
            netto_match, netto_betrag = max(netto_betraege, key=lambda x: x[1])
            brutto_berechnet = round(netto_betrag * 1.19, 2)
            
            for betrag, prio, orig, typ in alle_kandidaten[:]:
                if abs(betrag - brutto_berechnet) < 0.5:
                    alle_kandidaten.remove((betrag, prio, orig, typ))
                    alle_kandidaten.append((betrag, 95, orig, "brutto_verifiziert"))
    
    # 5. Fallback: Höchster Betrag in der ganzen Rechnung
    all_betraege = []
    for m in re.finditer(r'(\d{1,3}(?:\.\d{3})*,\d{2})', text):
        b = parse_betrag(m.group(1))
        if 0 < b < 50000:
            all_betraege.append(b)
    if all_betraege:
        max_betrag = max(all_betraege)
        if not any(prio >= 80 for betrag, prio, orig, typ in alle_kandidaten):
            alle_kandidaten.append((max_betrag, 10, format_betrag(max_betrag), "max"))
    
    # Wähle den besten Kandidaten
    if alle_kandidaten:
        alle_kandidaten.sort(key=lambda x: (x[1], x[0]), reverse=True)
        best_betrag, best_prio, best_orig, best_typ = alle_kandidaten[0]
        return {"betrag": best_betrag, "original": best_orig, "typ": best_typ}
    
    return None


def finde_datum(text):
    """Extrahiert das Rechnungsdatum - verbesserte Version"""
    text_lower = text.lower()
    
    # 1. HÖCHSTE Priorität: Explizites "Rechnungsdatum" oder "Rechnung ... Datum"
    # Muster: "Rechnungsdatum: 21.04.2026" oder "Rechnung Nr. ... Datum: 21.04.2026"
    rechnungs_patterns = [
        r'ReCHNUNGSDATUM[:\s]+(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})',
        r'ReCHNUNG[^\n]{0,100}DATUM[:\s]+(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})',
        r'DATUM[:\s]+(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})[^\n]{0,50}RECHNUNG',
    ]
    
    for pattern in rechnungs_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            tag = int(match.group(1))
            monat = int(match.group(2))
            jahr = int(match.group(3))
            if 1 <= tag <= 31 and 1 <= monat <= 12 and 2020 <= jahr <= 2030:
                return {"tag": tag, "monat": monat, "jahr": jahr, "iso": f"{jahr}-{monat:02d}-{tag:02d}"}
    
    # 2. Suche nach "Datum" in Verbindung mit "Rechnung" in der Nähe
    # Prüfe alle "Datum" Vorkommen
    datum_matches = list(re.finditer(r'(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})', text))
    beste_kandidaten = []
    
    for match in datum_matches:
        tag = int(match.group(1))
        monat = int(match.group(2))
        jahr = int(match.group(3))
        
        if not (1 <= tag <= 31 and 1 <= monat <= 12 and 2020 <= jahr <= 2030):
            continue
        
        # Kontext um das Datum prüfen
        context_start = max(0, match.start() - 100)
        context_end = min(len(text), match.end() + 100)
        context = text[context_start:context_end].lower()
        
        # Prio basierend auf Kontext
        prio = 0
        if 'rechnungsdatum' in context:
            prio = 100
        elif 'rechnung' in context and 'datum' in context:
            prio = 90
        elif 'rechnung' in context:
            prio = 80
        elif 'datum' in context:
            prio = 70
        else:
            prio = 10  # Fallback
        
        # Reduziere Prio bei Verbrauchs-/Zeitraum-Kontext
        if re.search(r'verbrauch|zeitraum|von.*bis|abrechnungszeitraum|lieferzeitraum', context):
            prio -= 50
        
        # Reduziere Prio bei "Auftragsdatum", "Lieferscheindatum", etc.
        if re.search(r'auftragsdatum|lieferscheindatum|versanddatum|bestelldatum', context):
            prio -= 30
        
        if prio > 0:
            beste_kandidaten.append((prio, jahr, monat, tag, f"{jahr}-{monat:02d}-{tag:02d}"))
    
    # Wähle den besten Kandidaten (höchste Prio)
    if beste_kandidaten:
        beste_kandidaten.sort(reverse=True)
        _, jahr, monat, tag, iso = beste_kandidaten[0]
        return {"tag": tag, "monat": monat, "jahr": jahr, "iso": iso}
    
    return None


def ocr_pdf_online(datei_pfad):
    """Verwendet OCR.space API für gescannte PDFs (kostenlos)"""
    if not REQUESTS_VERFUEGBAR:
        return None
    
    api_key = CONFIG.get("ocr_space_api_key", "")
    if not api_key:
        return None
    
    # Prüfe Dateigröße (max 1MB für Free-Tier)
    datei_groesse = os.path.getsize(datei_pfad)
    if datei_groesse > 1024 * 1024:  # 1 MB
        print(f"  ⚠️  Datei zu groß für Free-OCR ({datei_groesse / 1024 / 1024:.2f} MB > 1 MB)")
        return None
    
    try:
        print("  🌐 Versuche Online-OCR (OCR.space)...")
        
        # API Endpoint
        url = "https://api.ocr.space/parse/image"
        
        # Datei hochladen
        with open(datei_pfad, 'rb') as f:
            response = requests.post(
                url,
                files={"file": (os.path.basename(datei_pfad), f)},
                data={
                    "apikey": api_key,
                    "language": "ger",
                    "OCREngine": "2",  # Bessere Erkennung
                    "isTable": "false",
                    "detectOrientation": "true",
                    "scale": "true"
                },
                timeout=120  # Erhöht auf 120 Sekunden
            )
        
        if response.status_code == 200:
            result = response.json()
            
            if result.get("OCRExitCode") == 1:
                parsed_text = ""
                for parsed_result in result.get("ParsedResults", []):
                    parsed_text += parsed_result.get("ParsedText", "") + "\n"
                
                if parsed_text.strip():
                    print("  ✅ Online-OCR erfolgreich")
                    return parsed_text
                else:
                    print("  ⚠️  Online-OCR: Kein Text erkannt")
                    return None
            else:
                error_msg = result.get("ErrorMessage", ["Unbekannter Fehler"])[0]
                print(f"  ⚠️  Online-OCR Fehler: {error_msg}")
                return None
        else:
            print(f"  ⚠️  Online-OCR HTTP Fehler: {response.status_code}")
            return None
            
    except requests.exceptions.Timeout:
        print("  ⚠️  Online-OCR Timeout (120s) - Datei möglicherweise zu groß")
        return None
    except Exception as e:
        print(f"  ⚠️  Online-OCR Fehler: {e}")
        return None


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
            
            # 1. Versuche lokales OCR (Tesseract)
            ocr_text = ocr_pdf(datei_pfad)
            if ocr_text:
                text = ocr_text
                ist_gescannt = False  # OCR hat funktioniert
                print("  ✅ Lokales OCR erfolgreich (Tesseract)")
            else:
                # 2. Versuche Online-OCR (OCR.space)
                ocr_text = ocr_pdf_online(datei_pfad)
                if ocr_text:
                    text = ocr_text
                    ist_gescannt = True  # Online-OCR war nötig
                    print("  ✅ Online-OCR erfolgreich (OCR.space)")
                else:
                    # Beide OCR-Methoden fehlgeschlagen
                    if OCR_VERFUEGBAR or REQUESTS_VERFUEGBAR:
                        return {
                            "datei": os.path.basename(datei_pfad),
                            "pfad": datei_pfad,
                            "gescannt": True,
                            "extrahiert": False,
                            "hinweis": "Gescanntes PDF - OCR fehlgeschlagen (lokal und online)"
                        }
                    else:
                        return {
                            "datei": os.path.basename(datei_pfad),
                            "pfad": datei_pfad,
                            "gescannt": True,
                            "extrahiert": False,
                            "hinweis": "Gescanntes PDF - keine OCR verfügbar (lokal: pip install pdf2image pytesseract | online: pip install requests)"
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
    """Generiert neuen Dateinamen: 26-001_Lieferant_2026-04-21_Summe.pdf"""
    nummer_str = f"{CONFIG['prefix']}{nummer:03d}"
    lieferant_safe = daten["lieferant"]
    datum_str = daten["datum"]["iso"] if daten["datum"] else "0000-00-00"
    preis_str = format_betrag(daten["gesamtpreis"])
    
    return f"{nummer_str}_{lieferant_safe}_{datum_str}_{preis_str}EUR.pdf"


def erstelle_excel(ergebnisse, ziel_pfad, quell_ordner):
    """Aktualisiert oder erstellt die Excel-Datei mit fortlaufender Liste"""
    from collections import defaultdict
    
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
                        "datum_iso": row[1],  # ISO Format für Sortierung
                        "datum_de": row[2],    # Deutsch Format
                        "lieferant": row[3],
                        "gesamtpreis": row[4],
                        "seiten": row[5],
                        "original": row[6],
                        "neu_name": row[7] if len(row) > 7 else ""
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
        lfd_nr = f"{CONFIG['prefix']}{idx:03d}"
        datum_iso = e["datum"]["iso"] if e["datum"] else "-"
        # Deutsch Format: DD.MM.YYYY
        if datum_iso != "-":
            parts = datum_iso.split("-")
            datum_de = f"{parts[2]}.{parts[1]}.{parts[0]}"
        else:
            datum_de = "-"
        
        # Neuer Dateiname für Spalte
        lieferant_safe = e["lieferant"]
        preis_str = format_betrag(e["gesamtpreis"])
        neu_name = f"{lfd_nr}_{lieferant_safe}_{datum_iso}_{preis_str}EUR.pdf"
        
        neue_eintraege.append({
            "lfd_nr": lfd_nr,
            "datum_iso": datum_iso,
            "datum_de": datum_de,
            "lieferant": e["lieferant"],
            "gesamtpreis": format_betrag(e["gesamtpreis"]),
            "seiten": e["seiten"],
            "original": e["datei"],
            "neu_name": neu_name
        })
    
    # Alle Einträge kombinieren
    alle_eintraege = bestehende_eintraege + neue_eintraege
    
    # Sortiere nach Datum (älteste zuerst, dann Lfd. Nr. aufsteigend)
    def sort_key(e):
        datum = e["datum_iso"] if e["datum_iso"] != "-" else "0000-00-00"
        lfd = e["lfd_nr"]
        return (datum, lfd)
    
    alle_eintraege.sort(key=sort_key)
    
    # Gruppiere nach Jahr-Monat für Summen
    monats_gruppen = defaultdict(list)
    for e in alle_eintraege:
        if e["datum_iso"] != "-":
            key = e["datum_iso"][:7]  # YYYY-MM
            monats_gruppen[key].append(e)
        else:
            monats_gruppen["0000-00"].append(e)
    
    # Erstelle neues Workbook
    wb = Workbook()
    
    # Haupt-Tabelle
    ws_haupt = wb.active
    ws_haupt.title = "Rechnungen"
    
    # Header
    headers = ["Lfd. Nr.", "Rechnungsdatum (ISO)", "Rechnungsdatum", "Lieferant", 
               "Gesamtpreis inkl. MwSt (€)", "Anzahl Seiten", 
               "Ursprünglicher Dateiname", "Neuer Dateiname"]
    for col, header in enumerate(headers, 1):
        cell = ws_haupt.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    # Alle Daten schreiben - nach Monat gruppiert mit Summen
    current_row = 2
    monats_summen = {}
    
    for monat_key in sorted(monats_gruppen.keys()):
        eintraege = monats_gruppen[monat_key]
        monats_summe = 0.0
        
        for e in eintraege:
            ws_haupt.cell(row=current_row, column=1, value=e["lfd_nr"])
            ws_haupt.cell(row=current_row, column=2, value=e["datum_iso"])
            ws_haupt.cell(row=current_row, column=3, value=e["datum_de"])
            ws_haupt.cell(row=current_row, column=4, value=e["lieferant"])
            ws_haupt.cell(row=current_row, column=5, value=e["gesamtpreis"])
            ws_haupt.cell(row=current_row, column=6, value=e["seiten"])
            ws_haupt.cell(row=current_row, column=7, value=e["original"])
            ws_haupt.cell(row=current_row, column=8, value=e["neu_name"])
            
            # Summe berechnen
            try:
                betrag_str = str(e["gesamtpreis"]).replace(".", "").replace(",", ".")
                monats_summe += float(betrag_str)
            except:
                pass
            
            current_row += 1
        
        # Monatssumme einfügen
        monats_name = monat_key
        if monat_key != "0000-00":
            parts = monat_key.split("-")
            monats_name = f"{parts[1]}.{parts[0]}"
        
        sum_cell = ws_haupt.cell(row=current_row, column=1, value=f"Summe {monats_name}")
        sum_cell.font = Font(bold=True)
        sum_cell.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
        
        sum_value = ws_haupt.cell(row=current_row, column=5, value=format_betrag(monats_summe))
        sum_value.font = Font(bold=True)
        sum_value.fill = PatternFill(start_color="E6F3FF", end_color="E6F3FF", fill_type="solid")
        
        monats_summen[monat_key] = monats_summe
        current_row += 1
    
    # Spaltenbreiten anpassen
    ws_haupt.column_dimensions['A'].width = 12
    ws_haupt.column_dimensions['B'].width = 20  # ISO Datum (hidden später)
    ws_haupt.column_dimensions['C'].width = 18  # Deutsch
    ws_haupt.column_dimensions['D'].width = 30
    ws_haupt.column_dimensions['E'].width = 25
    ws_haupt.column_dimensions['F'].width = 15
    ws_haupt.column_dimensions['G'].width = 45
    ws_haupt.column_dimensions['H'].width = 50
    
    # ISO-Spalte verstecken (nur für interne Sortierung)
    ws_haupt.column_dimensions['B'].hidden = True
    
    # Monatsauswertung
    ws_monat = wb.create_sheet("Monatsauswertung")
    
    for col, header in enumerate(["Jahr-Monat", "Monatsname", "Anzahl Rechnungen", "Gesamtsumme (€)"], 1):
        cell = ws_monat.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    for idx, (key, summe) in enumerate(sorted(monats_summen.items()), 1):
        if key != "0000-00":
            parts = key.split("-")
            jahr = parts[0]
            monat = int(parts[1])
            monatsname = datetime(int(jahr), monat, 1).strftime("%B %Y")
            anzahl = len(monats_gruppen[key])
            
            ws_monat.cell(row=idx + 1, column=1, value=key)
            ws_monat.cell(row=idx + 1, column=2, value=monatsname)
            ws_monat.cell(row=idx + 1, column=3, value=anzahl)
            ws_monat.cell(row=idx + 1, column=4, value=format_betrag(summe))
    
    ws_monat.column_dimensions['A'].width = 15
    ws_monat.column_dimensions['B'].width = 20
    ws_monat.column_dimensions['C'].width = 20
    ws_monat.column_dimensions['D'].width = 20
    
    # Jahresauswertung
    ws_jahr = wb.create_sheet("Jahresauswertung")
    jahres_daten = defaultdict(lambda: {"anzahl": 0, "summe": 0.0})
    
    for key, summe in monats_summen.items():
        if key != "0000-00":
            jahr = key.split("-")[0]
            jahres_daten[jahr]["anzahl"] += len(monats_gruppen[key])
            jahres_daten[jahr]["summe"] += summe
    
    for col, header in enumerate(["Jahr", "Anzahl Rechnungen", "Gesamtsumme (€)"], 1):
        cell = ws_jahr.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
    
    for idx, (jahr, daten) in enumerate(sorted(jahres_daten.items()), 1):
        ws_jahr.cell(row=idx + 1, column=1, value=jahr)
        ws_jahr.cell(row=idx + 1, column=2, value=daten["anzahl"])
        ws_jahr.cell(row=idx + 1, column=3, value=format_betrag(daten["summe"]))
    
    ws_jahr.column_dimensions['A'].width = 15
    ws_jahr.column_dimensions['B'].width = 20
    ws_jahr.column_dimensions['C'].width = 20
    
    wb.save(ziel_pfad)
    
    return len(neue_eintraege), len(alle_eintraege)


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
