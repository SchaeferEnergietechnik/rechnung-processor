#!/usr/bin/env node

/**
 * Rechnungsbearbeitung - PDF-Rechnungen verarbeiten (Standalone .exe Version)
 * 
 * Funktionen:
 * - Liest PDFs aus dem Ordner, in dem die .exe liegt
 * - Extrahiert: Gesamtpreis (inkl. MwSt), Lieferant, Rechnungsdatum
 * - Benennt um direkt im Ordner: 26-xxx_YYYY-MM-DD_Lieferant_Gesamtpreis.pdf
 * - Schreibt Excel direkt in den Ordner
 * 
 * Usage: Einfach die .exe im Ordner mit den PDFs ausführen
 */

const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');

// Konfiguration
const CONFIG = {
  startNummer: 1,
  prefix: '26-',
  excelDatei: 'Rechnungsuebersicht.xlsx'
};

// Bekannte Lieferanten mit korrekten Umlauten
const LIEFERANTEN_MAP = {
  'ttcher': 'Böttcher',
  'bottcher': 'Böttcher',
  'stadtwerke wittenberge': 'Stadtwerke Wittenberge',
  'stadtwerke': 'Stadtwerke',
  'wittenberge': 'Stadtwerke Wittenberge',
  're-invent': 'RE-INvent',
  'invent': 'RE-INvent',
  'voelkner': 'Voelkner',
  'steinke': 'Steinke'
};

// Stoppwörter
const STOPP_WORTER = ['umsatzsteuer', 'ust-id', 'steuernummer', 'bankverbindung', 'bic', 'iban', 
  'öffnungszeiten', 'montag', 'dienstag', 'mittwoch', 'donnerstag', 'freitag', 
  'strasse', 'ferchlipp', 'lichterfelde', 'altmärkische', 'deutschland', 'g.e.s.', 'energietechnik'];

/**
 * Betrag normalisieren
 */
function parseBetrag(betragStr) {
  if (!betragStr) return 0;
  const clean = betragStr.replace(/\./g, '').replace(',', '.');
  return parseFloat(clean) || 0;
}

/**
 * Formatiert Betrag
 */
function formatBetrag(betrag) {
  return betrag.toFixed(2).replace('.', ',');
}

/**
 * Extrahiert den Lieferanten aus dem Text (mit Umlaut-Fix)
 */
function findeLieferant(text) {
  const textLower = text.toLowerCase();
  
  // Prüfe auf bekannte Lieferanten (inkl. Teilstrings)
  for (const [key, value] of Object.entries(LIEFERANTEN_MAP)) {
    if (textLower.includes(key.toLowerCase())) {
      return value;
    }
  }
  
  // Suche nach Firmen mit GmbH/AG/KG
  const firmaMuster = /([A-ZÄÖÜ][a-zäöüßA-ZÄÖÜ\s]{2,40}(?:GmbH|AG|KG|OHG|e\.?K|UG))/i;
  const match = text.match(firmaMuster);
  if (match) {
    const kandidat = match[1].trim();
    if (!STOPP_WORTER.some(w => textLower.includes(w))) {
      return bereinigeLieferant(kandidat);
    }
  }
  
  // Erste Zeilen durchgehen
  const zeilen = text.split(/\r?\n/).map(z => z.trim()).filter(z => z.length > 0);
  for (const zeile of zeilen.slice(0, 15)) {
    if (/\b(GmbH|AG|KG|OHG|e\.?K|UG)\b/i.test(zeile)) {
      const bereinigt = bereinigeLieferant(zeile);
      if (bereinigt.length > 3) {
        const zeileLower = zeile.toLowerCase();
        if (!STOPP_WORTER.some(w => zeileLower.includes(w))) {
          return bereinigt;
        }
      }
    }
  }
  
  return 'Unbekannt';
}

/**
 * Bereinigt den Lieferantennamen
 */
function bereinigeLieferant(name) {
  let fixed = name
    .replace(/ttcher/i, 'Böttcher')
    .replace(/bottcher/i, 'Böttcher')
    .replace(/schaefer/i, 'Schäfer')
    .replace(/ae/g, 'ä')
    .replace(/oe/g, 'ö')
    .replace(/ue/g, 'ü')
    .replace(/ss/g, 'ß');
  
  return fixed
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .trim()
    .substring(0, 40);
}

/**
 * Findet die Endsumme/Gesamtsumme
 */
function findeGesamtpreis(text) {
  const musterPrioritaet = [
    /(?:rechnungsbetrag|zu zahlen|fällig|noch zu zahlend)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})/i,
    /(?:gesamtbetrag|endbetrag|summe)[^\d]{0,50}(?:brutto)?[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})/i,
    /(?:zahlbar|betrag)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})/i,
    /(?:gesamt|total)[^\d]*(\d{1,3}(?:\.\d{3})*,\d{2})/i
  ];
  
  for (const muster of musterPrioritaet) {
    const match = text.match(muster);
    if (match) {
      const betrag = parseBetrag(match[1]);
      if (betrag > 0 && betrag < 50000) {
        return { betrag, original: match[1] };
      }
    }
  }
  
  // Fallback: Höchster realistischer Betrag
  const betragMuster = /(\d{1,3}(?:\.\d{3})*,\d{2})/g;
  let match;
  const betraege = [];
  
  while ((match = betragMuster.exec(text)) !== null) {
    const betrag = parseBetrag(match[1]);
    if (betrag > 0 && betrag < 10000) {
      betraege.push(betrag);
    }
  }
  
  if (betraege.length > 0) {
    const maxBetrag = Math.max(...betraege);
    return { betrag: maxBetrag, original: maxBetrag.toFixed(2).replace('.', ',') };
  }
  
  return null;
}

/**
 * Extrahiert das Rechnungsdatum
 */
function findeDatum(text) {
  // Zuerst: Suche nach explizitem "Rechnungsdatum"
  const rechnungsDatumMuster = /rechnungsdatum[\s:]+(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})/i;
  const explizitesDatum = text.match(rechnungsDatumMuster);
  if (explizitesDatum) {
    const tag = parseInt(explizitesDatum[1]);
    const monat = parseInt(explizitesDatum[2]);
    const jahr = parseInt(explizitesDatum[3]);
    if (tag >= 1 && tag <= 31 && monat >= 1 && monat <= 12 && jahr >= 2020 && jahr <= 2030) {
      return { tag, monat, jahr, iso: `${jahr}-${String(monat).padStart(2, '0')}-${String(tag).padStart(2, '0')}` };
    }
  }
  
  // Alternative: Suche nach "Datum" oder "Rechnung" + Datum
  const datumMitKontextMuster = /(?:datum|rechnung|invoice)[\s\S]{0,30}(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})/gi;
  let kontextMatch;
  while ((kontextMatch = datumMitKontextMuster.exec(text)) !== null) {
    const tag = parseInt(kontextMatch[1]);
    const monat = parseInt(kontextMatch[2]);
    const jahr = parseInt(kontextMatch[3]);
    if (tag >= 1 && tag <= 31 && monat >= 1 && monat <= 12 && jahr >= 2020 && jahr <= 2030) {
      const contextBefore = text.substring(Math.max(0, kontextMatch.index - 50), kontextMatch.index).toLowerCase();
      if (!/verbrauch|zeitraum|von|bis|abrechnung/i.test(contextBefore)) {
        return { tag, monat, jahr, iso: `${jahr}-${String(monat).padStart(2, '0')}-${String(tag).padStart(2, '0')}` };
      }
    }
  }
  
  // Fallback: Alle Datumsangaben sammeln
  const muster = /(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})/g;
  const daten = [];
  let match;
  
  while ((match = muster.exec(text)) !== null) {
    const tag = parseInt(match[1]);
    const monat = parseInt(match[2]);
    const jahr = parseInt(match[3]);
    
    if (tag >= 1 && tag <= 31 && monat >= 1 && monat <= 12 && jahr >= 2020 && jahr <= 2030) {
      const context = text.substring(Math.max(0, match.index - 80), match.index + 80).toLowerCase();
      const istVerbrauch = /verbrauch|zeitraum|von.*bis|abrechnungszeitraum/i.test(context);
      const istRechnung = /rechnung|rechnungsdatum|fällig|zahlbar/i.test(context);
      
      const gewicht = istRechnung ? 3 : (istVerbrauch ? 0 : 1);
      if (gewicht > 0) {
        daten.push({ tag, monat, jahr, gewicht });
      }
    }
  }
  
  if (daten.length === 0) return null;
  
  const zaehler = {};
  daten.forEach(d => {
    const key = `${d.jahr}-${String(d.monat).padStart(2, '0')}-${String(d.tag).padStart(2, '0')}`;
    zaehler[key] = (zaehler[key] || 0) + d.gewicht;
  });
  
  const haeufigstes = Object.entries(zaehler)
    .sort((a, b) => b[1] - a[1])[0][0];
  
  const [jahr, monat, tag] = haeufigstes.split('-').map(Number);
  return { tag, monat, jahr, iso: haeufigstes };
}

/**
 * Verarbeitet eine einzelne PDF-Datei
 */
async function verarbeitePDF(dateiPfad) {
  try {
    const buffer = fs.readFileSync(dateiPfad);
    const data = await pdfParse(buffer);
    const text = data.text || '';
    
    const hatText = text.trim().length > 100;
    
    if (!hatText) {
      return {
        datei: path.basename(dateiPfad),
        pfad: dateiPfad,
        gescannt: true,
        extrahiert: false,
        hinweis: 'Gescanntes PDF - OCR nicht verfügbar'
      };
    }
    
    // Extrahiere alle Daten
    const lieferant = findeLieferant(text);
    const gesamtpreis = findeGesamtpreis(text);
    const datum = findeDatum(text);
    
    return {
      datei: path.basename(dateiPfad),
      pfad: dateiPfad,
      lieferant,
      gesamtpreis: gesamtpreis ? gesamtpreis.betrag : 0,
      gesamtpreisOriginal: gesamtpreis ? gesamtpreis.original : '-',
      datum,
      seiten: data.numpages,
      extrahiert: true
    };
    
  } catch (error) {
    return {
      datei: path.basename(dateiPfad),
      pfad: dateiPfad,
      fehler: error.message,
      extrahiert: false
    };
  }
}

/**
 * Generiert neuen Dateinamen mit Datum
 */
function generiereDateiname(daten, nummer) {
  const nummerStr = `${CONFIG.prefix}${String(nummer).padStart(3, '0')}`;
  const datumStr = daten.datum ? daten.datum.iso : '0000-00-00';
  const lieferantSafe = daten.lieferant;
  const preisStr = formatBetrag(daten.gesamtpreis);
  
  return `${nummerStr}_${datumStr}_${lieferantSafe}_${preisStr}EUR.pdf`;
}

/**
 * Erstellt die Excel-Datei
 */
function erstelleExcel(ergebnisse, zielPfad) {
  // Hauptdaten
  const hauptDaten = ergebnisse
    .filter(e => e.extrahiert)
    .map((e, idx) => ({
      'Lfd. Nr.': `${CONFIG.prefix}${String(idx + 1).padStart(3, '0')}`,
      'Rechnungsdatum': e.datum ? e.datum.iso : '-',
      'Lieferant': e.lieferant,
      'Gesamtpreis inkl. MwSt (€)': formatBetrag(e.gesamtpreis),
      'Anzahl Seiten': e.seiten,
      'Ursprünglicher Dateiname': e.datei
    }));
  
  // Nicht verarbeitet
  const nichtVerarbeitet = ergebnisse
    .filter(e => !e.extrahiert)
    .map(e => ({
      'Datei': e.datei,
      'Grund': e.gescannt ? e.hinweis : (e.fehler || 'Unbekannt')
    }));
  
  // Monatliche Auswertung
  const monatsDaten = {};
  ergebnisse
    .filter(e => e.extrahiert && e.datum)
    .forEach(e => {
      const key = `${e.datum.jahr}-${String(e.datum.monat).padStart(2, '0')}`;
      if (!monatsDaten[key]) {
        monatsDaten[key] = { jahr: e.datum.jahr, monat: e.datum.monat, anzahl: 0, summe: 0 };
      }
      monatsDaten[key].anzahl++;
      monatsDaten[key].summe += e.gesamtpreis;
    });
  
  const monatsTabelle = Object.values(monatsDaten)
    .sort((a, b) => (a.jahr * 12 + a.monat) - (b.jahr * 12 + b.monat))
    .map(m => ({
      'Jahr': m.jahr,
      'Monat': m.monat,
      'Monatsname': new Date(m.jahr, m.monat - 1).toLocaleString('de-DE', { month: 'long' }),
      'Anzahl Rechnungen': m.anzahl,
      'Gesamtsumme (€)': formatBetrag(m.summe)
    }));
  
  // Jährliche Auswertung
  const jahresDaten = {};
  ergebnisse
    .filter(e => e.extrahiert)
    .forEach(e => {
      const jahr = e.datum ? e.datum.jahr : new Date().getFullYear();
      if (!jahresDaten[jahr]) {
        jahresDaten[jahr] = { jahr, anzahl: 0, summe: 0 };
      }
      jahresDaten[jahr].anzahl++;
      jahresDaten[jahr].summe += e.gesamtpreis;
    });
  
  const jahresTabelle = Object.values(jahresDaten)
    .sort((a, b) => a.jahr - b.jahr)
    .map(j => ({
      'Jahr': j.jahr,
      'Anzahl Rechnungen': j.anzahl,
      'Gesamtsumme (€)': formatBetrag(j.summe)
    }));
  
  // Erstelle Workbook
  const wb = XLSX.utils.book_new();
  
  if (hauptDaten.length > 0) {
    const wsHaupt = XLSX.utils.json_to_sheet(hauptDaten);
    XLSX.utils.book_append_sheet(wb, wsHaupt, 'Rechnungen');
  }
  
  if (nichtVerarbeitet.length > 0) {
    const wsFehler = XLSX.utils.json_to_sheet(nichtVerarbeitet);
    XLSX.utils.book_append_sheet(wb, wsFehler, 'Nicht verarbeitet');
  }
  
  if (monatsTabelle.length > 0) {
    const wsMonat = XLSX.utils.json_to_sheet(monatsTabelle);
    XLSX.utils.book_append_sheet(wb, wsMonat, 'Monatsauswertung');
  }
  
  if (jahresTabelle.length > 0) {
    const wsJahr = XLSX.utils.json_to_sheet(jahresTabelle);
    XLSX.utils.book_append_sheet(wb, wsJahr, 'Jahresauswertung');
  }
  
  XLSX.writeFile(wb, zielPfad);
}

/**
 * Hauptfunktion - Verarbeitet PDFs im aktuellen Ordner
 */
async function main() {
  // Ordner ist automatisch der Ordner, in dem das Script/.exe liegt
  const quellOrdner = process.cwd();
  
  console.log(`📁 Verarbeite PDFs in: ${quellOrdner}\n`);
  
  // Finde alle PDFs (ohne bereits umbenannte)
  const dateien = fs.readdirSync(quellOrdner)
    .filter(f => f.toLowerCase().endsWith('.pdf'))
    .map(f => path.join(quellOrdner, f))
    .filter(f => !path.basename(f).startsWith(CONFIG.prefix)); // Bereits verarbeitete überspringen
  
  if (dateien.length === 0) {
    console.log('⚠️  Keine neuen PDF-Dateien gefunden.');
    console.log('   (Dateien mit "26-" am Anfang werden übersprungen)');
    return;
  }
  
  console.log(`📄 ${dateien.length} neue PDF(s) gefunden. Verarbeite...\n`);
  
  // Verarbeite alle PDFs
  const ergebnisse = [];
  for (let i = 0; i < dateien.length; i++) {
    const pfad = dateien[i];
    const dateiName = path.basename(pfad);
    console.log(`[${i + 1}/${dateien.length}] ${dateiName}...`);
    
    const daten = await verarbeitePDF(pfad);
    ergebnisse.push(daten);
    
    if (daten.extrahiert) {
      const neuerName = generiereDateiname(daten, i + 1);
      daten.neuerName = neuerName;
      
      const zielPfad = path.join(quellOrdner, neuerName);
      
      // Umbenennen (verschieben)
      try {
        fs.renameSync(pfad, zielPfad);
        console.log(`  ✅ ${neuerName}`);
      } catch (err) {
        console.log(`  ⚠️  Konnte nicht umbenennen: ${err.message}`);
      }
      
      console.log(`     Lieferant: ${daten.lieferant}`);
      console.log(`     Datum: ${daten.datum ? daten.datum.iso : '-'}`);
      console.log(`     Gesamtpreis: ${formatBetrag(daten.gesamtpreis)} €`);
    } else if (daten.gescannt) {
      console.log(`  ⚠️  Übersprungen: ${daten.hinweis}`);
    } else {
      console.log(`  ❌ Fehler: ${daten.fehler}`);
    }
    console.log('');
  }
  
  // Excel erstellen (direkt im Ordner)
  const excelPfad = path.join(quellOrdner, CONFIG.excelDatei);
  erstelleExcel(ergebnisse, excelPfad);
  
  // Zusammenfassung
  const erfolgreich = ergebnisse.filter(e => e.extrahiert).length;
  const fehler = ergebnisse.filter(e => !e.extrahiert).length;
  const gesamtSumme = ergebnisse
    .filter(e => e.extrahiert)
    .reduce((sum, e) => sum + e.gesamtpreis, 0);
  
  console.log('\n📊 Zusammenfassung:');
  console.log(`   Erfolgreich verarbeitet: ${erfolgreich}`);
  console.log(`   Fehler: ${fehler}`);
  console.log(`   Gesamtsumme: ${formatBetrag(gesamtSumme)} €`);
  console.log(`\n📂 Excel erstellt: ${CONFIG.excelDatei}`);
  
  // Warte auf Tastendruck bei .exe
  console.log('\n-------------------------------------------');
  console.log('Drücken Sie Enter zum Beenden...');
  process.stdin.once('data', () => process.exit(0));
}

main().catch(err => {
  console.error('❌ Fehler:', err);
  process.exit(1);
});
