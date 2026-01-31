// --- KONFIGURATION: Activity Review ---
const REVIEW_TARGET_SHEET = 'KK_TIMELINE'; // Wo die Daten gelesen werden
const REVIEW_OUTPUT_SHEET = 'AI_ACTIVITY_REVIEWS'; // Wo die Reviews gespeichert werden
const REVIEW_API_MODEL = 'gemini-3-pro-preview'; // Oder dein bevorzugtes Modell
// --- Globale Konstanten (müssen im Projekt definiert sein) ---
// const SOURCE_TIMELINE_SHEET = 'timeline';
// const TIMELINE_SHEET_NAME = 'KK_TIMELINE';
// const LOG_SHEET_NAME = 'AI_LOG';
// ------------------------------------

/**
 * V32-FIX: Generiert einen kurzen Review zur absolvierten Aktivität.
 * FIX: Schreibt exakt in die Spaltenstruktur deines Screenshots:
 * [Datum, Load Ist, Load Soll, ATL Ist, ATL Soll, CTL Ist, CTL Soll, ACWR Ist, ACWR Soll, Text]
 * FIX: Hängt UNTEN an (appendRow).
 * FIX: Extrahiert sauber den Text aus dem JSON.
 */
function generateActivityReview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- KONFIGURATION ---
  const TIMELINE_SHEET_NAME = 'timeline';
  const PLAN_SHEET_NAME = 'AI_REPORT_PLAN';
  const REVIEW_SHEET_NAME = 'AI_ACTIVITY_REVIEWS';
  
  try {
    logToSheet('INFO', '[ActivityReview V32] Starte Review-Generierung (Mapping Fix)...');

    // 1. TIMELINE DATEN (IST) LADEN
    const timelineSheet = ss.getSheetByName(TIMELINE_SHEET_NAME);
    if (!timelineSheet) throw new Error(`Blatt '${TIMELINE_SHEET_NAME}' fehlt.`);
    
    const tData = timelineSheet.getDataRange().getValues();
    const tHeaders = tData[0];
    const idxDate = tHeaders.indexOf('date');
    const idxIsToday = tHeaders.indexOf('is_today');
    const idxLoad = tHeaders.indexOf('load_fb_day');
    const idxSport = tHeaders.indexOf('Sport_x');
    const idxAtl = tHeaders.indexOf('fbATL_obs');
    const idxCtl = tHeaders.indexOf('fbCTL_obs');
    const idxAcwr = tHeaders.indexOf('fbACWR_obs');

    // Finde HEUTE (is_today = 1)
    let istRow = null;
    for (let i = 1; i < tData.length; i++) {
      if (tData[i][idxIsToday] == 1) {
        istRow = tData[i];
        break;
      }
    }

    if (!istRow) {
      logToSheet('WARN', '[ActivityReview] Kein Tag mit is_today=1 gefunden. Abbruch.');
      return;
    }

    const istLoad = parseGermanFloat(istRow[idxLoad]);
    
    // Kleiner Check: Wenn Load 0 und kein Sport eingetragen, evtl. abbrechen? 
    // Wir lassen es laufen, falls du auch Ruhetage bewertet haben willst.

    const todayDateObj = new Date(istRow[idxDate]);
    const todayString = Utilities.formatDate(todayDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

    const todayData = {
      date: todayString,
      sport_x: istRow[idxSport],
      load_fb_day: istLoad,
      fbATL_obs: istRow[idxAtl],
      fbCTL_obs: istRow[idxCtl],
      fbACWR_obs: istRow[idxAcwr]
    };

    // 2. PLAN DATEN (SOLL) LADEN
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    let planData = { load: "N/A", zone: "N/A", beispiel_1: "-", beispiel_2: "-" };
    
    if (planSheet && planSheet.getLastRow() > 1) {
      const pData = planSheet.getDataRange().getValues();
      const pHeaders = pData[0];
      const pIdxDate = pHeaders.indexOf('Datum');
      const pIdxLoad = pHeaders.indexOf('Empfohlener Load (ESS)');
      const pIdxZone = pHeaders.indexOf('Empfohlene Zone (KI)');
      const pIdxB1 = pHeaders.indexOf('Beispiel-Training 1 (KI)');
      const pIdxB2 = pHeaders.indexOf('Beispiel-Training 2 (KI)');

      for (let i = 1; i < pData.length; i++) {
        let pDate = pData[i][pIdxDate];
        if (pDate instanceof Date) {
          pDate = Utilities.formatDate(pDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        if (pDate === todayString) {
          planData = {
            load: pData[i][pIdxLoad],
            zone: (pIdxZone > -1) ? pData[i][pIdxZone] : "N/A",
            beispiel_1: (pIdxB1 > -1) ? pData[i][pIdxB1] : "-",
            beispiel_2: (pIdxB2 > -1) ? pData[i][pIdxB2] : "-"
          };
          break;
        }
      }
    }

    // 3. PROMPT GENERIEREN
    const prompt = buildActivityReviewPrompt(todayData, planData);

    // 4. KI AUFRUFEN & REINIGEN
    logToSheet('INFO', '[ActivityReview] Sende Daten an KI...');
    
    let reviewText = "Fehler bei KI-Generierung.";
    try {
       let rawResponse = callGeminiAPI(prompt);
       
       // Clean Markdown Code Blocks
       let cleanJson = rawResponse.replace(/```json/g, "").replace(/```/g, "").trim();

       // Versuche JSON zu parsen
       if (cleanJson.startsWith('{') || cleanJson.startsWith('[')) {
          try {
             const json = JSON.parse(cleanJson);
             if (json.analysis) reviewText = json.analysis;
             else if (json.text) reviewText = json.text;
             else if (json.response) reviewText = json.response;
             else if (json.review) reviewText = json.review;
             else reviewText = JSON.stringify(json); 
          } catch(e) { 
             reviewText = cleanJson; 
          }
       } else {
          reviewText = cleanJson;
       }
       
    } catch(e) {
       logToSheet('ERROR', `[ActivityReview] KI-API Fehler: ${e.message}`);
       return;
    }

    // 5. SPEICHERN (UNTEN ANFÜGEN - MIT KORREKTER SPALTENZUORDNUNG)
    let revSheet = ss.getSheetByName(REVIEW_SHEET_NAME);
    if (!revSheet) {
      // Fallback Header, falls Blatt neu erstellt wird (Passt zur Struktur im Screenshot)
      revSheet = ss.insertSheet(REVIEW_SHEET_NAME);
      const h = ['Datum', 'Load (IST)', 'Load (SOLL)', 'ATL (IST)', 'ATL (SOLL)', 'CTL (IST)', 'CTL (SOLL)', 'ACWR (IST)', 'ACWR (SOLL)', 'KI Review Text'];
      revSheet.appendRow(h);
      revSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    }

    // MAPPING V32: Exakt passend zu deinem Screenshot
    const newRow = [
      todayDateObj,            // A: Datum
      todayData.load_fb_day,   // B: Load (IST)
      planData.load,           // C: Load (SOLL)
      todayData.fbATL_obs,     // D: ATL (IST)
      "N/A",                   // E: ATL (SOLL) - Daten fehlen im Plan-Sheet meist
      todayData.fbCTL_obs,     // F: CTL (IST)
      "N/A",                   // G: CTL (SOLL)
      todayData.fbACWR_obs,    // H: ACWR (IST)
      "N/A",                   // I: ACWR (SOLL)
      reviewText               // J: KI Review Text
    ];

    revSheet.appendRow(newRow);
    
    // Formatierung
    const lastRow = revSheet.getLastRow();
    revSheet.getRange(lastRow, 1).setNumberFormat('dd.MM.yyyy');
    revSheet.getRange(lastRow, 10).setWrap(true); // Textumbruch für Review

    logToSheet('INFO', '[ActivityReview] Review erfolgreich gespeichert (Spalten korrigiert).');

  } catch (e) {
    logToSheet('ERROR', `[ActivityReview V32] CRASH: ${e.message}`);
  }
}

/**
 * Baut den spezifischen Prompt für den Aktivitäts-Review.
 */
function buildReviewPrompt(datum, ist, soll) {
  // Sicherstellen, dass Werte Zahlen sind
  const istLoad = typeof ist.load === 'number' ? ist.load : 0;
  const sollLoad = typeof soll.load === 'number' ? soll.load : 0;
  const istAtl = typeof ist.atl === 'number' ? ist.atl : 0;
  const sollAtl = typeof soll.atl === 'number' ? soll.atl : 0;
  const istCtl = typeof ist.ctl === 'number' ? ist.ctl : 0;
  const sollCtl = typeof soll.ctl === 'number' ? soll.ctl : 0;
  const istAcwr = (typeof ist.acwr === 'number' && !isNaN(ist.acwr)) ? ist.acwr : 0;
  const sollAcwr = typeof soll.acwr === 'number' ? soll.acwr : 0;

  const diffLoad = istLoad - sollLoad;
  const diffAtl = istAtl - sollAtl;
  const diffCtl = istCtl - sollCtl;

  let datumString = "Unbekanntes Datum";
  if (datum instanceof Date && !isNaN(datum.valueOf())) {
      datumString = datum.toLocaleDateString('de-DE');
  } else {
     if (typeof logToSheet === 'function') logToSheet('WARN', '[Review Prompt] Ungültiges Datum empfangen: ' + datum);
  }

  let prompt = `Du bist "Coach Kira", eine KI-Fitnessexpertin. Analysiere die Abweichung zwischen geplanter und tatsächlicher Trainingsbelastung für den ${datumString}.

GEPLANTE WERTE (SOLL) für Ende des Tages:
- Geplanter Load (ESS): ${sollLoad.toFixed(0)}
- Prognostizierter ATL: ${sollAtl.toFixed(0)}
- Prognostizierter CTL: ${sollCtl.toFixed(0)}
- Prognostizierter ACWR: ${sollAcwr.toFixed(2)}

TATSÄCHLICHE WERTE (IST) nach der Aktivität (Ende des Tages):
- Tatsächlicher Load (EPOC): ${istLoad.toFixed(0)} (Differenz: ${diffLoad >= 0 ? '+' : ''}${diffLoad.toFixed(0)})
- Beobachteter ATL: ${istAtl.toFixed(0)} (Differenz: ${diffAtl >= 0 ? '+' : ''}${diffAtl.toFixed(0)})
- Beobachteter CTL: ${istCtl.toFixed(0)} (Differenz: ${diffCtl >= 0 ? '+' : ''}${diffCtl.toFixed(0)})
- Beobachteter ACWR: ${istAcwr.toFixed(2)}

DEINE AUFGABE:
Bewerte kurz und prägnant (max. 3-4 Sätze) die Abweichung.
- War der tatsächliche Load höher oder niedriger als geplant? Ist die Abweichung signifikant (+/- 10%)?
- Wie haben sich ATL und CTL im Vergleich zur Prognose entwickelt? Entspricht das der Erwartung basierend auf der Load-Abweichung?
- Gibt es Besonderheiten oder Empfehlungen für morgen basierend auf diesem Ergebnis (z.B. Plan beibehalten, vorsichtiger sein, mehr Erholung)?

ANTWORTE AUSSCHLIESSLICH IM FOLGENDEN JSON-FORMAT (ohne Markdown):
{
  "review_text": "DEINE BEWERTUNG HIER..."
}
`;
  return prompt;
}

/**
 * Ruft die Gemini API speziell für den Review auf.
 */
function callGeminiForReview(url, prompt) {
  const payload = {
    "contents": [ { "parts": [ { "text": prompt } ] } ]
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts && jsonResponse.candidates[0].content.parts[0]) {
            let textResponse = jsonResponse.candidates[0].content.parts[0].text;
            textResponse = textResponse.replace(/^```json\n?/, '').replace(/\n?```$/, '');
             try {
                return JSON.parse(textResponse);
             } catch (parseError) {
                 logToSheet('ERROR', `[Review API Call] Fehler beim Parsen der bereinigten JSON-Antwort: ${parseError.message}. Bereinigte Antwort: ${textResponse}`);
                 return null;
             }
        } else {
             logToSheet('ERROR', `[Review API Call] Unerwartete Antwortstruktur von Gemini: ${responseBody}`);
             return null;
        }
      } else {
        logToSheet('ERROR', `[Review API Call] Fehler bei API-Anfrage (Code ${responseCode}): ${responseBody}`);
        return null;
      }
  } catch (e) {
      logToSheet('ERROR', `[Review API Call] Fehler beim API-Aufruf oder Parsen: ${e.message}.`);
      return null;
  }
}

/**
 * Schreibt das Ergebnis des Reviews IMMER in Zeile 2 des Output-Blatts.
 * (Version 1.7 - Überschreibt Zeile 2)
 */
function writeReviewToSheet(datum, istLoad, istAtl, istCtl, istAcwr, sollLoad, sollAtl, sollCtl, sollAcwr, reviewText) {
  // Check Parameter (rudimentär)
  if (typeof istLoad === 'undefined' || typeof sollLoad === 'undefined') {
      const errorMsg = `[Write Review] FEHLER: Ungültige Load-Parameter empfangen. IST=${istLoad}, SOLL=${sollLoad}`;
      if (typeof logToSheet === 'function') logToSheet('ERROR', errorMsg); else Logger.log(errorMsg);
      try { SpreadsheetApp.getUi().alert("Interner Fehler beim Schreiben des Reviews (ungültige Load-Daten). Siehe Logs."); } catch(e){}
      return;
  }

  // Hole Spreadsheet-Objekt direkt hier
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
      const errorMsg = "[Write Review] FEHLER: Konnte aktives Spreadsheet nicht abrufen.";
      if (typeof logToSheet === 'function') logToSheet('ERROR', errorMsg); else Logger.log(errorMsg);
      try { SpreadsheetApp.getUi().alert("Interner Fehler beim Schreiben des Reviews (Spreadsheet nicht gefunden). Siehe Logs."); } catch(e){}
      return;
  }

  let outputSheet = ss.getSheetByName(REVIEW_OUTPUT_SHEET);
  const headers = [
      'Datum',
      'Load (IST)', 'Load (SOLL)',
      'ATL (IST)', 'ATL (SOLL)',
      'CTL (IST)', 'CTL (SOLL)',
      'ACWR (IST)', 'ACWR (SOLL)',
      'KI Review Text'
  ];

  // Sicherstellen, dass das Blatt und die Header existieren
  if (!outputSheet) {
    outputSheet = ss.insertSheet(REVIEW_OUTPUT_SHEET);
    outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    outputSheet.setFrozenRows(1);
    if (typeof logToSheet === 'function') logToSheet('INFO', `[Write Review] Blatt '${REVIEW_OUTPUT_SHEET}' wurde neu erstellt.`);
    // Stelle sicher, dass Zeile 2 existiert, falls das Blatt neu ist
    if (outputSheet.getMaxRows() < 2) {
        outputSheet.insertRowAfter(1);
    }
  } else {
    // Header prüfen und ggf. überschreiben
    const existingHeaders = outputSheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (JSON.stringify(existingHeaders) !== JSON.stringify(headers)) {
        if (typeof logToSheet === 'function') logToSheet('WARN', `[Write Review] Header in '${REVIEW_OUTPUT_SHEET}' stimmen nicht überein. Überschreibe sie.`);
        outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        outputSheet.setFrozenRows(1);
    }
    // Stelle sicher, dass Zeile 2 existiert, auch wenn das Blatt schon da war
    if (outputSheet.getMaxRows() < 2) {
        outputSheet.insertRowAfter(1);
        if (typeof logToSheet === 'function') logToSheet('INFO', `[Write Review] Zeile 2 in '${REVIEW_OUTPUT_SHEET}' wurde hinzugefügt.`);
    }
  }

  // Datum prüfen
  let displayDatum = datum;
  if (!(datum instanceof Date) || isNaN(datum.valueOf())) {
      if (typeof logToSheet === 'function') logToSheet('WARN', '[Write Review] Ungültiges Datum beim Schreiben empfangen: ' + datum + ". Verwende Fallback.");
      try {
          displayDatum = new Date(datum);
          if (isNaN(displayDatum.valueOf())) displayDatum = new Date();
      } catch(e) {
          displayDatum = new Date();
      }
  }

  // Daten vorbereiten
  const safeIstLoad = typeof istLoad === 'number' ? istLoad : 0;
  const safeSollLoad = typeof sollLoad === 'number' ? sollLoad : 0;
  const safeIstAtl = typeof istAtl === 'number' ? istAtl : 0;
  const safeSollAtl = typeof sollAtl === 'number' ? sollAtl : 0;
  const safeIstCtl = typeof istCtl === 'number' ? istCtl : 0;
  const safeSollCtl = typeof sollCtl === 'number' ? sollCtl : 0;
  const safeIstAcwr = (typeof istAcwr === 'number' && !isNaN(istAcwr)) ? istAcwr : 0;
  const safeSollAcwr = typeof sollAcwr === 'number' ? sollAcwr : 0;

  const newData = [
    displayDatum,
    safeIstLoad, safeSollLoad,
    safeIstAtl, safeSollAtl,
    safeIstCtl, safeSollCtl,
    safeIstAcwr, safeSollAcwr,
    reviewText
  ];

  // --- MODIFIZIERT (V1.7): Schreibe IMMER in Zeile 2 ---
  const targetRange = outputSheet.getRange(2, 1, 1, newData.length);
  targetRange.setValues([newData]); // Überschreibt den Inhalt von Zeile 2
  if (typeof logToSheet === 'function') logToSheet('INFO', `[Write Review] Zeile 2 in '${REVIEW_OUTPUT_SHEET}' wurde überschrieben.`);
  // --- ENDE MODIFIZIERUNG ---


  // Formatierungen für Zeile 2
  outputSheet.getRange(2, 1, 1, 1).setNumberFormat('yyyy-mm-dd');
  outputSheet.getRange(2, 2, 1, 7).setNumberFormat('0');
  outputSheet.getRange(2, 8, 1, 2).setNumberFormat('0.00');
  outputSheet.getRange(2, 10, 1, 1).setWrap(true);

  // Spaltenbreiten anpassen (optional)
  // outputSheet.autoResizeColumns(1, headers.length);
}

/* // --- ANLEITUNG ZUR ANPASSUNG VON onOpen() ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Coach Kira')
    .addItem('KI-Planung & Charts starten (vX.X)', 'runKiraGeminiSupervisor')
    .addSeparator()
    .addItem('Aktivitäts-Review starten', 'generateActivityReview')
    .addToUi();
}
*/

// --- BENÖTIGTE GLOBALE HILFSFUNKTIONEN (müssen im Projekt existieren) ---
/*
function arrayToObject(headers, data) { ... }
function parseGermanFloat(value) { ... }
function logToSheet(level, message) { ... }
function copyTimelineData() { ... }
*/