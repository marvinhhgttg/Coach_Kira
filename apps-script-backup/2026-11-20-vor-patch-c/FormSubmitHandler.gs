// --- KONFIGURATION: Formular-Handler ---
const TARGET_SHEET_NAME = 'timeline'; // Das Blatt, in das geschrieben wird
const FORM_RESPONSE_SHEET_NAME = 'Formularantworten 5'; // Das Blatt, das die Formularantworten empfängt
const IS_TODAY_COLUMN_NAME = 'is_today'; // Name der Spalte, die den heutigen Tag markiert (Wert 1)
// ------------------------------------

/**
 * (V34 - Fügt Aerobic_TE/Anaerobic_TE hinzu): Wird durch Formular-Trigger ausgelöst.
 * @param {Object} e - Das Event-Objekt, das die Formularantwort enthält.
 */
function onFormSubmit(e) {
  // Annahme: Alle Hilfsfunktionen (logToSheet, parseGermanFloat, etc.) sind global definiert.
  
  if (typeof logToSheet !== 'function') { Logger.log("FEHLER: logToSheet fehlt!"); return; }
  
  logToSheet('INFO', 'Formularantwort empfangen. Starte Verarbeitung (V34 - mit TE-Werten)...');

  // --- GRUNDLAGEN ---
  if (!e || !e.namedValues) {
    logToSheet('ERROR', 'Fehler: Event-Objekt (e) oder e.namedValues nicht gefunden.');
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME); // 'timeline'
  if (!targetSheet) {
    logToSheet('ERROR', `Fehler: Zielblatt '${TARGET_SHEET_NAME}' nicht gefunden.`);
    return;
  }
  if (typeof parseGermanFloat !== 'function') { 
    logToSheet('ERROR', "FEHLER: Benötigte Hilfsfunktion (parseGermanFloat) fehlt!"); 
    return; 
  }

  // --- Schritt 1: Übermittlungstyp & Zieldatum bestimmen ---
  const submissionTypeKey = "Was willst Du eingeben?";
  const dateKey = "Datum"; 
  let submissionType = "";
  let targetDateStr = ""; 
  let targetRowIndex = -1; 
  if (e.namedValues.hasOwnProperty(submissionTypeKey)) {
      submissionType = e.namedValues[submissionTypeKey][0];
  } else {
      logToSheet('WARN', `[Form Submit] Konnte den Übermittlungstyp (Spalte "${submissionTypeKey}") nicht finden.`);
  }
  if (e.namedValues.hasOwnProperty(dateKey) && e.namedValues[dateKey][0]) {
      try {
          const dateRaw = e.namedValues[dateKey][0]; 
          const parts = dateRaw.match(/^(\d{2})\.(\d{2})\.(\d{4})/); 
          if (!parts) throw new Error("Datumsformat nicht erkannt (erwartet DD.MM.YYYY...).");
          const dateObj = new Date(Date.UTC(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1])));
          if (isNaN(dateObj.getTime())) throw new Error("Ungültiges Datum nach dem Parsen.");
          targetDateStr = Utilities.formatDate(dateObj, "UTC", "yyyy-MM-dd");
      } catch (dateError) {
          logToSheet('ERROR', `[Form Submit] Fehler beim Verarbeiten des "Datum"-Felds: ${dateError}. Rohwert war: ${e.namedValues[dateKey][0]}`);
          return; 
      }
  } else {
      logToSheet('ERROR', `[Form Submit] Kein Datum im Formular gefunden (Spalte "${dateKey}") oder Wert leer.`);
      return; 
  }
  logToSheet('INFO', `[Form Submit] Typ: "${submissionType}", Zieldatum: ${targetDateStr}`);


  // --- Schritt 2: Finde die Zielzeile ---
  if (submissionType === "Morgens") {
    logToSheet('INFO', `[Form Submit] Typ "Morgens" erkannt. Starte Tageswechsel...`);
    if (typeof advanceTodayFlag_silent !== 'function') {
        logToSheet('ERROR', "[Form Submit] FEHLER: Funktion 'advanceTodayFlag_silent' nicht gefunden!");
        return;
    }
    const success = advanceTodayFlag_silent(); 
    if (success) {
      logToSheet('INFO', `[Form Submit] Tageswechsel war erfolgreich.`);
      targetRowIndex = findRowByIsToday(targetSheet); 
      if (targetRowIndex === -1) {
         logToSheet('ERROR', `[Form Submit] TAGESWECHSEL OK, aber konnte die NEUE 'is_today=1' Zeile danach nicht finden. Breche ab.`);
         return;
      }
      logToSheet('INFO', `[Form Submit] Neue Zielzeile ist die 'is_today=1' Zeile: ${targetRowIndex}`);
    } else {
      logToSheet('ERROR', `[Form Submit] TAGESWECHSEL FEHLGESCHLAGEN. Breche Eintragung ab.`);
      return; 
    }
  } else {
    logToSheet('INFO', `[Form Submit] Typ "${submissionType}". Suche Zeile für Datum ${targetDateStr}...`);
    const targetDisplayData = targetSheet.getDataRange().getDisplayValues(); 
    const targetHeaders = targetDisplayData[0];
    const dateColIndex = targetHeaders.indexOf('date'); 
    if (dateColIndex === -1) {
      logToSheet('ERROR',`Fehler: Spalte 'date' im Blatt '${TARGET_SHEET_NAME}' nicht gefunden.`);
      return;
    }
    let found = false;
    for (let i = 1; i < targetDisplayData.length; i++) { 
        if (targetDisplayData[i][dateColIndex] === targetDateStr) { 
            targetRowIndex = i + 1;
            found = true;
            break;
        }
    }
    if (!found) {
       const targetDateStrDisplay = `${targetDateStr.split('-')[2]}.${targetDateStr.split('-')[1]}.${targetDateStr.split('-')[0]}`;
       logToSheet('WARN', `[Form Submit] Datum ${targetDateStr} nicht gefunden. Versuche DD.MM.YYYY (${targetDateStrDisplay})...`);
       for (let i = 1; i < targetDisplayData.length; i++) { 
          if (targetDisplayData[i][dateColIndex] === targetDateStrDisplay) {
              targetRowIndex = i + 1;
              found = true;
              break;
          }
       }
    }
    if (targetRowIndex === -1) {
       logToSheet('ERROR', `[Form Submit] Konnte keine Zeile für das Zieldatum ${targetDateStr} (oder DD.MM.YYYY) in '${TARGET_SHEET_NAME}' finden. Daten werden NICHT geschrieben.`);
       return; 
    }
    logToSheet('INFO', `[Form Submit] Zieldatenzeile in '${TARGET_SHEET_NAME}' gefunden: ${targetRowIndex}`);
  }

  // --- Schritt 3: Definiere das Mapping & Spaltenindizes (ANGEPASST FÜR V34) ---
  let relevantColumns = [];
  if (submissionType === "Morgens") {
    relevantColumns = ['garminATL', 'garminCTL', 'garminACWR', 'sleep_hours', 'sleep_score_0_100', 'rhr_bpm', 'hrv_status', 'hrv_threshholds', 'Garmin_Training_Readiness', 'Trainingszustand'];
  } else if (submissionType === "Nach Aktivität") {
    // --- NEU (V34): 'Aerobic_TE', 'Anaerobic_TE' hinzugefügt ---
    relevantColumns = ['load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs', 'garminEnduranceScore', 'Sport_x', 'Zone', 'fb_TR_obs', 'Aerobic_TE', 'Anaerobic_TE'];
  } else if (submissionType === "Abends") {
    relevantColumns = ['kcal_in', 'kcal_out', 'deficit', 'carb_g', 'protein_g', 'fat_g', 'fiber_g'];
  } else {
     // Fallback (z.B. wenn Typ leer ist), fügt die neuen Spalten auch hier hinzu
     relevantColumns = [
        'garminATL', 'garminCTL', 'garminACWR', 'sleep_hours', 'sleep_score_0_100',
        'rhr_bpm', 'hrv_status', 'hrv_threshholds', 'Garmin_Training_Readiness',
        'load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs', 'garminEnduranceScore',
        'kcal_in', 'kcal_out', 'deficit', 'carb_g', 'protein_g', 'fat_g', 'fiber_g',
        'Sport_x', 'Zone', 'fb_TR_obs', 'Aerobic_TE', 'Anaerobic_TE' // <-- NEU (V34)
    ];
    logToSheet('WARN', `[Form Submit] Unbekannter oder fehlender Übermittlungstyp. Versuche alle relevanten Spalten zu aktualisieren.`);
  }
  
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0]; 
  const targetColIndices = {};
  relevantColumns.forEach(colName => {
    const index = targetHeaders.indexOf(colName);
    if (index !== -1) {
      targetColIndices[colName] = index + 1;
    } else {
      if (e.namedValues.hasOwnProperty(colName)) {
        logToSheet('WARN', `[Form Submit] Relevante Spalte '${colName}' in '${TARGET_SHEET_NAME}' nicht gefunden.`);
      }
    }
  });


  // --- Schritt 4: Schreibe die relevanten Daten in die gefundene Datumszeile (ANGEPASST FÜR V34) ---
  logToSheet('INFO', '[Form Submit] Schreibe Daten aus Formularantwort in Zieldatenzeile...');
  let updateCount = 0;
  for (const colName in targetColIndices) {
    if (e.namedValues.hasOwnProperty(colName) && e.namedValues[colName][0] !== "") {
      const sourceValue = e.namedValues[colName][0];
      const targetCol = targetColIndices[colName];
      
      let parsedValue;
      // --- (V33): Text-Spalten ---
      if (colName === 'Sport_x' || colName === 'Zone' || colName === 'hrv_threshholds') {
        parsedValue = sourceValue;
      } else {
        // Alle anderen (auch Aerobic_TE) werden als Zahl geparst
        parsedValue = parseGermanFloat(sourceValue);
      }
      
      try {
           targetSheet.getRange(targetRowIndex, targetCol).setValue(parsedValue);
           updateCount++;
      } catch (error) {
        logToSheet('ERROR', `[Form Submit] Fehler beim Schreiben in Spalte '${colName}' (Wert: ${parsedValue}): ${error}`);
      }
    }
  }
  logToSheet('INFO', `[Form Submit] ${updateCount} Spalten in Zeile ${targetRowIndex} (Datum: ${targetDateStr}) von '${TARGET_SHEET_NAME}' wurden aktualisiert.`);

  // --- Schritt 5 & 6: Trigger (Unverändert) ---
  const heuteStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  if (targetRowIndex !== -1 && targetDateStr === heuteStr) {
    
    // --- Schritt 5: "Nach Aktivität" Spezial-Aktionen (Async Review) ---
    if (submissionType === "Nach Aktivität") {
      logToSheet('INFO', "[Form Submit] Typ 'Nach Aktivität' (HEUTE) erkannt. Starte Spezial-Aktionen (V34)...");
      const actDoneColIndex = targetHeaders.indexOf('activity_done');
            if (actDoneColIndex !== -1) {
                try {
                    targetSheet.getRange(targetRowIndex, actDoneColIndex + 1).setValue('x');
                    logToSheet('INFO', `[Form Submit] Spalte 'activity_done' automatisch auf 'x' gesetzt.`);
                } catch (e) {
                    logToSheet('ERROR', `[Form Submit] Fehler beim Setzen von 'activity_done': ${e.message}`);
                }
            } else {
                logToSheet('WARN', `[Form Submit] Spalte 'activity_done' nicht im Sheet gefunden.`);
            }

      // AKTION 1 (Wahrheit kopieren, SYNCHRON)
      if (typeof copyObservedToForecast === 'function') {
        logToSheet('INFO', `[Form Submit] Aktion (1/3): Kopiere "Wahrheit" (IST -> SOLL)...`);
        copyObservedToForecast(targetRowIndex); // Synchron
      } else {
        logToSheet('ERROR', "[Form Submit] FEHLER: Funktion 'copyObservedToForecast' nicht gefunden!");
      }

      // AKTION 2 (Activity Review - ASYNCHRON)
      if (typeof triggerActivityReview === 'function') {
          logToSheet('INFO', `[Form Submit] Aktion (2/3): Starte Activity Review (Asynchron)...`);
          triggerActivityReview(); // ASYNCHRON
      } else {
          logToSheet('ERROR', "[Form Submit] FEHLER: Funktion 'triggerActivityReview' nicht gefunden!");
      }
    } // Ende "Nach Aktivität"

    // --- Schritt 6: Globaler Trigger für HEUTE (LÄUFT ZUM SCHLUSS) ---
    logToSheet('INFO', `[Form Submit] Aktion (3/3): Asynchroner KI-Report (Hauptlauf) wird in 10s gestartet...`);
    if (typeof triggerFullAnalysis_WebApp === 'function') {
        triggerFullAnalysis_WebApp(); 
    } else {
        logToSheet('ERROR', "[Form Submit] FEHLER: Funktion 'triggerFullAnalysis_WebApp' nicht gefunden!");
    }

  } else if (targetRowIndex !== -1) {
       logToSheet('INFO', `[Form Submit] Eintragung für VERGANGENHEIT (${targetDateStr}) erkannt. KI-Report wird NICHT neu ausgelöst.`);
  }

  logToSheet('INFO', '[Form Submit] Formularverarbeitung abgeschlossen.');
}

/**
 * (V13.1 HELPER) Findet die Zeile mit 'is_today=1' und gibt den 1-basierten Zeilenindex zurück.
 * @param {Sheet} sheet - Das zu durchsuchende 'timeline'-Blatt.
 * @returns {number} Der 1-basierte Zeilenindex (z.B. 70) oder -1, wenn nicht gefunden.
 */
function findRowByIsToday(sheet) {
  try {
    const data = sheet.getDataRange().getValues(); // Hole aktuelle Werte
    const headers = data[0];
    const isTodayColIndex = headers.indexOf(IS_TODAY_COLUMN_NAME); // 'is_today'

    if (isTodayColIndex === -1) {
      logToSheet('ERROR', `[findRowByIsToday] Spalte '${IS_TODAY_COLUMN_NAME}' nicht gefunden.`);
      return -1;
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][isTodayColIndex] == 1) {
        return i + 1; // 1-basierte Zeilennummer
      }
    }

    logToSheet('WARN', `[findRowByIsToday] Konnte keine Zeile mit 'is_today=1' finden.`);
    return -1;
    
  } catch (e) {
    logToSheet('ERROR', `[findRowByIsToday] Fehler bei der Suche: ${e.message}`);
    return -1;
  }
}

// --- BENÖTIGTE GLOBALE KONSTANTEN (müssen im Projekt existieren) ---
/*
const TARGET_SHEET_NAME = 'timeline';
const FORM_RESPONSE_SHEET_NAME = 'Formularantworten 5';
// IS_TODAY_COLUMN_NAME wird nicht mehr für die Zeilenfindung gebraucht, aber ggf. von anderen Funktionen
const IS_TODAY_COLUMN_NAME = 'is_today';
*/

// --- BENÖTIGTE GLOBALE HILFSFUNKTIONEN (müssen im Projekt existieren) ---
/*
function parseGermanFloat(value) { ... }
function logToSheet(level, message) { ... }
function generateActivityReview() { ... }
function copyTimelineData() { ... } // Wird von generateActivityReview benötigt
*/