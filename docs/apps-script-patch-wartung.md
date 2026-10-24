# Apps-Script-Patch: Wartung (Heartbeat, AI_LOG, AI_REPORT_PLAN, Doppelstarts)

Stand: 19.10.2026 · alle Änderungen in `KiraGeminiSupervisor`

**Was behoben wird**

1. **Heartbeat zeigt immer „Spreadsheet“:** Der Supervisor ruft `hb_(ss, runId, stufe, text)` auf, `hb_` erwartet aber `(stufe, text)`. Deshalb landete das Spreadsheet-Objekt als Stufe im Heartbeat. Das neue `hb_` versteht beide Aufrufformen. Ab dann steht dort z. B. `AI_WAIT_LOCAL`, `WRITE` oder `END_OK`.
2. **AI_LOG seit 28.09. leer:** Die Ursache sehe ich in meiner Kopie nicht, `logToSheet` sieht dort korrekt aus. Vermutlich überschreibt eine zweite Skriptdatei die Funktion, oder es gibt einen stillen Fehler. Deshalb:
   - **Neues `logToSheet`:** schafft bei vollem Blatt Platz, kürzt sehr lange Meldungen, räumt in Blöcken auf und merkt sich den letzten Fehler (`CK_LOG_LAST_ERROR`).
   - **Neuer Diagnose-Modus `?mode=diag`:** zeigt, welche `logToSheet`-Version wirklich aktiv ist, ob ein Testeintrag ankommt, und listet alle Trigger.
3. **AI_REPORT_PLAN steht auf Januar:** `writeOutputToSheets` schreibt das Blatt nicht mehr, der Abschnitt dafür ist irgendwann weggefallen. Gelesen wird es aber noch: von der Kalender-Synchronisation, von „KI-Plan anwenden“, vom Planungsmodul und von `planJson`. Neu schreibt der Speicherteil des Supervisors (`ckSup_finish_`) den empfohlenen Plan wieder in die bekannten Spalten. Das gilt für GPT-4o und für `coach-kira`.
4. **Doppelte Supervisor-Läufe:** Ein Formulareintrag und ein Start aus dem Cockpit lösen jeweils einen Lauf aus, heute im Abstand von 49 Sekunden. Neu gilt:
   - Läuft schon ein Supervisor (jünger als 6 Minuten), beendet sich der zweite Start sofort und merkt einen **Nachlauf** vor.
   - Ist der erste fertig, startet der Nachlauf einmal neu, mit den dann frischesten Daten.
   - So gibt es keine parallelen GPT-4o-Aufrufe mehr, und es fehlen keine neuen Morgenwerte.

---

## Änderung 1 – `hb_` (ganze Funktion)

**Streiche**

```javascript
function hb_(stage, extra) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("AI_HEARTBEAT") || ss.insertSheet("AI_HEARTBEAT");

    // Header einmalig setzen
    if (sh.getLastRow() === 0) {
      sh.getRange("A1:D1").setValues([["timestamp", "stage", "runId", "extra"]]);
      sh.setFrozenRows(1);
    }

    const runId = PropertiesService.getScriptProperties().getProperty("HB_RUN_ID") ||
      ("RUN_" + Utilities.getUuid().slice(0, 8));

    PropertiesService.getScriptProperties().setProperty("HB_RUN_ID", runId);

    // In feste Zellen schreiben (übersichtlich)
    sh.getRange("A2").setValue(new Date());
    sh.getRange("B2").setValue(stage || "");
    sh.getRange("C2").setValue(runId);
    sh.getRange("D2").setValue(extra || "");

    SpreadsheetApp.flush(); // damit du es sofort siehst
  } catch (e) {
    // Herzschlag darf nie den Job killen
    Logger.log("HB ERROR: " + (e && e.stack ? e.stack : e));
  }
}
```

**Ersetze durch**

```javascript
function hb_(a, b, c, d) {
  try {
    // Zwei Aufrufformen: hb_(stage, extra) und hb_(ss, runId, stage, extra) (Supervisor)
    let stage, extra, runIdArg = null;
    if (a && typeof a === 'object' && typeof a.getSheetByName === 'function') {
      runIdArg = b; stage = c; extra = d;
    } else {
      stage = a; extra = b;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("AI_HEARTBEAT") || ss.insertSheet("AI_HEARTBEAT");

    // Header einmalig setzen
    if (sh.getLastRow() === 0) {
      sh.getRange("A1:D1").setValues([["timestamp", "stage", "runId", "extra"]]);
      sh.setFrozenRows(1);
    }

    let runId = runIdArg ? String(runIdArg) : PropertiesService.getScriptProperties().getProperty("HB_RUN_ID");
    if (!runId) {
      runId = "RUN_" + Utilities.getUuid().slice(0, 8);
      PropertiesService.getScriptProperties().setProperty("HB_RUN_ID", runId);
    }

    // In feste Zellen schreiben (übersichtlich)
    sh.getRange("A2:D2").setValues([[new Date(), String(stage || ""), runId, String(extra || "").slice(0, 500)]]);
    SpreadsheetApp.flush(); // damit du es sofort siehst
  } catch (e) {
    // Herzschlag darf nie den Job killen
    Logger.log("HB ERROR: " + (e && e.stack ? e.stack : e));
  }
}
```

---

## Änderung 2 – `logToSheet` (ganze Funktion)

**Streiche**

```javascript
function logToSheet(level, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME, ss.getSheets().length);
      logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Level', 'Message']]).setFontWeight('bold');
      logSheet.setColumnWidth(1, 150); logSheet.setColumnWidth(2, 80); logSheet.setColumnWidth(3, 800);
      logSheet.setFrozenRows(1); // Friere die Kopfzeile
    }

    const timestamp = new Date();
    let logMessage = (typeof message === 'object') ? JSON.stringify(message) : message;

    // 1. Neue Zeile OBEN einfügen (Zeile 2) - Das Neueste steht oben
    logSheet.appendRow([timestamp, level, logMessage]);


    // --- CLEANER: nur die letzten 500 Logs behalten (appendRow-kompatibel) ---
const MAX_ROWS = 500;
const lastRow = logSheet.getLastRow(); // inkl. Header

// Header (1) + 500 Logs = 501 Zeilen
if (lastRow > MAX_ROWS + 1) {
  const numRowsToDelete = lastRow - (MAX_ROWS + 1);
  // Lösche die ÄLTESTEN Logs direkt unter dem Header (Zeile 2 ...)
  logSheet.deleteRows(2, numRowsToDelete);
}


  } catch (e) {
    Logger.log(`Konnte nicht ins Log-Sheet schreiben: ${e.message}`);
  }
```

**Ersetze durch**

```javascript
function logToSheet(level, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME, ss.getSheets().length);
      logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Level', 'Message']]).setFontWeight('bold');
      logSheet.setColumnWidth(1, 150); logSheet.setColumnWidth(2, 80); logSheet.setColumnWidth(3, 800);
      logSheet.setFrozenRows(1); // Friere die Kopfzeile
    }

    let logMessage = (typeof message === 'object') ? JSON.stringify(message) : String(message);
    if (logMessage.length > 45000) logMessage = logMessage.slice(0, 45000) + ' …(gekürzt)';

    // Neue Zeile unten anhängen; Platz schaffen, falls das Blatt voll ist
    const next = logSheet.getLastRow() + 1;
    if (next > logSheet.getMaxRows()) logSheet.insertRowsAfter(logSheet.getMaxRows(), 100);
    logSheet.getRange(next, 1, 1, 3).setValues([[new Date(), String(level || ''), logMessage]]);

    // CLEANER: ab 600 Einträgen auf die letzten 500 kürzen (in Blöcken, spart Laufzeit)
    const MAX_ROWS = 500;
    const lastRow = logSheet.getLastRow(); // inkl. Header
    if (lastRow > MAX_ROWS + 101) logSheet.deleteRows(2, lastRow - (MAX_ROWS + 1));

  } catch (e) {
    Logger.log(`Konnte nicht ins Log-Sheet schreiben: ${e.message}`);
    try {
      PropertiesService.getScriptProperties().setProperty('CK_LOG_LAST_ERROR',
        new Date().toISOString() + ' ' + String(e && e.message ? e.message : e).slice(0, 400));
    } catch (x) {}
  }
}
```

---

## Änderung 3 – `runKiraGeminiSupervisor`: Start-Sperre (Anfang der Funktion)

**Streiche**

```javascript
function runKiraGeminiSupervisor() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HB_RUN_ID = Utilities.getUuid().slice(0, 8);
  hb_(ss, HB_RUN_ID, "START", "Supervisor gestartet");
```

**Ersetze durch**

```javascript
function runKiraGeminiSupervisor() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HB_RUN_ID = Utilities.getUuid().slice(0, 8);
  if (!ckSup_begin_(HB_RUN_ID)) return; // läuft schon – Nachlauf ist vorgemerkt
  hb_(ss, HB_RUN_ID, "START", "Supervisor gestartet");
```

## Änderung 4 – `runKiraGeminiSupervisor`: Sperre im Lokal-Zweig lösen

**Streiche**

```javascript
      hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
      logToSheet('INFO', '[Supervisor] Lokal: Auftrag für den Mac bereitgestellt (coach-kira).');
      return;
```

**Ersetze durch**

```javascript
      hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
      logToSheet('INFO', '[Supervisor] Lokal: Auftrag für den Mac bereitgestellt (coach-kira).');
      ckSup_end_();
      return;
```

## Änderung 5 – `runKiraGeminiSupervisor`: Sperre nach dem Speichern lösen

**Streiche**

```javascript
    ckSup_finish_(ss, apiResponse, ckSupCtx, HB_RUN_ID);
```

**Ersetze durch**

```javascript
    ckSup_finish_(ss, apiResponse, ckSupCtx, HB_RUN_ID);
    ckSup_end_();
```

## Änderung 6 – `runKiraGeminiSupervisor`: Sperre auch im Fehlerfall lösen (Ende der Funktion)

**Streiche**

```javascript
  } catch (e) {
    hb_(ss, HB_RUN_ID, "END_ERROR", e && e.message ? e.message : String(e));
    logToSheet('ERROR', `🛑 FEHLER im Supervisor: ${e.message}`);
    logToSheet('ERROR', `Stack: ${e.stack}`);
  }
}
```

**Ersetze durch**

```javascript
  } catch (e) {
    hb_(ss, HB_RUN_ID, "END_ERROR", e && e.message ? e.message : String(e));
    logToSheet('ERROR', `🛑 FEHLER im Supervisor: ${e.message}`);
    logToSheet('ERROR', `Stack: ${e.stack}`);
    ckSup_end_();
  }
}
```

---

## Änderung 7 – `ckSup_finish_`: AI_REPORT_PLAN mitschreiben

**Streiche** (in `ckSup_finish_`, direkt nach den `const`-Zeilen)

```javascript
  const teResult = ctx.teResult;
```

**Ersetze durch**

```javascript
  const teResult = ctx.teResult;

  // AI_REPORT_PLAN wieder aktuell halten (wird von Kalender, „KI-Plan anwenden“ und Plan-Daten gelesen)
  try { ckSup_writeReportPlan_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] AI_REPORT_PLAN nicht geschrieben: ' + e.message); }
```

---

## Änderung 8 – `doGet`: Diagnose-Modus

**Streiche**

```javascript
  if (p.mode === 'supervisorJob') return ckSup_getJob_();
```

**Ersetze durch**

```javascript
  if (p.mode === 'supervisorJob') return ckSup_getJob_();
  if (p.mode === 'diag') return ckDash_diag_();
```

---

## Änderung 9 – neue Funktionen (ans Ende der Datei)

```javascript
// ============================================================================
// Wartung 19.10.2026: Doppelstarts verhindern, AI_REPORT_PLAN schreiben, Diagnose
// ============================================================================
/** true = dieser Lauf darf starten. Läuft schon einer (< 6 min), wird ein Nachlauf vorgemerkt. */
function ckSup_begin_(runId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return true; // im Zweifel laufen lassen
  try {
    const props = PropertiesService.getScriptProperties();
    const running = Number(props.getProperty('CK_SUP_RUNNING') || 0);
    if (running && Date.now() - running < 6 * 60 * 1000) {
      props.setProperty('CK_SUP_RERUN', '1');
      try { logToSheet('INFO', '[Supervisor] Läuft bereits – Nachlauf vorgemerkt (' + runId + ').'); } catch (x) {}
      return false;
    }
    props.setProperty('CK_SUP_RUNNING', String(Date.now()));
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Lauf beendet: Sperre lösen; war ein Nachlauf vorgemerkt, einmal neu starten (frische Daten). */
function ckSup_end_() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('CK_SUP_RUNNING');
    if (props.getProperty('CK_SUP_RERUN')) {
      props.deleteProperty('CK_SUP_RERUN');
      ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() === 'runKiraGeminiSupervisor' && t.getEventType() === ScriptApp.EventType.CLOCK) {
          try { ScriptApp.deleteTrigger(t); } catch (x) {}
        }
      });
      ScriptApp.newTrigger('runKiraGeminiSupervisor').timeBased().after(5000).create();
      try { logToSheet('INFO', '[Supervisor] Nachlauf gestartet.'); } catch (x) {}
    }
  } catch (e) {
    try { logToSheet('WARN', '[Supervisor] ckSup_end_: ' + e.message); } catch (x) {}
  }
}

/** Schreibt empfohlener_plan aus der KI-Antwort nach AI_REPORT_PLAN (gleiche Spalten wie früher). */
function ckSup_writeReportPlan_(ss, apiResponse) {
  let data = apiResponse;
  if (typeof data === 'string') {
    const s = data.replace(/`{3}(?:json)?/g, '').trim();
    try { data = JSON.parse(s); } catch (e) { data = JSON.parse(s.substring(s.indexOf('{'), s.lastIndexOf('}') + 1)); }
  }
  const plan = (data && Array.isArray(data.empfohlener_plan)) ? data.empfohlener_plan : [];
  if (!plan.length) return;
  const toDate = function (v) {
    const t = String(v || '').trim();
    let m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
    m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
    return t;
  };
  const header = ['Datum', 'Tag', 'Original Load (ESS)', 'Empfohlener Load (ESS)', 'Empfohlene Zone (KI)',
    'Beispiel-Training 1 (KI)', 'Beispiel-Training 2 (KI)', 'Prognostizierte Auswirkung (KI)', 'Wetterempfehlung'];
  const rows = plan.map(function (d) {
    return [toDate(d.datum), d.tag || '', Number(d.original_load_ess) || 0, Number(d.empfohlener_load_ess) || 0,
      d.empfohlene_zone || '', d.beispiel_training_1 || '', d.beispiel_training_2 || '',
      d.prognostizierte_auswirkung || '', d.wetterempfehlung || ''];
  });
  let sh = ss.getSheetByName(OUTPUT_PLAN_SHEET);
  if (!sh) sh = ss.insertSheet(OUTPUT_PLAN_SHEET);
  const lr = sh.getLastRow();
  if (lr > 1) sh.getRange(2, 1, lr - 1, Math.max(9, sh.getLastColumn())).clearContent();
  sh.getRange(1, 1, 1, 9).setValues([header]);
  sh.getRange(2, 1, rows.length, 9).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-MM-dd');
  logToSheet('INFO', '[Supervisor] AI_REPORT_PLAN aktualisiert (' + rows.length + ' Tage ab ' + String(plan[0].datum) + ').');
}

/** Diagnose für Wartung: welche Funktionen sind aktiv, Log-Zustand, Trigger. Schreibt einen Testeintrag. */
function ckDash_diag_() {
  const out = { ok: true, timestamp: ckDash_now_() };
  try {
    const props = PropertiesService.getScriptProperties();
    out.logToSheetHead = String(logToSheet).slice(0, 220);
    out.hbHead = String(hb_).slice(0, 120);
    out.logSheetName = LOG_SHEET_NAME;
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    out.logSheet = sh ? { lastRow: sh.getLastRow(), maxRows: sh.getMaxRows(), lastCol: sh.getLastColumn() } : null;
    const before = sh ? sh.getLastRow() : 0;
    logToSheet('DEBUG', '[Diag] Testeintrag');
    const sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    out.testLogWritten = sh2 ? (sh2.getLastRow() !== before || String(sh2.getRange(sh2.getLastRow(), 3).getValue()).indexOf('[Diag]') >= 0) : false;
    out.lastLogError = props.getProperty('CK_LOG_LAST_ERROR');
    out.supRunning = props.getProperty('CK_SUP_RUNNING');
    out.supRerun = props.getProperty('CK_SUP_RERUN');
    out.supervisorEngine = props.getProperty('SUPERVISOR_ENGINE');
    out.triggers = ScriptApp.getProjectTriggers().map(function (t) {
      return { fn: t.getHandlerFunction(), type: String(t.getEventType()), source: String(t.getTriggerSource()) };
    });
  } catch (e) {
    out.ok = false;
    out.error = String(e && e.message ? e.message : e);
  }
  return ckDash_json_(out);
}
```

---

## Danach

1. Speichern, dann Bereitstellen → Bereitstellungen verwalten → Stift → „Neue Version“ → Bereitstellen.
2. Kurz Bescheid geben – ich rufe `?mode=diag` auf, prüfe AI_LOG und Heartbeat und starte einen Supervisor-Lauf, der AI_REPORT_PLAN neu schreibt.
