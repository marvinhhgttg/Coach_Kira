# Apps-Script-Patch: Supervisor lokal

Stand: 18.10.2026 · betrifft `runKiraGeminiSupervisor` (KI-Report, Blätter `AI_REPORT_STATUS` / `AI_REPORT_PLAN` usw.)

**Ausgangslage**

Der Supervisor ruft heute **OpenAI GPT-4o** (`callOpenAI`, Modell `gpt-4o-2024-08-06`, festes JSON-Schema) auf, nicht Gemini. Google-Server können deinen Mac nicht erreichen. Deshalb arbeitet der lokale Modus in zwei Hälften:

1. Der Supervisor sammelt die Daten, berechnet die Scores und baut den Prompt wie bisher. Im Lokal-Modus legt er dann einen **Auftrag** ab (Script-Cache, 6 Stunden gültig) und beendet sich. Im Heartbeat steht „Wartet auf Mac (coach-kira)“.
2. Der Mac-Dienst holt den Auftrag ab (`?mode=supervisorJob`) und rechnet mit `coach-kira`. Ollama hält sich dabei an dasselbe JSON-Schema wie GPT-4o. Das Ergebnis schickt er zurück (`saveSupervisor`).
3. `saveSupervisor` führt **exakt den bisherigen Speicherteil** aus: `writeOutputToSheets`, die TE-Korrektur in D11/C11/E11, die Telegram-Meldung und `schedulePostWork_`.

**Umschalten:** Skripteigenschaft `SUPERVISOR_ENGINE` = `local`. Fehlt sie oder steht sie auf etwas anderem, läuft alles wie bisher mit GPT-4o. Es gibt keinen automatischen Wechsel in die Cloud.

---

## Änderung 1 – `runKiraGeminiSupervisor`: KI-Aufruf und Speichern trennen

**Streiche** (ab `let apiResponse = null;` bis einschließlich der `END_OK`-Zeile):

```javascript
    let apiResponse = null;
    try {
        hb_(ss, HB_RUN_ID, "AI_CALL", "OpenAI Request raus");
        logToSheet('INFO', '📞 Rufe OpenAI (GPT-4o) mit JSON-Schema auf...');
        apiResponse = callOpenAI(prompt);
    } catch (apiError) {
        logToSheet('ERROR', `API Fehler: ${apiError.message}`);
        throw apiError; 
    }

    logToSheet('INFO', 'Antwort erhalten.');
    hb_(ss, HB_RUN_ID, "AI_OK", "OpenAI Antwort da");
    hb_(ss, HB_RUN_ID, "WRITE", "writeOutputToSheets startet");

    // ------------------------------------------------------------
    // SCHRITT 4: SPEICHERN (Hier schreibt die KI den Text in F11!)
    // ------------------------------------------------------------
    writeOutputToSheets(
        ss, 
        apiResponse, 
        fitnessScores, 
        nutritionScore, 
        scoreResult.num, 
        scoreResult.ampel, 
        subScores.recoveryScore, 
        subScores.recoveryAmpel, 
        subScores.trainingScore, 
        subScores.trainingAmpel
    );

    // +++ ANZEIGE-FIX (NUR FORMAT & AMPEL) +++
    // Wir lassen den Text in F11 (den die KI gerade geschrieben hat) in Ruhe!
    try {
       const statusSheet = ss.getSheetByName('AI_REPORT_STATUS');
       if (statusSheet && teResult.val > 0) {
           // Wir fixieren nur die Zahl als Text-Format (gegen "0,106...")
           statusSheet.getRange("D11").setNumberFormat("@").setValue(teResult.text);
           
           // Wir fixieren Score & Ampel (falls writeOutput geschlampt hat)
           let ampel = (teResult.val >= 0.2) ? "LILA" : (teResult.val >= 0.1) ? "GRÜN" : "ROT";
           let score = (teResult.val >= 0.2) ? 100 : (teResult.val >= 0.1) ? 80 : 25;
           
           statusSheet.getRange("C11").setValue(ampel);
           statusSheet.getRange("E11").setValue(score);
           
           // WICHTIG: KEIN setValue für F11 hier!
       }
    } catch(e) { console.log("Cell Fix Fehler: " + e.message); }

    logToSheet('INFO', `Berichte & Scores gespeichert.`);

    // ------------------------------------------------------------
    // SCHRITT 5: NACHARBEIT
    // ------------------------------------------------------------
    try {
       const uhrzeit = new Date().toLocaleTimeString('de-DE', {hour: '2-digit', minute:'2-digit'});
       const appUrl = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec"; 
       const configSheet = ss.getSheetByName("KK_CONFIG");
       if (configSheet) configSheet.getRange("B9").setValue(appUrl);
       if (typeof sendTelegram === 'function') {
         try {
            sendTelegram(`🫡 Commander, Analyse fertig (${uhrzeit})!\n📱 Dashboard: ${appUrl}`);
         } catch(e) { Utilities.sleep(1000); sendTelegram("🫡 Analyse fertig (Retry)"); }
       }
    } catch(e) {}

    // ------------------------------------------------------------
    // SCHRITT 6: NACHARBEIT (ASYNC)
    // ------------------------------------------------------------
    // Diese Schritte können (je nach Datenmenge) viel Zeit kosten. Damit der Supervisor
    // schneller fertig ist, laufen sie asynchron in einem Trigger.
    schedulePostWork_();

    Logger.log("✅ Supervisor-Prozess komplett abgeschlossen.");
    logToSheet('INFO', `✅ KI-Bewertung V137 erfolgreich.`);
    hb_(ss, HB_RUN_ID, "END_OK", "Supervisor fertig");

```

**Ersetze durch**

```javascript
    // Lokal-Modus: Auftrag an den Mac geben und hier beenden (Ergebnis kommt per saveSupervisor)
    const ckSupCtx = { fitnessScores: fitnessScores, nutritionScore: nutritionScore, scoreResult: scoreResult, subScores: subScores, teResult: teResult };
    if (ckSup_engine_() === 'local') {
      ckSup_enqueue_(prompt, ckSupCtx, HB_RUN_ID);
      hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
      logToSheet('INFO', '[Supervisor] Lokal: Auftrag für den Mac bereitgestellt (coach-kira).');
      return;
    }

    let apiResponse = null;
    try {
        hb_(ss, HB_RUN_ID, "AI_CALL", "OpenAI Request raus");
        logToSheet('INFO', '📞 Rufe OpenAI (GPT-4o) mit JSON-Schema auf...');
        apiResponse = callOpenAI(prompt);
    } catch (apiError) {
        logToSheet('ERROR', `API Fehler: ${apiError.message}`);
        throw apiError; 
    }

    logToSheet('INFO', 'Antwort erhalten.');
    ckSup_finish_(ss, apiResponse, ckSupCtx, HB_RUN_ID);
```

---

## Änderung 2 – `doGet`

**Streiche**

```javascript
  if (p.mode === 'activityReviews') return ckDash_activityReviews_(p);
```

**Ersetze durch**

```javascript
  if (p.mode === 'activityReviews') return ckDash_activityReviews_(p);
  if (p.mode === 'supervisorJob') return ckSup_getJob_();
```

---

## Änderung 3 – `doPost`

**Streiche**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveActivityReview') || /"kind"\s*:\s*"saveActivityReview"/.test(ckBody)) {
    return ckDash_saveActivityReview_(e);
  }
```

**Ersetze durch**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveActivityReview') || /"kind"\s*:\s*"saveActivityReview"/.test(ckBody)) {
    return ckDash_saveActivityReview_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveSupervisor') || /"kind"\s*:\s*"saveSupervisor"/.test(ckBody)) {
    return ckSup_save_(e);
  }
```

---

## Änderung 4 – neue Funktionen (ans Ende der Datei)

```javascript
// ============================================================================
// Supervisor lokal (Mac-Dienst, coach-kira) – Auftrag ablegen, abholen, speichern
// ============================================================================
function ckSup_engine_() {
  return String(PropertiesService.getScriptProperties().getProperty('SUPERVISOR_ENGINE') || 'cloud').toLowerCase();
}

function ckSup_schema_() {
  // identisch mit dem Schema aus callOpenAI()
    return {
    "name": "fitness_report",
    "strict": true,
    "schema": {
      "type": "object",
      "properties": {
        "empfehlung_zukunft": {
          "type": "object",
          "properties": {
            "plan_status": { "type": "string", "enum": ["FREIGABE", "BEOBACHTUNG", "INTERVENTION"] },
            "text": { "type": "string", "description": "Ausführliche Gesamtanalyse" }
          },
          "required": ["plan_status", "text"],
          "additionalProperties": false
        },
        "bewertung_ernaehrung_7d": {
          "type": "object",
          "properties": {
            "text": { "type": "string", "description": "Kurzbewertung Ernährung" }
          },
          "required": ["text"],
          "additionalProperties": false
        },
        "einzelscore_kommentare": {
          "type": "array",
          "description": "Liste der Kommentare für jeden einzelnen Score",
          "items": {
            "type": "object",
            "properties": {
              "metrik": { "type": "string" },
              "text_info": { "type": "string", "description": "Kurzer Kommentar (1 Satz)" }
            },
            "required": ["metrik", "text_info"],
            "additionalProperties": false
          }
        },
        "empfohlener_plan": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "datum": { "type": "string" },
              "tag": { "type": "string" },
              "original_load_ess": { "type": "number" },
              "empfohlener_load_ess": { "type": "number" },
              "empfohlene_zone": { "type": "string" },
              "beispiel_training_1": { "type": "string" },
              "beispiel_training_2": { "type": "string" },
              "prognostizierte_auswirkung": { 
                "type": "string", 
                "description": "Detaillierte Erklärung." 
              },
              "wetterempfehlung": { "type": "string" }
            },
            "required": ["datum", "tag", "original_load_ess", "empfohlener_load_ess", "empfohlene_zone", "beispiel_training_1", "beispiel_training_2", "prognostizierte_auswirkung", "wetterempfehlung"],
            "additionalProperties": false
          }
        }
      },
      "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare", "empfohlener_plan"],
      "additionalProperties": false
    }
  };
}

function ckSup_enqueue_(prompt, ctx, runId) {
  const cache = CacheService.getScriptCache();
  const job = {
    id: String(runId) + '-' + Date.now(),
    runId: String(runId),
    createdAt: ckDash_now_(),
    system: 'Du bist ein präziser Daten-Analyst. Deine Aufgabe ist es, die Daten zu analysieren und NUR das angeforderte JSON zu füllen.',
    prompt: String(prompt),
    ctx: ctx
  };
  const s = JSON.stringify(job);
  const size = 90000;
  const parts = {};
  const n = Math.ceil(s.length / size);
  for (let i = 0; i < n; i++) parts['ckSup_job_' + i] = s.substr(i * size, size);
  parts['ckSup_job_n'] = String(n);
  parts['ckSup_job_id'] = job.id;
  cache.putAll(parts, 21600); // 6 Stunden
}

function ckSup_readJob_() {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get('ckSup_job_n') || '0', 10);
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push('ckSup_job_' + i);
  const got = cache.getAll(keys);
  let s = '';
  for (let i = 0; i < n; i++) {
    if (got['ckSup_job_' + i] == null) return null;
    s += got['ckSup_job_' + i];
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}

function ckSup_clearJob_() {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get('ckSup_job_n') || '0', 10);
  const keys = ['ckSup_job_n', 'ckSup_job_id'];
  for (let i = 0; i < n; i++) keys.push('ckSup_job_' + i);
  cache.removeAll(keys);
}

function ckSup_getJob_() {
  try {
    const job = ckSup_readJob_();
    if (!job) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), engine: ckSup_engine_(), job: null });
    return ckDash_json_({
      ok: true, timestamp: ckDash_now_(), engine: ckSup_engine_(),
      job: { id: job.id, createdAt: job.createdAt, system: job.system, prompt: job.prompt, schema: ckSup_schema_().schema }
    });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckSup_save_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let runId = 'local';
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const job = ckSup_readJob_();
    if (!job) throw new Error('Kein offener Supervisor-Auftrag (abgelaufen oder schon gespeichert).');
    if (String(b.id || '') !== job.id) throw new Error('Auftrags-ID passt nicht (neuerer Auftrag vorhanden).');
    const content = String(b.content || '').trim();
    if (!content) throw new Error('Leeres Ergebnis.');
    runId = job.runId || runId;
    ckSup_clearJob_();
    logToSheet('INFO', '[Supervisor] Lokales Ergebnis empfangen (' + String(b.model || 'coach-kira').slice(0, 40) + ').');
    ckSup_finish_(ss, content, job.ctx, runId);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), id: job.id });
  } catch (err) {
    try { hb_(ss, runId, "END_ERROR", String(err && err.message ? err.message : err)); } catch (x) {}
    try { logToSheet('ERROR', '[Supervisor] Lokal speichern fehlgeschlagen: ' + (err && err.message ? err.message : err)); } catch (x) {}
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** Speicherteil des Supervisors – unverändert aus runKiraGeminiSupervisor übernommen. */
function ckSup_finish_(ss, apiResponse, ctx, HB_RUN_ID) {
  const fitnessScores = ctx.fitnessScores;
  const nutritionScore = ctx.nutritionScore;
  const scoreResult = ctx.scoreResult;
  const subScores = ctx.subScores;
  const teResult = ctx.teResult;

    hb_(ss, HB_RUN_ID, "AI_OK", "OpenAI Antwort da");
    hb_(ss, HB_RUN_ID, "WRITE", "writeOutputToSheets startet");

    // ------------------------------------------------------------
    // SCHRITT 4: SPEICHERN (Hier schreibt die KI den Text in F11!)
    // ------------------------------------------------------------
    writeOutputToSheets(
        ss, 
        apiResponse, 
        fitnessScores, 
        nutritionScore, 
        scoreResult.num, 
        scoreResult.ampel, 
        subScores.recoveryScore, 
        subScores.recoveryAmpel, 
        subScores.trainingScore, 
        subScores.trainingAmpel
    );

    // +++ ANZEIGE-FIX (NUR FORMAT & AMPEL) +++
    // Wir lassen den Text in F11 (den die KI gerade geschrieben hat) in Ruhe!
    try {
       const statusSheet = ss.getSheetByName('AI_REPORT_STATUS');
       if (statusSheet && teResult.val > 0) {
           // Wir fixieren nur die Zahl als Text-Format (gegen "0,106...")
           statusSheet.getRange("D11").setNumberFormat("@").setValue(teResult.text);
           
           // Wir fixieren Score & Ampel (falls writeOutput geschlampt hat)
           let ampel = (teResult.val >= 0.2) ? "LILA" : (teResult.val >= 0.1) ? "GRÜN" : "ROT";
           let score = (teResult.val >= 0.2) ? 100 : (teResult.val >= 0.1) ? 80 : 25;
           
           statusSheet.getRange("C11").setValue(ampel);
           statusSheet.getRange("E11").setValue(score);
           
           // WICHTIG: KEIN setValue für F11 hier!
       }
    } catch(e) { console.log("Cell Fix Fehler: " + e.message); }

    logToSheet('INFO', `Berichte & Scores gespeichert.`);

    // ------------------------------------------------------------
    // SCHRITT 5: NACHARBEIT
    // ------------------------------------------------------------
    try {
       const uhrzeit = new Date().toLocaleTimeString('de-DE', {hour: '2-digit', minute:'2-digit'});
       const appUrl = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec"; 
       const configSheet = ss.getSheetByName("KK_CONFIG");
       if (configSheet) configSheet.getRange("B9").setValue(appUrl);
       if (typeof sendTelegram === 'function') {
         try {
            sendTelegram(`🫡 Commander, Analyse fertig (${uhrzeit})!\n📱 Dashboard: ${appUrl}`);
         } catch(e) { Utilities.sleep(1000); sendTelegram("🫡 Analyse fertig (Retry)"); }
       }
    } catch(e) {}

    // ------------------------------------------------------------
    // SCHRITT 6: NACHARBEIT (ASYNC)
    // ------------------------------------------------------------
    // Diese Schritte können (je nach Datenmenge) viel Zeit kosten. Damit der Supervisor
    // schneller fertig ist, laufen sie asynchron in einem Trigger.
    schedulePostWork_();

    Logger.log("✅ Supervisor-Prozess komplett abgeschlossen.");
    logToSheet('INFO', `✅ KI-Bewertung V137 erfolgreich.`);
    hb_(ss, HB_RUN_ID, "END_OK", "Supervisor fertig");

}
```

---

## Danach

1. Speichern, dann Bereitstellen → Bereitstellungen verwalten → Stift → „Neue Version“ → Bereitstellen.
2. **Noch nicht** auf lokal stellen. Gib mir Bescheid – ich teste zuerst, dass der Cloud-Weg unverändert läuft.
3. Für den Lokal-Test: Projekteinstellungen → Skripteigenschaften → `SUPERVISOR_ENGINE` = `local`. Zurück auf GPT-4o: Eigenschaft löschen oder `cloud` eintragen.
