# Apps-Script-Patch: KI-frei, Stufe 3 (Streiche / Ersetze)

Stand 26.10.2026 · Grundlage: die bereitgestellte Stufe-1+2-Fassung von `KiraGeminiSupervisor.gs` (11.043 Zeilen).

Ergebnis: kein Prompt-Text und keine KI-Logik mehr im Apps Script. Der Supervisor-Auftrag kommt als Datenpaket, es gibt den neuen Modus `briefingData`, der Supervisor-Start nimmt die Wahl lokal/Cloud mit, und Telegram-Fragen gehen an den Mac. Für dieselben Daten baut der Mac zeichengenau denselben Prompt wie bisher (getestet in vier Szenarien). Mac und Cockpit sind schon umgestellt und laufen mit Stufe 2 und Stufe 3.

**So arbeitest du:** Die Schritte gehen von unten nach oben durch die Datei, also in der Reihenfolge 1 → 14 abarbeiten. Die Zeilenangaben sind nur zur Orientierung (Stand vor Schritt 1). Gesucht wird immer mit Cmd+F nach der ersten Codezeile des Streiche-Blocks.

---

## 1. Telegram-Funktionen am Dateiende anfügen

Ganz ans **Ende der Datei**, nach der letzten schließenden `}`:

**Füge ein**
```javascript
// ============================================================================
// KI-frei (Stufe 3, 2026-10-26): Telegram-Fragen – Warteschlange für den Mac
//   Bot → ckTg_enqueue_ · Mac holt mode=telegramJobs · antwortet POST kind=telegramAnswer
// ============================================================================
const CK_TG_QUEUE_KEY = 'CK_TG_QUEUE';
const CK_TG_VERLAUF_KEY = 'CK_TG_VERLAUF';

function ckTg_read_(key) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(key) || '[]'); } catch (e) { return []; }
}

function ckTg_send_(chatId, text) {
  // ohne Markdown, damit Sonderzeichen in Kiras Antwort nichts kaputt machen
  const cfg = getTelegramConfig_();
  if (!cfg.token) { logToSheet('ERROR', '[Telegram] TELEGRAM_BOT_TOKEN fehlt in ScriptProperties'); return; }
  UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chatId, text: String(text).slice(0, 4000) })
  });
}

function ckTg_enqueue_(chatId, text, user) {
  const cfg = getTelegramConfig_();
  if (cfg.chatId && String(chatId) !== String(cfg.chatId)) return; // nur Marcs Chat
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const q = ckTg_read_(CK_TG_QUEUE_KEY);
    q.push({ id: 'tg-' + Date.now(), chatId: String(chatId), frage: String(text).slice(0, 2000), user: String(user || ''), ts: ckDash_now_() });
    PropertiesService.getScriptProperties().setProperty(CK_TG_QUEUE_KEY, JSON.stringify(q.slice(-10)));
  } finally {
    lock.releaseLock();
  }
  ckTg_send_(chatId, 'Kira denkt nach – die Antwort kommt vom Mac (lokal), meist in 1–2 Minuten.');
}

function ckTg_jobs_() {
  return ckDash_json_({
    ok: true, timestamp: ckDash_now_(),
    jobs: ckTg_read_(CK_TG_QUEUE_KEY).map(function (j) { return { id: j.id, frage: j.frage, ts: j.ts }; }),
    verlauf: ckTg_read_(CK_TG_VERLAUF_KEY)
  });
}

function ckTg_answer_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const id = String(b.id || '');
    const text = String(b.text || '').trim();
    if (!id || !text) return ckDash_json_({ ok: false, error: 'id oder text fehlt.' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let job = null;
    try {
      const q = ckTg_read_(CK_TG_QUEUE_KEY);
      job = q.filter(function (j) { return j.id === id; })[0] || null;
      if (!job) return ckDash_json_({ ok: false, error: 'Frage nicht (mehr) in der Warteschlange.' });
      PropertiesService.getScriptProperties().setProperty(CK_TG_QUEUE_KEY, JSON.stringify(q.filter(function (j) { return j.id !== id; })));
      const v = ckTg_read_(CK_TG_VERLAUF_KEY);
      v.push({ frage: job.frage.slice(0, 500), antwort: text.slice(0, 1500) });
      PropertiesService.getScriptProperties().setProperty(CK_TG_VERLAUF_KEY, JSON.stringify(v.slice(-3)));
    } finally {
      lock.releaseLock();
    }
    ckTg_send_(job.chatId, text);
    const m = String(b.model || 'coach-kira').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const eng = String(b.engine || 'local') === 'cloud' ? 'Cloud' : 'lokal';
    logChatMessage('Marc (Telegram)', job.frage);
    logChatMessage('Kira (' + eng + ', ' + m + ', Telegram)', text);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), id: id });
  } catch (err) {
    return ckDash_json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}
```

## 2. Prompt-Bausteine im unteren Dateiteil streichen

### a) `ckX_planungsBlock_` (Zeile 10911–10943)

**Streiche** von
```javascript
/** Textblock für den Supervisor-Prompt: aktive Ziele/Touren und der Saisonrahmen der laufenden Woche */
function ckX_planungsBlock_() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
// ============================================================
```
(diese Zeile bleibt stehen).

### b) `ckRegel_monotonieBlock_` (Zeile 10885–10890)

**Streiche** von
```javascript
/** Block für den Supervisor-Prompt */
function ckRegel_monotonieBlock_() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
/** GET mode=loadFactors · KK_LOAD_FACTORS (ESS pro Stunde je Sport/Zone) und KK_ELEV_FACTORS (nur lesen) */
```
(diese Zeile bleibt stehen).

### c) `ckRegel_monotonie_` (Zeile 10874–10883)

**Streiche** von
```javascript
/** Monotonie-Regel (gleicher Wortlaut wie im Cockpit, lib/monotony.ts) */
function ckRegel_monotonie_() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
/** Block für den Supervisor-Prompt */
```
(diese Zeile bleibt stehen).

### d) `ckX_trefferPromptBlock_` (Zeile 10855–10872)

**Streiche** von
```javascript
/** Textblock für den Supervisor-Prompt: letzte Wochenbilanz der Kira-Treffsicherheit */
function ckX_trefferPromptBlock_() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
/** Monotonie-Regel (gleicher Wortlaut wie im Cockpit, lib/monotony.ts) */
```
(diese Zeile bleibt stehen).

## 3. `ckSup_getJob_()`: Auftrag als Daten ausliefern

**Streiche**
```javascript
      job: { id: job.id, createdAt: job.createdAt, system: job.system, prompt: job.prompt, schema: ckSup_schema_().schema }
```

**Ersetze durch**
```javascript
      job: { id: job.id, createdAt: job.createdAt, engine: job.engine || 'local', data: job.data || null }
```

## 4. `ckSup_enqueue_` ersetzen und `ckSup_promptData_` davor einfügen

**Streiche** die komplette Funktion
```javascript
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
```

**Ersetze durch**
```javascript
function ckSup_promptData_(data, fitnessScores, nutritionScore, gesamtScoreNum, gesamtScoreAmpel, zoneFactors, elevFactors, weekConfigCSV, isActivityDone, wetterDaten) {
  const heute = (data && data.heute) || {};
  const tz = Session.getScriptTimeZone();

  // TE-Balance heute aus AI_DATA_FORECAST (null = nicht gefunden → Prompt nutzt Monotonie-Varianz)
  let teBalanceHeute = null;
  try {
    const fcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI_DATA_FORECAST');
    if (fcSheet && fcSheet.getLastRow() > 1) {
      const fcData = fcSheet.getDataRange().getValues();
      const fcHeaders = fcData[0].map(h => h.toString().trim());
      const fcHeadersLc = fcHeaders.map(h => h.toLowerCase());

      const teIdx = fcHeadersLc.indexOf('te balance') > -1
        ? fcHeadersLc.indexOf('te balance')
        : fcHeadersLc.indexOf('te balance (% intensiv)');

      if (teIdx !== -1) {
        const tz = Session.getScriptTimeZone();
        const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
        let todayRowIdx = -1;

        // 1) Bevorzugt: is_today == 1
        const isTodayIdx = fcHeadersLc.indexOf('is_today');
        if (isTodayIdx !== -1) {
          for (let r = 1; r < fcData.length; r++) {
            if (parseGermanFloat(fcData[r][isTodayIdx]) === 1) {
              todayRowIdx = r;
              break;
            }
          }
        }

        // 2) Fallback: Datum == heute
        if (todayRowIdx === -1) {
          const dateIdx = fcHeadersLc.indexOf('datum') > -1
            ? fcHeadersLc.indexOf('datum')
            : fcHeadersLc.indexOf('date');

          if (dateIdx !== -1) {
            for (let r = 1; r < fcData.length; r++) {
              const rawDate = fcData[r][dateIdx];
              const dObj = rawDate instanceof Date ? rawDate : new Date(rawDate);
              if (dObj instanceof Date && !isNaN(dObj.getTime())) {
                const rowKey = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
                if (rowKey === todayKey) {
                  todayRowIdx = r;
                  break;
                }
              }
            }
          }
        }

        if (todayRowIdx !== -1) {
          let teRaw = fcData[todayRowIdx][teIdx];
          let teNum = Number.NaN;

          if (typeof teRaw === 'number') {
            teNum = teRaw;
          } else {
            teNum = parseGermanFloat(String(teRaw).replace('%', ''));
          }

          if (Number.isFinite(teNum)) {
            if (Math.abs(teNum) <= 1.0 && teNum !== 0) teNum = teNum * 100;
            teBalanceHeute = teNum.toFixed(1).replace('.', ',') + "%";
          }
        }
      }
    }
  } catch (e) {
    teBalanceHeute = null;
  }

  // Datum und Wochentag von heute
  let heuteDateObj = heute['date'];
  if (!(heuteDateObj instanceof Date)) heuteDateObj = new Date(heuteDateObj);
  const heuteDatum = Utilities.formatDate(heuteDateObj, tz, "yyyy-MM-dd");
  let heuteTag = heute['date.2'];
  if (!heuteTag || heuteTag === "") heuteTag = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"][heuteDateObj.getDay()];

  // Wissensbasis KK_WISSEN: aktive Inhalte (Spalte B = 'x', Spalte C = Text); null = Blatt fehlt
  let wissen = null;
  const wSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KK_WISSEN');
  if (wSheet) {
    wissen = [];
    const w = wSheet.getDataRange().getValues();
    for (let i = 1; i < w.length; i++) {
      if (String(w[i][1]).toLowerCase() === 'x' && w[i][2]) wissen.push(String(w[i][2]));
    }
  }

  // Treffsicherheit (letzte Wochen-Auswertung) und Planung (Ziele, Saisonrahmen) aus den KI-Extras
  let treffer = null;
  let planung = null;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (rows[i][1] === 'treffer_woche' && rows[i][2]) { treffer = { datum: rows[i][0], text: String(rows[i][2]) }; break; }
      }
      const heuteKey = Utilities.formatDate(new Date(), tz || 'Europe/Berlin', 'yyyy-MM-dd');
      const ziele = [];
      let saison = null;
      rows.forEach(function (r) {
        if (r[1] === 'ziel' && r[0] >= heuteKey) ziele.push({ datum: r[0], text: String(r[2]) });
        if (r[1] === 'saison' && (!saison || r[0] > saison[0])) saison = r;
      });
      let wochen = [];
      if (saison && saison[3]) { try { wochen = JSON.parse(saison[3]).wochen || []; } catch (e) {} }
      planung = { heute: heuteKey, ziele: ziele, wochen: wochen };
    }
  } catch (e) {}

  // JSON-sicher (Date → ISO-Text), wie es der Mac per supervisorJob bekommt
  return JSON.parse(JSON.stringify({
    v: 1,
    heute: heute,
    heute_datum: heuteDatum,
    heute_tag: heuteTag,
    zukunft_plan: (data && data.zukunft && data.zukunft.plan) || [],
    baseline: (data && data.baseline) || {},
    history: (data && data.history) || '',
    monotony_varianz: data ? data.monotony_varianz : 0,
    rtp_status: data ? data.rtp_status : '',
    fitnessScores: fitnessScores,
    nutritionScore: nutritionScore,
    gesamtScoreNum: gesamtScoreNum,
    gesamtScoreAmpel: gesamtScoreAmpel,
    zoneFactors: zoneFactors,
    elevFactors: elevFactors,
    weekConfigCSV: weekConfigCSV,
    isActivityDone: !!isActivityDone,
    wetterDaten: wetterDaten,
    teBalanceHeute: teBalanceHeute,
    wissen: wissen,
    treffer: treffer,
    planung: planung
  }));
}

function ckSup_enqueue_(promptData, ctx, runId) {
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const engine = String(props.getProperty('CK_SUP_ENGINE_NEXT') || 'local') === 'cloud' ? 'cloud' : 'local';
  props.deleteProperty('CK_SUP_ENGINE_NEXT');
  const job = {
    id: String(runId) + '-' + Date.now(),
    runId: String(runId),
    createdAt: ckDash_now_(),
    engine: engine,
    data: promptData,
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
```

## 5. `ckSup_schema_` streichen

Das JSON-Schema liegt jetzt in der Prompt-Datei von Cockpit und Mac.

### `ckSup_schema_` (Zeile 10279–10344)

**Streiche** von
```javascript
function ckSup_schema_() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
function ckSup_enqueue_(prompt, ctx, runId) {
```
(diese Zeile bleibt stehen).

## 6. `ckDash_briefingPrompt_` ersetzen (Daten statt Prompt)

**Streiche** die komplette Funktion
```javascript
function ckDash_briefingPrompt_(p) {
  try {
    const type = String((p && p.type) || '');
    const b = ckDash_buildBriefingMessages_(type);
    // Monotonie-Regel an die System-Anweisung anhängen
    if (b.messages && b.messages.length && b.messages[0].role === 'system') {
      b.messages[0].content = String(b.messages[0].content) + '\n\n' + ckRegel_monotonie_();
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, format: b.format, messages: b.messages });

  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}
```

**Ersetze durch**
```javascript
function ckDash_briefingPrompt_(p) {
  // KI-frei (Stufe 3): Prompt-Bau liegt in kira_prompts (Cockpit/Mac). Daten: mode=briefingData.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'briefingPrompt entfällt – bitte mode=briefingData nutzen.' });
}

/** KI-frei (Stufe 3): Daten für Status-, Plan- und konsolidiertes Briefing. */
function ckDash_briefingData_(p) {
  try {
    const type = String((p && p.type) || '');
    let daten;
    if (type === 'status') daten = { status: ckDash_getReportStatusSummary_() };
    else if (type === 'plan') daten = { status: ckDash_getReportStatusSummary_(), sim: ckDash_collectPlanBriefingContext_() };
    else if (type === 'consolidated') daten = { context: ckDash_collectConsolidatedContext_() };
    else throw new Error('Unbekannter Briefing-Typ: ' + type);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, daten: JSON.parse(JSON.stringify(daten)) });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}
```

## 7. `ckDash_buildBriefingMessages_` streichen

### `ckDash_buildBriefingMessages_` (Zeile 9980–10098)

**Streiche** von
```javascript
// ============================================================================
// Briefings lokal oder in der Cloud – Prompt holen, lokales Ergebnis speichern
// ============================================================================
function ckDash_buildBriefingMessages_(type) {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
function ckDash_briefingPrompt_(p) {
```
(diese Zeile bleibt stehen).

## 8. `getDeepKnowledgeBase` streichen

### `getDeepKnowledgeBase` (Zeile 5944–5971)

**Streiche** von
```javascript
/**
 * Liest das Expertenwissen aus dem Sheet 'KK_WISSEN'.
 * Nur Zeilen, bei denen in Spalte B ein 'x' steht.
 */
function getDeepKnowledgeBase() {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
// --- PINNWAND / NOTIZEN FUNKTIONEN ---
```
(diese Zeile bleibt stehen).

## 9. `createDashboardPrompt_V39` streichen (wird nirgends aufgerufen)

### `createDashboardPrompt_V39` (Zeile 4316–4434)

**Streiche** von
```javascript
/**
 * V41-CONTEXT: Erstellt den Prompt für das Dashboard.
 * Korrektur: Bei "Activity Done" wird nun die physiologische Passung bewertet,
 * statt eines sinnlosen Soll/Ist-Zahlenvergleichs.
 */
function createDashboardPrompt_V39(isActivityDone) {
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
/**
```
(diese Zeile bleibt stehen).

## 10. Telegram: freie Fragen in `handleCommand()`

**Streiche**
```javascript
  const user = contents.message.from.first_name || "Commander";

  if (text.startsWith("/status")) {
```

**Ersetze durch**
```javascript
  const user = contents.message.from.first_name || "Commander";

  // KI-frei (Stufe 3): freie Fragen beantwortet Kira lokal auf dem Mac
  if (text && !text.startsWith('/')) {
    ckTg_enqueue_(chatId, text, user);
    return;
  }

  if (text.startsWith("/status")) {
```

## 11. `doPost()`: Antwort vom Mac annehmen

**Streiche**
```javascript
  if ((e && e.parameter && e.parameter.mode === 'chatLog') || /"kind"\s*:\s*"chatLog"/.test(ckBody)) {
```

**Ersetze durch**
```javascript
  if (/"kind"\s*:\s*"telegramAnswer"/.test(ckBody)) {
    return ckTg_answer_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'chatLog') || /"kind"\s*:\s*"chatLog"/.test(ckBody)) {
```

## 12. `ckDash_queueAction_()`: KI-Wahl beim Supervisor-Start merken

**Streiche**
```javascript
        functionName: functionName
      });
    }

    const lock = LockService.getScriptLock();
```

**Ersetze durch**
```javascript
        functionName: functionName
      });
    }

    // KI-frei (Stufe 3): gewünschte KI (lokal/Cloud) für den nächsten Supervisor-Auftrag merken
    if (functionName === 'runKiraGeminiSupervisor') {
      PropertiesService.getScriptProperties().setProperty('CK_SUP_ENGINE_NEXT', String(p.engine || '') === 'cloud' ? 'cloud' : 'local');
    }

    const lock = LockService.getScriptLock();
```

## 13. `doGet()`: zwei neue Modi

**Streiche**
```javascript
  if (p.mode === 'briefingPrompt') return ckDash_briefingPrompt_(p);
```

**Ersetze durch**
```javascript
  if (p.mode === 'briefingPrompt') return ckDash_briefingPrompt_(p);
  if (p.mode === 'briefingData') return ckDash_briefingData_(p);
  if (p.mode === 'telegramJobs') return ckTg_jobs_();
```

## 14. Supervisor-Prompt im oberen Dateiteil

### a) `buildGeminiPrompt` (Zeile 2126–2133)

**Streiche** von
```javascript
/**
 * Baut den detaillierten Text-Prompt (V42 - mit Wetter)
 * Ruft die Master-Prompt-Funktion auf und übergibt alle Argumente.
 */
function buildGeminiPrompt(data, fitnessScores, nutritionScore, gesamtScoreNum, gesamtScoreAmpel, zoneFactors, elevFactors, weekConfigCSV, isActivityDone, wetterDaten) { // <-- NEU (V42)
```
bis einschließlich der schließenden Klammer der Funktion
```javascript
}
```
direkt vor
```javascript
/**
```
(diese Zeile bleibt stehen).

### b) `createMasterGeminiPrompt` (Zeile 1535–2124, rund 590 Zeilen)

**Streiche** von
```javascript
/**
 * V145-SMART-ENTRY: Intelligenter RTP-Einstieg.
 * Formel: RTP-Startpunkt = 10 - Krankheitstage.
 * Kurze Krankheit -> Später Einstieg (kurzes RTP). Lange Krankheit -> Start bei 1.
 */
function createMasterGeminiPrompt(data, fitnessScores, nutritionScore, gesamtScoreNum, gesamtScoreAmpel, zoneFactors, elevFactors, weekConfigCSV, isActivityDone, wetterDaten) {
```
bis einschließlich
```javascript
  return prompt;
}
```
Achtung: Mitten in der Funktion stehen Zeilen, die nur aus `}` bestehen (das JSON-Beispiel im Prompt-Text). Das richtige Ende ist `  return prompt;` mit der `}` direkt darunter. Danach kam der Kommentar über `buildGeminiPrompt`, den du in a) schon gestrichen hast.

### c) `WISSENSBLOCK_V105` (Zeile 1500–1533)

**Streiche** von
```javascript
const WISSENSBLOCK_V105 = `
--- KERNWISSEN (Firstbeat Whitepapers) ---
```
bis einschließlich
```javascript
`;
```
direkt vor
```javascript
/**
```
(diese Zeile bleibt stehen).

### d) `runKiraGeminiSupervisor()`: Daten statt Prompt

**Streiche**
```javascript
    // Prompt erstellen
    const prompt = createMasterGeminiPrompt(
```

**Ersetze durch**
```javascript
    // KI-frei (Stufe 3): nur Daten – den Prompt baut der Mac (kira_prompts)
    const promptData = ckSup_promptData_(
```

und etwas weiter unten

**Streiche**
```javascript
    ckSup_enqueue_(prompt, ckSupCtx, HB_RUN_ID);
```

**Ersetze durch**
```javascript
    ckSup_enqueue_(promptData, ckSupCtx, HB_RUN_ID);
```

---

## Danach

1. Speichern. Oben im Editor darf kein roter Syntaxfehler stehen.
2. **Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Version: Neue Version → Bereitstellen.**
3. Optional: auf der Trigger-Seite die vier deaktivierten Zeit-Trigger löschen. `onFormSubmit` stehen lassen.

## Prüfen

- Im Script nach `Du bist`, `createMasterGeminiPrompt` und `ckSup_schema_` suchen: keine Treffer mehr.
- `…/exec?mode=briefingData&type=status` liefert `{ ok: true, daten: … }`.
- KI-Report im Cockpit starten: Das Mac-Protokoll zeigt „Auftrag … übernommen (lokal, Prompt … Zeichen)“, mit Schalter Cloud „(Cloud, …)“.
- Telegram: eine Frage ohne `/` schicken. Sofort kommt „Kira denkt nach…“, nach 1–2 Minuten die Antwort vom Mac.

Passt irgendwo ein Suchtext nicht genau, sag mir den Schritt.
