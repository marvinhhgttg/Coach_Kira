# Apps-Script-Patch: KI-frei, Stufe 1 + 2

Stand 26.10.2026 · Grundlage: deine eingefügte `KiraGeminiSupervisor.gs` (Build `2026-10-17__kira_chat`, 12.599 Zeilen) und `ActivityReview.gs`.

## Was der Patch macht

Danach ruft das Apps Script **keine KI mehr auf** und liest **keinen KI-Schlüssel** mehr.

| Bereich | Bisher | Danach |
|---|---|---|
| Supervisor / KI-Report | Mac, wenn `SUPERVISOR_ENGINE=local`, sonst GPT-4o | immer Auftrag an den Mac, GPT-4o-Zweig gestrichen |
| Activity Review nach „Nach Aktivität“ | Gemini-Trigger (scheitert seit Monaten: „API key not valid“) | kein Trigger mehr, den Kommentar schreibt der Mac |
| Historien-Analyse, 3-Tage-Vorschläge, PlanApp-Analysen | Gemini | entfernt (Hinweis „KI läuft im Cockpit“) |
| Statusanalyse, Plan-Briefing, Konsolidierte Analyse (`mode=…`) | Perplexity im Apps Script | entfernt – macht das Cockpit (lokal oder Cloud über den Cockpit-Server) |
| Kira-Chat (`mode=chat`) | Perplexity im Apps Script | entfernt – macht das Cockpit |
| Ablage `saveBriefing`, `chatLog` | immer „lokal“ beschriftet | beschriftet „lokal“ oder „Cloud“ je nach Quelle |
| Menü „Coach Kira“ im Sheet | mit „Aktivitäts-Review“ und „Historien-Analyse“ | ohne diese zwei Einträge |

Unverändert bleiben: Datenaufbereitung und Prompt-Bau für Supervisor und Briefings (kommt in Stufe 3 auf den Mac), Speicherung, Telegram-Meldungen und `/status`, Wetter, Kalender, Garmin-Endpunkte, alle `doGet`/`doPost`-Routen.

Die Funktionsnamen bleiben als kurze Platzhalter bestehen, damit nichts anderes ins Leere ruft (z. B. alte HTML-Seiten).

## Weg A (empfohlen): komplette Datei ersetzen

1. In `KiraGeminiSupervisor.gs` alles markieren (Cmd+A) und löschen.
2. Inhalt der Datei **KiraGeminiSupervisor-KI-frei-Stufe1-2.gs** einfügen, speichern.
3. Weiter mit „Danach“ unten.

Das ist genau deine eingefügte Fassung mit den Änderungen aus Weg B – sonst nichts. Hast du seit dem Einfügen noch etwas am Script geändert, nimm Weg B.

## Weg B: Schritt für Schritt (Streiche / Ersetze)

### 1. Supervisor: GPT-4o-Zweig streichen

In `runKiraGeminiSupervisor()`:

**Streiche**
```javascript
    if (ckSup_engine_() === 'local') {
      ckSup_enqueue_(prompt, ckSupCtx, HB_RUN_ID);
      hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
      logToSheet('INFO', '[Supervisor] Lokal: Auftrag für den Mac bereitgestellt (coach-kira).');
      ckSup_end_();
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
    ckSup_end_();
```

**Ersetze durch**
```javascript
    // KI-frei (2026-10-26): Der Auftrag geht immer an den Mac. Das Ergebnis kommt per saveSupervisor zurück.
    ckSup_enqueue_(prompt, ckSupCtx, HB_RUN_ID);
    hb_(ss, HB_RUN_ID, "AI_WAIT_LOCAL", "Wartet auf Mac (coach-kira)");
    logToSheet('INFO', '[Supervisor] Auftrag für den Mac bereitgestellt.');
    ckSup_end_();
    return;
```

### 2. Menü: zwei Einträge streichen

In `onOpen()`:

**Streiche**
```javascript
    .addItem('Aktivitäts-Review starten', 'generateActivityReview') // Neuer Eintrag
    .addSeparator() 
    .addItem('Historien-Analyse starten', 'runHistoricalAnalysis') 
    .addSeparator() // <-- NEU
```

### 3. Funktionen durch Platzhalter ersetzen

Jeweils die **komplette Funktion** (von `function Name(` bis zur schließenden `}` direkt vor der nächsten Funktion) ersetzen.

#### `callGeminiAPI` (ab Zeile 2926, endet vor `function clearMemory(`)

**Ersetze die komplette Funktion durch**
```javascript
function callGeminiAPI(promptText) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt.
  throw new Error('callGeminiAPI: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}
```

#### `askKira` (ab Zeile 4908, endet vor `function createChatPrompt_V45(`)

**Ersetze die komplette Funktion durch**
```javascript
function askKira(frage) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt. Fragen: Cockpit-Chat oder Telegram (lokal auf dem Mac).
  return 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.';
}
```

#### `callGeminiStructured` (ab Zeile 6525, endet vor `function callOpenAI(`)

**Ersetze die komplette Funktion durch**
```javascript
function callGeminiStructured(promptText) {
  // KI-frei (2026-10-26): Gemini-Aufrufe entfernt.
  throw new Error('callGeminiStructured: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}
```

#### `callOpenAI` (ab Zeile 6615, endet vor `function getWeatherIconMap(`)

**Ersetze die komplette Funktion durch**
```javascript
function callOpenAI(promptText) {
  // KI-frei (2026-10-26): OpenAI-Aufrufe entfernt. Der Supervisor läuft auf dem Mac (coach-kira).
  throw new Error('callOpenAI: KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.');
}
```

#### `runHistoricalAnalysis` (ab Zeile 3602, endet vor `function buildHistoryPrompt(`)

**Ersetze die komplette Funktion durch**
```javascript
function runHistoricalAnalysis(e) {
  // KI-frei (2026-10-26): Historien-Analyse (Gemini) entfernt.
  logToSheet('INFO', '[History] Historien-Analyse ist entfernt (KI läuft nicht mehr im Apps Script).');
}
```

#### `triggerHistoryAnalysis_WebApp` (ab Zeile 4283, endet vor `function getLoadFactors(`)

**Ersetze die komplette Funktion durch**
```javascript
function triggerHistoryAnalysis_WebApp() {
  // KI-frei (2026-10-26): Historien-Analyse (Gemini) entfernt.
  return 'Die Historien-Analyse läuft nicht mehr im Apps Script. Langzeit-Auswertungen findest du im Cockpit (Analyse, Berichte).';
}
```

#### `triggerActivityReview` (ab Zeile 4562, endet vor `function applyKiraPlanToTimeline(`)

**Ersetze die komplette Funktion durch**
```javascript
function triggerActivityReview() {
  // KI-frei (2026-10-26): Den Aktivitätskommentar schreibt der Mac (coach-kira) selbst.
  logToSheet('INFO', '[ActivityReview] Läuft lokal auf dem Mac – kein Apps-Script-Trigger mehr.');
  return true; // Wichtig für das Form-Skript
}
```

#### `getNext3DaySuggestions` (ab Zeile 5755, endet vor `function buildNext3DayPrompt(`)

**Ersetze die komplette Funktion durch**
```javascript
function getNext3DaySuggestions() {
  // KI-frei (2026-10-26): 3-Tage-Vorschläge (Gemini) entfernt – Wochenplan im Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}
```

#### `getStrategicKiraBriefing` (ab Zeile 8447, endet vor `function getBriefHistoryContext(`)

**Ersetze die komplette Funktion durch**
```javascript
function getStrategicKiraBriefing(simData) {
  // KI-frei (2026-10-26): Gemini-Briefing entfernt – Plan-Briefing im Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}
```

#### `getCoachKiraAnalysis` (ab Zeile 9683, endet vor `function slimSimLast_(`)

**Ersetze die komplette Funktion durch**
```javascript
function getCoachKiraAnalysis(simLast) {
  // KI-frei (2026-10-26): Gemini-Analyse entfernt – Cockpit nutzen.
  return { success: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' };
}
```

#### `ckDash_runStatusAiAnalysis_` (ab Zeile 10567, endet vor `function ckDash_statusAiAnalysis_(`)

**Ersetze die komplette Funktion durch**
```javascript
function ckDash_runStatusAiAnalysis_(p) {
  // KI-frei (2026-10-26): Statusanalyse wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}
```

#### `ckDash_generatePlanBriefing_` (ab Zeile 10749, endet vor `function ckDash_collectPlanBriefingContext_(`)

**Ersetze die komplette Funktion durch**
```javascript
function ckDash_generatePlanBriefing_(p) {
  // KI-frei (2026-10-26): Plan-Briefing wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}
```

#### `ckDash_runConsolidatedAnalysis_` (ab Zeile 10976, endet vor `function ckDash_collectConsolidatedContext_(`)

**Ersetze die komplette Funktion durch**
```javascript
function ckDash_runConsolidatedAnalysis_(p) {
  // KI-frei (2026-10-26): Konsolidierte Analyse wird im Cockpit erstellt (lokal oder Cloud) und per saveBriefing gespeichert.
  return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}
```

#### `ckDash_chat_` (ab Zeile 11409, endet vor `function ckDash_chatLog_(`)

**Ersetze die komplette Funktion durch**
```javascript
function ckDash_chat_(e) {
  // KI-frei (2026-10-26): Der Chat läuft im Cockpit (lokal oder Cloud über den Cockpit-Server). Ablage weiter per chatLog.
  return ckDash_json_({ ok: false, error: 'KI läuft nicht mehr im Apps Script – lokal (Mac) oder Cloud über das Cockpit.' });
}
```

#### `ckSup_engine_` (ab Zeile 11830, endet vor `function ckSup_schema_(`)

**Ersetze die komplette Funktion durch**
```javascript
function ckSup_engine_() {
  return 'local'; // KI-frei (2026-10-26): Supervisor-KI läuft immer auf dem Mac (lokal oder Cloud über den Cockpit-Server).
}
```

### 4. Tote Prompt-Bausteine streichen

Diese Funktionen ruft danach niemand mehr auf. Jeweils komplett streichen (inklusive Kommentarblock direkt darüber):

- `buildHistoryPrompt` (direkt nach `runHistoricalAnalysis`)
- `buildActivityReviewPrompt`
- `createChatPrompt_V45`
- `buildNext3DayPrompt`

Wenn du unsicher bist: stehen lassen schadet nicht, sie werden nur nicht mehr benutzt.

### 5. Ablage „lokal“ oder „Cloud“ beschriften

In `ckDash_chatLog_(e)`:

**Streiche**
```javascript
    if (!q || !a) return ckDash_json_({ ok: false, error: 'Frage oder Antwort fehlt.' });
    logChatMessage('Marc (Cockpit)', q);
    logChatMessage('Kira (lokal, ' + m + ')', a);
```

**Ersetze durch**
```javascript
    const eng = String(b.engine || 'local') === 'cloud' ? 'Cloud' : 'lokal';
    if (!q || !a) return ckDash_json_({ ok: false, error: 'Frage oder Antwort fehlt.' });
    logChatMessage('Marc (Cockpit)', q);
    logChatMessage('Kira (' + eng + ', ' + m + ')', a);
```

In `ckDash_saveBriefing_(e)`:

**Streiche**
```javascript
    const raw = JSON.stringify({ engine: 'local', model: model, savedAt: ckDash_now_() });
```

**Ersetze durch**
```javascript
    const engine = String(b.engine || 'local') === 'cloud' ? 'cloud' : 'local';
    const raw = JSON.stringify({ engine: engine, model: model, savedAt: ckDash_now_() });
```

## Danach (für Weg A und B)

1. **Datei `ActivityReview.gs` löschen** (Drei-Punkte-Menü an der Datei → Löschen). Nichts ruft sie mehr auf.
2. **Speichern**, dann **Bereitstellen → Bereitstellungen verwalten → Bearbeiten (Stift) → Version: Neue Version → Bereitstellen**. Die URL bleibt gleich.
3. **Skripteigenschaften löschen** (Zahnrad „Projekteinstellungen“ → ganz unten „Skripteigenschaften“):
   - `GEMINI_API_KEY`
   - `OPENAI_API_KEY`
   - `PERPLEXITY_API_KEY` (dein Perplexity-Schlüssel liegt jetzt im Tresor des Cockpit-Servers)
   - `SUPERVISOR_ENGINE` (wird nicht mehr gebraucht)

   **Nicht** löschen: `TELEGRAM_*`, `OPENWEATHER_API_KEY`, `WEBAPP_API_KEY` und alles andere.
4. Optional: bei Google AI Studio und OpenAI die alten Schlüssel widerrufen.

## Prüfen

- Cockpit → Eingabe „Nach Aktivität“ senden: In `AI_LOG` steht „[ActivityReview] Läuft lokal auf dem Mac“, kein Gemini-Fehler mehr. Der Supervisor-Auftrag geht an den Mac (Heartbeat `AI_WAIT_LOCAL`).
- `…/exec?mode=runStatusAiAnalysis` liefert `ok:false` mit dem Hinweis – das Cockpit nutzt diesen Weg nicht mehr.
- Cockpit-Chat und Briefings laufen wie gehabt (lokal oder Cloud); in `AI_CHAT_LOG` steht jetzt „Kira (Cloud, sonar-pro)“ bzw. „Kira (lokal, coach-kira)“.

## Danach: Stufe 3

Prompt-Bau für Supervisor und Briefings wandert auf den Mac (gemeinsame Prompt-Datei für Cockpit und Mac), das Apps Script liefert nur noch Daten. Dazu Telegram-Fragen, die der Mac lokal beantwortet. Kommt als eigener Patch, abgestimmt mit dem Mac-Update.
