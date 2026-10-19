# Apps-Script-Patch: Original-Load im KI-Plan

Stand: 19.10.2026 · Nachtrag zum Wartungs-Patch

**Ursache**

Im Prompt des Supervisors (`createMasterGeminiPrompt`) steht für normale Tage `Geplanter Load: 190`. Die JSON-Anweisung verlangt aber, `original_load_ess` „stumpf aus `[ORIGINAL-PLAN: ...]`“ zu kopieren. Dieser Marker kommt nur an RTP-Tagen (Rückkehr nach Krankheit) vor. Das Beispiel im Prompt enthält außerdem `"original_load_ess": 0`. Deshalb übernimmt das Modell die 0. Betroffen sind GPT-4o und `coach-kira` gleichermaßen.

**Behebung (doppelt abgesichert)**

1. Der Prompt schreibt den Marker jetzt an jedem Tag. Damit „sieht“ die KI den geplanten Load korrekt und kann ihre Empfehlung daran ausrichten.
2. Nach der Antwort setzt `ckSup_fixOriginalLoads_` den Original-Load fest aus der timeline (`coachE_ESS_day`). Damit stimmen `AI_REPORT_PLAN` und `AI_FUTURE_STATUS` immer, egal was die KI schreibt.

---

## Änderung 1 – `createMasterGeminiPrompt` (eine Zeile)

**Streiche**

```javascript
    let loadInstruction = `Geplanter Load: ${tag.original_load_ess}`; 
```

**Ersetze durch**

```javascript
    let loadInstruction = `Geplanter Load [ORIGINAL-PLAN: ${tag.original_load_ess}]`; 
```

---

## Änderung 2 – `ckSup_finish_` (aus dem Wartungs-Patch)

**Streiche**

```javascript
  // AI_REPORT_PLAN wieder aktuell halten (wird von Kalender, „KI-Plan anwenden“ und Plan-Daten gelesen)
  try { ckSup_writeReportPlan_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] AI_REPORT_PLAN nicht geschrieben: ' + e.message); }
```

**Ersetze durch**

```javascript
  // Original-Load nicht der KI überlassen: fest aus der timeline übernehmen
  try { apiResponse = ckSup_fixOriginalLoads_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] Original-Load nicht korrigiert: ' + e.message); }

  // AI_REPORT_PLAN wieder aktuell halten (wird von Kalender, „KI-Plan anwenden“ und Plan-Daten gelesen)
  try { ckSup_writeReportPlan_(ss, apiResponse); } catch (e) { logToSheet('WARN', '[Supervisor] AI_REPORT_PLAN nicht geschrieben: ' + e.message); }
```

---

## Änderung 3 – neue Funktion (ans Ende der Datei)

```javascript
/** Setzt original_load_ess im KI-Plan auf den geplanten Wert aus der timeline (coachE_ESS_day). */
function ckSup_fixOriginalLoads_(ss, apiResponse) {
  let data = apiResponse;
  const wasString = (typeof data === 'string');
  if (wasString) {
    const s = data.replace(/`{3}(?:json)?/g, '').trim();
    try { data = JSON.parse(s); } catch (e) { data = JSON.parse(s.substring(s.indexOf('{'), s.lastIndexOf('}') + 1)); }
  }
  const plan = (data && Array.isArray(data.empfohlener_plan)) ? data.empfohlener_plan : [];
  if (!plan.length) return apiResponse;

  const sh = ss.getSheetByName('timeline');
  if (!sh) return apiResponse;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const dIdx = head.indexOf('date');
  const eIdx = head.indexOf('coachE_ESS_day');
  if (dIdx < 0 || eIdx < 0) return apiResponse;
  const n = sh.getLastRow() - 1;
  const dates = sh.getRange(2, dIdx + 1, n, 1).getValues();
  const ess = sh.getRange(2, eIdx + 1, n, 1).getValues();
  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const map = {};
  for (let i = 0; i < n; i++) {
    const d = dates[i][0];
    if (d instanceof Date) map[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = Number(String(ess[i][0]).replace(',', '.')) || 0;
  }
  let changed = 0;
  plan.forEach(function (day) {
    const key = String(day.datum || '').slice(0, 10);
    if (key in map && Number(day.original_load_ess) !== map[key]) { day.original_load_ess = map[key]; changed++; }
  });
  if (changed) logToSheet('INFO', '[Supervisor] Original-Load an ' + changed + ' Tag(en) aus der timeline übernommen.');
  return wasString ? JSON.stringify(data) : data;
}
```

---

## Danach

Speichern, neue Version bereitstellen, kurz Bescheid geben – ich starte einen Supervisor-Lauf und prüfe AI_REPORT_PLAN.
