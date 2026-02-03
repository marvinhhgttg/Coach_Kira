# Commander Map — Coach_Kira (1 Seite)

### Zweck & Grundprinzip
**Coach_Kira = Google Apps Script + Google Sheet “Coach Kira”**  
- Repo ist primär **Backup/Versionierung** (Sync via **clasp**).  
- Betrieb: **Daten lesen → Metriken rechnen → LLM(s) → Outputs in Sheets/UI → optional Notifications**.

### Die 3 Bausteine
1) **Google Sheet (Backend)**
- **`timeline`** = Formel-/Source (bewusst: Formeln “leben” hier)
- **`KK_TIMELINE`** = Values-/Target (schnelle Reads, v.a. Today-Row)
- **`week_config`** = Wochen-Template (wird gecached)
- **`KK_CONFIG`** = key/value Ops-Config (z.B. `CALENDAR_ID`, `LAT`, `LON`)
- **`AI_*` Tabs** = Reports/Data/Logs (Outputs & Monitoring)
- **`LOOKER_DATA`** = BI-Exportfläche

2) **Apps Script (Orchestrator)**
- Main: **`KiraGeminiSupervisor.js`**
- Nebenläufer: `ActivityReview.js`, `FormSubmitHandler.js`
- UI: `WebApp_V2.html`, `PlanApp.html`, `charts.html`, `ChatApp.html`, `Tactical_Log.html`, …

3) **Externe APIs**
- LLM: **OpenAI**, **Gemini**
- Wetter: OpenWeather/Open‑Meteo
- Optional: Telegram / Strava / Calendar/Drive Advanced Services

---

### Entry Points (wie es “läuft”)
**Main Run:**
- `runKiraGeminiSupervisor()` = täglicher Supervisor

**Async Post-Work (Performance-Pattern):**
- `schedulePostWork_()` → Trigger ~15s → `runPostWork_()` (Heavy Exports, dann Trigger Cleanup)

**WebApp:**
- `doGet(e)` = HTML/JSON routing
- `doPost(e)` = primär Telegram Webhook Ingestion (+ Background trigger)

**Stabile Annahme im Code:**
- Today wird über **`is_today`** in timeline/KK_TIMELINE gefunden.

---

### WebApp Routen (Frontends)
Basis: `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`

**Pages (`?page=` → HTML Template / Repo-Datei)**
- `page=plan` → `PlanApp.html`
- `page=planfb` → `PlanAppFB.html`
- `page=charts` → `charts.html`
- `page=chat` → `ChatApp.html`
- `page=log` → `Tactical_Log.html`
- `page=calc` → `CalculatorApp.html`
- `page=rtp` → `RTP_Smoothstep_Simulator.html`
- `page=prime` → `PrimeRangeFinder.html`
- (kein `page` param) → `WebApp_V2.html`

**JSON / API**
- `?mode=build` → Build-ID (`KK_BUILD_ID`)
- `?mode=json` → Dashboard JSON (cached + lock)
- `?mode=timeline&days=N&future=N` → Timeline Payload (cached + lock)
- `?format=json` → Widget API (`getPlanForWidget()`)

---

### Datenkontrakt (Sheet “API” – die heilige Zeile)
Die Header-Zeile ist die API. Hauptfelder:
- `date`, `is_today`
- `Sport_x`/`sport_x`, `Zone`/`zone`/`coach_zone`
- `coachE_ESS_day` / `coache_ess_day`
- `Target_Aerobic_TE`, `Target_Anaerobic_TE`
- Recovery: `sleep_hours`, `sleep_score_0_100`, `rhr_bpm`, `hrv_status`
- Observed: `load_fb_day`, `fbATL_obs`, `fbCTL_obs`, `fbACWR_obs`
- Forecast: `coachE_*_forecast`

**Stolperfallen**
- Deutsche Dezimalzahlen (`"7,5"`)
- Prozent als String (`"16,9%"`) vs Ratio (`0.169`)
- Aliases/Misch-Casing → siehe `COLUMN_ALIASES.md`

---

### Config/Secrets — wo was hingehört
**ScriptProperties (Secrets):**
- `OPENAI_API_KEY`, `GEMINI_API_KEY`, optional `OPENWEATHER_API_KEY`
- (Empfohlen) Telegram Token/ChatId **hier**, nicht im Code

**DocumentProperties / Cache:**
- PlanApp Snapshot (`PLANAPP_SNAPSHOT_V1`) + Caches (Faktoren, week_config)

**Sheet (`KK_CONFIG`):**
- key/value: `CALENDAR_ID`, `LAT`, `LON`, toggles (aber keine Secrets)

---

### Ops / Debug (wo zuerst schauen)
1) Apps Script → **Executions** (Stacktraces)
2) Sheet → **`AI_LOG`** (präfixte Logs)
3) Trigger-Liste (Duplikate / hängen gebliebene Trigger)

**Smoke Suite nach Changes (`TESTING.md`)**
- `runKiraGeminiSupervisor()`
- prüfen, ob `runPostWork_` läuft + sich selbst entfernt
- `?mode=build`, `?mode=json`, `?mode=timeline`
- Pages laden: `/exec`, `page=plan`, `page=charts`

---

### Don’ts (damit nichts heimlich kaputtgeht)
- **Keine Header-Renames** ohne Repo-Search + Alias-Plan.
- **Keine Full-Sheet Reads** (`getDataRange()` auf 50k rows) in kritischen Pfaden.
- **Keine Secrets ins Repo**.
- WebApp ist aktuell **ANYONE_ANONYMOUS** → State-Mutating Endpunkte müssen guarded sein.

---

## Top-10 Stellen im Code (Ankerpunkte)

1) `KiraGeminiSupervisor.js::runKiraGeminiSupervisor()`
- Hauptlauf (Supervisor): lädt Daten aus `KK_TIMELINE`/`timeline`, berechnet Scores, baut Prompt, ruft LLM, schreibt Reports.
- Schreibt u.a. `AI_REPORT_STATUS`, `AI_REPORT_PLAN` (indirekt über `writeOutputToSheets(...)`), Log/Heartbeat.

2) `KiraGeminiSupervisor.js::schedulePostWork_()` + `runPostWork_()`
- Async Heavy Work via Trigger (~15s): `updateHistoryWithRIS()` und `exportLookerChartsData()`.
- Wichtig für Stabilität/Timeouts: Trigger wird vorher dedupliziert und danach gelöscht.

3) `KiraGeminiSupervisor.js::doGet(e)`
- Routing + Caching/Locking:
  - `?mode=json` → `getDashboardDataAsStringV76()` (Cache key `WEBAPP_FULL_PAYLOAD`)
  - `?mode=timeline` → `getTimelinePayload(days,future)` (Cache key `TIMELINE_PAYLOAD_*`)
  - `?format=json` → `getPlanForWidget()`
  - `?page=...` → HTML Templates (`WebApp_V2`, `PlanApp`, `charts`, `ChatApp`, `Tactical_Log`, …)

4) `KiraGeminiSupervisor.js::doPost(e)` + `processTelegramBackground()`
- Telegram webhook ingestion (dedupe `LAST_UPDATE_ID`, payload in `LAST_TELEGRAM_DATA`, Trigger `processTelegramBackground`).
- **Achtung:** im Code sind `TELEGRAM_TOKEN`/`MY_CHAT_ID` hardcoded (Repo public → rotieren/auslagern).

5) `KiraGeminiSupervisor.js::callOpenAI(promptText)` + `callGeminiAPI(promptText)` + `callGeminiStructured(promptText)`
- Externe AI Calls. Gemini-Key kommt aus ScriptProperties (`GEMINI_API_KEY`).
- Diese Funktionen sind die „Kosten/RateLimit“-Hotspots.

6) `KiraGeminiSupervisor.js::getDashboardDataAsStringV76()`
- Baut das Dashboard JSON:
  - Liest Timeline per **Full `getDataRange()`** (KK_TIMELINE/timeline) → potenzieller Performance-Hotspot.
  - Baut `planData`, `nerdStats`, zieht Reviews aus `AI_ACTIVITY_REVIEWS`, KI-Plan Map aus `AI_REPORT_PLAN`.
  - Enthält mehrere Firewalls/Hotfixes (z.B. ACWR Firewall).

7) `KiraGeminiSupervisor.js::getTimelinePayload(days, futureDays)`
- Liefert Timeline als JSON (inkl. Default `futureDays=14`).
- Enthält `missingHeaders`/requiredGroups Checks für Charts.

8) `KiraGeminiSupervisor.js::copyTimelineData()`
- Sync von `timeline` (Formeln) → `KK_TIMELINE` (Values) mit Date-Mapping.
- Kritisch, weil viele Reads/Charts darauf bauen.

9) `KiraGeminiSupervisor.js::getWeekConfig()`
- Lädt Wochen-Template (`week_config`) + Caching.
- Bricht oft indirekt, wenn Struktur/Range verändert wird.

10) `KiraGeminiSupervisor.js::askKira(frage)` + `getChatHistoryForUI(limit)`
- Backend für `ChatApp.html` (UI ruft `getChatHistoryForUI(50)` und `askKira(txt)` via `google.script.run`).
- Zentral für „Secure Channel“ Chat.

**Nebenanker (wichtig, aber außerhalb Top-10):**
- `ActivityReview.js::callGeminiForReview(...)` + Review-Write (schreibt aktuell immer Zeile 2)
- `FormSubmitHandler.js` (Form-Input → Timeline/Reviews Pfade, `findRowByIsToday` etc.)
