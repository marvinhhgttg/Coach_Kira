# Coach_Kira — Testing (Minimal Smoke Suite)

Coach_Kira is a Google Apps Script system; most failures are runtime/integration issues.
This document defines a minimal set of **smoke tests** you can run after changes.

---

## 1) Pre-flight

- Ensure ScriptProperties contain required keys:
  - `OPENAI_API_KEY`
  - `GEMINI_API_KEY`
  - (optional) `OPENWEATHER_API_KEY`

- Ensure critical sheets exist:
  - `timeline`, `KK_TIMELINE`
  - `AI_LOG`
  - `AI_REPORT_STATUS`
  - `AI_DATA_FORECAST`, `AI_DATA_HISTORY`
  - `week_config`

---

## 2) Supervisor run

### Test S1 — manual supervisor
Run in Apps Script editor:
- `runKiraGeminiSupervisor()`

Expected:
- `AI_LOG` receives:
  - start line
  - TE-balance found line
  - “Antwort erhalten.”
  - end line
- No execution timeout

### Test S2 — async post-work
After S1:
- In Triggers / Executions:
  - `runPostWork_` runs and completes
  - The trigger cleans itself up

---

## 3) WebApp endpoints

### Test W1 — build id
GET:
- `/exec?mode=build`

Expected:
- JSON contains `build` (matches `KK_BUILD_ID`).

### Test W2 — dashboard JSON
GET:
- `/exec?mode=json`

Expected:
- Valid JSON
- Second call is cached (faster)

### Test W3 — timeline JSON
GET:
- `/exec?mode=timeline&days=14&future=14`

Expected:
- Valid JSON
- Includes history + (optional) future payload

### Test W4 — HTML pages
Open:
- `/exec` → main
- `/exec?page=plan`
- `/exec?page=charts`

Expected:
- pages load without template-eval errors

---

## 4) Telegram webhook (if used)

### Test T1 — doPost ingestion
Send a Telegram message to the bot.

Expected:
- `doPost(e)` stores `LAST_UPDATE_ID` and `LAST_TELEGRAM_DATA`
- `processTelegramBackground` trigger runs

---

## 5) Performance checks

- Avoid full-sheet reads on 50k-row sheets.
- Prefer:
  - header-only reads
  - narrow column reads
  - single-row writes

---

*Last updated: 2026-02-01*
