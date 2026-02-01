# Coach_Kira — Secrets & Config Baseline

This document defines the **configuration contract** of the Coach_Kira system.
It’s meant to answer quickly:

- Which settings live in **ScriptProperties** vs **DocumentProperties** vs **Sheets**
- Which keys are **required**
- Which keys are **safe to commit** (almost none)

> Do **not** store secret values in GitHub.

---

## 1) Where config lives

### A) ScriptProperties (secrets + runtime state)
Use `PropertiesService.getScriptProperties()`.

**Use for:**
- API keys / tokens
- Integration credentials
- Small runtime state (dedupe ids, run ids)

**Never commit values.**

### B) DocumentProperties (caches + per-spreadsheet state)
Use `PropertiesService.getDocumentProperties()`.

**Use for:**
- Caches (factors, week_config)
- UI snapshots tied to this spreadsheet (PlanApp snapshot)

### C) UserProperties (user-specific settings)
Used rarely. Consider only if per-user is required.

### D) Sheets (user-editable parameters)
In the spreadsheet:
- `KK_CONFIG` (operational config)
- `config` (numeric constants)
- factor tables (`KK_LOAD_FACTORS`, `KK_ELEV_FACTORS`)
- templates (`week_config`)

---

## 2) ScriptProperties keys (baseline)

### 2.1 AI providers
- `OPENAI_API_KEY`
- `GEMINI_API_KEY`

### 2.2 Weather
- `OPENWEATHER_API_KEY`

### 2.3 Telegram integration (NOTE: currently hardcoded)
In `KiraGeminiSupervisor.js` there are hardcoded constants:
- `TELEGRAM_TOKEN`
- `MY_CHAT_ID`

**Recommendation:** move these into ScriptProperties:
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

### 2.4 Runtime state written by code
- `HB_RUN_ID` (run identifier written by heartbeat helper)
- `LAST_UPDATE_ID` (Telegram webhook dedupe)
- `LAST_TELEGRAM_DATA` (stored payload for background processing)

---

## 3) DocumentProperties keys (baseline)

### 3.1 PlanApp snapshot
- `PLANAPP_SNAPSHOT_V1` (constant `PLANAPP_SNAPSHOT_KEY`)
  - JSON blob
  - intended to be valid only for the current calendar day

### 3.2 Caches
The code uses ScriptCache for factors/week_config, but some paths persist to DocumentProperties.
Expect keys like:
- load factor cache keys
- elevation factor cache keys
- week_config cache keys

**Guideline:** whenever you add a cache key, document it here and specify TTL/invalidations.

---

## 4) Sheet-based config keys (baseline)

### 4.1 `KK_CONFIG` (recommended contract)
Store as a 2-column key/value table.

Common expected keys (from logs/behavior):
- `CALENDAR_ID` (Calendar integration)
- `LAT`, `LON` (Weather)

If you add a key, add it here with:
- description
- allowed formats
- example

### 4.2 `config`
Numeric constants, often with German decimals.
Examples found:
- `sleep_hours_default`
- `sleep_score_default`
- `alpha_ATL`, `alpha_CTL`, caps, etc.

---

## 5) Build/version markers

- `KK_BUILD_ID` (string constant in code)
- WebApp: `/exec?mode=build` returns `{ ok: true, build: KK_BUILD_ID }`

---

## 6) Quick setup checklist

1) ScriptProperties
   - set `OPENAI_API_KEY`
   - set `GEMINI_API_KEY`
   - (optional) set `OPENWEATHER_API_KEY`
   - (recommended) set Telegram token/chat id if moved out of code

2) Sheet tabs exist
   - `timeline`, `KK_TIMELINE`
   - `AI_LOG`, `AI_REPORT_STATUS`, `AI_DATA_HISTORY`, `AI_DATA_FORECAST`
   - `week_config`, `KK_CONFIG`

3) `KK_CONFIG`
   - set `CALENDAR_ID` if calendar sync is used
   - set `LAT/LON` if weather is used

---

*Last updated: 2026-02-01*
