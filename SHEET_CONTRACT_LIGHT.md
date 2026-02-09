# Coach_Kira — Sheet Contract (Light) + Glossary

This document is a **practical reference** for developing Coach_Kira.
It describes:

- **Which spreadsheet tabs exist** and what they are used for
- **Which columns are expected** ("contract") for the core tabs
- **How values are formatted** (common pitfalls: German decimals, percent strings)
- A **glossary** of the domain terms used throughout code + sheets

> Spreadsheet: **Coach Kira** (`1ydTDKQ1VmYL-H0U3E_xe1tdYJlEXfydYMYpKZWW5xsQ`)
>
> Source code: `KiraGeminiSupervisor.js` (main), plus `ActivityReview.js`, `FormSubmitHandler.js`

---

## 0) Conventions & pitfalls (read first)

### 0.1 Header row is the API
Most code locates columns via `headers.indexOf('<name>')`. That means:

- Column names are **case-sensitive** in many places.
- Typos or renames silently break logic.
- Some functions use **different casing** (e.g. `Sport_x` vs `sport_x`, `coachE_ESS_day` vs `coache_ess_day`).

**Rule:** Treat header names like an API. If you rename headers, search the repo for them.

### 0.2 German decimals
The sheet often uses German number formatting.

- Many values are stored as strings like `"7,5"`.
- Some code uses helpers like `parseGermanFloat_()`.

**Rule:** If you add numeric columns that the script reads, keep formatting consistent.

### 0.3 Percent values
Some values are stored like `"16,9%"`, but downstream code often expects **0.169**.

**Rule:** Always define whether a column stores:
- a ratio in `[0..1]` or
- a percent string with `%`.

### 0.4 "timeline" vs "KK_TIMELINE" (two timelines)
- `timeline` is the **formula/source** sheet.
- `KK_TIMELINE` is the **values/target** sheet.

Some functions intentionally write into `timeline` (so formulas keep rolling).
Some reads prefer `KK_TIMELINE` for performance.

---

## 1) Sheet inventory (observed)

Extracted via Sheets API (`gog sheets metadata`):

- `timeline` (50k rows) — source sheet (formulas)
- `KK_TIMELINE` (50k rows) — values sheet (synced)
- `week_config` — weekly template
- `KK_CONFIG` — operational config
- `config` — numeric constants (aux)
- `KK_LOAD_FACTORS`, `KK_ELEV_FACTORS` — factor tables
- `AI_DATA_FORECAST`, `AI_DATA_HISTORY` — data surfaces
- `AI_REPORT_STATUS`, `AI_REPORT_PLAN`, `AI_REPORT_HISTORY` — outputs
- `AI_FUTURE_STATUS` — future/outlook
- `AI_LOG` — central run log
- `AI_CHAT_LOG` — chat log
- `AI_HEARTBEAT` — run state/heartbeats
- `LOOKER_DATA` — BI export
- `AI_PLAN_BACKUP` — plan backup
- `AI_ACTIVITY_REVIEWS` — activity review outputs
- `phys_baseline`, `KK_WISSEN`, `KK_NOTES` — reference/notes
- plus UI/helper tabs (`Site_Embed`, `Radar_VIEW`, …)

---

## 2) Core contracts

### 2.1 `timeline` / `KK_TIMELINE` (core daily time series)

These sheets are the backbone. Many features depend on them.

#### Required columns (high priority)
These are directly used in multiple flows:

- `date` — date of the row
- `is_today` — marker identifying the current "today" row (typically `1`)
- `Sport_x` (sometimes `sport_x`) — activity type
- `Zone` (sometimes `zone` or `coach_zone`) — intensity/zone label
- `coachE_ESS_day` (sometimes `coache_ess_day`) — planned load/ESS for the day
- `Target_Aerobic_TE` (sometimes `target_aerobic_te`) — planned aerobic TE
- `Target_Anaerobic_TE` (sometimes `target_anaerobic_te`) — planned anaerobic TE
- `fix` — a fix/override flag (used in planning flows)

#### Commonly used derived columns
Used for scoring, risk checks, or UI summaries:

- `load_fb_day` — observed load
- `fKEI` — clamp factor applied to KEI (formula-based)
- `Load*` / `load_star` — KEI-adjusted load (formula-based)
- `CTL*` / `ctl_star` — CTL computed from `Load*` (formula-based)
- `ctl_gap` — optional: `CTL − CTL*` delta (formula-based)
- `fbATL_obs`, `fbCTL_obs`, `fbACWR_obs` — observed training metrics
- `coachE_ATL_forecast`, `coachE_CTL_forecast`, `coachE_ACWR_forecast` — forecast metrics
- `Monotony7`, `Strain7` — 7d monotony/strain
- `coachE_ATL_morning`, `coachE_CTL_morning` — morning metrics
- `Week_Phase` (sometimes `week_phase`) — phase label
- `Weekday` — weekday label (if present)

#### Recovery / readiness related columns
Used in readiness/training readiness logic:

- `sleep_hours`
- `sleep_score_0_100`
- `rhr_bpm`
- `hrv_status`
- `hrv_threshholds`

#### Nutrition / energy columns
Used in prompt/context:

- `deficit`
- `carb_g`, `protein_g`, `fat_g`

#### TE balance
- `te_balance` or `te balance` (header matching is fuzzy: `includes('te_balance')` or `includes('te balance')`)

Contract definition:
- Allowed formats: number, `"16,9"`, `"16,9%"` (code tries to normalize)
- Interpretation: should represent TE balance; code often converts percent to ratio.

#### Semantics
- `is_today` must have **exactly one active row** (or last occurrence wins in some scans).
- `date` should be parseable (Date object or ISO-like string).

#### KEI-adjusted load formulas (in `timeline`)
- `fKEI = CLAMP(0.85, 1 + 0.05 * KEI, 1.15)`
- `Load* = Load * fKEI`
- `CTL*` uses the same 42‑day smoothing as `CTL`, but on `Load*`
- Optional: `ctl_gap = CTL − CTL*`

#### Write patterns
- Planning updates often write into `timeline`.
- Sync functions copy (at least) the today row into `KK_TIMELINE`.

---

### 2.2 `week_config` (weekly template)

Role:
- Stores a weekly plan template.
- `getWeekConfig()` reads it and caches it.

Contract:
- The code currently reads `getDataRange().getValues()` and converts to a CSV-like string.
- Header/structure is treated as **template data** more than strict schema.

Guidelines:
- Avoid very large ranges (keep it compact).
- If you change structure, update the prompt builder accordingly.

---

### 2.3 `KK_CONFIG` (operational configuration)

Role:
- User-editable configuration for integrations and runtime switches.

Observed usages:
- Calendar sync expects a `CALENDAR_ID` entry.
- Weather expects `LAT`/`LON` and an API key in ScriptProperties.

Contract (recommended pattern):
- Two-column table: `key | value` (or similar).
- Keys should be stable identifiers (uppercase snake case recommended).

Common keys (from code behavior/logs):
- `CALENDAR_ID`
- `LAT`, `LON`

Secrets should NOT be stored here (use ScriptProperties).

---

### 2.4 `config` (numeric constants)

Role:
- Legacy/aux constants (alphas, defaults). Example values include:
  - `sleep_hours_default`, `sleep_score_default`
  - `alpha_ATL`, `alpha_CTL`, caps, etc.

Contract:
- Typically `name | value` rows.
- Many values are stored with German decimals.

Guideline:
- Prefer `KK_CONFIG` for operational toggles/IDs.
- Prefer ScriptProperties for secrets.

---

### 2.5 `AI_LOG` (central logging)

Role:
- Append-only run log (`logToSheet(level, message)`).

Contract (recommended columns):
- Timestamp (local time)
- Level (`DEBUG|INFO|WARN|ERROR`)
- Message (string)

Guidelines:
- Keep messages prefixed (`[History]`, `[Kalender]`, `[PostWork]`, …) for grepability.

---

### 2.6 `AI_REPORT_STATUS` (daily status output)

Role:
- Human-readable (and/or UI-consumed) status report.

Observed header contract:
- `Metrik`
- `Status_Ampel`
- `Status_Wert_Num`
- `Status_Score_100`
- `Text_Info`

Guideline:
- Treat these headers as stable API.

---

### 2.7 `AI_DATA_FORECAST` / `AI_DATA_HISTORY`

Role:
- Tabular data surfaces used by supervisor + reports.

Observed:
- Supervisor uses `AI_DATA_FORECAST` as a **fallback** source for TE balance.

Guidelines:
- Keep the first row as headers.
- Avoid large recalculations in Apps Script; prefer formulas where possible.

---

### 2.8 `LOOKER_DATA`

Role:
- Export surface for Looker Studio.

Guidelines:
- Keep schema stable once Looker is connected.
- When changing schema, version the Looker data or provide a compatibility layer.

---

### 2.9 `AI_CHAT_LOG`

Role:
- Stores chat interactions.

Code notes:
- The code expects `AI_CHAT_LOG` with specific columns (often A=timestamp, B=sender, C=message).

---

### 2.10 `AI_HEARTBEAT`

Role:
- Lightweight run-state events (START/END/phase markers), keyed by a run id.

Guideline:
- Use it for timeline/debugging of long runs.

---

## 3) ScriptProperties (secrets) — contract

These are not in the sheet, but they are part of the system contract.

Required/commonly used keys:
- `OPENAI_API_KEY`
- `GEMINI_API_KEY`
- `OPENWEATHER_API_KEY`

Runtime keys written by the system:
- `HB_RUN_ID` (temporary run identifier)
- `LAST_UPDATE_ID`, `LAST_TELEGRAM_DATA` (telegram integration state)

Guideline:
- Document new ScriptProperties keys here whenever you add one.

---

## 4) Glossary

### Training / load
- **ESS**: (CoachE) daily load value used for planning (`coachE_ESS_day`).
- **ATL**: Acute Training Load (short-term load).
- **CTL**: Chronic Training Load (long-term load).
- **ACWR**: Acute:Chronic Workload Ratio.
- **Monotony7 / Strain7**: derived 7-day metrics used for risk/strain evaluation.

### Intensity / TE
- **TE**: Training Effect.
  - **Aerobic TE**: endurance/systemic training effect.
  - **Anaerobic TE**: high-intensity training effect.
- **TE Balance**: a balance metric; appears as `te_balance` column and is logged as percent.

### Readiness / health
- **RTP**: Return-to-Play logic (illness/injury recovery / protocol status).
- **Training Readiness**: readiness score derived from sleep/HRV/etc.

### Operational
- **Supervisor**: the orchestrator run (`runKiraGeminiSupervisor`).
- **PostWork**: heavy tasks run asynchronously after supervisor.
- **Forecast vs Observed**: planned values vs measured values (`*_forecast` vs `*_obs`).

---

## 5) Recommended next improvements (for maintainability)

1) **Normalize headers**
   - Pick ONE naming convention (e.g. lowercase snake_case) and build a mapping layer.
   - Today the code mixes `Sport_x` and `sport_x`, `coachE_*` and `coache_*`.

2) **Centralize column lookup**
   - Implement `getCol_(headers, [aliases...])` and use it everywhere.

3) **Introduce a validation function**
   - `validateTimelineContract_(sheet)` that checks required columns and logs a clear error.

---

*Last updated: 2026-02-01*
