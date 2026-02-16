# Coach_Kira — WebApp Routes (doGet/doPost)

This doc maps the **public WebApp URLs** of the Apps Script deployment to the
**HTML templates** and **server-side handlers** in this repository.

Base URL (prod deployment example):

- `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`

In this project, routing is implemented in **`KiraGeminiSupervisor.js` → `doGet(e)`**.

---

## 1) HTML pages (`?page=...`)

The WebApp serves HTML via `HtmlService.createTemplateFromFile('<NAME>')`.

| URL | Template | Repo file |
|---|---|---|
| `/exec` *(no page param)* | `WebApp_V2` | `WebApp_V2.html` |
| `/exec?page=plan` | `PlanApp` | `PlanApp.html` |
| `/exec?page=planfb` | `PlanAppFB` | `PlanAppFB.html` |
| `/exec?page=charts` | `charts` | `charts.html` |
| `/exec?page=charts_v2` | `charts_V2` | `charts_V2.html` |
| `/exec?page=chat` | `ChatApp` | `ChatApp.html` |
| `/exec?page=calc` | `CalculatorApp` | `CalculatorApp.html` |
| `/exec?page=log` | `Tactical_Log` | `Tactical_Log.html` |
| `/exec?page=rtp` | `RTP_Smoothstep_Simulator` | `RTP_Smoothstep_Simulator.html` |
| `/exec?page=prime` | `PrimeRangeFinder` | `PrimeRangeFinder.html` |
| `/exec?page=lab` | `LoadLab` | *(not present in repo; may exist only in GAS or was renamed)* |

### Shared template variables
All HTML templates receive:

- `template.pubUrl` — the WebApp base URL used by frontends to call back.
- `template._debug_build` — current build id (`KK_BUILD_ID`).

Other templates may set additional debug fields (e.g. `charts`).

---

## 2) JSON / API modes (`?mode=...` / `?format=...`)

These endpoints return JSON (or text) via `ContentService`.

| URL | Purpose | Handler |
|---|---|---|
| `/exec?mode=build` | Returns build id for debugging | `doGet(e)` (build branch) |
| `/exec?format=json` | “Plan Widget API” (7-day preview for Scriptable, etc.) | `getPlanForWidget()` |
| `/exec?mode=json` | Dashboard JSON payload (cached; lock-protected) | `getDashboardDataAsStringV76()` |
| `/exec?mode=timeline` | Timeline JSON payload (cached; supports history + future) | `getTimelinePayload(days,futureDays)` |

### Timeline API details

`/exec?mode=timeline` supports:

- `days=<N>` — history window (e.g. 14, 90). If omitted/empty → ALL history.
- `future=<N>` — forecast window (default 14). If set to 0/empty → no forecast.

Cache key is derived from both `days` and `future`.

---

## 3) WebApp POST (`doPost(e)`)

`doPost(e)` is currently used primarily for **Telegram webhook ingestion**.

### Telegram webhook
Expected payload:
- JSON body compatible with Telegram updates.

Behavior:
- De-duplicates by `update_id` (stored in ScriptProperties).
- Stores last update data (`LAST_TELEGRAM_DATA`).
- Schedules background processing via a time-based trigger:
  - `processTelegramBackground`

---

## 4) Operational notes

### Caching and locking
- `?mode=json` and `?mode=timeline` use:
  - `CacheService.getUserCache()`
  - `LockService.getScriptLock()`

This reduces quota and prevents concurrent expensive reads.

### X-Frame-Options
The HTML output sets:
- `ALLOWALL`

This is intentional to support embedding.

### Security
The Apps Script WebApp is configured in `appsscript.json` with:

- `access: ANYONE_ANONYMOUS`

Ensure that any route that changes state is properly guarded (tokens, internal-only,
rate limiting, etc.).

---

## 5) Where to edit routing

- `KiraGeminiSupervisor.js` → `function doGet(e)`
- Add new pages by adding a `page === '...'` branch mapping to a new `*.html` file.

---

*Last updated: 2026-02-01*
