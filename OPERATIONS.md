# Coach_Kira — Operations / Runbook

This runbook documents how to operate Coach_Kira safely: deployments, triggers, rollbacks, and debugging.

---

## 1) Deployments (Apps Script)

### 1.1 WebApp deployments
The WebApp is deployed from Apps Script and typically accessed at:

- `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`

Notes:
- Apps Script has both a **HEAD/Test deployment** and **versioned deployments**.
- Running functions from the editor often uses a test/head context.

### 1.2 Versioning / rollback strategy
Use **versioned deployments** for production.

Rollback approach:
1) Apps Script → Deploy → Manage deployments
2) Select the production deployment
3) Switch to the previous version

---

## 2) Repo ↔ Apps Script sync (clasp)

This repo is primarily a backup of the Apps Script project.

### 2.1 Browser editor → GitHub (backup)
```bash
clasp pull
git add -A
git commit -m "Backup: ..."
git push
```

### 2.2 GitHub → Browser editor (push)
```bash
git pull
clasp push
```

⚠️ `clasp push` overwrites server code. Ensure browser edits are saved.

---

## 3) Triggers (important)

Coach_Kira relies on time-based triggers to reduce main-run execution time and to process webhooks.

### 3.1 Known trigger handlers (from code)
- `runKiraGeminiSupervisor` (scheduled supervisor)
- `runPostWork_` (async post-work; should delete its own trigger)
- `processTelegramBackground` (Telegram webhook processing)
- `runHistoricalAnalysis`
- `generateActivityReview`

### 3.2 Trigger hygiene
- Avoid duplicates: code contains helpers that delete triggers for a function name.
- If a function is slow and scheduled frequently, ensure triggers don’t pile up.

---

## 4) WebApp behavior & caching

WebApp routing is documented in `WEBAPP_ROUTES.md`.

Caching/locking:
- JSON endpoints use `CacheService.getUserCache()` and `LockService.getScriptLock()`.

If you see timeouts:
- Check for lock contention
- Reduce sheet reads (avoid full `getDataRange()` on 50k sheets)

---

## 5) Debugging checklist

### 5.1 Where to look first
1) Apps Script → **Executions** (failed runs, stack traces)
2) Sheet `AI_LOG` (application log with prefixes)
3) Triggers list (duplicates / stuck triggers)

### 5.2 Common failure modes
- Timeouts / quota limits due to full-sheet reads (timeline is ~50k rows)
- Missing columns after sheet edits (header changed)
- Missing ScriptProperties keys (OpenAI/Gemini/Weather)
- Trigger duplication

### 5.3 “Safe mode” tactics
- Temporarily disable heavy post-processing by short-circuiting `schedulePostWork_()`
- Lower data windows (history days) in endpoints

---

## 6) Security notes

`appsscript.json` currently configures WebApp access as `ANYONE_ANONYMOUS`.

Operational implications:
- Ensure any state-mutating endpoint is guarded (token, allowlist, etc.).
- Never hardcode secrets in repo.

---

*Last updated: 2026-02-01*
