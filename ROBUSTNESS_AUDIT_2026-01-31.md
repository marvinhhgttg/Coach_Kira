# Coach_Kira — Robustness Audit (2026-01-31)

Context:
- Project type: Google Sheets bound Apps Script project
- Repo acts as review/backup; runtime is Apps Script (server-side) + HtmlService UIs (client-side)
- Baseline deployment exists ("Coach Kira v794 - Baseline")

This audit focuses on **robustness**: reliability, data correctness, quotas/performance, and safe UI rendering.

---

## Executive summary

Overall the project is operational and has pragmatic safeguards (fallbacks, logging, defensive checks). Biggest remaining risk areas are:

1) **UI rendering safety** (`innerHTML` with interpolated values) — potential XSS/markup breakage from unescaped strings.
2) **Locale/number parsing consistency** — mitigated by recent PR, but still a cross-cutting concern.
3) **Error handling consistency** — some silent catches and mixed logging facilities.
4) **Performance / quotas** — several `getDataRange().getValues()` patterns across large sheets; usually fine but can spike runtime.

---

## Findings (prioritized)

### High — UI injection / broken markup risk in HtmlService
**Where:** `PlanApp.html`, `PlanAppFB.html`, `WebApp_V2.html`, `PrimeRangeFinder.html`, `RTP_Smoothstep_Simulator.html`, `Tactical_Log.html`, `ChatApp.html`, `charts.html`.

**Why it matters:** Many templates do `element.innerHTML = `...${value}...``. If `value` contains `<`, `&`, backticks, or HTML fragments, it can:
- break UI rendering
- render unintended markup
- in worst case, enable script injection (XSS) inside the HtmlService iframe context

**Typical sources:** error strings, free-text notes, server-returned strings, user inputs.

**Recommended fix:** add a small `escapeHtml()` helper and use it for all interpolated dynamic values.

---

### Medium — Mixed logging & silent catches hide problems
**Where:** `KiraGeminiSupervisor.js` and various HTML files use `console.log`, some server code uses `Logger.log`, and operational logging uses `logToSheet`.

**Why it matters:** when something intermittently fails (quota/parse edge-case), it becomes hard to trace.

**Recommended fix:**
- Prefer `logToSheet(level, msg)` for server-side operational events.
- Use `Logger.log` only for dev/diagnostics.
- Replace empty `catch {}` blocks with at least a `WARN` log in the critical path.

---

### Medium — Performance / quota hotspots
**Where:** frequent patterns of:
- `sheet.getDataRange().getValues()` on `timeline`/`KK_TIMELINE` (500x57)
- repeated reads of the same ranges within one run

**Why it matters:** Usually OK, but when combined with network calls (Gemini), chart exports, multiple write-backs and flushes, runtime can approach Apps Script limits.

**Recommended fix:**
- Load a range once, pass arrays around.
- Prefer column-limited ranges over full `getDataRange()`.
- Batch writes (`setValues` in blocks) — already done in many places.

---

### Low — Duplicated helper concepts
**Where:** multiple numeric parsing helpers exist (`parseGermanFloat`, `parseGermanFloat_`, `extractNumericValue_`, `cleanNumberFromKI`).

**Why it matters:** future changes can diverge.

**Recommended fix:** document "source of truth" helpers and gradually route callers to them.

---

## Proposed PRs

1) **PR: Robustness audit report** (this file) — documentation only.
2) **PR: HTML safety hardening** — add `escapeHtml()` and apply to dynamic `innerHTML` usages for untrusted/variable strings.
3) (Optional later) **PR: Logging consistency** — replace server-side `console.log` and add minimal warn logs to silent catches.

---

## Notes / boundaries

- No code is pushed to Google Apps Script without Marc’s explicit approval.
- Any changes should be kept small and reviewable (one purpose per PR).
