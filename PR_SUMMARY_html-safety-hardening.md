# PR Summary — HTML Safety Hardening

Branch: `pr/html-safety-hardening`

Goal: Reduce risk of broken markup / HTML injection in HtmlService UIs by avoiding unescaped dynamic values in `innerHTML` template strings.

## Summary of changes (high level)

### WebApp_V2.html
- Added client-side escaping helpers:
  - `escapeHtml(input)`
  - `escapeAttr(input)`
- Escaped dynamic values inserted via `innerHTML` in multiple render paths, including:
  - Plan table: day tag/date, sport, zone, load
  - Activity review table: `rev.datum`, `rev.text`
  - Suggestions UI: weekday/date and input `value="..."` attributes
  - Score group header: group name

### PlanApp.html
- Escaped recommendation values rendered into HTML (`escapeHtml(val)`) in the strategic briefing UI.

### PlanAppFB.html
- Hardened gate info rendering:
  - switched from `innerHTML` to `textContent`
  - avoids injecting raw strings into HTML

### ChatApp.html
- Hardened history-load error fallback:
  - switched from `innerHTML` to `textContent` for the fallback message

### PrimeRangeFinder.html
- Added a local `escapeHtml()` helper.
- Escaped dynamic values in KPI pills and table cells that may contain flags / derived text.

### RTP_Smoothstep_Simulator.html
- Added a local `escapeHtml()` helper.
- Escaped dynamic values in KPI pills and `dayType` table cell.

### Tactical_Log.html
- Added a local `escapeHtml()` helper.
- Escaped dynamic values rendered into the log table (date/phase/sport/load/TE/CTL trend/RHR/SG display).

## Notes
- Static HTML snippets (e.g. placeholders/spinner icons) were kept as-is.
- `.DS_Store` is present locally but is not intended to be committed.
- No changes are pushed to Google Apps Script automatically; apply only with Marc’s explicit approval.
