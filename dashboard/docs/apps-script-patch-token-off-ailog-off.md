# Apps-Script-Patch: Token-Schutz entfernen + AI_LOG abschalten

Beide Änderungen liegen in `KiraGeminiSupervisor.gs`. Danach **eine** neue Version bereitstellen.

---

## 1. Token-Schutz entfernen

Im Block, der mit diesem Kommentar beginnt:

```javascript
// ============================================================
// Coach Kira Dashboard API – Minimalendpunkte für pptx.app
// Stand: MVP
// ============================================================

function ckDash_json_(obj) { … }

function ckDash_now_() { … }

function ckDash_requireToken_(p) {      // ← diese Funktion
  const expected = PropertiesService
    .getScriptProperties()
    .getProperty('DASHBOARD_ACTION_TOKEN');
  …
}
```

die **dritte** Funktion `ckDash_requireToken_` komplett (von `function ckDash_requireToken_(p) {` bis zur zugehörigen schließenden `}` vor `function ckDash_getSheet_`) ersetzen durch:

```javascript
/**
 * (2026-09-28) Token-Schutz auf Wunsch deaktiviert.
 * Wieder einschalten: Prüfung gegen ScriptProperty DASHBOARD_ACTION_TOKEN zurückholen (siehe unten).
 */
function ckDash_requireToken_(p) {
  return { ok: true };
}
```

---

## 2. AI_LOG abschalten (Blatt bleibt erhalten)

Ganz oben in der Datei steht:

```javascript
const LOG_SHEET_NAME = 'AI_LOG';
```

**Direkt darunter** eine Zeile einfügen:

```javascript
const AI_LOG_ENABLED = false; // (2026-09-28) Schreiben ins Blatt AI_LOG abgeschaltet; Blatt bleibt bestehen
```

Dann in `function logToSheet(level, message) {` **als erste Zeile nach der öffnenden Klammer** einfügen:

```javascript
  if (typeof AI_LOG_ENABLED !== 'undefined' && !AI_LOG_ENABLED) {
    console.log(`[${level}] ${typeof message === 'object' ? JSON.stringify(message) : message}`);
    return;
  }
```

Die Funktion beginnt danach so:

```javascript
function logToSheet(level, message) {
  if (typeof AI_LOG_ENABLED !== 'undefined' && !AI_LOG_ENABLED) {
    console.log(`[${level}] ${typeof message === 'object' ? JSON.stringify(message) : message}`);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    …
```

**Wirkung:**
- Ins Blatt `AI_LOG` wird nichts mehr geschrieben. Das Blatt und sein bisheriger Inhalt bleiben stehen.
- Die Meldungen gehen weiterhin in das Apps-Script-Protokoll (Editor → „Ausführungen“). Fehlersuche ist also weiter möglich.
- Nebeneffekt: Jeder Lauf wird etwas schneller, weil die rund 100 Blatt-Schreibzugriffe pro Nacharbeit entfallen.
- Wieder einschalten: `AI_LOG_ENABLED = true`.

---

## 3. Speichern und neu bereitstellen

**Bereitstellen → Bereitstellung verwalten → Bleistift → Version: Neue Version → Bereitstellen**, dann „fertig“ sagen.

---

## Zum Wiederherstellen des Token-Schutzes

```javascript
function ckDash_requireToken_(p) {
  const expected = PropertiesService.getScriptProperties().getProperty('DASHBOARD_ACTION_TOKEN');
  if (!expected) return { ok: false, error: 'missing_dashboard_token_config' };
  if (!p || p.token !== expected) return { ok: false, error: 'unauthorized' };
  return { ok: true };
}
```
