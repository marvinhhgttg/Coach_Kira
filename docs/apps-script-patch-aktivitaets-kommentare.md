# Apps-Script-Patch: Aktivitäts-Kommentare lokal

Stand: 18.10.2026 · betrifft das Blatt `AI_ACTIVITY_REVIEWS`

**Ausgangslage**

Der letzte Eintrag in `AI_ACTIVITY_REVIEWS` stammt vom 07.03.2026. Seitdem kommen keine Gemini-Kommentare mehr an; die Funktion `generateActivityReview` liegt außerdem in einer anderen Skriptdatei als `KiraGeminiSupervisor`. Es muss also nichts abgeschaltet werden.

**Was der Patch macht (rein ergänzend)**

- `saveActivityReview` (POST): Der Mac-Dienst schreibt einen lokal erzeugten Kommentar als neue Zeile ins Blatt.
  - Format wie bisher: Spalten A–J, `Datum` als echtes Datum (der Supervisor liest die Zeilen weiter mit `instanceof Date`).
  - Neu: Spalte K `Quelle`, z. B. `{"engine":"local","model":"gemma4:e4b-mlx"}`.
- `activityReviews` (GET): liefert alle Kommentare für das Log im Cockpit (Datum, Text, Quelle). Pro Tag gilt der jeweils neueste Eintrag.

Bestehende Funktionen bleiben unverändert.

---

## Änderung 1 – `doGet`

**Streiche**

```javascript
  if (p.mode === 'briefingSources') return ckDash_briefingSources_();
```

**Ersetze durch**

```javascript
  if (p.mode === 'briefingSources') return ckDash_briefingSources_();
  if (p.mode === 'activityReviews') return ckDash_activityReviews_(p);
```

---

## Änderung 2 – `doPost`

**Streiche**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveBriefing') || /"kind"\s*:\s*"saveBriefing"/.test(ckBody)) {
    return ckDash_saveBriefing_(e);
  }
```

**Ersetze durch**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveBriefing') || /"kind"\s*:\s*"saveBriefing"/.test(ckBody)) {
    return ckDash_saveBriefing_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveActivityReview') || /"kind"\s*:\s*"saveActivityReview"/.test(ckBody)) {
    return ckDash_saveActivityReview_(e);
  }
```

---

## Änderung 3 – neue Funktionen (ans Ende der Datei)

```javascript
// ============================================================================
// Aktivitäts-Kommentare lokal (Mac-Dienst) – speichern und lesen
// ============================================================================
function ckDash_activitySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('AI_ACTIVITY_REVIEWS');
  if (!sh) {
    sh = ss.insertSheet('AI_ACTIVITY_REVIEWS');
    sh.getRange(1, 1, 1, 10).setValues([[
      'Datum', 'Load (IST)', 'Load (SOLL)', 'ATL (IST)', 'ATL (SOLL)',
      'CTL (IST)', 'CTL (SOLL)', 'ACWR (IST)', 'ACWR (SOLL)', 'KI Review Text'
    ]]);
  }
  if (String(sh.getRange(1, 11).getValue() || '') === '') {
    sh.getRange(1, 11).setValue('Quelle');
  }
  return sh;
}

function ckDash_saveActivityReview_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const date = String(b.date || '');
    const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) throw new Error('date fehlt oder ist nicht YYYY-MM-DD.');
    const text = String(b.text || '').slice(0, 4000).trim();
    if (!text) throw new Error('Kein Kommentar-Text übermittelt.');
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    const ist = b.ist || {};
    const soll = b.soll || {};
    const val = function (v) { return (v === null || v === undefined || v === '') ? 'N/A' : v; };
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);

    const sh = ckDash_activitySheet_();
    sh.appendRow([
      d,
      val(ist.load), val(soll.load),
      val(ist.atl), val(soll.atl),
      val(ist.ctl), val(soll.ctl),
      val(ist.acwr), val(soll.acwr),
      text,
      JSON.stringify({ engine: 'local', model: model, savedAt: ckDash_now_() })
    ]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd.MM.yyyy');
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), date: date });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_activityReviews_(p) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AI_ACTIVITY_REVIEWS');
    if (!sh || sh.getLastRow() < 2) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), reviews: [] });
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const lastCol = Math.max(10, Math.min(11, sh.getLastColumn()));
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const byDate = {};
    rows.forEach(function (r) {
      if (!(r[0] instanceof Date)) return;
      const date = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
      let text = String(r[9] || '').trim();
      if (!text) return;
      if (text.charAt(0) === '{') {
        try { const j = JSON.parse(text); text = String(j.coach_feedback || j.review || j.text || text); } catch (x) {}
      }
      let engine = 'cloud';
      let model = 'gemini';
      try {
        const q = JSON.parse(String(r[10] || '{}'));
        if (q && q.engine === 'local') engine = 'local';
        if (q && q.model) model = String(q.model);
      } catch (x) {}
      byDate[date] = {
        date: date, text: text, engine: engine, model: model,
        loadIst: r[1], loadSoll: r[2], acwrIst: r[7]
      };
    });
    const limit = Math.max(1, Math.min(2000, parseInt((p && p.limit) || '500', 10) || 500));
    const reviews = Object.keys(byDate).sort().slice(-limit).map(function (k) { return byDate[k]; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), reviews: reviews });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}
```

---

## Danach

1. Speichern.
2. Bereitstellen → Bereitstellungen verwalten → Stift → Version „Neue Version“ → Bereitstellen.
3. Kurz Bescheid geben – ich teste `?mode=activityReviews` und den ersten lokalen Kommentar.
