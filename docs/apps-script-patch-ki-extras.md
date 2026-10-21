# Apps-Script-Patch: KI-Extras

Stand: 21.10.2026 · für die sechs neuen KI-Funktionen

## Was sich ändert

**Tabelle**

- Neues Blatt `AI_KIRA_EXTRAS`. Es wird beim ersten Speichern automatisch angelegt.
- Spalten: Datum, Typ, Text, Daten (JSON), Quelle, Gespeichert.
- Ein Eintrag je Typ und Tag; ein erneutes Speichern ersetzt ihn.
- Typen: `treffer_tag`, `treffer_woche`, `selbsttest`, `plausi`, `ernaehrung`, `alltag`, `workout`.
- Sonst werden keine Blätter oder Spalten verändert.

**Neue Routen der Web-App**

| Route | Zweck | genutzt von |
|---|---|---|
| `POST kind=saveExtra` | Ergebnis speichern | Mac-Dienst, Cockpit |
| `GET mode=extras&typ=…&n=…` | Ergebnisse lesen | Cockpit, Mac-Dienst |
| `GET mode=calendarDays&days=14` | Kalender-Belegung je Tag | Alltag-Check |
| `GET mode=physBaseline` | Gewicht, Max-HF, Laktatschwelle, FTP aus `phys_baseline` | Workout-Generator, Ernährung |

**Kalender und Datenschutz**

Die Web-App ist ohne Anmeldung erreichbar. Deshalb liefert `calendarDays` **keine Termintitel**, nur die Belegung je Tag:

- Anzahl Termine,
- belegte Stunden (6 bis 22 Uhr),
- früh oder abends belegt, von–bis,
- Merker für Reise, Urlaub und Gesundheit (Arzt, Physio).

Die Merker erkennt das Skript an Stichwörtern im Titel; der Titel selbst verlässt Google nicht. Ausgelassen werden Kalender mit „Coach Kira“, „Feiertage“ oder „Geburtstage“ im Namen sowie Termine, die mit „[Coach Kira]“ beginnen. Beim ersten Aufruf fragt Google eventuell noch einmal nach der Kalender-Berechtigung.

**Supervisor-Prompt**

Nach dem Wetterblock kommt ein kurzer Block „Kira-Treffsicherheit“ hinzu, sobald es eine Wochenbilanz gibt. So lernt der KI-Report aus deinen echten Reaktionen. Solange die Bilanz fehlt, bleibt der Block leer.

---

## Änderung 1 – `doGet`: drei neue Routen

**Streiche**

```javascript
  if (p.mode === 'diag') return ckDash_diag_();
```

**Ersetze durch**

```javascript
  if (p.mode === 'diag') return ckDash_diag_();
  // KI-Extras (Treffsicherheit, Selbsttest, Plausibilität, Ernährung, Alltag-Check)
  if (p.mode === 'extras') return ckX_get_(p);
  if (p.mode === 'calendarDays') return ckX_calendar_(p);
  if (p.mode === 'physBaseline') return ckX_phys_();
```

---

## Änderung 2 – `doPost`: Speichern

**Streiche**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveSupervisor') || /"kind"\s*:\s*"saveSupervisor"/.test(ckBody)) {
    return ckSup_save_(e);
  }
```

**Ersetze durch**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveSupervisor') || /"kind"\s*:\s*"saveSupervisor"/.test(ckBody)) {
    return ckSup_save_(e);
  }
  if ((e && e.parameter && e.parameter.mode === 'saveExtra') || /"kind"\s*:\s*"saveExtra"/.test(ckBody)) {
    return ckX_save_(e);
  }
```

---

## Änderung 3 – `createMasterGeminiPrompt`: Treffsicherheit

**Streiche**

```javascript
**NEU (V42): WETTER PROGNOSE (HEUTE & MORGEN):**
${wetterDaten}
---
```

**Ersetze durch**

```javascript
**NEU (V42): WETTER PROGNOSE (HEUTE & MORGEN):**
${wetterDaten}
---
${ckX_trefferPromptBlock_()}
```

---

## Änderung 4 – neue Funktionen (ans Ende der Datei)

```javascript
// ============================================================
// Coach Kira – KI-Extras (Treffsicherheit, Selbsttest, Plausibilität,
// Ernährung, Alltag-Check) · Speicher: Blatt AI_KIRA_EXTRAS
// ============================================================
const CKX_SHEET = 'AI_KIRA_EXTRAS';
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout'];

function ckX_sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CKX_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CKX_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['Datum', 'Typ', 'Text', 'Daten (JSON)', 'Quelle', 'Gespeichert']]);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

/** POST kind=saveExtra · {typ, datum:'YYYY-MM-DD', text, data, quelle} · ersetzt Eintrag mit gleichem Typ+Datum */
function ckX_save_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const typ = String(b.typ || '');
    if (CKX_TYPES.indexOf(typ) < 0) throw new Error('Unbekannter Typ: ' + typ);
    const datum = String(b.datum || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(datum)) throw new Error('datum fehlt oder ist nicht YYYY-MM-DD.');
    const text = String(b.text || '').slice(0, 6000);
    const data = b.data === undefined ? '' : JSON.stringify(b.data).slice(0, 40000);
    const quelle = String(b.quelle || 'lokal').replace(/[^\w:.\- ]/g, '').slice(0, 60);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = ckX_sheet_();
      const row = [datum, typ, text, data, quelle, ckDash_now_()];
      const n = sh.getLastRow() - 1;
      if (n > 0) {
        const keys = sh.getRange(2, 1, n, 2).getDisplayValues();
        for (let i = n - 1; i >= 0; i--) {
          if (keys[i][0] === datum && keys[i][1] === typ) {
            sh.getRange(i + 2, 1, 1, 6).setValues([row]);
            return ckDash_json_({ ok: true, timestamp: ckDash_now_(), updated: true });
          }
        }
      }
      sh.appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), updated: false });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** GET mode=extras&typ=a,b&n=30 · neueste zuerst */
function ckX_get_(p) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (!sh || sh.getLastRow() < 2) return ckDash_json_({ ok: true, timestamp: ckDash_now_(), items: [] });
    const types = String(p.typ || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    const n = Math.max(1, Math.min(400, Number(p.n) || 30));
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();
    const items = [];
    for (let i = rows.length - 1; i >= 0 && items.length < n; i--) {
      const r = rows[i];
      if (types.length && types.indexOf(r[1]) < 0) continue;
      let data = null;
      if (r[3]) { try { data = JSON.parse(r[3]); } catch (x) { data = null; } }
      items.push({ datum: r[0], typ: r[1], text: r[2], data: data, quelle: r[4], gespeichert: r[5] });
    }
    items.sort(function (a, b) { return a.datum < b.datum ? 1 : a.datum > b.datum ? -1 : 0; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), items: items });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/**
 * GET mode=calendarDays&days=14 · nur Belegung, KEINE Termintitel
 * (die Web-App ist öffentlich erreichbar).
 */
function ckX_calendar_(p) {
  try {
    const days = Math.max(1, Math.min(21, Number(p.days) || 14));
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const skip = /coach\s*kira|feiertag|holiday|geburtstag|birthday|wochennummer|week numbers/i;
    const reise = /reise|flug|flight|hotel|dienstreise|messe|zug nach|bahn/i;
    const frei = /urlaub|ferien|frei\b|vacation/i;
    const krank = /krank|arzt|zahnarzt|physio|op\b/i;
    const cals = CalendarApp.getAllCalendars().filter(function (c) { return !skip.test(c.getName()); });
    const start = new Date(); start.setHours(0, 0, 0, 0);
    const out = [];
    for (let d = 0; d < days; d++) {
      const a = new Date(start.getTime() + d * 86400000);
      const b = new Date(a.getTime() + 86400000);
      const dayStart = new Date(a.getTime() + 6 * 3600000), dayEnd = new Date(a.getTime() + 22 * 3600000);
      let busyMin = 0, allDay = 0, evening = false, morning = false, isReise = false, isFrei = false, isKrank = false, count = 0;
      let first = null, last = null;
      cals.forEach(function (c) {
        c.getEvents(a, b).forEach(function (ev) {
          if (/^\[coach kira\]/i.test(ev.getTitle())) return;
          const t = ev.getTitle() || '';
          count++;
          if (reise.test(t)) isReise = true;
          if (frei.test(t)) isFrei = true;
          if (krank.test(t)) isKrank = true;
          if (ev.isAllDayEvent()) { allDay++; return; }
          const s = new Date(Math.max(ev.getStartTime().getTime(), dayStart.getTime()));
          const e2 = new Date(Math.min(ev.getEndTime().getTime(), dayEnd.getTime()));
          if (e2 > s) busyMin += (e2 - s) / 60000;
          if (ev.getEndTime().getHours() >= 18 || ev.getStartTime().getHours() >= 17) evening = true;
          if (ev.getStartTime().getHours() < 8) morning = true;
          if (!first || ev.getStartTime() < first) first = ev.getStartTime();
          if (!last || ev.getEndTime() > last) last = ev.getEndTime();
        });
      });
      out.push({
        date: Utilities.formatDate(a, tz, 'yyyy-MM-dd'),
        termine: count,
        ganztaegig: allDay,
        belegt_h: Math.round(Math.min(busyMin, 16 * 60) / 6) / 10,
        frueh: morning,
        abends: evening,
        von: first ? Utilities.formatDate(first, tz, 'HH:mm') : '',
        bis: last ? Utilities.formatDate(last, tz, 'HH:mm') : '',
        reise: isReise,
        frei: isFrei,
        gesundheit: isKrank
      });
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), calendars: cals.length, days: out });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** GET mode=physBaseline · Gewicht, Max-HF, Laktatschwelle, FTP … aus phys_baseline */
function ckX_phys_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('phys_baseline');
    if (!sh) throw new Error('Blatt phys_baseline fehlt.');
    const v = sh.getRange(1, 1, 2, sh.getLastColumn()).getValues();
    const o = {};
    v[0].forEach(function (h, i) { if (String(h).trim()) o[String(h).trim()] = v[1][i]; });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), baseline: o });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** Textblock für den Supervisor-Prompt: letzte Wochenbilanz der Kira-Treffsicherheit */
function ckX_trefferPromptBlock_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (!sh || sh.getLastRow() < 2) return '';
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i][1] === 'treffer_woche' && rows[i][2]) {
        return '\n---\n**KIRA-TREFFSICHERHEIT (Rückblick, Stand ' + rows[i][0] + '):**\n' +
          'So hat Marc in den letzten Wochen auf deine Empfehlungen reagiert. Berücksichtige das bei Plan und Empfehlung.\n' +
          String(rows[i][2]).slice(0, 1500) + '\n---\n';
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

```

---

## Danach

Speichern, neue Version bereitstellen, kurz Bescheid geben. Ich teste dann alle vier Routen, fülle die Treffsicherheit mit den vorhandenen Plan-Briefings rückwirkend und schalte die Mac-Funktionen frei.
