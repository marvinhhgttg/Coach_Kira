# Apps-Script-Patch: Cockpit-Eingabe statt Formular (`submitCockpitEntry`)

Drei Änderungen, alle in `KiraGeminiSupervisor.gs`. Danach **neue Version bereitstellen**.

---

## 1. Route in `doGet` einfügen

Direkt **unter** dem Block für `saveWellbeing`:

```javascript
if (p.action === 'saveWellbeing') {
  return ckDash_saveWellbeing_(p);
}
```

diese drei Zeilen einfügen:

```javascript
if (p.action === 'submitCockpitEntry') {
  return ckDash_submitCockpitEntry_(p);
}
```

---

## 2. `advanceTodayFlag_silent` ersetzen (Tageswechsel absichern)

Die bestehende Funktion `function advanceTodayFlag_silent() { … }` **komplett** durch diese Version ersetzen. Sie setzt `is_today` auf die Zeile mit dem **heutigen Datum**, statt blind eine Zeile weiterzuschieben. Doppeltes Absenden (Formular oder Cockpit) verschiebt dadurch nichts mehr, und ein ausgelassener Tag wird automatisch übersprungen.

```javascript
/**
 * (V2 2026-09-27) Setzt 'is_today' in 'timeline' auf die Zeile mit dem HEUTIGEN Datum.
 * Idempotent: mehrfacher Aufruf am selben Tag ändert nichts. Wird von onFormSubmit ("Morgens") genutzt.
 * @returns {boolean} true, wenn die heutige Zeile gefunden und markiert wurde.
 */
function advanceTodayFlag_silent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SOURCE_TIMELINE_SHEET); // 'timeline'
  if (!sheet) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Quellblatt '${SOURCE_TIMELINE_SHEET}' nicht gefunden.`);
    return false;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const isTodayIdx = headers.indexOf('is_today');
  const dateIdx = headers.indexOf('date');
  if (isTodayIdx < 0 || dateIdx < 0) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Spalte 'is_today' oder 'date' fehlt in '${SOURCE_TIMELINE_SHEET}'.`);
    return false;
  }

  const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const dates = sheet.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
  const flags = sheet.getRange(2, isTodayIdx + 1, lastRow - 1, 1).getValues();

  let todayRow = -1;
  let oldRow = -1;
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    const key = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').trim().slice(0, 10);
    if (key === todayKey) todayRow = i + 2;
    if (Number(String(flags[i][0]).replace(',', '.')) === 1 && oldRow < 0) oldRow = i + 2;
  }
  if (todayRow < 0) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Keine Zeile für heute (${todayKey}) in '${SOURCE_TIMELINE_SHEET}'.`);
    return false;
  }
  if (oldRow === todayRow) {
    logToSheet('INFO', `[advanceTodayFlag_silent] 'is_today' steht bereits auf heute (Zeile ${todayRow}) – keine Änderung.`);
    return true;
  }

  try {
    const out = flags.map((_, i) => [i + 2 === todayRow ? 1 : 0]);
    sheet.getRange(2, isTodayIdx + 1, out.length, 1).setValues(out);
    logToSheet('INFO', `[advanceTodayFlag_silent] 'is_today' von Zeile ${oldRow} auf ${todayRow} (${todayKey}) gesetzt.`);
    return true;
  } catch (e) {
    logToSheet('ERROR', `[advanceTodayFlag_silent] Fehler beim Schreiben: ${e.message}`);
    return false;
  }
}
```

> Hinweis: Das Formular trägt Morgenwerte damit immer in die **heutige** Zeile ein. So war es auch bisher gedacht. Nachträge für einen früheren Morgen gingen auch vorher nicht sauber.

---

## 3. Neue Funktion anhängen

Am **Ende** der Datei einfügen:

```javascript
/**
 * submitCockpitEntry — Cockpit-Eingabe anstelle des Google-Formulars.
 * Baut ein Formular-Event (namedValues) und ruft die UNVERÄNDERTE onFormSubmit(e) auf.
 * Damit laufen Schreiben, Tageswechsel, IST→SOLL, Activity Review und KI-Report wie beim Formular.
 * Parameter: token, type (Morgens | Nach Aktivität | Abends), date (yyyy-MM-dd), values (JSON-Objekt)
 */
function ckDash_submitCockpitEntry_(p) {
  const ALLOWED = {
    'Morgens': ['garminATL', 'garminCTL', 'garminACWR', 'sleep_hours', 'sleep_score_0_100', 'rhr_bpm',
                'hrv_status', 'hrv_threshholds', 'Garmin_Training_Readiness', 'Trainingszustand'],
    'Nach Aktivität': ['load_fb_day', 'fbATL_obs', 'fbCTL_obs', 'fbACWR_obs', 'garminEnduranceScore',
                       'Sport_x', 'Zone', 'fb_TR_obs', 'Aerobic_TE', 'Anaerobic_TE'],
    'Abends': ['kcal_in', 'kcal_out', 'deficit', 'carb_g', 'protein_g', 'fat_g', 'fiber_g']
  };
  const TEXT_COLS = ['Sport_x', 'Zone', 'hrv_threshholds', 'Trainingszustand'];
  const lock = LockService.getDocumentLock();
  const started = Date.now();

  try {
    const auth = ckDash_requireToken_(p);
    if (!auth.ok) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: auth.error });

    const type = String(p.type || '').trim();
    if (!ALLOWED[type]) throw new Error('Unbekannter Typ: ' + type);
    if (typeof onFormSubmit !== 'function') throw new Error('onFormSubmit() nicht gefunden.');

    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const now = new Date();
    const todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const y = new Date(now.getTime()); y.setDate(y.getDate() - 1);
    const yesterdayKey = Utilities.formatDate(y, tz, 'yyyy-MM-dd');
    const dateKey = String(p.date || todayKey).trim().slice(0, 10);
    if (type === 'Morgens' && dateKey !== todayKey) throw new Error('Morgenwerte nur für heute (' + todayKey + ').');
    if (dateKey !== todayKey && dateKey !== yesterdayKey) throw new Error('Nur heute oder gestern erlaubt (' + dateKey + ').');

    let values = {};
    try { values = JSON.parse(String(p.values || '{}')); } catch (_) { throw new Error('values ist kein gültiges JSON.'); }

    const [yy, mm, dd] = dateKey.split('-');
    const namedValues = {
      'Zeitstempel': [Utilities.formatDate(now, tz, 'dd.MM.yyyy HH:mm:ss')],
      'Was willst Du eingeben?': [type],
      'Datum': [dd + '.' + mm + '.' + yy]
    };
    const used = [];
    ALLOWED[type].forEach(function(col) {
      if (!Object.prototype.hasOwnProperty.call(values, col)) return;
      let v = values[col];
      if (v === null || v === undefined || String(v).trim() === '') return;
      v = TEXT_COLS.indexOf(col) >= 0 ? String(v).trim() : String(v).trim().replace('.', ','); // wie Formular: Dezimalkomma
      namedValues[col] = [v];
      used.push(col);
    });
    if (!used.length) throw new Error('Keine Werte übergeben.');

    lock.waitLock(30000);
    onFormSubmit({ namedValues: namedValues, values: [], source: 'cockpit' });
    SpreadsheetApp.flush();

    // Rücklesen zur Bestätigung
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOURCE_TIMELINE_SHEET);
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    const dateIdx = headers.indexOf('date');
    const dates = sh.getRange(2, dateIdx + 1, sh.getLastRow() - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < dates.length; i++) {
      const d = dates[i][0];
      const k = (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
      if (k === dateKey) { row = i + 2; break; }
    }
    const readback = {};
    if (row > 0) {
      const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
      used.forEach(function(col) {
        const idx = headers.indexOf(col);
        if (idx >= 0) readback[col] = rowVals[idx] instanceof Date ? String(rowVals[idx]) : rowVals[idx];
      });
    }

    // Protokoll
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName('COCKPIT_ENTRIES');
    if (!log) {
      log = ss.insertSheet('COCKPIT_ENTRIES');
      log.getRange(1, 1, 1, 5).setValues([['timestamp', 'type', 'date', 'row', 'values_json']]).setFontWeight('bold');
    }
    log.appendRow([namedValues['Zeitstempel'][0], type, dateKey, row, JSON.stringify(values)]);
    try { logToSheet('INFO', '[Dashboard] submitCockpitEntry ' + type + ' ' + dateKey + ' → Zeile ' + row + ' (' + used.join(', ') + ')'); } catch (_) {}

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      type: type,
      date: dateKey,
      row: row,
      written: used,
      readback: readback,
      triggersStarted: dateKey === todayKey,
      durationMs: Date.now() - started
    });
  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
```

---

## 4. Speichern und neu bereitstellen

**Bereitstellen → Bereitstellung verwalten → Bleistift → Version: Neue Version → Bereitstellen**, danach „fertig“ sagen.

Ich teste danach nur ohne Token (muss `unauthorized` liefern). Einen echten Eintrag mache ich nicht selbst, weil er Tageswechsel und KI-Report auslöst. Den ersten Eintrag machst du im Cockpit.
