# Apps-Script-Patch: Befinden speichern (`saveWellbeing`)

## 1. Route in `doGet` einfügen

In `KiraGeminiSupervisor.gs` in `doGet(e)` **direkt unter** diesen Block:

```javascript
if (p.action === 'saveSimulatedPlan') {
  return ckDash_saveSimulatedPlan_(p);
}
```

diese drei Zeilen einfügen:

```javascript
if (p.action === 'saveWellbeing') {
  return ckDash_saveWellbeing_(p);
}
```

## 2. Neue Funktion anhängen

Am **Ende** von `KiraGeminiSupervisor.gs` (unterhalb aller anderen Funktionen) einfügen:

```javascript
/**
 * saveWellbeing — speichert das morgendliche Befinden (1–5) für heute oder gestern.
 * Schreibt in 'timeline' (Quelle) UND 'KK_TIMELINE' (Spiegel), Spalte 'befinden_1_5'.
 * Parameter: token, date (yyyy-MM-dd), value (1–5)
 */
function ckDash_saveWellbeing_(p) {
  var COL_NAME = 'befinden_1_5';
  var lock = LockService.getScriptLock();
  try {
    var auth = ckDash_requireToken_(p);
    if (!auth.ok) {
      return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: auth.error });
    }

    var value = parseInt(String(p.value || '').trim(), 10);
    if (!(value >= 1 && value <= 5)) throw new Error('value muss 1–5 sein.');

    var tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var y = new Date();
    y.setDate(y.getDate() - 1);
    var yesterdayKey = Utilities.formatDate(y, tz, 'yyyy-MM-dd');

    var dateKey = String(p.date || todayKey).trim().slice(0, 10);
    if (dateKey !== todayKey && dateKey !== yesterdayKey) {
      throw new Error('Nur heute oder gestern speicherbar (' + dateKey + ').');
    }

    lock.waitLock(10000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = [];
    ['timeline', 'KK_TIMELINE'].forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) throw new Error('Blatt ' + name + ' nicht gefunden.');

      var lastCol = sh.getLastColumn();
      var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
        return String(h || '').trim().toLowerCase();
      });
      var colIdx = headers.indexOf(COL_NAME);
      var dateIdx = headers.indexOf('date');
      if (colIdx < 0) throw new Error('Spalte ' + COL_NAME + ' fehlt in ' + name + '.');
      if (dateIdx < 0) throw new Error('Spalte date fehlt in ' + name + '.');

      var lastRow = sh.getLastRow();
      var dates = sh.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues();
      var row = -1;
      for (var i = 0; i < dates.length; i++) {
        var d = dates[i][0];
        var key = (d instanceof Date)
          ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
          : String(d || '').trim().slice(0, 10);
        if (key === dateKey) { row = i + 2; break; }
      }
      if (row < 0) throw new Error('Datum ' + dateKey + ' nicht in ' + name + ' gefunden.');

      sh.getRange(row, colIdx + 1).setValue(value);
      written.push(name + '!R' + row);
    });

    SpreadsheetApp.flush();
    try { logToSheet('INFO', '[Dashboard] saveWellbeing ' + dateKey + ' = ' + value + ' → ' + written.join(', ')); } catch (_) {}

    return ckDash_json_({
      ok: true,
      timestamp: ckDash_now_(),
      date: dateKey,
      value: value,
      written: written
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

## 3. Speichern und neu bereitstellen

1. Speichern (Cmd+S)
2. **Bereitstellen → Bereitstellung verwalten → Bleistift → Version: Neue Version → Bereitstellen**

Danach „fertig“ sagen. Ich teste dann, dass ein Aufruf **ohne Token** abgelehnt wird (`unauthorized`). Einen echten Schreibtest mache ich nicht, der passiert beim ersten Klick im Cockpit.
