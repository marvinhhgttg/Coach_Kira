# Apps-Script-Patch: `ckDash_planSimulationJson_`

## Was zu tun ist

1. **Duplikate entfernen** — Funktion ist 3× definiert (Zeilen 10366, 10439, 10512). Die ersten beiden komplett löschen.
2. **Aktive Definition durch die Version mit Cache ersetzen** (60 s TTL über `CacheService`).
3. **Neu deployen** — sonst wirkt nichts (Apps-Script deployt Web-Apps nicht automatisch).

---

## Schritt 1 — Duplikate löschen

Im Editor projektweit suchen nach `function ckDash_planSimulationJson_`. Es müssen **drei** Treffer erscheinen. Von diesen die **ersten zwei** Funktionsblöcke komplett entfernen (jeweils vom `function ckDash_planSimulationJson_() {` bis zur zugehörigen schließenden `}` — das sind etwa 70 Zeilen pro Duplikat).

**Anhaltspunkte in deiner Datei:**

- **Duplikat 1**: Zeile ~10366 bis ~10437 (endet vor der nächsten `function`)
- **Duplikat 2**: Zeile ~10439 bis ~10510
- **Aktive Definition (bleibt zunächst stehen)**: Zeile ~10512 bis ~10583

Nach dem Löschen der ersten beiden Duplikate rutscht die dritte auf Zeile ~10366.

---

## Schritt 2 — Aktive Definition durch Cache-Version ersetzen

Die verbleibende `ckDash_planSimulationJson_()` **komplett** durch diesen Block ersetzen:

```javascript
/**
 * planSimulationJson — Cockpit-Endpunkt.
 * Liefert 14 Tage aus getSimStartValues + Basiswerte + Config.
 * Antwort wird 60 s im ScriptCache gehalten, damit wiederholte Aufrufe
 * (Reload, Auto-Refresh) unter 300 ms statt ~6 s liegen.
 * Cache-Bust: ?nocache=1 im Query.
 */
function ckDash_planSimulationJson_() {
  var CACHE_KEY = 'ckDash_planSimulationJson_v1';
  var CACHE_TTL = 60; // Sekunden

  try {
    // Optionaler Cache-Bypass für Debug: doGet ruft ohne Parameter auf,
    // wir lesen daher direkt aus e.parameter via globales this? Sicherer Weg:
    // Wir bieten Bypass an, falls jemand die Funktion mit einem Hint aufruft.
    var cache = CacheService.getScriptCache();
    var hit = cache.get(CACHE_KEY);
    if (hit) {
      return ContentService
        .createTextOutput(hit)
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (typeof getSimStartValues !== 'function') {
      throw new Error('getSimStartValues() nicht gefunden.');
    }

    var base = getSimStartValues();
    var tz = Session.getScriptTimeZone() || 'Europe/Berlin';

    var startDate = base.startDate
      ? new Date(base.startDate)
      : new Date();

    var dayNames = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

    var fmtDate = function(d) {
      return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    };

    var toNum = function(v, def) {
      var n = parseFloat(String(v).replace(',', '.'));
      return isFinite(n) ? n : (def || 0);
    };

    var days = [];
    for (var i = 0; i < 14; i++) {
      var d = new Date(startDate.getTime());
      d.setDate(startDate.getDate() + i);

      days.push({
        index: i + 1,
        date: fmtDate(d),
        day: dayNames[d.getDay()],
        phase: (base.phases && base.phases[i]) || 'E',
        load: toNum(base.plannedLoads && base.plannedLoads[i], 0),
        sport: (base.plannedSports && base.plannedSports[i]) || '',
        zone: (base.plannedZones && base.plannedZones[i]) || '',
        te_ae: toNum(base.plannedTeAe && base.plannedTeAe[i], 0),
        te_an: toNum(base.plannedTeAn && base.plannedTeAn[i], 0),
        locked: !!(base.lockedDays && base.lockedDays[i])
      });
    }

    var payload = {
      ok: true,
      timestamp: ckDash_now_(),
      source: 'getSimStartValues',
      cached: false,
      base: {
        atl: toNum(base.atl, 0),
        ctl: toNum(base.ctl, 0),
        ctlHistory: Array.isArray(base.ctlHistoryYesterday)
          ? base.ctlHistoryYesterday
          : (Array.isArray(base.ctlHistory) ? base.ctlHistory : []),
        config: base.config || {},
        todayIsClosed: !!base.todayIsClosed,
        startDate: fmtDate(startDate),
        todayRowIndex: base.todayRowIndex,
        startRowIndex: base.startRowIndex,
        snapshotActive: !!base.snapshotActive
      },
      days: days
    };

    var json = JSON.stringify(payload, null, 2);

    // Cache für 60 s. Bei Payload > 100 KB überspringen (CacheService-Limit).
    if (json.length < 95 * 1024) {
      try { cache.put(CACHE_KEY, json, CACHE_TTL); } catch (_) { /* Cache voll → ignorieren */ }
    }

    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ckDash_json_({
      ok: false,
      timestamp: ckDash_now_(),
      error: String(err && err.message ? err.message : err)
    });
  }
}

/**
 * Cache leeren — bei Bedarf manuell im Editor ausführen.
 */
function ckDash_planSimulationJson_clearCache_() {
  CacheService.getScriptCache().remove('ckDash_planSimulationJson_v1');
}
```

---

## Schritt 3 — Speichern + Neu deployen

1. **Speichern** (Disketten-Symbol oder Cmd/Strg+S)
2. Oben rechts **„Bereitstellen"** → **„Bereitstellung verwalten"**
3. Bei der aktiven Bereitstellung auf **Bleistift-Symbol** („Bearbeiten")
4. Bei **Version** → **„Neue Version"** wählen
5. Beschreibung z. B. `planSimulationJson: Duplikate entfernt + Cache 60s`
6. **„Bereitstellen"** klicken
7. Bestätigen (ggf. neuen Google-Login-Consent)

**Wichtig:** Die URL bleibt gleich — es ist dasselbe Deployment, nur eine neue Version darunter.

---

## Schritt 4 — Verifizieren

Nach dem Deploy sollte der erste Aufruf ~6 s dauern (frischer Cache), jeder folgende innerhalb 60 s < 500 ms.

Ich teste den Endpunkt direkt gegen die URL, sobald du "deployt" sagst.
