# Apps-Script-Patch: Garmin km/hm

Stand: 24.10.2026 · Datei: dein Apps-Script-Projekt (dort, wo `CKX_TYPES` und `ckX_save_` stehen)

## Was sich ändert

- Der neue Mac-Dienst `kira_garmin.py` liest km und Höhenmeter aus der offiziellen Garmin Personal API. Er schreibt sie über die neue POST-Weiche `saveDayMetrics` in die vorhandenen Spalten `distance_km_day` und `elev_m_day`, und zwar in **timeline** (Quelle) und **KK_TIMELINE** (Spiegel), wie beim Befinden-Patch.
- **Modus `leer`** (Standard, für das Nachziehen der Historie): füllt nur leere Zellen, vorhandene Werte bleiben.
- **Modus `alle`**: überschreibt die Werte. Der Dienst nutzt ihn nur für heute und gestern.
- **`probe`**: schreibt nichts, zählt nur.
- **Doppelte Datumszeilen** (z. B. 30.04.) werden beide befüllt.
- **Formeln:** Steht in einer der beiden Spalten eine Formel, bricht der Aufruf ab, ohne zu schreiben.
- **Neuer Extra-Typ `garmin_api`** in **AI_KIRA_EXTRAS**: je Tag ein Eintrag mit den API-Werten, mit denen die Cockpit-Eingabe vorausfüllt.
- **Was gleich bleibt:**
  - keine neuen Spalten und kein neues Blatt;
  - das Formular, `onFormSubmit` und die Garmin-Skills bleiben unverändert;
  - es gibt keine automatische Eingabe, du speicherst weiter selbst.

---

## Änderung 1 – `CKX_TYPES` um `garmin_api` ergänzen

**Streiche**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis'];
```

**Ersetze durch**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api'];
```

---

## Änderung 2 – `doPost`: neue Weiche

**Streiche**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveExtra') || /"kind"\s*:\s*"saveExtra"/.test(ckBody)) {
    return ckX_save_(e);
  }
```

**Ersetze durch**

```javascript
  if ((e && e.parameter && e.parameter.mode === 'saveExtra') || /"kind"\s*:\s*"saveExtra"/.test(ckBody)) {
    return ckX_save_(e);
  }
  if (/"kind"\s*:\s*"saveDayMetrics"/.test(ckBody)) {
    return ckGA_saveDays_(e);
  }
```

---

## Änderung 3 – neue Funktion (ans Ende derselben Datei, in der `ckX_save_` steht)

```javascript
// ============================================================
// Coach Kira – Garmin km/hm (Mac-Dienst kira_garmin.py, Garmin Personal API)
// POST kind=saveDayMetrics · {rows:[{datum:'YYYY-MM-DD', km, hm}], modus:'leer'|'alle', probe:true|false}
// Schreibt distance_km_day / elev_m_day in 'timeline' (Quelle) und 'KK_TIMELINE' (Spiegel).
// modus 'leer' (Standard): nur leere Zellen füllen · 'alle': vorhandene Werte überschreiben
// probe: nichts schreiben, nur zählen
// ============================================================
function ckGA_saveDays_(e) {
  const lock = LockService.getScriptLock();
  let gesperrt = false;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const rows = Array.isArray(b.rows) ? b.rows : [];
    if (!rows.length) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'rows fehlt' });
    if (rows.length > 400) return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: 'max 400 Tage je Aufruf' });
    const modus = b.modus === 'alle' ? 'alle' : 'leer';
    const probe = b.probe === true;
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const heute = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    const tage = {};
    let verworfen = 0;
    rows.forEach(function (r) {
      const datum = String((r && r.datum) || '');
      const km = Number(r && r.km);
      const hm = Number(r && r.hm);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(datum) || datum > heute ||
          !isFinite(km) || km < 0 || km > 1000 || !isFinite(hm) || hm < 0 || hm > 20000) {
        verworfen++;
        return;
      }
      tage[datum] = { km: Math.round(km * 100) / 100, hm: Math.round(hm) };
    });

    lock.waitLock(20000);
    gesperrt = true;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leer = function (x) { return x === '' || x === null; };
    const blaetter = {};

    ['timeline', 'KK_TIMELINE'].forEach(function (name) {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error("Blatt '" + name + "' nicht gefunden.");
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
      const cD = head.indexOf('date'), cK = head.indexOf('distance_km_day'), cH = head.indexOf('elev_m_day');
      [['date', cD], ['distance_km_day', cK], ['elev_m_day', cH]].forEach(function (p) {
        if (p[1] < 0) throw new Error("Spalte '" + p[0] + "' fehlt in Blatt '" + name + "'.");
      });
      const n = sh.getLastRow() - 1;
      const st = { geschrieben: 0, unveraendert: 0, behalten: 0, ohne_zeile: 0 };
      if (n < 1) { st.ohne_zeile = Object.keys(tage).length; blaetter[name] = st; return; }

      const rK = sh.getRange(2, cK + 1, n, 1), rH = sh.getRange(2, cH + 1, n, 1);
      const fK = rK.getFormulas(), fH = rH.getFormulas();
      for (let i = 0; i < n; i++) {
        if (fK[i][0]) throw new Error("Formel in Spalte 'distance_km_day' von Blatt '" + name + "' (Zeile " + (i + 2) + ').');
        if (fH[i][0]) throw new Error("Formel in Spalte 'elev_m_day' von Blatt '" + name + "' (Zeile " + (i + 2) + ').');
      }
      const dates = sh.getRange(2, cD + 1, n, 1).getValues();
      const vK = rK.getValues(), vH = rH.getValues();
      const zeilen = {};
      for (let i = 0; i < n; i++) {
        const d = dates[i][0];
        const key = d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
        if (key) (zeilen[key] = zeilen[key] || []).push(i);
      }

      Object.keys(tage).forEach(function (datum) {
        const idx = zeilen[datum];
        if (!idx) { st.ohne_zeile++; return; }
        const w = tage[datum];
        idx.forEach(function (i) {
          const aK = vK[i][0], aH = vH[i][0];
          if (aK === w.km && aH === w.hm) { st.unveraendert++; return; }
          if (modus !== 'alle' && (!leer(aK) || !leer(aH))) { st.behalten++; return; }
          vK[i][0] = w.km;
          vH[i][0] = w.hm;
          st.geschrieben++;
        });
      });

      if (!probe && st.geschrieben > 0) {
        rK.setValues(vK);
        rH.setValues(vH);
      }
      blaetter[name] = st;
    });

    return ckDash_json_({
      ok: true, timestamp: ckDash_now_(), modus: modus, probe: probe,
      tage: Object.keys(tage).length, verworfen: verworfen, blaetter: blaetter,
    });
  } catch (err) {
    try { logToSheet('ERROR', '[ckGA_saveDays_] ' + err); } catch (e2) {}
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  } finally {
    if (gesperrt) lock.releaseLock();
  }
}
```

---

## Bereitstellen

Speichern → Bereitstellen → Bereitstellungen verwalten → bestehende Web-App bearbeiten → Version „Neue Version“ → Bereitstellen. Die URL bleibt gleich.

## Test (ändert nichts, `probe`)

```bash
curl -sL -X POST "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec" \
  -H 'Content-Type: text/plain;charset=utf-8' \
  -d '{"kind":"saveDayMetrics","probe":true,"rows":[{"datum":"2026-10-24","km":10.5,"hm":120}]}'
```

Erwartet: `"ok": true`, `"probe": true`, `"tage": 1` und unter `blaetter` je Blatt die Zähler. Zeigt curl eine HTML-Seite, sag mir Bescheid, dann prüfe ich es von hier.
