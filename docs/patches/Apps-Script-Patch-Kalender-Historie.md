# Apps-Script-Patch: Kalender-Historie

Für die Kachel **Alltag und Erholung** (Analyse → „Was wirkt auf deine Leistung?“). Neuer Modus `calendarHistorie`: Belegung vergangener Tage (nur Stunden, keine Termintitel), gefiltert auf deine drei Kalender, ganztägige Termine ausgenommen.

## Schritt 1 – Weiche ergänzen

Suche in `doGet` die Zeile:

```javascript
  if (p.mode === 'calendarDays') return ckX_calendar_(p);
```

**Füge direkt darunter ein:**

```javascript
  if (p.mode === 'calendarHistorie') return ckX_calendarHistorie_(p);
```

## Schritt 2 – Funktion ans Dateiende

**Füge ganz unten in der Datei ein:**

```javascript
/**
 * GET mode=calendarHistorie&von=YYYY-MM-DD&bis=YYYY-MM-DD&kal=Name|Name&ganztag=0
 * Vergangene Belegung je Tag (für „Alltag und Erholung“) · KEINE Termintitel.
 * Ein Abruf je Kalender für den ganzen Zeitraum (max. 500 Tage).
 */
function ckX_calendarHistorie_(p) {
  try {
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const parse = function (s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '')); return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null; };
    const heute = new Date(); heute.setHours(0, 0, 0, 0);
    const bis = parse(p.bis) || new Date(heute.getTime() - 86400000);
    let von = parse(p.von) || new Date(bis.getFullYear(), bis.getMonth(), bis.getDate() - 90);
    const maxVon = new Date(bis.getFullYear(), bis.getMonth(), bis.getDate() - 500);
    if (von < maxVon) von = maxVon;
    const ende = new Date(bis.getFullYear(), bis.getMonth(), bis.getDate() + 1);
    const nurKal = String(p.kal || '').split('|').map(function (x) { return x.trim().toLowerCase(); }).filter(String);
    const ohneGanztag = String(p.ganztag) === '0';
    const skip = /coach\s*kira|feiertag|holiday|geburtstag|birthday|wochennummer|week numbers/i;
    const cals = CalendarApp.getAllCalendars().filter(function (c) {
      const n = String(c.getName() || '').trim().toLowerCase();
      return nurKal.length ? nurKal.indexOf(n) >= 0 : !skip.test(c.getName());
    });
    const tage = {};
    const tag = function (d) {
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (!tage[k]) tage[k] = { iv: [], termine: 0, abends: false, frueh: false };
      return tage[k];
    };
    const seen = {};
    cals.forEach(function (c) {
      c.getEvents(von, ende).forEach(function (ev) {
        if (/^\[coach kira\]/i.test(ev.getTitle())) return;
        if (ev.isAllDayEvent()) { if (ohneGanztag) return; tag(ev.getStartTime()).termine++; return; }
        const s0 = ev.getStartTime(), e0 = ev.getEndTime();
        const key = (ev.getTitle() || '') + '|' + s0.getTime() + '|' + e0.getTime();
        if (seen[key]) return;
        seen[key] = true;
        // je Kalendertag zwischen 6 und 22 Uhr zerlegen (mehrtägige Termine)
        let d = new Date(s0.getFullYear(), s0.getMonth(), s0.getDate());
        let erster = true;
        while (d < e0 && d < ende) {
          const a = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 6, 0, 0);
          const b = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 22, 0, 0);
          const s = Math.max(s0.getTime(), a.getTime()), e = Math.min(e0.getTime(), b.getTime());
          if (d >= von) {
            const t = tag(d);
            if (erster) t.termine++;
            if (e > s) t.iv.push([s, e]);
            if (erster && s0.getHours() < 8) t.frueh = true;
            if (e0.getTime() >= new Date(d.getFullYear(), d.getMonth(), d.getDate(), 18, 0, 0).getTime() && s0.getTime() < b.getTime()) t.abends = true;
          }
          erster = false;
          d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
        }
      });
    });
    const out = [];
    for (let d = new Date(von); d < ende; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      const t = tage[k] || { iv: [], termine: 0, abends: false, frueh: false };
      t.iv.sort(function (x, y) { return x[0] - y[0]; });
      let min = 0, cs = null, ce = null;
      t.iv.forEach(function (x) {
        if (ce === null || x[0] > ce) { if (ce !== null) min += (ce - cs) / 60000; cs = x[0]; ce = x[1]; }
        else if (x[1] > ce) ce = x[1];
      });
      if (ce !== null) min += (ce - cs) / 60000;
      out.push({ date: k, termine: t.termine, belegt_h: Math.round(Math.min(min, 16 * 60) / 6) / 10, abends: t.abends, frueh: t.frueh });
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), filter: { kalender: cals.map(function (c) { return c.getName(); }), ohneGanztag: ohneGanztag }, days: out });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}
```

## Schritt 3 – Bereitstellen

Bereitstellen → Bereitstellungen verwalten → Bearbeiten (Stift) → Version: **Neue Version** → Bereitstellen. Die URL bleibt gleich.

## Test

Im Browser öffnen (sollte `"ok": true` und eine Liste `days` zeigen):

```
https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec?mode=calendarHistorie&von=2026-10-01&bis=2026-11-02&kal=Marc|Marc%20(geteilt)|Natalia%20%26%20Marc&ganztag=0
```
