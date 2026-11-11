# Apps-Script-Patch: Zeitfenster (Entlastung, Pause, Aktivurlaub)

Ziel: Du sagst dem System, wann eine Entlastungswoche passt (Dienstreise, Familienwoche), wann kein Training möglich ist und wann ein Aktivurlaub ansteht. Der Saisonrahmen legt seine Wochen danach; Cockpit und Mac lesen denselben Stand.

Was sich ändert:
1. `saveExtra` akzeptiert den neuen Typ `zeitfenster` (ein Eintrag je Zeitfenster im Blatt `AI_KIRA_EXTRAS`, Löschen über das vorhandene `loeschen: true`).
2. Neuer Lese-Modus `mode=calendarFenster&von=YYYY-MM-DD&bis=YYYY-MM-DD`: findet in deinen Kalendern ganztägige oder lange Termine (ab 4 h) mit Reise- oder Urlaubsbezug und gibt nur Zeiträume zurück – **keine Termintitel**, nur das gefundene Stichwort (z. B. „dienstreise“, „urlaub“). Die Web-App ist öffentlich erreichbar, deshalb bleiben Titel draußen. Gespeichert wird nichts; das Cockpit schlägt die Zeiträume nur vor.

## Datei: KiraGeminiSupervisor.gs

### Änderung 1 – Streiche

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag',
  'strategie'];
```

### Änderung 1 – Ersetze

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag',
  'strategie', 'zeitfenster'];
```

### Änderung 2 – Streiche

```javascript
  if (p.mode === 'calendarDays') return ckX_calendar_(p);
```

### Änderung 2 – Ersetze

```javascript
  if (p.mode === 'calendarDays') return ckX_calendar_(p);
  if (p.mode === 'calendarFenster') return ckX_calFenster_(p);
```

### Änderung 3 – Streiche

```javascript
/** GET mode=physBaseline · Gewicht, Max-HF, Laktatschwelle, FTP … aus phys_baseline */
```

### Änderung 3 – Ersetze

```javascript
/**
 * GET mode=calendarFenster&von=YYYY-MM-DD&bis=YYYY-MM-DD
 * Reise-/Urlaubszeiträume aus allen Kalendern (ganztägig, mehrtägig oder ab 4 h) – KEINE Termintitel,
 * nur Stichwort. Standard: heute bis +200 Tage, höchstens 500 Tage.
 */
function ckX_calFenster_(p) {
  try {
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const ymd = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };
    const parse = function (s, def) {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
      return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0) : def;
    };
    const heute = new Date(); heute.setHours(0, 0, 0, 0);
    const a = parse(p.von, heute);
    let b = parse(p.bis, new Date(heute.getFullYear(), heute.getMonth(), heute.getDate() + 200));
    if ((b - a) / 86400000 > 500) b = new Date(a.getFullYear(), a.getMonth(), a.getDate() + 500);
    const bEnde = new Date(b.getFullYear(), b.getMonth(), b.getDate() + 1);
    const skip = /coach\s*kira|feiertag|holiday|geburtstag|birthday|wochennummer|week numbers/i;
    const muster = [
      { typ: 'reise', re: /(dienstreise|geschäftsreise|messe|konferenz|workshop|seminar|reise|flug|flight|hotel)/i },
      { typ: 'frei', re: /(urlaub|ferien|vacation|skiurlaub|bergurlaub|wanderurlaub)/i },
    ];
    const tage = {}; // datum -> {typ, stichwort}
    const seen = {};
    CalendarApp.getAllCalendars().filter(function (c) { return !skip.test(c.getName()); }).forEach(function (c) {
      c.getEvents(a, bEnde).forEach(function (ev) {
        const t = ev.getTitle() || '';
        if (/^\[coach kira\]/i.test(t)) return;
        const key = t + '|' + ev.getStartTime().getTime() + '|' + ev.getEndTime().getTime();
        if (seen[key]) return;
        seen[key] = true;
        let treffer = null;
        for (let i = 0; i < muster.length && !treffer; i++) {
          const m = muster[i].re.exec(t);
          if (m) treffer = { typ: muster[i].typ, stichwort: m[1].toLowerCase() };
        }
        if (!treffer) return;
        const s = ev.getStartTime(), e = ev.getEndTime();
        const lang = ev.isAllDayEvent() || (e - s) >= 4 * 3600000 || ymd(s) !== ymd(new Date(e.getTime() - 1));
        if (!lang) return;
        // alle Kalendertage des Termins markieren (Ende exklusiv)
        const d = new Date(Math.max(s.getTime(), a.getTime())); d.setHours(0, 0, 0, 0);
        const ende = new Date(Math.min(e.getTime() - 1, bEnde.getTime() - 1));
        while (d <= ende) {
          const k = ymd(d);
          if (!tage[k] || (tage[k].typ === 'frei' && treffer.typ === 'reise')) tage[k] = treffer;
          d.setDate(d.getDate() + 1);
        }
      });
    });
    // zusammenhängende Tage gleichen Typs zu Zeiträumen zusammenfassen
    const keys = Object.keys(tage).sort();
    const out = [];
    keys.forEach(function (k) {
      const x = tage[k];
      const last = out[out.length - 1];
      const prev = last ? new Date(last.bis + 'T12:00:00') : null;
      const cur = new Date(k + 'T12:00:00');
      if (last && last.typ === x.typ && Math.round((cur - prev) / 86400000) === 1) {
        last.bis = k;
        if (last.stichworte.indexOf(x.stichwort) < 0) last.stichworte.push(x.stichwort);
      } else out.push({ von: k, bis: k, typ: x.typ, stichworte: [x.stichwort] });
    });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), von: ymd(a), bis: ymd(b), fenster: out });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** GET mode=physBaseline · Gewicht, Max-HF, Laktatschwelle, FTP … aus phys_baseline */
```

Danach: Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Neue Version → Bereitstellen (gleiche URL).

Prüfen: Plan → Ziele & Saison → Kachel „Zeitfenster“. Dort stehen Vorschläge aus dem Kalender; eigene Zeitfenster bleiben nach dem Neuladen erhalten.
