# Apps-Script-Patch: Alltag-Kalender

Ziel: Der Alltag-Check im Cockpit berücksichtigt nur die Kalender **Marc**, **Marc (geteilt)** und **Marc & Natalia**. Ganztägige Termine zählen nicht mehr.

Was sich ändert:
1. `mode=calendarDays` versteht zwei neue, freiwillige Angaben:
   - `kal=Marc|Marc (geteilt)|Marc & Natalia`: nur diese Kalender. Der Name muss exakt stimmen, Groß- und Kleinschreibung ist egal.
   - `ganztag=0`: ganztägige Termine werden komplett übergangen. Sie zählen weder als Termin noch für Reise, Urlaub oder Arzt.
2. Die Antwort meldet zurück, welche Kalender tatsächlich gefunden wurden (`filter.kalender`). Das Cockpit warnt, falls ein Name nicht passt.

Ohne diese Angaben bleibt alles wie bisher. Morgen-Briefing, Supervisor und Mac lesen den Kalender also unverändert. Termintitel werden weiterhin nicht ausgegeben.

## Datei: KiraGeminiSupervisor.gs, Funktion `ckX_calendar_`

### Änderung 1 – Streiche

```javascript
    const cals = CalendarApp.getAllCalendars().filter(function (c) { return !skip.test(c.getName()); });
```

### Änderung 1 – Ersetze

```javascript
    // optional: nur bestimmte Kalender (kal=Name|Name|…, exakter Name ohne Groß/Klein) und ganztag=0 = ganztägige Termine ignorieren
    const nurKal = String(p.kal || '').split('|').map(function (x) { return x.trim().toLowerCase(); }).filter(String);
    const ohneGanztag = String(p.ganztag) === '0';
    const cals = CalendarApp.getAllCalendars().filter(function (c) {
      const n = String(c.getName() || '').trim().toLowerCase();
      return nurKal.length ? nurKal.indexOf(n) >= 0 : !skip.test(c.getName());
    });
```

### Änderung 2 – Streiche

```javascript
          const t = ev.getTitle() || '';
          // gleicher Termin in mehreren Kalendern nur einmal zählen
```

### Änderung 2 – Ersetze

```javascript
          if (ohneGanztag && ev.isAllDayEvent()) return;
          const t = ev.getTitle() || '';
          // gleicher Termin in mehreren Kalendern nur einmal zählen
```

### Änderung 3 – Streiche

```javascript
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), calendars: cals.length, days: out });
```

### Änderung 3 – Ersetze

```javascript
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), calendars: cals.length, days: out,
      filter: { kalender: nurKal.length ? cals.map(function (c) { return c.getName(); }) : null, ganztaegig: !ohneGanztag } });
```

## Danach

Speichern, dann **Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Version: Neue Version → Bereitstellen**. Die URL bleibt gleich.

Prüfen im Browser, die Antwort muss `"filter"` mit den drei Kalendernamen enthalten:

```
https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec?mode=calendarDays&days=3&kal=Marc|Marc%20(geteilt)|Marc%20%26%20Natalia&ganztag=0
```

Fehlt ein Name in `filter.kalender`, heißt der Kalender in Google Kalender anders. Dann den Namen in `ALLTAG_KALENDER` im Cockpit (`src/lib/extras.ts`) anpassen.
