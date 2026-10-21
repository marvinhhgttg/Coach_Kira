# Apps-Script-Patch: Kalender-Belegung ohne Doppelzählung

Stand: 21.10.2026 · Nachtrag zu „KI-Extras“, nur in `ckX_calendar_`

**Problem:** Termine, die in mehreren Kalendern stehen (z. B. Familien- und eigener Kalender), sowie sich überschneidende Termine wurden mehrfach gezählt. Am Sa 24.10. ergaben sich so 9 h Belegung, obwohl die Termine nur von 13:30 bis 18:00 reichen.

**Lösung:**

- Gleiche Termine (Titel, Beginn, Ende) zählen nur einmal.
- Die belegte Zeit wird als Vereinigung der Zeitfenster berechnet.
- Sonst ändert sich nichts.

---

## Änderung 1

**Streiche**

```javascript
      let busyMin = 0, allDay = 0, evening = false, morning = false, isReise = false, isFrei = false, isKrank = false, count = 0;
```

**Ersetze durch**

```javascript
      let allDay = 0, evening = false, morning = false, isReise = false, isFrei = false, isKrank = false, count = 0;
      const iv = [];
      const seen = {};
```

---

## Änderung 2

**Streiche**

```javascript
          if (/^\[coach kira\]/i.test(ev.getTitle())) return;
          const t = ev.getTitle() || '';
          count++;
```

**Ersetze durch**

```javascript
          if (/^\[coach kira\]/i.test(ev.getTitle())) return;
          const t = ev.getTitle() || '';
          // gleicher Termin in mehreren Kalendern nur einmal zählen
          const key = t + '|' + ev.getStartTime().getTime() + '|' + ev.getEndTime().getTime();
          if (seen[key]) return;
          seen[key] = true;
          count++;
```

---

## Änderung 3

**Streiche**

```javascript
          if (e2 > s) busyMin += (e2 - s) / 60000;
```

**Ersetze durch**

```javascript
          if (e2 > s) iv.push([s.getTime(), e2.getTime()]);
```

---

## Änderung 4 (erste Zeile der Ausgabe je Tag)

**Streiche**

```javascript
      out.push({
```

**Ersetze durch**

```javascript
      // Überschneidungen zusammenfassen
      iv.sort(function (x, y) { return x[0] - y[0]; });
      let busyMin = 0, curS = null, curE = null;
      iv.forEach(function (x) {
        if (curE === null || x[0] > curE) { if (curE !== null) busyMin += (curE - curS) / 60000; curS = x[0]; curE = x[1]; }
        else if (x[1] > curE) curE = x[1];
      });
      if (curE !== null) busyMin += (curE - curS) / 60000;
      out.push({
```

---

Danach speichern, neue Version bereitstellen, kurz Bescheid geben.
