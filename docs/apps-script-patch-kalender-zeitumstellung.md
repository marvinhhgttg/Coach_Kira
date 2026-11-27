# Apps-Script-Patch: Kalender bei Zeitumstellung

Stand: 22.10.2026 · nur in `ckX_calendar_`

**Problem:** Die Tage wurden in festen 24-Stunden-Schritten gezählt. Am 25.10. (Umstellung auf Winterzeit, 25 Stunden) landet der nächste Schritt auf So 23:00. Dadurch erscheint der 25.10. doppelt, und alle folgenden Tage sind um einen Tag verschoben: Die Termine von Mo 26.10. stehen beim 25.10.

**Lösung:** Kalendertage statt fester 24 Stunden.

**Streiche**

```javascript
      const a = new Date(start.getTime() + d * 86400000);
      const b = new Date(a.getTime() + 86400000);
      const dayStart = new Date(a.getTime() + 6 * 3600000), dayEnd = new Date(a.getTime() + 22 * 3600000);
```

**Ersetze durch**

```javascript
      // Kalendertage statt fester 24 h (sonst doppelter Tag bei der Zeitumstellung)
      const a = new Date(start.getFullYear(), start.getMonth(), start.getDate() + d, 0, 0, 0);
      const b = new Date(start.getFullYear(), start.getMonth(), start.getDate() + d + 1, 0, 0, 0);
      const dayStart = new Date(a.getFullYear(), a.getMonth(), a.getDate(), 6, 0, 0);
      const dayEnd = new Date(a.getFullYear(), a.getMonth(), a.getDate(), 22, 0, 0);
```

Danach speichern, neue Version bereitstellen, kurz Bescheid geben.
