# Apps-Script-Patch: Strategie speichern

Ziel: Events (Formspitze) und Szenarien A/B/C aus der Strategie-Ansicht liegen dauerhaft im Sheet statt nur im Browser.

Was sich ändert:
- `saveExtra` akzeptiert den neuen Typ `strategie`. Gespeichert wird genau ein Eintrag (Datum-Schlüssel `2025-08-26`) im Blatt `AI_KIRA_EXTRAS`; Lesen läuft über das vorhandene `mode=extras&typ=strategie`.
- Sonst nichts: keine neuen Blätter, keine neuen Modi.

Ohne den Patch funktioniert das Cockpit weiter wie bisher (dann mit dem Hinweis „Nur im Browser – Apps-Script-Patch fehlt“).

## Datei: KiraGeminiSupervisor.gs

### Streiche

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag'];
```

### Ersetze

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag',
  'strategie'];
```

Danach: Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Neue Version → Bereitstellen (gleiche URL).

Prüfen: Strategie-Ansicht öffnen. Unter „6 · Events & Formspitze“ steht dann „Im Sheet gespeichert · Stand …“. Events und Szenarien, die bisher im Browser lagen, werden beim ersten Öffnen automatisch übernommen.
