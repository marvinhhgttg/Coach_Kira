# Apps-Script-Patch: Ereignis-Marker

Stand: 24.10.2026 · Datei: dein Apps-Script-Projekt (dort, wo `CKX_TYPES` steht, eingefügt mit dem Patch „Zukunftsplanung“)

## Was sich ändert

- Neuer Eintragstyp `ereignis` im Blatt **AI_KIRA_EXTRAS** – keine neue Spalte, kein neues Blatt, KK_TIMELINE bleibt unverändert.
- Je Tag ein Eintrag: Datum, Typ `ereignis`, Text (z. B. „krank, Reise“), Daten-JSON `{"arten":["krank","reise"],"notiz":"…"}`, Quelle `cockpit`.
- Speichern, Lesen und Löschen laufen über die vorhandenen Wege (`saveExtra`, `mode=extras`, `loeschen:true`).

## Änderung 1 – `CKX_TYPES` um `ereignis` ergänzen

**Streiche**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose'];
```

**Ersetze durch**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis'];
```

## Bereitstellen

Speichern → Bereitstellen → Bereitstellungen verwalten → bestehende Web-App bearbeiten → Version „Neue Version“ → Bereitstellen. Die URL bleibt gleich.

## Test

Im Cockpit unter Eingabe → „Ereignisse“ einen Marker für heute setzen und wieder löschen. Ohne den Patch zeigt die Kachel „Patch ‚Ereignis-Marker‘ noch nicht aktiv“.
