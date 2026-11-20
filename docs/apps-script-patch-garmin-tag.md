# Apps-Script-Patch: Garmin-Tagesdetails

Stand: 25.10.2026 · Datei: dein Apps-Script-Projekt (dort, wo `CKX_TYPES` und `ckX_get_` stehen)

## Was sich ändert

- Neuer Extra-Typ `garmin_tag` in **AI_KIRA_EXTRAS**: je Tag ein Eintrag mit Tageswerten, Schlaf, HRV und allen Aktivitätsdetails aus der Garmin Personal API, geschrieben vom Mac-Dienst.
- Abrufgrenze für `ckX_get_` auf 1000 Einträge hochgesetzt, damit das Cockpit die ganze Historie (~430 Tage) auf einmal laden kann.
- Keine neuen Spalten, Blätter oder Formularänderungen.

---

## Änderung 1 – `CKX_TYPES` um `garmin_tag` ergänzen

**Streiche**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api'];
```

**Ersetze durch**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose', 'ereignis', 'garmin_api', 'garmin_tag'];
```

---

## Änderung 2 – Abrufgrenze in `ckX_get_` von 400 auf 1000

**Streiche** (steht in `function ckX_get_`)

```javascript
    const n = Math.max(1, Math.min(400, Number(p.n) || 30));
```

**Ersetze durch**

```javascript
    const n = Math.max(1, Math.min(1000, Number(p.n) || 30));
```

---

## Bereitstellen

Speichern → Bereitstellen → Bereitstellungen verwalten → bestehende Web-App bearbeiten → Version „Neue Version“ → Bereitstellen. Die URL bleibt gleich.

## Test

Gib mir nach dem Bereitstellen kurz Bescheid. Ich speichere von hier einen Testeintrag `garmin_tag` und lösche ihn wieder. Ohne Patch kommt „Unbekannter Typ: garmin_tag“, mit Patch `ok: true`.
