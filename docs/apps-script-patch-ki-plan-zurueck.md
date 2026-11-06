# Apps-Script-Patch: KI-Plan zurück + gesperrte Tage absichern

Stand: 19.10.2026 · macht den Patch „KI-Plan im Supervisor streichen“ rückgängig

**Was passiert**

1. Der Supervisor erzeugt wieder seinen 8-Tage-Plan (`empfohlener_plan` → `AI_REPORT_PLAN`). Das sind die Änderungen 1 bis 3, also exakt die Umkehrung des Streich-Patches.
2. Neu ist die Absicherung gesperrter Tage: Für Tage mit `fix = x` in der timeline setzt `ckSup_fixOriginalLoads_` nach der KI-Antwort den empfohlenen Load fest auf deinen Plan. Bei 0 ESS werden außerdem die Zone auf „Ruhetag (Fixed)“ und die Beispiele auf „Geplanter Ruhetag“ gesetzt.
3. Die Leer-Logik in `ckSup_writeReportPlan_` (Änderung 4 des Streich-Patches) darf bleiben. Sie greift nur, wenn die KI keinen Plan liefert.

Im Cockpit ersetzt der Supervisor-Plan anschließend die Kachel „Plan-Briefing“ auf der Plan-Seite.

---

## Änderung 1 – JSON-Schema wieder mit Plan (an beiden Stellen: `callOpenAI` und `ckSup_schema_`)

**Streiche**

```javascript
        }
      },
      "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare"],
```

**Ersetze durch**

```javascript
        },
        "empfohlener_plan": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "datum": { "type": "string" },
              "tag": { "type": "string" },
              "original_load_ess": { "type": "number" },
              "empfohlener_load_ess": { "type": "number" },
              "empfohlene_zone": { "type": "string" },
              "beispiel_training_1": { "type": "string" },
              "beispiel_training_2": { "type": "string" },
              "prognostizierte_auswirkung": { 
                "type": "string", 
                "description": "Detaillierte Erklärung." 
              },
              "wetterempfehlung": { "type": "string" }
            },
            "required": ["datum", "tag", "original_load_ess", "empfohlener_load_ess", "empfohlene_zone", "beispiel_training_1", "beispiel_training_2", "prognostizierte_auswirkung", "wetterempfehlung"],
            "additionalProperties": false
          }
        }
      },
      "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare", "empfohlener_plan"],
```

---

## Änderung 2 – `createMasterGeminiPrompt`: Aufgabe 3

**Streiche**

```javascript
3. KEINEN Trainingsplan: Den 14-Tage-Plan erstellt das Plan-Briefing separat. Gib 'empfohlener_plan' NICHT aus und ignoriere die Plan-Regeln weiter unten; die Prognose-Daten dienen nur als Kontext für deine Gesamtbewertung.
```

**Ersetze durch**

```javascript
3. 'empfohlener_plan': Einen angepassten 8-TAGE-PLAN (HEUTE + 7 TAGE ZUKUNFT).
```

---

## Änderung 3 – `createMasterGeminiPrompt`: JSON-Beispiel

**Streiche**

```javascript
    /* ... und ALLE anderen aus der Liste ... */
  ]
}
```

**Ersetze durch**

```javascript
    /* ... und ALLE anderen aus der Liste ... */
  ],
  "empfohlener_plan": [
    {
      "datum": "${cleanDate}",
      "tag": "${cleanTag}",
      "original_load_ess": 0, // KOPIERE HIER STUMPFE DEN WERT AUS "[ORIGINAL-PLAN: ...]"! NICHT ÄNDERN!
      "empfohlener_load_ess": ${todayPlan_LoadRec},
      "empfohlene_zone": "${todayPlan_Zone}", 
      "beispiel_training_1": "${todayPlan_Sport1}",
      "beispiel_training_2": "${todayPlan_Sport2}",
      "prognostizierte_auswirkung": "${todayPlan_Text}",
      "wetterempfehlung": "Wetter egal."
    }
    /* ... 7 weitere Tage ... */
  ]
}
```

---

## Änderung 4 – `ckSup_fixOriginalLoads_`: Spalten für Sperre, Zone, Sport

**Streiche**

```javascript
  const eIdx = head.indexOf('coachE_ESS_day');
  if (dIdx < 0 || eIdx < 0) return apiResponse;
```

**Ersetze durch**

```javascript
  const eIdx = head.indexOf('coachE_ESS_day');
  const fIdx = head.indexOf('fix');
  const zIdx = head.indexOf('Zone');
  const sIdx = head.indexOf('Sport_x');
  if (dIdx < 0 || eIdx < 0) return apiResponse;
```

---

## Änderung 5 – `ckSup_fixOriginalLoads_`: Werte lesen

**Streiche**

```javascript
  const ess = sh.getRange(2, eIdx + 1, n, 1).getValues();
```

**Ersetze durch**

```javascript
  const ess = sh.getRange(2, eIdx + 1, n, 1).getValues();
  const fix = fIdx >= 0 ? sh.getRange(2, fIdx + 1, n, 1).getValues() : null;
  const zon = zIdx >= 0 ? sh.getRange(2, zIdx + 1, n, 1).getDisplayValues() : null;
  const spo = sIdx >= 0 ? sh.getRange(2, sIdx + 1, n, 1).getDisplayValues() : null;
  const locked = {};
```

---

## Änderung 6 – `ckSup_fixOriginalLoads_`: gesperrte Tage merken

**Streiche**

```javascript
    if (d instanceof Date) map[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = Number(String(ess[i][0]).replace(',', '.')) || 0;
```

**Ersetze durch**

```javascript
    if (!(d instanceof Date)) continue;
    const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    map[k] = Number(String(ess[i][0]).replace(',', '.')) || 0;
    if (fix && String(fix[i][0]).trim().toLowerCase() === 'x') {
      locked[k] = { zone: zon ? String(zon[i][0] || '') : '', sport: spo ? String(spo[i][0] || '') : '' };
    }
```

---

## Änderung 7 – `ckSup_fixOriginalLoads_`: gesperrte Tage durchsetzen

**Streiche**

```javascript
    if (key in map && Number(day.original_load_ess) !== map[key]) { day.original_load_ess = map[key]; changed++; }
  });
```

**Ersetze durch**

```javascript
    if (key in map && Number(day.original_load_ess) !== map[key]) { day.original_load_ess = map[key]; changed++; }
    // Gesperrte Tage (fix = x) bleiben wie geplant – die KI darf sie nicht verändern
    if (locked[key]) {
      const lk = locked[key];
      if (Number(day.empfohlener_load_ess) !== map[key]) { day.empfohlener_load_ess = map[key]; changed++; }
      day.empfohlene_zone = map[key] > 0 ? (lk.zone || day.empfohlene_zone || '') : 'Ruhetag (Fixed)';
      if (map[key] <= 0) { day.beispiel_training_1 = 'Geplanter Ruhetag'; day.beispiel_training_2 = 'Mobility / Spaziergang'; }
      else if (lk.sport) day.beispiel_training_1 = lk.sport + ' (fest geplant)';
    }
  });
```

---

## Danach

Speichern, neue Version bereitstellen, kurz Bescheid geben – ich starte einen Supervisor-Lauf und prüfe, dass der Plan zurück ist und Fr 23.10. bei 0 ESS bleibt.
