# Apps-Script-Patch: KI-Plan im Supervisor streichen

Stand: 19.10.2026

**Warum**

Den 14-Tage-Plan liefert das **Plan-Briefing** (Plan-Seite, `AI_PLAN_BRIEFING`). Der Supervisor erzeugt zusätzlich einen 8-Tage-Plan (`empfohlener_plan` → `AI_REPORT_PLAN`), den im Cockpit niemand anzeigt:

- Die Plan-Seite nutzt deine eigene Simulation (`planSimulationJson`).
- Die Kalender-Synchronisation (`syncToGoogleCalendar`) liest die timeline.
- `syncPlanToCalendar` (liest `AI_REPORT_PLAN`) wird nirgends aufgerufen.
- Gelesen wird `AI_REPORT_PLAN` nur noch vom Menüpunkt „KI-PLAN ANWENDEN“ und vom alten HTML-Dashboard.

**Effekt**

- Der KI-Report wird kürzer und schneller. Beim lokalen Lauf sind die 8 Tage mit je fünf Textfeldern der größte Teil der Antwort.
- Es gibt keinen zweiten, widersprüchlichen Plan mehr.
- `AI_REPORT_PLAN` wird beim nächsten Lauf geleert. „KI-PLAN ANWENDEN“ meldet dann „Blatt leer“, statt einen veralteten Plan in die timeline zu schreiben.

Unverändert bleiben Gesamtbewertung, Plan-Status (FREIGABE/BEOBACHTUNG/INTERVENTION), Einzelscore-Kommentare und Ernährung.

---

## Änderung 1 – JSON-Schema: `empfohlener_plan` raus

**An zwei Stellen:** in `callOpenAI` (Variable `reportSchema`) **und** in `ckSup_schema_`. Der Text ist an beiden Stellen identisch.

**Streiche**

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

**Ersetze durch**

```javascript
        }
      },
      "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare"],
```

---

## Änderung 2 – `createMasterGeminiPrompt`: Aufgabe 3

**Streiche**

```javascript
3. 'empfohlener_plan': Einen angepassten 8-TAGE-PLAN (HEUTE + 7 TAGE ZUKUNFT).
```

**Ersetze durch**

```javascript
3. KEINEN Trainingsplan: Den 14-Tage-Plan erstellt das Plan-Briefing separat. Gib 'empfohlener_plan' NICHT aus und ignoriere die Plan-Regeln weiter unten; die Prognose-Daten dienen nur als Kontext für deine Gesamtbewertung.
```

---

## Änderung 3 – `createMasterGeminiPrompt`: JSON-Beispiel am Ende

**Streiche**

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

**Ersetze durch**

```javascript
    /* ... und ALLE anderen aus der Liste ... */
  ]
}
```

---

## Änderung 4 – `ckSup_writeReportPlan_`: veralteten Plan leeren

**Streiche**

```javascript
  if (!plan.length) return;
  const toDate
```

**Ersetze durch**

```javascript
  if (!plan.length) {
    // Kein KI-Plan mehr (Plan-Briefing übernimmt): altes Blatt leeren, damit nichts Veraltetes gelesen wird
    const old = ss.getSheetByName(OUTPUT_PLAN_SHEET);
    if (old && old.getLastRow() > 1) old.getRange(2, 1, old.getLastRow() - 1, Math.max(9, old.getLastColumn())).clearContent();
    return;
  }
  const toDate
```

---

## Danach

Speichern, neue Version bereitstellen, kurz Bescheid geben – ich starte einen lokalen Supervisor-Lauf, vergleiche die Laufzeit (heute früh 272 s) und prüfe, dass Scores und Texte weiter ankommen.
