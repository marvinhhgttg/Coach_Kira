# Apps-Script-Patch: Monotonie-Regel

Stand: 22.10.2026

**Was sich ändert**

Bisher bekamen KI-Report (Supervisor) und Briefings nur die Monotonie-Zahlen, aber keine Regel, wie sie zu bewerten sind. Jetzt erhalten beide denselben Regeltext wie das Cockpit:

- Ein Einzelwert ist meist kein Problem; maßgeblich ist die rollierende 7-Tage-Monotonie im Verlauf.
- Zonen: < 1,5 · 1,5–2,0 · > 2,0 · > 2,5 · > 3,0.
- Muster 1–2 / 3–4 / 5–7 Tage / mehrere Wochen.
- Steuerung: Ruhetag binnen 24–48 h, Entlastungstag, keine Qualität bei schlechter Erholung.
- Strain = Wochenlast × Monotonie.

Keine Änderung an Tabellen, Spalten oder Berechnungen.

1. Neue Funktionen `ckRegel_monotonie_` und `ckRegel_monotonieBlock_`.
2. Der Supervisor-Prompt bekommt den Regelblock direkt nach der Treffsicherheit.
3. Die lokalen Briefings (Status, Plan, konsolidiert) bekommen die Regel an die System-Anweisung angehängt.

---

## Änderung 1 – `createMasterGeminiPrompt`

**Streiche**

```javascript
${ckX_trefferPromptBlock_()}
```

**Ersetze durch**

```javascript
${ckX_trefferPromptBlock_()}
${ckRegel_monotonieBlock_()}
```

---

## Änderung 2 – `ckDash_briefingPrompt_`

**Streiche**

```javascript
    const b = ckDash_buildBriefingMessages_(type);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, format: b.format, messages: b.messages });
```

**Ersetze durch**

```javascript
    const b = ckDash_buildBriefingMessages_(type);
    // Monotonie-Regel an die System-Anweisung anhängen
    if (b.messages && b.messages.length && b.messages[0].role === 'system') {
      b.messages[0].content = String(b.messages[0].content) + '\n\n' + ckRegel_monotonie_();
    }
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, format: b.format, messages: b.messages });
```

---

## Änderung 3 – neue Funktionen (ans Ende der Datei)

```javascript
/** Monotonie-Regel (gleicher Wortlaut wie im Cockpit, lib/monotony.ts) */
function ckRegel_monotonie_() {
  return [
    "Monotonie-Regel: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Maßgeblich ist die rollierende 7-Tage-Monotonie (Ø/σ der Tageslasten) im Verlauf.",
    "Zonen: < 1,5 gute Variation; 1,5–2,0 normal; > 2,0 aufmerksam (vor allem bei hoher Wochenlast); > 2,5 ungünstig, wenn es länger anhält; > 3,0 Warnzone.",
    "Muster: 1–2 ähnliche Tage normal; 3–4 ähnliche, fordernde Tage beobachten (bei schlechterem Schlaf, HRV, Ruhepuls, schweren Beinen oder weniger Motivation reduzieren); 5–7 Tage ohne deutliche Entlastung relevant; mehrere Wochen hohe Monotonie = Risiko für Ermüdung, Stagnation, Überlastung, Infekte.",
    "Steuerung: Über etwa 2,0 und schon 4–5 fordernde Tage → nächster Tag sehr locker oder frei. Bleibt der Wert über 2,5 → innerhalb 24–48 h echter Ruhetag oder sehr kurzer Regenerationsreiz. Über 3,0 → mindestens ein echter Entlastungstag oder deutliche Reduktion. Hohe Monotonie plus schlechte Erholungswerte → keine Qualitätseinheit durchdrücken.",
    "Entscheidend ist die Kombination mit der Gesamtlast (Strain = Wochenlast × Monotonie). Monotonie senkst du durch klaren Rhythmus: harte Reize klar hart, lockere Tage wirklich locker, mindestens ein echter Erholungstag bei hoher Gesamtlast – nicht durch weniger Gesamtlast."
  ].join('\n');
}

/** Block für den Supervisor-Prompt */
function ckRegel_monotonieBlock_() {
  return '\n---\n**MONOTONIE-REGEL (verbindlich für Bewertung und Plan):**\n' + ckRegel_monotonie_() +
    '\nBewerte die Prognose-Monotonie oben nach dieser Regel: ein einzelner hoher Tageswert ist kein Grund zum Kürzen; ' +
    'entscheidend sind fordernde Tage ohne Entlastung, anhaltend hohe Werte und die Wochenlast.\n---\n';
}

```

---

Danach speichern, neue Version bereitstellen, kurz Bescheid geben.
