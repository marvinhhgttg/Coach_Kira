# Apps-Script-Patch: Briefings lokal oder in der Cloud

Stand: 18.10.2026 · betrifft KI-Analyse (Status), konsolidierte Analyse und Plan-Briefing

**Was der Patch macht**

Der Patch ist rein ergänzend: Die bestehenden Perplexity-Funktionen bleiben unverändert und werden weiter für „Cloud“ genutzt.

Neu dazu kommen drei Modi:

- `briefingPrompt` (GET) liefert den fertigen Prompt mit denselben Daten wie bisher, aber ohne Perplexity-Aufruf. Damit rechnet das Cockpit lokal mit `coach-kira`.
- `saveBriefing` (POST) speichert das lokale Ergebnis über die vorhandenen Schreibfunktionen in `AI_STATUS_ANALYSIS`, `AI_CONSOLIDATED_ANALYSIS` bzw. `AI_PLAN_BRIEFING`. In der Spalte `raw` wird `{"engine":"local","model":"coach-kira"}` vermerkt.
- `briefingSources` (GET) zeigt für jedes Briefing, ob der letzte Eintrag lokal oder aus der Cloud stammt.

Gegenüber den Originalen sind zwei Fehler behoben, aber nur in den neuen Funktionen:

- Die Prompt-Zeilen werden mit echten Zeilenumbrüchen verbunden (`'\n'` statt `'\\n'`).
- Codeblöcke (```json) werden beim Plan-Briefing korrekt entfernt.

---

## Änderung 1 – `doGet`

**Streiche**

```javascript
  e = e || { parameter: {} };
  const p = e.parameter || {};
```

**Ersetze durch**

```javascript
  e = e || { parameter: {} };
  const p = e.parameter || {};

  // Briefings lokal (Cockpit + Ollama)
  if (p.mode === 'briefingPrompt') return ckDash_briefingPrompt_(p);
  if (p.mode === 'briefingSources') return ckDash_briefingSources_();
```

---

## Änderung 2 – `doPost`

**Streiche**

```javascript
  const ckBody = (e && e.postData && e.postData.contents) || '';
```

**Ersetze durch**

```javascript
  const ckBody = (e && e.postData && e.postData.contents) || '';
  if ((e && e.parameter && e.parameter.mode === 'saveBriefing') || /"kind"\s*:\s*"saveBriefing"/.test(ckBody)) {
    return ckDash_saveBriefing_(e);
  }
```

(Falls du den Patch „Kira lokal protokollieren“ schon eingespielt hast: die neuen Zeilen einfach direkt unter die `const ckBody = …`-Zeile setzen, die `chatLog`-Abfrage bleibt darunter.)

---

## Änderung 3 – neue Funktionen (ans Ende der Datei)

```javascript
// ============================================================================
// Briefings lokal oder in der Cloud – Prompt holen, lokales Ergebnis speichern
// ============================================================================
function ckDash_buildBriefingMessages_(type) {
  if (type === 'status') {
    const status = ckDash_getReportStatusSummary_();
    return {
      format: 'text',
      messages: [
        { role: 'system', content: 'Du bist ein präziser deutschsprachiger Ausdauer-, Trainings- und Recovery-Coach.' },
        { role: 'user', content: [
          'Du bist Coach Kira.',
          'Analysiere die folgenden Status-Items knapp, konkret und handlungsorientiert.',
          'Sprich Marc direkt per Du an.',
          'Fokus: Training, Recovery, Schlaf, ACWR, Readiness, Ernährung.',
          'Gib maximal 6 Bulletpoints plus eine klare Empfehlung für heute.',
          '',
          JSON.stringify(status, null, 2)
        ].join('\n') }
      ]
    };
  }

  if (type === 'plan') {
    const status = ckDash_getReportStatusSummary_();
    const sim = ckDash_collectPlanBriefingContext_();
    return {
      format: 'json',
      messages: [
        { role: 'system', content: [
          'Du bist Coach Kira, eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach und Recovery-Analystin.',
          'Du gibst präzise, konservative, handlungsorientierte Empfehlungen für Marc.',
          'Du nutzt ACWR, ATL, CTL, Training Readiness, Schlaf, HRV, RHR, TE und KEI/Smart Gains.',
          'Du sprichst Marc direkt per Du an.',
          'Du antwortest als valides JSON ohne Markdown-Codeblock.'
        ].join(' ') },
        { role: 'user', content: [
          'Erstelle ein Plan-Briefing für das Plan Cockpit.',
          'Nutze Command-Center-Status, aktuelle PlanApp-Daten und Garmin-/Coach-Kira-Kontext.',
          '',
          'Output-Format exakt:',
          '{',
          '  "briefing": "mehrzeiliger deutscher Text mit konkreten Empfehlungen",',
          '  "recommendations": [',
          '    {"date":"YYYY-MM-DD", "day":"Mo", "recommended_load_ess":130, "recommended_zone":"Z2", "rationale":"kurze Begründung"}',
          '  ]',
          '}',
          '',
          'Regeln:',
          '- recommendations muss genau die sichtbaren 14 Tage enthalten.',
          '- ESS nur realistische Werte 0 bis 300.',
          '- Wenn Recovery/Readiness kritisch ist: Last reduzieren.',
          '- Wenn ACWR > 1.3: keine zusätzliche Steigerung.',
          '- Wenn Schlafdauer oder Readiness schlecht ist: Z2 oder Off bevorzugen.',
          '- Keine medizinischen Diagnosen.',
          '',
          'STATUS:',
          JSON.stringify(status, null, 2),
          '',
          'PLAN_CONTEXT:',
          JSON.stringify(sim, null, 2)
        ].join('\n') }
      ]
    };
  }

  if (type === 'consolidated') {
    const context = ckDash_collectConsolidatedContext_();
    const skillPerspectives = [
      { name: '/garmin-gesamtanalyse', role: 'Gesamtbild aus Trainings-, Recovery-, Schlaf-, HRV-, Ernährungs- und Belastungsdaten.' },
      { name: '/garmin-anomalie-erkennung', role: 'Auffälligkeiten, Datenbrüche, unerwartete Sprünge, mögliche Überlastung oder Recovery-Risiken.' },
      { name: '/garmin-ernaehrung-trainingskorrelation', role: 'Zusammenhang zwischen Energie-/Makroaufnahme, Defizit/Überschuss, Load, Recovery und Schlaf.' },
      { name: '/garmin-wochenplanung', role: 'Adaptive 7- bis 14-Tage-Planung mit Load-, Zone- und Ruhetagslogik.' },
      { name: '/garmin-trainingsstatus', role: 'Trainingszustand, ATL/CTL/ACWR, TE Balance, KEI, Monotony/Strain und Fortschrittslogik.' },
      { name: '/garmin-erholungsstatus', role: 'Schlafdauer, Schlafscore, HRV, RHR, Training Readiness, Stress und Erholungsbedarf.' },
      { name: '/garmin-gesundheitsstatus', role: 'Gesundheitsnahe, nicht-diagnostische Risikoeinschätzung und Watchlist.' }
    ];
    return {
      format: 'text',
      messages: [
        { role: 'system', content: [
          'Du bist Coach Kira: eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach, Recovery-Analystin und Garmin-Datenexpertin.',
          'Du synthetisierst mehrere Analyseperspektiven, statt nur Statuswerte nachzuerzählen.',
          'Du sprichst Marc direkt per Du an.',
          'Keine medizinischen Diagnosen. Bei Gesundheitsrisiken vorsichtig formulieren.',
          'Nutze klare Überschriften und Bulletpoints. Keine JSON-Antwort.'
        ].join(' ') },
        { role: 'user', content: [
          'Erstelle eine konsolidierte Coach-Kira-Lagebeurteilung.',
          '',
          'Die folgenden Garmin-Skill-Perspektiven sollen gedanklich verwendet, synthetisiert und zusammengefasst werden:',
          JSON.stringify(skillPerspectives, null, 2),
          '',
          'Pflichtstruktur:',
          '## Gesamtbild',
          '## Auffälligkeiten und Anomalien',
          '## Erholung und Gesundheitsstatus',
          '## Training und Belastungssteuerung',
          '## Ernährung und Trainingskorrelation',
          '## Plan 14 Tage',
          '## Empfehlung für heute',
          '## Watchlist',
          '',
          'Konkrete Regeln:',
          '- Nenne konkrete Werte, wenn sie entscheidungsrelevant sind.',
          '- Gib klare Handlungsprioritäten.',
          '- Weise auf Widersprüche hin, z.B. guter KEI aber niedrige Readiness.',
          '- Verdichte, nicht ausufern.',
          '- Maximal 900 Wörter.',
          '',
          'DATENKONTEXT:',
          JSON.stringify(context, null, 2)
        ].join('\n') }
      ]
    };
  }

  throw new Error('Unbekannter Briefing-Typ: ' + type);
}

function ckDash_briefingPrompt_(p) {
  try {
    const type = String((p && p.type) || '');
    const b = ckDash_buildBriefingMessages_(type);
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), type: type, format: b.format, messages: b.messages });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_parseBriefingJsonLocal_(content) {
  let txt = String(content || '').trim();
  txt = txt.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  try {
    const parsed = JSON.parse(txt);
    return {
      briefing: String(parsed.briefing || ''),
      recommendations: Array.isArray(parsed.recommendations) ? parsed.recommendations : []
    };
  } catch (e) {
    return { briefing: txt, recommendations: [] };
  }
}

function ckDash_saveBriefing_(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const type = String(b.type || '');
    const text = String(b.text || '').slice(0, 30000).trim();
    const model = String(b.model || 'lokal').replace(/[^\w:.\-]/g, '').slice(0, 40);
    if (!text) throw new Error('Kein Briefing-Text übermittelt.');
    const raw = JSON.stringify({ engine: 'local', model: model, savedAt: ckDash_now_() });

    if (type === 'status') {
      ckDash_writeStatusAiAnalysis_(text, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), analysis: text });
    }
    if (type === 'consolidated') {
      ckDash_writeConsolidatedAnalysis_(text, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), analysis: text });
    }
    if (type === 'plan') {
      const parsed = ckDash_parseBriefingJsonLocal_(text);
      ckDash_writePlanBriefing_(parsed.briefing, parsed.recommendations, raw);
      return ckDash_json_({ ok: true, timestamp: ckDash_now_(), briefing: parsed.briefing, recommendations: parsed.recommendations });
    }
    throw new Error('Unbekannter Briefing-Typ: ' + type);
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

function ckDash_briefingSources_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defs = [
      ['status', 'AI_STATUS_ANALYSIS', 3],
      ['consolidated', 'AI_CONSOLIDATED_ANALYSIS', 3],
      ['plan', 'AI_PLAN_BRIEFING', 4]
    ];
    const out = {};
    defs.forEach(function (d) {
      const sh = ss.getSheetByName(d[1]);
      if (!sh || sh.getLastRow() < 2) { out[d[0]] = null; return; }
      const row = sh.getRange(sh.getLastRow(), 1, 1, d[2]).getValues()[0];
      let engine = 'cloud';
      let model = null;
      try {
        const r = JSON.parse(String(row[d[2] - 1] || '{}'));
        if (r && r.engine === 'local') engine = 'local';
        model = (r && r.model) ? String(r.model) : null;
      } catch (e) {}
      out[d[0]] = { engine: engine, model: model, timestamp: row[0] };
    });
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), sources: out });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}
```

---

## Danach

1. Speichern und warten, bis „Gespeichert“ angezeigt wird.
2. Bereitstellen → Bereitstellungen verwalten → Stift → Version „Neue Version“ → Bereitstellen.
3. Kurz Bescheid geben – ich prüfe `?mode=briefingPrompt&type=status` und teste dann alle drei Briefings lokal.
