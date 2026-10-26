# Apps-Script-Patch: KI-frei, Stufe 3 (Prompt-Bau raus, Telegram lokal)

Stand 26.10.2026 · Grundlage: die bereitgestellte Stufe-1+2-Fassung (`KiraGeminiSupervisor-KI-frei-Stufe1-2.gs`).

## Was der Patch macht

Danach enthält das Apps Script **keinen Prompt-Text und keine KI-Logik mehr**. Es liefert nur noch Daten, speichert Ergebnisse und verschickt Telegram-Nachrichten.

| Bereich | Bisher (Stufe 2) | Danach (Stufe 3) |
|---|---|---|
| Supervisor-Auftrag (`mode=supervisorJob`) | fertiger Prompt + Schema aus dem Apps Script | **Datenpaket** (`job.data`) – den Prompt baut der Mac mit `kira_prompts.mjs` |
| Briefings Status / Plan / Konsolidiert | `mode=briefingPrompt` liefert fertige Nachrichten | **`mode=briefingData`** liefert nur Daten; Prompt baut Cockpit bzw. Mac (gleiche Datei `kiraPrompts.ts`) |
| KI-Wahl beim Supervisor-Start im Cockpit | immer lokal | Schalter **lokal/Cloud** wird mitgegeben (`engine`); automatische Läufe (Eingabe) bleiben lokal |
| Telegram | nur `/status` | zusätzlich **freie Fragen**: Bot antwortet „Kira denkt nach…“, der Mac beantwortet lokal mit coach-kira |
| Prompt-Funktionen | `createMasterGeminiPrompt`, `getDeepKnowledgeBase`, `WISSENSBLOCK_V105`, Monotonie-Regeltext, Treffer-/Planungsblock, `ckSup_schema_`, `ckDash_buildBriefingMessages_`, `createDashboardPrompt_V39` | **entfernt** (liegen 1:1 in `kiraPrompts.ts`) |

Geprüft: Für dieselben Daten erzeugt der neue Weg **zeichengenau denselben** Supervisor-Prompt wie bisher (vier Szenarien: voll, leer, RTP, krank) und **identische** Briefing-Nachrichten.

Unverändert: alle Daten- und Speicherfunktionen (`saveSupervisor`, `saveBriefing`, `chatLog`, Extras, Garmin), `/status`, Wetter, Kalender, Formular-Eingang.

Neue Skripteigenschaften legt das Script selbst an: `CK_SUP_ENGINE_NEXT` (kurzlebig), `CK_TG_QUEUE`, `CK_TG_VERLAUF`.

## Einspielen (empfohlen: komplette Datei)

1. In `KiraGeminiSupervisor.gs` alles markieren (Cmd+A), löschen, Inhalt von **KiraGeminiSupervisor-KI-frei-Stufe3.gs** einfügen, speichern.
2. **Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Version: Neue Version → Bereitstellen.** URL bleibt gleich.
3. Optional aufräumen: auf der Trigger-Seite die vier deaktivierten Zeit-Trigger löschen (`runDataRefreshOnly`, `generateActivityReview`, `runKiraGeminiSupervisor`, `runHistoricalAnalysis`). Den Trigger **onFormSubmit** stehen lassen.

Mac und Cockpit sind schon umgestellt und funktionieren mit Stufe 2 **und** Stufe 3 – du musst dich nicht beeilen. Nach der Bereitstellung schalten sie automatisch auf den neuen Datenweg.

## Die Änderungen im Einzelnen (falls du lieber vergleichst)

1. **`runKiraGeminiSupervisor()`** – statt `const prompt = createMasterGeminiPrompt(` jetzt `const promptData = ckSup_promptData_(` (gleiche Argumente); `ckSup_enqueue_(prompt, …)` → `ckSup_enqueue_(promptData, …)`.
2. **Neu `ckSup_promptData_(…)`** direkt vor `ckSup_enqueue_` – sammelt heute, Plan, Baseline, Historie, Scores, Faktoren, Wetter, TE-Balance (AI_DATA_FORECAST), aktive KK_WISSEN-Texte, letzte Treffsicherheit, Ziele und Saisonwochen. Nur Daten.
3. **`ckSup_enqueue_(promptData, ctx, runId)`** – Auftrag enthält `data` und `engine` statt `system`/`prompt`.
4. **`ckSup_getJob_()`** – liefert `{ id, createdAt, engine, data }`.
5. **`ckDash_queueAction_()`** – merkt sich beim Supervisor-Start `engine` (lokal/Cloud) aus dem Cockpit.
6. **`ckDash_briefingPrompt_`** → kurzer Hinweis; **neu `ckDash_briefingData_`**; in `doGet` zwei neue Zeilen: `briefingData`, `telegramJobs`.
7. **`handleCommand()`** – freie Texte (ohne `/`) gehen an `ckTg_enqueue_`.
8. **`doPost()`** – neuer Zweig `kind: telegramAnswer` vor `chatLog`.
9. **Neu am Dateiende:** `ckTg_read_`, `ckTg_send_`, `ckTg_enqueue_`, `ckTg_jobs_`, `ckTg_answer_` (nur Marcs Chat-ID, max. 10 offene Fragen, Antwort ohne Markdown, Frage+Antwort auch in `AI_CHAT_LOG` als „Telegram“).
10. **Entfernt:** die oben genannten Prompt-Funktionen und `WISSENSBLOCK_V105`.

## Prüfen

- `…/exec?mode=briefingData&type=status` liefert `{ ok: true, daten: { status: … } }`.
- Cockpit → KI-Report (Supervisor) starten: Mac-Protokoll zeigt „Auftrag … übernommen (lokal, Prompt … Zeichen)“; mit Schalter Cloud „(Cloud, …)“.
- Telegram: eine Frage ohne `/` schicken, z. B. „Kann ich morgen Intervalle machen?“ → sofort „Kira denkt nach…“, nach 1–2 Minuten Kiras Antwort (der Mac muss wach sein).
- Status-, Plan- und konsolidiertes Briefing im Cockpit (lokal und Cloud) laufen wie gehabt.
