# Etappe 2 – Ein Votum statt zwei

Stand 20.11.2026. Nur v2 (coachkira-v2.pplx.app), nur lesend. Live bleibt unverändert.

## Ausgangslage: drei Stellen, zwei Regelwerke

| Stelle | Wo | Regel | Ergebnis |
|---|---|---|---|
| Cockpit „Heute“ | `src/screens/RecoveryPanel.tsx` → `decideVote`/`buildDays` | Readiness-Basis, Warnsignale (HRV, RHR, Schlaf), ACWR, Monotonie, Befinden | Quality / Train / Easy / Rest |
| Widgets (iPhone, iPad, Mac) | `server/widget.mjs` | **dieselbe** Funktion wie das Cockpit | wie Cockpit |
| Supervisor | Apps Script `KiraGeminiSupervisor.gs` → `AI_REPORT_STATUS` | eigene Scores (Recovery, Training, Gesamt) | GRÜN / GELB / ROT / LILA + Plan-Status |

Cockpit und Widgets sind schon einheitlich. Der Supervisor rechnet seine Ampel getrennt davon. Status, Briefing und Plan lesen seine Ampel aus `AI_REPORT_*`.

## Wichtiger Befund

- `AI_REPORT_STATUS` enthält nur den Stand des Tages. Der Supervisor überschreibt ihn bei jedem Lauf.
- `AI_REPORT_HISTORY` wird seit dem 13.02.2026 nicht mehr aktualisiert.
- Deshalb kann niemand nachträglich Tag für Tag vergleichen. Die Reihe muss **ab jetzt** entstehen.
- Am **20.11.2026** zeigt das Cockpit „Rest“ (rot): Readiness 32, zwei Warnsignale. Der Supervisor zeigt am selben Tag „GELB · 58“ mit Plan-Status INTERVENTION. Sein eigener Text spricht trotzdem von „kritisch“. Die beiden Voten liegen also eine Stufe auseinander.

## Was jetzt gebaut ist

- `src/kern/votum.ts`: Über diese Stelle laufen alle Votum-Funktionen. Sie definiert eine gemeinsame Stufenleiter: rot 0, gelb 1, grün 2, blau/lila 3. Dazu kommt die Funktion `vergleiche()`. Die Widgets nutzen den Einstieg bereits.
- `server/votum.mjs`: Rechnet das Cockpit-Votum mit derselben Funktion und liest die Supervisor-Ampel lesend über `mode=status` aus „v2 Schattenbetrieb“.
- Ein täglicher Lauf ab 12 Uhr startet nach Morgenformular und Supervisor-Lauf. Er schreibt nur, wenn der Supervisor heute mit END_OK fertig ist. Speicherort ist `vt_historie` in der v2-Datenbank.
- `GET /api/rechenkern/votum` liefert den heutigen Vergleich, die letzten 30 Tage und die Trefferquote.
- In v2 steht unter Technik die Kachel **„Votum-Abgleich“** mit beiden Voten, dem Urteil und den Tagesfeldern.

**Apps Script:** Patch C ist nicht nötig, weil die vorhandene Route `mode=status` reicht.

## Entscheidung und Umsetzung (20.11.2026, Abend)

**Entscheidung (Marc):** Es gilt die Cockpit-Regel.

**Umgesetzt in v2:**
- Die Regel liegt jetzt in `src/kern/votumRegeln.ts`. Dorthin sind `decideVote`, `buildDays`, die Stufen und Schwellen umgezogen.
- `RecoveryPanel.tsx` liest die Regel nur noch aus dem Kern. Cockpit, Widgets und Server nutzen dieselbe Datei.
- Gegenprobe: Ich habe alle 452 Tage vom 26.08.2025 bis 20.11.2026 vor und nach dem Umzug gerechnet. Votum, Grund, Begründung und Stufen der Einzelwerte sind **byte-identisch**.
- In der Anzeige von v2 heißt die Supervisor-Ampel jetzt „Supervisor-Score · Zweitmeinung“, der Heartbeat-Eintrag „Letzter Supervisor-Lauf“. Das Tagesvotum kommt allein aus der Cockpit-Regel.
- Die Kachel „Votum-Abgleich“ schreibt die Zweitmeinung weiter mit. So sieht man, wann und warum beide Voten auseinanderlaufen.

**Apps Script ist bewusst nicht geändert.** Der Supervisor läuft über einen Zeit-Trigger. Ein Trigger führt immer den zuletzt *gespeicherten* Code aus, nicht eine Bereitstellung. Deshalb würde jede Änderung im Editor sofort auch Live treffen, nicht nur „v2 Schattenbetrieb“. Patch C kommt erst mit der Umschaltung.

## Warum Gelb, obwohl das Cockpit Rest sagt (20.11.2026)

Der Supervisor kennt kein Veto. Seine Gesamtampel ist ein gewichteter Mittelwert der gültigen Einzelscores (`calculateGesamtScore`), heute 10, weil Protein-Invest offen ist. HRV, Readiness und KEI zählen doppelt. Die Grenzen sind: unter 50 rot, unter 80 gelb, sonst grün.

| Einzelscore | Wert | Score | Gewicht |
|---|---|---|---|
| RHR | 53 | 100 (LILA) | 1 |
| Schlafdauer | 7,2 h | 85 | 1 |
| Schlafscore | 64 | 64 | 1 |
| HRV | 38 ms | 50 | 2 |
| Training Readiness | 35 | 35 | 2 |
| TE-Balance | 26,6 % | 100 (LILA) | 1 |
| ACWR | 0,72 | 10 | 1 |
| Training Status | – | 70 | 1 |
| 7-Tage-Bilanz Ernährung | – | 100 | 1 |
| KEI | −146 | 25 | 2 |

Daraus ergibt sich (Summe Score × Gewicht) / Summe Gewichte = 749 / 13 = **57,6 → 58 → GELB**.

Gute Werte wie Ruhepuls, TE-Balance, Schlafdauer und Ernährung gleichen die schlechte Readiness rechnerisch aus. Die Cockpit-Regel geht anders vor. Sie startet bei der Readiness (32, also Basis Rest), und Warnsignale können das Votum nur senken.

**Off-Tag:** Heute ist ein Off-Tag (Sport Off, ESS 0). Er ist **nicht** der Grund für das Gelb. Er drückt den Supervisor-Wert sogar nach unten, weil das niedrige ACWR von 0,72 dort mit 10 Punkten bewertet wird. Das Cockpit-Votum Rest passt zum Off-Tag.

**Readiness 35 statt 32:** Der Supervisor ist um 11:28 Uhr gelaufen und hat zu diesem Zeitpunkt den Wert 35 gesehen. Das Cockpit nutzt den späteren Wert 32.

## Nächste Schritte

1. Zweitmeinung 14 Tage weiter mitschreiben. Das dient als Beleg und als Prüfung, ob Kern-Votum und Daten stabil sind.
2. **Patch C bei der Umschaltung:**
   - Der Supervisor übernimmt das Kern-Votum über `GET /api/rechenkern/votum` als Zeile „Tagesvotum“ in `AI_REPORT_STATUS`.
   - Die KI-Texte sollen sich daran ausrichten statt an der Mittelwert-Ampel.
   - Patch C kommt als Streiche-Ersetze-Patch und wird gemeinsam mit der Umschaltung von Live eingespielt.
3. Soll-ESS aus dem Kern ableiten (Etappe 2b).

## Ursprünglicher Plan für Etappe 2

1. **Beobachten (14 Tage):**
   - Täglich beide Voten mitschreiben.
   - Abweichungen nach Grund ordnen (`bind`: readiness, hrv, acwr, mono …).
2. **Entscheiden (du):** Welche Regel gilt künftig? Mein Vorschlag ist die Cockpit-Regel. Sie ist nachvollziehbar, sie ist dein Erfahrungsmodell, und Widgets und Heute nutzen sie schon.
3. **Regel in den Kern verschieben:**
   - `decideVote`/`buildDays` physisch nach `src/kern/votum.ts` verschieben. Das Frontend greift nur noch über den Kern zu.
   - Prüfen, dass jeder Tag der Historie dasselbe Votum ergibt.
4. **Supervisor anbinden:** Dafür ist eine zusätzliche Route im Kern nötig, die das Kern-Votum liefert.
   - Patch C, additiv: Der Supervisor schreibt dieses Votum als zusätzliche Zeile „Kern-Votum“ in `AI_REPORT_STATUS`.
   - Die alte Ampel bleibt daneben bestehen.
   - Zuerst nur in „v2 Schattenbetrieb“. Live erst, wenn du umschaltest.
5. **Soll-ESS:** Danach leitet der Kern das Tages-Soll ab. Das betrifft die Plan-Simulation und das Widget „Soll heute“. Auch das läuft zuerst im Schattenbetrieb gegen die Sheet-Werte.

## Abnahmekriterien Etappe 2

- 14 Tage lang tägliche Einträge ohne Lücke.
- Jede Abweichung hat einen benannten Grund.
- Nach deiner Entscheidung zeigen Cockpit, Widgets und Supervisor in v2 dieselbe Stufe.

## Rückweg

- Alles ist additiv. Die Kachel und die Route lassen sich ohne Folgen löschen.
- Live (v978, coachkira.pplx.app) ist nicht berührt.
