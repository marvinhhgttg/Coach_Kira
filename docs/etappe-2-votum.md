# Etappe 2 – Ein Votum statt zwei (Vorbereitung)

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

## Plan für Etappe 2

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
