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

## Patch C eingespielt (20.11.2026, 19:16)

Patch C ist eingespielt (Marc: „Bitte wende Patch C selbst an“). Er ist rein additiv: keine Zeile entfernt, 94 Zeilen ergänzt.

**Sicherung:**
- Vorher: `apps-script-backup/2026-11-20-vor-patch-c/`, mit allen 5 Skriptdateien und 10 HTML-Dateien, vollständig.
- Nachher: `apps-script-backup/2026-11-20-nach-patch-c/`.

**Änderungen in `KiraGeminiSupervisor.gs`:**
1. **Neue Funktion `ckKern_votumProtokoll_(ss)`.** Sie läuft am Ende jedes Supervisor-Laufs, direkt vor END_OK, und ist in try/catch gekapselt. Ein Fehler bricht den Supervisor deshalb nie ab. Sie arbeitet in drei Schritten:
   - Sie liest das Kern-Votum (Cockpit-Regel) aus `coachkira-v2 …/api/rechenkern/votum`.
   - Sie liest die eigene, gerade geschriebene Ampel aus `AI_REPORT_STATUS`.
   - Sie schreibt eine Zeile je Tag in das **neue Blatt `AI_KERN_VOTUM`**. Läuft der Supervisor mehrmals am Tag, gilt der letzte Lauf.

   Spalten: Datum, Kern-Votum, Kern-Stufe, Grund, Begründung, Supervisor-Ampel, Supervisor-Stufe, Gesamtscore, Recovery Score, Training Score, Plan Status, Urteil, Zeit, Quelle.
2. **Neue Route `doGet ?mode=kernVotum[&tage=120]`**, nur lesend.
3. **`ckKern_votumJetzt()`** zum manuellen Test im Editor.

**Änderung in `appsscript.json`:** `https://coachkira-v2.pplx.app/` steht jetzt in `urlFetchWhitelist`. Ohne den Eintrag meldet Google: „URL nicht auf die Zulassungsliste gesetzt“.

**Bereitstellung:**
- Nur die v2-Bereitstellung (ID `AKfycbwfLq…`) läuft jetzt auf **v981**. Im Dialog heißt sie „Unbenannt“, neue Beschreibung: „v2 Schattenbetrieb – Patch C (Kern-Votum)“.
- Live `AKfycbxCEN…` bleibt auf **v978** und kennt `mode=kernVotum` nicht. Das ist geprüft.

**Test:**
- `ckKern_votumJetzt` hat um 19:15 diese Zeile geschrieben: 2026-11-20 · REST · readiness · GELB · 58 · cockpit_strenger.
- `mode=kernVotum` liefert die Zeile.
- v2 liest seine Historie jetzt aus diesem Blatt. Sie bleibt deshalb auch nach einer Neuveröffentlichung erhalten.

**Bitte beachten:** Der Supervisor-Trigger führt den gespeicherten Code aus, auch für Live. Zwei Dinge sind bei Live deshalb neu:
- Ein zusätzlicher Abruf von wenigen Sekunden am Ende jedes Laufs.
- Das Schreiben ins neue Blatt `AI_KERN_VOTUM`.

Was Live anzeigt, bleibt unverändert, denn kein bestehendes Blatt und keine bestehende Route wurde verändert.

**Rückweg:**
- Den Inhalt von `2026-11-20-vor-patch-c/KiraGeminiSupervisor.gs` und `appsscript.json` zurück in den Editor kopieren und speichern.
- Das Blatt `AI_KERN_VOTUM` löschen.

## Stand Etappe 2

| Teil | Stand |
|---|---|
| Eine Regel (Cockpit-Regel) im Kern | erledigt, 452 Tage identisch |
| Cockpit, Widgets und Server nutzen sie | erledigt |
| Supervisor protokolliert Kern-Votum und eigene Ampel (Patch C) | erledigt, ab heute dauerhaft im Sheet |
| v2 zeigt die Supervisor-Ampel nur als Zweitmeinung | erledigt |
| Abnahme: 14 Tage lückenlos, jede Abweichung mit Grund | läuft, 1/14 |
| KI-Texte (Briefing, Analyse) richten sich nach dem Kern-Votum | umgesetzt 20.11.2026, siehe unten (Schalter unter Technik) |
| Soll-ESS aus dem Kern (2b) | Schattenbetrieb seit 20.11.2026, siehe unten |

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
2. **Patch C:** erledigt, siehe oben. Offen ist noch, dass die KI-Texte bei der Umschaltung dem Kern-Votum folgen. Der frühere Plan dazu lautete:
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

## Abnahme (Stand 20.11.2026, 19:45)

- **Rückwirkend geprüft:**
  - Etappe 1: Alle 17 Formeln stimmen über die gesamte Historie mit dem Sheet überein, das sind 3.422 Werte ohne Abweichung. Die Plan-Kette trifft 13/13 Tage.
  - Etappe 2: Das Votum ist über 452 Tage vor und nach dem Umzug identisch.
- **Dauerbetrieb:**
  - Die tägliche Prüfung läuft als Automation „Coach Kira – Abnahme Etappe 1+2 (täglich)“, jeden Tag um 13:30 Uhr vom 21.11. bis 04.12.2026.
  - Sie meldet sich nur bei einem Fehler, sonst erst am 04.12. mit dem Abnahmeprotokoll.
  - Die Automation prüft selbst und zählt unabhängig von Neuveröffentlichungen.
- **Formales Urteil:** Am 04.12.2026, wenn 14 Tage in Folge bestanden sind.

## Umschaltung der KI-Texte auf das Kern-Votum (20.11.2026, Abend)

Marc hat entschieden, nicht bis zum 04.12. zu warten, sondern mit Rückfallweg sofort umzuschalten.

- **Was sich ändert:** Briefing (Status, Plan, Konsolidiert), Supervisor-Analyse und Telegram bekommen den Block „TAGESVOTUM (MASSGEBLICH …)“ mit Votum und Gründen. Der Gesamtscore bzw. die Supervisor-Ampel ist im Prompt nur noch Zweitmeinung. Der Text soll einmal „Grundlage: Tagesvotum …“ nennen.
- **Rückfall:** Fehlt für heute ein Tagesvotum, z. B. weil die Morgenwerte noch nicht erfasst sind, gilt wie bisher die Ampel, und der Text nennt „Grundlage: Supervisor-Ampel“.
- **Schalter:** Technik › Kira-Texte › „Texte nach Tagesvotum (Kern) statt Supervisor-Ampel“, gespeichert als `texte_einstellung.votumQuelle` (`kern` | `ampel`, Standard `kern`). Er wirkt sofort und ohne neue Veröffentlichung auf Mac und Server.
- **Code:**
  - `src/lib/kernVotumKi.ts`: Tagesvotum mit Rückfall. Nutzt `buildDays`, die Regel ist dieselbe wie im Kern.
  - `src/lib/kiraPrompts.ts`: `votumBlock`. Fehlt `kernVotum`, bleibt der Prompt unverändert.
  - `server/kira_prompts_cli.ts`: rechnet das Votum aus der Timeline. Den Schalter übergibt der Server, der Mac liest ihn aus `/api/texte/status`. Neue Art `votum` zum Prüfen.
  - `server/texte.mjs`: Cloud-Ersatz mit Schalter.
  - `src/lib/briefingLocal.ts`: Briefing aus dem Cockpit.
- **Mac:** `~/coach-kira-runner/kira_prompts.mjs` ist ersetzt. Prüfsumme 1a88e006…, gleich der Live-Kopie. Die Sicherung liegt in `kira_prompts.vor-umschaltung.mjs`.
- **Nachweis:**
  - Der Befehl `votum` liefert am 20.11. REST mit denselben Gründen wie `/api/rechenkern/votum`.
  - Die Prompts zeigen den Block in allen drei Arten, in der Stellung „ampel“ den Rückfall-Hinweis.
  - Der Schalter ist lokal getestet: Ohne Sitzung kommt 401, mit Sitzung lässt er sich umstellen, der Cloud-Ersatz bleibt unberührt.
- **Rückweg:** Den Schalter ausschalten. Oder auf dem Mac `kira_prompts.vor-umschaltung.mjs` zurückkopieren und live die vorige Version veröffentlichen.
- **Offen:** Der erste echte Lauf auf dem Mac. Im Mac-Protokoll steht die Prompt-Länge des Supervisors: am 20.11. waren es 27.573 Zeichen, mit dem Block etwa 700 mehr. Die tägliche Abnahme notiert die Texte-Quelle.

## Etappe 2b – Soll-Rahmen aus dem Kern (Schattenbetrieb v2, 20.11.2026)

- **Befund:** `coachE_ESS_day` in `timeline` ist keine Formel, sondern ein Planwert. Ihn schreiben Plan, KI-Plan oder Sperren; für vergangene Tage kopiert ihn `copyObservedToForecast` aus der Ist-Last. Der Kern kann diesen Wert nicht nachrechnen und soll ihn auch nicht ersetzen, denn er ist deine Planung.
- **Kern:** `src/kern/soll.ts` mit `lastEmpfehlung`, unverändert aus `RecoveryPanel.tsx` verschoben. Das Band je Votum wird mit der Last für „Fitness halten“ multipliziert, die Obergrenze ist durch ACWR ≤ 1,3 gedeckelt. Dazu kommt `lageImRahmen`.
- **Nachweis Umzug:** Über 451 Tage (26.08.2025–20.11.2026) ist das Ergebnis identisch mit der bisherigen Cockpit-Funktion, an 429 Tagen gibt es einen Rahmen.
- **Schattenabgleich:**
  - `/api/rechenkern/soll` in v2 liefert den Rahmen, den Plan heute (Plan-Simulation bzw. `coachE_ESS_day`), die Lage und die Ist-Last.
  - Der Server schreibt ab 12 Uhr täglich `soll_historie` mit.
  - Der Rückblick über 30 Tage zeigt die Ist-Last gegen den Rahmen, am 20.11.: 12 im Rahmen, 8 darüber, 9 darunter.
  - In v2 zeigt Technik die Kachel „Soll-Abgleich“.
- **Erster Tag (20.11.):** Votum Rest, Rahmen 0–20 ESS, Plan 0 ESS (Off), also im Rahmen.
- **Live:** unverändert. Die tägliche Abnahme beobachtet 2b mit, als Beobachtung ohne Abnahmekriterium.
