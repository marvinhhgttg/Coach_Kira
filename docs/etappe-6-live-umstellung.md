# Etappe 6 – Live auf den Kern umstellen

Stand 20.11.2026. **Umgesetzt am 20.11.2026, abends.** Marc wollte nicht bis zum 04.12. warten („Fang direkt an“). Ergebnis und Nachweise stehen am Ende.

## Ziel

Live (coachkira.pplx.app) läuft auf derselben Codebasis wie v2. Dann gilt:
- Alle Zahlen kommen aus dem Rechenkern (`src/kern/`).
- Der Verlauf kommt aus dem Formelblatt `timeline`, der Wahrheit, statt aus der Wertekopie `KK_TIMELINE`.
- Der einzige Unterschied zwischen Live und v2 ist die Startvariable `CK_INSTANZ`.

v2 bleibt danach als Testinstanz für künftige Umbauten bestehen.

## Voraussetzungen (Abnahme bis 04.12.)

- Etappe 1: 14 Tage in Folge ohne Abweichung, alle 17 Formeln, Kette exakt.
- Etappe 2: 14 Tage Kern-Votum im Blatt `AI_KERN_VOTUM`, jede Abweichung zur Supervisor-Ampel mit Grund.
- Umschaltung der KI-Texte ohne Rückschaltung, Mac-Protokoll ohne „FEHLER beim Prompt-Bau“.

## Was sich live ändert (vorab gemessen am 20.11.2026)

**Code** (Unterschiede Live → v2, ohne Daten):
- **Neu:** `src/kern/` (formeln, votum, votumRegeln, soll), `server/rechenkern*.{ts,mjs}`, `server/prognose.ts`, `server/votum*`, die Kacheln Rechenkern, Votum- und Soll-Abgleich (nur in v2 sichtbar) und `src/lib/v2guard.ts` (nur auf dem v2-Host aktiv).
- **Geändert:**
  - `src/lib/ctlModel.ts`: holt die Formeln aus dem Kern.
  - `RecoveryPanel.tsx`: Regeln und Soll-Rahmen kommen aus dem Kern.
  - `CommandCenter.tsx`: v2-Kacheln.
  - `main.tsx`: v2-Hinweis nur auf dem v2-Host.
  - `server.mjs`: Instanz-Schalter.
  - `texte.mjs`: passiv nur in v2.
  - `widget_cli.ts`: importiert aus dem Kern.

**Zahlen:**
- Das Widget rechnet heute mit Live- und Kern-Code identisch, alle 45 Felder stimmen überein.
- Votum, Formeln und Soll-Rahmen sind über 451 bzw. 452 Tage identisch zur bisherigen Live-Rechnung (Etappe 1, 2 und 2b).

**Datenquelle `KK_TIMELINE` → `timeline`** (451 Tage verglichen):

| Feld | Tage mit anderem Wert |
|---|---|
| Tagesvotum | 0 |
| Tages-ESS (`coachE_ESS_day`) | 17 |
| Readiness, HRV, Ruhepuls, Schlaf, Last, ACWR | je 2 |
| Soll-Rahmen (Empfehlung heute) | 14 |

Die betroffenen Tage zeigen danach den Wert aus `timeline`, also den richtigen.

## Schritte am Umstellungstag

1. **Sichern:**
   - Git-Tag `vor-etappe-6`.
   - Tar-Archiv des Live-Arbeitsordners ohne `node_modules`.
   - Kopie von `data.db` nur lokal, nicht ins Repo.
2. **Zusammenführen:**
   - Den Live-Arbeitsordner auf den v2-Stand bringen. Ausgenommen sind `data.db` und die Startbefehle.
   - Starten mit `CK_SETUP_HASH=… CK_MAC_HASH=… CK_MAC_PFLICHT=an NODE_ENV=production node server.mjs`, wie bisher ohne `CK_INSTANZ`, also live.
3. **Datenquelle umstellen:**
   - Die Umleitung von `mode=chartData` auf `/api/rechenkern/chartData` (Blatt `timeline`) gilt dann auch live, nicht nur in v2. Sie wird dafür aus `v2guard.ts` herausgelöst.
   - Der Schreibschutz bleibt nur in v2.
4. **Routen absichern:**
   - `/api/rechenkern/*` live nur mit Sitzung oder Mac-Schlüssel (Regel in `zugang.mjs`), weil die Abgleiche Rechenzeit kosten.
   - Die täglichen Abgleich-Takte laufen live nicht (`aktiv: false`), sondern nur in v2.
5. **Lokal prüfen, dann veröffentlichen:**
   - Widget-JSON alt gegen neu: Alle Felder müssen gleich sein, außer den 2 bekannten Tagen.
   - Heute, Plan und Analyse als Bildschirmfoto alt gegen neu.
   - Anmeldung, Kopplungen, Mac-Puls und Schalter unter Technik.
6. **Beobachten (7 Tage):**
   - Die tägliche Prüfung vergleicht zusätzlich das Live-Widget mit v2. Beide müssen identisch sein.

## Offen zu entscheiden

- **Soll-Rahmen im Widget:** Neues Feld `sollRahmen` in `/api/widget`, im großen Widget als Zeile „Rahmen 0–20 ESS“. Dafür ist eine kleine Swift-Änderung nötig, der Bau läuft über Xcode.
- **Quelle für `AI_KERN_VOTUM`:**
  - Der Supervisor liest das Kern-Votum heute von coachkira-v2 (Patch C).
  - Vorschlag: so lassen, solange v2 läuft.
  - Später ein additiver Patch E mit einer Skripteigenschaft für die URL, damit live die Quelle sein kann.

## Rückweg

- Den Live-Arbeitsordner aus dem Tar-Archiv bzw. dem Tag `vor-etappe-6` wiederherstellen und veröffentlichen.
- Apps Script und Mac sind in Etappe 6 nicht betroffen.

## Umsetzung (20.11.2026)

**Sicherung:**
- Git-Tag `vor-etappe-6`.
- `backups/live-vor-etappe-6.tgz` (Live-Ordner ohne `node_modules`), `backups/data.db.vor-etappe-6` und `backups/runcmd.vor-etappe-6.txt`, alle nur im Arbeitsbereich, nicht im Repo.

**Code:**
- Der Live-Ordner ist jetzt identisch mit v2. Unterschiede gibt es nur bei der Datenbank und beim Startbefehl: live ohne `CK_INSTANZ`, v2 mit `CK_INSTANZ=v2`.
- `src/lib/api.ts` › `fetchChartData`: liest zuerst `timeline` über `/api/rechenkern/chartData`, bei Fehler oder nach 90 s wie bisher Apps Script `mode=chartData` (`KK_TIMELINE`).
  - Im Browser über `SERVER_BASE`.
  - Auf dem Server über `CK_ZEITREIHE_URL`, die `server.mjs` auf den eigenen Port setzt.
  - `CK_ZEITREIHE_URL=aus` schaltet den alten Weg wieder ein.
- `server/kira_prompts_cli.ts`: dieselbe Quelle, auch auf dem Mac (öffentliche Route des Live-Servers). Das Mac-Werkzeug ist ersetzt, die Sicherung heißt `kira_prompts.vor-etappe-6.mjs`.
- `server/rechenkern_api.mjs`: Gleichzeitige Abrufe von `timeline` teilen sich eine Anfrage. Ist Apps Script gestört, gilt der letzte gute Stand.
- `server/zugang.mjs`: `/api/rechenkern/abgleich` live nur mit Sitzung oder Mac-Schlüssel. `chartData`, `votum` und `soll` bleiben lesbar, genau wie die Apps-Script-Web-App.

**Messung vor dem Umschalten, korrigiert:**
- Die erste Schätzung oben (Tages-ESS 17 Tage, Morgenwerte 2 Tage) betraf nur die Felder für Votum und Soll.
- Der Vergleich über alle 62 Spalten nach Datum zeigt mehr Abweichungen, alle vor dem 10.10.2026. Seitdem sind beide Blätter identisch.

| Spalte | Tage anders |
|---|---|
| `ctl_star`, `ctl_gap` | 310 |
| `coachE_ATL_morning` / `coachE_CTL_morning` | 229 / 226 |
| `coachE_Smart_Gains` | 135 |
| `TE_Balance_Trend` | 122 |
| `Strain7` / `Monotony7` | 76 / 74 |
| `load_star` | 22 |
| `coachE_ESS_day` | 17 |
| Forecast-ATL/CTL/ACWR | 13–14 |
| alle Tageswerte am 28./29.04.2026 | 2 (in `KK_TIMELINE` vertauscht) |

- Das Tagesvotum ist an allen 451 Tagen gleich.

**Nachweise:**
- **Widget:** alt (Live-Sicherung) gegen neu, alle 45 Felder gleich.
- **Zeilenzahl:** `chartData` neu gegen Apps Script für 7, 14, 28, 30, 60, 90, 120 Tage und die ganze Historie, jeweils gleiche Anzahl Zeilen, gleiches erstes und letztes Datum und gleiches Datumsformat.
- **Rückfall:** geprüft mit unerreichbarer URL und mit `aus`, beide Male Apps Script.
- **Ansichten alt gegen neu** (Heute, Plan, Analyse): keine Fehler im Browser. Die Unterschiede liegen nur in historischen Auswertungen, zum Beispiel ESS-Klassen 147–207 → 147–209, HRV-Bänder ±1, Ø Monotonie über 120 Tage 1,59 → 1,54, CTL-Modellfehler ±22,7 → ±19,8.
- **Live danach:**
  - `chartData` mit Quelle `timeline`.
  - PIN-Einrichtung, Mac, Text-Schalter und Warteschlange unverändert.
  - Abgleich ohne Sitzung 401.
  - Votum und Soll-Rahmen live und in v2 identisch (Rest, 0–20 ESS).
- **v2:** Abgleich weiter 17/17, Kette 13/13.

**Beobachtung:** Die tägliche Prüfung vergleicht bis 27.11. zusätzlich live gegen v2 (Votum, Soll-Rahmen, Quelle `timeline`) und meldet Abweichungen.

**Rückweg:**
- Schnell: in `/tmp/runcmd.txt` `CK_ZEITREIHE_URL=aus` voranstellen und neu veröffentlichen. Dann liest live wieder `KK_TIMELINE`, der Code bleibt neu.
- Vollständig: `backups/live-vor-etappe-6.tgz` in den Live-Ordner entpacken und veröffentlichen.
- Mac: `kira_prompts.vor-etappe-6.mjs` zurückkopieren.
