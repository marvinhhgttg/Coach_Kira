# Etappe 1 – Rechenkern im Schattenbetrieb

Stand 20.11.2026. Baseline: Git-Tag `baseline-2026-11-20` (Commit e1265ec), Sheet-Kopie „Coach Kira – Baseline 2026-11-20“, Apps Script v978.

## Ziel

Alle Formeln (ATL, CTL, ACWR, Monotonie, Strain, fKEI, Smart Gains, Tourform, Votum) laufen an einer Stelle, einem TypeScript-Rechenkern auf dem Server. Apps Script liefert dafür nur Rohdaten.

Diese Etappe ist reiner Schattenbetrieb:
- Das Sheet bleibt die Wahrheit.
- Das Live-Cockpit bleibt unverändert.
- Der Kern rechnet in v2 nur mit und vergleicht jede Zahl mit dem Sheet.

Umgeschaltet wird erst, wenn die Historie ab 26.08.2025 und 14 Tage im laufenden Betrieb ohne Abweichung übereinstimmen.

## Warum so

- Grundsatz: Basis ist die Vergangenheit. Der neue Kern muss die vorhandene Historie exakt nachrechnen, bevor er irgendetwas ersetzt.
- Das Apps-Script-Projekt hat zwei Ausführungswege:
  - **Web-App-Bereitstellungen** laufen auf einer festen Version. Live nutzt v978.
  - **Trigger** (Formular, Supervisor-Takt) laufen immer auf dem aktuellen Code.

  Deshalb sind die Patches dieser Etappe rein additiv. Es kommen eine neue Funktion und eine neue Route dazu, keine bestehende Funktion ändert sich. Die Trigger verhalten sich dadurch weiterhin genauso.

## Bausteine

| Nr. | Baustein | Wo | Wer |
|---|---|---|---|
| 1.1 | Route `mode=rohdaten` (nur lesend) | Apps Script, Patch A | Marc spielt ein |
| 1.2 | Eigene Bereitstellung „v2 Schattenbetrieb“ | Apps Script | Marc |
| 1.3 | Rechenkern `server/rechenkern.ts` (die Formeln aus `timeline`, 1:1 portiert) | Cockpit v2 | lokal (coach-kira) portiert je Formel, Cloud prüft |
| 1.4 | Abgleich `/api/rechenkern/abgleich` und Seite „Abgleich“ unter Technik | Cockpit v2 | Cloud |
| 1.5 | Täglicher Abgleich und Abnahme | v2 | automatisch, nur lesend |

## Abnahmekriterien

- Historie ab 26.08.2025: In jeder Spalte weichen höchstens 0,5 % der Werte ab (absolut 0,05). Jede Abweichung ist erklärt.
- 14 Tage am Stück im laufenden Betrieb ohne neue Abweichung.
- Live-Cockpit, Formular, Garmin-Skills und Trigger zeigen kein verändertes Verhalten.

## Rückweg

- Bereitstellung „v2 Schattenbetrieb“ archivieren. Live läuft weiter auf v978.
- Die Funktion `ckDash_rohdaten_` und die eine Zeile in `doGet` können stehen bleiben oder wieder gestrichen werden. Sie werden nur über die Route aufgerufen.

---

## Patch A – Route `mode=rohdaten` (Apps Script, `KiraGeminiSupervisor.gs`)

Was er tut: Er liefert alle Zeilen aus `KK_TIMELINE` ab dem 26.08.2025, mit allen 62 Spalten als Rohwerte. Dazu kommen alle benannten Einzelzell-Bereiche als Konstanten (z. B. `alpha_CTL`, `S_CTL`, `B_CTL`, `wH_CTL`, `wS_CTL`). Optional (`formeln=1`) liefert er die Formeln einer Zeile aus `timeline` als Vorlage für die Portierung.

Er schreibt nichts, und keine bestehende Funktion ändert sich.

### A1 – in `doGet`

Streiche:
```js
  if (p.mode === 'loadFactors') return ckX_loadFactors_();
```

Ersetze durch:
```js
  if (p.mode === 'loadFactors') return ckX_loadFactors_();
  // Etappe 1: Rohdaten für den Rechenkern (nur lesend)
  if (p.mode === 'rohdaten') return ckDash_rohdaten_(p);
```

### A2 – neue Funktion direkt vor `ckDash_chartData_`

Streiche:
```js
function ckDash_chartData_(p) {
```

Ersetze durch:
```js
/**
 * Etappe 1 (Schattenbetrieb): Rohdaten + Konstanten für den Rechenkern im Cockpit v2.
 * Nur lesend. Parameter: ab=yyyy-MM-dd (Standard 2025-08-26), formeln=1, zeile=<Zeilennummer>.
 */
function ckDash_rohdaten_(p) {
  try {
    p = p || {};
    const ab = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ab || '')) ? String(p.ab) : '2025-08-26';
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ckDash_getSheet_('KK_TIMELINE');
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(function (h) { return String(h || '').trim(); });
    const dateIdx = headers.indexOf('date');
    if (dateIdx < 0) throw new Error('Spalte date fehlt in KK_TIMELINE');

    const wert = function (v) {
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      if (v === '' || v === null || v === undefined) return null;
      return v;
    };

    const zeilen = [];
    for (let i = 1; i < values.length; i++) {
      const d = values[i][dateIdx];
      if (!(d instanceof Date)) continue;
      if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') < ab) continue;
      zeilen.push([i + 1].concat(values[i].map(wert)));
    }

    const konstanten = {};
    ss.getNamedRanges().forEach(function (nr) {
      const r = nr.getRange();
      if (r.getNumRows() === 1 && r.getNumColumns() === 1) konstanten[nr.getName()] = wert(r.getValue());
    });

    const out = {
      ok: true,
      timestamp: ckDash_now_(),
      blatt: 'KK_TIMELINE',
      ab: ab,
      spalten: ['row'].concat(headers),
      zeilen: zeilen,
      konstanten: konstanten
    };

    if (String(p.formeln || '') === '1' && zeilen.length) {
      const tl = ckDash_getSheet_('timeline');
      const n = Number(p.zeile) || zeilen[zeilen.length - 1][0];
      out.formeln = {
        zeile: n,
        spalten: tl.getRange(1, 1, 1, tl.getLastColumn()).getValues()[0],
        werte: tl.getRange(n, 1, 1, tl.getLastColumn()).getFormulas()[0]
      };
    }
    return ckDash_json_(out);
  } catch (err) {
    return ckDash_json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function ckDash_chartData_(p) {
```

### Einspielen, ohne Live zu berühren

1. Patch A1 und A2 einfügen und speichern. Die Trigger bleiben unverändert, weil keine bestehende Funktion geändert ist.
2. **Bereitstellen → Neue Bereitstellung** (nicht „Bereitstellungen verwalten“):
   - Typ: Web-App
   - Beschreibung: „v2 Schattenbetrieb“
   - Ausführen als: ich
   - Zugriff: wie die Live-Bereitstellung
3. Die neue `/exec`-URL an mich geben. Nur coachkira-v2 nutzt sie.
4. Kontrolle: Die Live-Bereitstellung „v978 – Perplexity Dashboard“ steht weiter auf Version 978.

Test im Browser: `<v2-URL>?mode=rohdaten&formeln=1` liefert `ok: true`, mehr als 400 Zeilen und `formeln.werte` mit den `coachE_*`-Formeln.

---

## Ergebnis nach Patch A (20.11.2026)

- Die Bereitstellung „v2 Schattenbetrieb“ (`AKfycbwfLq…`) liefert 582 Zeilen ab 26.08.2025 und 12 Konstanten. Die Live-Bereitstellung kennt die Route nicht und läuft also weiter auf v978.
- Im Sheet stehen 11 Spalten als Formeln. Die Werte für ESS, ATL, CTL und ACWR samt Prognosen schreibt Apps Script selbst. Sie sind Eingaben für den Kern, keine Formeln.
- Rechenkern `server/rechenkern.ts`: Den Entwurf hat das lokale Modell coach-kira geschrieben, die Cloud hat ihn korrigiert. Gegen das Formelblatt `timeline` stimmen **alle 11 Formeln zu 100 %**, geprüft werden alle Zellen mit der heutigen Formelversion.
- Befunde (bewusst 1:1 übernommen, nichts geändert):
  1. **Zwei Wahrheiten im Sheet:** Die Wertekopie `KK_TIMELINE`, aus der das Cockpit liest, weicht vom Formelblatt `timeline` ab. Beim Tages-ESS betrifft das 17 von 452 Tagen, bei ctl_star 255 Tage.
  2. **Ältere Formelversionen:** Historische Zeilen haben teils andere Formeln, zum Beispiel greift die CTL-Morgenformel Anfang August 2026 auf den Vortag zu. Diese Werte gelten als Wahrheit der Vergangenheit.
  3. **Strain7** nutzt ein Fenster, das gestern endet: `ROWS($AF$2:…ROW())` = ROW()−1.
  4. **SG_Flags** setzt „M“, wenn die Monotonie leer ist. Google Sheets wertet leeren Text als größer als jede Zahl.
  5. **ctl_gap** bezieht sich auf die feste Zeile 10 (`INDEX($10:$10; …)`).
  6. **Fehlerwerte** wie `#REF!` und `#VALUE!` in den ersten Tagen der Historie.

## Patch B – Formelblatt und alle Formeln (Apps Script, Funktion `ckDash_rohdaten_` aus Patch A)

Was er tut: Mit `blatt=timeline` liest die Route das Formelblatt statt der Wertekopie. Mit `formeln=alle` liefert sie zusätzlich die Formeln aller gelieferten Zeilen, damit v2 die Formelversion jeder Zelle prüfen kann. Ohne diese Parameter verhält sich die Route wie bisher.

Er bleibt nur lesend und betrifft nur die neue Funktion.

### B1

Streiche:
```js
    const sh = ckDash_getSheet_('KK_TIMELINE');
```

Ersetze durch:
```js
    const blatt = String(p.blatt || '') === 'timeline' ? 'timeline' : 'KK_TIMELINE';
    const sh = ckDash_getSheet_(blatt);
```

### B2

Streiche:
```js
      blatt: 'KK_TIMELINE',
```

Ersetze durch:
```js
      blatt: blatt,
```

### B3

Streiche:
```js
    if (String(p.formeln || '') === '1' && zeilen.length) {
```

Ersetze durch:
```js
    if (String(p.formeln || '') === 'alle' && zeilen.length) {
      // Formeln aller gelieferten Zeilen aus `timeline` (nur Zellen mit Formel) für den Versionsabgleich
      const tlA = ckDash_getSheet_('timeline');
      const von = zeilen[0][0];
      const bis = zeilen[zeilen.length - 1][0];
      const breite = tlA.getLastColumn();
      const kopf = tlA.getRange(1, 1, 1, breite).getValues()[0];
      const fA = tlA.getRange(von, 1, bis - von + 1, breite).getFormulas();
      const formelnAlle = {};
      for (let r = 0; r < fA.length; r++) {
        for (let c = 0; c < fA[r].length; c++) {
          if (!fA[r][c]) continue;
          const name = String(kopf[c] || '').trim();
          if (!formelnAlle[name]) formelnAlle[name] = {};
          formelnAlle[name][von + r] = fA[r][c];
        }
      }
      out.formelnAlle = formelnAlle;
    }
    if (String(p.formeln || '') === '1' && zeilen.length) {
```

### Einspielen

1. B1 bis B3 in `ckDash_rohdaten_` einfügen und speichern.
2. **Bereitstellen → Bereitstellungen verwalten → „v2 Schattenbetrieb“ → Bearbeiten → Version: Neue Version → Bereitstellen.** Die URL bleibt gleich. Die Live-Bereitstellung bitte nicht anfassen, sie bleibt auf v978.
3. Test: `<v2-URL>?mode=rohdaten&blatt=timeline&formeln=alle` liefert `"blatt": "timeline"` und `formelnAlle`.
4. Danach zeigt coachkira-v2 unter Technik den Abgleich „Rechenkern · Abgleich mit dem Sheet“.

---

## Stand nach Patch B (20.11.2026, 17:19)

- „v2 Schattenbetrieb“ steht auf **Version 980** (Patch A + B). Live bleibt auf v978 und kennt die Route nicht.
- coachkira-v2 → Technik → „Rechenkern · Abgleich mit dem Sheet“ rechnet live nach. Alle 11 Formeln stimmen zu 100 %, und es gibt eine Sheet-Fehlerzelle (2025-09-01, `#VALUE!`).
- Drift zwischen der Wertekopie `KK_TIMELINE` und `timeline`: ESS an 17 von 452 Tagen, ctl_star an 255 Tagen, Morgenwerte an über 200 Tagen. Hier fehlt noch die Entscheidung, welches Blatt gilt.
- Noch offen: der automatische tägliche Abgleich (Baustein 1.5) und die 14-Tage-Beobachtung.

## Entscheidung und Etappe 1b (20.11.2026)

- **Entscheidung von Marc: Das Formelblatt `timeline` ist die Wahrheit.** `KK_TIMELINE` ist nur noch eine Kopie.
- v2 liest seine Verlaufsdaten (`mode=chartData`) jetzt über `/api/rechenkern/chartData` aus `timeline`. Format und Zeitfenster sind gleich wie bei Live, auch die Zeitumstellungs-Eigenheit, bei der aus 120 Tagen 118 Zeilen werden. Abweichungen gibt es nur an den Drift-Tagen.
- v2 ist jetzt auch im Browser nur lesend. `src/lib/v2guard.ts` fängt alle POST-Aufrufe und alle `action=…`-Aufrufe ans Apps Script ab.
- Der tägliche Abgleich läuft ab 10 Uhr und zählt die Tage in Folge ohne Abweichung. Am 20.11.2026 ist er bei 1/14. Die Liste liegt in der v2-Datenbank, bei einer Neuveröffentlichung kann sie verloren gehen.
- Noch aus `KK_TIMELINE` lesen in v2: `status`, `planJsonV2`, `briefingData` und die übrigen Apps-Script-Routen. Sie kommen als Nächstes dran.

## Etappe 1c und 1d (20.11.2026)

**Warum die beiden Blätter auseinanderlaufen (Ursache geklärt):** `copyTimelineData` kopiert nur heute und die Zukunft von `timeline` nach `KK_TIMELINE`. Vergangene Tage bleiben in `KK_TIMELINE` eingefroren, so wie sie am jeweiligen Tag waren. `timeline` rechnet sie neu, wenn später etwas nachgetragen wird. Das ist so gebaut und kein Fehler. Nach der Entscheidung gilt `timeline`.

Hinweis: Werte in `timeline` können sich rückwirkend ändern, etwa wenn eine Konstante wie `alpha_CTL` angepasst wird. `KK_TIMELINE` hält dagegen den damaligen Stand fest.

**Status, Plan und Briefings** lesen nicht direkt aus `KK_TIMELINE`. Sie kommen aus den Berichtsblättern (`AI_REPORT_*`, `AI_*_ANALYSIS`), die der Supervisor schreibt. Der Supervisor liest dafür die Heute-Zeile, und die ist in beiden Blättern gleich. Umgestellt werden sie erst, wenn der Supervisor-Teil in den Kern wandert (Etappe 2).

**1c – Prognoseformeln:** Der Kern rechnet jetzt auch `coachE_ATL_forecast`, `coachE_CTL_forecast`, `coachE_ACWR_forecast` sowie `fbATL_obs`, `fbCTL_obs` und `fbACWR_obs`. Den Entwurf hat das lokale Modell coach-kira geschrieben, die Cloud hat ihn geprüft. Ergebnis: **17 von 17 Spalten exakt**, die Prognosen auch über 41 bzw. 130 Zukunftstage.

**1d – Plan-Kette und Was-wäre-wenn:** `simuliere()` schreibt ATL, CTL und ACWR ab morgen aus den eigenen Ergebnissen fort. Gegen die Sheet-Prognose stimmen alle 13 Plantage exakt. Neue Route `GET /api/rechenkern/simulation?ess={"JJJJ-MM-TT":Wert}`. In v2 steht unter Technik der „Plan-Rechner“. Er dient nur zur Anzeige und speichert nichts.

**Nächste Schritte:**
1. Das Plan-Cockpit in v2 nutzt den Plan-Rechner statt des Umwegs über das Sheet (`planSimulationJson`).
2. Etappe 2 vorbereiten: erledigt, siehe `docs/etappe-2-votum.md`.
3. Absicherung von Live (Etappe 3), sobald du das freigibst.

## Etappe 1 technisch abgeschlossen (20.11.2026, Abend)

**Eine Formelquelle:** `src/kern/formeln.ts` enthält jetzt alle Lastformeln: ATL- und CTL-Schritt, Ø28, Monotonie, Smart Gains und fKEI. Diese Stellen nutzen sie:
- der Server (`server/rechenkern.ts`, `server/prognose.ts`) für Abgleich, Plan-Kette und Plan-Rechner,
- das Frontend über `src/lib/ctlModel.ts`. Die Datei reicht die Formeln nur noch weiter, die Schnittstelle bleibt gleich. Darüber laufen Plan-Cockpit, Recovery, Chat-Kontext, Paket 3 und Prognose-Check,
- die Widgets über `server/widget.mjs`.

Vorher standen dieselben Koeffizienten doppelt im Code, einmal im Frontend und einmal im Server. Die Zahlen sind identisch geblieben. Nach dem Umbau stimmen weiterhin 17/17 Spalten exakt, die Plan-Kette trifft 13/13 Tage, und die Seite zeigt keine Fehler.

**Apps Script:** Für diesen Schritt war keine Änderung nötig. Live bleibt auf v978, „v2 Schattenbetrieb“ auf v980.

**Formale Abnahme:** Sie gilt erst nach 14 Tagen in Folge ohne Abweichung. Der Zähler läuft ab heute neu. Er beginnt nach jeder Veröffentlichung von v2 wieder bei null, weil die Server-Datenbank dabei neu angelegt wird.

## Ablauf nach Patch A

1. **Formeln ziehen:** Der Server in v2 holt `rohdaten&formeln=1` und legt die Formeln aller berechneten Spalten (`coachE_*`, `fKEI`, `load_star`, `ctl_star`, `ctl_gap`, `Monotony7`, `Strain7`, `SG_*`, `TE_Balance_Trend`) als Vorlage ab.
2. **Portieren (hybrid):** Das lokale Modell coach-kira schreibt je Formel eine TypeScript-Funktion. Die Cloud prüft jede Funktion gegen die Historie. Eine Formel gilt als fertig, wenn die Historie ohne Abweichung passt.
3. **Abgleich:** `GET /api/rechenkern/abgleich` rechnet alle Tage nach. Es zeigt je Spalte die Zahl der Treffer, die größte Abweichung und die ersten fünf Tage mit Abweichung. Die Seite „Abgleich“ unter Technik in v2 zeigt das als Tabelle.
4. **Täglich:** v2 gleicht morgens nach dem Formular ab, nur lesend, und merkt sich das Ergebnis.
5. **Abnahme:** Sind die Kriterien erfüllt, plane ich Etappe 1b. Dann liest das Cockpit die Kernwerte. Die Sheet-Formeln bleiben als Gegenprobe, und Apps Script rechnet nichts mehr doppelt.

## Bekannte Punkte aus der Sicherheitsprüfung (betrifft Live, Thema Etappe 3)

- `/api/auth/setup` lässt sich nach der Einrichtung erneut aufrufen. Damit ließe sich der PIN überschreiben.
- `/api/widget/koppeln`, `/geraete`, `/widerrufen` und `/api/llm` prüfen nur die Herkunft. Die Prüfung lässt fehlende Origin-Header und jede `*.pplx.app`-Seite durch.
- Die Mac-Endpunkte (`/api/lokal/*`, `/api/texte/mac|claim`, `/api/garmin/*`) haben kein gemeinsames Geheimnis.
- `npm audit`: eine kritische Lücke über express/proxy-addr, dazu mehrere hohe.

In v2 sind alle schreibenden API-Aufrufe gesperrt. Live bleibt wie gewünscht unverändert. Die Absicherung von Live ist Etappe 3 und kann auf Wunsch vorgezogen werden.
