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
