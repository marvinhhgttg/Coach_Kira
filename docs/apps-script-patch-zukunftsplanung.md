# Apps-Script-Patch: Zukunftsplanung

Stand: 22.10.2026 · für Ziel- und Tourenplaner, Saisonplaner, Bereitschafts-Prognose, Wochen-Autopilot, Lastfaktoren und Umplanen

## Was sich ändert

**Tabelle**

Keine neuen Blätter oder Spalten. Alles landet im vorhandenen Blatt `AI_KIRA_EXTRAS`, mit sechs neuen Typen:

| Typ | Datum = | Inhalt |
|---|---|---|
| `ziel` | Zieldatum | Tour/Event mit Anforderungen und Vorbereitungsplan |
| `saison` | Erstellungstag | Rahmen für 12–24 Wochen (Phase, Ziel-Wochenlast, Sportmix) |
| `autopilot` | Montag der Planwoche | Vorschlag für die neue Woche (vom Mac, sonntagabends) |
| `umplanung` | Tag der Störung | erkannte Störung (Infekt-Verdacht, Reise, ausgefallener Tag) |
| `lastfaktoren` | Erstellungstag | Vorschläge für `KK_LOAD_FACTORS` (nur Vorschlag, nichts wird überschrieben) |
| `prognose` | Erstellungstag | Güte des Bereitschaftsmodells (zur Kontrolle) |

**Web-App**

1. `saveExtra` kann Einträge auch löschen (`loeschen: true`), damit du Ziele im Cockpit entfernen kannst.
2. Neue Route `GET mode=loadFactors`: liest `KK_LOAD_FACTORS` und `KK_ELEV_FACTORS`, nur lesend.
3. Der Supervisor-Prompt bekommt einen Block mit deinen aktiven Zielen und dem Saisonrahmen der laufenden Woche. So richtet der KI-Report den Wochenplan daran aus.

`KK_LOAD_FACTORS` ändere ich nicht automatisch. Das Cockpit zeigt dir Vorschläge, übernehmen musst du sie selbst.

---

## Änderung 1 – erlaubte Typen

**Streiche**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout'];
```

**Ersetze durch**

```javascript
const CKX_TYPES = ['treffer_tag', 'treffer_woche', 'selbsttest', 'plausi', 'ernaehrung', 'alltag', 'workout',
  'ziel', 'saison', 'autopilot', 'umplanung', 'lastfaktoren', 'prognose'];
```

---

## Änderung 2 – `ckX_save_`: Löschen

**Streiche**

```javascript
      const sh = ckX_sheet_();
      const row = [datum, typ, text, data, quelle, ckDash_now_()];
      const n = sh.getLastRow() - 1;
```

**Ersetze durch**

```javascript
      const sh = ckX_sheet_();
      const row = [datum, typ, text, data, quelle, ckDash_now_()];
      const n = sh.getLastRow() - 1;
      // Eintrag löschen (z. B. ein Ziel entfernen): {typ, datum, loeschen:true}
      if (b.loeschen === true) {
        if (n > 0) {
          const keys0 = sh.getRange(2, 1, n, 2).getDisplayValues();
          for (let i = n - 1; i >= 0; i--) {
            if (keys0[i][0] === datum && keys0[i][1] === typ) sh.deleteRow(i + 2);
          }
        }
        return ckDash_json_({ ok: true, timestamp: ckDash_now_(), deleted: true });
      }
```

---

## Änderung 3 – `doGet`: neue Route

**Streiche**

```javascript
  if (p.mode === 'physBaseline') return ckX_phys_();
```

**Ersetze durch**

```javascript
  if (p.mode === 'physBaseline') return ckX_phys_();
  if (p.mode === 'loadFactors') return ckX_loadFactors_();
```

---

## Änderung 4 – `createMasterGeminiPrompt`: Ziele und Saison

**Streiche**

```javascript
${ckRegel_monotonieBlock_()}
```

**Ersetze durch**

```javascript
${ckRegel_monotonieBlock_()}
${ckX_planungsBlock_()}
```

---

## Änderung 5 – neue Funktionen (ans Ende der Datei)

```javascript
/** GET mode=loadFactors · KK_LOAD_FACTORS (ESS pro Stunde je Sport/Zone) und KK_ELEV_FACTORS (nur lesen) */
function ckX_loadFactors_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const read = function (name) {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return [];
      const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      const h = v[0].map(function (x) { return String(x).trim(); });
      return v.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
        const o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
      });
    };
    return ckDash_json_({ ok: true, timestamp: ckDash_now_(), load: read('KK_LOAD_FACTORS'), elev: read('KK_ELEV_FACTORS') });
  } catch (err) {
    return ckDash_json_({ ok: false, timestamp: ckDash_now_(), error: String(err && err.message ? err.message : err) });
  }
}

/** Textblock für den Supervisor-Prompt: aktive Ziele/Touren und der Saisonrahmen der laufenden Woche */
function ckX_planungsBlock_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CKX_SHEET);
    if (!sh || sh.getLastRow() < 2) return '';
    const tz = Session.getScriptTimeZone() || 'Europe/Berlin';
    const heute = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
    const ziele = [];
    let saison = null;
    rows.forEach(function (r) {
      if (r[1] === 'ziel' && r[0] >= heute) ziele.push(r);
      if (r[1] === 'saison' && (!saison || r[0] > saison[0])) saison = r;
    });
    let out = '';
    if (ziele.length) {
      ziele.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
      out += '**ZIELE / TOUREN (darauf hinplanen):**\n' + ziele.slice(0, 5).map(function (r) { return '- ' + r[0] + ': ' + String(r[2]).slice(0, 300); }).join('\n') + '\n';
    }
    if (saison && saison[3]) {
      try {
        const d = JSON.parse(saison[3]);
        const wochen = d.wochen || [];
        const w = wochen.filter(function (x) { return x.start <= heute; }).pop() || wochen[0];
        const n = wochen[wochen.indexOf(w) + 1];
        if (w) out += '**SAISONRAHMEN diese Woche (ab ' + w.start + '):** Phase ' + w.phase + ', Ziel-Wochenlast ca. ' + w.ziel_ess + ' ESS, Fokus: ' + (w.fokus || '-') + (n ? ' · nächste Woche: ' + n.phase + ', ca. ' + n.ziel_ess + ' ESS' : '') + '\n';
      } catch (e) {}
    }
    return out ? '\n---\n' + out + 'Richte Plan und Empfehlung an Zielen und Saisonrahmen aus, ohne die Monotonie- und ACWR-Regeln zu verletzen.\n---\n' : '';
  } catch (e) {
    return '';
  }
}
```

---

Danach speichern, neue Version bereitstellen, kurz Bescheid geben.
