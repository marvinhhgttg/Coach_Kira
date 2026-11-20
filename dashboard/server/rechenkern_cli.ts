// Abgleich Rechenkern ↔ Sheet (Etappe 1, Schattenbetrieb, nur lesend)
import { rechneZeile, KERN_SPALTEN, type Zeile, type Zelle } from './rechenkern';

type Roh = { ok: boolean; error?: string; spalten: string[]; zeilen: Zelle[][]; konstanten: Record<string, Zelle>; formelnAlle?: Record<string, Record<string, string>> };

// Zellbezüge relativ zur eigenen Zeile schreiben (AF340 in Zeile 342 → AF[-2]); $-Bezüge bleiben absolut
export function normFormel(f: string, zeile: number): string {
  return f
    .replace(/(?<![A-Z$_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g, (m, d1, col, d2, n) => (d2 === '$' ? m : `${d1}${col}[${Number(n) - zeile}]`))
    .replace(/\s+/g, '');
}

const leer = (v: unknown) => v === null || v === undefined || v === '';
const istFehler = (v: unknown) => typeof v === 'string' && /^#(REF|VALUE|N\/A|DIV\/0|NUM|NAME|ERROR)/.test(v);

export function abgleich(tl: Roh, kk: Roh | null, heute: string) {
  const zu = (r: Roh) => r.zeilen.map((z) => Object.fromEntries(r.spalten.map((s, j) => [s, z[j]])) as Zeile);
  const rows = zu(tl);
  const iHeute = rows.findIndex((r) => r.date === heute);
  const ergebnis: Record<string, unknown>[] = [];
  for (const k of KERN_SPALTEN) {
    const formeln = tl.formelnAlle?.[k] || {};
    const vorlage = iHeute >= 0 && formeln[String(rows[iHeute].row)] ? normFormel(formeln[String(rows[iHeute].row)], Number(rows[iHeute].row)) : null;
    let geprueft = 0, treffer = 0, sheetFehler = 0, alteVersion = 0, festwert = 0, maxDiff = 0;
    const beispiele: string[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r.date) > heute) continue;
      const f = formeln[String(r.row)];
      if (!f) { festwert++; continue; }
      if (!vorlage || normFormel(f, Number(r.row)) !== vorlage) { alteVersion++; continue; }
      const soll = r[k];
      if (istFehler(soll)) { sheetFehler++; continue; }
      const ist = rechneZeile(rows, i, tl.konstanten)[k];
      geprueft++;
      let gleich: boolean;
      if (leer(soll) || leer(ist)) gleich = leer(soll) && leer(ist);
      else if (typeof soll === 'number' && typeof ist === 'number') {
        const d = Math.abs(soll - ist);
        maxDiff = Math.max(maxDiff, d);
        gleich = d <= Math.max(0.05, Math.abs(soll) * 0.005);
      } else gleich = String(soll).trim() === String(ist).trim();
      if (gleich) treffer++;
      else if (beispiele.length < 5) beispiele.push(`${r.date}: Sheet ${JSON.stringify(soll)} · Kern ${JSON.stringify(ist)}`);
    }
    ergebnis.push({ spalte: k, geprueft, treffer, sheetFehler, alteVersion, festwert, maxDiff, beispiele });
  }

  // Drift: Wertekopie KK_TIMELINE gegenüber dem Formelblatt timeline (gleiche Zeilennummer)
  let drift: Record<string, unknown>[] = [];
  if (kk?.ok) {
    const kkRows = new Map(zu(kk).map((r) => [String(r.row), r]));
    const spalten = ['coachE_ESS_day', 'coachE_ATL_forecast', 'coachE_CTL_forecast', 'coachE_ACWR_forecast', ...KERN_SPALTEN];
    drift = spalten.map((s) => {
      let n = 0, ab = 0; const bsp: string[] = [];
      for (const r of rows) {
        if (String(r.date) > heute) continue;
        const k2 = kkRows.get(String(r.row)); if (!k2) continue;
        const a = r[s], b = k2[s]; n++;
        const gleich = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= Math.max(0.05, Math.abs(a) * 0.005) : String(a ?? '').trim() === String(b ?? '').trim() || (istFehler(a) && istFehler(b));
        if (!gleich) { ab++; if (bsp.length < 3) bsp.push(`${r.date}: timeline ${JSON.stringify(a)} · KK ${JSON.stringify(b)}`); }
      }
      return { spalte: s, verglichen: n, abweichend: ab, beispiele: bsp };
    });
  }
  return { ok: true, heute, zeilen: rows.length, spalten: ergebnis, drift };
}
