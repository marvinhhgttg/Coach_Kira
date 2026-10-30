/**
 * Prognose-Check: Sind die Prognosen eingetroffen?
 * - echt: gespeicherte Mac-Prognosen (Extra „prognose“, je Tag die letzte) gegen die Ist-Werte.
 * - modell: Rückrechnung mit den echten Lasten über die ganze Historie (reiner Modellfehler).
 * Vergleich jeweils mit „bleibt wie heute“ (naiv). Fest berechnet, ohne KI.
 */
import { ctlStep, esStep } from './ctlModel';

export interface IstTag { date: string; load: number | null; ctl: number | null; es: number | null; ready: number | null }
export interface PrognoseItem { datum: string; data: { vorausblick?: any; erstellt?: string } }
export interface Fehler { n: number; mae: number; bias: number; maeNaiv: number }
type H3 = '7' | '14' | '28';
export interface CheckErgebnis {
  seit: string | null;
  echt: {
    fitness: Record<H3, Fehler & { lastPlan: number; lastIst: number }>;
    es: Record<H3, Fehler>;
    bereitschaft: Record<'1' | '2' | '3', Fehler>;
    qualitaetstag: { n: number; eingetroffen: number; faelle: { datum: string; fuer: string; prog: number | null; ist: number | null; ok: boolean }[] };
  };
  modell: { fitness: Record<H3, Fehler>; es: Record<H3, Fehler> };
}

const r1 = (v: number) => Math.round(v * 10) / 10;
const addTage = (iso: string, n: number) => {
  const d = new Date(iso + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};
const tageZwischen = (a: string, b: string) => Math.round((Date.parse(b + 'T12:00:00Z') - Date.parse(a + 'T12:00:00Z')) / 86400000);
const zahl = (v: unknown): number | null => (v == null || v === '' || !Number.isFinite(Number(v)) ? null : Number(v));

class Samm {
  e: number[] = [];
  b: number[] = [];
  nv: number[] = [];
  add(prog: number, ist: number, naiv: number | null) {
    this.e.push(Math.abs(prog - ist));
    this.b.push(prog - ist);
    if (naiv != null) this.nv.push(Math.abs(naiv - ist));
  }
  get(): Fehler {
    const m = (x: number[]) => (x.length ? x.reduce((a, v) => a + v, 0) / x.length : 0);
    return { n: this.e.length, mae: r1(m(this.e)), bias: r1(m(this.b)), maeNaiv: r1(m(this.nv)) };
  }
}

function rueckrechnung(ist: IstTag[], feld: 'ctl' | 'es', win: number, h: number): Fehler {
  const s = new Samm();
  const L = ist.map((t) => t.load ?? 0);
  for (let i = win; i + h < ist.length; i++) {
    const start = ist[i][feld];
    const ziel = ist[i + h][feld];
    if (start == null || ziel == null) continue;
    let x = start;
    for (let k = i + 1; k <= i + h; k++) {
      const w = L.slice(Math.max(0, k - win + 1), k + 1);
      const mean = w.reduce((a, v) => a + v, 0) / w.length;
      x = feld === 'ctl' ? ctlStep(x, L[k], mean) : esStep(x, L[k], mean);
    }
    s.add(x, ziel, start);
  }
  return s.get();
}

export function prognoseCheck(items: PrognoseItem[], ist0: IstTag[], heute: string): CheckErgebnis {
  const ist = ist0.filter((t) => t.date < heute).sort((a, b) => a.date.localeCompare(b.date));
  // lückenlose Tagesreihe für die Rückrechnung (fehlende Tage = Last 0, Werte leer)
  const voll: IstTag[] = [];
  if (ist.length) {
    const byD = new Map(ist.map((t) => [t.date, t]));
    for (let d = ist[0].date; d <= ist[ist.length - 1].date; d = addTage(d, 1))
      voll.push(byD.get(d) || { date: d, load: null, ctl: null, es: null, ready: null });
  }
  const byDate = new Map(voll.map((t) => [t.date, t]));
  const H: H3[] = ['7', '14', '28'];

  const fit = Object.fromEntries(H.map((h) => [h, new Samm()])) as Record<H3, Samm>;
  const lastP = Object.fromEntries(H.map((h) => [h, [] as number[]])) as Record<H3, number[]>;
  const lastI = Object.fromEntries(H.map((h) => [h, [] as number[]])) as Record<H3, number[]>;
  const es = Object.fromEntries(H.map((h) => [h, new Samm()])) as Record<H3, Samm>;
  const ber = { '1': new Samm(), '2': new Samm(), '3': new Samm() };
  const qf: CheckErgebnis['echt']['qualitaetstag']['faelle'] = [];
  const its = [...items].filter((x) => x?.data?.vorausblick && /^\d{4}-\d{2}-\d{2}$/.test(x.datum)).sort((a, b) => a.datum.localeCompare(b.datum));

  for (const it of its) {
    const v = it.data.vorausblick;
    const tage: any[] = Array.isArray(v.tage) ? v.tage : [];
    const planLast = new Map<string, number>(tage.filter((t) => t?.datum).map((t) => [String(t.datum), Number(t.last) || 0]));
    for (const h of H) {
      const ziel = addTage(it.datum, Number(h));
      const t = byDate.get(ziel);
      const pc = zahl(v['ctl_' + h]);
      if (pc != null && t?.ctl != null) {
        fit[h].add(pc, t.ctl, zahl(v.ctl_heute));
        let ok = true, sp = 0, si = 0;
        for (let k = 1; k <= Number(h); k++) {
          const d = addTage(it.datum, k);
          if (!planLast.has(d)) { ok = false; break; }
          sp += planLast.get(d)!;
          si += byDate.get(d)?.load ?? 0;
        }
        if (ok) { lastP[h].push(sp); lastI[h].push(si); }
      }
      const pe = zahl(v['es_' + h]);
      if (pe != null && t?.es != null) es[h].add(pe, t.es, zahl(v.es_heute));
    }
    const naivR = byDate.get(it.datum)?.ready ?? null;
    for (const t of tage) {
      const k = String(tageZwischen(it.datum, String(t?.datum || '')));
      const pb = zahl(t?.bereitschaft);
      const ir = byDate.get(String(t?.datum))?.ready;
      if ((k === '1' || k === '2' || k === '3') && pb != null && ir != null) ber[k].add(pb, ir, naivR);
    }
    const q = v.qualitaetstag;
    if (q && q.datum) {
      const ir = byDate.get(String(q.datum))?.ready ?? null;
      if (ir != null) {
        const pb = zahl(q.bereitschaft);
        qf.push({ datum: it.datum, fuer: String(q.datum), prog: pb, ist: ir, ok: ir >= 55 && (pb == null || ir >= pb - 10) });
      }
    }
  }
  const mean = (x: number[]) => (x.length ? r1(x.reduce((a, v) => a + v, 0) / x.length) : 0);
  // je Zieltag nur der letzte Vorschlag zählt
  const qMap = new Map<string, (typeof qf)[number]>();
  for (const f of qf) qMap.set(f.fuer, f);
  const qs = [...qMap.values()].sort((a, b) => a.fuer.localeCompare(b.fuer));

  return {
    seit: its.length ? its[0].datum : null,
    echt: {
      fitness: Object.fromEntries(H.map((h) => [h, { ...fit[h].get(), lastPlan: mean(lastP[h]), lastIst: mean(lastI[h]) }])) as any,
      es: Object.fromEntries(H.map((h) => [h, es[h].get()])) as any,
      bereitschaft: { '1': ber['1'].get(), '2': ber['2'].get(), '3': ber['3'].get() },
      qualitaetstag: { n: qs.length, eingetroffen: qs.filter((x) => x.ok).length, faelle: qs },
    },
    modell: {
      fitness: Object.fromEntries(H.map((h) => [h, rueckrechnung(voll, 'ctl', 28, Number(h))])) as any,
      es: Object.fromEntries(H.map((h) => [h, rueckrechnung(voll, 'es', 56, Number(h))])) as any,
    },
  };
}
