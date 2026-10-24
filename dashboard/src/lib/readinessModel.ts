/**
 * Bereitschafts-Prognose: lernt aus Marcs eigenen Daten, wie Trainingsbereitschaft und HRV
 * am Morgen nach einem Tag auf Last, ACWR, Monotonie und den Ausgangswert reagieren.
 * Ridge-Regression, lokal im Browser – kein Sprachmodell nötig.
 */
import { fetchChartData } from './api';

export interface HistRow {
  date: string;
  load: number | null;
  readiness: number | null;
  hrv: number | null;
  acwr: number | null;
  mono: number | null;
}

export interface Model {
  w: number[];
  mean: number[];
  sd: number[];
  n: number;
  mae: number; // mittlerer Fehler auf den letzten 20 % (zeitlich getrennt)
  r2: number;
  baselineMae: number; // Fehler der naiven Prognose „wie heute“
}

export interface ReadinessForecast {
  date: string;
  readiness: number;
  hrv: number | null;
  band: [number, number];
}

const num = (v: unknown): number | null => {
  if (v === '' || v == null) return null;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};

export async function fetchHistory(): Promise<HistRow[]> {
  const r: any = await fetchChartData('180d' as any, [
    'load_fb_day',
    'Garmin_Training_Readiness',
    'hrv_status',
    'fbACWR_obs',
    'Monotony7',
  ]);
  return ((r && r.data) || []).map((x: any) => ({
    date: String(x.date),
    load: num(x.load_fb_day),
    readiness: num(x.Garmin_Training_Readiness),
    hrv: num(x.hrv_status),
    acwr: num(x.fbACWR_obs),
    mono: num(x.Monotony7),
  }));
}

// Merkmale für „Morgen danach“ aus Tag D: Last D, Last D-1, Ø Last 7 Tage, ACWR D, Monotonie D, Ausgangswert D, Ruhetag-Flag
function feats(load: number, loadPrev: number, mean7: number, acwr: number, mono: number, start: number): number[] {
  return [load, loadPrev, mean7, acwr, Math.min(mono, 4), start, load <= 20 ? 1 : 0];
}

function solve(A: number[][], b: number[]): number[] {
  // Gauß mit Pivotisierung
  const n = b.length;
  const M = A.map((r, i) => [...r, b[i]]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
    [M[c], M[p]] = [M[p], M[c]];
    const d = M[c][c] || 1e-9;
    for (let k = c; k <= n; k++) M[c][k] /= d;
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = M[r][c];
      for (let k = c; k <= n; k++) M[r][k] -= f * M[c][k];
    }
  }
  return M.map((r) => r[n]);
}

function fit(X: number[][], y: number[], lambda = 3): Omit<Model, 'mae' | 'r2' | 'baselineMae' | 'n'> {
  const p = X[0].length;
  const mean = Array.from({ length: p }, (_, j) => X.reduce((a, r) => a + r[j], 0) / X.length);
  const sd = Array.from({ length: p }, (_, j) => Math.sqrt(X.reduce((a, r) => a + (r[j] - mean[j]) ** 2, 0) / X.length) || 1);
  const Z = X.map((r) => [1, ...r.map((v, j) => (v - mean[j]) / sd[j])]);
  const k = p + 1;
  const A = Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => Z.reduce((a, r) => a + r[i] * r[j], 0) + (i === j && i > 0 ? lambda : 0)));
  const bb = Array.from({ length: k }, (_, i) => Z.reduce((a, r, t) => a + r[i] * y[t], 0));
  return { w: solve(A, bb), mean, sd };
}

function predict(m: { w: number[]; mean: number[]; sd: number[] }, x: number[]) {
  return m.w[0] + x.reduce((a, v, j) => a + m.w[j + 1] * ((v - m.mean[j]) / m.sd[j]), 0);
}

function dataset(h: HistRow[], target: 'readiness' | 'hrv') {
  const X: number[][] = [];
  const y: number[] = [];
  const naive: number[] = [];
  for (let i = 7; i < h.length - 1; i++) {
    const d = h[i];
    const n = h[i + 1];
    const t = n[target];
    const s = d[target];
    if (t == null || s == null || d.load == null) continue;
    const prev = h[i - 1].load ?? 0;
    const w7 = h.slice(i - 6, i + 1).map((r) => r.load ?? 0);
    X.push(feats(d.load, prev, w7.reduce((a, b) => a + b, 0) / 7, d.acwr ?? 1, d.mono ?? 1.5, s));
    y.push(t);
    naive.push(s);
  }
  return { X, y, naive };
}

export function trainModel(h: HistRow[], target: 'readiness' | 'hrv'): Model | null {
  const { X, y, naive } = dataset(h, target);
  if (X.length < 30) return null;
  const cut = Math.floor(X.length * 0.8);
  const mTest = fit(X.slice(0, cut), y.slice(0, cut));
  const err = y.slice(cut).map((v, i) => v - predict(mTest, X[cut + i]));
  const mae = err.reduce((a, e) => a + Math.abs(e), 0) / err.length;
  const ym = y.slice(cut).reduce((a, b) => a + b, 0) / err.length;
  const ss = y.slice(cut).reduce((a, v) => a + (v - ym) ** 2, 0) || 1;
  const r2 = 1 - err.reduce((a, e) => a + e * e, 0) / ss;
  const baselineMae = y.slice(cut).reduce((a, v, i) => a + Math.abs(v - naive[cut + i]), 0) / err.length;
  const full = fit(X, y);
  return { ...full, n: X.length, mae, r2, baselineMae };
}

/** Prognose für die Plantage: rekursiv, Ausgangswert = Prognose des Vortags */
export function forecastPlan(
  h: HistRow[],
  mR: Model | null,
  mH: Model | null,
  plan: { date: string; load: number; acwr: number; mono: number }[],
): ReadinessForecast[] {
  if (!mR || !plan.length) return [];
  const lastR = [...h].reverse().find((r) => r.readiness != null)?.readiness ?? 60;
  const lastH = [...h].reverse().find((r) => r.hrv != null)?.hrv ?? null;
  const loads = h.map((r) => r.load ?? 0);
  // Plantag 0 = heute: Ausgangswert heute ist gemessen
  let r = lastR;
  let hv = lastH;
  const out: ReadinessForecast[] = [];
  for (let i = 0; i < plan.length; i++) {
    const d = plan[i];
    loads.push(d.load);
    const prev = loads[loads.length - 2] ?? 0;
    const w7 = loads.slice(-7);
    const m7 = w7.reduce((a, b) => a + b, 0) / w7.length;
    const xr = feats(d.load, prev, m7, d.acwr, d.mono, r);
    const rn = Math.max(1, Math.min(100, predict(mR, xr)));
    let hn: number | null = null;
    if (mH && hv != null) hn = Math.max(10, predict(mH, feats(d.load, prev, m7, d.acwr, d.mono, hv)));
    const next = i + 1 < plan.length ? plan[i + 1].date : addDay(d.date);
    out.push({ date: next, readiness: Math.round(rn), hrv: hn != null ? Math.round(hn) : null, band: [Math.round(rn - mR.mae), Math.round(rn + mR.mae)] });
    r = rn;
    hv = hn;
  }
  return out;
}

function addDay(d: string) {
  const t = new Date(d + 'T12:00:00');
  t.setDate(t.getDate() + 1);
  return t.toLocaleDateString('sv-SE');
}

/** Wichtigste Einflüsse in Worten (Vorzeichen der standardisierten Gewichte) */
export function explainModel(m: Model): string[] {
  const names = ['Last des Tages', 'Last am Vortag', 'Ø Last 7 Tage', 'ACWR', 'Monotonie', 'Ausgangswert', 'Ruhetag'];
  return m.w
    .slice(1)
    .map((w, i) => ({ n: names[i], w }))
    .sort((a, b) => Math.abs(b.w) - Math.abs(a.w))
    .slice(0, 4)
    .map((x) => `${x.n}: ${x.w > 0 ? '+' : '−'}${Math.abs(x.w).toFixed(1).replace('.', ',')} Punkte je Standardabweichung`);
}
