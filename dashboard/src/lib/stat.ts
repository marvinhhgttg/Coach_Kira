/** Kleine, robuste Statistik-Helfer für die Analysen (Median, Theil-Sen, Effektstärke, Datum). */
import type { GarminAkt, GarminTage } from './garmin';

export const median = (xs: number[]): number | null => {
  const s = xs.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!s.length) return null;
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};
export const mittel = (xs: number[]): number | null => {
  const s = xs.filter((x) => Number.isFinite(x));
  return s.length ? s.reduce((a, b) => a + b, 0) / s.length : null;
};
export const sd = (xs: number[]): number | null => {
  const m = mittel(xs);
  if (m == null || xs.length < 2) return null;
  return Math.sqrt(xs.reduce((a, x) => a + (x - m) ** 2, 0) / (xs.length - 1));
};
export const quantil = (xs: number[], q: number): number | null => {
  const s = xs.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!s.length) return null;
  const i = (s.length - 1) * q, lo = Math.floor(i), hi = Math.ceil(i);
  return s[lo] + (s[hi] - s[lo]) * (i - lo);
};

/** Robuste Steigung (Theil-Sen) je Tag; Paare mit < 1 Tag Abstand zählen nicht. */
export function theilSen(pts: { t: number; y: number }[]): { b: number; a: number } | null {
  if (pts.length < 3) return null;
  const st: number[] = [];
  const n = Math.min(pts.length, 400);
  const p = pts.length > n ? pts.filter((_, i) => i % Math.ceil(pts.length / n) === 0) : pts;
  for (let i = 0; i < p.length; i++) for (let j = i + 1; j < p.length; j++) if (p[j].t - p[i].t >= 1) st.push((p[j].y - p[i].y) / (p[j].t - p[i].t));
  const b = median(st);
  if (b == null) return null;
  const a = median(p.map((q) => q.y - b * q.t))!;
  return { b, a };
}

/** Effektstärke (Cohen d) zweier Gruppen. */
export function cohenD(a: number[], b: number[]): number | null {
  if (a.length < 3 || b.length < 3) return null;
  const ma = mittel(a)!, mb = mittel(b)!, sa = sd(a)!, sb = sd(b)!;
  const s = Math.sqrt(((a.length - 1) * sa ** 2 + (b.length - 1) * sb ** 2) / (a.length + b.length - 2));
  return s > 0 ? (ma - mb) / s : null;
}

export const tageSeit = (d: string, basis = '2025-01-01') => (new Date(d + 'T12:00:00').getTime() - new Date(basis + 'T12:00:00').getTime()) / 86400000;
export const plusTage = (d: string, n: number) => { const x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + n); return x.toLocaleDateString('sv-SE'); };
export const wochentag = (d: string) => (new Date(d + 'T12:00:00').getDay() + 6) % 7; // 0 = Mo
export const WT = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
export const wochenStart = (d: string) => plusTage(d, -wochentag(d));
export const kurz = (d: string) => new Date(d + 'T12:00:00').toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: '2-digit' });
export const monatName = (ym: string) => new Date(ym + '-15T12:00:00').toLocaleDateString('de-DE', { month: 'short', year: '2-digit' });
export const de = (v: number | null | undefined, d = 0) => (v == null || !Number.isFinite(v) ? '–' : v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d }));
export const vz = (v: number | null | undefined, d = 0) => (v == null || !Number.isFinite(v) ? '–' : `${v > 0 ? '+' : ''}${de(v, d)}`);
export const hhmm = (min: number) => { const m = ((Math.round(min) % 1440) + 1440) % 1440; return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`; };
export const tempoTxt = (minProKm: number | null | undefined) => {
  if (minProKm == null || !Number.isFinite(minProKm)) return '–';
  const s = Math.round(minProKm * 60);
  return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
};
export const dauerTxt = (s: number | null | undefined) => {
  if (s == null || !Number.isFinite(s)) return '–';
  const m = Math.round(s / 60);
  if (m >= 6000) return `${Math.round(m / 60).toLocaleString('de-DE')} h`;
  return m >= 60 ? `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')} h` : `${m} min`;
};

export const LAUF = new Set(['running']);
export const RAD = new Set(['road_biking', 'cycling', 'gravel_cycling', 'mountain_biking']);
export const RAD_ALLE = new Set([...RAD, 'indoor_cycling', 'virtual_ride']);
export const BERG = new Set(['hiking', 'mountaineering', 'backcountry_skiing', 'trail_running']);

export type E = { datum: string; a: GarminAkt };
export function eintraege(tage: GarminTage | null): E[] {
  const out: E[] = [];
  if (!tage) return out;
  for (const [d, g] of tage) for (const a of g.akt || []) out.push({ datum: d, a });
  return out.sort((x, y) => (x.datum + (x.a.start || '')).localeCompare(y.datum + (y.a.start || '')));
}
export const vBewegt = (a: GarminAkt) => (a.km && a.bewegt_s ? a.km / (a.bewegt_s / 3600) : a.v_kmh || 0);
export const minuten = (a: GarminAkt) => (a.bewegt_s || a.dauer_s || 0) / 60;
/** Herzschläge pro km (Ø HF × min/km). */
export const schlaegeProKm = (a: GarminAkt) => { const v = vBewegt(a); return v > 0 && a.hf_avg ? (a.hf_avg * 60) / v : null; };
/** Flach, mit Trainingsreiz, ohne Bummeln – wie in der Effizienz-Kachel. */
export function flachGeeignet(a: GarminAkt, sport: 'lauf' | 'rad'): boolean {
  if (!a.km || !a.hf_avg) return false;
  const hpk = (a.hm || 0) / a.km, v = vBewegt(a), m = minuten(a);
  if (sport === 'lauf') return LAUF.has(a.typ || '') && hpk <= 5 && m >= 20 && a.km >= 3 && a.hf_avg >= 115 && v >= 5 && v <= 22;
  return RAD.has(a.typ || '') && hpk <= 4 && m >= 30 && a.km >= 10 && a.hf_avg >= 105 && v >= 15 && v <= 50;
}
export const startMin = (a: GarminAkt) => { const m = /^(\d{1,2}):(\d{2})/.exec(a.start || ''); return m ? Number(m[1]) * 60 + Number(m[2]) : null; };
/** "22:14" → Minuten relativ zu Mitternacht (Abend negativ). */
export const zuNachtMin = (s?: string) => { const m = /^(\d{1,2}):(\d{2})/.exec(s || ''); if (!m) return null; const x = Number(m[1]) * 60 + Number(m[2]); return x >= 12 * 60 ? x - 1440 : x; };
/** Harte Einheit: Qualitätsreiz laut Garmin Training Effect oder hohe Last. */
export const hart = (a: GarminAkt) => (a.te_ae || 0) >= 3.5 || (a.te_an || 0) >= 2 || (a.last || 0) >= 150;
