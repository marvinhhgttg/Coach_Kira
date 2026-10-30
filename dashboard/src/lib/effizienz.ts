/**
 * Effizienz-Verläufe aus den Garmin-Aktivitäten (garmin_tag): wie viel Herzschlag bzw. Leistung kostet
 * eine Einheit Tempo – getrennt nach flachen Strecken und Strecken mit vielen Höhenmetern.
 *
 * Laufen: Herzschläge pro km (Ø HF × min/km) und Laufleistung je km/h (Ø Watt ÷ Ø km/h)
 * Rad:    Herzschläge pro km (Ø HF × min/km) und Normalized Power je km/h (NP ÷ Ø km/h)
 * Bei allen vier Kennzahlen gilt: niedriger = effizienter (gleiches Tempo mit weniger Puls bzw. Leistung).
 * Rein deterministisch – das Sprachmodell formuliert nur das Fazit.
 */
import type { GarminAkt, GarminTage } from './garmin';

export type Sportart = 'lauf' | 'rad';
export type Gelaende = 'flach' | 'berg';
export type Untergrund = 'alle' | 'strasse' | 'gravel';
export type Kennzahl = 'hf' | 'leistung';

export const LAUF_TYPEN = new Set(['running']); // Trail bewusst nicht: Untergrund verzerrt HF und Leistung je Tempo
export const RAD_TYPEN = new Set(['road_biking', 'cycling', 'gravel_cycling', 'mountain_biking']);
const STRASSE = new Set(['road_biking', 'cycling']);

/** Voreinstellung der Geländegrenzen in Höhenmetern je km (dazwischen = gemischt, nicht gezählt). */
export const GRENZEN_STANDARD: Record<Sportart, { flach: number; berg: number }> = {
  lauf: { flach: 5, berg: 10 },
  rad: { flach: 4, berg: 7 },
};
/** Mindestumfang, damit Ø-Werte aussagekräftig sind (keine Ein- und Auslaufschnipsel). */
const MIN: Record<Sportart, { minuten: number; km: number }> = { lauf: { minuten: 20, km: 3 }, rad: { minuten: 30, km: 10 } };
/** Mindest-Ø-HF, damit die Einheit ein Trainingsreiz war. */
export const MIN_HF: Record<Sportart, number> = { lauf: 115, rad: 105 };

export interface EffPunkt {
  datum: string;
  ts: number;
  typ: string;
  name?: string;
  km: number;
  hm: number;
  hmProKm: number;
  minuten: number;
  hf: number;
  v: number; // km/h
  tempo: number; // min/km
  watt: number | null; // Lauf: Ø Leistung · Rad: NP
  schlaegeProKm: number;
  wattProKmh: number | null;
}

export function sportVon(typ?: string): Sportart | null {
  if (!typ) return null;
  if (LAUF_TYPEN.has(typ)) return 'lauf';
  if (RAD_TYPEN.has(typ)) return 'rad';
  return null;
}

function punkt(datum: string, a: GarminAkt, sport: Sportart): EffPunkt | null {
  const hf = a.hf_avg, km = a.km;
  const minuten = (a.bewegt_s || a.dauer_s || 0) / 60;
  if (!hf || !km || minuten <= 0) return null;
  // Tempo aus der Bewegungszeit – Pausen (Ampel, Foto, Hütte) verfälschen sonst das Verhältnis
  const v = a.bewegt_s ? km / (a.bewegt_s / 3600) : a.v_kmh;
  if (!v || v <= 0) return null;
  if (minuten < MIN[sport].minuten || km < MIN[sport].km) return null;
  // Ohne Trainingsreiz (Familien-/Genussausfahrt, Spaziergang-Lauf) sagt das Verhältnis nichts über die Form
  if (hf < MIN_HF[sport]) return null;
  if (sport === 'lauf' && (v < 5 || v > 22)) return null; // Gehpausen-/GPS-Ausreißer
  if (sport === 'rad' && (v < 8 || v > 50)) return null;
  const hm = a.hm || 0;
  if (sport === 'rad' && hm / km < 4 && v < 15) return null; // flach und unter 15 km/h = gebummelt
  const w = sport === 'lauf' ? a.watt : a.watt_np;
  const watt = typeof w === 'number' && w > 0 ? w : null;
  return {
    datum, ts: new Date(datum + 'T12:00:00').getTime(), typ: a.typ || '', name: a.name,
    km, hm, hmProKm: hm / km, minuten, hf, v, tempo: 60 / v, watt,
    schlaegeProKm: (hf * 60) / v,
    wattProKmh: watt != null ? watt / v : null,
  };
}

/** Alle geeigneten Einheiten einer Sportart (chronologisch). */
export function effPunkte(tage: GarminTage | null, sport: Sportart, untergrund: Untergrund = 'alle'): EffPunkt[] {
  if (!tage) return [];
  const out: EffPunkt[] = [];
  for (const [d, g] of tage) {
    for (const a of g.akt || []) {
      if (sportVon(a.typ) !== sport) continue;
      if (sport === 'rad' && untergrund !== 'alle') {
        const istStrasse = STRASSE.has(a.typ || '');
        if ((untergrund === 'strasse') !== istStrasse) continue;
      }
      const p = punkt(d, a, sport);
      if (p) out.push(p);
    }
  }
  return out.sort((a, b) => a.ts - b.ts);
}

export function nachGelaende(ps: EffPunkt[], g: Gelaende, grenzen: { flach: number; berg: number }): EffPunkt[] {
  return ps.filter((p) => (g === 'flach' ? p.hmProKm <= grenzen.flach : p.hmProKm >= grenzen.berg));
}

const median = (xs: number[]) => {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

export const wertVon = (p: EffPunkt, k: Kennzahl): number | null => (k === 'hf' ? p.schlaegeProKm : p.wattProKmh);

export interface EffBlock { wert: number | null; n: number; tempo: number | null; hf: number | null; v: number | null; hmProKm: number | null }

export interface EffTrend {
  n: number;
  start: string | null; ende: string | null;
  /** Werte der Ausgleichsgeraden am Anfang und Ende des Zeitraums */
  linieStart: number | null; linieEnde: number | null;
  /** Veränderung in % über den Zeitraum (negativ = effizienter) */
  prozent: number | null;
  /** Veränderung je 90 Tage in % */
  prozent90: number | null;
  /** Median der ersten und letzten 90 Tage, mit Tempo/HF zur Einordnung */
  frueh: EffBlock;
  spaet: EffBlock;
}

/** Robuste Ausgleichsgerade über die Zeit + Vergleich erste vs. letzte 90 Tage. */
export function trend(ps: EffPunkt[], k: Kennzahl): EffTrend {
  const xs = ps.filter((p) => wertVon(p, k) != null);
  const leer: EffBlock = { wert: null, n: 0, tempo: null, hf: null, v: null, hmProKm: null };
  if (xs.length < 2) return { n: xs.length, start: xs[0]?.datum ?? null, ende: xs[0]?.datum ?? null, linieStart: null, linieEnde: null, prozent: null, prozent90: null, frueh: leer, spaet: leer };
  const t0 = xs[0].ts, t1 = xs[xs.length - 1].ts;
  const T = xs.map((p) => (p.ts - t0) / 86400000), Y = xs.map((p) => wertVon(p, k)!);
  // Robuste Trendlinie (Theil-Sen: Median aller Paar-Steigungen) – einzelne Ausreißer kippen den Trend nicht
  const steig: number[] = [];
  for (let i = 0; i < T.length; i++) for (let j = i + 1; j < T.length; j++) if (T[j] - T[i] >= 1) steig.push((Y[j] - Y[i]) / (T[j] - T[i]));
  const b = median(steig) ?? 0;
  const a0 = median(Y.map((y, i) => y - b * T[i])) ?? Y[0];
  const tage = (t1 - t0) / 86400000;
  const ls = a0, le = a0 + b * tage;
  const block = (sel: EffPunkt[]): EffBlock => ({
    wert: median(sel.map((p) => wertVon(p, k)!)), n: sel.length,
    tempo: median(sel.map((p) => p.tempo)), hf: median(sel.map((p) => p.hf)), v: median(sel.map((p) => p.v)),
    hmProKm: median(sel.map((p) => p.hmProKm)),
  });
  const frueh = block(xs.filter((p) => p.ts <= t0 + 90 * 86400000));
  const spaet = block(xs.filter((p) => p.ts >= t1 - 90 * 86400000));
  return {
    n: xs.length, start: xs[0].datum, ende: xs[xs.length - 1].datum,
    linieStart: ls, linieEnde: le,
    prozent: tage >= 30 && ls > 0 ? ((le - ls) / ls) * 100 : null,
    prozent90: tage >= 30 && ls > 0 ? ((b * 90) / ls) * 100 : null,
    frueh, spaet,
  };
}

/** Gleitender Median über die letzten n Einheiten (für die Verlaufslinie). */
export function gleitend(ps: EffPunkt[], k: Kennzahl, n = 5): (number | null)[] {
  const out: (number | null)[] = [];
  const vals: number[] = [];
  for (const p of ps) {
    const w = wertVon(p, k);
    if (w != null) vals.push(w);
    out.push(w == null ? null : median(vals.slice(-n)));
  }
  return out;
}
