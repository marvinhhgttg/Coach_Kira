/**
 * 23.11.2026: Garmin-Farbmuster für alle Ringe, Bögen und Balken – jeweils mit den individuellen Grenzwerten.
 * Farben wie in Garmin Connect: Rot · Orange · Grün · Blau · Lila. Nur Darstellung, Kiras Votum-Regeln bleiben unverändert.
 */
import { TR_BAENDER } from './trFarbe';

export const G = { rot: '#D9433A', orange: '#E5822F', gruen: '#5DB75B', blau: '#3A78CF', lila: '#6E45E9', grau: '#9aa5b1' } as const;

export type Band = { von: number; bis: number; farbe: string; name: string };
export type Skala = { min: number; max: number; baender: Band[] };

export function bandVon(s: Skala | null | undefined, v: number | null | undefined): Band | null {
  if (!s || v == null || !Number.isFinite(v)) return null;
  if (v < s.baender[0].von) return s.baender[0];
  return s.baender.find((b) => v >= b.von && v < b.bis) ?? s.baender[s.baender.length - 1];
}
export const farbeVon = (s: Skala | null | undefined, v: number | null | undefined) => bandVon(s, v)?.farbe ?? G.grau;
export const nameVon = (s: Skala | null | undefined, v: number | null | undefined) => bandVon(s, v)?.name ?? '';

const r1 = (x: number) => Math.round(x * 10) / 10;

/** Trainingsbereitschaft – Garmins feste Grenzen */
export const TR_SKALA: Skala = { min: 0, max: 100, baender: TR_BAENDER.map((b) => ({ ...b })) };

/** Body Battery – Garmins Stufen */
export const BB_SKALA: Skala = {
  min: 0, max: 100, baender: [
    { von: 0, bis: 26, farbe: G.rot, name: 'Sehr niedrig' },
    { von: 26, bis: 51, farbe: G.orange, name: 'Niedrig' },
    { von: 51, bis: 76, farbe: G.gruen, name: 'Mittel' },
    { von: 76, bis: 101, farbe: G.blau, name: 'Hoch' },
  ],
};

/** Schlafscore – Garmins Stufen */
export const SCHLAF_SCORE_SKALA: Skala = {
  min: 0, max: 100, baender: [
    { von: 0, bis: 60, farbe: G.rot, name: 'Schlecht' },
    { von: 60, bis: 80, farbe: G.orange, name: 'Ordentlich' },
    { von: 80, bis: 90, farbe: G.gruen, name: 'Gut' },
    { von: 90, bis: 101, farbe: G.blau, name: 'Ausgezeichnet' },
  ],
};

/** Kiras Scores (Supervisor, Training, Recovery, KEI, Ernährung) – Kiras Schwellen 30/50/70 */
export const SCORE_SKALA: Skala = {
  min: 0, max: 100, baender: [
    { von: 0, bis: 30, farbe: G.rot, name: 'Schlecht' },
    { von: 30, bis: 50, farbe: G.orange, name: 'Niedrig' },
    { von: 50, bis: 70, farbe: G.gruen, name: 'Mäßig' },
    { von: 70, bis: 101, farbe: G.blau, name: 'Gut' },
  ],
};

/** Schlafdauer gegen das persönliche Schlafziel (Kira: 80 / 85 / 95 %) */
export function schlafSkala(zielH: number): Skala {
  return {
    min: zielH * 0.6, max: zielH * 1.2, baender: [
      { von: 0, bis: zielH * 0.8, farbe: G.rot, name: 'Zu kurz' },
      { von: zielH * 0.8, bis: zielH * 0.85, farbe: G.orange, name: 'Knapp' },
      { von: zielH * 0.85, bis: zielH * 0.95, farbe: G.gruen, name: 'Ordentlich' },
      { von: zielH * 0.95, bis: 99, farbe: G.blau, name: 'Ziel erreicht' },
    ],
  };
}

/** HRV gegen das persönliche Normalband (Garmin: Niedrig · Unausgeglichen · Ausgeglichen) */
export function hrvSkala(lo: number | null | undefined, hi: number | null | undefined): Skala | null {
  if (lo == null || hi == null || !(hi > lo)) return null;
  const rand = Math.max(6, (hi - lo) * 0.7);
  return {
    min: r1(lo - rand), max: r1(hi + rand), baender: [
      { von: -999, bis: lo - 2, farbe: G.rot, name: 'Niedrig' },
      { von: lo - 2, bis: lo, farbe: G.orange, name: 'Unausgeglichen' },
      { von: lo, bis: hi, farbe: G.gruen, name: 'Ausgeglichen' },
      { von: hi, bis: 9999, farbe: G.orange, name: 'Unausgeglichen' },
    ],
  };
}

/** Ruhepuls-Abweichung zur persönlichen 28-Tage-Basis (Kira: −5 / +1 / +3 / +5 bpm) */
export const RHR_DELTA_SKALA: Skala = {
  min: -8, max: 8, baender: [
    { von: -99, bis: -5, farbe: G.lila, name: 'Auffällig niedrig' },
    { von: -5, bis: 1, farbe: G.blau, name: 'Normal' },
    { von: 1, bis: 3, farbe: G.gruen, name: 'Leicht erhöht' },
    { von: 3, bis: 5, farbe: G.orange, name: 'Erhöht' },
    { von: 5, bis: 99, farbe: G.rot, name: 'Deutlich erhöht' },
  ],
};

/** Befinden 1–5 (Kira: 2 / 3 / 4) */
export const BEFINDEN_SKALA: Skala = {
  min: 1, max: 5, baender: [
    { von: 0, bis: 2, farbe: G.rot, name: 'Schlecht' },
    { von: 2, bis: 3, farbe: G.orange, name: 'Mäßig' },
    { von: 3, bis: 4, farbe: G.gruen, name: 'Ok' },
    { von: 4, bis: 9, farbe: G.blau, name: 'Gut' },
  ],
};

/** ACWR mit Kiras Korridor (Garmin-Logik der akuten Last: Niedrig · Optimal · Hoch · Sehr hoch) */
export function acwrSkala(warn: number, hoch: number): Skala {
  return {
    min: 0.4, max: 2.0, baender: [
      { von: -9, bis: 0.8, farbe: G.blau, name: 'Niedrig' },
      { von: 0.8, bis: warn, farbe: G.gruen, name: 'Optimal' },
      { von: warn, bis: hoch, farbe: G.orange, name: 'Hoch' },
      { von: hoch, bis: 99, farbe: G.rot, name: 'Sehr hoch' },
    ],
  };
}

/** Tageslast gegen Kiras persönlichen Soll-Rahmen lo–hi (über +15 % = sehr hoch) */
export function lastSkala(lo: number | null | undefined, hi: number | null | undefined): Skala | null {
  if (hi == null || !(hi > 0)) return null;
  const u = lo != null && lo > 0 ? lo : 0;
  const b: Band[] = [];
  if (u > 0) b.push({ von: -1, bis: u, farbe: G.blau, name: 'Unter Soll' });
  b.push({ von: u > 0 ? u : -1, bis: hi + 0.5, farbe: G.gruen, name: 'Im Rahmen' });
  b.push({ von: hi + 0.5, bis: hi * 1.15, farbe: G.orange, name: 'Über Soll' });
  b.push({ von: hi * 1.15, bis: 1e9, farbe: G.rot, name: 'Deutlich drüber' });
  return { min: 0, max: Math.round(hi * 1.3), baender: b };
}

/** Wochenlast gegen Kiras Wochenbudget (85 % / 100 % / 115 %) */
export function wocheSkala(budget: number | null | undefined): Skala | null {
  if (budget == null || !(budget > 0)) return null;
  return {
    min: 0, max: Math.round(budget * 1.3), baender: [
      { von: -1, bis: budget * 0.85, farbe: G.blau, name: 'Unter Budget' },
      { von: budget * 0.85, bis: budget + 0.5, farbe: G.gruen, name: 'Im Budget' },
      { von: budget + 0.5, bis: budget * 1.15, farbe: G.orange, name: 'Über Budget' },
      { von: budget * 1.15, bis: 1e9, farbe: G.rot, name: 'Deutlich drüber' },
    ],
  };
}

/** Tourform / Zielbereitschaft in % (Kira: 60 / 80) */
export const TOURFORM_SKALA: Skala = {
  min: 0, max: 100, baender: [
    { von: 0, bis: 60, farbe: G.orange, name: 'Aufbau' },
    { von: 60, bis: 80, farbe: G.gruen, name: 'Auf Kurs' },
    { von: 80, bis: 101, farbe: G.blau, name: 'Bereit' },
  ],
};

/** Kiras Stufen → Garmin-Farben, wenn nur die Stufe bekannt ist */
type Lv = 'blau' | 'gruen' | 'gelb' | 'orange' | 'rot' | 'grau' | 'leer';
export function stufeHex(art: 'hoeher-besser' | 'hrv' | 'last', lv: Lv): string | undefined {
  if (lv === 'leer') return undefined;
  if (lv === 'grau') return G.grau;
  if (art === 'hrv') return lv === 'rot' ? G.rot : lv === 'orange' ? G.orange : G.gruen;
  if (art === 'last') return lv === 'orange' || lv === 'rot' ? G.rot : lv === 'gelb' ? G.orange : G.gruen;
  return lv === 'rot' ? G.rot : lv === 'orange' ? G.orange : lv === 'gelb' ? G.gruen : G.blau;
}
