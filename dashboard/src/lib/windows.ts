import type { Range } from './api';

/** Einheitliche Zeitfenster im ganzen Dashboard. 0 = alle Daten. */
export type WinKey = 14 | 30 | 60 | 120 | 240 | 360 | 0;
export const WINDOWS: WinKey[] = [14, 30, 60, 120, 240, 360, 0];
/** API-Bereich für „alle Daten“ (Apps Script begrenzt ohne Patch auf 365 Tage). */
export const ALL_RANGE: Range = '9999d';

export const winLabel = (w: WinKey) => (w === 0 ? 'Alle' : `${w} T`);
export const winTitle = (w: WinKey) => (w === 0 ? 'alle Daten' : `${w} Tage`);
export function sliceWin<T>(xs: T[], w: WinKey): T[] {
  return w === 0 ? xs : xs.slice(-w);
}

const STEPS: [number, Range][] = [
  [14, '14d'], [30, '30d'], [60, '60d'], [90, '90d'], [120, '120d'], [180, '180d'], [240, '240d'], [360, '360d'],
];
/** Abrufbereich für ein Fenster, optional mit Vorlauf (lead) und Mindestlänge (min) in Tagen. */
export function rangeFor(w: WinKey, lead = 0, min = 0): Range {
  if (w === 0) return ALL_RANGE;
  const n = Math.max(w + lead, min);
  for (const [d, r] of STEPS) if (n <= d) return r;
  return ALL_RANGE;
}
export const rangeDays = (r: Range | null | undefined) => (!r ? 0 : r === ALL_RANGE ? 99999 : Number(r.replace('d', '')));
