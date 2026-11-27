import type { Range } from './api';

/** Eigener Zeitraum (von–bis, inklusive). */
export interface CustomWin { von: string; bis: string }
/** Einheitliche Zeitfenster im ganzen Dashboard. 0 = alle Daten, Objekt = eigener Zeitraum. */
export type WinPreset = 14 | 30 | 60 | 120 | 240 | 360 | 0;
export type WinKey = WinPreset | CustomWin;
export const WINDOWS: WinPreset[] = [14, 30, 60, 120, 240, 360, 0];
/** API-Bereich für „alle Daten“. */
export const ALL_RANGE: Range = '9999d';
/** Beginn der Datenhistorie. */
export const DATEN_AB = '2025-08-26';

export const isCustom = (w: WinKey): w is CustomWin => typeof w === 'object' && w !== null;
const heute = () => new Date().toLocaleDateString('sv-SE');
const dm = (iso: string, jahr = false) => `${iso.slice(8, 10)}.${iso.slice(5, 7)}.${jahr ? iso.slice(2, 4) : ''}`;
const tageZwischen = (a: string, b: string) => Math.round((Date.parse(b + 'T12:00:00Z') - Date.parse(a + 'T12:00:00Z')) / 86400000);

export const winLabel = (w: WinKey) => (isCustom(w) ? `${dm(w.von)}–${dm(w.bis, true)}` : w === 0 ? 'Alle' : `${w} T`);
export const winTitle = (w: WinKey) =>
  isCustom(w) ? `${dm(w.von, true)} bis ${dm(w.bis, true)} (${tageZwischen(w.von, w.bis) + 1} Tage)` : w === 0 ? 'alle Daten' : `${w} Tage`;
/** Gleichheit für aktive Chips. */
export const winGleich = (a: WinKey, b: WinKey) => (isCustom(a) || isCustom(b) ? isCustom(a) && isCustom(b) && a.von === b.von && a.bis === b.bis : a === b);
/** Zeitraum als von–bis (für datumsbasierte Filter). */
export function winVonBis(w: WinKey): { von: string; bis: string } {
  if (isCustom(w)) return w;
  if (w === 0) return { von: '0000-00-00', bis: '9999-12-31' };
  const d = new Date();
  d.setDate(d.getDate() - w + 1);
  return { von: d.toLocaleDateString('sv-SE'), bis: '9999-12-31' };
}
/** Anzahl Tage im Fenster (für Tagesmittel o. Ä.). */
export const winTage = (w: WinKey) => (isCustom(w) ? tageZwischen(w.von, w.bis) + 1 : w);

const datumVon = (x: any): string | undefined => x?.date ?? x?.datum ?? x?.d?.date;
export function sliceWin<T>(xs: T[], w: WinKey): T[] {
  if (isCustom(w)) return xs.filter((x) => { const d = datumVon(x); return !!d && d >= w.von && d <= w.bis; });
  return w === 0 ? xs : xs.slice(-w);
}

const STEPS: [number, Range][] = [
  [14, '14d'], [30, '30d'], [60, '60d'], [90, '90d'], [120, '120d'], [180, '180d'], [240, '240d'], [360, '360d'],
];
/** Abrufbereich für ein Fenster, optional mit Vorlauf (lead) und Mindestlänge (min) in Tagen. */
export function rangeFor(w: WinKey, lead = 0, min = 0): Range {
  if (!isCustom(w) && w === 0) return ALL_RANGE;
  const tage = isCustom(w) ? tageZwischen(w.von, heute()) + 1 : w;
  const n = Math.max(tage + lead, min);
  for (const [d, r] of STEPS) if (n <= d) return r;
  return ALL_RANGE;
}
export const rangeDays = (r: Range | null | undefined) => (!r ? 0 : r === ALL_RANGE ? 99999 : Number(r.replace('d', '')));
