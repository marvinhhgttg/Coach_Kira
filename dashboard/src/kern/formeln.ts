/**
 * Coach-Kira-Rechenkern – die eine Quelle für alle Lastformeln.
 * Genutzt von Frontend (ctlModel, Plan-Cockpit, Recovery, Chat), Widgets und Server (Abgleich, Plan-Rechner).
 * 1:1 aus dem Formelblatt `timeline` (Wahrheit), geprüft gegen die Historie ab 26.08.2025.
 *
 * ATL_t = −8,819849 + 0,83645661·ATL_{t−1} + 1,24059231·Last_t
 * CTL_t = 0,7·(−3,95042 + 0,97283824·CTL_{t−1} + 0,2336734·Last_t) + 0,3·7,1·Ø28(Last)
 * ACWR  = ATL / CTL
 */
export const ATL_MODEL = { a: -8.819849, b: 0.83645661, c: 1.24059231 } as const;
export const CTL_MODEL = { a: -3.95042, b: 0.97283824, c: 0.2336734, g: 0.7, k: 7.1, win: 28 } as const;

export function atlStep(atl: number, last: number) {
  return ATL_MODEL.a + ATL_MODEL.b * atl + ATL_MODEL.c * last;
}

/** mittel28: Ø-Tageslast der letzten 28 Tage inkl. heute. */
export function ctlStep(ctl: number, last: number, mittel28: number) {
  const m = CTL_MODEL;
  return m.g * (m.a + m.b * ctl + m.c * last) + (1 - m.g) * m.k * mittel28;
}

/** Ø der letzten `win` Werte (bei kürzerer Reihe: Ø der vorhandenen). */
export function windowMean(reihe: number[], win: number = CTL_MODEL.win) {
  const w = reihe.slice(-win);
  return w.length ? w.reduce((a, b) => a + b, 0) / w.length : 0;
}

/** Monotonie = Ø/σ (Grundgesamtheit); null bei < 2 Werten oder σ = 0 – wie Monotony7 im Sheet. */
export function monotonie(w: number[]): number | null {
  if (w.length < 2) return null;
  const mu = w.reduce((s, x) => s + x, 0) / w.length;
  const sd = Math.sqrt(w.reduce((s, x) => s + (x - mu) ** 2, 0) / w.length);
  return sd === 0 ? null : mu / sd;
}

/** Smart Gains (KEI) = ((CTL − CTL_vor7) · 10) / (MAX(ACWR; 0,8) · (1 + Monotonie)) · 0,1 */
export function smartGains(ctl: number, ctlVor7: number, acwr: number, mono: number) {
  return ((ctl - ctlVor7) * 10) / (Math.max(acwr, 0.8) * (1 + mono)) * 0.1;
}

/** fKEI = MIN(1,15; MAX(0,85; 1 + 0,05·KEI)) */
export const fKEI = (kei: number) => Math.min(1.15, Math.max(0.85, 1 + 0.05 * kei));
