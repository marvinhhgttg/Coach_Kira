/**
 * Fortschreibung ATL/CTL (nach Belastung des Tages), abgeglichen mit KK_TIMELINE
 * (413 Tage, 26.08.2025–12.10.2026).
 *
 * ATL_t = -8,819849 + 0,83645661·ATL_{t-1} + 1,24059231·Load_t          (unverändert)
 * CTL_t = 0,7·(-3,95042 + 0,97283824·CTL_{t-1} + 0,2336734·Load_t)
 *       + 0,3·7,1·Ø(Load_{t-27..t})                                       (seit 12.10.2026)
 *
 * Die alte reine Rekursion reagierte zu träge (Aufbau: −18 nach 7 T, −52 nach 28 T).
 * Die beobachtete CTL ≈ 7,04 × Ø-Tageslast der letzten 28 Tage; die Mischung behält die
 * gute Kurzfrist-Dynamik der Rekursion und zieht mittelfristig zum 28-Tage-Mittel.
 */
export const ATL_MODEL = { a: -8.819849, b: 0.83645661, c: 1.24059231 };
export const CTL_MODEL = { a: -3.95042, b: 0.97283824, c: 0.2336734, g: 0.7, k: 7.1, win: 28 };

export function atlStep(atl: number, load: number) {
  return ATL_MODEL.a + ATL_MODEL.b * atl + ATL_MODEL.c * load;
}

/** mean28: Ø-Tageslast der letzten 28 Tage inkl. heutigem Load. */
export function ctlStep(ctl: number, load: number, mean28: number) {
  const m = CTL_MODEL;
  return m.g * (m.a + m.b * ctl + m.c * load) + (1 - m.g) * m.k * mean28;
}

/** Ø der letzten `win` Werte (bei kürzerer Reihe: Ø der vorhandenen). */
export function windowMean(series: number[], win = CTL_MODEL.win) {
  const w = series.slice(-win);
  return w.length ? w.reduce((a, b) => a + b, 0) / w.length : 0;
}
