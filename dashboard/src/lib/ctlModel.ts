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
// Seit Etappe 1: Formeln kommen aus dem gemeinsamen Rechenkern (eine Quelle für Cockpit, Widgets, Server).
export { ATL_MODEL, CTL_MODEL, atlStep, ctlStep, windowMean } from '../kern/formeln';

/**
 * Näherung Garmin-Ausdauerwert (garminEnduranceScore), gefittet auf KK_TIMELINE
 * (Tage ab 24.12.2025): O_t = 155,91 + 0,96845·O_{t-1} + 0,1537·Load_t + 0,2993·Ø56(Load).
 * Fehler mit echten Lasten: 7 T ±133, 28 T ±206, 56 T ±220 (Bias ≈ 0). Lange Touren erzeugt
 * Garmin-Sprünge (bis ±380), die die Formel nicht abbildet.
 */
export const ES_MODEL = { a: 155.91432, b: 0.96845, c: 0.1537, d: 0.2993, win: 56 };
export function esStep(es: number, load: number, mean56: number) {
  return ES_MODEL.a + ES_MODEL.b * es + ES_MODEL.c * load + ES_MODEL.d * mean56;
}
